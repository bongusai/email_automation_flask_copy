from flask import Flask, render_template, request, jsonify, send_from_directory
import os, csv, threading, time, smtplib, re, imaplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr
from utils import read_senders, read_recipients
from datetime import date
from dotenv import load_dotenv

load_dotenv()
app = Flask(__name__)
BASE = os.path.dirname(__file__)
UPLOAD = os.path.join(BASE, 'uploads')
os.makedirs(UPLOAD, exist_ok=True)
LOG = os.path.join(BASE, 'send_log.csv')

state = {
    'senders': [],
    'recipients': [],
    'is_sending': False,
    'log': [],
    'notifications': [],
    'progress_index': 0,
    'resumable': False
}

def clean(t):
    return re.sub(r'[\x00-\x1f\x7f-\x9f\xa0]', ' ', str(t or '')).strip()

def write_log(r):
    head = ['timestamp','event','sender','recipient','status','error']
    new = not os.path.exists(LOG)
    with open(LOG, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if new: w.writerow(head)
        w.writerow(r)

def notify(tp, msg):
    from datetime import datetime
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    state['notifications'].append({'timestamp': ts, 'type': tp, 'message': msg})
    state['notifications'] = state['notifications'][-200:]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_senders', methods=['POST'])
def upload_senders():
    f = request.files['file']
    path = os.path.join(UPLOAD, 'senders.xlsx')
    f.save(path)
    s = read_senders(path)
    for x in s:
        x['email'] = clean(x['email'])
        x['api_key'] = clean(x.get('api_key'))
        x['name'] = clean(x.get('name') or x['email'])
    state['senders'] = s
    return jsonify({'count': len(s)})

@app.route('/upload_recipients', methods=['POST'])
def upload_recipients():
    f = request.files['file']
    path = os.path.join(UPLOAD, 'recipients.xlsx')
    f.save(path)
    r = read_recipients(path)
    for x in r:
        x['email'] = clean(x['email'])
        x['first_name'] = clean(x.get('first_name'))
    state['recipients'] = r
    return jsonify({'count': len(r)})

@app.route('/start', methods=['POST'])
def start():
    # Reset log only if this is a new session (not a resume)
    if not state.get('resumable', False):
        if os.path.exists(LOG):
            os.remove(LOG)
        state['progress_index'] = 0

    if state['is_sending']:
        return jsonify({'status': 'already'})

    d = request.json
    gap = float(d.get('gap', 5))
    event, datex, loc, count = [clean(d.get(k)) for k in ['event', 'date', 'location', 'count']]
    subj = clean(d.get('subject')) or f"Attendee list - {event}"
    template = d.get('template') or ''
    highlight = d.get('color') or '#d6336c'
    txtcolor = d.get('txtcolor') or '#1f2937'

    senders = [{'email': clean(s['email']), 'pw': clean(s['api_key']), 'name': clean(s['name']), 'paused': False}
               for s in state['senders']]
    recipients = state['recipients'][:]
    if not senders or not recipients:
        return jsonify({'status': 'error'})

    current_index = state.get('progress_index', 0)
    state['is_sending'] = True
    state['resumable'] = True

    def loop():
        idx = current_index
        total_senders = len(senders)
        for i in range(idx, len(recipients)):
            if not state['is_sending']:
                # Save progress when stopped
                state['progress_index'] = i
                break

            r = recipients[i]
            rcpt = clean(r.get('email'))
            fname = clean(r.get('first_name'))
            s = senders[i % total_senders]

            # Skip paused senders
            active_senders = [x for x in senders if not x['paused']]
            if not active_senders:
                ts = time.strftime('%Y-%m-%d %H:%M:%S')
                msg = "All senders paused (Gmail limit reached) — stopping."
                state['log'].append(f"[{ts}] {msg}")
                notify('critical', msg)
                state['is_sending'] = False
                break
            s = active_senders[i % len(active_senders)]

            try:
                html_body = (
                    template.replace('{event}', f"<span style='color:{highlight};font-weight:bold'>{event}</span>")
                            .replace('{date}', f"<span style='color:{highlight};font-weight:bold'>{datex}</span>")
                            .replace('{location}', f"<span style='color:{highlight};font-weight:bold'>{loc}</span>")
                            .replace('{count}', f"<span style='color:{highlight};font-weight:bold'>{count}</span>")
                            .replace('{first_name}', f"<span style='color:{highlight};font-weight:bold'>{fname}</span>")
                            .replace('{sender_name}', f"<span style='color:{highlight};font-weight:bold'>{s['name']}</span>")
                            .replace('\n', '<br>')
                )
                html = f"<div style='color:{txtcolor};font-family:Segoe UI,Arial,sans-serif'>{html_body}</div>"

                msg = MIMEMultipart("alternative")
                msg["From"] = formataddr((str(Header(s['name'], 'utf-8')), s['email']))
                msg["To"] = rcpt
                msg["Subject"] = Header(subj.replace('{event}', event), 'utf-8')
                msg.attach(MIMEText(html, "html", "utf-8"))

                with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=25) as sv:
                    sv.login(s['email'], s['pw'])
                    sv.send_message(msg)

                ts = time.strftime('%Y-%m-%d %H:%M:%S')
                state['log'].append(f"[{ts}] {rcpt} sent through {s['email']}")
                write_log([ts, event, s['email'], rcpt, 'sent', ''])
            except Exception as e:
                err = str(e).encode('utf-8', 'ignore').decode('utf-8')
                ts = time.strftime('%Y-%m-%d %H:%M:%S')
                state['log'].append(f"[{ts}] ERROR {rcpt} via {s['email']}: {err}")
                write_log([ts, event, s['email'], rcpt, 'error', err])

                # ✅ Detect Gmail limit or auth issue and pause this sender
                if any(k in err.lower() for k in ['daily user sending limit', 'quota', '535', '5.7.1']):
                    s['paused'] = True
                    msg = f"⚠️ Sender {s['email']} paused (Gmail limit or quota reached)"
                    state['log'].append(f"[{ts}] {msg}")
                    notify('warning', msg)

                    # If all senders paused → stop
                    if all(x['paused'] for x in senders):
                        msg2 = "All senders paused (Gmail limits reached) — Stopping."
                        state['log'].append(f"[{ts}] {msg2}")
                        notify('critical', msg2)
                        state['is_sending'] = False
                        break

            slept = 0.0
            while slept < gap:
                if not state['is_sending']:
                    state['progress_index'] = i + 1
                    break
                time.sleep(0.5)
                slept += 0.5

        if state['is_sending']:
            state['log'].append('Finished sending.')
            state['progress_index'] = len(recipients)
        state['is_sending'] = False

    threading.Thread(target=loop, daemon=True).start()
    return jsonify({'status': 'started'})

@app.route('/stop', methods=['POST'])
def stop():
    state['is_sending'] = False
    return jsonify({'stopped': 1})

@app.route('/new_event', methods=['POST'])
def new_event():
    """Reset all memory, log, and progress"""
    state.update({
        'senders': [],
        'recipients': [],
        'is_sending': False,
        'log': [],
        'notifications': [],
        'progress_index': 0,
        'resumable': False
    })
    if os.path.exists(LOG):
        os.remove(LOG)
    return jsonify({'status': 'reset'})

@app.route('/status')
def status():
    sent_count = sum(1 for line in state['log'] if 'sent through' in line)
    total_count = len(state['recipients'])
    return jsonify({
        'is_sending': state['is_sending'],
        'log': state['log'][-200:],
        'sent_count': sent_count,
        'total_count': total_count
    })

@app.route('/download_log')
def download_log():
    if os.path.exists(LOG):
        return send_from_directory(os.path.dirname(LOG), os.path.basename(LOG), as_attachment=True)
    return 'No log', 404

@app.route('/quota')
def quota():
    """Fetch Gmail quota via IMAP login per sender"""
    results = []
    for s in state['senders']:
        email = s['email']; pw = s['api_key']
        used, limit = 0, 500
        try:
            imap = imaplib.IMAP4_SSL("imap.gmail.com")
            imap.login(email, pw)
            imap.select('"[Gmail]/Sent Mail"')
            typ, data = imap.search(None, 'SINCE', date.today().strftime("%d-%b-%Y"))
            if typ == 'OK':
                msgs = len(data[0].split())
                used = msgs
            imap.logout()
        except Exception:
            pass
        remaining = max(0, limit - used)
        results.append({'email': email, 'used': used, 'daily_limit': limit, 'remaining': remaining})
    return jsonify(results)

@app.route('/notifications')
def get_notifications():
    return jsonify(state['notifications'])

@app.route('/notifications/clear', methods=['POST'])
def clear_notifications():
    state['notifications'] = []
    return jsonify({'ok': 1})

if __name__ == '__main__':
    app.run(debug=True)
