# app.py
import os
import csv
import json
import threading
import time
import smtplib
import re
import imaplib
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, jsonify, send_from_directory
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr
from utils import read_senders, read_recipients
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__, template_folder='templates')
BASE = os.path.dirname(__file__)
UPLOAD = os.path.join(BASE, "uploads")
os.makedirs(UPLOAD, exist_ok=True)
USERS_JSON = os.path.join(UPLOAD, "users.json")
LOG_FILE = os.path.join(BASE, "send_log.csv")

# in-memory state for sending
state = {
    "senders": [],
    "is_sending": False,
    "paused": False,
    "log": [],
    "notifications": [],
    "current_user": None,
    "current_event_id": None,
    "current_event_subject": None,
    "current_recipient_index": 0,
    "current_event_total": 0,
    "current_event_sent": 0,
    "stop_clear_now": False
}

# helpers
def clean(t):
    return re.sub(r"[\x00-\x1f\x7f-\x9f\xa0]", " ", str(t or "")).strip()

def write_log_row(row):
    header = ["timestamp", "event", "sender", "recipient", "status", "error"]
    write_header = not os.path.exists(LOG_FILE)
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if write_header:
            w.writerow(header)
        w.writerow(row)

def load_users():
    if not os.path.exists(USERS_JSON):
        return {}
    with open(USERS_JSON, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except Exception:
            return {}

def save_users(data):
    with open(USERS_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def notify_user(user_email, ntype, message):
    ts = datetime.now().isoformat()
    state["notifications"].append({"timestamp": ts, "type": ntype, "message": message, "user": user_email})
    state["notifications"] = state["notifications"][-500:]
    users = load_users()
    if user_email:
        users.setdefault(user_email, {})
        users[user_email].setdefault("notifications", [])
        users[user_email]["notifications"].append({"timestamp": ts, "type": ntype, "message": message})
        users[user_email]["notifications"] = users[user_email]["notifications"][-500:]
        save_users(users)

def cleanup_old_data():
    users = load_users()
    cutoff = datetime.now() - timedelta(days=7)
    changed = False
    for email, profile in list(users.items()):
        new_events = []
        for ev in profile.get("events", []):
            created = None
            if "created_at" in ev:
                try:
                    created = datetime.fromisoformat(ev["created_at"])
                except Exception:
                    created = None
            if created is None or created >= cutoff:
                new_events.append(ev)
            else:
                changed = True
        profile["events"] = new_events
        new_hist = []
        for h in profile.get("history", []):
            try:
                t = datetime.fromisoformat(h.get("time"))
                if t >= cutoff:
                    new_hist.append(h)
                else:
                    changed = True
            except Exception:
                changed = True
        profile["history"] = new_hist
        users[email] = profile
    if changed:
        save_users(users)

cleanup_old_data()

@app.route("/")
def index():
    return render_template("index.html")

# ---------- user management ----------
@app.route("/login_user", methods=["POST"])
def login_user():
    data = request.json or {}
    name = clean(data.get("name"))
    email = clean(data.get("email")).lower()
    if not email:
        return jsonify({"ok": False, "error": "Email required"}), 400
    users = load_users()
    if email not in users:
        users[email] = {"name": name or email.split("@")[0], "email": email, "created": datetime.now().isoformat(), "events": [], "history": [], "pitch": "", "notifications": []}
        save_users(users)
    resp = jsonify({"ok": True, "profile": users[email]})
    resp.set_cookie("user_email", email, max_age=30*24*3600)
    return resp

@app.route("/get_profile")
def get_profile():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    users = load_users()
    profile = users.get(user_email)
    if not profile:
        return jsonify({"ok": False})
    return jsonify({"ok": True, "profile": profile})

@app.route("/delete_profile", methods=["POST"])
def delete_profile():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    users = load_users()
    if user_email in users:
        del users[user_email]
        save_users(users)
    resp = jsonify({"ok": True})
    resp.set_cookie("user_email", "", expires=0)
    return resp

# ---------- pitch endpoints ----------
@app.route("/save_pitch", methods=["POST"])
def save_pitch():
    data = request.get_json() or {}
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False, "error": "Not logged in"}), 401
    users = load_users()
    users.setdefault(user_email, {})
    users[user_email]["pitch"] = data.get("pitch", "")
    save_users(users)
    return jsonify({"ok": True})

@app.route("/get_pitch")
def get_pitch():
    user_email = request.cookies.get("user_email")
    users = load_users()
    if not user_email:
        return jsonify({"ok": False})
    pitch = users.get(user_email, {}).get("pitch", "")
    return jsonify({"ok": True, "pitch": pitch})

# ---------- events CRUD ----------
def generate_event_id():
    return f"evt_{int(time.time()*1000)}"

@app.route("/save_event", methods=["POST"])
def save_event():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False, "error": "Not logged in"}), 401
    d = request.json or {}
    users = load_users()
    profile = users.get(user_email, {"events": [], "history": [], "notifications": [], "pitch": ""})
    ev_id = d.get("id") or generate_event_id()
    event_obj = {
        "id": ev_id,
        "event": clean(d.get("event")),
        "date": clean(d.get("date")),
        "location": clean(d.get("location")),
        "count": clean(d.get("count")),
        "subject": clean(d.get("subject") or ""),
        "recipients_filename": d.get("recipients_filename") or "",
        "recipients_count": int(d.get("recipients_count") or 0),
        "created_at": datetime.now().isoformat()
    }
    replaced = False
    for i, ex in enumerate(profile.get("events", [])):
        if ex.get("id") == ev_id:
            profile["events"][i] = event_obj
            replaced = True
            break
    if not replaced:
        profile.setdefault("events", []).append(event_obj)
    users[user_email] = profile
    save_users(users)
    return jsonify({"ok": True, "event": event_obj})

@app.route("/list_events")
def list_events():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    users = load_users()
    profile = users.get(user_email, {})
    return jsonify({"ok": True, "events": profile.get("events", [])})

@app.route("/delete_event", methods=["POST"])
def delete_event():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    d = request.json or {}
    ev_id = d.get("id")
    users = load_users()
    profile = users.get(user_email)
    if not profile:
        return jsonify({"ok": False})
    profile["events"] = [e for e in profile.get("events", []) if e.get("id") != ev_id]
    users[user_email] = profile
    save_users(users)
    return jsonify({"ok": True})

# ---------- history endpoints ----------
@app.route("/history")
def history():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    users = load_users()
    profile = users.get(user_email, {})
    return jsonify({"ok": True, "history": profile.get("history", [])})

@app.route("/delete_history", methods=["POST"])
def delete_history():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    d = request.json or {}
    idx = d.get("index")
    users = load_users()
    profile = users.get(user_email, {})
    hist = profile.get("history", [])
    if idx is not None and 0 <= int(idx) < len(hist):
        hist.pop(int(idx))
    profile["history"] = hist
    users[user_email] = profile
    save_users(users)
    return jsonify({"ok": True})

@app.route("/clear_history", methods=["POST"])
def clear_history():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    users = load_users()
    profile = users.get(user_email, {})
    profile["history"] = []
    users[user_email] = profile
    save_users(users)
    return jsonify({"ok": True})

# ---------- uploads ----------
@app.route("/upload_recipients_file", methods=["POST"])
def upload_recipients_file():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False, "error": "Not logged in"}), 401
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "No file"}), 400
    fname = f"{int(time.time())}_{clean(f.filename)}"
    path = os.path.join(UPLOAD, fname)
    f.save(path)
    try:
        recs = read_recipients(path)
        count = len(recs)
    except Exception:
        count = 0
    return jsonify({"ok": True, "filename": fname, "count": count})

@app.route("/upload_senders", methods=["POST"])
def upload_senders():
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "No file"}), 400
    path = os.path.join(UPLOAD, "senders.xlsx")
    f.save(path)
    try:
        s = read_senders(path)
        state["senders"] = [{
            "email": clean(x["email"]),
            "api_key": clean(x.get("api_key")),
            "name": clean(x.get("name") or x.get("email")),
            "paused": False
        } for x in s]
        return jsonify({"ok": True, "count": len(s)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

# ---------- notifications endpoints ----------
@app.route("/notifications")
def notifications():
    user_email = request.cookies.get("user_email")
    users = load_users()
    user_notifs = users.get(user_email, {}).get("notifications", []) if user_email else []
    live = [n for n in state.get("notifications", []) if n.get("user") == user_email]
    merged = (user_notifs or []) + live
    return jsonify({"ok": True, "user": user_email, "notifications": merged})

@app.route("/notifications/delete", methods=["POST"])
def notifications_delete():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    data = request.json or {}
    ts = data.get("timestamp")
    users = load_users()
    profile = users.get(user_email, {})
    profile["notifications"] = [n for n in profile.get("notifications", []) if n.get("timestamp") != ts]
    users[user_email] = profile
    save_users(users)
    state["notifications"] = [n for n in state["notifications"] if not (n.get("user")==user_email and n.get("timestamp")==ts)]
    return jsonify({"ok": True})

@app.route("/clear_user_notifications", methods=["POST"])
def clear_user_notifications():
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False})
    users = load_users()
    profile = users.get(user_email, {})
    profile["notifications"] = []
    users[user_email] = profile
    save_users(users)
    state["notifications"] = [n for n in state["notifications"] if n.get("user") != user_email]
    return jsonify({"ok": True})

# ---------- send control endpoints ----------
@app.route("/pause", methods=["POST"])
def pause():
    if not state.get("is_sending"):
        return jsonify({"ok": False, "error": "Not sending"}), 400
    state["paused"] = True
    return jsonify({"ok": True})

@app.route("/resume", methods=["POST"])
def resume():
    if not state.get("is_sending"):
        return jsonify({"ok": False, "error": "Not sending"}), 400
    state["paused"] = False
    return jsonify({"ok": True})

@app.route("/stop_clear", methods=["POST"])
def stop_clear():
    if not state.get("is_sending"):
        return jsonify({"ok": False, "error": "Not sending"}), 400
    state["stop_clear_now"] = True
    return jsonify({"ok": True})

@app.route("/start_send_queue", methods=["POST"])
def start_send_queue():
    if state.get("is_sending"):
        return jsonify({"ok": False, "error": "Already sending"}), 400
    user_email = request.cookies.get("user_email")
    if not user_email:
        return jsonify({"ok": False, "error": "Not logged in"}), 401
    d = request.json or {}
    try:
        gap = float(d.get("gap", 5))
    except Exception:
        gap = 5.0
    subject_template = clean(d.get("subject") or "")
    highlight = d.get("color") or "#d6336c"
    txtcolor = d.get("txtcolor") or "#1f2937"

    users_all = load_users()
    profile = users_all.get(user_email)
    if not profile:
        return jsonify({"ok": False, "error": "Profile not found"}), 400
    events = profile.get("events", [])
    senders = state.get("senders", [])
    if not senders:
        return jsonify({"ok": False, "error": "Senders not uploaded"}), 400
    shared_body = profile.get("pitch") or body_default_text()

    state["is_sending"] = True
    state["paused"] = False
    state["current_user"] = user_email
    state["current_event_id"] = None
    state["current_event_subject"] = None
    state["current_recipient_index"] = 0
    state["current_event_total"] = 0
    state["current_event_sent"] = 0
    state["stop_clear_now"] = False

    def send_loop():
        nonlocal gap, subject_template, highlight, txtcolor, shared_body
        while state.get("is_sending"):
            users_local_all = load_users()
            profile_local = users_local_all.get(user_email, {})
            events_local = profile_local.get("events", [])
            if not events_local:
                break
            ev = events_local[0]
            ev_id = ev.get("id")
            state["current_event_id"] = ev_id
            state["current_event_subject"] = ev.get("subject") or subject_template or f"Attendee list - {ev.get('event')}"
            recipients_file = ev.get("recipients_filename")
            recipients_path = os.path.join(UPLOAD, recipients_file) if recipients_file else None
            try:
                recipients = read_recipients(recipients_path) if recipients_path else []
            except Exception as e:
                recipients = []
                state["log"].append(f"[{datetime.now().isoformat()}] ERROR reading recipients for {ev.get('event')}: {e}")
            total_recipients = len(recipients)
            state["current_event_total"] = total_recipients
            idx = state.get("current_recipient_index", 0)
            sent_count_for_event = 0

            while idx < total_recipients:
                if not state.get("is_sending"):
                    break
                if state.get("paused"):
                    time.sleep(0.5)
                    continue
                if state.get("stop_clear_now"):
                    ts = datetime.now().isoformat()
                    profile_local.setdefault("history", []).append({
                        "event": ev.get("event"),
                        "recipients": total_recipients,
                        "sent": sent_count_for_event,
                        "time": ts
                    })
                    users_all = load_users()
                    user_profile = users_all.get(user_email, {})
                    user_profile["events"] = [e for e in user_profile.get("events", []) if e.get("id") != ev_id]
                    users_all[user_email] = user_profile
                    save_users(users_all)
                    state["stop_clear_now"] = False
                    state["current_recipient_index"] = 0
                    break

                active = [s for s in state.get("senders", []) if not s.get("paused")]
                if not active:
                    msg = "All senders paused — stopping"
                    state["log"].append(f"[{datetime.now().isoformat()}] {msg}")
                    notify_user(user_email, "critical", msg)
                    state["is_sending"] = False
                    break
                sender = active[idx % len(active)]
                recipient = recipients[idx]
                rcpt_email = clean(recipient.get("email"))
                first_name = clean(recipient.get("first_name") or "")

                body_html = (shared_body
                             .replace("{event}", f"<span style='color:{highlight};font-weight:bold'>{ev.get('event')}</span>")
                             .replace("{date}", f"<span style='color:{highlight};font-weight:bold'>{ev.get('date')}</span>")
                             .replace("{location}", f"<span style='color:{highlight};font-weight:bold'>{ev.get('location')}</span>")
                             .replace("{count}", f"<span style='color:{highlight};font-weight:bold'>{ev.get('count')}</span>")
                             .replace("{first_name}", f"<span style='color:{highlight};font-weight:bold'>{first_name}</span>")
                             .replace("{sender_name}", f"<span style='color:{highlight};font-weight:bold'>{sender.get('name')}</span>")
                             .replace("\n", "<br>"))
                html = f"<div style='color:{txtcolor};font-family:Segoe UI,Arial,sans-serif'>{body_html}</div>"

                try:
                    msg = MIMEMultipart("alternative")
                    msg["From"] = formataddr((str(Header(sender.get("name", sender.get("email")), "utf-8")), sender.get("email")))
                    msg["To"] = rcpt_email
                    subj = ev.get("subject") or subject_template or f"Attendee list - {ev.get('event')}"
                    msg["Subject"] = Header(subj.replace("{event}", ev.get("event")), "utf-8")
                    msg.attach(MIMEText(html, "html", "utf-8"))

                    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=25) as sv:
                        sv.login(sender.get("email"), sender.get("api_key"))
                        sv.send_message(msg)

                    ts = datetime.now().isoformat()
                    state["log"].append(f"[{ts}] {rcpt_email} sent through {sender.get('email')}")
                    write_log_row([ts, ev.get("event"), sender.get("email"), rcpt_email, "sent", ""])
                    sent_count_for_event += 1
                    state["current_event_sent"] = sent_count_for_event
                except Exception as e:
                    err = str(e)
                    ts = datetime.now().isoformat()
                    state["log"].append(f"[{ts}] ERROR {rcpt_email} via {sender.get('email')}: {err}")
                    write_log_row([ts, ev.get("event"), sender.get("email"), rcpt_email, "error", err])
                    if any(k in err.lower() for k in ["quota", "daily", "535", "5.7.1", "authentication"]):
                        for s in state["senders"]:
                            if s.get("email") == sender.get("email"):
                                s["paused"] = True
                                break
                        notify_user(user_email, "warning", f"Sender {sender.get('email')} paused — {err}")
                        active = [s for s in state.get("senders", []) if not s.get("paused")]
                        if not active:
                            notify_user(user_email, "critical", "All senders paused — stopping send loop")
                            state["is_sending"] = False
                            break
                idx += 1
                state["current_recipient_index"] = idx

                slept = 0.0
                while slept < gap:
                    if not state["is_sending"] or state.get("paused") or state.get("stop_clear_now"):
                        break
                    time.sleep(0.5)
                    slept += 0.5

            # finished or interrupted
            if idx >= total_recipients:
                ts = datetime.now().isoformat()
                profile_local.setdefault("history", []).append({
                    "event": ev.get("event"),
                    "recipients": total_recipients,
                    "sent": sent_count_for_event,
                    "time": ts
                })
                users_all = load_users()
                user_profile = users_all.get(user_email, {})
                user_profile["events"] = [e for e in user_profile.get("events", []) if e.get("id") != ev_id]
                users_all[user_email] = user_profile
                save_users(users_all)
                state["current_recipient_index"] = 0
                state["current_event_sent"] = 0
                state["current_event_total"] = 0
            else:
                # partial - keep index for resume
                pass

            time.sleep(0.5)

        # finish / cleanup
        state["is_sending"] = False
        state["paused"] = False
        state["current_user"] = None
        state["current_event_id"] = None
        state["current_event_subject"] = None
        state["current_recipient_index"] = 0
        state["current_event_total"] = 0
        state["current_event_sent"] = 0

    t = threading.Thread(target=send_loop, daemon=True)
    t.start()
    return jsonify({"ok": True})

@app.route("/stop", methods=["POST"])
def stop_endpoint():
    state["is_sending"] = False
    state["paused"] = False
    return jsonify({"ok": True})

@app.route("/status")
def status():
    cl = state.get("log", [])[-500:]
    user_email = state.get("current_user")
    cur_event_name = None
    cur_event_id = state.get("current_event_id")
    total = state.get("current_event_total", 0)
    sent_count = state.get("current_event_sent", 0)
    if user_email and cur_event_id:
        users = load_users()
        profile = users.get(user_email, {})
        for ev in profile.get("events", []):
            if ev.get("id") == cur_event_id:
                cur_event_name = ev.get("event")
                total = ev.get("recipients_count") or total
                break
    queue_total = 0
    if user_email:
        users = load_users()
        profile = users.get(user_email, {})
        for ev in profile.get("events", []):
            queue_total += int(ev.get("recipients_count") or 0)
    return jsonify({
        "is_sending": state.get("is_sending", False),
        "is_paused": state.get("paused", False),
        "log": cl,
        "current_event": cur_event_name,
        "current_event_id": cur_event_id,
        "current_subject": state.get("current_event_subject"),
        "sent_count": sent_count,
        "current_total": total,
        "queue_total_recipients": queue_total,
        "notifications": state.get("notifications", [])[-200:]
    })

@app.route("/clear_log", methods=["POST"])
def clear_log():
    # clear in-memory log and remove file if exists
    state["log"] = []
    try:
        if os.path.exists(LOG_FILE):
            os.remove(LOG_FILE)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    return jsonify({"ok": True})

@app.route("/download_log")
def download_log():
    if os.path.exists(LOG_FILE):
        return send_from_directory(os.path.dirname(LOG_FILE), os.path.basename(LOG_FILE), as_attachment=True)
    return "No log", 404

@app.route("/quota")
def quota():
    if not state.get("senders"):
        return jsonify({"error": "No senders uploaded"}), 400
    results = []
    for s in state.get("senders", []):
        email = s.get("email"); pw = s.get("api_key"); used, limit = 0, 500
        try:
            imap = imaplib.IMAP4_SSL("imap.gmail.com")
            imap.login(email, pw)
            imap.select('"[Gmail]/Sent Mail"')
            typ, data = imap.search(None, "SINCE", date.today().strftime("%d-%b-%Y"))
            if typ == "OK":
                used = len(data[0].split()) if data and data[0] else 0
            imap.logout()
        except Exception:
            pass
        results.append({"email": email, "used": used, "daily_limit": limit, "remaining": max(0, limit - used)})
    return jsonify(results)

@app.route("/notifications/clear", methods=["POST"])
def clear_notifications():
    state["notifications"] = []
    return jsonify({"ok": True})

@app.route("/admin/cleanup", methods=["POST"])
def admin_cleanup():
    cleanup_old_data()
    return jsonify({"ok": True})

def body_default_text():
    return ("Hi,\n\nWe are pleased to inform you that the attendee {event}\n{date}\n{location}\n\n"
            "The list contains {count} pre-registered attendees with email addresses, contact numbers, and more in an Excel format.\n"
            "We do charge for our services; would you like to See What is the price?\n\nRegards,\n{sender_name}")

if __name__ == "__main__":
    app.run(debug=True, threaded=True)
