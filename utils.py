import pandas as pd
import validators

def read_senders(path):
    df = pd.read_excel(path, engine='openpyxl')
    df.columns = [c.strip().lower() for c in df.columns]

    if 'email' not in df.columns:
        raise ValueError("sender_emails.xlsx must contain an 'email' column")

    # For Gmail setup, app password column can be named either 'api_key' or 'app_password'
    if 'api_key' not in df.columns and 'app_password' in df.columns:
        df['api_key'] = df['app_password']
    elif 'api_key' not in df.columns:
        df['api_key'] = None

    # daily_limit optional
    if 'daily_limit' not in df.columns:
        df['daily_limit'] = None

    if 'name' not in df.columns:
        df['name'] = ''

    senders = df.to_dict(orient='records')

    for s in senders:
        # Clean name for encoding issues
        if 'name' in s:
            s['name'] = str(s['name']).replace('\xa0', ' ')
        email = str(s.get('email'))
        if not validators.email(email):
            raise ValueError(f"Invalid sender email: {email}")
        # warn if no api key
        if not s.get('api_key'):
            print(f"⚠️ Warning: Sender '{email}' has no App Password/API key. It may fail to send.")
    return senders

def read_recipients(path):
    df = pd.read_excel(path, engine='openpyxl')
    df.columns = [c.strip().lower() for c in df.columns]

    if 'email' not in df.columns:
        cols = df.columns.tolist()
        if len(cols) >= 1:
            df = df.rename(columns={cols[0]: 'email'})
        else:
            raise ValueError("recipient_emails.xlsx must contain an 'email' column")

    if 'first_name' not in df.columns:
        df['first_name'] = ''

    recipients = df.to_dict(orient='records')

    for r in recipients:
        if 'first_name' in r:
            r['first_name'] = str(r['first_name']).replace('\xa0', ' ')
        email = str(r.get('email'))
        if not validators.email(email):
            raise ValueError(f"Invalid recipient email: {email}")

    return recipients
