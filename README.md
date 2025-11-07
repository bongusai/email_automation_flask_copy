
# Email Automation - Flask + SendGrid (email_automation_flask)

This project implements a web UI (Flask) to upload sender and recipient Excel files and send emails using SendGrid API keys (recommended for deliverability).

## Features
- Upload `sender_emails.xlsx` and `recipient_emails.xlsx`
- Rotate senders: recipient1 → sender1, recipient2 → sender2, ... repeat
- Gap between emails (seconds)
- Subject and templated body with placeholders: {event}, {date}, {location}, {count}, {first_name}
- Start / Stop sending from UI, live log, clear history
- Saves logs to `send_log.csv`
- Marks senders as paused when API returns rate-limit/authorization errors

## File formats
### sender_emails.xlsx
Columns (headers are case-insensitive):
- email (required) — "From" address
- api_key (required) — SendGrid API key for this sender (or leave blank to use default SENDGRID_API_KEY)
- name (optional) — sender name shown in "From"

Example row:
```
your.sender@example.com, SG.xxxxx..., Sender Name
```

### recipient_emails.xlsx
- email (required)
- first_name (optional)

## Setup
1. Create a Python virtual environment and activate it.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Create a `.env` file with (optional) default SendGrid API key:
   ```
   SENDGRID_API_KEY=SG.xxxxxx
   ```
4. Run the app:
   ```bash
   python app.py
   ```
5. Open `http://127.0.0.1:5000` in your browser.

## Notes on SendGrid
- Using SendGrid avoids Gmail SMTP quotas and improves deliverability.
- SendGrid has free tier limits; for high volume consider a paid plan.
- Configure domain authentication (SPF/DKIM) in SendGrid for best deliverability.

## Security
- Keep API keys private. Do NOT commit `.env` or your keys to public repos.
