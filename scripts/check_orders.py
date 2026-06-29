import json
import base64
import urllib.request
import urllib.parse
import os
import subprocess
from datetime import datetime, timedelta

import anthropic

CLIENT_ID = os.environ["GMAIL_CLIENT_ID"]
CLIENT_SECRET = os.environ["GMAIL_CLIENT_SECRET"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]

ACCOUNTS = [
    {
        "email": "matt@farmlindproduce.com",
        "refresh_token": os.environ["GMAIL_REFRESH_TOKEN_1"]
    },
    {
        "email": "farmlindproduce@gmail.com",
        "refresh_token": os.environ["GMAIL_REFRESH_TOKEN_2"]
    }
]

VENDORS = {
    "Goodness Gardens": "orders@goodnessgardens.com",
    "Dagele Brothers": "office@dagelebrothersproduce.com",
    "Dottavio": "anthony@mdottavioproduce.com",
}

_test_email = os.environ.get("TEST_EMAIL", "").strip()
NOTIFY_EMAILS = [_test_email] if _test_email else ["matt@farmlindproduce.com", "farmlindproduce@gmail.com"]

REMINDER_SUBJECT = "Order Check -"

claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
STATE_FILE = ".order_state.json"


def env_is_true(name):
    return os.environ.get(name, "").lower() == "true"


def get_access_token(refresh_token):
    data = urllib.parse.urlencode({
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token"
    }).encode()
    req = urllib.request.Request("https://oauth2.googleapis.com/token", data=data)
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())["access_token"]


def gmail_get(access_token, path, params=None):
    url = f"https://gmail.googleapis.com/gmail/v1{path}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {access_token}"})
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def gmail_post(access_token, path, body):
    url = f"https://gmail.googleapis.com/gmail/v1{path}"
    data = json.dumps(body).encode()
    req = urllib.request.Request(
        url, data=data,
        headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def get_todays_emails(access_token, days=1):
    """Get all emails sent in the last N days (default: today only)."""
    emails = []
    try:
        result = gmail_get(access_token, "/users/me/messages", {
            "q": f"in:sent newer_than:{days}d",
            "maxResults": 100
        })
        for m in result.get("messages", []):
            full = gmail_get(access_token, f"/users/me/messages/{m['id']}", {"format": "full"})
            headers = full.get("payload", {}).get("headers", [])
            to_val = next((h["value"] for h in headers if h.get("name","").lower() == "to"), "")
            subj = next((h["value"] for h in headers if h.get("name","").lower() == "subject"), "")
            if to_val and subj:
                emails.append({"to": to_val.lower(), "subject": subj})
    except Exception as e:
        print(f"  Error fetching emails: {e}")

    return emails


def check_vendor_in_email(vendor_email, email_to_header):
    """Check if vendor email is in the To header."""
    return vendor_email.lower() in email_to_header.lower()


def load_state():
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(state):
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(state, f, indent=2)

        try:
            subprocess.run(["git", "config", "user.email", "github-actions@farmlind.local"],
                         capture_output=True, check=False)
            subprocess.run(["git", "config", "user.name", "GitHub Actions"],
                         capture_output=True, check=False)

            subprocess.run(["git", "add", STATE_FILE], check=True, capture_output=True)
            result = subprocess.run(["git", "commit", "-m", "Update order check state"],
                                  capture_output=True, text=True)

            if result.returncode == 0:
                print(f"State committed to git")
                push_result = subprocess.run(["git", "push"], capture_output=True, text=True)
                if push_result.returncode == 0:
                    print("State pushed to remote")
            else:
                if "nothing to commit" not in result.stderr and "nothing to commit" not in result.stdout:
                    print(f"Git commit returned: {result.stderr}")
        except Exception as e:
            print(f"Warning: failed to persist state to git: {e}")
    except Exception as e:
        print(f"Error saving state file: {e}")


def build_reminder_body(vendors_ordered, vendors_not_ordered, today):
    lines = [
        f"Order Check Report - {today.strftime('%A, %B %d, %Y')}",
        "=" * 50,
        ""
    ]

    if vendors_ordered:
        lines.append("VENDORS ORDERED:")
        for vendor in sorted(vendors_ordered):
            lines.append(f"  ✓ {vendor}")
    else:
        lines.append("NO VENDORS ORDERED TODAY")

    lines.append("")

    if vendors_not_ordered:
        lines.append("VENDORS NOT ORDERED:")
        for vendor in sorted(vendors_not_ordered):
            lines.append(f"  ✗ {vendor}")
    else:
        lines.append("ALL VENDORS ORDERED!")

    lines += ["", "- Farmlind Order Bot"]
    return "\n".join(lines)


def send_reminder(access_token, sender_email, vendors_ordered, vendors_not_ordered, today):
    subject = f"{REMINDER_SUBJECT} {today.strftime('%A %b %d')}"
    body = build_reminder_body(vendors_ordered, vendors_not_ordered, today)
    raw = f"From: {sender_email}\r\nTo: {', '.join(NOTIFY_EMAILS)}\r\nSubject: {subject}\r\n\r\n{body}"
    encoded = base64.urlsafe_b64encode(raw.encode()).decode()
    gmail_post(access_token, "/users/me/messages/send", {"raw": encoded})
    print("✓ Reminder sent.")


def main():
    today = datetime.now().date()
    now_ms = int(datetime.now().timestamp() * 1000)
    print(f"=== Order Check for {today.strftime('%A, %B %d, %Y')} ===\n")

    # Get access tokens
    access_tokens = []
    for account in ACCOUNTS:
        try:
            token = get_access_token(account["refresh_token"])
            access_tokens.append((account["email"], token))
            print(f"✓ Authenticated: {account['email']}")
        except Exception as e:
            print(f"✗ Auth failed for {account['email']}: {e}")

    if not access_tokens:
        print("No valid tokens, exiting.")
        return

    print()

    # Fetch today's emails from all accounts
    all_todays_emails = []
    for email_addr, token in access_tokens:
        emails = get_todays_emails(token, days=1)
        all_todays_emails.extend(emails)
        print(f"  Found {len(emails)} emails from {email_addr}")

    if not all_todays_emails:
        print("\n✓ No emails sent today — nothing to check.")
        return

    print(f"\n✓ Total emails today: {len(all_todays_emails)}\n")

    # Check which vendors were ordered
    vendors_ordered = []
    vendors_not_ordered = []

    for vendor_name, vendor_email in VENDORS.items():
        ordered = any(check_vendor_in_email(vendor_email, email["to"]) for email in all_todays_emails)
        if ordered:
            vendors_ordered.append(vendor_name)
            print(f"✓ {vendor_name}: ORDERED")
        else:
            vendors_not_ordered.append(vendor_name)
            print(f"✗ {vendor_name}: NOT ORDERED")

    print()

    # Check if reminder was already sent today
    state = load_state()
    today_key = today.isoformat()
    state_today = state.get(today_key, {})

    already_sent = state_today.get("sent", False)
    force = env_is_true("FORCE_SEND")
    dry = env_is_true("DRY_RUN")
    scan_only = env_is_true("SCAN_ONLY")

    # Track when we first detected emails
    first_email_time = state_today.get("first_email_time")
    if all_todays_emails and not first_email_time:
        # First time we've seen emails today - record the time
        first_email_time = now_ms
        state_today["first_email_time"] = first_email_time
        state[today_key] = state_today
        save_state(state)
        print(f"✓ First email detected at {datetime.fromtimestamp(first_email_time/1000).strftime('%H:%M:%S')}")

    if scan_only:
        print("📊 SCAN MODE: Monitoring for orders...\n")
        if first_email_time:
            time_since_first = (now_ms - first_email_time) / 1000 / 60  # minutes
            print(f"First email was {time_since_first:.1f} minutes ago")
            if time_since_first >= 30:
                print("✓ 30 minutes passed - ready to send email")
            else:
                minutes_left = 30 - time_since_first
                print(f"Will send email in {minutes_left:.1f} minutes\n")
        print(build_reminder_body(vendors_ordered, vendors_not_ordered, today))
        return

    # Decide whether to send
    send_email_now = False

    if force:
        send_email_now = True
    elif already_sent:
        print("✓ Reminder already sent today — nothing to do.")
        return
    elif first_email_time:
        time_since_first = (now_ms - first_email_time) / 1000 / 60  # minutes
        if time_since_first >= 30:
            send_email_now = True
        else:
            print(f"First email was {time_since_first:.1f} minutes ago")
            print(f"Will send reminder in {30 - time_since_first:.1f} minutes")
            return

    # Send reminder if conditions met
    if send_email_now and (vendors_not_ordered or force):
        sender_email, sender_token = access_tokens[0]

        if dry:
            print("--- DRY RUN: email NOT sent ---\n")
            print(build_reminder_body(vendors_ordered, vendors_not_ordered, today))
        else:
            send_reminder(sender_token, sender_email, vendors_ordered, vendors_not_ordered, today)

            # Save state
            state_today["sent"] = True
            state[today_key] = state_today
            save_state(state)
    elif send_email_now:
        print("✓ All vendors ordered — no reminder needed.")
        state_today["sent"] = True
        state[today_key] = state_today
        save_state(state)


if __name__ == "__main__":
    main()
