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
    {"email": "matt@farmlindproduce.com", "refresh_token": os.environ["GMAIL_REFRESH_TOKEN_1"]},
    {"email": "farmlindproduce@gmail.com", "refresh_token": os.environ["GMAIL_REFRESH_TOKEN_2"]}
]

VENDORS = {
    "Goodness Gardens": "orders@goodnessgardens.com",
    "Dagele Brothers": "office@dagelebrothersproduce.com",
    "Dottavio": "anthony@mdottavioproduce.com",
}

_test_email = os.environ.get("TEST_EMAIL", "").strip()
NOTIFY_EMAILS = [_test_email] if _test_email else ["matt@farmlindproduce.com", "farmlindproduce@gmail.com"]

REMINDER_SUBJECT = "Order Check -"
STATE_FILE = ".order_state.json"

claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


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


def decode_b64(data):
    return base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="ignore")


def extract_text_from_part(part):
    mime = part.get("mimeType", "")
    body_data = part.get("body", {}).get("data", "")
    if mime == "text/plain" and body_data:
        return decode_b64(body_data)
    for sub in part.get("parts", []):
        result = extract_text_from_part(sub)
        if result:
            return result
    return ""


def get_todays_emails(access_token, days=1):
    """Get all emails sent in the last N days."""
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
                body = extract_text_from_part(full.get("payload", {}))
                emails.append({"to": to_val.lower(), "subject": subj, "body": body})
    except Exception as e:
        print(f"  Error fetching emails: {e}")
    return emails


def check_vendor_in_email(vendor_email, email_to_header):
    """Check if vendor email is in the To header."""
    return vendor_email.lower() in email_to_header.lower()


def extract_items_from_emails(emails, vendor_email):
    """Use Claude to extract ordered items for a vendor."""
    matching_emails = [e for e in emails if check_vendor_in_email(vendor_email, e["to"])]

    if not matching_emails:
        return []

    combined_text = "\n\n".join([f"Subject: {e['subject']}\n{e['body']}" for e in matching_emails])

    try:
        response = claude.messages.create(
            model="claude-opus-4-8",
            max_tokens=500,
            messages=[{
                "role": "user",
                "content": f"Extract all items/products ordered in these emails. Return as a simple list, one per line:\n\n{combined_text}"
            }]
        )
        items_text = response.content[0].text.strip()
        return [item.strip() for item in items_text.split("\n") if item.strip()]
    except Exception as e:
        print(f"  Error extracting items: {e}")
        return []


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
                subprocess.run(["git", "push"], capture_output=True, text=True)
        except Exception as e:
            print(f"Warning: failed to persist state to git: {e}")
    except Exception as e:
        print(f"Error saving state file: {e}")


def build_reminder_body(vendor_items, vendors_ordered, today):
    lines = [
        f"Order Summary - {today.strftime('%A, %B %d, %Y')}",
        "=" * 50,
        ""
    ]

    for vendor in sorted(vendor_items.keys()):
        items = vendor_items[vendor]
        lines.append(f"{vendor}:")
        if items:
            for item in items:
                lines.append(f"  • {item}")
        else:
            lines.append("  (no items found)")
        lines.append("")

    missing_vendors = [v for v in VENDORS.keys() if v not in vendors_ordered]
    if missing_vendors:
        lines.append("VENDORS NOT ORDERED:")
        for v in missing_vendors:
            lines.append(f"  ✗ {v}")
    else:
        lines.append("✓ All vendors ordered!")

    lines += ["", "- Farmlind Order Bot"]
    return "\n".join(lines)


def send_reminder(access_token, sender_email, vendor_items, vendors_ordered, today):
    subject = f"{REMINDER_SUBJECT} {today.strftime('%A %b %d')}"
    body = build_reminder_body(vendor_items, vendors_ordered, today)
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

    # Fetch today's emails
    all_todays_emails = []
    for email_addr, token in access_tokens:
        emails = get_todays_emails(token, days=1)
        all_todays_emails.extend(emails)
        print(f"  Found {len(emails)} emails from {email_addr}")

    if not all_todays_emails:
        print("\n✓ No emails sent today — nothing to check.")
        return

    print(f"✓ Total emails today: {len(all_todays_emails)}\n")

    # Check which vendors were ordered
    vendors_ordered = []
    vendor_items = {}

    for vendor_name, vendor_email in VENDORS.items():
        ordered = any(check_vendor_in_email(vendor_email, email["to"]) for email in all_todays_emails)
        if ordered:
            vendors_ordered.append(vendor_name)
            print(f"✓ {vendor_name}: ORDERED")
            items = extract_items_from_emails(all_todays_emails, vendor_email)
            vendor_items[vendor_name] = items
            if items:
                print(f"  Items: {', '.join(items[:3])}{'...' if len(items) > 3 else ''}")
        else:
            print(f"✗ {vendor_name}: NOT ORDERED")

    print()

    # Track first email and 30-minute timer
    state = load_state()
    today_key = today.isoformat()
    state_today = state.get(today_key, {})

    first_email_time = state_today.get("first_email_time")
    if all_todays_emails and not first_email_time:
        first_email_time = now_ms
        state_today["first_email_time"] = first_email_time
        state[today_key] = state_today
        save_state(state)
        print(f"First email detected at {datetime.fromtimestamp(first_email_time/1000).strftime('%H:%M:%S')}")

    already_sent = state_today.get("sent", False)
    force = env_is_true("FORCE_SEND")
    dry = env_is_true("DRY_RUN")
    scan_only = env_is_true("SCAN_ONLY")

    if scan_only:
        print("📊 SCAN MODE: Monitoring for orders...\n")
        if first_email_time:
            time_since_first = (now_ms - first_email_time) / 1000 / 60
            print(f"First email was {time_since_first:.1f} minutes ago")
            if time_since_first >= 30:
                print("✓ 30 minutes passed - ready to send email")
            else:
                minutes_left = 30 - time_since_first
                print(f"Will send email in {minutes_left:.1f} minutes")
        return

    # Decide whether to send
    send_email_now = False

    if force:
        send_email_now = True
    elif already_sent:
        print("✓ Reminder already sent today — nothing to do.")
        return
    elif first_email_time:
        time_since_first = (now_ms - first_email_time) / 1000 / 60
        if time_since_first >= 30:
            send_email_now = True
        else:
            print(f"First email was {time_since_first:.1f} minutes ago")
            print(f"Will send reminder in {30 - time_since_first:.1f} minutes")
            return

    # Send reminder
    if send_email_now:
        sender_email, sender_token = access_tokens[0]

        if dry:
            print("--- DRY RUN: email NOT sent ---\n")
            print(build_reminder_body(vendor_items, vendors_ordered, today))
        else:
            send_reminder(sender_token, sender_email, vendor_items, vendors_ordered, today)
            state_today["sent"] = True
            state[today_key] = state_today
            save_state(state)
    else:
        print("✓ Email not ready to send yet.")


if __name__ == "__main__":
    main()
