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


def fetch_emails_from_vendor(access_token, vendor_email, days=30):
    """Fetch emails from last N days."""
    emails = []
    try:
        result = gmail_get(access_token, "/users/me/messages", {
            "q": f"from:{vendor_email} newer_than:{days}d",
            "maxResults": 100
        })
        for m in result.get("messages", []):
            full = gmail_get(access_token, f"/users/me/messages/{m['id']}", {"format": "full"})
            headers = full.get("payload", {}).get("headers", [])
            to_val = next((h["value"] for h in headers if h.get("name","").lower() == "to"), "")
            subj = next((h["value"] for h in headers if h.get("name","").lower() == "subject"), "")
            if to_val and subj:
                body = extract_text_from_part(full.get("payload", {}))
                ts = int(full.get("internalDate", 0))
                emails.append({"to": to_val.lower(), "subject": subj, "body": body, "ts": ts})
    except Exception as e:
        print(f"  Error fetching emails: {e}")
    return emails


def vendor_in_email(vendor_email, email_to):
    """Check if vendor is in the To header."""
    return vendor_email.lower() in email_to.lower()


def extract_items(emails, vendor_email):
    """Extract items ordered for a vendor using Claude."""
    matching = [e for e in emails if vendor_in_email(vendor_email, e["to"])]
    if not matching:
        return []

    combined = "\n\n".join([f"Subject: {e['subject']}\n{e['body']}" for e in matching])

    try:
        response = claude.messages.create(
            model="claude-opus-4-8",
            max_tokens=500,
            messages=[{
                "role": "user",
                "content": f"Extract all items/products ordered. Return as list, one per line:\n\n{combined}"
            }]
        )
        items = response.content[0].text.strip().split("\n")
        return [i.strip() for i in items if i.strip()]
    except Exception as e:
        print(f"  Error extracting: {e}")
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


def build_email(today_items, prev_items, today):
    lines = [
        "ORDER CHECK REPORT",
        f"{today.strftime('%A, %B %d, %Y')}",
        "=" * 70,
        ""
    ]

    for vendor in sorted(VENDORS.keys()):
        lines.append("")
        lines.append("=" * 70)
        lines.append(f"{vendor.upper()}")
        lines.append("=" * 70)

        today_list = today_items.get(vendor, [])
        prev_list = prev_items.get(vendor, [])

        if today_list:
            lines.append("")
            lines.append("ORDERED TODAY:")
            for i, item in enumerate(today_list, 1):
                lines.append(f"  {i}. {item}")
        else:
            lines.append("")
            lines.append("ORDERED TODAY: NOTHING")

        if prev_list:
            missing = [item for item in prev_list if item not in today_list]
            if missing:
                lines.append("")
                lines.append("MISSING (ordered before, not today):")
                for i, item in enumerate(missing, 1):
                    lines.append(f"  {i}. {item}")
            else:
                lines.append("")
                lines.append("MISSING: Nothing - all previous items ordered!")
        else:
            lines.append("")
            lines.append("MISSING: No previous orders to compare")

    lines.append("")
    lines.append("=" * 70)
    lines.append("")
    lines.append("- Farmlind Order Bot")
    return "
".join(lines)
".join(lines)


def send_email(access_token, sender, today_items, prev_items, today):
    subject = f"{REMINDER_SUBJECT} {today.strftime('%A %b %d')}"
    body = build_email(today_items, prev_items, today)
    raw = f"From: {sender}\r\nTo: {', '.join(NOTIFY_EMAILS)}\r\nSubject: {subject}\r\n\r\n{body}"
    encoded = base64.urlsafe_b64encode(raw.encode()).decode()
    gmail_post(access_token, "/users/me/messages/send", {"raw": encoded})
    print("✓ Reminder sent.")


def main():
    today = datetime.now().date()
    now_ms = int(datetime.now().timestamp() * 1000)
    print(f"=== Order Check for {today.strftime('%A, %B %d, %Y')} ===\n")

    # Authenticate both accounts
    access_tokens = []
    for account in ACCOUNTS:
        try:
            token = get_access_token(account["refresh_token"])
            access_tokens.append((account["email"], token))
            print(f"✓ Authenticated: {account['email']}")
        except Exception as e:
            print(f"✗ Auth failed for {account['email']}: {e}")

    if not access_tokens:
        print("No valid tokens.")
        return

    print()

    # Get TODAY's emails from both accounts
    all_today = []
    for email_addr, token in access_tokens:
        emails = fetch_emails(token, days=1)
        all_today.extend(emails)
        print(f"  Found {len(emails)} emails from {email_addr}")

    if not all_today:
        print("\n✓ No emails sent today.")
        return

    print(f"✓ Total: {len(all_today)} emails\n")

    # Extract items ordered TODAY for each vendor
    today_items = {}
    vendors_ordered_today = []

    for vendor, vendor_email in VENDORS.items():
        if any(vendor_in_email(vendor_email, e["to"]) for e in all_today):
            vendors_ordered_today.append(vendor)
            items = extract_items(all_today, vendor_email)
            today_items[vendor] = items
            print(f"✓ {vendor}: {len(items)} items")
        else:
            print(f"✗ {vendor}: not ordered")
            today_items[vendor] = []

    print()

    # Get PREVIOUS orders for each vendor (all emails before today)
    prev_items = {}
    for email_addr, token in access_tokens:
        all_emails = fetch_emails(token, days=365)  # Get last year
        today_ts = int(datetime.combine(today, datetime.min.time()).timestamp() * 1000)

        for vendor, vendor_email in VENDORS.items():
            # Find most recent email before today
            prev_emails = [e for e in all_emails if e["ts"] < today_ts and vendor_in_email(vendor_email, e["to"])]
            if prev_emails and vendor not in prev_items:
                items = extract_items(prev_emails, vendor_email)
                prev_items[vendor] = items

    # Handle first email tracking and 30-minute timer
    state = load_state()
    today_key = today.isoformat()
    state_today = state.get(today_key, {})

    first_email_time = state_today.get("first_email_time")
    if all_today and not first_email_time:
        first_email_time = now_ms
        state_today["first_email_time"] = first_email_time
        state[today_key] = state_today
        save_state(state)

    already_sent = state_today.get("sent", False)
    force = env_is_true("FORCE_SEND")
    dry = env_is_true("DRY_RUN")
    scan_only = env_is_true("SCAN_ONLY")

    if scan_only:
        print("📊 SCAN MODE\n")
        if first_email_time:
            mins_ago = (now_ms - first_email_time) / 1000 / 60
            print(f"First email: {mins_ago:.1f} min ago")
            if mins_ago >= 30:
                print("✓ Ready to send")
            else:
                print(f"Will send in {30 - mins_ago:.1f} min")
        return

    if already_sent and not force:
        print("✓ Already sent today.")
        return

    send_now = force
    if not force and first_email_time:
        mins_since = (now_ms - first_email_time) / 1000 / 60
        if mins_since < 30:
            print(f"First email was {mins_since:.1f} min ago - will send in {30 - mins_since:.1f} min")
            return
        send_now = True

    if send_now and vendors_ordered_today:
        sender_email, sender_token = access_tokens[0]

        if dry:
            print("--- DRY RUN ---\n")
            print(build_email(today_items, prev_items, today))
        else:
            send_email(sender_token, sender_email, today_items, prev_items, today)
            state_today["sent"] = True
            state[today_key] = state_today
            save_state(state)
    elif send_now:
        print("✓ No vendors ordered today.")
        state_today["sent"] = True
        state[today_key] = state_today
        save_state(state)


if __name__ == "__main__":
    main()
