import json
import base64
import urllib.request
import urllib.parse
import os
from datetime import datetime, timedelta
from collections import defaultdict

import anthropic
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
GMAIL_CLIENT_ID = os.environ["GMAIL_CLIENT_ID"]
GMAIL_CLIENT_SECRET = os.environ["GMAIL_CLIENT_SECRET"]
GMAIL_REFRESH_TOKEN = os.environ["GMAIL_REFRESH_TOKEN_1"]

ALL_EMAILS = ["matt@farmlindproduce.com", "babbe331@gmail.com"]
TEST_EMAILS = ["babbe331@gmail.com"]
NOTIFY_EMAILS = TEST_EMAILS if os.environ.get("TEST_MODE") == "true" else ALL_EMAILS

RESTAURANT_CHANNEL = "C03E4TEBZU0"
STORE_CHANNEL = "C03D0GJ5F0F"
CHANNELS = [RESTAURANT_CHANNEL, STORE_CHANNEL]

HISTORY_FILE = os.path.join(os.path.dirname(__file__), "order_tracking_history.json")

slack_client = WebClient(token=SLACK_BOT_TOKEN)
claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY, timeout=120.0, max_retries=2)


# ── Vendor order history ───────────────────────────────────────────────────────

def load_vendor_history():
    """Load vendor last order dates from JSON file."""
    if not os.path.exists(HISTORY_FILE):
        return {}
    try:
        with open(HISTORY_FILE, "r") as f:
            return json.load(f)
    except Exception as e:
        print(f"  Warning: could not load history file: {e}")
        return {}


def save_vendor_history(data):
    """Save vendor last order dates to JSON file."""
    try:
        with open(HISTORY_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"  Warning: could not save history file: {e}")


# ── Gmail helpers ───────────────────────────────────────────────────────────────

def get_access_token(refresh_token=None):
    data = urllib.parse.urlencode({
        "client_id": GMAIL_CLIENT_ID,
        "client_secret": GMAIL_CLIENT_SECRET,
        "refresh_token": refresh_token or GMAIL_REFRESH_TOKEN,
        "grant_type": "refresh_token"
    }).encode()
    req = urllib.request.Request("https://oauth2.googleapis.com/token", data=data)
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())["access_token"]


def send_email(subject, body_text):
    """Send email via Gmail API."""
    try:
        access_token = get_access_token()
        recipients = ", ".join(NOTIFY_EMAILS)

        mime = (
            f"From: matt@farmlindproduce.com\r\n"
            f"To: {recipients}\r\n"
            f"Subject: {subject}\r\n"
            f"MIME-Version: 1.0\r\n"
            f"Content-Type: text/plain; charset=utf-8\r\n\r\n"
            f"{body_text}"
        )

        msg_b64 = base64.urlsafe_b64encode(mime.encode()).decode()
        body = {"raw": msg_b64}

        req = urllib.request.Request(
            "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
            data=json.dumps(body).encode(),
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req) as resp:
            result = json.loads(resp.read())
            print(f"  Email sent: {result.get('id')}")
    except Exception as e:
        print(f"  Error sending email: {e}")


# ── Slack message processing ───────────────────────────────────────────────────

def extract_vendor_name_from_message(message_text):
    """
    Use Claude to extract vendor name from message.
    Handles both "Vendor Name" at top and "add to Vendor Name" style messages.
    """
    if not message_text or len(message_text.strip()) < 2:
        return None

    response = claude.messages.create(
        model="claude-opus-4-8",
        max_tokens=50,
        messages=[{
            "role": "user",
            "content": f'Extract the vendor name from this message. It could be at the top or mentioned as "add to [vendor]". Return ONLY the vendor name, nothing else:\n\n"{message_text}"'
        }]
    )

    vendor = response.content[0].text.strip()
    return vendor if vendor and len(vendor) > 1 else None


def extract_vendor_from_file(file_url):
    """Extract vendor name from PDF, image, or document using Claude vision."""
    try:
        response = claude.messages.create(
            model="claude-opus-4-8",
            max_tokens=100,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "url",
                            "url": file_url
                        }
                    },
                    {
                        "type": "text",
                        "text": "Extract only the vendor name from the first line of this document or image. Return ONLY the vendor name, nothing else."
                    }
                ]
            }]
        )

        vendor = response.content[0].text.strip()
        return vendor if vendor and len(vendor) > 1 else None
    except Exception as e:
        print(f"  Error processing file {file_url}: {e}")
        return None


def get_message_reactions(channel_id, message_ts):
    """Check if message has any emoji reactions."""
    try:
        response = slack_client.reactions_get(channel=channel_id, timestamp=message_ts)
        return len(response.get("message", {}).get("reactions", [])) > 0
    except SlackApiError as e:
        if e.response["error"] != "message_not_found":
            print(f"  Slack error checking reactions: {e}")
        return False


def get_slack_message_link(channel_id, message_ts):
    """Generate Slack message permalink from channel ID and timestamp."""
    # Convert timestamp "1234567890.123456" to "p1234567890123456"
    ts_clean = message_ts.replace(".", "")
    return f"https://farmlindproduce.slack.com/archives/{channel_id}/p{ts_clean}"


def fetch_channel_messages(channel_id, hours=24):
    """Fetch recent messages from a channel."""
    try:
        oldest_ts = (datetime.now() - timedelta(hours=hours)).timestamp()
        response = slack_client.conversations_history(
            channel=channel_id,
            oldest=oldest_ts,
            limit=100
        )
        return response.get("messages", [])
    except SlackApiError as e:
        print(f"  Slack error fetching messages: {e}")
        return []


# ── Main order tracking logic ───────────────────────────────────────────────────

def check_orders():
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting order tracking check...")

    vendor_history = load_vendor_history()
    today = datetime.now().date()

    # Track orders from today and check for processed status
    orders_today = defaultdict(lambda: {"has_reaction": False, "messages": []})

    # Fetch messages from both channels
    for channel_id in CHANNELS:
        print(f"  Scanning channel {channel_id}...")
        messages = fetch_channel_messages(channel_id, hours=24)

        for msg in messages:
            if "text" in msg:
                # Text message
                vendor = extract_vendor_name_from_message(msg["text"])
            elif "files" in msg:
                # PDF, image, or document
                vendor = None
                for file in msg["files"]:
                    file_url = file.get("url_private")
                    if file_url:
                        vendor = extract_vendor_from_file(file_url)
                        if vendor:
                            break
            else:
                vendor = None

            if vendor:
                has_reaction = get_message_reactions(channel_id, msg["ts"])
                message_link = get_slack_message_link(channel_id, msg["ts"])
                orders_today[vendor]["has_reaction"] = has_reaction
                orders_today[vendor]["messages"].append({
                    "channel": channel_id,
                    "ts": msg["ts"],
                    "text": msg.get("text", "")[:100],
                    "link": message_link
                })

                # Update last order date for this vendor
                vendor_history[vendor] = today.isoformat()

    # Save updated history
    save_vendor_history(vendor_history)

    # Build email reports
    unprocessed_today = [v for v, data in orders_today.items() if not data["has_reaction"]]

    four_days_ago = (today - timedelta(days=4)).isoformat()
    no_recent_orders = [
        v for v, last_date in vendor_history.items()
        if last_date < four_days_ago
    ]

    # Send email
    subject = f"Order Tracking Report — {today.strftime('%m/%d/%Y')}"

    body = f"Order Tracking Report\n{today.strftime('%m/%d/%Y')}\n\n"

    body += "VENDORS WHO ORDERED BUT HAVEN'T BEEN MARKED YET (NO EMOJI):\n"
    if unprocessed_today:
        for v in sorted(unprocessed_today):
            body += f"  • {v}\n"
            for msg in orders_today[v]["messages"]:
                body += f"    {msg['link']}\n"
    else:
        body += "  None — all orders marked!\n"

    body += f"\nVENDORS WITH NO ORDERS IN 4+ DAYS:\n"
    if no_recent_orders:
        for v in sorted(no_recent_orders):
            last = vendor_history[v]
            body += f"  • {v} (last order: {last})\n"
    else:
        body += "  None — everyone ordering regularly!\n"

    body += f"\n\nTotal vendors in history: {len(vendor_history)}\n"
    body += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} EST\n"

    print(f"\n  Unprocessed orders today: {len(unprocessed_today)}")
    print(f"  No orders in 4+ days: {len(no_recent_orders)}")
    print(f"\nEmail body:\n{body}")

    send_email(subject, body)


if __name__ == "__main__":
    check_orders()
