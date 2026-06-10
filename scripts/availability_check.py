import json
import base64
import urllib.request
import urllib.parse
import io
import os
from datetime import datetime

import anthropic
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
GMAIL_CLIENT_ID = os.environ["GMAIL_CLIENT_ID"]
GMAIL_CLIENT_SECRET = os.environ["GMAIL_CLIENT_SECRET"]
GMAIL_REFRESH_TOKEN = os.environ["GMAIL_REFRESH_TOKEN_1"]

NOTIFY_EMAILS = ["matt@farmlindproduce.com", "farmlindproduce@gmail.com"]
CHANNEL_NAME = "availability-list-changes"

claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


# ── Slack helpers ──────────────────────────────────────────────────────────────

def slack_get(path, params=None):
    url = f"https://slack.com/api{path}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"})
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())
    if not data.get("ok"):
        raise RuntimeError(f"Slack API error on {path}: {data.get('error')}")
    return data


def find_channel_id(name):
    cursor = None
    while True:
        params = {"limit": 200, "exclude_archived": "true", "types": "public_channel,private_channel"}
        if cursor:
            params["cursor"] = cursor
        data = slack_get("/conversations.list", params)
        for ch in data.get("channels", []):
            if ch["name"] == name:
                return ch["id"]
        cursor = data.get("response_metadata", {}).get("next_cursor")
        if not cursor:
            break
    raise RuntimeError(f"Channel #{name} not found — make sure the bot is invited to it")


def fetch_channel_history(channel_id, limit=200):
    data = slack_get("/conversations.history", {"channel": channel_id, "limit": limit})
    return data.get("messages", [])


def download_slack_file(url):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"})
    with urllib.request.urlopen(req) as resp:
        return resp.read()


# ── Claude helpers ─────────────────────────────────────────────────────────────

EXTRACT_PROMPT = """
You are parsing a produce availability/price list. Extract every item that has a price.

Return ONLY a JSON array. Each element:
{
  "item": "<produce item name, normalized>",
  "price": "<price as a string, e.g. '1.25' or '1.25/lb'>",
  "unit": "<unit if present, e.g. 'lb', 'each', 'case', 'bunch', blank if unknown>",
  "date": "<date if mentioned in the document, YYYY-MM-DD format, or blank>"
}

Rules:
- Normalize item names (title case, remove extra spaces).
- Keep the price exactly as shown, no currency symbol.
- If no date appears in the document, leave date blank.
- Do NOT include items with no price.
- Return valid JSON only — no markdown, no explanation.
"""


def extract_prices_from_content(content_blocks, message_ts):
    msg_date = datetime.fromtimestamp(float(message_ts)).strftime("%Y-%m-%d")
    result = claude.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": content_blocks + [{"type": "text", "text": EXTRACT_PROMPT}]
        }]
    )
    raw = result.content[0].text.strip()
    try:
        items = json.loads(raw)
    except json.JSONDecodeError:
        # Try to salvage a partial JSON array
        start, end = raw.find("["), raw.rfind("]")
        if start != -1 and end != -1:
            try:
                items = json.loads(raw[start:end + 1])
            except Exception:
                print(f"  Could not parse Claude response: {raw[:200]}")
                return []
        else:
            return []

    # Fill in message date where the document had none
    for item in items:
        if not item.get("date"):
            item["date"] = msg_date
    return items


def parse_price_float(price_str):
    """Return the numeric part of a price string as float, or None."""
    if not price_str:
        return None
    cleaned = price_str.replace("$", "").split("/")[0].strip()
    try:
        return float(cleaned)
    except ValueError:
        return None


# ── Gmail helpers ──────────────────────────────────────────────────────────────

def get_access_token():
    data = urllib.parse.urlencode({
        "client_id": GMAIL_CLIENT_ID,
        "client_secret": GMAIL_CLIENT_SECRET,
        "refresh_token": GMAIL_REFRESH_TOKEN,
        "grant_type": "refresh_token"
    }).encode()
    req = urllib.request.Request("https://oauth2.googleapis.com/token", data=data)
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())["access_token"]


def send_spreadsheet_email(access_token, xlsx_bytes, subject, body_text):
    attachment_b64 = base64.b64encode(xlsx_bytes).decode()
    filename = f"availability_comparison_{datetime.now().strftime('%Y%m%d')}.xlsx"
    boundary = "==farmlind_boundary=="

    recipients = ", ".join(NOTIFY_EMAILS)
    mime = (
        f"From: matt@farmlindproduce.com\r\n"
        f"To: {recipients}\r\n"
        f"Subject: {subject}\r\n"
        f"MIME-Version: 1.0\r\n"
        f"Content-Type: multipart/mixed; boundary=\"{boundary}\"\r\n\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: text/plain; charset=utf-8\r\n\r\n"
        f"{body_text}\r\n\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n"
        f"Content-Transfer-Encoding: base64\r\n"
        f"Content-Disposition: attachment; filename=\"{filename}\"\r\n\r\n"
        f"{attachment_b64}\r\n\r\n"
        f"--{boundary}--"
    )
    encoded = base64.urlsafe_b64encode(mime.encode()).decode()
    url = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send"
    req_data = json.dumps({"raw": encoded}).encode()
    req = urllib.request.Request(
        url, data=req_data,
        headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


# ── Spreadsheet builder ────────────────────────────────────────────────────────

GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL   = PatternFill("solid", fgColor="FFC7CE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BOLD = Font(bold=True)

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)


def build_spreadsheet(price_data):
    """
    price_data: list of {item, price, unit, date, vendor_hint}
    Groups by item name, finds the two most recent distinct dates, computes change.
    Returns xlsx bytes.
    """
    # Sort by date descending
    price_data.sort(key=lambda x: x.get("date", ""), reverse=True)

    # Build: item -> list of (date, price_str, price_float)
    item_history: dict[str, list] = {}
    for row in price_data:
        item = row["item"]
        date = row.get("date", "")
        price_str = row.get("price", "")
        price_f = parse_price_float(price_str)
        unit = row.get("unit", "")
        if item not in item_history:
            item_history[item] = []
        item_history[item].append((date, price_str, price_f, unit))

    # For each item, pick the two most recent distinct dates
    comparison_rows = []
    for item, entries in sorted(item_history.items()):
        # Sort by date desc
        entries.sort(key=lambda x: x[0], reverse=True)
        # Get entries for the most recent date
        dates_seen = []
        by_date = {}
        for date, price_str, price_f, unit in entries:
            if date not in by_date:
                by_date[date] = (price_str, price_f, unit)
                dates_seen.append(date)

        if not dates_seen:
            continue

        latest_date = dates_seen[0]
        latest_price_str, latest_price_f, latest_unit = by_date[latest_date]

        if len(dates_seen) >= 2:
            prev_date = dates_seen[1]
            prev_price_str, prev_price_f, prev_unit = by_date[prev_date]
        else:
            prev_date = ""
            prev_price_str = ""
            prev_price_f = None
            prev_unit = latest_unit

        # Compute change
        if latest_price_f is not None and prev_price_f is not None:
            change = latest_price_f - prev_price_f
            change_pct = (change / prev_price_f * 100) if prev_price_f else 0
        else:
            change = None
            change_pct = None

        unit_display = latest_unit or prev_unit

        comparison_rows.append({
            "item": item,
            "unit": unit_display,
            "prev_date": prev_date,
            "prev_price": prev_price_str,
            "latest_date": latest_date,
            "latest_price": latest_price_str,
            "change": change,
            "change_pct": change_pct,
        })

    # Sort: biggest price increases first, then alphabetical
    comparison_rows.sort(key=lambda x: (-(x["change"] or 0), x["item"]))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Price Comparison"

    headers = ["Item", "Unit", "Previous Date", "Previous Price", "Latest Date", "Latest Price", "$ Change", "% Change", "Status"]
    col_widths = [30, 10, 14, 14, 14, 14, 10, 10, 12]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN_BORDER
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 18

    for row_idx, row in enumerate(comparison_rows, 2):
        change = row["change"]
        change_pct = row["change_pct"]

        if change is not None:
            if change > 0:
                fill = RED_FILL
                status = "UP"
            elif change < 0:
                fill = GREEN_FILL
                status = "DOWN"
            else:
                fill = None
                status = "SAME"
        else:
            fill = YELLOW_FILL
            status = "NEW" if not row["prev_price"] else "—"

        values = [
            row["item"],
            row["unit"],
            row["prev_date"],
            row["prev_price"],
            row["latest_date"],
            row["latest_price"],
            f"+{change:.2f}" if change and change > 0 else (f"{change:.2f}" if change is not None else ""),
            f"+{change_pct:.1f}%" if change_pct and change_pct > 0 else (f"{change_pct:.1f}%" if change_pct is not None else ""),
            status,
        ]

        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if col > 1 else "left")
            if fill:
                cell.fill = fill
            if col in (7, 8) and change:
                cell.font = Font(bold=True, color="9C0006" if change > 0 else "375623")

    # Freeze header row
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("Fetching availability price data from Slack...")

    channel_id = find_channel_id(CHANNEL_NAME)
    print(f"Found channel #{CHANNEL_NAME}: {channel_id}")

    messages = fetch_channel_history(channel_id, limit=200)
    print(f"Fetched {len(messages)} messages")

    all_price_data = []

    for msg in messages:
        ts = msg.get("ts", "0")
        msg_date = datetime.fromtimestamp(float(ts)).strftime("%Y-%m-%d")
        content_blocks = []

        # Plain text body
        text = msg.get("text", "").strip()
        if text:
            content_blocks.append({"type": "text", "text": text})

        # File attachments
        for f in msg.get("files", []):
            mime = f.get("mimetype", "")
            url = f.get("url_private_download") or f.get("url_private")
            filename = f.get("name", "")
            if not url:
                continue
            print(f"  Downloading file: {filename} ({mime})")
            try:
                file_bytes = download_slack_file(url)
                if "pdf" in mime or filename.lower().endswith(".pdf"):
                    content_blocks.append({
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": base64.b64encode(file_bytes).decode()
                        }
                    })
                elif mime.startswith("image/"):
                    content_blocks.append({
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": mime,
                            "data": base64.b64encode(file_bytes).decode()
                        }
                    })
                elif filename.lower().endswith((".xlsx", ".xls")):
                    import openpyxl as _oxl
                    wb = _oxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
                    text_content = "\n".join(
                        " ".join(str(c) for c in row if c is not None)
                        for sheet in wb.worksheets
                        for row in sheet.iter_rows(values_only=True)
                    )
                    content_blocks.append({"type": "text", "text": text_content})
            except Exception as e:
                print(f"  File download/parse error: {e}")

        if not content_blocks:
            continue

        print(f"  Parsing message from {msg_date}...")
        try:
            items = extract_prices_from_content(content_blocks, ts)
            print(f"    Extracted {len(items)} items")
            all_price_data.extend(items)
        except Exception as e:
            print(f"  Parse error: {e}")

    if not all_price_data:
        print("No price data found in channel — nothing to send.")
        return

    print(f"\nTotal items extracted: {len(all_price_data)}")
    print("Building spreadsheet...")
    xlsx_bytes = build_spreadsheet(all_price_data)

    today_str = datetime.now().strftime("%B %d, %Y")
    subject = f"Availability Price Comparison — {today_str}"
    body = (
        f"Hi Matt,\n\n"
        f"Here is the latest availability price comparison from #{CHANNEL_NAME}.\n\n"
        f"The spreadsheet shows each item's most recent price vs. the previous price:\n"
        f"  - RED   = price went UP\n"
        f"  - GREEN = price went DOWN\n"
        f"  - YELLOW = new item (no previous price to compare)\n\n"
        f"Items are sorted by largest price increase first.\n\n"
        f"— Farmlind Availability Bot"
    )

    print("Sending email...")
    token = get_access_token()
    send_spreadsheet_email(token, xlsx_bytes, subject, body)
    print("Done — spreadsheet sent.")


if __name__ == "__main__":
    main()
