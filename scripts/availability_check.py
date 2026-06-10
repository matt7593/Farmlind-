import json
import base64
import urllib.request
import urllib.parse
import io
import os
from datetime import datetime
from collections import defaultdict

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


def get_user_display_name(user_id):
    try:
        data = slack_get("/users.info", {"user": user_id})
        profile = data.get("user", {}).get("profile", {})
        return profile.get("display_name") or profile.get("real_name") or user_id
    except Exception:
        return user_id


def download_slack_file(url):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"})
    with urllib.request.urlopen(req) as resp:
        return resp.read()


# ── Claude helpers ─────────────────────────────────────────────────────────────

EXTRACT_PROMPT = """
You are parsing a produce availability/price list from a vendor.

Extract:
1. The vendor/company name (look for it in the document header, footer, or letterhead). If you cannot find a company name, return "Unknown Vendor".
2. Every item that has a price.

Return ONLY a JSON object in this exact format:
{
  "vendor": "<company name>",
  "items": [
    {
      "item": "<produce item name, normalized to Title Case>",
      "price": <price as a number, e.g. 1.25>,
      "unit": "<unit if present, e.g. 'lb', 'each', 'case', 'bunch', or blank>"
    }
  ]
}

Rules:
- Normalize item names to Title Case, remove extra spaces.
- Price must be a number only — no $ symbol, no slashes.
- If a price range is given (e.g. 1.00-1.50), use the lower number.
- Do NOT include items with no price.
- Return valid JSON only — no markdown, no explanation.
"""


def extract_prices_from_content(content_blocks, sender_name):
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
        data = json.loads(raw)
    except json.JSONDecodeError:
        start, end = raw.find("{"), raw.rfind("}")
        if start != -1 and end != -1:
            try:
                data = json.loads(raw[start:end + 1])
            except Exception:
                print(f"  Could not parse Claude response: {raw[:200]}")
                return None
        else:
            return None

    # If Claude couldn't find a vendor name, fall back to the Slack sender
    if not data.get("vendor") or data["vendor"] == "Unknown Vendor":
        data["vendor"] = sender_name

    return data


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


def send_email_with_attachment(access_token, subject, body_text, xlsx_bytes):
    attachment_b64 = base64.b64encode(xlsx_bytes).decode()
    filename = f"availability_{datetime.now().strftime('%Y%m%d')}.xlsx"
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


# ── Email body builder ─────────────────────────────────────────────────────────

def build_email_body(item_vendor_map, today_str):
    lines = [
        f"Availability Price Comparison — {today_str}",
        "=" * 50,
        "",
        "Items are sorted alphabetically.",
        "Vendors listed cheapest to most expensive.",
        "",
        "=" * 50,
        "",
    ]

    for item in sorted(item_vendor_map.keys()):
        vendors = item_vendor_map[item]  # list of (vendor, price, unit)
        vendors_sorted = sorted(vendors, key=lambda x: x[1])

        prices = [v[1] for v in vendors_sorted]
        unit = vendors_sorted[0][2] if vendors_sorted[0][2] else ""
        unit_str = f"/{unit}" if unit else ""

        low = min(prices)
        high = max(prices)
        cheapest_vendor = vendors_sorted[0][0]

        lines.append(f"  {item.upper()}")
        if low == high:
            lines.append(f"  Price: ${low:.2f}{unit_str} (all vendors same price)")
        else:
            lines.append(f"  Range: ${low:.2f} - ${high:.2f}{unit_str}  |  Cheapest: {cheapest_vendor}")
        lines.append("")

        for vendor, price, unit in vendors_sorted:
            u = f"/{unit}" if unit else ""
            lines.append(f"    ${price:.2f}{u}  —  {vendor}")

        lines.append("")
        lines.append("-" * 50)
        lines.append("")

    lines.append("— Farmlind Availability Bot")
    return "\n".join(lines)


# ── Spreadsheet builder ────────────────────────────────────────────────────────

GREEN_FILL  = PatternFill("solid", fgColor="C6EFCE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
RED_FILL    = PatternFill("solid", fgColor="FFC7CE")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)


def build_spreadsheet(item_vendor_map):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Price Comparison"

    headers = ["Item", "Price Range", "Cheapest Option", "Vendor", "Price", "Unit"]
    col_widths = [28, 16, 22, 28, 12, 10]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN_BORDER
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 18
    ws.freeze_panes = "A2"

    row_idx = 2
    for item in sorted(item_vendor_map.keys()):
        vendors = sorted(item_vendor_map[item], key=lambda x: x[1])
        prices = [v[1] for v in vendors]
        low, high = min(prices), max(prices)
        cheapest_vendor = vendors[0][0]
        unit = vendors[0][2] if vendors[0][2] else ""
        price_range = f"${low:.2f} - ${high:.2f}" if low != high else f"${low:.2f}"

        item_start_row = row_idx

        for i, (vendor, price, vunit) in enumerate(vendors):
            u = vunit or unit

            # Item and range only on first row of the group
            if i == 0:
                ws.cell(row=row_idx, column=1, value=item).font = Font(bold=True)
                ws.cell(row=row_idx, column=2, value=price_range)
                ws.cell(row=row_idx, column=3, value=cheapest_vendor)

            ws.cell(row=row_idx, column=4, value=vendor)
            ws.cell(row=row_idx, column=5, value=f"${price:.2f}")
            ws.cell(row=row_idx, column=6, value=u)

            # Color: green = cheapest, red = most expensive, yellow = middle
            if price == low:
                fill = GREEN_FILL
            elif price == high and low != high:
                fill = RED_FILL
            else:
                fill = YELLOW_FILL

            for col in range(1, 7):
                cell = ws.cell(row=row_idx, column=col)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="center" if col > 1 else "left")
                if col >= 4:
                    cell.fill = fill

            row_idx += 1

        # Merge item/range/cheapest cells across the group rows
        if len(vendors) > 1:
            for col in (1, 2, 3):
                ws.merge_cells(
                    start_row=item_start_row, start_column=col,
                    end_row=row_idx - 1, end_column=col
                )
                ws.cell(row=item_start_row, column=col).alignment = Alignment(
                    horizontal="left" if col == 1 else "center",
                    vertical="center"
                )

        # Empty spacer row between items
        row_idx += 1

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

    # item -> list of (vendor, price, unit)
    item_vendor_map = defaultdict(list)

    for msg in messages:
        ts = msg.get("ts", "0")
        user_id = msg.get("user", "")
        sender_name = get_user_display_name(user_id) if user_id else "Unknown"

        content_blocks = []

        text = msg.get("text", "").strip()
        if text:
            content_blocks.append({"type": "text", "text": text})

        for f in msg.get("files", []):
            mime = f.get("mimetype", "")
            url = f.get("url_private_download") or f.get("url_private")
            filename = f.get("name", "")
            if not url:
                continue
            print(f"  Downloading: {filename} ({mime})")
            try:
                file_bytes = download_slack_file(url)
                if "pdf" in mime or filename.lower().endswith(".pdf"):
                    content_blocks.append({
                        "type": "document",
                        "source": {"type": "base64", "media_type": "application/pdf",
                                   "data": base64.b64encode(file_bytes).decode()}
                    })
                elif mime.startswith("image/"):
                    content_blocks.append({
                        "type": "image",
                        "source": {"type": "base64", "media_type": mime,
                                   "data": base64.b64encode(file_bytes).decode()}
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
                print(f"  File error: {e}")

        if not content_blocks:
            continue

        msg_date = datetime.fromtimestamp(float(ts)).strftime("%Y-%m-%d")
        print(f"  Parsing message from {sender_name} ({msg_date})...")
        try:
            result = extract_prices_from_content(content_blocks, sender_name)
            if not result:
                continue
            vendor = result.get("vendor", sender_name)
            items = result.get("items", [])
            print(f"    Vendor: {vendor} — {len(items)} items")
            for entry in items:
                item = entry.get("item", "").strip()
                price = entry.get("price")
                unit = entry.get("unit", "") or ""
                if item and price is not None:
                    try:
                        item_vendor_map[item].append((vendor, float(price), unit))
                    except (ValueError, TypeError):
                        pass
        except Exception as e:
            print(f"  Parse error: {e}")

    if not item_vendor_map:
        print("No price data found in channel — nothing to send.")
        return

    print(f"\nTotal unique items: {len(item_vendor_map)}")

    today_str = datetime.now().strftime("%B %d, %Y")
    subject = f"Availability Price Comparison — {today_str}"

    print("Building email and spreadsheet...")
    body = build_email_body(item_vendor_map, today_str)
    xlsx_bytes = build_spreadsheet(item_vendor_map)

    print("Sending email...")
    token = get_access_token()
    send_email_with_attachment(token, subject, body, xlsx_bytes)
    print("Done — email sent with spreadsheet attached.")


if __name__ == "__main__":
    main()
