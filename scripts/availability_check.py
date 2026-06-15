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

NOTIFY_EMAILS = ["matt@farmlindproduce.com", "farmlindproduce@gmail.com", "babbe331@gmail.com"]
CHANNEL_NAME = "availability-list-changes"
REFERENCE_FILE = os.path.join(os.path.dirname(__file__), "product_reference.xlsx")

claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


# ── Product reference loader ───────────────────────────────────────────────────

def load_product_reference():
    """
    Load known product names from the internal availability spreadsheet.
    Returns a deduplicated list of product name strings Claude can use
    to map vendor abbreviations to real names.
    """
    if not os.path.exists(REFERENCE_FILE):
        print("  No product_reference.xlsx found — skipping reference lookup")
        return []

    names = set()
    try:
        wb = openpyxl.load_workbook(REFERENCE_FILE, read_only=True, data_only=True)

        # List sheet col A — full product names used internally
        if "List" in wb.sheetnames:
            ws = wb["List"]
            for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
                val = row[0]
                if val and isinstance(val, str):
                    val = val.strip().replace("\xa0", "")
                    # Skip section headers (all caps) and empty
                    if val and not val.isupper():
                        names.add(val)

            # Reserva Variants col 35 (AJ) — official system names
            for row in ws.iter_rows(min_col=35, max_col=35, values_only=True):
                val = row[0]
                if val and isinstance(val, str):
                    val = val.strip()
                    if val and val not in ("Reserva names",):
                        names.add(val)

        # Product IDs sheet — display_name column
        if "Product IDs" in wb.sheetnames:
            ws2 = wb["Product IDs"]
            for row in ws2.iter_rows(min_col=2, max_col=2, values_only=True):
                val = row[0]
                if val and isinstance(val, str) and val.strip() != "display_name":
                    names.add(val.strip())

    except Exception as e:
        print(f"  Warning: could not load product reference: {e}")

    print(f"  Loaded {len(names)} product names from reference file")
    return sorted(names)


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


def has_messages_today(messages):
    """Return True if any message was posted today (ET)."""
    from datetime import timezone, timedelta
    et = timezone(timedelta(hours=-4))  # EDT; use -5 for EST
    today = datetime.now(et).date()
    for msg in messages:
        ts = float(msg.get("ts", 0))
        msg_date = datetime.fromtimestamp(ts, tz=et).date()
        if msg_date == today:
            return True
    return False


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

def build_extract_prompt(product_names):
    reference_section = ""
    if product_names:
        sample = product_names[:300]
        reference_section = (
            "\n\nKNOWN PRODUCT NAMES (use these to resolve abbreviations and shorthand):\n"
            + "\n".join(f"- {n}" for n in sample)
            + "\n\nWhen you see an abbreviation or short name, match it to the closest known product name above."
        )

    return f"""You are parsing a produce availability/price list. Extract the vendor name and all priced items.

VENDOR NAME RULES — identify the vendor using these exact mappings:
- "TMK", "Tmk" → "TMK"
- "Aurpack", "Aurback", "Auerbach", "Aurebach", "Auerpak", "Aureback" → "Aurpack"
- "Nathel", "Nathel list", "Nathel prices" → "Nathel"
- "Top line", "Top Line" → "Top Line"
- "A and j", "A and J", "A & J", "A&J" → "A & J Produce Corp"
- "Dagele", "DAGELE" → "Dagele Brothers"
- "Triple J" → "Triple J"
- "Andy boy", "Andy Boy" → "Andy Boy"
- "Stews", "Stew", "Stew Leonard" → "Stew Leonards"
- If the content says "Sunday Specials" or "Sunday prices" → vendor is "Sunday Specials"
- If the content says "Tuesday Specials" → vendor is "Tuesday Specials"
- If the content says "Monday Specials" → vendor is "Monday Specials"
- If the content says "Wednesday Specials" → vendor is "Wednesday Specials"
- If no vendor can be identified → return null for vendor

Return ONLY a JSON object:
{{
  "vendor": "<vendor name or null>",
  "items": [
    {{
      "item": "<produce item name ONLY — no sizes, counts, or grades — in Title Case>",
      "price": <price as a number only, no $ sign>,
      "unit": "<size/count/weight/grade — e.g. '72ct', '36ct', '2.5 inch', '5 pack', '50lb', '12/3lb bags', 'case', 'lb', 'bunch' — include ALL distinguishing size or count info here>"
    }}
  ]
}}

Rules:
- Item name must be the CLEAN product name only — NO counts, sizes, or grades in the name. Put all of that in the unit field instead.
  - WRONG: "Fuji Apple 72ct", "Grapefruit 36ct", "Lime 230ct"
  - RIGHT: item="Fuji Apple", unit="72ct" / item="Grapefruit", unit="36ct" / item="Lime", unit="230ct"
- If the same item appears at multiple prices, put the distinguishing size/count/grade in the unit field so they can be told apart.
- Price must be a number only — no $ symbol, no slashes.
- If a price range is given (e.g. 1.00-1.50), use the lower number.
- Do NOT include items with no price.
- Return valid JSON only — no markdown, no explanation.{reference_section}"""


def extract_prices_from_content(content_blocks, sender_name, product_names):
    prompt = build_extract_prompt(product_names)
    result = claude.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": content_blocks + [{"type": "text", "text": prompt}]
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

    # Skip messages where Claude couldn't identify a real vendor name
    if not data.get("vendor") or data["vendor"] in ("Unknown Vendor", ""):
        data["vendor"] = None

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

    headers = ["Item", "Price Range", "Cheapest Option", "Vendor", "Price", "Unit/Weight/Count"]
    col_widths = [28, 16, 22, 28, 12, 18]

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


# ── Item name normalization ────────────────────────────────────────────────────

import re
import string

SPECIALS_VENDORS = {"Sunday Specials", "Tuesday Specials", "Monday Specials", "Wednesday Specials"}

ITEM_ALIASES = {
    "strawberries": "Strawberry",
    "peaches": "Peach",
    "avocados": "Avocado",
    "mangoes": "Mango",
    "pineapples": "Pineapple",
    "lemons": "Lemon",
    "limes": "Lime",
    "oranges": "Orange",
    "grapes": "Grape",
    "cherries": "Cherry",
    "blueberries": "Blueberry",
    "raspberries": "Raspberry",
    "blackberries": "Blackberry",
    "plums": "Plum",
    "nectarines": "Nectarine",
    "apricots": "Apricot",
    "apples, fuji": "Fuji Apple",
    "fuji": "Fuji Apple",
    "apples fuji": "Fuji Apple",
    "granny smith": "Granny Smith Apple",
    "apples granny smith": "Granny Smith Apple",
    "granny": "Granny Smith Apple",
    "pears, bartlett": "Bartlett Pear",
    "onions, sweet globe": "Sweet Globe Onion",
    "oranges, navel": "Navel Orange",
    "hass avocados": "Hass Avocado",
}

def normalize_item_name(name):
    """Normalize item names so minor variations map to the same key."""
    name = re.sub(r"[,\.\/\\]", " ", name)
    name = " ".join(name.split())
    lower = name.lower()
    if lower in ITEM_ALIASES:
        return ITEM_ALIASES[lower]
    return name.title()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("Fetching availability price data from Slack...")

    print("Loading product reference...")
    product_names = load_product_reference()

    channel_id = find_channel_id(CHANNEL_NAME)
    print(f"Found channel #{CHANNEL_NAME}: {channel_id}")

    messages = fetch_channel_history(channel_id, limit=200)
    print(f"Fetched {len(messages)} messages")

    if not os.environ.get("FORCE_SEND") and not has_messages_today(messages):
        print("No new messages posted today — skipping email.")
        return

    # item -> list of (vendor, price, unit)
    # Keys are normalized item names; values track (vendor, price, unit)
    item_vendor_map = defaultdict(list)
    # canonical name lookup: normalized_key -> display name
    item_display_name = {}
    # track which vendors we've already processed (specials only use most recent)
    seen_vendors = set()

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
            result = extract_prices_from_content(content_blocks, sender_name, product_names)
            if not result:
                continue
            vendor = result.get("vendor")
            if not vendor:
                print(f"    Skipping — no vendor name found in content")
                continue
            # Only use the most recent message per vendor (Slack returns newest first)
            # This means if a vendor hasn't posted recently, their last price is still used
            if vendor in seen_vendors:
                print(f"    Skipping older {vendor} message")
                continue
            seen_vendors.add(vendor)
            items = result.get("items", [])
            print(f"    Vendor: {vendor} — {len(items)} items")
            for entry in items:
                item = entry.get("item", "").strip()
                price = entry.get("price")
                unit = entry.get("unit", "") or ""
                if item and price is not None:
                    try:
                        key = normalize_item_name(item)
                        if key not in item_display_name:
                            item_display_name[key] = item
                        entry = (vendor, float(price), unit.strip())
                        if entry not in item_vendor_map[key]:
                            item_vendor_map[key].append(entry)
                    except (ValueError, TypeError):
                        pass
        except Exception as e:
            print(f"  Parse error: {e}")

    if not item_vendor_map:
        print("No price data found in channel — nothing to send.")
        return

    # Rebuild map using display names for output
    display_map = defaultdict(list)
    for key, entries in item_vendor_map.items():
        display_name = item_display_name.get(key, key)
        display_map[display_name] = entries

    print(f"\nTotal unique items: {len(display_map)}")

    today_str = datetime.now().strftime("%B %d, %Y")
    subject = f"Availability Price Comparison — {today_str}"

    print("Building spreadsheet...")
    xlsx_bytes = build_spreadsheet(display_map)

    body = (
        f"Hi Matt,\n\n"
        f"Attached is today's availability price comparison across all vendors.\n\n"
        f"GREEN = cheapest option   |   RED = most expensive   |   YELLOW = middle\n\n"
        f"— Farmlind Availability Bot"
    )

    print("Sending email...")
    token = get_access_token()
    send_email_with_attachment(token, subject, body, xlsx_bytes)
    print("Done — email sent with spreadsheet attached.")


if __name__ == "__main__":
    main()
