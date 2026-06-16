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

ALL_EMAILS  = ["matt@farmlindproduce.com", "farmlindproduce@gmail.com", "babbe331@gmail.com", "howard@farmlindproduce.com"]
TEST_EMAILS = ["babbe331@gmail.com"]
NOTIFY_EMAILS = TEST_EMAILS if os.environ.get("TEST_MODE") == "true" else ALL_EMAILS
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
- "Nathel", "Nathel & Nathel", "Nathel and Nathel", "Nathel list", "Nathel prices", "PD371" → "Nathel"
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
- ALWAYS write the variety/descriptor BEFORE the base fruit/vegetable name (e.g., "Fuji Apple" NOT "Apple Fuji", "Navel Orange" NOT "Orange Navel", "Hass Avocado" NOT "Avocado Hass", "Roma Tomato" NOT "Tomato Roma", "Bartlett Pear" NOT "Pear Bartlett").
- Use SINGULAR form for all item names (e.g., "Strawberry" not "Strawberries", "Avocado" not "Avocados", "Lime" not "Limes").
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
    """Matrix layout: vendors across the top (columns), items down the side (rows).
    Each price cell holds a real number (currency-formatted) so spreadsheet formulas
    work directly. Size/count/unit info is attached as a cell note."""
    from openpyxl.comments import Comment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Price Comparison"

    # ── Collect the full vendor list (columns) ──
    all_vendors = set()
    for entries in item_vendor_map.values():
        for vendor, _price, _unit in entries:
            all_vendors.add(vendor)
    # Sort alphabetically, but push the "Specials" vendors to the end
    def vendor_sort_key(v):
        return (1 if v in SPECIALS_VENDORS else 0, v.lower())
    vendors = sorted(all_vendors, key=vendor_sort_key)

    # ── Header row ──
    # Col 1 = Item, Col 2 = Cheapest, Col 3 = Price Range, then one column per vendor
    ws.cell(row=1, column=1, value="Item")
    ws.cell(row=1, column=2, value="Cheapest")
    ws.cell(row=1, column=3, value="Price Range")
    vendor_col = {}
    for i, vendor in enumerate(vendors):
        col = 4 + i
        vendor_col[vendor] = col
        ws.cell(row=1, column=col, value=vendor)

    for col in range(1, 4 + len(vendors)):
        cell = ws.cell(row=1, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 16
    for vendor in vendors:
        ws.column_dimensions[openpyxl.utils.get_column_letter(vendor_col[vendor])].width = 14

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "D2"  # lock item name + cheapest + range columns and the header row

    # ── Item rows ──
    row_idx = 2
    for item in sorted(item_vendor_map.keys()):
        entries = item_vendor_map[item]

        # Collapse to one (price, unit) per vendor — keep the cheapest if a vendor
        # listed the same item at multiple sizes; remember all sizes for the note.
        per_vendor = {}  # vendor -> {"price": float, "unit": str, "all": [(price, unit)]}
        for vendor, price, unit in entries:
            slot = per_vendor.setdefault(vendor, {"price": price, "unit": unit, "all": []})
            slot["all"].append((price, unit))
            if price < slot["price"]:
                slot["price"] = price
                slot["unit"] = unit

        prices = [v["price"] for v in per_vendor.values()]
        low, high = min(prices), max(prices)
        cheapest_vendor = min(per_vendor.items(), key=lambda kv: kv[1]["price"])[0]
        price_range = f"${low:.2f} - ${high:.2f}" if low != high else f"${low:.2f}"

        # Item name + cheapest + range
        item_cell = ws.cell(row=row_idx, column=1, value=item)
        item_cell.font = Font(bold=True)
        item_cell.alignment = Alignment(horizontal="left", vertical="center")
        item_cell.border = THIN_BORDER
        for col, val in ((2, cheapest_vendor), (3, price_range)):
            c = ws.cell(row=row_idx, column=col, value=val)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = THIN_BORDER

        # Price per vendor
        for vendor, col in vendor_col.items():
            cell = ws.cell(row=row_idx, column=col)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            slot = per_vendor.get(vendor)
            if not slot:
                continue  # vendor doesn't carry this item — leave blank

            price = slot["price"]
            cell.value = price
            cell.number_format = '"$"#,##0.00'

            # Color: green = cheapest, red = most expensive, yellow = middle
            if price == low:
                cell.fill = GREEN_FILL
            elif price == high and low != high:
                cell.fill = RED_FILL
            else:
                cell.fill = YELLOW_FILL

            # Attach size/unit info as a cell note
            note_lines = []
            for p, u in sorted(slot["all"]):
                if u:
                    note_lines.append(f"${p:.2f} — {u}")
                else:
                    note_lines.append(f"${p:.2f}")
            note_text = "\n".join(note_lines)
            if note_text:
                cell.comment = Comment(note_text, "Farmlind Bot")

        row_idx += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Item name normalization ────────────────────────────────────────────────────

import re
import string

SPECIALS_VENDORS = {"Sunday Specials", "Tuesday Specials", "Monday Specials", "Wednesday Specials"}

ITEM_ALIASES = {
    # Plurals → singular
    "strawberries": "Strawberry",
    "peaches": "Peach",
    "avocados": "Avocado",
    "mangoes": "Mango",
    "mangos": "Mango",
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
    "tomatoes": "Tomato",
    "peppers": "Pepper",
    "onions": "Onion",
    "potatoes": "Potato",
    "cucumbers": "Cucumber",
    "zucchinis": "Zucchini",
    "watermelons": "Watermelon",
    "cantaloupes": "Cantaloupe",
    "honeydews": "Honeydew",
    "pears": "Pear",
    "apples": "Apple",
    "bananas": "Banana",
    "grapefruits": "Grapefruit",
    "tangerines": "Tangerine",
    "clementines": "Clementine",
    "pomegranates": "Pomegranate",
    "figs": "Fig",
    "dates": "Date",
    "kiwis": "Kiwi",
    "papayas": "Papaya",
    "guavas": "Guava",
    "lychees": "Lychee",
    "persimmons": "Persimmon",
    "quinces": "Quince",
    "kumquats": "Kumquat",
    "artichokes": "Artichoke",
    "asparagus tips": "Asparagus",
    "broccolis": "Broccoli",
    "cabbages": "Cabbage",
    "carrots": "Carrot",
    "cauliflowers": "Cauliflower",
    "celeries": "Celery",
    "corns": "Corn",
    "eggplants": "Eggplant",
    "garlic bulbs": "Garlic",
    "lettuces": "Lettuce",
    "mushrooms": "Mushroom",
    "parsnips": "Parsnip",
    "radishes": "Radish",
    "spinaches": "Spinach",
    "squashes": "Squash",
    "turnips": "Turnip",
    "beets": "Beet",
    "leeks": "Leek",
    "shallots": "Shallot",
    "scallions": "Scallion",
    "chives": "Chive",

    # Fuji Apple variants
    "fuji": "Fuji Apple",
    "fuji apple": "Fuji Apple",
    "fuji apples": "Fuji Apple",
    "apple fuji": "Fuji Apple",
    "apples fuji": "Fuji Apple",
    "apples, fuji": "Fuji Apple",
    "apple, fuji": "Fuji Apple",

    # Granny Smith variants
    "granny": "Granny Smith Apple",
    "granny smith": "Granny Smith Apple",
    "granny smith apple": "Granny Smith Apple",
    "granny smith apples": "Granny Smith Apple",
    "apple granny smith": "Granny Smith Apple",
    "apples granny smith": "Granny Smith Apple",
    "apples, granny smith": "Granny Smith Apple",

    # Gala Apple variants
    "gala": "Gala Apple",
    "gala apple": "Gala Apple",
    "gala apples": "Gala Apple",
    "apple gala": "Gala Apple",
    "apples gala": "Gala Apple",
    "apples, gala": "Gala Apple",

    # Honeycrisp Apple variants
    "honeycrisp": "Honeycrisp Apple",
    "honeycrisp apple": "Honeycrisp Apple",
    "honeycrisp apples": "Honeycrisp Apple",
    "apple honeycrisp": "Honeycrisp Apple",
    "apples honeycrisp": "Honeycrisp Apple",

    # Hass Avocado variants
    "hass": "Hass Avocado",
    "hass avocado": "Hass Avocado",
    "hass avocados": "Hass Avocado",
    "avocado hass": "Hass Avocado",
    "avocados hass": "Hass Avocado",

    # Navel Orange variants
    "navel": "Navel Orange",
    "navel orange": "Navel Orange",
    "navel oranges": "Navel Orange",
    "orange navel": "Navel Orange",
    "oranges navel": "Navel Orange",
    "oranges, navel": "Navel Orange",

    # Bartlett Pear variants
    "bartlett": "Bartlett Pear",
    "bartlett pear": "Bartlett Pear",
    "bartlett pears": "Bartlett Pear",
    "pear bartlett": "Bartlett Pear",
    "pears bartlett": "Bartlett Pear",
    "pears, bartlett": "Bartlett Pear",

    # Roma Tomato variants
    "roma": "Roma Tomato",
    "roma tomato": "Roma Tomato",
    "roma tomatoes": "Roma Tomato",
    "tomato roma": "Roma Tomato",
    "tomatoes roma": "Roma Tomato",

    # Bell Pepper variants
    "bell pepper": "Bell Pepper",
    "bell peppers": "Bell Pepper",
    "pepper bell": "Bell Pepper",
    "peppers bell": "Bell Pepper",
    "green pepper": "Green Bell Pepper",
    "green peppers": "Green Bell Pepper",
    "red pepper": "Red Bell Pepper",
    "red peppers": "Red Bell Pepper",
    "yellow pepper": "Yellow Bell Pepper",
    "yellow peppers": "Yellow Bell Pepper",

    # Sweet Globe Onion
    "sweet globe onion": "Sweet Globe Onion",
    "onions, sweet globe": "Sweet Globe Onion",
    "sweet onion": "Sweet Onion",
    "sweet onions": "Sweet Onion",

    # Yellow Onion variants
    "yellow onion": "Yellow Onion",
    "yellow onions": "Yellow Onion",
    "onion yellow": "Yellow Onion",

    # Red Onion variants
    "red onion": "Red Onion",
    "red onions": "Red Onion",

    # Russet Potato variants
    "russet": "Russet Potato",
    "russet potato": "Russet Potato",
    "russet potatoes": "Russet Potato",
    "potato russet": "Russet Potato",

    # Green/Red Grape variants
    "green grape": "Green Grape",
    "green grapes": "Green Grape",
    "red grape": "Red Grape",
    "red grapes": "Red Grape",
    "black grape": "Black Grape",
    "black grapes": "Black Grape",

    # Misc common shorthand
    "cukes": "Cucumber",
    "zukes": "Zucchini",
    "zucchini squash": "Zucchini",
    "yellow squash": "Yellow Squash",
    "butternut": "Butternut Squash",
    "butternut squash": "Butternut Squash",
    "iceberg": "Iceberg Lettuce",
    "iceberg lettuce": "Iceberg Lettuce",
    "romaine": "Romaine Lettuce",
    "romaine lettuce": "Romaine Lettuce",
    "baby spinach": "Baby Spinach",
    "grape tomato": "Grape Tomato",
    "grape tomatoes": "Grape Tomato",
    "cherry tomato": "Cherry Tomato",
    "cherry tomatoes": "Cherry Tomato",
    "beefsteak tomato": "Beefsteak Tomato",
    "beefsteak tomatoes": "Beefsteak Tomato",
}

def normalize_item_name(name):
    """Normalize item names so minor variations map to the same key."""
    name = re.sub(r"[,\.\/\\]", " ", name)
    name = " ".join(name.split())
    lower = name.lower()
    if lower in ITEM_ALIASES:
        return ITEM_ALIASES[lower]
    return name.title()


STATIC_LISTS_DIR = os.path.join(os.path.dirname(__file__), "static_price_lists")


def merge_result_into_map(result, item_vendor_map, item_display_name, seen_vendors):
    """Merge one extraction result into the running item->vendor map.
    Returns the vendor name if merged, or None if skipped."""
    vendor = result.get("vendor")
    if not vendor:
        print("    Skipping — no vendor name found in content")
        return None
    # Only use the most recent source per vendor (Slack is processed first, newest-first)
    if vendor in seen_vendors:
        print(f"    Skipping older {vendor} source")
        return None
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
                row = (vendor, float(price), unit.strip())
                if row not in item_vendor_map[key]:
                    item_vendor_map[key].append(row)
            except (ValueError, TypeError):
                pass
    return vendor


def process_static_lists(item_vendor_map, item_display_name, seen_vendors, product_names):
    """Parse any baseline price lists bundled in static_price_lists/ and merge them in.
    These are vendors whose lists are NOT posted to Slack. A newer Slack post for the
    same vendor (processed earlier) takes priority via seen_vendors."""
    if not os.path.isdir(STATIC_LISTS_DIR):
        return
    for filename in sorted(os.listdir(STATIC_LISTS_DIR)):
        path = os.path.join(STATIC_LISTS_DIR, filename)
        if not os.path.isfile(path):
            continue
        lower = filename.lower()
        try:
            with open(path, "rb") as fh:
                file_bytes = fh.read()
            content_blocks = []
            if lower.endswith(".pdf"):
                content_blocks.append({
                    "type": "document",
                    "source": {"type": "base64", "media_type": "application/pdf",
                               "data": base64.b64encode(file_bytes).decode()}
                })
            elif lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp")):
                ext = lower.rsplit(".", 1)[-1]
                media = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
                content_blocks.append({
                    "type": "image",
                    "source": {"type": "base64", "media_type": media,
                               "data": base64.b64encode(file_bytes).decode()}
                })
            elif lower.endswith((".xlsx", ".xls")):
                import openpyxl as _oxl
                wb = _oxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
                text_content = "\n".join(
                    " ".join(str(c) for c in row if c is not None)
                    for sheet in wb.worksheets
                    for row in sheet.iter_rows(values_only=True)
                )
                content_blocks.append({"type": "text", "text": text_content})
            else:
                continue

            print(f"  Parsing static price list: {filename}...")
            result = extract_prices_from_content(content_blocks, "Static List", product_names)
            if result:
                merge_result_into_map(result, item_vendor_map, item_display_name, seen_vendors)
        except Exception as e:
            print(f"  Static list error ({filename}): {e}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("Fetching availability price data from Slack...")

    print("Loading product reference...")
    product_names = load_product_reference()

    channel_id = find_channel_id(CHANNEL_NAME)
    print(f"Found channel #{CHANNEL_NAME}: {channel_id}")

    messages = fetch_channel_history(channel_id, limit=200)
    print(f"Fetched {len(messages)} messages")

    from datetime import timezone, timedelta
    et = timezone(timedelta(hours=-4))
    today_weekday = datetime.now(et).weekday()  # 1=Tuesday, 5=Saturday
    is_scheduled_day = today_weekday in (1, 5)

    if not os.environ.get("FORCE_SEND") and not is_scheduled_day and not has_messages_today(messages):
        print("No new messages today and not a scheduled send day — skipping email.")
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
            if result:
                merge_result_into_map(result, item_vendor_map, item_display_name, seen_vendors)
        except Exception as e:
            print(f"  Parse error: {e}")

    # Merge in baseline price lists not posted to Slack (e.g. Nathel).
    # Slack was processed first, so a newer Slack post for the same vendor wins.
    print("\nProcessing static (non-Slack) price lists...")
    process_static_lists(item_vendor_map, item_display_name, seen_vendors, product_names)

    if not item_vendor_map:
        print("No price data found — nothing to send.")
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
