import json
import base64
import urllib.request
import urllib.parse
import io
import os
from datetime import datetime

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
    "Goodness Gardens": ["goodness"],
    "Dagele Brothers":  ["dagele"],
    "Dottavio":         ["dottavio"],
}

NOTIFY_EMAILS = ["matt@farmlindproduce.com", "farmlindproduce@gmail.com"]

claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


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
    return base64.urlsafe_b64decode(data + "==")


def extract_text_from_part(part):
    mime = part.get("mimeType", "")
    body_data = part.get("body", {}).get("data", "")
    if mime == "text/plain" and body_data:
        return decode_b64(body_data).decode("utf-8", errors="ignore")
    for sub in part.get("parts", []):
        result = extract_text_from_part(sub)
        if result:
            return result
    return ""


def extract_attachments(payload):
    attachments = []
    def walk(part):
        filename = part.get("filename", "")
        mime = part.get("mimeType", "")
        body = part.get("body", {})
        data = body.get("data", "")
        if filename and data:
            attachments.append((filename, mime, decode_b64(data), None))
        elif filename and body.get("attachmentId"):
            attachments.append((filename, mime, None, body["attachmentId"]))
        for sub in part.get("parts", []):
            walk(sub)
    walk(payload)
    return attachments


def fetch_attachment_data(access_token, message_id, attachment_id):
    result = gmail_get(access_token, f"/users/me/messages/{message_id}/attachments/{attachment_id}")
    return decode_b64(result["data"])


def parse_items_with_claude(content_blocks):
    message = claude.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": content_blocks + [{
                "type": "text",
                "text": (
                    "This is a produce order. Extract every ordered item as a simple list, "
                    "one item per line, including quantities if present. "
                    "Return only the list, no other text. If there are no order items, reply NONE."
                )
            }]
        }]
    )
    text = message.content[0].text.strip()
    if text.upper() == "NONE":
        return []
    return [l.strip() for l in text.splitlines() if l.strip()]


def body_has_any_vendor_section(body_text):
    all_keywords = [kw for kws in VENDORS.values() for kw in kws]
    for line in body_text.splitlines():
        if any(kw in line.strip().lower() for kw in all_keywords):
            return True
    return False


def extract_vendor_section(body_text, vendor_keywords):
    """Extract only the lines belonging to this vendor's section of the email body."""
    lines = body_text.splitlines()
    section_lines = []
    in_section = False

    all_keywords = [kw for kws in VENDORS.values() for kw in kws]

    for line in lines:
        line_lower = line.strip().lower()

        is_our_vendor = any(kw in line_lower for kw in vendor_keywords)
        is_other_vendor = any(kw in line_lower for kw in all_keywords) and not is_our_vendor

        if is_our_vendor:
            in_section = True
            continue
        elif is_other_vendor and in_section:
            break
        elif in_section and line.strip():
            section_lines.append(line.strip())

    return "\n".join(section_lines)


def get_items_from_message(access_token, message_id, payload, vendor_keywords):
    """Get order items for a specific vendor from an email."""
    body_text = extract_text_from_part(payload)

    if body_text.strip():
        section = extract_vendor_section(body_text, vendor_keywords)
        if section.strip():
            try:
                return parse_items_with_claude([{"type": "text", "text": section}])
            except Exception as e:
                print(f"    Claude parse error: {e}")
        # Body has vendor-labeled sections but not ours — wrong email for this vendor
        if body_has_any_vendor_section(body_text):
            return []

    # Body is empty or has no vendor labels — try attachments
    items = []
    for att in extract_attachments(payload):
        filename, mime, data, att_id = att
        if att_id:
            data = fetch_attachment_data(access_token, message_id, att_id)
        print(f"    Processing attachment: {filename}")
        try:
            content_block = None
            if mime == "application/pdf" or filename.lower().endswith(".pdf"):
                content_block = {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": base64.b64encode(data).decode()}}
            elif mime.startswith("image/"):
                content_block = {"type": "image", "source": {"type": "base64", "media_type": mime, "data": base64.b64encode(data).decode()}}
            elif filename.lower().endswith((".xlsx", ".xls")):
                import openpyxl
                wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
                text = "\n".join(" ".join(str(c) for c in row if c is not None) for sheet in wb.worksheets for row in sheet.iter_rows(values_only=True))
                content_block = {"type": "text", "text": text}
            elif filename.lower().endswith((".docx", ".doc")):
                from docx import Document
                doc = Document(io.BytesIO(data))
                content_block = {"type": "text", "text": "\n".join(p.text for p in doc.paragraphs if p.text.strip())}
            else:
                try:
                    content_block = {"type": "text", "text": data.decode("utf-8", errors="ignore")}
                except Exception:
                    pass
            if content_block:
                items.extend(parse_items_with_claude([content_block]))
        except Exception as e:
            print(f"    Attachment error: {e}")

    return items


ORDER_RECIPIENTS = "to:orders@goodnessgardens.com OR to:Office@dagelebrothersproduce.com OR to:anthony@mdottavioproduce.com"


def fetch_order_emails(access_token, date, max_results=20):
    """Fetch sent order emails on a given date."""
    after = date.strftime("%Y/%m/%d")
    from datetime import timedelta
    next_day = (date + timedelta(days=1)).strftime("%Y/%m/%d")
    result = gmail_get(access_token, "/users/me/messages", {
        "q": f"in:sent subject:order ({ORDER_RECIPIENTS}) after:{after} before:{next_day}",
        "maxResults": max_results
    })
    messages = []
    for m in result.get("messages", []):
        full = gmail_get(access_token, f"/users/me/messages/{m['id']}", {"format": "full"})
        messages.append(full)
    return messages


def find_latest_order_timestamp(access_tokens, today):
    """Return the internalDate (ms) of the most recent order email sent today, or None."""
    after = today.strftime("%Y/%m/%d")
    latest = None
    for email_addr, token in access_tokens:
        try:
            result = gmail_get(token, "/users/me/messages", {
                "q": f"in:sent subject:order ({ORDER_RECIPIENTS}) after:{after}",
                "maxResults": 10
            })
            for m in result.get("messages", []):
                full = gmail_get(token, f"/users/me/messages/{m['id']}", {"format": "metadata"})
                ts = int(full.get("internalDate", 0))
                if latest is None or ts > latest:
                    latest = ts
        except Exception:
            pass
    return latest


def reminder_already_sent_after(access_token, since_ms):
    """Return True if an Order Check reminder was already sent after since_ms timestamp."""
    from datetime import timedelta
    since_dt = datetime.fromtimestamp(since_ms / 1000)
    after = since_dt.strftime("%Y/%m/%d")
    try:
        result = gmail_get(access_token, "/users/me/messages", {
            "q": f"in:sent subject:\"Order Check -\" after:{after}",
            "maxResults": 10
        })
        for m in result.get("messages", []):
            full = gmail_get(access_token, f"/users/me/messages/{m['id']}", {"format": "metadata"})
            if int(full.get("internalDate", 0)) > since_ms:
                return True
    except Exception:
        pass
    return False


def find_previous_order_items(access_token, today, vendor_keywords, max_search=30):
    """Find the most recent past order email that contains items for this vendor."""
    result = gmail_get(access_token, "/users/me/messages", {
        "q": f"in:sent subject:order ({ORDER_RECIPIENTS})",
        "maxResults": max_search
    })
    for m in result.get("messages", []):
        full = gmail_get(access_token, f"/users/me/messages/{m['id']}", {"format": "full"})
        ts = int(full.get("internalDate", 0)) / 1000
        msg_date = datetime.fromtimestamp(ts).date()
        if msg_date >= today.date():
            continue
        body = extract_text_from_part(full.get("payload", {}))
        section = extract_vendor_section(body, vendor_keywords)
        if section.strip():
            return msg_date, [full]
    return None, []


def normalize_items(items):
    """Ask Claude to return just the product names, no quantities."""
    response = claude.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": (
                "Strip all leading quantities and numbers from these produce order items and return only the product names, one per line. "
                "Keep descriptors like colors, sizes, and grades (e.g. 'jumbo', 'red', 'fancy', '60ct'). "
                "No bullet points, no numbers at the start, no extra text.\n\n"
                + "\n".join(items)
            )
        }]
    )
    return [l.strip().lower() for l in response.content[0].text.strip().splitlines() if l.strip()]



def send_reminder(access_token, sender_email, missing_by_vendor, vendor_today, today, skipped_vendors=None):
    subject = f"Order Check - {today.strftime('%A %b %d')}"
    lines = [
        "Hi Matt,",
        "",
        "Here is a summary of tonight's orders and what may be missing:",
        "",
        "=" * 40,
    ]
    for vendor, missing in missing_by_vendor.items():
        sent = vendor_today.get(vendor, [])
        lines.append("")
        lines.append(f"  {vendor.upper()}")
        lines.append(f"  {'-' * 36}")
        lines.append(f"  This week you sent these items:")
        for item in sent:
            lines.append(f"    - {item}")
        lines.append("")
        lines.append(f"  You may have forgotten these items ({len(missing)}):")
        for item in missing:
            lines.append(f"    - {item}")
        lines.append("")
        lines.append("=" * 40)
    if skipped_vendors:
        lines.append("")
        lines.append("  *** YOU DID NOT SEND AN ORDER TO: ***")
        for vendor in skipped_vendors:
            lines.append(f"    - {vendor}")
        lines.append("")
        lines.append("=" * 40)
    lines += ["", "- Farmlind Order Bot"]

    body = "\n".join(lines)
    raw = f"From: {sender_email}\r\nTo: {', '.join(NOTIFY_EMAILS)}\r\nSubject: {subject}\r\n\r\n{body}"
    encoded = base64.urlsafe_b64encode(raw.encode()).decode()
    gmail_post(access_token, "/users/me/messages/send", {"raw": encoded})
    print("Reminder sent.")


def main(override_date=None):
    today = override_date if override_date else datetime.now()
    print(f"Checking orders for {today.strftime('%Y-%m-%d %A')}")

    access_tokens = []
    for account in ACCOUNTS:
        try:
            token = get_access_token(account["refresh_token"])
            access_tokens.append((account["email"], token))
            print(f"Authenticated: {account['email']}")
        except Exception as e:
            print(f"Auth failed for {account['email']}: {e}")

    if not access_tokens:
        print("No valid tokens, exiting.")
        return

    # When running on a schedule, only proceed if a new order was sent at least 3 minutes ago
    # and we haven't already sent a reminder for it
    if not override_date and not os.environ.get("FORCE_SEND"):
        from datetime import timedelta
        latest_ts = find_latest_order_timestamp(access_tokens, today)
        if not latest_ts:
            print("No order sent today — nothing to do.")
            return
        order_age_minutes = (today.timestamp() * 1000 - latest_ts) / 60000
        if order_age_minutes < 3:
            print(f"Most recent order sent {order_age_minutes:.1f} min ago — waiting for 3 min mark.")
            return
        sender_email, sender_token = access_tokens[0]
        if reminder_already_sent_after(sender_token, latest_ts):
            print("Reminder already sent for this order — nothing to do.")
            return

    vendor_today = {}
    vendor_previous = {}

    for vendor_name, keywords in VENDORS.items():
        print(f"\nChecking {vendor_name}...")
        today_items = []
        prev_items = []

        seen_today_ids = set()
        for email_addr, token in access_tokens:
            try:
                todays_msgs = fetch_order_emails(token, today)
                for msg in todays_msgs:
                    if msg["id"] in seen_today_ids:
                        continue
                    seen_today_ids.add(msg["id"])
                    items = get_items_from_message(token, msg["id"], msg.get("payload", {}), keywords)
                    if items:
                        print(f"  Found today's order from {email_addr} ({len(items)} items)")
                        today_items.extend(items)
            except Exception as e:
                print(f"  Error fetching today for {email_addr}: {e}")

        if today_items:
            # Find the most recent previous date across all accounts
            best_prev_date = None
            for email_addr, token in access_tokens:
                try:
                    pd, _ = find_previous_order_items(token, today, keywords)
                    if pd and (best_prev_date is None or pd > best_prev_date):
                        best_prev_date = pd
                except Exception:
                    pass

            if best_prev_date:
                print(f"  Previous order date: {best_prev_date}")
                seen_prev_ids = set()
                seen_prev_sections = set()
                for email_addr, token in access_tokens:
                    try:
                        prev_msgs = fetch_order_emails(token, datetime.combine(best_prev_date, datetime.min.time()))
                        for msg in prev_msgs:
                            if msg["id"] in seen_prev_ids:
                                continue
                            seen_prev_ids.add(msg["id"])
                            body = extract_text_from_part(msg.get("payload", {}))
                            section = extract_vendor_section(body, keywords).strip()
                            if section in seen_prev_sections:
                                continue
                            seen_prev_sections.add(section)
                            items = get_items_from_message(token, msg["id"], msg.get("payload", {}), keywords)
                            if items:
                                print(f"  Found previous order from {email_addr} ({len(items)} items)")
                                prev_items.extend(items)
                    except Exception as e:
                        print(f"  Error fetching previous for {email_addr}: {e}")

        vendor_today[vendor_name] = today_items
        vendor_previous[vendor_name] = prev_items
        print(f"  Today: {len(today_items)} items, Previous: {len(prev_items)} items")

    ordered_today = any(items for items in vendor_today.values())
    if not ordered_today:
        print("\nNo orders sent today — nothing to compare.")
        return

    missing_by_vendor = {}
    for vendor_name in VENDORS:
        current_items = vendor_today[vendor_name]
        prev_items = vendor_previous[vendor_name]
        if not current_items or not prev_items:
            continue
        try:
            today_names = normalize_items(current_items)
            prev_names = normalize_items(prev_items)

            missing = []
            seen = set()
            for name in prev_names:
                if name in seen:
                    continue
                seen.add(name)
                found = any(name in today_name or today_name in name for today_name in today_names)
                if not found:
                    missing.append(name)

            if missing:
                missing_by_vendor[vendor_name] = missing
        except Exception as e:
            print(f"  Comparison error for {vendor_name}: {e}")

    # Check for vendors Matt has ordered from before but skipped entirely today
    skipped_vendors = []
    for vendor_name, keywords in VENDORS.items():
        if not vendor_today[vendor_name]:
            # See if he's ever ordered from this vendor before
            for email_addr, token in access_tokens:
                try:
                    pd, _ = find_previous_order_items(token, today, keywords)
                    if pd:
                        skipped_vendors.append(vendor_name)
                        break
                except Exception:
                    pass

    if not missing_by_vendor and not skipped_vendors:
        print("\nAll previous items accounted for. No reminder needed.")
        return

    sender_email, sender_token = access_tokens[0]
    print(f"\nMissing items: {missing_by_vendor}")
    print(f"Skipped vendors: {skipped_vendors}")
    send_reminder(sender_token, sender_email, missing_by_vendor, vendor_today, today, skipped_vendors)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        override = datetime.strptime(sys.argv[1], "%Y-%m-%d")
        main(override_date=override)
    else:
        main()
