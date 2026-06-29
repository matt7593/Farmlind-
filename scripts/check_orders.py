#!/usr/bin/env python3
import json, base64, urllib.request, urllib.parse, os, subprocess
from datetime import datetime
import anthropic

CLIENT_ID, CLIENT_SECRET = os.environ["GMAIL_CLIENT_ID"], os.environ["GMAIL_CLIENT_SECRET"]
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

TEST_EMAIL = os.environ.get("TEST_EMAIL", "").strip()
NOTIFY = [TEST_EMAIL] if TEST_EMAIL else ["matt@farmlindproduce.com", "farmlindproduce@gmail.com"]
STATE = ".order_state.json"
claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

def get_token(rt):
    data = urllib.parse.urlencode({"client_id": CLIENT_ID, "client_secret": CLIENT_SECRET, "refresh_token": rt, "grant_type": "refresh_token"}).encode()
    with urllib.request.urlopen(urllib.request.Request("https://oauth2.googleapis.com/token", data=data)) as r:
        return json.loads(r.read())["access_token"]

def gmail_get(tok, path, params=None):
    url = f"https://gmail.googleapis.com/gmail/v1{path}" + (f"?{urllib.parse.urlencode(params)}" if params else "")
    with urllib.request.urlopen(urllib.request.Request(url, headers={"Authorization": f"Bearer {tok}"})) as r:
        return json.loads(r.read())

def gmail_post(tok, path, body):
    url = f"https://gmail.googleapis.com/gmail/v1{path}"
    data = json.dumps(body).encode()
    with urllib.request.urlopen(urllib.request.Request(url, data=data, headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json"})) as r:
        return json.loads(r.read())

def get_body(part):
    mime, b64 = part.get("mimeType", ""), part.get("body", {}).get("data", "")
    if mime == "text/plain" and b64:
        return base64.urlsafe_b64decode(b64 + "==").decode("utf-8", errors="ignore")
    for sub in part.get("parts", []):
        r = get_body(sub)
        if r: return r
    return ""

def get_sent_emails(tok, days=30):
    emails = []
    try:
        for m in gmail_get(tok, "/users/me/messages", {"q": f"in:sent newer_than:{days}d", "maxResults": 100}).get("messages", []):
            full = gmail_get(tok, f"/users/me/messages/{m['id']}", {"format": "full"})
            headers = full.get("payload", {}).get("headers", [])
            to_val = next((h["value"] for h in headers if h.get("name", "").lower() == "to"), "")
            body = get_body(full.get("payload", {}))
            ts = int(full.get("internalDate", 0))
            if to_val and body:
                emails.append({"to": to_val.lower(), "body": body, "ts": ts})
    except Exception as e: print(f"Error: {e}")
    return emails

def vendor_in_to(vendor_email, to_header):
    return vendor_email.lower() in to_header.lower()

def items(emails):
    if not emails: return []
    try:
        r = claude.messages.create(model="claude-opus-4-8", max_tokens=1000, messages=[{"role": "user", "content": f"List items ordered, one per line:\n\n" + "\n\n".join([e["body"] for e in emails])}])
        return [x.strip() for x in r.content[0].text.strip().split("\n") if x.strip()]
    except: return []

def load_state():
    try:
        with open(STATE) as f: return json.load(f)
    except: return {}

def save_state(s):
    with open(STATE, "w") as f: json.dump(s, f, indent=2)
    subprocess.run(["git", "config", "user.email", "github-actions@farmlind.local"], capture_output=True, check=False)
    subprocess.run(["git", "config", "user.name", "GitHub Actions"], capture_output=True, check=False)
    subprocess.run(["git", "add", STATE], capture_output=True, check=False)
    subprocess.run(["git", "commit", "-m", "Update state"], capture_output=True)
    subprocess.run(["git", "push"], capture_output=True)

def build_body(today, prev, date):
    lines = ["ORDER SUMMARY", str(date.strftime("%A, %B %d")), "=" * 50, ""]
    for v in sorted(VENDORS.keys()):
        t, p = today.get(v, []), prev.get(v, [])
        lines += ["", v, "-" * len(v)]
        if t:
            lines.append("TODAY:")
            lines += [f"  {i}. {x}" for i, x in enumerate(t, 1)]
        else:
            lines.append("Not ordered")
        if p and (m := [x for x in p if x not in t]):
            lines += ["", "MISSING:"] + [f"  {i}. {x}" for i, x in enumerate(m, 1)]
    return "\n".join(lines)

def send(tok, sender, today, prev, date):
    body = build_body(today, prev, date)
    raw = f"From: {sender}\r\nTo: {', '.join(NOTIFY)}\r\nSubject: Orders\r\n\r\n{body}"
    gmail_post(tok, "/users/me/messages/send", {"raw": base64.urlsafe_b64encode(raw.encode()).decode()})

def main():
    today = datetime.now().date()
    now = int(datetime.now().timestamp() * 1000)
    tokens = []
    for a in ACCOUNTS:
        try: tokens.append((a["email"], get_token(a["refresh_token"]))); print(f"✓ {a['email']}")
        except Exception as e: print(f"✗ {e}")
    if not tokens: return
    
    all_sent = []
    for _, tok in tokens:
        all_sent.extend(get_sent_emails(tok, 1))
    
    if not all_sent:
        print(f"  Looking for: {vendor_email.lower()}")
        print("No emails sent today")
        return
    
    print(f"Found {len(all_sent)} sent emails\n")
    
    today_items, prev_items, found = {}, {}, []
    
    for vendor, vendor_email in VENDORS.items():
        print(f"{vendor}:")
        
        today_emails = [e for e in all_sent if e["ts"] >= today_ts and vendor_email.lower() in e["to"]]
        prev_emails = [e for e in all_sent if e["ts"] < today_ts and vendor_email.lower() in e["to"]]
        
        if today_emails:
            found_vendors.append(vendor)
            today_items[vendor] = extract_items(today_emails)
            print(f"  Today: {len(today_items[vendor])} items")
        else:
            today_items[vendor] = []
            print(f"  No email sent")
        
        if prev_emails:
            most_recent = max(prev_emails, key=lambda x: x["ts"])
            prev_items[vendor] = extract_items([most_recent])
            print(f"  Previous: {len(prev_items[vendor])} items")
        
        print()
    state, key = load_state(), today.isoformat()
    st = state.get(key, {})
    ft = st.get("ft")
    if found and not ft: st["ft"] = ft = now; state[key] = st; save_state(state)
    
    scan = os.environ.get("SCAN_ONLY", "").lower() == "true"
    force = os.environ.get("FORCE_SEND", "").lower() == "true"
    dry = os.environ.get("DRY_RUN", "").lower() == "true"
    sent = st.get("sent", False)
    
    if scan:
        if ft: mins = (now - ft) / 1000 / 60; print(f"First: {mins:.0f}m - {'Ready' if mins >= 30 else f'wait {30-mins:.0f}m'}")
        return
    if sent and not force: print("Already sent"); return
    if not force and ft and (now - ft) / 1000 / 60 < 30: mins_left = 30 - (now - ft) / 1000 / 60; print(f"Wait {mins_left:.0f}m"); return
    if (force or (ft and (now - ft) / 1000 / 60 >= 30)) and found:
        sender, tok = tokens[0]
        if dry: print(build_body(today_items, prev_items, today))
        else: send(tok, sender, today_items, prev_items, today); st["sent"] = True; state[key] = st; save_state(state); print("Email sent")

if __name__ == "__main__": main()
