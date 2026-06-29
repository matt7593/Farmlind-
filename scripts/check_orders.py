import json, base64, urllib.request, urllib.parse, os, subprocess
from datetime import datetime
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

TEST_EMAIL = os.environ.get("TEST_EMAIL", "").strip()
NOTIFY = [TEST_EMAIL] if TEST_EMAIL else ["matt@farmlindproduce.com"]
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

def get_sent_emails(tok):
    emails = []
    try:
        for m in gmail_get(tok, "/users/me/messages", {"q": "in:sent", "maxResults": 200}).get("messages", []):
            full = gmail_get(tok, f"/users/me/messages/{m['id']}", {"format": "full"})
            headers = full.get("payload", {}).get("headers", [])
            to_val = next((h["value"] for h in headers if h.get("name", "").lower() == "to"), "")
            body = get_body(full.get("payload", {}))
            ts = int(full.get("internalDate", 0))
            if to_val and body:
                emails.append({"to": to_val.lower(), "body": body, "ts": ts})
    except Exception as e:
        print(f"Error: {e}")
    return emails

def extract_items(emails):
    if not emails: return []
    try:
        text = "\n\n".join([e["body"] for e in emails])
        r = claude.messages.create(model="claude-opus-4-8", max_tokens=1000, messages=[{"role": "user", "content": f"Extract items, one per line:\n\n{text}"}])
        return [x.strip() for x in r.content[0].text.strip().split("\n") if x.strip()]
    except:
        return []

def load_state():
    try:
        with open(STATE) as f: return json.load(f)
    except: return {}

def save_state(s):
    with open(STATE, "w") as f: json.dump(s, f, indent=2)
    subprocess.run(["git", "config", "user.email", "github-actions@farmlind.local"], capture_output=True)
    subprocess.run(["git", "config", "user.name", "GitHub Actions"], capture_output=True)
    subprocess.run(["git", "add", STATE], capture_output=True)
    subprocess.run(["git", "commit", "-m", "Update"], capture_output=True)
    subprocess.run(["git", "push"], capture_output=True)

def main():
    today = datetime.now().date()
    now = int(datetime.now().timestamp() * 1000)
    today_ts = int((datetime.now().timestamp() - 86400) * 1000)  # Last 24 hours

    tokens = []
    for a in ACCOUNTS:
        try: 
            tokens.append((a["email"], get_token(a["refresh_token"])))
            print(f"✓ {a['email']}")
        except Exception as e: 
            print(f"✗ {e}")

    if not tokens: return
    print()

    all_sent = []
    for _, tok in tokens:
        all_sent.extend(get_sent_emails(tok))

    print(f"Found {len(all_sent)} total emails\n")
    
    if all_sent:
        print("Sample emails:")
        for e in all_sent[:3]:
            print(f"  To: {e['to'][:70]}")
        print()

    today_items = {}
    prev_items = {}
    found = []

    for vendor, ve in VENDORS.items():
        print(f"{vendor} (looking for: {ve}):")
        te = [e for e in all_sent if e["ts"] >= today_ts and ve.lower() in e["to"]]
        pe = [e for e in all_sent if e["ts"] < today_ts and ve.lower() in e["to"]]
        
        print(f"  Today: {len(te)}, Previous: {len(pe)}")

        if te:
            found.append(vendor)
            today_items[vendor] = extract_items(te)
            print(f"  Items today: {len(today_items[vendor])}")
        else:
            today_items[vendor] = []

        if pe:
            pr = max(pe, key=lambda x: x["ts"])
            prev_items[vendor] = extract_items([pr])

        print()

    print(f"Found vendors: {found}\n")

    if not found:
        print("NO ORDERS SENT TODAY")
        return

    state, key = load_state(), today.isoformat()
    st = state.get(key, {})
    if found and not st.get("ft"): st["ft"] = now; state[key] = st; save_state(state)

    force = os.environ.get("FORCE_SEND", "").lower() == "true"
    dry = os.environ.get("DRY_RUN", "").lower() == "true"
    
    lines = ["ORDER SUMMARY\n"]
    for v in sorted(VENDORS.keys()):
        t, p = today_items.get(v, []), prev_items.get(v, [])
        lines.append(f"{v}:")
        if t:
            lines.append("Ordered today:")
            lines.extend([f"  • {x}" for x in t])
        else:
            lines.append("Not ordered")
        if p:
            m = [x for x in p if x not in t]
            if m:
                lines.append("Missing:")
                lines.extend([f"  ✗ {x}" for x in m])
        lines.append("")

    body = "\n".join(lines)
    if dry:
        print(body)
    elif found:
        sender, tok = tokens[0]
        raw = f"From: {sender}\r\nTo: {', '.join(NOTIFY)}\r\nSubject: Orders\r\n\r\n{body}"
        gmail_post(tok, "/users/me/messages/send", {"raw": base64.urlsafe_b64encode(raw.encode()).decode()})
        print("✓ Email sent")

if __name__ == "__main__":
    main()
