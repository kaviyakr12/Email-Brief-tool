import win32com.client
import pythoncom
import webview
import threading
import pystray
import requests
import json
from PIL import Image
from datetime import datetime, timedelta
import os
import re

# ─────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────
OLLAMA_URL   = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "mistral"

# The logged-in user's name/email — filled at startup from Outlook
MY_NAME  = ""
MY_EMAIL = ""

# ─────────────────────────────────────────────────────────────────────
#  KEYWORDS — for CATEGORY routing only (not priority)
# ─────────────────────────────────────────────────────────────────────
LEAVE_KEYWORDS = [
    "leave", "sick leave", "out of office", "not available",
    "casual leave", "medical leave", "cl", "el", "pl",
    "out of desk", "away from desk", "half day", "vacation",
    "wfh", "work from home", "day off", "annual leave",
    "personal leave", "maternity leave", "paternity leave",
    "ooo", "time off", "absent"
]
APPROVAL_KEYWORDS = [
    "approval", "approve", "approved", "request approval",
    "sign off", "authorization", "consent", "sanction",
    "pending approval", "awaiting approval", "needs approval",
    "for your approval", "kindly approve"
]
MEETING_KEYWORDS = [
    "meeting", "meet", "invite", "invitation", "calendar",
    "conference", "call scheduled", "zoom", "teams call",
    "google meet", "webex", "standup", "stand-up", "sync",
    "catchup", "catch-up", "all hands", "town hall",
    "workshop", "seminar", "interview scheduled", "appointment"
]

# ─────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────
def normalize(text):
    text = (text or "").lower()
    text = text.replace("_", " ").replace("-", " ")
    return " ".join(text.split())

def match_any(text, words):
    t = normalize(text)
    for w in words:
        w = normalize(w)
        if " " in w:
            if w in t:
                return True
        else:
            if re.search(rf"\b{re.escape(w)}\b", t):
                return True
    return False

def keyword_category(subject):
    if match_any(subject, LEAVE_KEYWORDS):    return "leave"
    if match_any(subject, APPROVAL_KEYWORDS): return "approvals"
    if match_any(subject, MEETING_KEYWORDS):  return "meetings"
    return "emails"


# ─────────────────────────────────────────────────────────────────────
#  LLM-POWERED PERSON-AWARE PRIORITY
# ─────────────────────────────────────────────────────────────────────
def llm_priority(subject, body, sender, to_list, cc_list, my_name, my_email):
    """
    Ask Mistral to decide the priority of this email FOR THIS SPECIFIC PERSON.
    Returns "high", "medium", or "low".
    Falls back to keyword priority if Ollama is unavailable.
    """
    # Build a compact context for the LLM
    to_str  = ", ".join(to_list[:5])   if to_list  else "unknown"
    cc_str  = ", ".join(cc_list[:5])   if cc_list  else "none"
    snippet = (body or "")[:600].replace("\r\n", " ").replace("\n", " ").strip()

    prompt = f"""You are an intelligent email assistant. 
Decide the priority of this email SPECIFICALLY FOR: {my_name} ({my_email}).

Email details:
- From: {sender}
- To: {to_str}
- CC: {cc_str}
- Subject: {subject}
- Body snippet: {snippet}

Rules:
- HIGH: {my_name} is directly addressed and must take action urgently, or it is time-sensitive/critical for them.
- MEDIUM: {my_name} needs to read it or act soon but it is not urgent.
- LOW: {my_name} is just CC'd for awareness, or the email doesn't require any action from them.

Consider:
- If {my_name} is only in CC and the action is clearly for someone else → LOW.
- If {my_name} is the TO recipient and action is required → HIGH or MEDIUM.
- Managers CC'd on routine approvals that their TL handles → LOW.
- If the email says "please do X" and {my_name} is addressed → HIGH.

Respond with exactly one word: high, medium, or low. Nothing else."""

    try:
        resp = requests.post(
            OLLAMA_URL,
            json={"model": OLLAMA_MODEL, "prompt": prompt, "stream": False},
            timeout=15
        )
        if resp.status_code == 200:
            text = resp.json().get("response", "").strip().lower()
            if "high"   in text: return "high"
            if "medium" in text: return "medium"
            if "low"    in text: return "low"
    except Exception as e:
        print(f"LLM priority error: {e}")

    # Fallback: simple keyword
    combined = subject + " " + (body or "")[:200]
    HIGH_WORDS   = ["urgent","asap","critical","action required","deadline","overdue","escalat","escalation","breach","blocker","p0","p1","eod","cob", "as soon as possible", "end of the day", "First half"]
    MEDIUM_WORDS = ["important","follow up","reminder","request","pending","review","please","need","require","waiting"]
    if match_any(combined, HIGH_WORDS):   return "high"
    if match_any(combined, MEDIUM_WORDS): return "medium"
    return "low"


# ─────────────────────────────────────────────────────────────────────
#  OUTLOOK CONNECTION
# ─────────────────────────────────────────────────────────────────────
def get_namespace():
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def get_my_identity(ns):
    """Get the current user's name and email from Outlook."""
    global MY_NAME, MY_EMAIL
    try:
        acct = ns.CurrentUser
        MY_NAME  = acct.Name  or ""
        MY_EMAIL = acct.Address or ""
    except Exception as e:
        print(f"Could not get identity: {e}")


# ─────────────────────────────────────────────────────────────────────
#  MEETINGS — from Calendar folder
# ─────────────────────────────────────────────────────────────────────
def get_meetings(ns):
    meetings = []
    try:
        calendar = ns.GetDefaultFolder(9)
        items    = calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        now   = datetime.now()
        today = now.date()
        start_str   = f"{today.month}/{today.day}/{today.year} 12:00 AM"
        end_str     = f"{today.month}/{today.day}/{today.year} 11:59 PM"
        restriction = f"[Start] >= '{start_str}' AND [Start] <= '{end_str}'"
        restricted  = items.Restrict(restriction)

        for appt in restricted:
            try:
                start     = appt.Start
                end       = appt.End
                subject   = appt.Subject or "(No Subject)"
                entry_id  = appt.EntryID
                organizer = getattr(appt, "Organizer", "") or ""
                location  = getattr(appt, "Location",  "") or ""

                start_dt = datetime(start.year, start.month, start.day, start.hour, start.minute)
                end_dt   = datetime(end.year,   end.month,   end.day,   end.hour,   end.minute)
                duration = int((end_dt - start_dt).total_seconds() / 60)
                diff_min = (start_dt - now).total_seconds() / 60
                is_ongoing = (-duration <= diff_min < 0)

                if is_ongoing:
                    priority, status = "high",   "  NOW"
                elif diff_min < 0:
                    priority, status = "low",    " ✓ Done"
                elif diff_min < 30:
                    priority, status = "high",   ""
                elif diff_min < 120:
                    priority, status = "medium", ""
                else:
                    priority, status = "low",    ""

                time_str = start_dt.strftime('%I:%M %p')
                preview  = f"Organizer: {organizer}" if organizer else ""
                if location:
                    preview += (f" | {location}" if preview else f"Location: {location}")

                meetings.append({
                    "text":     f"{time_str} – {end_dt.strftime('%I:%M %p')} ({duration}m){status} — {subject}",
                    "subject":  subject,
                    "sender":   organizer or "Unknown Organizer",
                    "time":     time_str,
                    "preview":  preview,
                    "id":       entry_id,
                    "priority": priority
                })
            except Exception as e:
                print(f"Meeting item error: {e}")
                continue
    except Exception as e:
        print(f"Calendar fetch error: {e}")
    return meetings


# ─────────────────────────────────────────────────────────────────────
#  EMAILS — last 24 hours from Inbox, with LLM person-aware priority
# ─────────────────────────────────────────────────────────────────────
def get_emails(ns):
    emails, approvals, vacation = [], [], []
    try:
        inbox = ns.GetDefaultFolder(6)
        msgs  = inbox.Items
        msgs.Sort("[ReceivedTime]", True)

        now     = datetime.now()
        last_24 = now - timedelta(days=1)

        for msg in msgs:
            try:
                rt       = msg.ReceivedTime
                received = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute)
                if received < last_24:
                    break

                subject  = msg.Subject or ""
                sender   = str(msg.SenderName or "")
                body     = msg.Body or ""
                time_str = received.strftime('%H:%M')
                preview  = body[:150].replace("\n", " ").replace("\r", " ").strip()

                # Extract TO and CC recipients
                to_list, cc_list = [], []
                try:
                    for r in msg.Recipients:
                        name = getattr(r, "Name", "") or ""
                        addr = getattr(r, "Address", "") or ""
                        rtype = getattr(r, "Type", 1)  # 1=TO, 2=CC, 3=BCC
                        if rtype == 1:
                            to_list.append(f"{name} <{addr}>")
                        elif rtype == 2:
                            cc_list.append(f"{name} <{addr}>")
                except Exception:
                    pass

                # LLM decides priority FOR THIS SPECIFIC USER
                priority = llm_priority(
                    subject, body, sender,
                    to_list, cc_list,
                    MY_NAME, MY_EMAIL
                )

                entry = {
                    "text":     f"{time_str} — {subject}",
                    "subject":  subject,
                    "sender":   sender,
                    "time":     time_str,
                    "preview":  preview,
                    "id":       msg.EntryID,
                    "priority": priority,
                    "to":       ", ".join(to_list[:3]),
                    "cc":       ", ".join(cc_list[:3]),
                }

                cat = keyword_category(subject)
                if cat == "leave":
                    vacation.append(entry)
                elif cat == "approvals":
                    approvals.append(entry)
                else:
                    emails.append(entry)

            except Exception:
                continue
    except Exception as e:
        print(f"Email fetch error: {e}")
    return emails, approvals, vacation


# ─────────────────────────────────────────────────────────────────────
#  MAIN DATA FETCH
# ─────────────────────────────────────────────────────────────────────
def get_data():
    try:
        ns = get_namespace()
        get_my_identity(ns)
    except Exception as e:
        print(f"Outlook connect error: {e}")
        return {
            "emails": [], "meetings": [],
            "approvals": [], "vacation": [],
            "error": "Cannot connect to Outlook. Please ensure Outlook is open.",
            "user": ""
        }

    emails, approvals, vacation = get_emails(ns)
    meetings                    = get_meetings(ns)

    return {
        "emails":    emails,
        "meetings":  meetings,
        "approvals": approvals,
        "vacation":  vacation,
        "error":     "",
        "user":      MY_NAME
    }


# ─────────────────────────────────────────────────────────────────────
#  OPEN ITEM
# ─────────────────────────────────────────────────────────────────────
def open_item(entry_id):
    try:
        ns   = get_namespace()
        item = ns.GetItemFromID(entry_id)
        item.Display()
    except Exception as e:
        print(f"Open item error: {e}")


# ─────────────────────────────────────────────────────────────────────
#  JS API
# ─────────────────────────────────────────────────────────────────────
class API:
    def load(self):
        return get_data()

    def refresh(self):
        return get_data()

    def open_item(self, entry_id):
        open_item(entry_id)

    def hide(self):
        window.hide()

    def maximize(self):
        window.maximize()

    def restore(self):
        window.restore()


# ─────────────────────────────────────────────────────────────────────
#  WINDOW
# ─────────────────────────────────────────────────────────────────────
def create_window():
    global window
    screen = webview.screens[0]
    window = webview.create_window(
        "BriefDesk",
        os.path.abspath("index.html"),
        width=420,
        height=640,
        x=screen.width  - 440,
        y=screen.height - 680,
        resizable=True,
        js_api=API()
    )
    webview.start()


# ─────────────────────────────────────────────────────────────────────
#  TRAY
# ─────────────────────────────────────────────────────────────────────
def setup_tray():
    icon = pystray.Icon(
        "BriefDesk",
        Image.open("icon.ico"),
        "BriefDesk",
        pystray.Menu(
            pystray.MenuItem("Open",    lambda i, x: window.show()),
            pystray.MenuItem("Refresh", lambda i, x: window.evaluate_js("loadData()"))
        )
    )
    icon.run()


# ─────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    threading.Thread(target=setup_tray, daemon=True).start()
    create_window()