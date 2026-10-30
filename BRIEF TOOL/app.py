import win32com.client
import pythoncom
import webview
import threading
import pystray
from PIL import Image
from datetime import datetime, timedelta
import os
import re

# ================= KEYWORDS =================
VACATION_KEYWORDS = [
    "leave", "sick leave", "out of office", "not available",
    "casual leave", "medical leave", "cl", "el", "pl",
    "out of desk", "away from desk", "half day", "vacation"
]

APPROVAL_KEYWORDS = [
    "approval", "approve", "approved", "request approval",
    "sign off", "authorization", "consent"
]

MEETING_KEYWORDS = [
    "urgent", "asap", "as soon as possible", "immediate", "critical",
    "important", "priority", "high priority", "most priority",
    "top most priority", "action required", "deadline",
    "eod", "end of the day", "cob", "close of business"
]

HIGH_WORDS = [
    "urgent", "asap", "as soon as possible", "immediate", "critical",
    "important", "priority", "high priority", "most priority",
    "top most priority", "action required", "deadline",
    "eod", "end of the day", "cob", "close of business"
]

HIGH_WORDS    = ["urgent", "asap", "immediate", "critical", "important", "deadline", "action required", "priority"]
MEDIUM_WORDS  = ["request", "pending", "review", "follow up", "reminder"]

# ================= NORMALIZE =================
def normalize(text):
    text = (text or "").lower()
    text = text.replace("_", " ").replace("-", " ")
    return " ".join(text.split())

def match(text, words):
    text = normalize(text)
    return any(re.search(rf"\b{re.escape(w)}\b", text) for w in words)

# ================= PRIORITY =================
def classify(subject, sender):
    combined = normalize(subject + " " + sender)
    if match(combined, HIGH_WORDS):
        return "high"
    if match(combined, MEDIUM_WORDS) or "manager" in combined:
        return "medium"
    return "low"

# ================= OUTLOOK CONNECTION =================
def get_namespace():
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")

# ================= MEETINGS =================
def get_meetings(ns):
    meetings = []
    try:
        calendar = ns.GetDefaultFolder(9)
        items    = calendar.Items

        # CRITICAL: IncludeRecurrences must be set BEFORE Sort
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

                start_dt = datetime(start.year, start.month, start.day,
                                    start.hour, start.minute)
                end_dt   = datetime(end.year, end.month, end.day,
                                    end.hour, end.minute)

                duration   = int((end_dt - start_dt).total_seconds() / 60)
                diff_min   = (start_dt - now).total_seconds() / 60
                is_ongoing = (-duration <= diff_min < 0)

                if is_ongoing:
                    priority, status = "high",   " 🔴 NOW"
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
                    "text":     (f"{time_str} – {end_dt.strftime('%I:%M %p')} "
                                 f"({duration}m){status} — {subject}"),
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

# ================= EMAILS =================
def get_emails(ns):
    """
    Fetches ALL emails from the past 24 hours with no cap.
    Each email is routed into emails / approvals / vacation bucket.
    """
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
                received = datetime(rt.year, rt.month, rt.day,
                                    rt.hour, rt.minute)
                # Stop as soon as we go past 24 hours — list is sorted newest first
                if received < last_24:
                    break

                subject  = msg.Subject or ""
                sender   = str(msg.SenderName or "")
                priority = classify(subject, sender)
                time_str = received.strftime('%H:%M')

                entry = {
                    "text":     f"{time_str} — {subject}",
                    "subject":  subject,
                    "sender":   sender,
                    "time":     time_str,
                    "id":       msg.EntryID,
                    "priority": priority
                }

                if match(subject, APPROVAL_KEYWORDS):
                    approvals.append(entry)
                elif match(subject, VACATION_KEYWORDS):
                    vacation.append(entry)
                else:
                    emails.append(entry)

            except Exception:
                continue

    except Exception as e:
        print(f"Email fetch error: {e}")

    # No cap — return everything received in the last 24 hours
    return emails, approvals, vacation

# ================= MAIN DATA FETCH =================
def get_data():
    try:
        ns = get_namespace()
    except Exception as e:
        print(f"Outlook connect error: {e}")
        return {
            "emails": [], "meetings": [],
            "approvals": [], "vacation": [],
            "error": "Cannot connect to Outlook. Please ensure Outlook is open."
        }

    emails, approvals, vacation = get_emails(ns)
    meetings                    = get_meetings(ns)

    return {
        "emails":    emails,
        "meetings":  meetings,
        "approvals": approvals,
        "vacation":  vacation,
        "error":     ""
    }

# ================= OPEN ITEM =================
def open_item(entry_id):
    try:
        ns   = get_namespace()
        item = ns.GetItemFromID(entry_id)
        item.Display()
    except Exception as e:
        print(f"Open item error: {e}")

# ================= JS API =================
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

# ================= WINDOW =================
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

# ================= TRAY =================
def setup_tray():
    icon = pystray.Icon(
        "BriefDesk",
        Image.open("icon.ico"),
        "BriefDesk",
        pystray.Menu(
            pystray.MenuItem("Open",     lambda i, x: window.show()),
            pystray.MenuItem("Refresh",  lambda i, x: window.evaluate_js("loadData()"))
        )
    )
    icon.run()

# ================= MAIN =================
if __name__ == "__main__":
    threading.Thread(target=setup_tray, daemon=True).start()
    create_window()