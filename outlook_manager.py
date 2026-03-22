"""
Outlook Inbox Manager - GUI Tool via COM Automation
Requires: pip install pywin32
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import re
from enum import Enum
from pathlib import Path
from datetime import datetime

try:
    import win32com.client
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
    import win32com.client

# ─────────────────────────────────────────────
BASE = Path(r"C:\Users\dirk\OutlookSort")
RULES_FILE    = BASE / "outlook_rules.json"
KEYWORDS_FILE = BASE / "outlook_keywords.json"
STATS_FILE    = BASE / "outlook_stats.json"

ALL_FOLDERS = [
    "Finanzen",
    "Finanzen/Zahlungen",
    "Finanzen/Kreditkarte",
    "Finanzen/Bank und Versicherung",
    "Finanzen/Rechnungen",
    "Shopping",
    "Shopping/Amazon",
    "Shopping/Pakete und Lieferungen",
    "Shopping/Shops und Einkauf",
    "Shopping/Kleinanzeigen",
    "Auto",
    "Auto/Autohaus Molnar",
    "Auto/Autoverkauf und Suche",
    "Immobilien",
    "Immobilien/Böhringen Lichtensteinstr",
    "Immobilien/Wittlingen Schwalbenstr",
    "Kinder und Schule",
    "Kinder und Schule/GEG",
    "Beruf",
    "Beruf/Fierthbauer",
    "Beruf/Holzher",
    "Beruf/Weiterbildung",
    "Beruf/Bewerbung",
    "Vereine",
    "Vereine/TSV Wittlingen",
    "Vereine/TSV Böhringen",
    "Ehrenamt",
    "Ehrenamt/Ortschaftsrat",
    "Hobby",
    "Persönlich",
    "Persönlich/Gesundheit",
    "Persönlich/Urlaub und Reise",
    "Persönlich/Passes",
    "Persönlich/Steuer",
    "Haus und Energie",
    "Digital",
    "Digital/Software und Lizenzen",
    "Digital/Netzwerk und Backup",
    "Digital/Scanned",
    "Online-Dienste",
    "Online-Dienste/Web und Hosting",
    "Online-Dienste/Vergleichsportale",
    "Online-Dienste/Telekom und Anbieter",
    "Online-Dienste/Web.de",
    "Newsletter",
    "Newsletter/Heise CT",
    "Newsletter/Medium",
    "Familie",
    "Familie/Julia",
    "Familie/Hannes",
    "Familie/Martina und Steffen",
    "Familie/Peter Koval",
    "Familie/Archiv",
    "Kirche",
    "Persönlich/Genealogie",
    "_Archiv",
]

QUICK_FOLDERS = [
    ("Amazon",     "Shopping/Amazon"),
    ("Pakete",     "Shopping/Pakete und Lieferungen"),
    ("Shops",      "Shopping/Shops und Einkauf"),
    ("Rechnungen", "Finanzen/Rechnungen"),
    ("Zahlungen",  "Finanzen/Zahlungen"),
    ("Bank/Vers.", "Finanzen/Bank und Versicherung"),
    ("Software",   "Digital/Software und Lizenzen"),
    ("Online",     "Online-Dienste"),
    ("Urlaub",     "Persönlich/Urlaub und Reise"),
    ("Gesundheit", "Persönlich/Gesundheit"),
    ("Auto",       "Auto/Autoverkauf und Suche"),
    ("Kinder",     "Kinder und Schule"),
    ("Ehrenamt",   "Ehrenamt/Ortschaftsrat"),
    ("Heise",      "Newsletter/Heise CT"),
    ("Familie",    "Familie"),
    ("Kirche",     "Kirche"),
    ("Archiv",     "_Archiv"),
]


def strip_html(html: str) -> str:
    text = re.sub(r"<style[^>]*>.*?</style>", " ", html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<script[^>]*>.*?</script>", " ", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</?p[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    for ent, ch in [("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"),
                    ("&gt;", ">"), ("&quot;", '"'), ("&#39;", "'")]:
        text = text.replace(ent, ch)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


# ─────────────────────────────────────────────
# Rule Engine
# ─────────────────────────────────────────────

class RuleEngine:
    def __init__(self):
        self.rules = []
        self._load()

    def _load(self):
        if RULES_FILE.exists():
            try:
                with open(RULES_FILE, "r", encoding="utf-8") as f:
                    self.rules = json.load(f)
            except Exception:
                self.rules = []

    def save(self):
        with open(RULES_FILE, "w", encoding="utf-8") as f:
            json.dump(self.rules, f, indent=2, ensure_ascii=False)

    def add(self, pattern: str, action: str, dest: str = ""):
        pattern = pattern.lower()
        self.rules = [r for r in self.rules if r["pattern"] != pattern]
        self.rules.append({"pattern": pattern, "action": action, "dest": dest})
        self.save()

    def delete(self, index: int):
        del self.rules[index]
        self.save()

    def find(self, sender_email: str):
        email = sender_email.lower()
        for rule in self.rules:
            if rule["pattern"] in email:
                return rule
        return None


# ─────────────────────────────────────────────
# Suggestion Engine
# ─────────────────────────────────────────────

class Confidence(Enum):
    HIGH    = "HIGH"
    MEDIUM  = "MEDIUM"
    UNKNOWN = "UNKNOWN"

class Suggestion:
    def __init__(self, folder: str, confidence: Confidence, reason: str, action: str = "move"):
        self.folder     = folder
        self.confidence = confidence
        self.reason     = reason
        self.action     = action

class SuggestionEngine:
    def __init__(self, rule_engine: RuleEngine):
        self.rules = rule_engine
        self.keyword_rules = []
        self._load()

    def _load(self):
        if KEYWORDS_FILE.exists():
            try:
                with open(KEYWORDS_FILE, "r", encoding="utf-8") as f:
                    self.keyword_rules = json.load(f)
            except Exception:
                self.keyword_rules = []

    def reload(self):
        self._load()

    def suggest(self, sender_email: str, sender_name: str, subjects: list) -> Suggestion:
        email_l = sender_email.lower()
        name_l  = sender_name.lower()

        # Layer 1: exact rule
        rule = self.rules.find(sender_email)
        if rule:
            return Suggestion(
                folder=rule.get("dest", ""),
                confidence=Confidence.HIGH,
                reason=f"Regel: {rule['pattern']}",
                action=rule["action"]
            )

        # Layer 2: sender keyword
        sender_text = f"{email_l} {name_l}"
        for kr in self.keyword_rules:
            if kr.get("field") != "sender":
                continue
            for kw in kr["keywords"]:
                if kw.lower() in sender_text:
                    return Suggestion(
                        folder=kr["folder"],
                        confidence=Confidence.MEDIUM,
                        reason=f'Absender: "{kw}"'
                    )

        # Layer 3: subject keyword majority-vote
        subj_lower = [s.lower() for s in subjects]
        folder_hits: dict[str, tuple[int, str]] = {}
        for kr in self.keyword_rules:
            if kr.get("field") != "subject":
                continue
            hits = 0
            first_kw = ""
            for kw in kr["keywords"]:
                kw_l = kw.lower()
                for s in subj_lower:
                    if kw_l in s:
                        hits += 1
                        if not first_kw:
                            first_kw = kw
            if hits > 0:
                folder = kr["folder"]
                if hits > folder_hits.get(folder, (0, ""))[0]:
                    folder_hits[folder] = (hits, first_kw)

        if folder_hits:
            best = max(folder_hits, key=lambda f: folder_hits[f][0])
            cnt, kw = folder_hits[best]
            return Suggestion(
                folder=best,
                confidence=Confidence.MEDIUM,
                reason=f'Betreff: "{kw}" ({cnt}×)'
            )

        return Suggestion(folder="", confidence=Confidence.UNKNOWN, reason="Kein Treffer")


# ─────────────────────────────────────────────
# Stats Tracker
# ─────────────────────────────────────────────

class StatsTracker:
    def __init__(self):
        self._data = {"folder_usage": {}, "sender_mail_counts": {}}
        self._load()

    def _load(self):
        if STATS_FILE.exists():
            try:
                with open(STATS_FILE, "r", encoding="utf-8") as f:
                    d = json.load(f)
                self._data["folder_usage"]      = d.get("folder_usage", {})
                self._data["sender_mail_counts"] = d.get("sender_mail_counts", {})
            except Exception:
                pass

    def _save(self):
        try:
            with open(STATS_FILE, "w", encoding="utf-8") as f:
                json.dump(self._data, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def record_move(self, sender_email: str, dest_folder: str, count: int):
        today = datetime.now().strftime("%Y-%m-%d")
        fu = self._data["folder_usage"]
        if dest_folder not in fu:
            fu[dest_folder] = {"count": 0, "last": today}
        fu[dest_folder]["count"] += count
        fu[dest_folder]["last"]   = today
        self._save()

    def record_sender_counts(self, groups: list):
        smc = self._data["sender_mail_counts"]
        for email, data in groups:
            smc[email.lower()] = len(data["items"])
        self._save()

    def neue_muster(self, rule_engine: RuleEngine, threshold: int = 3) -> list:
        """Senders with >= threshold mails and no rule."""
        smc = self._data["sender_mail_counts"]
        result = [(email, cnt) for email, cnt in smc.items()
                  if cnt >= threshold and not rule_engine.find(email)]
        return sorted(result, key=lambda x: x[1], reverse=True)

    def top_folders(self, n: int = 10) -> list:
        fu = self._data["folder_usage"]
        return sorted(fu.items(), key=lambda x: x[1]["count"], reverse=True)[:n]


# ─────────────────────────────────────────────
# New Category Analyzer
# ─────────────────────────────────────────────

class NewCategoryAnalyzer:
    def __init__(self, rule_engine: RuleEngine, suggestion_engine: SuggestionEngine):
        self.rules   = rule_engine
        self.suggest = suggestion_engine

    @staticmethod
    def _domain(email: str) -> str:
        try:
            host = email.split("@", 1)[1]
            parts = host.split(".")
            return ".".join(parts[-2:]) if len(parts) >= 2 else host
        except Exception:
            return email

    def analyze(self, groups: list, top_n: int = 20) -> list:
        domain_map: dict[str, dict] = {}
        for email, data in groups:
            if self.rules.find(email):
                continue
            subjects = []
            for item in data["items"][:20]:
                try:
                    s = getattr(item, "Subject", "") or ""
                    if s:
                        subjects.append(s)
                except Exception:
                    pass
            sug = self.suggest.suggest(email, data["name"], subjects)
            # Include both UNKNOWN and MEDIUM (MEDIUM = keyword suggestion but no rule yet)
            domain = self._domain(email)
            if domain not in domain_map:
                domain_map[domain] = {
                    "domain": domain,
                    "mail_count": 0,
                    "senders": [],
                    "subjects": [],
                    "suggestion": sug,
                }
            c = domain_map[domain]
            c["mail_count"] += len(data["items"])
            c["senders"].append(email)
            c["subjects"].extend(subjects[:5])
            # Use highest-confidence suggestion for domain
            if sug.confidence == Confidence.HIGH or \
               (sug.confidence == Confidence.MEDIUM and c["suggestion"].confidence == Confidence.UNKNOWN):
                c["suggestion"] = sug

        result = sorted(domain_map.values(), key=lambda x: x["mail_count"], reverse=True)
        return result[:top_n]

    def show_dialog(self, parent, groups: list, on_assign_callback):
        clusters = self.analyze(groups)
        if not clusters:
            messagebox.showinfo("Analyse", "Alle Absender haben bereits Regeln.", parent=parent)
            return

        win = tk.Toplevel(parent)
        win.title("Neue Kategorien analysieren")
        win.geometry("920x580")
        win.configure(bg="#1e1e2e")

        tk.Label(win,
                 text=f"Unklassifizierte Absender — {len(clusters)} Domains ohne Regel",
                 bg="#1e1e2e", fg="#89b4fa",
                 font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, padx=12, pady=(10, 2))
        tk.Label(win,
                 text="Wähle Cluster → Ordner zuweisen → Regel wird gespeichert",
                 bg="#1e1e2e", fg="#a6adc8",
                 font=("Segoe UI", 9, "italic")).pack(anchor=tk.W, padx=12, pady=(0, 6))

        cols = ("Domain", "Mails", "Vorschlag", "Absender / Beispiel-Betreff")
        frame = tk.Frame(win, bg="#1e1e2e")
        frame.pack(fill=tk.BOTH, expand=True, padx=12)

        tree = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        tree.heading("Domain",   text="Domain")
        tree.heading("Mails",    text="Mails")
        tree.heading("Vorschlag",text="Vorschlag")
        tree.heading("Absender / Beispiel-Betreff", text="Absender / Beispiel-Betreff")
        tree.column("Domain",    width=160)
        tree.column("Mails",     width=55, anchor=tk.CENTER)
        tree.column("Vorschlag", width=200)
        tree.column("Absender / Beispiel-Betreff", width=400)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(frame, orient=tk.VERTICAL,
                      command=tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        for c in clusters:
            sug = c["suggestion"]
            sug_txt = (f"{sug.folder} ({sug.reason})"
                       if sug.confidence != Confidence.UNKNOWN
                       else "—")
            detail = (", ".join(c["senders"][:2])
                      + (f" | {c['subjects'][0]}" if c["subjects"] else ""))
            tree.insert("", tk.END,
                        values=(c["domain"], c["mail_count"], sug_txt, detail[:80]))

        def _on_select(event):
            sel = tree.selection()
            if sel:
                idx = tree.index(sel[0])
                sug = clusters[idx]["suggestion"]
                if sug.confidence != Confidence.UNKNOWN:
                    dest_var.set(sug.folder)

        tree.bind("<<TreeviewSelect>>", _on_select)

        bbar = tk.Frame(win, bg="#1e1e2e", pady=8)
        bbar.pack(fill=tk.X, padx=12)

        dest_var = tk.StringVar()
        tk.Label(bbar, text="Zielordner:", bg="#1e1e2e",
                 fg="#a6adc8", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        ttk.Combobox(bbar, textvariable=dest_var, values=ALL_FOLDERS,
                     width=34, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(4, 8))

        def _assign():
            path = dest_var.get().strip()
            if not path:
                messagebox.showwarning("Kein Ordner", "Bitte Zielordner wählen.", parent=win)
                return
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Keine Auswahl",
                                        "Bitte Cluster in der Liste auswählen.", parent=win)
                return
            idx = tree.index(sel[0])
            on_assign_callback(clusters[idx], path)
            tree.delete(sel[0])
            clusters.pop(idx)

        tk.Button(bbar, text="Ordner zuweisen & Regel merken",
                  bg="#a6e3a1", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 9, "bold"), padx=8,
                  command=_assign).pack(side=tk.LEFT, padx=2)
        tk.Button(bbar, text="Schließen", command=win.destroy,
                  bg="#313244", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=8).pack(side=tk.LEFT, padx=4)


# ─────────────────────────────────────────────
# Outlook Bridge
# ─────────────────────────────────────────────

class OutlookBridge:
    def __init__(self):
        self.ol = None
        self.inbox = None
        self.trash = None
        self._folder_cache = {}

    def connect(self):
        self.ol = win32com.client.GetActiveObject("Outlook.Application")
        ns = self.ol.GetNamespace("MAPI")
        self.inbox = ns.GetDefaultFolder(6)
        self.trash = ns.GetDefaultFolder(3)
        self._folder_cache = {}

    def load_groups(self, year_from: int = None, year_to: int = None, max_items: int = 5000):
        """Load inbox items with optional date range and hard item cap to prevent RPC crashes."""
        q = chr(34)
        filters = []
        if year_from is not None:
            filters.append(
                f"{q}urn:schemas:httpmail:datereceived{q} >= '{year_from}-01-01 00:00:00'"
            )
        if year_to is not None:
            filters.append(
                f"{q}urn:schemas:httpmail:datereceived{q} <= '{year_to}-12-31 23:59:59'"
            )

        if filters:
            dasl = "@SQL=" + " AND ".join(filters)
            items = self.inbox.Items.Restrict(dasl)
        else:
            items = self.inbox.Items

        groups = {}
        count = 0
        try:
            for item in items:
                if count >= max_items:
                    break
                count += 1
                try:
                    email = (getattr(item, "SenderEmailAddress", "") or "").lower().strip()
                    name  = getattr(item, "SenderName", "") or email
                    if not email:
                        email = "(unbekannt)"
                    if email not in groups:
                        groups[email] = {"name": name, "items": []}
                    groups[email]["items"].append(item)
                except Exception:
                    continue
        except Exception as e:
            # RPC or COM error mid-iteration — return what we have so far
            if groups:
                pass  # partial result is fine
            else:
                raise
        return sorted(groups.items(), key=lambda x: len(x[1]["items"]), reverse=True), count

    def resolve_folder(self, path: str):
        if path in self._folder_cache:
            return self._folder_cache[path]
        cur = self.inbox
        for part in path.split("/"):
            cur = cur.Folders.Item(part)
        self._folder_cache[path] = cur
        return cur

    def move_items(self, items, dest_path: str):
        folder = self.resolve_folder(dest_path)
        count, errors = 0, 0
        for item in list(items):
            try:
                item.Move(folder)
                count += 1
            except Exception:
                errors += 1
        return count, errors

    def delete_items(self, items):
        count, errors = 0, 0
        for item in list(items):
            try:
                item.Move(self.trash)
                count += 1
            except Exception:
                errors += 1
        return count, errors

    def get_body(self, item) -> str:
        try:
            html = getattr(item, "HTMLBody", "")
            if html:
                return strip_html(html)
            return getattr(item, "Body", "") or "(kein Inhalt)"
        except Exception:
            return "(Inhalt nicht lesbar)"


# ─────────────────────────────────────────────
# Main GUI
# ─────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook Inbox Manager")
        self.geometry("1440x860")
        self.minsize(1100, 640)
        self.configure(bg="#1e1e2e")

        self.bridge   = OutlookBridge()
        self.rules    = RuleEngine()
        self.suggeng  = SuggestionEngine(self.rules)
        self.stats    = StatsTracker()
        self.analyzer = NewCategoryAnalyzer(self.rules, self.suggeng)

        self.groups             = []
        self.current_idx        = 0
        self.current_mail_items = []
        self.year_from_var      = tk.IntVar(value=2025)
        self.year_to_var        = tk.IntVar(value=2026)
        self.year_filter_var    = tk.BooleanVar(value=True)
        self.max_items_var      = tk.IntVar(value=2000)
        self.sort_mode          = tk.StringVar(value="count")  # "count" | "alpha"
        self._display_order     = []
        self._current_display_idx = 0

        self._build_ui()
        self._bind_keys()
        self._connect_and_load()

    # ── UI ────────────────────────────────────

    def _build_ui(self):
        # Toolbar
        bar = tk.Frame(self, bg="#181825", pady=7)
        bar.pack(fill=tk.X)

        tk.Label(bar, text="Outlook Inbox Manager", bg="#181825", fg="#cdd6f4",
                 font=("Segoe UI", 13, "bold")).pack(side=tk.LEFT, padx=12)

        # Date range filter
        tk.Checkbutton(bar, text="Von:", variable=self.year_filter_var,
                       bg="#181825", fg="#a6adc8", selectcolor="#313244",
                       activebackground="#181825",
                       font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(16, 0))
        tk.Spinbox(bar, from_=2010, to=2030, textvariable=self.year_from_var,
                   width=6, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(2, 0))
        tk.Label(bar, text="Bis:", bg="#181825", fg="#a6adc8",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(4, 0))
        tk.Spinbox(bar, from_=2010, to=2030, textvariable=self.year_to_var,
                   width=6, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(2, 0))
        tk.Label(bar, text="Max:", bg="#181825", fg="#a6adc8",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(6, 0))
        tk.Spinbox(bar, from_=500, to=10000, increment=500, textvariable=self.max_items_var,
                   width=6, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(2, 0))

        tk.Button(bar, text=" Laden ", command=self._reload,
                  bg="#313244", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=6).pack(side=tk.LEFT, padx=6)
        tk.Button(bar, text=" 🔌 Reconnect ", command=self._reconnect,
                  bg="#f38ba8", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=6).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text=" ⚡ Regeln anwenden ", command=self._apply_all_rules,
                  bg="#a6e3a1", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 9, "bold"), padx=6).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text=" Regeln ", command=self._open_rules_window,
                  bg="#89b4fa", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=6).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text=" Neue Kategorien ", command=self._open_analyzer,
                  bg="#cba6f7", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=6).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text=" Statistik ", command=self._open_stats,
                  bg="#89dceb", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=6).pack(side=tk.LEFT, padx=2)

        self.status_var = tk.StringVar(value="Verbinde…")
        tk.Label(bar, textvariable=self.status_var, bg="#181825", fg="#f38ba8",
                 font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=12)

        # Content
        content = tk.Frame(self, bg="#1e1e2e")
        content.pack(fill=tk.BOTH, expand=True)

        # Left: group list
        left = tk.Frame(content, bg="#181825", width=330)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0), pady=8)
        left.pack_propagate(False)

        tk.Label(left, text="Absender-Gruppen", bg="#181825", fg="#89b4fa",
                 font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, padx=8, pady=(6, 2))
        tk.Label(left, text="✓=Regel  ~=Vorschlag  ·=unbekannt",
                 bg="#181825", fg="#6c7086",
                 font=("Segoe UI", 8)).pack(anchor=tk.W, padx=8)

        sort_row = tk.Frame(left, bg="#181825")
        sort_row.pack(anchor=tk.W, padx=6, pady=(2, 0))
        tk.Label(sort_row, text="Sortierung:", bg="#181825", fg="#6c7086",
                 font=("Segoe UI", 8)).pack(side=tk.LEFT)
        for label, val in (("Anzahl", "count"), ("A–Z", "alpha"), ("Neueste", "recent")):
            tk.Radiobutton(sort_row, text=label, variable=self.sort_mode, value=val,
                           command=self._fill_group_list,
                           bg="#181825", fg="#a6adc8", selectcolor="#313244",
                           activebackground="#181825",
                           font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)

        self.group_lb = tk.Listbox(
            left, bg="#181825", fg="#cdd6f4",
            selectbackground="#313244", selectforeground="#cba6f7",
            font=("Segoe UI", 9), relief=tk.FLAT, activestyle="none",
            highlightthickness=0, exportselection=False
        )
        gsb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.group_lb.yview)
        self.group_lb.configure(yscrollcommand=gsb.set)
        self.group_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0), pady=4)
        gsb.pack(side=tk.LEFT, fill=tk.Y, pady=4)
        self.group_lb.bind("<<ListboxSelect>>", self._on_group_select)

        # Right
        right = tk.Frame(content, bg="#1e1e2e")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.sender_var = tk.StringVar(value="")
        tk.Label(right, textvariable=self.sender_var, bg="#1e1e2e", fg="#cba6f7",
                 font=("Segoe UI", 10, "bold"), anchor=tk.W).pack(fill=tk.X)

        self.hint_var = tk.StringVar(value="")
        self.hint_label = tk.Label(right, textvariable=self.hint_var, bg="#1e1e2e",
                                    fg="#a6e3a1", font=("Segoe UI", 9, "italic"), anchor=tk.W)
        self.hint_label.pack(fill=tk.X, pady=(0, 2))

        # Vertical paned: mail list + preview
        vpane = tk.PanedWindow(right, orient=tk.VERTICAL, bg="#1e1e2e",
                                sashwidth=6, sashrelief=tk.FLAT)
        vpane.pack(fill=tk.BOTH, expand=True)

        # Mail list
        mail_frame = tk.Frame(vpane, bg="#1e1e2e")
        sel_row = tk.Frame(mail_frame, bg="#1e1e2e")
        sel_row.pack(fill=tk.X, pady=(0, 2))
        self.sel_info_var = tk.StringVar(value="")
        tk.Label(sel_row, textvariable=self.sel_info_var, bg="#1e1e2e",
                 fg="#a6adc8", font=("Segoe UI", 8)).pack(side=tk.LEFT)
        tk.Button(sel_row, text="Alle", command=self._select_all_mails,
                  bg="#313244", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 8), padx=4, pady=1).pack(side=tk.LEFT, padx=4)
        tk.Button(sel_row, text="Keine", command=self._deselect_all_mails,
                  bg="#313244", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 8), padx=4, pady=1).pack(side=tk.LEFT)

        ml_inner = tk.Frame(mail_frame, bg="#1e1e2e")
        ml_inner.pack(fill=tk.BOTH, expand=True)
        self.mail_lb = tk.Listbox(
            ml_inner, bg="#181825", fg="#cdd6f4",
            selectbackground="#45475a", selectforeground="#f9e2af",
            font=("Consolas", 9), relief=tk.FLAT, activestyle="none",
            highlightthickness=0, exportselection=False,
            selectmode=tk.EXTENDED
        )
        msb = ttk.Scrollbar(ml_inner, orient=tk.VERTICAL, command=self.mail_lb.yview)
        self.mail_lb.configure(yscrollcommand=msb.set)
        self.mail_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        msb.pack(side=tk.RIGHT, fill=tk.Y)
        self.mail_lb.bind("<<ListboxSelect>>", self._on_mail_select)
        self.mail_lb.bind("<Double-Button-1>", self._on_mail_doubleclick)
        vpane.add(mail_frame, minsize=120)

        # Preview
        preview_frame = tk.Frame(vpane, bg="#1e1e2e")
        prev_hdr = tk.Frame(preview_frame, bg="#1e1e2e")
        prev_hdr.pack(fill=tk.X)
        self.preview_title_var = tk.StringVar(value="Vorschau")
        tk.Label(prev_hdr, textvariable=self.preview_title_var, bg="#1e1e2e",
                 fg="#89b4fa", font=("Segoe UI", 9, "bold"), anchor=tk.W).pack(side=tk.LEFT)
        self.preview_text = scrolledtext.ScrolledText(
            preview_frame, bg="#181825", fg="#cdd6f4",
            font=("Segoe UI", 9), relief=tk.FLAT, wrap=tk.WORD,
            state=tk.DISABLED, height=10, highlightthickness=0, padx=8, pady=6
        )
        self.preview_text.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        vpane.add(preview_frame, minsize=80)

        # Action bar
        abar = tk.Frame(right, bg="#1e1e2e", pady=4)
        abar.pack(fill=tk.X)
        tk.Label(abar, text="Zielordner:", bg="#1e1e2e", fg="#a6adc8",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.dest_var = tk.StringVar()
        ttk.Combobox(abar, textvariable=self.dest_var, values=ALL_FOLDERS,
                     width=36, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(4, 8))

        bs = {"relief": tk.FLAT, "font": ("Segoe UI", 9, "bold"), "padx": 8, "pady": 4}
        tk.Button(abar, text="[J] Auswahl verschieben", bg="#a6e3a1", fg="#1e1e2e",
                  command=self._do_move, **bs).pack(side=tk.LEFT, padx=2)
        tk.Button(abar, text="[L] Auswahl löschen", bg="#f38ba8", fg="#1e1e2e",
                  command=self._do_delete, **bs).pack(side=tk.LEFT, padx=2)
        tk.Button(abar, text="[N] Überspringen", bg="#fab387", fg="#1e1e2e",
                  command=self._do_skip, **bs).pack(side=tk.LEFT, padx=2)
        self.save_rule_var = tk.BooleanVar(value=True)
        tk.Checkbutton(abar, text="Regel merken", variable=self.save_rule_var,
                       bg="#1e1e2e", fg="#a6adc8", selectcolor="#313244",
                       activebackground="#1e1e2e",
                       font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=8)

        # Quick folders
        qbar = tk.Frame(right, bg="#1e1e2e", pady=2)
        qbar.pack(fill=tk.X)
        tk.Label(qbar, text="Schnell:", bg="#1e1e2e", fg="#a6adc8",
                 font=("Segoe UI", 8)).pack(side=tk.LEFT)
        for label, path in QUICK_FOLDERS:
            tk.Button(qbar, text=label, bg="#313244", fg="#cdd6f4",
                      relief=tk.FLAT, font=("Segoe UI", 8), padx=5, pady=2,
                      command=lambda p=path: self._quick_move(p)).pack(side=tk.LEFT, padx=1)

    def _bind_keys(self):
        for key in ("j", "J"):
            self.bind(f"<{key}>", lambda e: self._do_move())
        for key in ("l", "L"):
            self.bind(f"<{key}>", lambda e: self._do_delete())
        for key in ("n", "N"):
            self.bind(f"<{key}>", lambda e: self._do_skip())
        self.bind("<Down>",      lambda e: self._nav(+1))
        self.bind("<Up>",        lambda e: self._nav(-1))
        self.bind("<Escape>",    lambda e: self.focus_set())
        self.bind("<Control-a>", lambda e: self._select_all_mails())

    # ── Connect & Load ────────────────────────

    def _connect_and_load(self):
        try:
            self.bridge.connect()
            self.status_var.set("Verbunden ✓")
            self._reload()
        except Exception as e:
            self.status_var.set("❌ Nicht verbunden — Outlook starten & Reconnect drücken")
            messagebox.showerror("Verbindungsfehler",
                                  f"Outlook nicht gefunden.\nOutlook starten und dann "
                                  f"'Reconnect' drücken.\n\nDetail: {e}")

    def _reconnect(self):
        self.status_var.set("Verbinde…")
        self.update_idletasks()
        try:
            self.bridge.connect()
            self.status_var.set("Verbunden ✓")
            self._reload()
        except Exception as e:
            self.status_var.set("❌ Nicht verbunden — Outlook läuft?")
            messagebox.showerror("Reconnect fehlgeschlagen",
                                  f"Outlook nicht erreichbar.\n\nDetail: {e}")

    def _reload(self):
        self.status_var.set("Lade Posteingang…")
        self.update_idletasks()
        try:
            use_filter = self.year_filter_var.get()
            year_from  = self.year_from_var.get() if use_filter else None
            year_to    = self.year_to_var.get()   if use_filter else None
            max_items  = self.max_items_var.get()
            self.groups, loaded = self.bridge.load_groups(year_from, year_to, max_items)
            self.stats.record_sender_counts(self.groups)
            self._fill_group_list()
            total  = sum(len(g[1]["items"]) for g in self.groups)
            if use_filter:
                yr_txt = f"{year_from}–{year_to}"
            else:
                yr_txt = "alle Jahre"
            capped = f" (Max {max_items} erreicht — Zeitraum einengen!)" if loaded >= max_items else ""
            self.status_var.set(
                f"{total} Mails · {len(self.groups)} Absender · {yr_txt}{capped}"
            )
            if self.groups:
                self._select(0)
        except Exception as e:
            err = str(e)
            if "RPC" in err or "dispatch" in err.lower() or "com" in err.lower():
                self.status_var.set("❌ RPC-Fehler — Outlook neu starten & Reconnect drücken")
                messagebox.showerror("RPC-Fehler",
                                      "Outlook hat die Verbindung getrennt (RPC-Fehler).\n\n"
                                      "1. Outlook neu starten\n"
                                      "2. 'Reconnect' drücken\n"
                                      "3. Zeitraum einengen oder Max reduzieren\n\n"
                                      f"Detail: {e}")
            else:
                self.status_var.set("Ladefehler")
                messagebox.showerror("Ladefehler", str(e))

    def _fill_group_list(self):
        self.group_lb.delete(0, tk.END)
        mode = self.sort_mode.get()
        if mode == "alpha":
            display_groups = sorted(self.groups,
                                    key=lambda g: (g[1]["name"] or g[0]).lower())
        elif mode == "recent":
            def _newest(g):
                best = datetime.min
                for item in g[1]["items"]:
                    try:
                        rt = item.ReceivedTime
                        if rt > best:
                            best = rt
                    except Exception:
                        pass
                return best
            display_groups = sorted(self.groups, key=_newest, reverse=True)
        else:
            display_groups = self.groups  # already sorted by count from load_groups
        # Keep a mapping from display order → original groups index for _select()
        self._display_order = [self.groups.index(g) for g in display_groups]
        for email, data in display_groups:
            n = len(data["items"])
            subjects_sample = []
            for item in data["items"][:8]:
                try:
                    s = getattr(item, "Subject", "") or ""
                    if s:
                        subjects_sample.append(s)
                except Exception:
                    pass
            sug = self.suggeng.suggest(email, data["name"], subjects_sample)

            if sug.confidence == Confidence.HIGH:
                marker, conf_color = "✓ ", "#a6e3a1"
            elif sug.confidence == Confidence.MEDIUM:
                marker, conf_color = "~ ", "#f9e2af"
            else:
                marker, conf_color = "· ", "#6c7086"

            name = data["name"] if data["name"] != email else email
            self.group_lb.insert(tk.END, f"{marker}[{n:3d}]  {name}")
            idx = self.group_lb.size() - 1
            # Volume color for many mails, else confidence color
            if n >= 10:
                self.group_lb.itemconfig(idx, fg="#f38ba8")
            elif n >= 5:
                self.group_lb.itemconfig(idx, fg="#fab387")
            else:
                self.group_lb.itemconfig(idx, fg=conf_color)

    # ── Group Selection ───────────────────────

    def _on_group_select(self, _event):
        sel = self.group_lb.curselection()
        if sel:
            self._select(sel[0])

    def _select(self, display_idx: int):
        if not self.groups:
            return
        display_idx = max(0, min(display_idx, len(self.groups) - 1))
        self.current_idx = getattr(self, "_display_order", list(range(len(self.groups))))[display_idx]
        self._current_display_idx = display_idx
        self.group_lb.selection_clear(0, tk.END)
        self.group_lb.selection_set(display_idx)
        self.group_lb.see(display_idx)

        email, data = self.groups[self.current_idx]
        n = len(data["items"])
        self.sender_var.set(f"{data['name']}  <{email}>  — {n} Mail(s)")

        self.current_mail_items = sorted(
            data["items"],
            key=lambda m: getattr(m, "ReceivedTime", datetime.min),
            reverse=True
        )

        self.mail_lb.delete(0, tk.END)
        for item in self.current_mail_items:
            try:
                rt = item.ReceivedTime
                ds = f"{rt.year}-{rt.month:02d}-{rt.day:02d}"
                subj = (getattr(item, "Subject", "") or "(kein Betreff)")[:95]
                self.mail_lb.insert(tk.END, f"  {ds}  {subj}")
            except Exception:
                self.mail_lb.insert(tk.END, "  (Fehler)")

        # Suggestion
        subjects = [getattr(it, "Subject", "") or "" for it in self.current_mail_items]
        sug = self.suggeng.suggest(email, data["name"], subjects)

        if sug.confidence == Confidence.HIGH:
            hint = f"✓ Regel: {sug.reason}"
            if sug.action == "move":
                self.dest_var.set(sug.folder)
            self.hint_label.config(fg="#a6e3a1")
        elif sug.confidence == Confidence.MEDIUM:
            hint = f"~ Vorschlag ({sug.reason}): {sug.folder}"
            self.dest_var.set(sug.folder)
            self.hint_label.config(fg="#f9e2af")
        else:
            hint = "· Kein Vorschlag — unbekannter Absender"
            self.hint_label.config(fg="#6c7086")

        self.hint_var.set(hint)
        self._clear_preview()
        self.sel_info_var.set(f"0 von {n} ausgewählt  (Strg+Klick = mehrere, Strg+A = alle)")
        self.focus_set()

    def _nav(self, delta: int):
        self._select(self._current_display_idx + delta)

    # ── Mail Preview & Selection ──────────────

    def _on_mail_select(self, _event):
        sel = self.mail_lb.curselection()
        n_total = len(self.current_mail_items)
        n_sel   = len(sel)
        hint = f"{n_sel} von {n_total} ausgewählt"
        if n_sel == 0:
            hint += "  (Strg+Klick = mehrere, Strg+A = alle)"
        self.sel_info_var.set(hint)
        if sel and self.current_mail_items:
            last = sel[-1]
            if last < len(self.current_mail_items):
                self._show_preview(self.current_mail_items[last])

    def _on_mail_doubleclick(self, _event):
        idx = self.mail_lb.nearest(_event.y)
        if 0 <= idx < len(self.current_mail_items):
            try:
                self.current_mail_items[idx].Display()
            except Exception as e:
                messagebox.showerror("Fehler", f"Mail konnte nicht geöffnet werden:\n{e}")

    def _show_preview(self, item):
        try:
            rt   = item.ReceivedTime
            subj = getattr(item, "Subject", "") or "(kein Betreff)"
            name = getattr(item, "SenderName", "") or ""
            body = self.bridge.get_body(item)
            self.preview_title_var.set(
                f"{rt.year}-{rt.month:02d}-{rt.day:02d}  |  {subj[:70]}"
            )
            self.preview_text.configure(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, f"Von: {name}\nBetreff: {subj}\n\n{body}")
            self.preview_text.configure(state=tk.DISABLED)
        except Exception:
            self._clear_preview()

    def _clear_preview(self):
        self.preview_title_var.set("Vorschau")
        self.preview_text.configure(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.configure(state=tk.DISABLED)

    def _select_all_mails(self):
        self.mail_lb.selection_set(0, tk.END)
        self._on_mail_select(None)

    def _deselect_all_mails(self):
        self.mail_lb.selection_clear(0, tk.END)
        self.sel_info_var.set(f"0 von {len(self.current_mail_items)} ausgewählt")

    # ── Action Helpers ────────────────────────

    def _get_target_items(self):
        sel = self.mail_lb.curselection()
        if sel:
            items = [self.current_mail_items[i]
                     for i in sel if i < len(self.current_mail_items)]
            return items, len(items) < len(self.current_mail_items)
        return list(self.current_mail_items), False

    def _remove_items_from_group(self, done_items):
        email, data = self.groups[self.current_idx]
        done_ids = {id(x) for x in done_items}
        data["items"] = [it for it in data["items"] if id(it) not in done_ids]
        if not data["items"]:
            self._remove_current_group()
        else:
            self._select(self.current_idx)
            n = len(data["items"])
            rule = self.rules.find(email)
            marker = "✓ " if rule else "~ " if self.suggeng.suggest(
                email, data["name"], []).confidence == Confidence.MEDIUM else "· "
            name = data["name"] if data["name"] != email else email
            self.group_lb.delete(self.current_idx)
            self.group_lb.insert(self.current_idx, f"{marker}[{n:3d}]  {name}")
            color = "#f38ba8" if n >= 10 else "#fab387" if n >= 5 else "#cdd6f4"
            self.group_lb.itemconfig(self.current_idx, fg=color)
            self.group_lb.selection_set(self.current_idx)
        total = sum(len(g[1]["items"]) for g in self.groups)
        self.status_var.set(f"{total} Mails · {len(self.groups)} Absender verbleibend")

    # ── Actions ───────────────────────────────

    def _current(self):
        if not self.groups or self.current_idx >= len(self.groups):
            return None, None
        return self.groups[self.current_idx]

    def _do_move(self):
        email, data = self._current()
        if data is None:
            return
        dest = self.dest_var.get().strip()
        if not dest:
            messagebox.showwarning("Kein Ordner", "Bitte Zielordner wählen.")
            return
        target, is_partial = self._get_target_items()
        if not target:
            return
        try:
            moved, errors = self.bridge.move_items(target, dest)
            self.stats.record_move(email, dest, moved)
            self.status_var.set(f"✓ {moved} Mails → {dest}")
            if self.save_rule_var.get() and not is_partial:
                self.rules.add(email, "move", dest)
                self._fill_group_list()
            self._remove_items_from_group(target)
            if errors:
                messagebox.showwarning("Teilfehler", f"{errors} Mail(s) nicht verschoben.")
        except Exception as e:
            messagebox.showerror("Fehler", str(e))

    def _do_delete(self):
        email, data = self._current()
        if data is None:
            return
        target, is_partial = self._get_target_items()
        if not target:
            return
        n     = len(target)
        label = "ausgewählte" if is_partial else "alle"
        if not messagebox.askyesno("Löschen?",
                                    f"{n} {label} Mail(s) von '{email}' löschen?"):
            return
        deleted, errors = self.bridge.delete_items(target)
        self.status_var.set(f"🗑 {deleted} Mails gelöscht")
        if self.save_rule_var.get() and not is_partial:
            self.rules.add(email, "delete")
        self._remove_items_from_group(target)
        if errors:
            messagebox.showwarning("Teilfehler", f"{errors} Mail(s) nicht gelöscht.")

    def _do_skip(self):
        email, data = self._current()
        if data is None:
            return
        if self.save_rule_var.get():
            self.rules.add(email, "skip")
        self._nav(+1)

    def _quick_move(self, path: str):
        self.dest_var.set(path)
        self._do_move()

    def _remove_current_group(self):
        display_idx = self._current_display_idx
        self.groups.pop(self.current_idx)
        self._fill_group_list()   # rebuild display order after removal
        self.current_mail_items = []
        self._clear_preview()
        if self.groups:
            self._select(min(display_idx, len(self.groups) - 1))
        else:
            self.sender_var.set("Posteingang aufgeräumt!")
            self.hint_var.set("")
            self.mail_lb.delete(0, tk.END)
            self.sel_info_var.set("")

    # ── Apply All Rules ───────────────────────

    def _apply_all_rules(self):
        if not self.rules.rules:
            messagebox.showinfo("Keine Regeln", "Noch keine Regeln gespeichert.")
            return
        if not messagebox.askyesno("Alle Regeln anwenden",
                                    f"{len(self.rules.rules)} Regeln anwenden?"):
            return
        self.status_var.set("Wende Regeln an…")
        self.update_idletasks()
        moved_total = deleted_total = errors_total = unmatched = 0
        rpc_error = False
        for email, data in list(self.groups):
            rule = self.rules.find(email)
            if not rule:
                unmatched += 1
                continue
            try:
                if rule["action"] == "move" and rule["dest"]:
                    m, e = self.bridge.move_items(data["items"], rule["dest"])
                    self.stats.record_move(email, rule["dest"], m)
                    moved_total  += m
                    errors_total += e
                elif rule["action"] == "delete":
                    d, e = self.bridge.delete_items(data["items"])
                    deleted_total += d
                    errors_total  += e
            except Exception as ex:
                err = str(ex)
                if "RPC" in err or "dispatch" in err.lower():
                    rpc_error = True
                    break
                errors_total += len(data["items"])
        if rpc_error:
            messagebox.showwarning("RPC-Fehler während Regeln anwenden",
                                    "Verbindung zu Outlook wurde unterbrochen.\n"
                                    "Outlook neu starten → Reconnect → erneut versuchen.")
            self.status_var.set("❌ RPC-Fehler — Reconnect drücken")
            return
        self._reload()
        messagebox.showinfo("Fertig",
                             f"Verschoben:  {moved_total}\n"
                             f"Gelöscht:    {deleted_total}\n"
                             f"Ohne Regel:  {unmatched}\n"
                             f"Fehler:      {errors_total}")

    # ── New Category Analyzer ─────────────────

    def _open_analyzer(self):
        def on_assign(cluster: dict, folder_path: str):
            domain = cluster["domain"]
            # Add domain as sender keyword rule
            self.suggeng.keyword_rules.append({
                "keywords": [domain],
                "folder": folder_path,
                "field": "sender"
            })
            with open(KEYWORDS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.suggeng.keyword_rules, f, indent=2, ensure_ascii=False)
            # Add exact rules for all senders in cluster
            for sender_email in cluster["senders"]:
                self.rules.add(sender_email, "move", folder_path)
            self._fill_group_list()
            self.status_var.set(f"Regel für @{domain} → {folder_path} gespeichert")

        self.analyzer.show_dialog(self, self.groups, on_assign)

    # ── Statistics ────────────────────────────

    def _open_stats(self):
        win = tk.Toplevel(self)
        win.title("Statistik & Neue Muster")
        win.geometry("680x520")
        win.configure(bg="#1e1e2e")

        tk.Label(win, text="Statistik", bg="#1e1e2e", fg="#89b4fa",
                 font=("Segoe UI", 11, "bold")).pack(anchor=tk.W, padx=12, pady=(10, 4))

        txt = scrolledtext.ScrolledText(
            win, bg="#181825", fg="#cdd6f4",
            font=("Consolas", 9), relief=tk.FLAT, wrap=tk.WORD,
            highlightthickness=0, padx=8, pady=6
        )
        txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        lines = []
        top = self.stats.top_folders(12)
        if top:
            lines.append("Top Zielordner:\n")
            for folder, data in top:
                lines.append(f"  {data['count']:4d}×  {folder}")
        neue = self.stats.neue_muster(self.rules, threshold=3)
        if neue:
            lines.append(f"\n\nNeue Muster (≥3 Mails, keine Regel) — {len(neue)} Absender:\n")
            for email, count in neue[:20]:
                lines.append(f"  {count:4d}×  {email}")
        lines.append(f"\n\nGespeicherte Regeln: {len(self.rules.rules)}")
        lines.append(f"Keyword-Regeln:      {len(self.suggeng.keyword_rules)}")

        txt.insert(tk.END, "\n".join(lines) if lines else "Noch keine Daten.")
        txt.configure(state=tk.DISABLED)

        tk.Button(win, text="Schließen", command=win.destroy,
                  bg="#313244", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=8).pack(pady=6)

    # ── Rules Window ──────────────────────────

    def _open_rules_window(self):
        win = tk.Toplevel(self)
        win.title("Regeln verwalten")
        win.geometry("760x480")
        win.configure(bg="#1e1e2e")

        nb = ttk.Notebook(win)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Tab 1: sender rules
        tab1 = tk.Frame(nb, bg="#1e1e2e")
        nb.add(tab1, text=f"Absender-Regeln ({len(self.rules.rules)})")
        self._build_rules_tab(tab1, self.rules.rules,
                               ("Muster", "Aktion", "Zielordner"),
                               lambda idx: self.rules.delete(idx) or self.rules.save())

        # Tab 2: keyword rules
        tab2 = tk.Frame(nb, bg="#1e1e2e")
        nb.add(tab2, text=f"Keyword-Regeln ({len(self.suggeng.keyword_rules)})")
        self._build_kw_tab(tab2)

        tk.Button(win, text="Schließen", command=win.destroy,
                  bg="#313244", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=8).pack(pady=6)

    def _build_rules_tab(self, parent, rules_list, cols, delete_fn):
        frame = tk.Frame(parent, bg="#1e1e2e")
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=16)
        for col in cols:
            tree.heading(col, text=col)
        tree.column(cols[0], width=260)
        tree.column(cols[1], width=80)
        tree.column(cols[2], width=320)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(frame, orient=tk.VERTICAL,
                      command=tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        def refresh():
            tree.delete(*tree.get_children())
            for r in rules_list:
                tree.insert("", tk.END,
                             values=(r.get("pattern", ""), r.get("action", ""),
                                     r.get("dest", "")))
        refresh()

        def delete_sel():
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            pat = rules_list[idx].get("pattern", "")
            if messagebox.askyesno("Löschen?", f"Regel für '{pat}' entfernen?",
                                    parent=tree.winfo_toplevel()):
                delete_fn(idx)
                refresh()

        tk.Button(parent, text="Ausgewählte löschen", command=delete_sel,
                  bg="#f38ba8", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 9), padx=8).pack(pady=4)

    def _build_kw_tab(self, parent):
        frame = tk.Frame(parent, bg="#1e1e2e")
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        cols = ("Schlüsselwörter", "Feld", "Zielordner")
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=16)
        tree.heading("Schlüsselwörter", text="Schlüsselwörter")
        tree.heading("Feld", text="Feld")
        tree.heading("Zielordner", text="Zielordner")
        tree.column("Schlüsselwörter", width=300)
        tree.column("Feld", width=70)
        tree.column("Zielordner", width=280)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(frame, orient=tk.VERTICAL,
                      command=tree.yview).pack(side=tk.RIGHT, fill=tk.Y)
        for kr in self.suggeng.keyword_rules:
            tree.insert("", tk.END, values=(
                ", ".join(kr.get("keywords", [])),
                kr.get("field", ""),
                kr.get("folder", "")
            ))
        tk.Label(parent,
                 text=f"Bearbeite {KEYWORDS_FILE.name} zum Hinzufügen neuer Regeln",
                 bg="#1e1e2e", fg="#6c7086",
                 font=("Segoe UI", 8, "italic")).pack(pady=4)


if __name__ == "__main__":
    app = App()
    app.mainloop()
