# ============================================================
# AdmitGuide - AI admission consultant web app
# ============================================================
# HOW TO RUN (apne computer par):
#   1. pip install -r requirements.txt
#   2. python app.py
#   3. Browser me kholen: http://127.0.0.1:5000
#
# NOTE: Ye app baghair kisi AI API key ke bhi chalti hai.
# Agar AI_API_KEY neeche laga denge to chat asal AI se
# baat karega, warna built-in consultant brain jawab dega.
# ============================================================

# ---------- SETTINGS (sirf ye lines badalni hain) ----------
AI_API_KEY = ""   # <-- yahan apni AI API key paste karen (khali = built-in brain)
AI_API_BASE = "https://api.openai.com/v1"  # OpenAI jaisi koi bhi API
AI_MODEL = "gpt-4o-mini"

ADMIN_PASSWORD = "admin123"   # /admin page ka password
# ------------------------------------------------------------

import re
import uuid
import openpyxl
from flask import Flask, request, jsonify, render_template, session, redirect

app = Flask(__name__)
# Har restart par naya secret - demo ke liye theek hai
app.secret_key = "admitguide-" + uuid.uuid4().hex

# ---------- University database: Excel se ek dafa load ----------
COLUMNS = ["Country", "University", "City", "World Rank",
           "Tuition/Year (USD)", "Scholarship", "IELTS Required",
           "Intake Deadlines", "Popular Programs", "Language",
           "Visa Notes", "Living Cost/Year (USD)", "Official Website",
           "Admissions Email", "Admissions Phone", "Intl Office Page"]

wb = openpyxl.load_workbook("universities.xlsx")
ws = wb["Universities"]

UNIVERSITIES = []
for row in ws.iter_rows(min_row=2, values_only=True):
    u = dict(zip(COLUMNS, row))
    u["University"] = (u["University"] or "").strip()
    if u["University"]:          # khali rows hata do
        UNIVERSITIES.append(u)

print("AdmitGuide ready: %d universities loaded." % len(UNIVERSITIES))


def num(x):
    """Kisi bhi value ko number banao, na ban sake to None."""
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def money_fmt(x):
    """12345 -> $12,345"""
    n = num(x)
    return "$%s" % format(int(n), ",d") if n is not None else "N/A"


# ---------- Har country ke documents (spec se) ----------
CHECKLISTS = {
    "USA": ["Academic transcripts", "Passport", "Statement of Purpose (SOP)",
            "2 Letters of Recommendation", "Financial proof (bank statement)",
            "IELTS / TOEFL score"],
    "UK": ["Academic transcripts", "Passport", "Statement of Purpose (SOP)",
           "1-2 Letters of Recommendation", "CAS documents from university",
           "Financial proof (28-day bank statement)"],
    "Germany": ["Academic transcripts", "Passport",
                "Blocked account proof (~11,904 EUR/year)",
                "IELTS / German language certificate", "Uni-assist application"],
    "Italy": ["Academic transcripts", "Passport",
              "Pre-enrollment via Universitaly portal",
              "Financial proof", "Declaration of Value (DOV)"],
    "Japan": ["Academic transcripts", "Passport",
              "Certificate of Eligibility (COE) documents",
              "Financial proof", "MEXT scholarship form (agar apply karna ho)"],
    "China": ["Academic transcripts", "Passport", "JW201 form",
              "Physical examination record", "Financial proof",
              "CSC scholarship form (agar apply karna ho)"],
    "Russia": ["Academic transcripts", "Passport",
               "Invitation letter from university",
               "Medical certificates (HIV test samet)",
               "Government quota application (agar apply karna ho)"],
}

COUNTRY_ALIASES = {
    "usa": "USA", "america": "USA", "united states": "USA",
    "uk": "UK", "britain": "UK", "england": "UK", "united kingdom": "UK",
    "germany": "Germany",
    "italy": "Italy",
    "japan": "Japan",
    "china": "China",
    "russia": "Russia",
}
ALL_COUNTRIES = ["USA", "UK", "Germany", "Italy", "Japan", "China", "Russia"]

# ---------- User ke jawabon se number / country nikalna ----------

def parse_money(text):
    """'10000', '10k', '$12,000' -> 10000.0"""
    t = (text or "").lower().replace(",", "").replace("$", "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*k\b", t)
    if m:
        return float(m.group(1)) * 1000
    m = re.search(r"\d+(?:\.\d+)?", t)
    return float(m.group(0)) if m else None


def parse_ielts(text):
    """'6.5', 'band 7' -> 6.5 ; 'nahi diya' / 'no' -> None"""
    t = (text or "").lower()
    if any(w in t for w in ["nahi", "no ", "not yet", "abhi nahi", "none"]):
        return None
    m = re.search(r"\d(?:\.\d)?", t)
    v = float(m.group(0)) if m else None
    return v if v and 1 <= v <= 9 else None


def parse_countries(text):
    """'USA aur Germany' -> ['USA', 'Germany'] ; 'sab' -> saare 7"""
    t = (text or "").lower()
    if "sab" in t or "all" in t:
        return ALL_COUNTRIES[:]
    found = []
    for alias, canon in COUNTRY_ALIASES.items():
        if alias in t and canon not in found:
            found.append(canon)
    return found


def find_university(name):
    """Naam se university dhoondo (thori spelling ghalat bhi chale gi)."""
    t = (name or "").lower().strip()
    if not t:
        return None
    for u in UNIVERSITIES:                       # pehle exact jaisa
        if t == u["University"].lower():
            return u
    for u in UNIVERSITIES:                       # phir hissa milna
        if t in u["University"].lower() or u["University"].lower() in t:
            return u
    words = [w for w in t.split() if len(w) > 3]
    for u in UNIVERSITIES:                       # phir lafz milna
        ul = u["University"].lower()
        if words and all(w in ul for w in words):
            return u
    return None


# ---------- University matching (spec ke mutabiq) ----------

def match_universities(profile):
    """
    profile = {"marks":..., "ielts": float|None, "budget": float|None,
               "field": str, "countries": [..]}
    Returns: {"dream": [...], "match": [...], "safe": [...]}  (max 2+4+4)
    """
    countries = profile.get("countries") or ALL_COUNTRIES
    budget = profile.get("budget")
    ielts = profile.get("ielts")
    field = (profile.get("field") or "").lower().strip()

    def ok(u):
        if u["Country"] not in countries:
            return False
        t = num(u["Tuition/Year (USD)"])
        if budget and t is not None and t > budget:
            return False
        req = num(u["IELTS Required"])
        if ielts and req is not None and req > ielts:
            return False
        return True

    pool = [u for u in UNIVERSITIES if ok(u)]

    # Field ka keyword Popular Programs me mile to behtar
    if field:
        fwords = [w for w in re.findall(r"[a-z]+", field) if len(w) > 2]
        fpool = [u for u in pool
                 if any(w in (u["Popular Programs"] or "").lower() for w in fwords)]
        if fpool:          # agar field se kuch mila to wahi rakho, warna sab
            pool = fpool

    def rank_key(u):
        r = num(u["World Rank"])
        return r if r else 99999

    pool.sort(key=rank_key)
    dream = [u for u in pool if (num(u["World Rank"]) or 99999) <= 50][:2]
    match = [u for u in pool if 50 < (num(u["World Rank"]) or 99999) <= 200][:4]
    safe = [u for u in pool if (num(u["World Rank"]) or 99999) > 200][:4]
    return {"dream": dream, "match": match, "safe": safe}


def card(u, why):
    """Chat me dikhane wale card ka JSON."""
    return {"name": u["University"], "country": u["Country"],
            "city": u["City"] or "-", "rank": u["World Rank"] or "N/A",
            "tuition": money_fmt(u["Tuition/Year (USD)"]),
            "scholarship": u["Scholarship"] or "-",
            "ielts": u["IELTS Required"] or "-",
            "deadlines": u["Intake Deadlines"] or "-",
            "website": u["Official Website"] or "",
            "email": u["Admissions Email"] or "",
            "phone": u["Admissions Phone"] or "",
            "why": why}


def gap_line(u, profile):
    """Har shortlisted university ke liye kya kami hai - ek line."""
    gaps = []
    req = num(u["IELTS Required"])
    mine = profile.get("ielts")
    if req and mine and req > mine:
        gaps.append("IELTS %s chahiye, apke paas %s hai" %
                    (u["IELTS Required"], mine))
    if not mine:
        gaps.append("IELTS abhi nahi diya - %s darkar hai" % u["IELTS Required"])
    if not gaps:
        return "✅ %s: IELTS requirement poori hai." % u["University"]
    return "⚠️ %s: %s." % (u["University"], "; ".join(gaps))


# ---------- Built-in consultant brain (baghair AI key ke) ----------
# Har visitor ka alag session: sawal-jawab yahan yaad rehta hai.
SESSIONS = {}

QUESTIONS = [
    ("marks",
     "Assalam-o-Alaikum! 👋 Main **AdmitGuide** hun, apka AI admission consultant.\n\n"
     "Pehle apki profile banate hain taake best universities shortlist kar sakun.\n\n"
     "**Sawal 1/5:** Apke academic marks kya hain? (jaise: FSc 85%, ya GPA 3.4)"),
    ("ielts",
     "**Sawal 2/5:** IELTS ya TOEFL ka score kya hai?\n(Agar abhi test nahi diya to bas **'nahi diya'** likhen)"),
    ("budget",
     "**Sawal 3/5:** Saal ka kul budget kitna hai USD me?\n(Tuition + rehne ka kharcha mila kar — jaise: 10000)"),
    ("field",
     "**Sawal 4/5:** Kis field me parhna chahte hain?\n(jaise: Computer Science, Business, Medicine)"),
    ("countries",
     "**Sawal 5/5:** Kin countries me jana chahenge?\nUSA, UK, Germany, Italy, Japan, China, Russia — "
     "ek se zyada likh sakte hain, ya **'sab'** likhen"),
]


def brain(session_id, text):
    """Rule-based consultant: profile -> shortlist -> gap -> checklists."""
    s = SESSIONS.setdefault(session_id,
                            {"state": "q0", "profile": {}, "shortlist": []})
    t = (text or "").strip()

    # ---- pehli dafa chat khula: seedha pehla sawal ----
    if t == "START":
        s["state"] = "q0"
        return {"reply": QUESTIONS[0][1], "buttons": []}

    # ---- command buttons ----
    if t == "CMD:RESTART":
        SESSIONS[session_id] = {"state": "q0", "profile": {}, "shortlist": []}
        return {"reply": QUESTIONS[0][1], "buttons": []}
    if t == "CMD:DEADLINES":
        lines = ["⏰ **Intake Deadlines** (shortlisted universities):\n"]
        for u in s["shortlist"]:
            lines.append("• **%s** (%s): %s" %
                         (u["University"], u["Country"], u["Intake Deadlines"] or "-"))
        lines.append("\n⚠️ Jo deadline sab se qareeb hai, us par **sab se pehle** apply karen!")
        return {"reply": "\n".join(lines), "buttons": after_buttons()}
    if t == "CMD:VISA":
        seen, lines = set(), ["🛂 **Visa notes** (apki shortlist ke countries):\n"]
        for u in s["shortlist"]:
            if u["Country"] not in seen:
                seen.add(u["Country"])
                lines.append("• **%s:** %s" % (u["Country"], u["Visa Notes"] or "-"))
        lines.append("\nMock visa interview ke liye bas poochhen: **'mera interview lo'** 🙂")
        return {"reply": "\n".join(lines), "buttons": after_buttons()}
    if t == "CMD:CHECKLIST":
        s["state"] = "checklist_ask"
        return {"reply": "Kis university ka **document checklist** chahiye? Naam likhen (jaise: TU Munich)."}
    if t == "CMD:SOP":
        s["state"] = "sop_ask"
        return {"reply": "✍️ **SOP helper:** Apni kahani mukhtasir batayen — background, goals, "
                         "aur ye field/country **kyun** chuna? (2-3 lines kaafi hain)"}
    if "mera interview lo" in t.lower() and s["shortlist"]:
        c = s["shortlist"][0]["Country"]
        return {"reply": "🎤 **Mock visa interview** (%s):\n\n"
                         "**Sawal 1:** Yehi country kyun chuna, koi aur kyun nahi?\n\n"
                         "Jawab likhen, main agla sawal poochhunga!" % c,
                "buttons": after_buttons()}

    state = s["state"]
    p = s["profile"]

    # ---- checklist ka naam aaya ----
    if state == "checklist_ask":
        u = find_university(t)
        if not u:
            return {"reply": "Ye naam nahi mila. Dobara likhen (jaise: TU Munich)."}
        items = CHECKLISTS.get(u["Country"], [])
        lines = ["📋 **%s** — documents:\n" % u["University"]]
        lines += ["%d. %s" % (i + 1, x) for i, x in enumerate(items)]
        return {"reply": "\n".join(lines), "buttons": after_buttons()}

    # ---- SOP draft ----
    if state == "sop_ask":
        p["story"] = t
        s["state"] = "shortlist"
        field = p.get("field", "mera field")
        countries = ", ".join(p.get("countries", []))
        draft = ("✍️ **Apka SOP draft** (isay apne lafzon me aur behtar karen):\n\n"
                 "\"Mera naam ek mehnati talib-e-ilm hun jise %s me gehra dilchaspi hai. "
                 "%s Mera academic record (%s) aur mustaqbil ke goals mujhe %s me "
                 "taleem ke liye tayyar karte hain. Mujhe yaqeen hai ke wahan ki taleem "
                 "se main apne khandan aur mulk ka naam roshan karunga.\"\n\n"
                 "Tip: har university ke liye 1-2 lines uske programs ke baare me "
                 "khud izafa karen — copy-paste SOP kamzor lagta hai!" %
                 (field, t, p.get("marks", "-"), countries))
        return {"reply": draft, "buttons": after_buttons()}

    # ---- profile ke 5 sawal, ek ek karke ----
    if state.startswith("q"):
        idx = int(state[1:])
        key = QUESTIONS[idx][0]
        if key == "ielts":
            p[key] = parse_ielts(t)
        elif key == "budget":
            p[key] = parse_money(t)
        elif key == "countries":
            p[key] = parse_countries(t) or ALL_COUNTRIES
        else:
            p[key] = t
        idx += 1
        if idx < len(QUESTIONS):                    # agla sawal
            s["state"] = "q%d" % idx
            return {"reply": QUESTIONS[idx][1]}
        # ---- 5on mukammal: shortlist banao ----
        res = match_universities(p)
        cards, flat = [], []
        for grp, label in [("dream", "🌟 Dream"), ("match", "🎯 Match"), ("safe", "🛡️ Safe")]:
            for u in res[grp]:
                flat.append(u)
                why = ("%s | Rank #%s | Tuition %s/saal" %
                       (label, u["World Rank"] or "N/A",
                        money_fmt(u["Tuition/Year (USD)"])))
                if p.get("budget"):
                    why += " — apke budget me ✅"
                cards.append(card(u, why))
        s["shortlist"] = flat
        s["state"] = "shortlist"
        if not flat:
            return {"reply": "😔 Maazrat! Apke filters par koi university poori nahi utri. "
                             "**'dobara shuru karo'** likhen aur budget ya countries wasee karen.",
                     "buttons": [{"label": "🔁 Dobara shuru karo", "payload": "CMD:RESTART"}]}
        lines = ["🎉 **Apki shortlist tayyar hai!** (%d universities)\n" % len(flat)]
        lines += [gap_line(u, p) for u in flat]
        lines.append("\nNeeche har university ka card hai — **Official Website** par click "
                     "karen to seedha university ki site khulegi. 👇")
        return {"reply": "\n".join(lines), "cards": cards,
                "buttons": after_buttons()}

    # ---- shortlist ke baad free text ----
    if state == "shortlist":
        u = find_university(t.replace("checklist", "").strip())
        if u:
            items = CHECKLISTS.get(u["Country"], [])
            lines = ["📋 **%s** — documents:\n" % u["University"]]
            lines += ["%d. %s" % (i + 1, x) for i, x in enumerate(items)]
            return {"reply": "\n".join(lines), "buttons": after_buttons()}
        return {"reply": "Samajh nahi aya — neeche options me se chunen 👇",
                "buttons": after_buttons()}

    return {"reply": QUESTIONS[0][1]}


def after_buttons():
    return [{"label": "📋 Documents", "payload": "CMD:CHECKLIST"},
            {"label": "⏰ Deadlines", "payload": "CMD:DEADLINES"},
            {"label": "🛂 Visa info", "payload": "CMD:VISA"},
            {"label": "✍️ SOP help", "payload": "CMD:SOP"},
            {"label": "🔁 Nayi profile", "payload": "CMD:RESTART"}]


# ---------- AI API mode (key lagi ho to) ----------
_agent_prompt_cache = None

def agent_prompt():
    global _agent_prompt_cache
    if _agent_prompt_cache is None:
        with open("agent_prompt.txt", encoding="utf-8") as f:
            _agent_prompt_cache = f.read()
    return _agent_prompt_cache


def ai_reply(history):
    """AI_API_KEY lagi ho to asal AI se jawab lao."""
    import requests
    r = requests.post(
        AI_API_BASE.rstrip("/") + "/chat/completions",
        headers={"Authorization": "Bearer " + AI_API_KEY,
                 "Content-Type": "application/json"},
        json={"model": AI_MODEL,
              "messages": [{"role": "system", "content": agent_prompt()}] + history[-20:]},
        timeout=60)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]

# ============================================================
# ROUTES
# ============================================================

@app.route("/")
def index():
    return render_template("index.html", total=len(UNIVERSITIES))


@app.route("/chat")
def chat_page():
    return render_template("chat.html")


@app.route("/api/chat", methods=["POST"])
def api_chat():
    """Chat ka main endpoint: frontend yahan message bhejta hai."""
    data = request.get_json(force=True)
    session_id = data.get("session_id") or uuid.uuid4().hex
    text = data.get("message", "")

    if AI_API_KEY:                                   # asal AI mode
        s = SESSIONS.setdefault(session_id, {"history": []})
        s["history"].append({"role": "user", "content": text})
        try:
            reply = ai_reply(s["history"])
        except Exception as e:
            reply = "AI se rabta nahi ho saka: %s" % e
        s["history"].append({"role": "assistant", "content": reply})
        return jsonify({"reply": reply, "session_id": session_id})

    out = brain(session_id, text)                    # built-in brain
    out["session_id"] = session_id
    return jsonify(out)


@app.route("/api/match")
def api_match():
    """?budget=10000&ielts=6.5&countries=USA,Germany&field=CS"""
    q = request.args
    profile = {
        "ielts": num(q.get("ielts")),
        "budget": num(q.get("budget")),
        "field": q.get("field", ""),
        "countries": [c.strip() for c in q.get("countries", "").split(",") if c.strip()] or ALL_COUNTRIES,
    }
    res = match_universities(profile)
    out = {}
    for grp in ("dream", "match", "safe"):
        out[grp] = [card(u, grp) for u in res[grp]]
    return jsonify(out)


@app.route("/api/checklist")
def api_checklist():
    """?university=TU Munich  -> us country ka document checklist"""
    u = find_university(request.args.get("university", ""))
    if not u:
        return jsonify({"error": "University nahi mili"}), 404
    return jsonify({"university": u["University"], "country": u["Country"],
                    "checklist": CHECKLISTS.get(u["Country"], [])})


# ---------- Admin: table dekho, cell edit karo, Excel me save ----------

@app.route("/admin", methods=["GET", "POST"])
def admin():
    if request.method == "POST" and request.form.get("password") == ADMIN_PASSWORD:
        session["is_admin"] = True
    if not session.get("is_admin"):
        return render_template("admin_login.html")
    return render_template("admin.html", columns=COLUMNS,
                           universities=UNIVERSITIES, total=len(UNIVERSITIES))


@app.route("/admin/save", methods=["POST"])
def admin_save():
    if not session.get("is_admin"):
        return jsonify({"error": "login required"}), 403
    rows = request.get_json(force=True).get("rows", [])
    # Pehle backup, phir save
    wb.save("universities.backup.xlsx")
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append(r[:len(COLUMNS)])
    wb.save("universities.xlsx")
    # Memory me bhi taaza karo
    UNIVERSITIES.clear()
    for row in ws.iter_rows(min_row=2, values_only=True):
        u = dict(zip(COLUMNS, row))
        u["University"] = (u["University"] or "").strip()
        if u["University"]:
            UNIVERSITIES.append(u)
    return jsonify({"ok": True, "saved": len(UNIVERSITIES)})


@app.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return redirect("/")


if __name__ == "__main__":
    # debug=False taake do dafa load na ho
    app.run(host="0.0.0.0", port=5000, debug=False)
