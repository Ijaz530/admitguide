# 🎓 AdmitGuide — AI Admission Consultant

Cursor ki zaroorat **nahi** — ye app pehle se bani hui hai! Bas chalao.

## Chalane ka tareeqa (2 minute)

**Apne computer par:**
```bash
cd admitguide
pip install -r requirements.txt
python app.py
```
Phir browser me kholo: **http://127.0.0.1:5000**

**Online (free hosting) — Render.com:**
1. Ye folder GitHub par dalo
2. render.com par "New Web Service" → GitHub repo connect karo
3. Build command: `pip install -r requirements.txt` — Start command: `python app.py`
4. Free plan kaafi hai shuruwat ke liye

**Domain:** Namecheap se ~$12/saal (jaise admitguide.pk) → Render me "Custom Domain" me jor do.

## Kya kya bana hua hai

| Page | Kaam |
|---|---|
| `/` | Landing page — Start Chat button |
| `/chat` | Chat consultant: 5 sawal → Dream/Match/Safe shortlist → gap analysis → documents/deadlines/visa/SOP |
| `/api/match` | Direct matching API: `?budget=10000&ielts=6.5&countries=USA,Germany&field=CS` |
| `/api/checklist?university=TU Munich` | Us country ka document checklist |
| `/admin` | University table — cell edit karo, **Save** dabao, Excel update (password: `admin123`) |

## AI key lagana (optional)

`app.py` ke top par `AI_API_KEY = ""` me apni key paste karo —
phir chat asal AI se baat karega (baghair key ke built-in brain chalta hai, jo bilkul free hai).

## Files

- `app.py` — poora backend (Flask)
- `universities.xlsx` — 121 universities ka database (wahi Excel)
- `agent_prompt.txt` — AI ka system prompt
- `templates/` — landing + chat + admin pages
- `static/style.css` — styling
