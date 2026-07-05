# JobHunt LatAm
Local Flask web app that finds remote jobs open to LatAm applicants across 10 free job boards, flags work-authorization eligibility by country, scores each role against your resume, and generates tailored .docx edits that keep your original formatting. Runs offline; optional keyless AI mode.

<img width="964" height="987" alt="01-JobHunt" src="https://github.com/user-attachments/assets/e0b87654-9ce0-4271-9888-e7048a64f16a" />
<img width="962" height="772" alt="02-JobHunt-fixed2" src="https://github.com/user-attachments/assets/75f6bbe1-40cf-4568-8955-0b2df3ff5e8f" /> 
<img width="694" height="594" alt="03-JobHunt-fixed2" src="https://github.com/user-attachments/assets/eb8bc051-56b7-41c0-a54f-85f8ec712031" />



# JobHunt LatAm

A local Python web app that finds **remote jobs open to LatAm-based applicants** (including roles posted from the US, Canada, and EMEA that accept LatAm/Americas/worldwide candidates), matches them to **your resume**, suggests **role-specific resume changes**, and exports a **tailored `.docx`** with one click.

## What it does

1. **Job search** — aggregates live listings from **10 auto-searched sources** (no API keys): **Remotive, RemoteOK, Arbeitnow, Jobicy, Himalayas, The Muse, We Work Remotely, Working Nomads, Jobspresso,** and **Get on Board** (LatAm). Results are **filtered and ranked by how well they match your keywords**, so typing e.g. `DevOps` surfaces real DevOps roles first (not loosely-related ones). Dozens more boards that need login/keys (LinkedIn, Indeed, Wellfound, Upwork, Built In, Computrabajo, InfoJobs…) are listed under **More job boards** to open in your browser. Searches by your keywords; tick/untick any source. **Click any source's name (↗) to open its own site pre-searched with your current keyword.** A couple of boards can't be auto-searched in-app and open as links instead: **hiring.cafe** (it closed its public API) and **RemoteYeah** (needs a free account).
2. **Eligibility filter** — pick the **country you're based in** (or "All LatAm"). Each job is labelled `Eligible (LatAm)`, `Eligible (worldwide)`, `Eligible via Spain (may sponsor)`, `Possibly eligible`, or `Not eligible (needs local work auth)`. Jobs that are "remote **from** USA / Portugal / Europe / etc." are flagged **not eligible** because they require local work authorization you don't have. Only worldwide/anywhere roles, LatAm roles, your own country, and **Spain** (which sometimes sponsors) count as eligible. Toggle **"Eligible only"** to hide the rest.
3. **Resume matching** — upload a **`.pdf`, `.docx`, or `.txt`** resume. Matching uses **only your work-experience section** (Education, Courses and Certifications are ignored). Each job gets a **0–100 match score** plus the skills you already match and the gaps to close.
4. **Tailored suggestions** — click **✨ Tailor resume** on any job for a fit breakdown, before→after sentence edits, keywords to add, and an optional tailored summary. **If your resume already fits the role, it tells you so and recommends no changes** (nothing to apply).
5. **Apply changes → download** — when you upload a **`.docx`**, the chosen edits are applied **in place to your original file, preserving the exact formatting** (two-column header, bold, bullets, fonts) — only the selected sentences/words change. Your original file on disk is never modified; you get a tailored copy. (PDF/TXT uploads can't keep Word formatting, so a new plain `.docx` is generated instead — upload `.docx` for format-preserving edits.)
6. **Salary filter** — set a minimum salary. It only hides jobs that *disclose* a lower salary; jobs with no salary info still appear. Jobicy and Himalayas provide the most salary data.

## Local vs. AI engine

The engine **only affects the ✨ Tailor resume feature — job search is identical either way.** Open **⚙ Settings** to choose it (Local is the default):

- **Local only (no key)** — fully offline & free: keyword/skill-overlap scoring, keyword insertion, manual tips. Default.
- **Free AI (no key, no login)** — a keyless hosted LLM (Pollinations) that runs the query in the backend. No signup or API key. Produces the same **before→after sentence edits** as a paid key. Caveat: it's a shared free public service, so it can be slower/rate-limited, and your prompt (including resume text) is sent to it — prefer Local or your own key for sensitive data.
- **OpenAI** / **Anthropic (Claude)** — paste your own key for the most reliable, highest-quality tailoring. Keys stay in memory on your machine.

## Run it — easiest way

The app **installs its own dependencies automatically** into whatever Python runs it, so you normally don't need to install anything by hand.

- **Windows:** double-click **`run.bat`** (or run it in a terminal).
- **macOS / Linux:** run **`./run.sh`** in a terminal.

Then open **http://127.0.0.1:5000** in your browser. To stop it, press `Ctrl+C` in the terminal.

### Manual way (if you prefer)

```bash
cd jobhunt-latam
python -m pip install -r requirements.txt   # use python -m pip so it hits the right Python
python app.py
```

> Python 3.9+ recommended. If `python` isn't found, try `py` (Windows) or `python3` (macOS/Linux).
> Note: the first launch may pause briefly while it installs Flask, requests, python-docx and pdfplumber.

## Supported resume formats

`.pdf` (text-based — not scanned images), `.docx`, and `.txt`. Old `.doc` files aren't supported: open in Word and **Save As → .docx or PDF**. If a PDF is a scanned image with no selectable text, the app will tell you and you can upload a `.docx` instead.

## How the LatAm heuristic works

A job is flagged eligible when its location/description mentions LatAm, Latin America, the Americas, a LatAm country, Americas time zones (GMT-3…GMT-6 / UTC-3…UTC-6), or global/worldwide/anywhere remote — and **not** flagged when it says US-only / EU-only / UK-only / Canada-only, etc. Plain "remote" with no stated geo is shown as *Possibly eligible* so you don't miss anything.

## Job sources used

| Source | Salary data | Notes |
|---|---|---|
| Remotive | sometimes | Strong remote coverage, keyword search |
| RemoteOK | sometimes | Tech-heavy, tags |
| Arbeitnow | rarely | EU + global remote |
| Jobicy | **often** | Good geo labels (incl. Latin America) + salary |
| Himalayas | **often** | Location restrictions + timezones + salary |
| The Muse | rarely | Curated roles, "Flexible / Remote" |
| We Work Remotely | rarely | Popular remote board (RSS feed) |
| Working Nomads | sometimes | Curated remote roles across categories |
| Jobspresso | rarely | Hand-screened remote jobs (RSS feed) |
| Get on Board | **often** | LatAm-focused tech jobs + salary (great for Costa Rica) |
| hiring.cafe | — | Closed its public API; opens as a keyword link (search still works in your browser) |


### Adding even more sources later

`app.py` keeps each source as a small adapter function registered in the `SOURCES` dict. To add a keyed API (e.g. Adzuna or JSearch/RapidAPI) for richer US/Canada/EMEA salary data, write a `fetch_xxx(keywords)` that returns the same normalized job dict and add it to `SOURCES` (and a checkbox in `templates/index.html`).

## Notes

- All state is in-memory and single-user (it's a personal local tool). Restarting the server clears the uploaded resume.
- This tool helps you find and tailor — it does **not** auto-submit applications. Use each job's **Apply ↗** link to apply on the source site.
- If one job source is temporarily down, the app skips it and still returns results from the others.
