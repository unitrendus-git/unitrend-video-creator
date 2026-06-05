# Video Creation Dashboard — Setup Guide
**Project:** UNI-T NA Video Creation Pipeline  
**Location:** `Marketing\Pete\Video Creation Dashboard\`  
**Last updated:** 2026-06-05

This guide covers every service registration, API key, and local installation step required to run the Video Creation Dashboard from end to end. Follow in order — later steps depend on earlier ones.

---

## Overview of what you're registering for

| Service | Purpose | Cost | Priority |
|---|---|---|---|
| Anthropic API | Script generation from PDF/PPTX | ~$5–15/mo at this volume | Required |
| Edge TTS | Text-to-speech (voiceover audio) | Free | Required (Phase 1) |
| ElevenLabs | Production-quality TTS voices | $22/mo (Starter) | Phase 2 |
| FFmpeg | Video assembly (frames + audio + lower-thirds) | Free, open source | Required |
| Streamlit Cloud | Host the dashboard (web UI) | Free tier available | Required |
| GitHub | Private repo to deploy from | Free | Required |

> **Already have:** Anthropic API key (`Anthropic Pete Key.txt` in `Marketing\Pete\`). Skip Section 1 registration — just retrieve the key value.

---

## Section 1 — Anthropic API Key

### You already have this.

Retrieve it from `Marketing\Pete\Anthropic Pete Key.txt`. Copy the `sk-ant-...` key value — you'll paste it into Streamlit secrets in Section 5.

If you ever need to create a new key or rotate it:
1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Sign in with your Anthropic account
3. Left sidebar → **API Keys** → **Create Key**
4. Name it something like `unitrend-video-creator`
5. Copy immediately — it won't show again

**Where it goes:** Streamlit secrets file (Section 5) and local `.streamlit/secrets.toml` (Section 6e).

---

## Section 2 — Edge TTS (Free, no registration)

Edge TTS is a Python library that wraps Microsoft's text-to-speech engine. No account, no API key, no registration required.

It installs as a Python package. That's it — handled in Section 6b (Installation).

**Voices available without registration:** Ryan (US Male), Jenny (US Female), Eric (US Male), and ~300 others. Run `edge-tts --list-voices` after install to see all options.

---

## Section 3 — ElevenLabs (Phase 2 — do when ready to upgrade voices)

1. Go to [elevenlabs.io](https://elevenlabs.io)
2. Click **Sign Up** → create account with your Uni-Trend email
3. After login, go to **Profile → API Keys** (bottom left)
4. Click **Create API Key** → name it `unitrend-video-creator`
5. Enable endpoints: **Text to Speech** (Access), **Voices** (Read), **Voice Generation** (Access) — leave all others on No Access
6. Copy the key (`sk_...`) and save it to `Marketing\Pete\Video Creation Dashboard\` as `elevenlabs-key.txt`

**Plan:** Free tier gives 10,000 characters/month (roughly 5–6 short videos). **Starter plan at $22/mo** gives 30,000 characters — enough for ~15–20 product videos per month. Upgrade when you're past prototype stage.

**Where the key goes:** Add to Streamlit secrets alongside the Anthropic key (Section 5).

---

## Section 4 — GitHub Private Repository

You'll push the dashboard code here so Streamlit Cloud can deploy it. The repo must stay **private** because it references your secrets structure.

1. Go to [github.com](https://github.com) → sign in (or create a free account)
2. Click **+** (top right) → **New repository**
3. Name it: `unitrend-video-creator`
4. Set to **Private**
5. Do NOT initialize with README (you'll push existing files)
6. Click **Create repository**
7. GitHub shows you a push command — keep this page open for Section 8

> **Note:** If you already have a GitHub account from the social listening project, use the same account. Just create a new repo.

---

## Section 5 — Streamlit Cloud Account

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Click **Sign in with GitHub** — use the same GitHub account from Section 4
3. After login, you'll land on your apps dashboard — nothing to configure yet
4. Come back here after you push code to GitHub (Section 8)

---

## Section 6 — Local Installation

Everything below runs on your local Windows machine in a terminal (PowerShell or Command Prompt). The project folder is:

```
C:\Users\pjsto\OneDrive - Uni-Trend\Documents - Uni-T NA share point\Marketing\Pete\Video Creation Dashboard\
```

### 6a — Python environment ✅ Complete

Virtual environment created and activated. `(venv)` confirmed in prompt.

```powershell
cd "C:\Users\pjsto\OneDrive - Uni-Trend\Documents - Uni-T NA share point\Marketing\Pete\Video Creation Dashboard"
python -m venv venv
venv\Scripts\activate
```

> **Note:** If PowerShell blocks the activate script, run this once first:
> `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

### 6b — Install Python dependencies ✅ Complete

All packages installed cleanly.

```powershell
pip install -r requirements.txt
```

Installs: `streamlit`, `anthropic`, `edge-tts`, `python-pptx`, `pdf2image`, `Pillow`, `ffmpeg-python`, `requests`, `python-dotenv`

### 6c — Install Poppler ✅ Complete — v26.02.0

Installed at: `C:\Program Files\Poppler\poppler-26.02.0\`  
PATH entry added: `C:\Program Files\Poppler\poppler-26.02.0\Library\bin`  
Verified: `pdftoppm version 26.02.0`

To reinstall if needed:
1. Download from [github.com/oschwartz10612/poppler-windows/releases](https://github.com/oschwartz10612/poppler-windows/releases)
2. Extract to `C:\Program Files\Poppler\`
3. Add `...\Library\bin` to System PATH
4. Restart terminal and verify: `pdftoppm -v`

### 6d — Install FFmpeg ✅ Complete — v8.1.1

Installed at: `C:\Program Files\ffmpeg\ffmpeg-8.1.1-essentials_build\`  
PATH entry added: `C:\Program Files\ffmpeg\ffmpeg-8.1.1-essentials_build\bin`  
Verified: `ffmpeg version 8.1.1-essentials_build-www.gyan.dev`

To reinstall if needed:
1. Download from [gyan.dev/ffmpeg/builds](https://www.gyan.dev/ffmpeg/builds/) → `ffmpeg-release-essentials.zip`
2. Extract to `C:\Program Files\ffmpeg\`
3. Add `...\bin` to System PATH
4. Restart terminal and verify: `ffmpeg -version`

### 6e — Set up local secrets

Create the Streamlit secrets file so the app can find your API keys locally:

```powershell
mkdir .streamlit
New-Item .streamlit\secrets.toml
notepad .streamlit\secrets.toml
```

Paste and fill in:

```toml
ANTHROPIC_API_KEY = "sk-ant-YOUR-KEY-HERE"
ELEVENLABS_API_KEY = "sk_YOUR-KEY-HERE"
```

Retrieve the Anthropic key from `Marketing\Pete\Anthropic Pete Key.txt`.  
Retrieve the ElevenLabs key from `Marketing\Pete\Video Creation Dashboard\elevenlabs-key.txt`.

> **Important:** `secrets.toml` must never be committed to GitHub. The `.gitignore` file already excludes it.

### 6f — Run the app locally to test

```powershell
streamlit run app.py
```

Dashboard opens at `http://localhost:8501`. Upload a PDF and confirm script generation works before deploying to the cloud.

---

## Section 7 — Project files

All files live in `Marketing\Pete\Video Creation Dashboard\`:

| File | Status | What it is |
|---|---|---|
| `VIDEO-CREATOR-SETUP.md` | ✅ Done | This guide |
| `requirements.txt` | ✅ Done | Python package list |
| `.gitignore` | ✅ Done | Keeps secrets and venv out of GitHub |
| `.streamlit/secrets.toml` | ⬜ Section 6e | Your API keys — local only, never committed |
| `app.py` | ⬜ Next build | Main Streamlit dashboard |
| `video_creator_utils.py` | ⬜ Next build | Helper functions: PDF extraction, TTS, FFmpeg assembly |
| `README.md` | ⬜ Next build | Project README for the GitHub repo |

---

## Section 8 — Deploy to Streamlit Cloud

Once local testing passes:

1. Initialize Git in the project folder:

```powershell
git init
git add .
git commit -m "Initial commit — video creator dashboard"
```

2. Connect to the GitHub repo from Section 4:

```powershell
git remote add origin https://github.com/YOUR-USERNAME/unitrend-video-creator.git
git branch -M main
git push -u origin main
```

3. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
4. Select your GitHub repo → set **Main file path** to `app.py` → click **Deploy**
5. After deploy, go to **App settings → Secrets** and paste:

```toml
ANTHROPIC_API_KEY = "sk-ant-YOUR-KEY-HERE"
ELEVENLABS_API_KEY = "sk_YOUR-KEY-HERE"
```

6. Click **Save** — the app restarts with live keys

Your dashboard is now live at a `*.streamlit.app` URL. Control access under **Settings → Sharing**.

---

## Section 9 — Quick reference: where each key lives

| Key | Local location | Cloud location |
|---|---|---|
| Anthropic API key | `.streamlit/secrets.toml` | Streamlit Cloud → App Settings → Secrets |
| ElevenLabs API key | `.streamlit/secrets.toml` | Streamlit Cloud → App Settings → Secrets |
| Source copy of Anthropic key | `Marketing\Pete\Anthropic Pete Key.txt` | — |
| Source copy of ElevenLabs key | `Marketing\Pete\Video Creation Dashboard\elevenlabs-key.txt` | — |

---

## Build order

| Step | Status |
|---|---|
| Dashboard UI — React prototype | ✅ Done |
| Environment setup (venv, packages, Poppler, FFmpeg) | ✅ Done |
| Secrets file (`.streamlit/secrets.toml`) | ⬜ Section 6e |
| `app.py` — Streamlit Python dashboard | ⬜ Next |
| Script generation (Claude API + PDF input) | ⬜ Next |
| Edge TTS — validate timing and pacing | ⬜ Next |
| FFmpeg assembly — frames + audio + lower-thirds | ⬜ Next |
| ElevenLabs voice upgrade | ⬜ Phase 2 |
| 16:9 + 9:16 dual output from same source | ⬜ Phase 2 |
