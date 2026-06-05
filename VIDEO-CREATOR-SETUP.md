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

## Section 3 — ElevenLabs ✅ Complete

1. Account created at [elevenlabs.io](https://elevenlabs.io)
2. API key created — named `unitrend-video-creator`
3. Endpoints enabled: **Text to Speech** (Access), **Voices** (Read), **Voice Generation** (Access)
4. Key saved to `Marketing\Pete\Video Creation Dashboard\elevenlabs-key.txt`

**Plan:** Free tier gives 10,000 characters/month. **Starter plan at $22/mo** gives 30,000 characters — enough for ~15–20 product videos per month. Upgrade when past prototype stage.

---

## Section 4 — GitHub Private Repository ✅ Complete

- Repo: `unitrendus-git/unitrend-video-creator` (Private)
- URL: `https://github.com/unitrendus-git/unitrend-video-creator.git`
- Account: `unitrendus@gmail.com`
- Git identity configured: `Pete Stoermer / unitrendus@gmail.com`

---

## Section 5 — Streamlit Cloud Account

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Click **Sign in with GitHub** — use `unitrendus@gmail.com`
3. After login, you'll land on your apps dashboard
4. Come back here after `app.py` is built and pushed (Section 8)

---

## Section 6 — Local Installation ✅ Complete

All steps verified. Terminal: **PowerShell** (not Command Prompt).

Reactivate venv at the start of every session:
```powershell
cd "C:\Users\pjsto\OneDrive - Uni-Trend\Documents - Uni-T NA share point\Marketing\Pete\Video Creation Dashboard"
venv\Scripts\activate
```

### 6a — Python environment ✅
Virtual environment created and activated. `(venv)` confirmed in prompt.

> If PowerShell blocks activation: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

### 6b — pip packages ✅
All packages installed from `requirements.txt`.

### 6c — Poppler ✅ v26.02.0
- Installed at: `C:\Program Files\Poppler\poppler-26.02.0\`
- PATH: `C:\Program Files\Poppler\poppler-26.02.0\Library\bin`
- Verified: `pdftoppm version 26.02.0`

### 6d — FFmpeg ✅ v8.1.1
- Installed at: `C:\Program Files\ffmpeg\ffmpeg-8.1.1-essentials_build\`
- PATH: `C:\Program Files\ffmpeg\ffmpeg-8.1.1-essentials_build\bin`
- Verified: `ffmpeg version 8.1.1-essentials_build-www.gyan.dev`

### 6e — Secrets file ✅
- Location: `.streamlit\secrets.toml`
- Contains: `ANTHROPIC_API_KEY` and `ELEVENLABS_API_KEY`
- Never committed to GitHub (excluded by `.gitignore`)

### 6f — Run locally
```powershell
streamlit run app.py
```
Dashboard opens at `http://localhost:8501`.

---

## Section 7 — Project files

All files live in `Marketing\Pete\Video Creation Dashboard\`:

| File | Status | What it is |
|---|---|---|
| `VIDEO-CREATOR-SETUP.md` | ✅ Done | This guide |
| `requirements.txt` | ✅ Done | Python package list |
| `.gitignore` | ✅ Done | Keeps secrets and venv out of GitHub |
| `.streamlit/secrets.toml` | ✅ Done | API keys — local only, never committed |
| `app.py` | ✅ Done | Main Streamlit dashboard |
| `video_creator_utils.py` | ⬜ Next | Helper functions: PDF extraction, TTS, FFmpeg assembly |
| `README.md` | ⬜ Later | Project README for the GitHub repo |

---

## Section 8 — Deploy to Streamlit Cloud

Initial files pushed to GitHub (commit `957ed89`). Once `app.py` is built and tested locally:

1. Stage and push new files:
```powershell
git add .
git commit -m "Add app.py — video creator dashboard"
git push
```

2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select `unitrendus-git/unitrend-video-creator` → main file: `app.py` → **Deploy**
4. Go to **App settings → Secrets** and paste:

```toml
ANTHROPIC_API_KEY = "sk-ant-YOUR-KEY-HERE"
ELEVENLABS_API_KEY = "sk_YOUR-KEY-HERE"
```

5. Click **Save** — app restarts with live keys

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
| Secrets file | ✅ Done |
| GitHub repo + initial push | ✅ Done |
| `app.py` — Streamlit dashboard | ✅ Done |
| Script generation (Claude API + PDF input) | ✅ Done (inside app.py) |
| Edge TTS — validate timing and pacing | ⬜ Next |
| FFmpeg assembly — frames + audio + lower-thirds | ⬜ Next |
| Streamlit Cloud deploy | ⬜ After local test passes |
| ElevenLabs voice upgrade | ⬜ Phase 2 |
| 16:9 + 9:16 dual output from same source | ⬜ Phase 2 |
