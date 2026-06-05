import streamlit as st
import anthropic
import asyncio
import os
import json
import tempfile
import time
import datetime
import base64
import requests
import pdfplumber
from pptx import Presentation
import edge_tts

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="UNI-T Video Creator",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main .block-container { max-width: 960px; padding-top: 2rem; }
    .scene-header {
        font-size: 0.78rem; font-weight: 600; text-transform: uppercase;
        letter-spacing: 0.06em; color: #555; margin-bottom: 2px;
    }
    .timing-badge { font-size: 0.75rem; color: #999; font-style: italic; }
    div[data-testid="stHorizontalBlock"] { align-items: flex-start; }
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
SCENE_CONFIG = {
    "hero":        {"label": "Hero",          "timing": "0–5s",   "target_words": 15},
    "problem":     {"label": "Problem",       "timing": "5–20s",  "target_words": 45},
    "features":    {"label": "Key Features",  "timing": "20–50s", "target_words": 55},
    "application": {"label": "Application",   "timing": "50–65s", "target_words": 35},
    "cta":         {"label": "Call to Action","timing": "65–75s", "target_words": 20},
}

EL_VOICES = [
    ("Adam — Deep, authoritative (ElevenLabs)",      "pNInz6obpgDQGcFmaJgB", "el"),
    ("Antoni — Natural, conversational (ElevenLabs)", "ErXwobaYiN019PkySvjV", "el"),
    ("Josh — Clear, professional (ElevenLabs)",       "TxGEqnHWrfWFTfGW9XjX", "el"),
    ("Arnold — Strong, confident (ElevenLabs)",       "VR6AewLTigWG4xSOukaG", "el"),
    ("Rachel — Warm, professional F (ElevenLabs)",    "21m00Tcm4TlvDq8ikWAM", "el"),
]
EDGE_VOICES = [
    ("Guy — US Male (Edge TTS, free)",        "en-US-GuyNeural",         "edge"),
    ("Jenny — US Female (Edge TTS, free)",     "en-US-JennyNeural",       "edge"),
    ("Eric — US Male (Edge TTS, free)",        "en-US-EricNeural",        "edge"),
    ("Christopher — US Male (Edge TTS, free)", "en-US-ChristopherNeural", "edge"),
]
ALL_VOICES   = EL_VOICES + EDGE_VOICES
VOICE_LABELS = [v[0] for v in ALL_VOICES]
VOICE_IDS    = [v[1] for v in ALL_VOICES]
VOICE_TIERS  = [v[2] for v in ALL_VOICES]

EL_MODELS = {
    "eleven_v3":              "v3 — most expressive (recommended)",
    "eleven_multilingual_v2": "Multilingual v2 — best narration",
    "eleven_turbo_v2":        "Turbo v2 — fast, good quality",
}

SHEET_TAB = "Video Scripts"
SHEET_HEADERS = [
    "script_id", "product_name", "version", "created_at", "source_file",
    "voice_label", "el_model", "stability", "style", "total_words",
    "est_duration", "hero", "problem", "features_vo", "features_callouts",
    "application", "cta_vo", "cta_line",
]

BRAND_SYSTEM_PROMPT = """You are a technical video scriptwriter for UNI-T North America, a B2B test and measurement brand headquartered in Fort Worth, TX.

VOICE RULES (non-negotiable):
- Name the engineer's problem BEFORE the solution — never lead with features
- Specs always written ascending order (low to high) — e.g. "20V / 80V / 160V" not "160V / 80V / 20V"
- No competitor names anywhere in the script
- CTA always references uni-trendus.com
- Warranty expressed as "Industry-leading 3+2 warranty"
- No aerospace application targeting — safe alternatives: industrial, R&D, manufacturing, transportation
- Tone: confident, direct, peer-to-peer (engineer speaking to engineer)
- Tagline available if appropriate: Design. Debug. Deploy.™

SCRIPT STRUCTURE (75 seconds total, ~130 words/minute spoken pace, target ~165 words total):
- Hero (0–5s, ~15 words): Product name + one-line positioning statement
- Problem (5–20s, ~45 words): The engineer's pain point this product solves
- Key Features (20–50s, ~55 words): 3 key specs/capabilities — one at a time, clear and factual
- Application (50–65s, ~35 words): Real-world use case — who uses this and how
- CTA (65–75s, ~20 words): Direct call to action ending with uni-trendus.com

OUTPUT: Valid JSON only. No markdown, no preamble, no explanation. Schema:
{
  "product_name": "string",
  "hero": {"voiceover": "string"},
  "problem": {"voiceover": "string"},
  "features": {"voiceover": "string", "callouts": ["string", "string", "string"]},
  "application": {"voiceover": "string"},
  "cta": {"voiceover": "string", "cta_line": "string"}
}"""

# ── Google Sheets ──────────────────────────────────────────────────────────────
def get_secret(key, fallback=""):
    try:
        return st.secrets[key]
    except Exception:
        return fallback

@st.cache_resource(ttl=300)
def sheets_connect():
    if not GSPREAD_OK:
        return None
    sa_raw = get_secret("GOOGLE_SERVICE_ACCOUNT")
    if not sa_raw:
        return None
    try:
        sa_info = json.loads(sa_raw)
        pk = sa_info.get("private_key", "")
        if pk and r"\n" in pk:
            sa_info["private_key"] = pk.replace(r"\n", chr(10))
        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.sidebar.warning(f"Sheets auth failed: {e}")
        return None

def sheets_open():
    gc = sheets_connect()
    if not gc:
        return None
    try:
        return gc.open_by_key(get_secret("GSHEET_ID"))
    except Exception as e:
        st.sidebar.warning(f"Could not open sheet: {e}")
        return None

def sheets_ensure_tab(sh):
    if not sh:
        return
    existing = [ws.title for ws in sh.worksheets()]
    if SHEET_TAB not in existing:
        ws = sh.add_worksheet(title=SHEET_TAB, rows=1000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS)

def sheets_get_scripts(sh) -> list[dict]:
    if not sh:
        return []
    try:
        ws = sh.worksheet(SHEET_TAB)
        records = ws.get_all_records()
        # Newest first — skip deletion log rows
        return [r for r in reversed(records)
                if r.get("script_id", "") and
                not str(r.get("script_id", "")).startswith("DELETED")]
    except Exception:
        return []

def sheets_delete_script(sh, script_id: str):
    """Delete script row by script_id and append a deletion log entry."""
    if not sh:
        return False
    try:
        ws   = sh.worksheet(SHEET_TAB)
        data = ws.get_all_values()
        for i, row in enumerate(data[1:], 2):
            if row and row[0] == script_id:
                ws.delete_rows(i)
                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                log_row = [f"DELETED:{script_id}", script_id, "", now,
                           "", "deleted", "", "", "", "", "", "", "", "", "", "", "", ""]
                ws.append_row(log_row, value_input_option="USER_ENTERED")
                return True
        return False
    except Exception as e:
        st.error(f"Delete error: {e}")
        return False

def sheets_next_version(records: list[dict], product_name: str) -> int:
    safe = product_name.upper().replace(" ", "_")
    existing = [
        int(r.get("version", 0))
        for r in records
        if r.get("script_id", "").startswith(safe + "_v")
    ]
    return max(existing, default=0) + 1

def sheets_save_script(sh, scenes: dict, product_name: str, source_file: str,
                        voice_label: str, el_model: str,
                        stability: float, style: float) -> str:
    if not sh:
        return ""
    try:
        ws      = sh.worksheet(SHEET_TAB)
        records = ws.get_all_records()
        version = sheets_next_version(records, product_name)
        safe    = product_name.upper().replace(" ", "_")
        sid     = f"{safe}_v{version}"
        now     = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        total_w = sum(count_words(s.get("voiceover", "")) for s in scenes.values())
        callouts = scenes.get("features", {}).get("callouts", [])
        row = [
            sid, product_name, version, now, source_file,
            voice_label, el_model,
            round(stability, 2), round(style, 2),
            total_w, estimate_duration(total_w),
            scenes.get("hero", {}).get("voiceover", ""),
            scenes.get("problem", {}).get("voiceover", ""),
            scenes.get("features", {}).get("voiceover", ""),
            " | ".join(callouts) if callouts else "",
            scenes.get("application", {}).get("voiceover", ""),
            scenes.get("cta", {}).get("voiceover", ""),
            scenes.get("cta", {}).get("cta_line", "Available at uni-trendus.com"),
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        return sid
    except Exception as e:
        st.error(f"Sheets save error: {e}")
        return ""

def sheets_load_script(record: dict) -> dict:
    callouts_raw = record.get("features_callouts", "")
    callouts = [c.strip() for c in callouts_raw.split("|") if c.strip()] if callouts_raw else []
    return {
        "hero":        {"voiceover": record.get("hero", ""),        "callouts": [], "cta_line": ""},
        "problem":     {"voiceover": record.get("problem", ""),     "callouts": [], "cta_line": ""},
        "features":    {"voiceover": record.get("features_vo", ""), "callouts": callouts, "cta_line": ""},
        "application": {"voiceover": record.get("application", ""), "callouts": [], "cta_line": ""},
        "cta":         {"voiceover": record.get("cta_vo", ""),      "callouts": [],
                        "cta_line": record.get("cta_line", "Available at uni-trendus.com")},
    }

# ── Helpers ────────────────────────────────────────────────────────────────────
def extract_text_from_pdf(file_bytes: bytes) -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        pages = []
        with pdfplumber.open(tmp_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    pages.append(t)
        return "\n\n".join(pages)
    finally:
        os.unlink(tmp_path)

def extract_text_from_pptx(file_bytes: bytes) -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        prs = Presentation(tmp_path)
        slides = []
        for i, slide in enumerate(prs.slides, 1):
            parts = [shape.text.strip() for shape in slide.shapes
                     if hasattr(shape, "text") and shape.text.strip()]
            if parts:
                slides.append(f"[Slide {i}]\n" + "\n".join(parts))
        return "\n\n".join(slides)
    finally:
        os.unlink(tmp_path)

def count_words(text: str) -> int:
    return len(text.strip().split()) if text.strip() else 0

def estimate_duration(total_words: int) -> float:
    return round((total_words / 130) * 60, 1)

def generate_script(product_name: str, file_bytes: bytes, is_pdf: bool,
                    extracted_text: str) -> dict:
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    if is_pdf:
        b64 = base64.standard_b64encode(file_bytes).decode("utf-8")
        user_content = [
            {"type": "document",
             "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
            {"type": "text",
             "text": f"Product name: {product_name or 'extract from document'}.\n\n"
                     "Generate a 75-second product video script following the exact JSON schema."},
        ]
    else:
        user_content = [{"type": "text",
                         "text": f"Product name: {product_name or 'unknown'}.\n\n"
                                 f"PPTX content:\n\n{extracted_text[:8000]}\n\n"
                                 "Generate a 75-second product video script following the exact JSON schema."}]
    resp = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=1000,
        system=BRAND_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
    )
    raw = resp.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

def generate_tts_elevenlabs(text: str, voice_id: str, output_path: str,
                             model_id: str, stability: float, style: float):
    api_key = get_secret("ELEVENLABS_API_KEY")
    if not api_key:
        raise ValueError("ELEVENLABS_API_KEY not set in secrets.toml")
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}"
    headers = {"xi-api-key": api_key, "Content-Type": "application/json", "Accept": "audio/mpeg"}
    payload = {
        "text": text,
        "model_id": model_id,
        "voice_settings": {
            "stability": stability,
            "similarity_boost": 0.75,
            "style": style,
            "use_speaker_boost": True,
        },
    }
    r = requests.post(url, json=payload, headers=headers, timeout=90)
    if r.status_code != 200:
        raise ValueError(f"ElevenLabs error {r.status_code}: {r.text[:200]}")
    with open(output_path, "wb") as f:
        f.write(r.content)

def build_voiceover(scenes: dict, model_id: str = "eleven_turbo_v2") -> str:
    is_v3 = model_id == "eleven_v3"
    # v3 uses [pause] audio tags; v1/v2 use SSML <break> tags
    BREAKS = {
        "hero":        "[pause]" if is_v3 else '<break time="1.2s"/>',
        "problem":     "[short pause]" if is_v3 else '<break time="0.8s"/>',
        "features":    "[short pause]" if is_v3 else '<break time="0.8s"/>',
        "application": "[short pause]" if is_v3 else '<break time="0.8s"/>',
        "cta":         "",
    }
    parts = []
    for k in SCENE_CONFIG:
        vo = scenes.get(k, {}).get("voiceover", "").strip()
        if vo:
            parts.append(vo + BREAKS.get(k, ""))
    text = " ".join(parts)
    text = text.replace("UNI-T", "Unity").replace("Uni-T", "Unity").replace("uni-t", "Unity")
    return text

# ── Session state ──────────────────────────────────────────────────────────────
DEFAULTS = {
    "scenes": {}, "product_name": "", "status": "idle",
    "log": [], "audio_path": None, "file_bytes": None, "file_name": "",
    "pending_delete": None, "audio_history": [],
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

def add_log(msg):
    st.session_state.log.append(f"{time.strftime('%H:%M:%S')} — {msg}")

def reset():
    for k, v in DEFAULTS.items():
        st.session_state[k] = v if not isinstance(v, (dict, list)) else type(v)()
    st.session_state.status = "idle"

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🎬 UNI-T Video Creator")
    st.divider()

    sh = sheets_open()
    if sh:
        try:
            sheets_ensure_tab(sh)
            st.success("✓ Google Sheets connected")
        except Exception as e:
            st.warning(f"Sheets tab error: {e}")
    else:
        st.info("Google Sheets not connected — add GOOGLE_SERVICE_ACCOUNT and GSHEET_ID to secrets.")

    st.divider()

    # ── Audio settings ─────────────────────────────────────────────────────────
    st.markdown("**🎙 Audio settings**")
    voice_idx = st.selectbox(
        "Voice",
        options=list(range(len(ALL_VOICES))),
        format_func=lambda i: VOICE_LABELS[i],
        index=0,
        help="ElevenLabs voices require ELEVENLABS_API_KEY in secrets.toml",
    )
    voice_id   = VOICE_IDS[voice_idx]
    voice_tier = VOICE_TIERS[voice_idx]

    if voice_tier == "el":
        el_model_key = st.selectbox(
            "Model",
            options=list(EL_MODELS.keys()),
            format_func=lambda k: EL_MODELS[k],
            index=0,
            help="Multilingual v2 sounds more natural but takes ~2x longer to generate",
        )
        with st.expander("Fine-tune voice", expanded=False):
            stability = st.slider("Stability", 0.0, 1.0, 0.45, 0.05,
                help="Higher = more consistent. Lower = more expressive.")
            style = st.slider("Style", 0.0, 1.0, 0.15, 0.05,
                help="Higher = more dramatic. Keep 0.1–0.25 for B2B narration.")
            st.caption("💡 v3: use `[pause]` and `[short pause]`. v2: use `<break time=\"0.5s\"/>`")
    else:
        el_model_key = "eleven_turbo_v2"
        stability    = 0.5
        style        = 0.0

    st.markdown("[📖 SSML reference →](https://elevenlabs.io/docs/best-practices/prompting/controls)")
    st.caption("Opens in new tab — emphasis, pauses, pacing, pronunciation")

    st.divider()

    # ── Load / delete saved scripts ────────────────────────────────────────────
    st.markdown("**📂 Load saved script**")
    if sh:
        records = sheets_get_scripts(sh)
        if records:
            options = {
                f"{r['script_id']} — {r['created_at']} ({r['total_words']}w)": i
                for i, r in enumerate(records)
            }
            selected_label  = st.selectbox("Select version", list(options.keys()),
                                            label_visibility="collapsed")
            selected_idx    = options[selected_label]
            selected_record = records[selected_idx]
            selected_sid    = selected_record.get("script_id", "")

            load_col, del_col = st.columns([3, 1])
            with load_col:
                if st.button("↑ Load", use_container_width=True):
                    st.session_state.scenes       = sheets_load_script(selected_record)
                    st.session_state.product_name = selected_record.get("product_name", "")
                    st.session_state.status       = "done"
                    st.session_state.pending_delete = None
                    add_log(f"Loaded {selected_sid} from Sheets")
                    st.rerun()
            with del_col:
                if st.button("🗑", use_container_width=True, help="Delete this version"):
                    st.session_state.pending_delete = selected_sid

            # Confirmation dialog
            if st.session_state.pending_delete == selected_sid:
                st.warning(f"Delete **{selected_sid}**? Cannot be undone.")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✓ Confirm", use_container_width=True, type="primary"):
                        with st.spinner("Deleting…"):
                            ok = sheets_delete_script(sh, selected_sid)
                        if ok:
                            add_log(f"Deleted {selected_sid} from Sheets")
                            st.session_state.pending_delete = None
                            st.rerun()
                with c2:
                    if st.button("✕ Cancel", use_container_width=True):
                        st.session_state.pending_delete = None
                        st.rerun()
        else:
            st.caption("No saved scripts yet.")
    else:
        st.caption("Connect Sheets to enable load.")

    st.divider()
    if st.button("↺ Reset", use_container_width=True):
        reset()
        st.rerun()

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("## 🎬 UNI-T Video Creator")
st.caption("PDF or PPTX → Claude script → ElevenLabs audio → ready for FFmpeg assembly")
st.divider()

left, right = st.columns([1, 1.4], gap="large")

# ═══════════════════════════════════════════════════════════════════════════════
# LEFT COLUMN
# ═══════════════════════════════════════════════════════════════════════════════
with left:

    st.markdown("**Step 1 — Source file**")
    uploaded = st.file_uploader("PDF or PPTX", type=["pdf", "pptx", "ppt"],
                                 label_visibility="collapsed")
    if uploaded:
        st.session_state.file_bytes = uploaded.read()
        st.session_state.file_name  = uploaded.name
        st.success(f"✓ {uploaded.name} ({len(st.session_state.file_bytes)/1024:.0f} KB)")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Step 2 — Configure**")
    product_name = st.text_input("Product / series name",
                                  value=st.session_state.product_name,
                                  placeholder="e.g. UDP5000 Series")
    st.session_state.product_name = product_name

    output_format = st.radio("Output format",
        ["16:9 — YouTube / LinkedIn", "9:16 — Shorts / Reels", "Both formats"],
        index=2, horizontal=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Step 3 — Generate**")

    gen_col, audio_col = st.columns(2)
    with gen_col:
        generate_btn = st.button("✦ Generate script", use_container_width=True, type="primary",
            disabled=st.session_state.file_bytes is None
                     or st.session_state.status in ["generating", "extracting"])
    with audio_col:
        audio_btn = st.button("▶ Generate audio", use_container_width=True,
            disabled=not st.session_state.scenes)

    # Save to Sheets
    if st.session_state.scenes:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Save to Sheets", use_container_width=True, disabled=not sh,
                     help="Saves current script as a new version"):
            with st.spinner("Saving…"):
                sid = sheets_save_script(sh, st.session_state.scenes,
                    st.session_state.product_name or "Unknown",
                    st.session_state.file_name or "",
                    VOICE_LABELS[voice_idx], el_model_key, stability, style)
            if sid:
                st.success(f"Saved as **{sid}**")
                add_log(f"Saved to Sheets as {sid}")
            else:
                st.error("Save failed — check Sheets connection.")

    # Generate logic
    if generate_btn and st.session_state.file_bytes:
        is_pdf = st.session_state.file_name.lower().endswith(".pdf")
        st.session_state.status = "extracting"
        add_log(f"Loading {st.session_state.file_name}")
        with st.spinner("Extracting text…"):
            try:
                extracted = (extract_text_from_pdf(st.session_state.file_bytes) if is_pdf
                             else extract_text_from_pptx(st.session_state.file_bytes))
                add_log(f"Extracted {count_words(extracted)} words")
            except Exception as e:
                st.error(f"Extraction error: {e}")
                st.session_state.status = "error"
                extracted = ""
        if extracted or is_pdf:
            st.session_state.status = "generating"
            add_log("Sending to Claude API…")
            with st.spinner("Generating script with Claude…"):
                try:
                    parsed = generate_script(st.session_state.product_name,
                                             st.session_state.file_bytes, is_pdf, extracted)
                    if parsed.get("product_name") and not st.session_state.product_name:
                        st.session_state.product_name = parsed["product_name"]
                    scenes = {}
                    for key in SCENE_CONFIG:
                        raw = parsed.get(key, {})
                        scenes[key] = {
                            "voiceover": raw.get("voiceover", ""),
                            "callouts":  raw.get("callouts", []),
                            "cta_line":  raw.get("cta_line", "Available at uni-trendus.com"),
                        }
                    st.session_state.scenes = scenes
                    st.session_state.status = "done"
                    total_w = sum(count_words(s["voiceover"]) for s in scenes.values())
                    add_log(f"Script ready — {total_w}w / ~{estimate_duration(total_w)}s")
                    st.rerun()
                except Exception as e:
                    st.error(f"Script generation error: {e}")
                    st.session_state.status = "error"
                    add_log(f"Error: {e}")

    # Audio logic
    if audio_btn and st.session_state.scenes:
        live_scenes = {}
        for k in SCENE_CONFIG:
            live_scenes[k] = {
                "voiceover": st.session_state.get(f"vo_{k}",
                             st.session_state.scenes.get(k, {}).get("voiceover", "")),
                "callouts":  st.session_state.scenes.get(k, {}).get("callouts", []),
                "cta_line":  st.session_state.scenes.get(k, {}).get("cta_line", ""),
            }
        full_text = build_voiceover(live_scenes, el_model_key)
        add_log(f"Generating audio — {VOICE_LABELS[voice_idx]}")
        with st.spinner("Generating audio…"):
            try:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".mp3")
                tmp.close()
                if voice_tier == "el":
                    generate_tts_elevenlabs(full_text, voice_id, tmp.name,
                                            el_model_key, stability, style)
                else:
                    asyncio.run(edge_tts.Communicate(full_text, voice_id).save(tmp.name))
                st.session_state.audio_path = tmp.name
                # Add to audio history
                safe_name = (st.session_state.product_name or "product").lower().replace(" ", "_")
                ts = time.strftime("%Y-%m-%d_%H-%M")
                audio_label = f"{st.session_state.product_name or 'Product'} — {VOICE_LABELS[voice_idx].split(' —')[0]} — {EL_MODELS.get(el_model_key, el_model_key).split(' —')[0]} — {ts}"
                audio_filename = f"AUDIO_{safe_name}_{VOICE_LABELS[voice_idx].split(' —')[0].lower().replace(' ','_')}_{ts}.mp3"
                st.session_state.audio_history.insert(0, {
                    "label":    audio_label,
                    "filename": audio_filename,
                    "path":     tmp.name,
                    "ts":       ts,
                })
                add_log("Audio ready")
                st.rerun()
            except Exception as e:
                st.error(f"TTS error: {e}")
                add_log(f"TTS error: {e}")

    # Audio history
    if st.session_state.audio_history:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**Audio history**")
        st.caption("Most recent first. Files persist until you reset.")
        for i, entry in enumerate(st.session_state.audio_history):
            if not os.path.exists(entry["path"]):
                continue
            with st.container():
                st.caption(entry["label"])
                col_play, col_dl = st.columns([3, 1])
                with col_play:
                    with open(entry["path"], "rb") as f:
                        audio_bytes = f.read()
                    st.audio(audio_bytes, format="audio/mp3")
                with col_dl:
                    st.download_button(
                        "↓",
                        data=audio_bytes,
                        file_name=entry["filename"],
                        mime="audio/mpeg",
                        key=f"dl_audio_{i}",
                        use_container_width=True,
                        help=f"Download {entry['filename']}",
                    )
                st.markdown("<div style='margin-bottom:4px'></div>", unsafe_allow_html=True)

    # Metrics
    if st.session_state.scenes:
        st.markdown("<br>", unsafe_allow_html=True)
        total_w   = sum(count_words(s["voiceover"]) for s in st.session_state.scenes.values())
        est_sec   = estimate_duration(total_w)
        on_target = 65 <= est_sec <= 85
        m1, m2, m3 = st.columns(3)
        m1.metric("Words", total_w)
        m2.metric("Est. duration", f"{est_sec}s")
        m3.metric("Target", "✓ On target" if on_target else "⚠ Adjust")

    # Activity log
    if st.session_state.log:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("Activity log", expanded=False):
            for entry in reversed(st.session_state.log):
                st.caption(entry)

    # Pipeline
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Pipeline**")
    for icon, label in [
        ("✅", "Upload + extract"),
        ("✅", "Script generation"),
        ("✅", "TTS audio (ElevenLabs)"),
        ("✅", "Save / version control (Sheets)"),
        ("⬜", "FFmpeg video assembly"),
        ("⬜", "16:9 + 9:16 dual output"),
        ("⬜", "YouTube publish"),
    ]:
        st.caption(f"{icon} {label}")

# ═══════════════════════════════════════════════════════════════════════════════
# RIGHT COLUMN
# ═══════════════════════════════════════════════════════════════════════════════
with right:
    if not st.session_state.scenes:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.info("Upload a file and click **Generate script** — or load a saved version from the sidebar.")
    else:
        st.markdown(f"**Script — {st.session_state.product_name or 'Product'}**")
        st.caption("Edit any scene. Click **Save to Sheets** on the left when happy with a version.")


        for scene_key, config in SCENE_CONFIG.items():
            scene = st.session_state.scenes.get(scene_key, {})
            st.markdown(
                f"<div class='scene-header'>{config['label']}"
                f"<span class='timing-badge'> · {config['timing']}</span></div>",
                unsafe_allow_html=True,
            )
            new_vo = st.text_area(
                f"vo_{scene_key}", value=scene.get("voiceover", ""),
                height=100, label_visibility="collapsed", key=f"vo_{scene_key}",
                help="SSML tags supported — see reference above",
            )
            st.session_state.scenes[scene_key]["voiceover"] = new_vo

            wc, target = count_words(new_vo), config["target_words"]
            delta = wc - target
            color = "#2d7a4f" if abs(delta) <= 8 else "#b85c00"
            st.markdown(
                f"<div style='font-size:0.75rem;color:{color};margin-bottom:4px'>"
                f"{wc} words (target ~{target})"
                f"{'  ✓' if abs(delta) <= 8 else f'  {delta:+d}'}</div>",
                unsafe_allow_html=True,
            )

            if scene_key == "features":
                callout_text = "\n".join(scene.get("callouts", []))
                new_callouts = st.text_area("On-screen callouts (one per line)",
                    value=callout_text, height=80, key=f"callouts_{scene_key}",
                    help="These become animated lower-thirds in the video")
                st.session_state.scenes[scene_key]["callouts"] = [
                    c.strip() for c in new_callouts.split("\n") if c.strip()]

            if scene_key == "cta":
                new_cta = st.text_input("CTA line",
                    value=scene.get("cta_line", "Available at uni-trendus.com"),
                    key="cta_line")
                st.session_state.scenes[scene_key]["cta_line"] = new_cta

            st.markdown("---")

        # Export
        script_md  = f"# Video Script — {st.session_state.product_name or 'Product'}\n\n"
        script_md += f"Generated: {time.strftime('%Y-%m-%d %H:%M')}\n\n---\n\n"
        for key, config in SCENE_CONFIG.items():
            s = st.session_state.scenes.get(key, {})
            script_md += f"## {config['label']} ({config['timing']})\n\n"
            script_md += f"**Voiceover:** {s.get('voiceover', '')}\n\n"
            if key == "features" and s.get("callouts"):
                script_md += "**Callouts:**\n" + "".join(f"- {c}\n" for c in s["callouts"]) + "\n"
            if key == "cta":
                script_md += f"**CTA line:** {s.get('cta_line', '')}\n\n"
        total_w    = sum(count_words(s["voiceover"]) for s in st.session_state.scenes.values())
        script_md += f"---\n\nTotal: ~{total_w} words / ~{estimate_duration(total_w)}s\n"
        safe_name  = (st.session_state.product_name or "product").lower().replace(" ", "_")
        date_str   = time.strftime("%Y-%m-%d")

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button("⬇ Download script (.md)", data=script_md,
                file_name=f"VIDEO-SCRIPT_{safe_name}_{date_str}.md",
                mime="text/markdown", use_container_width=True)
        with dl2:
            st.download_button("⬇ Download scene data (.json)",
                data=json.dumps(st.session_state.scenes, indent=2),
                file_name=f"VIDEO-SCENES_{safe_name}_{date_str}.json",
                mime="application/json", use_container_width=True)
