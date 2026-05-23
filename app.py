"""AEY — Voice-to-Table MVP web app.

Run locally:
    py app.py

Run with gunicorn (production / Docker):
    gunicorn -w 1 -b 0.0.0.0:5000 app:app --timeout 600

Required env vars:
    ANTHROPIC_API_KEY  Claude API key (https://console.anthropic.com)
    GROQ_API_KEY       Groq API key  (https://console.groq.com)

External dep: ffmpeg on PATH (used to compress audio before sending to Groq).
"""
from __future__ import annotations
import io
import json
import os
import subprocess
import tempfile

from anthropic import Anthropic
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Cm, Pt
from flask import Flask, jsonify, render_template, request, send_file, send_from_directory
from groq import Groq

# ── Prompts ──────────────────────────────────────────────────────────
UK_BAT_PROMPT = (
    "UK bat dusk survey. Species: Common pipistrelle, Soprano pipistrelle, "
    "Nathusius pipistrelle, Noctule, Leisler's bat, Serotine, Brown long-eared bat, "
    "Daubenton's bat, Natterer's bat, Whiskered bat, Brandt's bat, Barbastelle, "
    "Bechstein's bat, Greater horseshoe, Lesser horseshoe, Alcathoe. "
    "Behaviour: foraging, commuting, emerging, social call, re-entry. "
    "Locations: north building side, south building side, vantage point, VP1, VP2."
)

PARSE_INSTRUCTIONS = """You are parsing a UK bat dusk-survey voice transcript into structured rows.

Rules:
- One row per distinct observation. Split combined statements into multiple rows.
- time: HH:MM 24-hour if mentioned (e.g. "21:35"), else empty string.
- species: full UK bat species name. Expand abbreviations: CP/Pip = Common pipistrelle, SP = Soprano pipistrelle, NP = Nathusius pipistrelle, NO = Noctule, BLE = Brown long-eared.
- count: number as string, default "1".
- behaviour: one of Foraging, Commuting, Flying, Emerged, Re-entry, Social, Other.
- location: building side / vantage point / feature where observation occurred.
- direction: compass direction if mentioned (e.g. "SE", "N"), else empty string.
- notes: anything else — roost evidence detail, "Consistent", count caveats.
- If a statement contains a structural detail that looks like roost evidence (emerged from a feature, gap, arch, tile, soffit, etc.), include that detail in notes and prefix with "Possible roost feature: ".

Return ONLY a JSON array. No prose, no markdown fences."""

# ── App ──────────────────────────────────────────────────────────────
app = Flask(__name__)
CLAUDE = Anthropic()  # ANTHROPIC_API_KEY
GROQ = Groq()         # GROQ_API_KEY


def get_audio_duration(path: str) -> float:
    """Audio duration in seconds via ffprobe (ships with ffmpeg)."""
    result = subprocess.run(
        [
            "ffprobe", "-i", path,
            "-show_entries", "format=duration",
            "-v", "quiet", "-of", "csv=p=0",
        ],
        check=True, capture_output=True, text=True,
    )
    return float(result.stdout.strip() or "0")


def compress_audio(input_path: str, output_path: str) -> None:
    """Strip silences, downmix to mono, compress to fit Groq's 25 MB limit.

    The silenceremove filter drops gaps of ≥0.5s below -40 dB. In a typical
    UK bat dusk survey this removes 50-70% of the audio (long quiet stretches
    between observations), which slashes upload time, transcription cost,
    and Whisper drift risk for long recordings.

    Speech quality at mono 32 kbps / 16 kHz is fine — Whisper's native rate.
    """
    subprocess.run(
        [
            "ffmpeg", "-y", "-i", input_path,
            "-af", "silenceremove=stop_periods=-1:stop_duration=0.5:stop_threshold=-40dB",
            "-ac", "1",
            "-b:a", "32k",
            "-ar", "16000",
            output_path,
        ],
        check=True,
        capture_output=True,
    )


GROQ_MAX_BYTES = 25 * 1024 * 1024  # Groq's hard limit on /audio/transcriptions


# ── Routes ───────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory("static", filename)


@app.route("/healthz")
def healthz():
    return "ok"


@app.route("/api/transcribe", methods=["POST"])
def transcribe():
    audio = request.files["audio"]
    suffix = os.path.splitext(audio.filename)[1] or ".m4a"

    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as f:
        audio.save(f.name)
        raw_path = f.name
    compressed_path = raw_path + ".compressed.m4a"

    try:
        original_seconds = get_audio_duration(raw_path)
        compress_audio(raw_path, compressed_path)
        compressed_seconds = get_audio_duration(compressed_path)
        compressed_size = os.path.getsize(compressed_path)

        if compressed_size > GROQ_MAX_BYTES:
            return jsonify({
                "error": (
                    f"Audio is still {compressed_size / 1024 / 1024:.1f} MB after silence stripping. "
                    "Groq's hard limit is 25 MB. Chunked transcription is the next feature; for now, "
                    "split the recording in two and process each half separately."
                ),
            }), 413

        with open(compressed_path, "rb") as f:
            transcription = GROQ.audio.transcriptions.create(
                file=(os.path.basename(compressed_path), f.read()),
                model="whisper-large-v3-turbo",
                prompt=UK_BAT_PROMPT,
                response_format="text",
                language="en",
            )

        text = transcription if isinstance(transcription, str) else transcription.text
        return jsonify({
            "transcript": text.strip(),
            "original_seconds": original_seconds,
            "compressed_seconds": compressed_seconds,
            "compressed_bytes": compressed_size,
        })
    finally:
        for p in (raw_path, compressed_path):
            try:
                os.unlink(p)
            except OSError:
                pass


@app.route("/api/parse", methods=["POST"])
def parse():
    transcript = request.json["transcript"]
    user_msg = f"{PARSE_INSTRUCTIONS}\n\nTranscript:\n\"\"\"\n{transcript}\n\"\"\""

    with CLAUDE.messages.stream(
        model="claude-opus-4-7",
        max_tokens=8192,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": user_msg}],
    ) as stream:
        msg = stream.get_final_message()

    text = "".join(b.text for b in msg.content if b.type == "text").strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    rows = json.loads(text)
    return jsonify({"rows": rows})


@app.route("/api/docx", methods=["POST"])
def docx():
    data = request.json
    rows = data["rows"]
    meta = data.get("meta", {})

    doc = Document()
    doc.add_heading(f"Bat Activity Log — {meta.get('site') or 'Site'}", level=1)

    p = doc.add_paragraph()
    p.add_run("Date: ").bold = True
    p.add_run(f"{meta.get('date', '')}    ")
    p.add_run("Surveyor: ").bold = True
    p.add_run(f"{meta.get('surveyor', '')}    ")
    p.add_run("Survey type: ").bold = True
    p.add_run(f"{meta.get('survey_type', 'Dusk emergence')}")

    headers = ["Time", "Species", "Count", "Behaviour", "Location", "Direction", "Notes"]
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Light Grid Accent 1"

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(10)

    for r_idx, row in enumerate(rows, start=1):
        cells = [
            row.get("time", ""),
            row.get("species", ""),
            row.get("count", "1"),
            row.get("behaviour", ""),
            row.get("location", ""),
            row.get("direction", ""),
            row.get("notes", ""),
        ]
        for c_idx, val in enumerate(cells):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = str(val)
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(10)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    widths_cm = [1.4, 3.2, 1.0, 2.2, 3.6, 1.2, 3.4]
    for col, w in zip(table.columns, widths_cm):
        for cell in col.cells:
            cell.width = Cm(w)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    site = meta.get("site") or "bat_survey"
    safe_site = "".join(c if c.isalnum() else "_" for c in site)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=f"{safe_site}_activity_log.docx",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
