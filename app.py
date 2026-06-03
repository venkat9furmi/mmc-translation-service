import os
import io
import re
import google.generativeai as genai
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt

app = Flask(__name__)

# ─── CONFIG ────────────────────────────────────────────────────────────────────
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

SKIP_EXACT = {
    "commentary", "time", "picture:sound", "picture: sound",
    "end", "[no commentary]", "time in", "clock into part 2",
    "break bumper out", "break bumper in", "still to come",
    "opening titles", "menu", "closing sequence",
    "semi-final, fixtures", "ucl final, fixture", "uwcl final, fixture",
    "no", "item", "dur", "r/t", "",
}

# ─── HELPERS ───────────────────────────────────────────────────────────────────

def cell_text(cell):
    return "\n".join(p.text for p in cell.paragraphs if p.text.strip())


def should_skip(text):
    t = text.strip()
    if not t:
        return True
    if t.lower() in SKIP_EXACT:
        return True
    if re.match(r"^\d{2}:\d{2}:\d{2}$", t):
        return True
    return False


def build_prompt(numbered_cells, language):
    lang_name = "German" if language == "de" else "French"
    prefix    = "DE:" if language == "de" else "FR:"
    items     = "\n\n".join(
        f"[{i+1}] {text}" for i, (_, text) in enumerate(numbered_cells)
    )
    return f"""You are a senior UEFA TV magazine script translator for MMC Sport.
Translate the following script lines to {lang_name} for professional voice-over broadcast.

RULES:
- Keep the original English line exactly as-is
- Add the {lang_name} translation on the NEXT line, starting with "{prefix} "
- Natural broadcast {lang_name} — NOT word-for-word translation
- Short spoken sentences, active voice, rhythm suitable for VO
- Do NOT translate: [TRANSCRIPT MISSING], names, scores, technical cues
- If untranslatable → write: {prefix} [NO TRANSLATION NEEDED]

OUTPUT FORMAT — follow exactly:
[1] Original English line
{prefix} {lang_name} translation

[2] Original English line
{prefix} {lang_name} translation

INPUT LINES:
{items}"""


def call_gemini(prompt, api_key):
    genai.configure(api_key=api_key)
    model    = genai.GenerativeModel("gemini-2.0-flash")
    response = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            temperature=0.2,
            max_output_tokens=8192,
        ),
    )
    return response.text


def parse_gemini_output(output, keys):
    translations = {}
    blocks = re.split(r"\[(\d+)\]", output)
    i = 1
    while i < len(blocks) - 1:
        num_str = blocks[i].strip()
        content = blocks[i + 1].strip() if i + 1 < len(blocks) else ""
        try:
            idx = int(num_str) - 1
            if 0 <= idx < len(keys):
                translations[keys[idx]] = content
        except ValueError:
            pass
        i += 2
    return translations


def translate_batch(batch_items, language, api_key):
    prompt = build_prompt(batch_items, language)
    output = call_gemini(prompt, api_key)
    keys   = [k for k, _ in batch_items]
    return parse_gemini_output(output, keys)


def add_translation_to_cell(cell, translated_block):
    lines = translated_block.strip().splitlines()
    translation_lines = [l for l in lines if re.match(r"^(DE:|FR:)", l.strip())]
    if not translation_lines:
        return
    translation_text = "\n".join(translation_lines)
    para = cell.add_paragraph()
    run  = para.add_run(translation_text)
    run.font.size = Pt(10)


# ─── ROUTES ────────────────────────────────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "MMC Sport Translation Service"})


@app.route("/translate", methods=["POST"])
def translate():
    try:
        api_key  = request.form.get("api_key") or GEMINI_API_KEY
        language = request.form.get("language", "de").lower()

        if not api_key:
            return jsonify({"error": "GEMINI_API_KEY missing"}), 400
        if language not in ("de", "fr"):
            return jsonify({"error": "language must be 'de' or 'fr'"}), 400
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400

        file = request.files["file"]
        if not file.filename.lower().endswith(".docx"):
            return jsonify({"error": "Only .docx files are supported"}), 400

        doc_bytes = file.read()
        doc       = Document(io.BytesIO(doc_bytes))

        # Collect cells to translate
        cells_to_translate = {}
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                n_cols = len(row.cells)
                if t_idx == 0 and n_cols >= 2:
                    text = cell_text(row.cells[1])
                    if not should_skip(text) and "FEATURE" in text.upper():
                        cells_to_translate[f"T{t_idx}R{r_idx}C1"] = text
                if t_idx >= 1 and n_cols >= 4:
                    text = cell_text(row.cells[3])
                    if not should_skip(text):
                        cells_to_translate[f"T{t_idx}R{r_idx}C3"] = text

        if not cells_to_translate:
            return jsonify({"error": "No translatable content found"}), 400

        # Translate in batches of 15
        all_translations = {}
        items      = list(cells_to_translate.items())
        batch_size = 15

        for i in range(0, len(items), batch_size):
            batch  = items[i : i + batch_size]
            result = translate_batch(batch, language, api_key)
            all_translations.update(result)

        # Apply translations
        doc = Document(io.BytesIO(doc_bytes))
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                n_cols = len(row.cells)
                if t_idx == 0 and n_cols >= 2:
                    key = f"T{t_idx}R{r_idx}C1"
                    if key in all_translations:
                        add_translation_to_cell(row.cells[1], all_translations[key])
                if t_idx >= 1 and n_cols >= 4:
                    key = f"T{t_idx}R{r_idx}C3"
                    if key in all_translations:
                        add_translation_to_cell(row.cells[3], all_translations[key])

        # Save and return
        output      = io.BytesIO()
        doc.save(output)
        output.seek(0)

        lang_tag    = "GER" if language == "de" else "FRA"
        base_name   = file.filename.replace(".docx", "")
        output_name = f"{base_name}_{lang_tag}_translated.docx"

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=output_name,
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─── MAIN ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
