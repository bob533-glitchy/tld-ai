"""
Tiny AI chat site that can fill out the questions in an uploaded .docx.

Run:
    pip install -r requirements.txt
    export GEMINI_API_KEY=your-key
    python app.py
Then open http://localhost:5000
"""
import copy
import json
import os
import re
import uuid

import requests
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from flask import Flask, jsonify, request, send_from_directory

GROQ_KEY = os.environ.get("GROQ_API_KEY", "")
GROQ_BASE = "https://api.groq.com/openai/v1"
GATEWAY_KEY = os.environ.get("AI_GATEWAY_API_KEY", "")  # Vercel AI Gateway
OPENAI_KEY = os.environ.get("OPENAI_API_KEY", "")
API_KEY = os.environ.get("GEMINI_API_KEY", "")  # Gemini
GATEWAY_BASE = "https://ai-gateway.vercel.sh/v1"
# Provider priority: Groq, Vercel AI Gateway, OpenAI, then Gemini.
# MODEL can be a comma-separated list; the next one is tried if one fails.
# Vercel model names look like "provider/model", e.g. "openai/gpt-4o-mini".
if GROQ_KEY:
    _default_models = "llama-3.3-70b-versatile,llama-3.1-8b-instant"
elif GATEWAY_KEY:
    _default_models = "openai/gpt-4o-mini,anthropic/claude-haiku-4.5"
elif OPENAI_KEY:
    _default_models = "gpt-4o-mini,gpt-5-mini"
else:
    _default_models = "gemini-3.8-flash,gemini-3.8-flash-lite"
MODEL = os.environ.get("MODEL", _default_models)
MAX_UPLOAD_MB = 10
MAX_DOC_CHARS = 60_000

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUT_DIR, exist_ok=True)

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

class LLMError(Exception):
    pass


_MODEL_EXPLICIT = bool(os.environ.get("MODEL"))
_groq_cache = []


def groq_models():
    """Ask Groq which chat models this key can use, best-for-free-tier first."""
    global _groq_cache
    if _groq_cache:
        return _groq_cache
    try:
        r = requests.get(
            f"{GROQ_BASE}/models", headers={"Authorization": f"Bearer {GROQ_KEY}"}, timeout=15
        )
        ids = [m["id"] for m in r.json()["data"] if m.get("active", True)]
    except Exception:
        return []
    skip = ("whisper", "tts", "guard", "playai", "orpheus", "embed", "transcri", "safeguard")
    ids = [i for i in ids if not any(w in i.lower() for w in skip)]

    def rank(i):
        i = i.lower()
        for n, word in enumerate(("instant", "versatile", "gpt-oss", "llama", "qwen")):
            if word in i:
                return n
        return 9

    _groq_cache = sorted(ids, key=rank)[:4]
    return _groq_cache


def _model_list():
    if GROQ_KEY and not _MODEL_EXPLICIT:
        found = groq_models()
        if found:
            return found
    return [m.strip() for m in MODEL.split(",") if m.strip()]


def _gemini_request(model, system, messages, max_tokens, json_mode):
    body = {
        "systemInstruction": {"parts": [{"text": system}]},
        "contents": [
            {
                "role": "model" if m["role"] == "assistant" else "user",
                "parts": [{"text": m["content"]}],
            }
            for m in messages
        ],
        "generationConfig": {"maxOutputTokens": max_tokens},
    }
    if json_mode:
        body["generationConfig"]["responseMimeType"] = "application/json"
    r = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        headers={"x-goog-api-key": API_KEY, "Content-Type": "application/json"},
        json=body,
        timeout=35,
    )

    def extract(j):
        return "".join(p.get("text", "") for p in j["candidates"][0]["content"]["parts"])

    return r, extract


def _openai_request(model, system, messages, max_tokens, json_mode, base="https://api.openai.com/v1", key=None):
    body = {
        "model": model,
        "messages": [{"role": "system", "content": system}] + messages,
    }
    if base == "https://api.openai.com/v1":
        body["max_completion_tokens"] = max_tokens
    elif base == GROQ_BASE:
        body["max_completion_tokens"] = min(max_tokens, 4096)
    r = requests.post(
        f"{base}/chat/completions",
        headers={"Authorization": f"Bearer {key or OPENAI_KEY}", "Content-Type": "application/json"},
        json=body,
        timeout=60,
    )
    return r, lambda j: j["choices"][0]["message"]["content"] or ""


def _groq_request(model, system, messages, max_tokens, json_mode):
    return _openai_request(model, system, messages, max_tokens, json_mode, base=GROQ_BASE, key=GROQ_KEY)


def _gateway_request(model, system, messages, max_tokens, json_mode):
    return _openai_request(model, system, messages, max_tokens, json_mode, base=GATEWAY_BASE, key=GATEWAY_KEY)


def call_llm(system, messages, max_tokens=2000, json_mode=False):
    """messages = [{"role": "user"|"assistant", "content": str}, ...]"""
    if GROQ_KEY:
        send, name = _groq_request, "Groq"
    elif GATEWAY_KEY:
        send, name = _gateway_request, "Vercel AI Gateway"
    elif OPENAI_KEY:
        send, name = _openai_request, "OpenAI"
    elif API_KEY:
        send, name = _gemini_request, "Gemini"
    else:
        raise LLMError("Missing GROQ_API_KEY (or AI_GATEWAY_API_KEY / OPENAI_API_KEY / GEMINI_API_KEY) on the server.")

    errors = []
    for model in _model_list():
        try:
            r, extract = send(model, system, messages, max_tokens, json_mode)
        except requests.RequestException:
            errors.append(f"{model} did not answer in time")
            continue
        if r.status_code == 200:
            try:
                return extract(r.json())
            except (KeyError, IndexError, TypeError):
                errors.append(f"{model} returned no answer")
                continue
        try:
            err = r.json()["error"]
            msg = err["message"] if isinstance(err, dict) else str(err)
        except Exception:
            msg = r.text[:200]
        errors.append(f"{model} error {r.status_code}: {msg}")
        if r.status_code not in (404, 429, 500, 503, 504):
            break  # e.g. bad key or bad request: another model won't help
    raise LLMError(f"{name} is unavailable right now. " + (" | ".join(errors) or "no model configured"))


CHAT_SYSTEM = "You are a helpful, concise AI assistant."

DOCX_SYSTEM = """You fill out documents. You will receive the text of a Word document,
split into numbered blocks like [12] some text. Some blocks are questions, prompts or
form fields; others are instructions or context.

Find every question or field that expects an answer and answer it. Return ONLY a JSON
array, no prose and no code fences. Each item:
  {"id": <block number>, "answer": "<your answer>", "placement": "after" | "replace" | "append"}

placement rules:
- "after":   the block is a question; the answer goes in a new paragraph right below it.
             Use this by default.
- "replace": the block is an empty answer line or placeholder (blank, "____", "Answer:")
             or an empty table cell meant for the answer. The answer replaces its text.
- "append":  the block is a label like "Name:" or "Date:" and the answer follows on the same line.

Rules: answer in the same language as the question unless the user's note says otherwise.
Be accurate and concise. If a question needs personal info you cannot know, write a short
placeholder like "[your answer here]". Never invent facts. Do not answer blocks that are
just instructions or headings."""


# ---------------------------------------------------------------- docx helpers
def iter_blocks(doc):
    """Yield (label, Paragraph) for body paragraphs and table-cell paragraphs, in order."""
    body = doc.element.body
    t_index = 0
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield "", Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            table = Table(child, doc)
            seen = set()
            for r, row in enumerate(table.rows):
                for c, cell in enumerate(row.cells):
                    if id(cell._tc) in seen:  # merged cells repeat
                        continue
                    seen.add(id(cell._tc))
                    for p in cell.paragraphs:
                        yield f"(table {t_index} r{r+1}c{c+1}) ", p
            t_index += 1


def extract_blocks(doc):
    blocks = []
    for label, par in iter_blocks(doc):
        blocks.append((label, par))
    return blocks


def blocks_to_text(blocks):
    lines = []
    for i, (label, par) in enumerate(blocks):
        text = par.text.strip()
        lines.append(f"[{i}] {label}{text if text else '(empty)'}")
    return "\n".join(lines)


def set_text(par, text):
    """Write text (may contain newlines) into a paragraph run."""
    lines = text.split("\n")
    run = par.add_run(lines[0])
    for line in lines[1:]:
        run.add_break()
        run.add_text(line)
    return run


def insert_paragraph_after(par, text):
    new_p = copy.deepcopy(par._p)
    for child in list(new_p):
        if child.tag != qn("w:pPr"):
            new_p.remove(child)
    pPr = new_p.find(qn("w:pPr"))
    if pPr is not None:  # don't continue the question's numbering/bullets
        num = pPr.find(qn("w:numPr"))
        if num is not None:
            pPr.remove(num)
    par._p.addnext(new_p)
    new_par = Paragraph(new_p, par._parent)
    run = set_text(new_par, text)
    if par.runs and par.runs[0]._r.rPr is not None:
        run._r.insert(0, copy.deepcopy(par.runs[0]._r.rPr))
    return new_par


def apply_answers(blocks, answers):
    done = 0
    for item in answers:
        try:
            idx = int(item["id"])
            answer = str(item["answer"]).strip()
            placement = item.get("placement", "after")
            _, par = blocks[idx]
        except (KeyError, ValueError, TypeError, IndexError):
            continue
        if not answer:
            continue
        if placement == "replace":
            template = par.runs[0]._r.rPr if par.runs else None
            for r in list(par.runs):
                r._r.getparent().remove(r._r)
            run = set_text(par, answer)
            if template is not None:
                run._r.insert(0, copy.deepcopy(template))
        elif placement == "append":
            sep = "" if par.text.endswith((" ", "\t")) or not par.text else " "
            run = set_text(par, sep + answer)
            if len(par.runs) > 1 and par.runs[0]._r.rPr is not None:
                run._r.insert(0, copy.deepcopy(par.runs[0]._r.rPr))
        else:
            insert_paragraph_after(par, answer)
        done += 1
    return done


def parse_json_array(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()
    start, end = text.find("["), text.rfind("]")
    if start == -1 or end == -1:
        raise ValueError("Model did not return a JSON array")
    return json.loads(text[start : end + 1])


def fill_docx(file_storage, note):
    doc = Document(file_storage)
    blocks = extract_blocks(doc)
    doc_text = blocks_to_text(blocks)
    if len(doc_text) > MAX_DOC_CHARS:
        raise ValueError("That document is too long for this demo.")
    if not any(par.text.strip() for _, par in blocks):
        raise ValueError("That document looks empty.")

    user = f"User note: {note or '(none)'}\n\nDocument:\n{doc_text}"
    text = call_llm(DOCX_SYSTEM, [{"role": "user", "content": user}], max_tokens=8000, json_mode=True)
    answers = parse_json_array(text)
    count = apply_answers(blocks, answers)

    file_id = uuid.uuid4().hex
    doc.save(os.path.join(OUT_DIR, f"{file_id}.docx"))
    return file_id, count


# ---------------------------------------------------------------------- routes
@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.post("/api/chat")
def chat():
    message = (request.form.get("message") or "").strip()
    try:
        history = json.loads(request.form.get("history") or "[]")
    except json.JSONDecodeError:
        history = []
    upload = request.files.get("file")

    try:
        if upload and upload.filename.lower().endswith(".docx"):
            file_id, count = fill_docx(upload, message)
            name = os.path.splitext(upload.filename)[0] + "-filled.docx"
            reply = (
                f"Done. I answered {count} question{'s' if count != 1 else ''} "
                "and saved them into your document."
                if count
                else "I couldn't find any questions to answer in that document."
            )
            return jsonify(
                reply=reply,
                download={"url": f"/download/{file_id}?name={name}", "name": name}
                if count
                else None,
            )
        if upload:
            return jsonify(error="Only .docx files are supported."), 400

        if not message:
            return jsonify(error="Say something first."), 400
        msgs = [
            {"role": m["role"], "content": m["content"]}
            for m in history[-20:]
            if m.get("role") in ("user", "assistant") and m.get("content")
        ]
        msgs.append({"role": "user", "content": message})
        return jsonify(reply=call_llm(CHAT_SYSTEM, msgs))
    except LLMError as e:
        return jsonify(error=str(e)), 500
    except Exception as e:  # keep the demo friendly
        return jsonify(error=str(e)), 500


@app.get("/download/<file_id>")
def download(file_id):
    if not re.fullmatch(r"[0-9a-f]{32}", file_id):
        return "Not found", 404
    name = re.sub(r"[^\w.\- ]", "_", request.args.get("name", "filled.docx"))
    return send_from_directory(OUT_DIR, f"{file_id}.docx", as_attachment=True, download_name=name)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
