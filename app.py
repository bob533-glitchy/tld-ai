"""
Tiny AI chat site that can fill out the questions in an uploaded .docx.

Run:
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY=sk-ant-...
    python app.py
Then open http://localhost:5000
"""
import copy
import json
import os
import re
import uuid

import anthropic
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from flask import Flask, jsonify, request, send_from_directory

MODEL = os.environ.get("MODEL", "claude-sonnet-5-5")
MAX_UPLOAD_MB = 10
MAX_DOC_CHARS = 60_000

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUT_DIR, exist_ok=True)

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from the environment

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
    resp = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        system=DOCX_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )
    answers = parse_json_array(resp.content[0].text)
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
        resp = client.messages.create(
            model=MODEL, max_tokens=2000, system=CHAT_SYSTEM, messages=msgs
        )
        return jsonify(reply=resp.content[0].text)
    except anthropic.AuthenticationError:
        return jsonify(error="Missing or invalid ANTHROPIC_API_KEY on the server."), 500
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
