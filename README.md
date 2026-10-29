# My AI

A chat site (like Bull-33) that can also **fill out the questions in a .docx** you attach.

## Run it

```bash
pip install -r requirements.txt
export GEMINI_API_KEY=your-key     # free at aistudio.google.com/apikey
python app.py
```

Open http://localhost:5000. Type to chat, or click **+**, pick a .docx, and send. Add a note
(e.g. "answer in French") if you want. You get a download link for the filled document.

## How the docx part works

1. `python-docx` reads every paragraph and table cell and numbers them `[0]`, `[1]`, ...
2. The numbered text goes to Gemini, which returns JSON: which block is a question, the answer,
   and where it goes (`after` the question, `replace` a blank line, or `append` after a label).
3. `apply_answers()` writes each answer into the original file, keeping its formatting, and saves it to `outputs/`.

## Put it online

Deploy the folder to Render, Railway, Fly.io, or any host that runs Python. Set
`GEMINI_API_KEY` as an environment variable there (never put it in the HTML), and run with
`gunicorn app:app` instead of `python app.py` (add `gunicorn` to requirements).

## Before you share it publicly

- Add rate limiting or a login, or strangers can run up your API bill.
- Clean out the `outputs/` folder regularly; filled documents may contain private info.
- Change the model with the `MODEL` environment variable (default `gemini-3.8-flash`).
