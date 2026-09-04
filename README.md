# Document Formatting Tool — Hosting-Ready Package

This package wraps your original formatting pipeline in a web
backend + browser frontend so anyone in the office can use it without
VS Code, Python, or the command line.

**Nothing about the pipeline's logic was changed.** `backend/pipeline/`
contains your original files exactly as you shared them
(`format_doc.py`, `policy_extractor.py`, `structure_classifier.py`,
`formatting_engine.py`, `integrity_validator.py`, `toc_card_builder.py`,
`docx_fast.py`, `ooxml_fonts.py`). The only new file, `backend/app.py`,
calls the same `format_doc.run(...)` function your CLI command already
calls — verified to produce byte-for-byte identical document content to
running `python format_doc.py --reference ... --input ... --output ...`
directly.

## Structure

```
formatting-tool/
├── backend/
│   ├── app.py              — FastAPI wrapper (new)
│   ├── requirements.txt    — pinned dependencies
│   └── pipeline/           — your original scripts (unchanged)
├── frontend/
│   └── index.html          — simple upload page (new)
└── deploy/
    ├── DEPLOY.md                     — step-by-step Ubuntu VM setup
    ├── formatting-tool.service       — systemd unit
    └── nginx_formatting-tool.conf    — reverse-proxy config
```

## Try it locally first (optional, recommended)

```bash
cd backend
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
uvicorn app:app --reload --port 8000
```
Then open `http://127.0.0.1:8000/` in a browser, upload a reference doc
and a target doc, and confirm the formatted result downloads correctly.

## Deploy to your office's Ubuntu VM

Follow `deploy/DEPLOY.md` step by step — it's written for exactly the
free, self-managed Ubuntu VM your IT person mentioned, using systemd +
Nginx (no cloud billing, no paid services).

## API reference (for anyone extending this later)

- `GET /api/health` → `{"status": "ok"}`
- `POST /api/format` — multipart form fields `reference` and `target`
  (both `.docx`) → returns the formatted `.docx` as a file download, or
  a JSON error body with a `detail` message on failure (e.g. content
  drift detected, non-.docx upload, file too large).

## What's next (optional upgrades, not required to launch)

- **SharePoint/Power Automate integration** and a **native Word
  add-in** are natural next steps once this is validated with a few
  users — see the earlier roadmap discussion for how those layer on
  top of this same backend without any pipeline changes.
- **Auth**: currently open to anyone who can reach the URL. See
  `deploy/DEPLOY.md` section 8 for free ways to restrict it to your
  organization (LAN-only, Basic Auth, or Entra ID via OAuth2 Proxy).
