# AI CV Parser for Google Sheets

A Google Apps Script template that reads PDF resumes from a Google Drive folder,
extracts structured fields with an LLM (Anthropic Claude or any OpenRouter model),
de-duplicates by candidate name, and writes the results into a spreadsheet.

## What it does
- Converts PDFs to text (OCR-capable) via the Advanced Drive Service.
- Sends each document to your chosen LLM and returns a fixed JSON schema.
- Skips already-processed files and duplicate names across runs.
- Handles rate limits with retry + exponential backoff.

## Setup
1. Open your spreadsheet → **Extensions → Apps Script**, paste `Code.gs`.
2. In the editor, **Services (+)** → add **Drive API** (identifier `Drive`, v3).
3. Save and reload the spreadsheet. A **⚙️ Setup** menu appears.
4. Configure everything through that menu — nothing is edited in code:
   - **1. Set API Provider** — `claude` or `openrouter`
   - **2. Set API Key** — stored in Script Properties, never in the source
   - **3. Set Model** — use a current, valid model ID from your provider
   - **4. Set Drive Folder** — paste the folder URL or ID with your PDFs
   - **5. Initialize Sheet** — writes the example header row
5. Run **▶ Import & Parse Files**.

## A note on the data schema
The output columns and the parsing prompt in this template are a **generic,
simplified example**. The author works with confidential HR data and deliberately
does **not** publish production column mappings, role definitions, or business
logic. Anyone adapting this must supply their own real schema and keep the sheet
headers in sync with the JSON fields in `parseRecordText_()`.

## Configuration & secrets
All keys and resource IDs live in **Script Properties** via the ⚙️ Setup menu.
No API key, token, folder ID, or spreadsheet ID is hardcoded anywhere in this repo.

## Attribution
100% AI-generated, 100% human-directed.
