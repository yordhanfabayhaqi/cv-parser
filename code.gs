/**
 * AI CV Parser for Google Sheets
 * 100% AI-generated, 100% human-directed.
 * Directed and iterated by Yordhan Fitrians Akhmad B. based on real HR pain points —
 * refined repeatedly for business fit, user experience, and workflow comfort.
 * Published as a generic, reusable template. Configure everything via the ⚙️ Setup menu.
 *
 * Stack: Google Apps Script + (Claude API or OpenRouter) + Google Drive + Sheets
 *
 * First-time setup:
 *   1. Spreadsheet → Extensions → Apps Script. Paste this into Code.gs.
 *   2. Editor → "Services" (+ icon) → add "Drive API" (identifier: Drive, v3).
 *      This enables PDF → text conversion (with OCR for image-based PDFs).
 *   3. Save, reload the spreadsheet, then use the "⚙️ Setup" menu to configure everything.
 *
 * No API key, token, or resource ID is stored in this file. Every credential and ID
 * is entered once through the ⚙️ Setup menu and persisted in Script Properties.
 */

// ────────────────────────────────────────────────────────────────────────────
// CONSTANTS (non-secret defaults only)
// ────────────────────────────────────────────────────────────────────────────
const DEFAULT_SHEET_NAME = 'CV Parser';

const PROVIDER_CLAUDE     = 'claude';
const PROVIDER_OPENROUTER = 'openrouter';

const CLAUDE_URL      = 'https://api.anthropic.com/v1/messages';
const OPENROUTER_URL  = 'https://openrouter.ai/api/v1/chat/completions';
const ANTHROPIC_VER   = '2023-06-01';

const DEFAULT_CLAUDE_MODEL     = 'claude-haiku-4-5-20251001';
const DEFAULT_OPENROUTER_MODEL = 'google/gemini-2.5-flash';

const MAX_RUN_SECONDS  = 300;     // stop gracefully before Apps Script's 6-min limit
const MAX_RETRIES      = 4;
const POLITE_DELAY_MS  = 1200;    // pause between successful calls

// ── DEMONSTRATION OUTPUT SCHEMA ─────────────────────────────────────────────
// These headers are a generic EXAMPLE only and are NOT tied to any real hiring
// process. Whoever adapts this template should replace them with the columns
// their own workflow needs, and keep the JSON fields in parseRecordText_() in
// sync. This is a demonstration schema, not a production mapping.
const EXAMPLE_HEADERS = [
  'No.',
  'Candidate Name',
  'Summary',
  'Latest Role Title',
  'Latest Role Description',
  'Years of Experience',
  'Contact Number',
  'Contact Email',
  'Source File',
  'LinkedIn',
  'GitHub'
];

// ────────────────────────────────────────────────────────────────────────────
// CONFIG LAYER  (Script Properties — nothing hardcoded)
// ────────────────────────────────────────────────────────────────────────────
function getConfig_() {
  const p = PropertiesService.getScriptProperties();
  const provider = p.getProperty('PROVIDER') || PROVIDER_CLAUDE;

  const cfg = {
    provider:        provider,
    folderId:        p.getProperty('FOLDER_ID')         || '',
    sheetName:       p.getProperty('SHEET_NAME')        || DEFAULT_SHEET_NAME,
    claudeKey:       p.getProperty('CLAUDE_API_KEY')    || '',
    openrouterKey:   p.getProperty('OPENROUTER_API_KEY')|| '',
    claudeModel:     p.getProperty('CLAUDE_MODEL')      || DEFAULT_CLAUDE_MODEL,
    openrouterModel: p.getProperty('OPENROUTER_MODEL')  || DEFAULT_OPENROUTER_MODEL
  };
  cfg.activeKey   = provider === PROVIDER_CLAUDE ? cfg.claudeKey   : cfg.openrouterKey;
  cfg.activeModel = provider === PROVIDER_CLAUDE ? cfg.claudeModel : cfg.openrouterModel;
  return cfg;
}

function saveConfig_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

// ────────────────────────────────────────────────────────────────────────────
// MENU
// ────────────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Setup')
    .addItem('1. Set API Provider',   'setProvider')
    .addItem('2. Set API Key',        'setApiKey')
    .addItem('3. Set Model',          'setModel')
    .addItem('4. Set Drive Folder',   'setDriveFolder')
    .addItem('5. Initialize Sheet',   'initializeSheet')
    .addSeparator()
    .addItem('Show Current Settings', 'showSettings')
    .addItem('Test API Connection',   'testApiConnection')
    .addSeparator()
    .addItem('▶ Import & Parse Files', 'importAndParseCVs')
    .addItem('♻️ Reset Processed Log', 'resetProcessedLog')
    .addToUi();
}

// ────────────────────────────────────────────────────────────────────────────
// SETUP HANDLERS
// ────────────────────────────────────────────────────────────────────────────
function setProvider() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Set API Provider',
    'Current: ' + getConfig_().provider + '\n\n' +
    'Type one of:\n' +
    '  • claude      — Anthropic API direct\n' +
    '  • openrouter  — OpenRouter (any model)',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const val = res.getResponseText().trim().toLowerCase();
  if (val !== PROVIDER_CLAUDE && val !== PROVIDER_OPENROUTER) {
    ui.alert('Must be "claude" or "openrouter".');
    return;
  }
  saveConfig_('PROVIDER', val);
  ui.alert('✓ Provider set to: ' + val + '\n\nNow set the API key and (optionally) the model.');
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const provider = getConfig_().provider;
  const propKey  = provider === PROVIDER_CLAUDE ? 'CLAUDE_API_KEY' : 'OPENROUTER_API_KEY';
  const expected = provider === PROVIDER_CLAUDE ? 'sk-ant-' : 'sk-or-';
  const res = ui.prompt(
    'Set API Key for ' + provider,
    'Paste your ' + (provider === PROVIDER_CLAUDE ? 'Anthropic' : 'OpenRouter') +
    ' API key (starts with "' + expected + '"):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key.startsWith(expected)) {
    ui.alert('That does not look like a valid key for ' + provider + '.');
    return;
  }
  saveConfig_(propKey, key);
  ui.alert('✓ Key saved for ' + provider + '. (Stored in Script Properties, not in code.)');
}

function setModel() {
  const ui = SpreadsheetApp.getUi();
  const cfg = getConfig_();
  const provider = cfg.provider;
  const propKey  = provider === PROVIDER_CLAUDE ? 'CLAUDE_MODEL' : 'OPENROUTER_MODEL';
  const current  = provider === PROVIDER_CLAUDE ? cfg.claudeModel : cfg.openrouterModel;

  const suggestions = provider === PROVIDER_CLAUDE
    ? '  • claude-haiku-4-5-20251001  (recommended — cheap & fast)\n' +
      '  • claude-sonnet-4-6           (more accurate, slower)\n' +
      '  Use the exact model IDs from your provider\'s current docs.'
    : '  • google/gemini-2.5-flash     (recommended — very cheap & fast)\n' +
      '  • anthropic/claude-haiku-4-5\n' +
      '  • openai/gpt-4o-mini\n' +
      '  • meta-llama/llama-3.3-70b-instruct';

  const res = ui.prompt(
    'Set Model (' + provider + ')',
    'Current: ' + current + '\n\nSuggestions:\n' + suggestions + '\n\nEnter model ID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const model = res.getResponseText().trim();
  if (!model) return;
  saveConfig_(propKey, model);
  ui.alert('✓ Model set to: ' + model);
}

function setDriveFolder() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Set Drive Folder',
    'Paste the Drive folder URL or folder ID that contains the source PDFs:',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const input = res.getResponseText().trim();
  const match = input.match(/folders\/([a-zA-Z0-9_-]+)/);
  const folderId = match ? match[1] : input;
  try {
    const folder = DriveApp.getFolderById(folderId);
    saveConfig_('FOLDER_ID', folderId);
    ui.alert('✓ Folder saved.\nName: ' + folder.getName());
  } catch (e) {
    ui.alert('✗ Could not access that folder. Check the ID/URL and sharing permissions.');
  }
}

function initializeSheet() {
  const ui  = SpreadsheetApp.getUi();
  const cfg = getConfig_();
  const ss  = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) sheet = ss.insertSheet(cfg.sheetName);

  // Writes the generic EXAMPLE headers (see EXAMPLE_HEADERS note above).
  sheet.getRange(1, 1, 1, EXAMPLE_HEADERS.length)
       .setValues([EXAMPLE_HEADERS])
       .setFontWeight('bold');
  sheet.setFrozenRows(1);
  ui.alert('✓ Sheet "' + cfg.sheetName + '" initialized with example headers.\n' +
           'Adapt these columns to your own workflow as needed.');
}

function showSettings() {
  const cfg = getConfig_();
  const processed = JSON.parse(PropertiesService.getScriptProperties().getProperty('PROCESSED_FILES') || '{}');

  let folderName = '(not set)';
  if (cfg.folderId) {
    try { folderName = DriveApp.getFolderById(cfg.folderId).getName(); }
    catch (e) { folderName = '(invalid or inaccessible ID)'; }
  }

  const msg =
    'Provider:  ' + cfg.provider + '\n' +
    'Sheet:     ' + cfg.sheetName + '\n' +
    'Folder:    ' + folderName + '\n\n' +
    '— Claude —\n' +
    '  Key:   ' + (cfg.claudeKey ? '✓ set' : '✗ not set') + '\n' +
    '  Model: ' + cfg.claudeModel + '\n\n' +
    '— OpenRouter —\n' +
    '  Key:   ' + (cfg.openrouterKey ? '✓ set' : '✗ not set') + '\n' +
    '  Model: ' + cfg.openrouterModel + '\n\n' +
    'Already parsed: ' + Object.keys(processed).length + ' files';

  SpreadsheetApp.getUi().alert('Setup — Current Settings', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function resetProcessedLog() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('Reset processed log?',
    'Every PDF becomes eligible for re-parsing on the next run. Spreadsheet rows are NOT deleted.',
    ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteProperty('PROCESSED_FILES');
    ui.alert('✓ Cleared.');
  }
}

function testApiConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = callLLM_('Reply with the single word: OK');
    ui.alert('✓ Connection OK.\nProvider: ' + getConfig_().provider + '\nResponse: ' + result.slice(0, 100));
  } catch (e) {
    ui.alert('✗ Error:\n\n' + e.message);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// DEDUPLICATION HELPERS
// ────────────────────────────────────────────────────────────────────────────
function normalizeName_(name) {
  if (!name) return '';
  return String(name).toLowerCase().trim().replace(/\s+/g, ' ');
}

function getExistingNames_(sheet) {
  const lastRow = sheet.getLastRow();
  const set = {};
  if (lastRow < 2) return set;
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // column B (Candidate Name) from row 2
  for (let i = 0; i < values.length; i++) {
    const norm = normalizeName_(values[i][0]);
    if (norm) set[norm] = (i + 2); // store row number for reference
  }
  return set;
}

// ────────────────────────────────────────────────────────────────────────────
// MAIN — Import & Parse
// ────────────────────────────────────────────────────────────────────────────
function importAndParseCVs() {
  const ui  = SpreadsheetApp.getUi();
  const cfg = getConfig_();

  if (!cfg.folderId) { ui.alert('Set the Drive folder first (⚙️ Setup → 4).'); return; }
  if (!cfg.activeKey) {
    ui.alert('No API key set for "' + cfg.provider + '". Use ⚙️ Setup → 2.');
    return;
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(cfg.sheetName);
  if (!sheet) { ui.alert('Sheet "' + cfg.sheetName + '" not found. Run ⚙️ Setup → 5 first.'); return; }

  const folder    = DriveApp.getFolderById(cfg.folderId);
  const processed = JSON.parse(PropertiesService.getScriptProperties().getProperty('PROCESSED_FILES') || '{}');
  const existing  = getExistingNames_(sheet);
  const startTime = Date.now();

  const files = folder.getFilesByType(MimeType.PDF);
  let parsed = 0, skipped = 0, errors = 0, duplicates = 0;
  const errorLog = [];
  const dupLog   = [];

  while (files.hasNext()) {
    if ((Date.now() - startTime) / 1000 > MAX_RUN_SECONDS) {
      errorLog.push('⏱ Time limit — run again to continue with remaining files.');
      break;
    }

    const file   = files.next();
    const fileId = file.getId();
    if (processed[fileId]) { skipped++; continue; }

    try {
      // 1. Extract text (OCR-capable conversion)
      const cvText = pdfToText_(file);
      if (!cvText || cvText.length < 50) {
        throw new Error('Empty/too-short extracted text (' + cvText.length + ' chars)');
      }

      // 2. Send to the configured LLM
      const data = parseRecordText_(cvText);

      // 3. Duplicate check on Candidate Name (column B)
      const normName = normalizeName_(data.name);
      if (normName && existing[normName]) {
        duplicates++;
        dupLog.push('"' + data.name + '" (already in row ' + existing[normName] + ') ← ' + file.getName());
        processed[fileId] = new Date().toISOString();
        continue;
      }

      // 4. Append row (order matches EXAMPLE_HEADERS)
      const nextRow = Math.max(sheet.getLastRow() + 1, 2);
      sheet.getRange(nextRow, 1, 1, EXAMPLE_HEADERS.length).setValues([[
        nextRow - 1,
        data.name || '',
        data.biography || '',
        data.latest_experience_title || '',
        data.latest_experience_description || '',
        data.years_of_experience || '',
        data.contact_number || '',
        data.contact_email || '',
        file.getUrl(),
        data.linkedin || '',
        data.github || ''
      ]]);

      if (normName) existing[normName] = nextRow;
      processed[fileId] = new Date().toISOString();
      parsed++;

      if (parsed % 5 === 0) {
        saveConfig_('PROCESSED_FILES', JSON.stringify(processed));
        SpreadsheetApp.flush();
      }

      Utilities.sleep(POLITE_DELAY_MS);
    } catch (e) {
      errors++;
      errorLog.push(file.getName() + ': ' + e.message);
      Logger.log(file.getName() + ' — ' + e.message);
    }
  }

  saveConfig_('PROCESSED_FILES', JSON.stringify(processed));

  let summary = 'Parsed:             ' + parsed +
                '\nDuplicates skipped: ' + duplicates +
                '\nAlready done:       ' + skipped +
                '\nErrors:             ' + errors +
                '\nProvider:           ' + cfg.provider;

  if (dupLog.length) {
    summary += '\n\n— Duplicates —\n' + dupLog.slice(0, 15).join('\n');
    if (dupLog.length > 15) summary += '\n… and ' + (dupLog.length - 15) + ' more.';
  }
  if (errorLog.length) {
    summary += '\n\n— Errors —\n' + errorLog.slice(0, 8).join('\n');
    if (errorLog.length > 8) summary += '\n… and ' + (errorLog.length - 8) + ' more (View → Logs).';
  }
  ui.alert('Import Complete', summary, ui.ButtonSet.OK);
}

// ────────────────────────────────────────────────────────────────────────────
// PDF → TEXT  (Advanced Drive Service — enable "Drive API" in Services)
// ────────────────────────────────────────────────────────────────────────────
function pdfToText_(pdfFile) {
  let tempId = null;
  try {
    const blob = pdfFile.getBlob();
    const doc = Drive.Files.create(
      { name: 'parser_tmp_' + Utilities.getUuid(), mimeType: 'application/vnd.google-apps.document' },
      blob,
      { ocrLanguage: 'en' }   // change OCR language if your source files differ
    );
    tempId = doc.id;
    const text = DocumentApp.openById(tempId).getBody().getText();
    return text.trim();
  } catch (e) {
    throw new Error('PDF→text failed (is the Drive API service enabled?): ' + e.message);
  } finally {
    if (tempId) {
      try { DriveApp.getFileById(tempId).setTrashed(true); } catch (_) {}
    }
  }
}

// ────────────────────────────────────────────────────────────────────────────
// LLM PARSING
// ────────────────────────────────────────────────────────────────────────────
function parseRecordText_(cvText) {
  const prompt =
    'You are extracting structured data from a CV/resume for a hiring pipeline. ' +
    'Return ONLY a single valid JSON object — no markdown fences, no commentary. ' +
    'Use exactly these fields:\n\n' +
    '{\n' +
    '  "name": "Full name as written on the CV",\n' +
    '  "biography": "Concise 2-3 sentence professional summary. Use the CV\'s own summary if present (paraphrased), otherwise synthesize from background and skills.",\n' +
    '  "latest_experience_title": "Most recent role as \'Title at Company\'",\n' +
    '  "latest_experience_description": "2-3 sentences paraphrasing key responsibilities and achievements in the most recent role. Do not list bullets verbatim.",\n' +
    '  "years_of_experience": "Total years of professional experience as a NUMBER (e.g. 3, 1.5, 0.5). Sum all paid roles and internships. Round to 1 decimal.",\n' +
    '  "contact_number": "Phone number with country code",\n' +
    '  "contact_email": "Primary email",\n' +
    '  "linkedin": "LinkedIn URL or username, empty string if absent",\n' +
    '  "github": "GitHub URL or username, empty string if absent"\n' +
    '}\n\n' +
    'Rules:\n' +
    '- Empty string for missing fields (0 for years_of_experience).\n' +
    '- Do not invent data.\n' +
    '- Output must parse with JSON.parse with no preprocessing.\n\n' +
    'CV TEXT:\n---\n' + cvText + '\n---';

  const raw = callLLM_(prompt);

  let text = raw.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
  const first = text.indexOf('{');
  const last  = text.lastIndexOf('}');
  if (first === -1 || last === -1) {
    throw new Error('No JSON in LLM response: ' + raw.slice(0, 200));
  }
  return JSON.parse(text.slice(first, last + 1));
}

function callLLM_(prompt) {
  const provider = getConfig_().provider;
  if (provider === PROVIDER_CLAUDE)     return callClaude_(prompt);
  if (provider === PROVIDER_OPENROUTER) return callOpenRouter_(prompt);
  throw new Error('Unknown provider: ' + provider);
}

function callClaude_(prompt) {
  const cfg = getConfig_();
  const payload = {
    model: cfg.claudeModel,
    max_tokens: 1500,
    messages: [{ role: 'user', content: prompt }]
  };
  const data = fetchWithRetry_(CLAUDE_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': cfg.claudeKey, 'anthropic-version': ANTHROPIC_VER },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  let out = '';
  for (const b of data.content) { if (b.type === 'text') out += b.text; }
  return out;
}

function callOpenRouter_(prompt) {
  const cfg = getConfig_();
  const payload = {
    model: cfg.openrouterModel,
    max_tokens: 1500,
    messages: [{ role: 'user', content: prompt }]
  };
  const data = fetchWithRetry_(OPENROUTER_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + cfg.openrouterKey,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'CV Parser Template'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (!data.choices || !data.choices[0]) {
    throw new Error('OpenRouter unexpected response: ' + JSON.stringify(data).slice(0, 400));
  }
  return data.choices[0].message.content || '';
}

function fetchWithRetry_(url, options) {
  let lastError = '';
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code === 200) return JSON.parse(body);

    if (code === 429 || code === 529 || code >= 500) {
      const headers = response.getAllHeaders();
      const retryAfter = parseInt(headers['retry-after'] || headers['Retry-After'] || '0', 10);
      const sleepMs = retryAfter > 0 ? (retryAfter * 1000) : Math.min(60000, Math.pow(2, attempt) * 5000);
      Logger.log('HTTP ' + code + ' — sleeping ' + sleepMs + 'ms (attempt ' + (attempt + 1) + ')');
      Utilities.sleep(sleepMs);
      lastError = 'HTTP ' + code + ': ' + body.slice(0, 300);
      continue;
    }

    throw new Error('HTTP ' + code + ': ' + body.slice(0, 500));
  }
  throw new Error('Max retries exceeded — ' + lastError);
}
