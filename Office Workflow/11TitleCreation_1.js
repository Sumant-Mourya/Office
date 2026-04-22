/**
 * CONFIGURATION
 */
const API_KEY = '';
const MODEL_NAME = 'gemini-2.5-flash';
const SAFE_BATCH_SIZE = 15;
const GENERATION_PARALLEL_CHUNK_SIZE = 3;
const GENERATION_MAX_RETRIES = 3;
const GENERATION_RETRY_BASE_DELAY_MS = 1500;
const INLINE_IMAGE_MAX_BYTES = 7 * 1024 * 1024;

const GENERATE_ENDPOINT_BASE = '';
const TRIGGER_DELAY_MS = 60000;
const RUNNING_FLAG_KEY = 'IS_TRIGGER_RUNNING';
const ENABLE_DETAILED_LOGS = true;
const LOG_SHEET_NAME = 'log';
const CONFIG_SHEET_NAME = 'configuration';
const TARGET_SHEET_NAME = 'Title Creation';

const EXPECTED_TITLE_HEADERS = [
  'Parent SKU',
  'designcode',
  'item-name',
  'imageurl',
  'Usersay',
  'Title',
  'Description',
  'Bullet point 1',
  'Bullet point 2',
  'Bullet point 3',
  'Bullet point 4',
  'Bullet point 5',
  'Keyword',
  'weight'
];

const HEADER_LIST_TEXT = EXPECTED_TITLE_HEADERS.join(',');
const TITLE_COL_IDX = Object.freeze({
  ITEM_NAME: 2,
  IMAGE_URL: 3,
  TITLE: 5
});

/**
 * INITIAL START
 */
function startAutomatedProcess() {
  const props = PropertiesService.getScriptProperties();
  const isRunning = props.getProperty(RUNNING_FLAG_KEY) === 'true';

  if (isRunning) {
    const removed = _clearMainTriggers();
    props.setProperty(RUNNING_FLAG_KEY, 'false');
    _log(`TRIGGER STOP: processBatch trigger stopped. Removed ${removed} trigger(s).`);
    return;
  }

  const runtime = _getRuntimeContext(true);
  if (!runtime) {
    return;
  }

  const headerCheck = _validateTitleCreationHeaders(runtime.targetSheet);
  if (!headerCheck.ok) {
    _notifyUser(
      `${headerCheck.message}\n\nPlease correct sheet format in this format:\n${HEADER_LIST_TEXT}\n\nCurrent headers:\n${headerCheck.actualHeaderList}`,
      true,
      'Title Generation - Invalid Sheet Format'
    );
    return;
  }

  const nextPendingIndex = _findNextPendingRowIndex(runtime.targetSheet);
  if (nextPendingIndex === -1) {
    _notifyUser('No pending rows found to process.', true, 'Title Generation');
    return;
  }
  props.setProperty('START_INDEX', String(nextPendingIndex));

  _registerMainTrigger();
  props.setProperty(RUNNING_FLAG_KEY, 'true');
  _log('TRIGGER START: processBatch trigger started. Runs every 60 seconds.');
}

/**
 * MAIN PROCESSING FUNCTION (main trigger)
 */
function processBatch() {
  const runLock = LockService.getScriptLock();
  if (!runLock.tryLock(0)) {
    _log('TRIGGER SKIP: Previous processBatch run is still active. Current trigger execution cancelled.');
    return;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(RUNNING_FLAG_KEY) !== 'true') {
      _log('TRIGGER SKIP: Running flag is false. Exiting current trigger execution.');
      return;
    }

    const runtime = _getRuntimeContext(false);
    if (!runtime) {
      props.setProperty(RUNNING_FLAG_KEY, 'false');
      _clearMainTriggers();
      return;
    }

    const sheet = runtime.targetSheet;
    const languageMode = runtime.languageMode;
    const headerCheck = _validateTitleCreationHeaders(sheet);
    if (!headerCheck.ok) {
      _notifyUser(
        `${headerCheck.message}\n\nPlease correct sheet format in this format:\n${HEADER_LIST_TEXT}\n\nCurrent headers:\n${headerCheck.actualHeaderList}`,
        true,
        'Title Generation - Invalid Sheet Format'
      );
      props.setProperty(RUNNING_FLAG_KEY, 'false');
      _clearMainTriggers();
      return;
    }

    const data = sheet.getDataRange().getValues();

    let startIndex = parseInt(props.getProperty('START_INDEX'), 10) || 1;
    const firstPendingFromTop = _findNextPendingRowIndexFromData(data, 1);
    if (firstPendingFromTop !== -1 && (startIndex < 1 || startIndex >= data.length || firstPendingFromTop < startIndex)) {
      startIndex = firstPendingFromTop;
      props.setProperty('START_INDEX', String(startIndex));
    }

    if (firstPendingFromTop === -1) {
      props.deleteProperty('START_INDEX');
      props.setProperty(RUNNING_FLAG_KEY, 'false');
      _clearMainTriggers();
      _log('TRIGGER STOP: No pending rows found.');
      return;
    }

    let rowsProcessedInThisRun = 0;
    let reachedEndOfSheet = true;
    const candidates = [];

    for (let i = startIndex; i < data.length; i++) {
      if (candidates.length >= SAFE_BATCH_SIZE) {
        props.setProperty('START_INDEX', i.toString());
        reachedEndOfSheet = false;
        break;
      }

      const itemName = data[i][TITLE_COL_IDX.ITEM_NAME];
      const imageUrl = data[i][TITLE_COL_IDX.IMAGE_URL];
      const existingTitle = data[i][TITLE_COL_IDX.TITLE];

      if (existingTitle && existingTitle.toString().trim() !== '') {
        continue;
      }

      if (!imageUrl || !itemName) {
        continue;
      }

      candidates.push({
        rowIndex: i,
        itemName: String(itemName),
        imageUrl: String(imageUrl)
      });
    }

    if (candidates.length) {
      const imageResponses = UrlFetchApp.fetchAll(candidates.map(c => ({
        url: c.imageUrl,
        muteHttpExceptions: true
      })));

      const generationJobs = [];
      for (let i = 0; i < candidates.length; i++) {
        const candidate = candidates[i];
        try {
          const imageResponse = imageResponses[i];
          if (imageResponse.getResponseCode() !== 200) {
            throw new Error(`Image fetch failed with code ${imageResponse.getResponseCode()}`);
          }

          const imgBlob = imageResponse.getBlob();
          const imagePart = _prepareImagePart(imgBlob, candidate.imageUrl);
          generationJobs.push({
            rowIndex: candidate.rowIndex,
            itemName: candidate.itemName,
            request: _buildGenerateRequest(candidate.itemName, imagePart, languageMode)
          });
        } catch (e) {
          _log(`ROW ERROR: ${candidate.rowIndex + 1} | ${e.message}`, 'ERROR');
        }
      }

      if (generationJobs.length) {
        for (let chunkStart = 0; chunkStart < generationJobs.length; chunkStart += GENERATION_PARALLEL_CHUNK_SIZE) {
          const chunk = generationJobs.slice(chunkStart, chunkStart + GENERATION_PARALLEL_CHUNK_SIZE);
          const generationResponses = UrlFetchApp.fetchAll(chunk.map(j => j.request));

          for (let i = 0; i < chunk.length; i++) {
            const job = chunk[i];
            try {
              const response = _retryGenerateIfNeeded(generationResponses[i], job.request, job.rowIndex + 1);
              let result = _parseGenerateResponse(response);
              result = _sanitizeResultByLanguageMode(result, languageMode);

              const bullets = _normalizeBullets(result.bullets);
              const rowValues = [[
                (result.title || '').substring(0, 200),
                result.description || '',
                bullets[0],
                bullets[1],
                bullets[2],
                bullets[3],
                bullets[4],
                _truncateByBytes(result.keywords || '', 249),
                parseInt(result.weight, 10) || 0
              ]];

              // F-N: Title, Description, Bullet 1..5, Keyword, Weight
              sheet.getRange(job.rowIndex + 1, 6, 1, 9).setValues(rowValues);

              rowsProcessedInThisRun++;
              _log(`ROW SUCCESS: ${job.rowIndex + 1} | ${job.itemName}`);
            } catch (e) {
              _log(`ROW ERROR: ${job.rowIndex + 1} | ${e.message}`, 'ERROR');
            }
          }
        }
      }
    }

    if (reachedEndOfSheet) {
      props.deleteProperty('START_INDEX');
      props.setProperty(RUNNING_FLAG_KEY, 'false');
      _clearMainTriggers();
      _log('TRIGGER STOP: ALL ROWS COMPLETED SUCCESSFULLY.');
    }
  } finally {
    runLock.releaseLock();
  }
}

function _registerMainTrigger() {
  _clearMainTriggers();
  ScriptApp.newTrigger('processBatch').timeBased().everyMinutes(1).create();
  _log('TRIGGER REGISTERED: processBatch scheduled with 1-minute interval.');
}

function _isMainTriggerRunning() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processBatch') {
      return true;
    }
  }
  return false;
}

function _clearMainTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  return removed;
}

function _log(message, level) {
  if (!ENABLE_DETAILED_LOGS && level !== 'ERROR') {
    return;
  }
  const tz = Session.getScriptTimeZone() || 'GMT';
  const ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const lvl = level || 'INFO';
  const formatted = `[${ts}] [${lvl}] ${message}`;
  Logger.log(formatted);

  try {
    const logSheet = _getOrCreateLogSheet();
    if (logSheet) {
      logSheet.appendRow([ts, lvl, message]);
    }
  } catch (e) {
    Logger.log(`[${ts}] [ERROR] LOG_WRITE_FAILED: ${e.message}`);
  }
}

function _getOrCreateLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return null;
  }

  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['timestamp', 'level', 'message']);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function _getRuntimeContext(showDialog) {
  const hostSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!hostSpreadsheet) {
    _notifyUser('Active spreadsheet not found.', showDialog, 'Title Generation - Error');
    return null;
  }

  const configSheet = _getSheetByNameCaseInsensitive(hostSpreadsheet, CONFIG_SHEET_NAME);
  if (!configSheet) {
    _notifyUser(`Sheet '${CONFIG_SHEET_NAME}' not found in active spreadsheet.`, showDialog, 'Title Generation - Error');
    return null;
  }

  const targetSheetId = String(configSheet.getRange('H2').getValue() || '').trim();
  const languageChoice = String(configSheet.getRange('I2').getValue() || '').trim();

  if (!targetSheetId) {
    _notifyUser('Sheet ID not found. Please enter correct Sheet ID in configuration!H2.', showDialog, 'Title Generation - Invalid Sheet ID');
    return null;
  }

  let targetSpreadsheet;
  try {
    targetSpreadsheet = SpreadsheetApp.openById(targetSheetId);
  } catch (e) {
    _notifyUser('Sheet ID not found. Please enter correct Sheet ID.', showDialog, 'Title Generation - Invalid Sheet ID');
    _log(`ERROR: Invalid target sheet ID '${targetSheetId}' | ${e.message}`, 'ERROR');
    return null;
  }

  const targetSheet = _getSheetByNameCaseInsensitive(targetSpreadsheet, TARGET_SHEET_NAME);
  if (!targetSheet) {
    _notifyUser(`Sheet '${TARGET_SHEET_NAME}' not found in target spreadsheet.`, showDialog, 'Title Generation - Error');
    return null;
  }

  return {
    targetSpreadsheet: targetSpreadsheet,
    targetSheet: targetSheet,
    languageMode: _resolveLanguageMode(languageChoice)
  };
}

function _getSheetByNameCaseInsensitive(ss, name) {
  if (!ss || !name) return null;
  const target = String(name).trim().toLowerCase();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getName()).trim().toLowerCase() === target) {
      return sheets[i];
    }
  }
  return null;
}

function _normalizeHeaderToken(value) {
  return String(value === undefined || value === null ? '' : value)
    .trim()
    .replace(/\s+/g, '')
    .toLowerCase();
}

function _validateTitleCreationHeaders(sheet) {
  if (!sheet) {
    return {
      ok: false,
      message: `Sheet '${TARGET_SHEET_NAME}' not found.`,
      actualHeaderList: ''
    };
  }

  const requiredHeaderCount = EXPECTED_TITLE_HEADERS.length;
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < requiredHeaderCount || sheet.getLastRow() < 1) {
    return {
      ok: false,
      message: 'Required header columns are missing.',
      actualHeaderList: ''
    };
  }

  const actualHeaders = sheet.getRange(1, 1, 1, requiredHeaderCount).getValues()[0];
  const expectedNormalized = EXPECTED_TITLE_HEADERS.map(_normalizeHeaderToken);
  const actualNormalized = actualHeaders.map(_normalizeHeaderToken);

  for (let i = 0; i < expectedNormalized.length; i++) {
    if (expectedNormalized[i] !== actualNormalized[i]) {
      return {
        ok: false,
        message: `Header mismatch at column ${i + 1}.`,
        actualHeaderList: actualHeaders.map(h => String(h || '').trim()).join(',')
      };
    }
  }

  return {
    ok: true,
    message: '',
    actualHeaderList: actualHeaders.map(h => String(h || '').trim()).join(',')
  };
}

function _findNextPendingRowIndexFromData(data, fromIndex) {
  if (!data || data.length <= 1) {
    return -1;
  }

  const start = Math.max(1, parseInt(fromIndex, 10) || 1);
  for (let i = start; i < data.length; i++) {
    const row = data[i] || [];
    const itemName = row[TITLE_COL_IDX.ITEM_NAME];
    const imageUrl = row[TITLE_COL_IDX.IMAGE_URL];
    const existingTitle = row[TITLE_COL_IDX.TITLE];

    const hasItem = itemName !== undefined && itemName !== null && String(itemName).trim() !== '';
    const hasImage = imageUrl !== undefined && imageUrl !== null && String(imageUrl).trim() !== '';
    const hasTitle = existingTitle !== undefined && existingTitle !== null && String(existingTitle).trim() !== '';

    if (hasItem && hasImage && !hasTitle) {
      return i;
    }
  }

  return -1;
}

function _findNextPendingRowIndex(sheet) {
  if (!sheet) {
    return -1;
  }
  const data = sheet.getDataRange().getValues();
  return _findNextPendingRowIndexFromData(data, 1);
}

function _resolveLanguageMode(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (normalized.indexOf('english') !== -1) {
    return 'english';
  }
  if (normalized.indexOf('country') !== -1) {
    return 'country';
  }
  return 'country';
}

function _notifyUser(message, showDialog, title) {
  const msg = String(message || 'Unknown error.');
  _log(msg, 'ERROR');
  if (!showDialog) {
    return;
  }

  try {
    SpreadsheetApp.getUi().alert(title || 'Title Generation', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    _log(`UI ALERT FAILED: ${e.message}`, 'ERROR');
  }
}

function _prepareImagePart(blob, sourceUrl) {
  const bytes = blob.getBytes();
  const mimeType = blob.getContentType() || 'image/jpeg';

  // Faster path on Vertex: send inline image directly in generateContent call.
  if (bytes.length <= INLINE_IMAGE_MAX_BYTES) {
    return {
      inlineData: {
        mimeType: mimeType,
        data: Utilities.base64Encode(bytes)
      }
    };
  }

  // Fallback for large images: use public image URL reference.
  if (!sourceUrl) {
    throw new Error(`Image is too large for inline payload (${bytes.length} bytes) and no URL fallback is available.`);
  }

  return {
    fileData: {
      fileUri: String(sourceUrl),
      mimeType: mimeType
    }
  };
}

function _buildLanguageInstruction(languageMode) {
  if (languageMode === 'english') {
    return 'Generate title, description, bullets and keywords in English only, regardless of item-name language. Output must be strictly English.';
  }
  return 'Generate title, description, bullets and keywords in the same language as item-name text. For all outputs, avoid accented/diacritic characters (for example: ë, ï, ü, ÿ, à, è, ù, â, ê, î, ô, û) and use plain letters.';
}

function _buildGenerateRequest(itemName, imagePart, languageMode) {
  const url = `${GENERATE_ENDPOINT_BASE}/${MODEL_NAME}:generateContent?key=${encodeURIComponent(API_KEY)}`;
  const languageInstruction = _buildLanguageInstruction(languageMode);

  const payload = {
    contents: [{
      role: 'user',
      parts: [
        {
          text: `System: Amazon Jewelry SEO Expert. Goal: Max discovery.
                       Task: Analyze "${itemName}" and the image.
                       Return JSON ONLY:
                       {
                         "title": "SEO Title (180-200 chars). High density keywords. Format: Brand + Material + Gemstone + Type + Style + Occasion.",
                         "description": "Engaging product description (approx 500-1000 chars).",
                         "bullets": ["Material Quality & Finish", "Gemstone Cut/Clarity", "Detailed Design/Size", "Versatile Gifting/Occasions", "Quality Care/Packaging"],
                         "keywords": "Maximum possible amount of generic search terms, synonyms, and trends. No commas. No brand names. Maximize 249 bytes.",
                         "weight": "Estimate weight. Return ONLY the minimum integer in grams."
                       }
                       Additional language rule: ${languageInstruction}`
        },
        imagePart
      ]
    }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.4 }
  };

  return {
    url: url,
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
}

function _parseGenerateResponse(response) {
  if (response.getResponseCode() !== 200) {
    throw new Error(`API Error: ${response.getContentText()}`);
  }

  const resJson = JSON.parse(response.getContentText());
  const raw = (((resJson.candidates || [])[0] || {}).content || {}).parts || [];
  const text = raw.length && raw[0].text ? raw[0].text : '';
  if (!text) {
    throw new Error(`Empty model response: ${response.getContentText()}`);
  }

  return _parseJsonSafely(text);
}

function _isRetryableGenerateStatus(statusCode) {
  return statusCode === 429 || statusCode === 500 || statusCode === 503 || statusCode === 504;
}

function _executeGenerateRequest(request) {
  return UrlFetchApp.fetch(request.url, {
    method: request.method,
    contentType: request.contentType,
    payload: request.payload,
    muteHttpExceptions: request.muteHttpExceptions
  });
}

function _retryGenerateIfNeeded(initialResponse, request, rowNumber) {
  let response = initialResponse;
  let attempt = 0;

  while (response && _isRetryableGenerateStatus(response.getResponseCode()) && attempt < GENERATION_MAX_RETRIES) {
    attempt++;
    const delayMs = (GENERATION_RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1)) + Math.floor(Math.random() * 500);
    _log(`ROW RETRY: ${rowNumber} | HTTP ${response.getResponseCode()} | retry ${attempt}/${GENERATION_MAX_RETRIES} | waiting ${delayMs}ms`);
    Utilities.sleep(delayMs);
    response = _executeGenerateRequest(request);
  }

  return response;
}

function _generateContent(itemName, imagePart, languageMode) {
  const request = _buildGenerateRequest(itemName, imagePart, languageMode);
  const response = _retryGenerateIfNeeded(_executeGenerateRequest(request), request, 0);

  return _parseGenerateResponse(response);
}

function _stripDiacritics(value) {
  const text = String(value === undefined || value === null ? '' : value);
  try {
    return text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (e) {
    return text;
  }
}

function _sanitizeResultByLanguageMode(result, languageMode) {
  if (languageMode !== 'country') {
    return result || {};
  }

  const output = result || {};
  output.title = _stripDiacritics(output.title || '');
  output.description = _stripDiacritics(output.description || '');
  output.keywords = _stripDiacritics(output.keywords || '');

  const bullets = Array.isArray(output.bullets) ? output.bullets : [];
  output.bullets = bullets.map(_stripDiacritics);
  return output;
}

function _parseJsonSafely(text) {
  const trimmed = (text || '').trim();
  try {
    return JSON.parse(trimmed);
  } catch (e) {
    const match = trimmed.match(/\{[\s\S]*\}/);
    if (match && match[0]) {
      return JSON.parse(match[0]);
    }
    throw new Error(`Invalid JSON returned by model: ${trimmed}`);
  }
}

function _normalizeBullets(bullets) {
  const arr = Array.isArray(bullets) ? bullets : [];
  const normalized = [];
  for (let i = 0; i < 5; i++) {
    normalized.push(arr[i] ? String(arr[i]) : '');
  }
  return normalized;
}

function _truncateByBytes(str, maxBytes) {
  if (!str) return '';
  let result = String(str);
  while (Utilities.newBlob(result, 'UTF-8').getBytes().length > maxBytes) {
    result = result.slice(0, -1);
  }
  return result;
}