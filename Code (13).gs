/*********************************************************************
 * ANTAKSHARI — ONLINE (Google Apps Script)
 * ---------------------------------------------------------------
 * Server-side code. This is the online replacement for the offline
 * "choose a CSV + choose a media folder on your device" flow:
 *   - The song/word list now lives in a tab of THIS spreadsheet.
 *   - The media files (mp3/mp4/images) and lyrics .txt files now
 *     live in a Google Drive folder, referenced by filename exactly
 *     like the old MP3Link / LyricsLink columns.
 *
 * SETUP (see README.md for full step-by-step):
 *   1. Set SHEET_NAME below to the tab that holds your song/word list
 *      (same columns as the old CSV: Category, Subcategory,
 *      QuestionNumber, MP3Link, Title, Seconds, LyricsLink, Answer, Words).
 *   2. Upload your media + lyrics files to one Drive folder, copy its
 *      folder ID from the URL, and paste it into MEDIA_FOLDER_ID below.
 *   3. Deploy > New deployment > Web app.
 *        Execute as:  Me
 *        Who has access:  Anyone with the link (or your organization)
 *********************************************************************/

// ---- CONFIGURE THESE TWO VALUES ----
const SHEET_NAME = 'Songs';                          // tab name holding the song/word list
const MEDIA_FOLDER_ID = 'PUT_YOUR_DRIVE_FOLDER_ID_HERE'; // Drive folder with mp3/mp4/image/lyrics files

const DEFAULT_SECONDS = 30;
const CACHE_KEY = 'antakshari_data_v1';
const CACHE_SECONDS = 300; // 5 minutes — bump this down while you're still editing the sheet

// =====================================================================
// WEB APP ENTRY POINT
// =====================================================================
function doGet(e) {
  // JSON API mode (used when Index.html is hosted outside Apps Script):
  //   <exec-url>?fn=getSongsAndWords
  //   <exec-url>?fn=getSong&args=["Category","Sub","1"]
  if (e && e.parameter && e.parameter.fn) {
    const allowed = { getSongsAndWords: getSongsAndWords, getSong: getSong };
    let out;
    try {
      const fn = allowed[e.parameter.fn];
      if (!fn) throw new Error('Unknown function: ' + e.parameter.fn);
      const args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
      out = { result: fn.apply(null, args) };
    } catch (err) {
      out = { error: String(err && err.message ? err.message : err) };
    }
    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Normal mode: serve the page from Apps Script.
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Antakshari')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================================
// SHEET LOADING (mirrors the old client-side CSV parser)
// =====================================================================
const HEADER_ALIASES = {
  category: 'Category',
  subcategory: 'Subcategory',
  subcat: 'Subcategory',
  questionnumber: 'QuestionNumber',
  question: 'QuestionNumber',
  qno: 'QuestionNumber',
  mp3link: 'MP3Link',
  filename: 'MP3Link',
  file: 'MP3Link',
  media: 'MP3Link',
  title: 'Title',
  seconds: 'Seconds',
  time: 'Seconds',
  lyricslink: 'LyricsLink',
  lyrics: 'LyricsLink',
  lyricsfile: 'LyricsLink',
  answer: 'Answer',
  words: 'Words',
  word: 'Words'
};

function normalizeHeader_(h) {
  return String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Reads the sheet into { songs: [...], words: [...] }, same shape the
// offline version got back from parsing the CSV. Cached briefly so
// switching categories/questions doesn't re-read the sheet every time.
function loadData_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet tab "' + SHEET_NAME + '" not found. Check SHEET_NAME in Code.gs.');
  }
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { songs: [], words: [] };

  const rawHeaders = values.shift();
  const headers = rawHeaders.map(h => HEADER_ALIASES[normalizeHeader_(h)] || null);

  const songs = [];
  const words = [];
  values.forEach(r => {
    if (r.every(cell => String(cell).trim() === '')) return;
    const obj = { Category: '', Subcategory: '', QuestionNumber: '', MP3Link: '', Title: '', Seconds: '', LyricsLink: '', Answer: '', Words: '' };
    headers.forEach((h, i) => {
      if (h && r[i] !== undefined && r[i] !== null) obj[h] = String(r[i]);
    });
    const w = obj.Words.trim();
    if (w) words.push(w);
    if (obj.Category.trim() === '' || obj.QuestionNumber.trim() === '') return; // words-only row
    songs.push(obj);
  });

  const data = { songs, words };
  try {
    cache.put(CACHE_KEY, JSON.stringify(data), CACHE_SECONDS);
  } catch (e) {
    // Sheet is too large to cache (>100KB) — fine, we'll just re-read each time.
  }
  return data;
}

// Call this from the Apps Script editor (Run > clearCache) after editing
// the sheet, if you don't want to wait out the cache TTL.
function clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  return true;
}

// =====================================================================
// PUBLIC API — called from Index.html via google.script.run
// =====================================================================

// Bulk fetch: the client loads the whole song+word list once at startup
// (same as the old CSV parse) and does all its category/subcategory/
// question filtering locally from that.
function getSongsAndWords() {
  return loadData_();
}

// Resolves a single song's playable media URL, timer length, lyrics text
// and answer. Called lazily, only when a song is actually played, so we
// only touch Drive for files that get used.
function getSong(category, subcategory, questionNumber) {
  const { songs } = loadData_();
  const row = songs.find(s =>
    s.Category.trim() === String(category).trim() &&
    (!subcategory || s.Subcategory.trim() === String(subcategory).trim()) &&
    s.QuestionNumber.trim() === String(questionNumber).trim()
  );
  const label = category + (subcategory ? ' / ' + subcategory : '') + ' Q' + questionNumber;
  if (!row) throw new Error('No song found for ' + label);

  const media = resolveMediaUrl_(row.MP3Link, label);

  let secs = parseInt(row.Seconds, 10);
  if (isNaN(secs) || secs <= 0) secs = DEFAULT_SECONDS;
  const title = row.Title ? row.Title : label;

  let lyricsText = '';
  if (row.LyricsLink && row.LyricsLink.trim()) {
    lyricsText = resolveLyricsText_(row.LyricsLink);
  }

  return {
    mediaType: media.mediaType,
    streamUrl: media.url,
    title: title,
    seconds: secs,
    lyricsText: lyricsText,
    answer: row.Answer || ''
  };
}

// =====================================================================
// DRIVE HELPERS
// =====================================================================

// Finds a file by exact name anywhere in the media folder (not subfolders).
function findDriveFile_(name) {
  const folder = DriveApp.getFolderById(MEDIA_FOLDER_ID);
  const base = String(name).trim().split(/[\\/]/).pop(); // strip any path, keep the filename
  const it = folder.getFilesByName(base);
  if (it.hasNext()) return it.next();
  return null;
}

function guessMediaType_(name) {
  if (/\.(mp4|mov|webm|m4v|avi|mkv)(\?|$)/i.test(name)) return 'video';
  if (/\.(jpe?g|png|gif|webp|bmp|svg)(\?|$)/i.test(name)) return 'image';
  return 'audio';
}

// Returns a playable URL for the given MP3Link value: a direct URL is
// passed straight through; a bare filename is looked up in the Drive
// media folder, shared "anyone with the link can view", and turned into
// a direct-download link the <audio>/<video>/<img> tag can stream.
function resolveMediaUrl_(ref, label) {
  const str = String(ref || '').trim();
  if (/^https?:\/\//i.test(str)) {
    return { url: str, mediaType: guessMediaType_(str) };
  }
  const file = findDriveFile_(str);
  if (!file) {
    throw new Error('Could not find media file "' + ref + '" in the Drive media folder for ' + label);
  }
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Sharing settings may be locked down by a Workspace admin policy — the
    // file will still work for anyone who already has folder access.
  }
  const url = 'https://drive.google.com/uc?export=download&id=' + file.getId();
  return { url, mediaType: guessMediaType_(file.getName()) };
}

// Returns the lyrics text for a LyricsLink value: a direct URL is fetched
// as plain text; a bare filename is read straight out of the Drive file.
function resolveLyricsText_(ref) {
  const str = String(ref || '').trim();
  if (!str) return '';
  if (/^https?:\/\//i.test(str)) {
    try {
      const res = UrlFetchApp.fetch(str, { muteHttpExceptions: true });
      return res.getResponseCode() === 200 ? res.getContentText() : '';
    } catch (e) {
      return '';
    }
  }
  const file = findDriveFile_(str);
  if (!file) return '';
  try {
    return file.getBlob().getDataAsString();
  } catch (e) {
    return '';
  }
}
