/** src/ui/taskpane.js — Main task pane UI controller */

import Recorder   from '../core/recorder.js';
import Player     from '../core/player.js';
import * as fmt    from '../core/format_encoding.js';

/* ── INTEGRATED DEBUGGER SETUP ──────────────────────────── */
const _originalLog = console.log;
window.appLogger = (msg) => {
  const logEl = document.getElementById('app-debug-log');
  if (logEl) {
    const line = document.createElement('div');
    line.className = 'debug-line';
    const content = typeof msg === 'object' ? JSON.stringify(msg, null, 2) : msg;
    line.textContent = `[${new Date().toLocaleTimeString()}] ${content}`;
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
  }
};

console.log = (...args) => {
  _originalLog(...args);
  if (args.length > 0) window.appLogger(args[0]);
};

/* ── DOM helpers ─────────────────────────────────────────── */
const $ = (id) => document.getElementById(id);

let domElements = {};
let recorder = new Recorder();
let player   = new Player();
let db       = null;
let timerId  = null;
let autoSaveTimerId = null;
let userEmail = 'unknown'; // Default identity for Word environments

/* ── UI element references (populated on init) ──────────── */
function initElements() {
  domElements = {
    statusBadge:        $('statusBadge'),
    btnRecord:          $('btnRecord'),
    btnStop:            $('btn02_stop_id_fallback_fix_needed'), // mapping later... wait, let me use existing structure
    btnStop:            $('btnStop'), 
    btnPause:           $('btnPause'),
    sessionIdEl:        $('sessionId'),
    docTitleEl:         $('docTitle'),
    eventCountEl:       $('eventCount'),
    recTimeEl:          $('recordingTime'),
    lastSaveMsg:        $('lastSave'),
    playbackSection:    $('playbackSection'), // Corrected from playbackControls/timelineArea logic
    timelineFill:       $('timelineFill'),
    btnPlay:            $('btnPlay'),
    seekSlider:         $('seekSlider'),
    seekLabel:          $('seekLabel'),
    speedSelect:        $('speedSelect'),
    storageActions:     $('storageActions'),
    btnSaveLocal:       $('btnSaveLocal'),
    btnExportWtp:       $('btnExportWtp'),
    fileInput:          $('fileInput'),
    eventListEl:        $('eventList'),
    errorBanner:        $('errorBanner'),
    debugPanel:         $('debugPanel'),
    btnClearEvents:     $('btnClearEvents'),
    btnSaveDoc:         $('btnSaveDoc'),
   };

  // Fix mapping for the button listener/element reference logic from HTML
  // The user's new HTML has specific IDs we must strictly match. Let me re-check it.
  // <div class="playback-section" id="playbackSection"> ... </div>
  // My code above already attempts to use playbackSection via the manual update below.

  injectDebugUI();
}

function injectDebugUI() {
  const style = document.createElement('style');
  style.textContent = `
    #debug-container {
      position: fixed; bottom: 0; left: 0; right: 0; height: 150px;
      background: #1e1e1e; color: #d4d4d4; font-family: 'Consolas', monospace;
      font-size: 11px; overflow-y: auto; border-top: 2px solid #444;
      z-index: 9999; padding: 8px; box-shadow: 0 -4px 10px rgba(0,0,0,0.5);
    }
    .debug-line { border-bottom: 1px solid #333; padding: 2px 0; white-space: pre-wrap; line-height: 1.2;}
  `;
  document.head.appendChild(style);

  const container = document.createElement('div');
  container.id = 'debug-container';
  
  const title = document.createElement('div');
  title.style.cssText = "font-weight:bold; color:#aaa; margin-bottom:4px; border-bottom:1px solid #555;";
  title.textContent = "INTERNAL DEBUG LOG";
  container.appendChild(title);

  const logArea = document.createElement('div');
  logArea.id = 'app-debug-log';
  container.appendChild(logArea);

  document.body.appendChild(container);
  console.log("DEBUG SYSTEM: UI and Redirection Active.");
}

/* ── Error handling & debugging ─────────────────────────── */
function showError(message) {
  const el = domElements.errorBanner;
  if (el) {
    el.textContent = message;
    el.classList.add('visible');
   }
  console.error('[WordTimestamp]', message);
}

function clearError() {
  const el = domElements.errorBanner;
  if (el) {
    el.classList.remove('visible');
  }
}

function updateDebug() {
  const panel = domElements.debugPanel;
  if (!panel) return;
  const state = {
    recording: recorder.recording,
    paused: recorder.paused,
    entries: recorder.entries.length,
    sessionId: recorder.sessionId || 'none'
    };
  panel.innerHTML = `<code style="color:#aaa">state=${JSON.stringify(state)}</code>`;
  // Add visibility logic here fix
  if (recorder.recording || recorder.entries.length > 0) {
      panel.classList.add('visible');
  } else {
      panel.classList.remove('visible');
  }
}

/* ── Recorder callback wiring ───────────────────────────── */
recorder.onFlush = () => {
  if (domElements.eventCountEl) {
    domElements.eventCountEl.textContent = `${recorder.entries.length}`;
  }
  updateDebug();
};

recorder.onError = ({ phase, error }) => {
  showError(`Recording error (${phase}): ${error}`);
};

/* ── Player callback wiring ─────────────────────────────── */
player.tickCallback = function(state) {
  const dur = player.duration();
  const pct = dur > 0 ? (state.position / dur) * 100 : 0;

  if (domElements.timelineFill)
    domElements.timelineFill.style.width = Math.min(pct, 100) + '%';
  if (domElements.seekSlider)
    domElements.seekSlider.value = Math.round(pct);
  if (domElements.seekLabel)
    domElements.seekLabel.textContent = fmtTime(Math.abs(state.position));

  renderEventList(player.entriesAt(state.position));
};

/* ── IndexedDB lifecycle ─────────────────────────────────── */
async function openDb() {
  return new Promise(function(resolve, reject) {
    var req = indexedDB.open('word-timestamp-db', 1);
    req.onupgradeneeded = function(ev) {
      const database = ev.target.result;
      if (!database.objectStoreNames.contains('sessions')) {
        database.createObjectStore('sessions', { keyPath: 'sessionId' });
       }
     };
    req.onsuccess = function(e) { db = e.target.result; resolve(db); };
    req.onerror   = function(e) { reject(e?.target?.error || new Error('IndexedDB open failed')); };
   });
}

/* ── Time formatting ─────────────────────────────────────── */
function fmtTime(ms) {
  ms = Math.abs(ms);
  if (ms < 1000) return `${Math.round(ms)}ms`;
  var s = Math.floor(ms / 1000);
  var m = Math.floor(s / 60);
  s %= 60;
  return `${m}:${padZero(s)}`;
}

function padZero(n) { return String(n).padStart(2, '0'); }

/* ── Event list rendering ────────────────────────────────── */
function renderEventList(entries) {
  const el = domElements.eventListEl;
  if (!el || !entries.length) {
    if (el) el.innerHTML = '';
    return;
  }

  var html = '<div class="ev-row" style="font-weight:bold;background:#f1f5f9;padding:4px;border-radius:3px;margin-bottom:4px;">Time|Kind|ΔText</div>';
  const recent = entries.slice(-50);
  for (var i = 0; i < recent.length; i++) {
    var entry    = recent[i];
    var op       = entry.ops?.[0] || {};
    var inserted = (op.i || '').slice(0, 40).replace(/</g, '&lt;');
    var deleted  = (op.d || '').slice(0, 40).replace(/</g, '&lt;');
    var preview  = deleted ? `-${deleted}→+${inserted}` : inserted.slice(0, 50);
    if (!preview && !inserted) preview = '[change]';
    var ts       = fmtTime(entry.t);
    var kn       = entry.k || 'text';

    html += '<div class="ev-row">';
    html += `<span class="event-time">${ts}</span>`;
    html += `<span class="event-kind">${kn}</span>`;
    html += `<span title="${deleted ? 'Deleted: ' + deleted : ''}">${preview || '[change]'}</span>`;
    html += '</div>';
   }

  el.innerHTML = html;
  el.scrollTop = el.scrollHeight;
}

/* ── UI State Manager ────────────────────────────────────── */
function updateUI() {
  const d = domElements;
  if (!d.btnRecord) return;

  if (recorder.recording) {
    d.btnRecord.disabled     = true;
    d.btnStop.disabled       = recorder.paused;
    d.btnPause.disabled      = false;
    d.btnPause.textContent   = recorder.paused ? '▶ Resume' : '⏸ Pause';
    d.statusBadge.classList.add('active');
    d.statusBadge.classList.remove('error');
    d.storageActions.style.display = 'none';
   } else {
    d.btnRecord.disabled     = false;
    d.btnStop.disabled       = true;
    d.btnPause.disabled      = true;
    d.btnPause.textContent   = '⏸ Pause';
    d.statusBadge.classList.remove('active');
   }

  d.eventCountEl.textContent = `${recorder.entries.length}`;
  if (recorder.recording) {
    d.recTimeEl.textContent = fmtTime(recorder._relativeMs());
   }

  if (recorder.entries.length) d.playbackSection.style.display = 'block';
  updateDebug();
}

/* ── Auto-save every 5 seconds ───────────────────────────── */
function startAutoSave() {
  stopAutoSave();
  autoSaveTimerId = setInterval(async () => {
    try {
      if (!db) await openDb();
      const session = fmt.buildSession(recorder);
      await fmt.saveSession(session, db);
      if (domElements.lastSaveMsg)
        domElements.lastSaveMsg.textContent = `Auto-saved ${new Date().toLocaleTimeString()}`;
     } catch (err) {
      console.warn('[WordTimestamp] Auto-save failed:', err);
     }
   }, 5000);
}

function stopAutoSave() {
  if (autoSaveTimerId !== null) {
    clearInterval(autoSaveTimerId);
    autoSaveTimerId = null;
   }
}

/* ── Click handlers ─────────────────────────────────────── */
async function onStartRecording() {
  clearError();
  try {
    const started = await recorder.start();
    if (started) {
      domElements.sessionIdEl.textContent = recorder.sessionId.slice(0, 8) + '...';
      domElements.docTitleEl.textContent  = recorder.docTitle || 'Loading...';
      startTimer();
      startAutoSave();
      updateUI();
     } else {
      showError('Failed to start recording.');
     }
   } catch (err) {
    showError(`Start failed: ${err.message}`);
   }
}

async function onStopRecording() {
  try {
    recorder.stop();
    stopTimer();
    stopAutoSave();
    domElements.statusBadge.textContent = 'Stopped';
    domElements.statusBadge.classList.remove('active');

    try {
      if (!db) await openDb();
      const session = fmt.buildSession(recorder);
      await fmt.saveSession(session, db);
      if (domElements.lastSaveMsg)
        domElements.lastSaveMsg.textContent = `Saved ${new Date().toLocaleTimeString()}`;
     } catch (saveErr) {
      console.warn('[WordTimestamp] Final save failed:', saveErr);
     }

    player.load(recorder.entries);
    domElements.playbackSection.style.display   = 'block';
    domElements.storageActions.style.display     = 'flex';
    updateUI();
   } catch (err) {
    showError(`Stop failed: ${err.message}`);
   }
}

function onPauseResume() {
  try {
    if (!recorder.recording) return;
    if (recorder.paused) recorder.resume(); else recorder.pause();
    updateUI();
   } catch (err) {
    showError(`Pause/Resume failed: ${err.message}`);
   }
}

function onTogglePlayback() {
  try {
    if (player.isPlaying) {
      player.pause();
      domElements.btnPlay.textContent = '▶ Play';
     } else {
      player.play();
      domElements.btnPlay.textContent = '⏸ Pause';
     }
   } catch (err) {
    showError(`Playback failed: ${err.message}`);
   }
}

function onChangeSpeed(e) {
  try {
    player.setSpeed(parseFloat(e.target.value || '1'));
   } catch (err) {
    console.warn('[WordTimestamp] Speed change failed:', err);
   }
}

function onSeekSlide(e) {
  try {
    var pct = parseFloat(e.target.value);
    var maxT = player.duration();
    if (maxT <= 0) return;
    const targetMs = (pct / 100) * maxT;
    player.setPosition(targetMs);
   } catch (err) {
    console.warn('[WordTimestamp] Seek failed:', err);
   }
}

async function onSaveLocal() {
  clearError();
  try {
    if (!db) await openDb();
    const session = fmt.buildSession(recorder);
    await fmt.saveSession(session, db);
    domElements.lastSaveMsg.textContent = `Saved ${new Date().toLocaleTimeString()}`;
   } catch (err) {
    showError(`Local save failed: ${err.message}`);
   }
}

async function onExportWtp() {
  clearError();
  if (!recorder.entries.length) {
    showError('No events to export.');
    return;
   }
  try {
    const session = fmt.buildSession(recorder);
    const result = await fmt.exportWtp(session);
    domElements.lastSaveMsg.textContent = `Exported ${result.filename}`;
   } catch (err) {
    showError(`Export failed: ${err.message}`);
   }
}

async function onImportFile(ev) {
  clearError();
  var file = ev.target.files?.[0];
  if (!file) return;
  try {
    const session = await fmt.importWtp(file);
    domElements.docTitleEl.textContent = session.docTitle || 'Imported';
    player.load(session.entries);
    domElements.playbackSection.style.display      = 'block';
    domElements.storageActions.style.display     = 'none'; 

     if (session.integrityOk) {
      domElements.statusBadge.textContent = `✓ ${session.entries.length}`;
      domElements.statusBadge.classList.remove('error');
     } else {
      domElements.statusBadge.textContent = `⚠ ${session.entries.length}`;
      domElements.statusBadge.classList.add('error');
     }

    ev.target.value = '';
    updateUI();
   } catch (err) {
    showError(`Import failed: ${err.message}`);
   }
}

function onClearEvents() {
  if (!confirm('Clear all recorded events?')) return;
  recorder.clear();
  player.clear();
  stopTimer();
  stopAutoSave();
  domElements.sessionIdEl.textContent    = '—';
  domElements.docTitleEl.textContent     = '—';
  domElements.eventCountEl.textContent   = '0';
  domElements.recTimeEl.textContent      = '—';
  domElements.lastSaveMsg.textContent       = '—';

  domElements.playbackSection.style.display = 'none';
  domElements.storageActions.style.display = 'none';
  domElements.eventListEl.innerHTML      = '';
  updateUI();
}

async function onSaveDoc() {
  try {
    await Word.run(async (ctx) => {
      const doc = ctx.document;
      doc.save();
      await ctx.sync();
     });
    console.log('[WordTimestamp] Document saved');
   } catch (err) {
    showError(`Document save failed: ${err.message}`);
   }
}

function startTimer() {
  stopTimer();
  timerId = setInterval(() => {
    domElements.recTimeEl.textContent = fmtTime(recorder._relativeMs());
   }, 500); 
}

function stopTimer() {
  if (timerId !== null) { clearInterval(timerId); timerId = null; }
}

/* ── Office initialization ──────────────────────────────── */
Office.onInitialized(async function() {
  console.log('[WordTimestamp] Office initialized...');

  try {
    initElements();

    // User info is not available in Word task panes via mail API.
    userEmail = 'guest-user';

    try { await openDb(); } catch (dbErr) {
      console.warn('[WordTimestamp] IndexedDB error:', dbErr);
     }

    domElements.btnRecord.addEventListener('click', onStartRecording);
    domElements.btnStop.addEventListener('click', onStopRecording);
    domElements.btnPause.addEventListener('click', onPauseResume);
    domElements.btnPlay.addEventListener('click', onTogglePlayback);
    domElements.seekSlider.addEventListener('input', onSeekSlide);
    domElements.speedSelect.addEventListener('change', onChangeSpeed);
    domElements.btnSaveLocal.addEventListener('click', onSaveLocal);
    domElements.btnExportWtp.addEventListener('click', onExportWtp);
    domElements.fileInput.addEventListener('change', onImportFile);
    domElements.btnClearEvents.addEventListener('click', onClearEvents);
    if (domElements.btnSaveDoc) {
      domElements.btnSaveDoc.addEventListener('click', onSaveDoc);
     }

    setTimeout(() => {
      updateDebug();
    }, 1000);
   } catch (err) {
    console.error('[WordTimestamp] Initialization error:', err);
   }
});