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

/* ── DOM HELPERS ─────────────────────────────────────────── */
const $ = (id) => document.getElementById(id);

let domElements = {};
let recorder = new Recorder();
let player   = new Player();
let db       = null;
let timerId  = null;
let autoSaveTimerId = null;
let userEmail = 'unknown'; // Default identity for Word task panes

function initElements() {
  domElements = {
    statusBadge:        $('statusBadge'),
    btnRecord:          $('btnRecord'),
    btnStop:            $('btnStop'),
    btnPause:           $('btnPause'),
    sessionIdEl:        $('sessionId'),
    docTitleEl:         $('docTitle'),
    eventCountEl:       $('eventCount'),
    recTimeEl:          $('recordingTime'),
    lastSaveMsg:        $('lastSaveMsg'), // Match HTML id="lastSave" in HTML setup if needed, but we use lastSaveMsg here. Let's ensure symmetry.
    playbackSection:    $('playbackSection'),
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

  injectDebugUI();
}

function injectDebugUI() {
  const style = document.createElement('style');
  style.textContent = `
    #debug-container {
      position: fixed; bottom: 0; left: 0; right: 0; height: 150px;
      background: #1e2e3b; color: #d4d4d4; font-family: 'Consolas', monospace;
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

/* ── ERROR HANDLING ───────────────────────────────────────── */
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
  if (el) el.classList.remove('visible');
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
  
  if (recorder.recording || recorder.entries.length > 0) {
      panel.classList.add('visible');
  } else {
      panel.classList.remove('visible');
  }
}

/* ── RECORDER CALLBACK WIRING ───────────────────────────── */
recorder.onFlush = () => {
  if (domElements.eventCountEl) {
    domElements.eventCountEl.textContent = `${recorder.entries.length}`;
  }
  updateDebug();
};

recorder.onError = ({ phase, error }) => {
  showError(`Recording error (${phase}): ${error}`);
};

/* ── PLAYER CALLBACK WIRING ─────────────────────────────── */
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

/* ── INDEXED DB LIFECYCLE ─────────────────────────────────── */
async function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('word-timestamp-db', 1);
    req.onupgradeneeded = (ev) => {
      const db = ev.target.result;
      if (!db.objectStoreNames.contains('sessions')) {
        db.createObjectStore('sessions', { keyPath: 'sessionId' });
       }
     };
    req.onsuccess = (e) => { db = e.target.result; resolve(db); };
    req.onerror = (e) => reject(e?.target?.error || new Error('IndexedDB open failed'));
   });
}

/* ── TIME FORMATTING ─────────────────────────────────────── */
function fmtTime(ms) {
  ms = Math.abs(ms);
  if (ms < 1000) return `${Math.round(ms)}ms`;
  const s = Math.floor(ms / 1000);
  const m = Math.floor(s / 60);
  return `${m}:${String(s % 60).padStart(2, '0')}`;
}

/* ── EVENT LIST RENDERING ────────────────────────────────── */
function renderEventList(entries) {
  const el = domElements.eventListEl;
  if (!el) return;
  if (entries.length === 0) {
    el.innerHTML = '';
    return;
  }

  let html = '<div class="ev-row" style="font-weight:bold;background:#f1f5f9;padding:4px;border-radius:3px;margin-bottom:4px;">Time|Kind|ΔText</div>';
  const recent = entries.slice(-50);
  for (let i = 0; i < recent.length; i++) {
    const entry = recent[i];
    const op = entry.ops?.[0] || {};
    const inserted = (op.i || '').slice(0, 40).replace(/</g, '&lt;');
    const deleted  = (op.d || '').slice(0, 40).replace(/</g, '&lt;');
    const preview  = deleted ? `-${deleted}→+${inserted}` : inserted.slice(0, 50);
    const ts       = fmtTime(entry.t);
    const kn       = entry.k || 'text';

    html += `<div class="ev-row">
      <span class="event-time">${ts}</span>
      <span class="event-kind">${kn}</span>
      <span title="${deleted ? 'Deleted: ' + deleted : ''}">${preview || '[change]'}</span>
    </div>`;
   }

  el.innerHTML = html;
  el.scrollTop = el.scrollHeight;
}

/* ── UI STATE MANAGER ────────────────────────────────────── */
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

  if (recorder.entries.length > 0) d.playbackSection.style.display = 'block';
  updateDebug();
}

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

/* ── CLICK HANDLERS ───────────────────────────────────────── */
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
    const pct = parseFloat(e.target.value);
    const maxT = player.duration();
    if (maxT <= 0) return;
    player.setPosition((pct / 100) * maxT);
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

    // Identity is not easily available in Word task panes via standard API.
    userEmail = 'guest-user';

    try { await openDb(); } catch (dbErr) {
      console.warn('[WordTimestamp] IndexedDB error:', dbErr);
     }

    console.log(`[WordTimestamp] User: ${userEmail}`);

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