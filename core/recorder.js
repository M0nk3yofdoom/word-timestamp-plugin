/** src/core/recorder.js — Captures Word document change events with timestamps */

class Recorder {
  constructor() {
    this.recording = false;
    this.paused = false;
    this.startTime = null;
    this.sessionId = '';
    this.docTitle = null;
    this.userName = null;
    this.entries = [];
    this._pauseElapsed = 0;
    this._pausedAt = null;
    this._handlerRef = null; // Store handler reference for cleanup
    this.onFlush = null;    // optional: async callback after each batch
    this.onError = null;    // optional: error callback
  }

  async start() {
    if (this.recording) return false;

    this.sessionId = typeof crypto !== 'undefined' && crypto.randomUUID
       ? crypto.randomUUID()
       : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    this.entries = [];
    this.recording = true;
    this.paused = false;
    this.startTime = Date.now();
    this._pauseElapsed = 0;
    this._pausedAt = null;

    try {
      await Word.run(async (ctx) => {
        const doc = ctx.document;
        doc.load('name');
        await ctx.sync();
        this.docTitle = doc.name || 'Untitled';

        const self = this;
        this._handlerRef = (event) => {
          if (!self.recording || self.paused) {
            event.completed?.();
            return;
          }

          try {
            // --- ROBUST CHANGE DETECTION ENGINE ---
            let changes = [];
            if (event && event.changes) { // Modern Word API
              changes = event.changes;
            } else if (event && event.all) { // Older/Other versions
              changes = event.all;
            } else if (Array.isArray(event)) { // Direct array fallback
              changes = event;
            }

            if (changes.length === 0) {
              event.completed?.();
              return;
            }

            for (let i = 0; i < changes.length; i++) {
              const change = changes[i];
              const entry = Recorder._parseChange(change, self._relativeMs());
              if (entry) {
                self.entries.push(entry);
              }
            }

            if (self.onFlush && typeof self.onFlush === 'function') {
              try { self.onFlush(); } catch (_) {}
            }
          } catch (err) {
            console.error('[WordTimestamp] Error processing change event:', err);
            if (self.onError) {
              try { self.onError({ phase: 'recording', error: err.message }); } catch (_) {}
            }
          }

          event.completed?.();
        };

        doc.onChanged.add(this._handlerRef);
        await ctx.sync();
      });

      console.log(`[WordTimestamp] Recording started. Session: ${this.sessionId}`);
      return true;
    } catch (err) {
      console.error('[WordTimestamp] Failed to start recording:', err);
      this.recording = false;
      if (this.onError) {
        try { this.onError({ phase: 'start', error: err.message }); } catch (_) {}
      }
      return false;
    }
  }

  static _parseChange(change, timestamp) {
    if (!change) return null;

    const entry = { t: timestamp, k: 'text', ops: [] };

    try {
      // Some versions wrap the change in an extra array layer
      const changes = Array.isArray(change) ? change : [change];
      for (const c of changes) {
        entry.ops.push({
          // Support various property names for compatibility across Word API versions
          d: c.oldText !== undefined ? c.oldText : (c.before || ''),
          i: c.newText !== undefined ? c.newText : (c.after || ''),
          o: c.id || 0,
          r: c.rangeId || ''
        });
      }
    } catch (err) {
      console.error('[WordTimestamp] Error parsing change:', err);
      entry.ops = [{ d: '', i: String(change).slice(0, 50), o: 0 }];
    }

    return entry;
  }

  stop() {
    if (!this.recording) return false;
    const self = this;
    Word.run(async (ctx) => {
      try {
        const doc = ctx.document;
        if (self._handlerRef) {
          doc.onChanged.remove(self._handlerRef);
          self._handlerRef = null;
        }
        await ctx.sync();
      } catch (err) {
        console.warn('[WordTimestamp] Warning during stop:', err);
      }
    }).catch((err) => {
      console.error('[WordTimestamp] Error during stop cleanup:', err);
    });

    this.recording = false;
    this.paused = false;
    return true;
  }

  pause() {
    if (!this.recording || this.paused) return false;
    if (this._pausedAt !== null) this._pauseElapsed += Date.now() - this._pausedAt;
    this.paused = true;
    this._pausedAt = Date.now();
    return true;
  }

  resume() {
    if (!this.recording || !this.paused) return false;
    if (this._pausedAt !== null) this._pauseElapsed += Date.now() - this._pausedAt;
    this.paused = false;
    this._pausedAt = null;
    return true;
  }

  _relativeMs() {
    if (!this.startTime) return 0;
    let elapsed = Date.now() - this.startTime - this._pauseElapsed;
    if (this._pausedAt !== null) elapsed -= Date.now() - this._pausedAt;
    return Math.max(0, Math.round(elapsed));
  }

  clear() {
    this.entries = [];
    this.sessionId = '';
    this.docTitle = null;
    this.recording = false;
    this.paused = false;
    this._handlerRef = null;
  }

  getSnapshot() {
    return {
      version: 1,
      recordedAt: new Date().toISOString(),
      sessionId: this.sessionId,
      docTitle: this.docTitle || 'Untitled',
      entryCount: this.entries.length,
      durationMs: this._relativeMs(),
      entries: [...this.entries]
    };
  }
}

export { Recorder };
export default Recorder;