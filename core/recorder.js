/** src/core/recorder.js - Captures Word document change events with timestamps */

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
    this._handlerRef = null;
    this.onFlush = null;
    this.onError = null;
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
       // STEP 1: Get document title (async is fine for data reads)
      await Word.run(async (ctx) => {
        const doc = ctx.document;
        doc.load('name');
        await ctx.sync();
        this.docTitle = doc.name || 'Untitled';
        });

       // STEP 2: Define the event handler
       // IMPORTANT: The callback itself can reference async state freely,
       // but it must be DEFINED and ADDED in a synchronous Word.run block.
      const self = this;
      this._handlerRef = (event) => {
        if (!self.recording || self.paused) {
          event.completed?.();
          return;
          }

        try {
           // event.changes is an array of OnChangedEventDetails objects, each with:
           //    oldText - text before the change
           //    newText - text after the change
          const changes = event?.changes || [];

          for (let i = 0; i < changes.length; i++) {
            const change = changes[i];
            const entry = Recorder._parseChange(change, self._relativeMs());
            if (entry) {
              self.entries.push(entry);
              }
            }

           // Notify UI of new entries
          if (self.onFlush && typeof self.onFlush === 'function') {
            try { self.onFlush(); } catch (_) {}
            }

          } catch (err) {
          console.error('[WordTimestamp] Error processing change event:', err);
          if (self.onError) {
            try { self.onError({ phase: 'recording', error: err.message }); } catch (_) {}
            }
          }

         // Call completed() to signal Office.js this batch is done.
        event.completed?.();
        };

       // STEP 3: Register handler in a SYNCHRONOUS Word.run block and sync.
       // Office.js proxy handlers MUST be added in the same synchronous call stack
       // as the context object. Must also call ctx.sync() to persist the registration.
      await Word.run(async (ctx) => {
        ctx.document.onChanged.add(this._handlerRef);
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

  /** Parse a single Word OnChangedEventDetails change into our compact format */
  static _parseChange(change, timestamp) {
    if (!change) return null;

    const entry = { t: timestamp, k: 'text', ops: [] };
    const op = {
      d: change.oldText !== undefined ? change.oldText : '',
      i: change.newText !== undefined ? change.newText : '',
      o: 0,
      r: change.rangeId || ''
      };
    entry.ops.push(op);
    return entry;
    }

  stop() {
    if (!this.recording) return false;

     // Remove handler in a synchronous Word.run block (same async-boundary rule).
    const self = this;
    try {
      Word.run((ctx) => {
        if (self._handlerRef) {
          ctx.document.onChanged.remove(self._handlerRef);
          self._handlerRef = null;
          }
        });
       } catch (err) {
      console.error('[WordTimestamp] Error during stop:', err);
       }

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
    this.userName = null;
    this.recording = false;
    this.paused = false;
    this._handlerRef = null;
    this.startTime = null;
    this._pauseElapsed = 0;
    this._pausedAt = null;
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

 /** Process an externally-sourced entry (for importing events) */
  processEntry(entry) {
    if (!entry) return;
    this.entries.push(entry);
    if (this.onFlush && typeof this.onFlush === 'function') {
      try { this.onFlush(); } catch (_) {}
      }
    }
}

export { Recorder };
export default Recorder;
