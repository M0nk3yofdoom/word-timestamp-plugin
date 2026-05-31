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
      // STEP 1: Get document title in a single sync-profile run
      await Word.run(async (ctx) => {
        const doc = ctx.document;
        doc.load('name');
        await ctx.sync();
        this.docTitle = doc.name || 'Untitled';
      });

      // STEP 2: Define the event handler synchronously
      const self = this;
      this._handlerRef = (event) => {
        if (!self.recording || self.paused) {
          try { event.completed(); } catch (_) {}
          return;
        }

        try {
          // Process multiple changes in a batch
          const changes = event?.changes || [];
          for (let i = 0; i < changes.length; i++) {
            const change = changes[i];
            const entry = Recorder._parseChange(change, self._relativeMs());
            if (entry) self.entries.push(entry);
          }

          // Defer UI updates to avoid blocking the event processing loop
          queueMicrotask(() => {
            if (self.onFlush && typeof self.onFlush === 'function') {
              try { self.onFlush(); } catch (_) {}
            }
          });

        } catch (err) {
          console.error('[WordTimestamp] Error processing change event:', err);
          if (self.onError) {
            try { self.onError({ phase: 'recording', error: err.message }); } catch (_) {}
          }
        } finally {
          // Signal completion to the Office JS engine
          try { event.completed(); } catch (_) {}
        }
      };

      // STEP 3: Register handler in a sync-profile Word.run call
      await Word.run((ctx) => {
        ctx.document.onChanged.add(this._handlerRef);
        return ctx.sync();
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
    entry.ops.push({
      d: change.oldText || '',
      i: change.newText || '',
      o: change.id || 0,
      r: change.rangeId || ''
    });
    return entry;
  }

  stop() {
    if (!this.recording) return false;
    const self = this;

    // Remove handler synchronously via a fresh context to ensure stability
    Word.run((ctx) => {
      if (self._handlerRef) {
        ctx.document.onChanged.remove(self._handlerRef);
        self._handlerRef = null;
      }
      return ctx.sync();
    }).catch(err => console.warn('[WordTimestamp] Cleanup error:', err));

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
    // Ensure listener is detached first if it was active
    if (this.recording && this._handlerRef) {
      // This is slightly simplified for clear(); we'll rely on stop doing the heavy lifting
      // But let's ensure there's no dangling reference
      this._handlerRef = null;
    }

    this.entries = [];
    this.sessionId = '';
    this.docTitle = null;
    this.userName = null;
    this.recording = false;
    this.paused = false;
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

  processEntry(entry) {
    if (!entry) return;
    this.entries.push(entry);
    queueMicrotask(() => {
      if (this.onFlush && typeof this.onFlush === 'function') {
        try { this.onFlush(); } catch (_) {}
      }
    });
  }
}

export { Recorder };
export default Recorder;
