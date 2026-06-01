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
      // STEP 1: Fetch doc title in an async context
      await Word.run(async (ctx) => {
        const doc = ctx.document;
        doc.load('name');
        await ctx.sync();
        this.docTitle = doc.name || 'Untitled';
      });

      // STEP 2: Create the handler function (it's a closure)
      const self = this;
      this._handlerRef = (event) => {
        if (!self.recording || self.paused) {
          try { event.completed(); } catch (_) {}
          return;
        }

        try {
          const changes = event?.changes || [];
          for (let i = 0; i < changes.length; i++) {
            const change = changes[i];
            const entry = Recorder._parseChange(change, self._relativeMs());
            if (entry) self.entries.push(entry);
          }

          // Defer UI updates to avoid blocking the event loop
          queueMicrotask(() => {
            if (self.onFlush && typeof self.onFlush === 'function') {
              try { self.onFlush(); } catch (_) {}
            }
          });

        } catch (err) {
          console.error('[WordTimestamp] Event processing error:', err);
          if (self.onError) {
            try { self.onError({ phase: 'recording', error: err.message }); } catch (_) {}
          }
        } finally {
          // IMPORTANT: Signal completion to the Office host engine
          try { event.completed(); } catch (_) {}
        }
      };

      // STEP 3: Register handler in a synchronous context boundary.
      // We must return from Word.run before we enter any further await points for this to be stable.
      await Word.run((ctx) => {
        ctx.document.onChanged.add(this._handlerRef);
        return ctx.sync();
      });

      console.log(`[WordTimestamp] Recording started: ${this.sessionId}`);
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

    // Remove handler using a synchronous run to ensure the proxy is detached correctly
    Word.run((ctx) => {
      if (self._handlerRef) {
        ctx.document.onChanged.remove(self._handlerRef);
        self._handlerRef = null;
      }
      return ctx.sync();
    }).catch(err => console.warn('[WordTimestamp] Stop cleanup warning:', err));

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
    if (this.recording && this._handlerRef) {
      Word.run((ctx) => {
        ctx.document.onChanged.remove(this._handlerRef);
        this._handlerRef = null;
        return ctx.sync();
      }).catch(() => {}); 
    }
    this.entries = [];
    this.sessionId = '';
    this.docTitle = null;
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
}

export { Recorder };
export default Recorder;
