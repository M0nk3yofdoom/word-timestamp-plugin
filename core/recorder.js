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

    // Capture user identity and doc title from Office context
    try {
      await Word.run(async (ctx) => {
        const doc = ctx.document;
        doc.load('name');
        await ctx.sync();
        this.docTitle = doc.name || 'Untitled';

        // Store handler reference so we can remove it on stop
        const self = this;
        this._handlerRef = (event) => {
          if (!self.recording || self.paused) {
            event.completed?.();
            return;
          }

          try {
            // Process each change in the batch
            const changes = event.all || [];
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

            // Auto-save callback
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

        // Register the change handler
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

  /** Parse a single Word change event into our compact format */
  static _parseChange(change, timestamp) {
    if (!change) return null;

    const entry = {
      t: timestamp,
      k: 'text',
      ops: []
    };

    // Capture old and new text from the changed range
    try {
      const ops = [];
      for (const c of change) {
        const op = {
          d: c.oldText || '',
          i: c.newText || '',
          o: c.id || 0, // Use the change ID as offset reference
          r: c.rangeId || '' // Store range ID for reference
        };
        ops.push(op);
      }
      entry.ops = ops;
    } catch (_) {
      // Fallback: generic change record
      const genericChange = Array.isArray(change) ? change : [change];
      for (const c of genericChange) {
        entry.ops.push({
          d: '',
          i: String(c).slice(0, 100),
          o: 0
        });
      }
    }

    return entry;
  }

  stop() {
    if (!this.recording) return false;

    // Remove the change handler - we need to do this in Word context
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
    console.log(`[WordTimestamp] Recording stopped. Captured ${this.entries.length} events.`);
    return true;
  }

  pause() {
    if (!this.recording || this.paused) return false;
    if (this._pausedAt !== null) {
      this._pauseElapsed += Date.now() - this._pausedAt;
    }
    this.paused = true;
    this._pausedAt = Date.now();
    console.log('[WordTimestamp] Recording paused');
    return true;
  }

  resume() {
    if (!this.recording || !this.paused) return false;
    if (this._pausedAt !== null) {
      this._pauseElapsed += Date.now() - this._pausedAt;
    }
    this.paused = false;
    this._pausedAt = null;
    console.log('[WordTimestamp] Recording resumed');
    return true;
  }

  _relativeMs() {
    if (!this.startTime) return 0;
    let elapsed = Date.now() - this.startTime - this._pauseElapsed;
    if (this._pausedAt !== null) {
      elapsed -= Date.now() - this._pausedAt;
    }
    return Math.max(0, Math.round(elapsed));
  }

  clear() {
    this.entries = [];
    this.sessionId = '';
    this.docTitle = null;
    this.userName = null;
    this.recording = false;
    this.paused = false;
    this.startTime = null;
    this._pauseElapsed = 0;
    this._pausedAt = null;
    this._handlerRef = null;
  }

  /** Get a snapshot of current recording data for export */
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
