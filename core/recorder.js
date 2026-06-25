/** src/core/recorder.js - Captures Word document changes with precise timestamps for authorship evidence */

class Recorder {
  constructor() {
    this.recording = false;
    this.paused = false;
    this.startTime = null;
    this.sessionId = '';
    this.docTitle = null;
    this.userName = null;
    this.authorId = null;
    this.authorName = null;
    this.entries = [];
    this._pauseElapsed = 0;
    this._pausedAt = null;
    this._handlerRef = null;
    this.onFlush = null;
    this.onError = null;
    this.sessionStartTime = null;
  }

  async start() {
    if (this.recording) return false;

    this.sessionId = typeof crypto !== 'undefined' && crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    this.entries = [];
    this.recording = true;
    this.paused = false;
    this.startTime = Date.now();
    this.sessionStartTime = new Date().toISOString();
    this._pauseElapsed = 0;
    this._pausedAt = null;

    // Signal that recording state has changed
    if (this.onFlush) this.onFlush();

    try {
      // PHASE 1: Get the document meta-data and user info (async is fine here)
      await Word.run(async (ctx) => {
        const doc = ctx.document;
        doc.load('name');
        await ctx.sync();
        this.docTitle = doc.name || 'Untitled';
        
        // Try to get user info if available
        try {
          const userProfile = await ctx.document.getFilePropertiesAsync();
          this.authorName = userProfile.author || 'Unknown Author';
        } catch (e) {
          this.authorName = 'Unknown Author';
        }
      });

      // PHASE 2: Define the handler logic inside-this scope synchronously
      const self = this;
      this._handlerRef = (event) => {
        if (!self.recording || self.paused) {
          try { event.completed(); } catch(e) {}
          return;
        }

        try {
          const changes = event?.changes || [];
          for (let i = 0; i < changes.length; i++) {
            const change = changes[i];
            const entry = Recorder._parseChange(change, self._relativeMs());
            if (entry) self.entries.push(entry);
          }

          // Pulse the UI layer via microtask to keep loop responsive
          queueMicrotask(() => {
            if (self.onFlush) self.onFlush();
          });
        } catch (err) {
          console.error('[WordTimestamp] Change processing error:', err);
          if (self.onError) self.onError({ phase: 'recording', error: err.message });
        } finally {
          try { event.completed(); } catch(e) {}
        }
      };

      // PHASE 3: Register within a dedicated Synchronous Runtime call
      // This is the core fix for the "no events found" issue.
      await Word.run(async (ctx) => {
        ctx.document.onChanged.add(this._handlerRef);
        await ctx.sync();
      });

      console.log(`[WordTimestamp] Session Started: ${this.sessionId} by ${this.authorName}`);
      return true;

    } catch (err) {
      console.error('[WordTimestamp] Start failure:', err);
      this.recording = false;
      if (this.onError) this.onError({ phase: 'start', error: err.message });
      return false;
    }
  }

  static _parseChange(change, timestamp) {
    if (!change) return null;
    
    // Determine change type
    let changeType = 'text';
    let location = 'unknown';
    let paragraphIndex = -1;
    let sectionIndex = -1;
    
    if (change) {
      // Try to determine what type of change occurred
      if (change.newText !== undefined && change.oldText !== undefined) {
        changeType = 'text';
      } else if (change.format !== undefined) {
        changeType = 'formatting';
      } else if (change.image !== undefined) {
        changeType = 'image';
      }
      
      // Try to extract location info
      if (change.paragraphIndex !== undefined) {
        paragraphIndex = change.paragraphIndex;
      }
      if (change.sectionIndex !== undefined) {
        sectionIndex = change.sectionIndex;
      }
      if (change.location) {
        location = change.location;
      }
    }
    
    const entry = { 
      t: timestamp, 
      k: changeType,
      loc: location,
      para: paragraphIndex,
      sect: sectionIndex,
      ops: [] 
    };
    
    entry.ops.push({
      d: change.oldText || '',
      i: change.newText || '',
      o: change.id || 0,
      r: change.rangeId || '',
      fmt: change.format || null,
      img: change.image || null
    });
    
    return entry;
  }

  setAuthorInfo(authorId, authorName) {
    this.authorId = authorId;
    this.authorName = authorName;
  }

  getAuthorInfo() {
    return {
      authorId: this.authorId,
      authorName: this.authorName
    };
  }

  stop() {
    if (!this.recording) return false;
    const self = this;

    // Synchronous removal of the handler bridge
    Word.run((ctx) => {
      if (self._handlerRef) {
        ctx.document.onChanged.remove(self._handlerRef);
        self._handlerRef = null;
      }
      return ctx.sync();
    }).catch(e => console.warn('[WordTimestamp] Stop cleanup incomplete:', e));

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
    // Force cleanup of existing listener before clearing state
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
    this._handlerRef = null;
    this.startTime = null;
    this._pauseElapsed = 0;
    this._pausedAt = null;
    this.sessionStartTime = null;
  }

  getSnapshot() {
    return {
      version: 1,
      recordedAt: new Date().toISOString(),
      sessionStartTime: this.sessionStartTime,
      sessionId: this.sessionId,
      docTitle: this.docTitle || 'Untitled',
      authorId: this.authorId,
      authorName: this.authorName,
      entryCount: this.entries.length,
      durationMs: this._relativeMs(),
      entries: [...this.entries]
    };
  }
}

export { Recorder };
export default Recorder;
