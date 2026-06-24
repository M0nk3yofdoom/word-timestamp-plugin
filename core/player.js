/** src/core/player.js — Timeline playback engine with enhanced navigation */

class Player {
  constructor() {
    this.entries = [];
    this.isPlaying = false;
    this.speed = 1;
    this.positionMs = 0;
    this._rafId = null;
    this._lastFrameTime = null;
    this.tickCallback = null;
    this.timelineZoom = 1;
    this.authorTracks = {};
    this.currentAuthor = null;
  }

  load(entries, metadata = {}) {
    this.entries = [...entries];
    this.timelineMetadata = metadata;
    this.pause();
    this.positionMs = 0;
    this.authorTracks = {};

    // Group entries by author for tracking
    this.entries.forEach(entry => {
      const author = entry.authorName || 'unknown';
      if (!this.authorTracks[author]) {
        this.authorTracks[author] = [];
      }
      this.authorTracks[author].push(entry);
    });

    if (this.tickCallback) {
      try { this.tickCallback({ position: 0, playing: false }); } catch (_) {}
    }
  }

  play() {
    if (this.isPlaying) return;
    const maxT = this.duration();
    if (this.positionMs >= maxT && maxT > 0) { this.positionMs = 0; }
    this.isPlaying = true;
    this._lastFrameTime = performance.now();
    this._loop();
  }

  pause() {
    this.isPlaying = false;
    if (this._rafId !== null) { cancelAnimationFrame(this._rafId); this._rafId = null; }
  }

  setSpeed(speed) { this.speed = Math.max(0.25, Math.min(32, speed)); }

  setPosition(ms) {
    this.positionMs = Math.max(0, ms);
    if (this.tickCallback) { try { this.tickCallback({ position: this.positionMs, playing: this.isPlaying }); } catch (_) {} }
  }

  setZoom(zoom) {
    this.timelineZoom = Math.max(0.1, Math.min(10, zoom));
  }

  duration() {
    if (!this.entries.length) return 0;
    const last = this.entries[this.entries.length - 1];
    return Math.round(last.t + 50);
  }

  entriesAt(pos) {
    return this.entries.filter(e => e.t <= pos).sort((a, b) => a.t - b.t);
  }

  getAllTimestamps() {
    return this.entries.map((e, idx) => ({ t: e.t, idx }));
  }

  prevChange(stepMs = 100) {
    const prevEntry = this.getPreviousEntry();
    if (prevEntry) {
      this.positionMs = prevEntry.t;
      if (this.tickCallback) {
        try { this.tickCallback({ position: this.positionMs, playing: this.isPlaying }); } catch (_) {}
      }
    }
  }

  nextChange(stepMs = 100) {
    const nextEntry = this.getNextEntry();
    if (nextEntry) {
      this.positionMs = nextEntry.t;
      if (this.tickCallback) {
        try { this.tickCallback({ position: this.positionMs, playing: this.isPlaying }); } catch (_) {}
      }
    }
  }

  getPreviousEntry() {
    for (let i = this.entries.length - 1; i >= 0; i--) {
      if (this.entries[i].t < this.positionMs) {
        return this.entries[i];
      }
    }
    return null;
  }

  getNextEntry() {
    for (let i = 0; i < this.entries.length; i++) {
      if (this.entries[i].t > this.positionMs) {
        return this.entries[i];
      }
    }
    return null;
  }

  goToEntry(index) {
    if (index < 0 || index >= this.entries.length) return;
    this.positionMs = this.entries[index].t + 1;
    if (this.tickCallback) {
      try { this.tickCallback({ position: this.positionMs, playing: this.isPlaying }); } catch (_) {}
    }
  }

  getEntriesByAuthor(authorName) {
    return this.authorTracks[authorName] || [];
  }

  clear() {
    this.pause();
    this.entries = [];
    this.positionMs = 0;
    this.authorTracks = {};
    this.timelineZoom = 1;
  }

  _loop() {
    if (!this.isPlaying) return;

    this._rafId = requestAnimationFrame(() => {
      const now = performance.now();
      const dt = now - (this._lastFrameTime || now);
      this._lastFrameTime = now;
      this.positionMs += dt * this.speed;

      if (this.tickCallback) {
        try { this.tickCallback({ position: this.positionMs, playing: true }); } catch (_) {}
      }

      const maxT = this.duration();
      if (maxT > 0 && this.positionMs >= maxT) {
        this.pause();
        return;
      }

      this._loop();
    });
  }
}

export { Player };
export default Player;
