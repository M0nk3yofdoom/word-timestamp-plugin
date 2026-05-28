/** src/core/player.js — Timeline playback engine */

class Player {
  constructor() {
    this.entries = [];
    this.isPlaying = false;
    this.speed = 1;
    this.positionMs = 0;
    this._rafId = null;
    this._lastFrameTime = null;
    this.tickCallback = null;
   }

  load(entries) {
    this.entries = [...entries];
    this.pause();
    this.positionMs = 0;
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

  duration() {
    if (!this.entries.length) return 0;
    const last = this.entries[this.entries.length - 1];
   return Math.round(last.t + 50);
    }

  entriesAt(pos) { return this.entries.filter(e => e.t <= pos).sort((a, b) => a.t - b.t); }

  clear() { this.pause(); this.entries = []; this.positionMs = 0; }

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
      if (maxT > 0 && this.positionMs >= maxT) { this.pause(); return; }

     this._loop();
     });
    }
}

export { Player };
export default Player;
