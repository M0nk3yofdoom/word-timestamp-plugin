/** src/core/format_encoding.js — Export/import + integrity hash */

/** Compute SHA-256 hex digest of a string */
export async function sha256hex(text) {
  const data = new TextEncoder().encode(text);
  const hash = await crypto.subtle.digest('SHA-256', data);
  return Array.from(new Uint8Array(hash))
       .map(b => b.toString(16).padStart(2, '0'))
       .join('');
}

/** Build a complete session object ready for storage/export */
export function buildSession(recorder) {
  const snapshot = recorder.getSnapshot();
  return {
    version: 1,
    recordedAt: snapshot.recordedAt,
    sessionId: snapshot.sessionId,
    docTitle: snapshot.docTitle,
    startTime: snapshot.sessionStartTime,
    authorId: recorder.authorId,
    authorName: recorder.authorName,
    integrityHash: '', // populated after hashing
    entries: snapshot.entries
  };
}

/** Save session to IndexedDB */
export async function saveSession(session, db) {
  const json = JSON.stringify(session);
  session.integrityHash = `sha256-${await sha256hex(json)}`;

  return new Promise((resolve, reject) => {
    const tx = db.transaction(['sessions'], 'readwrite');
    const store = tx.objectStore('sessions');
    const req = store.put({ ...session, sId: session.sessionId });
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(tx.error || new Error('Save failed'));
     });
}

/** Load session from IndexedDB */
export async function loadSession(sessionId, db) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(['sessions'], 'readonly');
    const get = tx.objectStore('sessions').get(sessionId);
    get.onsuccess = () => resolve(get.result || null);
    get.onerror = () => reject(tx.error || new Error('Load failed'));
     });
}

/** Export session as .wtp file (gzip-compressed JSON with integrity hash) */
export async function exportWtp(session) {
   // Build canonical JSON for hashing BEFORE compression
  const text = JSON.stringify({
    version: session.version,
    recordedAt: session.recordedAt,
    sessionId: session.sessionId,
    docTitle: session.docTitle,
    integrityHash: '', // hash will be set separately
    entries: session.entries
    }, null, 2);

    // Compute hash of the uncompressed content
  const hex = await sha256hex(text);
  const fullSession = {
     ...session,
    integrityHash: `sha256-${hex}`
     };
  const finalJson = JSON.stringify(fullSession, null, 2);

    // Compress for storage
  const compressed = await new Response(finalJson)
       .pipeThrough(new CompressionStream('gzip'))
       .arrayBuffer();

  const blob = new Blob([compressed], { type: 'application/gzip' });
  const url = URL.createObjectURL(blob);

  const a = document.createElement('a');
  a.href = url;
  a.download = `${session.docTitle ?? 'recording'}-${session.sessionId}.wtp`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  return {
    filename: `${session.docTitle ?? 'recording'}-${session.sessionId}.wtp`,
    size: compressed.byteLength
     };
}

/** Import .wtp file and parse session data */
export async function importWtp(file) {
  const decompressed = await new Response(file)
       .pipeThrough(new DecompressionStream('gzip'))
       .text();
  const parsed = JSON.parse(decompressed);

    // Verify integrity if hash is present
  if (parsed.integrityHash && typeof parsed.integrityHash === 'string') {
    const tempJson = JSON.stringify({
      version: parsed.version,
      recordedAt: parsed.recordedAt,
      sessionId: parsed.sessionId,
      docTitle: parsed.docTitle,
      startTime: parsed.startTime,
      authorId: parsed.authorId,
      authorName: parsed.authorName,
      integrityHash: '', // strip for comparison
      entries: parsed.entries
    }, null, 2);
    const computed = `sha256-${await sha256hex(tempJson)}`;

    if (computed !== parsed.integrityHash) {
      console.warn('[WordTimestamp] Integrity check failed — data may have been altered');
      return { ...parsed, integrityOk: false };
       }
     }

     return { ...parsed, integrityOk: true };
}

/** List all saved sessions from IndexedDB */
export async function listSessions(db) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(['sessions'], 'readonly');
    const store = tx.objectStore('sessions');
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(tx.error || new Error('List failed'));
     });
}

/** Delete a session from IndexedDB */
export async function deleteSession(sessionId, db) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(['sessions'], 'readwrite');
    const store = tx.objectStore('sessions');
    const req = store.delete(sessionId);
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(tx.error || new Error('Delete failed'));
     });
}
