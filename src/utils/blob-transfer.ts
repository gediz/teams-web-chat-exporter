// Transient Blob handoff between the MV3 service worker and the offscreen
// document.
//
// Why this exists: a Chromium MV3 service worker has no
// URL.createObjectURL, so the only download URL it can build itself is a
// base64 data: URL. For a large export that base64 string exceeds V8's
// maximum string length and throws "Invalid string length" (issue #27).
// The fix is to mint a real blob: URL in the offscreen document, which
// has a DOM. But chrome.runtime messaging cannot carry binary (it
// serializes messages to a JSON-compatible form, which is why the SVG
// rasterizer ships bytes as number[]), and a number[] of a few hundred
// MB is fatal.
//
// IndexedDB solves the transport: it stores a Blob by reference
// (disk-backed), so the SW can park the export Blob here and the
// offscreen document can read it back without ever materializing a giant
// string or array. Entries are removed as soon as they are read (or on
// failure), so nothing is left behind.
//
// Both the SW and the offscreen document run on the same extension
// origin, so they share this database. This module is only reached on
// Chromium MV3 (Firefox MV2's background page has createObjectURL and
// never needs the handoff).

const DB_NAME = 'tce-blob-transfer';
const STORE = 'blobs';
const DB_VERSION = 1;

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('indexedDB.open failed'));
  });
}

/** Park a Blob under `key`. Rejects if IndexedDB is unavailable or quota
 *  is exceeded; callers treat a rejection as "fall back to the data: URL". */
export async function putTransferBlob(key: string, blob: Blob): Promise<void> {
  const db = await openDb();
  try {
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(blob, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error('put failed'));
      tx.onabort = () => reject(tx.error || new Error('put aborted'));
    });
  } finally {
    db.close();
  }
}

/** Read and delete the Blob at `key` in a single transaction. Returns
 *  null when the key is absent or the value is not a Blob. */
export async function takeTransferBlob(key: string): Promise<Blob | null> {
  const db = await openDb();
  try {
    return await new Promise<Blob | null>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      const store = tx.objectStore(STORE);
      const getReq = store.get(key);
      // Delete in the same transaction once the read is queued, so the
      // entry never lingers regardless of how the caller uses the Blob.
      getReq.onsuccess = () => { store.delete(key); };
      tx.oncomplete = () => resolve(getReq.result instanceof Blob ? getReq.result : null);
      tx.onerror = () => reject(tx.error || new Error('take failed'));
      tx.onabort = () => reject(tx.error || new Error('take aborted'));
    });
  } finally {
    db.close();
  }
}

/** Best-effort delete, used to clean up a staged Blob when the offscreen
 *  mint never happened (creation failed, message lost). */
export async function deleteTransferBlob(key: string): Promise<void> {
  const db = await openDb();
  try {
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error('delete failed'));
    });
  } finally {
    db.close();
  }
}

/** Drop every staged Blob. Called once on service-worker startup so an
 *  orphan left by a previous worker that died mid-handoff (between put and
 *  the offscreen read) does not sit in storage. Entries here are always
 *  transient, so clearing on a fresh start is safe. */
export async function clearTransferBlobs(): Promise<void> {
  const db = await openDb();
  try {
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error('clear failed'));
    });
  } finally {
    db.close();
  }
}
