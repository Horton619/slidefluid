'use strict';

/**
 * SlideFluid 3.0 — Preload script
 *
 * Exposes a narrow, typed API surface to the renderer via contextBridge.
 * The renderer never has direct access to Node or Electron internals.
 *
 * All methods are async (return Promises) unless noted.
 * Event listeners use a subscribe/unsubscribe pattern.
 */

const { contextBridge, ipcRenderer } = require('electron');

// ---------------------------------------------------------------------------
// Helper — one-shot typed event subscription
// ---------------------------------------------------------------------------

function makeListener(channel) {
  return (callback) => {
    const handler = (event, payload) => callback(payload);
    ipcRenderer.on(channel, handler);
    // Return an unsubscribe function
    return () => ipcRenderer.removeListener(channel, handler);
  };
}

// ---------------------------------------------------------------------------
// Exposed API
// ---------------------------------------------------------------------------

contextBridge.exposeInMainWorld('slidefluid', {

  // --- Settings ---

  /** Returns the full settings object */
  getSettings: () => ipcRenderer.invoke('settings:getAll'),

  /** Set a single setting key/value */
  setSetting: (key, value) => ipcRenderer.invoke('settings:set', key, value),

  /** Bulk-update settings */
  setSettings: (obj) => ipcRenderer.invoke('settings:setAll', obj),


  // --- File dialogs ---

  /** Open a multi-select PDF file picker. Returns string[] of paths. */
  openFilePicker: () => ipcRenderer.invoke('dialog:openFiles'),

  /**
   * Open a folder picker. Persists the chosen folder in settings.
   * Returns the chosen path string, or null if cancelled.
   */
  openFolderPicker: (defaultPath) =>
    ipcRenderer.invoke('dialog:openFolder', defaultPath),


  // --- Shell ---

  /** Open a folder in Finder / Explorer. */
  openFolder: (folderPath) => ipcRenderer.invoke('shell:openFolder', folderPath),


  // --- Conversion ---

  /**
   * Start a conversion batch.
   * @param {object} payload
   *   files      string[]  — absolute paths to PDF files
   *   outputDir  string    — absolute path to output folder
   *   dpi        72|144
   *   fillMode   'black'|'color_match'|'smear'
   *   suffix     string    — filename suffix (may be '')
   * @returns {Promise<{ok: boolean, error?: string}>}
   */
  startConversion: (payload) => ipcRenderer.invoke('conversion:start', payload),

  /** Cancel the running conversion. */
  cancelConversion: () => ipcRenderer.invoke('conversion:cancel'),

  /** Notify main that the job is done (clears the ref). */
  notifyJobDone: () => ipcRenderer.send('conversion:jobDone'),

  /**
   * Check whether an output file already exists and what to do.
   * Returns: 'ok' | 'overwrite' | 'skip' | 'rename' | 'cancel'
   */
  checkOverwrite: (filePath) => ipcRenderer.invoke('overwrite:check', filePath),


  // --- Conversion events (subscribe returns an unsubscribe fn) ---

  /**
   * Fired for each IPC JSON line from the Python backend.
   * Payload mirrors the Python schema:
   *   {type: 'start'|'progress'|'done'|'error'|'batch_done', ...}
   */
  onConversionMessage: makeListener('conversion:message'),

  /** Python process stderr output (usually warnings). */
  onConversionStderr: makeListener('conversion:stderr'),

  /**
   * Fired when the Python process exits.
   * Payload: {code: number, signal: string|null, cancelled: boolean}
   */
  onConversionExit: makeListener('conversion:exit'),

  /** Fired if spawn itself fails (e.g. binary not found). */
  onConversionSpawnError: makeListener('conversion:spawn_error'),


  // --- Preflight ---

  /** Trigger a preflight check. Result arrives via onPreflightResult. */
  runPreflight: () => ipcRenderer.invoke('preflight:run'),

  /**
   * Fired when preflight completes.
   * Payload: {results: {[check]: {ok, message}}, exitCode: number}
   */
  onPreflightResult: makeListener('preflight:result'),


  // --- Log ---

  /** Returns the absolute path to the log file. */
  getLogPath: () => ipcRenderer.invoke('log:getPath'),

  /** Returns the last n lines of the log as string[]. */
  getLogTail: (n) => ipcRenderer.invoke('log:getTail', n || 100),

  /** Opens the log file in the system default text editor. */
  openLogFile: () => ipcRenderer.invoke('log:openFile'),

  // --- App info ---

  getVersion: () => ipcRenderer.invoke('app:getVersion'),
  getPlatform: () => ipcRenderer.invoke('app:getPlatform'),

  // --- PDF utilities ---

  /** Get page count and aspect ratio for a single PDF (uses pdfinfo). */
  getPdfInfo: (filePath) => ipcRenderer.invoke('pdf:info', filePath),

  /** Expand an array of file/folder paths to a flat list of supported files (PDF, DOCX, TXT). */
  scanPaths: (paths) => ipcRenderer.invoke('pdf:scan', paths),

  /**
   * Get slide count and word count for a .txt or .docx file.
   * Returns {ok, slideCount, wordCount} or {ok: false, error}.
   */
  getDocxInfo: (filePath) => ipcRenderer.invoke('docx:info', filePath),

  /**
   * Render a single PDF page as a PNG data URL (uses pdftoppm).
   * @param {string} filePath
   * @param {number} page  1-based page number
   * @returns {Promise<{ok: boolean, dataUrl?: string, error?: string}>}
   */
  getPdfPageImage: (filePath, page) =>
    ipcRenderer.invoke('pdf:getPageImage', filePath, page || 1),

  // --- File system ---

  /** Write a UTF-8 text file. Used for conversion reports. */
  writeFile: (filePath, content) => ipcRenderer.invoke('fs:writeFile', filePath, content),

  // --- Auto-updater ---

  /** Trigger an update check against GitHub Releases. */
  checkForUpdates: () => ipcRenderer.invoke('update:check'),

  /** Begin downloading an available update. */
  downloadUpdate: () => ipcRenderer.invoke('update:download'),

  /** Quit and install a downloaded update immediately. */
  installUpdate: () => ipcRenderer.send('update:install'),

  /**
   * Fired when update state changes.
   * Payload: { state: 'checking'|'current'|'available'|'downloading'|'ready'|'error', version?, percent?, message? }
   */
  onUpdateStatus: makeListener('update:status'),
});
