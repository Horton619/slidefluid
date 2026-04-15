'use strict';

/**
 * SlideFluid 3.0 — Electron main process
 *
 * Responsibilities:
 *   - Create and manage the BrowserWindow
 *   - Resolve paths to bundled Python backend and Poppler binaries
 *   - Spawn the Python subprocess and pipe IPC JSON back to the renderer
 *   - Handle OS-level file dialogs (open, folder picker)
 *   - Persist settings via electron-store (app preferences)
 *   - Open the output folder in Finder / Explorer
 */

const {
  app,
  BrowserWindow,
  ipcMain,
  dialog,
  shell,
  nativeTheme,
} = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const os = require('os');
const { autoUpdater } = require('electron-updater');

// ---------------------------------------------------------------------------
// Simple JSON settings store (no external dep — we write it ourselves)
// ---------------------------------------------------------------------------

class SettingsStore {
  constructor() {
    const userData = app.getPath('userData');
    this.filePath = path.join(userData, 'slidefluid-settings.json');
    this._data = this._load();
  }

  _defaults() {
    return {
      outputDir: null,
      dpi: 144,
      fillMode: 'black',
      autoOpenOnComplete: true,
      overwriteBehavior: 'ask',      // ask | overwrite | skip | rename
      filenameSuffix: '',
      skin: 'professional',          // professional | fun
      previewMode: 'graphical',      // graphical | live  (live = PDF.js, Phase 8)
      writeReport: true,             // write a .txt conversion report to output folder
    };
  }

  _load() {
    try {
      if (fs.existsSync(this.filePath)) {
        const raw = fs.readFileSync(this.filePath, 'utf8');
        return Object.assign(this._defaults(), JSON.parse(raw));
      }
    } catch (e) {
      console.warn('Settings load failed, using defaults:', e.message);
    }
    return this._defaults();
  }

  save() {
    try {
      fs.mkdirSync(path.dirname(this.filePath), { recursive: true });
      fs.writeFileSync(this.filePath, JSON.stringify(this._data, null, 2), 'utf8');
    } catch (e) {
      console.error('Settings save failed:', e.message);
    }
  }

  get(key) { return this._data[key]; }
  set(key, value) { this._data[key] = value; this.save(); }
  getAll() { return { ...this._data }; }
  setAll(obj) { Object.assign(this._data, obj); this.save(); }
}

// ---------------------------------------------------------------------------
// Logger — writes timestamped lines to userData/slidefluid.log
// ---------------------------------------------------------------------------

class Logger {
  constructor(filePath) {
    this.filePath = filePath;
    try {
      fs.mkdirSync(path.dirname(filePath), { recursive: true });
      this._rotate();
    } catch (_) {}
  }

  _rotate() {
    if (!fs.existsSync(this.filePath)) return;
    try {
      if (fs.statSync(this.filePath).size > 2 * 1024 * 1024) {
        const old = this.filePath + '.1';
        if (fs.existsSync(old)) fs.unlinkSync(old);
        fs.renameSync(this.filePath, old);
      }
    } catch (_) {}
  }

  _write(level, msg) {
    const ts = new Date().toISOString();
    try {
      fs.appendFileSync(this.filePath, `[${ts}] [${level}] ${msg}\n`, 'utf8');
    } catch (_) {}
  }

  info(msg)  { this._write('INFO',  String(msg)); }
  warn(msg)  { this._write('WARN',  String(msg)); }
  error(msg) { this._write('ERROR', String(msg)); }

  getTail(n = 100) {
    try {
      if (!fs.existsSync(this.filePath)) return [];
      return fs.readFileSync(this.filePath, 'utf8')
        .split('\n').filter(Boolean).slice(-n);
    } catch (e) {
      return [`[Logger read error: ${e.message}]`];
    }
  }
}

// Convenience — safe to call before logger is initialized (falls back to console)
function log(level, msg) {
  if (global.logger) global.logger[level](msg);
  else console[level === 'info' ? 'log' : level](msg);
}

// ---------------------------------------------------------------------------
// Path resolution — works in dev and in packaged app
// ---------------------------------------------------------------------------

/**
 * Resolve the path to the bundled Python backend executable.
 * In dev:       ../backend/slidefluid_convert.py  (run via system python3)
 * In packaged:  <resources>/backend/slidefluid_convert  (PyInstaller binary)
 */
function resolvePythonBackend() {
  if (app.isPackaged) {
    const exe = process.platform === 'win32'
      ? 'slidefluid_convert.exe'
      : 'slidefluid_convert';
    return path.join(process.resourcesPath, 'backend', exe);
  }
  // Dev mode: use system python3 + the script
  return null; // signals to use python3 + script path
}

function resolveBackendScript() {
  // Dev mode only
  return path.join(__dirname, '..', 'backend', 'slidefluid_convert.py');
}

/**
 * Resolve the path to bundled Poppler binaries.
 * In dev:       ../vendor/poppler/<platform>/
 * In packaged:  <resources>/poppler/
 */
function resolvePopplerPath() {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, 'poppler');
  }
  // Dev: check vendor dir, fall back to system Poppler (null = use PATH)
  const platform = process.platform === 'win32' ? 'win' :
                   process.platform === 'darwin' ? 'mac' : 'linux';
  const vendorPath = path.join(__dirname, '..', 'vendor', 'poppler', platform);
  if (fs.existsSync(vendorPath)) return vendorPath;
  return null; // use system Poppler from PATH
}

// ---------------------------------------------------------------------------
// Subprocess manager
// ---------------------------------------------------------------------------

class ConversionJob {
  constructor({ files, outputDir, dpi, fillMode, suffix, overwrite, win, slideTheme }) {
    this.files = files;
    this.outputDir = outputDir;
    this.dpi = dpi;
    this.fillMode = fillMode;
    this.suffix = suffix;
    this.overwrite = overwrite;
    this.win = win;
    this.slideTheme = slideTheme || 'light';
    this.proc = null;
    this.cancelled = false;
  }

  _buildArgs() {
    const backendExe = resolvePythonBackend();
    const popplerPath = resolvePopplerPath();

    const scriptArgs = [
      '--ipc',
      '--dpi', String(this.dpi),
      '--fill', this.fillMode,
      '--slide-theme', this.slideTheme,
      '--output-dir', this.outputDir,
      '--overwrite',   // Electron handles overwrite UX; always pass --overwrite here
    ];

    if (this.suffix) scriptArgs.push('--suffix', this.suffix);
    if (popplerPath) scriptArgs.push('--poppler-path', popplerPath);
    scriptArgs.push(...this.files);

    if (backendExe) {
      // Packaged: run the PyInstaller binary directly
      return { cmd: backendExe, args: scriptArgs };
    } else {
      // Dev: prefer venv python if present, fall back to system python3
      const venvPython = path.join(__dirname, '..', 'venv', 'bin', 'python3');
      const cmd = fs.existsSync(venvPython) ? venvPython : 'python3';
      return { cmd, args: [resolveBackendScript(), ...scriptArgs] };
    }
  }

  start() {
    const { cmd, args } = this._buildArgs();
    log('info', `spawn: ${cmd} ${args.slice(0, 6).join(' ')} ...`);

    this.proc = spawn(cmd, args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      env: { ...process.env },
    });

    let lineBuffer = '';

    this.proc.stdout.on('data', (chunk) => {
      lineBuffer += chunk.toString('utf8');
      let newline;
      while ((newline = lineBuffer.indexOf('\n')) !== -1) {
        const line = lineBuffer.slice(0, newline).trim();
        lineBuffer = lineBuffer.slice(newline + 1);
        if (!line) continue;
        try {
          const msg = JSON.parse(line);
          this._handleMessage(msg);
        } catch (e) {
          log('warn', `non-JSON stdout: ${line}`);
        }
      }
    });

    this.proc.stderr.on('data', (chunk) => {
      const text = chunk.toString('utf8').trim();
      if (text) {
        log('warn', `python stderr: ${text}`);
        this._send('conversion:stderr', { message: text });
      }
    });

    this.proc.on('close', (code, signal) => {
      log('info', `python exited: code=${code} signal=${signal}`);
      this._send('conversion:exit', {
        code,
        signal,
        cancelled: this.cancelled,
      });
      this.proc = null;
    });

    this.proc.on('error', (err) => {
      log('error', `spawn error: ${err.message}`);
      this._send('conversion:spawn_error', { message: err.message });
    });
  }

  cancel() {
    if (this.proc) {
      this.cancelled = true;
      this.proc.kill('SIGTERM');
    }
  }

  _handleMessage(msg) {
    // Forward all IPC messages straight to the renderer
    this._send('conversion:message', msg);
  }

  _send(channel, payload) {
    if (this.win && !this.win.isDestroyed()) {
      this.win.webContents.send(channel, payload);
    }
  }
}

// ---------------------------------------------------------------------------
// Preflight job (reuses subprocess)
// ---------------------------------------------------------------------------

function runPreflight(win) {
  const backendExe = resolvePythonBackend();
  const popplerPath = resolvePopplerPath();

  const scriptArgs = ['--preflight', '--ipc'];
  if (popplerPath) scriptArgs.push('--poppler-path', popplerPath);

  const venvPython = path.join(__dirname, '..', 'venv', 'bin', 'python3');
  const devPython = fs.existsSync(venvPython) ? venvPython : 'python3';
  const cmd = backendExe || devPython;
  const args = backendExe ? scriptArgs : [resolveBackendScript(), ...scriptArgs];

  const proc = spawn(cmd, args, {
    stdio: ['ignore', 'pipe', 'pipe'],
    env: { ...process.env },
  });

  let output = '';
  proc.stdout.on('data', (d) => (output += d.toString()));
  proc.stderr.on('data', (d) => log('warn', `preflight stderr: ${d.toString().trim()}`));

  proc.on('close', (code) => {
    // Parse the preflight_result JSON line
    let results = null;
    for (const line of output.split('\n')) {
      try {
        const msg = JSON.parse(line.trim());
        if (msg.type === 'preflight_result') {
          results = msg.results;
          break;
        }
      } catch (_) {}
    }

    // Append output-folder-writable and disk-space checks (done in main process)
    const settings = global.settings;
    const outputDir = settings ? settings.get('outputDir') : null;

    const folderCheck = checkOutputFolderWritable(outputDir);
    const diskCheck = checkDiskSpace(outputDir);

    if (results) {
      results.output_folder = folderCheck;
      results.disk_space = diskCheck;
      results.app_version = {
        ok: true,
        message: `v${app.getVersion()} — ${process.platform} ${process.arch}`,
      };
    }

    if (win && !win.isDestroyed()) {
      win.webContents.send('preflight:result', { results, exitCode: code });
    }
  });

  proc.on('error', (err) => {
    if (win && !win.isDestroyed()) {
      win.webContents.send('preflight:result', {
        results: null,
        error: err.message,
      });
    }
  });
}

function checkOutputFolderWritable(dir) {
  if (!dir) return { ok: false, message: 'No output folder configured' };
  try {
    const testFile = path.join(dir, `.slidefluid_write_test_${Date.now()}`);
    fs.writeFileSync(testFile, 'ok');
    fs.unlinkSync(testFile);
    return { ok: true, message: dir };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

function checkDiskSpace(dir) {
  // Node doesn't have a native disk-space API; we approximate via fs.statfsSync
  // Available in Node 19+. Fall back gracefully if not present.
  if (!dir) return { ok: false, message: 'No output folder configured' };
  try {
    if (typeof fs.statfsSync === 'function') {
      const stat = fs.statfsSync(dir);
      const freeMB = Math.floor((stat.bavail * stat.bsize) / (1024 * 1024));
      const ok = freeMB >= 500;
      return {
        ok,
        message: `${freeMB} MB free${ok ? '' : ' — WARNING: low disk space'}`,
      };
    }
    return { ok: true, message: 'Disk space check not available on this platform' };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

// ---------------------------------------------------------------------------
// 6.6 — Poppler startup check
// ---------------------------------------------------------------------------

function _checkPopplerBin() {
  return new Promise((resolve) => {
    const popplerPath = resolvePopplerPath();
    const pdfinfoBin = popplerPath ? path.join(popplerPath, 'pdfinfo') : 'pdfinfo';
    const proc = spawn(pdfinfoBin, ['-v'], { stdio: 'ignore' });
    proc.on('error', (err) => {
      // ENOENT = binary not found; any other error means it exists but errored (still ok)
      resolve(err.code !== 'ENOENT' && err.code !== 'EACCES');
    });
    proc.on('close', () => resolve(true));
    // Safety timeout — resolve true so we don't block the app forever
    setTimeout(() => { try { proc.kill(); } catch (_) {} resolve(true); }, 4000);
  });
}

// ---------------------------------------------------------------------------
// Window creation
// ---------------------------------------------------------------------------

let mainWindow = null;
let currentJob = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 640,
    minWidth: 720,
    minHeight: 520,
    backgroundColor: '#070910',
    titleBarStyle: 'default',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));

  // Open DevTools in dev mode
  if (!app.isPackaged) {
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  }

  mainWindow.on('closed', () => {
    if (currentJob) currentJob.cancel();
    mainWindow = null;
  });

  setupAutoUpdater(mainWindow);
}

// ---------------------------------------------------------------------------
// App lifecycle
// ---------------------------------------------------------------------------

app.whenReady().then(async () => {
  global.logger = new Logger(path.join(app.getPath('userData'), 'slidefluid.log'));
  global.logger.info(`--- SlideFluid ${app.getVersion()} started (${process.platform} ${process.arch}) ---`);

  global.settings = new SettingsStore();
  global.logger.info(`Settings loaded from ${global.settings.filePath}`);

  // 6.6 — Silent Poppler preflight before opening the window
  const popplerOk = await _checkPopplerBin();
  if (!popplerOk) {
    await dialog.showMessageBox({
      type: 'error',
      title: 'Poppler not found',
      message: 'SlideFluid requires Poppler to convert PDFs.',
      detail: 'Install Poppler via Homebrew and restart:\n\n  brew install poppler',
      buttons: ['Quit'],
    });
    app.quit();
    return;
  }

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('will-quit', () => {
  if (global.logger) global.logger.info('App quitting');
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ---------------------------------------------------------------------------
// IPC handlers — called from renderer via contextBridge
// ---------------------------------------------------------------------------

// --- Settings ---

ipcMain.handle('settings:getAll', () => {
  return global.settings.getAll();
});

ipcMain.handle('settings:set', (event, key, value) => {
  global.settings.set(key, value);
  return true;
});

ipcMain.handle('settings:setAll', (event, obj) => {
  global.settings.setAll(obj);
  return true;
});

// --- File dialogs ---

ipcMain.handle('dialog:openFiles', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select files to convert',
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Supported Files', extensions: ['pdf', 'docx', 'txt'] },
      { name: 'PDF Files',       extensions: ['pdf'] },
      { name: 'Word Documents',  extensions: ['docx'] },
      { name: 'Text Files',      extensions: ['txt'] },
    ],
  });
  return result.canceled ? [] : result.filePaths;
});

ipcMain.handle('dialog:openFolder', async (event, defaultPath) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Choose output folder',
    defaultPath: defaultPath || app.getPath('documents'),
    properties: ['openDirectory', 'createDirectory'],
  });
  if (result.canceled) return null;
  const folder = result.filePaths[0];
  global.settings.set('outputDir', folder);
  return folder;
});

// --- Shell ---

ipcMain.handle('shell:openFolder', (event, folderPath) => {
  if (folderPath && fs.existsSync(folderPath)) {
    shell.openPath(folderPath);
    return true;
  }
  return false;
});

// --- Conversion ---

ipcMain.handle('conversion:start', async (event, payload) => {
  if (currentJob) {
    return { ok: false, error: 'A conversion is already running.' };
  }

  const {
    files,
    outputDir,
    dpi,
    fillMode,
    suffix,
    slideTheme,
  } = payload;

  // Validate output dir writable before spawning
  const folderCheck = checkOutputFolderWritable(outputDir);
  if (!folderCheck.ok) {
    return { ok: false, error: `Output folder not writable: ${folderCheck.message}` };
  }

  currentJob = new ConversionJob({
    files,
    outputDir,
    dpi,
    fillMode: fillMode || 'black',
    suffix: suffix || '',
    overwrite: true, // Electron renderer handles overwrite UX via overwrite:check
    win: mainWindow,
    slideTheme: slideTheme || 'light',
  });

  currentJob.start();
  return { ok: true };
});

ipcMain.handle('conversion:cancel', () => {
  if (currentJob) {
    currentJob.cancel();
    currentJob = null;
    return true;
  }
  return false;
});

// Called by renderer when a conversion job finishes (cleans up ref)
ipcMain.on('conversion:jobDone', () => {
  currentJob = null;
});

// --- Overwrite check (renderer asks before sending job) ---
ipcMain.handle('overwrite:check', async (event, filePath) => {
  // filePath is the proposed output .pptx path
  if (!fs.existsSync(filePath)) return 'ok'; // no conflict

  const behavior = global.settings.get('overwriteBehavior');
  if (behavior === 'overwrite') return 'overwrite';
  if (behavior === 'skip') return 'skip';
  if (behavior === 'rename') return 'rename';

  // 'ask' — show dialog
  const result = await dialog.showMessageBox(mainWindow, {
    type: 'question',
    title: 'File exists',
    message: `${path.basename(filePath)} already exists in the output folder.`,
    buttons: ['Overwrite', 'Overwrite All', 'Skip', 'Skip All', 'Rename', 'Cancel All'],
    defaultId: 0,
    cancelId: 5,
  });
  return ['overwrite', 'overwrite_all', 'skip', 'skip_all', 'rename', 'cancel'][result.response];
});

// --- Preflight ---

ipcMain.handle('preflight:run', () => {
  runPreflight(mainWindow);
  return true;
});

// --- PDF page image (used by the preview button) ---

ipcMain.handle('pdf:getPageImage', async (event, filePath, page) => {
  try {
    const popplerPath = resolvePopplerPath();
    const pdftoppmBin = popplerPath ? path.join(popplerPath, 'pdftoppm') : 'pdftoppm';
    const pageStr = String(page || 1);
    const tmpPrefix = path.join(os.tmpdir(), `sfpreview_${Date.now()}`);

    await new Promise((resolve, reject) => {
      const proc = spawn(pdftoppmBin, [
        '-r', '96',
        '-f', pageStr, '-l', pageStr,
        '-png',
        filePath,
        tmpPrefix,
      ], { stdio: ['ignore', 'ignore', 'pipe'] });
      proc.on('error', reject);
      proc.on('close', (code) => {
        if (code === 0) resolve();
        else reject(new Error(`pdftoppm exited ${code}`));
      });
    });

    // pdftoppm names files like prefix-1.png / prefix-01.png depending on page count
    const base = path.basename(tmpPrefix);
    const pngFiles = fs.readdirSync(os.tmpdir())
      .filter(f => f.startsWith(base + '-') && f.endsWith('.png'))
      .map(f => path.join(os.tmpdir(), f));

    if (pngFiles.length === 0) throw new Error('No PNG output generated');

    const data = fs.readFileSync(pngFiles[0]);
    pngFiles.forEach(f => { try { fs.unlinkSync(f); } catch (_) {} });

    return { ok: true, dataUrl: `data:image/png;base64,${data.toString('base64')}` };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

// --- PDF info / scan (used by renderer before conversion) ---

ipcMain.handle('pdf:info', async (event, filePath) => {
  try {
    const popplerPath = resolvePopplerPath();
    const pdfinfoBin = popplerPath ? path.join(popplerPath, 'pdfinfo') : 'pdfinfo';
    const { execSync } = require('child_process');
    const output = execSync(`"${pdfinfoBin}" "${filePath}"`, { timeout: 5000 }).toString();

    const pagesMatch = output.match(/Pages:\s*(\d+)/);
    const pageCount = pagesMatch ? parseInt(pagesMatch[1], 10) : null;

    // "Page size:      792 x 612 pts" or "612 x 792 pts (landscape)"
    const sizeMatch = output.match(/Page size:\s*([\d.]+)\s*x\s*([\d.]+)/);
    let ar = null, widthPt = null, heightPt = null;
    if (sizeMatch) {
      widthPt = parseFloat(sizeMatch[1]);
      heightPt = parseFloat(sizeMatch[2]);
      ar = widthPt / heightPt;
    }

    return { ok: true, pageCount, ar, widthPt, heightPt };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('pdf:scan', async (event, itemPaths) => {
  const SUPPORTED_EXTS = /\.(pdf|docx|txt)$/i;

  function scanDir(dir) {
    const results = [];
    try {
      const entries = fs.readdirSync(dir, { withFileTypes: true });
      for (const entry of entries) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) results.push(...scanDir(full));
        else if (SUPPORTED_EXTS.test(entry.name)) results.push(full);
      }
    } catch (_) {}
    return results.sort();
  }

  const result = [];
  const seen = new Set();
  for (const itemPath of itemPaths) {
    try {
      const stat = fs.statSync(itemPath);
      const paths = stat.isDirectory() ? scanDir(itemPath) : [itemPath];
      for (const p of paths) {
        if (!seen.has(p) && SUPPORTED_EXTS.test(p)) {
          seen.add(p);
          result.push(p);
        }
      }
    } catch (_) {}
  }
  return result;
});

// --- DOCX / TXT info ---

ipcMain.handle('docx:info', async (event, filePath) => {
  const backendExe = resolvePythonBackend();
  const venvPython = path.join(__dirname, '..', 'venv', 'bin', 'python3');
  const devPython  = fs.existsSync(venvPython) ? venvPython : 'python3';
  const cmd  = backendExe || devPython;
  const args = backendExe
    ? ['--docx-info', filePath]
    : [resolveBackendScript(), '--docx-info', filePath];

  return new Promise((resolve) => {
    let output = '';
    const proc = spawn(cmd, args, { stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env } });
    proc.stdout.on('data', (d) => (output += d.toString()));
    proc.stderr.on('data', (d) => log('warn', `docx:info stderr: ${d.toString().trim()}`));
    proc.on('close', () => {
      for (const line of output.split('\n')) {
        try {
          const msg = JSON.parse(line.trim());
          if (msg.type === 'docx_info') {
            resolve({ ok: msg.ok, slideCount: msg.slideCount, wordCount: msg.wordCount, message: msg.message });
            return;
          }
        } catch (_) {}
      }
      resolve({ ok: false, error: 'No docx_info response from backend' });
    });
    proc.on('error', (err) => resolve({ ok: false, error: err.message }));
  });
});

// --- Log ---

ipcMain.handle('log:getPath', () => global.logger ? global.logger.filePath : null);

ipcMain.handle('log:getTail', (event, n) =>
  global.logger ? global.logger.getTail(n || 100) : []
);

ipcMain.handle('log:openFile', () => {
  if (global.logger && fs.existsSync(global.logger.filePath)) {
    shell.openPath(global.logger.filePath);
    return true;
  }
  return false;
});

// --- App info ---

ipcMain.handle('app:getVersion', () => app.getVersion());
ipcMain.handle('app:getPlatform', () => ({
  platform: process.platform,
  arch: process.arch,
  version: app.getVersion(),
}));

// --- File write (used by renderer to save conversion report) ---

ipcMain.handle('fs:writeFile', async (event, filePath, content) => {
  try {
    fs.writeFileSync(filePath, content, 'utf8');
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

// --- Auto-updater ---

autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = true;

function setupAutoUpdater(win) {
  const send = (payload) => {
    if (win && !win.isDestroyed()) win.webContents.send('update:status', payload);
  };

  autoUpdater.on('checking-for-update',  ()     => send({ state: 'checking' }));
  autoUpdater.on('update-not-available', ()     => send({ state: 'current', version: app.getVersion() }));
  autoUpdater.on('update-available',     (info) => send({ state: 'available', version: info.version }));
  autoUpdater.on('update-downloaded',    (info) => send({ state: 'ready', version: info.version }));
  autoUpdater.on('download-progress',    (p)    => send({ state: 'downloading', percent: Math.round(p.percent) }));
  autoUpdater.on('error', (err) => {
    log('warn', `Auto-updater error: ${err.message}`);
    send({ state: 'error', message: err.message });
  });
}

ipcMain.handle('update:check', () => {
  try { autoUpdater.checkForUpdates(); } catch (_) { /* no-op in dev/unsigned builds */ }
  return true;
});

ipcMain.handle('update:download', () => {
  try { autoUpdater.downloadUpdate(); } catch (_) {}
  return true;
});

ipcMain.on('update:install', () => {
  autoUpdater.quitAndInstall();
});
