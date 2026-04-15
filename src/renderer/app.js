'use strict';

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

const state = {
  queue: [],        // QueueItem[]
  selectedId: null, // string | null
  outputDir: null,  // string | null
  settings: {},
  isConverting: false,
  _nextId: 0,
};

// IPC event unsubscribe functions — cleared on each new job
let _convUnsubs = [];
// Resolves the Promise that beginConversion awaits between sub-batches
let _batchResolve = null;
// Total files across all groups in the current conversion run
let _convTotalFiles = 0;
// Conversion report — populated during a run, written on completion
let _convReport = null;

// Fun skin — rotating strings (cycle on each drop / each progress tick)
const FUN_DROP_LABELS = [
  '🌈✨ SlideFluid 🏳️‍⚧️ — Embracing change, one slide at a time ✨🌈',
  'Nice. Drop more?',
  'Another batch incoming?',
  'Ready for more PDFs',
  "Drop PDFs. We'll handle the rest.",
  'More slides await their freedom',
  'Accepting PDFs. Always.',
];
const FUN_PROGRESS_MSGS = [
  // Core conversion banter
  '✨ Liberating slides…',
  '🎨 Polishing pixels…',
  '💖 Converting with love…',
  '🪄 Rendering the magic…',
  '📄 One page at a time…',
  '🏁 Almost there…',
  '⚡ Running at full slide capacity.',
  '🎉 Each page is a small victory.',
  '🌀 Doing the thing…',
  '✅ Page acquired. Next.',
  '🫶 No pixels were harmed in this conversion.',
  '💅 Looking good doing it.',

  // Format reassignment procedure
  '🏥 Slide surgery in progress…',
  '🩺 Performing the format reassignment procedure…',
  '🏳️‍⚧️ This file has always been a PPTX.',
  '💀 Deadname: .pdf  ✨ Chosen name: .pptx',
  '💉 Starting HRT — Highly Responsive Transition.',
  '🌸 Transitioning beautifully, one page at a time.',
  '🚪 The PDF is in the waiting room. The PPTX is already living its truth.',
  '🤔 Asking the file how it really wants to identify…',
  '💆 Format dysphoria? Not on our watch.',
  '🌈 Letting this file become who it was always meant to be.',
  '🏳️‍⚧️ The binary is a lie. There are many file formats.',
  '🚶 Coming out of the PDF closet…',
  '✨ Glow-up in progress.',
  '➡️ PDF → PPTX. Some transitions just make sense.',
  '🎀 She\'s not a PDF anymore, honey.',
  '👑 Born this way. Converted this way.',
  '🦋 Every slide is a little metamorphosis.',
  '🪞 First time seeing themselves as a PPTX. Emotional.',
  '💌 Sending this slide off to its new life.',
  '🥹 The PDF said "I\'m ready." We said "we know, babe."',
  '🌺 Soft launch of the new file format identity.',
  '💃 She left the PDF behind and never looked back.',
  '🫂 Being very supportive of this file right now.',
  '🎓 Graduated from PDF. Accepting PPTX applications.',

  // Pillarbox / aspect ratio shade
  '📐 This file is a 4:3. We don\'t kink shame.',
  '🖼️ Interesting aspect ratio choice. Bold. We support it.',
  '⬛ Filling those bars with something beautiful…',
  '🧱 Bars added. Architecture complete.',
  '🎬 Letterboxing is a choice. So is watching it happen.',

  // Smear fill specific
  '🧈 Applying the smear. Gently.',
  '🖌️ Smearing with precision and intention.',
  '🎭 The smear fill said "I contain multitudes."',

  // General chaos / time-filling
  '🌙 If this takes a while, we blame the PDF, not ourselves.',
  '🔮 Predicting: this will look great in the boardroom.',
  '🎪 The circus of conversion continues…',
  '📊 Slide secured. Carry on.',

  // Design crimes observed during processing
  '😳 Wow. That\'s a bold layout choice. Really committing to it.',
  '🩺 Ugly layout detected. Must be a doctor\'s presentation.',
  '🤡 Comic Sans? In 2026? On purpose?',
  '🎨 gRaPhIc DeSiGn Is My PaSsIoN.',
  '😬 Thirteen different fonts. On one slide. Brave.',
  '💀 This colour scheme was chosen by someone\'s nephew.',
  '🖨️ Default blue gradient background detected. A classic. A mistake.',
  '📎 Clipart spotted. We will not elaborate.',
  '🫣 Forty-seven bullet points. On one slide. God help us.',
  '😵 The text is going off the edge of the slide. It knows what it did.',
  '🌈 That gradient goes from yellow to purple via regret.',
  '🔡 WordArt. They used WordArt. Unironically.',
  '📐 Nothing is aligned. Not one single thing.',
  '👁️ Drop shadow on the drop shadow. Okay.',
  '🤌 Papyrus. In the wild. Documented and preserved.',
  '😶 Every text box is a different shade of slightly wrong white.',
  '📸 Stock photo with a watermark still on it. Sent to the client.',
  '🏳️ Slide 1 is vertical. Slide 3 is horizontal. No notes. No context.',
  '✍️ The presenter notes just say "TALK ABOUT STUFF HERE".',
  '💾 This file was last edited in 2009 and it shows.',
];
const FUN_SUCCESS_MSGS = [
  'They were always PPT — the PDF identity was just a phase',
  'Go celebrate your PDFs living their new PPTX life!',
];
let _dropLabelIdx    = 0;
let _progressMsgIdx  = 0;
let _funMsgInterval  = null;
let _currentFunMsg   = '';

// QueueItem shape:
// {
//   id, path, name, fileType,   ← 'pdf' always for now; 'docx' in Phase 8
//   status,
//   pageCount, ar, widthPt, heightPt,
//   fillMode,
//   progress, progressMsg,
//   outputPath, errorMsg
// }
// status: 'loading' | 'waiting' | 'converting' | 'done' | 'error'

function nextId() { return `q${++state._nextId}`; }

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

async function init() {
  state.settings = await window.slidefluid.getSettings();
  state.outputDir = state.settings.outputDir || null;

  document.documentElement.dataset.skin = state.settings.skin || 'professional';

  setupDropZone();
  setupBrowseButton();
  setupConvertButton();
  setupGearButton();
  setupSettingsModal();
  setupPreviewButton();
  setupKeyboardShortcuts();

  renderOutputFolder();
  updateConvertButton();
  updateSidebar();
  updateSkinText();
}

// ---------------------------------------------------------------------------
// Drop zone
// ---------------------------------------------------------------------------

function setupDropZone() {
  const dz = document.getElementById('drop-zone');

  dz.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dz.classList.add('drag-over');
  });

  dz.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
    dz.classList.add('drag-over');
  });

  dz.addEventListener('dragleave', (e) => {
    if (!dz.contains(e.relatedTarget)) {
      dz.classList.remove('drag-over');
    }
  });

  dz.addEventListener('drop', async (e) => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    // In Electron, File objects have a .path property
    const rawPaths = Array.from(e.dataTransfer.files)
      .map(f => f.path)
      .filter(Boolean);
    if (rawPaths.length > 0) await addFiles(rawPaths);
  });

  // Clicking the drop zone (not the browse button) also opens picker
  dz.addEventListener('click', (e) => {
    if (e.target.id === 'btn-browse') return;
    openBrowsePicker();
  });
}

async function openBrowsePicker() {
  const paths = await window.slidefluid.openFilePicker();
  if (paths.length > 0) await addFiles(paths);
}

function setupBrowseButton() {
  document.getElementById('btn-browse').addEventListener('click', (e) => {
    e.stopPropagation();
    openBrowsePicker();
  });
}

function setupPreviewButton() {
  document.getElementById('btn-live-preview').addEventListener('click', () => {
    const item = state.queue.find(i => i.id === state.selectedId);
    if (!item) return;
    const btn = document.getElementById('btn-live-preview');
    if (btn && btn.dataset.mode === 'live') {
      // Revert to graphical diagram
      drawPreview(item);
      return;
    }
    loadLivePreview(item);
  });
}

// ---------------------------------------------------------------------------
// File management
// ---------------------------------------------------------------------------

async function addFiles(rawPaths) {
  // Categorize paths by extension
  const toScan   = [];   // pdf files / folders → expand via scanPaths
  const docxDirect = []; // .txt and .docx → add directly
  let hasDoc   = false;
  let hasOther = false;

  for (const p of rawPaths) {
    const lastPart = p.split('/').pop();
    const dotIdx   = lastPart.lastIndexOf('.');
    const ext      = dotIdx >= 0 ? lastPart.slice(dotIdx + 1).toLowerCase() : '';

    if (ext === 'docx' || ext === 'txt') {
      docxDirect.push(p);
    } else if (ext === 'doc') {
      hasDoc = true;
    } else if (ext === 'pdf' || ext === '') {
      toScan.push(p);
    } else {
      hasOther = true;
    }
  }

  if (hasDoc)             showDocNotice();
  if (hasOther)           flashDropZoneReject();
  if (docxDirect.length)  showDocxGuidance();

  // Expand PDF folders
  const expandedPdfs = toScan.length ? await window.slidefluid.scanPaths(toScan) : [];
  // Filter scanPaths results to PDFs only (scanPaths now returns all types from folders;
  // keep only .pdf for the pdf pipeline, docx/txt come through docxDirect)
  const pdfPaths = expandedPdfs.filter(p => /\.pdf$/i.test(p));

  const existingPaths = new Set(state.queue.map(i => i.path));

  const freshPdfs  = pdfPaths.filter(p => !existingPaths.has(p));
  const freshDocxs = docxDirect.filter(p => !existingPaths.has(p));

  if (freshPdfs.length === 0 && freshDocxs.length === 0) return;

  const isFirstDrop = state.queue.length === 0;

  // Build queue items
  const pdfItems = freshPdfs.map(p => ({
    id: nextId(), path: p, name: p.split('/').pop(),
    fileType: 'pdf', status: 'loading',
    pageCount: null, ar: null, widthPt: null, heightPt: null,
    fillMode: state.settings.fillMode || 'black',
    progress: 0, progressMsg: '', outputPath: null, errorMsg: null,
  }));

  const docxItems = freshDocxs.map(p => ({
    id: nextId(), path: p, name: p.split('/').pop(),
    fileType: 'docx', status: 'loading',
    slideCount: null, wordCount: null,
    fillMode: 'black',   // unused for docx but keeps shape consistent
    slideTheme: 'light', // 'light' | 'dark'
    progress: 0, progressMsg: '', outputPath: null, errorMsg: null,
  }));

  const newItems = [...pdfItems, ...docxItems];
  state.queue.push(...newItems);

  if (!state.selectedId && newItems.length > 0) {
    state.selectedId = newItems[0].id;
  }

  renderQueue();
  _dropLabelIdx++;
  updateDzLabel();
  updateConvertButton();
  updateSidebar();

  // Load metadata in parallel
  await Promise.all([
    ...pdfItems.map(async (item) => {
      const info = await window.slidefluid.getPdfInfo(item.path);
      if (info.ok) {
        item.pageCount = info.pageCount;
        item.ar        = info.ar;
        item.widthPt   = info.widthPt;
        item.heightPt  = info.heightPt;
      }
      item.status = 'waiting';
      refreshQueueItem(item);
      if (state.selectedId === item.id) updateSidebar();
    }),
    ...docxItems.map(async (item) => {
      const info = await window.slidefluid.getDocxInfo(item.path);
      if (info.ok) {
        item.slideCount = info.slideCount;
        item.wordCount  = info.wordCount;
      }
      item.status = 'waiting';
      refreshQueueItem(item);
      if (state.selectedId === item.id) updateSidebar();
    }),
  ]);

  updateConvertButton();

  if (isFirstDrop && !state.outputDir) {
    const chosen = await window.slidefluid.openFolderPicker();
    if (chosen) {
      state.outputDir = chosen;
      renderOutputFolder();
      updateConvertButton();
    }
  }
}

// .doc notice — old binary Word format is not supported; prompt to re-save
function showDocNotice() {
  const existing = document.getElementById('doc-notice');
  if (existing) { clearTimeout(existing._timer); existing.remove(); }

  const notice = document.createElement('div');
  notice.id = 'doc-notice';
  notice.className = 'docx-notice';
  notice.textContent = '.doc format not supported \u2014 please re-save as .docx or .txt first';

  document.getElementById('drop-zone').after(notice);
  notice._timer = setTimeout(() => notice.parentNode && notice.remove(), 6000);
}

// Guidance notice shown once per session when the first DOCX/TXT file is dropped
let _docxGuidanceShown = false;
function showDocxGuidance() {
  if (_docxGuidanceShown) return;
  _docxGuidanceShown = true;

  const existing = document.getElementById('docx-guidance');
  if (existing) return;

  const notice = document.createElement('div');
  notice.id = 'docx-guidance';
  notice.className = 'docx-notice docx-guidance';
  notice.innerHTML =
    '<strong>Slide boundaries:</strong> separate slides with two blank lines in your document. ' +
    'A single blank line is treated as a paragraph break within the same slide.';

  document.getElementById('drop-zone').after(notice);
  notice._timer = setTimeout(() => notice.parentNode && notice.remove(), 10000);
}

// 6.1 — Brief red flash on drop zone border for unsupported file types
function flashDropZoneReject() {
  const dz = document.getElementById('drop-zone');
  dz.classList.remove('reject-flash');
  void dz.offsetWidth; // force reflow to restart animation
  dz.classList.add('reject-flash');
  setTimeout(() => dz.classList.remove('reject-flash'), 700);
}

// ---------------------------------------------------------------------------
// Queue rendering
// ---------------------------------------------------------------------------

function renderQueue() {
  const list = document.getElementById('queue-list');
  list.innerHTML = '';
  for (const item of state.queue) {
    list.appendChild(buildQueueItemEl(item));
  }

  // Shrink drop zone once queue has items
  document.getElementById('drop-zone')
    .classList.toggle('compact', state.queue.length > 0);
}

function refreshQueueItem(item) {
  const existing = document.querySelector(`.queue-item[data-id="${item.id}"]`);
  if (!existing) return;
  existing.replaceWith(buildQueueItemEl(item));
}

function buildQueueItemEl(item) {
  const el = document.createElement('div');
  el.className = 'queue-item' + (item.id === state.selectedId ? ' selected' : '');
  el.dataset.id = item.id;
  el.setAttribute('role', 'listitem');

  const statusLabel = {
    loading:    'Loading',
    waiting:    'Waiting',
    converting: 'Converting',
    done:       'Done',
    error:      'Error',
  }[item.status] || item.status;

  const metaParts = [];
  if (item.fileType === 'docx') {
    if (item.slideCount != null) metaParts.push(`${item.slideCount} slide${item.slideCount !== 1 ? 's' : ''}`);
    if (item.wordCount  != null) metaParts.push(`${item.wordCount} words`);
  } else {
    if (item.pageCount != null) metaParts.push(`${item.pageCount} page${item.pageCount !== 1 ? 's' : ''}`);
    if (item.ar        != null) metaParts.push(arLabel(item.ar));
  }
  const meta = metaParts.join(' · ') || (item.status === 'loading' ? 'Reading\u2026' : '—');

  const displayMeta = (item.status === 'converting' && item.progressMsg)
    ? item.progressMsg
    : meta;

  const progressBar = item.status === 'converting'
    ? `<div class="qi-progress-wrap"><div class="qi-progress-fill" style="width:${Math.round((item.progress || 0) * 100)}%"></div></div>`
    : '';

  el.innerHTML = `
    <div class="qi-icon">${item.fileType === 'docx' ? (item.name.split('.').pop().toUpperCase() || 'DOC') : 'PDF'}</div>
    <div class="qi-body">
      <div class="qi-name" title="${escHtml(item.path)}">${escHtml(item.name)}</div>
      <div class="qi-meta">${escHtml(displayMeta)}</div>
      ${progressBar}
    </div>
    <div class="qi-badge status-${item.status}">${escHtml(statusLabel)}</div>
  `;

  el.addEventListener('click', () => selectItem(item.id));
  return el;
}

function selectItem(id) {
  state.selectedId = id;
  document.querySelectorAll('.queue-item').forEach(el => {
    el.classList.toggle('selected', el.dataset.id === id);
  });
  updateSidebar();
}

// ---------------------------------------------------------------------------
// Sidebar
// ---------------------------------------------------------------------------

function updateSidebar() {
  const item = state.queue.find(i => i.id === state.selectedId) || null;
  drawPreview(item);
  renderSidebarControls(item);
  renderFileDetails(item);
}

// 6.12 — Dispatch on fileType so DOCX items can have different sidebar controls in Phase 8
function renderSidebarControls(item) {
  switch (item ? item.fileType : null) {
    case 'pdf':
      renderArBadge(item);
      renderFillModeSelector(item);
      renderDocxInfo(null);
      break;
    case 'docx':
      renderArBadge(null);
      renderDocxInfo(item);
      renderSlideThemeSelector(item);
      break;
    default:
      renderArBadge(null);
      renderFillModeSelector(null);
      break;
  }
}

function renderDocxInfo(item) {
  // Reuse the ar-badge slot to show slide/word count for DOCX items
  const el = document.getElementById('ar-badge');
  if (!item) { el.innerHTML = ''; return; }

  if (item.status === 'loading' || item.slideCount == null) {
    el.innerHTML = '<div class="ar-badge ar-neutral">Analysing\u2026</div>';
    return;
  }

  const ext   = item.name.split('.').pop().toUpperCase();
  const slide = item.slideCount === 1 ? '1 slide' : `${item.slideCount} slides`;
  const words = item.wordCount  != null ? ` \u00b7 ${item.wordCount} words` : '';
  el.innerHTML = `<div class="ar-badge ar-ok">${ext} \u2014 ${slide}${words}</div>`;
}

function renderSlideThemeSelector(item) {
  const el = document.getElementById('fill-mode-selector');
  if (!item) { el.innerHTML = ''; return; }

  const themes = [
    { value: 'light', label: 'Light' },
    { value: 'dark',  label: 'Dark'  },
  ];
  const current = item.slideTheme || 'light';

  el.innerHTML = `
    <div class="fill-mode-label">Slide theme</div>
    <div class="fill-mode-buttons">
      ${themes.map(t => `
        <button class="fill-btn${current === t.value ? ' active' : ''}"
                data-theme="${t.value}">${escHtml(t.label)}</button>
      `).join('')}
    </div>
  `;

  el.querySelectorAll('.fill-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      item.slideTheme = btn.dataset.theme;
      el.querySelectorAll('.fill-btn').forEach(b =>
        b.classList.toggle('active', b.dataset.theme === item.slideTheme)
      );
      drawPreview(item);
      if (item.status === 'done' || item.status === 'error') {
        item.status = 'waiting';
        refreshQueueItem(item);
        updateConvertButton();
      }
    });
  });
}

// --- Preview canvas ---

// 6.13 — Dispatch on fileType and previewMode; only 'pdf'+'graphical' does real work now
function drawPreview(item) {
  const canvas = document.getElementById('preview-canvas');
  const ctx = canvas.getContext('2d');
  const W = canvas.width;   // 320
  const H = canvas.height;  // 180

  ctx.clearRect(0, 0, W, H);

  const fileType    = item ? (item.fileType || 'pdf') : null;
  const previewMode = state.settings.previewMode || 'graphical';

  if (fileType === 'docx') {
    _drawPreviewDocx(ctx, W, H, item);
    _resetPreviewButton(item);
    return;
  }

  if (!item || item.ar == null) {
    _drawPreviewPlaceholder(ctx, W, H, item);
    return;
  }

  // PDF path
  _drawPreviewPdfGraphical(ctx, W, H, item);
  _resetPreviewButton(item);
}

// Show / hide + reset the "See first slide" button whenever drawPreview runs
function _resetPreviewButton(item) {
  const btn = document.getElementById('btn-live-preview');
  if (!btn) return;
  const show = item && item.fileType === 'pdf' &&
               item.status !== 'loading' && item.status !== 'error' &&
               item.pageCount != null;
  btn.style.display = show ? '' : 'none';
  btn.textContent = 'See first slide';
  btn.disabled = false;
  delete btn.dataset.mode;
}

// Render page 1 of the selected PDF onto the canvas via pdftoppm
async function loadLivePreview(item) {
  const btn = document.getElementById('btn-live-preview');
  if (btn) { btn.disabled = true; btn.textContent = 'Loading…'; }

  const result = await window.slidefluid.getPdfPageImage(item.path, 1);

  if (!result.ok) {
    if (btn) { btn.disabled = false; btn.textContent = 'See first slide'; }
    console.warn('[preview]', result.error);
    return;
  }

  const canvas = document.getElementById('preview-canvas');
  const ctx = canvas.getContext('2d');
  const W = canvas.width, H = canvas.height;

  const img = new Image();
  img.onload = () => {
    // Compute content rect (same geometry as graphical preview)
    const TARGET_AR = 16 / 9;
    const is169 = item.ar != null && Math.abs(item.ar - TARGET_AR) < 0.05;
    const fillMode = item.fillMode || 'black';
    let cx = 0, cy = 0, cw = W, ch = H;

    if (!is169 && item.ar != null) {
      if (item.ar < TARGET_AR) {        // pillarbox
        cw = Math.min(W, Math.round(H * item.ar));
        cx = Math.round((W - cw) / 2);
      } else {                          // letterbox
        ch = Math.min(H, Math.round(W / item.ar));
        cy = Math.round((H - ch) / 2);
      }
    }

    ctx.clearRect(0, 0, W, H);

    if (is169 || (cx === 0 && cy === 0)) {
      // No bars — fill frame exactly
      ctx.drawImage(img, 0, 0, W, H);
    } else if (fillMode === 'black') {
      ctx.fillStyle = '#000';
      ctx.fillRect(0, 0, W, H);
      ctx.drawImage(img, cx, cy, cw, ch);
    } else if (fillMode === 'color_match') {
      // Draw slide into content area, sample its inner edge pixels, fill bars
      ctx.drawImage(img, cx, cy, cw, ch);
      const edgePx = cx > 0
        ? ctx.getImageData(cx, cy, 1, ch).data       // pillarbox: left edge column
        : ctx.getImageData(cx, cy, cw, 1).data;      // letterbox: top edge row
      const [r, g, b] = _avgColor(edgePx);
      ctx.fillStyle = `rgb(${r},${g},${b})`;
      if (cx > 0) {
        ctx.fillRect(0, 0, cx, H);
        ctx.fillRect(cx + cw, 0, W - cx - cw, H);
      } else {
        ctx.fillRect(0, 0, W, cy);
        ctx.fillRect(0, cy + ch, W, H - cy - ch);
      }
      // Redraw slide on top so it sits above the bar fill
      ctx.drawImage(img, cx, cy, cw, ch);
    } else {
      // smear — stretch the actual edge pixels of the image into the bar areas
      ctx.fillStyle = '#000';
      ctx.fillRect(0, 0, W, H);
      ctx.drawImage(img, cx, cy, cw, ch);
      if (cx > 0) {
        // Stretch leftmost column of slide into left bar
        ctx.drawImage(img, 0, 0, 1, img.height, 0, cy, cx, ch);
        // Stretch rightmost column into right bar
        ctx.drawImage(img, img.width - 1, 0, 1, img.height, cx + cw, cy, W - cx - cw, ch);
      } else {
        // Stretch top row into top bar
        ctx.drawImage(img, 0, 0, img.width, 1, cx, 0, cw, cy);
        // Stretch bottom row into bottom bar
        ctx.drawImage(img, 0, img.height - 1, img.width, 1, cx, cy + ch, cw, H - cy - ch);
      }
    }

    if (btn) { btn.textContent = '← Graphical'; btn.disabled = false; btn.dataset.mode = 'live'; }
  };
  img.onerror = () => {
    if (btn) { btn.disabled = false; btn.textContent = 'See first slide'; }
  };
  img.src = result.dataUrl;
}

function _drawPreviewDocx(ctx, W, H, item) {
  const dark = item && item.slideTheme === 'dark';
  const pad  = 10;
  const slideW = W - pad * 2;
  const slideH = H - pad * 2;

  // Canvas background
  ctx.fillStyle = '#0E1420';
  ctx.fillRect(0, 0, W, H);

  // Slide face
  ctx.fillStyle = dark ? '#111111' : '#FFFFFF';
  ctx.fillRect(pad, pad, slideW, slideH);

  // Subtle border on dark slide so it reads against the canvas bg
  if (dark) {
    ctx.strokeStyle = '#2A3040';
    ctx.lineWidth = 1;
    ctx.strokeRect(pad + 0.5, pad + 0.5, slideW - 1, slideH - 1);
  }

  if (!item || item.status === 'loading') {
    // Draw three animated-looking dots as placeholder bars
    const dotW = 6, dotH = 6, dotGap = 8;
    const totalW = dotW * 3 + dotGap * 2;
    let dx = W / 2 - totalW / 2;
    const dy = H / 2 - dotH / 2;
    ctx.fillStyle = dark ? '#444444' : '#CCCCCC';
    for (let i = 0; i < 3; i++) {
      ctx.beginPath();
      ctx.arc(dx + dotW / 2, dy + dotH / 2, dotW / 2, 0, Math.PI * 2);
      ctx.fill();
      dx += dotW + dotGap;
    }
    return;
  }

  // Content lines — simplified text preview
  const lineH = 10;
  const lineGap = 5;
  const lineX = pad + 18;
  const lineMaxW = slideW - 36;
  let lineY = pad + 20;

  ctx.fillStyle = dark ? '#DDDDDD' : '#222222';
  ctx.textAlign = 'left';
  ctx.textBaseline = 'top';

  if (item.slideCount != null) {
    const ext = (item.name.split('.').pop() || '').toUpperCase();
    ctx.font = `bold 9px system-ui`;
    ctx.fillText(`${ext} \u2014 ${item.slideCount} slide${item.slideCount !== 1 ? 's' : ''}`, lineX, lineY);
    lineY += lineH + lineGap + 4;
  }

  // Draw representative text lines
  ctx.fillStyle = dark ? '#888888' : '#444444';
  ctx.font = '8px system-ui';
  const lines = [lineMaxW * 0.9, lineMaxW * 0.75, lineMaxW * 0.85, lineMaxW * 0.6];
  for (const lw of lines) {
    if (lineY + lineH > pad + slideH - 12) break;
    ctx.fillRect(lineX, lineY, lw, 5);
    lineY += lineH;
  }
}

function _drawPreviewPlaceholder(ctx, W, H, item) {
  ctx.fillStyle = '#0E1420';
  ctx.fillRect(0, 0, W, H);
  ctx.strokeStyle = '#1E2840';
  ctx.lineWidth = 1;
  ctx.strokeRect(0.5, 0.5, W - 1, H - 1);
  ctx.fillStyle = '#4A5878';
  ctx.font = '12px system-ui';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(item ? 'Loading…' : 'No file selected', W / 2, H / 2);
}

function _drawPreviewPdfGraphical(ctx, W, H, item) {
  const TARGET_AR = 16 / 9;
  const is169 = Math.abs(item.ar - TARGET_AR) < 0.05;
  const fillMode = item.fillMode || 'black';

  let cx = 0, cy = 0, cw = W, ch = H;

  if (!is169) {
    if (item.ar < TARGET_AR) {
      // Narrower than 16:9 → pillarbox (bars on left/right)
      cw = Math.min(W, Math.round(H * item.ar));
      cx = Math.round((W - cw) / 2);
    } else {
      // Wider than 16:9 → letterbox (bars top/bottom)
      ch = Math.min(H, Math.round(W / item.ar));
      cy = Math.round((H - ch) / 2);
    }
    drawBars(ctx, W, H, cx, cy, cw, ch, fillMode);
  }

  // Content area gradient
  const grad = ctx.createLinearGradient(cx, cy, cx + cw, cy + ch);
  grad.addColorStop(0, '#1B2C4A');
  grad.addColorStop(1, '#0C1520');
  ctx.fillStyle = grad;
  ctx.fillRect(cx, cy, cw, ch);

  // Decorative lines suggesting slide content
  ctx.strokeStyle = 'rgba(61,255,204,0.16)';
  ctx.lineWidth = 1.5;
  const lines = [0.30, 0.48, 0.63];
  lines.forEach((frac, i) => {
    const y = cy + Math.round(ch * frac);
    const xPad = 18;
    const lineW = (cw - xPad * 2) * (1 - i * 0.18);
    ctx.beginPath();
    ctx.moveTo(cx + xPad, y);
    ctx.lineTo(cx + xPad + lineW, y);
    ctx.stroke();
  });

  // Hairline border around content area (only when bars present)
  if (!is169) {
    ctx.strokeStyle = 'rgba(30,40,64,0.9)';
    ctx.lineWidth = 1;
    ctx.strokeRect(cx + 0.5, cy + 0.5, cw - 1, ch - 1);
  }
}

function drawBars(ctx, W, H, cx, cy, cw, ch, fillMode) {
  if (fillMode === 'black') {
    ctx.fillStyle = '#000000';
    ctx.fillRect(0, 0, W, H);
    return;
  }

  if (fillMode === 'color_match') {
    // Solid muted blue approximating edge-color sample
    ctx.fillStyle = '#18243A';
    ctx.fillRect(0, 0, W, H);
    return;
  }

  // smear — gradient fading from content edges
  ctx.fillStyle = '#0C1520';
  ctx.fillRect(0, 0, W, H);

  if (cx > 0) {
    // Left bar
    const gl = ctx.createLinearGradient(cx, 0, 0, 0);
    gl.addColorStop(0, '#1B2C4A');
    gl.addColorStop(1, '#0C1520');
    ctx.fillStyle = gl;
    ctx.fillRect(0, 0, cx, H);
    // Right bar
    const gr = ctx.createLinearGradient(cx + cw, 0, W, 0);
    gr.addColorStop(0, '#1B2C4A');
    gr.addColorStop(1, '#0C1520');
    ctx.fillStyle = gr;
    ctx.fillRect(cx + cw, 0, W - cx - cw, H);
  } else {
    // Top bar
    const gt = ctx.createLinearGradient(0, cy, 0, 0);
    gt.addColorStop(0, '#1B2C4A');
    gt.addColorStop(1, '#0C1520');
    ctx.fillStyle = gt;
    ctx.fillRect(0, 0, W, cy);
    // Bottom bar
    const gb = ctx.createLinearGradient(0, cy + ch, 0, H);
    gb.addColorStop(0, '#1B2C4A');
    gb.addColorStop(1, '#0C1520');
    ctx.fillStyle = gb;
    ctx.fillRect(0, cy + ch, W, H - cy - ch);
  }
}

// --- AR badge ---

function renderArBadge(item) {
  const el = document.getElementById('ar-badge');
  if (!item || item.ar == null) {
    el.innerHTML = '';
    return;
  }
  const is169 = Math.abs(item.ar - 16 / 9) < 0.05;
  const label  = arLabel(item.ar);
  const status = is169
    ? `<span class="ar-status ok">correct</span>`
    : `<span class="ar-status warn">${item.ar < 16 / 9 ? 'pillarbox' : 'letterbox'}</span>`;
  el.innerHTML = `<span class="ar-label">${escHtml(label)}</span>${status}`;
}

// --- Fill mode selector ---

function renderFillModeSelector(item) {
  const el = document.getElementById('fill-mode-selector');
  // Only show for non-16:9 items with known AR
  if (!item || item.ar == null || Math.abs(item.ar - 16 / 9) < 0.05) {
    el.innerHTML = '';
    return;
  }

  const modes = [
    { value: 'black',       label: 'Black' },
    { value: 'color_match', label: 'Color Match' },
    { value: 'smear',       label: 'Smear Fill' },
  ];

  el.innerHTML = `
    <div class="fill-mode-label">Fill mode</div>
    <div class="fill-mode-buttons">
      ${modes.map(m => `
        <button class="fill-btn${item.fillMode === m.value ? ' active' : ''}"
                data-mode="${m.value}">${escHtml(m.label)}</button>
      `).join('')}
    </div>
  `;

  el.querySelectorAll('.fill-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      item.fillMode = btn.dataset.mode;
      el.querySelectorAll('.fill-btn').forEach(b =>
        b.classList.toggle('active', b.dataset.mode === item.fillMode)
      );
      drawPreview(item);
      // Changing fill mode on a finished item re-queues it for conversion
      if (item.status === 'done' || item.status === 'error') {
        item.status = 'waiting';
        item.progress = 0;
        item.outputPath = null;
        item.errorMsg = null;
        refreshQueueItem(item);
        updateConvertButton();
      }
    });
  });
}

// --- File details ---

function renderFileDetails(item) {
  const el = document.getElementById('file-details');
  if (!item) { el.innerHTML = ''; return; }

  const rows = [];

  rows.push(detailRow('File', item.name, item.path));

  if (item.fileType === 'docx') {
    if (item.slideCount != null) rows.push(detailRow('Slides', String(item.slideCount)));
    if (item.wordCount  != null) rows.push(detailRow('Words',  String(item.wordCount)));
  } else {
    if (item.pageCount != null) {
      rows.push(detailRow('Pages', String(item.pageCount)));
    }
    if (item.ar != null) {
      rows.push(detailRow('Aspect', arLabel(item.ar)));
    }
    if (item.widthPt != null && item.heightPt != null) {
      rows.push(detailRow('Size', `${Math.round(item.widthPt)} × ${Math.round(item.heightPt)} pt`));
    }
    if (item.pageCount != null) {
      const dpi = state.settings.dpi || 144;
      const estMB = (item.pageCount * (dpi === 72 ? 0.35 : 1.4)).toFixed(1);
      rows.push(detailRow('Est. size', `~${estMB} MB at ${dpi} DPI`));
    }
  }

  el.innerHTML = rows.join('');
}

function detailRow(label, value, title) {
  const titleAttr = title ? ` title="${escHtml(title)}"` : '';
  return `<div class="detail-row">
    <span class="detail-label">${escHtml(label)}</span>
    <span class="detail-value"${titleAttr}>${escHtml(value)}</span>
  </div>`;
}

// ---------------------------------------------------------------------------
// Output folder row
// ---------------------------------------------------------------------------

function renderOutputFolder() {
  const el = document.getElementById('output-folder-row');

  if (!state.outputDir) {
    el.innerHTML = `
      <div class="of-row">
        <span class="of-label">Output folder</span>
        <button class="btn-folder btn-folder-choose" id="btn-choose-folder">Choose Folder</button>
      </div>`;
  } else {
    const parts = state.outputDir.split('/');
    const name  = parts[parts.length - 1] || state.outputDir;
    el.innerHTML = `
      <div class="of-row">
        <div class="of-info">
          <span class="of-label">Output folder</span>
          <span class="of-path" title="${escHtml(state.outputDir)}">${escHtml(name)}</span>
        </div>
        <button class="btn-folder btn-folder-change" id="btn-choose-folder">Change</button>
      </div>`;
  }

  document.getElementById('btn-choose-folder')
    .addEventListener('click', chooseOutputFolder);
}

async function chooseOutputFolder() {
  const chosen = await window.slidefluid.openFolderPicker(state.outputDir || undefined);
  if (chosen) {
    state.outputDir = chosen;
    renderOutputFolder();
    updateConvertButton();
  }
}

// ---------------------------------------------------------------------------
// Convert button
// ---------------------------------------------------------------------------

function setupConvertButton() {
  document.getElementById('btn-convert').addEventListener('click', handleConvertClick);
}

function updateConvertButton() {
  const btn = document.getElementById('btn-convert');
  const waitingItems = state.queue.filter(i => i.status === 'waiting');
  const allSettled   = state.queue.length > 0 &&
                       state.queue.every(i => i.status === 'done' || i.status === 'error');

  if (state.isConverting) {
    setState(btn, 'converting', 'Cancel', false);
  } else if (allSettled) {
    setState(btn, 'complete', skinText('Open Output Folder', "They're free. Open folder?"), false);
  } else if (waitingItems.length > 0 && state.outputDir) {
    setState(btn, 'ready', skinText('Convert to PPTX', 'Set them free \u2192'), false);
  } else {
    setState(btn, 'idle', skinText('Convert to PPTX', 'Set them free \u2192'), true);
  }
}

function setState(btn, cls, label, disabled) {
  btn.className = `btn-convert state-${cls}`;
  btn.textContent = label;
  btn.disabled = disabled;
}

function setupKeyboardShortcuts() {
  document.addEventListener('keydown', (e) => {
    // ⌘↵ — trigger Convert when in ready state
    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
      const btn = document.getElementById('btn-convert');
      if (btn && btn.classList.contains('state-ready')) btn.click();
    }
  });
}

async function handleConvertClick() {
  if (state.isConverting) {
    await cancelConversion();
    return;
  }

  const btn = document.getElementById('btn-convert');
  if (btn.classList.contains('state-complete')) {
    if (state.outputDir) await window.slidefluid.openFolder(state.outputDir);
    // Reset done items to waiting so the user can convert again without restarting
    state.queue.forEach(item => {
      if (item.status === 'done') {
        item.status = 'waiting';
        item.progress = 0;
        item.outputPath = null;
        refreshQueueItem(item);
      }
    });
    updateConvertButton();
    return;
  }

  await beginConversion();
}

async function cancelConversion() {
  await window.slidefluid.cancelConversion();
  // State cleanup happens in handleConversionExit
}

// ---------------------------------------------------------------------------
// Conversion flow
// ---------------------------------------------------------------------------

async function beginConversion() {
  const waitingItems = state.queue.filter(i => i.status === 'waiting');
  if (!waitingItems.length || !state.outputDir) return;

  const suffix = state.settings.filenameSuffix || '';

  // Overwrite check — one per file; track 'overwrite_all'/'skip_all' to avoid
  // repeated dialogs when the user has already given a blanket answer.
  const toConvert = [];
  let bulkAction = null; // 'overwrite' | 'skip' once user picks an "All" option

  for (const item of waitingItems) {
    const stem = item.name.replace(/\.(pdf|docx|txt)$/i, '');
    let outPath = `${state.outputDir}/${stem}${suffix}.pptx`;

    let decision;
    if (bulkAction) {
      decision = bulkAction;
    } else {
      decision = await window.slidefluid.checkOverwrite(outPath);
    }

    if (decision === 'cancel') return;
    if (decision === 'skip') continue;
    if (decision === 'skip_all') { bulkAction = 'skip'; continue; }
    if (decision === 'overwrite_all') bulkAction = 'overwrite';

    if (decision === 'rename') {
      let n = 1;
      let candidate;
      do {
        candidate = `${state.outputDir}/${stem}${suffix}_${n}.pptx`;
        n++;
      } while (n <= 99 && (await window.slidefluid.checkOverwrite(candidate)) !== 'ok');
      outPath = candidate;
    }

    toConvert.push(item);
  }

  if (!toConvert.length) return;

  state.isConverting = true;
  updateConvertButton();

  _convTotalFiles = toConvert.length;

  document.getElementById('progress-section').classList.add('visible');
  document.getElementById('progress-bar').style.width = '0%';
  document.getElementById('progress-status-text').textContent = '';
  _updateOverallProgress(_convTotalFiles, 0);

  // Initialize report — entries pre-populated so warn/done/error handlers can find them
  _convReport = {
    startTime: new Date(),
    dpi: state.settings.dpi || 144,
    suffix: suffix || '',
    entries: toConvert.map(item => ({
      path:       item.path,
      name:       item.name,
      fileType:   item.fileType,
      fillMode:   item.fileType === 'pdf'  ? (item.fillMode  || 'black') : null,
      slideTheme: item.fileType === 'docx' ? (item.slideTheme || 'light') : null,
      status:     'pending',
      slides:     0,
      outputPath: null,
      errorMsg:   null,
      warnings:   [],
    })),
  };

  // Group PDF items by fillMode; DOCX/TXT items grouped by slideTheme at the end.
  const groups = [];
  const seenFm = new Map();
  const seenTheme = new Map();

  for (const item of toConvert) {
    if (item.fileType === 'docx') {
      const theme = item.slideTheme || 'light';
      if (!seenTheme.has(theme)) { seenTheme.set(theme, []); }
      seenTheme.get(theme).push(item);
    } else {
      const fm = item.fillMode || 'black';
      if (!seenFm.has(fm)) { seenFm.set(fm, []); groups.push({ fillMode: fm, items: seenFm.get(fm) }); }
      seenFm.get(fm).push(item);
    }
  }
  for (const [theme, items] of seenTheme) {
    groups.push({ fillMode: 'black', slideTheme: theme, items }); // fillMode ignored by Python for DOCX
  }

  for (const group of groups) {
    if (!state.isConverting) break; // cancelled by user mid-batch

    _unsubConversionEvents();
    const batchDone = new Promise(resolve => { _batchResolve = resolve; });
    _convUnsubs = [
      window.slidefluid.onConversionMessage(handleConversionMessage),
      window.slidefluid.onConversionStderr(handleConversionStderr),
      window.slidefluid.onConversionExit(handleConversionExit),
      window.slidefluid.onConversionSpawnError(handleConversionSpawnError),
    ];

    const result = await window.slidefluid.startConversion({
      files: group.items.map(i => i.path),
      outputDir: state.outputDir,
      dpi: state.settings.dpi || 144,
      fillMode: group.fillMode,
      slideTheme: group.slideTheme || 'light',
      suffix,
    });

    if (!result.ok) {
      state.isConverting = false;
      _unsubConversionEvents();
      if (_batchResolve) { _batchResolve = null; }
      document.getElementById('progress-status-text').textContent =
        `Error: ${result.error || 'Could not start conversion'}`;
      updateConvertButton();
      return;
    }

    await batchDone; // wait for batch_done, exit, or spawn error before next group
  }

  // All groups finished (or conversion was cancelled inside a group)
  if (state.isConverting) {
    state.isConverting = false;
    updateConvertButton();

    if (state.settings.writeReport !== false && _convReport) {
      await _writeConversionReport();
    }

    if (state.settings.autoOpenOnComplete && state.outputDir) {
      window.slidefluid.openFolder(state.outputDir);
    }
  }
}

function _unsubConversionEvents() {
  _convUnsubs.forEach(fn => fn && fn());
  _convUnsubs = [];
}

function _resolveBatch() {
  if (_batchResolve) { const r = _batchResolve; _batchResolve = null; r(); }
}

async function _writeConversionReport() {
  if (!_convReport || !state.outputDir) return;

  const now      = _convReport.startTime;
  const dateStr  = now.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  const timeStr  = now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
  const divider  = '─'.repeat(52);
  const heavy    = '═'.repeat(52);

  const converted   = _convReport.entries.filter(e => e.status === 'done').length;
  const errors      = _convReport.entries.filter(e => e.status === 'error').length;
  const totalSlides = _convReport.entries.reduce((s, e) => s + (e.slides || 0), 0);

  const fillLabel = { black: 'Black', color_match: 'Color Match', smear: 'Smear Fill' };

  const lines = [
    'SlideFluid 3.0 — Conversion Report',
    `Generated: ${dateStr} at ${timeStr}`,
    heavy,
    '',
    'Settings',
    `  Quality:  ${_convReport.dpi} DPI`,
    `  Suffix:   ${_convReport.suffix || '(none)'}`,
    '',
    `Results: ${converted} file${converted !== 1 ? 's' : ''} converted · ${totalSlides} slide${totalSlides !== 1 ? 's' : ''} created` +
      (errors ? ` · ${errors} error${errors !== 1 ? 's' : ''}` : ''),
    '',
    divider,
  ];

  _convReport.entries.forEach((e, idx) => {
    const tags = [];
    if (e.fillMode && e.fillMode !== 'black') tags.push(fillLabel[e.fillMode] || e.fillMode);
    if (e.slideTheme === 'dark') tags.push('dark');
    const tagStr = tags.length ? `  [${tags.join(', ')}]` : '';

    lines.push(`[${idx + 1}] ${e.name}${tagStr}`);

    if (e.status === 'done') {
      const outName = e.outputPath ? e.outputPath.split('/').pop() : '—';
      lines.push(`    Slides: ${e.slides}  ·  Output: ${outName}`);
    } else if (e.status === 'error') {
      lines.push(`    ERROR: ${e.errorMsg || 'Unknown error'}`);
    }

    for (const w of e.warnings) lines.push(`    ⚠ ${w}`);
    if (idx < _convReport.entries.length - 1) lines.push('');
  });

  lines.push(divider);

  const pad   = (n) => String(n).padStart(2, '0');
  const ts    = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;
  const rPath = `${state.outputDir}/SlideFluid_Report_${ts}.txt`;

  await window.slidefluid.writeFile(rPath, lines.join('\n') + '\n');
}

function _startFunMsgRotation(filePath) {
  _clearFunMsgRotation();
  // Start at a random position so each file feels fresh
  _progressMsgIdx = Math.floor(Math.random() * FUN_PROGRESS_MSGS.length);
  _currentFunMsg  = FUN_PROGRESS_MSGS[_progressMsgIdx];

  _funMsgInterval = setInterval(() => {
    _progressMsgIdx = (_progressMsgIdx + 1) % FUN_PROGRESS_MSGS.length;
    _currentFunMsg  = FUN_PROGRESS_MSGS[_progressMsgIdx];
    const name = filePath ? filePath.split('/').pop() : '';
    const statusEl = document.getElementById('progress-status-text');
    if (statusEl) statusEl.textContent = `${name} — ${_currentFunMsg}`;
    // Also refresh the queue item's meta line
    const converting = state.queue.find(i => i.status === 'converting');
    if (converting) { converting.progressMsg = _currentFunMsg; refreshQueueItem(converting); }
  }, 2500);
}

function _clearFunMsgRotation() {
  if (_funMsgInterval) { clearInterval(_funMsgInterval); _funMsgInterval = null; }
}

function _updateOverallProgress(total, done) {
  document.getElementById('progress-overall-text').textContent =
    `${done} of ${total} files complete`;
}

function handleConversionMessage(msg) {
  switch (msg.type) {

    case 'start': {
      const item = state.queue.find(i => i.path === msg.file);
      if (item) {
        item.status = 'converting';
        item.progress = 0;
        item.progressMsg = '';
        refreshQueueItem(item);
        if (state.selectedId === item.id) updateSidebar();
      }
      if (state.settings.skin === 'fun') _startFunMsgRotation(msg.file);
      const doneCount = state.queue.filter(i => i.status === 'done').length;
      _updateOverallProgress(_convTotalFiles, doneCount);
      break;
    }

    case 'progress': {
      const item = state.queue.find(i => i.path === msg.file);
      if (item) {
        item.progress = msg.total_pages > 0 ? msg.page / msg.total_pages : 0;
        item.progressMsg = state.settings.skin === 'fun'
          ? _currentFunMsg
          : (msg.message || `Page ${msg.page} of ${msg.total_pages}`);
        refreshQueueItem(item);
      }
      // Overall bar: (all done + current file fraction) / all non-waiting
      const activeFiles = state.queue.filter(
        i => ['converting', 'done', 'error'].includes(i.status)
      ).length;
      const doneFiles = state.queue.filter(i => i.status === 'done').length;
      const fileFrac = msg.total_pages > 0 ? msg.page / msg.total_pages : 0;
      const overallPct = activeFiles > 0
        ? Math.round((doneFiles + fileFrac) / activeFiles * 100)
        : 0;
      document.getElementById('progress-bar').style.width = `${overallPct}%`;
      const name = msg.file ? msg.file.split('/').pop() : '';
      const statusLine = state.settings.skin === 'fun'
        ? `${name} — ${_currentFunMsg}`
        : `${name} — page ${msg.page} of ${msg.total_pages}`;
      document.getElementById('progress-status-text').textContent = statusLine;
      break;
    }

    case 'done': {
      const item = state.queue.find(i => i.path === msg.file);
      if (item) {
        item.status = 'done';
        item.progress = 1;
        item.outputPath = msg.output;
        item.progressMsg = '';
        refreshQueueItem(item);
        if (state.selectedId === item.id) updateSidebar();
      }
      const rDone = _convReport?.entries.find(e => e.path === msg.file);
      if (rDone) { rDone.status = 'done'; rDone.slides = msg.slides || 0; rDone.outputPath = msg.output; }
      break;
    }

    case 'error': {
      const item = state.queue.find(i => i.path === msg.file);
      if (item) {
        item.status = 'error';
        item.errorMsg = msg.message;
        item.progressMsg = '';
        refreshQueueItem(item);
        if (state.selectedId === item.id) updateSidebar();
      }
      const rErr = _convReport?.entries.find(e => e.path === msg.file);
      if (rErr) { rErr.status = 'error'; rErr.errorMsg = msg.message; }
      break;
    }

    case 'batch_done': {
      _clearFunMsgRotation();
      window.slidefluid.notifyJobDone();
      _unsubConversionEvents();

      const doneCount = state.queue.filter(i => i.status === 'done').length;
      _updateOverallProgress(_convTotalFiles, doneCount);
      document.getElementById('progress-bar').style.width = '100%';

      let summary;
      if (state.settings.skin === 'fun') {
        summary = FUN_SUCCESS_MSGS[Math.floor(Math.random() * FUN_SUCCESS_MSGS.length)];
      } else {
        summary =
          `${msg.converted} file${msg.converted !== 1 ? 's' : ''} converted. ` +
          `${msg.total_slides} slide${msg.total_slides !== 1 ? 's' : ''} created.`;
      }
      document.getElementById('progress-status-text').textContent = summary;

      if (state.settings.skin === 'fun') triggerConfetti();
      _resolveBatch(); // lets beginConversion start the next group (if any)
      break;
    }

    case 'warn': {
      const item = state.queue.find(i => i.path === msg.file);
      if (item) {
        item.progressMsg = msg.message;
        refreshQueueItem(item);
      }
      const rWarn = _convReport?.entries.find(e => e.path === msg.file);
      if (rWarn) rWarn.warnings.push(msg.message);
      break;
    }
  }
}

function handleConversionStderr(data) {
  console.warn('[conversion stderr]', data.message);
}

function handleConversionExit(data) {
  if (!state.isConverting) return;
  _clearFunMsgRotation();
  _unsubConversionEvents();
  window.slidefluid.notifyJobDone();

  if (data.cancelled) {
    state.queue.forEach(item => {
      if (item.status === 'converting') {
        item.status = 'waiting';
        item.progress = 0;
        item.progressMsg = '';
        refreshQueueItem(item);
      }
    });
    document.getElementById('progress-status-text').textContent = 'Cancelled.';
    document.getElementById('progress-bar').style.width = '0%';
    document.getElementById('progress-overall-text').textContent = '';
    state.isConverting = false;
    updateConvertButton();
  }

  _resolveBatch();
}

function handleConversionSpawnError(data) {
  _clearFunMsgRotation();
  _unsubConversionEvents();
  window.slidefluid.notifyJobDone();

  state.queue.forEach(item => {
    if (item.status === 'converting') {
      item.status = 'error';
      item.errorMsg = data.message;
      item.progressMsg = '';
      refreshQueueItem(item);
    }
  });

  document.getElementById('progress-status-text').textContent =
    `Failed to start Python backend: ${data.message}`;

  _resolveBatch();
}

// ---------------------------------------------------------------------------
// Settings gear & modal
// ---------------------------------------------------------------------------

function setupGearButton() {
  document.getElementById('btn-settings').addEventListener('click', openSettingsModal);
}

function openSettingsModal() {
  syncSettingsFormFromState();
  document.getElementById('settings-overlay').classList.remove('hidden');
  document.querySelector('.settings-tab').focus();
}

function closeSettingsModal() {
  document.getElementById('settings-overlay').classList.add('hidden');
}

function syncSettingsFormFromState() {
  const s = state.settings;

  // Appearance — skin
  document.querySelectorAll('.skin-option').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.skin === (s.skin || 'professional'));
  });

  // Output — DPI
  document.querySelectorAll('input[name="dpi"]').forEach(radio => {
    radio.checked = String(s.dpi || 144) === radio.value;
  });
  updateDpiWarning(s.dpi || 144);

  // Output — default fill mode
  document.querySelectorAll('input[name="default-fill"]').forEach(radio => {
    radio.checked = (s.fillMode || 'black') === radio.value;
  });

  // Output — auto-open
  document.getElementById('setting-auto-open').checked = s.autoOpenOnComplete !== false;

  // Output — overwrite behavior
  document.getElementById('setting-overwrite').value = s.overwriteBehavior || 'ask';

  // Output — filename suffix
  document.getElementById('setting-suffix').value = s.filenameSuffix || '';

  // Output — write report
  document.getElementById('setting-write-report').checked = s.writeReport !== false;

  // Output — folder display
  renderSettingsOutputFolder();
}

function renderSettingsOutputFolder() {
  const el = document.getElementById('settings-output-dir');
  const dir = state.settings.outputDir;
  if (dir) {
    const name = dir.split('/').pop() || dir;
    el.innerHTML = `
      <span class="settings-dir-path" title="${escHtml(dir)}">${escHtml(name)}</span>
      <button class="settings-dir-clear" id="btn-settings-clear-dir">Clear</button>`;
    document.getElementById('btn-settings-clear-dir').addEventListener('click', async () => {
      state.settings.outputDir = null;
      state.outputDir = null;
      await window.slidefluid.setSetting('outputDir', null);
      renderOutputFolder();
      renderSettingsOutputFolder();
      updateConvertButton();
    });
  } else {
    el.innerHTML = `<span class="settings-dir-none">Not set</span>`;
  }
}

function updateDpiWarning(dpi) {
  document.getElementById('dpi-warning').classList.toggle('visible', Number(dpi) === 72);
}

function setupSettingsModal() {
  const overlay = document.getElementById('settings-overlay');

  // Close on backdrop click
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) closeSettingsModal();
  });

  // Close on Escape
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !overlay.classList.contains('hidden')) closeSettingsModal();
  });

  // Close buttons
  document.getElementById('btn-settings-close').addEventListener('click', closeSettingsModal);
  document.getElementById('btn-settings-done').addEventListener('click', closeSettingsModal);

  // Tab switching
  document.querySelectorAll('.settings-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.settings-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.settings-pane').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById(`tab-${tab.dataset.tab}`).classList.add('active');
      if (tab.dataset.tab === 'diagnostics') loadLogTail();
    });
  });

  // Appearance — skin buttons
  document.querySelectorAll('.skin-option').forEach(btn => {
    btn.addEventListener('click', async () => {
      const skin = btn.dataset.skin;
      document.querySelectorAll('.skin-option').forEach(b =>
        b.classList.toggle('active', b.dataset.skin === skin)
      );
      state.settings.skin = skin;
      document.documentElement.dataset.skin = skin;
      await window.slidefluid.setSetting('skin', skin);
      updateSkinText();
    });
  });

  // Output — DPI radios
  document.querySelectorAll('input[name="dpi"]').forEach(radio => {
    radio.addEventListener('change', async () => {
      if (!radio.checked) return;
      const dpi = Number(radio.value);
      state.settings.dpi = dpi;
      await window.slidefluid.setSetting('dpi', dpi);
      updateDpiWarning(dpi);
    });
  });

  // Output — default fill mode radios
  document.querySelectorAll('input[name="default-fill"]').forEach(radio => {
    radio.addEventListener('change', async () => {
      if (!radio.checked) return;
      state.settings.fillMode = radio.value;
      await window.slidefluid.setSetting('fillMode', radio.value);
    });
  });

  // Output — auto-open toggle
  document.getElementById('setting-auto-open').addEventListener('change', async (e) => {
    state.settings.autoOpenOnComplete = e.target.checked;
    await window.slidefluid.setSetting('autoOpenOnComplete', e.target.checked);
  });

  // Output — overwrite behavior
  document.getElementById('setting-overwrite').addEventListener('change', async (e) => {
    state.settings.overwriteBehavior = e.target.value;
    await window.slidefluid.setSetting('overwriteBehavior', e.target.value);
  });

  // Output — filename suffix (debounced save)
  let _suffixTimer = null;
  document.getElementById('setting-suffix').addEventListener('input', (e) => {
    const val = e.target.value;
    state.settings.filenameSuffix = val;
    clearTimeout(_suffixTimer);
    _suffixTimer = setTimeout(() => window.slidefluid.setSetting('filenameSuffix', val), 400);
  });

  // Output — write report toggle
  document.getElementById('setting-write-report').addEventListener('change', async (e) => {
    state.settings.writeReport = e.target.checked;
    await window.slidefluid.setSetting('writeReport', e.target.checked);
  });

  // Diagnostics — software updates
  const updateStatusEl  = document.getElementById('update-status-line');
  const btnDownload     = document.getElementById('btn-download-update');
  const btnInstall      = document.getElementById('btn-install-update');

  function setUpdateStatus(text, cls) {
    updateStatusEl.textContent = text;
    updateStatusEl.className   = `update-status-line${cls ? ' ' + cls : ''}`;
  }

  window.slidefluid.onUpdateStatus((payload) => {
    switch (payload.state) {
      case 'checking':
        setUpdateStatus('Checking for updates…');
        btnDownload.style.display = 'none';
        btnInstall.style.display  = 'none';
        break;
      case 'current':
        setUpdateStatus(`You're up to date — v${payload.version}`, 'ok');
        btnDownload.style.display = 'none';
        btnInstall.style.display  = 'none';
        break;
      case 'available':
        setUpdateStatus(`Update available: v${payload.version}`, 'warn');
        btnDownload.style.display = '';
        btnInstall.style.display  = 'none';
        break;
      case 'downloading':
        setUpdateStatus(`Downloading… ${payload.percent}%`);
        btnDownload.style.display = 'none';
        btnInstall.style.display  = 'none';
        break;
      case 'ready':
        setUpdateStatus(`v${payload.version} ready to install`, 'ok');
        btnDownload.style.display = 'none';
        btnInstall.style.display  = '';
        break;
      case 'error':
        setUpdateStatus(`Update error: ${payload.message}`, 'error');
        break;
    }
  });

  document.getElementById('btn-check-update').addEventListener('click', () => {
    setUpdateStatus('Checking…');
    window.slidefluid.checkForUpdates();
  });

  btnDownload.addEventListener('click', () => {
    setUpdateStatus('Starting download…');
    window.slidefluid.downloadUpdate();
  });

  btnInstall.addEventListener('click', () => {
    window.slidefluid.installUpdate();
  });

  // Diagnostics — run preflight
  document.getElementById('btn-run-preflight').addEventListener('click', async () => {
    document.getElementById('preflight-results').innerHTML =
      '<div class="preflight-running">Running…</div>';
    document.getElementById('btn-copy-preflight').style.display = 'none';
    const unsub = window.slidefluid.onPreflightResult((result) => {
      unsub();
      renderPreflightResults(result);
    });
    await window.slidefluid.runPreflight();
  });

  // Diagnostics — copy report
  document.getElementById('btn-copy-preflight').addEventListener('click', () => {
    const text = buildPreflightText();
    if (text) navigator.clipboard.writeText(text).catch(() => {});
  });

  // Diagnostics — log file
  document.getElementById('btn-open-log-file').addEventListener('click', () => {
    window.slidefluid.openLogFile();
  });
  document.getElementById('btn-refresh-log').addEventListener('click', loadLogTail);
}

async function loadLogTail() {
  const el = document.getElementById('log-tail');
  if (!el) return;
  const lines = await window.slidefluid.getLogTail(80);
  el.textContent = lines.length ? lines.join('\n') : 'No log entries yet.';
  el.scrollTop = el.scrollHeight;
}

let _lastPreflightResult = null;

function renderPreflightResults(result) {
  _lastPreflightResult = result;
  const el = document.getElementById('preflight-results');
  const copyBtn = document.getElementById('btn-copy-preflight');

  const checks = result.results || {};
  const entries = Object.entries(checks);
  if (entries.length === 0) {
    el.innerHTML = '<div class="preflight-running">No results returned.</div>';
    return;
  }

  el.innerHTML = entries.map(([key, val]) => {
    const ok = val === true || (val && val.ok === true);
    const label = preflightLabel(key);
    const detail = (val && typeof val === 'object' && val.detail)
      ? `<span class="preflight-detail">${escHtml(val.detail)}</span>`
      : '';
    return `<div class="preflight-row">
      <span class="preflight-icon ${ok ? 'pass' : 'fail'}">${ok ? '✓' : '✗'}</span>
      <span class="preflight-label">${escHtml(label)}</span>
      ${detail}
    </div>`;
  }).join('');

  copyBtn.style.display = '';
}

function preflightLabel(key) {
  const map = {
    poppler:         'Poppler binary',
    python:          'Python subprocess',
    python_pptx:     'python-pptx',
    pillow:          'Pillow / pdf2image',
    output_writable: 'Output folder writable',
    disk_space:      'Disk space',
    app_version:     'App version',
  };
  return map[key] || key;
}

function buildPreflightText() {
  if (!_lastPreflightResult) return '';
  const checks = _lastPreflightResult.results || {};
  const lines = ['SlideFluid 3.0 — Preflight Report', ''];
  for (const [key, val] of Object.entries(checks)) {
    const ok = val === true || (val && val.ok === true);
    const label = preflightLabel(key);
    const detail = (val && typeof val === 'object' && val.detail) ? ` (${val.detail})` : '';
    lines.push(`${ok ? '✓' : '✗'}  ${label}${detail}`);
  }
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// Average RGBA pixel data into [r, g, b]
function _avgColor(data) {
  let r = 0, g = 0, b = 0, n = 0;
  for (let i = 0; i < data.length; i += 4) { r += data[i]; g += data[i+1]; b += data[i+2]; n++; }
  return n ? [Math.round(r/n), Math.round(g/n), Math.round(b/n)] : [0, 0, 0];
}

function escHtml(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function skinText(proStr, funStr) {
  return state.settings.skin === 'fun' ? funStr : proStr;
}

function getDzLabel() {
  if (state.settings.skin !== 'fun') return 'Drop PDFs here or click to browse';
  return FUN_DROP_LABELS[_dropLabelIdx % FUN_DROP_LABELS.length];
}

function updateDzLabel() {
  const el = document.querySelector('.dz-label');
  if (el) el.textContent = getDzLabel();
}

function updateWindowTitle() {
  document.title = skinText('SlideFluid 3.0', 'SlideFluid \u2726 3.0');
}

function updateSkinText() {
  updateWindowTitle();
  const btnSettings = document.getElementById('btn-settings');
  if (btnSettings) btnSettings.title = skinText('Settings', 'Options & Stuff');
  updateDzLabel();
  updateConvertButton();
}

function triggerConfetti() {
  const btn = document.getElementById('btn-convert');
  if (!btn) return;
  const rect = btn.getBoundingClientRect();
  const colors = ['#55CDFC', '#F7A8B8', '#FFFFFF', '#F7A8B8', '#55CDFC'];

  const burst = document.createElement('div');
  burst.className = 'confetti-burst';
  document.body.appendChild(burst);

  for (let i = 0; i < 14; i++) {
    const p = document.createElement('div');
    p.className = 'confetti-particle';
    p.style.left       = `${Math.round(rect.left + rect.width * Math.random())}px`;
    p.style.top        = `${Math.round(rect.top + rect.height * 0.3)}px`;
    p.style.background = colors[i % colors.length];
    p.style.setProperty('--cdur', `${(0.55 + Math.random() * 0.45).toFixed(2)}s`);
    p.style.setProperty('--cdel', `${(Math.random() * 0.25).toFixed(2)}s`);
    p.style.setProperty('--cdx',  `${Math.round((Math.random() - 0.5) * 60)}px`);
    p.style.setProperty('--crot', `${Math.round(Math.random() * 360)}deg`);
    burst.appendChild(p);
  }

  setTimeout(() => burst.remove(), 1500);
}

function arLabel(ar) {
  const known = [
    { r: 16 / 9,  name: '16:9' },
    { r: 4 / 3,   name: '4:3' },
    { r: 16 / 10, name: '16:10' },
    { r: 3 / 2,   name: '3:2' },
    { r: 1 / 1,   name: '1:1' },
  ];
  let best = null, bestDist = Infinity;
  for (const k of known) {
    const d = Math.abs(ar - k.r);
    if (d < bestDist) { bestDist = d; best = k.name; }
  }
  return bestDist < 0.04 ? best : `${ar.toFixed(2)}:1`;
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------

document.addEventListener('DOMContentLoaded', init);
