// ─── State machine ──────────────────────────────────────────────────────────

const sections = {
  upload:  document.getElementById('upload-section'),
  loading: document.getElementById('loading-section'),
  error:   document.getElementById('error-section'),
  results: document.getElementById('results-section'),
};

function setState(state) {
  for (const [key, el] of Object.entries(sections)) {
    el.classList.toggle('hidden', key !== state);
  }
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function safeGet(fn, fallback = null) {
  try {
    const v = fn();
    return v ?? fallback;
  } catch {
    return fallback;
  }
}

function formatDate(dateVal) {
  if (!dateVal) return '—';
  try {
    // MPXJ dates may be Java LocalDateTime or similar — coerce to string then parse
    const s = String(dateVal);
    // Typical ISO-like: "2024-01-15T00:00:00"
    const d = new Date(s);
    if (isNaN(d.getTime())) return s;
    return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
  } catch {
    return '—';
  }
}

function formatDuration(durVal) {
  if (!durVal) return '—';
  try {
    // Try .toString() first; MPXJ Duration has a toString that gives "X days" etc.
    const s = String(durVal);
    if (s && s !== 'null' && s !== 'undefined') return s;
    return '—';
  } catch {
    return '—';
  }
}

// ─── Rendering ──────────────────────────────────────────────────────────────

function renderResults(project, tasks) {
  // Project info bar
  const name   = safeGet(() => project.getName(), '');
  const start  = safeGet(() => project.getStartDate());
  const finish = safeGet(() => project.getFinishDate());

  document.getElementById('meta-name').textContent   = name || '(unnamed)';
  document.getElementById('meta-start').textContent  = formatDate(start);
  document.getElementById('meta-finish').textContent = formatDate(finish);
  document.getElementById('meta-count').textContent  = tasks.length;

  // Task table
  const tbody = document.getElementById('task-body');
  tbody.innerHTML = '';

  for (const task of tasks) {
    const id       = safeGet(() => task.getID(), '');
    const taskName = safeGet(() => task.getName(), '');
    const outline  = safeGet(() => task.getOutlineLevel(), 1) || 1;
    const tStart   = safeGet(() => task.getStart());
    const tFinish  = safeGet(() => task.getFinish());
    const dur      = safeGet(() => task.getDuration());
    const pct      = safeGet(() => task.getPercentageComplete(), 0);
    const isSummary = safeGet(() => task.getSummary(), false);

    const pctNum = Math.min(100, Math.max(0, Number(pct) || 0));
    const indent = (outline - 1) * 16; // px per level

    const tr = document.createElement('tr');
    if (isSummary) tr.classList.add('summary');

    tr.innerHTML = `
      <td class="col-id">${id ?? ''}</td>
      <td class="col-name">
        <span class="task-name" style="padding-left:${indent}px">${escapeHtml(taskName || '')}</span>
      </td>
      <td class="col-date">${formatDate(tStart)}</td>
      <td class="col-date">${formatDate(tFinish)}</td>
      <td class="col-dur">${formatDuration(dur)}</td>
      <td class="col-pct">
        <div class="progress-wrap">
          <div class="progress-track">
            <div class="progress-fill" style="width:${pctNum}%"></div>
          </div>
          <span class="progress-label">${pctNum}%</span>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  }

  setState('results');
}

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ─── MPXJ parsing ───────────────────────────────────────────────────────────

let mpxjModule = null;

async function getMpxj() {
  if (!mpxjModule) {
    mpxjModule = await import('https://cdn.jsdelivr.net/npm/mpxj/+esm');
  }
  return mpxjModule;
}

async function readProject(arrayBuffer) {
  const mpxj = await getMpxj();

  // The ESM build exposes a default export that is the Emscripten module factory
  // or a named ProjectFile / ProjectReader API depending on the version.
  // We use the FileReader / UniversalProjectReader approach.
  let reader;
  if (mpxj.UniversalProjectReader) {
    reader = new mpxj.UniversalProjectReader();
  } else if (mpxj.default && mpxj.default.UniversalProjectReader) {
    reader = new mpxj.default.UniversalProjectReader();
  } else {
    // Fall back: try every likely named export
    const Ctor =
      mpxj.MPPReader ||
      mpxj.default?.MPPReader ||
      mpxj.ProjectReader ||
      mpxj.default?.ProjectReader;
    if (!Ctor) throw new Error('Could not find a project reader in the mpxj module.');
    reader = new Ctor();
  }

  const uint8 = new Uint8Array(arrayBuffer);
  const project = await reader.read(uint8);

  const taskList = safeGet(() => project.getTasks(), null);
  if (!taskList) throw new Error('No task data found in file.');

  // getTasks() returns a Java List-like object; iterate with size()/get() or spread
  let tasks = [];
  if (typeof taskList[Symbol.iterator] === 'function') {
    tasks = [...taskList];
  } else {
    const size = safeGet(() => taskList.size(), 0);
    for (let i = 0; i < size; i++) {
      tasks.push(taskList.get(i));
    }
  }

  // Filter out the invisible root task (ID 0 / no name)
  tasks = tasks.filter(t => {
    const id   = safeGet(() => t.getID(), null);
    const name = safeGet(() => t.getName(), null);
    return id !== 0 && id !== null && name;
  });

  // Try to get project properties for the meta bar
  let props = null;
  try { props = project.getProjectProperties(); } catch { /* ok */ }
  const projectMeta = props || project;

  return { project: projectMeta, tasks };
}

// ─── File handling ──────────────────────────────────────────────────────────

async function handleFile(file) {
  if (!file) return;
  setState('loading');
  try {
    const buffer = await file.arrayBuffer();
    const { project, tasks } = await readProject(buffer);
    renderResults(project, tasks);
  } catch (err) {
    console.error(err);
    document.getElementById('error-message').textContent =
      err?.message ? `Error: ${err.message}` : 'Could not read the file. Is it a valid .mpp file?';
    setState('error');
  }
}

// ─── Event wiring ───────────────────────────────────────────────────────────

const dropZone  = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

// Drag events
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('dragend',   () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer?.files?.[0];
  if (file) handleFile(file);
});

// Keyboard activation of drop zone
dropZone.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') fileInput.click();
});

// File input
fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file) handleFile(file);
  fileInput.value = '';
});

// Reset buttons
document.getElementById('try-again-btn').addEventListener('click', () => setState('upload'));
document.getElementById('open-another-btn').addEventListener('click', () => setState('upload'));
