// ─── Version ────────────────────────────────────────────────────────────────
const VERSION = '1.1.0';

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

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function getText(el, tag) {
  return el?.querySelector(tag)?.textContent?.trim() ?? '';
}

function formatDate(iso) {
  if (!iso) return '—';
  // MSPDI dates: "2024-01-15T00:00:00" or "2024-01-15"
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

function formatDuration(ptDuration) {
  // ISO 8601 duration like "PT8H0M0S", "P5DT0H0M0S"
  if (!ptDuration) return '—';
  const m = ptDuration.match(/P(?:(\d+)D)?T(?:(\d+)H)?(?:(\d+)M)?/);
  if (!m) return ptDuration;
  const days = parseInt(m[1] || 0);
  const hrs  = parseInt(m[2] || 0);
  const mins = parseInt(m[3] || 0);
  const parts = [];
  if (days) parts.push(`${days}d`);
  if (hrs)  parts.push(`${hrs}h`);
  if (mins) parts.push(`${mins}m`);
  return parts.length ? parts.join(' ') : '0h';
}

// ─── MSPDI XML parser ────────────────────────────────────────────────────────
// Parses the Microsoft Project XML (MSPDI) format exported by MS Project.
// Namespace-aware: tries both ns-prefixed and un-prefixed selectors.

function parseMspdiXml(xmlText) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, 'application/xml');

  const parseErr = doc.querySelector('parsererror');
  if (parseErr) throw new Error('XML parse error: ' + parseErr.textContent.slice(0, 120));

  // MSPDI uses namespace http://schemas.microsoft.com/project
  // querySelector works on local names, so we use just the tag names.
  function q(parent, ...tags) {
    for (const tag of tags) {
      const el = parent.querySelector(tag);
      if (el) return el;
    }
    return null;
  }

  const proj = q(doc, 'Project');
  if (!proj) throw new Error('This does not appear to be a valid MSPDI XML file (no <Project> root element found).');

  // Project-level metadata
  const projName   = getText(proj, 'Name') || getText(proj, 'Title') || '';
  const projStart  = getText(proj, 'StartDate') || getText(proj, 'Start') || '';
  const projFinish = getText(proj, 'FinishDate') || getText(proj, 'Finish') || '';

  // Tasks
  const taskEls = Array.from(doc.querySelectorAll('Task'));
  if (!taskEls.length) throw new Error('No tasks found in the file. Make sure this is an MSPDI XML export from Microsoft Project.');

  const tasks = taskEls
    .map(t => ({
      id:           parseInt(getText(t, 'ID') || '0'),
      uid:          parseInt(getText(t, 'UID') || '0'),
      name:         getText(t, 'Name'),
      start:        getText(t, 'Start'),
      finish:       getText(t, 'Finish'),
      duration:     getText(t, 'Duration'),
      pct:          parseFloat(getText(t, 'PercentComplete') || '0'),
      outline:      parseInt(getText(t, 'OutlineLevel') || '1'),
      isSummary:    getText(t, 'Summary') === '1',
      isMilestone:  getText(t, 'Milestone') === '1',
    }))
    // Filter out the implicit root task (ID=0, UID=0, no name)
    .filter(t => !(t.id === 0 && !t.name));

  return { projName, projStart, projFinish, tasks };
}

// ─── Rendering ──────────────────────────────────────────────────────────────

function renderResults({ projName, projStart, projFinish, tasks }) {
  document.getElementById('meta-name').textContent   = projName || '(unnamed)';
  document.getElementById('meta-start').textContent  = formatDate(projStart);
  document.getElementById('meta-finish').textContent = formatDate(projFinish);
  document.getElementById('meta-count').textContent  = tasks.length;

  const tbody = document.getElementById('task-body');
  tbody.innerHTML = '';

  for (const task of tasks) {
    const pct    = Math.min(100, Math.max(0, task.pct || 0));
    const indent = Math.max(0, (task.outline - 1)) * 16;

    const tr = document.createElement('tr');
    if (task.isSummary)   tr.classList.add('summary');
    if (task.isMilestone) tr.classList.add('milestone');

    tr.innerHTML = `
      <td class="col-id">${task.id}</td>
      <td class="col-name">
        <span class="task-name" style="padding-left:${indent}px">${escapeHtml(task.name)}</span>
      </td>
      <td class="col-date">${formatDate(task.start)}</td>
      <td class="col-date">${formatDate(task.finish)}</td>
      <td class="col-dur">${formatDuration(task.duration)}</td>
      <td class="col-pct">
        <div class="progress-wrap">
          <div class="progress-track">
            <div class="progress-fill" style="width:${pct}%"></div>
          </div>
          <span class="progress-label">${pct}%</span>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  }

  setState('results');
}

// ─── File reading ────────────────────────────────────────────────────────────

async function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = e => resolve(e.target.result);
    reader.onerror = () => reject(new Error('Could not read file.'));
    reader.readAsText(file, 'utf-8');
  });
}

async function readFileAsBytes(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = e => resolve(new Uint8Array(e.target.result));
    reader.onerror = () => reject(new Error('Could not read file.'));
    reader.readAsArrayBuffer(file);
  });
}

// Check first few bytes for OLE compound document magic (D0 CF 11 E0)
function isBinaryMpp(bytes) {
  return bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0;
}

// Check if text looks like XML
function looksLikeXml(text) {
  return text.trimStart().startsWith('<');
}

async function handleFile(file) {
  if (!file) return;
  setState('loading');

  try {
    // Read first 4 bytes to detect file type
    const head = await readFileAsBytes(file.slice(0, 4));

    if (isBinaryMpp(head)) {
      // True binary OLE/MPP — cannot be parsed client-side without a native library.
      throw new Error(
        'Binary .mpp files cannot be parsed directly in the browser.\n\n' +
        'To use MPP Viewer, export your project as XML from Microsoft Project:\n' +
        '  File → Save As → "XML Format (*.xml)"\n\n' +
        'Then drop the .xml file here.'
      );
    }

    // Try to read as text (XML)
    const text = await readFileAsText(file);

    if (!looksLikeXml(text)) {
      throw new Error('Unrecognised file format. Please drop an MSPDI XML file exported from Microsoft Project.');
    }

    const data = parseMspdiXml(text);
    renderResults(data);

  } catch (err) {
    console.error(err);
    const msg = err?.message || 'Could not read the file.';
    document.getElementById('error-message').textContent = msg;
    setState('error');
  }
}

// ─── Event wiring ────────────────────────────────────────────────────────────

const dropZone  = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

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

dropZone.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') fileInput.click();
});

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file) handleFile(file);
  fileInput.value = '';
});

document.getElementById('try-again-btn').addEventListener('click', () => setState('upload'));
document.getElementById('open-another-btn').addEventListener('click', () => setState('upload'));
