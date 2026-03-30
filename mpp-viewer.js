// ─── Layout constants (must match CSS) ───────────────────────────────────────
const ROW_H    = 32;   // px — height of every data row
const HEADER_H = 36;   // px — height of column/timeline header row
const BAR_H    = 18;   // px — bar thickness inside the row
const BAR_TOP  = (ROW_H - BAR_H) / 2;

// ─── State machine ───────────────────────────────────────────────────────────

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

// ─── Helpers ─────────────────────────────────────────────────────────────────

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function getText(el, tag) {
  return el?.querySelector(tag)?.textContent?.trim() ?? '';
}

function formatDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

// Parse ISO 8601 duration → work-days (8 h/day as MS Project uses)
// MSPDI stores duration in work-hours: "PT8H" = 1 day, "PT40H" = 5 days
function formatDuration(pt) {
  if (!pt) return '—';
  const m = pt.match(/P(?:[\d.]+Y)?(?:[\d.]+M)?(?:([\d.]+)D)?(?:T(?:([\d.]+)H)?(?:([\d.]+)M)?(?:[\d.]+S)?)?/);
  if (!m) return pt;
  const totalH = (parseFloat(m[1] || 0) * 8) + parseFloat(m[2] || 0) + (parseFloat(m[3] || 0) / 60);
  if (totalH === 0) return '0 days';
  const days = totalH / 8;
  if (days < 1) return `${+(totalH.toFixed(1))} h`;
  const r = Math.round(days * 10) / 10;
  return `${r} day${r === 1 ? '' : 's'}`;
}

function parseDate(iso) {
  if (!iso) return null;
  const d = new Date(iso);
  return isNaN(d.getTime()) ? null : d;
}

function dayDiff(a, b) {
  return Math.floor((b - a) / 86400000);
}

// ─── MSPDI XML parser ─────────────────────────────────────────────────────────

function parseMspdiXml(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
  const err = doc.querySelector('parsererror');
  if (err) throw new Error('XML parse error: ' + err.textContent.slice(0, 120));

  const proj = doc.querySelector('Project');
  if (!proj) throw new Error('Not a valid MSPDI XML file — no <Project> element found.');

  const projName   = getText(proj, 'Name') || getText(proj, 'Title') || '';
  const projStart  = getText(proj, 'StartDate') || getText(proj, 'Start') || '';
  const projFinish = getText(proj, 'FinishDate') || getText(proj, 'Finish') || '';

  const taskEls = Array.from(doc.querySelectorAll('Task'));
  if (!taskEls.length) throw new Error('No tasks found. Make sure this is an MSPDI XML export from Microsoft Project.');

  const tasks = taskEls
    .map(t => ({
      id:          parseInt(getText(t, 'ID') || '0'),
      name:        getText(t, 'Name'),
      start:       getText(t, 'Start'),
      finish:      getText(t, 'Finish'),
      duration:    getText(t, 'Duration'),
      pct:         parseFloat(getText(t, 'PercentComplete') || '0'),
      outline:     parseInt(getText(t, 'OutlineLevel') || '1'),
      isSummary:   getText(t, 'Summary') === '1',
      isMilestone: getText(t, 'Milestone') === '1',
    }))
    .filter(t => !(t.id === 0 && !t.name));

  return { projName, projStart, projFinish, tasks };
}

// ─── Render: table left panel ─────────────────────────────────────────────────

function renderTable(tasks) {
  const tbody = document.getElementById('task-body');
  tbody.innerHTML = '';
  for (const task of tasks) {
    const pct    = Math.min(100, Math.max(0, task.pct || 0));
    const indent = Math.max(0, task.outline - 1) * 14;
    const tr = document.createElement('tr');
    if (task.isSummary)   tr.classList.add('summary');
    if (task.isMilestone) tr.classList.add('milestone');
    tr.innerHTML = `
      <td class="col-id">${task.id}</td>
      <td class="col-name"><span class="task-name" style="padding-left:${indent}px">${escapeHtml(task.name)}</span></td>
      <td class="col-dur">${formatDuration(task.duration)}</td>
      <td class="col-date">${formatDate(task.start)}</td>
      <td class="col-date">${formatDate(task.finish)}</td>
      <td class="col-pct">
        <div class="progress-wrap">
          <div class="progress-track"><div class="progress-fill" style="width:${pct}%"></div></div>
          <span class="progress-label">${pct}%</span>
        </div>
      </td>`;
    tbody.appendChild(tr);
  }
}

// ─── Render: Gantt right panel ────────────────────────────────────────────────

function renderGantt(tasks) {
  const headerEl   = document.getElementById('gantt-header');
  const barsEl     = document.getElementById('gantt-bars');
  headerEl.innerHTML = '';
  barsEl.innerHTML   = '';

  // Date range
  let minD = null, maxD = null;
  for (const t of tasks) {
    const s = parseDate(t.start), f = parseDate(t.finish);
    if (s && (!minD || s < minD)) minD = s;
    if (f && (!maxD || f > maxD)) maxD = f;
  }
  if (!minD || !maxD) {
    barsEl.innerHTML = '<div style="padding:20px 16px;color:var(--muted);font-size:13px">No date information found in this file.</div>';
    return;
  }

  // Snap to month boundaries
  minD = new Date(minD.getFullYear(), minD.getMonth(), 1);
  maxD = new Date(maxD.getFullYear(), maxD.getMonth() + 1, 0);

  const totalDays = dayDiff(minD, maxD) + 1;
  // Target ~900px timeline; clamp px/day to a readable range
  const pxPerDay  = Math.max(3, Math.min(40, 900 / totalDays));
  const totalW    = Math.round(totalDays * pxPerDay);

  // ── Month header cells ──
  headerEl.style.width = totalW + 'px';
  let cur = new Date(minD.getFullYear(), minD.getMonth(), 1);
  while (cur <= maxD) {
    const y = cur.getFullYear(), mo = cur.getMonth();
    const daysInMo = new Date(y, mo + 1, 0).getDate();
    const cell = document.createElement('div');
    cell.className  = 'gh-month';
    cell.style.width = Math.round(daysInMo * pxPerDay) + 'px';
    cell.textContent = cur.toLocaleString('default', { month: 'short', year: 'numeric' });
    headerEl.appendChild(cell);
    cur = new Date(y, mo + 1, 1);
  }

  // ── Canvas: grid lines + today + bar rows ──
  barsEl.style.width  = totalW + 'px';
  barsEl.style.height = (tasks.length * ROW_H) + 'px';

  // Month grid lines
  let dayOff = 0;
  cur = new Date(minD.getFullYear(), minD.getMonth(), 1);
  while (cur <= maxD) {
    const y = cur.getFullYear(), mo = cur.getMonth();
    if (dayOff > 0) {
      const ln = document.createElement('div');
      ln.className  = 'gantt-vline';
      ln.style.left = Math.round(dayOff * pxPerDay) + 'px';
      barsEl.appendChild(ln);
    }
    dayOff += new Date(y, mo + 1, 0).getDate();
    cur = new Date(y, mo + 1, 1);
  }

  // Today line
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (today >= minD && today <= maxD) {
    const tl = document.createElement('div');
    tl.className  = 'gantt-today';
    tl.style.left = Math.round(dayDiff(minD, today) * pxPerDay) + 'px';
    barsEl.appendChild(tl);
  }

  // One row div per task (must align pixel-perfectly with table rows)
  tasks.forEach((task, i) => {
    const row = document.createElement('div');
    row.className = 'gbar-row' + (task.isSummary ? ' summary' : '');
    row.style.top = (i * ROW_H) + 'px';

    const s = parseDate(task.start), f = parseDate(task.finish);
    if (s && f) {
      const left  = Math.round(dayDiff(minD, s) * pxPerDay);
      const width = Math.max(Math.round(dayDiff(s, f) * pxPerDay), task.isMilestone ? 0 : 2);

      if (task.isMilestone) {
        const d = document.createElement('div');
        d.className  = 'gantt-diamond';
        d.style.left = (left - 7) + 'px';
        row.appendChild(d);
      } else {
        const bar = document.createElement('div');
        bar.className = 'gantt-bar' + (task.isSummary ? ' summary-bar' : '');
        bar.style.cssText = `left:${left}px;width:${width}px;top:${BAR_TOP}px;height:${BAR_H}px`;
        if (!task.isSummary && task.pct > 0) {
          const done = document.createElement('div');
          done.className = 'gantt-done';
          done.style.width = Math.min(100, task.pct) + '%';
          bar.appendChild(done);
        }
        row.appendChild(bar);
      }
    }

    barsEl.appendChild(row);
  });

  // ── Scroll sync ──────────────────────────────────────────────────────────
  // gantt-bars-wrap drives everything; left table and gantt header follow
  const barsWrap   = document.getElementById('gantt-bars-wrap');
  const leftPanel  = document.getElementById('ms-left');
  const headerWrap = document.getElementById('gantt-header-wrap');

  let ticking = false;
  barsWrap.addEventListener('scroll', () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => {
      leftPanel.scrollTop     = barsWrap.scrollTop;
      headerWrap.scrollLeft   = barsWrap.scrollLeft;
      ticking = false;
    });
  });
  // Let left panel wheel events drive vertical scroll
  leftPanel.addEventListener('wheel', e => {
    e.preventDefault();
    barsWrap.scrollTop += e.deltaY;
  }, { passive: false });
  // Left panel scrollbar also syncs right
  leftPanel.addEventListener('scroll', () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => {
      barsWrap.scrollTop = leftPanel.scrollTop;
      ticking = false;
    });
  });
}

// ─── Render results (both panels) ────────────────────────────────────────────

function renderResults({ projName, projStart, projFinish, tasks }) {
  document.getElementById('meta-name').textContent   = projName || '(unnamed)';
  document.getElementById('meta-start').textContent  = formatDate(projStart);
  document.getElementById('meta-finish').textContent = formatDate(projFinish);
  document.getElementById('meta-count').textContent  = tasks.length;

  renderTable(tasks);
  renderGantt(tasks);
  setState('results');
}

// ─── File reading ─────────────────────────────────────────────────────────────

async function readAsText(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload  = e => res(e.target.result);
    r.onerror = () => rej(new Error('Could not read file.'));
    r.readAsText(file, 'utf-8');
  });
}

async function readAsBytes(blob) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload  = e => res(new Uint8Array(e.target.result));
    r.onerror = () => rej(new Error('Could not read file.'));
    r.readAsArrayBuffer(blob);
  });
}

function isBinaryMpp(b) {
  return b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0;
}

async function handleFile(file) {
  if (!file) return;
  setState('loading');
  try {
    const head = await readAsBytes(file.slice(0, 4));
    if (isBinaryMpp(head)) {
      throw new Error(
        'Binary .mpp files cannot be parsed directly in the browser.\n\n' +
        'Export your project as XML from Microsoft Project:\n' +
        '  File → Save As → "XML Format (*.xml)"\n\n' +
        'Then drop the .xml file here.'
      );
    }
    const text = await readAsText(file);
    if (!text.trimStart().startsWith('<')) {
      throw new Error('Unrecognised format. Please drop an MSPDI XML file exported from Microsoft Project.');
    }
    renderResults(parseMspdiXml(text));
  } catch (err) {
    console.error(err);
    document.getElementById('error-message').textContent = err?.message || 'Could not read the file.';
    setState('error');
  }
}

// ─── Event wiring ─────────────────────────────────────────────────────────────

const dropZone  = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('dragend',   () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const f = e.dataTransfer?.files?.[0];
  if (f) handleFile(f);
});
dropZone.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') fileInput.click();
});
fileInput.addEventListener('change', () => {
  const f = fileInput.files?.[0];
  if (f) handleFile(f);
  fileInput.value = '';
});

document.getElementById('try-again-btn').addEventListener('click',    () => setState('upload'));
document.getElementById('open-another-btn').addEventListener('click', () => setState('upload'));

// ─── Splitter drag ────────────────────────────────────────────────────────────
(function () {
  const splitter = document.getElementById('ms-splitter');
  const msLeft   = document.getElementById('ms-left');
  if (!splitter || !msLeft) return;

  let startX = 0, startW = 0, dragging = false;

  splitter.addEventListener('mousedown', e => {
    dragging = true;
    startX   = e.clientX;
    startW   = msLeft.offsetWidth;
    splitter.classList.add('dragging');
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  });

  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const newW = Math.max(180, Math.min(900, startW + e.clientX - startX));
    msLeft.style.width = newW + 'px';
  });

  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    splitter.classList.remove('dragging');
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
  });
})();
