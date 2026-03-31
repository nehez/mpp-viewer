// ─── Layout constants (must match CSS custom properties) ─────────────────────
const ROW_H    = 32;
const HEADER_H = 36;
const BAR_H    = 18;
const BAR_TOP  = (ROW_H - BAR_H) / 2;
const LINK_M   = 10;  // routing margin for link arrows

// ─── Module state ─────────────────────────────────────────────────────────────
const state = {
  allTasks:      [],
  filteredTasks: [],
  uidToTask:     {},   // uid → task
  minD: null, maxD: null,
  pxPerDay:   8,
  showLinks:  true,
  filterText:   '',
  filterStatus: 'all',  // all | not-started | in-progress | complete
  filterType:   'all',  // all | normal | summary | milestone
};

// ─── State machine ────────────────────────────────────────────────────────────
const sections = {
  upload:  document.getElementById('upload-section'),
  loading: document.getElementById('loading-section'),
  error:   document.getElementById('error-section'),
  results: document.getElementById('results-section'),
};
function setState(s) {
  for (const [k, el] of Object.entries(sections)) el.classList.toggle('hidden', k !== s);
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function getText(el, tag) { return el?.querySelector(tag)?.textContent?.trim() ?? ''; }

function formatDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return isNaN(d) ? iso : d.toLocaleDateString(undefined, { year:'numeric', month:'short', day:'numeric' });
}

function formatDuration(pt) {
  if (!pt) return '—';
  const m = pt.match(/P(?:[\d.]+Y)?(?:[\d.]+M)?(?:([\d.]+)D)?(?:T(?:([\d.]+)H)?(?:([\d.]+)M)?(?:[\d.]+S)?)?/);
  if (!m) return pt;
  const totalH = (parseFloat(m[1]||0)*8) + parseFloat(m[2]||0) + (parseFloat(m[3]||0)/60);
  if (totalH === 0) return '0 days';
  const days = totalH / 8;
  if (days < 1) return `${+(totalH.toFixed(1))} h`;
  const r = Math.round(days * 10) / 10;
  return `${r} day${r === 1 ? '' : 's'}`;
}

function parseDate(iso) {
  if (!iso) return null;
  const d = new Date(iso);
  return isNaN(d) ? null : d;
}

function dayDiff(a, b) { return Math.floor((b - a) / 86400000); }

function zoomLabel(px) {
  if (px < 2)  return 'Years';
  if (px < 4)  return 'Half Year';
  if (px < 7)  return 'Quarter';
  if (px < 15) return 'Month';
  if (px < 25) return '2 Weeks';
  if (px < 40) return 'Week';
  return 'Day';
}

// ─── MSPDI XML parser ─────────────────────────────────────────────────────────
function parseMspdiXml(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
  const parseErr = doc.querySelector('parsererror');
  if (parseErr) throw new Error('XML parse error: ' + parseErr.textContent.slice(0, 120));

  const proj = doc.querySelector('Project');
  if (!proj) throw new Error('Not a valid MSPDI XML file — no <Project> element found.');

  const projName   = getText(proj, 'Name') || getText(proj, 'Title') || '';
  const projStart  = getText(proj, 'StartDate') || getText(proj, 'Start') || '';
  const projFinish = getText(proj, 'FinishDate') || getText(proj, 'Finish') || '';

  const taskEls = Array.from(doc.querySelectorAll('Task'));
  if (!taskEls.length) throw new Error('No tasks found. Make sure this is an MSPDI XML export from Microsoft Project.');

  const tasks = taskEls
    .map(t => {
      // Parse predecessor links for this task
      const links = Array.from(t.querySelectorAll('PredecessorLink')).map(pl => ({
        predUID: parseInt(getText(pl, 'PredecessorUID') || '0'),
        type:    parseInt(getText(pl, 'Type')           || '1'),  // 0=FF,1=FS,2=SF,3=SS
        lag:     parseInt(getText(pl, 'LinkLag')        || '0'),
      })).filter(l => l.predUID > 0);

      return {
        uid:         parseInt(getText(t, 'UID')  || '0'),
        id:          parseInt(getText(t, 'ID')   || '0'),
        name:        getText(t, 'Name'),
        start:       getText(t, 'Start'),
        finish:      getText(t, 'Finish'),
        duration:    getText(t, 'Duration'),
        pct:         parseFloat(getText(t, 'PercentComplete') || '0'),
        outline:     parseInt(getText(t, 'OutlineLevel') || '1'),
        isSummary:   getText(t, 'Summary')   === '1',
        isMilestone: getText(t, 'Milestone') === '1',
        links,
      };
    })
    .filter(t => !(t.id === 0 && !t.name));

  return { projName, projStart, projFinish, tasks };
}

// ─── Filter logic ─────────────────────────────────────────────────────────────
function applyFilters() {
  const { filterText, filterStatus, filterType } = state;
  const q = filterText.trim().toLowerCase();

  state.filteredTasks = state.allTasks.filter(t => {
    if (q && !t.name.toLowerCase().includes(q)) return false;
    switch (filterStatus) {
      case 'not-started':  if (t.pct !== 0) return false; break;
      case 'in-progress':  if (t.pct === 0 || t.pct >= 100) return false; break;
      case 'complete':     if (t.pct < 100) return false; break;
    }
    switch (filterType) {
      case 'normal':    if (t.isSummary || t.isMilestone) return false; break;
      case 'summary':   if (!t.isSummary) return false; break;
      case 'milestone': if (!t.isMilestone) return false; break;
    }
    return true;
  });

  const total = state.allTasks.length;
  const shown = state.filteredTasks.length;
  document.getElementById('filter-count').textContent =
    shown === total ? `${total} tasks` : `${shown} / ${total}`;

  renderTable(state.filteredTasks);
  redrawGantt();
}

// ─── Table renderer ───────────────────────────────────────────────────────────
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

// ─── Gantt: dependency arrow SVG ──────────────────────────────────────────────
function buildLinkPath(fx, fy, tx, ty, type) {
  // FS (1): pred.right → succ.left — most common
  if (type === 1) {
    if (tx > fx + 2) {
      const mid = Math.round((fx + tx) / 2);
      return `M${fx},${fy} H${mid} V${ty} H${tx}`;
    }
    // wrap around when successor starts before predecessor finishes
    const dir  = fy <= ty ? 1 : -1;
    const elby = Math.round(fy + dir * ROW_H * 0.65);
    return `M${fx},${fy} H${fx+LINK_M} V${elby} H${tx-LINK_M} V${ty} H${tx}`;
  }
  // FF (0): pred.right → succ.right  (exit right, arrive from right → last segment goes left)
  if (type === 0) {
    const rx = Math.max(fx, tx) + LINK_M;
    return `M${fx},${fy} H${rx} V${ty} H${tx}`;
  }
  // SS (3): pred.left → succ.left   (exit left, arrive from left → last segment goes right)
  if (type === 3) {
    const lx = Math.min(fx, tx) - LINK_M;
    return `M${fx},${fy} H${lx} V${ty} H${tx}`;
  }
  // SF (2): pred.left → succ.right
  const rx = Math.max(fx, tx) + LINK_M;
  return `M${fx},${fy} H${rx} V${ty} H${tx}`;
}

function drawLinksSvg(tasks, uidToRow, totalW, totalH) {
  const NS  = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(NS, 'svg');
  svg.setAttribute('class', 'gantt-links-svg');
  svg.style.cssText =
    `position:absolute;top:0;left:0;width:${totalW}px;height:${totalH}px;` +
    `pointer-events:none;overflow:visible;z-index:5`;

  // Arrowhead marker (orient=auto follows last path segment direction)
  const defs = document.createElementNS(NS, 'defs');
  defs.innerHTML =
    `<marker id="gantt-ah" markerWidth="8" markerHeight="6"` +
    ` refX="7" refY="3" orient="auto" markerUnits="userSpaceOnUse">` +
    `<polygon points="0,0 8,3 0,6" class="gantt-ah-poly"/></marker>`;
  svg.appendChild(defs);

  const { minD, pxPerDay } = state;

  for (const task of tasks) {
    for (const link of (task.links || [])) {
      const pred = state.uidToTask[link.predUID];
      if (!pred) continue;

      const predIdx = uidToRow[pred.uid];
      const succIdx = uidToRow[task.uid];
      if (predIdx === undefined || succIdx === undefined) continue;

      const pS = parseDate(pred.start), pF = parseDate(pred.finish);
      const sS = parseDate(task.start), sF = parseDate(task.finish);
      if (!pS || !pF || !sS || !sF) continue;

      const pL = Math.round(dayDiff(minD, pS) * pxPerDay);
      const pR = Math.round(dayDiff(minD, pF) * pxPerDay);
      const sL = Math.round(dayDiff(minD, sS) * pxPerDay);
      const sR = Math.round(dayDiff(minD, sF) * pxPerDay);
      const pY = Math.round(predIdx * ROW_H + ROW_H / 2);
      const sY = Math.round(succIdx * ROW_H + ROW_H / 2);

      // Source/target x per link type
      let fx, tx;
      switch (link.type) {
        case 0: fx = pR; tx = sR; break;  // FF
        case 1: fx = pR; tx = sL; break;  // FS
        case 2: fx = pL; tx = sR; break;  // SF
        case 3: fx = pL; tx = sL; break;  // SS
        default: continue;
      }

      const path = document.createElementNS(NS, 'path');
      path.setAttribute('d', buildLinkPath(fx, pY, tx, sY, link.type));
      path.setAttribute('class', `link-path link-type-${link.type}`);
      path.setAttribute('marker-end', 'url(#gantt-ah)');
      svg.appendChild(path);
    }
  }
  return svg;
}

// ─── Gantt: main draw ─────────────────────────────────────────────────────────
function redrawGantt() {
  const { filteredTasks: tasks, minD, maxD, pxPerDay, showLinks } = state;

  const headerEl = document.getElementById('gantt-header');
  const barsEl   = document.getElementById('gantt-bars');
  headerEl.innerHTML = '';
  barsEl.innerHTML   = '';

  if (!minD || !maxD) {
    barsEl.innerHTML =
      '<div style="padding:20px 16px;color:var(--muted);font-size:13px">No date information found in this file.</div>';
    return;
  }

  const totalDays = dayDiff(minD, maxD) + 1;
  const totalW    = Math.round(totalDays * pxPerDay);
  const totalH    = tasks.length * ROW_H;

  // ── Timeline header ──
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

  // ── Canvas ──
  barsEl.style.width  = totalW + 'px';
  barsEl.style.height = Math.max(totalH, 1) + 'px';

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

  // Build UID→row-index map (only for filtered tasks)
  const uidToRow = {};
  tasks.forEach((t, i) => { uidToRow[t.uid] = i; });

  // Bar rows
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
        bar.className  = 'gantt-bar' + (task.isSummary ? ' summary-bar' : '');
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

  // SVG link arrows
  if (showLinks) {
    barsEl.appendChild(drawLinksSvg(tasks, uidToRow, totalW, Math.max(totalH, 1)));
  }

  // Update zoom label
  document.getElementById('zoom-label').textContent = zoomLabel(pxPerDay);
}

// ─── Scroll sync (one-time setup per file load) ───────────────────────────────
let _scrollAbort = null;
function setupScrollSync() {
  if (_scrollAbort) _scrollAbort.abort();
  _scrollAbort = new AbortController();
  const sig = _scrollAbort.signal;

  const barsWrap   = document.getElementById('gantt-bars-wrap');
  const leftPanel  = document.getElementById('ms-left');
  const headerWrap = document.getElementById('gantt-header-wrap');

  let busy = false;
  barsWrap.addEventListener('scroll', () => {
    if (busy) return; busy = true;
    requestAnimationFrame(() => {
      leftPanel.scrollTop    = barsWrap.scrollTop;
      headerWrap.scrollLeft  = barsWrap.scrollLeft;
      busy = false;
    });
  }, { signal: sig });

  leftPanel.addEventListener('wheel', e => {
    e.preventDefault();
    barsWrap.scrollTop += e.deltaY;
  }, { passive: false, signal: sig });

  leftPanel.addEventListener('scroll', () => {
    if (busy) return; busy = true;
    requestAnimationFrame(() => {
      barsWrap.scrollTop = leftPanel.scrollTop;
      busy = false;
    });
  }, { signal: sig });

  // Ctrl+Scroll zoom on the gantt area
  barsWrap.addEventListener('wheel', e => {
    if (!e.ctrlKey && !e.metaKey) return;
    e.preventDefault();
    const factor = e.deltaY < 0 ? 1.35 : 1 / 1.35;
    state.pxPerDay = Math.max(1, Math.min(80, state.pxPerDay * factor));
    redrawGantt();
  }, { passive: false, signal: sig });
}

// ─── Render results ───────────────────────────────────────────────────────────
function renderResults({ projName, projStart, projFinish, tasks }) {
  state.allTasks  = tasks;
  state.uidToTask = Object.fromEntries(tasks.map(t => [t.uid, t]));

  // Compute date range
  let minD = null, maxD = null;
  for (const t of tasks) {
    const s = parseDate(t.start), f = parseDate(t.finish);
    if (s && (!minD || s < minD)) minD = s;
    if (f && (!maxD || f > maxD)) maxD = f;
  }
  if (minD && maxD) {
    state.minD = new Date(minD.getFullYear(), minD.getMonth(), 1);
    state.maxD = new Date(maxD.getFullYear(), maxD.getMonth() + 1, 0);
    const totalDays = dayDiff(state.minD, state.maxD) + 1;
    // Auto-fit zoom to ~900px
    state.pxPerDay = Math.max(1.5, Math.min(40, 900 / totalDays));
  } else {
    state.minD = state.maxD = null;
    state.pxPerDay = 8;
  }

  // Meta bar
  document.getElementById('meta-name').textContent   = projName || '(unnamed)';
  document.getElementById('meta-start').textContent  = formatDate(projStart);
  document.getElementById('meta-finish').textContent = formatDate(projFinish);
  document.getElementById('meta-count').textContent  = tasks.length;

  // Reset filters & links toggle UI
  state.filterText   = '';
  state.filterStatus = 'all';
  state.filterType   = 'all';
  document.getElementById('filter-search').value  = '';
  document.getElementById('filter-status').value  = 'all';
  document.getElementById('filter-type').value    = 'all';
  state.showLinks = true;
  document.getElementById('btn-links').classList.add('active');

  setupScrollSync();
  applyFilters();   // populates filteredTasks, renders table + gantt
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
  return b[0]===0xD0 && b[1]===0xCF && b[2]===0x11 && b[3]===0xE0;
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
        '  File → Save As → "XML Format (*.xml)"\n\nThen drop the .xml file here.'
      );
    }
    const text = await readAsText(file);
    if (!text.trimStart().startsWith('<'))
      throw new Error('Unrecognised format. Please drop an MSPDI XML file exported from Microsoft Project.');
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
  e.preventDefault(); dropZone.classList.remove('drag-over');
  const f = e.dataTransfer?.files?.[0]; if (f) handleFile(f);
});
dropZone.addEventListener('keydown', e => { if (e.key==='Enter'||e.key===' ') fileInput.click(); });
fileInput.addEventListener('change', () => {
  const f = fileInput.files?.[0]; if (f) handleFile(f); fileInput.value = '';
});

document.getElementById('try-again-btn').addEventListener('click',    () => setState('upload'));
document.getElementById('open-another-btn').addEventListener('click', () => setState('upload'));

// Filter controls
document.getElementById('filter-search').addEventListener('input', e => {
  state.filterText = e.target.value; applyFilters();
});
document.getElementById('filter-status').addEventListener('change', e => {
  state.filterStatus = e.target.value; applyFilters();
});
document.getElementById('filter-type').addEventListener('change', e => {
  state.filterType = e.target.value; applyFilters();
});
document.getElementById('btn-clear-filters').addEventListener('click', () => {
  state.filterText = ''; state.filterStatus = 'all'; state.filterType = 'all';
  document.getElementById('filter-search').value = '';
  document.getElementById('filter-status').value = 'all';
  document.getElementById('filter-type').value   = 'all';
  applyFilters();
});

// Links toggle
document.getElementById('btn-links').addEventListener('click', () => {
  state.showLinks = !state.showLinks;
  document.getElementById('btn-links').classList.toggle('active', state.showLinks);
  redrawGantt();
});

// Zoom controls
document.getElementById('btn-zoom-in').addEventListener('click', () => {
  state.pxPerDay = Math.min(80, state.pxPerDay * 1.5); redrawGantt();
});
document.getElementById('btn-zoom-out').addEventListener('click', () => {
  state.pxPerDay = Math.max(1, state.pxPerDay / 1.5); redrawGantt();
});
document.getElementById('btn-zoom-fit').addEventListener('click', () => {
  const w = document.getElementById('gantt-bars-wrap').clientWidth || 900;
  if (state.minD && state.maxD) {
    state.pxPerDay = Math.max(1, w / (dayDiff(state.minD, state.maxD) + 1));
    redrawGantt();
  }
});

// ─── Splitter drag ────────────────────────────────────────────────────────────
(function () {
  const splitter = document.getElementById('ms-splitter');
  const msLeft   = document.getElementById('ms-left');
  let startX = 0, startW = 0, dragging = false;

  splitter.addEventListener('mousedown', e => {
    dragging = true; startX = e.clientX; startW = msLeft.offsetWidth;
    splitter.classList.add('dragging');
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    msLeft.style.width = Math.max(160, Math.min(1000, startW + e.clientX - startX)) + 'px';
  });
  document.addEventListener('mouseup', () => {
    if (!dragging) return; dragging = false;
    splitter.classList.remove('dragging');
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
  });
})();
