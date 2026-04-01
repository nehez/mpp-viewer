// ─── Layout constants (must match CSS custom properties) ─────────────────────
const ROW_H    = 32;
const HEADER_H = 36;
const BAR_H    = 18;
const BAR_TOP  = (ROW_H - BAR_H) / 2;
const LINK_M   = 10;

// ─── Module state ─────────────────────────────────────────────────────────────
const state = {
  allTasks:          [],
  filteredTasks:     [],
  uidToTask:         {},
  minD: null, maxD: null,
  pxPerDay:          8,
  showLinks:         true,
  filterText:        '',
  filterStatus:      'all',
  filterType:        'all',
  collapsedSummaries: new Set(),
  projectName: '',
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
  if (px < 0.15) return 'Decade+';
  if (px < 0.35) return 'Decade';
  if (px < 0.6)  return '5 Years';
  if (px < 1)    return 'Years';
  if (px < 2)    return 'Half Year';
  if (px < 4)    return 'Quarter';
  if (px < 7)    return 'Month';
  if (px < 15)   return '2 Weeks';
  if (px < 25)   return 'Week';
  if (px < 40)   return 'Day';
  return 'Hours';
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

  // Build resource UID → name map
  const resourceMap = {};
  Array.from(doc.querySelectorAll('Resources > Resource')).forEach(r => {
    const uid  = parseInt(getText(r, 'UID') || '0');
    const name = getText(r, 'Name');
    if (uid > 0 && name) resourceMap[uid] = name;
  });

  // Build task UID → resource names
  const taskResMap = {};
  Array.from(doc.querySelectorAll('Assignments > Assignment')).forEach(a => {
    const tuid = parseInt(getText(a, 'TaskUID') || '0');
    const ruid = parseInt(getText(a, 'ResourceUID') || '0');
    if (tuid > 0 && ruid > 0 && resourceMap[ruid]) {
      if (!taskResMap[tuid]) taskResMap[tuid] = [];
      taskResMap[tuid].push(resourceMap[ruid]);
    }
  });

  const taskEls = Array.from(doc.querySelectorAll('Task'));
  if (!taskEls.length) throw new Error('No tasks found. Make sure this is an MSPDI XML export from Microsoft Project.');

  const tasks = taskEls
    .map(t => {
      const uid = parseInt(getText(t, 'UID') || '0');
      const links = Array.from(t.querySelectorAll('PredecessorLink')).map(pl => ({
        predUID: parseInt(getText(pl, 'PredecessorUID') || '0'),
        type:    parseInt(getText(pl, 'Type')           || '1'),
        lag:     parseInt(getText(pl, 'LinkLag')        || '0'),
      })).filter(l => l.predUID > 0);

      return {
        uid,
        id:            parseInt(getText(t, 'ID') || '0'),
        name:          getText(t, 'Name'),
        start:         getText(t, 'Start'),
        finish:        getText(t, 'Finish'),
        duration:      getText(t, 'Duration'),
        pct:           parseFloat(getText(t, 'PercentComplete') || '0'),
        outline:       parseInt(getText(t, 'OutlineLevel') || '1'),
        outlineNumber: getText(t, 'OutlineNumber') || '',  // e.g. "1.2.3"
        isSummary:     getText(t, 'Summary')   === '1',
        isMilestone:   getText(t, 'Milestone') === '1',
        isActive:      getText(t, 'Active') !== '0',
        links,
        resources:     (taskResMap[uid] || []).join('; '),
        predStr:       '',
        succStr:       '',
      };
    })
    .filter(t => !(t.id === 0 && !t.name));

  // Compute parentUID using OutlineNumber ("1.2.3") when available — this is
  // order-independent and unambiguous. Fall back to pStack for files that omit it.
  {
    const numToUID = {};
    for (const task of tasks) {
      if (task.outlineNumber) numToUID[task.outlineNumber] = task.uid;
    }
    const pStack = [];
    for (const task of tasks) {
      if (task.outlineNumber) {
        const dot = task.outlineNumber.lastIndexOf('.');
        task.parentUID = dot > 0
          ? (numToUID[task.outlineNumber.slice(0, dot)] ?? null)
          : null;
      } else {
        while (pStack.length > 0 && pStack[pStack.length - 1].outline >= task.outline) pStack.pop();
        task.parentUID = pStack.length > 0 ? pStack[pStack.length - 1].uid : null;
      }
      // Keep pStack in sync for the fallback path
      while (pStack.length > 0 && pStack[pStack.length - 1].outline >= task.outline) pStack.pop();
      pStack.push({ uid: task.uid, outline: task.outline });
    }
  }

  return { projName, projStart, projFinish, tasks };
}

// ─── Collapse filter ──────────────────────────────────────────────────────────
// Uses stored parentUID references so it's immune to filter gaps or level jumps.
function collapseFilter(tasks) {
  if (state.collapsedSummaries.size === 0) return tasks;
  return tasks.filter(task => {
    let uid = task.parentUID;
    while (uid != null) {
      if (state.collapsedSummaries.has(uid)) return false;
      const p = state.uidToTask[uid];
      uid = p ? p.parentUID : null;
    }
    return true;
  });
}

// ─── Filter logic ─────────────────────────────────────────────────────────────
function applyFilters() {
  const { filterText, filterStatus, filterType } = state;
  const q = filterText.trim().toLowerCase();

  const baseFiltered = state.allTasks.filter(t => {
    if (q && !t.name.toLowerCase().includes(q) && !t.resources.toLowerCase().includes(q)) return false;
    switch (filterStatus) {
      case 'not-started': if (t.pct !== 0) return false; break;
      case 'in-progress': if (t.pct === 0 || t.pct >= 100) return false; break;
      case 'complete':    if (t.pct < 100) return false; break;
    }
    switch (filterType) {
      case 'normal':    if (t.isSummary || t.isMilestone) return false; break;
      case 'summary':   if (!t.isSummary) return false; break;
      case 'milestone': if (!t.isMilestone) return false; break;
    }
    return true;
  });

  state.filteredTasks = collapseFilter(baseFiltered);

  const total = state.allTasks.length;
  const shown = state.filteredTasks.length;
  document.getElementById('filter-count').textContent =
    shown === total ? `${total} tasks` : `${shown} / ${total}`;

  renderTable(state.filteredTasks);
  redrawGantt();
}

// ─── Links modal ─────────────────────────────────────────────────────────────
const TYPE_NAMES = ['FF', 'FS', 'SF', 'SS'];

function showLinksModal(taskUID, type) {
  const task = state.uidToTask[taskUID];
  if (!task) return;

  let items;
  if (type === 'pred') {
    items = task.links.map(l => {
      const t = state.uidToTask[l.predUID];
      return { id: t?.id ?? l.predUID, name: t?.name ?? `(UID ${l.predUID})`, rel: TYPE_NAMES[l.type] ?? 'FS' };
    });
  } else {
    items = state.allTasks
      .filter(t => t.links.some(l => l.predUID === taskUID))
      .map(t => {
        const link = t.links.find(l => l.predUID === taskUID);
        return { id: t.id, name: t.name, rel: TYPE_NAMES[link?.type] ?? 'FS' };
      });
  }

  document.getElementById('links-modal-title').textContent =
    (type === 'pred' ? 'Predecessors' : 'Successors') + ' — ' + task.name;

  const listEl = document.getElementById('links-modal-list');
  listEl.innerHTML = items.length
    ? items.map(it =>
        `<button class="lm-item" data-id="${it.id}">` +
        `<span class="lm-id">${it.id}</span>` +
        `<span class="lm-rel lm-rel-${it.rel.toLowerCase()}">${it.rel}</span>` +
        `<span class="lm-name">${escapeHtml(it.name)}</span>` +
        `</button>`).join('')
    : `<div class="lm-empty">No ${type === 'pred' ? 'predecessors' : 'successors'}</div>`;

  document.getElementById('links-modal').classList.remove('hidden');
}

function hideLinksModal() {
  document.getElementById('links-modal').classList.add('hidden');
}

document.getElementById('links-modal').addEventListener('click', e => {
  if (e.target.closest('.links-modal-scrim')) { hideLinksModal(); return; }
  const item = e.target.closest('.lm-item');
  if (item) { navigateToTaskId(parseInt(item.dataset.id)); hideLinksModal(); }
});
document.getElementById('links-modal-close').addEventListener('click', hideLinksModal);

// ─── Link cell helper ─────────────────────────────────────────────────────────
function linkCell(str, taskUID, type) {
  if (!str) return '';
  return `<div class="link-cell-wrap">` +
    `<span class="link-cell-text">${escapeHtml(str)}</span>` +
    `<button class="link-expand-btn" data-uid="${taskUID}" data-type="${type}" title="Show all ${type === 'pred' ? 'predecessors' : 'successors'}">↗</button>` +
    `</div>`;
}

// ─── Table renderer ───────────────────────────────────────────────────────────
function renderTable(tasks) {
  const tbody = document.getElementById('task-body');
  tbody.innerHTML = '';
  for (const task of tasks) {
    const indent      = Math.max(0, task.outline - 1) * 14;
    const isCollapsed = task.isSummary && state.collapsedSummaries.has(task.uid);
    const chevron     = task.isSummary
      ? `<span class="chevron">${isCollapsed ? '▶' : '▼'}</span>`
      : '<span class="chevron"></span>';

    const tr = document.createElement('tr');
    tr.dataset.uid = task.uid;
    if (task.isSummary)   tr.classList.add('summary');
    if (task.isMilestone) tr.classList.add('milestone');
    if (!task.isActive)   tr.classList.add('inactive');

    tr.innerHTML =
      `<td class="col-id">${task.id}</td>` +
      `<td class="col-uid">${task.uid}</td>` +
      `<td class="col-name"><span class="task-name" style="padding-left:${indent}px">${chevron}${escapeHtml(task.name)}</span></td>` +
      `<td class="col-dur">${formatDuration(task.duration)}</td>` +
      `<td class="col-start">${formatDate(task.start)}</td>` +
      `<td class="col-finish">${formatDate(task.finish)}</td>` +
      `<td class="col-pct">${task.pct}%</td>` +
      `<td class="col-pred">${linkCell(task.predStr, task.uid, 'pred')}</td>` +
      `<td class="col-succ">${linkCell(task.succStr, task.uid, 'succ')}</td>` +
      `<td class="col-res">${escapeHtml(task.resources)}</td>`;

    tbody.appendChild(tr);
  }
}

// ─── Navigate to task by ID ───────────────────────────────────────────────────
function navigateToTaskId(id) {
  const idx = state.filteredTasks.findIndex(t => t.id === id);
  if (idx === -1) return;

  const offset = Math.max(0, idx * ROW_H - ROW_H * 3);
  document.getElementById('gantt-bars-wrap').scrollTop = offset;
  document.getElementById('ms-left').scrollTop         = offset;

  const rows = document.querySelectorAll('#task-body tr');
  const tr   = rows[idx];
  if (tr) {
    tr.classList.add('navigate-flash');
    tr.addEventListener('animationend', () => tr.classList.remove('navigate-flash'), { once: true });
  }
}

// ─── Export visible tasks as formatted Excel (HTML/XLS) ───────────────────────
function exportExcel() {
  const cols = ['ID','UID','Task Name','Duration','Start','Finish','% Done',
                'Predecessors','Successors','Resources'];

  const hdrStyle = 'background:#1E3A5F;color:#fff;font-weight:bold;border:1px solid #aaa;padding:5px 8px;white-space:nowrap';
  const cellBase = 'border:1px solid #ddd;padding:4px 8px;white-space:nowrap';

  let trs = '<tr>' + cols.map(c => `<th style="${hdrStyle}">${c}</th>`).join('') + '</tr>\n';

  for (const t of state.filteredTasks) {
    const indent  = Math.max(0, t.outline - 1);
    const pad     = '\u00A0'.repeat(indent * 3);   // non-breaking spaces for indent
    const bg      = t.isSummary  ? '#FFF8E1'
                  : t.isMilestone ? '#EEF4FF'
                  : '#FFFFFF';
    const fw      = t.isSummary ? 'bold' : 'normal';
    const opacity = t.isActive  ? '' : ';color:#999;text-decoration:line-through';

    const rowStyle = `background:${bg}${opacity}`;
    // x:str attribute forces Excel to treat the cell as text, preventing numeric reinterpretation
    const td = (v, extra = '', forceText = false) => {
      const str = String(v ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      return `<td${forceText ? ' x:str' : ''} style="${cellBase};${rowStyle}${extra}">${str}</td>`;
    };

    trs += '<tr>' +
      td(t.id,                        ';text-align:right;color:#888',          true) +
      td(t.uid,                       ';text-align:right;color:#888',          true) +
      td(pad + t.name,                `;font-weight:${fw}`) +
      td(formatDuration(t.duration)) +
      td(formatDate(t.start),         ';color:#555') +
      td(formatDate(t.finish),        ';color:#555') +
      td(t.pct + '%',                 ';text-align:right;color:#555',          true) +
      td(t.predStr,                   ';color:#555;font-family:monospace',     true) +
      td(t.succStr,                   ';color:#555;font-family:monospace',     true) +
      td(t.resources,                 '',                                      true) +
    '</tr>\n';
  }

  const sheetName = (state.projectName || 'Project').slice(0, 31)
    .replace(/[[\]*?:/\\]/g, '_');

  const html = [
    '<html xmlns:o="urn:schemas-microsoft-com:office:office"',
    '      xmlns:x="urn:schemas-microsoft-com:office:excel"',
    '      xmlns="http://www.w3.org/TR/REC-html40">',
    '<head><meta charset="UTF-8">',
    '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>',
    `<x:ExcelWorksheet><x:Name>${sheetName}</x:Name>`,
    '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>',
    '</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->',
    '<style>table{border-collapse:collapse}body{font-family:Calibri,Arial,sans-serif;font-size:11pt}</style>',
    '</head><body>',
    '<table>',
    trs,
    '</table></body></html>',
  ].join('\n');

  const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = (state.projectName || 'project').replace(/[/\\?%*:|"<>]/g, '_') + '.xls';
  a.click();
  URL.revokeObjectURL(a.href);
}

// ─── Gantt: dependency arrow SVG ──────────────────────────────────────────────
function buildLinkPath(fx, fy, tx, ty, type) {
  if (type === 1) {
    if (tx > fx + 2) {
      const mid = Math.round((fx + tx) / 2);
      return `M${fx},${fy} H${mid} V${ty} H${tx}`;
    }
    const dir  = fy <= ty ? 1 : -1;
    const elby = Math.round(fy + dir * ROW_H * 0.65);
    return `M${fx},${fy} H${fx+LINK_M} V${elby} H${tx-LINK_M} V${ty} H${tx}`;
  }
  if (type === 0) {
    const rx = Math.max(fx, tx) + LINK_M;
    return `M${fx},${fy} H${rx} V${ty} H${tx}`;
  }
  if (type === 3) {
    const lx = Math.min(fx, tx) - LINK_M;
    return `M${fx},${fy} H${lx} V${ty} H${tx}`;
  }
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

      let fx, tx;
      switch (link.type) {
        case 0: fx = pR; tx = sR; break;
        case 1: fx = pR; tx = sL; break;
        case 2: fx = pL; tx = sR; break;
        case 3: fx = pL; tx = sL; break;
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

  // ── Timeline header (adaptive granularity) ──
  headerEl.style.width = totalW + 'px';

  if (pxPerDay >= 2) {
    // Month-level header
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
  } else if (pxPerDay >= 0.4) {
    // Quarter-level header
    let cur = new Date(minD.getFullYear(), Math.floor(minD.getMonth() / 3) * 3, 1);
    while (cur <= maxD) {
      const y = cur.getFullYear(), q = Math.floor(cur.getMonth() / 3);
      const endMo = q * 3 + 3;
      const nextQ = new Date(y, endMo, 1);
      const daysInQ = dayDiff(cur, nextQ);
      const cell = document.createElement('div');
      cell.className  = 'gh-month';
      cell.style.width = Math.round(daysInQ * pxPerDay) + 'px';
      cell.textContent = `Q${q + 1} ${y}`;
      headerEl.appendChild(cell);
      cur = nextQ;
    }
  } else {
    // Year-level header
    let cur = new Date(minD.getFullYear(), 0, 1);
    while (cur <= maxD) {
      const y = cur.getFullYear();
      const nextY = new Date(y + 1, 0, 1);
      const daysInY = dayDiff(cur, nextY);
      const cell = document.createElement('div');
      cell.className  = 'gh-month';
      cell.style.width = Math.round(daysInY * pxPerDay) + 'px';
      cell.textContent = String(y);
      headerEl.appendChild(cell);
      cur = nextY;
    }
  }

  // ── Canvas ──
  barsEl.style.width  = totalW + 'px';
  barsEl.style.height = Math.max(totalH, 1) + 'px';

  // Vertical grid lines (adaptive)
  if (pxPerDay >= 2) {
    // Month vlines
    let dayOff = 0;
    let cur = new Date(minD.getFullYear(), minD.getMonth(), 1);
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
  } else if (pxPerDay >= 0.4) {
    // Quarter vlines
    let cur = new Date(minD.getFullYear(), Math.floor(minD.getMonth() / 3) * 3, 1);
    while (cur <= maxD) {
      const y = cur.getFullYear(), q = Math.floor(cur.getMonth() / 3);
      const off = dayDiff(minD, cur);
      if (off > 0) {
        const ln = document.createElement('div');
        ln.className  = 'gantt-vline';
        ln.style.left = Math.round(off * pxPerDay) + 'px';
        barsEl.appendChild(ln);
      }
      cur = new Date(y, q * 3 + 3, 1);
    }
  } else {
    // Year vlines
    let cur = new Date(minD.getFullYear(), 0, 1);
    while (cur <= maxD) {
      const off = dayDiff(minD, cur);
      if (off > 0) {
        const ln = document.createElement('div');
        ln.className  = 'gantt-vline';
        ln.style.left = Math.round(off * pxPerDay) + 'px';
        barsEl.appendChild(ln);
      }
      cur = new Date(cur.getFullYear() + 1, 0, 1);
    }
  }

  // Today line
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (today >= minD && today <= maxD) {
    const tl = document.createElement('div');
    tl.className  = 'gantt-today';
    tl.style.left = Math.round(dayDiff(minD, today) * pxPerDay) + 'px';
    barsEl.appendChild(tl);
  }

  // Build UID→row-index map
  const uidToRow = {};
  tasks.forEach((t, i) => { uidToRow[t.uid] = i; });

  // Bar rows
  tasks.forEach((task, i) => {
    const classes = ['gbar-row'];
    if (task.isSummary) classes.push('summary');
    if (!task.isActive) classes.push('inactive');
    const row = document.createElement('div');
    row.className = classes.join(' ');
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
        if (task.isSummary) {
          // Bracket-style summary bar: thin top strip + downward triangle endpoints via CSS
          bar.style.cssText = `left:${left}px;width:${width}px;top:5px;height:5px`;
        } else {
          bar.style.cssText = `left:${left}px;width:${width}px;top:${BAR_TOP}px;height:${BAR_H}px`;
          if (task.pct > 0) {
            const done = document.createElement('div');
            done.className = 'gantt-done';
            done.style.width = Math.min(100, task.pct) + '%';
            bar.appendChild(done);
          }
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

  document.getElementById('zoom-label').textContent = zoomLabel(pxPerDay);
}

// ─── Scroll sync ──────────────────────────────────────────────────────────────
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
      leftPanel.scrollTop   = barsWrap.scrollTop;
      headerWrap.scrollLeft = barsWrap.scrollLeft;
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

  barsWrap.addEventListener('wheel', e => {
    if (!e.ctrlKey && !e.metaKey) return;
    e.preventDefault();
    const factor = e.deltaY < 0 ? 1.35 : 1 / 1.35;
    state.pxPerDay = Math.max(0.05, Math.min(80, state.pxPerDay * factor));
    redrawGantt();
  }, { passive: false, signal: sig });
}

// ─── Render results ───────────────────────────────────────────────────────────
function renderResults({ projName, projStart, projFinish, tasks }) {
  state.allTasks    = tasks;
  state.projectName = projName || 'project';
  state.uidToTask   = Object.fromEntries(tasks.map(t => [t.uid, t]));

  // Build predecessor display strings
  for (const task of tasks) {
    task.predStr = task.links.map(l => {
      const pred = state.uidToTask[l.predUID];
      const predId = pred ? pred.id : l.predUID;
      const typeName = ['FF','FS','SF','SS'][l.type] ?? 'FS';
      return typeName === 'FS' ? String(predId) : `${predId}${typeName}`;
    }).join(',');
    task.succStr = '';
  }

  // Build successor display strings
  for (const task of tasks) {
    for (const link of task.links) {
      const pred = state.uidToTask[link.predUID];
      if (!pred) continue;
      const typeName = ['FF','FS','SF','SS'][link.type] ?? 'FS';
      const entry = typeName === 'FS' ? String(task.id) : `${task.id}${typeName}`;
      pred.succStr = pred.succStr ? pred.succStr + ',' + entry : entry;
    }
  }

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
    state.pxPerDay = Math.max(0.05, Math.min(40, 900 / totalDays));
  } else {
    state.minD = state.maxD = null;
    state.pxPerDay = 8;
  }

  // Meta bar
  document.getElementById('meta-name').textContent   = projName || '(unnamed)';
  document.getElementById('meta-start').textContent  = formatDate(projStart);
  document.getElementById('meta-finish').textContent = formatDate(projFinish);
  document.getElementById('meta-count').textContent  = tasks.length;

  // Reset filters & collapsed state
  state.filterText   = '';
  state.filterStatus = 'all';
  state.filterType   = 'all';
  state.collapsedSummaries.clear();
  document.getElementById('filter-search').value = '';
  document.getElementById('filter-status').value = 'all';
  document.getElementById('filter-type').value   = 'all';
  state.showLinks = true;
  document.getElementById('btn-links').classList.add('active');

  setupScrollSync();
  applyFilters();
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
        'Binary .mpp files cannot be parsed in the browser — Microsoft has never\n' +
        'published the .mpp format specification.\n\n' +
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

// Outline picker (expand/collapse all dropdown)
const outlinePicker = document.getElementById('outline-picker');
document.getElementById('btn-outline').addEventListener('click', e => {
  e.stopPropagation();
  outlinePicker.style.display = outlinePicker.style.display === 'none' ? 'flex' : 'none';
});
document.getElementById('btn-expand-all').addEventListener('click', () => {
  outlinePicker.style.display = 'none';
  state.collapsedSummaries.clear();
  applyFilters();
});
document.getElementById('btn-collapse-all').addEventListener('click', () => {
  outlinePicker.style.display = 'none';
  state.allTasks.forEach(t => { if (t.isSummary) state.collapsedSummaries.add(t.uid); });
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
  state.pxPerDay = Math.max(0.05, state.pxPerDay / 1.5); redrawGantt();
});
document.getElementById('btn-zoom-fit').addEventListener('click', () => {
  const w = document.getElementById('gantt-bars-wrap').clientWidth || 900;
  if (state.minD && state.maxD) {
    state.pxPerDay = w / (dayDiff(state.minD, state.maxD) + 1);
    redrawGantt();
  }
});

// Task table click — expand button opens modal; summary row toggles collapse
document.getElementById('task-body').addEventListener('click', e => {
  const expandBtn = e.target.closest('.link-expand-btn');
  if (expandBtn) {
    e.stopPropagation();
    showLinksModal(parseInt(expandBtn.dataset.uid), expandBtn.dataset.type);
    return;
  }
  // Collapse/expand summary rows
  const tr = e.target.closest('tr.summary');
  if (!tr) return;
  const uid = parseInt(tr.dataset.uid);
  if (isNaN(uid)) return;
  if (state.collapsedSummaries.has(uid)) {
    state.collapsedSummaries.delete(uid);
  } else {
    state.collapsedSummaries.add(uid);
  }
  applyFilters();
});

// Column picker
const colPicker = document.getElementById('col-picker');
document.getElementById('btn-cols').addEventListener('click', e => {
  e.stopPropagation();
  colPicker.classList.toggle('hidden');
});
document.addEventListener('click', e => {
  if (!e.target.closest('.col-picker-wrap')) colPicker.classList.add('hidden');
  if (!e.target.closest('#outline-picker-wrap')) outlinePicker.style.display = 'none';
});
colPicker.addEventListener('change', e => {
  const col = e.target.dataset.col;
  if (!col) return;
  document.getElementById('task-table').classList.toggle(`hide-${col}`, !e.target.checked);
});

// Light/dark theme toggle — default is light (set on <html> in index.html)
(function () {
  const btn = document.getElementById('btn-theme');
  // Restore explicit dark preference from previous session
  if (localStorage.getItem('mpp-theme') === 'dark') {
    document.documentElement.dataset.theme = 'dark';
    btn.textContent = '☀';
  } else {
    document.documentElement.dataset.theme = 'light';
    btn.textContent = '☾';
  }
  btn.addEventListener('click', () => {
    const isDark = document.documentElement.dataset.theme === 'dark';
    if (isDark) {
      document.documentElement.dataset.theme = 'light';
      btn.textContent = '☾';
      localStorage.removeItem('mpp-theme');
    } else {
      document.documentElement.dataset.theme = 'dark';
      btn.textContent = '☀';
      localStorage.setItem('mpp-theme', 'dark');
    }
  });
})();

// Export CSV button
document.getElementById('btn-export').addEventListener('click', exportExcel);

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
    msLeft.style.width = Math.max(160, Math.min(window.innerWidth - 300, startW + e.clientX - startX)) + 'px';
  });
  document.addEventListener('mouseup', () => {
    if (!dragging) return; dragging = false;
    splitter.classList.remove('dragging');
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
  });
})();
