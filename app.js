/**
 * Crew Planner PWA
 * Analyse des plannings équipage (format Crew Duties XLSX)
 */

(function () {
  'use strict';

  const STORAGE_KEY = 'crewPlannerData_v1';
  const THEME_KEY = 'crewPlannerTheme';

  const KNOWN_CODES = [
    'OFF', 'H', 'SICK', 'MED', 'CVR', 'Rest Period', 'Days Off', 'Vacation',
    'O', 'ON', 'AV', 'STBA', 'STBH', 'STBO', 'SIM', 'T', '36SDO', 'EXT',
    'Request', 'TO', 'Office', 'SDO', 'M'
  ];

  const DEFAULT_NONWORKED = new Set([
    'OFF', 'H', 'SICK', 'MED', 'CVR', 'Rest Period', 'Days Off', 'Vacation'
  ]);

  const HEAT_COLORS = {
    nonworked: '#f59e0b',
    flight: '#22c55e',
    office: '#3b82f6',
    training: '#8b5cf6',
    other: '#64748b',
    empty: 'transparent'
  };

  let records = [];
  let filesMeta = [];
  let nonWorkedCodes = new Set(DEFAULT_NONWORKED);
  let charts = { pie: null, bar: null };
  let pendingConflicts = [];
  let pendingNewRecords = [];
  let pendingFileMeta = null;

  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  const el = {
    empty: $('#empty-state'),
    dashboard: $('#dashboard'),
    fileList: $('#file-list'),
    statsGrid: $('#stats-grid'),
    ratioNote: $('#ratio-note'),
    tableBody: $('#table-body'),
    tableCount: $('#table-count'),
    filterPerson: $('#filter-person'),
    filterFrom: $('#filter-from'),
    filterTo: $('#filter-to'),
    filterCode: $('#filter-code'),
    ratioMode: $('#ratio-mode'),
    pieMode: $('#pie-mode'),
    pairSelects: $('#pair-selects'),
    pairA: $('#pair-a'),
    pairB: $('#pair-b'),
    pieTitle: $('#pie-title'),
    heatmapPerson: $('#heatmap-person'),
    heatmapMonth: $('#heatmap-month'),
    heatmapGrid: $('#heatmap-grid'),
    heatmapLegend: $('#heatmap-legend'),
    modal: $('#modal-nonworked'),
    checkboxes: $('#nonworked-checkboxes'),
    modalConflict: $('#modal-conflict'),
    conflictList: $('#conflict-list'),
    toast: $('#toast'),
    fileInput: $('#file-input'),
  };

  function toast(msg, duration = 2800) {
    el.toast.textContent = msg;
    el.toast.classList.add('show');
    setTimeout(() => el.toast.classList.remove('show'), duration);
  }

  function parseDateFR(str) {
    if (str instanceof Date) return str;
    if (!str) return null;
    const s = String(str).trim();
    const m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (m) {
      const d = new Date(+m[3], +m[2] - 1, +m[1]);
      return isNaN(d) ? null : d;
    }
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }

  function formatDate(d) {
    if (!d) return '';
    return d.toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }

  function isoDate(d) {
    if (!d) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
  }

  function daysInYear(year) {
    return ((year % 4 === 0 && year % 100 !== 0) || year % 400 === 0) ? 366 : 365;
  }

  function daysBetween(a, b) {
    return Math.round(Math.abs(b - a) / 86400000) + 1;
  }

  function saveState() {
    try {
      const payload = {
        records: records.map((r) => ({ ...r, date: isoDate(r.date) })),
        filesMeta,
        nonWorked: [...nonWorkedCodes]
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (e) {
      console.warn('Save failed', e);
    }
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const data = JSON.parse(raw);
      records = (data.records || []).map((r) => ({
        ...r,
        date: parseDateFR(r.date)
      }));
      filesMeta = data.filesMeta || [];
      if (Array.isArray(data.nonWorked)) {
        nonWorkedCodes = new Set(data.nonWorked);
      }
    } catch (e) {
      console.warn('Load failed', e);
    }
  }

  function applyTheme(theme) {
    document.body.dataset.theme = theme;
    $$('.theme-btn').forEach((b) => b.classList.toggle('active', b.dataset.theme === theme));
    localStorage.setItem(THEME_KEY, theme);
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) {
      meta.content = theme === 'light' ? '#f8fafc' : theme === 'high-contrast' ? '#000000' : '#0f172a';
    }
    if (records.length) renderStatsAndCharts();
  }

  function parseWorkbook(wb, fileName) {
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const maxRow = range.e.r;
    const maxCol = range.e.c;

    const newRecords = [];
    let currentRole = '';
    let dateRow = null;

    for (let r = 0; r <= maxRow; r++) {
      const cellA = ws[XLSX.utils.encode_cell({ r: r, c: 0 })];
      const valA = cellA ? String(cellA.v || '').trim() : '';

      if (/cockpit/i.test(valA)) {
        currentRole = /FO/i.test(valA) ? 'FO' : (/CPT/i.test(valA) ? 'CPT' : valA);
        dateRow = [];
        for (let c = 1; c <= maxCol; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
          dateRow[c] = parseDateFR(cell ? cell.v : null);
        }
        continue;
      }

      if (valA && dateRow && !/cockpit|av\s*-|stbh|sim\s*-|36sdo|office|sick|rest period|days off|vacation|request|cvr|ext\s*-/i.test(valA)) {
        const person = valA;

        for (let c = 1; c <= maxCol; c++) {
          const d = dateRow[c];
          if (!d) continue;

          const cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
          let code = cell && cell.v != null ? String(cell.v).trim() : '';
          if (code === '') code = null;

          let note = null;
          const noteCell = ws[XLSX.utils.encode_cell({ r: r + 1, c: c })];
          if (noteCell && noteCell.v != null) {
            const n = String(noteCell.v).trim();
            const nextA = ws[XLSX.utils.encode_cell({ r: r + 1, c: 0 })];
            const nextAVal = nextA ? String(nextA.v || '').trim() : '';
            if (!nextAVal) note = n;
          }

          newRecords.push({
            id: person + '|' + isoDate(d) + '|' + fileName,
            person: person,
            role: currentRole || '—',
            date: d,
            code: code,
            note: note,
            source: fileName
          });
        }
      }
    }
    return newRecords;
  }

  async function handleFiles(fileList) {
    if (!fileList || !fileList.length) return;

    for (const file of fileList) {
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array', cellDates: false });
        const parsed = parseWorkbook(wb, file.name);

        if (parsed.length === 0) {
          toast('Aucun planning détecté dans « ' + file.name + ' »');
          continue;
        }

        const existingMap = new Map();
        records.forEach((r) => existingMap.set(r.person + '|' + isoDate(r.date), r));

        const conflicts = [];
        const cleanNew = [];

        parsed.forEach((r) => {
          const key = r.person + '|' + isoDate(r.date);
          const existing = existingMap.get(key);
          if (existing) {
            const same =
              (existing.code || '') === (r.code || '') &&
              (existing.note || '') === (r.note || '');
            if (!same) {
              conflicts.push({ key: key, existing: existing, incoming: r });
            } else {
              cleanNew.push(r);
            }
          } else {
            cleanNew.push(r);
          }
        });

        const dates = parsed.map((r) => r.date).filter(Boolean).sort((a, b) => a - b);
        const period = dates.length
          ? formatDate(dates[0]) + ' → ' + formatDate(dates[dates.length - 1])
          : '—';
        pendingFileMeta = { name: file.name, count: parsed.length, period: period };

        if (conflicts.length > 0) {
          pendingConflicts = conflicts;
          pendingNewRecords = cleanNew;
          openConflictModal(conflicts);
          return;
        }

        applyImport(cleanNew, pendingFileMeta);
      } catch (err) {
        console.error(err);
        toast('Erreur lecture « ' + file.name + ' »');
      }
    }
  }

  function applyImport(newRecs, meta) {
    const byKey = new Map();
    records.forEach((r) => byKey.set(r.person + '|' + isoDate(r.date), r));
    newRecs.forEach((r) => byKey.set(r.person + '|' + isoDate(r.date), r));
    records = Array.from(byKey.values());

    if (meta) {
      filesMeta = filesMeta.filter((f) => f.name !== meta.name);
      filesMeta.push(meta);
    }
    saveState();
    refreshUI();
    toast(newRecs.length + ' lignes importées / consolidées');
  }

  function openConflictModal(conflicts) {
    el.conflictList.innerHTML = conflicts
      .map((c, i) => {
        const d = formatDate(c.existing.date);
        return (
          '<div class="conflict-item" data-idx="' + i + '">' +
          '<div class="conflict-meta">' +
          '<strong>' + c.existing.person + '</strong> · ' + d + '<br>' +
          '<span style="color:var(--text-muted)">Existant : ' + (c.existing.code || '(vide)') +
          (c.existing.note ? ' (' + c.existing.note + ')' : '') + ' ← ' + c.existing.source + '</span><br>' +
          '<span style="color:var(--text-muted)">Nouveau : ' + (c.incoming.code || '(vide)') +
          (c.incoming.note ? ' (' + c.incoming.note + ')' : '') + ' ← ' + c.incoming.source + '</span>' +
          '</div>' +
          '<select class="conflict-choice" data-idx="' + i + '">' +
          '<option value="old">Garder l\'existant</option>' +
          '<option value="new" selected>Garder le nouveau</option>' +
          '</select></div>'
        );
      })
      .join('');
    el.modalConflict.classList.add('open');
  }

  function closeConflictModal() {
    el.modalConflict.classList.remove('open');
    pendingConflicts = [];
    pendingNewRecords = [];
    pendingFileMeta = null;
  }

  function resolveConflicts(globalChoice) {
    const choices = new Map();
    if (globalChoice) {
      pendingConflicts.forEach(function (_, i) { choices.set(i, globalChoice); });
    } else {
      el.conflictList.querySelectorAll('.conflict-choice').forEach(function (sel) {
        choices.set(+sel.dataset.idx, sel.value);
      });
    }

    const toAdd = pendingNewRecords.slice();
    pendingConflicts.forEach(function (c, i) {
      const choice = choices.get(i) || 'new';
      if (choice === 'new') toAdd.push(c.incoming);
    });

    applyImport(toAdd, pendingFileMeta);
    closeConflictModal();
  }

  function getFiltered() {
    let list = records.slice();

    const person = el.filterPerson.value;
    if (person) list = list.filter(function (r) { return r.person === person; });

    const from = el.filterFrom.value ? new Date(el.filterFrom.value + 'T00:00:00') : null;
    const to = el.filterTo.value ? new Date(el.filterTo.value + 'T23:59:59') : null;
    if (from) list = list.filter(function (r) { return r.date >= from; });
    if (to) list = list.filter(function (r) { return r.date <= to; });

    const code = el.filterCode.value;
    if (code === '__empty__') list = list.filter(function (r) { return !r.code; });
    else if (code) list = list.filter(function (r) { return r.code === code; });

    return list.sort(function (a, b) {
      return a.date - b.date || a.person.localeCompare(b.person);
    });
  }

  function isNonWorked(rec) {
    if (!rec.code) return false;
    return nonWorkedCodes.has(rec.code);
  }

  function dutyCategory(rec) {
    if (!rec.code) {
      if (rec.note && /flight/i.test(rec.note)) return 'flight';
      return 'other';
    }
    if (isNonWorked(rec)) return 'nonworked';
    if (/^(O|ON|Office)$/i.test(rec.code)) return 'office';
    if (/^(T|TO|SIM|Training)/i.test(rec.code)) return 'training';
    if (/flight/i.test(rec.note || '')) return 'flight';
    return 'other';
  }

  function computeStats(list) {
    if (!list.length) {
      return {
        totalDays: 0, nonWorked: 0, worked: 0,
        ratioPeriod: 0, ratioYear: 0,
        byPerson: {}, byCode: {}, years: [],
        minDate: null, maxDate: null, periodDays: 0
      };
    }

    const uniqueMap = new Map();
    list.forEach(function (r) {
      uniqueMap.set(r.person + '|' + isoDate(r.date), r);
    });
    const uniqueList = Array.from(uniqueMap.values());
    const totalDays = uniqueList.length;

    let nonWorked = 0;
    const byPerson = {};
    const byCode = {};
    const years = new Set();
    let minDate = uniqueList[0].date;
    let maxDate = uniqueList[0].date;

    uniqueList.forEach(function (r) {
      years.add(r.date.getFullYear());
      if (r.date < minDate) minDate = r.date;
      if (r.date > maxDate) maxDate = r.date;

      const codeKey = r.code || '(vide)';
      byCode[codeKey] = (byCode[codeKey] || 0) + 1;

      if (!byPerson[r.person]) byPerson[r.person] = { total: 0, nonWorked: 0 };
      byPerson[r.person].total += 1;
      if (isNonWorked(r)) {
        byPerson[r.person].nonWorked += 1;
        nonWorked += 1;
      }
    });

    const worked = totalDays - nonWorked;
    const ratioPeriod = totalDays ? (nonWorked / totalDays) * 100 : 0;

    let yearDays = 365;
    if (years.size === 1) yearDays = daysInYear([...years][0]);
    const ratioYear = yearDays ? (nonWorked / yearDays) * 100 : 0;

    return {
      totalDays: totalDays,
      nonWorked: nonWorked,
      worked: worked,
      ratioPeriod: ratioPeriod,
      ratioYear: ratioYear,
      byPerson: byPerson,
      byCode: byCode,
      years: [...years],
      minDate: minDate,
      maxDate: maxDate,
      periodDays: daysBetween(minDate, maxDate)
    };
  }

  function refreshUI() {
    const hasData = records.length > 0;
    el.empty.hidden = hasData;
    el.dashboard.hidden = !hasData;
    $('#btn-export').disabled = !hasData;
    $('#btn-nonworked').disabled = !hasData;

    if (!hasData) return;

    el.fileList.innerHTML = filesMeta
      .map(function (f) {
        return (
          '<div class="file-item">' +
          '<span title="' + f.name + '">' + f.name + '</span>' +
          '<span style="font-size:0.8rem;color:var(--text-muted)">' + f.count + ' entrées · ' + f.period + '</span>' +
          '</div>'
        );
      })
      .join('');

    const persons = [...new Set(records.map(function (r) { return r.person; }))].sort();
    const currentPerson = el.filterPerson.value;
    el.filterPerson.innerHTML =
      '<option value="">Toutes</option>' +
      persons.map(function (p) { return '<option value="' + p + '">' + p + '</option>'; }).join('');
    if (persons.indexOf(currentPerson) >= 0) el.filterPerson.value = currentPerson;

    const hmPerson = el.heatmapPerson.value;
    el.heatmapPerson.innerHTML =
      '<option value="">Toutes (vue consolidée)</option>' +
      persons.map(function (p) { return '<option value="' + p + '">' + p + '</option>'; }).join('');
    if (persons.indexOf(hmPerson) >= 0) el.heatmapPerson.value = hmPerson;

    el.pairA.innerHTML = persons.map(function (p) { return '<option value="' + p + '">' + p + '</option>'; }).join('');
    el.pairB.innerHTML = persons.map(function (p) { return '<option value="' + p + '">' + p + '</option>'; }).join('');
    if (persons.length >= 2) {
      el.pairA.value = persons[0];
      el.pairB.value = persons[1];
    }

    const codes = [...new Set(records.map(function (r) { return r.code; }).filter(Boolean))].sort();
    const currentCode = el.filterCode.value;
    el.filterCode.innerHTML =
      '<option value="">Tous</option><option value="__empty__">(vide)</option>' +
      codes.map(function (c) { return '<option value="' + c + '">' + c + '</option>'; }).join('');
    if (currentCode === '__empty__' || codes.indexOf(currentCode) >= 0) {
      el.filterCode.value = currentCode;
    }

    const allDates = records.map(function (r) { return r.date; }).filter(Boolean);
    if (allDates.length) {
      const min = new Date(Math.min.apply(null, allDates));
      const max = new Date(Math.max.apply(null, allDates));
      el.filterFrom.min = isoDate(min);
      el.filterFrom.max = isoDate(max);
      el.filterTo.min = isoDate(min);
      el.filterTo.max = isoDate(max);

      const months = new Set();
      allDates.forEach(function (d) {
        months.add(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'));
      });
      const monthList = [...months].sort();
      const curMonth = el.heatmapMonth.value;
      el.heatmapMonth.innerHTML = monthList
        .map(function (m) {
          const parts = m.split('-');
          const label = new Date(+parts[0], +parts[1] - 1, 1).toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' });
          return '<option value="' + m + '">' + label + '</option>';
        })
        .join('');
      if (monthList.indexOf(curMonth) >= 0) el.heatmapMonth.value = curMonth;
      else if (monthList.length) el.heatmapMonth.value = monthList[0];
    }

    renderStatsAndCharts();
    renderTable();
    renderHeatmap();
  }

  function renderStatsAndCharts() {
    const filtered = getFiltered();
    const stats = computeStats(filtered);
    const mode = el.ratioMode.value;

    let mainRatio = stats.ratioPeriod;
    let ratioLabel = 'Ratio non travaillés / période';
    let note = '';

    if (mode === 'year') {
      mainRatio = stats.ratioYear;
      ratioLabel = 'Ratio non travaillés / année';
      note = '* ' + stats.nonWorked + ' jours non travaillés sur ' +
        (stats.years.length === 1 ? daysInYear(stats.years[0]) : 365) +
        " jours de l'année. Données sur " + (stats.periodDays || '?') + ' jours calendaires.';
    } else if (mode === 'projected') {
      mainRatio = stats.ratioPeriod;
      ratioLabel = 'Taux annualisé (période)';
      note = '* Taux observé sur la période (' + stats.ratioPeriod.toFixed(1) +
        ' %). Si ce rythme se maintient sur une année complète, le ratio serait du même ordre.';
    } else {
      note = 'Période couverte : ' +
        (stats.minDate ? formatDate(stats.minDate) : '—') + ' → ' +
        (stats.maxDate ? formatDate(stats.maxDate) : '—') +
        ' (' + (stats.periodDays || 0) + ' jours).';
    }

    el.statsGrid.innerHTML =
      '<div class="stat-box"><div class="stat-value">' + stats.totalDays + '</div><div class="stat-label">Jours person-jour</div></div>' +
      '<div class="stat-box"><div class="stat-value">' + stats.worked + '</div><div class="stat-label">Jours travaillés</div></div>' +
      '<div class="stat-box"><div class="stat-value">' + stats.nonWorked + '</div><div class="stat-label">Jours non travaillés</div></div>' +
      '<div class="stat-box"><div class="stat-value">' + mainRatio.toFixed(1) + '%</div><div class="stat-label">' + ratioLabel + '</div></div>';
    el.ratioNote.textContent = note;

    const pieMode = el.pieMode.value;
    el.pairSelects.hidden = pieMode !== 'nonworked-pair';

    let pieLabels = [];
    let pieData = [];
    let pieTitle = 'Répartition des codes';

    if (pieMode === 'codes') {
      pieLabels = Object.keys(stats.byCode);
      pieData = Object.values(stats.byCode);
      pieTitle = 'Répartition des codes';
    } else if (pieMode === 'nonworked-all') {
      pieLabels = Object.keys(stats.byPerson);
      pieData = pieLabels.map(function (p) { return stats.byPerson[p].nonWorked; });
      pieTitle = 'Jours non travaillés – équipe';
    } else if (pieMode === 'nonworked-pair') {
      const a = el.pairA.value;
      const b = el.pairB.value;
      const na = stats.byPerson[a] ? stats.byPerson[a].nonWorked : 0;
      const nb = stats.byPerson[b] ? stats.byPerson[b].nonWorked : 0;
      pieLabels = [a || 'A', b || 'B'];
      pieData = [na, nb];
      pieTitle = 'Non travaillés : ' + a + ' vs ' + b;
    }

    el.pieTitle.textContent = pieTitle;
    const colors = generateColors(pieLabels.length);
    const textColor = getComputedStyle(document.body).getPropertyValue('--text').trim() || '#f1f5f9';
    const mutedColor = getComputedStyle(document.body).getPropertyValue('--text-muted').trim() || '#94a3b8';
    const borderColor = getComputedStyle(document.body).getPropertyValue('--bg-card').trim() || '#1e293b';
    const gridColor = getComputedStyle(document.body).getPropertyValue('--border').trim() || '#475569';

    if (charts.pie) charts.pie.destroy();
    const pieCtx = $('#chart-pie').getContext('2d');
    charts.pie = new Chart(pieCtx, {
      type: 'doughnut',
      data: {
        labels: pieLabels,
        datasets: [{
          data: pieData,
          backgroundColor: colors,
          borderWidth: 2,
          borderColor: borderColor
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: { color: textColor, boxWidth: 12, font: { size: 11 } }
          }
        }
      }
    });

    const persons = Object.keys(stats.byPerson);
    const nonWorkedData = persons.map(function (p) { return stats.byPerson[p].nonWorked; });
    const totalData = persons.map(function (p) { return stats.byPerson[p].total; });

    if (charts.bar) charts.bar.destroy();
    const barCtx = $('#chart-bar').getContext('2d');
    charts.bar = new Chart(barCtx, {
      type: 'bar',
      data: {
        labels: persons,
        datasets: [
          {
            label: 'Non travaillés',
            data: nonWorkedData,
            backgroundColor: '#f59e0b',
            borderRadius: 4
          },
          {
            label: 'Total jours',
            data: totalData,
            backgroundColor: '#3b82f6',
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { ticks: { color: mutedColor }, grid: { color: 'transparent' } },
          y: { beginAtZero: true, ticks: { color: mutedColor }, grid: { color: gridColor } }
        },
        plugins: { legend: { labels: { color: textColor } } }
      }
    });
  }

  function generateColors(n) {
    const base = [
      '#3b82f6', '#22c55e', '#f59e0b', '#ef4444', '#8b5cf6',
      '#06b6d4', '#ec4899', '#84cc16', '#f97316', '#6366f1'
    ];
    const out = [];
    for (let i = 0; i < n; i++) out.push(base[i % base.length]);
    return out;
  }

  function renderTable() {
    const filtered = getFiltered();
    el.tableCount.textContent = '(' + filtered.length + ' ligne' + (filtered.length > 1 ? 's' : '') + ')';

    const display = filtered.slice(0, 1000);
    el.tableBody.innerHTML = display
      .map(function (r) {
        const type = isNonWorked(r) ? 'Non travaillé' : 'Travaillé';
        const badgeClass = isNonWorked(r) ? 'badge-off' : 'badge-work';
        const codeLabel = r.code || '(vide)';
        return (
          '<tr>' +
          '<td>' + formatDate(r.date) + '</td>' +
          '<td>' + r.person + '</td>' +
          '<td>' + r.role + '</td>' +
          '<td><span class="badge">' + codeLabel + '</span></td>' +
          '<td>' + (r.note || '—') + '</td>' +
          '<td><span class="badge ' + badgeClass + '">' + type + '</span></td>' +
          '</tr>'
        );
      })
      .join('');

    if (filtered.length > 1000) {
      el.tableBody.innerHTML +=
        '<tr><td colspan="6" style="text-align:center;color:var(--text-muted)">… ' +
        (filtered.length - 1000) + ' lignes non affichées</td></tr>';
    }
  }

  function renderHeatmap() {
    const monthVal = el.heatmapMonth.value;
    if (!monthVal) {
      el.heatmapGrid.innerHTML = '<p style="color:var(--text-muted)">Aucune donnée</p>';
      return;
    }

    const parts = monthVal.split('-');
    const y = +parts[0];
    const mo = +parts[1];
    const personFilter = el.heatmapPerson.value;

    let list = records.filter(function (r) {
      return r.date.getFullYear() === y && r.date.getMonth() + 1 === mo;
    });
    if (personFilter) list = list.filter(function (r) { return r.person === personFilter; });

    el.heatmapLegend.innerHTML =
      '<span><i class="heatmap-swatch" style="background:' + HEAT_COLORS.nonworked + '"></i> Non travaillé</span>' +
      '<span><i class="heatmap-swatch" style="background:' + HEAT_COLORS.flight + '"></i> Vol / flight</span>' +
      '<span><i class="heatmap-swatch" style="background:' + HEAT_COLORS.office + '"></i> Bureau (O)</span>' +
      '<span><i class="heatmap-swatch" style="background:' + HEAT_COLORS.training + '"></i> Training / SIM</span>' +
      '<span><i class="heatmap-swatch" style="background:' + HEAT_COLORS.other + '"></i> Autre travaillé</span>';

    if (personFilter) {
      el.heatmapGrid.innerHTML = buildMonthCalendar(y, mo - 1, list, personFilter);
    } else {
      const persons = [...new Set(list.map(function (r) { return r.person; }))].sort();
      if (!persons.length) {
        el.heatmapGrid.innerHTML = '<p style="color:var(--text-muted)">Aucune donnée pour ce mois</p>';
        return;
      }
      el.heatmapGrid.innerHTML = persons
        .map(function (p) {
          const pl = list.filter(function (r) { return r.person === p; });
          return '<div><h4 style="margin-bottom:0.35rem;font-size:0.9rem">' + p + '</h4>' +
            buildMonthCalendar(y, mo - 1, pl, p) + '</div>';
        })
        .join('');
    }
  }

  function buildMonthCalendar(year, monthIndex, list, person) {
    const first = new Date(year, monthIndex, 1);
    const last = new Date(year, monthIndex + 1, 0);
    const startPad = (first.getDay() + 6) % 7;
    const daysInMonth = last.getDate();

    const byDay = new Map();
    list.forEach(function (r) {
      if (r.person === person || !person) byDay.set(r.date.getDate(), r);
    });

    const weekdays = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'];
    let html = '<div class="heatmap-month"><div class="heatmap-weekdays">' +
      weekdays.map(function (w) { return '<div>' + w + '</div>'; }).join('') +
      '</div><div class="heatmap-days">';

    for (let i = 0; i < startPad; i++) {
      html += '<div class="heatmap-day empty"></div>';
    }

    for (let d = 1; d <= daysInMonth; d++) {
      const rec = byDay.get(d);
      let bg = HEAT_COLORS.other;
      let codeLabel = '';
      let title = '';
      let textColor = '#fff';

      if (rec) {
        const cat = dutyCategory(rec);
        bg = HEAT_COLORS[cat] || HEAT_COLORS.other;
        codeLabel = rec.code || (rec.note === 'flight' ? 'vol' : '·');
        title = formatDate(rec.date) + ' – ' + (rec.code || '(vide)') +
          (rec.note ? ' / ' + rec.note : '');
        textColor = (cat === 'other' || !rec) ? '#fff' : '#fff';
      } else {
        bg = 'var(--bg)';
        title = 'Pas de donnée';
        textColor = 'var(--text-muted)';
      }

      html +=
        '<div class="heatmap-day" style="background:' + bg + ';color:' + textColor + '" title="' + title + '">' +
        '<span class="day-num">' + d + '</span>' +
        (codeLabel ? '<span class="day-code">' + codeLabel + '</span>' : '') +
        '</div>';
    }

    html += '</div></div>';
    return html;
  }

  function openNonWorkedModal() {
    const present = new Set(records.map(function (r) { return r.code; }).filter(Boolean));
    KNOWN_CODES.forEach(function (c) { present.add(c); });
    const sorted = [...present].sort();

    el.checkboxes.innerHTML = sorted
      .map(function (c) {
        return (
          '<label class="checkbox-item">' +
          '<input type="checkbox" value="' + c + '"' +
          (nonWorkedCodes.has(c) ? ' checked' : '') + ' />' +
          '<span>' + c + '</span></label>'
        );
      })
      .join('');

    el.modal.classList.add('open');
  }

  function closeNonWorkedModal() {
    el.modal.classList.remove('open');
  }

  function saveNonWorked() {
    const checked = new Set();
    el.checkboxes.querySelectorAll('input:checked').forEach(function (inp) {
      checked.add(inp.value);
    });
    nonWorkedCodes = checked;
    saveState();
    closeNonWorkedModal();
    renderStatsAndCharts();
    renderTable();
    renderHeatmap();
    toast('Configuration des jours non travaillés enregistrée');
  }

  function exportData() {
    const filtered = getFiltered();
    if (!filtered.length) {
      toast('Aucune donnée à exporter');
      return;
    }

    const header = ['Date', 'Personne', 'Rôle', 'Code', 'Note', 'Non_travaillé', 'Source'];
    const rows = filtered.map(function (r) {
      return [
        isoDate(r.date),
        r.person,
        r.role,
        r.code || '',
        r.note || '',
        isNonWorked(r) ? 'oui' : 'non',
        r.source
      ];
    });

    const csv = [header].concat(rows)
      .map(function (row) {
        return row.map(function (c) {
          return '"' + String(c).replace(/"/g, '""') + '"';
        }).join(';');
      })
      .join('\n');

    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'crew-planner-export-' + isoDate(new Date()) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
    toast('Export CSV téléchargé');
  }

  function initCollapsibles() {
    $$('.collapsible-header').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const targetId = btn.dataset.target;
        const body = document.getElementById(targetId);
        if (!body) return;
        const open = btn.getAttribute('aria-expanded') === 'true';
        btn.setAttribute('aria-expanded', String(!open));
        body.hidden = open;
      });
    });
  }

  function bindEvents() {
    $$('.theme-btn').forEach(function (btn) {
      btn.addEventListener('click', function () { applyTheme(btn.dataset.theme); });
    });

    $('#btn-import').addEventListener('click', function () { el.fileInput.click(); });
    $('#btn-import-empty').addEventListener('click', function () { el.fileInput.click(); });
    el.fileInput.addEventListener('change', function (e) {
      handleFiles(e.target.files);
      e.target.value = '';
    });

    $('#btn-export').addEventListener('click', exportData);

    $('#btn-clear').addEventListener('click', function () {
      if (confirm('Effacer toutes les données chargées ?')) {
        records = [];
        filesMeta = [];
        saveState();
        refreshUI();
        toast('Données effacées');
      }
    });

    $('#btn-nonworked').addEventListener('click', openNonWorkedModal);
    $('#modal-close').addEventListener('click', closeNonWorkedModal);
    $('#modal-cancel').addEventListener('click', closeNonWorkedModal);
    $('#modal-save').addEventListener('click', saveNonWorked);
    el.modal.addEventListener('click', function (e) {
      if (e.target === el.modal) closeNonWorkedModal();
    });

    $('#conflict-close').addEventListener('click', closeConflictModal);
    $('#conflict-keep-old').addEventListener('click', function () { resolveConflicts('old'); });
    $('#conflict-keep-new').addEventListener('click', function () { resolveConflicts('new'); });
    $('#conflict-apply').addEventListener('click', function () { resolveConflicts(null); });
    el.modalConflict.addEventListener('click', function (e) {
      if (e.target === el.modalConflict) closeConflictModal();
    });

    ['filter-person', 'filter-from', 'filter-to', 'filter-code', 'ratio-mode'].forEach(function (id) {
      $('#' + id).addEventListener('change', function () {
        renderStatsAndCharts();
        renderTable();
      });
    });

    el.pieMode.addEventListener('change', function () { renderStatsAndCharts(); });
    el.pairA.addEventListener('change', function () { renderStatsAndCharts(); });
    el.pairB.addEventListener('change', function () { renderStatsAndCharts(); });

    el.heatmapPerson.addEventListener('change', renderHeatmap);
    el.heatmapMonth.addEventListener('change', renderHeatmap);

    $('#btn-reset-filters').addEventListener('click', function () {
      el.filterPerson.value = '';
      el.filterFrom.value = '';
      el.filterTo.value = '';
      el.filterCode.value = '';
      renderStatsAndCharts();
      renderTable();
    });

    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        if (el.modal.classList.contains('open')) closeNonWorkedModal();
        if (el.modalConflict.classList.contains('open')) closeConflictModal();
      }
    });
  }

  function registerSW() {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js').catch(function (err) {
        console.warn('SW failed', err);
      });
    }
  }

  function init() {
    const savedTheme = localStorage.getItem(THEME_KEY) || 'dark';
    applyTheme(savedTheme);
    loadState();
    initCollapsibles();
    bindEvents();
    refreshUI();
    registerSW();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
