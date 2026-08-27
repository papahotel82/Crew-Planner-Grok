/**
 * Crew Planner PWA
 * Analyse des plannings équipage (format Crew Duties XLSX)
 */

(function () {
  'use strict';

  // ---------- Constants & defaults ----------
  const STORAGE_KEY = 'crewPlannerData_v1';
  const THEME_KEY = 'crewPlannerTheme';
  const NONWORKED_KEY = 'crewPlannerNonWorked';

  // Codes connus d'après la légende + ceux rencontrés
  const KNOWN_CODES = [
    'OFF', 'H', 'SICK', 'MED', 'CVR', 'Rest Period', 'Days Off', 'Vacation',
    'O', 'ON', 'AV', 'STBA', 'STBH', 'STBO', 'SIM', 'T', '36SDO', 'EXT',
    'Request', 'TO', 'Office', 'SDO'
  ];

  // Par défaut : non travaillés = repos / congés / maladie
  const DEFAULT_NONWORKED = new Set([
    'OFF', 'H', 'SICK', 'MED', 'CVR', 'Rest Period', 'Days Off', 'Vacation'
  ]);

  // ---------- State ----------
  let records = [];          // { id, person, role, date, code, note, source }
  let filesMeta = [];        // { name, count, period }
  let nonWorkedCodes = new Set(DEFAULT_NONWORKED);
  let charts = { pie: null, bar: null };

  // ---------- DOM refs ----------
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  const el = {
    empty: $('#empty-state'),
    dashboard: $('#dashboard'),
    fileList: $('#file-list'),
    statsGrid: $('#stats-grid'),
    tableBody: $('#table-body'),
    tableCount: $('#table-count'),
    filterPerson: $('#filter-person'),
    filterFrom: $('#filter-from'),
    filterTo: $('#filter-to'),
    filterCode: $('#filter-code'),
    modal: $('#modal-nonworked'),
    checkboxes: $('#nonworked-checkboxes'),
    toast: $('#toast'),
    fileInput: $('#file-input'),
  };

  // ---------- Utilities ----------
  function toast(msg, duration = 2800) {
    el.toast.textContent = msg;
    el.toast.classList.add('show');
    setTimeout(() => el.toast.classList.remove('show'), duration);
  }

  function parseDateFR(str) {
    // "01-01-2026" or Date object
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
    return d.toISOString().slice(0, 10);
  }

  function daysInYear(year) {
    return ((year % 4 === 0 && year % 100 !== 0) || year % 400 === 0) ? 366 : 365;
  }

  // ---------- Persistence (localStorage) ----------
  function saveState() {
    try {
      const payload = {
        records: records.map((r) => ({
          ...r,
          date: isoDate(r.date)
        })),
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

  // ---------- Theme ----------
  function applyTheme(theme) {
    document.body.dataset.theme = theme;
    $$('.theme-btn').forEach((b) => {
      b.classList.toggle('active', b.dataset.theme === theme);
    });
    localStorage.setItem(THEME_KEY, theme);
    // Update theme-color meta
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) {
      meta.content = theme === 'light' ? '#f8fafc' : theme === 'high-contrast' ? '#000000' : '#0f172a';
    }
  }

  // ---------- Parser XLSX (format Crew Duties) ----------
  /**
   * Parse un workbook SheetJS pour extraire les duties.
   * Structure attendue :
   * - Lignes d'en-tête contenant "Cockpit ..." + dates en colonnes
   * - Lignes personnes : nom en col A, codes dans les colonnes dates
   * - Lignes notes (souvent "flight") juste en dessous
   */
  function parseWorkbook(wb, fileName) {
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const maxRow = range.e.r;
    const maxCol = range.e.c;

    const newRecords = [];
    let currentRole = '';
    let dateRow = null; // array of Date|null indexed by col
    let lastPersonRow = -1;

    for (let r = 0; r <= maxRow; r++) {
      const cellA = ws[XLSX.utils.encode_cell({ r, c: 0 })];
      const valA = cellA ? String(cellA.v || '').trim() : '';

      // Détection en-tête de section (Cockpit CPT / FO)
      if (/cockpit/i.test(valA)) {
        currentRole = /FO/i.test(valA) ? 'FO' : (/CPT/i.test(valA) ? 'CPT' : valA);
        // Lire les dates sur cette ligne
        dateRow = [];
        for (let c = 1; c <= maxCol; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          const raw = cell ? cell.v : null;
          dateRow[c] = parseDateFR(raw);
        }
        lastPersonRow = -1;
        continue;
      }

      // Ligne personne (nom non vide + on a un header de dates)
      if (valA && dateRow && !/cockpit|av\s*-|stbh|sim\s*-|36sdo|office|sick|rest period|days off/i.test(valA)) {
        // C'est probablement un nom de personne
        const person = valA;
        lastPersonRow = r;

        for (let c = 1; c <= maxCol; c++) {
          const d = dateRow[c];
          if (!d) continue;

          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          let code = cell && cell.v != null ? String(cell.v).trim() : '';
          if (code === '') code = null; // vide = travaillé

          // Note éventuelle sur la ligne suivante
          let note = null;
          const noteCell = ws[XLSX.utils.encode_cell({ r: r + 1, c })];
          if (noteCell && noteCell.v != null) {
            const n = String(noteCell.v).trim();
            // Si la ligne suivante n'a pas de nom en A, c'est une note
            const nextA = ws[XLSX.utils.encode_cell({ r: r + 1, c: 0 })];
            const nextAVal = nextA ? String(nextA.v || '').trim() : '';
            if (!nextAVal || nextAVal === '') {
              note = n;
            }
          }

          newRecords.push({
            id: `${person}|${isoDate(d)}|${fileName}`,
            person,
            role: currentRole || '—',
            date: d,
            code,
            note,
            source: fileName
          });
        }
      }
    }

    return newRecords;
  }

  // ---------- Import ----------
  async function handleFiles(fileList) {
    if (!fileList || !fileList.length) return;

    let added = 0;
    for (const file of fileList) {
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array', cellDates: false });
        const parsed = parseWorkbook(wb, file.name);

        if (parsed.length === 0) {
          toast(`Aucun planning détecté dans « ${file.name} »`);
          continue;
        }

        // Consolidation : on ajoute / remplace par id (person+date+source)
        // Pour consolidation multi-fichiers : on garde toutes les sources,
        // mais on évite les doublons exacts person+date en privilégiant le dernier fichier.
        const byKey = new Map();
        records.forEach((r) => byKey.set(`${r.person}|${isoDate(r.date)}`, r));
        parsed.forEach((r) => {
          const key = `${r.person}|${isoDate(r.date)}`;
          byKey.set(key, r); // dernier gagne
        });
        records = Array.from(byKey.values());

        // Meta
        const dates = parsed.map((r) => r.date).filter(Boolean).sort((a, b) => a - b);
        const period = dates.length
          ? `${formatDate(dates[0])} → ${formatDate(dates[dates.length - 1])}`
          : '—';
        filesMeta = filesMeta.filter((f) => f.name !== file.name);
        filesMeta.push({ name: file.name, count: parsed.length, period });
        added += parsed.length;
      } catch (err) {
        console.error(err);
        toast(`Erreur lecture « ${file.name} »`);
      }
    }

    if (added > 0) {
      saveState();
      refreshUI();
      toast(`${added} lignes importées / consolidées`);
    }
  }

  // ---------- Filters & computed ----------
  function getFiltered() {
    let list = records.slice();

    const person = el.filterPerson.value;
    if (person) list = list.filter((r) => r.person === person);

    const from = el.filterFrom.value ? new Date(el.filterFrom.value) : null;
    const to = el.filterTo.value ? new Date(el.filterTo.value) : null;
    if (from) {
      from.setHours(0, 0, 0, 0);
      list = list.filter((r) => r.date >= from);
    }
    if (to) {
      to.setHours(23, 59, 59, 999);
      list = list.filter((r) => r.date <= to);
    }

    const code = el.filterCode.value;
    if (code === '__empty__') {
      list = list.filter((r) => !r.code);
    } else if (code) {
      list = list.filter((r) => r.code === code);
    }

    return list.sort((a, b) => a.date - b.date || a.person.localeCompare(b.person));
  }

  function isNonWorked(rec) {
    if (!rec.code) return false; // vide = travaillé
    return nonWorkedCodes.has(rec.code);
  }

  function computeStats(list) {
    if (!list.length) {
      return {
        totalDays: 0,
        nonWorked: 0,
        worked: 0,
        ratioPeriod: 0,
        ratioYear: 0,
        byPerson: {},
        byCode: {},
        years: []
      };
    }

    // Unique person-days (un jour par personne)
    const uniqueKeys = new Set(list.map((r) => `${r.person}|${isoDate(r.date)}`));
    const totalDays = uniqueKeys.size;

    let nonWorked = 0;
    const byPerson = {};
    const byCode = {};
    const years = new Set();

    list.forEach((r) => {
      const key = `${r.person}|${isoDate(r.date)}`;
      // on compte une seule fois par person-day
      // (déjà unique dans list si pas de doublons, mais on sécurise)
      years.add(r.date.getFullYear());

      const codeKey = r.code || '(vide / vol)';
      byCode[codeKey] = (byCode[codeKey] || 0) + 1;

      if (!byPerson[r.person]) {
        byPerson[r.person] = { total: 0, nonWorked: 0 };
      }
      byPerson[r.person].total += 1;
      if (isNonWorked(r)) {
        byPerson[r.person].nonWorked += 1;
        nonWorked += 1;
      }
    });

    // Correction : nonWorked et total doivent être sur uniques
    // Pour simplifier, on considère que list n'a pas de doublons person+date après consolidation
    const worked = totalDays - nonWorked;
    const ratioPeriod = totalDays ? (nonWorked / totalDays) * 100 : 0;

    // Ratio vs année : on prend l'année dominante ou la moyenne
    let yearDays = 365;
    if (years.size === 1) {
      yearDays = daysInYear([...years][0]);
    }
    // Ratio = non-travaillés / jours de l'année (pour la période couverte, indication)
    // Plus pertinent : non-travaillés / total jours de la période
    // On affiche aussi une projection annuelle si période < année
    const ratioYear = yearDays ? (nonWorked / yearDays) * 100 : 0;

    return {
      totalDays,
      nonWorked,
      worked,
      ratioPeriod,
      ratioYear,
      byPerson,
      byCode,
      years: [...years]
    };
  }

  // ---------- Render ----------
  function refreshUI() {
    const hasData = records.length > 0;
    el.empty.hidden = hasData;
    el.dashboard.hidden = !hasData;
    $('#btn-export').disabled = !hasData;
    $('#btn-nonworked').disabled = !hasData;

    if (!hasData) return;

    // File list
    el.fileList.innerHTML = filesMeta
      .map(
        (f) => `
      <div class="file-item">
        <span title="${f.name}">${f.name}</span>
        <span style="font-size:0.8rem;color:var(--text-muted)">${f.count} entrées · ${f.period}</span>
      </div>`
      )
      .join('');

    // Populate filters
    const persons = [...new Set(records.map((r) => r.person))].sort();
    const currentPerson = el.filterPerson.value;
    el.filterPerson.innerHTML =
      '<option value="">Toutes</option>' +
      persons.map((p) => `<option value="${p}">${p}</option>`).join('');
    if (persons.includes(currentPerson)) el.filterPerson.value = currentPerson;

    const codes = [...new Set(records.map((r) => r.code).filter(Boolean))].sort();
    const currentCode = el.filterCode.value;
    el.filterCode.innerHTML =
      '<option value="">Tous</option><option value="__empty__">(vide / vol)</option>' +
      codes.map((c) => `<option value="${c}">${c}</option>`).join('');
    if (currentCode === '__empty__' || codes.includes(currentCode)) {
      el.filterCode.value = currentCode;
    }

    // Date bounds
    const allDates = records.map((r) => r.date).filter(Boolean);
    if (allDates.length && !el.filterFrom.value) {
      const min = new Date(Math.min(...allDates));
      const max = new Date(Math.max(...allDates));
      el.filterFrom.min = isoDate(min);
      el.filterFrom.max = isoDate(max);
      el.filterTo.min = isoDate(min);
      el.filterTo.max = isoDate(max);
    }

    renderStatsAndCharts();
    renderTable();
  }

  function renderStatsAndCharts() {
    const filtered = getFiltered();
    const stats = computeStats(filtered);

    el.statsGrid.innerHTML = `
      <div class="stat-box">
        <div class="stat-value">${stats.totalDays}</div>
        <div class="stat-label">Jours person-jour</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${stats.worked}</div>
        <div class="stat-label">Jours travaillés</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${stats.nonWorked}</div>
        <div class="stat-label">Jours non travaillés</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${stats.ratioPeriod.toFixed(1)}%</div>
        <div class="stat-label">Ratio non travaillés / période</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${stats.ratioYear.toFixed(1)}%</div>
        <div class="stat-label">Ratio non travaillés / année*</div>
      </div>
    `;

    // Pie chart – répartition codes
    const pieLabels = Object.keys(stats.byCode);
    const pieData = Object.values(stats.byCode);
    const colors = generateColors(pieLabels.length);

    if (charts.pie) charts.pie.destroy();
    const pieCtx = $('#chart-pie').getContext('2d');
    charts.pie = new Chart(pieCtx, {
      type: 'doughnut',
      data: {
        labels: pieLabels,
        datasets: [{
          data: pieData,
          backgroundColor: colors,
          borderWidth: 1,
          borderColor: getComputedStyle(document.body).getPropertyValue('--bg-card') || '#1e293b'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: {
              color: getComputedStyle(document.body).getPropertyValue('--text') || '#f1f5f9',
              boxWidth: 12,
              font: { size: 11 }
            }
          }
        }
      }
    });

    // Bar chart – non worked per person
    const persons = Object.keys(stats.byPerson);
    const nonWorkedData = persons.map((p) => stats.byPerson[p].nonWorked);
    const totalData = persons.map((p) => stats.byPerson[p].total);

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
          x: {
            ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted') },
            grid: { color: 'transparent' }
          },
          y: {
            beginAtZero: true,
            ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted') },
            grid: { color: getComputedStyle(document.body).getPropertyValue('--border') }
          }
        },
        plugins: {
          legend: {
            labels: {
              color: getComputedStyle(document.body).getPropertyValue('--text')
            }
          }
        }
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
    el.tableCount.textContent = `${filtered.length} ligne${filtered.length > 1 ? 's' : ''}`;

    // Limite d'affichage pour perf (1000 max)
    const display = filtered.slice(0, 1000);
    el.tableBody.innerHTML = display
      .map((r) => {
        const type = isNonWorked(r) ? 'Non travaillé' : 'Travaillé';
        const badgeClass = isNonWorked(r) ? 'badge-off' : 'badge-work';
        return `
          <tr>
            <td>${formatDate(r.date)}</td>
            <td>${r.person}</td>
            <td>${r.role}</td>
            <td>${r.code ? `<span class="badge">${r.code}</span>` : '<span class="badge badge-work">vol</span>'}</td>
            <td>${r.note || '—'}</td>
            <td><span class="badge ${badgeClass}">${type}</span></td>
          </tr>`;
      })
      .join('');

    if (filtered.length > 1000) {
      el.tableBody.innerHTML += `<tr><td colspan="6" style="text-align:center;color:var(--text-muted)">… ${filtered.length - 1000} lignes supplémentaires non affichées</td></tr>`;
    }
  }

  // ---------- Modal non-worked ----------
  function openNonWorkedModal() {
    // Collect all codes present + known
    const present = new Set(records.map((r) => r.code).filter(Boolean));
    KNOWN_CODES.forEach((c) => present.add(c));
    const sorted = [...present].sort();

    el.checkboxes.innerHTML = sorted
      .map(
        (c) => `
      <label class="checkbox-item">
        <input type="checkbox" value="${c}" ${nonWorkedCodes.has(c) ? 'checked' : ''} />
        <span>${c}</span>
      </label>`
      )
      .join('');

    el.modal.classList.add('open');
  }

  function closeNonWorkedModal() {
    el.modal.classList.remove('open');
  }

  function saveNonWorked() {
    const checked = new Set();
    el.checkboxes.querySelectorAll('input:checked').forEach((inp) => {
      checked.add(inp.value);
    });
    nonWorkedCodes = checked;
    saveState();
    closeNonWorkedModal();
    renderStatsAndCharts();
    renderTable();
    toast('Configuration des jours non travaillés enregistrée');
  }

  // ---------- Export ----------
  function exportData() {
    const filtered = getFiltered();
    if (!filtered.length) {
      toast('Aucune donnée à exporter');
      return;
    }

    // CSV
    const header = ['Date', 'Personne', 'Rôle', 'Code', 'Note', 'Non_travaillé', 'Source'];
    const rows = filtered.map((r) => [
      isoDate(r.date),
      r.person,
      r.role,
      r.code || '',
      r.note || '',
      isNonWorked(r) ? 'oui' : 'non',
      r.source
    ]);

    const csv = [header, ...rows]
      .map((row) => row.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(';'))
      .join('\n');

    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `crew-planner-export-${isoDate(new Date())}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    toast('Export CSV téléchargé');
  }

  // ---------- Events ----------
  function bindEvents() {
    // Theme
    $$('.theme-btn').forEach((btn) => {
      btn.addEventListener('click', () => applyTheme(btn.dataset.theme));
    });

    // Import
    $('#btn-import').addEventListener('click', () => el.fileInput.click());
    $('#btn-import-empty').addEventListener('click', () => el.fileInput.click());
    el.fileInput.addEventListener('change', (e) => {
      handleFiles(e.target.files);
      e.target.value = '';
    });

    // Export
    $('#btn-export').addEventListener('click', exportData);

    // Clear
    $('#btn-clear').addEventListener('click', () => {
      if (confirm('Effacer toutes les données chargées ?')) {
        records = [];
        filesMeta = [];
        saveState();
        refreshUI();
        toast('Données effacées');
      }
    });

    // Non-worked modal
    $('#btn-nonworked').addEventListener('click', openNonWorkedModal);
    $('#modal-close').addEventListener('click', closeNonWorkedModal);
    $('#modal-cancel').addEventListener('click', closeNonWorkedModal);
    $('#modal-save').addEventListener('click', saveNonWorked);
    el.modal.addEventListener('click', (e) => {
      if (e.target === el.modal) closeNonWorkedModal();
    });

    // Filters
    ['filter-person', 'filter-from', 'filter-to', 'filter-code'].forEach((id) => {
      $(`#${id}`).addEventListener('change', () => {
        renderStatsAndCharts();
        renderTable();
      });
    });
    $('#btn-reset-filters').addEventListener('click', () => {
      el.filterPerson.value = '';
      el.filterFrom.value = '';
      el.filterTo.value = '';
      el.filterCode.value = '';
      renderStatsAndCharts();
      renderTable();
    });

    // Keyboard
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && el.modal.classList.contains('open')) {
        closeNonWorkedModal();
      }
    });
  }

  // ---------- Service Worker ----------
  function registerSW() {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js').catch((err) => {
        console.warn('SW registration failed', err);
      });
    }
  }

  // ---------- Init ----------
  function init() {
    const savedTheme = localStorage.getItem(THEME_KEY) || 'dark';
    applyTheme(savedTheme);
    loadState();
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
