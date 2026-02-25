  const SUPABASE_URL = 'https://zdgksbfvzrobeotekzpx.supabase.co';
  const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpkZ2tzYmZ2enJvYmVvdGVrenB4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzIwMTMxNDgsImV4cCI6MjA4NzU4OTE0OH0.tITnaa2hczGaqTjQOrX7bUmLFyjPBRU-oDfL1Yu97Q0';
  const CUSTOMIZE_PASSWORD = 'admin123';
  const THEME_PRESETS = [
    { bg: '#f5f0e8', paper: '#fffdf7', ink: '#1a1208', accent: '#c84b11', line: '#d9d0bc', cellHover: '#fff3e0', shadow: 'rgba(0,0,0,0.08)' },
    { bg: '#dfeeff', paper: '#f5faff', ink: '#0b2440', accent: '#2563eb', line: '#bcd0ea', cellHover: '#dbeafe', shadow: 'rgba(37,99,235,0.18)' },
    { bg: '#e3f7e9', paper: '#f5fff8', ink: '#0f2d1f', accent: '#15803d', line: '#b7dfc5', cellHover: '#dcfce7', shadow: 'rgba(21,128,61,0.18)' },
    { bg: '#ffe5ea', paper: '#fff7f9', ink: '#3c1420', accent: '#e11d48', line: '#efbdca', cellHover: '#ffe4ea', shadow: 'rgba(225,29,72,0.18)' },
    { bg: '#fff0d6', paper: '#fff8ea', ink: '#3d280a', accent: '#d97706', line: '#ecd1a4', cellHover: '#ffedd5', shadow: 'rgba(217,119,6,0.18)' },
    { bg: '#dfe3e8', paper: '#f1f4f7', ink: '#111827', accent: '#475569', line: '#c2cbd6', cellHover: '#e2e8f0', shadow: 'rgba(71,85,105,0.18)' },
    { bg: '#efe4ff', paper: '#f8f2ff', ink: '#2b1145', accent: '#7c3aed', line: '#d4baf0', cellHover: '#ede9fe', shadow: 'rgba(124,58,237,0.18)' },
    { bg: '#def6ee', paper: '#f3fffa', ink: '#0f3328', accent: '#0f766e', line: '#b9e1d6', cellHover: '#ccfbf1', shadow: 'rgba(15,118,110,0.18)' }
  ];
  const THEME_EVENT_PATTERNS = [
    { base: 'rgba(200,75,17,0.06)', deep: 'rgba(79,70,229,0.12)', meeting: 'rgba(220,38,38,0.12)', lunch: 'rgba(245,158,11,0.16)', break: 'rgba(245,158,11,0.16)', project: 'rgba(6,95,70,0.14)' },
    { base: 'rgba(37,99,235,0.08)', deep: 'rgba(29,78,216,0.16)', meeting: 'rgba(59,130,246,0.18)', lunch: 'rgba(14,116,144,0.18)', break: 'rgba(14,116,144,0.18)', project: 'rgba(6,182,212,0.18)' },
    { base: 'rgba(21,128,61,0.08)', deep: 'rgba(22,163,74,0.16)', meeting: 'rgba(101,163,13,0.18)', lunch: 'rgba(132,204,22,0.18)', break: 'rgba(132,204,22,0.18)', project: 'rgba(5,150,105,0.18)' },
    { base: 'rgba(225,29,72,0.08)', deep: 'rgba(190,24,93,0.16)', meeting: 'rgba(225,29,72,0.20)', lunch: 'rgba(244,114,182,0.18)', break: 'rgba(244,114,182,0.18)', project: 'rgba(236,72,153,0.18)' },
    { base: 'rgba(217,119,6,0.08)', deep: 'rgba(245,158,11,0.16)', meeting: 'rgba(234,88,12,0.18)', lunch: 'rgba(251,191,36,0.22)', break: 'rgba(251,191,36,0.22)', project: 'rgba(180,83,9,0.18)' },
    { base: 'rgba(71,85,105,0.08)', deep: 'rgba(51,65,85,0.16)', meeting: 'rgba(71,85,105,0.18)', lunch: 'rgba(100,116,139,0.18)', break: 'rgba(100,116,139,0.18)', project: 'rgba(30,41,59,0.18)' },
    { base: 'rgba(124,58,237,0.08)', deep: 'rgba(109,40,217,0.18)', meeting: 'rgba(168,85,247,0.18)', lunch: 'rgba(192,132,252,0.20)', break: 'rgba(192,132,252,0.20)', project: 'rgba(147,51,234,0.18)' },
    { base: 'rgba(15,118,110,0.08)', deep: 'rgba(13,148,136,0.16)', meeting: 'rgba(20,184,166,0.18)', lunch: 'rgba(45,212,191,0.20)', break: 'rgba(45,212,191,0.20)', project: 'rgba(5,150,105,0.18)' }
  ];
  const TABLE_CELL_PATTERNS = [
    ['rgba(248,239,229,0.74)', 'rgba(236,225,211,0.74)', 'rgba(244,232,220,0.74)', 'rgba(230,218,205,0.74)'],
    ['rgba(224,239,255,0.78)', 'rgba(213,232,252,0.78)', 'rgba(236,245,255,0.78)', 'rgba(204,225,249,0.78)'],
    ['rgba(227,246,233,0.78)', 'rgba(213,239,221,0.78)', 'rgba(240,252,244,0.78)', 'rgba(203,232,212,0.78)'],
    ['rgba(255,230,237,0.78)', 'rgba(248,216,226,0.78)', 'rgba(255,241,246,0.78)', 'rgba(243,205,218,0.78)'],
    ['rgba(255,241,217,0.80)', 'rgba(248,231,198,0.80)', 'rgba(255,248,232,0.80)', 'rgba(242,221,180,0.80)'],
    ['rgba(225,231,238,0.80)', 'rgba(210,219,230,0.80)', 'rgba(236,241,247,0.80)', 'rgba(198,208,221,0.80)'],
    ['rgba(239,227,255,0.80)', 'rgba(228,210,249,0.80)', 'rgba(248,241,255,0.80)', 'rgba(217,194,243,0.80)'],
    ['rgba(222,245,236,0.80)', 'rgba(205,236,224,0.80)', 'rgba(239,253,247,0.80)', 'rgba(190,226,212,0.80)']
  ];
  const THEME_PRESET_NAMES = ['Classic', 'Ocean', 'Forest', 'Rose', 'Sunset', 'Slate', 'Violet', 'Mint'];
  let supabaseClient = null;
  let isAppAuthenticated = false;
  let isAdmin = true;
  let nowLineTimer = null;
  let weekOffset = 0;
  const SCHEDULE_YEAR = 2026;
  let selectedEditableCell = null;
  let selectedRowIndex = null;
  let selectedColumnIndex = null;
  const changeHistory = [];
  const MAX_HISTORY = 60;
  let activeView = 'timetable';
  let calendarMonthDate = new Date(SCHEDULE_YEAR, new Date().getMonth(), 1);
  let selectedCountryCode = 'US';
  let holidayDates = new Map();
  let holidayList = [];
  let userCalendarEvents = {};
  let nowLineTheme = 'cyan';
  let activeThemePresetIndex = 0;
  let activeCellsPresetIndex = 0;
  let authMode = 'signin';
  let currentUserId = null;
  let columnThemePresetIndices = {};
  const DEFAULT_COLOR_PRESETS = ['#dbeafe', '#e9d5ff', '#dcfce7', '#fee2e2', '#fef3c7', '#cffafe', '#fce7f3', '#e5e7eb'];
  let colorPresets = [...DEFAULT_COLOR_PRESETS];
  const DEFAULT_TIME_SETTINGS = {
    intervalMinutes: 60,
    startMinute: 9 * 60,
    endMinute: 16 * 60,
    fullDay: false
  };
  let timeSettings = { ...DEFAULT_TIME_SETTINGS };

  const countryOptions = [
    'ZA', 'US', 'GB', 'AU', 'NZ', 'CA', 'IE', 'DE', 'FR', 'ES',
    'IT', 'NL', 'BE', 'CH', 'SE', 'NO', 'DK', 'FI', 'PL', 'IN', 'JP', 'BR'
  ];

  const defaultRows = [
    ['9:00 AM', 'Standup', 'Deep Work', 'Team Meeting', 'Deep Work', 'Review', '', ''],
    ['10:00 AM', 'Deep Work', 'Deep Work', 'Planning', 'Workshop', 'Deep Work', '', ''],
    ['11:00 AM', 'Deep Work', 'Meetings', 'Deep Work', 'Meetings', 'Admin', '', ''],
    ['12:00 PM', 'Lunch', 'Lunch', 'Lunch', 'Lunch', 'Lunch', '', ''],
    ['1:00 PM', 'Emails', 'Design', 'Emails', 'Design', '1:1s', '', ''],
    ['2:00 PM', 'Project A', 'Project B', 'Project A', 'Project B', 'Wrap Up', '', ''],
    ['3:00 PM', 'Project A', 'Project B', 'Project A', 'Project B', '', '', ''],
    ['4:00 PM', 'Review', 'Review', 'Review', 'Review', '', '', ''],
  ];

  const INITIAL_STATE = {
    title: 'My Weekly Schedule',
    subtitle: '',
    weekOffset: 0,
    headerLabels: ['Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
    rows: defaultRows.map(row => [...row]),
    timeSettings: { ...DEFAULT_TIME_SETTINGS },
    colorPresets: [...DEFAULT_COLOR_PRESETS],
    themePresetIndex: 0,
    cellsPresetIndex: 0,
    columnThemePresetIndices: {},
    customColors: [],
    userCalendarEvents: {},
    nowLineTheme: 'cyan'
  };

  function getColCount() {
    const headerRow = document.getElementById('header-row');
    return headerRow.querySelectorAll('th').length - 1; // exclude action col
  }

  function normalizeThemePresetIndex(index) {
    const parsed = Number(index);
    if (!Number.isInteger(parsed) || parsed < 0 || parsed >= THEME_PRESETS.length) {
      return 0;
    }
    return parsed;
  }

  function updateThemePresetActiveState() {
    const activePreviewIndex = normalizeThemePresetIndex(activeCellsPresetIndex);

    for (let i = 0; i < THEME_PRESETS.length; i++) {
      const button = document.getElementById(`theme-preset-${i}`);
      if (!button) continue;
      button.classList.toggle('active', i === activePreviewIndex);
      button.title = THEME_PRESET_NAMES[i] || `Theme ${i + 1}`;
    }

    const presetSelect = document.getElementById('color-preset-select');
    if (presetSelect) {
      presetSelect.value = String(activePreviewIndex);
      presetSelect.title = THEME_PRESET_NAMES[activePreviewIndex] || 'Theme preset';
    }
  }

  function hexToRgba(hex, alpha) {
    const sanitized = String(hex || '').replace('#', '').trim();
    if (!/^[0-9a-fA-F]{6}$/.test(sanitized)) {
      return `rgba(0, 0, 0, ${alpha})`;
    }

    const red = parseInt(sanitized.slice(0, 2), 16);
    const green = parseInt(sanitized.slice(2, 4), 16);
    const blue = parseInt(sanitized.slice(4, 6), 16);
    return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
  }

  function getThemeEventPattern(index) {
    const normalizedIndex = normalizeThemePresetIndex(index);
    return THEME_EVENT_PATTERNS[normalizedIndex] || THEME_EVENT_PATTERNS[0];
  }

  function getTableCellPattern(index) {
    const normalizedIndex = normalizeThemePresetIndex(index);
    return TABLE_CELL_PATTERNS[normalizedIndex] || TABLE_CELL_PATTERNS[0];
  }

  function getTableCellPatternColor(index, rowIndex, columnIndex) {
    const pattern = getTableCellPattern(index);
    const slot = Math.abs((rowIndex * 2) + columnIndex) % pattern.length;
    return pattern[slot];
  }

  function applyThemeVariables(theme, presetIndex = 0) {
    const pattern = getThemeEventPattern(presetIndex);
    const root = document.documentElement;
    root.style.setProperty('--bg', theme.bg);
    root.style.setProperty('--paper', theme.paper);
    root.style.setProperty('--ink', theme.ink);
    root.style.setProperty('--accent', theme.accent);
    root.style.setProperty('--line', theme.line);
    root.style.setProperty('--cell-hover', theme.cellHover);
    root.style.setProperty('--shadow', theme.shadow);
    root.style.setProperty('--header-grad-start', hexToRgba(theme.paper, 0.96));
    root.style.setProperty('--header-grad-end', hexToRgba(theme.bg, 0.92));
    root.style.setProperty('--panel-grad-start', hexToRgba(theme.paper, 0.82));
    root.style.setProperty('--panel-grad-end', hexToRgba(theme.bg, 0.90));
    root.style.setProperty('--toolbar-bg', hexToRgba(theme.paper, 0.94));
    root.style.setProperty('--table-head-bg', hexToRgba(theme.paper, 0.98));
    root.style.setProperty('--accent-soft', hexToRgba(theme.accent, 0.10));
    root.style.setProperty('--accent-mid', hexToRgba(theme.accent, 0.20));
    root.style.setProperty('--focus-bg', hexToRgba(theme.accent, 0.12));
    root.style.setProperty('--event-chip-base', pattern.base);
    root.style.setProperty('--event-chip-deep', pattern.deep);
    root.style.setProperty('--event-chip-meeting', pattern.meeting);
    root.style.setProperty('--event-chip-lunch', pattern.lunch);
    root.style.setProperty('--event-chip-break', pattern.break);
    root.style.setProperty('--event-chip-project', pattern.project);
  }

  function applyThemePreset(index, options = {}) {
    const { pushHistory = true } = options;
    const normalizedIndex = normalizeThemePresetIndex(index);

    if (pushHistory) {
      pushHistorySnapshot();
    }

    activeThemePresetIndex = normalizedIndex;
    applyThemeVariables(THEME_PRESETS[activeThemePresetIndex], activeThemePresetIndex);
    updateThemePresetActiveState();
    saveStateForCurrentUser();
  }

  function onThemePresetSelectChange() {
    const presetSelect = document.getElementById('color-preset-select');
    if (!presetSelect) return;
    const presetIndex = normalizeThemePresetIndex(presetSelect.value);
    presetSelect.title = THEME_PRESET_NAMES[presetIndex] || 'Theme preset';
  }

  function normalizeColumnThemePresetIndices(raw) {
    if (!raw || typeof raw !== 'object') {
      return {};
    }

    const normalized = {};
    Object.entries(raw).forEach(([columnKey, presetIndex]) => {
      const columnIndex = Number(columnKey);
      if (!Number.isInteger(columnIndex) || columnIndex < 0) return;
      normalized[columnIndex] = normalizeThemePresetIndex(presetIndex);
    });

    return normalized;
  }

  function clearColumnThemeVisual(columnIndex) {
    const headerCell = document.querySelectorAll('#header-row th')[columnIndex];
    if (headerCell) {
      headerCell.classList.remove('column-themed');
      headerCell.style.removeProperty('--column-theme-bg');
    }

    document.querySelectorAll('#tbody tr').forEach((row) => {
      const td = row.children[columnIndex];
      if (!td) return;
      td.classList.remove('column-themed');
      td.style.removeProperty('--column-theme-bg');
      td.style.removeProperty('--column-theme-cell-bg');
      td.style.removeProperty('--column-event-chip-base');
      td.style.removeProperty('--column-event-chip-deep');
      td.style.removeProperty('--column-event-chip-meeting');
      td.style.removeProperty('--column-event-chip-lunch');
      td.style.removeProperty('--column-event-chip-break');
      td.style.removeProperty('--column-event-chip-project');
    });
  }

  function applyColumnThemeVisual(columnIndex, presetIndex) {
    const normalizedPresetIndex = normalizeThemePresetIndex(presetIndex);
    const theme = THEME_PRESETS[normalizedPresetIndex];
    const pattern = getThemeEventPattern(normalizedPresetIndex);
    if (!theme) return;

    const headerCell = document.querySelectorAll('#header-row th')[columnIndex];
    if (headerCell) {
      headerCell.classList.add('column-themed');
      headerCell.style.setProperty('--column-theme-bg', hexToRgba(theme.paper, 0.98));
    }

    document.querySelectorAll('#tbody tr').forEach((row, rowIndex) => {
      const td = row.children[columnIndex];
      if (!td) return;
      td.classList.add('column-themed');
      td.style.setProperty('--column-theme-bg', hexToRgba(theme.cellHover, 0.95));
      td.style.setProperty('--column-theme-cell-bg', getTableCellPatternColor(normalizedPresetIndex, rowIndex, columnIndex));
      td.style.setProperty('--column-event-chip-base', pattern.base);
      td.style.setProperty('--column-event-chip-deep', pattern.deep);
      td.style.setProperty('--column-event-chip-meeting', pattern.meeting);
      td.style.setProperty('--column-event-chip-lunch', pattern.lunch);
      td.style.setProperty('--column-event-chip-break', pattern.break);
      td.style.setProperty('--column-event-chip-project', pattern.project);
    });
  }

  function renderColumnThemePresets() {
    const colCount = getColCount();

    for (let i = 0; i < colCount; i++) {
      clearColumnThemeVisual(i);
    }

    Object.entries(columnThemePresetIndices).forEach(([columnKey, presetIndex]) => {
      const columnIndex = Number(columnKey);
      if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= colCount) {
        delete columnThemePresetIndices[columnKey];
        return;
      }
      applyColumnThemeVisual(columnIndex, presetIndex);
    });
  }

  function getTargetColumnIndexForPreset() {
    if (selectedColumnIndex !== null && selectedColumnIndex >= 0 && selectedColumnIndex < getColCount()) {
      return selectedColumnIndex;
    }

    if (selectedEditableCell) {
      const td = selectedEditableCell.closest('td');
      if (td && td.parentElement) {
        const index = Array.from(td.parentElement.children).indexOf(td);
        if (index >= 0 && index < getColCount()) {
          return index;
        }
      }
    }

    return null;
  }

  function applyThemePresetToSelectedColumn(index, options = {}) {
    if (!isAdmin) return;

    const { pushHistory = true } = options;
    const targetColumn = getTargetColumnIndexForPreset();
    if (targetColumn === null) {
      alert('Select a column header or a cell in that column first.');
      return;
    }

    if (pushHistory) {
      pushHistorySnapshot();
    }

    const normalizedIndex = normalizeThemePresetIndex(index);
    columnThemePresetIndices[targetColumn] = normalizedIndex;
    activeCellsPresetIndex = normalizedIndex;
    renderColumnThemePresets();
    updateThemePresetActiveState();
    saveStateForCurrentUser();
  }

  function applyThemePresetToAllCells(index, options = {}) {
    if (!isAdmin) return;

    const { pushHistory = true } = options;
    if (pushHistory) {
      pushHistorySnapshot();
    }

    const normalizedIndex = normalizeThemePresetIndex(index);
    const next = {};
    for (let columnIndex = 1; columnIndex < getColCount(); columnIndex++) {
      next[columnIndex] = normalizedIndex;
    }

    columnThemePresetIndices = next;
    activeCellsPresetIndex = normalizedIndex;
    renderColumnThemePresets();
    updateThemePresetActiveState();
    saveStateForCurrentUser();
  }

  function applyThemePresetFromChip(index) {
    applyThemePresetToAllCells(index);
  }

  function shiftColumnThemePresetsAfterRemoval(removedIndex) {
    const next = {};
    Object.entries(columnThemePresetIndices).forEach(([columnKey, presetIndex]) => {
      const columnIndex = Number(columnKey);
      if (columnIndex === removedIndex) return;
      if (columnIndex > removedIndex) {
        next[columnIndex - 1] = presetIndex;
        return;
      }
      next[columnIndex] = presetIndex;
    });
    columnThemePresetIndices = next;
  }

  function clearCellSelection() {
    if (!selectedEditableCell) return;
    selectedEditableCell.classList.remove('selected-cell');
    selectedEditableCell = null;
  }

  function clearRowSelection() {
    document.querySelectorAll('#tbody tr.selected-row').forEach(row => row.classList.remove('selected-row'));
    selectedRowIndex = null;
  }

  function clearColumnSelection() {
    document.querySelectorAll('#header-row th.selected-column').forEach(th => th.classList.remove('selected-column'));
    document.querySelectorAll('#tbody td.selected-column').forEach(td => td.classList.remove('selected-column'));
    selectedColumnIndex = null;
    updateThemePresetActiveState();
  }

  function selectRowByIndex(rowIndex) {
    const rows = Array.from(document.querySelectorAll('#tbody tr'));
    if (rowIndex < 0 || rowIndex >= rows.length) return;

    clearCellSelection();
    clearColumnSelection();
    clearRowSelection();

    selectedRowIndex = rowIndex;
    rows[rowIndex].classList.add('selected-row');
  }

  function selectColumnByIndex(columnIndex) {
    if (columnIndex < 0 || columnIndex >= getColCount()) return;

    clearCellSelection();
    clearRowSelection();
    clearColumnSelection();

    selectedColumnIndex = columnIndex;
    const headerCell = document.querySelectorAll('#header-row th')[columnIndex];
    if (headerCell) headerCell.classList.add('selected-column');

    document.querySelectorAll('#tbody tr').forEach(row => {
      const td = row.children[columnIndex];
      if (td) td.classList.add('selected-column');
    });

    updateThemePresetActiveState();
  }

  function bindHeaderSelection() {
    document.querySelectorAll('#header-row th input').forEach(input => {
      if (input.dataset.selectionBound === '1') return;
      input.dataset.selectionBound = '1';
      input.addEventListener('click', (event) => {
        event.stopPropagation();
        const headerCell = input.closest('th');
        const headerCells = Array.from(document.querySelectorAll('#header-row th'));
        const index = headerCells.indexOf(headerCell);
        if (index >= 0 && index < getColCount()) {
          selectColumnByIndex(index);
        }
      });
    });
  }

  function setEditMode(enabled) {
    const editableElements = document.querySelectorAll(
      '#main-title, .subtitle-input, #timetable input, #timetable textarea, .row-del'
    );

    editableElements.forEach(el => {
      if (el.tagName === 'BUTTON') {
        el.disabled = !enabled;
      } else {
        el.disabled = !enabled;
      }
    });

    document.querySelectorAll('[data-requires-admin]').forEach(btn => {
      btn.disabled = !enabled;
    });

    document.querySelectorAll('[data-admin-action]').forEach(btn => {
      btn.disabled = !enabled;
    });

    const tableWrap = document.querySelector('.table-wrap');
    tableWrap.classList.toggle('locked', !enabled);

    const status = document.getElementById('auth-status');
    const unlockBtn = document.getElementById('customize-unlock-btn');
    const lockBtn = document.getElementById('customize-lock-btn');

    if (!status || !unlockBtn || !lockBtn) return;

    if (enabled) {
      status.textContent = 'Customize unlocked';
      status.style.color = 'var(--accent)';
      unlockBtn.style.display = 'none';
      lockBtn.style.display = 'inline-block';
    } else {
      status.textContent = 'Customize locked';
      status.style.color = '#888';
      unlockBtn.style.display = 'inline-block';
      lockBtn.style.display = 'none';
    }
  }

  function isCustomizePanelOpen() {
    const panel = document.getElementById('settings-panel');
    return panel?.classList.contains('open') === true;
  }

  function syncEditModeWithCustomizePanel() {
    setEditMode(isAdmin && isCustomizePanelOpen());
  }

  function toggleSettingsPanel() {
    const panel = document.getElementById('settings-panel');
    const toggleBtn = document.getElementById('settings-toggle-btn');
    if (!panel || !toggleBtn) return;

    const isOpen = panel.classList.toggle('open');
    toggleBtn.textContent = isOpen ? '✕ Close' : '⚙ Customize';
    syncEditModeWithCustomizePanel();
  }

  function setAuthMode(mode) {
    const normalizedMode = mode === 'signup' ? 'signup' : 'signin';
    authMode = normalizedMode;

    const title = document.getElementById('app-auth-title');
    const confirmPasswordInput = document.getElementById('app-confirm-password');
    const signInBtn = document.getElementById('signin-btn');
    const signUpBtn = document.getElementById('signup-btn');

    if (normalizedMode === 'signup') {
      title.textContent = 'Create your account';
      confirmPasswordInput.style.display = 'block';
      confirmPasswordInput.disabled = false;
      signUpBtn.classList.add('auth-active');
      signInBtn.classList.remove('auth-active');
      return;
    }

    title.textContent = 'Sign in to open your timetable';
    confirmPasswordInput.value = '';
    confirmPasswordInput.style.display = 'none';
    confirmPasswordInput.disabled = true;
    signInBtn.classList.add('auth-active');
    signUpBtn.classList.remove('auth-active');
  }

  function submitAuthByMode() {
    if (authMode === 'signup') {
      appSignUp();
      return;
    }

    appLogin();
  }

  function handleAuthEnter(event) {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    submitAuthByMode();
  }

  function getAuthEmailRedirectUrl() {
    if (window.location.protocol === 'http:' || window.location.protocol === 'https:') {
      return window.location.origin + window.location.pathname;
    }
    return null;
  }

  function getUserStateStorageKey(userId) {
    return `timetable_state_${userId}`;
  }

  function saveStateForCurrentUser() {
    if (!currentUserId || !isAppAuthenticated) return;

    try {
      const state = captureState();
      localStorage.setItem(getUserStateStorageKey(currentUserId), JSON.stringify(state));
    } catch (error) {
      // ignore storage errors
    }
  }

  function loadStateForCurrentUser() {
    if (!currentUserId) return;

    try {
      const raw = localStorage.getItem(getUserStateStorageKey(currentUserId));
      if (!raw) {
        restoreState(INITIAL_STATE);
        pushHistorySnapshot();
        return;
      }

      const parsed = JSON.parse(raw);
      restoreState(parsed);
      pushHistorySnapshot();
    } catch (error) {
      restoreState(INITIAL_STATE);
      pushHistorySnapshot();
    }
  }

  function appLogin() {
    const email = document.getElementById('app-email').value.trim();
    const password = document.getElementById('app-password').value;
    const appStatus = document.getElementById('app-login-status');

    if (authMode !== 'signin') {
      setAuthMode('signin');
      appStatus.textContent = 'Sign in with email and password';
      appStatus.style.color = '#888';
      return;
    }

    if (!supabaseClient) {
      appStatus.textContent = 'Set SUPABASE_URL and SUPABASE_ANON_KEY first';
      appStatus.style.color = '#b42318';
      return;
    }

    if (!email || !password) {
      appStatus.textContent = 'Enter email and password';
      appStatus.style.color = '#b42318';
      return;
    }

    supabaseClient.auth.signInWithPassword({ email, password })
      .then(({ error }) => {
        if (error) {
          appStatus.textContent = error.message;
          appStatus.style.color = '#b42318';
          return;
        }

        if (currentUserId) {
          loadStateForCurrentUser();
        }
        showAppShell();
        appStatus.textContent = 'Signed in';
        appStatus.style.color = '#888';
      });
  }

  function appSignUp() {
    const email = document.getElementById('app-email').value.trim();
    const password = document.getElementById('app-password').value;
    const confirmPassword = document.getElementById('app-confirm-password').value;
    const appStatus = document.getElementById('app-login-status');

    if (authMode !== 'signup') {
      setAuthMode('signup');
      appStatus.textContent = 'Create account: enter password and confirm password';
      appStatus.style.color = '#888';
      return;
    }

    if (!supabaseClient) {
      appStatus.textContent = 'Set SUPABASE_URL and SUPABASE_ANON_KEY first';
      appStatus.style.color = '#b42318';
      return;
    }

    if (!email || !password) {
      appStatus.textContent = 'Enter email and password';
      appStatus.style.color = '#b42318';
      return;
    }

    if (password.length < 6) {
      appStatus.textContent = 'Password must be at least 6 characters';
      appStatus.style.color = '#b42318';
      return;
    }

    if (password !== confirmPassword) {
      appStatus.textContent = 'Passwords do not match';
      appStatus.style.color = '#b42318';
      return;
    }

    const emailRedirectTo = getAuthEmailRedirectUrl();
    const signUpPayload = emailRedirectTo
      ? { email, password, options: { emailRedirectTo } }
      : { email, password };

    supabaseClient.auth.signUp(signUpPayload)
      .then(async ({ data, error }) => {
        if (error) {
          appStatus.textContent = error.message;
          appStatus.style.color = '#b42318';
          return;
        }

        if (data?.session) {
          await supabaseClient.auth.signOut();
        }

        setAuthMode('signin');
        document.getElementById('app-password').value = '';
        document.getElementById('app-confirm-password').value = '';
        appStatus.textContent = 'Account created. Please sign in.';
        appStatus.style.color = '#888';
      });
  }

  function showAppShell() {
    isAppAuthenticated = true;
    document.getElementById('app-lock').classList.add('app-hidden');
    document.getElementById('app-shell').classList.remove('app-hidden');
  }

  function initSupabaseAuth() {
    const appStatus = document.getElementById('app-login-status');

    if (!window.supabase || !SUPABASE_URL || !SUPABASE_ANON_KEY || SUPABASE_URL.includes('YOUR_PROJECT_ID') || SUPABASE_ANON_KEY.includes('YOUR_SUPABASE_ANON_KEY')) {
      appStatus.textContent = 'Add your Supabase URL + anon key to enable database login';
      appStatus.style.color = '#b42318';
      return;
    }

    supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

    supabaseClient.auth.getSession().then(({ data }) => {
      const session = data?.session;
      currentUserId = session?.user?.id || null;

      if (session) {
        loadStateForCurrentUser();
        showAppShell();
      }
    });

    supabaseClient.auth.onAuthStateChange((event, session) => {
      currentUserId = session?.user?.id || null;

      if (session) {
        loadStateForCurrentUser();
        showAppShell();
        return;
      }

      isAppAuthenticated = false;
      document.getElementById('app-shell').classList.add('app-hidden');
      document.getElementById('app-lock').classList.remove('app-hidden');
    });
  }

  function unlockCustomize() {
    const password = document.getElementById('customize-password').value;
    if (password === CUSTOMIZE_PASSWORD) {
      isAdmin = true;
      syncEditModeWithCustomizePanel();
      document.getElementById('customize-password').value = '';
      return;
    }

    alert('Invalid customize password.');
    isAdmin = false;
    setEditMode(false);
  }

  function lockCustomize() {
    isAdmin = false;
    syncEditModeWithCustomizePanel();
  }

  function makeCell(value = '', isFirst = false) {
    const td = document.createElement('td');
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.value = value;
    inp.spellcheck = false;
    if (isFirst) {
      inp.readOnly = true;
      inp.tabIndex = -1;
    }
    inp.addEventListener('input', () => applyEventChip(inp));
    inp.addEventListener('focus', () => selectEditableCell(inp));
    inp.addEventListener('click', (event) => event.stopPropagation());
    inp.addEventListener('change', () => {
      pushHistorySnapshot();
      updateNowLine();
    });
    td.addEventListener('click', () => {
      if (isFirst) {
        const tr = td.closest('tr');
        const rows = Array.from(document.querySelectorAll('#tbody tr'));
        const rowIndex = rows.indexOf(tr);
        if (rowIndex !== -1) {
          selectRowByIndex(rowIndex);
          return;
        }
      }
      selectEditableCell(inp);
    });
    inp.addEventListener('dblclick', () => expandToTextarea(inp, td));
    applyEventChip(inp);
    td.appendChild(inp);
    return td;
  }

  function expandToTextarea(inp, td) {
    const ta = document.createElement('textarea');
    ta.value = inp.value;
    ta.rows = 3;
    ta.spellcheck = false;
    ta.addEventListener('input', () => applyEventChip(ta));
    ta.addEventListener('focus', () => selectEditableCell(ta));
    td.replaceChild(ta, inp);
    applyEventChip(ta);
    ta.focus();
    ta.addEventListener('blur', () => {
      const newInp = document.createElement('input');
      newInp.type = 'text';
      newInp.value = ta.value;
      newInp.spellcheck = false;
      newInp.addEventListener('input', () => applyEventChip(newInp));
      newInp.addEventListener('focus', () => selectEditableCell(newInp));
      newInp.addEventListener('change', () => {
        pushHistorySnapshot();
        updateNowLine();
      });
      newInp.addEventListener('dblclick', () => expandToTextarea(newInp, td));
      td.replaceChild(newInp, ta);
      applyEventChip(newInp);
      pushHistorySnapshot();
      updateNowLine();
    });
  }

  function addRow(data = []) {
    if (!isAdmin && data.length === 0) return;
    if (isAdmin && data.length === 0) pushHistorySnapshot();

    const colCount = getColCount();
    const tr = document.createElement('tr');

    for (let i = 0; i < colCount; i++) {
      const td = makeCell(data[i] || '', i === 0);
      tr.appendChild(td);
    }

    // action cell
    const actionTd = document.createElement('td');
    actionTd.className = 'action-cell';
    const delBtn = document.createElement('button');
    delBtn.className = 'row-del';
    delBtn.textContent = '×';
    delBtn.title = 'Delete row';
    delBtn.onclick = () => {
      if (!isAdmin) return;
      pushHistorySnapshot();
      tr.remove();
      updateNowLine();
    };
    actionTd.appendChild(delBtn);
    tr.appendChild(actionTd);

    document.getElementById('tbody').appendChild(tr);
    renderColumnThemePresets();
    updateTodayHighlight();
    updateNowLine();
  }

  function removeRow() {
    if (!isAdmin) return;

    const rows = document.querySelectorAll('#tbody tr');
    if (!rows.length) return;

    const rowCount = rows.length;
    const choice = prompt(`Remove which row? Enter 1-${rowCount} (top to bottom)`, String(rowCount));
    if (choice === null) return;

    const rowNumber = Number(choice);
    if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > rowCount) {
      alert('Invalid row number.');
      return;
    }

    pushHistorySnapshot();
    rows[rowNumber - 1].remove();
    clearRowSelection();
    updateTodayHighlight();
    updateNowLine();
  }

  function addColumn() {
    if (!isAdmin) return;
    pushHistorySnapshot();

    const headerRow = document.getElementById('header-row');
    const actionTh = headerRow.lastElementChild;
    const th = document.createElement('th');
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.value = 'New Column';
    inp.spellcheck = false;
    th.appendChild(inp);
    headerRow.insertBefore(th, actionTh);

    document.querySelectorAll('#tbody tr').forEach(tr => {
      const actionTd = tr.lastElementChild;
      const td = makeCell('');
      tr.insertBefore(td, actionTd);
    });

    if (Object.keys(columnThemePresetIndices).length > 0) {
      applyThemePresetToAllCells(activeCellsPresetIndex, { pushHistory: false });
    }

    bindHeaderSelection();
    renderColumnThemePresets();
    updateTodayHighlight();
    updateNowLine();
  }

  function removeColumn() {
    if (!isAdmin) return;

    const minimumColumns = 8;
    const colCount = getColCount();
    if (colCount <= minimumColumns) return;

    const removableColumns = [];
    const headerInputs = document.querySelectorAll('#header-row th input');
    for (let i = minimumColumns; i < colCount; i++) {
      const title = (headerInputs[i]?.value || `Column ${i + 1}`).trim();
      removableColumns.push({
        number: i + 1,
        label: title
      });
    }

    const choicesText = removableColumns.map(col => `${col.number}: ${col.label}`).join('\n');
    const defaultChoice = String(removableColumns[removableColumns.length - 1].number);
    const choice = prompt(`Remove which column?\n${choicesText}\n\nEnter column number:`, defaultChoice);
    if (choice === null) return;

    const chosenColumn = Number(choice);
    const selected = removableColumns.find(col => col.number === chosenColumn);
    if (!selected) {
      alert('Invalid column number.');
      return;
    }

    pushHistorySnapshot();

    const headerRow = document.getElementById('header-row');
    const removableHeader = headerRow.children[selected.number - 1];
    if (removableHeader) {
      removableHeader.remove();
    }

    document.querySelectorAll('#tbody tr').forEach(tr => {
      const removableCell = tr.children[selected.number - 1];
      if (removableCell) {
        removableCell.remove();
      }
    });

    shiftColumnThemePresetsAfterRemoval(selected.number - 1);
    renderColumnThemePresets();

    clearColumnSelection();
    bindHeaderSelection();
    updateTodayHighlight();
    updateNowLine();
  }

  function removeSelectedItem() {
    if (!isAdmin) return;

    if (selectedRowIndex !== null) {
      const rows = document.querySelectorAll('#tbody tr');
      if (!rows.length || selectedRowIndex >= rows.length) {
        clearRowSelection();
        return;
      }

      pushHistorySnapshot();
      rows[selectedRowIndex].remove();
      clearRowSelection();
      updateTodayHighlight();
      updateNowLine();
      return;
    }

    if (selectedColumnIndex !== null) {
      const minimumColumns = 8;
      if (selectedColumnIndex + 1 <= minimumColumns) {
        alert('Base columns cannot be removed. Select an added column.');
        return;
      }

      pushHistorySnapshot();
      const headerRow = document.getElementById('header-row');
      const removableHeader = headerRow.children[selectedColumnIndex];
      if (removableHeader) {
        removableHeader.remove();
      }

      document.querySelectorAll('#tbody tr').forEach(tr => {
        const removableCell = tr.children[selectedColumnIndex];
        if (removableCell) {
          removableCell.remove();
        }
      });

      shiftColumnThemePresetsAfterRemoval(selectedColumnIndex);
      renderColumnThemePresets();

      clearColumnSelection();
      bindHeaderSelection();
      updateTodayHighlight();
      updateNowLine();
      return;
    }

    if (selectedEditableCell) {
      pushHistorySnapshot();
      selectedEditableCell.value = '';
      selectedEditableCell.classList.remove('custom-cell-color');
      selectedEditableCell.style.removeProperty('--custom-cell-color');
      applyEventChip(selectedEditableCell);
      updateNowLine();
      return;
    }

    alert('Select a row, column, or cell first.');
  }

  function clearAll() {
    if (!isAdmin) return;

    if (!confirm('Clear all cell content?')) return;
    pushHistorySnapshot();
    document.querySelectorAll('#tbody input, #tbody textarea').forEach(el => el.value = '');
    document.querySelectorAll('#tbody input, #tbody textarea').forEach(el => applyEventChip(el));
    updateNowLine();
  }

  function parseTimeToMinutes(text) {
    const normalized = (text || '').trim().toUpperCase();
    const match = normalized.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/);
    if (!match) return null;

    let hour = parseInt(match[1], 10);
    const minute = parseInt(match[2] || '0', 10);
    const meridiem = match[3];

    if (hour === 12) hour = 0;
    if (meridiem === 'PM') hour += 12;

    return hour * 60 + minute;
  }

  function formatMinutesToSlotLabel(minutes) {
    const h24 = Math.floor(minutes / 60) % 24;
    const mm = Math.floor(minutes % 60);
    const suffix = h24 >= 12 ? 'PM' : 'AM';
    const h12 = h24 % 12 || 12;
    return `${h12}:${String(mm).padStart(2, '0')} ${suffix}`;
  }

  function formatMinutesTo12h(minutes) {
    const wholeMinutes = Math.floor(minutes);
    const h24 = Math.floor(wholeMinutes / 60);
    const mm = wholeMinutes % 60;
    const seconds = Math.floor((minutes - wholeMinutes) * 60);
    const suffix = h24 >= 12 ? 'PM' : 'AM';
    const h12 = h24 % 12 || 12;
    return `${h12}:${String(mm).padStart(2, '0')}:${String(seconds).padStart(2, '0')} ${suffix}`;
  }

  function minuteToOptionValue(minutes) {
    return String(minutes).padStart(4, '0');
  }

  function optionValueToMinute(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) return 0;
    return Math.max(0, Math.min(1439, parsed));
  }

  function buildTimeSelectOptions() {
    const startSelect = document.getElementById('time-start-select');
    const endSelect = document.getElementById('time-end-select');
    if (!startSelect || !endSelect) return;

    const optionsHtml = [];
    for (let minutes = 0; minutes < 24 * 60; minutes += 15) {
      optionsHtml.push(`<option value="${minuteToOptionValue(minutes)}">${formatMinutesToSlotLabel(minutes)}</option>`);
    }

    startSelect.innerHTML = optionsHtml.join('');
    endSelect.innerHTML = optionsHtml.join('');
  }

  function normalizeTimeSettings(raw = {}) {
    const intervalMinutes = [15, 30, 60].includes(Number(raw.intervalMinutes))
      ? Number(raw.intervalMinutes)
      : DEFAULT_TIME_SETTINGS.intervalMinutes;

    const fullDay = raw.fullDay === true;

    if (fullDay) {
      return {
        intervalMinutes,
        startMinute: 0,
        endMinute: (24 * 60) - intervalMinutes,
        fullDay: true
      };
    }

    let startMinute = Number.isFinite(Number(raw.startMinute)) ? Number(raw.startMinute) : DEFAULT_TIME_SETTINGS.startMinute;
    let endMinute = Number.isFinite(Number(raw.endMinute)) ? Number(raw.endMinute) : DEFAULT_TIME_SETTINGS.endMinute;

    startMinute = Math.max(0, Math.min(1439, Math.floor(startMinute / 15) * 15));
    endMinute = Math.max(0, Math.min(1439, Math.floor(endMinute / 15) * 15));

    if (endMinute < startMinute) {
      endMinute = startMinute;
    }

    const stepCount = Math.max(0, Math.floor((endMinute - startMinute) / intervalMinutes));
    endMinute = startMinute + (stepCount * intervalMinutes);

    return {
      intervalMinutes,
      startMinute,
      endMinute,
      fullDay: false
    };
  }

  function setTimeSettingsUIState() {
    const intervalSelect = document.getElementById('time-interval-select');
    const modeSelect = document.getElementById('time-range-mode');
    const startSelect = document.getElementById('time-start-select');
    const endSelect = document.getElementById('time-end-select');
    if (!intervalSelect || !modeSelect || !startSelect || !endSelect) return;

    intervalSelect.value = String(timeSettings.intervalMinutes);
    modeSelect.value = timeSettings.fullDay ? 'full' : 'custom';
    startSelect.value = minuteToOptionValue(timeSettings.startMinute);
    endSelect.value = minuteToOptionValue(timeSettings.endMinute);

    const disableRange = timeSettings.fullDay;
    startSelect.disabled = disableRange;
    endSelect.disabled = disableRange;

    updateTimePresetActiveState();
  }

  function getMatchingTimePresetName(settings = timeSettings) {
    const normalized = normalizeTimeSettings(settings);

    if (normalized.fullDay) {
      return 'full';
    }

    if (normalized.startMinute === (1 * 60) && normalized.endMinute === (4 * 60)) {
      return 'early';
    }

    if (normalized.startMinute === (9 * 60) && normalized.endMinute === (17 * 60)) {
      return 'workday';
    }

    return null;
  }

  function updateTimePresetActiveState() {
    const earlyBtn = document.getElementById('time-preset-early');
    const workdayBtn = document.getElementById('time-preset-workday');
    const fullBtn = document.getElementById('time-preset-full');
    if (!earlyBtn || !workdayBtn || !fullBtn) return;

    const activePreset = getMatchingTimePresetName(timeSettings);

    earlyBtn.classList.toggle('time-preset-active', activePreset === 'early');
    workdayBtn.classList.toggle('time-preset-active', activePreset === 'workday');
    fullBtn.classList.toggle('time-preset-active', activePreset === 'full');
  }

  function getCurrentRowsData() {
    return Array.from(document.querySelectorAll('#tbody tr')).map(row => {
      const cells = Array.from(row.querySelectorAll('td input, td textarea'));
      return cells.map(cell => cell.value);
    });
  }

  function getTimeSlots(settings = timeSettings) {
    const normalized = normalizeTimeSettings(settings);
    const slots = [];
    for (let minutes = normalized.startMinute; minutes <= normalized.endMinute; minutes += normalized.intervalMinutes) {
      slots.push(minutes);
    }
    return slots;
  }

  function rebuildRowsForTimeSettings(existingRowsData = []) {
    const tbody = document.getElementById('tbody');
    const dayColumnCount = Math.max(0, getColCount() - 1);
    const slots = getTimeSlots(timeSettings);

    tbody.innerHTML = '';

    slots.forEach((slotMinute, rowIndex) => {
      const rowData = [];
      rowData[0] = formatMinutesToSlotLabel(slotMinute);
      for (let col = 1; col <= dayColumnCount; col++) {
        rowData[col] = existingRowsData[rowIndex]?.[col] || '';
      }
      addRow(rowData);
    });

    clearCellSelection();
    clearRowSelection();
    clearColumnSelection();
    bindHeaderSelection();
    updateTodayHighlight();
    updateNowLine();
  }

  function applyTimeSettings(nextSettings, options = {}) {
    const { pushHistory = false, existingRowsData = null } = options;
    if (pushHistory) {
      pushHistorySnapshot();
    }

    const rowsData = Array.isArray(existingRowsData) ? existingRowsData : getCurrentRowsData();
    timeSettings = normalizeTimeSettings(nextSettings);
    setTimeSettingsUIState();
    rebuildRowsForTimeSettings(rowsData);
  }

  function inferTimeSettingsFromRows(rows) {
    if (!Array.isArray(rows) || rows.length < 2) {
      return { ...DEFAULT_TIME_SETTINGS };
    }

    const first = parseTimeToMinutes(rows[0]?.[0]);
    const second = parseTimeToMinutes(rows[1]?.[0]);
    if (first === null || second === null) {
      return { ...DEFAULT_TIME_SETTINGS };
    }

    const intervalMinutes = second - first;
    if (![15, 30, 60].includes(intervalMinutes)) {
      return { ...DEFAULT_TIME_SETTINGS };
    }

    const startMinute = first;
    const endMinute = startMinute + (intervalMinutes * (rows.length - 1));
    const fullDay = startMinute === 0 && endMinute === (24 * 60) - intervalMinutes;

    return normalizeTimeSettings({
      intervalMinutes,
      startMinute,
      endMinute,
      fullDay
    });
  }

  function onTimeSettingsChange() {
    if (!isAdmin) return;

    const intervalSelect = document.getElementById('time-interval-select');
    const modeSelect = document.getElementById('time-range-mode');
    const startSelect = document.getElementById('time-start-select');
    const endSelect = document.getElementById('time-end-select');
    if (!intervalSelect || !modeSelect || !startSelect || !endSelect) return;

    const next = normalizeTimeSettings({
      intervalMinutes: Number(intervalSelect.value),
      fullDay: modeSelect.value === 'full',
      startMinute: optionValueToMinute(startSelect.value),
      endMinute: optionValueToMinute(endSelect.value)
    });

    applyTimeSettings(next, { pushHistory: true });
  }

  function applyTimePreset(presetName) {
    if (!isAdmin) return;

    let next = null;

    if (presetName === 'early') {
      next = normalizeTimeSettings({
        intervalMinutes: 60,
        startMinute: 1 * 60,
        endMinute: 4 * 60,
        fullDay: false
      });
    } else if (presetName === 'workday') {
      next = normalizeTimeSettings({
        intervalMinutes: 60,
        startMinute: 9 * 60,
        endMinute: 17 * 60,
        fullDay: false
      });
    } else if (presetName === 'full') {
      next = normalizeTimeSettings({
        intervalMinutes: timeSettings.intervalMinutes,
        fullDay: true
      });
    }

    if (!next) return;
    applyTimeSettings(next, { pushHistory: true });
  }

  function applyEventChip(el) {
    const value = (el.value || '').trim().toLowerCase();
    el.classList.remove('event-chip', 'event-deep', 'event-meeting', 'event-lunch', 'event-break', 'event-project');

    if (!value) return;
    el.classList.add('event-chip');

    if (value.includes('deep')) el.classList.add('event-deep');
    else if (value.includes('meeting') || value.includes('standup') || value.includes('1:1')) el.classList.add('event-meeting');
    else if (value.includes('lunch')) el.classList.add('event-lunch');
    else if (value.includes('break')) el.classList.add('event-break');
    else if (value.includes('project') || value.includes('workshop')) el.classList.add('event-project');
  }

  function selectEditableCell(el) {
    clearRowSelection();
    clearColumnSelection();
    if (selectedEditableCell && selectedEditableCell !== el) {
      selectedEditableCell.classList.remove('selected-cell');
    }
    selectedEditableCell = el;
    selectedEditableCell.classList.add('selected-cell');
    updateThemePresetActiveState();
  }

  function getColorActionTargets() {
    const getCellColumnIndex = (el) => {
      const td = el?.closest('td');
      if (!td || !td.parentElement) return -1;
      return Array.from(td.parentElement.children).indexOf(td);
    };

    if (selectedEditableCell) {
      const colIndex = getCellColumnIndex(selectedEditableCell);
      if (colIndex <= 0) {
        return [];
      }
      return [selectedEditableCell];
    }

    const activeEl = document.activeElement;
    if (activeEl && activeEl.matches('#tbody input, #tbody textarea')) {
      const colIndex = getCellColumnIndex(activeEl);
      if (colIndex > 0) {
        selectEditableCell(activeEl);
        return [activeEl];
      }
    }

    if (selectedColumnIndex !== null) {
      if (selectedColumnIndex <= 0) {
        return [];
      }

      return Array.from(document.querySelectorAll('#tbody tr'))
        .map((row) => row.children[selectedColumnIndex]?.querySelector('input, textarea'))
        .filter(Boolean);
    }

    if (selectedRowIndex !== null) {
      const row = document.querySelectorAll('#tbody tr')[selectedRowIndex];
      if (!row) return [];

      return Array.from(row.querySelectorAll('td input, td textarea'))
        .filter((el) => getCellColumnIndex(el) > 0);
    }

    return Array.from(document.querySelectorAll('#tbody tr'))
      .flatMap((row) => Array.from(row.querySelectorAll('td input, td textarea')))
      .filter((el) => getCellColumnIndex(el) > 0);
  }

  function applyCustomColorToSelected() {
    if (!isAdmin) return;

    const targets = getColorActionTargets();
    if (!targets.length) {
      alert('Select a day cell, row, or column (not the Time column).');
      return;
    }

    const color = document.getElementById('cell-color-picker').value;
    pushHistorySnapshot();
    targets.forEach((el) => {
      el.classList.add('custom-cell-color');
      el.style.setProperty('--custom-cell-color', color);
    });
  }

  function getSelectedRowIndexForColoring() {
    if (selectedRowIndex !== null) {
      return selectedRowIndex;
    }

    if (selectedEditableCell) {
      const row = selectedEditableCell.closest('tr');
      const rows = Array.from(document.querySelectorAll('#tbody tr'));
      const rowIndex = rows.indexOf(row);
      return rowIndex >= 0 ? rowIndex : null;
    }

    return null;
  }

  function getSelectedColumnIndexForColoring() {
    if (selectedColumnIndex !== null) {
      return selectedColumnIndex;
    }

    if (selectedEditableCell) {
      const td = selectedEditableCell.closest('td');
      if (td && td.parentElement) {
        const index = Array.from(td.parentElement.children).indexOf(td);
        if (index >= 0) {
          return index;
        }
      }
    }

    return null;
  }

  function applyCustomColorToSelectedRow() {
    if (!isAdmin) return;

    const rowIndex = getSelectedRowIndexForColoring();
    if (rowIndex === null) {
      alert('Select a row or a cell in that row first.');
      return;
    }

    const row = document.querySelectorAll('#tbody tr')[rowIndex];
    if (!row) return;

    const color = document.getElementById('cell-color-picker').value;
    const targets = Array.from(row.querySelectorAll('td input, td textarea')).filter((el) => {
      const td = el.closest('td');
      if (!td || !td.parentElement) return false;
      const columnIndex = Array.from(td.parentElement.children).indexOf(td);
      return columnIndex > 0;
    });

    if (!targets.length) return;

    pushHistorySnapshot();
    targets.forEach((el) => {
      el.classList.add('custom-cell-color');
      el.style.setProperty('--custom-cell-color', color);
    });
  }

  function applyCustomColorToSelectedColumn() {
    if (!isAdmin) return;

    const columnIndex = getSelectedColumnIndexForColoring();
    if (columnIndex === null || columnIndex <= 0) {
      alert('Select a day column or a cell in that column first.');
      return;
    }

    const color = document.getElementById('cell-color-picker').value;
    const targets = Array.from(document.querySelectorAll('#tbody tr'))
      .map((row) => row.children[columnIndex]?.querySelector('input, textarea'))
      .filter(Boolean);

    if (!targets.length) return;

    pushHistorySnapshot();
    targets.forEach((el) => {
      el.classList.add('custom-cell-color');
      el.style.setProperty('--custom-cell-color', color);
    });
  }

  function normalizeColorPresets(rawPresets) {
    if (!Array.isArray(rawPresets)) {
      return [...DEFAULT_COLOR_PRESETS];
    }

    const normalized = rawPresets
      .slice(0, DEFAULT_COLOR_PRESETS.length)
      .map((value, index) => {
        const candidate = String(value || '').trim();
        if (/^#[0-9a-fA-F]{6}$/.test(candidate)) {
          return candidate;
        }
        return DEFAULT_COLOR_PRESETS[index];
      });

    while (normalized.length < DEFAULT_COLOR_PRESETS.length) {
      normalized.push(DEFAULT_COLOR_PRESETS[normalized.length]);
    }

    return normalized;
  }

  function renderColorPresets() {
    const presetButtons = document.querySelectorAll('.color-preset-btn[data-preset-index]');
    presetButtons.forEach((button) => {
      const index = Number(button.dataset.presetIndex);
      const presetColor = colorPresets[index] || DEFAULT_COLOR_PRESETS[index] || '#dbeafe';
      button.style.setProperty('--preset-color', presetColor);
      button.title = `Preset color ${index + 1}: ${presetColor}`;
    });
  }

  function applyColorPreset(presetOrColor) {
    if (!isAdmin) return;

    let selectedColor = null;
    if (typeof presetOrColor === 'number') {
      selectedColor = colorPresets[presetOrColor] || null;
    } else if (typeof presetOrColor === 'string') {
      selectedColor = presetOrColor;
    }

    if (!selectedColor) return;

    const picker = document.getElementById('cell-color-picker');
    if (picker) {
      picker.value = selectedColor;
    }
    applyCustomColorToSelected();
  }

  function applyCellsPatternPresetFromPicker() {
    if (!isAdmin) return;

    const presetSelect = document.getElementById('color-preset-select');
    if (!presetSelect) return;

    const presetIndex = Number(presetSelect.value);
    applyThemePresetToAllCells(presetIndex);
  }

  function setColorPresetFromPicker() {
    applyCellsPatternPresetFromPicker();
  }

  function applyPageThemeFromPicker() {
    setColorPresetFromPicker();
  }

  function resetColorPresets() {
    if (!isAdmin) return;
    pushHistorySnapshot();
    colorPresets = [...DEFAULT_COLOR_PRESETS];
    activeCellsPresetIndex = 0;
    columnThemePresetIndices = {};
    renderColorPresets();
    renderColumnThemePresets();
    applyThemePreset(0, { pushHistory: false });
    clearAllCustomColorsInternal();

    const picker = document.getElementById('cell-color-picker');
    if (picker) {
      picker.value = colorPresets[0];
    }

    saveStateForCurrentUser();
  }

  function clearSelectedColor() {
    if (!isAdmin) return;

    const targets = getColorActionTargets();
    if (!targets.length) {
      alert('Select a day cell, row, or column (not the Time column).');
      return;
    }

    pushHistorySnapshot();
    targets.forEach((el) => {
      el.classList.remove('custom-cell-color');
      el.style.removeProperty('--custom-cell-color');
    });
  }

  function clearAllCustomColors() {
    if (!isAdmin) return;
    pushHistorySnapshot();
    document.querySelectorAll('#tbody input, #tbody textarea').forEach(el => {
      el.classList.remove('custom-cell-color');
      el.style.removeProperty('--custom-cell-color');
    });
  }

  function getCustomColorEntries() {
    const rows = Array.from(document.querySelectorAll('#tbody tr'));
    const colors = [];

    rows.forEach((row, rowIndex) => {
      const cells = Array.from(row.querySelectorAll('td input, td textarea'));
      cells.forEach((el, colIndex) => {
        const color = el.style.getPropertyValue('--custom-cell-color');
        if (color) {
          colors.push({ row: rowIndex, col: colIndex, color });
        }
      });
    });

    return colors;
  }

  function cloneUserCalendarEvents(events = {}) {
    return JSON.parse(JSON.stringify(events));
  }

  function captureState() {
    const headerInputs = Array.from(document.querySelectorAll('#header-row th input'));
    const rowElements = Array.from(document.querySelectorAll('#tbody tr'));
    const rows = rowElements.map(row => {
      const cells = Array.from(row.querySelectorAll('td input, td textarea'));
      return cells.map(cell => cell.value);
    });

    return {
      title: document.getElementById('main-title').value,
      subtitle: document.querySelector('.subtitle-input').value,
      weekOffset,
      headerLabels: headerInputs.map(h => h.value),
      rows,
      timeSettings: { ...timeSettings },
      colorPresets: [...colorPresets],
      themePresetIndex: activeThemePresetIndex,
      cellsPresetIndex: activeCellsPresetIndex,
      columnThemePresetIndices: { ...columnThemePresetIndices },
      customColors: getCustomColorEntries(),
      userCalendarEvents: cloneUserCalendarEvents(userCalendarEvents),
      nowLineTheme
    };
  }

  function applyCustomColors(customColors = []) {
    clearAllCustomColorsInternal();

    customColors.forEach(entry => {
      const row = document.querySelectorAll('#tbody tr')[entry.row];
      if (!row) return;
      const el = row.querySelectorAll('td input, td textarea')[entry.col];
      if (!el) return;
      el.classList.add('custom-cell-color');
      el.style.setProperty('--custom-cell-color', entry.color);
    });
  }

  function clearAllCustomColorsInternal() {
    document.querySelectorAll('#tbody input, #tbody textarea').forEach(el => {
      el.classList.remove('custom-cell-color');
      el.style.removeProperty('--custom-cell-color');
    });
  }

  function restoreState(state) {
    if (!state) return;

    document.getElementById('main-title').value = state.title || '';
    document.querySelector('.subtitle-input').value = state.subtitle || '';
    weekOffset = Number.isFinite(state.weekOffset) ? state.weekOffset : 0;

    const tbody = document.getElementById('tbody');
    tbody.innerHTML = '';

    const rows = Array.isArray(state.rows) && state.rows.length ? state.rows : defaultRows;
    const restoredTimeSettings = state.timeSettings
      ? normalizeTimeSettings(state.timeSettings)
      : inferTimeSettingsFromRows(rows);
    timeSettings = restoredTimeSettings;
    colorPresets = normalizeColorPresets(state.colorPresets);
    renderColorPresets();
    activeThemePresetIndex = normalizeThemePresetIndex(state.themePresetIndex);
    columnThemePresetIndices = normalizeColumnThemePresetIndices(state.columnThemePresetIndices);
    if (Number.isInteger(Number(state.cellsPresetIndex))) {
      activeCellsPresetIndex = normalizeThemePresetIndex(state.cellsPresetIndex);
    } else {
      const dataColumnEntry = Object.entries(columnThemePresetIndices).find(([columnKey]) => Number(columnKey) > 0);
      activeCellsPresetIndex = dataColumnEntry
        ? normalizeThemePresetIndex(dataColumnEntry[1])
        : 0;
    }
    applyThemePreset(activeThemePresetIndex, { pushHistory: false });
    setTimeSettingsUIState();
    rebuildRowsForTimeSettings(rows);
    renderColumnThemePresets();

    clearCellSelection();
    clearRowSelection();
    clearColumnSelection();
    bindHeaderSelection();

    renderWeekLabel();

    const headerInputs = Array.from(document.querySelectorAll('#header-row th input'));
    if (Array.isArray(state.headerLabels) && state.headerLabels.length === headerInputs.length) {
      headerInputs.forEach((input, i) => {
        input.value = state.headerLabels[i];
      });
    }

    applyCustomColors(Array.isArray(state.customColors) ? state.customColors : []);
    userCalendarEvents = cloneUserCalendarEvents(state.userCalendarEvents || {});
    setNowLineTheme(state.nowLineTheme || 'cyan');
    renderCalendar();
    updateTodayHighlight();
    updateNowLine();
  }

  function setNowLineTheme(theme) {
    nowLineTheme = theme === 'violet' ? 'violet' : 'cyan';
    const nowLine = document.getElementById('now-line');
    if (!nowLine) return;

    nowLine.classList.toggle('theme-violet', nowLineTheme === 'violet');
    nowLine.classList.toggle('theme-cyan', nowLineTheme === 'cyan');

    const themeSelect = document.getElementById('now-line-theme');
    if (themeSelect && themeSelect.value !== nowLineTheme) {
      themeSelect.value = nowLineTheme;
    }
  }

  function pushHistorySnapshot() {
    changeHistory.push(captureState());
    if (changeHistory.length > MAX_HISTORY) {
      changeHistory.shift();
    }
    saveStateForCurrentUser();
  }

  function undoPreviousChange() {
    if (!isAdmin) return;
    const previous = changeHistory.pop();
    if (!previous) return;
    restoreState(previous);
  }

  function clearDefaultPlaceholders() {
    if (!isAdmin) return;
    pushHistorySnapshot();

    document.getElementById('main-title').value = '';
    document.querySelector('.subtitle-input').value = '';

    const rows = Array.from(document.querySelectorAll('#tbody tr'));
    rows.forEach((row) => {
      const cells = Array.from(row.querySelectorAll('td input, td textarea'));
      cells.forEach((cell, colIndex) => {
        if (colIndex === 0) return;
        cell.value = '';
        applyEventChip(cell);
      });
    });

    clearAllCustomColorsInternal();
    updateNowLine();
  }

  function clearAllChanges() {
    if (!isAdmin) return;
    if (!confirm('Clear all changes and reset to default timetable?')) return;
    pushHistorySnapshot();
    changeHistory.length = 0;
    restoreState(INITIAL_STATE);
  }

  function switchView(viewName) {
    activeView = viewName === 'calendar' ? 'calendar' : 'timetable';

    document.getElementById('timetable-panel').classList.toggle('active', activeView === 'timetable');
    document.getElementById('calendar-panel').classList.toggle('active', activeView === 'calendar');

    document.getElementById('timetable-view-btn').classList.toggle('active', activeView === 'timetable');
    document.getElementById('calendar-view-btn').classList.toggle('active', activeView === 'calendar');

    if (activeView === 'calendar') {
      renderCalendar();
    }

    updateNowLine();
  }

  function getDetectedCountryCode() {
    try {
      const locale = navigator.language || 'en-US';
      const region = new Intl.Locale(locale).maximize().region;
      if (region) return region.toUpperCase();
    } catch (error) {
      // fallback below
    }
    return 'US';
  }

  function buildCountrySelect() {
    const select = document.getElementById('country-select');
    if (!select) return;

    const displayNames = typeof Intl.DisplayNames === 'function'
      ? new Intl.DisplayNames([navigator.language || 'en'], { type: 'region' })
      : null;

    select.innerHTML = '';
    countryOptions.forEach(code => {
      const option = document.createElement('option');
      option.value = code;
      option.textContent = `${displayNames ? displayNames.of(code) || code : code} (${code})`;
      select.appendChild(option);
    });

    const detected = getDetectedCountryCode();
    selectedCountryCode = countryOptions.includes(detected) ? detected : 'US';
    select.value = selectedCountryCode;
  }

  function onCountryChange() {
    const select = document.getElementById('country-select');
    selectedCountryCode = select?.value || 'US';
    loadPublicHolidays(calendarMonthDate.getFullYear());
  }

  function setHolidayStatus(text, isError = false) {
    const status = document.getElementById('holiday-status');
    if (!status) return;
    status.textContent = text;
    status.style.color = isError ? '#b42318' : '#888';
  }

  function normalizeCalendarEventType(value) {
    const raw = (value || '').trim().toLowerCase();
    if (raw === 'birthday' || raw === 'bday') return 'birthday';
    return 'event';
  }

  function getUserEventsForDate(isoKey) {
    if (!Array.isArray(userCalendarEvents[isoKey])) {
      userCalendarEvents[isoKey] = [];
    }
    return userCalendarEvents[isoKey];
  }

  function addCalendarEvent(isoKey) {
    if (!isAdmin) return;

    const title = (prompt('Event title (e.g. John birthday):', '') || '').trim();
    if (!title) return;

    const typeInput = prompt('Event type: birthday or event', 'event');
    const type = normalizeCalendarEventType(typeInput);

    pushHistorySnapshot();
    getUserEventsForDate(isoKey).push({
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      title,
      type,
      allDay: true
    });

    renderCalendar();
  }

  function editCalendarEvent(isoKey, eventId) {
    if (!isAdmin) return;
    const eventItem = getUserEventsForDate(isoKey).find(item => item.id === eventId);
    if (!eventItem) return;

    const nextTitle = (prompt('Edit event title:', eventItem.title) || '').trim();
    if (!nextTitle) return;

    const nextTypeInput = prompt('Event type: birthday or event', eventItem.type || 'event');
    const nextType = normalizeCalendarEventType(nextTypeInput);

    pushHistorySnapshot();
    eventItem.title = nextTitle;
    eventItem.type = nextType;
    renderCalendar();
  }

  function removeCalendarEvent(isoKey, eventId) {
    if (!isAdmin) return;
    const events = getUserEventsForDate(isoKey);
    const index = events.findIndex(item => item.id === eventId);
    if (index === -1) return;

    pushHistorySnapshot();
    events.splice(index, 1);
    if (!events.length) {
      delete userCalendarEvents[isoKey];
    }
    renderCalendar();
  }

  function getIsoDateKey(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  async function loadPublicHolidays(year) {
    setHolidayStatus(`Loading ${year} holidays...`);

    try {
      const response = await fetch(`https://date.nager.at/api/v3/PublicHolidays/${year}/${selectedCountryCode}`);
      if (!response.ok) throw new Error(`Holiday API returned ${response.status}`);

      const data = await response.json();
      holidayList = Array.isArray(data) ? data : [];
      holidayDates = new Map();

      holidayList.forEach(item => {
        if (!item?.date) return;
        if (!holidayDates.has(item.date)) {
          holidayDates.set(item.date, []);
        }
        holidayDates.get(item.date).push(item.localName || item.name || 'Holiday');
      });

      setHolidayStatus(`Loaded ${holidayList.length} public holidays for ${selectedCountryCode}`);
    } catch (error) {
      holidayList = [];
      holidayDates = new Map();
      setHolidayStatus('Could not load holidays right now', true);
    }

    renderCalendar();
  }

  function updateMonthButtons() {
    const prevBtn = document.getElementById('prev-month-btn');
    const nextBtn = document.getElementById('next-month-btn');
    if (!prevBtn || !nextBtn) return;

    prevBtn.disabled = calendarMonthDate.getMonth() === 0;
    nextBtn.disabled = calendarMonthDate.getMonth() === 11;
  }

  function changeMonth(delta) {
    const next = new Date(calendarMonthDate);
    next.setMonth(next.getMonth() + Number(delta));
    if (next.getFullYear() !== SCHEDULE_YEAR) return;

    calendarMonthDate = new Date(next.getFullYear(), next.getMonth(), 1);
    renderCalendar();
  }

  function renderCalendar() {
    const grid = document.getElementById('calendar-grid');
    if (!grid) return;

    const title = document.getElementById('calendar-title');
    const countryInfo = document.getElementById('country-info');
    const holidayCount = document.getElementById('holiday-count');

    title.textContent = calendarMonthDate.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
    countryInfo.textContent = `Country: ${selectedCountryCode}`;

    const monthStart = new Date(calendarMonthDate.getFullYear(), calendarMonthDate.getMonth(), 1);
    const monthEnd = new Date(calendarMonthDate.getFullYear(), calendarMonthDate.getMonth() + 1, 0);
    const startDay = (monthStart.getDay() + 6) % 7;

    const daysInMonth = monthEnd.getDate();
    const totalCells = Math.ceil((startDay + daysInMonth) / 7) * 7;
    const todayKey = getIsoDateKey(new Date());

    const visibleHolidayCount = holidayList.filter(item => {
      if (!item?.date) return false;
      const date = new Date(item.date);
      return date.getFullYear() === calendarMonthDate.getFullYear() && date.getMonth() === calendarMonthDate.getMonth();
    }).length;

    holidayCount.textContent = `Public Holidays: ${visibleHolidayCount}`;

    const dayHeaders = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    grid.innerHTML = dayHeaders.map(day => `<div class="calendar-dayhead">${day}</div>`).join('');

    for (let cell = 0; cell < totalCells; cell++) {
      const date = new Date(monthStart);
      date.setDate(1 - startDay + cell);

      const inMonth = date.getMonth() === calendarMonthDate.getMonth();
      const isoKey = getIsoDateKey(date);
      const holidayName = holidayDates.get(isoKey);

      const dayCell = document.createElement('div');
      dayCell.className = 'calendar-cell';
      if (!inMonth) dayCell.classList.add('outside');
      if (isoKey === todayKey) dayCell.classList.add('today');
      if (holidayName) dayCell.classList.add('holiday');

      const dateLabel = document.createElement('div');
      dateLabel.className = 'calendar-date';
      dateLabel.textContent = String(date.getDate());
      dayCell.appendChild(dateLabel);

      if (holidayName) {
        const holidayTag = document.createElement('div');
        holidayTag.className = 'holiday-tag';
        holidayTag.textContent = holidayName;
        dayCell.appendChild(holidayTag);
      }

      grid.appendChild(dayCell);
    }

    updateMonthButtons();
  }

  function getMondayOfWeek(date) {
    const d = new Date(date);
    const day = d.getDay() || 7;
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() - day + 1);
    return d;
  }

  function formatShortDate(date) {
    return date.toLocaleDateString(undefined, { day: '2-digit', month: 'short' });
  }

  function getYearRangeMondays(year) {
    const jan1 = new Date(year, 0, 1);
    const dec31 = new Date(year, 11, 31);

    const start = getMondayOfWeek(jan1);
    const end = getMondayOfWeek(dec31);

    return { start, end };
  }

  function isWeekInYear(monday, year) {
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    return monday.getFullYear() === year || sunday.getFullYear() === year;
  }

  function getDisplayedMonday() {
    const base = new Date();
    base.setDate(base.getDate() + (weekOffset * 7));
    return getMondayOfWeek(base);
  }

  function updateWeekButtons() {
    const { start, end } = getYearRangeMondays(SCHEDULE_YEAR);
    const monday = getDisplayedMonday();
    const prevBtn = document.getElementById('prev-week-btn');
    const nextBtn = document.getElementById('next-week-btn');

    if (!prevBtn || !nextBtn) return;

    prevBtn.disabled = monday <= start;
    nextBtn.disabled = monday >= end;
  }

  function updateWeekDayHeaders() {
    const monday = getDisplayedMonday();
    const headerInputs = document.querySelectorAll('#header-row th input');
    const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

    if (headerInputs.length < 8) return;

    headerInputs[0].value = 'Time';

    for (let i = 0; i < 7; i++) {
      const dayDate = new Date(monday);
      dayDate.setDate(monday.getDate() + i);
      headerInputs[i + 1].value = `${dayNames[i]} ${formatShortDate(dayDate)}`;
    }
  }

  function clampWeekToYear() {
    const { start, end } = getYearRangeMondays(SCHEDULE_YEAR);
    let monday = getDisplayedMonday();

    if (monday < start) {
      while (monday < start) {
        weekOffset += 1;
        monday = getDisplayedMonday();
      }
    } else if (monday > end) {
      while (monday > end) {
        weekOffset -= 1;
        monday = getDisplayedMonday();
      }
    }
  }

  function renderWeekLabel() {
    const monday = getDisplayedMonday();
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);

    const label = document.getElementById('week-label');
    label.textContent = `${formatShortDate(monday)} - ${formatShortDate(sunday)} (${SCHEDULE_YEAR})`;

    updateWeekDayHeaders();
    updateWeekButtons();
  }

  function changeWeek(delta) {
    weekOffset += Number(delta);
    clampWeekToYear();
    renderWeekLabel();
    updateTodayHighlight();
    updateNowLine();
  }

  function goToCurrentWeek() {
    weekOffset = 0;
    clampWeekToYear();
    renderWeekLabel();
    updateTodayHighlight();
    updateNowLine();
  }

  function getDayColumnForDisplayedWeek() {
    const now = new Date();
    const viewDate = new Date(now);
    viewDate.setDate(now.getDate() + (weekOffset * 7));
    const weekMonday = getMondayOfWeek(viewDate);

    if (!isWeekInYear(weekMonday, SCHEDULE_YEAR)) return null;

    const day = viewDate.getDay();
    if (day === 0) return 8;
    return day + 1;
  }

  function updateTodayHighlight() {
    const dayColumn = getDayColumnForDisplayedWeek();

    document.querySelectorAll('#header-row th').forEach(th => th.classList.remove('today-header'));
    document.querySelectorAll('#tbody td').forEach(td => td.classList.remove('today-col'));

    if (!dayColumn) return;

    const headerCell = document.querySelector(`#header-row th:nth-child(${dayColumn})`);
    if (headerCell) headerCell.classList.add('today-header');

    document.querySelectorAll('#tbody tr').forEach(row => {
      const td = row.children[dayColumn - 1];
      if (td) td.classList.add('today-col');
    });
  }

  function getVisibleDayColumn() {
    return getDayColumnForDisplayedWeek();
  }

  function updateNowLine() {
    const nowLine = document.getElementById('now-line');
    if (activeView !== 'timetable') {
      if (nowLine) nowLine.style.display = 'none';
      return;
    }

    const tableWrap = document.querySelector('.table-wrap');
    if (!nowLine || !tableWrap) {
      nowLine.style.display = 'none';
      return;
    }

    const now = new Date();
    const currentMinutes = now.getHours() * 60 + now.getMinutes() + (now.getSeconds() / 60);
    nowLine.setAttribute('data-time', formatMinutesTo12h(currentMinutes));
    const visibleStart = timeSettings.startMinute;
    const visibleDuration = timeSettings.fullDay
      ? 24 * 60
      : (timeSettings.endMinute - timeSettings.startMinute + timeSettings.intervalMinutes);

    if (currentMinutes < visibleStart || currentMinutes > visibleStart + visibleDuration) {
      nowLine.style.display = 'none';
      return;
    }

    const tbody = document.getElementById('tbody');
    if (!tbody || !tbody.offsetHeight) {
      nowLine.style.display = 'none';
      return;
    }

    const dayRatio = (currentMinutes - visibleStart) / visibleDuration;
    const top = tbody.offsetTop + (dayRatio * tbody.offsetHeight);
    const left = 0;
    const width = tableWrap.scrollWidth;

    nowLine.style.top = `${top}px`;
    nowLine.style.left = `${left}px`;
    nowLine.style.width = `${width}px`;
    nowLine.style.display = 'block';
  }

  function startNowLine() {
    if (nowLineTimer) clearInterval(nowLineTimer);
    updateNowLine();
    nowLineTimer = setInterval(updateNowLine, 1000);
    window.addEventListener('resize', updateNowLine);
    document.querySelector('.table-wrap').addEventListener('scroll', updateNowLine);
  }

  // Init
  applyThemePreset(0, { pushHistory: false });
  colorPresets = normalizeColorPresets(DEFAULT_COLOR_PRESETS);
  activeCellsPresetIndex = normalizeThemePresetIndex(INITIAL_STATE.cellsPresetIndex);
  columnThemePresetIndices = normalizeColumnThemePresetIndices(INITIAL_STATE.columnThemePresetIndices);
  renderColorPresets();
  buildTimeSelectOptions();
  timeSettings = normalizeTimeSettings(DEFAULT_TIME_SETTINGS);
  setTimeSettingsUIState();
  rebuildRowsForTimeSettings(defaultRows);
  bindHeaderSelection();
  pushHistorySnapshot();
  setEditMode(false);
  buildCountrySelect();
  renderWeekLabel();
  updateTodayHighlight();
  loadPublicHolidays(SCHEDULE_YEAR);
  renderCalendar();
  switchView('timetable');
  setNowLineTheme('cyan');
  startNowLine();
  setAuthMode('signin');
  window.addEventListener('beforeunload', saveStateForCurrentUser);
  document.getElementById('app-email').addEventListener('keydown', handleAuthEnter);
  document.getElementById('app-password').addEventListener('keydown', handleAuthEnter);
  document.getElementById('app-confirm-password').addEventListener('keydown', handleAuthEnter);
  initSupabaseAuth();