  const SUPABASE_URL = 'https://zdgksbfvzrobeotekzpx.supabase.co';
  const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpkZ2tzYmZ2enJvYmVvdGVrenB4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzIwMTMxNDgsImV4cCI6MjA4NzU4OTE0OH0.tITnaa2hczGaqTjQOrX7bUmLFyjPBRU-oDfL1Yu97Q0';
  const CUSTOMIZE_PASSWORD = 'admin123';
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
  let authMode = 'signin';
  let currentUserId = null;

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
    customColors: [],
    userCalendarEvents: {},
    nowLineTheme: 'cyan'
  };

  function getColCount() {
    const headerRow = document.getElementById('header-row');
    return headerRow.querySelectorAll('th').length - 1; // exclude action col
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

  function toggleSettingsPanel() {
    const panel = document.getElementById('settings-panel');
    const toggleBtn = document.getElementById('settings-toggle-btn');
    if (!panel || !toggleBtn) return;

    const isOpen = panel.classList.toggle('open');
    toggleBtn.textContent = isOpen ? '✕ Close' : '⚙ Customize';
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
      setEditMode(true);
      document.getElementById('customize-password').value = '';
      return;
    }

    alert('Invalid customize password.');
    isAdmin = false;
    setEditMode(false);
  }

  function lockCustomize() {
    isAdmin = false;
    setEditMode(false);
  }

  function makeCell(value = '', isFirst = false) {
    const td = document.createElement('td');
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.value = value;
    inp.spellcheck = false;
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

    bindHeaderSelection();
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

  function formatMinutesTo12h(minutes) {
    const wholeMinutes = Math.floor(minutes);
    const h24 = Math.floor(wholeMinutes / 60);
    const mm = wholeMinutes % 60;
    const seconds = Math.floor((minutes - wholeMinutes) * 60);
    const suffix = h24 >= 12 ? 'PM' : 'AM';
    const h12 = h24 % 12 || 12;
    return `${h12}:${String(mm).padStart(2, '0')}:${String(seconds).padStart(2, '0')} ${suffix}`;
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
  }

  function applyCustomColorToSelected() {
    if (!isAdmin) return;

    if (!selectedEditableCell) {
      const activeEl = document.activeElement;
      if (activeEl && activeEl.matches('#tbody input, #tbody textarea')) {
        selectEditableCell(activeEl);
      }
    }

    if (!selectedEditableCell) return;

    pushHistorySnapshot();
    const color = document.getElementById('cell-color-picker').value;
    selectedEditableCell.classList.add('custom-cell-color');
    selectedEditableCell.style.setProperty('--custom-cell-color', color);
  }

  function clearSelectedColor() {
    if (!isAdmin || !selectedEditableCell) return;
    pushHistorySnapshot();
    selectedEditableCell.classList.remove('custom-cell-color');
    selectedEditableCell.style.removeProperty('--custom-cell-color');
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
    rows.forEach(row => addRow(row));

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
    rows.forEach((row, rowIndex) => {
      const cells = Array.from(row.querySelectorAll('td input, td textarea'));
      cells.forEach((cell, colIndex) => {
        const defaultValue = defaultRows[rowIndex]?.[colIndex] || '';
        if ((cell.value || '') === defaultValue || colIndex > 0) {
          cell.value = colIndex === 0 ? (defaultRows[rowIndex]?.[0] || '') : '';
          applyEventChip(cell);
        }
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
    const dayRatio = currentMinutes / 1440;
    const top = dayRatio * tableWrap.scrollHeight;
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
  defaultRows.forEach(row => addRow(row));
  bindHeaderSelection();
  pushHistorySnapshot();
  setEditMode(true);
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