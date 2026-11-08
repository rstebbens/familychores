/**
 * Family Chores backend.
 * Runs as a Google Apps Script web application bound to the family spreadsheet.
 * Parent PIN is stored as a salted hash in Script Properties.
 */

const SHEET_ID = '1zKPw9gjkloEmlNGhQlxpZAlA3Wy5_FA7qw8JiNNfbxk';
const SHEETS = {
  log: 'Daily Log',
  chores: 'Chore Settings',
  leave: 'Leave Days',
  bonuses: 'Bonus Tasks'
};
const ZONE = 'Europe/London';

function tab(name) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sheet) throw Error('Missing tab: ' + name);
  return sheet;
}

function digest(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value)
    .map(byte => ('0' + (byte & 255).toString(16)).slice(-2))
    .join('');
}

function setParentPin() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Parent PIN',
    'Enter a new 6-8 digit PIN',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const pin = response.getResponseText().trim();
  if (!/^\d{6,8}$/.test(pin)) {
    ui.alert('Use 6-8 digits');
    return;
  }

  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({
    PIN_SALT: salt,
    PIN_HASH: digest(salt + ':' + pin)
  });

  ui.alert('PIN saved.');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Family Chores')
    .addItem('Set/change parent PIN', 'setParentPin')
    .addToUi();
}

function auth(pin) {
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty('PIN_SALT');
  const expectedHash = properties.getProperty('PIN_HASH');

  if (!expectedHash || digest(salt + ':' + String(pin)) !== expectedHash) {
    throw Error('Incorrect PIN');
  }
}

function today() {
  return Utilities.formatDate(new Date(), ZONE, 'yyyy-MM-dd');
}

function dateOk(value) {
  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(value) ||
    isNaN(Date.parse(value)) ||
    new Date(value + 'T12:00:00Z').toISOString().slice(0, 10) !== value
  ) {
    throw Error('Invalid date');
  }

  return value;
}

function chores() {
  return tab(SHEETS.chores)
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(row =>
      row[0] &&
      (row[4] === true || String(row[4]).toLowerCase() === 'true')
    )
    .map(row => ({
      id: String(row[0]),
      name: String(row[1]),
      weight: Number(row[3])
    }));
}

function entries() {
  return tab(SHEETS.log)
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(row => row[0])
    .map((row, index) => ({
      row: index + 2,
      id: String(row[0]),
      date: row[1] instanceof Date
        ? Utilities.formatDate(row[1], ZONE, 'yyyy-MM-dd')
        : String(row[1]),
      child: String(row[2]),
      choreId: String(row[3]),
      chore: String(row[4]),
      status: String(row[5]),
      submittedAt: String(row[6] || ''),
      reviewedAt: String(row[7] || '')
    }));
}


function optionalTab(name, headers) {
  const book = SpreadsheetApp.openById(SHEET_ID);
  let sheet = book.getSheetByName(name);
  if (!sheet) {
    sheet = book.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}
function leaveSheet() { return optionalTab(SHEETS.leave, ['Date', 'Child', 'Created At']); }
function bonusSheet() { return optionalTab(SHEETS.bonuses, ['ID', 'Child', 'Task', 'Amount', 'Status', 'Destination', 'Created At', 'Submitted At', 'Reviewed At']); }
function leaveDates(month) {
  return leaveSheet().getDataRange().getValues().slice(1)
    .filter(row => String(row[1]) === 'James')
    .map(row => row[0] instanceof Date ? Utilities.formatDate(row[0], ZONE, 'yyyy-MM-dd') : String(row[0]))
    .filter(date => date.startsWith(month));
}
function bonusTasks(month) {
  return bonusSheet().getDataRange().getValues().slice(1).filter(row => row[0] && String(row[1]) === 'James')
    .map(row => ({id:String(row[0]), name:String(row[2]), amount:Number(row[3]), status:String(row[4]), destination:String(row[5] || ''), created:String(row[6] || '')}))
    .filter(task => task.created.slice(0, 7) === month);
}
function apiSetLeave(pin, date, enabled) {
  auth(pin);
  dateOk(date);
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sheet = leaveSheet();
    const rows = sheet.getDataRange().getValues();
    const matches = [];
    for (let i = 1; i < rows.length; i++) {
      const d = rows[i][0] instanceof Date ? Utilities.formatDate(rows[i][0], ZONE, 'yyyy-MM-dd') : String(rows[i][0]);
      if (d === date && String(rows[i][1]) === 'James') matches.push(i + 1);
    }
    if (enabled && !matches.length) sheet.appendRow([date, 'James', new Date().toISOString()]);
    if (!enabled) for (const row of matches.reverse()) sheet.deleteRow(row);
    return {success:true};
  } finally { lock.releaseLock(); }
}
function apiCreateBonus(pin, name, amount) {
  auth(pin);
  name = String(name || '').trim();
  amount = Number(amount);
  if (!name || name.length > 120 || !Number.isFinite(amount) || amount <= 0 || amount > 100) throw Error('Enter a task and an amount between £0.01 and £100.');
  bonusSheet().appendRow([Utilities.getUuid(), 'James', name, Math.round(amount * 100) / 100, 'Open', '', new Date().toISOString(), '', '']);
  return {success:true};
}
function apiSubmitBonus(id, destination) {
  if (!['Cash','Savings'].includes(destination)) throw Error('Choose cash or savings');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sheet = bonusSheet(), rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id) && String(rows[i][1]) === 'James') {
        if (String(rows[i][4]) !== 'Open') throw Error('Already submitted');
        sheet.getRange(i + 1, 5, 1, 4).setValues([['Pending', destination, rows[i][6], new Date().toISOString()]]);
        return {success:true};
      }
    }
    throw Error('Bonus task not found');
  } finally { lock.releaseLock(); }
}
function apiReviewBonus(pin, id, approved) {
  auth(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sheet = bonusSheet(), rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id) && String(rows[i][1]) === 'James') {
        if (String(rows[i][4]) !== 'Pending') throw Error('Bonus is not pending');
        sheet.getRange(i + 1, 5).setValue(approved ? 'Approved' : 'Rejected');
        sheet.getRange(i + 1, 9).setValue(new Date().toISOString());
        return {success:true};
      }
    }
    throw Error('Bonus not found');
  } finally { lock.releaseLock(); }
}

function monthData(month, privileged) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
    throw Error('Invalid month');
  }

  const monthlyEntries = entries().filter(entry =>
    entry.child === 'James' && entry.date.startsWith(month)
  );
  const activeChores = chores();
  const leave = leaveDates(month);
  const extras = bonusTasks(month);
  const [year, monthNumber] = month.split('-').map(Number);
  const lastDay = new Date(year, monthNumber, 0).getDate();
  const currentDate = today();

  let points = 0;
  let possible = 0;
  let pending = 0;

  for (const chore of activeChores) {
    let applicable = 0;
    let approved = 0;

    for (let day = 1; day <= lastDay; day++) {
      const date = month + '-' + String(day).padStart(2, '0');
      // Future dates still count towards the full month's possible chores.
      const entry = monthlyEntries.find(item =>
        item.date === date && item.choreId === chore.id
      );

      if (leave.includes(date) || (entry && entry.status === 'Not Needed')) continue;

      applicable++;
      if (entry && entry.status === 'Approved') approved++;
      if (entry && entry.status === 'Pending') pending++;
    }

    if (applicable) {
      points += chore.weight * approved / applicable;
      possible += chore.weight;
    }
  }

  const percentage = possible
    ? Math.round(10000 * points / possible) / 100
    : 0;
  // Only award the monthly bonus once the entire month has finished.
  const monthFinished = currentDate > month + '-' + String(lastDay).padStart(2, '0');
  const bonus = monthFinished && percentage >= 85 && pending === 0 ? 20 : 0;

  return {
    month,
    child: 'James',
    chores: activeChores,
    leaveDays: leave,
    bonusTasks: extras,
    entries: monthlyEntries.map(entry =>
      privileged
        ? entry
        : {
            date: entry.date,
            choreId: entry.choreId,
            status: entry.status
          }
    ),
    totals: {
      completionPercent: percentage,
      pocketMoney: Math.round(percentage * 20) / 100,
      bonusCash: extras.filter(t => t.status === 'Approved' && t.destination === 'Cash').reduce((n,t) => n + t.amount, 0),
      bonusSavings: extras.filter(t => t.status === 'Approved' && t.destination === 'Savings').reduce((n,t) => n + t.amount, 0),
      savingsBonus: bonus,
      guaranteedSavings: 50,
      totalSavings: 50 + bonus + extras.filter(t => t.status === 'Approved' && t.destination === 'Savings').reduce((n,t) => n + t.amount, 0),
      pendingApprovals: pending
    }
  };
}

function submit(date, choreId) {
  dateOk(date);
  if (date > today()) throw Error('Future date');

  if (leaveDates(date.slice(0, 7)).includes(date)) throw Error('This is an approved leave day');
  const chore = chores().find(item => item.id === choreId);
  if (!chore) throw Error('Unknown chore');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const alreadySubmitted = entries().some(entry =>
      entry.date === date &&
      entry.choreId === choreId &&
      entry.child === 'James'
    );
    if (alreadySubmitted) throw Error('Already submitted');

    tab(SHEETS.log).appendRow([
      Utilities.getUuid(),
      date,
      'James',
      chore.id,
      chore.name,
      'Pending',
      new Date().toISOString(),
      '',
      '',
      ''
    ]);

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function review(pin, decisions) {
  auth(pin);

  if (
    !Array.isArray(decisions) ||
    !decisions.length ||
    decisions.length > 200
  ) {
    throw Error('Invalid decisions');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const allEntries = entries();
    const seenIds = new Set();

    for (const decision of decisions) {
      if (
        !['Approved', 'Rejected', 'Not Needed'].includes(decision.status) ||
        seenIds.has(decision.id)
      ) {
        throw Error('Invalid decision');
      }

      seenIds.add(decision.id);
      const entry = allEntries.find(item => item.id === decision.id);
      if (!entry || entry.status !== 'Pending') {
        throw Error('Already reviewed or missing');
      }
    }

    const logSheet = tab(SHEETS.log);

    for (const decision of decisions) {
      const entry = allEntries.find(item => item.id === decision.id);

      // Columns F-J: status, submitted, reviewed, reviewer, notes.
      logSheet.getRange(entry.row, 6, 1, 5).setValues([[
        decision.status,
        entry.submittedAt,
        new Date().toISOString(),
        'Parent',
        String(decision.notes || '').slice(0, 250)
      ]]);
    }

    return { reviewed: decisions.length };
  } finally {
    lock.releaseLock();
  }
}

function doGet(event) {
  const params = event.parameter || {};

  if (params.page === 'summary') {
    return publicDashboardResponse(params.callback);
  }

  if (params.page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Family Chores · Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (params.page === 'parent') {
    return HtmlService.createHtmlOutputFromFile('Parent')
      .setTitle('Family Chores · Parent');
  }

  if (params.page === 'child') {
    return HtmlService.createHtmlOutputFromFile('Child')
      .setTitle('Family Chores · James');
  }

  return HtmlService.createHtmlOutput(
    '<p>Open <a href="?page=child">James’s chores</a> ' +
    'or <a href="?page=parent">Parent sign-off</a>.</p>'
  );
}

// Functions called by google.script.run in the hosted HTML pages.
function apiMonth(month) {
  return monthData(month, false);
}

function apiSubmit(date, choreId) {
  return submit(date, choreId);
}

function apiParentMonth(pin, month) {
  auth(pin);
  return monthData(month, true);
}

function apiReview(pin, decisions) {
  return review(pin, decisions);
}

/**
 * Read-only overview for the public family dashboard.
 * Only returns counts, chore names and totals, never PINs or log IDs.
 */
function apiDashboard() {
  const date = today();
  const month = date.slice(0, 7);
  const data = monthData(month, false);
  const outstanding = data.leaveDays.includes(date) ? [] : data.chores.filter(chore =>
    !data.entries.some(entry =>
      entry.date === date && entry.choreId === chore.id
    )
  );

  return {
    completionPercent: data.totals.completionPercent,
    pocketMoney: data.totals.pocketMoney,
    bonusCash: data.totals.bonusCash,
    bonusSavings: data.totals.bonusSavings,
    leaveToday: data.leaveDays.includes(date),
    pendingApprovals: data.totals.pendingApprovals,
    totalSavings: data.totals.totalSavings,
    savingsBonus: data.totals.savingsBonus,
    outstanding: outstanding.map(chore => ({ name: chore.name })),
    bonusTasks: data.bonusTasks.filter(task => task.status === 'Open').map(task => ({name:task.name, amount:task.amount}))
  };
}

/**
 * Public, read-only JSONP endpoint for the GitHub Pages dashboard.
 *
 * Apps Script web apps redirect ContentService output to googleusercontent.com,
 * so ordinary cross-origin fetch() is unreliable without CORS headers.
 * A script tag can load JSONP without needing CORS or third-party cookies.
 *
 * Only the limited apiDashboard() summary is published. Never add PINs,
 * log identifiers, review notes or other private records to this response.
 */
function publicDashboardResponse(callback) {
  if (!/^[A-Za-z_$][\w$]{0,63}$/.test(String(callback || ''))) {
    return ContentService.createTextOutput('Invalid callback')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const summary = apiDashboard();
  const payload = JSON.stringify(summary)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  return ContentService.createTextOutput(callback + '(' + payload + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
