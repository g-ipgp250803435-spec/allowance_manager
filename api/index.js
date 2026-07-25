// Cached connection (add this near the top of api/index.js)
let _docCache = null;
let _docCacheTime = 0;

async function getDoc() {
  const now = Date.now();
  if (_docCache && (now - _docCacheTime) < 5000) return _docCache;
  const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID);
  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
  });
  await doc.loadInfo();
  _docCache = doc;
  _docCacheTime = now;
  return doc;
}

const { GoogleSpreadsheet } = require('google-spreadsheet');
const cors = require('cors');

function uuid() {
  return 'xxxx-xxxx-xxxx-xxxx-xxxx'.replace(/x/g, () => Math.floor(Math.random() * 16).toString(16));
}

const MONTH_MAP = {
  // English
  january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
  july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
  jan: 1, feb: 2, mar: 3, apr: 4, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
  // Malay
  januari: 1, februari: 2, mac: 3, mei: 5, jun: 6, julai: 7, ogos: 8, oktober: 10, disember: 12
};

function parseMonthYear(monthYearStr) {
  if (!monthYearStr) return { monthNum: 0, year: 0 };
  const parts = monthYearStr.trim().split(/\s+/);
  if (parts.length < 2) return { monthNum: 0, year: 0 };
  const monthName = parts[0].toLowerCase();
  const year = parseInt(parts[1]) || 0;
  const monthNum = MONTH_MAP[monthName] || 0;
  return { monthNum, year };
}

function compareMonthYear(aStr, bStr) {
  const a = parseMonthYear(aStr);
  const b = parseMonthYear(bStr);
  if (a.year !== b.year) return a.year - b.year;
  return a.monthNum - b.monthNum;
}

async function parseBody(req) {
  return new Promise(resolve => {
    let body = '';
    req.on('data', chunk => (body += chunk.toString()));
    req.on('end', () => {
      try { resolve(JSON.parse(body)); } catch { resolve({}); }
    });
  });
}

async function getDoc() {
  const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID);
  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
  });
  await doc.loadInfo();
  return doc;
}

async function getOrCreateSheet(doc, title, headers) {
  let sheet = doc.sheetsByTitle[title];
  if (!sheet) sheet = await doc.addSheet({ title, headerValues: headers });
  return sheet;
}

// In-memory Mock Storage for Local Development fallback
let mockBudgetSettings = { total: 200, refills: 100, savings_usage: 50 };
let mockSecureNotes = [];
let mockSharedLinks = [];

module.exports = async (req, res) => {
  await cors()(req, res, async () => {
    const { action } = req.query;
    let body = {};
    if (req.method === 'POST') {
      if (req.body && typeof req.body === 'object') {
        body = req.body;
      } else if (req.body && typeof req.body === 'string') {
        try { body = JSON.parse(req.body); } catch { body = {}; }
      } else {
        body = await parseBody(req);
      }
    }

    if (action === 'login') {
      if (body.username === process.env.ADMIN_USERNAME && body.password === process.env.ADMIN_PASSWORD) {
        return res.json({ success: true, token: 'admin-session' });
      }
      return res.status(401).json({ error: 'Invalid credentials' });
    }

    if (action !== 'getSharedDashboard') {
      if (body.token !== 'admin-session') return res.status(401).json({ error: 'Unauthorized' });
    }

    const useMock = !process.env.GOOGLE_SHEET_ID || !process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY;
    if (useMock) {
      let result;
      switch (action) {
        case 'getDashboard':
          result = {
            mamaAccount: 240,
            savingsBalance: 120,
            totalRemaining: 360,
            offsets: { wallet: 0, savings: 0 },
            moneyFlowSettings: { allowance: 430, mama: 300, wallet: 100, savings: 30 },
            alertSettings: { mama_threshold: 50, savings_threshold: 50, savings_goal: 500 },
            currencySettings: { primary_currency: 'MYR', selected_currency: 'MYR', rates: { MYR: 1, USD: 0.22, SGD: 0.3 } },
            allowanceStats: [
              { month: 'January 2026', dateReceived: '01/01/2026', allowance: 430, usage: 100, savings: 30, balance: 300 },
              { month: 'December 2025', dateReceived: '01/12/2025', allowance: 430, usage: 100, savings: 30, balance: 300 }
            ],
            savingsStats: [
              { month: 'January 2026', savings: 30, usage: 0, balance: 30 },
              { month: 'December 2025', savings: 30, usage: 10, balance: 20 }
            ],
            monthlyTransfers: [
              { month: 'Jan 2026', amount: 80 }
            ],
            bankAccounts: [
              { name: "Aieryl's Maybank", number: "153056659975" },
              { name: "Aieryl's Bank Rakyat", number: "2252698058" },
              { name: "Mama's Bank Rakyat", number: "2212319157" }
            ],
            lastUpdated: new Date().toISOString()
          };
          break;
        case 'getTransactions':
          result = [
            { date: new Date().toISOString(), type: 'refill', amount: 80, note: 'Refill Maybank', id: 'tx-1', attachment: '' },
            { date: new Date(Date.now() - 5*24*60*60*1000).toISOString(), type: 'income', amount: 300, note: 'Allowance Received', id: 'tx-2', attachment: '' },
            { date: new Date(Date.now() - 10*24*60*60*1000).toISOString(), type: 'savings_usage', amount: 10, note: 'Used savings Dec', id: 'tx-3', attachment: '' }
          ];
          break;
        case 'getHistory':
          result = [
            { id: 'hist-1', type: 'addTransaction', details: { type: 'refill', amount: 80, note: 'Refill Maybank' }, timestamp: new Date().toISOString(), undone: false }
          ];
          break;
        case 'getBudgetSettings':
          result = mockBudgetSettings;
          break;
        case 'setBudgetSettings':
          mockBudgetSettings = { total: body.total, refills: body.refills, savings_usage: body.savings_usage };
          result = { success: true };
          break;
        case 'getSecureNotes':
          result = mockSecureNotes;
          break;
        case 'saveSecureNote':
          const existingNoteIdx = mockSecureNotes.findIndex(n => n.id === body.id);
          const newNote = { id: body.id || uuid(), encrypted_data: body.encrypted_data, iv: body.iv, salt: body.salt };
          if (existingNoteIdx >= 0) {
            mockSecureNotes[existingNoteIdx] = newNote;
          } else {
            mockSecureNotes.push(newNote);
          }
          result = { success: true };
          break;
        case 'deleteSecureNote':
          mockSecureNotes = mockSecureNotes.filter(n => n.id !== body.id);
          result = { success: true };
          break;
        case 'getSharedLinks':
          result = mockSharedLinks;
          break;
        case 'createSharedLink':
          const newLink = {
            id: uuid(),
            token: uuid().replace(/-/g, ''),
            expiry: new Date(Date.now() + parseFloat(body.expiryHours) * 60 * 60 * 1000).toISOString(),
            password_hash: body.passwordHash || '',
            created_at: new Date().toISOString()
          };
          mockSharedLinks.push(newLink);
          result = { success: true, id: newLink.id, token: newLink.token, expiry: newLink.expiry };
          break;
        case 'revokeSharedLink':
          mockSharedLinks = mockSharedLinks.filter(l => l.id !== body.id && l.token !== body.id);
          result = { success: true };
          break;
        case 'getSharedDashboard':
          const link = mockSharedLinks.find(l => l.token === body.token);
          if (!link) {
            result = { error: 'Link invalid' };
          } else if (new Date() > new Date(link.expiry)) {
            result = { error: 'Link expired' };
          } else if (link.password_hash && link.password_hash !== body.passwordHash) {
            result = { error: 'Password incorrect', password_required: true };
          } else {
            result = {
              success: true,
              mamaAccount: 240,
              savingsBalance: 120,
              totalRemaining: 360,
              moneyFlowSettings: { allowance: 430, mama: 300, wallet: 100, savings: 30 }
            };
          }
          break;
        case 'addTransaction':
        case 'deleteTransaction':
        case 'requestMoney':
        case 'useSavings':
        case 'topUpSavings':
        case 'setOffsets':
        case 'setMoneyFlowSettings':
        case 'setAlertSettings':
        case 'getCurrencySettings':
        case 'setCurrencySettings':
          result = { success: true };
          break;
        default:
          result = { error: 'Unknown action' };
      }
      return res.json(result);
    }

    try {
      const doc = await getDoc();
      let result;
      switch (action) {
        case 'getDashboard': result = await getDashboardData(doc); break;
        case 'requestMoney': result = await requestMoney(doc, body.amount); break;
        case 'processNewMonth': result = await processNewMonth(doc, body.month, body.year, body.dateReceived); break;
        case 'updateDateReceived': result = await updateDateReceived(doc, body.month, body.newDate); break;
        case 'useSavings': result = await useSavings(doc, body.amount, body.month); break;
        case 'topUpSavings': result = await topUpSavings(doc, body.months, body.totalAmount); break;
        case 'setOffsets': result = await setOffsets(doc, body.walletOffset, body.savingsOffset); break;
        case 'setMoneyFlowSettings': result = await setMoneyFlowSettings(doc, body.allowance, body.mama, body.wallet, body.savings); break;
        case 'setAlertSettings': result = await setAlertSettings(doc, body.mamaThreshold, body.savingsThreshold, body.savingsGoal); break;
        case 'getCurrencySettings': result = await getCurrencySettings(doc); break;
        case 'setCurrencySettings': result = await setCurrencySettings(doc, body.primary, body.selected, body.customRates); break;
        case 'getHistory': result = await getHistory(doc); break;
        case 'getTransactions': result = await getTransactions(doc); break;
        case 'addTransaction': result = await addTransaction(doc, body.type, body.amount, body.note, body.attachment); break;
        case 'deleteTransaction': result = await deleteTransaction(doc, body.transactionId); break;
        case 'redoAction': result = await redoAction(doc, body.actionId); break;
        case 'undoAction': result = await undoAction(doc, body.actionId); break;
        case 'backup': result = await backupData(doc); break;
        case 'restore': result = await restoreData(doc, body.backup); break;
        case 'getBudgetSettings': result = await getBudgetSettings(doc); break;
        case 'setBudgetSettings': result = await setBudgetSettings(doc, body.total, body.refills, body.savings_usage); break;
        case 'getSecureNotes': result = await getSecureNotes(doc); break;
        case 'saveSecureNote': result = await saveSecureNote(doc, body.id, body.encrypted_data, body.iv, body.salt); break;
        case 'deleteSecureNote': result = await deleteSecureNote(doc, body.id); break;
        case 'getSharedLinks': result = await getSharedLinks(doc); break;
        case 'createSharedLink': result = await createSharedLink(doc, body.expiryHours, body.passwordHash); break;
        case 'revokeSharedLink': result = await revokeSharedLink(doc, body.id); break;
        case 'getSharedDashboard': result = await getSharedDashboard(doc, body.token, body.passwordHash); break;
        default: result = { error: 'Unknown action' };
      }
      res.json(result);
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message });
    }
  });
};

// ---------- Alert Settings ----------
async function getAlertSettings(doc) {
  const sheet = await getOrCreateSheet(doc, 'alert_settings', ['mama_threshold', 'savings_threshold', 'savings_goal']);
  const rows = await sheet.getRows();
  if (rows.length === 0) {
    const defaults = { mama_threshold: '50', savings_threshold: '50', savings_goal: '500' };
    await sheet.addRow(defaults);
    return { mama_threshold: 50, savings_threshold: 50, savings_goal: 500 };
  }
  return {
    mama_threshold: isNaN(parseFloat(rows[0].mama_threshold)) ? 50 : parseFloat(rows[0].mama_threshold),
    savings_threshold: isNaN(parseFloat(rows[0].savings_threshold)) ? 50 : parseFloat(rows[0].savings_threshold),
    savings_goal: isNaN(parseFloat(rows[0].savings_goal)) ? 500 : parseFloat(rows[0].savings_goal)
  };
}

async function setAlertSettings(doc, mama, savings, goal) {
  const sheet = await getOrCreateSheet(doc, 'alert_settings', ['mama_threshold', 'savings_threshold', 'savings_goal']);
  let rows = await sheet.getRows();
  if (rows.length === 0) {
    await sheet.addRow({ mama_threshold: mama, savings_threshold: savings, savings_goal: goal });
  } else {
    rows[0].mama_threshold = mama;
    rows[0].savings_threshold = savings;
    rows[0].savings_goal = goal;
    await rows[0].save();
  }
  await addHistory(doc, 'setAlertSettings', { mama, savings, goal });
  return { success: true };
}

async function getCurrencySettings(doc) {
  const sheet = await getOrCreateSheet(doc, 'currency_settings', ['primary_currency', 'selected_currency', 'rates_json', 'last_updated']);
  const rows = await sheet.getRows();
  let data = {
    primary_currency: 'MYR',
    selected_currency: 'MYR',
    rates_json: '{}',
    last_updated: ''
  };

  if (rows.length > 0) {
    data = {
      primary_currency: rows[0].primary_currency || 'MYR',
      selected_currency: rows[0].selected_currency || 'MYR',
      rates_json: rows[0].rates_json || '{}',
      last_updated: rows[0].last_updated || ''
    };
  } else {
    await sheet.addRow(data);
  }

  const now = new Date();
  let rates = {};
  try {
    rates = JSON.parse(data.rates_json);
  } catch (e) {
    rates = {};
  }

  const lastUpdatedDate = data.last_updated ? new Date(data.last_updated) : null;
  const hoursSinceUpdate = lastUpdatedDate ? (now - lastUpdatedDate) / (1000 * 60 * 60) : 24;

  if (hoursSinceUpdate >= 24 || !rates || Object.keys(rates).length === 0) {
    try {
      // Use Node 18+ global fetch
      const res = await fetch('https://open.er-api.com/v6/latest/MYR');
      if (res.ok) {
        const json = await res.json();
        if (json && json.rates) {
          rates = json.rates;
          data.rates_json = JSON.stringify(rates);
          data.last_updated = now.toISOString();

          const updatedRows = await sheet.getRows();
          if (updatedRows.length > 0) {
            updatedRows[0].rates_json = data.rates_json;
            updatedRows[0].last_updated = data.last_updated;
            await updatedRows[0].save();
          }
        }
      }
    } catch (err) {
      console.error('Failed to fetch exchange rates:', err);
    }
  }

  if (!rates || Object.keys(rates).length === 0) {
    rates = { MYR: 1, USD: 0.21, SGD: 0.28, EUR: 0.19, GBP: 0.17 };
  }

  return {
    primary_currency: data.primary_currency,
    selected_currency: data.selected_currency,
    rates,
    last_updated: data.last_updated
  };
}

async function setCurrencySettings(doc, primary, selected, customRates = null) {
  const sheet = await getOrCreateSheet(doc, 'currency_settings', ['primary_currency', 'selected_currency', 'rates_json', 'last_updated']);
  let rows = await sheet.getRows();
  const rates_json = customRates ? JSON.stringify(customRates) : '{}';
  const now = new Date().toISOString();

  if (rows.length === 0) {
    await sheet.addRow({
      primary_currency: primary,
      selected_currency: selected,
      rates_json,
      last_updated: now
    });
  } else {
    rows[0].primary_currency = primary;
    rows[0].selected_currency = selected;
    if (customRates) {
      rows[0].rates_json = rates_json;
      rows[0].last_updated = now;
    }
    await rows[0].save();
  }
  await addHistory(doc, 'setCurrencySettings', { primary, selected });
  return { success: true };
}

async function backupData(doc) {
  const sheetsToBackup = [
    { title: 'transactions', headers: ['Date', 'Type', 'Amount', 'Note', 'ID', 'Attachment'] },
    { title: 'allowance_stats', headers: ['Month', 'Date Received', 'Allowance Amount', 'Usage', 'Savings', 'Balance'] },
    { title: 'savings_stats', headers: ['Month', 'Savings', 'Usage', 'Balance'] },
    { title: 'manual_offsets', headers: ['wallet_offset', 'savings_offset'] },
    { title: 'money_flow_settings', headers: ['allowance', 'mama', 'wallet', 'savings'] },
    { title: 'alert_settings', headers: ['mama_threshold', 'savings_threshold', 'savings_goal'] },
    { title: 'currency_settings', headers: ['primary_currency', 'selected_currency', 'rates_json', 'last_updated'] },
    { title: 'history', headers: ['ID', 'Type', 'Details', 'Timestamp', 'Undone'] },
    { title: 'budget_settings', headers: ['category', 'budget_amount'] },
    { title: 'secure_notes', headers: ['id', 'encrypted_data', 'iv', 'salt'] },
    { title: 'shared_links', headers: ['id', 'token', 'expiry', 'password_hash', 'created_at'] }
  ];

  const backup = {};
  for (const s of sheetsToBackup) {
    const sheet = await getOrCreateSheet(doc, s.title, s.headers);
    const rows = await sheet.getRows();
    backup[s.title] = rows.map(r => {
      const rowData = {};
      s.headers.forEach(h => {
        rowData[h] = r[h] !== undefined ? r[h] : '';
      });
      return rowData;
    });
  }
  return { backup };
}

async function restoreData(doc, backup) {
  if (!backup || typeof backup !== 'object') {
    throw new Error('Invalid backup data format');
  }

  const sheetsToRestore = [
    { title: 'transactions', headers: ['Date', 'Type', 'Amount', 'Note', 'ID', 'Attachment'] },
    { title: 'allowance_stats', headers: ['Month', 'Date Received', 'Allowance Amount', 'Usage', 'Savings', 'Balance'] },
    { title: 'savings_stats', headers: ['Month', 'Savings', 'Usage', 'Balance'] },
    { title: 'manual_offsets', headers: ['wallet_offset', 'savings_offset'] },
    { title: 'money_flow_settings', headers: ['allowance', 'mama', 'wallet', 'savings'] },
    { title: 'alert_settings', headers: ['mama_threshold', 'savings_threshold', 'savings_goal'] },
    { title: 'currency_settings', headers: ['primary_currency', 'selected_currency', 'rates_json', 'last_updated'] },
    { title: 'history', headers: ['ID', 'Type', 'Details', 'Timestamp', 'Undone'] },
    { title: 'budget_settings', headers: ['category', 'budget_amount'] },
    { title: 'secure_notes', headers: ['id', 'encrypted_data', 'iv', 'salt'] },
    { title: 'shared_links', headers: ['id', 'token', 'expiry', 'password_hash', 'created_at'] }
  ];

  for (const s of sheetsToRestore) {
    if (!backup[s.title]) continue;
    const sheet = await getOrCreateSheet(doc, s.title, s.headers);

    // Clear existing data (fast, 1 call)
    await sheet.clear();
    // Restore headers (fast, 1 call)
    await sheet.setHeaderRow(s.headers);

    // Add new rows in a batch (fast, 1 call)
    const backupRows = backup[s.title];
    const rowsToRestore = backupRows.map(r => {
      const rowObj = {};
      s.headers.forEach(h => {
        rowObj[h] = r[h] !== undefined ? r[h] : '';
      });
      return rowObj;
    });

    if (rowsToRestore.length > 0) {
      await sheet.addRows(rowsToRestore);
    }
  }

  await addHistory(doc, 'restoreData', { timestamp: new Date().toISOString() });
  return { success: true };
}

// ---------- Dashboard Data ----------
async function getDashboardData(doc) {
  const offsets = await getOffsets(doc);
  const balance = await calculateBalance(doc);
  const savingsBalance = await calculateSavingsBalance(doc);
  const stats = await getAllowanceStats(doc);
  const savingsStats = await getSavingsStats(doc);
  const transfers = await getMonthlyMomTransfers(doc);
  const moneyFlowSettings = await getMoneyFlowSettings(doc);
  const alertSettings = await getAlertSettings(doc);
  const currencySettings = await getCurrencySettings(doc);
  return {
    mamaAccount: balance + offsets.wallet,
    savingsBalance: savingsBalance + offsets.savings,
    totalRemaining: balance + offsets.wallet + savingsBalance + offsets.savings,
    offsets,
    moneyFlowSettings,
    alertSettings,
    currencySettings,
    allowanceStats: stats,
    savingsStats,
    monthlyTransfers: transfers,
    bankAccounts: [
      { name: "Aieryl's Maybank", number: "153056659975" },
      { name: "Aieryl's Bank Rakyat", number: "2252698058" },
      { name: "Mama's Bank Rakyat", number: "2212319157" }
    ],
    lastUpdated: new Date().toISOString()
  };
}

// ---------- Transactions Extra ----------
async function getTransactions(doc) {
  const sheet = await getOrCreateSheet(doc, 'transactions', ['Date', 'Type', 'Amount', 'Note', 'ID', 'Attachment']);
  const rows = await sheet.getRows();
  return rows.map(r => ({
    date: r.Date,
    type: r.Type,
    amount: Number(r.Amount) || 0,
    note: r.Note,
    id: r.ID,
    attachment: r.Attachment || ''
  }));
}

async function addTransaction(doc, type, amount, note, attachment = '') {
  const sheet = await getOrCreateSheet(doc, 'transactions', ['Date', 'Type', 'Amount', 'Note', 'ID', 'Attachment']);
  const txId = uuid();
  const dateStr = new Date().toISOString();
  await sheet.addRow({ Date: dateStr, Type: type, Amount: amount, Note: note, ID: txId, Attachment: attachment });
  await addHistory(doc, 'addTransaction', { type, amount, note, txId, attachment });
  return { success: true, transactionId: txId };
}

async function deleteTransaction(doc, transactionId) {
  const txSheet = await getOrCreateSheet(doc, 'transactions', ['Date', 'Type', 'Amount', 'Note', 'ID', 'Attachment']);
  const rows = await txSheet.getRows();
  const row = rows.find(r => r.ID === transactionId);
  if (!row) {
    throw new Error(`Transaction with ID ${transactionId} not found`);
  }

  // Backup data
  const txBackup = {
    id: row.ID,
    date: row.Date,
    type: row.Type,
    amount: Number(row.Amount) || 0,
    note: row.Note,
    attachment: row.Attachment || ''
  };

  // Check the history to see if there's any detailed adjustment we can reverse directly.
  // Otherwise, we do dynamic adjustments.
  const historySheet = await getOrCreateHistorySheet(doc);
  const histRows = await historySheet.getRows();

  // Find matching history action where details contain this txId
  let adjustments = [];
  let matchingHistAction = null;

  for (const hRow of histRows) {
    try {
      const details = JSON.parse(hRow.Details);
      if (details.txId === transactionId && hRow.Undone !== 'true' && hRow.Undone !== true) {
        matchingHistAction = hRow;
        adjustments = details.adjustments || [];
        break;
      }
    } catch (e) {}
  }

  const resultAdjustments = [];

  // Reversing adjustments depending on type
  if (txBackup.type === 'refill' || txBackup.type === 'requestMoney') {
    // Reversing a transfer request means we need to deduct row.Usage and increase row.Balance
    const statsSheet = doc.sheetsByTitle['allowance_stats'];
    if (statsSheet) {
      const statsRows = await statsSheet.getRows();
      if (adjustments && adjustments.length > 0) {
        // Use matching history row's adjustments
        for (const adj of adjustments) {
          const statsRow = statsRows.find(r => r.rowNumber === adj.row);
          if (statsRow) {
            statsRow.Usage = (Number(statsRow.Usage) || 0) - adj.deducted;
            statsRow.Balance = (Number(statsRow.Balance) || 0) + adj.deducted;
            await statsRow.save();
            resultAdjustments.push({ sheet: 'allowance_stats', row: adj.row, month: statsRow.Month, deducted: -adj.deducted });
          }
        }
      } else {
        // Fallback: reverse-chronological FIFO deduction from allowance_stats Usage
        statsRows.sort((a, b) => compareMonthYear(b.Month, a.Month)); // Reverse order
        let remaining = txBackup.amount;
        for (const statsRow of statsRows) {
          if (remaining <= 0) break;
          const usage = Number(statsRow.Usage) || 0;
          const restore = Math.min(remaining, usage);
          if (restore > 0) {
            statsRow.Usage = usage - restore;
            statsRow.Balance = (Number(statsRow.Balance) || 0) + restore;
            await statsRow.save();
            resultAdjustments.push({ sheet: 'allowance_stats', row: statsRow.rowNumber, month: statsRow.Month, deducted: -restore });
            remaining -= restore;
          }
        }
      }
    }
  } else if (txBackup.type === 'savings_usage' || txBackup.type === 'useSavings') {
    // Reversing savings usage means we deduct row.Usage and increase row.Balance in savings_stats
    const savSheet = doc.sheetsByTitle['savings_stats'];
    if (savSheet) {
      const savRows = await savSheet.getRows();
      if (adjustments && adjustments.length > 0) {
        for (const adj of adjustments) {
          const savRow = savRows.find(r => r.rowNumber === adj.row);
          if (savRow) {
            savRow.Usage = (Number(savRow.Usage) || 0) - adj.deducted;
            savRow.Balance = (Number(savRow.Balance) || 0) + adj.deducted;
            await savRow.save();
            resultAdjustments.push({ sheet: 'savings_stats', row: adj.row, month: savRow.Month, deducted: -adj.deducted });
          }
        }
      } else {
        // Fallback
        savRows.sort((a, b) => compareMonthYear(b.Month, a.Month));
        let remaining = txBackup.amount;
        for (const savRow of savRows) {
          if (remaining <= 0) break;
          const usage = Number(savRow.Usage) || 0;
          const restore = Math.min(remaining, usage);
          if (restore > 0) {
            savRow.Usage = usage - restore;
            savRow.Balance = (Number(savRow.Balance) || 0) + restore;
            await savRow.save();
            resultAdjustments.push({ sheet: 'savings_stats', row: savRow.rowNumber, month: savRow.Month, deducted: -restore });
            remaining -= restore;
          }
        }
      }
    }
  } else if (txBackup.type === 'savings_topup' || txBackup.type === 'topUpSavings') {
    // Reversing savings topup means we add back row.Usage and deduct row.Balance in savings_stats
    const savSheet = doc.sheetsByTitle['savings_stats'];
    if (savSheet) {
      const savRows = await savSheet.getRows();
      if (adjustments && adjustments.length > 0) {
        for (const adj of adjustments) {
          const savRow = savRows.find(r => r.rowNumber === adj.row);
          if (savRow) {
            savRow.Usage = (Number(savRow.Usage) || 0) + adj.restored;
            savRow.Balance = (Number(savRow.Balance) || 0) - adj.restored;
            await savRow.save();
            resultAdjustments.push({ sheet: 'savings_stats', row: adj.row, month: savRow.Month, restored: -adj.restored });
          }
        }
      } else {
        // Fallback
        savRows.sort((a, b) => compareMonthYear(a.Month, b.Month));
        let remaining = txBackup.amount;
        for (const savRow of savRows) {
          if (remaining <= 0) break;
          const currentBalance = Number(savRow.Balance) || 0;
          const deduct = Math.min(remaining, currentBalance);
          if (deduct > 0) {
            savRow.Usage = (Number(savRow.Usage) || 0) + deduct;
            savRow.Balance = currentBalance - deduct;
            await savRow.save();
            resultAdjustments.push({ sheet: 'savings_stats', row: savRow.rowNumber, month: savRow.Month, restored: -deduct });
            remaining -= deduct;
          }
        }
      }
    }
  }

  // Delete actual row
  await row.delete();

  // Mark matching history row as undone (to prevent further confusion)
  if (matchingHistAction) {
    matchingHistAction.Undone = true;
    await matchingHistAction.save();
  }

  // Add deleteTransaction action to history
  await addHistory(doc, 'deleteTransaction', { txBackup, resultAdjustments, originalHistoryActionId: matchingHistAction ? matchingHistAction.ID : null });
  return { success: true };
}

// ---------- Balances ----------
async function calculateBalance(doc) {
  const sheet = doc.sheetsByTitle['transactions'];
  if (!sheet) return 0;
  const rows = await sheet.getRows();
  let balance = 0;
  for (const r of rows) {
    if (r.Type === 'income') balance += Number(r.Amount);
    else if (r.Type === 'refill') balance -= Number(r.Amount);
  }
  return balance;
}

async function calculateSavingsBalance(doc) {
  const sheet = doc.sheetsByTitle['savings_stats'];
  if (!sheet) return 0;
  const rows = await sheet.getRows();
  let total = 0;
  for (const r of rows) total += Number(r.Balance) || 0;
  return total;
}

// ---------- Offsets ----------
async function getOffsets(doc) {
  const sheet = doc.sheetsByTitle['manual_offsets'];
  if (!sheet) return { wallet: 0, savings: 0 };
  const rows = await sheet.getRows();
  return { wallet: Number(rows[0]?.wallet_offset) || 0, savings: Number(rows[1]?.savings_offset) || 0 };
}

async function setOffsets(doc, wallet, savings) {
  let sheet = doc.sheetsByTitle['manual_offsets'];
  if (!sheet) {
    sheet = await doc.addSheet({ title: 'manual_offsets', headerValues: ['wallet_offset', 'savings_offset'] });
  }
  let rows = await sheet.getRows();
  while (rows.length < 2) {
    await sheet.addRow({ wallet_offset: 0, savings_offset: 0 });
    rows = await sheet.getRows();
  }
  const oldWallet = Number(rows[0].wallet_offset) || 0;
  const oldSavings = Number(rows[1].savings_offset) || 0;
  rows[0].wallet_offset = wallet;
  rows[1].savings_offset = savings;
  await rows[0].save();
  await rows[1].save();
  await addHistory(doc, 'setOffsets', { wallet, savings, oldWallet, oldSavings });
  return { success: true };
}

// ---------- Money Flow Settings ----------
async function getMoneyFlowSettings(doc) {
  const sheet = doc.sheetsByTitle['money_flow_settings'];
  if (!sheet) {
    return { allowance: 430, mama: 300, wallet: 100, savings: 30 };
  }
  const rows = await sheet.getRows();
  if (rows.length === 0) {
    return { allowance: 430, mama: 300, wallet: 100, savings: 30 };
  }
  return {
    allowance: Number(rows[0].allowance) || 430,
    mama: Number(rows[0].mama) || 300,
    wallet: Number(rows[0].wallet) || 100,
    savings: Number(rows[0].savings) || 30
  };
}

async function setMoneyFlowSettings(doc, allowance, mama, wallet, savings) {
  let sheet = doc.sheetsByTitle['money_flow_settings'];
  if (!sheet) {
    sheet = await doc.addSheet({ title: 'money_flow_settings', headerValues: ['allowance', 'mama', 'wallet', 'savings'] });
  }
  let rows = await sheet.getRows();
  if (rows.length === 0) {
    await sheet.addRow({ allowance, mama, wallet, savings });
  } else {
    rows[0].allowance = allowance;
    rows[0].mama = mama;
    rows[0].wallet = wallet;
    rows[0].savings = savings;
    await rows[0].save();
  }
  await addHistory(doc, 'setMoneyFlowSettings', { allowance, mama, wallet, savings });
  return { success: true };
}

// ---------- Monthly Allocation ----------
async function processNewMonth(doc, month, year, dateReceived) {
  const now = new Date();

  // Use passed parameters if provided
  let monthLabel;
  let dateStr = dateReceived || now.toLocaleDateString('en-GB');

  // Fetch configured bases first
  const flowSettings = await getMoneyFlowSettings(doc);
  const monthlyAllowance = flowSettings.allowance; // Default 430
  const savings = flowSettings.savings; // Default 30
  const walletAmount = flowSettings.wallet; // Default 100 (formerly maybank)
  const custodian = monthlyAllowance - savings - walletAmount;

  if (month && year) {
    // month is "January" - "December"
    monthLabel = `${month} ${year}`;
  } else {
    monthLabel = now.toLocaleString('default', { month: 'long', year: 'numeric' });
  }

  const txSheet = doc.sheetsByTitle['transactions'] || await doc.addSheet({ title: 'transactions', headerValues: ['Date', 'Type', 'Amount', 'Note', 'ID'] });
  const incomeId = uuid();
  const savingsId = uuid();
  await txSheet.addRow({ Date: now.toISOString(), Type: 'income', Amount: custodian, Note: `Monthly allocation to Custodian (${monthLabel})`, ID: incomeId });
  await txSheet.addRow({ Date: now.toISOString(), Type: 'savings', Amount: savings, Note: `Monthly savings (${monthLabel})`, ID: savingsId });

  const statsSheet = await getOrCreateSheet(doc, 'allowance_stats', ['Month', 'Date Received', 'Allowance Amount', 'Usage', 'Savings', 'Balance']);
  await statsSheet.addRow({ Month: monthLabel, 'Date Received': dateStr, 'Allowance Amount': monthlyAllowance, Usage: walletAmount, Savings: savings, Balance: custodian });

  const savSheet = await getOrCreateSheet(doc, 'savings_stats', ['Month', 'Savings', 'Usage', 'Balance']);
  await savSheet.addRow({ Month: monthLabel, Savings: savings, Usage: 0, Balance: savings });

  await addHistory(doc, 'newMonth', { month: monthLabel, incomeId, savingsId, dateReceived: dateStr });
  return { success: true, month: monthLabel };
}

// ---------- Update Date Received ----------
async function updateDateReceived(doc, month, newDate) {
  const statsSheet = await getOrCreateSheet(doc, 'allowance_stats', ['Month', 'Date Received', 'Allowance Amount', 'Usage', 'Savings', 'Balance']);
  const rows = await statsSheet.getRows();
  const row = rows.find(r => r.Month.trim().toLowerCase() === month.trim().toLowerCase());
  if (!row) {
    throw new Error(`Month row ${month} not found in allowance_stats`);
  }
  const oldDate = row['Date Received'];
  row['Date Received'] = newDate;
  await row.save();

  await addHistory(doc, 'updateDateReceived', { month, oldDate, newDate });
  return { success: true };
}

// ---------- Request Money ----------
async function requestMoney(doc, amount) {
  const txSheet = await getOrCreateSheet(doc, 'transactions', ['Date', 'Type', 'Amount', 'Note', 'ID', 'Attachment']);
  const txId = uuid();
  await txSheet.addRow({ Date: new Date().toISOString(), Type: 'refill', Amount: amount, Note: 'Transfer from Mama', ID: txId });

  const statsSheet = doc.sheetsByTitle['allowance_stats'];
  const rows = await statsSheet.getRows();
  const withBalance = rows.filter(r => Number(r.Balance) > 0);
  withBalance.sort((a, b) => compareMonthYear(a.Month, b.Month));

  let remaining = amount;
  const adjustments = [];
  for (const row of withBalance) {
    if (remaining <= 0) break;
    const bal = Number(row.Balance), deduct = Math.min(remaining, bal);
    if (deduct > 0) {
      row.Usage = Number(row.Usage)+deduct; row.Balance = bal-deduct;
      await row.save();
      adjustments.push({ row: row.rowNumber, deducted: deduct });
      remaining -= deduct;
    }
  }
  await addHistory(doc, 'requestMoney', { amount, txId, adjustments });
  return { transactionId: txId, adjustments };
}

// ---------- Use Savings ----------
async function useSavings(doc, amount, month) {
  const txSheet = doc.sheetsByTitle['transactions'];
  const txId = uuid();
  const noteSuffix = month ? ` (${month})` : '';
  await txSheet.addRow({ Date: new Date().toISOString(), Type: 'savings_usage', Amount: amount, Note: `Used savings${noteSuffix}`, ID: txId });

  const savSheet = doc.sheetsByTitle['savings_stats'];
  const rows = await savSheet.getRows();
  let withBalance = rows.filter(r => Number(r.Balance) > 0);
  if (month) {
    withBalance = withBalance.filter(r => r.Month.trim().toLowerCase() === month.trim().toLowerCase());
    if (withBalance.length === 0) {
      throw new Error(`No available savings balance found for ${month}`);
    }
    const available = Number(withBalance[0].Balance);
    if (amount > available) {
      throw new Error(`Withdrawal amount exceeding available balance for ${month} (Available: RM${available.toFixed(2)})`);
    }
  } else {
    withBalance.sort((a, b) => compareMonthYear(a.Month, b.Month));
  }

  let remaining = amount;
  const adjustments = [];
  for (const row of withBalance) {
    if (remaining <= 0) break;
    const bal = Number(row.Balance), deduct = Math.min(remaining, bal);
    if (deduct > 0) {
      row.Usage = Number(row.Usage)+deduct; row.Balance = bal-deduct;
      await row.save();
      adjustments.push({ row: row.rowNumber, deducted: deduct });
      remaining -= deduct;
    }
  }
  await addHistory(doc, 'useSavings', { amount, month, txId, adjustments });
  return { transactionId: txId, adjustments };
}

// ---------- Top Up Savings ----------
async function topUpSavings(doc, months, totalAmount) {
  const txSheet = doc.sheetsByTitle['transactions'];
  const txId = uuid();
  await txSheet.addRow({ Date: new Date().toISOString(), Type: 'savings_topup', Amount: totalAmount, Note: 'Top-up savings', ID: txId });

  const savSheet = doc.sheetsByTitle['savings_stats'];
  const rows = await savSheet.getRows();
  const adjustments = [];
  let remaining = totalAmount;

  months.sort((a, b) => compareMonthYear(a.month, b.month));

  for (const sm of months) {
    if (remaining <= 0) break;
    const row = rows.find(r => r.Month.trim() === sm.month);
    if (!row) continue;
    const currentUsage = Number(row.Usage)||0, currentBalance = Number(row.Balance)||0;
    const restore = Math.min(remaining, currentUsage);
    if (restore > 0) {
      row.Usage = currentUsage - restore; row.Balance = currentBalance + restore;
      await row.save();
      adjustments.push({ row: row.rowNumber, month: sm.month, restored: restore, previousUsage: currentUsage, previousBalance: currentBalance });
      remaining -= restore;
    }
  }
  await addHistory(doc, 'topUpSavings', { totalAmount, months, txId, adjustments, unusedAmount: remaining });
  return { transactionId: txId, adjustments, unusedAmount: remaining };
}

// ---------- Stats ----------
async function getAllowanceStats(doc) {
  const sheet = doc.sheetsByTitle['allowance_stats'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows.map(r => ({
    month: r.Month, dateReceived: r['Date Received'],
    allowance: Number(r['Allowance Amount']), usage: Number(r.Usage),
    savings: Number(r.Savings), balance: Number(r.Balance)
  }));
}

async function getSavingsStats(doc) {
  const sheet = doc.sheetsByTitle['savings_stats'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows.map(r => ({
    month: r.Month, savings: Number(r.Savings), usage: Number(r.Usage), balance: Number(r.Balance)
  }));
}

async function getMonthlyMomTransfers(doc) {
  const sheet = doc.sheetsByTitle['transactions'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  const monthly = {};
  for (const r of rows) {
    if (r.Type !== 'refill') continue;
    const d = new Date(r.Date);
    const key = d.toLocaleString('default', { month: 'short', year: 'numeric' });
    monthly[key] = (monthly[key] || 0) + Number(r.Amount);
  }
  return Object.entries(monthly).map(([month, amount]) => ({ month, amount }));
}

// ---------- History ----------
async function getOrCreateHistorySheet(doc) {
  return await getOrCreateSheet(doc, 'history', ['ID', 'Type', 'Details', 'Timestamp', 'Undone']);
}

async function addHistory(doc, type, details) {
  const sheet = await getOrCreateHistorySheet(doc);
  await sheet.addRow({ ID: uuid(), Type: type, Details: JSON.stringify(details), Timestamp: new Date().toISOString(), Undone: false });
}

async function getHistory(doc) {
  const sheet = doc.sheetsByTitle['history'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows.map(r => ({
    id: r.ID, type: r.Type, details: JSON.parse(r.Details), timestamp: r.Timestamp, undone: r.Undone === 'true'
  }));
}

async function redoAction(doc, actionId) {
  const sheet = doc.sheetsByTitle['history'];
  if (!sheet) return { error: 'No history' };
  const rows = await sheet.getRows();
  const action = rows.find(r => r.ID === actionId);
  if (!action) return { error: 'Action not found' };
  if (action.Undone !== 'true' && action.Undone !== true) return { error: 'Action is not undone' };

  const details = JSON.parse(action.Details);
  // Re-execute based on type
  switch (action.Type) {
    case 'requestMoney': await requestMoney(doc, details.amount); break;
    case 'newMonth': await processNewMonth(doc, details.month ? details.month.split(' ')[0] : undefined, details.month ? details.month.split(' ')[1] : undefined, details.dateReceived); break;
    case 'updateDateReceived': await updateDateReceived(doc, details.month, details.newDate); break;
    case 'useSavings': await useSavings(doc, details.amount, details.month); break;
    case 'topUpSavings': await topUpSavings(doc, details.months, details.totalAmount); break;
    case 'setOffsets': await setOffsets(doc, details.wallet, details.savings); break;
    case 'addTransaction': await addTransaction(doc, details.type, details.amount, details.note, details.attachment); break;
    case 'deleteTransaction': await deleteTransaction(doc, details.txBackup ? details.txBackup.id : null); break;
  }
  action.Undone = false;
  await action.save();
  return { success: true };
}

// ---------- Budget Settings ----------
async function getBudgetSettings(doc) {
  const sheet = await getOrCreateSheet(doc, 'budget_settings', ['category', 'budget_amount']);
  const rows = await sheet.getRows();
  const settings = { total: 0, refills: 0, savings_usage: 0 };
  rows.forEach(r => {
    if (r.category) {
      settings[r.category] = parseFloat(r.budget_amount) || 0;
    }
  });
  return settings;
}

async function setBudgetSettings(doc, total, refills, savings_usage) {
  const sheet = await getOrCreateSheet(doc, 'budget_settings', ['category', 'budget_amount']);
  const rows = await sheet.getRows();

  const categories = { total, refills, savings_usage };

  for (const cat of Object.keys(categories)) {
    const existing = rows.find(r => r.category === cat);
    if (existing) {
      existing.budget_amount = String(categories[cat]);
      await existing.save();
    } else {
      await sheet.addRow({ category: cat, budget_amount: String(categories[cat]) });
    }
  }
  return { success: true };
}

// ---------- Secure Notes ----------
async function getSecureNotes(doc) {
  const sheet = await getOrCreateSheet(doc, 'secure_notes', ['id', 'encrypted_data', 'iv', 'salt']);
  const rows = await sheet.getRows();
  return rows.map(r => ({
    id: r.id,
    encrypted_data: r.encrypted_data,
    iv: r.iv,
    salt: r.salt
  }));
}

async function saveSecureNote(doc, id, encrypted_data, iv, salt) {
  const sheet = await getOrCreateSheet(doc, 'secure_notes', ['id', 'encrypted_data', 'iv', 'salt']);
  const rows = await sheet.getRows();
  const existing = rows.find(r => r.id === id);
  if (existing) {
    existing.encrypted_data = encrypted_data;
    existing.iv = iv;
    existing.salt = salt;
    await existing.save();
  } else {
    await sheet.addRow({ id: id || uuid(), encrypted_data, iv, salt });
  }
  return { success: true };
}

async function deleteSecureNote(doc, id) {
  const sheet = await getOrCreateSheet(doc, 'secure_notes', ['id', 'encrypted_data', 'iv', 'salt']);
  const rows = await sheet.getRows();
  const existing = rows.find(r => r.id === id);
  if (existing) {
    await existing.delete();
    return { success: true };
  }
  return { error: 'Note not found' };
}

// ---------- Shared Links ----------
async function getSharedLinks(doc) {
  const sheet = await getOrCreateSheet(doc, 'shared_links', ['id', 'token', 'expiry', 'password_hash', 'created_at']);
  const rows = await sheet.getRows();
  return rows.map(r => ({
    id: r.id,
    token: r.token,
    expiry: r.expiry,
    password_hash: r.password_hash,
    created_at: r.created_at
  }));
}

async function createSharedLink(doc, expiryHours, passwordHash) {
  const sheet = await getOrCreateSheet(doc, 'shared_links', ['id', 'token', 'expiry', 'password_hash', 'created_at']);
  const id = uuid();
  const token = uuid().replace(/-/g, '');
  const now = new Date();
  const expiry = new Date(now.getTime() + parseFloat(expiryHours) * 60 * 60 * 1000).toISOString();
  await sheet.addRow({
    id,
    token,
    expiry,
    password_hash: passwordHash || '',
    created_at: now.toISOString()
  });
  return { success: true, id, token, expiry };
}

async function revokeSharedLink(doc, id) {
  const sheet = await getOrCreateSheet(doc, 'shared_links', ['id', 'token', 'expiry', 'password_hash', 'created_at']);
  const rows = await sheet.getRows();
  const existing = rows.find(r => r.id === id || r.token === id);
  if (existing) {
    await existing.delete();
    return { success: true };
  }
  return { error: 'Link not found' };
}

async function getSharedDashboard(doc, token, passwordHash) {
  const sheet = await getOrCreateSheet(doc, 'shared_links', ['id', 'token', 'expiry', 'password_hash', 'created_at']);
  const rows = await sheet.getRows();
  const link = rows.find(r => r.token === token);
  if (!link) {
    return { error: 'Link invalid' };
  }

  const now = new Date();
  const expiryDate = new Date(link.expiry);
  if (now > expiryDate) {
    return { error: 'Link expired' };
  }

  if (link.password_hash && link.password_hash.trim() !== '') {
    if (!passwordHash || passwordHash !== link.password_hash) {
      return { error: 'Password incorrect', password_required: true };
    }
  }

  // Fetch minimal dashboard data
  const offsets = await getOffsets(doc);
  const balance = await calculateBalance(doc);
  const savingsBalance = await calculateSavingsBalance(doc);
  const moneyFlowSettings = await getMoneyFlowSettings(doc);

  return {
    success: true,
    mamaAccount: balance + offsets.wallet,
    savingsBalance: savingsBalance + offsets.savings,
    totalRemaining: balance + offsets.wallet + savingsBalance + offsets.savings,
    moneyFlowSettings
  };
}

async function undoAction(doc, actionId) {
  const sheet = doc.sheetsByTitle['history'];
  if (!sheet) return { error: 'No history' };
  const rows = await sheet.getRows();
  const action = rows.find(r => r.ID === actionId);
  if (!action) return { error: 'Action not found' };
  if (action.Undone === 'true' || action.Undone === true) return { error: 'Already undone' };

  const details = JSON.parse(action.Details);
  // Reverse the action
  switch (action.Type) {
    case 'requestMoney':
      // Delete tx and restore allowances
      const txSheet = doc.sheetsByTitle['transactions'];
      if (txSheet) {
        const txRows = await txSheet.getRows();
        const tx = txRows.find(r => r.ID === details.txId);
        if (tx) { await tx.delete(); }
      }
      const statsSheet = doc.sheetsByTitle['allowance_stats'];
      if (statsSheet) {
        const statsRows = await statsSheet.getRows();
        for (const adj of details.adjustments) {
          const row = statsRows.find(r => r.rowNumber === adj.row);
          if (row) { row.Usage = Number(row.Usage)-adj.deducted; row.Balance = Number(row.Balance)+adj.deducted; await row.save(); }
        }
      }
      break;
    case 'useSavings':
      const tx2 = doc.sheetsByTitle['transactions'];
      if (tx2) {
        const tx2Rows = await tx2.getRows();
        const tx2Found = tx2Rows.find(r => r.ID === details.txId);
        if (tx2Found) await tx2Found.delete();
      }
      const savSheet = doc.sheetsByTitle['savings_stats'];
      if (savSheet) {
        const savRows = await savSheet.getRows();
        for (const adj of details.adjustments) {
          const row = savRows.find(r => r.rowNumber === adj.row);
          if (row) { row.Usage = Number(row.Usage)-adj.deducted; row.Balance = Number(row.Balance)+adj.deducted; await row.save(); }
        }
      }
      break;
    case 'topUpSavings':
      const tx3 = doc.sheetsByTitle['transactions'];
      if (tx3) {
        const tx3Rows = await tx3.getRows();
        const tx3Found = tx3Rows.find(r => r.ID === details.txId);
        if (tx3Found) await tx3Found.delete();
      }
      const savSheet3 = doc.sheetsByTitle['savings_stats'];
      if (savSheet3) {
        const savRows3 = await savSheet3.getRows();
        for (const adj of details.adjustments) {
          const row = savRows3.find(r => r.rowNumber === adj.row);
          if (row) {
            row.Usage = Number(row.Usage) + adj.restored;
            row.Balance = Number(row.Balance) - adj.restored;
            await row.save();
          }
        }
      }
      break;
    case 'newMonth':
      const tx4 = doc.sheetsByTitle['transactions'];
      if (tx4) {
        const tx4Rows = await tx4.getRows();
        const tx4Income = tx4Rows.find(r => r.ID === details.incomeId);
        if (tx4Income) await tx4Income.delete();
        const tx4Rows2 = await tx4.getRows();
        const tx4Savings = tx4Rows2.find(r => r.ID === details.savingsId);
        if (tx4Savings) await tx4Savings.delete();
      }
      const statsSheet4 = doc.sheetsByTitle['allowance_stats'];
      if (statsSheet4) {
        const statsRows4 = await statsSheet4.getRows();
        const rowToDelete = statsRows4.find(r => r.Month === details.month);
        if (rowToDelete) await rowToDelete.delete();
      }
      const savSheet4 = doc.sheetsByTitle['savings_stats'];
      if (savSheet4) {
        const savRows4 = await savSheet4.getRows();
        const rowToDelete = savRows4.find(r => r.Month === details.month);
        if (rowToDelete) await rowToDelete.delete();
      }
      break;
    case 'updateDateReceived':
      await updateDateReceived(doc, details.month, details.oldDate);
      break;
    case 'setOffsets':
      await setOffsets(doc, details.oldWallet, details.oldSavings);
      break;
    case 'addTransaction':
      const txSheet5 = doc.sheetsByTitle['transactions'];
      if (txSheet5) {
        const tx5Rows = await txSheet5.getRows();
        const tx5Found = tx5Rows.find(r => r.ID === details.txId);
        if (tx5Found) await tx5Found.delete();
      }
      break;
    case 'deleteTransaction':
      // To undo a deleteTransaction action, we restore the original deleted transaction
      // and reverse the adjustments that we applied during deletion.
      const txSheetRestore = await getOrCreateSheet(doc, 'transactions', ['Date', 'Type', 'Amount', 'Note', 'ID', 'Attachment']);
      await txSheetRestore.addRow({
        Date: details.txBackup.date,
        Type: details.txBackup.type,
        Amount: details.txBackup.amount,
        Note: details.txBackup.note,
        ID: details.txBackup.id,
        Attachment: details.txBackup.attachment || ''
      });

      // Restore adjustments (opposite of what deleteTransaction did, so we use the stored adjustments)
      if (details.resultAdjustments && details.resultAdjustments.length > 0) {
        for (const adj of details.resultAdjustments) {
          if (adj.sheet === 'allowance_stats') {
            const statsSheet = doc.sheetsByTitle['allowance_stats'];
            if (statsSheet) {
              const statsRows = await statsSheet.getRows();
              const statsRow = statsRows.find(r => r.rowNumber === adj.row || r.Month === adj.month);
              if (statsRow) {
                // Deducted was negative of subtraction, so we subtract adj.deducted to revert to previous state
                statsRow.Usage = (Number(statsRow.Usage) || 0) - adj.deducted;
                statsRow.Balance = (Number(statsRow.Balance) || 0) + adj.deducted;
                await statsRow.save();
              }
            }
          } else if (adj.sheet === 'savings_stats') {
            const savSheet = doc.sheetsByTitle['savings_stats'];
            if (savSheet) {
              const savRows = await savSheet.getRows();
              const savRow = savRows.find(r => r.rowNumber === adj.row || r.Month === adj.month);
              if (savRow) {
                if (adj.deducted !== undefined) {
                  savRow.Usage = (Number(savRow.Usage) || 0) - adj.deducted;
                  savRow.Balance = (Number(savRow.Balance) || 0) + adj.deducted;
                } else if (adj.restored !== undefined) {
                  savRow.Usage = (Number(savRow.Usage) || 0) - adj.restored;
                  savRow.Balance = (Number(savRow.Balance) || 0) + adj.restored;
                }
                await savRow.save();
              }
            }
          }
        }
      }

      // Mark original history action as active again if one was marked undone
      if (details.originalHistoryActionId) {
        const histRows = await sheet.getRows();
        const origHistAction = histRows.find(r => r.ID === details.originalHistoryActionId);
        if (origHistAction) {
          origHistAction.Undone = false;
          await origHistAction.save();
        }
      }
      break;
  }
  action.Undone = true;
  await action.save();
  return { success: true };
}
