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

    if (body.token !== 'admin-session') return res.status(401).json({ error: 'Unauthorized' });

    try {
      const doc = await getDoc();
      let result;
      switch (action) {
        case 'getDashboard': result = await getDashboardData(doc); break;
        case 'requestMoney': result = await requestMoney(doc, body.amount); break;
        case 'processNewMonth': result = await processNewMonth(doc); break;
        case 'useSavings': result = await useSavings(doc, body.amount); break;
        case 'topUpSavings': result = await topUpSavings(doc, body.months, body.totalAmount); break;
        case 'setOffsets': result = await setOffsets(doc, body.walletOffset, body.savingsOffset); break;
        case 'getHistory': result = await getHistory(doc); break;
        case 'getTransactions': result = await getTransactions(doc); break;
        case 'addTransaction': result = await addTransaction(doc, body.type, body.amount, body.note); break;
        case 'redoAction': result = await redoAction(doc, body.actionId); break;
        case 'undoAction': result = await undoAction(doc, body.actionId); break;
        default: result = { error: 'Unknown action' };
      }
      res.json(result);
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message });
    }
  });
};

// ---------- Dashboard Data ----------
async function getDashboardData(doc) {
  const offsets = await getOffsets(doc);
  const balance = await calculateBalance(doc);
  const savingsBalance = await calculateSavingsBalance(doc);
  const stats = await getAllowanceStats(doc);
  const savingsStats = await getSavingsStats(doc);
  const transfers = await getMonthlyMomTransfers(doc);
  return {
    mamaAccount: balance + offsets.wallet,
    savingsBalance: savingsBalance + offsets.savings,
    totalRemaining: balance + offsets.wallet + savingsBalance + offsets.savings,
    offsets,
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
  const sheet = await getOrCreateSheet(doc, 'transactions', ['Date', 'Type', 'Amount', 'Note', 'ID']);
  const rows = await sheet.getRows();
  return rows.map(r => ({
    date: r.Date,
    type: r.Type,
    amount: Number(r.Amount) || 0,
    note: r.Note,
    id: r.ID
  }));
}

async function addTransaction(doc, type, amount, note) {
  const sheet = await getOrCreateSheet(doc, 'transactions', ['Date', 'Type', 'Amount', 'Note', 'ID']);
  const txId = uuid();
  const dateStr = new Date().toISOString();
  await sheet.addRow({ Date: dateStr, Type: type, Amount: amount, Note: note, ID: txId });
  await addHistory(doc, 'addTransaction', { type, amount, note, txId });
  return { success: true, transactionId: txId };
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

// ---------- Monthly Allocation ----------
async function processNewMonth(doc) {
  const now = new Date();
  const monthLabel = now.toLocaleString('default', { month: 'long', year: 'numeric' });
  const dateStr = now.toLocaleDateString('en-GB');
  const monthlyAllowance = 430, savings = 30, maybank = 100;
  const custodian = monthlyAllowance - savings - maybank;

  const txSheet = doc.sheetsByTitle['transactions'] || await doc.addSheet({ title: 'transactions', headerValues: ['Date', 'Type', 'Amount', 'Note', 'ID'] });
  const incomeId = uuid();
  const savingsId = uuid();
  await txSheet.addRow({ Date: now.toISOString(), Type: 'income', Amount: custodian, Note: `Monthly allocation to Custodian (${monthLabel})`, ID: incomeId });
  await txSheet.addRow({ Date: now.toISOString(), Type: 'savings', Amount: savings, Note: `Monthly savings (${monthLabel})`, ID: savingsId });

  const statsSheet = await getOrCreateSheet(doc, 'allowance_stats', ['Month', 'Date Received', 'Allowance Amount', 'Usage', 'Savings', 'Balance']);
  await statsSheet.addRow({ Month: monthLabel, 'Date Received': dateStr, 'Allowance Amount': monthlyAllowance, Usage: maybank, Savings: savings, Balance: custodian });

  const savSheet = await getOrCreateSheet(doc, 'savings_stats', ['Month', 'Savings', 'Usage', 'Balance']);
  await savSheet.addRow({ Month: monthLabel, Savings: savings, Usage: 0, Balance: savings });

  await addHistory(doc, 'newMonth', { month: monthLabel, incomeId, savingsId });
  return { success: true, month: monthLabel };
}

// ---------- Request Money ----------
async function requestMoney(doc, amount) {
  const txSheet = await getOrCreateSheet(doc, 'transactions', ['Date', 'Type', 'Amount', 'Note', 'ID']);
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
async function useSavings(doc, amount) {
  const txSheet = doc.sheetsByTitle['transactions'];
  const txId = uuid();
  await txSheet.addRow({ Date: new Date().toISOString(), Type: 'savings_usage', Amount: amount, Note: 'Used savings', ID: txId });

  const savSheet = doc.sheetsByTitle['savings_stats'];
  const rows = await savSheet.getRows();
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
  await addHistory(doc, 'useSavings', { amount, txId, adjustments });
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
    case 'newMonth': await processNewMonth(doc); break;
    case 'useSavings': await useSavings(doc, details.amount); break;
    case 'topUpSavings': await topUpSavings(doc, details.months, details.totalAmount); break;
    case 'setOffsets': await setOffsets(doc, details.wallet, details.savings); break;
    case 'addTransaction': await addTransaction(doc, details.type, details.amount, details.note); break;
  }
  action.Undone = false;
  await action.save();
  return { success: true };
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
  }
  action.Undone = true;
  await action.save();
  return { success: true };
}
