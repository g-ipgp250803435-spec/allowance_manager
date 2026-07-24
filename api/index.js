const { GoogleSpreadsheet } = require('google-spreadsheet');
const cors = require('cors');

function uuid() {
  return 'xxxx-xxxx-xxxx-xxxx-xxxx'.replace(/x/g, () => Math.floor(Math.random() * 16).toString(16));
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
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
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
    const body = req.method === 'POST' ? await parseBody(req) : {};

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
    await sheet.addRow({ wallet_offset: 0, savings_offset: 0 });
    await sheet.addRow({ wallet_offset: 0, savings_offset: 0 });
  }
  const rows = await sheet.getRows();
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
  const txSheet = doc.sheetsByTitle['transactions'];
  if (!txSheet) return { error: 'No transactions sheet' };
  const txId = uuid();
  await txSheet.addRow({ Date: new Date().toISOString(), Type: 'refill', Amount: amount, Note: 'Transfer from Mama', ID: txId });

  const statsSheet = doc.sheetsByTitle['allowance_stats'];
  const rows = await statsSheet.getRows();
  const withBalance = rows.filter(r => Number(r.Balance) > 0);
  const monthOrder = { Januari:1,Februari:2,Mac:3,April:4,Mei:5,Jun:6,Julai:7,Ogos:8,September:9,Oktober:10,November:11,Disember:12 };
  withBalance.sort((a,b) => {
    const [am,ay]=a.Month.split(' '),[bm,by]=b.Month.split(' ');
    return (parseInt(ay)-parseInt(by))||((monthOrder[am]||0)-(monthOrder[bm]||0));
  });

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
  const monthOrder = { Januari:1,Februari:2,Mac:3,April:4,Mei:5,Jun:6,Julai:7,Ogos:8,September:9,Oktober:10,November:11,Disember:12 };
  withBalance.sort((a,b) => {
    const [am,ay]=a.Month.split(' '),[bm,by]=b.Month.split(' ');
    return (parseInt(ay)-parseInt(by))||((monthOrder[am]||0)-(monthOrder[bm]||0));
  });

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

  const monthOrder = { Januari:1,Februari:2,Mac:3,April:4,Mei:5,Jun:6,Julai:7,Ogos:8,September:9,Oktober:10,November:11,Disember:12 };
  months.sort((a,b) => {
    const [am,ay]=a.month.split(' '),[bm,by]=b.month.split(' ');
    return (parseInt(ay)-parseInt(by))||((monthOrder[am]||0)-(monthOrder[bm]||0));
  });

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
  if (action.Undone !== 'true') return { error: 'Action is not undone' };

  const details = JSON.parse(action.Details);
  // Re-execute based on type
  switch (action.Type) {
    case 'requestMoney': await requestMoney(doc, details.amount); break;
    case 'processNewMonth': await processNewMonth(doc); break;
    case 'useSavings': await useSavings(doc, details.amount); break;
    case 'topUpSavings': await topUpSavings(doc, details.months, details.totalAmount); break;
    case 'setOffsets': await setOffsets(doc, details.wallet, details.savings); break;
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
  if (action.Undone === 'true') return { error: 'Already undone' };

  const details = JSON.parse(action.Details);
  // Reverse the action
  switch (action.Type) {
    case 'requestMoney':
      // Delete tx and restore allowances
      const txSheet = doc.sheetsByTitle['transactions'];
      const txRows = await txSheet.getRows();
      const tx = txRows.find(r => r.ID === details.txId);
      if (tx) { await tx.delete(); }
      const statsSheet = doc.sheetsByTitle['allowance_stats'];
      const statsRows = await statsSheet.getRows();
      for (const adj of details.adjustments) {
        const row = statsRows.find(r => r.rowNumber === adj.row);
        if (row) { row.Usage = Number(row.Usage)-adj.deducted; row.Balance = Number(row.Balance)+adj.deducted; await row.save(); }
      }
      break;
    case 'useSavings':
      const tx2 = doc.sheetsByTitle['transactions'];
      const tx2Rows = await tx2.getRows();
      const tx2Found = tx2Rows.find(r => r.ID === details.txId);
      if (tx2Found) await tx2Found.delete();
      const savSheet = doc.sheetsByTitle['savings_stats'];
      const savRows = await savSheet.getRows();
      for (const adj of details.adjustments) {
        const row = savRows.find(r => r.rowNumber === adj.row);
        if (row) { row.Usage = Number(row.Usage)-adj.deducted; row.Balance = Number(row.Balance)+adj.deducted; await row.save(); }
      }
      break;
    case 'setOffsets':
      await setOffsets(doc, details.oldWallet, details.oldSavings);
      break;
  }
  action.Undone = true;
  await action.save();
  return { success: true };
}
