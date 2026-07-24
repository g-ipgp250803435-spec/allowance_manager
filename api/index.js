const { GoogleSpreadsheet } = require('google-spreadsheet');
const cors = require('cors');

// Simple UUID generator
function uuid() {
  return 'xxxx-xxxx-xxxx'.replace(/x/g, () => Math.floor(Math.random() * 16).toString(16));
}

// Parse JSON body
async function parseBody(req) {
  return new Promise(resolve => {
    let body = '';
    req.on('data', chunk => (body += chunk.toString()));
    req.on('end', () => {
      try { resolve(JSON.parse(body)); } catch { resolve({}); }
    });
  });
}

// Google Sheets authentication
async function getDoc() {
  const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID);
  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  });
  await doc.loadInfo();
  return doc;
}

// Helper to get or create a sheet
async function getOrCreateSheet(doc, title, headers) {
  let sheet = doc.sheetsByTitle[title];
  if (!sheet) {
    sheet = await doc.addSheet({ title, headerValues: headers });
  }
  return sheet;
}

module.exports = async (req, res) => {
  await cors()(req, res, async () => {
    const { action } = req.query;
    const body = req.method === 'POST' ? await parseBody(req) : {};

    // ----- Login (no token) -----
    if (action === 'login') {
      if (body.username === process.env.ADMIN_USERNAME &&
          body.password === process.env.ADMIN_PASSWORD) {
        return res.json({ success: true, token: 'admin-session' });
      }
      return res.status(401).json({ error: 'Invalid credentials' });
    }

    // All other actions require token
    if (body.token !== 'admin-session') {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    try {
      const doc = await getDoc();
      let result;

      switch (action) {
        case 'getDashboard':
          result = await getDashboardData(doc);
          break;
        case 'requestMoney':
          result = await requestMoney(doc, body.amount);
          break;
        case 'processNewMonth':
          result = await processNewMonth(doc);
          break;
        case 'useSavings':
          result = await useSavings(doc, body.amount);
          break;
        case 'topUpSavings':
          result = await topUpSavings(doc, body.months, body.totalAmount);
          break;
        case 'setOffsets':
          result = await setOffsets(doc, body.walletOffset, body.savingsOffset);
          break;
        default:
          result = { error: 'Unknown action' };
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
  return {
    wallet: Number(rows[0]?.wallet_offset) || 0,
    savings: Number(rows[1]?.savings_offset) || 0
  };
}

async function setOffsets(doc, wallet, savings) {
  const sheet = doc.sheetsByTitle['manual_offsets'];
  if (!sheet) {
    const newSheet = await doc.addSheet({ title: 'manual_offsets', headerValues: ['wallet_offset', 'savings_offset'] });
    await newSheet.addRow({ wallet_offset: 0, savings_offset: 0 });
    await newSheet.addRow({ wallet_offset: 0, savings_offset: 0 });
  }
  const rows = await sheet.getRows();
  rows[0].wallet_offset = wallet;
  rows[1].savings_offset = savings;
  await rows[0].save();
  await rows[1].save();
  return { success: true };
}

// ---------- Monthly Allocation ----------
async function processNewMonth(doc) {
  const now = new Date();
  const monthLabel = now.toLocaleString('default', { month: 'long', year: 'numeric' });
  const dateStr = now.toLocaleDateString('en-GB'); // dd/MM/yyyy
  const monthlyAllowance = 430, savings = 30, maybank = 100;
  const custodian = monthlyAllowance - savings - maybank;

  // Transactions
  const txSheet = doc.sheetsByTitle['transactions'] ||
    await doc.addSheet({ title: 'transactions', headerValues: ['Date', 'Type', 'Amount', 'Note', 'ID'] });
  await txSheet.addRow({ Date: now.toISOString(), Type: 'income', Amount: custodian, Note: `Monthly allocation to Custodian (${monthLabel})`, ID: uuid() });
  await txSheet.addRow({ Date: now.toISOString(), Type: 'savings', Amount: savings, Note: `Monthly savings (${monthLabel})`, ID: uuid() });

  // Allowance stats
  const statsSheet = await getOrCreateSheet(doc, 'allowance_stats', ['Month', 'Date Received', 'Allowance Amount', 'Usage', 'Savings', 'Balance']);
  await statsSheet.addRow({ Month: monthLabel, 'Date Received': dateStr, 'Allowance Amount': monthlyAllowance, Usage: maybank, Savings: savings, Balance: custodian });

  // Savings stats
  const savSheet = await getOrCreateSheet(doc, 'savings_stats', ['Month', 'Savings', 'Usage', 'Balance']);
  await savSheet.addRow({ Month: monthLabel, Savings: savings, Usage: 0, Balance: savings });

  return { success: true, month: monthLabel };
}

// ---------- Request Money (deduct oldest first) ----------
async function requestMoney(doc, amount) {
  const txSheet = doc.sheetsByTitle['transactions'];
  if (!txSheet) return { error: 'No transactions sheet' };
  const txId = uuid();
  await txSheet.addRow({ Date: new Date().toISOString(), Type: 'refill', Amount: amount, Note: 'Transfer from Mama', ID: txId });

  const statsSheet = doc.sheetsByTitle['allowance_stats'];
  if (!statsSheet) return { error: 'No allowance_stats sheet' };

  const rows = await statsSheet.getRows();
  const withBalance = rows.filter(r => Number(r.Balance) > 0);
  // Sort oldest first using month name + year
  const monthOrder = { Januari:1,Februari:2,Mac:3,April:4,Mei:5,Jun:6,Julai:7,Ogos:8,September:9,Oktober:10,November:11,Disember:12 };
  withBalance.sort((a, b) => {
    const [am, ay] = a.Month.split(' '), [bm, by] = b.Month.split(' ');
    const ayear = parseInt(ay), byear = parseInt(by);
    if (ayear !== byear) return ayear - byear;
    return (monthOrder[am] || 0) - (monthOrder[bm] || 0);
  });

  let remaining = amount;
  const adjustments = [];
  for (const row of withBalance) {
    if (remaining <= 0) break;
    const bal = Number(row.Balance);
    const deduct = Math.min(remaining, bal);
    if (deduct > 0) {
      row.Usage = Number(row.Usage) + deduct;
      row.Balance = bal - deduct;
      await row.save();
      adjustments.push({ row: row.rowNumber, deducted: deduct });
      remaining -= deduct;
    }
  }
  return { transactionId: txId, adjustments };
}

// ---------- Use Savings ----------
async function useSavings(doc, amount) {
  const txSheet = doc.sheetsByTitle['transactions'];
  if (!txSheet) return { error: 'No transactions sheet' };
  const txId = uuid();
  await txSheet.addRow({ Date: new Date().toISOString(), Type: 'savings_usage', Amount: amount, Note: 'Used savings', ID: txId });

  const savSheet = doc.sheetsByTitle['savings_stats'];
  if (!savSheet) return { error: 'No savings_stats sheet' };

  const rows = await savSheet.getRows();
  const withBalance = rows.filter(r => Number(r.Balance) > 0);
  const monthOrder = { Januari:1,Februari:2,Mac:3,April:4,Mei:5,Jun:6,Julai:7,Ogos:8,September:9,Oktober:10,November:11,Disember:12 };
  withBalance.sort((a, b) => {
    const [am, ay] = a.Month.split(' '), [bm, by] = b.Month.split(' ');
    const ayear = parseInt(ay), byear = parseInt(by);
    if (ayear !== byear) return ayear - byear;
    return (monthOrder[am] || 0) - (monthOrder[bm] || 0);
  });

  let remaining = amount;
  const adjustments = [];
  for (const row of withBalance) {
    if (remaining <= 0) break;
    const bal = Number(row.Balance);
    const deduct = Math.min(remaining, bal);
    if (deduct > 0) {
      row.Usage = Number(row.Usage) + deduct;
      row.Balance = bal - deduct;
      await row.save();
      adjustments.push({ row: row.rowNumber, deducted: deduct });
      remaining -= deduct;
    }
  }
  return { transactionId: txId, adjustments };
}

// ---------- Top Up Savings ----------
async function topUpSavings(doc, months, totalAmount) {
  const txSheet = doc.sheetsByTitle['transactions'];
  if (!txSheet) return { error: 'No transactions sheet' };
  const txId = uuid();
  await txSheet.addRow({ Date: new Date().toISOString(), Type: 'savings_topup', Amount: totalAmount, Note: 'Top-up savings', ID: txId });

  const savSheet = doc.sheetsByTitle['savings_stats'];
  if (!savSheet) return { error: 'No savings_stats sheet' };

  const rows = await savSheet.getRows();
  const adjustments = [];
  let remaining = totalAmount;

  // Sort months oldest first (as provided)
  const monthOrder = { Januari:1,Februari:2,Mac:3,April:4,Mei:5,Jun:6,Julai:7,Ogos:8,September:9,Oktober:10,November:11,Disember:12 };
  const sortedMonths = months.slice().sort((a, b) => {
    const [am, ay] = a.month.split(' '), [bm, by] = b.month.split(' ');
    return (parseInt(ay) - parseInt(by)) || ((monthOrder[am] || 0) - (monthOrder[bm] || 0));
  });

  for (const sm of sortedMonths) {
    if (remaining <= 0) break;
    const row = rows.find(r => r.Month.trim() === sm.month);
    if (!row) continue;
    const currentUsage = Number(row.Usage) || 0;
    const currentBalance = Number(row.Balance) || 0;
    const maxRestorable = currentUsage;
    const restore = Math.min(remaining, maxRestorable);
    if (restore > 0) {
      row.Usage = currentUsage - restore;
      row.Balance = currentBalance + restore;
      await row.save();
      adjustments.push({ row: row.rowNumber, month: sm.month, restored: restore, previousUsage: currentUsage, previousBalance: currentBalance });
      remaining -= restore;
    }
  }
  return { transactionId: txId, adjustments, unusedAmount: remaining };
}

// ---------- Stats (getAllowanceStats, getSavingsStats, getMonthlyMomTransfers) ----------
async function getAllowanceStats(doc) {
  const sheet = doc.sheetsByTitle['allowance_stats'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows.map(r => ({
    month: r.Month,
    dateReceived: r['Date Received'],
    allowance: Number(r['Allowance Amount']),
    usage: Number(r.Usage),
    savings: Number(r.Savings),
    balance: Number(r.Balance)
  }));
}

async function getSavingsStats(doc) {
  const sheet = doc.sheetsByTitle['savings_stats'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows.map(r => ({
    month: r.Month,
    savings: Number(r.Savings),
    usage: Number(r.Usage),
    balance: Number(r.Balance)
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
