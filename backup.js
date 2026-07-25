const https = require('https');
const fs    = require('fs');
const path  = require('path');
const XLSX  = require(path.join(__dirname, 'xlsx.min.js'));

const PROJECT_ID  = 'miznonpool';
const API_KEY     = 'AIzaSyD_iu3GJPjkVU3ATCI0qonO2YA_y0RvX5c';
const COLLECTIONS = ['customers', 'products', 'invoices', 'settings'];
const BACKUP_DIR  = path.join(__dirname, 'backups');
const KEEP_DAYS   = 30;

function post(url, body) {
  return new Promise((resolve, reject) => {
    const buf = Buffer.from(body);
    const req = https.request(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': buf.length },
    }, res => {
      let raw = '';
      res.on('data', c => raw += c);
      res.on('end', () => {
        try { resolve(JSON.parse(raw)); }
        catch(e) { reject(new Error('Parse error: ' + raw.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    req.write(buf);
    req.end();
  });
}

// Get a short-lived Firebase ID token via anonymous sign-in
async function getAuthToken() {
  const data = await post(
    `https://identitytoolkit.googleapis.com/v1/accounts:signUp?key=${API_KEY}`,
    JSON.stringify({ returnSecureToken: true })
  );
  if (!data.idToken) throw new Error('Auth failed: ' + JSON.stringify(data));
  return data.idToken;
}

function get(url, token) {
  return new Promise((resolve, reject) => {
    const headers = token ? { Authorization: `Bearer ${token}` } : {};
    https.get(url, { headers }, res => {
      let raw = '';
      res.on('data', c => raw += c);
      res.on('end', () => {
        try { resolve(JSON.parse(raw)); }
        catch(e) { reject(new Error('Parse error: ' + raw.slice(0, 200))); }
      });
    }).on('error', reject);
  });
}

// Convert Firestore REST format to plain JS objects
function fsToJS(val) {
  if (!val || typeof val !== 'object') return val;
  if ('stringValue'  in val) return val.stringValue;
  if ('integerValue' in val) return Number(val.integerValue);
  if ('doubleValue'  in val) return val.doubleValue;
  if ('booleanValue' in val) return val.booleanValue;
  if ('nullValue'    in val) return null;
  if ('arrayValue'   in val) return (val.arrayValue.values || []).map(fsToJS);
  if ('mapValue'     in val) {
    const obj = {};
    for (const [k, v] of Object.entries(val.mapValue.fields || {})) obj[k] = fsToJS(v);
    return obj;
  }
  return val;
}

async function fetchCollection(name, token) {
  const base = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/${name}?key=${API_KEY}&pageSize=300`;
  let all = [], nextPage = null;
  do {
    const url  = nextPage ? `${base}&pageToken=${nextPage}` : base;
    const data = await get(url, token);
    (data.documents || []).forEach(doc => {
      const obj = {};
      for (const [k, v] of Object.entries(doc.fields || {})) obj[k] = fsToJS(v);
      all.push(obj);
    });
    nextPage = data.nextPageToken || null;
  } while (nextPage);
  return all;
}

async function run() {
  if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });

  process.stdout.write('  מתחבר ל-Firebase...');
  let token = null;
  try {
    token = await getAuthToken();
    console.log(' ✓ (מאומת)');
  } catch(e) {
    console.log(' ⚠ ללא אימות (Anonymous Auth לא מופעל) — ממשיך בגישה פתוחה');
  }

  const now    = new Date();
  const stamp  = now.toISOString().replace(/T/, '_').replace(/:/g, '-').slice(0, 19);
  const result = { _backup_time: now.toISOString() };

  for (const col of COLLECTIONS) {
    process.stdout.write(`  גובה: ${col}...`);
    result[col] = await fetchCollection(col, token);
    console.log(` ${result[col].length} רשומות`);
  }

  const filename     = path.join(BACKUP_DIR, `backup_${stamp}.json`);
  const xlsxFilename = path.join(BACKUP_DIR, `backup_${stamp}.xlsx`);
  fs.writeFileSync(filename, JSON.stringify(result, null, 2), 'utf8');
  console.log(`\nנשמר JSON: ${filename}`);

  // Build Excel workbook
  const wb = XLSX.utils.book_new();

  // Sheet 1: Invoices (one row per item)
  const invRows = [['מס׳ חשבון', 'תאריך', 'שעה', 'שם לקוח', 'מזהה לקוח', 'מוצר', 'מחיר יחידה', 'כמות', 'סה"כ שורה', 'סה"כ חשבון', 'הנחה %']];
  (result.invoices || [])
    .slice()
    .sort((a, b) => {
      const da = (a.date || '') + ' ' + (a.time || '');
      const db = (b.date || '') + ' ' + (b.time || '');
      return da < db ? -1 : da > db ? 1 : 0;
    })
    .forEach(inv => {
      (inv.items || []).forEach(item => {
        invRows.push([
          inv.num,
          inv.date || '',
          inv.time || '',
          inv.customerName || '',
          inv.customerId || '',
          item.name || '',
          item.price,
          item.qty,
          item.price * item.qty,
          inv.total,
          inv.discount ? inv.discount + '%' : '',
        ]);
      });
    });
  const wsInv = XLSX.utils.aoa_to_sheet(invRows);
  wsInv['!cols'] = [12, 12, 8, 20, 12, 22, 14, 8, 14, 14, 8].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsInv, 'חשבוניות');

  // Sheet 2: Customers
  const custRows = [['מזהה', 'שם', 'טלפון', 'תאריך הצטרפות']];
  (result.customers || []).forEach(c => custRows.push([c.id, c.name, c.phone || '', c.joinDate || '']));
  const wsCust = XLSX.utils.aoa_to_sheet(custRows);
  wsCust['!cols'] = [12, 24, 16, 14].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsCust, 'מנויים');

  // Sheet 3: Products
  const prodRows = [['מזהה', 'שם מוצר', 'מחיר', 'מלאי']];
  (result.products || []).forEach(p => prodRows.push([p.id, p.name, p.price, p.stock ?? '']));
  const wsProd = XLSX.utils.aoa_to_sheet(prodRows);
  wsProd['!cols'] = [8, 24, 10, 8].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsProd, 'מוצרים');

  const xlsxBuf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  fs.writeFileSync(xlsxFilename, xlsxBuf);
  console.log(`נשמר Excel: ${xlsxFilename}`);

  // Remove backups older than KEEP_DAYS
  const cutoff = Date.now() - KEEP_DAYS * 86400 * 1000;
  fs.readdirSync(BACKUP_DIR)
    .filter(f => f.startsWith('backup_') && (f.endsWith('.json') || f.endsWith('.xlsx')))
    .forEach(f => {
      const full = path.join(BACKUP_DIR, f);
      if (fs.statSync(full).mtimeMs < cutoff) {
        fs.unlinkSync(full);
        console.log(`נמחק ישן: ${f}`);
      }
    });
}

run().catch(err => {
  console.error('שגיאה בגיבוי:', err.message);
  process.exit(1);
});
