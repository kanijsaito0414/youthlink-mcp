import { google } from 'googleapis';
import express from 'express';
import cors from 'cors';

const SPREADSHEET_ID = '1cSQWEdlaoAPmCYsHVnz5td2yaqNBrfTyVyIIycG9fds';
const MONTH_PREFIX: Record<number, string> = {
  1: '1', 2: '2', 3: '3', 4: '４', 5: '5', 6: '6',
  7: '7', 8: '8', 9: '9', 10: '10', 11: '11', 12: '12'
};

const TYPE_RULES: [string, string][] = [
  ['ジョブハント', 'GP'], ['DYM', 'GP'], ['シンアド', 'GP'],
  ['GRIT', 'GP'], ['ワンキャリア', 'GP'], ['ワンキャリ', 'GP'],
  ['チアキャリ', 'GP'], ['ルビーイン', 'GP'],
  ['FC)', '面談'], ['Ai)', '面談'], ['AI)', '面談'],
  ['ジール', '面談'], ['ジョーカツ', '面談'], ['UZUZ', '面談'],
  ['スターマイン', '面談'], ['正直エージェント', '面談'],
  ['イロダス', '面談'], ['HRteam', '面談'],
  ['アクセス', 'その他'], ['みん就', 'その他'], ['社長飯', 'その他'],
  ['キャリガク', 'その他'], ['Winc', 'その他'], ['Shabell', 'その他'],
];

function getSheets() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY || '{}');
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

function detectType(name: string): string {
  for (const [key, val] of TYPE_RULES) {
    if (name.includes(key)) return val;
  }
  return 'GP';
}

function parseDate(s: string) {
  const md = s.match(/^(\d{1,2})[\/月](\d{1,2})/);
  if (md) return { month: +md[1], day: +md[2], fmt: `${md[1]}/${md[2]}` };
  const d = new Date(s);
  return { month: d.getMonth() + 1, day: d.getDate(), fmt: `${d.getMonth()+1}/${d.getDate()}` };
}

const app = express();
app.use(express.json());
app.use(cors());

app.get('/', (_req, res) => {
  res.json({ status: 'ok', name: 'YouthLink MCP Server', version: '1.0.0' });
});

app.get('/health', (_req, res) => {
  res.json({ status: 'healthy' });
});

app.post('/add', async (req, res) => {
  try {
    const { 案件名, 日付, 学生名, 担当 = '齋藤', 案件タイプ, 時間 = '', 経由 = '', 受注額 = '', 支払い額 = '', 備考 = '' } = req.body;
    if (!案件名 || !日付 || !学生名) {
      return res.status(400).json({ error: '案件名・日付・学生名は必須です' });
    }
    const date = parseDate(日付);
    const type = 案件タイプ || detectType(案件名);
    const m = MONTH_PREFIX[date.month] || String(date.month);
    const sheetName = `${m}月（${type}）${担当}`;
    const sheets = getSheets();
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const names = meta.data.sheets?.map(s => s.properties?.title || '') || [];
    if (!names.includes(sheetName)) {
      return res.status(404).json({ error: `シート「${sheetName}」が存在しません`, available: names.filter(n => n.includes('月')) });
    }
    const existing = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName });
    const rows = existing.data.values || [];
    let lastRow = 0;
    for (let i = rows.length - 1; i >= 0; i--) {
      if (rows[i].some((c: any) => c !== '' && c !== 'FALSE' && c !== false) && rows[i][0] !== '案件名' && rows[i][0] !== '対面案件名') {
        lastRow = i + 1; break;
      }
    }
    const newRow = lastRow + 1;
    const profit = (受注額 && 支払い額) ? Number(受注額) - Number(支払い額) : '';
    const rowData = [案件名, date.fmt, 時間, 学生名, 受注額 ? Number(受注額) : '', 支払い額 ? Number(支払い額) : '', profit, '', '', 経由, 'FALSE', 備考];
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A${newRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [rowData] },
    });
    res.json({ success: true, sheet: sheetName, row: newRow, message: `✅ ${学生名}を${sheetName}の${newRow}行目に追加しました` });
  } catch (e: any) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/read', async (req, res) => {
  try {
    const { シート名, 範囲 } = req.body;
    const sheets = getSheets();
    const range = 範囲 ? `${シート名}!${範囲}` : シート名;
    const result = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range });
    res.json({ data: result.data.values || [] });
  } catch (e: any) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/sheets', async (_req, res) => {
  try {
    const sheets = getSheets();
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const names = meta.data.sheets?.map(s => s.properties?.title || '') || [];
    res.json({ sheets: names });
  } catch (e: any) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`YouthLink MCP Server running on port ${PORT}`);
});
