import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import express from 'express';

const SPREADSHEET_ID = '1apDuNqifsWD0IkWHAPpO31hWbVkoMSUc997NRVqcy0Y';

const MONTH_PREFIX: Record<number, string> = {
  1: '1', 2: '2', 3: '3', 4: '４', 5: '5', 6: '6',
  7: '7', 8: '8', 9: '9', 10: '10', 11: '11', 12: '12'
};

const TYPE_RULES: [string, string][] = [
  ['ジョブハント', 'GP'], ['DYM', 'GP'], ['シンアド', 'GP'],
  ['GRIT', 'GP'], ['ワンキャリア', 'GP'], ['ワンキャリ', 'GP'],
  ['チアキャリ', 'GP'], ['ルビーイン', 'GP'], ['ウェルハンティング', 'GP'],
  ['FC)', '面談'], ['Ai)', '面談'], ['AI)', '面談'],
  ['ジール', '面談'], ['ジョーカツ', '面談'], ['UZUZ', '面談'],
  ['スターマイン', '面談'], ['正直エージェント', '面談'],
  ['イロダス', '面談'], ['HRteam', '面談'], ['0円就活', '面談'],
  ['アクセス', 'その他'], ['みん就', 'その他'], ['社長飯', 'その他'],
  ['キャリガク', 'その他'], ['Winc', 'その他'], ['Shabell', 'その他'],
  ['SRB', 'その他'], ['しゃべる', 'その他'], ['アスキャリ', 'その他'],
];

function getAuthClient() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY || '{}');
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function getSheetsClient() {
  const auth = getAuthClient();
  return google.sheets({ version: 'v4', auth });
}

function detectType(caseName: string): string {
  for (const [key, val] of TYPE_RULES) {
    if (caseName.includes(key)) return val;
  }
  return 'GP';
}

function parseDate(dateStr: string): { month: number; day: number; formatted: string } {
  const str = String(dateStr).trim();
  const md = str.match(/^(\d{1,2})[\/月](\d{1,2})/);
  if (md) {
    const month = parseInt(md[1]);
    const day = parseInt(md[2]);
    return { month, day, formatted: `${month}/${day}` };
  }
  const ymd = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymd) {
    const month = parseInt(ymd[2]);
    const day = parseInt(ymd[3]);
    return { month, day, formatted: `${month}/${day}` };
  }
  const d = new Date(str);
  return { month: d.getMonth() + 1, day: d.getDate(), formatted: `${d.getMonth() + 1}/${d.getDate()}` };
}

function buildSheetName(month: number, type: string, person: string): string {
  const m = MONTH_PREFIX[month] || String(month);
  return `${m}月（${type}）${person}`;
}

async function getSheetNames(): Promise<string[]> {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  return res.data.sheets?.map(s => s.properties?.title || '') || [];
}

async function readSheet(sheetName: string, range?: string): Promise<any[][]> {
  const sheets = await getSheetsClient();
  const fullRange = range ? `${sheetName}!${range}` : sheetName;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: fullRange });
  return res.data.values || [];
}

async function appendRow(sheetName: string, rowData: any[]): Promise<{ row: number }> {
  const sheets = await getSheetsClient();
  const existing = await readSheet(sheetName);
  let lastDataRow = 0;
  for (let i = existing.length - 1; i >= 0; i--) {
    const row = existing[i];
    const hasContent = row.some((cell: any) => cell !== '' && cell !== false && cell !== null && cell !== undefined && cell !== 'FALSE');
    if (hasContent && row[0] !== '案件名' && row[0] !== '対面案件名') { lastDataRow = i + 1; break; }
  }
  const newRow = lastDataRow + 1;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A${newRow}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [rowData] },
  });
  return { row: newRow };
}

async function updateCell(sheetName: string, cellRange: string, value: any): Promise<void> {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!${cellRange}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[value]] },
  });
}

const server = new Server(
  { name: 'youthlink-sheets-mcp', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: 'add_booking',
      description: '学生の予約をスプレッドシートに追加する',
      inputSchema: {
        type: 'object',
        properties: {
          案件名: { type: 'string' },
          日付: { type: 'string' },
          学生名: { type: 'string' },
          担当: { type: 'string', enum: ['齋藤', '濱谷'], default: '齋藤' },
          案件タイプ: { type: 'string', enum: ['GP', '面談', 'その他'] },
          時間: { type: 'string' },
          経由: { type: 'string' },
          受注額: { type: 'number' },
          支払い額: { type: 'number' },
          備考: { type: 'string' },
        },
        required: ['案件名', '日付', '学生名'],
      },
    },
    {
      name: 'read_sheet',
      description: '指定シートの内容を読む',
      inputSchema: { type: 'object', properties: { シート名: { type: 'string' }, 範囲: { type: 'string' } }, required: ['シート名'] },
    },
    {
      name: 'list_sheets',
      description: 'シートタブ一覧を取得',
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'update_booking_status',
      description: '既存行のステータスを更新',
      inputSchema: { type: 'object', properties: { シート名: { type: 'string' }, 行番号: { type: 'number' }, 列: { type: 'string' }, 値: {} }, required: ['シート名', '行番号', '列', '値'] },
    },
    {
      name: 'search_student',
      description: '学生名で全月横断検索',
      inputSchema: { type: 'object', properties: { 学生名: { type: 'string' } }, required: ['学生名'] },
    },
    {
      name: 'get_monthly_summary',
      description: '指定月の送客サマリーを集計',
      inputSchema: { type: 'object', properties: { 月: { type: 'number' }, 担当: { type: 'string', enum: ['齋藤', '濱谷', '両方'] } }, required: ['月'] },
    },
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  try {
    switch (name) {
      case 'add_booking': {
        const { 案件名, 日付, 学生名, 担当 = '齋藤', 案件タイプ, 時間 = '', 経由 = '', 受注額 = '', 支払い額 = '', 備考 = '' } = args as any;
        const dateInfo = parseDate(日付);
        const type = 案件タイプ || detectType(案件名);
        const sheetName = buildSheetName(dateInfo.month, type, 担当);
        const sheetNames = await getSheetNames();
        if (!sheetNames.includes(sheetName)) {
          return { content: [{ type: 'text', text: `⚠️ シート「${sheetName}」が存在しません。利用可能: ${sheetNames.filter(s => s.includes('月')).join(', ')}` }] };
        }
        const profit = (受注額 && 支払い額) ? Number(受注額) - Number(支払い額) : '';
        const rowData = [案件名, dateInfo.formatted, 時間, 学生名, 受注額 ? Number(受注額) : '', 支払い額 ? Number(支払い額) : '', profit, '', '', 経由, 'FALSE', 備考];
        const result = await appendRow(sheetName, rowData);
        return { content: [{ type: 'text', text: `✅ 追加完了\nシート: ${sheetName}\n行: ${result.row}行目\n案件: ${案件名} / ${dateInfo.formatted}\n学生: ${学生名}` }] };
      }
      case 'read_sheet': {
        const { シート名, 範囲 } = args as any;
        const data = await readSheet(シート名, 範囲);
        if (data.length === 0) return { content: [{ type: 'text', text: `「${シート名}」にデータがありません。` }] };
        const formatted = data.filter(row => row.some((cell: any) => cell !== '' && cell !== 'FALSE')).map((row, i) => `${i + 1}: ${row.join(' | ')}`).join('\n');
        return { content: [{ type: 'text', text: `【${シート名}】\n${formatted}` }] };
      }
      case 'list_sheets': {
        const names = await getSheetNames();
        return { content: [{ type: 'text', text: `【月次】\n${names.filter(n => n.includes('月')).join('\n')}\n\n【その他】\n${names.filter(n => !n.includes('月')).join('\n')}` }] };
      }
      case 'update_booking_status': {
        const { シート名, 行番号, 列, 値 } = args as any;
        const colMap: Record<string, string> = { '案件名': 'A', '日付': 'B', '時間': 'C', '学生名': 'D', '参加者氏名': 'D', '受注額': 'E', '支払い額': 'F', '利益': 'G', '着座': 'H', '成果': 'I', '経由': 'J', '支払い確認': 'K', '備考': 'L' };
        await updateCell(シート名, `${colMap[列] || 列}${行番号}`, 値);
        return { content: [{ type: 'text', text: `✅ ${シート名} ${行番号}行目「${列}」→「${値}」に更新しました。` }] };
      }
      case 'search_student': {
        const { 学生名 } = args as any;
        const sheetNames = await getSheetNames();
        const results: string[] = [];
        for (const sn of sheetNames.filter(n => n.includes('月') && (n.includes('齋藤') || n.includes('濱谷')))) {
          const data = await readSheet(sn);
          data.forEach((row, i) => { if (row[3] && String(row[3]).includes(学生名)) results.push(`・${sn} / ${i + 1}行目 / ${row[0]} ${row[1]} / 受注:${row[4] || '-'} 支払:${row[5] || '-'}`); });
        }
        return { content: [{ type: 'text', text: results.length ? `【${学生名} の記録 (${results.length}件)】\n${results.join('\n')}` : `「${学生名}」の記録はありません。` }] };
      }
      case 'get_monthly_summary': {
        const { 月, 担当 = '両方' } = args as any;
        const sheetNames = await getSheetNames();
        const targets = sheetNames.filter(n => { const m = MONTH_PREFIX[月]; co
