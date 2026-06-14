const XLSX = require('xlsx');
const path = require('path');

// 脂肪酸成分表編 Excel（第1表 可食部100g当たり）の構造分析スクリプト。
// シート構成・ヘッダ階層・食品番号列・対象カラム（飽和脂肪酸／n-3／n-6）を確認する。
// 本表（convert-excel.js）への join 設計の前提を把握するために使う。

// 取り込み対象の3カラム。脂肪酸成分表編の成分識別子（行4）を基準にする。
const TARGET_IDS = {
  FASAT: '飽和脂肪酸',
  FAPUN3: 'n-3系多価不飽和脂肪酸',
  FAPUN6: 'n-6系多価不飽和脂肪酸',
};

function clean(v) {
  return v == null ? '' : v.toString().replace(/\r\n|\r|\n/g, '').trim();
}

function analyze(filePath) {
  console.log(`分析中: ${filePath}`);
  console.log('='.repeat(60));

  const workbook = XLSX.readFile(filePath);
  console.log(`シート数: ${workbook.SheetNames.length}`);
  console.log(`シート名: ${workbook.SheetNames.map(s => s.trim()).join(', ')}`);
  console.log();

  // 「表全体」シートを代表として、ヘッダ階層を詳しく見る。
  const mainSheet = workbook.SheetNames.find(s => s.trim() === '表全体') || workbook.SheetNames[0];
  const ws = workbook.Sheets[mainSheet];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

  console.log(`代表シート: ${mainSheet}`);
  console.log(`行数: ${rows.length} / 列数: ${(rows[3] || []).length}`);
  console.log();

  // ヘッダ階層：行3=日本語名 / 行4=成分識別子 / 行5=単位。データは行6以降。
  const nameRow = rows[3] || [];
  const idRow = rows[4] || [];
  const unitRow = rows[5] || [];

  console.log('ヘッダ階層（列ごと: 日本語名 / 成分識別子 / 単位）:');
  for (let i = 0; i < nameRow.length; i++) {
    console.log(`  [${i}] ${clean(nameRow[i])} | ${clean(idRow[i])} | ${clean(unitRow[i])}`);
  }
  console.log();

  // 食品番号列の特定（5桁数字文字列が並ぶ列）。
  const sampleData = rows.slice(6, 16);
  const codeCol = (rows[3] || []).findIndex(h => clean(h) === '食品番号');
  console.log(`食品番号列: index=${codeCol}（例: ${sampleData.map(r => r && r[codeCol]).filter(Boolean).slice(0, 5).join(', ')}）`);
  const nameCol = (rows[3] || []).findIndex(h => clean(h) === '食品名');
  console.log(`食品名列: index=${nameCol}`);
  console.log();

  // 対象3カラムの列位置を成分識別子から特定。
  console.log('取り込み対象カラム:');
  for (const [id, label] of Object.entries(TARGET_IDS)) {
    const idx = idRow.findIndex(v => clean(v) === id);
    console.log(`  ${label} (${id}): index=${idx} / 単位=${clean(unitRow[idx])} / 日本語=${clean(nameRow[idx])}`);
  }
  console.log();

  // データサンプル（対象カラムのみ）。特殊値（- / Tr / 括弧付き）の出現を確認。
  console.log('データサンプル（食品番号・食品名・対象3カラム）:');
  const idxs = Object.keys(TARGET_IDS).map(id => idRow.findIndex(v => clean(v) === id));
  for (const r of sampleData) {
    if (!r) continue;
    console.log(`  ${r[codeCol]} ${clean(r[nameCol])}: ` +
      idxs.map(i => JSON.stringify(r[i])).join(' / '));
  }
  console.log('='.repeat(60));
}

if (require.main === module) {
  const filePath = process.argv[2] ||
    path.join('raw-data', '20260327-mxt_kagsei-mext-000029402_09.xlsx');
  analyze(filePath);
}

module.exports = { analyze, TARGET_IDS };
