const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 脂肪酸成分表編（八訂・増補2023）第1表 → JSON 変換と本表への join。
//
// 本表の変換（convert-excel.js）で生成済みの data/foods.json を入力に取り、
// 食品番号で脂肪酸成分表編を join して各レコードに次の3カラムを追加する：
//   - 飽和脂肪酸
//   - n-3系多価不飽和脂肪酸
//   - n-6系多価不飽和脂肪酸
// 脂肪酸成分表編に収載のない食品は3カラムとも null（= 未測定）。
//
// 特殊値規約は本表（convert-excel.js）と一部を意図的に変える。
// 本表は括弧付き推計値を 'negligible' に潰すが、脂肪酸は推計値率が高く実数を残す価値が
// 大きいため、栄養計算の精度を優先して推計値の実数を保持する：
//   'Tr' / '(Tr)' → 'trace'（推計値のトレースも trace 扱い）
//   '-' → null（未測定）
//   '(0.22)' などの括弧付き数値 → 0.22（推計値の実数を保持。本表の 'negligible' とは異なる）
//   '0.22' → 0.22 / 数値 → 数値

// 取り込み対象。脂肪酸成分表編の成分識別子（ヘッダ行4）→ foods.json でのキー名。
const TARGET = {
  FASAT: '飽和脂肪酸',
  FAPUN3: 'n-3系多価不飽和脂肪酸',
  FAPUN6: 'n-6系多価不飽和脂肪酸',
};

const FATTY_ACID_KEYS = Object.values(TARGET);

function clean(v) {
  return v == null ? '' : v.toString().replace(/\r\n|\r|\n/g, '').trim();
}

// 脂肪酸カラム用の特殊値正規化。推計値（括弧付き）の実数を保持する点が本表と異なる。
function normalize(value) {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  if (trimmed === '') return null;
  if (trimmed === 'Tr' || trimmed === '(Tr)') return 'trace';
  if (trimmed === '-') return null;
  // 括弧付き推計値は中の数値を実数として保持（本表は 'negligible' に潰すが、脂肪酸では残す）。
  const paren = trimmed.match(/^\(([0-9.]+)\)$/);
  if (paren) return parseFloat(paren[1]);
  if (trimmed.match(/^\d+\.?\d*$/)) return parseFloat(trimmed);
  return trimmed;
}

// 脂肪酸成分表編 Excel から { 食品番号(int) → {飽和脂肪酸, n-3, n-6} } を作る。
function buildFattyAcidMap(excelFilePath) {
  const workbook = XLSX.readFile(excelFilePath);
  const sheetName = workbook.SheetNames.find(s => s.trim() === '表全体') || workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

  // ヘッダ階層：行3=日本語名 / 行4=成分識別子。データは行6以降。
  const headerRow = rows[3] || [];
  const idRow = rows[4] || [];

  const codeCol = headerRow.findIndex(h => clean(h) === '食品番号');
  const colOf = {};
  for (const id of Object.keys(TARGET)) {
    const idx = idRow.findIndex(v => clean(v) === id);
    if (idx === -1) throw new Error(`脂肪酸成分表編に成分識別子 ${id} が見つかりません`);
    colOf[id] = idx;
  }

  const map = new Map();
  for (let i = 5; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const code = row[codeCol];
    // 食品番号は5桁の数字文字列。これ以外（単位行・空行）はスキップ。
    if (!code || !code.toString().match(/^\d{5}$/)) continue;

    const foodCode = parseInt(code, 10); // foods.json の 食品番号（数値）と揃える
    const entry = {};
    for (const [id, key] of Object.entries(TARGET)) {
      entry[key] = normalize(row[colOf[id]]);
    }
    map.set(foodCode, entry);
  }
  return map;
}

function join(foodsPath, excelFilePath) {
  const foods = JSON.parse(fs.readFileSync(foodsPath, 'utf8'));
  const faMap = buildFattyAcidMap(excelFilePath);

  let matched = 0;
  const unmatchedFoodCodes = []; // 本表にあって脂肪酸編に無い食品

  for (const food of foods) {
    const entry = faMap.get(food.食品番号);
    if (entry) {
      Object.assign(food, entry);
      matched++;
    } else {
      // 収載外は3カラムとも null（未測定）。
      for (const key of FATTY_ACID_KEYS) food[key] = null;
      unmatchedFoodCodes.push(food.食品番号);
    }
  }

  // 脂肪酸編にあって本表に無い食品番号（収載差分の逆方向。通常は無いはず）。
  const foodCodeSet = new Set(foods.map(f => f.食品番号));
  const faOnly = [...faMap.keys()].filter(code => !foodCodeSet.has(code));

  return { foods, stats: { total: foods.length, matched, unmatched: unmatchedFoodCodes.length, faOnly } };
}

// metadata.json に脂肪酸データの出典・収載食品数・null 規約を追記する。
// 本表の `yarn convert` が metadata.json を再生成するため、join 後にこの後段で上書きする。
function updateMetadata(metadataPath, stats) {
  const metadata = JSON.parse(fs.readFileSync(metadataPath, 'utf8'));
  metadata.fatty_acid = {
    title: '日本食品標準成分表（八訂）増補2023年 脂肪酸成分表編',
    source: '文部科学省 科学技術・学術審議会 資源調査分科会',
    url: 'https://www.mext.go.jp/a_menu/syokuhinseibun/mext_00001.html',
    table: '第1表 可食部100g当たりの脂肪酸成分表（本表）',
    columns: FATTY_ACID_KEYS,
    unit: 'g/100g',
    foods_with_fatty_acid: stats.matched, // 脂肪酸データ収載食品数
    total_foods: stats.total, // 本表食品数
    foods_without_fatty_acid: stats.unmatched, // 収載外（脂肪酸3カラムが null）
    null_convention: '本表に収載され脂肪酸成分表編に収載のない食品は、脂肪酸3カラムが null（未測定）。',
    value_convention: '括弧付き推計値は実数として保持（本表の negligible 正規化とは異なる）。Tr / (Tr) は trace、- は null。',
  };
  fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2), 'utf8');
}

if (require.main === module) {
  const foodsPath = path.join(__dirname, '../data/foods.json');
  const excelFilePath = process.argv[2] ||
    path.join(__dirname, '../raw-data/20260327-mxt_kagsei-mext-000029402_09.xlsx');

  if (!fs.existsSync(foodsPath)) {
    console.error(`foods.json が見つかりません: ${foodsPath}\n先に \`yarn convert\` を実行してください。`);
    process.exit(1);
  }
  if (!fs.existsSync(excelFilePath)) {
    console.error(`脂肪酸成分表編 Excel が見つかりません: ${excelFilePath}`);
    process.exit(1);
  }

  console.log('脂肪酸成分表編の join を開始します...');
  const { foods, stats } = join(foodsPath, excelFilePath);

  fs.writeFileSync(foodsPath, JSON.stringify(foods, null, 2), 'utf8');

  const metadataPath = path.join(__dirname, '../data/metadata.json');
  if (fs.existsSync(metadataPath)) {
    updateMetadata(metadataPath, stats);
    console.log('✓ metadata.json 更新（fatty_acid セクション）');
  }

  console.log(`✓ foods.json 更新: ${stats.total}食品`);
  console.log(`  脂肪酸データ収載: ${stats.matched} / 収載外(null): ${stats.unmatched}`);
  if (stats.faOnly.length > 0) {
    console.log(`  ⚠ 脂肪酸編にのみ存在する食品番号: ${stats.faOnly.length}件 (${stats.faOnly.slice(0, 10).join(', ')}...)`);
  }
}

module.exports = { join, buildFattyAcidMap, updateMetadata, normalize, TARGET, FATTY_ACID_KEYS };
