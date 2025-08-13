const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

class FoodCompositionConverter {
  constructor(excelFilePath) {
    this.excelFilePath = excelFilePath;
    this.workbook = XLSX.readFile(excelFilePath);
    this.outputDir = path.join(__dirname, '../data');
  }

  extractHeaders(worksheet) {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // ヘッダーは2-5行目に分かれているため、結合する
    const headerRow2 = jsonData[1] || [];
    const headerRow3 = jsonData[2] || [];
    const headerRow4 = jsonData[3] || [];
    const headerRow5 = jsonData[4] || [];

    const headers = [];
    for (let i = 0; i < Math.max(headerRow2.length, headerRow3.length, headerRow4.length, headerRow5.length); i++) {
      // より具体的な行から順に取得し、最初に見つかった有効な値を使用
      let bestHeader = '';
      
      // 5行目（最も具体的）から順に確認
      const candidates = [
        headerRow5[i],
        headerRow4[i], 
        headerRow3[i],
        headerRow2[i]
      ];
      
      for (const candidate of candidates) {
        if (candidate && candidate.toString().trim() !== '') {
          const cleaned = candidate.toString().replace(/\r\n|\r|\n/g, '').trim();
          // 意味のある内容があれば採用
          if (cleaned !== '可食部100g当たり' && cleaned !== '可食部100 g当たり' && 
              cleaned !== '無機質' && cleaned !== 'ビタミン' && 
              cleaned.length > 1) {
            bestHeader = cleaned;
            break;
          }
        }
      }
      
      // 有効なヘッダーが見つからない場合は複数行を結合
      if (!bestHeader) {
        const parts = [
          headerRow2[i] || '',
          headerRow3[i] || '',
          headerRow4[i] || '',
          headerRow5[i] || ''
        ].filter(part => part && part.trim() && part !== '');

        if (parts.length > 0) {
          const cleanParts = parts.map(part => 
            part.toString().replace(/\r\n|\r|\n/g, '').trim()
          ).filter((part, index, arr) => arr.indexOf(part) === index);
          
          bestHeader = cleanParts.join('_');
        } else {
          bestHeader = `column_${i}`;
        }
      }
      
      headers[i] = bestHeader;
    }

    return headers;
  }

  cleanHeaders(headers) {
    return headers.map((header, index) => {
      // column_X 形式のヘッダーを適切な名前に変換
      const knownColumns = {
        6: 'エネルギーkcal',
        14: '利用可能炭水化物単糖当量',
        17: '利用可能炭水化物質量計', 
        32: 'ヨウ素',
        38: 'レチノール',
        39: 'αカロテン',
        40: 'βカロテン', 
        41: 'βクリプトキサンチン'
      };
      
      if (header && header.match(/^column_\d+$/)) {
        return knownColumns[index] || `unknown_column_${index}`;
      }
      
      if (!header || header.trim() === '') {
        return knownColumns[index] || `unknown_column_${index}`;
      }
      
      // 特定の列の名前をより適切に変更
      if (index === 3) return '食品名';
      
      return header
        .replace(/\s+/g, '')
        .replace(/[（）()]/g, '')
        .replace(/[・]/g, '')
        .replace(/[　]/g, '')
        .replace(/可食部100g食品名/g, '食品名')
        .replace(/可食部100g/g, '')
        .replace(/_/g, '')
        .toLowerCase();
    });
  }

  convertSheetToJson(sheetName, startRow = 5) {
    const worksheet = this.workbook.Sheets[sheetName];
    if (!worksheet) return null;

    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const headers = this.extractHeaders(worksheet);
    const cleanHeaders = this.cleanHeaders(headers);
    
    const foods = [];
    
    for (let i = startRow; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || row.length === 0) continue;
      
      // 食品番号（2列目）が存在するかチェック - これが実際の食品データの判定基準
      const foodCode = row[1]; // 食品番号
      if (!foodCode || typeof foodCode !== 'string' || !foodCode.match(/^\d{5}$/)) {
        continue; // 5桁の数字でない場合はスキップ（単位行やヘッダー行）
      }
      
      const food = {};
      
      for (let j = 0; j < cleanHeaders.length; j++) {
        const header = cleanHeaders[j];
        if (header) {
          const value = row[j];
          // 数値の場合は数値として保存、文字列の場合はそのまま
          if (typeof value === 'number') {
            food[header] = value;
          } else if (typeof value === 'string' && value.trim() !== '') {
            const trimmed = value.trim();
            
            // "Tr" (トレース), "-" (測定していない), "(0)" (検出されない)などの特殊値を処理
            if (trimmed === 'Tr') {
              food[header] = 'trace';
            } else if (trimmed === '-') {
              food[header] = null;
            } else if (trimmed.match(/^\([0-9.]+\)$/)) {
              food[header] = 'negligible';
            } else if (trimmed.match(/^\d+\.?\d*$/)) {
              // 数値文字列を数値に変換
              food[header] = parseFloat(trimmed);
            } else {
              // その他の文字列（食品名、備考など）はそのまま
              food[header] = trimmed;
            }
          } else {
            food[header] = null;
          }
        }
      }
      
      foods.push(food);
    }
    
    return foods;
  }

  generateCategories() {
    const categories = [];
    
    this.workbook.SheetNames.forEach((sheetName, index) => {
      if (sheetName === '表全体') return; // 全体データはスキップ
      
      const worksheet = this.workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      // 食品群番号を抽出（シート名から）
      const categoryMatch = sheetName.match(/^(\d+)/);
      const categoryId = categoryMatch ? parseInt(categoryMatch[1]) : index;
      
      categories.push({
        id: categoryId,
        name: sheetName.replace(/^\d+/, ''), // 番号を除去
        sheet_name: sheetName,
        food_count: jsonData.length - 5 // ヘッダー行を除く
      });
    });
    
    return categories.sort((a, b) => a.id - b.id);
  }

  generateMetadata() {
    return {
      title: '日本食品標準成分表（八訂）増補2023年',
      version: '8th_edition_supplement_2023',
      source: '文部科学省 科学技術・学術審議会 資源調査分科会',
      url: 'https://www.mext.go.jp/a_menu/syokuhinseibun/mext_00001.html',
      license: '政府標準利用規約2.0 (CC BY 4.0互換)',
      generated_at: new Date().toISOString(),
      data_date: '2023-04-28',
      total_foods: this.workbook.Sheets['表全体'] ? 
        XLSX.utils.sheet_to_json(this.workbook.Sheets['表全体'], { header: 1 }).length - 5 : 0,
      categories: this.workbook.SheetNames.length - 1 // 表全体を除く
    };
  }

  async convert() {
    console.log('Excel→JSON変換を開始します...');
    
    // 出力ディレクトリを作成
    if (!fs.existsSync(this.outputDir)) {
      fs.mkdirSync(this.outputDir, { recursive: true });
    }

    try {
      // 1. 全食品データ（表全体シート）
      console.log('全食品データを変換中...');
      const allFoods = this.convertSheetToJson('表全体');
      if (allFoods) {
        fs.writeFileSync(
          path.join(this.outputDir, 'foods.json'),
          JSON.stringify(allFoods, null, 2),
          'utf8'
        );
        console.log(`✓ foods.json: ${allFoods.length}食品`);
      }

      // 2. カテゴリデータ
      console.log('カテゴリデータを生成中...');
      const categories = this.generateCategories();
      fs.writeFileSync(
        path.join(this.outputDir, 'categories.json'),
        JSON.stringify(categories, null, 2),
        'utf8'
      );
      console.log(`✓ categories.json: ${categories.length}カテゴリ`);

      // 3. メタデータ
      console.log('メタデータを生成中...');
      const metadata = this.generateMetadata();
      fs.writeFileSync(
        path.join(this.outputDir, 'metadata.json'),
        JSON.stringify(metadata, null, 2),
        'utf8'
      );
      console.log(`✓ metadata.json`);

      console.log('\n変換完了！');
      console.log(`出力先: ${this.outputDir}`);

    } catch (error) {
      console.error('変換エラー:', error);
      throw error;
    }
  }
}

// メイン実行
if (require.main === module) {
  const excelFilePath = path.join(__dirname, '../raw-data/20230428-mxt_kagsei-mext_00001_012.xlsx');
  
  if (!fs.existsSync(excelFilePath)) {
    console.error(`Excelファイルが見つかりません: ${excelFilePath}`);
    process.exit(1);
  }

  const converter = new FoodCompositionConverter(excelFilePath);
  converter.convert().catch(error => {
    console.error('変換失敗:', error);
    process.exit(1);
  });
}

module.exports = FoodCompositionConverter;