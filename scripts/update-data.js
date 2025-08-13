#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

class DataUpdater {
  constructor() {
    this.rootDir = path.join(__dirname, '..');
    this.rawDataDir = path.join(this.rootDir, 'raw-data');
    this.dataDir = path.join(this.rootDir, 'data');
    this.scriptsDir = path.join(this.rootDir, 'scripts');
  }

  log(message) {
    console.log(`[${new Date().toISOString()}] ${message}`);
  }

  error(message) {
    console.error(`[ERROR] ${message}`);
  }

  // 1. 新しいExcelファイルの検証
  validateExcelFile(filename) {
    const filepath = path.join(this.rawDataDir, filename);
    
    if (!fs.existsSync(filepath)) {
      throw new Error(`Excelファイルが見つかりません: ${filepath}`);
    }

    // ファイルサイズチェック（最低1MB以上）
    const stats = fs.statSync(filepath);
    if (stats.size < 1024 * 1024) {
      throw new Error(`ファイルサイズが小さすぎます: ${stats.size} bytes`);
    }

    this.log(`✓ Excelファイル検証完了: ${filename} (${(stats.size / 1024 / 1024).toFixed(1)}MB)`);
    return filepath;
  }

  // 2. 変換スクリプトのファイルパス自動更新
  updateConverterScript(filename) {
    const converterPath = path.join(this.scriptsDir, 'convert-excel.js');
    let content = fs.readFileSync(converterPath, 'utf8');

    // 既存のファイルパスを見つけて置換
    const oldPattern = /const excelFilePath = path\.join\(__dirname, '\.\.\/raw-data\/[^']+\.xlsx'\);/;
    const newPath = `const excelFilePath = path.join(__dirname, '../raw-data/${filename}');`;
    
    if (!oldPattern.test(content)) {
      throw new Error('convert-excel.jsのファイルパス設定が見つかりません');
    }

    content = content.replace(oldPattern, newPath);
    fs.writeFileSync(converterPath, content, 'utf8');
    
    this.log(`✓ 変換スクリプト更新完了: ${filename}`);
  }

  // 3. データ構造の分析
  analyzeStructure() {
    this.log('データ構造を分析中...');
    try {
      const output = execSync('yarn analyze', { 
        cwd: this.rootDir, 
        encoding: 'utf8',
        stdio: 'pipe'
      });
      
      // シート数と列数を抽出
      const sheetMatches = output.match(/シート数: (\d+)/);
      const sheetCount = sheetMatches ? parseInt(sheetMatches[1]) : 0;
      
      this.log(`✓ データ構造分析完了: ${sheetCount}シート検出`);
      
      if (sheetCount !== 19) {
        this.log(`⚠️  警告: シート数が変更されています (期待値: 19, 実際: ${sheetCount})`);
      }
      
      return { sheetCount };
    } catch (error) {
      throw new Error(`データ構造分析に失敗: ${error.message}`);
    }
  }

  // 4. データ変換実行
  convertData() {
    this.log('Excel→JSON変換を実行中...');
    try {
      const output = execSync('yarn convert', { 
        cwd: this.rootDir, 
        encoding: 'utf8',
        stdio: 'pipe'
      });
      
      this.log('✓ データ変換完了');
      return output;
    } catch (error) {
      throw new Error(`データ変換に失敗: ${error.message}`);
    }
  }

  // 5. データ検証
  validateData() {
    this.log('データ品質をチェック中...');
    
    const foodsPath = path.join(this.dataDir, 'foods.json');
    const categoriesPath = path.join(this.dataDir, 'categories.json');
    const metadataPath = path.join(this.dataDir, 'metadata.json');

    // ファイル存在確認
    if (!fs.existsSync(foodsPath)) {
      throw new Error('foods.json が生成されていません');
    }
    if (!fs.existsSync(categoriesPath)) {
      throw new Error('categories.json が生成されていません');
    }
    if (!fs.existsSync(metadataPath)) {
      throw new Error('metadata.json が生成されていません');
    }

    // データ読み込み
    const foods = JSON.parse(fs.readFileSync(foodsPath, 'utf8'));
    const categories = JSON.parse(fs.readFileSync(categoriesPath, 'utf8'));
    const metadata = JSON.parse(fs.readFileSync(metadataPath, 'utf8'));

    // 基本検証
    if (!Array.isArray(foods) || foods.length === 0) {
      throw new Error('foods.json にデータがありません');
    }
    
    if (!Array.isArray(categories) || categories.length === 0) {
      throw new Error('categories.json にデータがありません');
    }

    // 必須フィールドチェック
    const requiredFields = ['食品群', '食品番号', '食品名'];
    const missingFields = foods.filter(food => 
      requiredFields.some(field => !food.hasOwnProperty(field) || food[field] === null)
    );

    if (missingFields.length > 0) {
      this.log(`⚠️  警告: ${missingFields.length}食品で必須フィールドが不足しています`);
    }

    // 統計情報
    const stats = {
      foods: foods.length,
      categories: categories.length,
      minFoodCount: Math.min(...categories.map(c => c.food_count)),
      maxFoodCount: Math.max(...categories.map(c => c.food_count))
    };

    this.log(`✓ データ検証完了:`);
    this.log(`  - 食品数: ${stats.foods}`);
    this.log(`  - カテゴリ数: ${stats.categories}`);
    this.log(`  - カテゴリ別食品数: ${stats.minFoodCount} - ${stats.maxFoodCount}`);
    
    if (stats.foods < 2500) {
      this.log(`⚠️  警告: 食品数が少ないです (${stats.foods} < 2500)`);
    }

    return stats;
  }

  // 6. メタデータの更新
  updateMetadata(stats, filename) {
    const metadataPath = path.join(this.dataDir, 'metadata.json');
    const metadata = JSON.parse(fs.readFileSync(metadataPath, 'utf8'));
    
    // バージョン推測（ファイル名から年度を抽出）
    const yearMatch = filename.match(/(\d{4})/);
    const year = yearMatch ? yearMatch[1] : new Date().getFullYear();
    
    // メタデータ更新
    metadata.version = `8th_edition_supplement_${year}`;
    metadata.total_foods = stats.foods;
    metadata.source_file = filename;
    metadata.updated_at = new Date().toISOString();
    
    fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2), 'utf8');
    this.log(`✓ メタデータ更新完了: バージョン ${metadata.version}`);
  }

  // 7. 配信用ファイル生成
  buildDistribution() {
    this.log('配信用ファイルを生成中...');
    try {
      execSync('yarn build', { 
        cwd: this.rootDir, 
        stdio: 'pipe'
      });
      this.log('✓ 配信用ファイル生成完了');
    } catch (error) {
      throw new Error(`配信用ファイル生成に失敗: ${error.message}`);
    }
  }

  // 8. 変更レポート生成
  generateReport(stats, oldStats = null) {
    const reportPath = path.join(this.rootDir, 'update-report.md');
    const timestamp = new Date().toISOString().split('T')[0];
    
    let report = `# データ更新レポート - ${timestamp}\n\n`;
    report += `## 更新結果\n\n`;
    report += `- 食品数: ${stats.foods}\n`;
    report += `- カテゴリ数: ${stats.categories}\n`;
    
    if (oldStats) {
      const foodDiff = stats.foods - oldStats.foods;
      report += `- 食品数変化: ${foodDiff > 0 ? '+' : ''}${foodDiff}\n`;
    }
    
    report += `\n## 実行ログ\n\n`;
    report += `- 実行日時: ${new Date().toISOString()}\n`;
    report += `- Node.js バージョン: ${process.version}\n`;
    
    fs.writeFileSync(reportPath, report, 'utf8');
    this.log(`✓ 更新レポート生成: ${reportPath}`);
  }

  // メイン処理
  async update(filename) {
    const startTime = Date.now();
    this.log(`=== データ更新開始: ${filename} ===`);

    try {
      // 1. ファイル検証
      this.validateExcelFile(filename);

      // 2. スクリプト更新
      this.updateConverterScript(filename);

      // 3. 構造分析
      const structure = this.analyzeStructure();

      // 4. データ変換
      this.convertData();

      // 5. データ検証
      const stats = this.validateData();

      // 6. メタデータ更新
      this.updateMetadata(stats, filename);

      // 7. 配信用ファイル生成
      this.buildDistribution();

      // 8. レポート生成
      this.generateReport(stats);

      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      this.log(`=== データ更新完了 (${duration}秒) ===`);
      
      console.log('\n🎉 データ更新が正常に完了しました！');
      console.log('\n次の手順:');
      console.log('1. update-report.md で結果を確認');
      console.log('2. 必要に応じてREADME.mdの食品数を更新');
      console.log('3. git add . && git commit で変更をコミット');

    } catch (error) {
      this.error(`データ更新に失敗: ${error.message}`);
      process.exit(1);
    }
  }
}

// CLI実行
if (require.main === module) {
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    console.log('使用方法: node scripts/update-data.js <Excelファイル名>');
    console.log('例: node scripts/update-data.js 20240428-mxt_kagsei-mext_00001_013.xlsx');
    process.exit(1);
  }

  const updater = new DataUpdater();
  updater.update(args[0]);
}

module.exports = DataUpdater;