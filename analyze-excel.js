const XLSX = require('xlsx');

function analyzeExcelStructure(filePath) {
    console.log(`分析中: ${filePath}`);
    console.log('='.repeat(50));
    
    // Excelファイルを読み込み
    const workbook = XLSX.readFile(filePath);
    
    console.log(`シート数: ${workbook.SheetNames.length}`);
    console.log(`シート名: ${workbook.SheetNames.join(', ')}`);
    console.log();
    
    // 各シートの基本情報を表示
    workbook.SheetNames.forEach(sheetName => {
        console.log(`シート: ${sheetName}`);
        console.log('-'.repeat(30));
        
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        console.log(`行数: ${jsonData.length}`);
        
        if (jsonData.length > 0) {
            console.log(`列数: ${jsonData[0].length}`);
            console.log('列名（1行目）:');
            jsonData[0].forEach((col, i) => {
                console.log(`  ${i+1}. ${col}`);
            });
            console.log();
            
            console.log('データサンプル（最初の5行）:');
            jsonData.slice(0, 5).forEach((row, i) => {
                console.log(`行${i+1}:`, row);
            });
        }
        
        console.log('='.repeat(50));
        console.log();
    });
}

if (require.main === module) {
    const filePath = 'raw-data/20230428-mxt_kagsei-mext_00001_012.xlsx';
    analyzeExcelStructure(filePath);
}

module.exports = { analyzeExcelStructure };