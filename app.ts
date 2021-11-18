const XLSX = require('xlsx');
const Aggregate = require('./Aggregate');
const excelData = XLSX.readFile('./assets/20211027_homes_cloudworks.xlsx');

const resultSheet = excelData.Sheets[excelData.SheetNames[1]]; //2番目のシートを結果用シートとする

let sheet_0 = excelData.Sheets[excelData.SheetNames[0]]; // 生データ（1番目シート）の読み込み
const baseCellRange = sheet_0['!ref']; // 生データのデータ初期範囲設定を保持しておく

// 各集計エリアと要素キーワード位置を設定する
const readSheetRange = ['E1:E108', 'G1:G108', 'I1:I108', 'K1:K108', 'M1:M108', 'O1:O108', 'Q1:Q108'];
const readKeyWord = ['E1', 'G1', 'I1', 'K1', 'M1', 'O1', 'Q1'];

for (let i = 0; i < readSheetRange.length; i++) {
  // データの集計
  sheet_0['!ref'] = readSheetRange[i];
  const sheet_json = XLSX.utils.sheet_to_json(sheet_0);
  const key = sheet_0[readKeyWord[i]].v;
  const aggregate = new Aggregate(sheet_json, key);
  console.log(aggregate.allCounts());

  // インメモリ内に集計結果の書き込み準備
  XLSX.utils.sheet_add_aoa(resultSheet, aggregate.allCounts(), { origin: { r: 1 , c: 3 * i }}); // ex)E1集計はA2:Bxの範囲
}

// sheet_0['!ref'] = 'E1:E108';
// const sheet_json = XLSX.utils.sheet_to_json(sheet_0);
// const key = sheet_0['E1'].v;

// const aggregate_E1 = new Aggregate(sheet_json, key);
// console.log(aggregate_E1.allCounts());


// /* 書き込み設定 */
// let resultSheet = excelData.Sheets[excelData.SheetNames[1]]; //2番目のシート
// XLSX.utils.sheet_add_aoa(resultSheet,aggregate_E1.allCounts(),{origin:{r:1,c:0}}); //E1集計はA2:Bxの範囲


sheet_0['!ref'] = baseCellRange; // 書き込み前sheet1のセル範囲を戻す
XLSX.writeFile(excelData, './assets/result.xlsx');

console.log('集計完了');
