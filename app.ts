const XLSX = require('xlsx');
const Aggregate = require('./Aggregate');

const excelData = XLSX.readFile('./assets/20211027_homes_cloudworks.xlsx');
let sheet_0 = excelData.Sheets[excelData.SheetNames[0]];
const baseCellRange = sheet_0['!ref'];
sheet_0['!ref'] = 'E1:E108';
const sheet_json = XLSX.utils.sheet_to_json(sheet_0);
const key = sheet_0['E1'].v;

const aggregate_E1 = new Aggregate(sheet_json, key);
// console.log(aggregate.allCounts());


/* 書き込み設定 */
let resultSheet = excelData.Sheets[excelData.SheetNames[1]]; //2番目のシート
XLSX.utils.sheet_add_aoa(resultSheet,aggregate_E1.allCounts(),{origin:{r:1,c:0}}); //E1集計はA2:Bxの範囲


sheet_0['!ref'] = baseCellRange; // 書き込み前sheet1のセル範囲を戻す
XLSX.writeFile(excelData, './assets/result.xlsx');
