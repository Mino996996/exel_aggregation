const XLSX = require('xlsx');
const Aggregate = require('./Aggregate');

const excelData = XLSX.readFile('./assets/20211027_homes_cloudworks.xlsx');
let sheet_0 = excelData.Sheets[excelData.SheetNames[0]];
const cellRange = sheet_0['!ref'];
sheet_0['!ref'] = 'E1:E108';
const sheet_json = XLSX.utils.sheet_to_json(sheet_0);
const key = sheet_0['E1'].v;

const aggregate = new Aggregate(sheet_json, key);
// console.log(aggregate.allCounts());


/* 書き込み設定 */
let resultSheet = excelData.Sheets[excelData.SheetNames[1]]; //2番目のシート

let d = [
  ["1番目","second"],
  [1,"あ","いうえお","さしす"],
];
// 選択したシートws に r=1番目(２行目),c=0番目(A列)から書き込む
XLSX.utils.sheet_add_aoa(resultSheet,d,{origin:{r:1,c:0}});


sheet_0['!ref'] = cellRange; // 書き込み前sheet1のセル範囲を戻す
XLSX.writeFile(excelData, './assets/result.xlsx');
