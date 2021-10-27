const XLSX = require('xlsx');

const excelData = XLSX.readFile('./assets/data_sample.xlsx');

let sheet = excelData.Sheets[excelData.SheetNames[0]];

// console.log(sheet['!ref']);

sheet['!ref'] = 'E1:E108';

const sheet_json = XLSX.utils.sheet_to_json(sheet);

console.log(sheet_json);

// let count = 0;
// for(let sell of sheet_json) {
//   console.log(`${count} :${sell}`);
//   count +=1;
// }
