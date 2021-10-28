const XLSX = require('xlsx');
const Aggregate = require('./Aggregate');

const excelData = XLSX.readFile('./assets/data_sample.xlsx');

let sheet = excelData.Sheets[excelData.SheetNames[0]];
sheet['!ref'] = 'E1:E108';
const sheet_json = XLSX.utils.sheet_to_json(sheet);
const key = sheet['E1'].v;

const aggregate = new Aggregate(sheet_json, key);

console.log(aggregate.allCounts());

