const XLSX = require('xlsx');
const wb = XLSX.readFile('./坪山区青春小店介绍.xlsx');
const sheetName = wb.SheetNames[0];
const ws = wb.Sheets[sheetName];

console.log('Sheet Name:', sheetName);
console.log('Range:', ws['!ref']);

const range = XLSX.utils.decode_range(ws['!ref']);
console.log('\nFirst 10 rows:');
for(let r=0; r<=Math.min(10, range.e.r); r++) {
    let row = [];
    for(let c=0; c<=range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({r,c})];
        row.push(cell ? cell.v : '');
    }
    console.log(`${r}: ${row.join(' | ')}`);
}