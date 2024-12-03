import fs from 'fs';
import xlsx from 'xlsx';

// Write Into Excel
const workbook = xlsx.readFile('Apply History.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

const newRow = {
    A: '',
    B: '',
    C: '',
};

xlsx.utils.sheet_add_aoa(sheet, [Object.values(newRow)], { origin: 0 });

xlsx.writeFile(workbook, 'Apply History.xlsx');
console.log(sheet);


// change JSON file
// const filePath = './sequence.json'
// const sequenceFile = JSON.parse(fs.readFileSync(filePath, 'utf8'));

// let sequence = parseInt(sequenceFile.sequence) + 1
// const data = JSON.stringify({sequence})
// fs.writeFileSync(filePath, data);
