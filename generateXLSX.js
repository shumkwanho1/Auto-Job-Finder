import XLSX from 'xlsx';
import fs from 'fs';

fs.unlinkSync('Apply History.xlsx');

const data = JSON.stringify({ sequence: 0 })
fs.writeFileSync('./sequence.json', data);

const workbook = XLSX.utils.book_new();
const worksheet = []
XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
XLSX.writeFile(workbook, "Apply History.xlsx");