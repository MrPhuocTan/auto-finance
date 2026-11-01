import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from '../vendor/xlsx.mjs';
import { getSampleDataset } from '../src/core/sampleData.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');
const outputFilePath = path.join(projectRoot, 'public', 'mau-cong-no-ke-toan.xlsx');

function buildSampleWorkbook() {
  const { sheetName, rawGrid } = getSampleDataset();
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(rawGrid);

  // Set clean column widths suitable for Vietnamese accounting ledgers
  worksheet['!cols'] = [
    { wch: 15 }, // Ngày chứng từ
    { wch: 14 }, // Số chứng từ
    { wch: 35 }, // Khách hàng / Đối tác
    { wch: 62 }, // Diễn giải nội dung
    { wch: 18 }, // Số tiền (VNĐ)
    { wch: 22 }, // Ghi chú
  ];

  // Merge top company/report title rows (Rows 1..3 -> indices 0..2) across columns A..F
  worksheet['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 5 } },
  ];

  // Apply standard accounting integer format (#,##0) to all numeric amount cells in Column E (index 4)
  for (let r = 5; r < rawGrid.length; r++) {
    const cellRef = XLSX.utils.encode_cell({ r, c: 4 });
    const cell = worksheet[cellRef];
    if (cell && typeof cell.v === 'number') {
      cell.z = '#,##0';
    }
  }

  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  return workbook;
}

function main() {
  const workbook = buildSampleWorkbook();
  fs.mkdirSync(path.dirname(outputFilePath), { recursive: true });

  const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
  fs.writeFileSync(outputFilePath, buffer);

  const stats = fs.statSync(outputFilePath);
  console.log(
    `[generateSampleExcel] Created ${path.relative(projectRoot, outputFilePath)} (${stats.size} bytes)`
  );
}

main();
