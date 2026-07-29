// R-40: auto-add a branded .xlsx whenever the research contains tabular data — one sheet
// per table block, grouped under the nearest preceding h1/h2 heading for the sheet name.
const ExcelJS = require('exceljs');

function hasTableBlock(blocks) {
  return (blocks || []).some((b) => b && String(b.type).toLowerCase() === 'table');
}

async function xlsxBuffer(job, brandKit) {
  const blocks = (job.input && job.input.blocks) || [];
  if (!hasTableBlock(blocks)) return null;

  const accent = ((brandKit && brandKit.accent_color) || '4393A5').replace('#', '').toUpperCase();

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Intelligence Machine';
  wb.created = new Date();

  let sectionTitle = '';
  let sheetIndex = 0;
  for (const b of blocks) {
    const t = String(b.type || '').toLowerCase();
    if (t === 'h1' || t === 'h2') { sectionTitle = b.text || sectionTitle; continue; }
    if (t !== 'table') continue;
    sheetIndex += 1;
    const name = (sectionTitle || `Table ${sheetIndex}`).replace(/[\\/*?:[\]]/g, ' ').slice(0, 31) || `Sheet${sheetIndex}`;
    const sheet = wb.addWorksheet(name);
    const headers = Array.isArray(b.headers) ? b.headers.map(String) : [];
    const rows = Array.isArray(b.rows) ? b.rows : [];
    if (headers.length) {
      const headerRow = sheet.addRow(headers);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + accent } };
      });
      sheet.columns = headers.map(() => ({ width: 22 }));
      sheet.views = [{ state: 'frozen', ySplit: 1 }];
    }
    rows.forEach((r) => sheet.addRow((r || []).map(String)));
  }

  if (!wb.worksheets.length) return null;
  return Buffer.from(await wb.xlsx.writeBuffer());
}

module.exports = { xlsxBuffer, hasTableBlock };
