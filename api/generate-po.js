import ExcelJS from 'exceljs';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  const { templateBase64, sheetName, headerRowIdx, dataStartRowIdx, rows } = req.body || {};
  if (
    !templateBase64 ||
    sheetName === undefined ||
    headerRowIdx === undefined ||
    dataStartRowIdx === undefined ||
    !Array.isArray(rows)
  ) {
    res.status(400).json({ error: '필수 값이 누락되었습니다.' });
    return;
  }

  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(templateBase64, 'base64'));
    const ws = wb.getWorksheet(sheetName) || wb.worksheets[0];

    // 헤더 행에 옅은 배경색 + 볼드체를 넣어 아래 데이터와 구분되게 함
    const headerRowNum = headerRowIdx + 1; // ExcelJS 행/열 번호는 1부터 시작
    const headerRow = ws.getRow(headerRowNum);
    const usedCols = Math.max(ws.columnCount || 0, headerRow.cellCount || 0, 1);
    for (let c = 1; c <= usedCols; c++) {
      const cell = headerRow.getCell(c);
      if (cell.value === null || cell.value === undefined || String(cell.value).trim() === '') continue;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3EAF5' } };
      cell.font = { ...(cell.font || {}), bold: true };
    }
    headerRow.commit();

    // 발주 데이터 채우기
    rows.forEach((rowValues, i) => {
      const rowNum = dataStartRowIdx + 1 + i; // 1-based
      const row = ws.getRow(rowNum);
      Object.entries(rowValues).forEach(([colIndexStr, val]) => {
        const colNum = Number(colIndexStr) + 1; // 1-based
        const cell = row.getCell(colNum);
        const trimmed = String(val).trim();
        const num = Number(trimmed);
        const isNum = trimmed !== '' && !isNaN(num) && /^-?\d+(\.\d+)?$/.test(trimmed);
        cell.value = isNum ? num : val;
      });
      row.commit();
    });

    const buf = await wb.xlsx.writeBuffer();
    res.status(200).json({ base64: Buffer.from(buf).toString('base64') });
  } catch (err) {
    console.error('generate-po error:', err);
    res.status(500).json({ error: err.message || '발주서 생성 중 오류가 발생했습니다.' });
  }
}
