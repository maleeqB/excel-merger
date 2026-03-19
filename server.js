const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.post('/merge', upload.array('files'), (req, res) => {
  try {
    if (!req.files || req.files.length < 2) {
      return res.status(400).json({ error: 'Please upload at least 2 Excel files.' });
    }

    // Parse all uploaded files
    const allRows = [];
    let headerRow = null;
    let totalColIndex = -1;
    let usernameColIndex = -1;

    for (const file of req.files) {
      const workbook = XLSX.read(file.buffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      if (data.length === 0) continue;

      const fileHeader = data[0];

      // Find Username and Total columns (case-insensitive, partial match for Total)
      const uIdx = fileHeader.findIndex(h =>
        String(h).trim().toLowerCase() === 'username'
      );
      const tIdx = fileHeader.findIndex(h =>
        String(h).trim().toLowerCase().startsWith('total')
      );

      if (uIdx === -1 || tIdx === -1) {
        return res.status(400).json({
          error: `File "${file.originalname}" is missing a Username or Total column.`
        });
      }

      // Use the first file's header as the canonical header
      if (!headerRow) {
        headerRow = fileHeader;
        usernameColIndex = uIdx;
        totalColIndex = tIdx;
      }

      // Add data rows (skip header)
      for (let i = 1; i < data.length; i++) {
        allRows.push({ row: data[i], uIdx, tIdx });
      }
    }

    // Merge logic: deduplicate by username, prefer row where Total != 0
    const merged = new Map();

    for (const { row, uIdx, tIdx } of allRows) {
      const username = String(row[uIdx] || '').trim();
      if (!username) continue;

      const totalVal = Number(row[tIdx]) || 0;

      if (!merged.has(username)) {
        merged.set(username, { row, totalVal });
      } else {
        const existing = merged.get(username);
        // Keep the one with non-zero total; if both non-zero, keep existing
        if (existing.totalVal === 0 && totalVal !== 0) {
          merged.set(username, { row, totalVal });
        }
      }
    }

    // Build output
    const outputData = [headerRow];
    for (const { row } of merged.values()) {
      outputData.push(row);
    }

    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.aoa_to_sheet(outputData);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Merged');

    const buffer = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

    const totalInput = allRows.length;
    const totalOutput = merged.size;
    const duplicatesRemoved = totalInput - totalOutput;

    res.setHeader('Content-Disposition', 'attachment; filename="merged.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('X-Stats', JSON.stringify({ totalInput, totalOutput, duplicatesRemoved }));
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Something went wrong processing the files.' });
  }
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Excel Merger running at http://localhost:${PORT}`);
});
