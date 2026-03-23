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

    // Group all rows by username
    const grouped = new Map(); // username -> [{row, totalVal}]

    for (const { row, uIdx, tIdx } of allRows) {
      const username = String(row[uIdx] || '').trim();
      if (!username) continue;
      const totalVal = Number(row[tIdx]) || 0;
      if (!grouped.has(username)) grouped.set(username, []);
      grouped.get(username).push({ row, totalVal });
    }

    // Build output: for each username decide which rows to keep
    const outputData = [headerRow];
    const flagged = []; // usernames with multiple non-zero entries

    for (const [username, entries] of grouped) {
      const nonZero = entries.filter(e => e.totalVal !== 0);

      if (nonZero.length > 1) {
        // Multiple non-zero rows: keep all of them consecutively and flag
        flagged.push(username);
        for (const e of nonZero) outputData.push(e.row);
      } else if (nonZero.length === 1) {
        // Exactly one non-zero: keep it
        outputData.push(nonZero[0].row);
      } else {
        // All zero: keep first occurrence
        outputData.push(entries[0].row);
      }
    }

    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.aoa_to_sheet(outputData);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Merged');

    const buffer = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

    const totalInput = allRows.length;
    const totalOutput = outputData.length - 1; // exclude header
    const duplicatesRemoved = totalInput - totalOutput;

    res.setHeader('Content-Disposition', 'attachment; filename="merged.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('X-Stats', JSON.stringify({ totalInput, totalOutput, duplicatesRemoved, flagged }));
    res.setHeader('Access-Control-Expose-Headers', 'X-Stats');
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
