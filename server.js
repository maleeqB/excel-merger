const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

// A workbook may declare a range it does not hold: the <dimension> element is
// attacker-controlled metadata, so sheet_to_json has to be bounded at read time.
const MAX_ROWS = 10000;

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.post('/merge', upload.array('files'), (req, res) => {
  try {
    if (!req.files || req.files.length < 2) {
      return res.status(400).json({ error: 'Please upload at least 2 files.' });
    }

    // Parse all uploaded files
    const allRows = [];
    let headerRow = null;
    let totalColIndex = -1;
    let usernameColIndex = -1;

    for (const file of req.files) {
      const isCSV = file.originalname.toLowerCase().endsWith('.csv');
      const workbook = isCSV
        ? XLSX.read(file.buffer.toString('utf8'), { type: 'string', sheetRows: MAX_ROWS + 1 })
        : XLSX.read(file.buffer, { type: 'buffer', sheetRows: MAX_ROWS + 1 });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      // sheetRows makes SheetJS range-check against the rows it actually parsed
      // rather than trusting <dimension>; a workbook longer than the cap is
      // refused here instead of being silently truncated.
      if (sheet && sheet['!fullref'] && sheet['!fullref'] !== sheet['!ref']) {
        return res.status(413).json({
          error: `File "${file.originalname}" has more than ${MAX_ROWS} rows.`
        });
      }
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
    let zeroDropped = 0;

    for (const [username, entries] of grouped) {
      const nonZero = entries.filter(e => e.totalVal !== 0);

      if (nonZero.length > 1) {
        const uniqueScores = new Set(nonZero.map(e => e.totalVal));
        if (uniqueScores.size === 1) {
          // Same score across files — plain duplicate, keep first
          outputData.push(nonZero[0].row);
          zeroDropped += entries.length - 1;
        } else {
          // Genuinely different scores — keep all and flag
          flagged.push(username);
          for (const e of nonZero) outputData.push(e.row);
        }
      } else if (nonZero.length === 1) {
        // Exactly one non-zero: keep it, drop the zero-score duplicates
        outputData.push(nonZero[0].row);
        zeroDropped += entries.length - 1;
      } else {
        // All zero: keep first occurrence
        outputData.push(entries[0].row);
        zeroDropped += entries.length - 1;
      }
    }

    const newSheet = XLSX.utils.aoa_to_sheet(outputData);
    const csvOutput = XLSX.utils.sheet_to_csv(newSheet);

    const totalInput = allRows.length;
    const totalOutput = outputData.length - 1; // exclude header

    res.setHeader('Content-Disposition', 'attachment; filename="merged.csv"');
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('X-Stats', JSON.stringify({ totalInput, totalOutput, zeroDropped, flagged }));
    res.setHeader('Access-Control-Expose-Headers', 'X-Stats');
    res.send(csvOutput);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Something went wrong processing the files.' });
  }
});

app.use((err, req, res, next) => {
  console.error(err);
  const status = err.status || err.statusCode || 500;
  res.status(status).json({ error: err.message || 'Something went wrong.' });
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Excel Merger running at http://localhost:${PORT}`);
});
