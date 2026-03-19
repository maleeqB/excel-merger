const XLSX = require('xlsx');

function makeSheet(data) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return wb;
}

// File 1: has some students, some with Total=0 (incomplete)
const file1 = [
  ['S/N', 'Username', 'First Name', 'Last Name', 'Start time', 'End time', 'IP Address', 'Submitter', 'GES_107_CA (40)', 'GES_107_EXAM (60)', 'Total (100)'],
  [1, 'STU/001/2021', 'Alice', 'Johnson', '2024-01-10 09:00', '2024-01-10 10:00', '192.168.1.1', 'System', 32, 45, 77],
  [2, 'STU/002/2021', 'Bob', 'Smith', '2024-01-10 09:05', '2024-01-10 10:05', '192.168.1.2', 'System', 0, 0, 0],   // duplicate, total=0
  [3, 'STU/003/2021', 'Carol', 'White', '2024-01-10 09:10', '2024-01-10 10:10', '192.168.1.3', 'System', 28, 50, 78],
  [4, 'STU/004/2021', 'David', 'Brown', '2024-01-10 09:15', '2024-01-10 10:15', '192.168.1.4', 'System', 35, 55, 90],
  [5, 'STU/005/2021', 'Eve', 'Davis', '2024-01-10 09:20', '2024-01-10 10:20', '192.168.1.5', 'System', 0, 0, 0],   // duplicate, total=0
];

// File 2: overlapping usernames (with leading/trailing spaces to test trim), different totals
const file2 = [
  ['S/N', 'Username', 'First Name', 'Last Name', 'Start time', 'End time', 'IP Address', 'Submitter', 'GES_107_CA (40)', 'GES_107_EXAM (60)', 'Total (100)'],
  [1, ' STU/002/2021 ', 'Bob', 'Smith', '2024-01-11 09:00', '2024-01-11 10:00', '10.0.0.1', 'System', 30, 48, 78],  // same as file1 row2 but has score
  [2, 'STU/005/2021',   'Eve', 'Davis', '2024-01-11 09:05', '2024-01-11 10:05', '10.0.0.2', 'System', 22, 40, 62],  // same as file1 row5 but has score
  [3, 'STU/006/2021',   'Frank', 'Miller', '2024-01-11 09:10', '2024-01-11 10:10', '10.0.0.3', 'System', 38, 52, 90], // new student
  [4, 'STU/007/2021',   'Grace', 'Wilson', '2024-01-11 09:15', '2024-01-11 10:15', '10.0.0.4', 'System', 40, 58, 98], // new student
];

XLSX.writeFile(makeSheet(file1), 'test_file1.xlsx');
XLSX.writeFile(makeSheet(file2), 'test_file2.xlsx');

console.log('Created test_file1.xlsx and test_file2.xlsx');
console.log('\nFile 1 has 5 students (STU/001 to STU/005), two with Total=0');
console.log('File 2 has 4 students — STU/002 & STU/005 overlap but with real scores');
console.log('\nExpected merged result: 7 unique students, no zeroed-out rows');
