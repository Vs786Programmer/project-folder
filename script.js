// Container for Handsontable
const container = document.getElementById('spreadsheet');
const saveButton = document.getElementById('saveButton');

// Initial data (example CSV data)
const initialData = [
  ['Name', 'Age', 'Grade'],
  ['Alice', '14', 'A'],
  ['Bob', '15', 'B'],
  ['Charlie', '13', 'A'],
];

let hot;

// Initialize Handsontable
hot = new Handsontable(container, {
  data: initialData,
  rowHeaders: true,
  colHeaders: true,
  licenseKey: 'non-commercial-and-evaluation',
  stretchH: 'all', // Makes columns stretch to fit the container
  manualColumnResize: true, // Allows resizing columns
  manualRowResize: true, // Allows resizing rows
  contextMenu: true, // Enables right-click menu
  height: 400, // Sets table height
});

// Export functionality
saveButton.addEventListener('click', () => {
  const updatedData = hot.getData(); // Get the edited data
  const ws = XLSX.utils.aoa_to_sheet(updatedData); // Convert array to worksheet
  const wb = XLSX.utils.book_new(); // Create a new workbook
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1'); // Append worksheet
  XLSX.writeFile(wb, 'Edited_File.xlsx'); // Export file
});
