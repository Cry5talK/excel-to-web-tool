// Add Excel export library (SheetJS)
document.write(`
    <script src="https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js"><\/script>
  `);
function calculate() {
    // 1. Get input values
    const hours = parseFloat(document.getElementById('hours').value) || 0;
    const rate = parseFloat(document.getElementById('rate').value) || 0;
    const expenses = parseFloat(document.getElementById('expenses').value) || 0;
  
    // 2. Calculate (like Excel formulas)
    const subtotal = hours * rate;
    const total = subtotal + expenses;
  
    // 3. Display results
    document.getElementById('result').innerHTML = `
      <h3>Quote Summary:</h3>
      <p>Subtotal (Hours*Rate): $${subtotal.toFixed(2)}</p>
      <p>Expenses: $${expenses.toFixed(2)}</p>
      <p><strong>Total: $${total.toFixed(2)}</strong></p>
    `;
  }
  
  function exportToExcel() {
    // Get values
    const hours = parseFloat(document.getElementById('hours').value) || 0;
    const rate = parseFloat(document.getElementById('rate').value) || 0;
    const expenses = parseFloat(document.getElementById('expenses').value) || 0;
    const subtotal = hours * rate;
    const total = subtotal + expenses;
  
    // Create Excel workbook
    const workbook = XLSX.utils.book_new();
    
    // Data for Excel sheet
    const excelData = [
      ["Item", "Value"],
      ["Hours", hours],
      ["Rate", rate],
      ["Subtotal", subtotal],
      ["Expenses", expenses],
      ["TOTAL", total]
    ];
    
    // Create worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    
    // Add to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, "Quote");
    
    // Export
    XLSX.writeFile(workbook, "Freelance_Quote.xlsx");
  }