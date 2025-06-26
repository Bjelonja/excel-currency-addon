// taskpane.js – v5  (overwrite + full-range restore)
let originalSnapshot = null;   // deep copy of the first converted range
let originalAddress  = null;   // address of that range (e.g. "B3:D11")

Office.onReady(() => {
  document.getElementById("convert").addEventListener("click", convertCurrency);
  document.getElementById("restore").addEventListener("click", restoreOriginal);
});

async function convertCurrency() {
  try {
    await Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.load("values, address");
      await context.sync();

      // Save snapshot only once (first Convert press)
      if (!originalSnapshot) {
        originalSnapshot = sel.values.map(row => row.slice()); // deep copy
        originalAddress  = sel.address;
      }

      const from = document.getElementById("fromCurrency").value;
      const to   = document.getElementById("toCurrency").value;

      const url  = `https://api.frankfurter.app/latest?from=${from}&to=${to}`;
      const data = await fetch(url).then(r => r.json());
      if (!data.rates || !data.rates[to]) {
        console.error("No rate found in response:", data);
        return;
      }
      const rate = data.rates[to];

      // Overwrite selected cells with converted numbers
      sel.values = sel.values.map(row =>
        row.map(cell => {
          const n = parseFloat(cell);
          return isNaN(n) ? cell : (n * rate).toFixed(2);
        })
      );

      await context.sync();
    });
  } catch (err) {
    console.error("Currency add-in error:", err);
  }
}

async function restoreOriginal() {
  if (!originalSnapshot || !originalAddress) {
    console.warn("Nothing to restore – no conversion done yet.");
    return;
  }
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rng   = sheet.getRange(originalAddress);
      rng.values  = originalSnapshot;   // restore full block
      rng.select();                     // optional: re-select it
      await context.sync();
    });
  } catch (err) {
    console.error("Restore error:", err);
  }
}
