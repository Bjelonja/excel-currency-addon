// taskpane.js – v6  (works with or without the Restore button)
let originalSnapshot = null;   // deep copy of first converted range
let originalAddress  = null;   // address of that range (e.g. "B3:D11")

Office.onReady(() => {
  // Convert is mandatory
  document.getElementById("convert")
          .addEventListener("click", convertCurrency);

  // Restore is optional — wire only if the element exists
  const restoreBtn = document.getElementById("restore");
  if (restoreBtn) {
    restoreBtn.addEventListener("click", restoreOriginal);
  }
});

async function convertCurrency() {
  try {
    await Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.load(["values", "address"]);
      await context.sync();

      // Save snapshot only once (first Convert press)
      if (!originalSnapshot) {
        originalSnapshot = sel.values.map(row => row.slice()); // deep copy
        originalAddress  = sel.address;
      }

      const from = fromCurrency.value;
      const to   = toCurrency.value;

      const url  = `https://api.frankfurter.app/latest?from=${from}&to=${to}`;
      const data = await fetch(url).then(r => r.json());
      if (!data.rates || !data.rates[to]) {
        console.error("No rate found:", data);
        return;
      }
      const rate = data.rates[to];

      // Overwrite selected cells
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
      const rng = context.workbook.worksheets
                      .getActiveWorksheet()
                      .getRange(originalAddress);
      rng.values = originalSnapshot;
      rng.select();                 // optional: re-select block
      await context.sync();
    });
  } catch (err) {
    console.error("Restore error:", err);
  }
}
