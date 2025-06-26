// taskpane.js – v4  (overwrite + restore)
let originalSnapshot = null;   // stores the very first selection’s raw values

Office.onReady(() => {
  document.getElementById("convert").addEventListener("click", convertCurrency);
  document.getElementById("restore").addEventListener("click", restoreOriginal);
});

async function convertCurrency() {
  try {
    await Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.load("values, address");              // address only for debug
      await context.sync();

      // snapshot original values only once, the first time Convert is pressed
      if (!originalSnapshot) {
        originalSnapshot = sel.values.map(row => row.slice());
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

      // overwrite selected cells with converted numbers
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
  if (!originalSnapshot) {
    console.warn("Nothing to restore – no conversion done yet.");
    return;
  }
  try {
    await Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.values = originalSnapshot;
      await context.sync();
    });
  } catch (err) {
    console.error("Restore error:", err);
  }
}
