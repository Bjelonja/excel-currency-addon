// taskpane.js  –  v3  (switch to frankfurter.app + robust handling)
Office.onReady(() => {
  document.getElementById("convert").addEventListener("click", convertCurrency);
});

async function convertCurrency() {
  try {
    await Excel.run(async (context) => {
      // 1. selected range
      const sel = context.workbook.getSelectedRange();
      sel.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      // 2. currencies
      const from = document.getElementById("fromCurrency").value;
      const to   = document.getElementById("toCurrency").value;

      // 3. live rate  (Frankfurter)
      const url   = `https://api.frankfurter.app/latest?from=${from}&to=${to}`;
      const resp  = await fetch(url);
      const data  = await resp.json();

      if (!data.rates || !data.rates[to]) {
        console.error("No rate returned:", data);
        return;                         // stop gracefully
      }
      const rate = data.rates[to];

      // 4. build converted matrix
      const out = sel.values.map(row =>
        row.map(cell => {
          const n = parseFloat(cell);
          return isNaN(n) ? cell : (n * rate).toFixed(2);
        })
      );

      // 5. write next-column
      sel.getOffsetRange(0, sel.columnCount).values = out;
      await context.sync();
    });
  } catch (err) {
    console.error("Currency add-in error:", err);
  }
}
