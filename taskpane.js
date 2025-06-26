// taskpane.js  –  v2  (uses the exact selection)
Office.onReady(() => {
  document.getElementById("convert").addEventListener("click", convertCurrency);
});

async function convertCurrency() {
  try {
    await Excel.run(async (context) => {
      // 1. selected cells – not the whole column
      const sel = context.workbook.getSelectedRange();
      sel.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      // 2. currencies
      const from = document.getElementById("fromCurrency").value;
      const to   = document.getElementById("toCurrency").value;

      // 3. live rate
      const res   = await fetch(`https://api.exchangerate.host/latest?base=${from}&symbols=${to}`);
      const json  = await res.json();
      const rate  = json.rates[to];

      // 4. build converted matrix (preserves shape)
      const out = sel.values.map(row =>
        row.map(cell => {
          const n = parseFloat(cell);
          return isNaN(n) ? cell : (n * rate).toFixed(2);
        })
      );

      // 5. write to the first empty column to the right of the selection
      const dest = sel.getOffsetRange(0, sel.columnCount);
      dest.values = out;

      await context.sync();
    });
  } catch (err) {
    console.error("Currency add-in error:", err);
    // Optional: tiny in-sheet alert
    OfficeRuntime.displayDialogAsync(
      "about:blank",
      { height: 20, width: 30, displayInIframe: true },
      dlg => dlg.value.setHtml(`<p style='font-family:Arial;padding:12px'>Conversion failed:<br>${err.message}</p>`)
    );
  }
}
