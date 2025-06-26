Office.onReady(() => {
  document.getElementById("convert").onclick = convertCurrency;
});

async function convertCurrency() {
  try {
    await Excel.run(async (context) => {
      const fromCurrency = document.getElementById("fromCurrency").value;
      const toCurrency = document.getElementById("toCurrency").value;

      // Get user's selected range
      const range = context.workbook.getActiveCell().getEntireColumn().getUsedRange();
      range.load("values, rowCount, columnIndex");

      await context.sync();

      // Fetch live exchange rate
      const response = await fetch(`https://api.exchangerate.host/latest?base=${fromCurrency}&symbols=${toCurrency}`);
      const data = await response.json();
      const rate = data.rates[toCurrency];

      // Calculate new values
      const newValues = range.values.map(row => {
        const num = parseFloat(row[0]);
        return [isNaN(num) ? "N/A" : (num * rate).toFixed(2)];
      });

      // Insert values into next column
      const nextColumn = range.getOffsetRange(0, 1);
      nextColumn.values = newValues;

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
