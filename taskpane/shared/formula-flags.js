/* global Excel */

// Range.formulas returns the value for non-formula cells too. A literal text
// value such as "=1+1" therefore needs an explicit formula-presence check.
export async function getFormulaFlags(sheetName, addresses) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const checks = addresses.map((address) => {
      const formulaCells = sheet.getRange(address).getSpecialCellsOrNullObject("Formulas");
      return [address, formulaCells];
    });
    await context.sync();
    return {
      success: true,
      flags: Object.fromEntries(checks.map(([address, cells]) => [address, !cells.isNullObject])),
    };
  });
}
