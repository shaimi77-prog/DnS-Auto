import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const [baselinePath, patchedPath, reportPath] = process.argv.slice(2);
const baseline = await SpreadsheetFile.importXlsx(await FileBlob.load(baselinePath));
const patched = await SpreadsheetFile.importXlsx(await FileBlob.load(patchedPath));
const baselineSheets = baseline.worksheets.items;
const patchedSheets = patched.worksheets.items;
const report = {
  sheet_names_equal: JSON.stringify(baselineSheets.map((s) => s.name)) === JSON.stringify(patchedSheets.map((s) => s.name)),
  cell_difference_count: 0,
  formula_difference_count: 0,
  dimension_difference_count: 0,
  sheets: [],
};
for (const baselineSheet of baselineSheets) {
  const patchedSheet = patched.worksheets.getItem(baselineSheet.name);
  const baselineUsed = baselineSheet.getUsedRange();
  const patchedUsed = patchedSheet.getUsedRange();
  const baselineValues = baselineUsed?.values ?? [];
  const patchedValues = patchedUsed?.values ?? [];
  const baselineFormulas = baselineUsed?.formulas ?? [];
  const patchedFormulas = patchedUsed?.formulas ?? [];
  const dimensionsEqual = baselineValues.length === patchedValues.length &&
    (baselineValues[0]?.length ?? 0) === (patchedValues[0]?.length ?? 0);
  if (!dimensionsEqual) report.dimension_difference_count += 1;
  let cellDiff = 0;
  let formulaDiff = 0;
  const rows = Math.max(baselineValues.length, patchedValues.length);
  const cols = Math.max(baselineValues[0]?.length ?? 0, patchedValues[0]?.length ?? 0);
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      if (JSON.stringify(baselineValues[row]?.[col] ?? null) !== JSON.stringify(patchedValues[row]?.[col] ?? null)) cellDiff += 1;
      if (JSON.stringify(baselineFormulas[row]?.[col] ?? null) !== JSON.stringify(patchedFormulas[row]?.[col] ?? null)) formulaDiff += 1;
    }
  }
  report.cell_difference_count += cellDiff;
  report.formula_difference_count += formulaDiff;
  report.sheets.push({ name: baselineSheet.name, dimensions_equal: dimensionsEqual, cell_difference_count: cellDiff, formula_difference_count: formulaDiff });
}
await fs.writeFile(reportPath, JSON.stringify(report, null, 2), "utf8");
console.log(JSON.stringify(report));
