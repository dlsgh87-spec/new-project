import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
const outputDir = process.argv[3];

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table,definedName,drawing",
  maxChars: 12000,
  tableMaxRows: 20,
  tableMaxCols: 20,
  tableMaxCellChars: 120,
});
console.log("SUMMARY_START");
console.log(summary.ndjson);
console.log("SUMMARY_END");

const sheets = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 4000 });
console.log("SHEETS_START");
console.log(sheets.ndjson);
console.log("SHEETS_END");

await fs.mkdir(outputDir, { recursive: true });
for (const sheet of workbook.worksheets.items) {
  const used = sheet.getUsedRange();
  if (!used) continue;
  const address = used.address;
  const table = await workbook.inspect({
    kind: "table",
    sheetId: sheet.name,
    range: address,
    include: "values,formulas",
    maxChars: 30000,
    tableMaxRows: 100,
    tableMaxCols: 30,
    tableMaxCellChars: 200,
  });
  console.log(`TABLE_START:${sheet.name}:${address}`);
  console.log(table.ndjson);
  console.log(`TABLE_END:${sheet.name}`);

  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 1.5,
    format: "png",
  });
  const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(`${outputDir}/${safeName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
