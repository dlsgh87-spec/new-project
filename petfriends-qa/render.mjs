import fs from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const timer = setTimeout(() => { console.error('Preview exceeded 120 seconds'); process.exit(2); }, 120000);
console.log('Importing cached preview');
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(fileURLToPath(new URL('./preview-cache.xlsx', import.meta.url))));
console.log('Imported');
const sheets = [
  ['sales-labels-v8', 'SKU별 판매현황', 'A1:N20'],
];
for (const [name, sheetName, range] of sheets) {
  const image = await workbook.render({sheetName, range, scale:1, format:'png'});
  await fs.writeFile(new URL(`./${name}.png`, import.meta.url), new Uint8Array(await image.arrayBuffer()));
  console.log(`Rendered ${name}`);
}
clearTimeout(timer);
