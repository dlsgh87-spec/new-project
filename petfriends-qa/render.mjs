import fs from 'node:fs/promises';
import { SpreadsheetFile } from '@oai/artifact-tool';

const bytes = Buffer.from(await fs.readFile(new URL('./export.b64', import.meta.url), 'utf8'), 'base64');
const workbook = await SpreadsheetFile.importXlsx(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
const sheets = [
  ['dashboard-top', '펫프렌즈 운영 대시보드', 'A1:N34'],
  ['dashboard-bottom', '펫프렌즈 운영 대시보드', 'A35:N71'],
  ['sales', 'SKU별 판매현황', 'A1:P10'],
  ['stock', 'SKU별 재고', 'A1:Q10'],
  ['master', 'SKU 기준정보', 'A1:N10'],
  ['monthly', '월별 운영집계', 'A1:R12'],
  ['settings', '운영 설정', 'A1:C30'],
];
for (const [name, sheetName, range] of sheets) {
  const image = await workbook.render({sheetName, range, scale:1, format:'png'});
  await fs.writeFile(new URL(`./${name}.png`, import.meta.url), new Uint8Array(await image.arrayBuffer()));
  console.log(`Rendered ${name}`);
}
