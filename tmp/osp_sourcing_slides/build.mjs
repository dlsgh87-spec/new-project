import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile, FileBlob } from '@oai/artifact-tool';

const workspaceDir = 'C:/Users/Administrator/Documents/New project';
const tmp = path.join(workspaceDir, 'tmp/osp_sourcing_slides');
const out = path.join(workspaceDir, 'outputs/osp_sourcing_slides_20260928');
const skill = 'D:/CodexData/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/presentations';
const runtime = 'C:/Users/CHOIIH/.cache/codex-runtimes/codex-primary-runtime/dependencies';
process.env.RUNTIME_NODE_MODULES = path.join(runtime, 'node/node_modules');
await fs.mkdir(tmp, { recursive: true });
await fs.mkdir(out, { recursive: true });
const { finalizePresentation, resolvePresentationFont } = await import(pathToFileURL(path.join(skill, 'container_tools/artifact_tool_utils.mjs')).href);
const font = resolvePresentationFont({ fontFamily: 'Malgun Gothic', availableFonts: ['Malgun Gothic'] });
const p = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const C = { text: '#202626', sub: '#566161', accent: '#007E78', red: '#B03D3D', line: '#D6DDDB', pale: '#F0F5F3' };
const slides = [];

function text(s, value, x, y, w, h, size = 26, opts = {}) {
  const t = s.shapes.add({ geometry: 'textbox', name: opts.name || value.slice(0, 24), position: { left: x, top: y, width: w, height: h }, fill: 'none', line: { fill: 'none', width: 0 } });
  t.text = value;
  t.text.style = { typeface: font, fontSize: size, bold: !!opts.bold, color: opts.color || C.text, autoFit: 'none', wrap: 'square', verticalAlignment: 'middle', alignment: opts.align || 'left', insets: { top: 0, bottom: 0, left: 0, right: 0 } };
  return t;
}
function slide(title, subtitle, notes) {
  const s = p.slides.add();
  s.background.fill = '#FFFFFF';
  slides.push(s);
  text(s, title, 60, 40, 1160, 66, 44, { bold: true, name: 'Slide title' });
  text(s, subtitle, 60, 113, 1160, 43, 25, { color: C.sub });
  text(s, `OSP  2026.09.28`, 60, 677, 400, 24, 17, { color: C.sub });
  text(s, `${slides.length} / 5`, 1120, 677, 100, 24, 17, { color: C.sub, align: 'right' });
  s.speakerNotes.textFrame.setText(notes);
  return s;
}
function table(s, values, x, y, widths, heights, size = 24) {
  const t = s.tables.add({ rows: values.length, columns: widths.length, left: x, top: y, width: widths.reduce((a, b) => a + b, 0), height: heights.reduce((a, b) => a + b, 0), columnWidths: widths, values });
  t.styleOptions = { headerRow: true, bandedRows: false };
  t.borders.assign({ fill: C.line, width: 1, style: 'solid' });
  for (let r = 0; r < values.length; r++) {
    t.rows[r].height = heights[r];
    for (let c = 0; c < widths.length; c++) {
      const cell = t.getCell(r, c);
      cell.fill = r === 0 ? C.accent : r % 2 === 0 ? C.pale : '#FFFFFF';
      cell.text.style = { typeface: font, fontSize: size, color: r === 0 ? '#FFFFFF' : C.text, bold: r === 0, alignment: c === 0 ? 'left' : 'center', verticalAlignment: 'middle', autoFit: 'none', insets: { left: 12, right: 12, top: 6, bottom: 6 } };
    }
  }
  return t;
}

{
  const s = slide('동결건조 간식 개발 목표', '네이처펫 대비 중량당 가격 우위를 목표로 합니다',
    '출처: 2026년 9월 4일 및 7일 시장조사 대화 기록. 당시 쿠팡 링크 https://shop.coupang.com/A00071911/520990?platform=p . 표의 가격은 현재가가 아닌 과거 조사 기록입니다. 메추리 난황은 추가 비교제품이며 해당 브랜드를 확정하지 않습니다. 초기 정상가 10,900~11,900원 검토 후 9월 7일 정상가 약 12,000원, 행사 9,900원, 쿠팡 마진 35%, 행사 전 OSP 마진 20%로 역산했습니다. 9월 28일 표시가 11,900원과 10% 쿠폰을 검토했습니다. 중량·정상가·행사 운영 방식은 아직 확정하지 않았습니다.');
  table(s, [
    ['비교 제품', '중량', '당시 가격'],
    ['치킨 큐브', '300g', '9,900원'],
    ['치킨 텐더', '300g', '11,570원'],
    ['북어 / 열빙어', '140g / 210g', '각 10,900원'],
    ['추가 비교 메추리 난황', '200g', '9,290~9,580원'],
  ], 60, 192, [330, 190, 230], [52, 62, 62, 62, 62], 24);
  text(s, '최근 가격 검토안', 850, 191, 360, 40, 25, { bold: true });
  text(s, '11,900원', 850, 248, 360, 72, 58, { bold: true, color: C.accent });
  text(s, '상시 표시가', 850, 325, 360, 34, 24);
  text(s, '10% 쿠폰 적용 10,710원', 850, 377, 360, 36, 25);
  text(s, '행사 최종가 9,900원', 850, 443, 360, 42, 28, { bold: true });
  text(s, '치킨 300g 기준, 상시 결제가는 비교제품보다 약 8.2% 높습니다', 60, 553, 1150, 42, 28, { bold: true, color: C.red });
  text(s, '기존 기획 정상가 약 12,000원 / 행사 9,900원. 품목별 중량은 미확정', 60, 604, 1150, 35, 24, { color: C.sub });
  text(s, '비교가격은 9월 4일과 7일 조사 기록으로 현재가 재확인이 필요합니다', 60, 645, 1100, 24, 17, { color: C.sub });
}

{
  const s = slide('업체별 단가와 협의 현황', '기존 제품 견적은 후난이 낮지만 인쇄포장 완제품 원가는 미확정입니다',
    '출처: C:/Users/CHOIIH/Downloads/quotation 20260917.xlsx, 후난 9월 17일 견적 Sheet2. B열은 FOB 표기, 200~500kg 가격 구간입니다. 라노바 Yolanda의 최근 이메일은 FOB Tianjin이며 USD/kg 단위는 맥락상 비교 기준으로 사용했고 재확인이 필요합니다. 후난 대구 36.6은 일반 등급, 26은 트리밍 저가 등급입니다. 후난 혼합 열빙어는 암수 각 50%, 라노바 without seeds 19.75는 무란으로 수컷 100%를 뜻하지 않습니다. 난황 8.0은 막 포함 기준으로 통형 생산에 그대로 적용되는지 미확인, 라노바 8.37은 닭 통난황 견적입니다. 라노바 벌크 MOQ 150kg/품목, 비교 견적 요청은 200kg/품목. 후난 포장재 조달 및 충전은 업체, 디자인은 OSP가 담당하기로 협의했습니다. 후난 인쇄봉투 2만개/품목 기준 0.20~0.24달러/개와 색상당 초기 인쇄비 105달러, 소량 디지털 인쇄 단가는 미확정입니다. 원더풀키친은 공유된 견적이 없으며 수치 비교에서 제외했습니다.');
  table(s, [
    ['품목', '후난 페토', '라노바', '비교 시 확인사항'],
    ['치킨', '10.90', '14.75', '큐브 규격과 포장비'],
    ['대구', '36.60 / 26.00', '47.15', '후난 일반 / 저가 등급'],
    ['새우', '86.60', '124.30', '크기와 가열 여부'],
    ['알배기 열빙어', '37.50', '46.50', '원물 크기와 규격'],
    ['혼합 / 무란 열빙어', '27.50 혼합', '19.75 무란', '서로 다른 사양'],
    ['난황', '8.00 막 포함', '8.37 통난황', '후난 통형 가격 확인'],
  ], 60, 190, [275, 230, 220, 435], [49, 47, 47, 47, 47, 47, 47], 24);
  text(s, '후난   지정 상해 창고 운송비 부담 회신, 인쇄포장 단가 확인 필요', 60, 548, 1160, 38, 25);
  text(s, '라노바   인쇄파우치 포함 FOB 텐진 재견적 회신 대기', 60, 596, 1160, 38, 25);
  text(s, '단위 kg당 미달러. 수량·사양·인도 조건이 달라 최종 원가의 직접 비교는 아직 불가합니다', 60, 645, 1160, 24, 17, { color: C.sub });
}

{
  const s = slide('쿠폰과 광고비 반영 손익', '치킨 300g 기존 견적을 기준으로 계산한 참고 시나리오입니다',
    '원가 가정: 후난 치킨 300g 원물 3.27달러 + 소포장 노동·방사선살균·탈산소제·포워더 창고 운송 등 0.23달러 + 무지봉투·스티커 0.26달러 = 3.76달러. 가정 환율 1,400원을 적용한 5,264원이며 추가 수입·국내 물류비와 인쇄포장 변경분은 미반영입니다. 쿠팡 몫은 쿠폰 전 표시가의 35%로 고정하고 할인액은 OSP 전액 부담하는 모델입니다. 부가세 제외 OSP 매출 = [11,900×65%-(11,900-소비자 결제금액)]/1.1. 광고비는 쿠폰 차감 후 OSP 매출의 5% 또는 10%로 가정하며 실제 계약·집행 기준은 확인이 필요합니다. 잔여 마진율=(OSP 매출-원가-광고비)/OSP 매출. 기타 판관비를 차감한 영업이익률이 아닙니다. 상시 쿠폰 적용 후 광고비 10%와 잔여 마진 20%를 함께 확보하려면 총원가 상한=5,950×(1-0.10-0.20)=4,165원입니다. 20%를 광고 차감 후에도 유지한다는 조건부 수치로, 최종 목표 마진은 내부 결정사항입니다.');
  table(s, [
    ['표시가 11,900원', '소비자 결제', 'OSP 매출', '광고 5% 후', '광고 10% 후'],
    ['쿠폰 미적용', '11,900원', '7,032원', '20.1%', '15.1%'],
    ['상시 쿠폰 10%', '10,710원', '5,950원', '6.5%', '1.5%'],
    ['행사 최종가', '9,900원', '5,214원', '-6.0%', '-11.0%'],
  ], 60, 187, [300, 230, 220, 205, 205], [55, 64, 64, 64], 26);
  text(s, '1.5%', 60, 466, 330, 80, 65, { bold: true, color: C.red });
  text(s, '상시 쿠폰과 광고 10% 적용 시\n추가 물류비 반영 전 마진', 60, 552, 500, 74, 25);
  text(s, '4,165원', 660, 466, 550, 80, 65, { bold: true, color: C.accent });
  text(s, '광고 10% 차감 후에도 마진 20%를\n확보하려는 경우의 총원가 상한', 660, 552, 550, 74, 25);
  text(s, '원가 5,264원: 기존 무지봉투·환율 1,400원 가정. 인쇄포장 변경분·추가 물류비 미반영\nOSP 매출은 부가세 제외. 쿠팡은 쿠폰 전 가격의 35%, 광고는 쿠폰 차감 후 OSP 매출 기준 가정', 60, 633, 1160, 40, 17, { color: C.sub });
}

{
  const s = slide('샘플과 본품 운송 조건', '샘플 한국행 특송비와 본품 상해 창고 운송비를 구분합니다',
    '출처: Kerry의 후속 이메일. 후난 샘플 제품은 무료이고 express cost만 OSP 부담. 샘플 필요 g 및 중국 포워더 보유 여부를 질문했으나 정확한 특송료·발송일·수량은 미정입니다. 한국 직송은 OSP의 요청 방향이며 발송 완료가 아닙니다. 후난은 본품을 OSP 지정 상해 창고까지 추가 운송비 없이 보내겠다고 회신했습니다. 라노바 샘플 제품비 및 특송료 부담 조건은 미확정입니다. 첨부 사진 C:/Users/CHOIIH/Downloads/FD egg yolk.png 는 후난이 보낸 막 포함 동결건조 통난황 형태 참고자료이며 가열 여부와 발주 MOQ는 확정하지 않습니다. 물류 원칙 https://academy.iccwbo.org/incoterms/article/incoterms-2020-fca-or-fob/ : FCA와 FOB 모두 수출통관은 판매자 책임. 지정 창고 FCA는 도착차량 하역 준비 상태 인도, FOB는 본선 적재 완료 인도입니다. FCA 절감은 미확정이며 지정창고 이후 포워더 비용과 FOB 가산비용 비교가 필요합니다. 제조와 소포장은 후난에서 진행하므로 상해 인건비가 절감 근거가 아닙니다.');
  text(s, '샘플 평가', 60, 181, 490, 43, 30, { bold: true });
  text(s, '품질과 제형, 기호성 확인', 60, 237, 530, 43, 29, { bold: true, color: C.accent });
  text(s, '후난  제품 무상 / 특송비 OSP 부담\n라노바  샘플비와 특송비 조건 확인\n한국 직송료와 품목별 필요량 미정', 60, 294, 540, 122, 25);
  const bytes = await fs.readFile('C:/Users/CHOIIH/Downloads/FD egg yolk.png');
  s.images.add({ blob: bytes, contentType: 'image/png', alt: '후난에서 제시한 막 포함 동결건조 통난황 형태', fit: 'contain', position: { left: 60, top: 445, width: 225, height: 180 } });
  text(s, '통난황 형태 확인\n생산수량 추가 확인', 315, 484, 280, 87, 24);
  text(s, '본품 물류 비교', 660, 181, 550, 43, 30, { bold: true });
  table(s, [
    ['비교안', '인도 지점'],
    ['후난 FCA 상해', '지정 창고'],
    ['후난 FOB 상해', '본선 적재 완료'],
    ['라노바 FOB 텐진', '본선 적재 완료'],
  ], 660, 239, [310, 250], [50, 57, 57, 57], 24);
  text(s, '후난 공장에서 생산과 포장 진행\n상해 지정 창고까지 업체 운송비 부담', 660, 486, 550, 78, 24);
  text(s, 'FCA 절감 여부는 아직 미확정\n동일한 국내 창고 입고원가로 비교', 660, 579, 550, 69, 25, { bold: true, color: C.accent });
  text(s, '후난 FOB 상해 견적은 추가 확인 필요', 660, 649, 550, 24, 17, { color: C.sub });
}

{
  const s = slide('오늘 결정할 사항', '최종 판매가와 발주를 정하기 전에 확인할 항목입니다',
    '의사결정 요청이며 이미 승인되거나 확정된 사항이 아닙니다. 샘플 품목과 g수 및 특송비 예산을 결정한 뒤 한국 직송 특송 견적을 받아 승인합니다. 평시·행사 최소 허용 마진과 광고비 산정 기준을 내부 확정합니다. 후난에는 인쇄포장 완제품 FCA 상해 지정창고와 FOB 상해 견적을 동일 수량·사양으로 요청하고 라노바의 인쇄포장 FOB 텐진 견적을 회신받습니다. 제품 중량과 주문량은 포장비·해상운임·수입 부대비용·국내 물류비까지 합산한 후 최종 결정합니다. 공급업체 MOQ, 인쇄포장 MOQ 및 미사용 포장재 보관 조건도 확인이 필요합니다. 재고가 늘어나는 대량 인쇄를 단가만 보고 확정하지 않습니다.');
  const items = [
    ['01', '샘플 품목과 필요량', '품질·제형·기호성 테스트 규모 결정'],
    ['02', '샘플 특송비 예산', '한국 직송 견적 수령 후 발송 승인'],
    ['03', '평시와 행사 최소 마진', '쿠폰 전액 부담과 광고비 5~10% 반영'],
    ['04', '제품 중량과 주문량', '인쇄포장 및 국내 입고원가 확인 후 결정'],
  ];
  items.forEach(([num, title, detail], i) => {
    const y = 198 + i * 99;
    text(s, num, 60, y, 80, 58, 44, { bold: true, color: C.accent });
    text(s, title, 180, y, 480, 58, 31, { bold: true });
    text(s, detail, 675, y, 545, 58, 26);
  });
  text(s, '후속 요청  양사 인쇄포장 재견적과 포워더 국내 입고비용 확인', 60, 615, 1160, 40, 25, { color: C.sub });
}

const candidatePath = path.join(tmp, 'candidate.pptx');
await (await PresentationFile.exportPptx(p)).save(candidatePath);
const finalPath = path.join(out, 'OSP_동결건조_화면보고_20260928_v2.pptx');
const result = await finalizePresentation({
  workspaceDir, candidatePath, finalPath,
  pythonExecutable: path.join(runtime, 'python/python.exe'),
  integrityValidatorPath: path.join(skill, 'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath: path.join(skill, 'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs: ['--expected-slide-size-emu', '12192000,6858000', '--validate-bullet-geometry', '--validate-heading-fit', '--require-native-table-slide', '1', '--require-native-table-slide', '2', '--require-native-table-slide', '3', '--require-native-table-slide', '4'],
  explicitTotalSlideCount: 5,
  requiredNativeTableOwnerSlides: [1, 2, 3, 4],
  fontPolicy: { basis: 'design', families: [font] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(tmp, 'validation-v2.json'),
});
console.log(JSON.stringify({ finalPath, result }));
const reloaded = await PresentationFile.importPptx(await FileBlob.load(finalPath));
for (let i = 0; i < 5; i++) {
  const s = reloaded.slides.items[i];
  const blob = await reloaded.export({ slide: s, format: 'png', scale: 1.2 });
  await fs.writeFile(path.join(tmp, `slide-${i + 1}.png`), new Uint8Array(await blob.arrayBuffer()));
}
console.log('Rendered five final slides');
