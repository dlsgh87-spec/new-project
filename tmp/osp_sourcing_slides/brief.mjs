import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile, FileBlob } from '@oai/artifact-tool';

const workspaceDir = 'C:/Users/Administrator/Documents/New project';
const tmp = path.join(workspaceDir, 'tmp/osp_sourcing_slides/brief');
const out = path.join(workspaceDir, 'outputs/osp_sourcing_slides_20260928');
const skill = 'D:/CodexData/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/presentations';
const runtime = 'C:/Users/CHOIIH/.cache/codex-runtimes/codex-primary-runtime/dependencies';
process.env.RUNTIME_NODE_MODULES = path.join(runtime, 'node/node_modules');
await fs.mkdir(tmp, { recursive: true });
const { finalizePresentation, resolvePresentationFont } = await import(pathToFileURL(path.join(skill, 'container_tools/artifact_tool_utils.mjs')).href);
const font = resolvePresentationFont({ fontFamily: 'Malgun Gothic', availableFonts: ['Malgun Gothic'] });
const p = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const C = { text: '#202626', sub: '#566161', accent: '#007E78', line: '#D6DDDB', pale: '#F0F5F3' };

function text(s, value, x, y, w, h, size = 26, opts = {}) {
  const t = s.shapes.add({ geometry: 'textbox', name: opts.name || value.slice(0, 24), position: { left: x, top: y, width: w, height: h }, fill: 'none', line: { fill: 'none', width: 0 } });
  t.text = value;
  t.text.style = { typeface: font, fontSize: size, bold: !!opts.bold, color: opts.color || C.text, autoFit: 'none', wrap: 'square', verticalAlignment: 'middle', alignment: opts.align || 'left', insets: { top: 0, bottom: 0, left: 0, right: 0 } };
}
function slide(title, subtitle, notes) {
  const s = p.slides.add();
  s.background.fill = '#FFFFFF';
  text(s, title, 60, 40, 1160, 66, 44, { bold: true, name: 'Slide title' });
  text(s, subtitle, 60, 113, 1160, 43, 25, { color: C.sub });
  text(s, 'OSP  2026.09.28', 60, 677, 400, 24, 17, { color: C.sub });
  text(s, `${p.slides.items.length} / 3`, 1120, 677, 100, 24, 17, { color: C.sub, align: 'right' });
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
}

{
  const s = slide('현재 업무 진행사항', '양사 제품 견적 수령 후, 인쇄포장 견적과 샘플 조건을 확인 중입니다',
    '출처: 사용자 제공 후난 Kerry Xiong 및 라노바 Yolanda Wang 이메일, quotation 20260917.xlsx. 후난은 대구 등살 중심 제품과 여러 부위 트리밍 제품의 두 등급을 제시했고 비교 샘플 제공이 가능합니다. 후난의 막 포함 통난황 사진은 OSP가 원한 형태라고 확인했으나 생산 가능 수량과 통형 적용 단가는 추가 확인이 필요합니다. 후난 포장재 조달은 공급업체, 디자인은 OSP가 담당하기로 했습니다. 후난은 본품을 OSP 지정 상해 창고까지 추가 운송비 없이 보내겠다고 회신했습니다. FCA 조건의 세부 비용 및 수출통관 포함 여부는 최종 견적에서 확인해야 하며 샘플 한국행 무료 특송을 의미하지 않습니다. 라노바 제품별 FOB 텐진 견적을 받았고, 사용자 확인에 따라 인쇄포장 완제품 FOB 텐진 재견적 요청 메일을 발송한 상태입니다. 최종 발주나 포장규격은 미확정입니다.');
  table(s, [
    ['구분', '후난 페토', '라노바'],
    ['제품 견적', '수정 견적 수령\n대구 2등급 제시', '제품별 견적 수령\n알래스카 명태 가격 대기'],
    ['제품 확인', '통난황 사진상 형태 확인\n실물 샘플 평가 전', '통난황 견적 수령\n실물 샘플 평가 전'],
    ['인쇄포장', '업체 조달 / OSP 디자인\n최종 포장 단가 미확정', '인쇄포장 완제품 재견적 요청\n회신 대기'],
    ['본품 운송', '지정 상해 창고까지\n업체 운송비 부담 회신', 'FOB 텐진\n현재 견적 기준'],
  ], 60, 185, [250, 455, 455], [54, 90, 90, 90, 90], 24);
  text(s, '다음 단계   인쇄포장 단가 확인 및 샘플 수령 준비', 60, 618, 1160, 42, 28, { bold: true, color: C.accent });
}

{
  const s = slide('업체별 단가', '기존 제품 견적 비교 | kg당 미달러 | 인쇄포장 완제품 단가는 미확정',
    '출처: C:/Users/CHOIIH/Downloads/quotation\u00a020260917.xlsx, 후난 9월 17일 견적 Sheet2 B열, 200~500kg 견적 가격 구간. 후난 표는 FOB 표기이나 이후 상해 지정 창고 인도 협의가 있으므로 최종 인도조건과 포함 비용을 재확인해야 합니다. 라노바 Yolanda 이메일 FOB Tianjin 제품 가격. USD/kg 단위는 맥락상 비교 기준이며 최종 서면 견적에서 확인할 사항입니다. 후난 대구 36.60은 등살 중심, 26.00은 여러 부위 트리밍으로 지방이 다소 높은 별도 등급입니다. 후난 혼합 열빙어는 암수 각 50%, 라노바 without seeds 19.75는 무란이며 수컷 100%라고 확정할 수 없습니다. 후난 난황 8.00은 막 포함 견적으로 통형에도 같은 단가가 적용되는지 미확인입니다. 라노바 8.37은 닭 통난황 견적입니다. 알래스카 명태는 후난 공유 견적에 별도 제시되지 않았고 라노바는 아직 가격을 받지 못했다고 회신했습니다. 후난 제품별 포장비, 인쇄비, 국제운송 등은 이 표의 제품 단가에 일괄 포함되어 있지 않습니다. 라노바 벌크 최소주문량은 품목별 150kg이며 후난 표의 200~500kg는 가격 구간으로 최소주문량을 확정하는 근거가 아닙니다.');
  table(s, [
    ['품목', '후난 페토', '라노바', '비교 시 확인사항'],
    ['치킨', '10.90', '14.75', '큐브 규격과 포장비'],
    ['대구', '36.60 / 26.00', '47.15', '후난 등살 중심 / 트리밍'],
    ['새우', '86.60', '124.30', '크기와 가열 여부'],
    ['알배기 열빙어', '37.50', '46.50', '원물 크기와 규격'],
    ['혼합 / 무란 열빙어', '27.50 혼합', '19.75 무란', '서로 다른 사양'],
    ['난황', '8.00 막 포함', '8.37 통난황', '후난 통형 적용 가격 확인'],
    ['알래스카 명태', '미제시', '가격 대기', '견적 수령 후 검토'],
  ], 60, 180, [275, 230, 220, 435], [54, 47, 47, 47, 47, 47, 47, 47], 24);
  text(s, '단순 제품가격은 후난이 낮은 편이나, 사양과 인도 조건이 다릅니다', 60, 588, 1160, 40, 27, { bold: true });
  text(s, '후난 9/17 견적·라노바 FOB 텐진 메일 기준. 포장비와 운송 포함 범위 확인 후 최종 비교', 60, 636, 1160, 27, 18, { color: C.sub });
}

{
  const s = slide('샘플 평가 및 특송 협의', '샘플 목적: 품질·제형·기호성 테스트',
    '출처: 사용자 제공 Kerry 이메일. 후난은 샘플 제품 무료, 특송료만 OSP 부담이라고 안내했고 품목별 필요 중량(g)과 중국 포워더 보유 여부를 질문했습니다. 한국행 특송의 정확한 요금, 샘플 물량 및 출고일은 확정되지 않았습니다. 한국 직송은 OSP의 검토 방향이며 공급업체가 출고 완료했다는 의미가 아닙니다. 라노바는 연구개발부서와 샘플을 협의하고 특송 발송을 준비하겠다고 했으나 제품 샘플비 및 특송료의 부담 주체는 확인되지 않았습니다. 오늘의 논의안은 품목과 필요량을 정하고 한국 직송 특송료 견적을 수령한 뒤 OSP 비용 부담 승인 및 발송을 진행하는 것입니다. 예산액과 실제 발송 승인은 아직 확정하지 않았습니다. 후난의 공장부터 상해 지정 창고까지 무료 운송 회신은 본품에 대한 별도 조건이며 한국행 샘플 특송비 면제를 뜻하지 않습니다.');
  table(s, [
    ['구분', '후난 페토', '라노바'],
    ['샘플 제품비', '무상 제공', '확인 필요'],
    ['특송비', 'OSP 부담 / 금액 미정', '부담 주체·금액 확인 필요'],
    ['준비 상태', '품목별 필요량 회신 필요', '샘플 준비와 일정 확인 필요'],
  ], 60, 184, [250, 455, 455], [54, 65, 65, 65], 25);
  text(s, '오늘 논의할 사항', 60, 463, 1160, 42, 31, { bold: true, color: C.accent });
  text(s, '샘플 품목과 필요량 결정', 60, 523, 510, 42, 29, { bold: true });
  text(s, '한국 직송 특송료 견적 후 비용 승인', 625, 523, 595, 42, 28, { bold: true });
  text(s, '후난 샘플 특송비는 OSP 부담으로 검토하며, 금액 확인 후 발송을 진행합니다', 60, 585, 1160, 40, 25);
  text(s, '본품의 상해 지정 창고 운송비 부담 조건과 샘플 한국행 특송비는 별도입니다', 60, 638, 1160, 27, 18, { color: C.sub });
}

const candidatePath = path.join(tmp, 'candidate.pptx');
await (await PresentationFile.exportPptx(p)).save(candidatePath);
const finalPath = path.join(out, 'OSP_진행현황_단가_샘플_3장_20260928.pptx');
const result = await finalizePresentation({
  workspaceDir, candidatePath, finalPath,
  pythonExecutable: path.join(runtime, 'python/python.exe'),
  integrityValidatorPath: path.join(skill, 'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath: path.join(skill, 'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs: ['--expected-slide-size-emu', '12192000,6858000', '--validate-bullet-geometry', '--validate-heading-fit', '--require-native-table-slide', '1', '--require-native-table-slide', '2', '--require-native-table-slide', '3'],
  explicitTotalSlideCount: 3,
  requiredNativeTableOwnerSlides: [1, 2, 3],
  fontPolicy: { basis: 'design', families: [font] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(tmp, 'validation.json'),
});
console.log(JSON.stringify({ finalPath, status: result.packageIntegrity.status, slideCount: result.firstPartyImport.slideCount }));
const reloaded = await PresentationFile.importPptx(await FileBlob.load(finalPath));
for (let i = 0; i < reloaded.slides.items.length; i++) {
  const blob = await reloaded.export({ slide: reloaded.slides.items[i], format: 'png', scale: 1.2 });
  await fs.writeFile(path.join(tmp, `slide-${i + 1}.png`), new Uint8Array(await blob.arrayBuffer()));
}
console.log('Rendered all three final slides');
