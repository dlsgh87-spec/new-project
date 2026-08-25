import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const SOURCE_HTML = "C:\\Users\\CHOIIH\\Desktop\\중국박람회 사진\\PET_FAIR_ASIA_2026_미팅사진매칭_메모업체만_260820.html";
const FINAL_PPTX = "C:\\Users\\CHOIIH\\Desktop\\중국박람회 사진\\PET_FAIR_ASIA_2026_미팅정리_PPT_260825.pptx";
const TMP_DIR = "C:\\Users\\Administrator\\Documents\\New project\\outputs\\pet_fair_ppt_260825";
const RENDER_DIR = path.join(TMP_DIR, "rendered");

const C = {
  navy: "#15233B",
  ink: "#172033",
  muted: "#667085",
  line: "#D8E0EB",
  bg: "#F5F7FB",
  white: "#FFFFFF",
  gold: "#F5C542",
  goldSoft: "#FFF5CC",
  blue: "#275BD6",
  blueSoft: "#EAF1FF",
  green: "#11856F",
  greenSoft: "#E6F8F4",
  orange: "#B45309",
  orangeSoft: "#FFF2E6",
  red: "#B42318",
  redSoft: "#FFF1F0",
  slate: "#344054",
};

function decodeEntities(input = "") {
  return input
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)));
}

function stripTags(input = "") {
  return decodeEntities(
    input
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>|<\/div>|<\/li>|<\/tr>/gi, "\n")
      .replace(/<[^>]*>/g, "")
  )
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

function oneLine(input = "") {
  return stripTags(input).replace(/\s*\n\s*/g, " ").replace(/\s{2,}/g, " ").trim();
}

function shorten(text = "", max = 110) {
  const clean = oneLine(text);
  if (clean.length <= max) return clean;
  return clean.slice(0, max - 1).trimEnd() + "...";
}

function matchOne(html, regex) {
  const m = html.match(regex);
  return m ? m[1] : "";
}

function extractList(block, cls) {
  const listHtml = matchOne(block, new RegExp(`<ul class="${cls}">([\\s\\S]*?)<\\/ul>`, "i"));
  return [...listHtml.matchAll(/<li>([\s\S]*?)<\/li>/gi)].map((m) => oneLine(m[1])).filter(Boolean);
}

function extractChips(block) {
  return [...block.matchAll(/<span class="chip">([\s\S]*?)<\/span>/gi)].map((m) => oneLine(m[1])).filter(Boolean);
}

function extractInfo(block) {
  const items = {};
  const re = /<div class="info-item[^"]*">\s*<div class="label">([\s\S]*?)<\/div>\s*<div class="value">([\s\S]*?)<\/div>\s*<\/div>/gi;
  for (const m of block.matchAll(re)) {
    items[oneLine(m[1])] = stripTags(m[2]);
  }
  return items;
}

function extractQuotes(block) {
  const tbody = matchOne(block, /<tbody>([\s\S]*?)<\/tbody>/i);
  const rows = [];
  for (const rm of tbody.matchAll(/<tr>([\s\S]*?)<\/tr>/gi)) {
    const cells = [...rm[1].matchAll(/<td>([\s\S]*?)<\/td>/gi)].map((m) => oneLine(m[1]));
    if (cells.length) rows.push(cells);
  }
  return rows;
}

function extractPhotos(block) {
  const photos = [];
  const re = /<figure class="photo">([\s\S]*?)<\/figure>/gi;
  for (const fm of block.matchAll(re)) {
    const fig = fm[1];
    const img = fig.match(/<img src="(data:image\/[^"]+)" alt="([^"]*)"/i);
    const cap = matchOne(fig, /<figcaption>([\s\S]*?)<\/figcaption>/i);
    const file = matchOne(cap, /<span>([\s\S]*?)<\/span>/i);
    const caption = oneLine(cap.replace(/<span>[\s\S]*?<\/span>/i, ""));
    if (img) {
      photos.push({
        dataUrl: img[1],
        alt: oneLine(img[2]) || caption || "PET FAIR ASIA booth photo",
        caption,
        file: oneLine(file),
      });
    }
  }
  return photos;
}

function extractHtmlData(html) {
  const heroTitle = oneLine(matchOne(html, /<h1>([\s\S]*?)<\/h1>/i));
  const heroSubtitle = oneLine(matchOne(html, /<header class="hero">[\s\S]*?<p>([\s\S]*?)<\/p>/i));
  const stats = [...html.matchAll(/<div class="hero-stat"><b>([\s\S]*?)<\/b><span>([\s\S]*?)<\/span><\/div>/gi)]
    .map((m) => ({ value: oneLine(m[1]), label: oneLine(m[2]) }));

  const priorities = [...html.matchAll(/<div class="priority-card"><strong>([\s\S]*?)<\/strong><span>([\s\S]*?)<\/span><\/div>/gi)]
    .map((m) => ({ title: oneLine(m[1]), body: oneLine(m[2]) }));

  const vendors = [...html.matchAll(/<section class="vendor">([\s\S]*?)<\/section>/gi)].map((m) => {
    const block = m[1];
    return {
      title: oneLine(matchOne(block, /<h3>([\s\S]*?)<\/h3>/i)),
      headline: oneLine(matchOne(block, /<p class="headline">([\s\S]*?)<\/p>/i)),
      memo: oneLine(matchOne(block, /<div class="memo">([\s\S]*?)<\/div>/i)),
      tags: [...block.matchAll(/<span class="tag[^"]*">([\s\S]*?)<\/span>/gi)].map((tm) => oneLine(tm[1])),
      info: extractInfo(block),
      quotes: extractQuotes(block),
      risks: extractList(block, "risk-list"),
      actions: extractChips(block),
      photos: extractPhotos(block),
    };
  });

  const footer = matchOne(html, /<div class="footer">([\s\S]*?)<\/div>\s*<\/main>/i);
  const footerTitle = oneLine(matchOne(footer, /<h3>([\s\S]*?)<\/h3>/i)) || "추가 체크 항목";
  let footerItems = [...footer.matchAll(/<li>([\s\S]*?)<\/li>/gi)].map((m) => oneLine(m[1]));
  if (!footerItems.length) {
    footerItems = [
      "계약 주체와 수출 서류 발행 주체를 업체별로 확정합니다.",
      "FOB/CIF 부산 조건, 내륙 운임, 포장비 포함 여부를 분리 확인합니다.",
      "MOQ가 품목별 기준인지 전체 발주 기준인지 재대조합니다.",
      "성분표, 제조공정도, 검역 서류, 한국 사료 수입 대응 범위를 요청합니다.",
      "소용량 토핑, 캔, 동결건조 원물별 샘플 우선순위를 확정합니다.",
      "기능성 문구는 국내 표시 가능 범위 안에서 별도 검토합니다.",
    ];
  }

  return { heroTitle, heroSubtitle, stats, priorities, vendors, footerTitle, footerItems };
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

function addShape(slide, position, fill = "none", lineFill = "none", radius = undefined) {
  return slide.shapes.add({
    geometry: "roundRect",
    position,
    fill,
    line: { style: "solid", fill: lineFill, width: lineFill === "none" ? 0 : 1 },
    ...(radius ? { borderRadius: radius } : {}),
  });
}

function addText(slide, text, position, style = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  box.text = text || "";
  box.text.style = {
    typeface: "Malgun Gothic",
    fontSize: style.fontSize ?? 18,
    color: style.color ?? C.ink,
    bold: style.bold ?? false,
    alignment: style.alignment ?? "left",
    verticalAlignment: style.verticalAlignment ?? "top",
    lineSpacing: style.lineSpacing ?? 1.15,
    wrap: "square",
    autoFit: style.autoFit ?? "shrinkText",
    insets: style.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return box;
}

function addBulletText(slide, items, position, options = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  const paragraphs = items.filter(Boolean).map((item) => ({
    bulletCharacter: "•",
    marginLeft: 18,
    indent: -10,
    runs: [shorten(item, options.max ?? 110)],
  }));
  box.text.set(paragraphs.length ? paragraphs : ["확인 필요"]);
  box.text.style = {
    typeface: "Malgun Gothic",
    fontSize: options.fontSize ?? 17,
    color: options.color ?? C.ink,
    lineSpacing: 1.14,
    wrap: "square",
    autoFit: "shrinkText",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return box;
}

function addTopBar(slide, title, subtitle = "", page = "") {
  addText(slide, title, { left: 64, top: 38, width: 800, height: 50 }, { fontSize: 35, bold: true, color: C.navy, lineSpacing: 0.95 });
  if (subtitle) addText(slide, subtitle, { left: 66, top: 88, width: 770, height: 36 }, { fontSize: 17, color: C.muted });
  addShape(slide, { left: 64, top: 128, width: 1120, height: 3 }, C.gold, C.gold);
  if (page) addText(slide, page, { left: 1110, top: 44, width: 90, height: 26 }, { fontSize: 15, color: C.muted, alignment: "right" });
}

function addTag(slide, text, x, y, color = C.blue, fill = C.blueSoft) {
  addShape(slide, { left: x, top: y, width: Math.min(170, 56 + text.length * 12), height: 26 }, fill, "none", "rounded-full");
  addText(slide, text, { left: x + 10, top: y + 4, width: Math.min(150, 36 + text.length * 12), height: 18 }, { fontSize: 13, bold: true, color });
}

function addPhoto(slide, photo, pos, idx) {
  addShape(slide, pos, C.white, C.line, "rounded-lg");
  const imgPos = { left: pos.left + 8, top: pos.top + 8, width: pos.width - 16, height: pos.height - 48 };
  slide.images.add({
    dataUrl: photo.dataUrl,
    alt: photo.alt || `PET FAIR ASIA photo ${idx}`,
    fit: "contain",
    position: imgPos,
    geometry: "roundRect",
    borderRadius: "rounded-md",
  });
  addText(slide, shorten(photo.caption || photo.alt || "부스/샘플 사진", 38), { left: pos.left + 10, top: pos.top + pos.height - 35, width: pos.width - 20, height: 18 }, { fontSize: 12, bold: true, color: C.slate });
  addText(slide, shorten(photo.file || "", 42), { left: pos.left + 10, top: pos.top + pos.height - 18, width: pos.width - 20, height: 14 }, { fontSize: 10, color: C.muted });
}

function addNotes(slide, sourceLines) {
  slide.speakerNotes.textFrame.setText(["[Sources]", ...sourceLines]);
  slide.speakerNotes.setVisible(true);
}

function firstAvailable(info, labels, max = 120) {
  for (const label of labels) {
    if (info[label]) return shorten(info[label], max);
  }
  return "확인 필요";
}

function addInfoRows(slide, rows, x, y, w, rowH = 38) {
  rows.forEach((row, i) => {
    const top = y + i * rowH;
    const fill = i % 2 === 0 ? "#FFFFFF" : "#FAFCFF";
    addShape(slide, { left: x, top, width: w, height: rowH - 2 }, fill, C.line, "rounded-md");
    addText(slide, row.label, { left: x + 12, top: top + 9, width: 92, height: 18 }, { fontSize: 14, bold: true, color: C.muted });
    addText(slide, row.value, { left: x + 110, top: top + 8, width: w - 122, height: 22 }, { fontSize: 15, color: C.ink });
  });
}

function vendorRole(vendor) {
  if (/Wonderful/i.test(vendor.title)) return "캔+동결건조";
  if (/Today|Ningbo/i.test(vendor.title)) return "적육/대용량 캔";
  if (/Peto|Hunan/i.test(vendor.title)) return "소용량 토핑";
  if (/Jinchong/i.test(vendor.title)) return "소용량 토핑";
  if (/ELF|JiLan/i.test(vendor.title)) return "해산물 원물";
  if (/Naki/i.test(vendor.title)) return "프리미엄 원물";
  if (/Haci|Harts|Medellin/i.test(vendor.title)) return "덴탈/저키";
  const text = `${vendor.title} ${vendor.headline} ${(vendor.info["가능 품목"] || "")}`;
  if (/Today|캔|85g|170g|400g/i.test(text)) return "저가/적육 캔";
  if (/Wonderful/i.test(text)) return "캔+동결건조";
  if (/Peto|Jinchong|소용량|난황|캣닢/i.test(text)) return "소용량 토핑";
  if (/ELF|JiLan|해산물|새우|관자|가리비/i.test(text)) return "해산물 원물";
  if (/Naki|Dental|치아/i.test(text)) return "기능성 저키";
  if (/Haci|Medellin|oral/i.test(text)) return "덴탈/저키";
  return "후보";
}

function createDeck(data) {
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });

  // Slide 1
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.navy;
    addShape(slide, { left: 0, top: 0, width: 1280, height: 720 }, C.navy, "none");
    addShape(slide, { left: 0, top: 606, width: 1280, height: 114 }, C.gold, "none");
    addText(slide, "PET FAIR ASIA 2026", { left: 72, top: 74, width: 520, height: 30 }, { fontSize: 18, bold: true, color: C.gold });
    addText(slide, "중국 반려동물 박람회\n업체별 미팅 정리", { left: 72, top: 132, width: 780, height: 150 }, { fontSize: 54, bold: true, color: C.white, lineSpacing: 0.96 });
    addText(slide, data.heroSubtitle || "미팅 및 소싱 가능한 업체만 대상으로 정리했습니다.", { left: 76, top: 310, width: 720, height: 40 }, { fontSize: 20, color: "#E7EDF6" });
    const stats = data.stats.slice(0, 4);
    stats.forEach((s, i) => {
      const x = 74 + i * 258;
      addShape(slide, { left: x, top: 430, width: 230, height: 92 }, "#2A3850", "#FFFFFF", "rounded-lg");
      addText(slide, s.value, { left: x + 18, top: 446, width: 190, height: 34 }, { fontSize: 29, bold: true, color: C.white });
      addText(slide, s.label, { left: x + 18, top: 488, width: 190, height: 22 }, { fontSize: 14, color: "#E7EDF6" });
    });
    addText(slide, "Source HTML · 2026.08.24", { left: 74, top: 638, width: 360, height: 28 }, { fontSize: 17, bold: true, color: C.navy });
    addText(slide, "메모 업체 중심 · 사진 2장 기준 · RMB 210원 환산", { left: 74, top: 666, width: 720, height: 24 }, { fontSize: 16, color: C.navy });
    addNotes(slide, [`원본 HTML: ${SOURCE_HTML}`, "원본 HTML의 내장 이미지와 텍스트를 기반으로 PPT로 재구성."]);
  }

  // Slide 2
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.bg;
    addTopBar(slide, "핵심 판단은 세 갈래로 압축됩니다", "저가 캔, 소용량 토핑, 해산물 원물형을 중심으로 후속 미팅을 진행합니다.", "02");
    data.priorities.slice(0, 4).forEach((p, i) => {
      const x = 72 + (i % 2) * 568;
      const y = 172 + Math.floor(i / 2) * 192;
      addShape(slide, { left: x, top: y, width: 520, height: 150 }, C.white, C.line, "rounded-lg");
      addText(slide, p.title, { left: x + 24, top: y + 22, width: 470, height: 32 }, { fontSize: 24, bold: true, color: i === 0 ? C.green : C.navy });
      addText(slide, shorten(p.body, 160), { left: x + 24, top: y + 62, width: 460, height: 70 }, { fontSize: 17, color: C.slate, lineSpacing: 1.18 });
    });
    addShape(slide, { left: 72, top: 588, width: 1080, height: 62 }, C.goldSoft, "#E9C74F", "rounded-lg");
    addText(slide, "미팅 방향", { left: 94, top: 606, width: 116, height: 24 }, { fontSize: 18, bold: true, color: C.orange });
    addText(slide, "1순위 업체는 견적·수출서류·샘플을 바로 요청하고, 보류 업체는 제품 확인용으로만 유지합니다.", { left: 214, top: 606, width: 900, height: 24 }, { fontSize: 18, bold: true, color: C.ink });
    addNotes(slide, [`원본 HTML: ${SOURCE_HTML}`, "핵심 판단 및 우선순위 카드에서 발췌."]);
  }

  // Slide 3
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.white;
    addTopBar(slide, "업체별 역할을 먼저 보고 들어갑니다", "업체별 강점이 겹치므로 미팅 질문은 품목별로 좁혀 진행하는 편이 효율적입니다.", "03");
    const cols = [
      { label: "업체", x: 72, w: 250 },
      { label: "역할", x: 322, w: 160 },
      { label: "가능 품목", x: 482, w: 360 },
      { label: "미팅 포인트", x: 842, w: 342 },
    ];
    cols.forEach((c) => {
      addShape(slide, { left: c.x, top: 158, width: c.w, height: 42 }, C.navy, C.navy);
      addText(slide, c.label, { left: c.x + 12, top: 168, width: c.w - 24, height: 20 }, { fontSize: 16, bold: true, color: C.white });
    });
    data.vendors.slice(0, 7).forEach((v, i) => {
      const y = 204 + i * 58;
      const fill = i % 2 ? "#FAFCFF" : C.white;
      cols.forEach((c) => addShape(slide, { left: c.x, top: y, width: c.w, height: 54 }, fill, C.line));
      addText(slide, shorten(v.title, 32), { left: 84, top: y + 9, width: 226, height: 34 }, { fontSize: 15, bold: true, color: C.ink });
      addText(slide, vendorRole(v), { left: 334, top: y + 14, width: 136, height: 24 }, { fontSize: 15, color: C.green, bold: true });
      addText(slide, shorten(v.info["가능 품목"] || v.headline, 64), { left: 494, top: y + 8, width: 336, height: 38 }, { fontSize: 14, color: C.slate });
      const point = v.actions[0] || v.risks[0] || v.headline;
      addText(slide, shorten(point, 58), { left: 854, top: y + 8, width: 318, height: 38 }, { fontSize: 14, color: C.slate });
    });
    addNotes(slide, [`원본 HTML: ${SOURCE_HTML}`, "업체별 제목, 헤드라인, 가능 품목, 리스크/다음 액션에서 요약."]);
  }

  // Vendor slides
  data.vendors.slice(0, 7).forEach((v, index) => {
    const slide = presentation.slides.add();
    slide.background.fill = C.bg;
    const page = String(index + 4).padStart(2, "0");
    addTopBar(slide, v.title, shorten(v.headline, 110), page);
    let tagX = 68;
    v.tags.slice(0, 3).forEach((t, i) => {
      addTag(slide, t, tagX, 137, i === 0 ? C.green : C.blue, i === 0 ? C.greenSoft : C.blueSoft);
      tagX += Math.min(170, 56 + t.length * 12) + 8;
    });

    const photos = v.photos.slice(0, 2);
    const photoY = 162;
    if (photos[0]) addPhoto(slide, photos[0], { left: 775, top: photoY, width: 196, height: 210 }, 1);
    if (photos[1]) addPhoto(slide, photos[1], { left: 990, top: photoY, width: 196, height: 210 }, 2);

    addText(slide, "기본 정보", { left: 70, top: 172, width: 180, height: 26 }, { fontSize: 20, bold: true, color: C.navy });
    const infoRows = [
      { label: "상호명", value: firstAvailable(v.info, ["상호명"], 82) },
      { label: "주소", value: firstAvailable(v.info, ["주소"], 90) },
      { label: "홈페이지", value: firstAvailable(v.info, ["홈페이지"], 90) },
      { label: "한국수출", value: firstAvailable(v.info, ["한국 수출 경험", "한국 수출"], 88) },
      { label: "가능품목", value: firstAvailable(v.info, ["가능 품목"], 96) },
      { label: "강점", value: firstAvailable(v.info, ["강점"], 100) },
    ];
    addInfoRows(slide, infoRows, 70, 208, 672, 40);

    addText(slide, "견적 요약", { left: 70, top: 474, width: 180, height: 26 }, { fontSize: 20, bold: true, color: C.navy });
    const quoteItems = v.quotes.slice(0, 3).map((q) => `${q[0] || "품목"}: ${q[1] || ""}${q[2] ? ` / ${q[2]}` : ""}`);
    addShape(slide, { left: 70, top: 508, width: 672, height: 92 }, C.white, C.line, "rounded-lg");
    addBulletText(slide, quoteItems.length ? quoteItems : ["정식 견적서 재확인 필요"], { left: 90, top: 526, width: 630, height: 56 }, { fontSize: 15, max: 100, color: C.slate });

    addShape(slide, { left: 775, top: 398, width: 411, height: 202 }, C.white, C.line, "rounded-lg");
    addText(slide, "리스크 / 다음 액션", { left: 796, top: 418, width: 250, height: 26 }, { fontSize: 20, bold: true, color: C.navy });
    const actions = [...v.actions.slice(0, 2), ...v.risks.slice(0, 1)];
    addBulletText(slide, actions, { left: 798, top: 458, width: 360, height: 106 }, { fontSize: 15, max: 92, color: C.slate });

    if (v.memo) addText(slide, shorten(v.memo, 78), { left: 70, top: 630, width: 760, height: 22 }, { fontSize: 12, color: C.muted });
    addNotes(slide, [
      `원본 HTML: ${SOURCE_HTML}`,
      `업체: ${v.title}`,
      ...photos.map((p) => `이미지: ${p.file || p.caption || p.alt}`),
      v.memo ? `메모: ${v.memo}` : "메모: 원본 HTML 내 업체 섹션",
    ]);
  });

  // Quote comparison slide
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.white;
    addTopBar(slide, "견적은 정식 견적서 기준으로 다시 맞춥니다", "RMB 환산가는 비교용입니다. 포장비, MOQ, 수출서류 발행 주체는 별도 확인이 필요합니다.", "11");
    const quoteRows = [];
    for (const v of data.vendors) {
      for (const q of v.quotes.slice(0, 2)) {
        quoteRows.push({
          vendor: v.title,
          item: q[0] || "",
          price: q[1] || "",
          moq: q[2] || "",
          memo: q[3] || "",
        });
      }
    }
    const rows = quoteRows.slice(0, 8);
    const headers = [
      { label: "업체", x: 72, w: 230 },
      { label: "품목", x: 302, w: 240 },
      { label: "견적", x: 542, w: 270 },
      { label: "MOQ/메모", x: 812, w: 372 },
    ];
    headers.forEach((h) => {
      addShape(slide, { left: h.x, top: 158, width: h.w, height: 38 }, C.navy, C.navy);
      addText(slide, h.label, { left: h.x + 10, top: 167, width: h.w - 20, height: 18 }, { fontSize: 15, bold: true, color: C.white });
    });
    rows.forEach((r, i) => {
      const y = 200 + i * 48;
      const fill = i % 2 ? "#FAFCFF" : C.white;
      headers.forEach((h) => addShape(slide, { left: h.x, top: y, width: h.w, height: 46 }, fill, C.line));
      addText(slide, shorten(r.vendor, 26), { left: 84, top: y + 11, width: 206, height: 22 }, { fontSize: 15, bold: true, color: C.ink });
      addText(slide, shorten(r.item, 28), { left: 314, top: y + 11, width: 216, height: 22 }, { fontSize: 15, color: C.slate });
      addText(slide, shorten(r.price, 34), { left: 554, top: y + 11, width: 246, height: 22 }, { fontSize: 15, color: C.green, bold: true });
      addText(slide, shorten(`${r.moq}${r.memo ? ` / ${r.memo}` : ""}`, 52), { left: 824, top: y + 11, width: 348, height: 22 }, { fontSize: 15, color: C.slate });
    });
    addShape(slide, { left: 72, top: 634, width: 1112, height: 48 }, C.goldSoft, "#E9C74F", "rounded-lg");
    addText(slide, "체크 기준: FOB/CIF 조건, 포장비 포함 여부, MOQ 단위, 한국 수출 서류, 샘플 리드타임을 동일 포맷으로 회수합니다.", { left: 94, top: 649, width: 1060, height: 22 }, { fontSize: 16, bold: true, color: C.ink });
    addNotes(slide, [`원본 HTML: ${SOURCE_HTML}`, "업체별 견적 요약 표에서 발췌. 환산 기준은 원본 HTML의 1 RMB = 210원 기준."]);
  }

  // Final slide
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.bg;
    addTopBar(slide, data.footerTitle || "추가 체크 항목", "킥오프/후속 미팅에서는 아래 항목을 먼저 확정하면 됩니다.", "12");
    const leftItems = data.footerItems.slice(0, Math.ceil(data.footerItems.length / 2));
    const rightItems = data.footerItems.slice(Math.ceil(data.footerItems.length / 2));
    addShape(slide, { left: 78, top: 178, width: 520, height: 384 }, C.white, C.line, "rounded-lg");
    addText(slide, "미팅 전 확인", { left: 104, top: 204, width: 220, height: 28 }, { fontSize: 24, bold: true, color: C.navy });
    addBulletText(slide, leftItems, { left: 108, top: 254, width: 440, height: 250 }, { fontSize: 18, max: 100, color: C.slate });
    addShape(slide, { left: 676, top: 178, width: 520, height: 384 }, C.white, C.line, "rounded-lg");
    addText(slide, "후속 요청", { left: 702, top: 204, width: 220, height: 28 }, { fontSize: 24, bold: true, color: C.navy });
    addBulletText(slide, rightItems.length ? rightItems : ["정식 견적서 및 샘플 요청서를 동일 양식으로 회수"], { left: 706, top: 254, width: 440, height: 250 }, { fontSize: 18, max: 100, color: C.slate });
    addShape(slide, { left: 78, top: 596, width: 1118, height: 58 }, C.navy, C.navy, "rounded-lg");
    addText(slide, "결론: 1순위 업체는 캔·동결건조를 동시에 견적화하고, 나머지는 품목별 후보로 분리해 샘플 검증합니다.", { left: 106, top: 614, width: 1060, height: 24 }, { fontSize: 18, bold: true, color: C.white });
    addNotes(slide, [`원본 HTML: ${SOURCE_HTML}`, "추가 체크 항목 및 전체 업체 섹션에서 요약."]);
  }

  return presentation;
}

async function main() {
  await fs.mkdir(RENDER_DIR, { recursive: true });
  const html = await fs.readFile(SOURCE_HTML, "utf8");
  const data = extractHtmlData(html);
  const presentation = createDeck(data);

  for (const [i, slide] of presentation.slides.items.entries()) {
    const png = await presentation.export({ slide, format: "png", scale: 1 });
    await writeBlob(path.join(RENDER_DIR, `slide-${String(i + 1).padStart(2, "0")}.png`), png);
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(RENDER_DIR, `slide-${String(i + 1).padStart(2, "0")}.layout.json`), await layout.text(), "utf8");
  }

  const montage = await presentation.export({ format: "webp", montage: true, scale: 1 });
  await writeBlob(path.join(TMP_DIR, "deck-montage.webp"), montage);

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL_PPTX);

  const sourceNotes = [
    "PET FAIR ASIA 2026 PPT conversion source notes",
    `Source HTML: ${SOURCE_HTML}`,
    `Final PPTX: ${FINAL_PPTX}`,
    "All visible vendor copy and embedded photos were extracted from the source HTML.",
    "No external web research was used for this conversion.",
  ].join("\n");
  await fs.writeFile(path.join(TMP_DIR, "source-notes.txt"), sourceNotes, "utf8");
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
