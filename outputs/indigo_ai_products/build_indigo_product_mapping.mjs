import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/Users/Administrator/Documents/New project/outputs/indigo_ai_products";
const outputPath = `${outputDir}/인디고_증상별_제품_원료_매칭표.xlsx`;

const today = "2026-07-22";

const categoryRows = [
  ["강아지", "관절", "해당", "인디고 바이오뉴트리션 독 관절", "2kg / 6kg / 정기구독", "관절강화, 저알러지, 면역력강화, 소화개선, 식욕증진", "연어, 오리", "관절 건강 관련 기능 원료/주원료 확인 필요: 연어, 오리, 유산균 사료", "https://brand.naver.com/indigo/products/7292695319", "제품 상세 상품정보에 기능 '관절강화' 확인"],
  ["강아지", "관절", "해당", "인디고 에이지뉴트리션 독 시니어10+", "6kg", "관절강화, 항산화, 면역력강화, 피모관리, 소화개선", "연어, 오리", "시니어 관절/항산화/면역 케어 제품으로 검토 가능", "https://brand.naver.com/indigo/products/12089579003", "제품 상세 상품정보에 기능 '관절강화' 확인"],
  ["강아지", "모질", "해당", "인디고 바이오뉴트리션 독 스킨", "2kg / 6kg / 정기구독", "저알러지, 면역력강화, 소화개선, 식욕증진, 피모관리", "연어", "피부+피모 관련 제품으로 검토 가능", "https://brand.naver.com/indigo/products/7292810813", "제품 상세 상품정보에 기능 '피모관리' 확인"],
  ["강아지", "모질", "해당", "인디고 펫 오리진 프로바이오틱스 독", "30포", "피모관리, 면역력강화, 구강관리, 항산화, 저알러지", "상세 확인 필요", "프로바이오틱스/유산균 기반 피모관리 제품", "https://brand.naver.com/indigo/products/10973942717", "제품 상세 상품정보에 기능 '피모관리' 확인"],
  ["강아지", "소화기", "해당", "인디고 펫 오리진 프로바이오틱스 독", "30포", "피모관리, 면역력강화, 구강관리, 항산화, 저알러지", "상세 확인 필요", "장건강/유산균/프로바이오틱스 소구 가능", "https://brand.naver.com/indigo/products/10973942717", "제품명과 태그에 장건강/유산균 확인"],
  ["강아지", "소화기", "해당", "인디고 세븐 강아지 유기농 사료", "1.6kg / 5.2kg", "면역력강화, 소화개선, 식욕증진, 영양공급, 저알러지", "연어+닭고기 / 소고기+오리", "소화개선 기능 표기 제품", "https://brand.naver.com/indigo/products/9907973724", "제품 상세 상품정보에 기능 '소화개선' 확인"],
  ["강아지", "눈", "해당", "인디고 에이지뉴트리션 독 어덜트", "1.2kg", "소화개선, 면역력강화, 항산화, 눈건강, 피모관리", "오리, 소", "눈건강 기능 표기 제품", "https://brand.naver.com/indigo/products/11395025645", "제품 상세 상품정보에 기능 '눈건강' 확인"],
  ["강아지", "눈물", "검토", "인디고 하이포알러제닉 독", "1kg", "알러지케어", "곤충", "저알러지/가수분해/밀웜 관련 제품. 단, 눈물 개선 표현은 사용 금지", "https://brand.naver.com/indigo/products/11217916820", "스토어 검색어 '눈물'에서 샘플/체험팩 위주 노출, 본품은 알러지케어 제품으로 연결 검토"],
  ["강아지", "체중", "해당", "인디고 바이오뉴트리션 독 체중", "2kg / 6kg / 정기구독", "다이어트", "오리", "체중/다이어트 관련 제품으로 검토 가능", "https://brand.naver.com/indigo/products/10536943410", "제품명과 상세 상품정보에 다이어트 확인"],
  ["강아지", "치아", "검토", "인디고 펫 오리진 프로바이오틱스 독", "30포", "피모관리, 면역력강화, 구강관리, 항산화, 저알러지", "상세 확인 필요", "치아보다는 구강관리 표현으로 순화 권장", "https://brand.naver.com/indigo/products/10973942717", "제품 상세 상품정보에 기능 '구강관리' 확인"],
  ["강아지", "피부", "해당", "인디고 바이오뉴트리션 독 스킨", "2kg / 6kg / 정기구독", "저알러지, 면역력강화, 소화개선, 식욕증진, 피모관리", "연어", "피부+피모 관련 제품으로 검토 가능", "https://brand.naver.com/indigo/products/7292810813", "제품명에 피부+피모, 상세 기능 피모관리 확인"],
  ["강아지", "피부", "검토", "인디고 하이포알러제닉 독", "1kg", "알러지케어", "곤충", "저알러지/가수분해/밀웜 기반 알러지케어 제품", "https://brand.naver.com/indigo/products/11217916820", "피부 개선 표현은 금지, 알러지 고민 제품으로만 표현"],
  ["강아지", "신장", "확인필요", "직접 매칭 제품 없음", "", "", "", "", "", "스토어 검색/상품정보 기준 강아지 신장 직접 매칭 확인 안 됨"],
  ["강아지", "귀", "확인필요", "직접 매칭 제품 없음", "", "", "", "", "", "스토어 검색/상품정보 기준 귀 직접 매칭 확인 안 됨"],
  ["강아지", "심장", "확인필요", "직접 매칭 제품 없음", "", "", "", "", "", "스토어 검색/상품정보 기준 심장 직접 매칭 확인 안 됨"],
  ["강아지", "호흡기", "확인필요", "직접 매칭 제품 없음", "", "", "", "", "", "스토어 검색/상품정보 기준 호흡기 직접 매칭 확인 안 됨"],
  ["고양이", "음수량", "해당", "인디고 펫 오리진 데일리스틱", "56g 2종 콤보", "음수량증진, 식욕증진, 영양공급, 소화개선, 체중유지", "2종 콤보", "짜먹는 간식/수분 섭취 보조 방향으로 검토 가능", "https://brand.naver.com/indigo/products/13580919806", "제품 상세 상품정보에 기능 '음수량증진' 확인"],
  ["고양이", "음수량", "검토", "인디고 캣 흰살참치 그레이비 습식 캔", "80g 6종 콤보", "헤어볼, 관절강화, 눈건강, 체중유지, 피모관리", "6종 콤보", "습식 캔이므로 음수량 카테고리 보조 노출 검토 가능", "https://brand.naver.com/indigo/products/11325635644", "상품유형이 습식 캔. 기능에는 음수량 직접 표기 없음"],
  ["고양이", "모질", "해당", "인디고 펫 오리진 프로바이오틱스 캣", "30포", "면역력강화, 구강관리, 항산화, 피모관리, 저알러지", "상세 확인 필요", "프로바이오틱스/유산균 기반 피모관리 제품", "https://brand.naver.com/indigo/products/11117808507", "제품 상세 상품정보에 기능 '피모관리' 확인"],
  ["고양이", "모질", "해당", "인디고 캣 흰살참치 그레이비 습식 캔", "80g 6종 콤보", "헤어볼, 관절강화, 눈건강, 체중유지, 피모관리", "6종 콤보", "피모관리 기능 표기 제품", "https://brand.naver.com/indigo/products/11325635644", "제품 상세 상품정보에 기능 '피모관리' 확인"],
  ["고양이", "헤어볼", "해당", "인디고 바이오뉴트리션 캣 헤어볼", "2kg / 6kg / 정기구독", "헤어볼 케어", "닭, 연어", "헤어볼 케어 제품", "https://brand.naver.com/indigo/products/10537646485", "제품명/상세 기능에 헤어볼 케어 확인"],
  ["고양이", "헤어볼", "해당", "인디고 캣 흰살참치 그레이비 습식 캔", "80g 6종 콤보", "헤어볼, 관절강화, 눈건강, 체중유지, 피모관리", "6종 콤보", "헤어볼 기능 표기 제품", "https://brand.naver.com/indigo/products/11325635644", "제품 상세 상품정보에 기능 '헤어볼' 확인"],
  ["고양이", "변비", "검토", "인디고 펫 오리진 프로바이오틱스 캣", "30포", "면역력강화, 구강관리, 항산화, 피모관리, 저알러지", "상세 확인 필요", "장건강/유산균 제품으로 검토. 변비 개선 표현은 금지", "https://brand.naver.com/indigo/products/11117808507", "제품명과 태그에 장건강/유산균 확인"],
  ["고양이", "변비", "검토", "인디고 바이오뉴트리션 키튼", "2kg", "면역+장건강", "닭, 연어", "장건강 카테고리로 검토 가능", "https://brand.naver.com/indigo/products/7293000421", "제품 상세 상품정보에 기능 '면역+장건강' 확인"],
  ["고양이", "설사", "검토", "인디고 펫 오리진 프로바이오틱스 캣", "30포", "면역력강화, 구강관리, 항산화, 피모관리, 저알러지", "상세 확인 필요", "장건강/유산균 제품으로 검토. 설사 개선 표현은 금지", "https://brand.naver.com/indigo/products/11117808507", "제품명과 태그에 장건강/유산균 확인"],
  ["고양이", "체중관리", "해당", "인디고 세븐 인도어 팻다운", "1.4kg / 5.2kg", "다이어트", "닭, 연어", "저지방/다이어트 제품", "https://brand.naver.com/indigo/products/9908006256", "제품명과 상세 기능에 다이어트 확인"],
  ["고양이", "체중관리", "해당", "인디고 펫 오리진 데일리스틱", "56g 2종 콤보", "음수량증진, 식욕증진, 영양공급, 소화개선, 체중유지", "2종 콤보", "체중유지 기능 표기 제품", "https://brand.naver.com/indigo/products/13580919806", "제품 상세 상품정보에 기능 '체중유지' 확인"],
  ["고양이", "치아", "검토", "인디고 펫 오리진 프로바이오틱스 캣", "30포", "면역력강화, 구강관리, 항산화, 피모관리, 저알러지", "상세 확인 필요", "치아보다는 구강관리 표현으로 순화 권장", "https://brand.naver.com/indigo/products/11117808507", "제품 상세 상품정보에 기능 '구강관리' 확인"],
  ["고양이", "피부", "해당", "인디고 펫 오리진 프로바이오틱스 캣", "30포", "면역력강화, 구강관리, 항산화, 피모관리, 저알러지", "상세 확인 필요", "피모관리/저알러지 기능 표기 제품", "https://brand.naver.com/indigo/products/11117808507", "제품 상세 상품정보에 기능 '피모관리, 저알러지' 확인"],
  ["고양이", "피부", "검토", "인디고 하이포알러제닉 캣", "1kg", "알러지케어", "곤충", "저알러지/가수분해/밀웜 기반 알러지케어 제품", "https://brand.naver.com/indigo/products/11217939656", "피부 개선 표현은 금지, 알러지 고민 제품으로만 표현"],
  ["고양이", "신장", "해당", "인디고 바이오뉴트리션 캣 유리너리", "2kg / 6kg / 정기구독", "유리너리(비뇨계), 종합비타민, 면역력강화, 소화개선, 신장/요로", "오리, 연어", "신장/요로 카테고리 제품", "https://brand.naver.com/indigo/products/7293042225", "제품 상세 상품정보에 기능 '신장/요로' 확인"],
  ["고양이", "관절", "검토", "인디고 캣 흰살참치 그레이비 습식 캔", "80g 6종 콤보", "헤어볼, 관절강화, 눈건강, 체중유지, 피모관리", "6종 콤보", "관절강화 기능 표기 제품. 간식/습식 캔 성격이라 보조 노출 권장", "https://brand.naver.com/indigo/products/11325635644", "제품 상세 상품정보에 기능 '관절강화' 확인"],
  ["고양이", "눈", "검토", "인디고 캣 흰살참치 그레이비 습식 캔", "80g 6종 콤보", "헤어볼, 관절강화, 눈건강, 체중유지, 피모관리", "6종 콤보", "눈건강 기능 표기 제품. 간식/습식 캔 성격이라 보조 노출 권장", "https://brand.naver.com/indigo/products/11325635644", "제품 상세 상품정보에 기능 '눈건강' 확인"],
  ["고양이", "심장", "확인필요", "직접 매칭 제품 없음", "", "", "", "", "", "스토어 검색/상품정보 기준 심장 직접 매칭 확인 안 됨"],
];

const productRows = [
  ["강아지", "바이오뉴트리션", "인디고 바이오뉴트리션 독 관절", "2kg/6kg", "관절강화, 저알러지, 면역력강화, 소화개선, 식욕증진", "연어, 오리", "관절, 소화기", "https://brand.naver.com/indigo/products/7292695319"],
  ["강아지", "바이오뉴트리션", "인디고 바이오뉴트리션 독 스킨", "2kg/6kg", "저알러지, 면역력강화, 소화개선, 식욕증진, 피모관리", "연어", "피부, 모질, 소화기", "https://brand.naver.com/indigo/products/7292810813"],
  ["강아지", "바이오뉴트리션", "인디고 바이오뉴트리션 독 체중", "2kg/6kg", "다이어트", "오리", "체중", "https://brand.naver.com/indigo/products/10536943410"],
  ["강아지", "하이포알러제닉", "인디고 하이포알러제닉 독", "1kg", "알러지케어", "곤충", "눈물, 피부", "https://brand.naver.com/indigo/products/11217916820"],
  ["강아지", "펫 오리진", "인디고 펫 오리진 프로바이오틱스 독", "30포", "피모관리, 면역력강화, 구강관리, 항산화, 저알러지", "상세 확인 필요", "소화기, 치아/구강, 피부, 모질", "https://brand.naver.com/indigo/products/10973942717"],
  ["강아지", "세븐", "인디고 세븐 강아지 유기농 사료 연어+닭고기", "1.6kg", "영양공급, 식욕증진, 면역력강화, 소화개선, 저알러지", "연어+닭고기", "소화기, 피부 검토", "https://brand.naver.com/indigo/products/9907973724"],
  ["강아지", "세븐", "인디고 세븐 강아지 유기농 사료 소고기+오리", "1.6kg/5.2kg", "면역력강화, 소화개선, 식욕증진, 영양공급, 저알러지", "소고기+오리", "소화기, 피부 검토", "https://brand.naver.com/indigo/products/9907995508"],
  ["강아지", "에이지뉴트리션", "인디고 에이지뉴트리션 독 어덜트", "1.2kg", "소화개선, 면역력강화, 항산화, 눈건강, 피모관리", "오리, 소", "눈, 모질, 피부, 소화기", "https://brand.naver.com/indigo/products/11395025645"],
  ["강아지", "에이지뉴트리션", "인디고 에이지뉴트리션 독 시니어10+", "6kg", "관절강화, 항산화, 면역력강화, 피모관리, 소화개선", "연어, 오리", "관절, 모질, 피부, 소화기", "https://brand.naver.com/indigo/products/12089579003"],
  ["고양이", "바이오뉴트리션", "인디고 바이오뉴트리션 캣 헤어볼", "2kg/6kg", "헤어볼 케어", "닭, 연어", "헤어볼", "https://brand.naver.com/indigo/products/10537646485"],
  ["고양이", "바이오뉴트리션", "인디고 바이오뉴트리션 캣 유리너리", "2kg/6kg", "유리너리(비뇨계), 종합비타민, 면역력강화, 소화개선, 신장/요로", "오리, 연어", "신장, 변비/설사 검토", "https://brand.naver.com/indigo/products/7293042225"],
  ["고양이", "바이오뉴트리션", "인디고 바이오뉴트리션 키튼", "2kg", "면역+장건강", "닭, 연어", "변비/설사 검토", "https://brand.naver.com/indigo/products/7293000421"],
  ["고양이", "하이포알러제닉", "인디고 하이포알러제닉 캣", "1kg", "알러지케어", "곤충", "피부", "https://brand.naver.com/indigo/products/11217939656"],
  ["고양이", "펫 오리진", "인디고 펫 오리진 프로바이오틱스 캣", "30포", "면역력강화, 구강관리, 항산화, 피모관리, 저알러지", "상세 확인 필요", "변비/설사, 치아/구강, 피부, 모질", "https://brand.naver.com/indigo/products/11117808507"],
  ["고양이", "펫 오리진", "인디고 펫 오리진 데일리스틱", "56g", "음수량증진, 식욕증진, 영양공급, 소화개선, 체중유지", "2종 콤보", "음수량, 체중관리, 변비/설사 검토", "https://brand.naver.com/indigo/products/13580919806"],
  ["고양이", "캣 습식", "인디고 캣 흰살참치 그레이비 습식 캔", "80g", "헤어볼, 관절강화, 눈건강, 체중유지, 피모관리", "6종 콤보", "음수량 검토, 헤어볼, 관절, 눈, 체중관리, 모질", "https://brand.naver.com/indigo/products/11325635644"],
  ["고양이", "세븐", "인디고 세븐 인도어 프로틴업", "1.4kg/5.2kg", "고단백", "연어, 닭", "직접 매칭 낮음", "https://brand.naver.com/indigo/products/9908012931"],
  ["고양이", "세븐", "인디고 세븐 인도어 팻다운", "1.4kg/5.2kg", "다이어트", "닭, 연어", "체중관리", "https://brand.naver.com/indigo/products/9908006256"],
];

const searchRows = [
  ["관절", "19", "인디고 바이오뉴트리션 독 관절, 에이지뉴트리션 독 시니어 등"],
  ["모질", "2", "세븐 강아지 연어&닭고기+유산균 세트 등. 기능표기 기준으로는 피모관리 제품 우선"],
  ["눈물", "3", "샘플/체험팩 위주 노출. 본품은 하이포알러제닉 독 알러지케어로 검토"],
  ["체중", "22", "바이오뉴트리션 독 체중, 세븐 팻다운 등"],
  ["다이어트", "13", "바이오뉴트리션 독 체중, 세븐 인도어 팻다운 등"],
  ["피부", "13", "바이오뉴트리션 독 스킨, 하이포알러제닉, 프로바이오틱스 등"],
  ["스킨", "9", "바이오뉴트리션 독 스킨 등"],
  ["유리너리", "11", "바이오뉴트리션 캣 유리너리"],
  ["헤어볼", "34", "바이오뉴트리션 캣 헤어볼, 캣 습식 캔 등"],
  ["변비", "17", "스킨/하이포알러제닉/유리너리/헤어볼/프로바이오틱스 관련 상품 노출. 개선 표현 금지"],
  ["프로바이오틱스", "6", "펫 오리진 프로바이오틱스 독/캣"],
  ["유산균", "118", "다수 상품 공통 노출. 단독 근거보다 제품별 기능표기와 함께 사용 권장"],
  ["밀웜", "10", "하이포알러제닉 독/캣"],
  ["가수분해", "74", "하이포알러제닉, 세븐, 바이오뉴트리션 등 다수"],
  ["소화/소화기/눈/치아/덴탈/신장/귀/심장/호흡기/설사/음수량", "검색 총 개수 미확인", "네이버가 결과 없음 상태에서도 기본 추천상품을 노출하여 검색 결과 근거로 사용하지 않음"],
  ["글루코사민/칼슘/아마씨/어분/완두콩/타피오카/염화콜린", "검색 총 개수 미확인", "첨부 이미지 예시 원료와 직접 일치하는 스토어 검색 결과는 확인 어려움. 상세 원재료표 별도 확인 필요"],
];

const notes = [
  ["구분", "내용"],
  ["작성 기준", `네이버 브랜드스토어 인디고 상품 상세 및 스토어 내부 검색 확인일: ${today}`],
  ["표현 주의", "'질병 예방 원료'는 치료/예방 보장처럼 보일 수 있어 대외 노출 문구는 '건강관리 관련 원료', '기능성 원료', '관련 기능/주원료'로 순화 권장"],
  ["금지 권장", "관절염 예방, 알러지 치료, 눈물 개선, 변비/설사 개선, 신장질환 예방, 심장질환 예방 등 질병명+효과 단정 표현은 사용하지 않는 방향 권장"],
  ["매칭 기준", "상품 상세의 기능/주원료/관련 태그와 스토어 내부 검색 결과가 확인되는 제품만 '해당' 또는 '검토'로 분류"],
  ["중복 처리", "정기구독, 묶음, 샘플, 증정 구성은 동일 포뮬러로 보고 대표 제품 중심으로 정리"],
  ["추가 확인 필요", "첨부 이미지 예시 원료(글루코사민, 칼슘, 아마씨, 어분, 완두콩 식이섬유, 타피오카 전분, 염화콜린 등)는 상세 원재료표 이미지/OCR 또는 제조사 원료표로 별도 확인 필요"],
];

function rangeFor(sheet, rowCount, colCount, startRow = 0, startCol = 0) {
  return sheet.getRangeByIndexes(startRow, startCol, rowCount, colCount);
}

function writeSheet(sheet, title, headers, rows, widths) {
  sheet.showGridLines = false;
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    font: { bold: true, size: 14, color: "#1F2937" },
  };
  const headerRow = 2;
  rangeFor(sheet, 1, headers.length, headerRow - 1, 0).values = [headers];
  rangeFor(sheet, 1, headers.length, headerRow - 1, 0).format = {
    fill: "#263238",
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#FFFFFF" },
  };
  if (rows.length) {
    const dataRange = rangeFor(sheet, rows.length, headers.length, headerRow, 0);
    dataRange.values = rows;
    dataRange.format = {
      wrapText: true,
      borders: { preset: "all", style: "thin", color: "#D9E1E8" },
    };
  }
  const used = rangeFor(sheet, rows.length + 1, headers.length, headerRow - 1, 0);
  used.format.font = { name: "맑은 고딕", size: 10 };
  for (let i = 0; i < widths.length; i++) {
    sheet.getRangeByIndexes(0, i, Math.max(rows.length + 3, 3), 1).format.columnWidth = widths[i];
  }
  sheet.freezePanes.freezeRows(2);
}

await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const summary = workbook.worksheets.add("요약");
summary.showGridLines = false;
summary.getRange("A1").values = [["인디고 증상별 제품/원료 매칭표"]];
summary.getRange("A1").format = { font: { bold: true, size: 16, color: "#1F2937" } };
summary.getRange("A3:B9").values = notes;
summary.getRange("A3:B3").format = { fill: "#263238", font: { bold: true, color: "#FFFFFF" } };
summary.getRange("A4:A9").format = { fill: "#E8F4F8", font: { bold: true } };
summary.getRange("A3:B9").format = { wrapText: true, borders: { preset: "all", style: "thin", color: "#D9E1E8" } };
summary.getRange("A:A").format.columnWidth = 18;
summary.getRange("B:B").format.columnWidth = 110;

writeSheet(
  workbook.worksheets.add("카테고리별 매칭"),
  "카테고리별 제품 매칭",
  ["대상", "카테고리", "판정", "대표 제품", "용량/구성", "확인된 기능/태그", "확인된 주원료", "노출 후보 표현", "제품 URL", "근거/비고"],
  categoryRows,
  [10, 12, 12, 34, 18, 42, 24, 48, 55, 50],
);

writeSheet(
  workbook.worksheets.add("제품별 상세"),
  "대표 제품별 상세",
  ["대상", "라인", "대표 제품", "용량/구성", "확인된 기능", "확인된 주원료", "매칭 카테고리", "제품 URL"],
  productRows,
  [10, 18, 45, 18, 48, 28, 35, 55],
);

writeSheet(
  workbook.worksheets.add("검색/원료 확인"),
  "스토어 내부 검색 및 원료 확인 메모",
  ["검색어/원료", "검색 결과", "메모"],
  searchRows,
  [28, 20, 90],
);

const categorySheet = workbook.worksheets.getItem("카테고리별 매칭");
categorySheet.getRange("C3:C200").conditionalFormats.add("containsText", {
  text: "해당",
  format: { fill: "#DFF6DD", font: { color: "#166534", bold: true } },
});
categorySheet.getRange("C3:C200").conditionalFormats.add("containsText", {
  text: "검토",
  format: { fill: "#FFF4CE", font: { color: "#92400E", bold: true } },
});
categorySheet.getRange("C3:C200").conditionalFormats.add("containsText", {
  text: "확인필요",
  format: { fill: "#FDE2E1", font: { color: "#991B1B", bold: true } },
});

const inspect = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 6000,
  tableMaxRows: 5,
  tableMaxCols: 8,
});
console.log(inspect.ndjson);

const preview = await workbook.render({
  sheetName: "카테고리별 매칭",
  range: "A1:J20",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/preview_category.png`, new Uint8Array(await preview.arrayBuffer()));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);

console.log(outputPath);
