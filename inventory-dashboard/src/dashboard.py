from __future__ import annotations

from typing import Any


def quoted_sheet_name(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def build_dashboard_values(layout: Any, last_success: str, result_text: str, unmapped_count: int) -> list[list[Any]]:
    month = quoted_sheet_name(layout.sheet_name)
    start = layout.data_start_row
    end = layout.data_end_row
    naver_col = layout.naver_start_col_letter
    error_col = layout.error_col_letter

    values: list[list[Any]] = [
        ["OSP 재고 자동 동기화 Dashboard"],
        ["마지막 성공 업데이트 일시", last_success],
        ["마지막 실행 결과", result_text],
        [],
        ["지표", "값"],
        ["총 재고합계 낱개", f"=SUM({month}!D{start}:D{end})"],
        ["총 하은재고 낱개", f"=SUM({month}!G{start}:G{end})"],
        ["총 논산재고 낱개", f"=SUM({month}!J{start}:J{end})"],
        ["총 네이버재고 낱개", f"=SUM({month}!{naver_col}{start}:{naver_col}{end})"],
        ["품절 SKU 수", f'=COUNTIF({month}!D{start}:D{end},"<=0")'],
        ["30개 이하 SKU 수", f'=COUNTIFS({month}!D{start}:D{end},">0",{month}!D{start}:D{end},"<=30")'],
        ["네이버 품절 SKU 수", f'=COUNTIF({month}!{naver_col}{start}:{naver_col}{end},"<=0")'],
        ["미매핑 SKU 수", unmapped_count],
        [],
        [],
        ["코드", "상품명", "재고합계_낱개", "하은재고_낱개", "논산재고_낱개", "네이버재고_낱개", "전체재고상태", "네이버재고상태", "오류메시지"],
    ]

    dashboard_row = 17
    for sheet_row in range(start, end + 1):
        values.append(
            [
                f"={month}!B{sheet_row}",
                f"={month}!C{sheet_row}",
                f"={month}!D{sheet_row}",
                f"={month}!G{sheet_row}",
                f"={month}!J{sheet_row}",
                f"={month}!{naver_col}{sheet_row}",
                f'=IF(C{dashboard_row}<=0,"품절",IF(C{dashboard_row}<=30,"주의","정상"))',
                f'=IF(ISBLANK(F{dashboard_row}),"미조회",IF(F{dashboard_row}<=0,"네이버품절",IF(F{dashboard_row}<=30,"네이버주의","네이버정상")))',
                f"={month}!{error_col}{sheet_row}",
            ]
        )
        dashboard_row += 1
    return values


def build_dashboard_format_requests(sheet_id: int, row_count: int) -> list[dict[str, Any]]:
    table_start = 15
    table_end = table_start + row_count + 1
    requests: list[dict[str, Any]] = [
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 9},
                "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 16}}},
                "fields": "userEnteredFormat.textFormat",
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 4, "endRowIndex": 5, "startColumnIndex": 0, "endColumnIndex": 2},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.1, "green": 0.32, "blue": 0.58},
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": table_start, "endRowIndex": table_start + 1, "startColumnIndex": 0, "endColumnIndex": 9},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.1, "green": 0.32, "blue": 0.58},
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        },
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 16}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
        {
            "autoResizeDimensions": {
                "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 9}
            }
        },
    ]

    for text, color in [
        ("품절", {"red": 0.96, "green": 0.8, "blue": 0.8}),
        ("네이버품절", {"red": 0.96, "green": 0.8, "blue": 0.8}),
        ("주의", {"red": 1.0, "green": 0.93, "blue": 0.55}),
        ("네이버주의", {"red": 1.0, "green": 0.93, "blue": 0.55}),
        ("미조회", {"red": 0.86, "green": 0.86, "blue": 0.86}),
    ]:
        requests.append(
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [
                            {"sheetId": sheet_id, "startRowIndex": table_start + 1, "endRowIndex": table_end, "startColumnIndex": 6, "endColumnIndex": 8}
                        ],
                        "booleanRule": {
                            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": text}]},
                            "format": {"backgroundColor": color},
                        },
                    },
                    "index": 0,
                }
            }
        )

    requests.append(
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {"sheetId": sheet_id, "startRowIndex": table_start + 1, "endRowIndex": table_end, "startColumnIndex": 8, "endColumnIndex": 9}
                    ],
                    "booleanRule": {
                        "condition": {"type": "TEXT_CONTAINS", "values": [{"userEnteredValue": "미매핑"}]},
                        "format": {"backgroundColor": {"red": 0.88, "green": 0.8, "blue": 0.96}},
                    },
                },
                "index": 0,
            }
        }
    )
    return requests
