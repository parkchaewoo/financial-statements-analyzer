"""PDF 리포트 공통 기반 모듈

모든 PDF 리포트 생성기가 상속하는 베이스 클래스.
공유 인프라: 폰트, 색상, 스타일, 포맷팅 헬퍼, 테이블 빌더.
"""

import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from risk_analyzer import (
    assess_metric,
    INDICATOR_DESC,
    LEVEL_OK,
    LEVEL_CAUTION,
    LEVEL_WARNING,
    LEVEL_DANGER,
)
from trend_analyzer import (
    SITUATION_LABELS,
    SITUATION_COLORS,
    SITUATION_EMOJI,
    TREND_UP,
    TREND_DOWN,
    TREND_FLAT,
    TREND_VOLATILE,
    TREND_RECOVERY,
)

# ── 한글 폰트 등록 ──────────────────────────────────────────

_FONT_PATHS = [
    # BeeWare 번들 폰트 (Android/모바일)
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources", "KoreanFont.ttf"),
    # macOS
    "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
    "/Library/Fonts/NanumGothic.ttf",
    "/Library/Fonts/NanumGothicBold.ttf",
    "/System/Library/Fonts/Supplemental/NotoSansGothic-Regular.ttf",
    # Linux
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    # Windows
    "C:/Windows/Fonts/malgun.ttf",
]

FONT_NAME = "Helvetica"  # 기본값 (폴백)

for fp in _FONT_PATHS:
    if os.path.exists(fp):
        try:
            pdfmetrics.registerFont(TTFont("KoreanFont", fp))
            FONT_NAME = "KoreanFont"
            break
        except Exception:
            continue

# ── 색상 & 스타일 상수 ──────────────────────────────────────

HEADER_BG = colors.HexColor("#2C3E50")
HEADER_FG = colors.white
SECTION_BG = colors.HexColor("#3498DB")
SECTION_FG = colors.white
SUB_HEADER_BG = colors.HexColor("#ECF0F1")
POSITIVE_COLOR = colors.HexColor("#2980B9")
NEGATIVE_COLOR = colors.HexColor("#E74C3C")
BORDER_COLOR = colors.HexColor("#BDC3C7")
LIGHT_ROW_BG = colors.HexColor("#F8F9FA")

# ── TTM 컬럼 색상 ─────────────────────────────────────────────
TTM_HEADER_BG = colors.HexColor("#27AE60")       # TTM 헤더 (녹색)
TTM_CELL_BG = colors.HexColor("#E8F8F5")         # TTM 셀 (연한 녹색)
TTM_RISK_BG = colors.HexColor("#E3F2FD")         # TTM 위험 항목 (연한 파란색)

# ── 추이 분석 색상 ─────────────────────────────────────────────
TREND_POSITIVE_BG = colors.HexColor("#E8F5E9")  # 긍정 (연한 녹색)
TREND_NEGATIVE_BG = colors.HexColor("#FFEBEE")  # 부정 (연한 빨강)
TREND_NEUTRAL_BG = colors.HexColor("#F5F5F5")   # 중립 (연한 회색)

# ── 위험 하이라이트 색상 ─────────────────────────────────────

CAUTION_BG = colors.HexColor("#FFF9C4")   # 연노랑 - 주의
WARNING_BG = colors.HexColor("#FFE0B2")   # 연주황 - 위험
DANGER_BG = colors.HexColor("#FFCDD2")    # 연빨강 - 심각

RISK_LEVEL_BG = {
    LEVEL_CAUTION: CAUTION_BG,
    LEVEL_WARNING: WARNING_BG,
    LEVEL_DANGER: DANGER_BG,
}

# 위험 패널 라벨 배경
RISK_LABEL_BG = {
    LEVEL_CAUTION: colors.HexColor("#F9A825"),
    LEVEL_WARNING: colors.HexColor("#EF6C00"),
    LEVEL_DANGER: colors.HexColor("#C62828"),
}


# ── 모듈 레벨 포맷팅 함수 ──────────────────────────────────────

def _fmt_num(value, unit="eok", decimals=0) -> str:
    """숫자 포맷팅"""
    if value is None or value == 0:
        return "-"
    if unit == "eok":
        v = value / 100_000_000
        if abs(v) >= 1:
            return f"{v:,.{decimals}f}"
        return f"{v:,.1f}"
    elif unit == "pct":
        return f"{value:.{max(decimals, 1)}f}"
    elif unit == "ratio":
        return f"{value:.2f}"
    elif unit == "won":
        return f"{int(value):,}"
    elif unit == "raw":
        if isinstance(value, float):
            return f"{value:.{decimals}f}"
        return f"{value:,}"
    return str(value)


def _fmt_eok(value) -> str:
    """원 → 억원 변환 포맷팅"""
    if value is None or value == 0:
        return "-"
    v = value / 100_000_000
    if abs(v) >= 10:
        return f"{v:,.0f}"
    return f"{v:,.1f}"


def _fmt_amount(value, currency="KRW") -> str:
    """통화별 금액 포맷팅 - KRW: 억원, USD: $M/$B"""
    if value is None or value == 0:
        return "-"
    if currency == "KRW":
        return _fmt_eok(value)
    abs_v = abs(value)
    if abs_v >= 1e9:
        v = value / 1e9
        return f"${v:,.1f}B"
    elif abs_v >= 1e6:
        v = value / 1e6
        return f"${v:,.1f}M"
    elif abs_v >= 1e3:
        v = value / 1e3
        return f"${v:,.1f}K"
    return f"${value:,.0f}"


def _fmt_price(value, currency="KRW") -> str:
    """통화별 주가 포맷팅 - KRW: 정수, USD: $소수점2자리"""
    if value is None or value == 0:
        return "-"
    if currency == "KRW":
        return f"{int(value):,}"
    return f"${value:,.2f}"


def _find(d: dict, *keys: str) -> int:
    """딕셔너리에서 여러 키로 값 검색 (유연한 매칭)

    검색 순서: 정확 → sj_div 접두사 → 부분 매칭
    """
    if not d:
        return 0
    for key in keys:
        # 1) 정확한 매칭
        if key in d:
            return d[key]
        # 2) sj_div 접두사 포함 (예: IS_매출액, BS_자산총계)
        for k, v in d.items():
            if k.endswith(f"_{key}"):
                return v
        # 3) 부분 매칭 (키가 다른 키에 포함)
        for k, v in d.items():
            if key in k:
                return v
    return 0


# ── 베이스 클래스 ─────────────────────────────────────────────

class PDFReportBase:
    """모든 PDF 리포트 생성기의 공통 베이스 클래스"""

    def __init__(self, output_path: str):
        self.output_path = output_path
        self.styles = getSampleStyleSheet()
        self._setup_styles()
        self._currency = "KRW"

    def _setup_styles(self):
        """커스텀 스타일 설정"""
        self.styles.add(ParagraphStyle(
            "KTitle",
            fontName=FONT_NAME,
            fontSize=14,
            leading=18,
            textColor=HEADER_FG,
            alignment=1,
        ))
        self.styles.add(ParagraphStyle(
            "KNormal",
            fontName=FONT_NAME,
            fontSize=7,
            leading=9,
        ))
        self.styles.add(ParagraphStyle(
            "KSmall",
            fontName=FONT_NAME,
            fontSize=6,
            leading=8,
        ))
        self.styles.add(ParagraphStyle(
            "KHeader",
            fontName=FONT_NAME,
            fontSize=8,
            leading=10,
            textColor=SECTION_FG,
        ))
        # 위험 분석용 스타일
        self.styles.add(ParagraphStyle(
            "KDesc",
            fontName=FONT_NAME,
            fontSize=5.5,
            leading=7,
            textColor=colors.HexColor("#7F8C8D"),
        ))
        self.styles.add(ParagraphStyle(
            "KRiskTitle",
            fontName=FONT_NAME,
            fontSize=7,
            leading=9,
            textColor=colors.white,
        ))
        self.styles.add(ParagraphStyle(
            "KRiskBody",
            fontName=FONT_NAME,
            fontSize=6,
            leading=8,
        ))

    def _prepare(self, report_data: dict):
        """리포트 데이터 전처리 (통화 감지, display_years 설정)"""
        info = report_data.get("company_info", {})
        self._currency = info.get("currency", "KRW")

        include_quarterly = report_data.get("include_quarterly", False)
        if include_quarterly and len(report_data.get("years", [])) > 4:
            report_data["display_years"] = report_data["years"][-4:]
        else:
            report_data["display_years"] = report_data.get("years", [])

    def _build_doc(self, elements: list):
        """SimpleDocTemplate 생성 및 빌드"""
        doc = SimpleDocTemplate(
            self.output_path,
            pagesize=landscape(A4),
            leftMargin=8 * mm,
            rightMargin=8 * mm,
            topMargin=6 * mm,
            bottomMargin=6 * mm,
        )
        doc.build(elements)
        print(f"\n리포트 생성 완료: {self.output_path}")

    # ── 통화별 포맷 헬퍼 ────────────────────────────────────

    def _fmt_amt(self, value) -> str:
        return _fmt_amount(value, self._currency)

    def _fmt_prc(self, value) -> str:
        return _fmt_price(value, self._currency)

    def _unit_label(self) -> str:
        return "억원" if self._currency == "KRW" else "$M"

    def _price_unit(self) -> str:
        return "원" if self._currency == "KRW" else "$"

    # ── 제목 바 ──────────────────────────────────────────────

    def _build_title_bar(self, data: dict, title_text: str = None,
                         bg_color=None) -> Table:
        info = data.get("company_info", {})
        corp_name = info.get("corp_name", "")
        stock_code = info.get("stock_code", "")

        if title_text is None:
            title_text = f"1p Company Analysis - {corp_name} ({stock_code})"

        if bg_color is None:
            bg_color = HEADER_BG

        title_data = [[
            Paragraph(title_text, self.styles["KTitle"]),
        ]]
        t = Table(title_data, colWidths=[270 * mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), bg_color),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    # ── 기업 개요 ────────────────────────────────────────────

    def _build_overview_section(self, data: dict) -> Table:
        info = data.get("company_info", {})
        stock = data.get("stock_data", {})
        srim = data.get("srim", {})
        shareholders = data.get("shareholders", [])

        # 좌: 기업 정보
        pu = self._price_unit()
        is_intl = info.get("is_international", False)
        ex_rate = 0
        if is_intl:
            mktcap_str = _fmt_amount(stock.get("market_cap", 0), self._currency)
            ex_rate = stock.get("exchange_rate", 0)
            price_krw = stock.get("price_krw", 0)
            mktcap_eok = stock.get("market_cap_eok", 0)
            left_info = [
                ["기업명", info.get("corp_name", "")],
                [f"시가총액", f'{mktcap_str} ({mktcap_eok:,}억원)'],
                [f"주가", f'{self._fmt_prc(stock.get("price", 0))} ({price_krw:,}원)'],
                ["기준일", f'{stock.get("date", "")} (1{self._currency}={ex_rate:,.0f}원)'],
            ]
        else:
            left_info = [
                ["기업명", info.get("corp_name", "")],
                [f"시가총액({self._unit_label()})", f'{stock.get("market_cap_eok", 0):,}'],
                [f"주가({pu})", self._fmt_prc(stock.get("price", 0))],
                ["기준일", stock.get("date", "")],
            ]

        # 중: 주주구성
        mid_info = [["[주주구성]", ""]]
        for sh in shareholders[:4]:
            mid_info.append([sh.get("name", ""), f'{sh.get("ratio", 0):.2f}%'])
        while len(mid_info) < 5:
            mid_info.append(["", ""])

        # 우: S-RIM 밸류에이션
        roe_source = srim.get("roe_source", "")
        roe_label = f"ROE 예측({roe_source})" if roe_source else "ROE 예측(%)"
        roe_hist = srim.get("roe_hist", 0)
        coe_value = srim.get("coe_value", data.get("required_return", 8.0))
        coe_source = srim.get("coe_source", "")
        coe_label = f"COE({coe_source})" if coe_source else "COE(%)"
        if is_intl and ex_rate:
            srim_p = srim.get("srim_price", 0)
            buy_p = srim.get("buy_price", 0)
            right_info = [
                ["[S-RIM 밸류에이션]", ""],
                ["S-RIM 적정가", f'{self._fmt_prc(srim_p)} ({round(srim_p * ex_rate):,}원)'],
                ["매수시작가", f'{self._fmt_prc(buy_p)} ({round(buy_p * ex_rate):,}원)'],
                [roe_label, f'{srim.get("roe_forecast", 0):.1f}%'],
                ["과거 ROE 가중평균", f'{roe_hist:.1f}%'],
                [coe_label, f'{coe_value:.1f}%'],
                [f"W(매수/적정)", f"{srim.get('w_buy', 0.5)}/{srim.get('w_fair', 1.0)}"],
            ]
        else:
            right_info = [
                ["[S-RIM 밸류에이션]", ""],
                [f"S-RIM 적정가({pu})", self._fmt_prc(srim.get("srim_price", 0))],
                [f"매수시작가({pu})", self._fmt_prc(srim.get("buy_price", 0))],
                [roe_label, f'{srim.get("roe_forecast", 0):.1f}%'],
                ["과거 ROE 가중평균", f'{roe_hist:.1f}%'],
                [coe_label, f'{coe_value:.1f}%'],
                [f"W(매수/적정)", f"{srim.get('w_buy', 0.5)}/{srim.get('w_fair', 1.0)}"],
            ]
        # ROE < COE 경고
        roe_f = srim.get("roe_forecast", 0)
        if roe_f and coe_value and roe_f < coe_value:
            right_info.append(["※ ROE < COE", "적정가/매수시작가 역전"])

        # CAPM 세부 정보 추가
        if srim.get("beta") is not None:
            right_info.append(["Beta", f'{srim.get("beta", 1.0):.2f}'])

        if is_intl:
            left_t = self._mini_table(left_info, [22 * mm, 45 * mm])
            right_t = self._mini_table(right_info, [30 * mm, 40 * mm])
        else:
            left_t = self._mini_table(left_info, [30 * mm, 30 * mm])
            right_t = self._mini_table(right_info, [35 * mm, 30 * mm])
        mid_t = self._mini_table(mid_info, [30 * mm, 20 * mm])

        outer = Table([[left_t, mid_t, right_t]], colWidths=[70 * mm, 55 * mm, 75 * mm])
        outer.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        return outer

    # ── 테이블 헬퍼 ──────────────────────────────────────────

    def _section_table(self, rows: list, num_years: int,
                       highlights: dict = None,
                       description: str = None,
                       ttm_column: bool = False) -> Table:
        """섹션 테이블 생성 (헤더 + 설명 + 데이터 행 + 하이라이트)"""
        # 컬럼 너비 계산
        label_width = 32 * mm
        total_data_cols = num_years + (1 if ttm_column else 0)
        data_width = (270 * mm / 2 - label_width) / max(total_data_cols, 1)
        col_widths = [label_width] + [data_width] * total_data_cols

        desc_offset = 0

        # Paragraph으로 변환
        styled_rows = []

        # 헤더 행
        header_row = []
        for cell in rows[0]:
            header_row.append(Paragraph(str(cell), self.styles["KHeader"]))
        styled_rows.append(header_row)

        # 설명 행 (옵션)
        if description:
            desc_row = [Paragraph(description, self.styles["KDesc"])]
            desc_row += [Paragraph("", self.styles["KDesc"])] * total_data_cols
            styled_rows.append(desc_row)
            desc_offset = 1

        # 데이터 행
        for i, row in enumerate(rows[1:], 1):
            styled_row = []
            for j, cell in enumerate(row):
                styled_row.append(Paragraph(str(cell), self.styles["KSmall"]))
            styled_rows.append(styled_row)

        t = Table(styled_rows, colWidths=col_widths)

        style_commands = [
            ("BACKGROUND", (0, 0), (-1, 0), SECTION_BG),
            ("TEXTCOLOR", (0, 0), (-1, 0), SECTION_FG),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6.5),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
        ]

        # TTM 컬럼 스타일
        if ttm_column:
            ttm_col = total_data_cols
            style_commands.append(
                ("BACKGROUND", (ttm_col, 0), (ttm_col, 0), TTM_HEADER_BG)
            )
            data_start_row = 1 + desc_offset
            for r in range(data_start_row, len(styled_rows)):
                style_commands.append(
                    ("BACKGROUND", (ttm_col, r), (ttm_col, r), TTM_CELL_BG)
                )

        # 설명 행 스타일
        if description:
            style_commands.append(("SPAN", (0, 1), (-1, 1)))
            style_commands.append(("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F0F4F8")))
            style_commands.append(("TOPPADDING", (0, 1), (-1, 1), 0.5))
            style_commands.append(("BOTTOMPADDING", (0, 1), (-1, 1), 0.5))

        # 짝수 행 배경색
        data_start = 1 + desc_offset
        for i in range(data_start + 1, len(styled_rows), 2):
            last_year_col = num_years
            style_commands.append(("BACKGROUND", (0, i), (last_year_col, i), LIGHT_ROW_BG))

        # 위험 하이라이트
        if highlights:
            for (row_i, col_j), risk_info in highlights.items():
                actual_row = row_i + desc_offset
                level = risk_info.get("level", LEVEL_OK)
                bg = RISK_LEVEL_BG.get(level)
                if bg:
                    style_commands.append(
                        ("BACKGROUND", (col_j, actual_row), (col_j, actual_row), bg)
                    )

        t.setStyle(TableStyle(style_commands))
        return t

    def _quarterly_table(self, rows: list, q_keys: list,
                         description: str = None) -> Table:
        """분기 상세 테이블 — 풀 페이지 너비 사용"""
        num_q = len(q_keys)
        label_width = 38 * mm
        q_width = (270 * mm - label_width) / max(num_q, 1)
        col_widths = [label_width] + [q_width] * num_q

        desc_offset = 0
        styled_rows = []

        # 헤더
        header_row = []
        for cell in rows[0]:
            header_row.append(Paragraph(str(cell), self.styles["KHeader"]))
        styled_rows.append(header_row)

        # 설명
        if description:
            desc_row = [Paragraph(description, self.styles["KDesc"])]
            desc_row += [Paragraph("", self.styles["KDesc"])] * num_q
            styled_rows.append(desc_row)
            desc_offset = 1

        # 데이터
        for row in rows[1:]:
            styled_row = []
            for cell in row:
                styled_row.append(Paragraph(str(cell), self.styles["KSmall"]))
            styled_rows.append(styled_row)

        t = Table(styled_rows, colWidths=col_widths)

        style_commands = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B4F72")),
            ("TEXTCOLOR", (0, 0), (-1, 0), SECTION_FG),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6.5),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
        ]

        if description:
            style_commands.append(("SPAN", (0, 1), (-1, 1)))
            style_commands.append(("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F0F4F8")))

        data_start = 1 + desc_offset
        for i in range(data_start + 1, len(styled_rows), 2):
            style_commands.append(("BACKGROUND", (0, i), (-1, i), LIGHT_ROW_BG))

        t.setStyle(TableStyle(style_commands))
        return t

    def _mini_table(self, rows: list, col_widths: list) -> Table:
        """작은 정보 테이블"""
        styled_rows = []
        for i, row in enumerate(rows):
            styled_row = []
            for cell in row:
                styled_row.append(Paragraph(str(cell), self.styles["KSmall"]))
            styled_rows.append(styled_row)

        t = Table(styled_rows, colWidths=col_widths)
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
            ("BACKGROUND", (0, 0), (0, 0), SUB_HEADER_BG),
        ]))
        return t

    # ── 데이터 검색 헬퍼 ─────────────────────────────────────

    @staticmethod
    def _find_account(fs: dict, year: int, key: str) -> int:
        """financial_summary에서 계정 값 찾기 (유연한 매칭)"""
        year_data = fs.get(year, {})
        if not year_data:
            return 0

        # 정확한 매칭
        if key in year_data:
            return year_data[key]

        # sj_div 접두사 포함 매칭
        for k, v in year_data.items():
            if k.endswith(f"_{key}") or key in k:
                return v

        # 부분 매칭
        key_parts = key.replace("(", "").replace(")", "")
        for k, v in year_data.items():
            clean_k = k.replace("(", "").replace(")", "")
            if key_parts in clean_k:
                return v

        return 0

    @staticmethod
    def _find_in_dict(d: dict, *keys: str) -> int:
        """딕셔너리에서 여러 키로 값 찾기"""
        for key in keys:
            if key in d:
                return d[key]
            # 부분 매칭
            for k, v in d.items():
                if key in k:
                    return v
        return 0
