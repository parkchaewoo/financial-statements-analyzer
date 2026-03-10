"""PDF 리포트 생성 모듈 - reportlab 기반 기업 분석 리포트

기존 통합 리포트를 유지하면서 3개 서브 리포트 생성기에 위임.
- AnnualReportGenerator:    연도별+TTM 재무 분석
- QuarterlyReportGenerator: 분기별 재무 분석
- RiskReportGenerator:      상장폐지/관리종목 리스크 분석
"""

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

from pdf_report_base import (
    PDFReportBase,
    FONT_NAME,
    HEADER_BG, HEADER_FG,
    SECTION_BG, SECTION_FG,
    BORDER_COLOR, LIGHT_ROW_BG,
    POSITIVE_COLOR, NEGATIVE_COLOR,
    _fmt_price,
)
from pdf_annual_report import AnnualReportGenerator
from pdf_quarterly_report import QuarterlyReportGenerator
from pdf_risk_report import RiskReportGenerator


class PDFReportGenerator(PDFReportBase):
    """통합 기업 분석 PDF 리포트 생성기 (기존 호환)

    내부적으로 Annual / Quarterly / Risk 서브 생성기에 위임하여
    하나의 PDF 파일에 모든 분석을 포함한다.
    """

    def generate(self, report_data: dict):
        """메인 PDF 생성 — 3개 서브 리포트 결합"""
        self._prepare(report_data)

        elements = []

        # 연도별 + TTM 재무 분석 (페이지 1-4)
        annual = AnnualReportGenerator(self.output_path)
        elements.extend(annual.build_elements(report_data))

        # 분기별 재무 분석 (optional)
        include_quarterly = report_data.get("include_quarterly", False)
        if include_quarterly:
            quarterly = QuarterlyReportGenerator(self.output_path)
            q_elements = quarterly.build_elements(report_data)
            if q_elements:
                from reportlab.platypus import PageBreak
                elements.append(PageBreak())
                elements.extend(q_elements)

        # 리스크 분석 (페이지 5-6)
        from reportlab.platypus import PageBreak
        elements.append(PageBreak())
        risk = RiskReportGenerator(self.output_path)
        elements.extend(risk.build_elements(report_data))

        self._build_doc(elements)


# ── S-RIM 스크리너 PDF 생성 ──────────────────────────────────

class ScreenerPDFGenerator(PDFReportBase):
    """S-RIM 스크리너 결과를 PDF로 생성"""

    def __init__(self, output_path: str):
        super().__init__(output_path)
        # 스크리너 전용 스타일 추가
        self.styles.add(ParagraphStyle(
            "ScrTitle", fontName=FONT_NAME, fontSize=14,
            leading=18, textColor=HEADER_FG, alignment=1,
        ))
        self.styles.add(ParagraphStyle(
            "ScrNormal", fontName=FONT_NAME, fontSize=7, leading=9,
        ))
        self.styles.add(ParagraphStyle(
            "ScrSmall", fontName=FONT_NAME, fontSize=6, leading=8,
        ))
        self.styles.add(ParagraphStyle(
            "ScrHeader", fontName=FONT_NAME, fontSize=7,
            leading=9, textColor=SECTION_FG,
        ))
        self.styles.add(ParagraphStyle(
            "ScrInfo", fontName=FONT_NAME, fontSize=8,
            leading=10, textColor=colors.HexColor("#2C3E50"),
        ))

    def generate(self, results: list[dict], roe_source: str = "consensus",
                 required_return: float = 8.0, market: str = "",
                 currency: str = "KRW", coe_source: str = "manual"):
        """스크리너 결과 PDF 생성"""
        self._currency = currency
        from datetime import datetime

        doc = SimpleDocTemplate(
            self.output_path,
            pagesize=landscape(A4),
            leftMargin=10 * mm,
            rightMargin=10 * mm,
            topMargin=8 * mm,
            bottomMargin=8 * mm,
        )

        elements = []

        # ── 제목 ──
        title_data = [[
            Paragraph("S-RIM 저평가 종목 스크리너", self.styles["ScrTitle"]),
        ]]
        title_t = Table(title_data, colWidths=[266 * mm])
        title_t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), HEADER_BG),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        elements.append(title_t)
        elements.append(Spacer(1, 3 * mm))

        # ── 요약 정보 ──
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        undervalued = [r for r in results if r.get("discount_pct", 0) > 0]
        overvalued = [r for r in results if r.get("discount_pct", 0) <= 0]

        roe_label = {"consensus": "컨센서스 우선", "historical": "과거 가중평균"}.get(
            roe_source, roe_source)
        coe_label = "CAPM(종목별 자동)" if coe_source == "capm" else f"수동({required_return:.0f}%)"

        info_rows = [[
            Paragraph(f"기준일: {now}", self.styles["ScrInfo"]),
            Paragraph(f"대상: {market or '직접입력'} ({len(results)}개 종목)",
                      self.styles["ScrInfo"]),
            Paragraph(f"ROE: {roe_label} | COE: {coe_label}",
                      self.styles["ScrInfo"]),
            Paragraph(f"저평가: {len(undervalued)}개 | 고평가: {len(overvalued)}개",
                      self.styles["ScrInfo"]),
        ]]
        info_t = Table(info_rows, colWidths=[70 * mm, 70 * mm, 70 * mm, 56 * mm])
        info_t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ECF0F1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("BOX", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
        ]))
        elements.append(info_t)
        elements.append(Spacer(1, 3 * mm))

        # ── 저평가 종목 테이블 ──
        if undervalued:
            elements.append(self._build_results_table(undervalued, "저평가 종목", currency))
            elements.append(Spacer(1, 3 * mm))

        # ── 고평가 종목 테이블 ──
        if overvalued:
            elements.append(self._build_results_table(overvalued, "고평가 종목", currency))
            elements.append(Spacer(1, 3 * mm))

        # ── 범례 ──
        legend_text = (
            "할인율 = (S-RIM적정가 - 현재가) / S-RIM적정가 x 100  |  "
            "양수: 저평가 (매수 기회)  |  음수: 고평가  |  "
            "ROE예측: S-RIM 계산에 사용된 ROE"
        )
        legend = Paragraph(legend_text, self.styles["ScrSmall"])
        elements.append(legend)

        doc.build(elements)

    def _build_results_table(self, results: list[dict], title: str,
                             currency: str = "KRW") -> Table:
        """결과 테이블 생성"""
        col_headers = [
            "순위", "종목명", "코드", "현재가", "S-RIM적정가",
            "매수시작가", "할인율(%)", "ROE예측(%)", "COE(%)",
            "ROE소스", "OPM(%)", "PER", "PBR",
        ]

        col_widths = [
            8 * mm, 27 * mm, 14 * mm, 18 * mm, 18 * mm,
            18 * mm, 14 * mm, 14 * mm, 13 * mm,
            18 * mm, 13 * mm, 13 * mm, 12 * mm,
        ]

        # 섹션 헤더
        section_header = [Paragraph(f"[{title}] ({len(results)}개)", self.styles["ScrHeader"])]
        section_header += [Paragraph("", self.styles["ScrHeader"])] * (len(col_headers) - 1)

        # 컬럼 헤더
        header_row = [Paragraph(h, self.styles["ScrHeader"]) for h in col_headers]

        rows = [section_header, header_row]

        for rank, r in enumerate(results, 1):
            discount = r.get("discount_pct", 0)
            roe_src = r.get("roe_source", "")

            coe_val = r.get("coe_value", 0)
            corp_name = r.get("corp_name", "")
            corp_name_kr = r.get("corp_name_kr", "")
            if corp_name_kr:
                display_name = f"{corp_name}<br/><font size='5'>({corp_name_kr})</font>"
            else:
                display_name = corp_name
            row = [
                Paragraph(str(rank), self.styles["ScrSmall"]),
                Paragraph(display_name, self.styles["ScrSmall"]),
                Paragraph(r.get("stock_code", ""), self.styles["ScrSmall"]),
                Paragraph(_fmt_price(r.get("price", 0), currency), self.styles["ScrSmall"]),
                Paragraph(_fmt_price(r.get("srim_price", 0), currency), self.styles["ScrSmall"]),
                Paragraph(_fmt_price(r.get("buy_price", 0), currency), self.styles["ScrSmall"]),
                Paragraph(f'{discount:+.1f}', self.styles["ScrSmall"]),
                Paragraph(f'{r.get("roe_forecast", 0):.1f}', self.styles["ScrSmall"]),
                Paragraph(f'{coe_val:.1f}' if coe_val else "-", self.styles["ScrSmall"]),
                Paragraph(roe_src, self.styles["ScrSmall"]),
                Paragraph(f'{r.get("opm", 0):.1f}', self.styles["ScrSmall"]),
                Paragraph(f'{r.get("per", 0):.1f}' if r.get("per") else "-",
                          self.styles["ScrSmall"]),
                Paragraph(f'{r.get("pbr", 0):.2f}' if r.get("pbr") else "-",
                          self.styles["ScrSmall"]),
            ]
            rows.append(row)

        t = Table(rows, colWidths=col_widths)
        style_cmds = [
            # 섹션 헤더
            ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
            ("SPAN", (0, 0), (-1, 0)),
            # 컬럼 헤더
            ("BACKGROUND", (0, 1), (-1, 1), SECTION_BG),
            ("TEXTCOLOR", (0, 1), (-1, 1), SECTION_FG),
            # 전체
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6.5),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),  # 순위
            ("ALIGN", (2, 0), (2, -1), "CENTER"),   # 코드
            ("ALIGN", (3, 1), (-1, -1), "RIGHT"),   # 숫자 열
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 1), (-1, -1), 0.3, BORDER_COLOR),
        ]

        # 짝수행 배경
        for i in range(3, len(rows), 2):
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_ROW_BG))

        # 할인율 색상 (저평가=파랑, 고평가=빨강)
        for i in range(2, len(rows)):
            r = results[i - 2]
            discount = r.get("discount_pct", 0)
            if discount > 20:
                style_cmds.append(("TEXTCOLOR", (6, i), (6, i), POSITIVE_COLOR))
            elif discount < 0:
                style_cmds.append(("TEXTCOLOR", (6, i), (6, i), NEGATIVE_COLOR))

        t.setStyle(TableStyle(style_cmds))
        return t
