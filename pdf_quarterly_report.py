"""Quarterly Financial Analysis Report Generator

분기별 재무 분석 리포트를 생성하는 모듈.
PDFReportBase를 상속하여 분기 손익, 수익성, 현금흐름, 재무상태표,
TTM 비교, 분기 모멘텀 분석을 포함한 PDF 리포트를 생성한다.
"""

from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, PageBreak
from pdf_report_base import (
    PDFReportBase, FONT_NAME, SECTION_BG, SECTION_FG,
    BORDER_COLOR, LIGHT_ROW_BG, SUB_HEADER_BG,
    TREND_POSITIVE_BG, TREND_NEGATIVE_BG, TREND_NEUTRAL_BG,
    _fmt_num, _fmt_amount,
)
from trend_analyzer import (
    SITUATION_EMOJI,
    TREND_UP, TREND_DOWN, TREND_FLAT, TREND_VOLATILE, TREND_RECOVERY,
)


class QuarterlyReportGenerator(PDFReportBase):
    """분기별 재무 분석 리포트 생성기"""

    # ── 엔트리 포인트 ─────────────────────────────────────────

    def generate(self, report_data):
        """단독 PDF 리포트 생성"""
        elements = self.build_elements(report_data)
        self._build_doc(elements)

    def build_elements(self, report_data):
        """리포트 요소 리스트 반환 (결합 리포트용)"""
        self._prepare(report_data)
        derived = report_data.get("derived", {})
        quarterly = derived.get("_quarterly", {})
        q_keys = derived.get("_quarterly_keys", [])

        if not quarterly or not q_keys:
            return []

        elements = []
        info = report_data.get("company_info", {})
        corp_name = info.get("corp_name", "")

        # Title bar (color: #1B4F72)
        elements.append(self._build_title_bar(
            report_data,
            title_text=f"분기별 재무 분석 - {corp_name}",
            bg_color=colors.HexColor("#1B4F72")
        ))
        elements.append(Spacer(1, 3 * mm))

        # TTM vs Annual overview
        elements.append(self._build_quarterly_overview(report_data))
        elements.append(Spacer(1, 3 * mm))

        # Quarterly tables
        q_is = self._build_quarterly_income(report_data, q_keys, quarterly)
        if q_is:
            elements.append(q_is)
            elements.append(Spacer(1, 2 * mm))

        q_prof = self._build_quarterly_profitability(report_data, q_keys, quarterly)
        if q_prof:
            elements.append(q_prof)
            elements.append(Spacer(1, 2 * mm))

        q_cf = self._build_quarterly_cashflow(report_data, q_keys, quarterly)
        if q_cf:
            elements.append(q_cf)
            elements.append(Spacer(1, 2 * mm))

        q_bs = self._build_quarterly_balance_sheet(report_data, q_keys, quarterly)
        if q_bs:
            elements.append(q_bs)
            elements.append(Spacer(1, 2 * mm))

        # Momentum summary
        elements.append(self._build_quarterly_momentum(report_data, q_keys, quarterly))

        # Trend analysis page (if available)
        trend_result = report_data.get("trend_analysis")
        if trend_result and trend_result.get("details"):
            elements.append(PageBreak())
            trend_elems = self._build_quarterly_trend_page(report_data)
            if trend_elems:
                elements.extend(trend_elems)

        return elements

    # ── TTM vs 최근 연간 비교 ─────────────────────────────────

    def _build_quarterly_overview(self, data):
        """TTM vs 최근 연간 비교 패널"""
        years = data.get("display_years", data.get("years", []))
        fs = data.get("financial_summary", {})
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_fs = ttm.get("financial_summary", {})
        ttm_derived = derived.get("_ttm", {})

        # 최근 연간 연도
        latest_year = years[-1] if years else ""
        latest_fs = fs.get(latest_year, {}) if latest_year else {}
        latest_derived = derived.get(latest_year, {}) if latest_year else {}

        # 연간 값 구하기 헬퍼
        def _get_annual_amount(key, *alt_keys):
            val = self._find_account(fs, latest_year, key) if latest_year else 0
            if not val:
                for ak in alt_keys:
                    val = self._find_account(fs, latest_year, ak) if latest_year else 0
                    if val:
                        break
            return val

        def _get_ttm_amount(key, *alt_keys):
            val = self._find_in_dict(ttm_fs, key)
            if not val:
                for ak in alt_keys:
                    val = self._find_in_dict(ttm_fs, ak)
                    if val:
                        break
            return val

        # 지표 정의: (label, annual_value, ttm_value, unit)
        annual_revenue = _get_annual_amount("매출액", "영업수익")
        ttm_revenue = _get_ttm_amount("매출액", "영업수익")

        annual_op = _get_annual_amount("영업이익")
        ttm_op = _get_ttm_amount("영업이익")

        annual_ni = _get_annual_amount("당기순이익")
        ttm_ni = _get_ttm_amount("당기순이익")

        annual_opm = latest_derived.get("영업이익률(%)", 0)
        ttm_opm = ttm_derived.get("영업이익률(%)", 0)

        annual_roe = latest_derived.get("ROE(%)", 0)
        ttm_roe = ttm_derived.get("ROE(%)", 0)

        annual_fcf = latest_derived.get("FCF", 0)
        ttm_fcf = ttm_derived.get("FCF", 0)

        metrics = [
            ("매출액", annual_revenue, ttm_revenue, "amount"),
            ("영업이익", annual_op, ttm_op, "amount"),
            ("당기순이익", annual_ni, ttm_ni, "amount"),
            ("영업이익률(%)", annual_opm, ttm_opm, "pct"),
            ("ROE(%)", annual_roe, ttm_roe, "pct"),
            ("FCF", annual_fcf, ttm_fcf, "amount"),
        ]

        # 변화 방향 및 델타 계산
        def _arrow_and_delta(annual_val, ttm_val, unit):
            if annual_val is None or ttm_val is None:
                return "-", "-"
            if annual_val == 0 and ttm_val == 0:
                return "-", "-"

            if unit == "amount":
                delta = ttm_val - annual_val
                if annual_val != 0:
                    pct_change = (delta / abs(annual_val)) * 100
                else:
                    pct_change = 0
                if delta > 0:
                    arrow = "\u2191"
                    delta_str = f'<font color="#27AE60">{arrow} +{pct_change:.1f}%</font>'
                elif delta < 0:
                    arrow = "\u2193"
                    delta_str = f'<font color="#E74C3C">{arrow} {pct_change:.1f}%</font>'
                else:
                    arrow = "\u2192"
                    delta_str = f'{arrow} 0.0%'
            else:  # pct
                delta = ttm_val - annual_val
                if delta > 0.05:
                    arrow = "\u2191"
                    delta_str = f'<font color="#27AE60">{arrow} +{delta:.1f}%p</font>'
                elif delta < -0.05:
                    arrow = "\u2193"
                    delta_str = f'<font color="#E74C3C">{arrow} {delta:.1f}%p</font>'
                else:
                    arrow = "\u2192"
                    delta_str = f'{arrow} {delta:.1f}%p'

            return arrow, delta_str

        # 테이블 구성
        year_label = str(latest_year) if latest_year else "-"
        ttm_label = ttm.get("ttm_label", "TTM")
        header = [f"[TTM vs 최근 연간 비교]", f"최근연간({year_label})", ttm_label, "변화"]
        rows_data = [header]

        for label, ann_val, ttm_val, unit in metrics:
            _, delta_str = _arrow_and_delta(ann_val, ttm_val, unit)

            if unit == "amount":
                ann_str = self._fmt_amt(ann_val)
                ttm_str = self._fmt_amt(ttm_val)
            else:
                ann_str = _fmt_num(ann_val, "pct") if ann_val else "-"
                ttm_str = _fmt_num(ttm_val, "pct") if ttm_val else "-"

            rows_data.append([label, ann_str, ttm_str, delta_str])

        # Paragraph 변환
        styled_rows = []

        # 헤더 행
        header_row = []
        for cell in rows_data[0]:
            header_row.append(Paragraph(str(cell), self.styles["KHeader"]))
        styled_rows.append(header_row)

        # 데이터 행
        for row in rows_data[1:]:
            styled_row = []
            for j, cell in enumerate(row):
                styled_row.append(Paragraph(str(cell), self.styles["KSmall"]))
            styled_rows.append(styled_row)

        col_widths = [40 * mm, 50 * mm, 50 * mm, 50 * mm]
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

        # TTM 컬럼 (3번째 데이터 컬럼) 배경색
        ttm_col = 2
        style_commands.append(
            ("BACKGROUND", (ttm_col, 0), (ttm_col, 0), colors.HexColor("#27AE60"))
        )
        for r in range(1, len(styled_rows)):
            style_commands.append(
                ("BACKGROUND", (ttm_col, r), (ttm_col, r), colors.HexColor("#E8F8F5"))
            )

        # 짝수 행 배경색 (TTM 컬럼 제외)
        for i in range(2, len(styled_rows), 2):
            style_commands.append(("BACKGROUND", (0, i), (1, i), LIGHT_ROW_BG))
            style_commands.append(("BACKGROUND", (3, i), (3, i), LIGHT_ROW_BG))

        t.setStyle(TableStyle(style_commands))
        return t

    # ── 분기 손익계산서 ───────────────────────────────────────

    def _build_quarterly_income(self, data: dict, q_keys: list,
                                quarterly: dict) -> Table:
        """분기 손익계산서"""
        header = ["[분기 손익]"] + [k for k in q_keys]
        rows = [header]

        for label, key in [("매출액", "매출액"), ("영업이익", "영업이익"),
                           ("당기순이익", "당기순이익")]:
            row = [label]
            for q in q_keys:
                val = quarterly.get(q, {}).get(key, 0)
                row.append(self._fmt_amt(val))
            rows.append(row)

        # 성장률
        for label, key in [("매출 YoY(%)", "매출성장률_YoY(%)"),
                           ("영업이익 YoY(%)", "영업이익성장률_YoY(%)"),
                           ("매출 QoQ(%)", "매출성장률_QoQ(%)"),
                           ("영업이익 QoQ(%)", "영업이익성장률_QoQ(%)")]:
            row = [label]
            for q in q_keys:
                val = quarterly.get(q, {}).get(key)
                if val is not None:
                    row.append(f"{val:+.1f}")
                else:
                    row.append("-")
            rows.append(row)

        desc = f"단위: {self._unit_label()} | YoY: 전년 동분기 대비 | QoQ: 직전 분기 대비"
        return self._quarterly_table(rows, q_keys, description=desc)

    # ── 분기 수익성 ───────────────────────────────────────────

    def _build_quarterly_profitability(self, data: dict, q_keys: list,
                                       quarterly: dict) -> Table:
        """분기 수익성"""
        header = ["[분기 수익성]"] + [k for k in q_keys]
        rows = [header]

        for label, key in [("영업이익률(%)", "영업이익률(%)"),
                           ("ROE(%)", "ROE(%)"),
                           ("ROA(%)", "ROA(%)")]:
            row = [label]
            for q in q_keys:
                val = quarterly.get(q, {}).get(key, 0)
                row.append(_fmt_num(val, "pct") if val else "-")
            rows.append(row)

        desc = "분기별 수익성 지표 추이"
        return self._quarterly_table(rows, q_keys, description=desc)

    # ── 분기 현금흐름 ─────────────────────────────────────────

    def _build_quarterly_cashflow(self, data: dict, q_keys: list,
                                  quarterly: dict) -> Table:
        """분기 현금흐름"""
        header = ["[분기 현금흐름]"] + [k for k in q_keys]
        rows = [header]

        for label, key in [("영업활동CF", "영업활동CF"),
                           ("FCF", "FCF")]:
            row = [label]
            for q in q_keys:
                val = quarterly.get(q, {}).get(key, 0)
                row.append(self._fmt_amt(val))
            rows.append(row)

        desc = "분기별 현금흐름 추이"
        return self._quarterly_table(rows, q_keys, description=desc)

    # ── 분기 재무상태표 추이 ──────────────────────────────────

    def _build_quarterly_balance_sheet(self, data: dict, q_keys: list,
                                       quarterly: dict) -> Table:
        """분기 재무상태표 추이"""
        header = ["[분기 자산/부채/자본]"] + [k for k in q_keys]
        rows = [header]

        for label, key in [("자산총계", "자산총계"),
                           ("부채총계", "부채총계"),
                           ("자본총계", "자본총계")]:
            row = [label]
            for q in q_keys:
                val = quarterly.get(q, {}).get(key, 0)
                row.append(self._fmt_amt(val))
            rows.append(row)

        desc = "분기별 재무상태 추이 (시점 데이터)"
        return self._quarterly_table(rows, q_keys, description=desc)

    # ── 분기별 QoQ/YoY 추세 요약 (모멘텀) ────────────────────

    def _build_quarterly_momentum(self, data, q_keys, quarterly):
        """분기별 QoQ/YoY 추세 요약 테이블"""

        def _trend_arrow(val):
            """값에 따른 추세 화살표 및 색상 반환"""
            if val is None:
                return "-"
            if val > 0.5:
                return f'<font color="#27AE60">\u2191 +{val:.1f}%</font>'
            elif val < -0.5:
                return f'<font color="#E74C3C">\u2193 {val:.1f}%</font>'
            else:
                return f'\u2192 {val:.1f}%'

        def _direction_arrow(val):
            """방향 화살표만 (색상 포함)"""
            if val is None:
                return "-"
            if val > 0.5:
                return f'<font color="#27AE60">\u2191</font>'
            elif val < -0.5:
                return f'<font color="#E74C3C">\u2193</font>'
            else:
                return '\u2192'

        header = ["[분기 모멘텀]"] + [k for k in q_keys]
        rows_data = [header]

        # 매출 YoY 추세
        row_rev_yoy = ["매출 YoY"]
        for q in q_keys:
            val = quarterly.get(q, {}).get("매출성장률_YoY(%)")
            row_rev_yoy.append(_trend_arrow(val))
        rows_data.append(row_rev_yoy)

        # 영업이익 YoY 추세
        row_op_yoy = ["영업이익 YoY"]
        for q in q_keys:
            val = quarterly.get(q, {}).get("영업이익성장률_YoY(%)")
            row_op_yoy.append(_trend_arrow(val))
        rows_data.append(row_op_yoy)

        # OPM 방향
        row_opm = ["OPM 방향"]
        prev_opm = None
        for q in q_keys:
            opm = quarterly.get(q, {}).get("영업이익률(%)")
            if opm is not None and prev_opm is not None:
                delta = opm - prev_opm
                row_opm.append(_direction_arrow(delta))
            elif opm is not None:
                row_opm.append(f'{opm:.1f}%')
            else:
                row_opm.append("-")
            prev_opm = opm
        rows_data.append(row_opm)

        # 매출 QoQ 추세
        row_rev_qoq = ["매출 QoQ"]
        for q in q_keys:
            val = quarterly.get(q, {}).get("매출성장률_QoQ(%)")
            row_rev_qoq.append(_trend_arrow(val))
        rows_data.append(row_rev_qoq)

        # 영업이익 QoQ 추세
        row_op_qoq = ["영업이익 QoQ"]
        for q in q_keys:
            val = quarterly.get(q, {}).get("영업이익성장률_QoQ(%)")
            row_op_qoq.append(_trend_arrow(val))
        rows_data.append(row_op_qoq)

        # Paragraph 변환
        num_q = len(q_keys)
        label_width = 38 * mm
        q_width = (270 * mm - label_width) / max(num_q, 1)
        col_widths = [label_width] + [q_width] * num_q

        styled_rows = []

        # 헤더 행
        header_row = []
        for cell in rows_data[0]:
            header_row.append(Paragraph(str(cell), self.styles["KHeader"]))
        styled_rows.append(header_row)

        # 데이터 행
        for row in rows_data[1:]:
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
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
        ]

        # 짝수 행 배경색
        for i in range(2, len(styled_rows), 2):
            style_commands.append(("BACKGROUND", (0, i), (-1, i), LIGHT_ROW_BG))

        t.setStyle(TableStyle(style_commands))
        return t

    # ── 분기별 추이 분석 페이지 ─────────────────────────────────

    def _build_quarterly_trend_page(self, data: dict) -> list:
        """분기별 추이 분석 페이지 (trend_analysis 데이터 활용)"""
        trend = data.get("trend_analysis", {})
        if not trend or not trend.get("details"):
            return []

        elements = []
        situation = trend.get("situation", "")
        label = trend.get("situation_label", "분석 불가")
        color_hex = trend.get("situation_color", "#95A5A6")
        confidence = trend.get("confidence", 0)
        summary = trend.get("summary", "")
        details = trend.get("details", [])
        key_metrics = trend.get("key_metrics", {})

        info = data.get("company_info", {})
        corp_name = info.get("corp_name", "")

        # ── 제목바 ──
        title_data = [[
            Paragraph(
                f"분기별 추이 분석 (Quarterly Trend Analysis) - {corp_name}",
                self.styles["KTitle"],
            ),
        ]]
        t = Table(title_data, colWidths=[270 * mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1B4F72")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 3 * mm))

        # ── 종합 진단 패널 ──
        situation_color = colors.HexColor(color_hex)
        emoji = SITUATION_EMOJI.get(situation, "?")

        diag_rows = []
        diag_rows.append([
            Paragraph(
                f'<font color="white" size="10"><b>'
                f'  {label}  '
                f'</b></font>',
                self.styles["KTitle"],
            ),
            Paragraph(
                f'<font size="7"><b>진단 신뢰도: {confidence * 100:.0f}%</b></font>',
                self.styles["KNormal"],
            ),
        ])

        # 핵심 지표 요약
        cagr = key_metrics.get("revenue_cagr", 0)
        opm = key_metrics.get("opm_latest", 0)
        roe = key_metrics.get("roe_latest", 0)
        dr = key_metrics.get("debt_ratio_latest", 0)
        fcf_t = key_metrics.get("fcf_trend", "-")
        momentum = key_metrics.get("quarterly_momentum", "-")

        metrics_text = (
            f'매출CAGR <b>{cagr:+.1f}%</b> | '
            f'OPM <b>{opm:.1f}%</b> | '
            f'ROE <b>{roe:.1f}%</b> | '
            f'부채비율 <b>{dr:.0f}%</b> | '
            f'FCF <b>{fcf_t}</b>'
        )
        if momentum and momentum not in ("데이터 없음", "데이터 부족"):
            metrics_text += f' | 분기모멘텀 <b>{momentum}</b>'
        diag_rows.append([
            Paragraph(metrics_text, self.styles["KNormal"]),
            Paragraph("", self.styles["KNormal"]),
        ])

        diag_t = Table(diag_rows, colWidths=[200 * mm, 70 * mm])
        diag_t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), situation_color),
            ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#ECEFF1")),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F5F5F5")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, 0), "CENTER"),
            ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#1B4F72")),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.HexColor("#1B4F72")),
        ]))
        elements.append(diag_t)
        elements.append(Spacer(1, 2 * mm))

        # ── 종합 요약 ──
        summary_rows = [[
            Paragraph(
                '<font size="7"><b>[종합 진단]</b></font>',
                self.styles["KRiskTitle"],
            ),
        ]]
        summary_rows.append([
            Paragraph(
                f'<font size="6.5">{summary}</font>',
                self.styles["KNormal"],
            ),
        ])

        summary_t = Table(summary_rows, colWidths=[270 * mm])
        summary_t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B4F72")),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#D6EAF8")),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#1B4F72")),
        ]))
        elements.append(summary_t)
        elements.append(Spacer(1, 3 * mm))

        # ── 영역별 상세 분석 테이블 ──
        detail_rows = []

        detail_rows.append([
            Paragraph('<font color="white"><b>분석 영역</b></font>', self.styles["KHeader"]),
            Paragraph('<font color="white"><b>추세</b></font>', self.styles["KHeader"]),
            Paragraph('<font color="white"><b>점수</b></font>', self.styles["KHeader"]),
            Paragraph('<font color="white"><b>진단</b></font>', self.styles["KHeader"]),
            Paragraph('<font color="white"><b>상세 분석</b></font>', self.styles["KHeader"]),
        ])

        trend_labels = {
            TREND_UP: "\u2191 상승",
            TREND_DOWN: "\u2193 하락",
            TREND_FLAT: "\u2192 횡보",
            TREND_VOLATILE: "\u2195 변동",
            TREND_RECOVERY: "\u2197 회복",
        }

        for d in details:
            category = d.get("category", "")
            trend_dir = d.get("trend", TREND_FLAT)
            score = d.get("score", 0)
            title = d.get("title", "")
            comment = d.get("comment", "")
            sub_items = d.get("sub_items", [])

            trend_text = trend_labels.get(trend_dir, "\u2192 횡보")

            if score >= 2:
                score_text = '<font color="#2ECC71"><b>A+</b></font>'
            elif score == 1:
                score_text = '<font color="#3498DB"><b>A</b></font>'
            elif score == 0:
                score_text = '<font color="#7F8C8D"><b>B</b></font>'
            elif score == -1:
                score_text = '<font color="#E67E22"><b>C</b></font>'
            else:
                score_text = '<font color="#E74C3C"><b>D</b></font>'

            full_comment = comment
            if sub_items:
                si_parts = []
                for si in sub_items:
                    label_s = si.get("label", "")
                    value_s = si.get("value", "")
                    si_parts.append(f"{label_s} {value_s}")
                sub_line = " | ".join(si_parts)
                full_comment += f'<br/><font size="5" color="#555555">  \u25b8 {sub_line}</font>'

            detail_rows.append([
                Paragraph(f'<b>{category}</b>', self.styles["KNormal"]),
                Paragraph(trend_text, self.styles["KSmall"]),
                Paragraph(score_text, self.styles["KNormal"]),
                Paragraph(title, self.styles["KSmall"]),
                Paragraph(full_comment, self.styles["KSmall"]),
            ])

        col_widths = [30 * mm, 18 * mm, 15 * mm, 65 * mm, 142 * mm]
        detail_t = Table(detail_rows, colWidths=col_widths)

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B4F72")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6.5),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("ALIGN", (2, 0), (2, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
        ]

        for i, d in enumerate(details, start=1):
            score = d.get("score", 0)
            if score >= 1:
                bg = TREND_POSITIVE_BG
            elif score <= -1:
                bg = TREND_NEGATIVE_BG
            else:
                bg = TREND_NEUTRAL_BG
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), bg))

        detail_t.setStyle(TableStyle(style_cmds))
        elements.append(detail_t)
        elements.append(Spacer(1, 3 * mm))

        # ── 점수 범례 + 면책 조항 ──
        legend_rows = [[
            Paragraph(
                '<font size="5.5" color="#555555">'
                '<b>점수 기준:</b> '
                '<font color="#2ECC71"><b>A+</b></font> 매우 양호 | '
                '<font color="#3498DB"><b>A</b></font> 양호 | '
                '<font color="#7F8C8D"><b>B</b></font> 보통 | '
                '<font color="#E67E22"><b>C</b></font> 부진 | '
                '<font color="#E74C3C"><b>D</b></font> 심각'
                '</font>',
                self.styles["KDesc"],
            ),
        ]]
        legend_rows.append([
            Paragraph(
                '<font size="5" color="#999999">'
                '\u203b 본 분석은 과거 재무데이터 기반 자동 진단이며, 투자 권유가 아닙니다. '
                '업종 특성, 경영 환경 변화, 일회성 요인 등은 반영되지 않습니다.'
                '</font>',
                self.styles["KDesc"],
            ),
        ])

        legend_t = Table(legend_rows, colWidths=[270 * mm])
        legend_t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FAFAFA")),
            ("BOX", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
        ]))
        elements.append(legend_t)
        elements.append(Spacer(1, 2 * mm))

        # ── 강점/약점 패널 ──
        sw = trend.get("strengths_weaknesses", {})
        strengths = sw.get("strengths", [])
        weaknesses = sw.get("weaknesses", [])
        opportunities = sw.get("opportunities", [])
        risks_list = sw.get("risks", [])

        if strengths or weaknesses or opportunities or risks_list:
            sw_rows = []
            if strengths:
                s_text = ", ".join(strengths[:4])
                sw_rows.append([
                    Paragraph(
                        f'<font size="6" color="#1B5E20"><b>\u2713 강점:</b></font> '
                        f'<font size="5.5">{s_text}</font>',
                        self.styles["KNormal"]),
                ])
            if weaknesses:
                w_text = ", ".join(weaknesses[:4])
                sw_rows.append([
                    Paragraph(
                        f'<font size="6" color="#B71C1C"><b>\u2717 약점:</b></font> '
                        f'<font size="5.5">{w_text}</font>',
                        self.styles["KNormal"]),
                ])
            if opportunities:
                o_text = ", ".join(opportunities[:3])
                sw_rows.append([
                    Paragraph(
                        f'<font size="6" color="#0D47A1"><b>\u25b2 기회:</b></font> '
                        f'<font size="5.5">{o_text}</font>',
                        self.styles["KNormal"]),
                ])
            if risks_list:
                r_text = ", ".join(risks_list[:3])
                sw_rows.append([
                    Paragraph(
                        f'<font size="6" color="#E65100"><b>\u25bc 위험:</b></font> '
                        f'<font size="5.5">{r_text}</font>',
                        self.styles["KNormal"]),
                ])

            sw_t = Table(sw_rows, colWidths=[270 * mm])
            sw_style = [
                ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#1B4F72")),
            ]
            for i, row_data in enumerate(sw_rows):
                if i == 0 and strengths:
                    sw_style.append(("BACKGROUND", (0, i), (-1, i),
                                     colors.HexColor("#E8F5E9")))
                elif (i == 1 and weaknesses) or (i == 0 and not strengths and weaknesses):
                    sw_style.append(("BACKGROUND", (0, i), (-1, i),
                                     colors.HexColor("#FFEBEE")))
                else:
                    sw_style.append(("BACKGROUND", (0, i), (-1, i),
                                     colors.HexColor("#F5F5F5")))
            sw_t.setStyle(TableStyle(sw_style))
            elements.append(sw_t)

        return elements
