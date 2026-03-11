"""Annual + TTM Financial Analysis Report Generator

PDFReportBase를 상속하여 연도별 재무 분석 리포트를 생성하는 클래스.
pdf_report.py에서 연간 분석 관련 메서드를 분리하여 독립 모듈로 구성.
"""

from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, PageBreak
from pdf_report_base import (
    PDFReportBase, FONT_NAME, HEADER_BG, SECTION_BG, SECTION_FG,
    SUB_HEADER_BG, BORDER_COLOR, LIGHT_ROW_BG,
    TTM_HEADER_BG, TTM_CELL_BG,
    TREND_POSITIVE_BG, TREND_NEGATIVE_BG, TREND_NEUTRAL_BG,
    RISK_LEVEL_BG,
    _fmt_num, _fmt_amount, _fmt_price,
    assess_metric, LEVEL_OK,
    SITUATION_EMOJI, TREND_UP, TREND_DOWN, TREND_FLAT, TREND_VOLATILE, TREND_RECOVERY,
)


class AnnualReportGenerator(PDFReportBase):
    """연도별 재무 분석 리포트 생성기"""

    def generate(self, report_data: dict):
        """메인 PDF 생성 (standalone)"""
        elements = self.build_elements(report_data)
        self._build_doc(elements)

    def build_elements(self, report_data: dict) -> list:
        """리포트 요소 목록 반환 (합성 리포트용)"""
        self._prepare(report_data)

        info = report_data.get("company_info", {})
        corp_name = info.get("corp_name", "")
        stock_code = info.get("stock_code", "")

        elements = []

        # Page 1: Title + Overview + Main layout
        elements.append(self._build_title_bar(
            report_data,
            title_text=f"연도별 재무 분석 - {corp_name} ({stock_code})"
        ))
        elements.append(Spacer(1, 2 * mm))
        elements.append(self._build_overview_section(report_data))
        elements.append(Spacer(1, 2 * mm))
        elements.append(self._build_main_layout(report_data))

        # Pages 2-3: Trend analysis (if available)
        trend_result = report_data.get("trend_analysis")
        if trend_result and trend_result.get("details"):
            elements.append(PageBreak())
            trend_elements = self._build_trend_page(report_data)
            if trend_elements:
                elements.extend(trend_elements)
            elements.append(PageBreak())
            trend_page2 = self._build_trend_page2(report_data)
            if trend_page2:
                elements.extend(trend_page2)

        # Page 4: S-RIM explanation
        elements.append(PageBreak())
        elements.append(self._build_srim_explanation(report_data))

        return elements

    # ── 메인 레이아웃 (좌우 2단) ─────────────────────────────

    def _build_main_layout(self, data: dict) -> Table:
        left_elements = []
        right_elements = []

        # 좌측: 손익계산서, 수익성, 재무상태표, 현금흐름, 밸류에이션
        left_elements.append(self._build_income_statement(data))
        left_elements.append(Spacer(1, 1 * mm))
        left_elements.append(self._build_profitability(data))
        left_elements.append(Spacer(1, 1 * mm))
        left_elements.append(self._build_balance_sheet_summary(data))
        left_elements.append(Spacer(1, 1 * mm))
        left_elements.append(self._build_cash_flow(data))
        left_elements.append(Spacer(1, 1 * mm))
        left_elements.append(self._build_valuation(data))

        # 우측: 운전자본, CAPEX, 현금성, 차입금, 순차입금, 컨센서스
        right_elements.append(self._build_working_capital(data))
        right_elements.append(Spacer(1, 1 * mm))
        right_elements.append(self._build_capex_assets(data))
        right_elements.append(Spacer(1, 1 * mm))
        right_elements.append(self._build_cash_assets(data))
        right_elements.append(Spacer(1, 1 * mm))
        right_elements.append(self._build_borrowings(data))
        right_elements.append(Spacer(1, 1 * mm))
        right_elements.append(self._build_net_debt(data))
        right_elements.append(Spacer(1, 1 * mm))
        right_elements.append(self._build_consensus(data))

        # 좌우 합치기
        left_cell = Table([[e] for e in left_elements], colWidths=[145 * mm])
        left_cell.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))

        right_cell = Table([[e] for e in right_elements], colWidths=[120 * mm])
        right_cell.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))

        main = Table([[left_cell, right_cell]], colWidths=[148 * mm, 122 * mm])
        main.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        return main

    # ── 손익계산서 ───────────────────────────────────────────

    def _build_income_statement(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        fs = data.get("financial_summary", {})
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_fs = ttm.get("financial_summary", {})
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm_fs)

        header = ["[손익계산서]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]

        items = [
            ("매출액", "매출액", "영업수익"),
            ("영업이익", "영업이익"),
            ("세전이익", "법인세비용차감전계속사업이익"),
            ("당기순이익", "당기순이익"),
            ("당기순이익(지배)", "당기순이익(지배)"),
        ]

        for item in items:
            label = item[0]
            keys = item[1:]
            row = [label]
            for y in years:
                val = 0
                for key in keys:
                    val = self._find_account(fs, y, key)
                    if val:
                        break
                row.append(self._fmt_amt(val))
            if show_ttm:
                ttm_val = 0
                for key in keys:
                    ttm_val = self._find_in_dict(ttm_fs, key)
                    if ttm_val:
                        break
                row.append(self._fmt_amt(ttm_val))
            rows.append(row)

        # YoY 성장률
        growth_items = [
            ("매출성장률(%)", "매출성장률(%)"),
            ("영업이익성장률(%)", "영업이익성장률(%)"),
            ("순이익성장률(%)", "순이익성장률(%)"),
        ]
        highlights = {}
        for row_idx_offset, (label, key) in enumerate(growth_items):
            row = [label]
            row_idx = len(items) + 1 + row_idx_offset
            for col_idx, y in enumerate(years, 1):
                val = derived.get(y, {}).get(key)
                if val is not None:
                    row.append(f"{val:+.1f}")
                else:
                    row.append("-")
            if show_ttm:
                ttm_val = ttm_derived.get(key)
                if ttm_val is not None:
                    row.append(f"{ttm_val:+.1f}")
                else:
                    row.append("-")
            rows.append(row)

        desc = f"단위: {self._unit_label()} | 성장률: 전년대비 YoY(%)"
        if show_ttm and ttm.get("annualized"):
            used_q = ttm.get("quarters_used", [])
            desc += f"\n* TTM: {'+'.join(used_q)} 기준 연환산(×{ttm['ann_factor']:.2f})"
        return self._section_table(rows, len(years), highlights=highlights, description=desc, ttm_column=show_ttm)

    # ── 수익성 ───────────────────────────────────────────────

    def _build_profitability(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[수익성]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]
        highlights = {}

        metric_names = ["영업이익률(%)", "ROE(%)", "ROA(%)", "레버리지비율"]

        for row_idx, label in enumerate(metric_names, 1):
            row = [label]
            for col_idx, y in enumerate(years, 1):
                val = derived.get(y, {}).get(label, 0)
                if label == "레버리지비율":
                    row.append(_fmt_num(val, "ratio") if val else "-")
                else:
                    row.append(_fmt_num(val, "pct") if val else "-")
                # 위험 평가
                if val:
                    risk = assess_metric(label, val)
                    if risk["level"] != LEVEL_OK:
                        highlights[(row_idx, col_idx)] = risk
            if show_ttm:
                ttm_val = ttm_derived.get(label, 0)
                if label == "레버리지비율":
                    row.append(_fmt_num(ttm_val, "ratio") if ttm_val else "-")
                else:
                    row.append(_fmt_num(ttm_val, "pct") if ttm_val else "-")
            rows.append(row)

        desc = "OPM: 매출대비 영업이익 | ROE: 자기자본이익률(8%↑ 양호) | ROA: 총자산이익률"
        return self._section_table(rows, len(years), highlights=highlights, description=desc, ttm_column=show_ttm)

    # ── 재무상태표 요약 ──────────────────────────────────────

    def _build_balance_sheet_summary(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        fs = data.get("financial_summary", {})
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_bs = ttm.get("balance_sheet", {})
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[재무상태표]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]

        items = [
            ("자본총계", "자본총계"),
            ("자본총계(지배)", "자본총계(지배)"),
            ("이자발생부채", "이자발생부채"),
        ]

        for label, key in items:
            row = [label]
            for y in years:
                val = self._find_account(fs, y, key)
                row.append(self._fmt_amt(val))
            if show_ttm:
                ttm_val = self._find_in_dict(ttm_bs, key)
                row.append(self._fmt_amt(ttm_val))
            rows.append(row)

        # 순차입금 (derived)
        row = ["순차입금"]
        for y in years:
            val = derived.get(y, {}).get("순차입금", 0)
            row.append(self._fmt_amt(val))
        if show_ttm:
            row.append(self._fmt_amt(ttm_derived.get("순차입금", 0)))
        rows.append(row)

        desc = "자본총계: 순자산(자산-부채) | 순차입금: 이자발생부채-현금성자산 (음수=순현금)"
        return self._section_table(rows, len(years), description=desc, ttm_column=show_ttm)

    # ── 현금흐름표 ───────────────────────────────────────────

    def _build_cash_flow(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[현금흐름표]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]
        highlights = {}

        cf_items = [
            ("영업활동CF", "영업활동CF"),
            ("투자활동CF", "투자활동CF"),
            ("재무활동CF", "재무활동CF"),
            ("CAPEX", "CAPEX"),
            ("FCF", "FCF"),
        ]

        for row_idx, (label, key) in enumerate(cf_items, 1):
            row = [label]
            for col_idx, y in enumerate(years, 1):
                val = derived.get(y, {}).get(key, 0)
                row.append(self._fmt_amt(val))
                if key in ("영업활동CF", "FCF") and val:
                    risk = assess_metric(key, val)
                    if risk["level"] != LEVEL_OK:
                        highlights[(row_idx, col_idx)] = risk
            if show_ttm:
                row.append(self._fmt_amt(ttm_derived.get(key, 0)))
            rows.append(row)

        # PFCR
        pfcr_row_idx = len(cf_items) + 1
        row = ["PFCR"]
        for col_idx, y in enumerate(years, 1):
            val = derived.get(y, {}).get("PFCR", 0)
            row.append(_fmt_num(val, "ratio") if val else "-")
            if val:
                risk = assess_metric("PFCR", val)
                if risk["level"] != LEVEL_OK:
                    highlights[(pfcr_row_idx, col_idx)] = risk
        if show_ttm:
            ttm_pfcr = ttm_derived.get("PFCR", 0)
            row.append(_fmt_num(ttm_pfcr, "ratio") if ttm_pfcr else "-")
        rows.append(row)

        desc = "영업CF: 본업 현금창출 | FCF: 잉여현금(배당/투자여력) | PFCR: 시총/FCF(낮을수록 저평가)"
        return self._section_table(rows, len(years), highlights=highlights, description=desc, ttm_column=show_ttm)

    # ── 밸류에이션 ───────────────────────────────────────────

    def _build_valuation(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[밸류에이션]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]
        highlights = {}

        pu = self._price_unit()
        val_items = [
            ("PER(배)", "PER"),
            ("PBR(배)", "PBR"),
            (f"EPS({pu})", "EPS"),
            (f"BPS({pu})", "BPS"),
        ]

        for row_idx, (label, key) in enumerate(val_items, 1):
            row = [label]
            for col_idx, y in enumerate(years, 1):
                val = derived.get(y, {}).get(key, 0)
                if key in ("PER", "PBR"):
                    row.append(_fmt_num(val, "ratio") if val else "-")
                else:
                    row.append(self._fmt_prc(val) if val else "-")
                # 위험 평가 (PER, PBR만)
                if key in ("PER", "PBR") and val:
                    risk = assess_metric(label, val)
                    if risk["level"] != LEVEL_OK:
                        highlights[(row_idx, col_idx)] = risk
            if show_ttm:
                ttm_val = ttm_derived.get(key, 0)
                if key in ("PER", "PBR"):
                    row.append(_fmt_num(ttm_val, "ratio") if ttm_val else "-")
                else:
                    row.append(self._fmt_prc(ttm_val) if ttm_val else "-")
            rows.append(row)

        desc = ("PER: 이익대비 주가(낮을수록 저평가) | PBR: 순자산대비 주가(1 미만 저평가)\n"
                "EPS: 주당순이익(1주당 벌어들이는 이익) | BPS: 주당순자산(1주당 순자산가치)")
        return self._section_table(rows, len(years), highlights=highlights, description=desc, ttm_column=show_ttm)

    # ── 운전자본 ─────────────────────────────────────────────

    def _build_working_capital(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        bs = data.get("balance_sheet_detail", {})
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_bs = ttm.get("balance_sheet", {})
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[운전자본]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]

        wc_items = [
            ("매출채권", "매출채권및기타유동채권", "매출채권및기타채권", "매출채권"),
            ("재고자산", "재고자산"),
            ("매입채무", "매입채무및기타유동채무", "매입채무및기타채무", "매입채무"),
        ]

        for item_names in wc_items:
            label = item_names[0]
            row = [label]
            for y in years:
                val = 0
                for key in item_names:
                    val = self._find_in_dict(bs.get(y, {}), key)
                    if val:
                        break
                row.append(self._fmt_amt(val))
            if show_ttm:
                ttm_val = 0
                for key in item_names:
                    ttm_val = self._find_in_dict(ttm_bs, key)
                    if ttm_val:
                        break
                row.append(self._fmt_amt(ttm_val))
            rows.append(row)

        # 운전자본 합계
        row = ["운전자본"]
        for y in years:
            val = derived.get(y, {}).get("운전자본", 0)
            row.append(self._fmt_amt(val))
        if show_ttm:
            row.append(self._fmt_amt(ttm_derived.get("운전자본", 0)))
        rows.append(row)

        # 매출대비 비율
        row = ["매출대비 비율(%)"]
        for y in years:
            val = derived.get(y, {}).get("운전자본비율(%)", 0)
            row.append(_fmt_num(val, "pct") if val else "-")
        if show_ttm:
            ttm_val = ttm_derived.get("운전자본비율(%)", 0)
            row.append(_fmt_num(ttm_val, "pct") if ttm_val else "-")
        rows.append(row)

        desc = "매출채권+재고-매입채무. 영업에 묶인 운영자금. 매출대비 비율 높으면 자금 효율 낮음"
        return self._section_table(rows, len(years), description=desc, ttm_column=show_ttm)

    # ── CAPEX 자산 ───────────────────────────────────────────

    def _build_capex_assets(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        bs = data.get("balance_sheet_detail", {})
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_bs = ttm.get("balance_sheet", {})
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[CAPEX 자산]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]

        items = [
            ("유형자산", "유형자산"),
            ("무형자산", "무형자산", "영업권이외의무형자산"),
            ("자산총계", "자산총계"),
        ]

        for item_names in items:
            label = item_names[0]
            row = [label]
            for y in years:
                val = 0
                for key in item_names:
                    val = self._find_in_dict(bs.get(y, {}), key)
                    if val:
                        break
                row.append(self._fmt_amt(val))
            if show_ttm:
                ttm_val = 0
                for key in item_names:
                    ttm_val = self._find_in_dict(ttm_bs, key)
                    if ttm_val:
                        break
                row.append(self._fmt_amt(ttm_val))
            rows.append(row)

        # 비중
        for label_key in [("유형자산 비중(%)", "유형자산비중(%)"), ("무형자산 비중(%)", "무형자산비중(%)")]:
            row = [label_key[0]]
            for y in years:
                val = derived.get(y, {}).get(label_key[1], 0)
                row.append(_fmt_num(val, "pct") if val else "-")
            if show_ttm:
                ttm_val = ttm_derived.get(label_key[1], 0)
                row.append(_fmt_num(ttm_val, "pct") if ttm_val else "-")
            rows.append(row)

        return self._section_table(rows, len(years), ttm_column=show_ttm)

    # ── 현금성자산 ───────────────────────────────────────────

    def _build_cash_assets(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        bs = data.get("balance_sheet_detail", {})
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_bs = ttm.get("balance_sheet", {})
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[현금성자산]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]

        items = [
            ("현금및현금성자산", "현금및현금성자산"),
            ("단기금융자산", "단기금융상품", "단기금융자산"),
            ("당기손익-공정가치금융자산", "당기손익-공정가치측정금융자산", "단기투자자산"),
        ]

        for item_names in items:
            label = item_names[0]
            row = [label]
            for y in years:
                val = 0
                for key in item_names:
                    val = self._find_in_dict(bs.get(y, {}), key)
                    if val:
                        break
                row.append(self._fmt_amt(val))
            if show_ttm:
                ttm_val = 0
                for key in item_names:
                    ttm_val = self._find_in_dict(ttm_bs, key)
                    if ttm_val:
                        break
                row.append(self._fmt_amt(ttm_val))
            rows.append(row)

        # 합계
        row = ["합계"]
        for y in years:
            val = derived.get(y, {}).get("현금성자산합계", 0)
            row.append(self._fmt_amt(val))
        if show_ttm:
            row.append(self._fmt_amt(ttm_derived.get("현금성자산합계", 0)))
        rows.append(row)

        return self._section_table(rows, len(years), ttm_column=show_ttm)

    # ── 차입금내역 ───────────────────────────────────────────

    def _build_borrowings(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        bs = data.get("balance_sheet_detail", {})
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_bs = ttm.get("balance_sheet", {})
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[차입금내역]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]
        highlights = {}

        items = [
            ("단기차입금", "단기차입금"),
            ("유동성장기부채", "유동성장기부채"),
            ("사채", "사채"),
            ("장기차입금", "장기차입금"),
        ]

        for item_names in items:
            label = item_names[0]
            row = [label]
            for y in years:
                val = 0
                for key in item_names:
                    val = self._find_in_dict(bs.get(y, {}), key)
                    if val:
                        break
                row.append(self._fmt_amt(val))
            if show_ttm:
                ttm_val = 0
                for key in item_names:
                    ttm_val = self._find_in_dict(ttm_bs, key)
                    if ttm_val:
                        break
                row.append(self._fmt_amt(ttm_val))
            rows.append(row)

        # 이자발생부채
        row = ["이자발생부채"]
        for y in years:
            val = derived.get(y, {}).get("이자발생부채계산", 0)
            row.append(self._fmt_amt(val))
        if show_ttm:
            row.append(self._fmt_amt(ttm_derived.get("이자발생부채계산", 0)))
        rows.append(row)

        # 단기채/장기채 비중
        row_idx_short = len(items) + 2  # items + header + 이자발생부채
        for i, (label, key) in enumerate([("단기채비중(%)", "단기채비중(%)"), ("장기채비중(%)", "장기채비중(%)")]):
            row = [label]
            for col_idx, y in enumerate(years, 1):
                val = derived.get(y, {}).get(key, 0)
                row.append(_fmt_num(val, "pct") if val else "-")
                # 단기채비중만 위험 평가
                if key == "단기채비중(%)" and val:
                    risk = assess_metric("단기채비중(%)", val)
                    if risk["level"] != LEVEL_OK:
                        highlights[(row_idx_short + i, col_idx)] = risk
            if show_ttm:
                ttm_val = ttm_derived.get(key, 0)
                row.append(_fmt_num(ttm_val, "pct") if ttm_val else "-")
            rows.append(row)

        desc = "단기차입 비중 높으면 유동성 리스크. 장기채 비중 높을수록 안정적"
        return self._section_table(rows, len(years), highlights=highlights, description=desc, ttm_column=show_ttm)

    # ── 순차입금현황 ─────────────────────────────────────────

    def _build_net_debt(self, data: dict) -> Table:
        years = data.get("display_years", data.get("years", []))
        derived = data.get("derived", {})
        ttm = data.get("ttm") or {}
        ttm_derived = data.get("derived", {}).get("_ttm", {})
        ttm_label = ttm.get("ttm_label", "TTM")
        show_ttm = bool(ttm.get("financial_summary"))

        header = ["[순차입금현황]"] + [f"{y}" for y in years]
        if show_ttm:
            header.append(ttm_label)
        rows = [header]

        for label in ["이자발생부채계산", "현금성자산합계", "순차입금"]:
            row = [label]
            for y in years:
                val = derived.get(y, {}).get(label, 0)
                row.append(self._fmt_amt(val))
            if show_ttm:
                row.append(self._fmt_amt(ttm_derived.get(label, 0)))
            rows.append(row)

        return self._section_table(rows, len(years), ttm_column=show_ttm)

    # ── 컨센서스 ───────────────────────────────────────────────

    def _build_consensus(self, data: dict) -> Table:
        consensus = data.get("consensus", {})
        stock = data.get("stock_data", {})
        info = data.get("company_info", {})
        is_intl = info.get("is_international", False)

        source_label = "Yahoo Finance" if is_intl else "네이버 증권"
        rows = [
            [f"[컨센서스] 출처: {source_label}", "", ""]
        ]

        target = consensus.get("target_price", 0)
        opinion = consensus.get("opinion", 0)
        price = stock.get("price", 0)
        pu = self._price_unit()

        if target:
            upside = round((target - price) / price * 100, 1) if price else 0
            rows.append([f"목표주가({pu})", self._fmt_prc(target), f"(괴리율 {upside:+.1f}%)"])
        if opinion:
            rows.append(["투자의견", f"{opinion:.1f}", ""])

        items = consensus.get("items", [])
        if items:
            rows.append(["", "현재", "이전"])
            for item in items[:6]:
                label = item.get("label", "")
                values = item.get("values", [])
                current = values[0] if len(values) > 0 else "-"
                prev = values[-1] if len(values) > 1 else "-"
                rows.append([label, current, prev])

        if len(rows) <= 1:
            rows.append(["데이터 없음", "", ""])

        col_widths = [35 * mm, 30 * mm, 30 * mm]
        styled_rows = []
        for i, row in enumerate(rows):
            styled_row = []
            for cell in row:
                if i == 0:
                    styled_row.append(Paragraph(str(cell), self.styles["KHeader"]))
                else:
                    styled_row.append(Paragraph(str(cell), self.styles["KSmall"]))
            styled_rows.append(styled_row)

        t = Table(styled_rows, colWidths=col_widths)
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), SECTION_BG),
            ("TEXTCOLOR", (0, 0), (-1, 0), SECTION_FG),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6.5),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
        ]
        for i in range(2, len(styled_rows), 2):
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_ROW_BG))
        t.setStyle(TableStyle(style_cmds))
        return t

    # ── 추이 분석 페이지 ────────────────────────────────────────

    def _build_trend_page(self, data: dict) -> list:
        """재무 추이 분석 페이지"""
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
                f"재무 추이 분석 (Financial Trend Analysis) - {corp_name}",
                self.styles["KTitle"],
            ),
        ]]
        t = Table(title_data, colWidths=[270 * mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1A237E")),
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

        # 종합 진단 상단: 상황 분류 + 핵심 지표
        diag_rows = []

        # 상황 분류 헤더
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
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#1A237E")),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.HexColor("#1A237E")),
        ]))
        elements.append(diag_t)
        elements.append(Spacer(1, 2 * mm))

        # ── 종합 요약 ──
        summary_rows = [[
            Paragraph(
                f'<font size="7"><b>[종합 진단]</b></font>',
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
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A237E")),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#E8EAF6")),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#1A237E")),
        ]))
        elements.append(summary_t)
        elements.append(Spacer(1, 3 * mm))

        # ── 영역별 상세 분석 테이블 ──
        detail_rows = []

        # 헤더
        detail_rows.append([
            Paragraph(
                '<font color="white"><b>분석 영역</b></font>',
                self.styles["KHeader"],
            ),
            Paragraph(
                '<font color="white"><b>추세</b></font>',
                self.styles["KHeader"],
            ),
            Paragraph(
                '<font color="white"><b>점수</b></font>',
                self.styles["KHeader"],
            ),
            Paragraph(
                '<font color="white"><b>진단</b></font>',
                self.styles["KHeader"],
            ),
            Paragraph(
                '<font color="white"><b>상세 분석</b></font>',
                self.styles["KHeader"],
            ),
        ])

        trend_labels = {
            TREND_UP: "↑ 상승",
            TREND_DOWN: "↓ 하락",
            TREND_FLAT: "→ 횡보",
            TREND_VOLATILE: "↕ 변동",
            TREND_RECOVERY: "↗ 회복",
        }

        for d in details:
            category = d.get("category", "")
            trend_dir = d.get("trend", TREND_FLAT)
            score = d.get("score", 0)
            title = d.get("title", "")
            comment = d.get("comment", "")
            sub_items = d.get("sub_items", [])

            trend_text = trend_labels.get(trend_dir, "→ 횡보")

            # 점수 → 시각적 표현
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

            # sub_items를 comment 아래에 추가
            full_comment = comment
            if sub_items:
                si_parts = []
                for si in sub_items:
                    label_s = si.get("label", "")
                    value_s = si.get("value", "")
                    si_parts.append(f"{label_s} {value_s}")
                sub_line = " | ".join(si_parts)
                full_comment += f'<br/><font size="5" color="#555555">  ▸ {sub_line}</font>'

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
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A237E")),
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

        # 행별 배경색
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
                '※ 본 분석은 과거 재무데이터 기반 자동 진단이며, 투자 권유가 아닙니다. '
                '업종 특성, 경영 환경 변화, 일회성 요인 등은 반영되지 않습니다. '
                '투자 결정 시 반드시 다양한 정보를 종합적으로 고려하시기 바랍니다.'
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
                        f'<font size="6" color="#1B5E20"><b>✓ 강점:</b></font> '
                        f'<font size="5.5">{s_text}</font>',
                        self.styles["KNormal"]),
                ])
            if weaknesses:
                w_text = ", ".join(weaknesses[:4])
                sw_rows.append([
                    Paragraph(
                        f'<font size="6" color="#B71C1C"><b>✗ 약점:</b></font> '
                        f'<font size="5.5">{w_text}</font>',
                        self.styles["KNormal"]),
                ])
            if opportunities:
                o_text = ", ".join(opportunities[:3])
                sw_rows.append([
                    Paragraph(
                        f'<font size="6" color="#0D47A1"><b>▲ 기회:</b></font> '
                        f'<font size="5.5">{o_text}</font>',
                        self.styles["KNormal"]),
                ])
            if risks_list:
                r_text = ", ".join(risks_list[:3])
                sw_rows.append([
                    Paragraph(
                        f'<font size="6" color="#E65100"><b>▼ 위험:</b></font> '
                        f'<font size="5.5">{r_text}</font>',
                        self.styles["KNormal"]),
                ])

            sw_t = Table(sw_rows, colWidths=[270 * mm])
            sw_style = [
                ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#1A237E")),
            ]
            # 행별 배경
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

    # ── 추이 분석 상세 페이지 (page 2) ─────────────────────────

    def _build_trend_page2(self, data: dict) -> list:
        """추이 분석 상세 페이지 — 연도별 테이블, DuPont, 투자 체크리스트"""
        trend = data.get("trend_analysis", {})
        if not trend:
            return []

        elements = []
        info = data.get("company_info", {})
        corp_name = info.get("corp_name", "")

        # ── 제목바 ──
        title_data = [[
            Paragraph(
                f"재무 추이 분석 상세 (Detailed Analysis) - {corp_name}",
                self.styles["KTitle"],
            ),
        ]]
        t = Table(title_data, colWidths=[270 * mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1A237E")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 3 * mm))

        # ── 연도별 핵심 지표 추이 테이블 ──
        yearly_table = trend.get("yearly_table")
        if yearly_table:
            yt_years = yearly_table.get("years", [])
            yt_rows_data = yearly_table.get("rows", [])

            if yt_years and yt_rows_data:
                # 섹션 헤더
                sec_hdr = [[Paragraph(
                    '<font size="7" color="white"><b>연도별 핵심 지표 추이</b></font>',
                    self.styles["KTitle"],
                )]]
                sec_t = Table(sec_hdr, colWidths=[270 * mm])
                sec_t.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#283593")),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]))
                elements.append(sec_t)

                # 헤더 행
                header_row = [Paragraph('<font color="white"><b>지표</b></font>',
                                        self.styles["KHeader"])]
                for y in yt_years:
                    header_row.append(Paragraph(
                        f'<font color="white"><b>{y}</b></font>',
                        self.styles["KHeader"]))

                data_rows = [header_row]
                n_cols = len(yt_years)
                label_w = 40 * mm
                val_w = (270 * mm - label_w) / n_cols if n_cols else 40 * mm

                for row_info in yt_rows_data:
                    label = row_info.get("label", "")
                    values = row_info.get("values", [])
                    row = [Paragraph(f'<b>{label}</b>', self.styles["KSmall"])]
                    for v in values:
                        if v is None or v == 0:
                            row.append(Paragraph("-", self.styles["KSmall"]))
                        elif isinstance(v, float):
                            row.append(Paragraph(f"{v:,.1f}", self.styles["KSmall"]))
                        else:
                            row.append(Paragraph(f"{v:,.0f}", self.styles["KSmall"]))
                    # 패딩
                    while len(row) < n_cols + 1:
                        row.append(Paragraph("-", self.styles["KSmall"]))
                    data_rows.append(row)

                col_widths = [label_w] + [val_w] * n_cols
                yt_t = Table(data_rows, colWidths=col_widths)

                yt_style = [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A237E")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
                    ("FONTSIZE", (0, 0), (-1, -1), 6),
                    ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                    ("ALIGN", (0, 0), (0, -1), "LEFT"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                    ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
                ]
                # 교대 행 색상
                for i in range(1, len(data_rows)):
                    bg = colors.HexColor("#F5F5F5") if i % 2 == 0 else colors.white
                    yt_style.append(("BACKGROUND", (0, i), (-1, i), bg))

                yt_t.setStyle(TableStyle(yt_style))
                elements.append(yt_t)
                elements.append(Spacer(1, 3 * mm))

        # ── DuPont ROE 분해 ──
        dupont = trend.get("dupont")
        if dupont and dupont.get("years"):
            dp_years = dupont["years"]
            dp_nm = dupont["net_margin"]
            dp_at = dupont["asset_turnover"]
            dp_em = dupont["equity_multiplier"]
            dp_roe = dupont["roe"]
            dp_driver = dupont.get("main_driver", "")
            dp_comment = dupont.get("comment", "")

            # 섹션 헤더
            sec_hdr = [[Paragraph(
                '<font size="7" color="white"><b>DuPont ROE 분해  '
                '(ROE = 순이익률 × 자산회전율 × 재무레버리지)</b></font>',
                self.styles["KTitle"],
            )]]
            sec_t = Table(sec_hdr, colWidths=[270 * mm])
            sec_t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#283593")),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]))
            elements.append(sec_t)

            n_dp = len(dp_years)
            dp_label_w = 45 * mm
            dp_val_w = (270 * mm - dp_label_w) / n_dp if n_dp else 40 * mm

            dp_header = [Paragraph('<font color="white"><b>팩터</b></font>',
                                   self.styles["KHeader"])]
            for y in dp_years:
                dp_header.append(Paragraph(
                    f'<font color="white"><b>{y}</b></font>',
                    self.styles["KHeader"]))

            dp_rows = [dp_header]
            # 순이익률
            r1 = [Paragraph('<b>순이익률(%)</b>', self.styles["KSmall"])]
            for v in dp_nm:
                r1.append(Paragraph(f"{v:.1f}", self.styles["KSmall"]))
            dp_rows.append(r1)
            # 자산회전율
            r2 = [Paragraph('<b>자산회전율(배)</b>', self.styles["KSmall"])]
            for v in dp_at:
                r2.append(Paragraph(f"{v:.3f}", self.styles["KSmall"]))
            dp_rows.append(r2)
            # 재무레버리지
            r3 = [Paragraph('<b>재무레버리지(배)</b>', self.styles["KSmall"])]
            for v in dp_em:
                r3.append(Paragraph(f"{v:.2f}", self.styles["KSmall"]))
            dp_rows.append(r3)
            # ROE
            r4 = [Paragraph('<b>ROE(%)</b>', self.styles["KSmall"])]
            for v in dp_roe:
                r4.append(Paragraph(f"{v:.1f}", self.styles["KSmall"]))
            dp_rows.append(r4)

            dp_col_widths = [dp_label_w] + [dp_val_w] * n_dp
            dp_t = Table(dp_rows, colWidths=dp_col_widths)
            dp_t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A237E")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
                ("FONTSIZE", (0, 0), (-1, -1), 6.5),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F5F5F5")),
                ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#F5F5F5")),
                ("BACKGROUND", (0, 4), (-1, 4), colors.HexColor("#E8EAF6")),
            ]))
            elements.append(dp_t)

            # DuPont 코멘트
            if dp_comment:
                dp_comment_row = [[Paragraph(
                    f'<font size="6">  <b>주 변동 요인:</b> {dp_driver} — {dp_comment}</font>',
                    self.styles["KNormal"],
                )]]
                dp_ct = Table(dp_comment_row, colWidths=[270 * mm])
                dp_ct.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#E8EAF6")),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ]))
                elements.append(dp_ct)
            elements.append(Spacer(1, 3 * mm))

        # ── 투자 체크리스트 ──
        checklist = trend.get("checklist", [])
        if checklist:
            # 섹션 헤더
            sec_hdr = [[Paragraph(
                '<font size="7" color="white"><b>투자 체크리스트</b></font>',
                self.styles["KTitle"],
            )]]
            sec_t = Table(sec_hdr, colWidths=[270 * mm])
            sec_t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#283593")),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]))
            elements.append(sec_t)

            cl_rows = []
            cl_header = [
                Paragraph('<font color="white"><b>체크 항목</b></font>',
                          self.styles["KHeader"]),
                Paragraph('<font color="white"><b>상태</b></font>',
                          self.styles["KHeader"]),
                Paragraph('<font color="white"><b>상세</b></font>',
                          self.styles["KHeader"]),
            ]
            cl_rows.append(cl_header)

            status_colors = {
                "양호": "#2ECC71",
                "적정": "#3498DB",
                "저평가": "#2ECC71",
                "보통": "#7F8C8D",
                "주의": "#E67E22",
                "부진": "#E74C3C",
                "고평가": "#E74C3C",
            }

            for item in checklist:
                q = item.get("question", "")
                status = item.get("status", "보통")
                detail = item.get("detail", "")
                s_color = status_colors.get(status, "#7F8C8D")

                # 상태 아이콘
                if status in ("양호", "적정", "저평가"):
                    icon = "✓"
                elif status in ("부진", "고평가"):
                    icon = "✗"
                else:
                    icon = "△"

                cl_rows.append([
                    Paragraph(f'{q}', self.styles["KSmall"]),
                    Paragraph(
                        f'<font color="{s_color}"><b>{icon} {status}</b></font>',
                        self.styles["KSmall"]),
                    Paragraph(detail, self.styles["KSmall"]),
                ])

            cl_t = Table(cl_rows, colWidths=[80 * mm, 30 * mm, 160 * mm])
            cl_style = [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A237E")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
                ("FONTSIZE", (0, 0), (-1, -1), 6.5),
                ("ALIGN", (1, 0), (1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
            ]
            for i in range(1, len(cl_rows)):
                bg = colors.HexColor("#F5F5F5") if i % 2 == 0 else colors.white
                cl_style.append(("BACKGROUND", (0, i), (-1, i), bg))
            cl_t.setStyle(TableStyle(cl_style))
            elements.append(cl_t)
            elements.append(Spacer(1, 3 * mm))

        # ── 면책 조항 ──
        disc_rows = [[Paragraph(
            '<font size="5" color="#999999">'
            '※ 본 분석은 과거 재무데이터 기반 자동 진단이며, 투자 권유가 아닙니다. '
            '업종 특성, 경영 환경 변화, 일회성 요인 등은 반영되지 않습니다.'
            '</font>',
            self.styles["KDesc"],
        )]]
        disc_t = Table(disc_rows, colWidths=[270 * mm])
        disc_t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FAFAFA")),
        ]))
        elements.append(disc_t)

        return elements

    # ── S-RIM 설명 패널 ──────────────────────────────────────

    def _build_srim_explanation(self, data: dict) -> Table:
        """S-RIM 모델 설명 패널 (초보자 가이드) — 상세 버전"""
        srim = data.get("srim", {})
        stock = data.get("stock_data", {})
        info = data.get("company_info", {})
        is_intl = info.get("is_international", False)
        price = stock.get("price", 0)
        srim_price = srim.get("srim_price", 0)
        buy_price = srim.get("buy_price", 0)
        roe_f = srim.get("roe_forecast", 0)
        coe_val = srim.get("coe_value", 8.0)
        beta = srim.get("beta", 1.0)
        rf = srim.get("risk_free_rate", 4.0)
        mrp = srim.get("market_risk_premium", 5.5)
        coe_source = srim.get("coe_source", "")
        is_capm = "CAPM" in coe_source
        capm_coe = srim.get("capm_coe", 0)
        equity = srim.get("equity", 0)

        # 판정
        if srim_price and price:
            diff_pct = (srim_price - price) / srim_price * 100
            if diff_pct > 30:
                verdict = "현재 주가가 S-RIM 적정가 대비 크게 저평가되어 있습니다. 안전마진이 충분합니다."
            elif diff_pct > 0:
                verdict = "현재 주가가 S-RIM 적정가 대비 소폭 저평가되어 있습니다."
            elif diff_pct > -20:
                verdict = "현재 주가가 S-RIM 적정가에 근접하거나 소폭 고평가되어 있습니다."
            else:
                verdict = "현재 주가가 S-RIM 적정가 대비 크게 고평가되어 있습니다. 주의가 필요합니다."
        else:
            verdict = ""
            diff_pct = 0

        # ── 설명 텍스트 구성 ──
        lines = []

        # 1) S-RIM이란?
        lines.append(("title", "S-RIM(초과이익 모델)이란?"))
        lines.append(("text", "S-RIM(Surplus Return on Investment Model)은 기업의 자본(순자산)을 기반으로,"))
        lines.append(("text", "그 기업이 '기대 수익률 이상으로 벌어들이는 초과이익'의 가치를 더해 적정 주가를 산출하는 모델입니다."))
        lines.append(("text", "PER(주가수익비율)이 순이익 기반인 데 반해, S-RIM은 자본과 수익성(ROE)을 함께 고려하여"))
        lines.append(("text", "보다 보수적이고 안정적인 가치 평가가 가능합니다."))
        lines.append(("blank", ""))

        # 2) 공식
        lines.append(("section", "계산 공식"))
        lines.append(("formula", "V = BPS + W × BPS × (ROE − COE) / (1 + COE − W)"))
        lines.append(("text", "  ① BPS(주당순자산): 자본총계(지배) / 발행주식수"))
        lines.append(("text", "  ② ROE − COE: 초과이익률. 양수이면 기업이 기대 이상으로 수익을 냄"))
        lines.append(("text", "  ③ W(초과이익 지속계수): 초과이익이 미래에 얼마나 유지되는지 (0~1)"))
        lines.append(("text", "  ④ W=1일 때 → BPS + BPS × (ROE−COE) / COE (초과이익 영구지속)"))
        lines.append(("blank", ""))

        # EPS, BPS 설명
        lines.append(("section", "EPS(주당순이익)와 BPS(주당순자산)"))
        lines.append(("text", "  EPS(Earnings Per Share, 주당순이익)"))
        lines.append(("formula", "  EPS = 당기순이익(지배) / 발행주식수"))
        lines.append(("text", "    → 1주당 벌어들이는 순이익. 기업의 수익 창출 능력을 1주 단위로 표현"))
        lines.append(("text", "    → EPS가 꾸준히 증가하면 기업의 이익 성장력이 양호한 것"))
        lines.append(("text", "    → PER = 주가 / EPS 이므로, EPS가 높을수록 같은 주가에서 PER이 낮아짐 (저평가)"))
        lines.append(("blank", ""))
        lines.append(("text", "  BPS(Book value Per Share, 주당순자산)"))
        lines.append(("formula", "  BPS = 자본총계(지배) / 발행주식수"))
        lines.append(("text", "    → 1주당 순자산가치. 기업이 청산될 경우 주주가 받을 수 있는 이론적 금액"))
        lines.append(("text", "    → S-RIM 공식의 핵심 기초값으로, 적정가 계산의 출발점"))
        lines.append(("text", "    → PBR = 주가 / BPS 이므로, PBR < 1이면 주가가 순자산가치 이하 (저평가 가능성)"))
        lines.append(("text", "    → BPS가 꾸준히 증가하면 기업이 이익을 자본에 축적하고 있다는 의미"))
        lines.append(("blank", ""))
        # 본 종목의 최신 EPS/BPS 값 표시
        derived = data.get("derived", {})
        years = data.get("years", [])
        if years and derived:
            latest_y = years[-1]
            latest_d = derived.get(latest_y, {})
            eps_val = latest_d.get("EPS", 0)
            bps_val = latest_d.get("BPS", 0)
            if eps_val or bps_val:
                lines.append(("text", f"  본 종목의 최신({latest_y}년) 수치:"))
                if eps_val:
                    lines.append(("text", f"    EPS = {self._fmt_prc(eps_val)}"))
                if bps_val:
                    lines.append(("text", f"    BPS = {self._fmt_prc(bps_val)}"))
                if price and eps_val:
                    per_calc = round(price / eps_val, 2) if eps_val != 0 else 0
                    lines.append(("text", f"    → 현재 PER = 주가({self._fmt_prc(price)}) / EPS({self._fmt_prc(eps_val)}) = {per_calc:.1f}배"))
                if price and bps_val:
                    pbr_calc = round(price / bps_val, 2) if bps_val != 0 else 0
                    lines.append(("text", f"    → 현재 PBR = 주가({self._fmt_prc(price)}) / BPS({self._fmt_prc(bps_val)}) = {pbr_calc:.1f}배"))
                lines.append(("blank", ""))

        # W 지속계수 설명
        w_buy = srim.get("w_buy", 0.5)
        w_fair = srim.get("w_fair", 1.0)

        lines.append(("section", "초과이익 지속계수(W)란?"))
        lines.append(("text", "  S-RIM의 핵심 변수로, 기업의 초과이익(ROE > COE)이 미래에 얼마나 오래"))
        lines.append(("text", "  지속될 것인지를 0~1 사이의 계수로 표현합니다."))
        lines.append(("blank", ""))
        lines.append(("text", "  W = 1.0: 초과이익이 영구적으로 지속 (가장 낙관적)"))
        lines.append(("text", "    → 브랜드 파워, 독점적 시장 지위, 특허 등으로 경쟁 우위가 유지되는 기업"))
        lines.append(("text", "  W = 0.9: 초과이익이 상당 기간 지속되나 점진적으로 감소"))
        lines.append(("text", "    → 대부분의 우량 기업에 적용 가능한 중립적 가정"))
        lines.append(("text", "  W = 0.8: 초과이익이 빠르게 감소"))
        lines.append(("text", "    → 경쟁 심화, 기술 변화 등으로 수익성이 빠르게 하락할 수 있는 기업"))
        lines.append(("blank", ""))
        lines.append(("section", "W값에 따른 가격 산출 체계"))
        lines.append(("text", f"  매수시작가 = W={w_buy}(비관적) 기준"))
        lines.append(("text", "    → '최악의 경우에도 이 가격이면 안전하다'는 보수적 진입점"))
        lines.append(("text", f"  적정가 = W={w_fair}(낙관적) 기준"))
        lines.append(("text", "    → 초과이익이 영구 지속된다는 낙관적 가정의 공정가치"))
        lines.append(("blank", ""))

        # 실제 계산값
        lines.append(("section", "본 종목의 W 시나리오별 가격"))
        if is_intl:
            lines.append(("text", f"  매수시작가(W={w_buy}): {self._fmt_prc(buy_price)}"))
            lines.append(("text", f"  적정가(W={w_fair}): {self._fmt_prc(srim_price)}"))
        else:
            lines.append(("text", f"  매수시작가(W={w_buy}): {buy_price:,}원"))
            lines.append(("text", f"  적정가(W={w_fair}): {srim_price:,}원"))
        lines.append(("text", "  (W값은 사용자 설정에 따라 변경 가능)"))
        lines.append(("blank", ""))

        # 3) 핵심 입력값
        lines.append(("section", "핵심 입력값"))
        lines.append(("text", f"  ROE(자기자본이익률) = {roe_f:.1f}%"))
        lines.append(("text", "    → 기업이 자기자본 대비 1년간 벌어들이는 순이익 비율"))
        lines.append(("text", "    → ROE가 높을수록 자본 효율이 높은 기업 (예: 15% 이상이면 우수)"))
        lines.append(("text", f"  COE(자기자본비용) = {coe_val:.1f}%"))
        lines.append(("text", "    → 투자자가 이 주식에 투자할 때 '최소한 이 정도는 벌어야 한다'고 기대하는 수익률"))
        lines.append(("text", "    → 일반적으로 COE는 8~10% 범위가 시장에서 통용되는 기본값"))
        lines.append(("blank", ""))

        # 4) CAPM 파라미터 상세 설명
        lines.append(("section", "COE(자기자본비용) 상세 — CAPM 모델"))
        lines.append(("formula", f"COE = Rf + β × MRP = {rf:.2f}% + {beta:.2f} × {mrp:.1f}% = {capm_coe:.1f}%"))
        lines.append(("blank", ""))
        lines.append(("text", f"  Rf(무위험수익률) = {rf:.2f}%"))
        if is_intl:
            lines.append(("text", "    출처: Yahoo Finance '^TNX' (미국 10년 국채 수익률, 실시간 조회)"))
            lines.append(("text", "    의미: 미국 정부가 보증하는 국채 수익률로, 사실상 무위험 자산의 기대 수익률"))
        else:
            lines.append(("text", "    출처: 네이버 금융 (한국 국고채 3년물 수익률, 실시간 조회)"))
            lines.append(("text", "    의미: 한국 정부가 보증하는 국채 수익률로, 국내 투자의 무위험 기준 수익률"))
        lines.append(("text", "    국채는 정부가 원리금을 보증하므로 모든 투자의 수익률 기준점(baseline)입니다."))
        lines.append(("blank", ""))
        lines.append(("text", f"  β(베타) = {beta:.2f}"))
        if is_intl:
            lines.append(("text", "    출처: Yahoo Finance 종목 정보 (S&P500 대비 변동성)"))
        else:
            lines.append(("text", "    출처: 네이버 금융 / pykrx (KOSPI 대비 1년 일별수익률 기반 계산)"))
        lines.append(("text", "    의미: 시장 전체 대비 개별 종목의 변동성 배수"))
        lines.append(("text", "    β > 1: 시장보다 변동성이 큰 공격적 종목 (예: 성장주, 기술주)"))
        lines.append(("text", "    β < 1: 시장보다 변동성이 작은 방어적 종목 (예: 유틸리티, 필수소비재)"))
        lines.append(("text", "    β = 1: 시장과 동일한 변동성"))
        lines.append(("blank", ""))
        lines.append(("text", f"  MRP(시장위험프리미엄) = {mrp:.1f}%"))
        lines.append(("text", "    출처: 역사적 평균값 (글로벌 주식시장의 장기 평균 초과수익률)"))
        lines.append(("text", "    의미: 주식시장 전체가 무위험 자산 대비 추가로 제공하는 수익률"))
        lines.append(("text", "    미국 시장 장기(1926~현재) 평균 약 5~7%, 본 분석에서는 5.5% 적용"))
        lines.append(("blank", ""))

        # CAPM COE vs 기본값 비교
        lines.append(("section", "CAPM COE 해석 (기본값 8~10% 대비)"))
        if is_capm:
            used_coe = capm_coe
            lines.append(("text", f"  본 분석에서 사용된 COE: CAPM 자동계산 = {capm_coe:.1f}%"))
        else:
            used_coe = coe_val
            lines.append(("text", f"  본 분석에서 사용된 COE: 수동입력 = {coe_val:.1f}%"))
            if capm_coe:
                lines.append(("text", f"  (참고: CAPM 자동계산 시 COE = {capm_coe:.1f}%)"))

        if capm_coe > 10:
            lines.append(("text", f"  CAPM COE({capm_coe:.1f}%) > 10%: 시장 평균보다 높은 리스크를 가진 종목"))
            lines.append(("text", "    → 변동성이 크고 위험 프리미엄이 높아, 더 높은 수익률을 요구함"))
            lines.append(("text", "    → 성장주/기술주/소형주에서 흔히 나타남. 적정가가 보수적으로 산출됨"))
        elif capm_coe < 8:
            lines.append(("text", f"  CAPM COE({capm_coe:.1f}%) < 8%: 시장 평균보다 낮은 리스크의 안정적 종목"))
            lines.append(("text", "    → 변동성이 낮고 방어적 성격. 국채 금리가 낮을 때 발생하기도 함"))
            lines.append(("text", "    → 유틸리티/필수소비재/배당주에서 흔히 나타남. 적정가가 높게 산출됨"))
        else:
            lines.append(("text", f"  CAPM COE({capm_coe:.1f}%)는 일반적 범위(8~10%) 내에 있음"))
            lines.append(("text", "    → 시장 평균 수준의 리스크를 가진 종목으로 판단됩니다"))
        lines.append(("blank", ""))

        # 5) ROE vs COE 해석
        lines.append(("section", "ROE vs COE 해석"))
        if roe_f > coe_val:
            lines.append(("text", f"  ROE({roe_f:.1f}%) > COE({coe_val:.1f}%) → 초과이익 발생"))
            lines.append(("text", "    기업이 투자자 기대 수익률 이상으로 수익을 창출하고 있음"))
            lines.append(("text", "    → 적정가가 자본가치(BPS)보다 높게 산출됨 (프리미엄)"))
        elif roe_f < coe_val:
            lines.append(("text", f"  ROE({roe_f:.1f}%) < COE({coe_val:.1f}%) → 초과이익 미달"))
            lines.append(("text", "    기업이 투자자 기대 수익률에 미치지 못하는 수익을 냄"))
            lines.append(("text", "    → 적정가가 자본가치(BPS)보다 낮게 산출됨 (디스카운트)"))
        else:
            lines.append(("text", f"  ROE({roe_f:.1f}%) = COE({coe_val:.1f}%) → 적정 수익 수준"))
            lines.append(("text", "    기업이 정확히 기대 수익률만큼 벌고 있음 → 적정가 ≒ BPS"))
        lines.append(("blank", ""))

        # 6) 투자 판정
        lines.append(("section", "투자 판정"))
        if srim_price and price:
            lines.append(("text", f"  S-RIM 적정가(W={w_fair}) 대비: {diff_pct:+.1f}% {'할인(저평가)' if diff_pct > 0 else '할증(고평가)'}"))
        if buy_price and price:
            buy_diff = (buy_price - price) / buy_price * 100
            lines.append(("text", f"  매수시작가(W={w_buy}) 대비: {buy_diff:+.1f}%"))
        lines.append(("blank", ""))
        # 3-stage verdict
        if buy_price and srim_price and price:
            if price <= buy_price:
                verdict = "적극 매수 구간: 비관적 시나리오 이하. 안전마진 충분"
            elif price <= srim_price:
                verdict = "매수 관심 구간: 적정가 이하. 분할 매수 고려"
            else:
                verdict = "보유/관망 구간: 적정가 초과. 추가 매수보다 보유 유지 권장"
        if verdict:
            lines.append(("text", f"  → {verdict}"))
        lines.append(("blank", ""))

        # 7) 유의사항
        lines.append(("section", "유의사항"))
        lines.append(("text", "  - S-RIM은 과거/예측 ROE 기반이므로 미래 실적 변동 시 결과가 달라질 수 있음"))
        lines.append(("text", "  - 자본이 적거나 음수인 기업(자본잠식)에는 적용이 어려움"))
        lines.append(("text", "  - 고성장주는 ROE 변동이 크므로 S-RIM 단독 판단보다 PER/PBR 등과 병행 권장"))
        lines.append(("text", "  - COE(할인율) 선택에 따라 적정가가 크게 달라지므로 8~10% 범위에서 비교해 보세요"))

        # ── 테이블 구성 ──
        rows = []
        rows.append([
            Paragraph("[S-RIM 모델 해설 (초보자 가이드)]", self.styles["KRiskTitle"]),
        ])
        for tag, line in lines:
            if tag == "blank":
                rows.append([Paragraph("", self.styles["KDesc"])])
            elif tag == "title":
                rows.append([Paragraph(f"<b>{line}</b>", self.styles["KRiskBody"])])
            elif tag == "section":
                rows.append([Paragraph(f"<b>▸ {line}</b>", self.styles["KSmall"])])
            elif tag == "formula":
                rows.append([Paragraph(f"<b>  {line}</b>", self.styles["KSmall"])])
            else:
                rows.append([Paragraph(line, self.styles["KDesc"])])

        t = Table(rows, colWidths=[270 * mm])
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A5276")),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#1A5276")),
        ]
        t.setStyle(TableStyle(style_cmds))
        return t
