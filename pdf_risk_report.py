"""상장폐지/관리종목 리스크 분석 리포트 생성기

PDFReportBase를 상속하여 리스크 분석에 특화된 PDF 리포트를 생성한다.

- 기업 리스크 개요 (종합 등급 배지)
- 관리종목/상장폐지 경고 패널 (기존 pdf_report.py와 동일)
- 조건별 상세 분석 테이블
- 재무 지표 리스크 매트릭스
- 연간 vs TTM 비교
- 5년간 리스크 지표 추이
"""

from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, PageBreak

from pdf_report_base import (
    PDFReportBase, FONT_NAME,
    HEADER_BG, SECTION_BG, SECTION_FG,
    BORDER_COLOR, LIGHT_ROW_BG, SUB_HEADER_BG,
    TTM_RISK_BG,
    RISK_LEVEL_BG, RISK_LABEL_BG,
    CAUTION_BG, WARNING_BG, DANGER_BG,
    TREND_POSITIVE_BG, TREND_NEGATIVE_BG, TREND_NEUTRAL_BG,
    _fmt_num, _fmt_amount, _find,
)

from trend_analyzer import (
    SITUATION_EMOJI,
    TREND_UP, TREND_DOWN, TREND_FLAT, TREND_VOLATILE, TREND_RECOVERY,
)

from risk_analyzer import (
    assess_metric,
    LEVEL_OK, LEVEL_CAUTION, LEVEL_WARNING, LEVEL_DANGER,
)


# ── 종합 리스크 등급 색상 ──────────────────────────────────────

_GRADE_COLORS = {
    "safe":    colors.HexColor("#27AE60"),  # 안전 (녹색)
    "caution": colors.HexColor("#F1C40F"),  # 주의 (노랑)
    "warning": colors.HexColor("#E67E22"),  # 위험 (주황)
    "danger":  colors.HexColor("#C0392B"),  # 심각 (빨강)
}

_GRADE_LABELS = {
    "safe":    "안전",
    "caution": "주의",
    "warning": "위험",
    "danger":  "심각",
}


class RiskReportGenerator(PDFReportBase):
    """상장폐지/관리종목 리스크 분석 PDF 리포트 생성기"""

    # ================================================================
    #  Entry Points
    # ================================================================

    def generate(self, report_data: dict):
        """독립 실행 PDF 빌드"""
        elements = self.build_elements(report_data)
        self._build_doc(elements)

    def build_elements(self, report_data: dict) -> list:
        """요소 리스트를 반환 (다른 리포트에 합치기 위해 사용 가능)"""
        self._prepare(report_data)
        elements = []
        info = report_data.get("company_info", {})
        corp_name = info.get("corp_name", "")

        # Page 1: Title + Risk overview + Listing risk detail + Metric risk matrix
        elements.append(self._build_title_bar(
            report_data,
            title_text=f"상장폐지/관리종목 리스크 분석 - {corp_name}",
            bg_color=colors.HexColor("#C62828"),
        ))
        elements.append(Spacer(1, 3 * mm))

        elements.append(self._build_risk_overview(report_data))
        elements.append(Spacer(1, 3 * mm))

        risk_panel = self._build_risk_panel(report_data)
        if risk_panel:
            elements.append(risk_panel)
            elements.append(Spacer(1, 3 * mm))

        elements.append(self._build_listing_risk_detail(report_data))
        elements.append(Spacer(1, 3 * mm))

        elements.append(self._build_metric_risk_matrix(report_data))

        # Page 2: TTM deterioration + Historical trend
        elements.append(PageBreak())

        ttm_panel = self._build_ttm_deterioration(report_data)
        if ttm_panel:
            elements.append(ttm_panel)
            elements.append(Spacer(1, 3 * mm))

        elements.append(self._build_risk_historical_trend(report_data))

        # Risk-focused trend analysis (if available)
        trend_result = report_data.get("trend_analysis")
        if trend_result and trend_result.get("details"):
            elements.append(Spacer(1, 3 * mm))
            risk_trend = self._build_risk_trend_analysis(report_data)
            if risk_trend:
                elements.extend(risk_trend)

        return elements

    # ================================================================
    #  (a) Risk Overview  - 기업정보 + 시장 + 종합 리스크 등급 배지
    # ================================================================

    def _build_risk_overview(self, data: dict) -> Table:
        """기업 정보 + 종합 리스크 등급 배지 + 경고 요약 카운트"""
        info = data.get("company_info", {})
        stock = data.get("stock_data", {})
        risk_warnings = data.get("risk_warnings", [])
        derived = data.get("derived", {})
        years = data.get("years", [])

        corp_name = info.get("corp_name", "")
        stock_code = info.get("stock_code", "")
        corp_cls = info.get("corp_cls", "")
        is_intl = info.get("is_international", False)

        # 시장 판별
        if is_intl:
            exchange = info.get("exchange", "")
            market_name = exchange if exchange else "해외"
        else:
            market_map = {"Y": "KOSPI", "K": "KOSDAQ", "N": "KONEX", "E": "기타"}
            market_name = market_map.get(corp_cls, "")

        # ── 종합 리스크 등급 결정 (worst level) ──
        level_priority = {
            LEVEL_OK: 0,
            LEVEL_CAUTION: 1,
            LEVEL_WARNING: 2,
            LEVEL_DANGER: 3,
        }
        worst_level = LEVEL_OK

        for w in risk_warnings:
            lvl = w.get("level", LEVEL_OK)
            if level_priority.get(lvl, 0) > level_priority.get(worst_level, 0):
                worst_level = lvl

        # 재무지표 위험 카운트
        metric_risk_count = 0
        if years:
            latest = years[-1]
            d = derived.get(latest, {})
            check_metrics = [
                "영업이익률(%)", "ROE(%)", "ROA(%)", "레버리지비율",
                "PER(배)", "PBR(배)", "FCF", "영업활동CF", "PFCR",
                "단기채비중(%)",
            ]
            for name in check_metrics:
                val = d.get(name, 0)
                if name in ("PER(배)", "PBR(배)"):
                    val = d.get(name.replace("(배)", ""), 0)
                if val:
                    risk = assess_metric(name, val)
                    if risk["level"] != LEVEL_OK:
                        metric_risk_count += 1
                        if level_priority.get(risk["level"], 0) > level_priority.get(worst_level, 0):
                            worst_level = risk["level"]

        # 등급이 ok이면 "safe"
        if worst_level == LEVEL_OK:
            grade_key = "safe"
        else:
            grade_key = worst_level  # caution / warning / danger

        grade_color = _GRADE_COLORS.get(grade_key, colors.gray)
        grade_label = _GRADE_LABELS.get(grade_key, "")

        listing_count = len(risk_warnings)

        # ── 좌측: 기업 정보 ──
        left_rows = [
            ["기업명", corp_name],
            ["종목코드", stock_code],
            ["시장", market_name],
        ]
        left_t = self._mini_table(left_rows, [22 * mm, 45 * mm])

        # ── 중앙: 종합 리스크 등급 배지 ──
        badge_html = (
            f'<font size="10" color="white"'
            f' backColor="{grade_color.hexval()}">'
            f'&nbsp;&nbsp; {grade_label} &nbsp;&nbsp;</font>'
        )
        badge_para = Paragraph(badge_html, self.styles["KNormal"])

        badge_desc = Paragraph(
            "종합 리스크 등급 (최악 지표 기준)",
            self.styles["KDesc"],
        )

        badge_rows = [
            [badge_para],
            [badge_desc],
        ]
        badge_t = Table(badge_rows, colWidths=[60 * mm])
        badge_t.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))

        # ── 우측: 경고 요약 카운트 ──
        summary_text = (
            f"관리종목/상장폐지 경고: {listing_count}건\n"
            f"재무지표 위험: {metric_risk_count}건"
        )
        summary_rows = [
            ["경고 요약", ""],
            ["관리종목/상장폐지 경고", f"{listing_count}건"],
            ["재무지표 위험", f"{metric_risk_count}건"],
        ]
        right_t = self._mini_table(summary_rows, [35 * mm, 20 * mm])

        # ── 합치기 ──
        outer = Table(
            [[left_t, badge_t, right_t]],
            colWidths=[75 * mm, 70 * mm, 60 * mm],
        )
        outer.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOX", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
        ]))
        return outer

    # ================================================================
    #  (b) Risk Panel  — 기존 pdf_report.py 1048-1192 그대로 복사
    # ================================================================

    def _build_risk_panel(self, data: dict) -> Table | None:
        """위험 지표 요약 + 관리종목/상장폐지 경고 패널"""
        risk_warnings = data.get("risk_warnings", [])
        derived = data.get("derived", {})
        years = data.get("years", [])

        # 최신 연도의 위험 지표 수집
        metric_risks = []
        if years:
            latest = years[-1]
            d = derived.get(latest, {})
            check_metrics = [
                "영업이익률(%)", "ROE(%)", "ROA(%)", "레버리지비율",
                "PER(배)", "PBR(배)", "FCF", "영업활동CF", "PFCR",
                "단기채비중(%)",
            ]
            for name in check_metrics:
                val = d.get(name, 0)
                if name in ("PER(배)", "PBR(배)"):
                    val = d.get(name.replace("(배)", ""), 0)
                if val:
                    risk = assess_metric(name, val)
                    if risk["level"] != LEVEL_OK:
                        metric_risks.append({
                            "name": name,
                            "value": val,
                            **risk,
                        })

        if not risk_warnings and not metric_risks:
            return None

        rows = []

        # 헤더
        rows.append([
            Paragraph("[위험 분석 요약]", self.styles["KRiskTitle"]),
            Paragraph("", self.styles["KRiskBody"]),
            Paragraph("", self.styles["KRiskBody"]),
        ])

        # 관리종목/상장폐지 경고
        for w in risk_warnings:
            level = w.get("level", LEVEL_WARNING)
            level_label = {"danger": "심각", "warning": "위험", "caution": "주의"}.get(level, "")
            label_color = RISK_LABEL_BG.get(level, colors.gray)

            title_html = (
                f'<font color="white" backColor="{label_color.hexval()}">'
                f'&nbsp;{level_label}&nbsp;</font>&nbsp;'
                f'{w.get("title", "")}'
            )
            rows.append([
                Paragraph(title_html, self.styles["KRiskBody"]),
                Paragraph(w.get("detail", ""), self.styles["KDesc"]),
                Paragraph("", self.styles["KDesc"]),
            ])

        # 위험 지표 항목
        if metric_risks:
            if risk_warnings:
                rows.append([
                    Paragraph("", self.styles["KDesc"]),
                    Paragraph("", self.styles["KDesc"]),
                    Paragraph("", self.styles["KDesc"]),
                ])

            for mr in metric_risks:
                level = mr.get("level", LEVEL_CAUTION)
                level_label = {"danger": "심각", "warning": "위험", "caution": "주의"}.get(level, "")
                label_color = RISK_LABEL_BG.get(level, colors.gray)

                # 값 포맷팅
                val = mr["value"]
                if "(%)" in mr["name"] or mr["name"] in ("레버리지비율",):
                    val_str = f"{val:.1f}"
                elif mr["name"] in ("FCF", "영업활동CF"):
                    val_str = self._fmt_amt(val)
                else:
                    val_str = f"{val:.2f}"

                title_html = (
                    f'<font color="white" backColor="{label_color.hexval()}">'
                    f'&nbsp;{level_label}&nbsp;</font>&nbsp;'
                    f'{mr["name"]} = {val_str}'
                )
                desc = mr.get("desc", "")
                comment = mr.get("comment", "")
                detail = f"{comment}. {desc}" if desc else comment

                rows.append([
                    Paragraph(title_html, self.styles["KRiskBody"]),
                    Paragraph(detail, self.styles["KDesc"]),
                    Paragraph("", self.styles["KDesc"]),
                ])

        # 범례
        rows.append([
            Paragraph("", self.styles["KDesc"]),
            Paragraph("", self.styles["KDesc"]),
            Paragraph("", self.styles["KDesc"]),
        ])
        legend_html = (
            '<font backColor="#FFF9C4">&nbsp;노랑&nbsp;</font> 주의 &nbsp;&nbsp;'
            '<font backColor="#FFE0B2">&nbsp;주황&nbsp;</font> 위험 &nbsp;&nbsp;'
            '<font backColor="#FFCDD2">&nbsp;빨강&nbsp;</font> 심각'
        )
        rows.append([
            Paragraph(legend_html, self.styles["KDesc"]),
            Paragraph("", self.styles["KDesc"]),
            Paragraph("", self.styles["KDesc"]),
        ])

        col_widths = [90 * mm, 150 * mm, 30 * mm]
        t = Table(rows, colWidths=col_widths)

        style_cmds = [
            # 헤더
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495E")),
            ("SPAN", (0, 0), (-1, 0)),
            # 전체
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#34495E")),
        ]

        # 관리종목/상장폐지 행 배경
        row_offset = 1
        for i, w in enumerate(risk_warnings):
            level = w.get("level", LEVEL_WARNING)
            title = w.get("title", "")
            if "[TTM]" in title:
                # TTM 위험 항목은 연한 파란색 배경으로 구분
                bg = TTM_RISK_BG
            else:
                bg = RISK_LEVEL_BG.get(level, LIGHT_ROW_BG)
            style_cmds.append(("BACKGROUND", (0, row_offset + i), (-1, row_offset + i), bg))

        t.setStyle(TableStyle(style_cmds))
        return t

    # ================================================================
    #  (c) Listing Risk Detail  — 관리종목/상장폐지 조건별 상세 분석
    # ================================================================

    def _build_listing_risk_detail(self, data: dict) -> Table:
        """관리종목/상장폐지 조건별 상세 분석 테이블"""
        info = data.get("company_info", {})
        fs = data.get("financial_summary", {})
        stock = data.get("stock_data", {})
        derived = data.get("derived", {})
        years = data.get("years", [])

        corp_cls = info.get("corp_cls", "")
        is_kosdaq = corp_cls == "K"
        is_intl = info.get("is_international", False)
        market_name = "코스닥" if is_kosdaq else "코스피"

        latest = years[-1] if years else None
        fs_latest = fs.get(latest, {}) if latest else {}

        # 핵심 값 추출
        revenue = _find(fs_latest, "매출액")
        op_profit = _find(fs_latest, "영업이익")
        net_income = _find(fs_latest, "당기순이익")
        total_equity = _find(fs_latest, "자본총계")
        capital = _find(fs_latest, "자본금")
        total_liabilities = _find(fs_latest, "부채총계")
        market_cap = stock.get("market_cap", 0)

        revenue_eok = revenue / 1e8 if revenue else 0
        mktcap_eok = market_cap / 1e8 if market_cap else 0

        # 테이블 헤더
        rows = []
        header = [
            Paragraph("조건", self.styles["KHeader"]),
            Paragraph("기준", self.styles["KHeader"]),
            Paragraph("현재 값", self.styles["KHeader"]),
            Paragraph("판정", self.styles["KHeader"]),
            Paragraph("여유/초과", self.styles["KHeader"]),
        ]
        rows.append(header)

        # 행 단위 배경색 추적
        row_bg_map = {}  # row_index -> bg_color

        def _status_para(is_hit: bool) -> Paragraph:
            if is_hit:
                return Paragraph(
                    '<font color="white" backColor="#C62828">&nbsp;해당&nbsp;</font>',
                    self.styles["KSmall"],
                )
            return Paragraph(
                '<font color="white" backColor="#27AE60">&nbsp;비해당&nbsp;</font>',
                self.styles["KSmall"],
            )

        def _add_row(condition, threshold, current_val, is_hit, margin_str):
            idx = len(rows)
            rows.append([
                Paragraph(str(condition), self.styles["KSmall"]),
                Paragraph(str(threshold), self.styles["KSmall"]),
                Paragraph(str(current_val), self.styles["KSmall"]),
                _status_para(is_hit),
                Paragraph(str(margin_str), self.styles["KSmall"]),
            ])
            if is_hit:
                row_bg_map[idx] = DANGER_BG

        # ── 해외 주식 ──
        if is_intl:
            exchange = info.get("exchange", "")
            price = stock.get("price", 0)

            # 1. 주가 $1 미만
            hit = price > 0 and price < 1.0
            margin = f"${1.0 - price:.2f}" if price and price < 1.0 else f"${price - 1.0:.2f} 여유"
            _add_row(
                "주가 $1 미달",
                "$1.00 미만 30일 연속",
                f"${price:.2f}" if price else "-",
                hit,
                margin if price else "-",
            )

            # 2. 시가총액 미달
            mkt_cap_m = market_cap / 1_000_000 if market_cap else 0
            if "NASDAQ" in exchange.upper():
                threshold_m = 15
            else:
                threshold_m = 50
            hit = mkt_cap_m > 0 and mkt_cap_m < threshold_m
            margin = f"${threshold_m - mkt_cap_m:,.0f}M 부족" if hit else f"${mkt_cap_m - threshold_m:,.0f}M 여유"
            _add_row(
                "시가총액 미달",
                f"${threshold_m}M 미만",
                f"${mkt_cap_m:,.0f}M" if mkt_cap_m else "-",
                hit,
                margin if mkt_cap_m else "-",
            )

            # 3. Negative Equity
            hit = total_equity is not None and total_equity < 0
            eq_m = total_equity / 1_000_000 if total_equity else 0
            _add_row(
                "Negative Equity",
                "자본총계 < 0",
                f"${eq_m:,.0f}M",
                hit,
                f"${abs(eq_m):,.0f}M 부족" if hit else f"${eq_m:,.0f}M 여유",
            )

            # 4. 연속 적자
            consec_ni = 0
            for y in years:
                ni = _find(fs.get(y, {}), "당기순이익")
                if ni < 0:
                    consec_ni += 1
                else:
                    consec_ni = 0
            hit = consec_ni >= 3
            _add_row(
                "연속 순이익 적자",
                "3년 연속 + 자본 미달 시",
                f"{consec_ni}년 연속",
                hit,
                f"{3 - consec_ni}년 여유" if not hit else "기준 초과",
            )

            # 5. 부채비율
            if total_equity and total_equity > 0:
                debt_ratio = total_liabilities / total_equity * 100
            else:
                debt_ratio = 0
            hit = debt_ratio > 500
            _add_row(
                "부채비율",
                "> 500% 위험",
                f"{debt_ratio:.0f}%",
                hit,
                f"{debt_ratio - 500:.0f}%p 초과" if hit else f"{500 - debt_ratio:.0f}%p 여유",
            )

            # 6. 영업CF 연속 적자
            consec_cf = 0
            for y in years:
                dd = derived.get(y, {})
                cf = dd.get("영업활동CF", 0)
                if cf < 0:
                    consec_cf += 1
                else:
                    consec_cf = 0
            hit = consec_cf >= 3
            _add_row(
                "영업CF 연속 적자",
                "3년 연속",
                f"{consec_cf}년 연속",
                hit,
                f"{3 - consec_cf}년 여유" if not hit else "기준 초과",
            )

            # 7. 매출 급감
            rev_change_pct = 0
            if len(years) >= 2:
                rev_prev = _find(fs.get(years[-2], {}), "매출액")
                if revenue and rev_prev and rev_prev > 0:
                    rev_change_pct = (revenue - rev_prev) / abs(rev_prev) * 100
            hit = rev_change_pct < -50
            _add_row(
                "매출 급감",
                "전년 대비 -50% 이상",
                f"{rev_change_pct:+.0f}%",
                hit,
                f"{rev_change_pct + 50:.0f}%p" if not hit else f"{rev_change_pct:.0f}% 하락",
            )

        # ── 한국 주식 ──
        else:
            # 1. 자본잠식률
            impairment_ratio = 0
            if capital and capital > 0 and total_equity is not None:
                impairment_ratio = (capital - total_equity) / capital * 100
            hit_mgmt = impairment_ratio >= 50
            hit_delist = impairment_ratio >= 100
            hit = hit_mgmt or hit_delist
            if hit_delist:
                status_str = "상장폐지"
            elif hit_mgmt:
                status_str = "관리종목"
            else:
                status_str = ""
            _add_row(
                "자본잠식률",
                "50% 이상 -> 관리종목\n100% -> 상장폐지",
                f"{impairment_ratio:.1f}%",
                hit,
                f"{50 - impairment_ratio:.1f}%p 여유" if not hit else f"{status_str} 해당",
            )

            # 2. 매출액 기준
            if is_kosdaq:
                rev_threshold = 300
            else:
                rev_threshold = 500
            hit = revenue_eok < rev_threshold and revenue_eok > 0
            _add_row(
                "매출액",
                f"{market_name} {rev_threshold}억원 미만",
                f"{revenue_eok:,.0f}억원",
                hit,
                f"{rev_threshold - revenue_eok:,.0f}억원 부족" if hit else f"{revenue_eok - rev_threshold:,.0f}억원 여유",
            )

            # 3. 영업손실 연속
            consec_op_loss = 0
            for y in years:
                op = _find(fs.get(y, {}), "영업이익")
                if op < 0:
                    consec_op_loss += 1
                else:
                    consec_op_loss = 0
            if is_kosdaq:
                op_threshold = 4  # 코스닥도 4년 기준 적용
            else:
                op_threshold = 4
            hit = consec_op_loss >= op_threshold
            _add_row(
                "영업손실 연속",
                f"KOSPI {op_threshold}년 연속",
                f"{consec_op_loss}년 연속",
                hit,
                f"{op_threshold - consec_op_loss}년 여유" if not hit else "기준 초과",
            )

            # 4. 세전손실 > 자본의 50% (KOSDAQ only)
            if is_kosdaq:
                pretax = _find(fs_latest, "법인세차감전", "법인세비용차감전", "세전이익")
                if total_equity and total_equity > 0 and pretax < 0:
                    pretax_ratio = abs(pretax) / total_equity * 100
                else:
                    pretax_ratio = 0
                hit = pretax < 0 and total_equity and abs(pretax) > total_equity * 0.5
                _add_row(
                    "세전손실 > 자본 50%",
                    "코스닥: 세전손실 > 자기자본의 50%",
                    f"{pretax_ratio:.0f}%" if pretax < 0 else "해당없음",
                    hit,
                    f"{pretax_ratio - 50:.0f}%p 초과" if hit else f"{50 - pretax_ratio:.0f}%p 여유" if pretax < 0 else "-",
                )

            # 5. 자기자본 < 100억 (KOSDAQ only)
            if is_kosdaq:
                equity_eok = total_equity / 1e8 if total_equity else 0
                hit = total_equity is not None and total_equity > 0 and total_equity < 10_000_000_000
                _add_row(
                    "자기자본 100억 미만",
                    "코스닥: 자기자본 100억원 미만",
                    f"{equity_eok:,.0f}억원",
                    hit,
                    f"{100 - equity_eok:,.0f}억원 부족" if hit else f"{equity_eok - 100:,.0f}억원 여유",
                )

            # 6. 시가총액 미달
            if is_kosdaq:
                mkt_threshold = 40
            else:
                mkt_threshold = 50
            hit = mktcap_eok > 0 and mktcap_eok < mkt_threshold
            _add_row(
                "시가총액 미달",
                f"{market_name} {mkt_threshold}억원 미만",
                f"{mktcap_eok:,.0f}억원",
                hit,
                f"{mkt_threshold - mktcap_eok:,.0f}억원 부족" if hit else f"{mktcap_eok - mkt_threshold:,.0f}억원 여유",
            )

            # 7. 당기순이익 연속 적자
            consec_ni_loss = 0
            for y in years:
                ni = _find(fs.get(y, {}), "당기순이익")
                if ni < 0:
                    consec_ni_loss += 1
                else:
                    consec_ni_loss = 0
            hit = consec_ni_loss >= 3
            _add_row(
                "당기순이익 연속 적자",
                "3년 이상 연속 적자",
                f"{consec_ni_loss}년 연속",
                hit,
                f"{3 - consec_ni_loss}년 여유" if not hit else "기준 초과",
            )

            # 8. 부채비율
            if total_equity and total_equity > 0:
                debt_ratio = total_liabilities / total_equity * 100 if total_liabilities else 0
            else:
                debt_ratio = 0
            hit_warning = debt_ratio > 300
            hit_danger = debt_ratio > 500
            _add_row(
                "부채비율",
                "> 300% 주의, > 500% 위험",
                f"{debt_ratio:.0f}%",
                hit_warning,
                f"{debt_ratio - 300:.0f}%p 초과" if hit_warning else f"{300 - debt_ratio:.0f}%p 여유",
            )

            # 9. 영업CF 연속 적자
            consec_cf_loss = 0
            for y in years:
                dd = derived.get(y, {})
                cf = dd.get("영업활동CF", 0)
                if cf < 0:
                    consec_cf_loss += 1
                else:
                    consec_cf_loss = 0
            hit = consec_cf_loss >= 2
            _add_row(
                "영업CF 연속 적자",
                "2년 이상 연속",
                f"{consec_cf_loss}년 연속",
                hit,
                f"{2 - consec_cf_loss}년 여유" if not hit else "기준 초과",
            )

        # ── 테이블 생성 ──
        col_widths = [45 * mm, 65 * mm, 45 * mm, 30 * mm, 45 * mm]

        # 섹션 제목 행 삽입 (맨 앞)
        section_header = [
            Paragraph("[관리종목/상장폐지 조건별 상세 분석]", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
        ]
        rows.insert(0, section_header)

        t = Table(rows, colWidths=col_widths)

        style_cmds = [
            # 섹션 타이틀
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495E")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("SPAN", (0, 0), (-1, 0)),
            # 헤더
            ("BACKGROUND", (0, 1), (-1, 1), SECTION_BG),
            ("TEXTCOLOR", (0, 1), (-1, 1), SECTION_FG),
            # 전체
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
            ("ALIGN", (2, 2), (4, -1), "RIGHT"),
            ("ALIGN", (3, 2), (3, -1), "CENTER"),
        ]

        # 해당(위험) 행 배경색
        for row_idx, bg in row_bg_map.items():
            # +1 because we inserted the section header row
            actual_idx = row_idx + 1
            style_cmds.append(("BACKGROUND", (0, actual_idx), (-1, actual_idx), bg))

        # 짝수 행 배경
        for i in range(3, len(rows), 2):
            if (i - 1) not in {ri + 1 for ri in row_bg_map}:
                style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_ROW_BG))

        t.setStyle(TableStyle(style_cmds))
        return t

    # ================================================================
    #  (d) Metric Risk Matrix  — 전 재무 지표 리스크 레벨 그리드
    # ================================================================

    def _build_metric_risk_matrix(self, data: dict) -> Table:
        """전 재무 지표 리스크 레벨 그리드"""
        derived = data.get("derived", {})
        years = data.get("years", [])

        latest = years[-1] if years else None
        d = derived.get(latest, {}) if latest else {}

        check_metrics = [
            "영업이익률(%)", "ROE(%)", "ROA(%)", "레버리지비율",
            "PER(배)", "PBR(배)", "FCF", "영업활동CF", "PFCR",
            "단기채비중(%)",
        ]

        level_label_map = {
            LEVEL_OK: "정상",
            LEVEL_CAUTION: "주의",
            LEVEL_WARNING: "위험",
            LEVEL_DANGER: "심각",
        }

        level_color_map = {
            LEVEL_OK: colors.HexColor("#27AE60"),
            LEVEL_CAUTION: colors.HexColor("#F9A825"),
            LEVEL_WARNING: colors.HexColor("#EF6C00"),
            LEVEL_DANGER: colors.HexColor("#C62828"),
        }

        # 헤더
        rows = []
        header = [
            Paragraph("지표", self.styles["KHeader"]),
            Paragraph("현재 값", self.styles["KHeader"]),
            Paragraph("리스크 등급", self.styles["KHeader"]),
            Paragraph("코멘트", self.styles["KHeader"]),
        ]
        rows.append(header)

        cell_styles = {}  # (row, col) -> bg_color

        for name in check_metrics:
            val = d.get(name, 0)
            if name in ("PER(배)", "PBR(배)"):
                val = d.get(name.replace("(배)", ""), 0)

            risk = assess_metric(name, val)
            level = risk["level"]
            level_label = level_label_map.get(level, "정상")
            comment = risk.get("comment", "")

            # 값 포맷팅
            if val is None or val == 0:
                val_str = "-"
            elif "(%)" in name or name in ("레버리지비율",):
                val_str = f"{val:.1f}"
            elif name in ("FCF", "영업활동CF"):
                val_str = self._fmt_amt(val)
            else:
                val_str = f"{val:.2f}"

            # 등급 라벨 HTML
            lbl_color = level_color_map.get(level, colors.gray)
            grade_html = (
                f'<font color="white" backColor="{lbl_color.hexval()}">'
                f'&nbsp;{level_label}&nbsp;</font>'
            )

            row_idx = len(rows)
            rows.append([
                Paragraph(name, self.styles["KSmall"]),
                Paragraph(val_str, self.styles["KSmall"]),
                Paragraph(grade_html, self.styles["KSmall"]),
                Paragraph(comment, self.styles["KSmall"]),
            ])

            # 배경색 지정
            bg = RISK_LEVEL_BG.get(level)
            if bg:
                cell_styles[row_idx] = bg

        # 섹션 타이틀 삽입
        section_header = [
            Paragraph("[재무 지표 리스크 매트릭스]", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
        ]
        rows.insert(0, section_header)

        col_widths = [50 * mm, 40 * mm, 35 * mm, 105 * mm]
        t = Table(rows, colWidths=col_widths)

        style_cmds = [
            # 섹션 타이틀
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495E")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("SPAN", (0, 0), (-1, 0)),
            # 헤더
            ("BACKGROUND", (0, 1), (-1, 1), SECTION_BG),
            ("TEXTCOLOR", (0, 1), (-1, 1), SECTION_FG),
            # 전체
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
            ("ALIGN", (1, 2), (1, -1), "RIGHT"),
            ("ALIGN", (2, 2), (2, -1), "CENTER"),
        ]

        # 위험 행 배경색
        for row_idx, bg in cell_styles.items():
            actual_idx = row_idx + 1  # +1 for section header insert
            style_cmds.append(("BACKGROUND", (0, actual_idx), (-1, actual_idx), bg))

        # 짝수 행 배경 (위험 없는 행만)
        for i in range(3, len(rows), 2):
            if (i - 1) not in {ri + 1 for ri in cell_styles}:
                style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_ROW_BG))

        t.setStyle(TableStyle(style_cmds))
        return t

    # ================================================================
    #  (e) TTM Deterioration  — 연간 vs TTM 비교 패널
    # ================================================================

    def _build_ttm_deterioration(self, data: dict) -> Table | None:
        """연간 vs TTM 주요 지표 비교 테이블"""
        derived = data.get("derived", {})
        years = data.get("years", [])
        ttm_derived = data.get("ttm_derived", {})
        ttm_data = data.get("ttm_data", {})

        if not ttm_derived and not ttm_data:
            return None

        latest = years[-1] if years else None
        if not latest:
            return None

        annual = derived.get(latest, {})
        fs = data.get("financial_summary", {})
        fs_latest = fs.get(latest, {})

        # TTM financial_summary 데이터 (있으면)
        ttm_fs = ttm_data.get("financial_summary", {}) if ttm_data else {}

        # 비교할 지표 목록
        # (표시명, 연간키_fs, 연간키_derived, ttm키_fs, ttm키_derived, format_type)
        compare_items = [
            ("매출액",    "매출액",      None,            "매출액",    None,            "amount"),
            ("영업이익",  "영업이익",    None,            "영업이익",  None,            "amount"),
            ("영업이익률", None,         "영업이익률(%)", None,        "영업이익률(%)", "pct"),
            ("ROE",       None,         "ROE(%)",        None,        "ROE(%)",        "pct"),
            ("FCF",       None,         "FCF",           None,        "FCF",           "amount"),
        ]

        rows = []
        header = [
            Paragraph("지표", self.styles["KHeader"]),
            Paragraph("최근연간", self.styles["KHeader"]),
            Paragraph("TTM", self.styles["KHeader"]),
            Paragraph("변화", self.styles["KHeader"]),
            Paragraph("판정", self.styles["KHeader"]),
        ]
        rows.append(header)

        row_colors = {}  # row_idx -> bg

        for label, fs_key, derived_key, ttm_fs_key, ttm_derived_key, fmt_type in compare_items:
            # 연간 값
            if fs_key:
                annual_val = _find(fs_latest, fs_key)
            elif derived_key:
                annual_val = annual.get(derived_key, 0)
            else:
                annual_val = 0

            # TTM 값
            ttm_val = 0
            if ttm_fs_key and ttm_fs:
                ttm_val = _find(ttm_fs, ttm_fs_key)
            if not ttm_val and ttm_derived_key and ttm_derived:
                ttm_val = ttm_derived.get(ttm_derived_key, 0)

            # 포맷팅
            if fmt_type == "amount":
                annual_str = self._fmt_amt(annual_val)
                ttm_str = self._fmt_amt(ttm_val)
            else:  # pct
                annual_str = f"{annual_val:.1f}%" if annual_val else "-"
                ttm_str = f"{ttm_val:.1f}%" if ttm_val else "-"

            # 변화 계산
            if annual_val and ttm_val:
                if fmt_type == "pct":
                    change = ttm_val - annual_val
                    change_str = f"{change:+.1f}%p"
                else:
                    if annual_val != 0:
                        change_pct = (ttm_val - annual_val) / abs(annual_val) * 100
                        change_str = f"{change_pct:+.1f}%"
                        change = change_pct
                    else:
                        change_str = "-"
                        change = 0
            else:
                change_str = "-"
                change = 0

            # 판정: 악화/개선/유지
            # 매출액, 영업이익, 영업이익률, ROE, FCF 모두 증가가 긍정
            if change and change != 0:
                if change < -5:
                    verdict = "악화"
                    verdict_color = colors.HexColor("#E74C3C")
                elif change > 5:
                    verdict = "개선"
                    verdict_color = colors.HexColor("#27AE60")
                else:
                    verdict = "유지"
                    verdict_color = colors.HexColor("#7F8C8D")
            else:
                verdict = "-"
                verdict_color = colors.HexColor("#7F8C8D")

            verdict_html = f'<font color="{verdict_color.hexval()}">{verdict}</font>'

            row_idx = len(rows)
            rows.append([
                Paragraph(label, self.styles["KSmall"]),
                Paragraph(annual_str, self.styles["KSmall"]),
                Paragraph(ttm_str, self.styles["KSmall"]),
                Paragraph(change_str, self.styles["KSmall"]),
                Paragraph(verdict_html, self.styles["KSmall"]),
            ])

            # 악화 시 빨간 배경
            if verdict == "악화":
                row_colors[row_idx] = colors.HexColor("#FFEBEE")
            elif verdict == "개선":
                row_colors[row_idx] = colors.HexColor("#E8F5E9")

        # 섹션 타이틀 삽입
        section_header = [
            Paragraph("[연간 vs TTM 비교 분석]", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
            Paragraph("", self.styles["KHeader"]),
        ]
        rows.insert(0, section_header)

        col_widths = [40 * mm, 50 * mm, 50 * mm, 40 * mm, 30 * mm]
        t = Table(rows, colWidths=col_widths)

        style_cmds = [
            # 섹션 타이틀
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495E")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("SPAN", (0, 0), (-1, 0)),
            # 헤더
            ("BACKGROUND", (0, 1), (-1, 1), SECTION_BG),
            ("TEXTCOLOR", (0, 1), (-1, 1), SECTION_FG),
            # 전체
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6.5),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
            ("ALIGN", (1, 2), (4, -1), "RIGHT"),
            ("ALIGN", (4, 2), (4, -1), "CENTER"),
        ]

        # 행 배경색
        for row_idx, bg in row_colors.items():
            actual_idx = row_idx + 1  # +1 for section header
            style_cmds.append(("BACKGROUND", (0, actual_idx), (-1, actual_idx), bg))

        t.setStyle(TableStyle(style_cmds))
        return t

    # ================================================================
    #  (f) Risk Historical Trend  — 5년간 리스크 지표 추이
    # ================================================================

    def _build_risk_historical_trend(self, data: dict) -> Table:
        """5년간 리스크 관련 지표 추이 테이블"""
        fs = data.get("financial_summary", {})
        derived = data.get("derived", {})
        years = data.get("years", [])

        info = data.get("company_info", {})
        corp_cls = info.get("corp_cls", "")
        is_kosdaq = corp_cls == "K"
        is_intl = info.get("is_international", False)

        # 행 정의: (표시명, fs_key, derived_key, format_type, threshold_func)
        # threshold_func: value -> bool (True = 위험)
        if is_intl:
            rev_threshold_val = 0  # 해외주식은 매출 절대기준 없음
        elif is_kosdaq:
            rev_threshold_val = 300 * 1e8  # 300억
        else:
            rev_threshold_val = 500 * 1e8  # 500억

        metric_rows = [
            ("매출액",      "매출액",      None,          "amount",
             lambda v: v > 0 and v < rev_threshold_val and rev_threshold_val > 0),
            ("영업이익",    "영업이익",    None,          "amount",
             lambda v: v < 0),
            ("당기순이익",  "당기순이익",  None,          "amount",
             lambda v: v < 0),
            ("자본총계",    "자본총계",    None,          "amount",
             lambda v: v is not None and v < 0),
            ("부채비율",    None,          None,          "pct",
             lambda v: v > 300),
            ("영업활동CF",  None,          "영업활동CF",  "amount",
             lambda v: v < 0),
        ]

        # 헤더
        header = ["지표"] + [str(y) for y in years]
        rows = [header]

        highlight_cells = {}  # (row_idx, col_idx) -> True

        for label, fs_key, derived_key, fmt_type, threshold_fn in metric_rows:
            row = [label]
            row_idx = len(rows)

            for col_idx, y in enumerate(years, 1):
                fs_y = fs.get(y, {})
                d_y = derived.get(y, {})

                # 값 가져오기
                if label == "부채비율":
                    total_liab = self._find_account(fs, y, "부채총계")
                    total_eq = self._find_account(fs, y, "자본총계")
                    if total_eq and total_eq > 0:
                        val = total_liab / total_eq * 100
                    else:
                        val = 0
                elif fs_key:
                    val = _find(fs_y, fs_key)
                elif derived_key:
                    val = d_y.get(derived_key, 0)
                else:
                    val = 0

                # 포맷팅
                if fmt_type == "amount":
                    cell_str = self._fmt_amt(val)
                elif fmt_type == "pct":
                    cell_str = f"{val:.0f}%" if val else "-"
                else:
                    cell_str = str(val) if val else "-"

                row.append(cell_str)

                # 임계치 위반 하이라이트
                if val and threshold_fn(val):
                    highlight_cells[(row_idx, col_idx)] = True

            rows.append(row)

        # 섹션 테이블 빌드
        num_years = len(years)
        label_width = 32 * mm
        data_width = (270 * mm - label_width) / max(num_years, 1)
        col_widths = [label_width] + [data_width] * num_years

        styled_rows = []

        # 섹션 타이틀
        title_row = [Paragraph("[5년간 리스크 지표 추이]", self.styles["KHeader"])]
        title_row += [Paragraph("", self.styles["KHeader"])] * num_years
        styled_rows.append(title_row)

        # 헤더
        header_row = [Paragraph(str(c), self.styles["KHeader"]) for c in rows[0]]
        styled_rows.append(header_row)

        # 데이터 행
        for row in rows[1:]:
            styled_row = [Paragraph(str(c), self.styles["KSmall"]) for c in row]
            styled_rows.append(styled_row)

        t = Table(styled_rows, colWidths=col_widths)

        style_cmds = [
            # 섹션 타이틀
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495E")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("SPAN", (0, 0), (-1, 0)),
            # 헤더
            ("BACKGROUND", (0, 1), (-1, 1), SECTION_BG),
            ("TEXTCOLOR", (0, 1), (-1, 1), SECTION_FG),
            # 전체
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE", (0, 0), (-1, -1), 6.5),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, BORDER_COLOR),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ]

        # 짝수 행 배경
        for i in range(3, len(styled_rows), 2):
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_ROW_BG))

        # 임계치 위반 셀 하이라이트
        for (row_idx, col_idx), _ in highlight_cells.items():
            # styled_rows offset: +2 (section title + header)
            actual_row = row_idx + 1  # +1 for section title (data rows start at row_idx=1 in rows)
            style_cmds.append(
                ("BACKGROUND", (col_idx, actual_row), (col_idx, actual_row), DANGER_BG)
            )

        t.setStyle(TableStyle(style_cmds))
        return t

    # ================================================================
    #  (g) Risk-Focused Trend Analysis  — 리스크 관점 추이 분석
    # ================================================================

    def _build_risk_trend_analysis(self, data: dict) -> list:
        """리스크 관점에서 재무안정성/현금흐름/수익성 추이를 분석"""
        trend = data.get("trend_analysis", {})
        if not trend or not trend.get("details"):
            return []

        details = trend.get("details", [])
        situation = trend.get("situation", "")
        label = trend.get("situation_label", "분석 불가")
        color_hex = trend.get("situation_color", "#95A5A6")
        confidence = trend.get("confidence", 0)
        summary = trend.get("summary", "")

        # 리스크 관련 영역만 필터
        risk_categories = {"재무안정성", "현금흐름", "수익성"}
        risk_details = [d for d in details if d.get("category") in risk_categories]

        if not risk_details:
            return []

        elements = []

        # ── 섹션 헤더 ──
        header_data = [[
            Paragraph(
                '<font color="white"><b>[리스크 관점 추이 분석]</b></font>',
                self.styles["KHeader"],
            ),
        ]]
        header_t = Table(header_data, colWidths=[270 * mm])
        header_t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#C62828")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ]))
        elements.append(header_t)
        elements.append(Spacer(1, 1 * mm))

        # ── 종합 상황 요약 ──
        situation_color = colors.HexColor(color_hex)
        emoji = SITUATION_EMOJI.get(situation, "?")

        summary_rows = [[
            Paragraph(
                f'<font color="white" size="8"><b> {label} </b></font>',
                self.styles["KTitle"],
            ),
            Paragraph(
                f'<font size="6.5">신뢰도: <b>{confidence * 100:.0f}%</b></font>',
                self.styles["KNormal"],
            ),
        ]]
        summary_t = Table(summary_rows, colWidths=[180 * mm, 90 * mm])
        summary_t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), situation_color),
            ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#ECEFF1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, 0), "CENTER"),
            ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#C62828")),
        ]))
        elements.append(summary_t)
        elements.append(Spacer(1, 1 * mm))

        # ── 리스크 영역별 분석 테이블 ──
        trend_labels = {
            TREND_UP: "\u2191 상승",
            TREND_DOWN: "\u2193 하락",
            TREND_FLAT: "\u2192 횡보",
            TREND_VOLATILE: "\u2195 변동",
            TREND_RECOVERY: "\u2197 회복",
        }

        detail_rows = []
        detail_rows.append([
            Paragraph('<font color="white"><b>영역</b></font>', self.styles["KHeader"]),
            Paragraph('<font color="white"><b>추세</b></font>', self.styles["KHeader"]),
            Paragraph('<font color="white"><b>점수</b></font>', self.styles["KHeader"]),
            Paragraph('<font color="white"><b>리스크 진단</b></font>', self.styles["KHeader"]),
            Paragraph('<font color="white"><b>상세</b></font>', self.styles["KHeader"]),
        ])

        for d in risk_details:
            category = d.get("category", "")
            trend_dir = d.get("trend", TREND_FLAT)
            score = d.get("score", 0)
            title = d.get("title", "")
            comment = d.get("comment", "")
            sub_items = d.get("sub_items", [])

            trend_text = trend_labels.get(trend_dir, "\u2192 횡보")

            # 리스크 관점 점수 표현: 부정일수록 위험
            if score >= 2:
                score_text = '<font color="#2ECC71"><b>안전</b></font>'
            elif score == 1:
                score_text = '<font color="#3498DB"><b>양호</b></font>'
            elif score == 0:
                score_text = '<font color="#7F8C8D"><b>보통</b></font>'
            elif score == -1:
                score_text = '<font color="#E67E22"><b>주의</b></font>'
            else:
                score_text = '<font color="#E74C3C"><b>위험</b></font>'

            full_comment = comment
            if sub_items:
                si_parts = []
                for si in sub_items:
                    si_parts.append(f"{si.get('label', '')} {si.get('value', '')}")
                sub_line = " | ".join(si_parts)
                full_comment += f'<br/><font size="5" color="#555555">  \u25b8 {sub_line}</font>'

            detail_rows.append([
                Paragraph(f'<b>{category}</b>', self.styles["KNormal"]),
                Paragraph(trend_text, self.styles["KSmall"]),
                Paragraph(score_text, self.styles["KNormal"]),
                Paragraph(title, self.styles["KSmall"]),
                Paragraph(full_comment, self.styles["KSmall"]),
            ])

        col_widths = [28 * mm, 18 * mm, 15 * mm, 65 * mm, 144 * mm]
        detail_t = Table(detail_rows, colWidths=col_widths)

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#C62828")),
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

        # 점수 기반 배경색 (부정=위험색)
        for i, d in enumerate(risk_details, start=1):
            score = d.get("score", 0)
            if score <= -2:
                bg = DANGER_BG
            elif score == -1:
                bg = WARNING_BG
            elif score == 0:
                bg = CAUTION_BG
            elif score >= 1:
                bg = TREND_POSITIVE_BG
            else:
                bg = TREND_NEUTRAL_BG
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), bg))

        detail_t.setStyle(TableStyle(style_cmds))
        elements.append(detail_t)
        elements.append(Spacer(1, 2 * mm))

        # ── 리스크 궤적 해석 ──
        trajectory_parts = []
        for d in risk_details:
            cat = d.get("category", "")
            trend_dir = d.get("trend", TREND_FLAT)
            score = d.get("score", 0)
            if score <= -1:
                direction = "악화" if trend_dir in (TREND_DOWN, TREND_VOLATILE) else "부진"
                trajectory_parts.append(
                    f'<font color="#C62828"><b>{cat}</b>: {direction} 추세</font>'
                )
            elif score >= 1:
                direction = "개선" if trend_dir in (TREND_UP, TREND_RECOVERY) else "안정"
                trajectory_parts.append(
                    f'<font color="#27AE60"><b>{cat}</b>: {direction} 추세</font>'
                )
            else:
                trajectory_parts.append(
                    f'<font color="#7F8C8D"><b>{cat}</b>: 보통 수준</font>'
                )

        trajectory_text = " | ".join(trajectory_parts)
        trajectory_rows = [[
            Paragraph(
                f'<font size="6"><b>리스크 궤적:</b></font> '
                f'<font size="5.5">{trajectory_text}</font>',
                self.styles["KNormal"],
            ),
        ]]

        trajectory_t = Table(trajectory_rows, colWidths=[270 * mm])
        trajectory_t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF3E0")),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#C62828")),
        ]))
        elements.append(trajectory_t)

        return elements
