"""위험 분석 및 지표 설명 모듈

- 각 재무 지표의 의미 설명
- 위험 수준 판단 (정상/주의/위험/심각)
- 관리종목/상장폐지 조건 체크
"""

# ── 위험 등급 ────────────────────────────────────────────────

LEVEL_OK = "ok"           # 정상
LEVEL_CAUTION = "caution"  # 주의 (노란색)
LEVEL_WARNING = "warning"  # 위험 (주황색)
LEVEL_DANGER = "danger"    # 심각 (빨간색)


# ── 지표별 설명 사전 ─────────────────────────────────────────

INDICATOR_DESC = {
    "매출액": "기업의 총 영업수익. 성장성의 핵심 지표.",
    "영업이익": "매출에서 매출원가·판관비를 뺀 본업 이익.",
    "세전이익": "영업이익에 영업외수익·비용을 가감한 세전 이익.",
    "당기순이익": "법인세 차감 후 최종 이익. 적자 시 자본잠식 위험.",
    "당기순이익(지배)": "지배주주에게 귀속되는 순이익.",
    "영업이익률(%)": "매출 대비 영업이익 비율. 본업 수익성 척도. 업종별 차이 큼.",
    "ROE(%)": "자기자본이익률. 주주 자본 대비 수익성. 8% 이상이 일반적 목표.",
    "ROA(%)": "총자산이익률. 전체 자산 활용 효율. 업종 평균 비교 필요.",
    "레버리지비율": "ROE/ROA. 부채 활용도. 높을수록 부채 의존도가 높음.",
    "자본총계": "순자산(자산-부채). 기업의 재무 안정성 기반.",
    "이자발생부채": "이자를 지급해야 하는 차입금 총액.",
    "순차입금": "이자발생부채 - 현금성자산. 실질적 부채 부담.",
    "영업활동CF": "본업에서 창출한 현금. 순이익보다 현금흐름이 더 중요.",
    "투자활동CF": "설비투자·자산취득에 사용한 현금. 음수가 정상 (투자 중).",
    "재무활동CF": "차입·상환·배당 등 자금조달 활동의 현금흐름.",
    "CAPEX": "유형·무형자산 취득에 투입된 자본적 지출.",
    "FCF": "잉여현금흐름(영업CF-CAPEX). 배당·부채상환 여력.",
    "PFCR": "시가총액/FCF. 낮을수록 현금흐름 대비 저평가.",
    "PER(배)": "주가수익비율. 낮을수록 이익 대비 저평가. 업종 비교 필요.",
    "PBR(배)": "주가순자산비율. 1 미만이면 순자산 대비 저평가.",
    "EPS(원)": "주당순이익. 주당 벌어들인 이익.",
    "BPS(원)": "주당순자산. 청산가치의 근사치.",
    "운전자본": "매출채권+재고-매입채무. 영업에 묶인 자금.",
    "매출대비 비율(%)": "운전자본/매출. 높으면 자금 효율 낮음.",
    "유형자산 비중(%)": "총자산 중 유형자산 비율. 장치산업일수록 높음.",
    "단기채비중(%)": "전체 차입금 중 단기채 비율. 높으면 유동성 리스크.",
}


# ── 지표별 위험 수준 판단 ────────────────────────────────────

def assess_metric(name: str, value, context: dict = None) -> dict:
    """개별 지표의 위험 수준과 코멘트를 반환

    Args:
        name: 지표명
        value: 지표 값
        context: 추가 맥락 (예: 시장구분, 자본금 등)

    Returns:
        {"level": str, "comment": str, "desc": str}
    """
    ctx = context or {}
    desc = INDICATOR_DESC.get(name, "")

    if value is None or value == 0:
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    # ── 수익성 ──
    if name == "영업이익률(%)":
        if value < -10:
            return {"level": LEVEL_DANGER, "comment": "영업적자 심각", "desc": desc}
        if value < 0:
            return {"level": LEVEL_WARNING, "comment": "영업적자", "desc": desc}
        if value < 2:
            return {"level": LEVEL_CAUTION, "comment": "수익성 매우 낮음", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    if name == "ROE(%)":
        if value < -20:
            return {"level": LEVEL_DANGER, "comment": "대규모 적자", "desc": desc}
        if value < 0:
            return {"level": LEVEL_WARNING, "comment": "적자 (자본 감소 중)", "desc": desc}
        if value < 5:
            return {"level": LEVEL_CAUTION, "comment": "수익성 낮음", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    if name == "ROA(%)":
        if value < -10:
            return {"level": LEVEL_DANGER, "comment": "자산 대비 대규모 적자", "desc": desc}
        if value < 0:
            return {"level": LEVEL_WARNING, "comment": "적자", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    if name == "레버리지비율":
        if value > 5:
            return {"level": LEVEL_WARNING, "comment": "과도한 부채 레버리지", "desc": desc}
        if value > 3:
            return {"level": LEVEL_CAUTION, "comment": "부채 의존도 높음", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    # ── 밸류에이션 ──
    if name in ("PER", "PER(배)"):
        if value < 0:
            return {"level": LEVEL_WARNING, "comment": "적자 기업", "desc": desc}
        if value > 50:
            return {"level": LEVEL_CAUTION, "comment": "고평가 가능", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    if name in ("PBR", "PBR(배)"):
        if value < 0:
            return {"level": LEVEL_DANGER, "comment": "자본잠식 (순자산 음수)", "desc": desc}
        if value < 0.3:
            return {"level": LEVEL_CAUTION, "comment": "극단적 저평가 또는 구조적 문제", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    # ── 현금흐름 ──
    if name == "FCF":
        if value < 0:
            return {"level": LEVEL_WARNING, "comment": "잉여현금흐름 적자", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    if name == "영업활동CF":
        if value < 0:
            return {"level": LEVEL_DANGER, "comment": "영업 현금흐름 적자 - 본업 현금 유출", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    if name == "PFCR":
        if value < 0:
            return {"level": LEVEL_WARNING, "comment": "FCF 적자", "desc": desc}
        if value > 30:
            return {"level": LEVEL_CAUTION, "comment": "현금흐름 대비 고평가", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    # ── 차입금 ──
    if name == "단기채비중(%)":
        if value > 80:
            return {"level": LEVEL_WARNING, "comment": "단기 차입 집중 - 유동성 위험", "desc": desc}
        if value > 60:
            return {"level": LEVEL_CAUTION, "comment": "단기 차입 비중 높음", "desc": desc}
        return {"level": LEVEL_OK, "comment": "", "desc": desc}

    return {"level": LEVEL_OK, "comment": "", "desc": desc}


# ── 관리종목 / 상장폐지 조건 체크 ─────────────────────────────

def check_listing_risk(data: dict, derived: dict,
                       ttm_data: dict = None, ttm_derived: dict = None) -> list[dict]:
    """관리종목 지정 및 상장폐지 위험 조건 체크

    Args:
        data: 수집된 원본 데이터
        derived: 연도별 파생지표
        ttm_data: TTM 원본 데이터 (분기 데이터 기반, optional)
        ttm_derived: TTM 파생지표 (optional)

    Returns:
        list of {"level": str, "title": str, "detail": str}
    """
    warnings = []
    years = data.get("years", [])
    fs = data.get("financial_summary", {})
    stock_data = data.get("stock_data", {})
    company_info = data.get("company_info", {})

    corp_cls = company_info.get("corp_cls", "")  # Y:KOSPI, K:KOSDAQ
    is_kosdaq = corp_cls == "K"
    market_name = "코스닥" if is_kosdaq else "코스피"

    if not years:
        return warnings

    latest = years[-1]
    fs_latest = fs.get(latest, {})

    # ── 핵심 값 추출 ──
    from pdf_report_base import _find

    revenue = _find(fs_latest, "매출액")
    op_profit = _find(fs_latest, "영업이익")
    net_income = _find(fs_latest, "당기순이익")
    total_equity = _find(fs_latest, "자본총계")
    capital = _find(fs_latest, "자본금")
    market_cap = stock_data.get("market_cap", 0)

    revenue_eok = revenue / 1e8 if revenue else 0
    mktcap_eok = market_cap / 1e8 if market_cap else 0

    # ── 1. 자본잠식 체크 ──
    if capital and total_equity:
        impairment_ratio = (capital - total_equity) / capital * 100
        if impairment_ratio >= 100:
            warnings.append({
                "level": LEVEL_DANGER,
                "title": f"[상장폐지] 자본 전액잠식 ({impairment_ratio:.0f}%)",
                "detail": (
                    f"자본금({capital/1e8:,.0f}억) > 자본총계({total_equity/1e8:,.0f}억)\n"
                    f"  ▸ 관리종목 기준: 자본잠식률 50% 이상\n"
                    f"  ▸ 상장폐지 기준: 자본잠식률 100% (전액잠식) 또는 2년 연속 50% 이상\n"
                    f"  ※ 현재 전액잠식 — 즉시 상장폐지 사유에 해당"
                ),
            })
        elif impairment_ratio >= 50:
            warnings.append({
                "level": LEVEL_DANGER,
                "title": f"[관리종목] 자본잠식 50% 이상 ({impairment_ratio:.0f}%)",
                "detail": (
                    f"자본잠식률 {impairment_ratio:.0f}% "
                    f"(자본금 {capital/1e8:,.0f}억, 자본총계 {total_equity/1e8:,.0f}억)\n"
                    f"  ▸ 관리종목 기준: 자본잠식률 50% 이상\n"
                    f"  ▸ 상장폐지 기준: 자본잠식률 100% (전액잠식) 또는 2년 연속 50% 이상"
                ),
            })
        elif impairment_ratio >= 30:
            warnings.append({
                "level": LEVEL_WARNING,
                "title": f"자본잠식 접근 ({impairment_ratio:.0f}%)",
                "detail": (
                    f"자본잠식률 {impairment_ratio:.0f}%\n"
                    f"  ▸ 관리종목 기준: 50% 이상 → 현재까지 {50 - impairment_ratio:.0f}%p 여유\n"
                    f"  ▸ 상장폐지 기준: 100% (전액잠식)"
                ),
            })

    # ── 2. 매출액 기준 ──
    if is_kosdaq:
        if revenue_eok < 300:
            warnings.append({
                "level": LEVEL_DANGER,
                "title": f"[관리종목] 매출액 미달 ({revenue_eok:,.0f}억원)",
                "detail": (
                    f"매출액 {revenue_eok:,.0f}억원 < 기준 300억원 (코스닥)\n"
                    f"  ▸ 관리종목 기준: 매출액 300억원 미만 (2027년부터 500억원)\n"
                    f"  ▸ 상장폐지 기준: 2년 연속 매출액 미달"
                ),
            })
        elif revenue_eok < 500:
            warnings.append({
                "level": LEVEL_CAUTION,
                "title": f"매출액 관리종목 기준 접근 ({revenue_eok:,.0f}억원)",
                "detail": (
                    f"매출액 {revenue_eok:,.0f}억원 (기준 300억원의 {revenue_eok/300*100:.0f}%)\n"
                    f"  ▸ 관리종목 기준: 300억원 미만 (2027년부터 500억원으로 상향 예정)\n"
                    f"  ▸ 현재 기준까지 여유: {revenue_eok - 300:,.0f}억원"
                ),
            })
    else:
        if revenue_eok < 500:
            warnings.append({
                "level": LEVEL_DANGER,
                "title": f"[관리종목] 매출액 미달 ({revenue_eok:,.0f}억원)",
                "detail": (
                    f"매출액 {revenue_eok:,.0f}억원 < 기준 500억원 (코스피)\n"
                    f"  ▸ 관리종목 기준: 매출액 500억원 미만"
                ),
            })

    # ── 3. 영업손실 연속 체크 ──
    consec_loss = 0
    for y in years:
        op = _find(fs.get(y, {}), "영업이익")
        if op < 0:
            consec_loss += 1
        else:
            consec_loss = 0

    if not is_kosdaq and consec_loss >= 4:
        warnings.append({
            "level": LEVEL_DANGER,
            "title": f"[관리종목] 영업손실 {consec_loss}년 연속",
            "detail": (
                f"  ▸ KOSPI 관리종목 기준: 4년 연속 영업손실\n"
                f"  ▸ 현재 {consec_loss}년 연속 → 즉시 관리종목 편입 위험"
            ),
        })
    elif not is_kosdaq and consec_loss >= 3:
        warnings.append({
            "level": LEVEL_WARNING,
            "title": f"영업손실 {consec_loss}년 연속",
            "detail": (
                f"  ▸ KOSPI 관리종목 기준: 4년 연속 영업손실\n"
                f"  ▸ 현재 {consec_loss}년 연속 → {4 - consec_loss}년 후 관리종목 편입 위험"
            ),
        })
    elif consec_loss >= 2:
        warnings.append({
            "level": LEVEL_CAUTION,
            "title": f"영업손실 {consec_loss}년 연속",
            "detail": (
                f"  ▸ KOSPI 관리종목 기준: 4년 연속 영업손실 → {4 - consec_loss}년 여유\n"
                f"  ▸ 지속적 영업적자는 관리종목 지정 위험 요인"
            ),
        })

    # ── 4. 세전손실 > 자본의 50% (코스닥) ──
    if is_kosdaq and total_equity:
        pretax = _find(fs_latest, "법인세차감전", "법인세비용차감전")
        if pretax < 0 and abs(pretax) > total_equity * 0.5:
            ratio = abs(pretax) / total_equity * 100
            warnings.append({
                "level": LEVEL_DANGER,
                "title": "[관리종목] 세전손실 > 자기자본의 50%",
                "detail": (
                    f"세전손실 {abs(pretax)/1e8:,.0f}억 / 자기자본 {total_equity/1e8:,.0f}억 "
                    f"= {ratio:.0f}%\n"
                    f"  ▸ 코스닥 관리종목 기준: 세전손실 > 자기자본의 50%\n"
                    f"  ▸ 상장폐지 기준: 2년 연속 해당"
                ),
            })

    # ── 5. 자기자본 < 100억 (코스닥) ──
    if is_kosdaq and total_equity and total_equity < 10_000_000_000:
        warnings.append({
            "level": LEVEL_WARNING,
            "title": f"[관리종목] 자기자본 100억 미만 ({total_equity/1e8:,.0f}억)",
            "detail": (
                f"자기자본 {total_equity/1e8:,.0f}억원\n"
                f"  ▸ 코스닥 관리종목 기준: 자기자본 100억원 미만\n"
                f"  ▸ 기준까지 부족액: {(10_000_000_000 - total_equity)/1e8:,.0f}억원"
            ),
        })

    # ── 6. 시가총액 미달 ──
    if is_kosdaq and mktcap_eok < 40:
        warnings.append({
            "level": LEVEL_DANGER,
            "title": f"[상장폐지] 시가총액 미달 ({mktcap_eok:,.0f}억원)",
            "detail": (
                f"  ▸ 코스닥 상장폐지 기준: 시가총액 40억원 미만 (30일 연속)\n"
                f"  ▸ 현재 시가총액: {mktcap_eok:,.0f}억원"
            ),
        })
    elif not is_kosdaq and mktcap_eok < 50:
        warnings.append({
            "level": LEVEL_DANGER,
            "title": f"[상장폐지] 시가총액 미달 ({mktcap_eok:,.0f}억원)",
            "detail": (
                f"  ▸ 코스피 상장폐지 기준: 시가총액 50억원 미만 (30일 연속)\n"
                f"  ▸ 현재 시가총액: {mktcap_eok:,.0f}억원"
            ),
        })

    # ── 7. 당기순이익 연속 적자 ──
    consec_ni_loss = 0
    for y in years:
        ni = _find(fs.get(y, {}), "당기순이익")
        if ni < 0:
            consec_ni_loss += 1
        else:
            consec_ni_loss = 0

    if consec_ni_loss >= 3:
        warnings.append({
            "level": LEVEL_WARNING,
            "title": f"당기순이익 {consec_ni_loss}년 연속 적자",
            "detail": (
                f"지속적 순손실 → 자본잠식 진행의 원인\n"
                f"  ▸ 자본잠식 50% 이상 시 관리종목 지정\n"
                f"  ▸ 전액잠식 시 상장폐지"
            ),
        })

    # ── 8. 부채비율 체크 (일반 위험 지표) ──
    total_liabilities = _find(fs_latest, "부채총계")
    if total_equity and total_equity > 0:
        debt_ratio = total_liabilities / total_equity * 100
        if debt_ratio > 500:
            warnings.append({
                "level": LEVEL_WARNING,
                "title": f"부채비율 과다 ({debt_ratio:.0f}%)",
                "detail": f"부채비율 {debt_ratio:.0f}%. 재무 안정성 심각한 우려.",
            })
        elif debt_ratio > 300:
            warnings.append({
                "level": LEVEL_CAUTION,
                "title": f"부채비율 높음 ({debt_ratio:.0f}%)",
                "detail": f"부채비율 {debt_ratio:.0f}%. 재무 건전성 주의 필요.",
            })

    # ── 9. 영업CF 적자 연속 ──
    consec_cf_loss = 0
    for y in years:
        d = derived.get(y, {})
        cf = d.get("영업활동CF", 0)
        if cf < 0:
            consec_cf_loss += 1
        else:
            consec_cf_loss = 0

    if consec_cf_loss >= 2:
        warnings.append({
            "level": LEVEL_WARNING,
            "title": f"영업현금흐름 {consec_cf_loss}년 연속 적자",
            "detail": "본업에서 현금 유출이 지속. 유동성 위기 가능성.",
        })

    # ══════════════════════════════════════════════════════════════
    # TTM 기반 추세 경고 (분기 데이터 사용 가능 시)
    # ══════════════════════════════════════════════════════════════
    if ttm_data:
        ttm_bs = ttm_data.get("balance_sheet", {})
        ttm_is = ttm_data.get("financial_summary", {})

        ttm_equity = _find(ttm_bs, "자본총계")
        ttm_capital = _find(ttm_bs, "자본금")
        ttm_revenue = _find(ttm_is, "매출액", "영업수익")
        ttm_op_profit = _find(ttm_is, "영업이익")

        # ── TTM 자본잠식 검사 (최신 분기 BS 기준) ──
        if ttm_capital and ttm_equity and ttm_capital > ttm_equity:
            ttm_imp_ratio = (ttm_capital - ttm_equity) / ttm_capital * 100
            # 연간 기준에서 이미 잡힌 경고와 중복 방지: 연간은 정상이나 TTM에서 악화된 경우만
            annual_imp = 0
            if capital and total_equity:
                annual_imp = (capital - total_equity) / capital * 100
            if annual_imp < 50 and ttm_imp_ratio >= 30:
                warnings.append({
                    "level": LEVEL_WARNING,
                    "title": f"[TTM] 자본잠식 {ttm_imp_ratio:.1f}% (최신 분기 기준)",
                    "detail": (
                        f"연간 보고서 자본잠식률 {annual_imp:.1f}% → TTM 기준 {ttm_imp_ratio:.1f}%\n"
                        f"  ▸ 관리종목 기준: 자본잠식률 50% 이상\n"
                        f"  ▸ 최신 분기 기준 자본잠식이 악화되고 있습니다"
                    ),
                })

        # ── TTM 매출액 미달 접근 경고 ──
        if ttm_revenue and ttm_revenue > 0:
            ttm_rev_eok = ttm_revenue / 1e8
            if is_kosdaq:
                threshold_eok = 300
            else:
                threshold_eok = 500
            # 기준의 120% 이내이면서 연간에서는 잡히지 않은 경우
            if ttm_rev_eok < threshold_eok * 1.2 and revenue_eok >= threshold_eok:
                pct = ttm_rev_eok / threshold_eok * 100
                warnings.append({
                    "level": LEVEL_CAUTION,
                    "title": f"[TTM] 매출액 기준 접근 중 ({ttm_rev_eok:,.0f}억원)",
                    "detail": (
                        f"TTM 매출 {ttm_rev_eok:,.0f}억원 → {market_name} 관리종목 기준 "
                        f"{threshold_eok}억원의 {pct:.0f}%\n"
                        f"  ▸ 연간 매출 {revenue_eok:,.0f}억원은 기준 충족이나 TTM 기준 접근 중"
                    ),
                })
            elif ttm_rev_eok < threshold_eok and revenue_eok >= threshold_eok:
                warnings.append({
                    "level": LEVEL_WARNING,
                    "title": f"[TTM] 매출액 기준 미달 ({ttm_rev_eok:,.0f}억원)",
                    "detail": (
                        f"TTM 매출 {ttm_rev_eok:,.0f}억원 < {market_name} 기준 {threshold_eok}억원\n"
                        f"  ▸ 연간 보고서 기준은 충족({revenue_eok:,.0f}억원)이나\n"
                        f"  ▸ 최근 4분기 합산 기준 미달 — 조기 경고"
                    ),
                })

        # ── TTM 영업이익 적자 전환 ──
        if ttm_derived:
            ttm_opm = ttm_derived.get("영업이익률(%)", 0)
            latest_annual = derived.get(years[-1], {})
            latest_annual_opm = latest_annual.get("영업이익률(%)", 0)

            if latest_annual_opm and latest_annual_opm > 0 and ttm_opm and ttm_opm < 0:
                warnings.append({
                    "level": LEVEL_WARNING,
                    "title": "[TTM] 영업이익 적자 전환",
                    "detail": (
                        f"연간 영업이익률 {latest_annual_opm:.1f}% → TTM {ttm_opm:.1f}%로 적자 전환\n"
                        f"  ▸ KOSPI: 4년 연속 영업손실 시 관리종목\n"
                        f"  ▸ 최근 분기에서 영업적자가 시작되었습니다"
                    ),
                })

            # ── TTM ROE 음수 전환 ──
            ttm_roe = ttm_derived.get("ROE(%)", 0)
            latest_annual_roe = latest_annual.get("ROE(%)", 0)

            if latest_annual_roe and latest_annual_roe > 0 and ttm_roe and ttm_roe < 0:
                warnings.append({
                    "level": LEVEL_CAUTION,
                    "title": "[TTM] ROE 음수 전환",
                    "detail": (
                        f"연간 ROE {latest_annual_roe:.1f}% → TTM ROE {ttm_roe:.1f}%\n"
                        f"  ▸ 순이익 적자 전환 → 자본잠식 진행 가능성\n"
                        f"  ▸ 지속 시 관리종목/상장폐지 위험 증가"
                    ),
                })

    return warnings


def check_us_listing_risk(data: dict, derived: dict) -> list[dict]:
    """미국 거래소(NYSE/NASDAQ) 상장폐지 및 컴플라이언스 위험 체크

    Returns:
        list of {"level": str, "title": str, "detail": str}
    """
    warnings = []
    years = data.get("years", [])
    fs = data.get("financial_summary", {})
    stock_data = data.get("stock_data", {})
    company_info = data.get("company_info", {})

    exchange = company_info.get("exchange", "")
    currency = company_info.get("currency", "USD")

    if not years:
        return warnings

    latest = years[-1]
    fs_latest = fs.get(latest, {})

    from pdf_report_base import _find

    price = stock_data.get("price", 0)
    market_cap = stock_data.get("market_cap", 0)
    total_equity = _find(fs_latest, "자본총계")
    net_income = _find(fs_latest, "당기순이익")
    total_liabilities = _find(fs_latest, "부채총계")
    revenue = _find(fs_latest, "매출액")

    # ── 1. 주가 $1 미만 (Penny Stock / 상장폐지 경고) ──
    if price and price < 1.0:
        warnings.append({
            "level": LEVEL_DANGER,
            "title": "[Delisting] Stock price below $1.00",
            "detail": f"Current price: ${price:.2f}. "
                      "NYSE/NASDAQ require minimum $1.00 bid price. "
                      "30 consecutive trading days below $1 triggers delisting proceedings.",
        })
    elif price and price < 3.0:
        warnings.append({
            "level": LEVEL_CAUTION,
            "title": f"Low stock price (${price:.2f})",
            "detail": "Stock price approaching $1.00 threshold. "
                      "$1 미만 30거래일 연속 시 상장폐지 절차가 시작됩니다.",
        })

    # ── 2. 시가총액 미달 ──
    if market_cap:
        mkt_cap_m = market_cap / 1_000_000  # millions
        if mkt_cap_m < 15:
            warnings.append({
                "level": LEVEL_DANGER,
                "title": f"[Delisting] Market cap below $15M (${mkt_cap_m:,.0f}M)",
                "detail": "NASDAQ minimum market cap requirement: $15M. "
                          "Below this threshold triggers delisting.",
            })
        elif mkt_cap_m < 50:
            warnings.append({
                "level": LEVEL_WARNING,
                "title": f"Low market cap (${mkt_cap_m:,.0f}M)",
                "detail": "NYSE minimum market cap: $50M (30-day average). "
                          "시가총액이 지속적으로 하락할 경우 상장폐지 위험.",
            })

    # ── 3. Negative Equity (자본잠식) ──
    if total_equity and total_equity < 0:
        warnings.append({
            "level": LEVEL_DANGER,
            "title": f"[Delisting Risk] Negative Equity (${total_equity / 1e6:,.0f}M)",
            "detail": "자본총계가 음수입니다 (부채 > 자산). "
                      "NYSE/NASDAQ 상장 유지 조건 위반 가능. 즉시 상장폐지 위험.",
        })
    elif total_equity and total_equity > 0:
        # 자본이 급감 중인지 체크 (이전 연도 대비)
        if len(years) >= 2:
            prev_equity = _find(fs.get(years[-2], {}), "자본총계")
            if prev_equity and prev_equity > 0:
                equity_change = (total_equity - prev_equity) / abs(prev_equity) * 100
                if equity_change < -50:
                    warnings.append({
                        "level": LEVEL_WARNING,
                        "title": f"Equity declining rapidly ({equity_change:+.0f}%)",
                        "detail": f"전년 대비 자본총계 {equity_change:.0f}% 감소. "
                                  "자본잠식(Negative Equity) 진행 시 즉시 상장폐지 사유.",
                    })

    # ── 4. 연속 적자 체크 (NYSE: 적자 + 자본 미달 → 상장폐지) ──
    consec_ni_loss = 0
    for y in years:
        ni = _find(fs.get(y, {}), "당기순이익")
        if ni < 0:
            consec_ni_loss += 1
        else:
            consec_ni_loss = 0

    equity_usd = total_equity / 1_000_000 if total_equity else 0  # millions

    if consec_ni_loss >= 3 and equity_usd < 6:
        warnings.append({
            "level": LEVEL_DANGER,
            "title": f"[Delisting Risk] {consec_ni_loss}Y consecutive losses + equity < $6M",
            "detail": f"순이익 {consec_ni_loss}년 연속 적자 + 자기자본 ${equity_usd:,.0f}M. "
                      "NYSE는 최근 3년 중 2년 적자 + 자기자본 $6M 미만 시 상장폐지 절차 개시.",
        })
    elif consec_ni_loss >= 3:
        warnings.append({
            "level": LEVEL_WARNING,
            "title": f"Net loss {consec_ni_loss} consecutive years",
            "detail": f"순이익 {consec_ni_loss}년 연속 적자. "
                      "추가 적자 지속 + 자기자본 감소 시 상장폐지 위험 증가.",
        })
    elif consec_ni_loss >= 2:
        warnings.append({
            "level": LEVEL_CAUTION,
            "title": f"Net loss {consec_ni_loss} consecutive years",
            "detail": f"순이익 {consec_ni_loss}년 연속 적자. "
                      "1년 더 적자 발생 시 NYSE/NASDAQ 컴플라이언스 위반 위험.",
        })

    # ── 5. 부채비율 과다 ──
    if total_equity and total_equity > 0 and total_liabilities:
        debt_ratio = total_liabilities / total_equity * 100
        if debt_ratio > 500:
            warnings.append({
                "level": LEVEL_WARNING,
                "title": f"Excessive debt ratio ({debt_ratio:.0f}%)",
                "detail": f"부채비율 {debt_ratio:.0f}%. 재무 안정성 심각한 우려.",
            })
        elif debt_ratio > 300:
            warnings.append({
                "level": LEVEL_CAUTION,
                "title": f"High debt ratio ({debt_ratio:.0f}%)",
                "detail": f"부채비율 {debt_ratio:.0f}%. 재무 건전성 주의 필요.",
            })

    # ── 6. 영업CF 연속 적자 ──
    consec_cf_loss = 0
    for y in years:
        d = derived.get(y, {})
        cf = d.get("영업활동CF", 0)
        if cf < 0:
            consec_cf_loss += 1
        else:
            consec_cf_loss = 0

    if consec_cf_loss >= 3:
        warnings.append({
            "level": LEVEL_DANGER,
            "title": f"Operating CF negative {consec_cf_loss} consecutive years",
            "detail": "본업에서 현금 유출 지속. 유동성 위기 및 상장폐지 위험.",
        })
    elif consec_cf_loss >= 2:
        warnings.append({
            "level": LEVEL_WARNING,
            "title": f"Operating CF negative {consec_cf_loss} consecutive years",
            "detail": "본업에서 현금 유출 지속. 1년 더 지속 시 심각한 유동성 위기.",
        })
    elif consec_cf_loss >= 1:
        warnings.append({
            "level": LEVEL_CAUTION,
            "title": "Operating CF negative in latest year",
            "detail": "본업 현금흐름 적자. 연속 적자 시 유동성 위기 가능.",
        })

    # ── 7. 매출 급감 ──
    if len(years) >= 2:
        rev_prev = _find(fs.get(years[-2], {}), "매출액")
        if revenue and rev_prev and rev_prev > 0:
            rev_change = (revenue - rev_prev) / abs(rev_prev) * 100
            if rev_change < -50:
                warnings.append({
                    "level": LEVEL_DANGER,
                    "title": f"Revenue collapsed ({rev_change:+.0f}%)",
                    "detail": f"전년 대비 매출 {rev_change:.0f}% 감소. 사업 지속성 우려.",
                })
            elif rev_change < -30:
                warnings.append({
                    "level": LEVEL_WARNING,
                    "title": f"Revenue declining sharply ({rev_change:+.0f}%)",
                    "detail": f"전년 대비 매출 {rev_change:.0f}% 감소.",
                })

    return warnings


def get_metric_description(name: str) -> str:
    """지표명에 대한 설명 반환"""
    return INDICATOR_DESC.get(name, "")
