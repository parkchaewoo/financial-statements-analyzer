"""재무 추이 분석 모듈

재무제표 다년간 데이터를 분석하여 기업의 현재 상황과 추세를 진단합니다.

분석 영역:
  1. 매출 성장성 (Revenue Growth Trend)
  2. 수익성 추이 (Profitability Trend)
  3. 재무 안정성 (Financial Stability)
  4. 현금흐름 품질 (Cash Flow Quality)
  5. 종합 진단 (Overall Assessment)
"""

from __future__ import annotations

# ── 기업 상황 분류 ─────────────────────────────────────────────

SITUATION_HIGH_GROWTH = "high_growth"       # 고성장
SITUATION_STABLE_GROWTH = "stable_growth"   # 안정 성장
SITUATION_VALUE = "value"                   # 가치주 (저평가 안정)
SITUATION_TURNAROUND = "turnaround"         # 턴어라운드
SITUATION_MATURE = "mature"                 # 성숙/정체
SITUATION_DECLINING = "declining"           # 하향세
SITUATION_CRISIS = "crisis"                 # 위기

SITUATION_LABELS = {
    SITUATION_HIGH_GROWTH: "고성장",
    SITUATION_STABLE_GROWTH: "안정 성장",
    SITUATION_VALUE: "가치주",
    SITUATION_TURNAROUND: "턴어라운드",
    SITUATION_MATURE: "성숙/정체",
    SITUATION_DECLINING: "하향세",
    SITUATION_CRISIS: "위기",
}

SITUATION_COLORS = {
    SITUATION_HIGH_GROWTH: "#2ECC71",    # 녹색
    SITUATION_STABLE_GROWTH: "#3498DB",  # 파란색
    SITUATION_VALUE: "#9B59B6",          # 보라색
    SITUATION_TURNAROUND: "#F39C12",     # 주황색
    SITUATION_MATURE: "#95A5A6",         # 회색
    SITUATION_DECLINING: "#E67E22",      # 진한 주황
    SITUATION_CRISIS: "#E74C3C",         # 빨간색
}

SITUATION_EMOJI = {
    SITUATION_HIGH_GROWTH: "++",
    SITUATION_STABLE_GROWTH: "+",
    SITUATION_VALUE: "V",
    SITUATION_TURNAROUND: "T",
    SITUATION_MATURE: "=",
    SITUATION_DECLINING: "-",
    SITUATION_CRISIS: "!!",
}


# ── 추세 방향 ──────────────────────────────────────────────────

TREND_UP = "up"
TREND_FLAT = "flat"
TREND_DOWN = "down"
TREND_VOLATILE = "volatile"
TREND_RECOVERY = "recovery"    # 하락 후 반등


# ── 핵심 분석 함수 ─────────────────────────────────────────────

def analyze_trend(data: dict, derived: dict,
                  ttm_data: dict = None, ttm_derived: dict = None,
                  quarterly_derived: dict = None,
                  quarterly_keys: list = None) -> dict:
    """재무 추이 종합 분석

    Args:
        data: collect_data()의 반환값 (financial_summary, balance_sheet_detail, etc.)
        derived: compute_derived_metrics()의 연도별 파생지표
        ttm_data: TTM 원본 데이터 (optional)
        ttm_derived: TTM 파생지표 (optional)

    Returns:
        {
            "situation": str,            # 기업 상황 분류 코드
            "situation_label": str,      # 한글 라벨
            "situation_color": str,      # 색상 코드
            "confidence": float,         # 진단 신뢰도 (0~1)
            "summary": str,              # 종합 진단 요약 (2~3문장)
            "details": [                 # 영역별 상세 분석
                {
                    "category": str,     # 분석 영역
                    "trend": str,        # 추세 방향
                    "score": int,        # 점수 (-2 ~ +2)
                    "title": str,        # 한줄 제목
                    "comment": str,      # 상세 설명
                },
                ...
            ],
            "key_metrics": {             # 핵심 요약 지표
                "revenue_cagr": float,   # 매출 연평균 성장률
                "opm_latest": float,     # 최신 영업이익률
                "roe_latest": float,     # 최신 ROE
                "debt_ratio_latest": float,  # 최신 부채비율
                "fcf_trend": str,        # FCF 추세
            },
        }
    """
    from pdf_report_base import _find

    years = data.get("years", [])
    fs = data.get("financial_summary", {})
    bs = data.get("balance_sheet_detail", {})
    stock_data = data.get("stock_data", {})

    if len(years) < 2:
        return _empty_result("데이터 부족으로 추이 분석 불가 (최소 2개년 필요)")

    # ── 연도별 핵심 지표 추출 ──
    metrics = _extract_yearly_metrics(years, fs, bs, derived, _find)

    # ── 1. 매출 성장성 분석 ──
    revenue_analysis = _analyze_revenue_growth(metrics, years, _find, ttm_derived)

    # ── 2. 수익성 추이 분석 ──
    profitability_analysis = _analyze_profitability(metrics, years, ttm_derived)

    # ── 3. 재무 안정성 분석 ──
    stability_analysis = _analyze_financial_stability(metrics, years, _find, fs, ttm_data)

    # ── 4. 현금흐름 품질 분석 ──
    cashflow_analysis = _analyze_cashflow_quality(metrics, years, ttm_derived)

    # ── 5. 밸류에이션 분석 ──
    valuation_analysis = _analyze_valuation(metrics, years, stock_data, ttm_derived)

    # ── 6. 효율성 분석 ──
    efficiency_analysis = _analyze_efficiency(metrics, years, ttm_derived)

    details = [
        revenue_analysis,
        profitability_analysis,
        stability_analysis,
        cashflow_analysis,
        valuation_analysis,
        efficiency_analysis,
    ]

    # ── TTM 점수 보정 ──
    if ttm_derived:
        _apply_ttm_score_adjustment(details, metrics, years, ttm_derived)

    # ── 분기 모멘텀 분석 ──
    quarterly_momentum = _analyze_quarterly_momentum(
        quarterly_derived, quarterly_keys)

    # ── 종합 진단 ──
    scores = [d["score"] for d in details]
    avg_score = sum(scores) / len(scores) if scores else 0

    situation = _classify_situation(details, metrics, years, avg_score)
    confidence = _calc_confidence(years, metrics)
    summary = _generate_summary(situation, details, metrics, years,
                                ttm_derived, stock_data)

    # 핵심 요약 지표
    key_metrics = _extract_key_metrics(metrics, years, ttm_derived)
    key_metrics["quarterly_momentum"] = quarterly_momentum

    # DuPont 분해
    dupont = _build_dupont_analysis(metrics, years)

    # 연도별 핵심 지표 추이 테이블
    yearly_table = _build_yearly_table(metrics, years)

    # 투자 체크리스트
    checklist = _build_investment_checklist(details, metrics, years,
                                           ttm_derived, key_metrics)

    # 강점/약점 요약
    strengths_weaknesses = _build_strengths_weaknesses(
        details, metrics, years, ttm_derived, key_metrics)

    return {
        "situation": situation,
        "situation_label": SITUATION_LABELS.get(situation, "분석 불가"),
        "situation_color": SITUATION_COLORS.get(situation, "#95A5A6"),
        "confidence": confidence,
        "summary": summary,
        "details": details,
        "key_metrics": key_metrics,
        "dupont": dupont,
        "yearly_table": yearly_table,
        "checklist": checklist,
        "strengths_weaknesses": strengths_weaknesses,
    }


# ── 지표 추출 ──────────────────────────────────────────────────

def _extract_yearly_metrics(years, fs, bs, derived, _find):
    """연도별 핵심 지표를 정리"""
    metrics = {}
    for y in years:
        fs_y = fs.get(y, {})
        bs_y = bs.get(y, {})
        d_y = derived.get(y, {})

        revenue = _find(fs_y, "매출액", "영업수익")
        op_profit = _find(fs_y, "영업이익")
        net_income = _find(fs_y, "당기순이익")
        total_assets = _find(fs_y, "자산총계")
        total_equity = _find(fs_y, "자본총계")
        total_liabilities = _find(fs_y, "부채총계")
        capital = _find(fs_y, "자본금")

        # 순이익률, 자산회전율, 재무레버리지 (DuPont 분해용)
        net_margin = (net_income / revenue * 100) if revenue and revenue > 0 else 0
        asset_turnover = (revenue / total_assets) if total_assets and total_assets > 0 else 0
        equity_multiplier = (total_assets / total_equity) if (
            total_equity and total_equity > 0) else 0

        metrics[y] = {
            "revenue": revenue,
            "op_profit": op_profit,
            "net_income": net_income,
            "total_assets": total_assets,
            "total_equity": total_equity,
            "total_liabilities": total_liabilities,
            "capital": capital,
            # 수익성
            "opm": d_y.get("영업이익률(%)", 0),
            "roe": d_y.get("ROE(%)", 0),
            "roa": d_y.get("ROA(%)", 0),
            "net_margin": net_margin,
            "leverage": d_y.get("레버리지비율", 0),
            "eff_tax_rate": d_y.get("유효법인세율(%)", 0),
            # 밸류에이션
            "per": d_y.get("PER", 0),
            "pbr": d_y.get("PBR", 0),
            "eps": d_y.get("EPS", 0),
            "bps": d_y.get("BPS", 0),
            "pfcr": d_y.get("PFCR", 0),
            # 현금흐름
            "fcf": d_y.get("FCF", 0),
            "op_cf": d_y.get("영업활동CF", 0),
            "inv_cf": d_y.get("투자활동CF", 0),
            "fin_cf": d_y.get("재무활동CF", 0),
            "capex": d_y.get("CAPEX", 0),
            # 재무 안정성
            "debt_total": d_y.get("이자발생부채계산", 0),
            "cash_total": d_y.get("현금성자산합계", 0),
            "net_debt": d_y.get("순차입금", 0),
            "short_debt_ratio": d_y.get("단기채비중(%)", 0),
            "long_debt_ratio": d_y.get("장기채비중(%)", 0),
            # 효율성
            "working_capital": d_y.get("운전자본", 0),
            "wc_ratio": d_y.get("운전자본비율(%)", 0),
            "tangible_ratio": d_y.get("유형자산비중(%)", 0),
            "intangible_ratio": d_y.get("무형자산비중(%)", 0),
            "asset_turnover": asset_turnover,
            "equity_multiplier": equity_multiplier,
            # 성장률
            "rev_growth": d_y.get("매출성장률(%)", None),
            "op_growth": d_y.get("영업이익성장률(%)", None),
            "ni_growth": d_y.get("순이익성장률(%)", None),
        }
    return metrics


def _extract_key_metrics(metrics, years, ttm_derived):
    """핵심 요약 지표 추출"""
    latest = years[-1]
    m_latest = metrics.get(latest, {})

    # 매출 CAGR (첫 번째 0이 아닌 연도 사용)
    cagr = 0
    first_idx = None
    last_idx = None
    for i, y in enumerate(years):
        rev = metrics.get(y, {}).get("revenue", 0)
        if rev and rev > 0:
            if first_idx is None:
                first_idx = i
            last_idx = i

    if first_idx is not None and last_idx is not None and first_idx < last_idx:
        rev_first = metrics.get(years[first_idx], {}).get("revenue", 0)
        rev_last = metrics.get(years[last_idx], {}).get("revenue", 0)
        n = last_idx - first_idx
        if rev_first > 0 and rev_last > 0 and n > 0:
            cagr = ((rev_last / rev_first) ** (1 / n) - 1) * 100

    opm = m_latest.get("opm", 0)
    roe = m_latest.get("roe", 0)

    # 부채비율
    equity = m_latest.get("total_equity", 0)
    liab = m_latest.get("total_liabilities", 0)
    debt_ratio = (liab / equity * 100) if equity and equity > 0 else 0

    # FCF 추세
    fcf_values = [metrics.get(y, {}).get("fcf", 0) for y in years]
    fcf_positive = sum(1 for v in fcf_values if v and v > 0)
    if fcf_positive == len(years):
        fcf_trend = "지속 양호"
    elif fcf_positive >= len(years) * 0.7:
        fcf_trend = "대체로 양호"
    elif fcf_positive >= len(years) * 0.5:
        fcf_trend = "불안정"
    else:
        fcf_trend = "부진"

    # TTM 보정
    if ttm_derived:
        ttm_opm = ttm_derived.get("영업이익률(%)", 0)
        ttm_roe = ttm_derived.get("ROE(%)", 0)
        if ttm_opm:
            opm = ttm_opm
        if ttm_roe:
            roe = ttm_roe

    return {
        "revenue_cagr": round(cagr, 1),
        "opm_latest": round(opm, 1) if opm else 0,
        "roe_latest": round(roe, 1) if roe else 0,
        "debt_ratio_latest": round(debt_ratio, 0),
        "fcf_trend": fcf_trend,
    }


# ── 1. 매출 성장성 분석 ────────────────────────────────────────

def _analyze_revenue_growth(metrics, years, _find, ttm_derived=None):
    """매출 성장 추이 분석"""
    revenues = []
    growth_rates = []

    for y in years:
        r = metrics.get(y, {}).get("revenue", 0)
        revenues.append(r)
        g = metrics.get(y, {}).get("rev_growth", None)
        if g is not None:
            growth_rates.append(g)

    # CAGR 계산 (첫 번째 0이 아닌 연도 사용)
    first_nonzero_idx = None
    for i, r in enumerate(revenues):
        if r and r > 0:
            first_nonzero_idx = i
            break
    last_nonzero_idx = None
    for i in range(len(revenues) - 1, -1, -1):
        if revenues[i] and revenues[i] > 0:
            last_nonzero_idx = i
            break

    cagr = 0
    if (first_nonzero_idx is not None and last_nonzero_idx is not None
            and first_nonzero_idx < last_nonzero_idx):
        n = last_nonzero_idx - first_nonzero_idx
        if n > 0:
            cagr = ((revenues[last_nonzero_idx] / revenues[first_nonzero_idx])
                    ** (1 / n) - 1) * 100

    # 최근 성장률 추세
    avg_growth = sum(growth_rates) / len(growth_rates) if growth_rates else 0

    # 추세 방향 판단
    if len(growth_rates) >= 2:
        recent = growth_rates[-2:]
        early = growth_rates[:2] if len(growth_rates) >= 3 else growth_rates[:1]
        recent_avg = sum(recent) / len(recent)
        early_avg = sum(early) / len(early)

        if all(g < 0 for g in recent) and any(g >= 0 for g in early):
            trend = TREND_DOWN
        elif all(g >= 0 for g in recent) and any(g < 0 for g in early):
            trend = TREND_RECOVERY
        elif recent_avg > early_avg + 10:
            trend = TREND_UP
        elif recent_avg < early_avg - 10:
            trend = TREND_DOWN
        elif max(growth_rates) - min(growth_rates) > 30:
            trend = TREND_VOLATILE
        else:
            trend = TREND_FLAT if abs(avg_growth) < 5 else (TREND_UP if avg_growth > 0 else TREND_DOWN)
    else:
        trend = TREND_FLAT

    # 점수 산정 (-2 ~ +2)
    if cagr > 20:
        score = 2
    elif cagr > 8:
        score = 1
    elif cagr > -3:
        score = 0
    elif cagr > -10:
        score = -1
    else:
        score = -2

    # 추세 하락 중이면 감점
    if trend == TREND_DOWN:
        score = max(score - 1, -2)
    elif trend == TREND_RECOVERY:
        score = min(score + 1, 2)

    # 코멘트 생성
    title, comment, sub_items = _revenue_comment(
        cagr, avg_growth, trend, revenues, years, ttm_derived, growth_rates)

    return {
        "category": "매출 성장성",
        "trend": trend,
        "score": score,
        "title": title,
        "comment": comment,
        "sub_items": sub_items,
    }


def _revenue_comment(cagr, avg_growth, trend, revenues, years,
                     ttm_derived, growth_rates):
    """매출 성장성 코멘트 + sub_items 생성"""
    rev_eok_latest = revenues[-1] / 1e8 if revenues[-1] else 0
    n = len(years)

    if cagr > 20:
        title = f"매출 고성장 중 (CAGR {cagr:.1f}%)"
        comment = f"{n}년간 연평균 {cagr:.1f}% 성장. "
    elif cagr > 8:
        title = f"매출 안정 성장 (CAGR {cagr:.1f}%)"
        comment = f"{n}년간 연평균 {cagr:.1f}% 성장. "
    elif cagr > 0:
        title = f"매출 저성장 (CAGR {cagr:.1f}%)"
        comment = f"{n}년간 연평균 {cagr:.1f}% 저성장. "
    elif cagr > -5:
        title = f"매출 정체 (CAGR {cagr:.1f}%)"
        comment = f"{n}년간 매출이 정체 상태. "
    else:
        title = f"매출 감소세 (CAGR {cagr:.1f}%)"
        comment = f"{n}년간 연평균 {abs(cagr):.1f}% 감소. "

    if trend == TREND_DOWN:
        comment += "최근 매출 감소 추세가 지속되고 있어 주의가 필요합니다. "
    elif trend == TREND_RECOVERY:
        comment += "최근 매출이 반등하며 회복 징후를 보이고 있습니다. "
    elif trend == TREND_VOLATILE:
        comment += "매출 변동성이 크며 안정적 성장이 확인되지 않습니다. "
    elif trend == TREND_UP:
        comment += "성장 가속 추세로 긍정적 흐름입니다. "
    else:
        comment += "매출이 비교적 안정적으로 유지되고 있습니다. "

    # 성장 가속도 분석
    if len(growth_rates) >= 3:
        recent_2 = growth_rates[-2:]
        early = growth_rates[:-2]
        recent_avg = sum(recent_2) / len(recent_2)
        early_avg = sum(early) / len(early) if early else 0
        if recent_avg > early_avg + 5:
            comment += "최근 2년간 성장이 가속되고 있어 긍정적입니다. "
            accel_label = "가속"
        elif recent_avg < early_avg - 5:
            comment += "성장률이 둔화 추세로 모멘텀 약화가 우려됩니다. "
            accel_label = "감속"
        else:
            accel_label = "유지"
    else:
        accel_label = "-"

    # TTM 보정 코멘트
    if ttm_derived:
        ttm_rev = ttm_derived.get("_revenue", 0)
        if ttm_rev and revenues[-1]:
            ttm_vs_annual = (ttm_rev - revenues[-1]) / abs(revenues[-1]) * 100
            if abs(ttm_vs_annual) > 5:
                direction = "증가" if ttm_vs_annual > 0 else "감소"
                comment += f"TTM 기준 매출은 전년 대비 {abs(ttm_vs_annual):.0f}% {direction}. "

    # ── sub_items 생성 ──
    sub_items = [
        {"label": "매출 CAGR",
         "value": f"{cagr:+.1f}%",
         "assessment": "양호" if cagr > 8 else ("보통" if cagr > 0 else "부진")},
        {"label": "매출 규모",
         "value": f"{rev_eok_latest:,.0f}억원",
         "assessment": "-"},
    ]
    if len(growth_rates) >= 2:
        recent_avg_2y = sum(growth_rates[-2:]) / len(growth_rates[-2:])
        sub_items.append(
            {"label": "최근2년 평균 성장률",
             "value": f"{recent_avg_2y:+.1f}%",
             "assessment": "양호" if recent_avg_2y > 5 else (
                 "보통" if recent_avg_2y > 0 else "부진")})
    sub_items.append(
        {"label": "성장 가속도",
         "value": accel_label,
         "assessment": "양호" if accel_label == "가속" else (
             "주의" if accel_label == "감속" else "보통")})
    if len(growth_rates) >= 2:
        import statistics
        vol = statistics.stdev(growth_rates) if len(growth_rates) >= 2 else 0
        sub_items.append(
            {"label": "성장 변동성",
             "value": f"±{vol:.1f}%p",
             "assessment": "양호" if vol < 10 else ("보통" if vol < 25 else "주의")})

    return title, comment, sub_items


# ── 2. 수익성 추이 분석 ────────────────────────────────────────

def _analyze_profitability(metrics, years, ttm_derived=None):
    """수익성(OPM, ROE) 추이 분석"""
    opms = [metrics.get(y, {}).get("opm", 0) for y in years]
    roes = [metrics.get(y, {}).get("roe", 0) for y in years]
    net_incomes = [metrics.get(y, {}).get("net_income", 0) for y in years]

    latest_opm = opms[-1] if opms else 0
    latest_roe = roes[-1] if roes else 0

    # OPM 추세
    opm_trend = _calc_series_trend(opms)
    roe_trend = _calc_series_trend(roes)

    # 적자 전환 여부
    was_profitable = any(ni > 0 for ni in net_incomes[:-1]) if len(net_incomes) > 1 else False
    now_loss = net_incomes[-1] < 0 if net_incomes else False
    turned_loss = was_profitable and now_loss

    # 흑자 전환 여부
    was_loss = any(ni < 0 for ni in net_incomes[:-2]) if len(net_incomes) > 2 else False
    now_profit = net_incomes[-1] > 0 if net_incomes else False
    turned_profit = was_loss and now_profit and (
        len(net_incomes) >= 2 and net_incomes[-2] < 0
    )

    # 점수 산정
    if latest_roe > 15 and latest_opm > 10:
        score = 2
    elif latest_roe > 8 and latest_opm > 5:
        score = 1
    elif latest_roe > 0 and latest_opm > 0:
        score = 0
    elif latest_roe < 0 or latest_opm < 0:
        score = -1
    else:
        score = 0

    # 추세 보정
    if opm_trend == TREND_DOWN:
        score = max(score - 1, -2)
    elif opm_trend == TREND_UP and score < 2:
        score += 1

    if turned_loss:
        score = max(score - 1, -2)
    elif turned_profit:
        score = min(score + 1, 2)

    # 종합 추세
    if turned_loss:
        trend = TREND_DOWN
    elif turned_profit:
        trend = TREND_RECOVERY
    elif opm_trend == roe_trend:
        trend = opm_trend
    else:
        trend = opm_trend  # OPM 우선

    # 코멘트
    latest_roa = metrics.get(years[-1], {}).get("roa", 0) if years else 0
    latest_net_margin = metrics.get(years[-1], {}).get("net_margin", 0) if years else 0
    latest_tax_rate = metrics.get(years[-1], {}).get("eff_tax_rate", 0) if years else 0
    roas = [metrics.get(y, {}).get("roa", 0) for y in years]

    title, comment, sub_items = _profitability_comment(
        latest_opm, latest_roe, opm_trend, roe_trend,
        turned_loss, turned_profit, opms, roes, years, ttm_derived,
        latest_roa, latest_net_margin, latest_tax_rate, roas
    )

    return {
        "category": "수익성",
        "trend": trend,
        "score": score,
        "title": title,
        "comment": comment,
        "sub_items": sub_items,
    }


def _profitability_comment(opm, roe, opm_trend, roe_trend,
                           turned_loss, turned_profit,
                           opms, roes, years, ttm_derived,
                           roa=0, net_margin=0, tax_rate=0, roas=None):
    """수익성 코멘트 + sub_items 생성"""
    if turned_loss:
        title = "적자 전환 - 수익성 악화"
        comment = (
            f"최근 적자 전환. 영업이익률 {opm:.1f}%, ROE {roe:.1f}%. "
            "수익 구조 개선이 시급합니다. "
        )
    elif turned_profit:
        title = "흑자 전환 - 수익성 회복"
        comment = (
            f"적자에서 흑자로 전환. 영업이익률 {opm:.1f}%, ROE {roe:.1f}%. "
            "지속 가능성 확인이 필요합니다. "
        )
    elif opm > 15 and roe > 15:
        title = f"높은 수익성 유지 (OPM {opm:.1f}%, ROE {roe:.1f}%)"
        comment = "영업이익률과 ROE 모두 우수한 수준. "
    elif opm > 5 and roe > 8:
        title = f"양호한 수익성 (OPM {opm:.1f}%, ROE {roe:.1f}%)"
        comment = "안정적인 수익 구조를 유지하고 있습니다. "
    elif opm > 0 and roe > 0:
        title = f"낮은 수익성 (OPM {opm:.1f}%, ROE {roe:.1f}%)"
        comment = "흑자이나 수익성이 낮은 수준. 수익 구조 개선 필요. "
    else:
        title = f"적자 상태 (OPM {opm:.1f}%, ROE {roe:.1f}%)"
        comment = "영업적자 또는 순손실 상태. 근본적 구조개선 필요. "

    if opm_trend == TREND_UP:
        comment += "수익성이 개선 추세입니다. "
    elif opm_trend == TREND_DOWN:
        comment += "수익성이 하락 추세입니다. "
    else:
        comment += "수익성은 횡보 중입니다. "

    # 영업이익률 vs 순이익률 갭 분석
    if opm > 0 and net_margin > 0:
        gap = opm - net_margin
        if gap > 10:
            comment += f"영업이익률({opm:.1f}%)과 순이익률({net_margin:.1f}%) 갭이 {gap:.1f}%p로 영업외 비용 부담이 큽니다. "
        elif gap < 0:
            comment += f"순이익률({net_margin:.1f}%)이 영업이익률({opm:.1f}%)보다 높아 영업외 수익이 기여하고 있습니다. "

    # TTM 보정
    if ttm_derived:
        ttm_opm = ttm_derived.get("영업이익률(%)", 0)
        ttm_roe = ttm_derived.get("ROE(%)", 0)
        if ttm_opm and abs(ttm_opm - opm) > 3:
            direction = "개선" if ttm_opm > opm else "악화"
            comment += f"TTM 영업이익률 {ttm_opm:.1f}%로 {direction} 중. "

    # ── sub_items 생성 ──
    opm_dir = "↑" if opm_trend == TREND_UP else ("↓" if opm_trend == TREND_DOWN else "→")
    roe_dir = "↑" if roe_trend == TREND_UP else ("↓" if roe_trend == TREND_DOWN else "→")
    roa_trend_dir = _calc_series_trend(roas) if roas else TREND_FLAT
    roa_dir = "↑" if roa_trend_dir == TREND_UP else ("↓" if roa_trend_dir == TREND_DOWN else "→")

    sub_items = [
        {"label": "영업이익률(OPM)",
         "value": f"{opm:.1f}% {opm_dir}",
         "assessment": "양호" if opm > 10 else ("보통" if opm > 0 else "부진")},
        {"label": "ROE",
         "value": f"{roe:.1f}% {roe_dir}",
         "assessment": "양호" if roe > 10 else ("보통" if roe > 0 else "부진")},
        {"label": "ROA",
         "value": f"{roa:.1f}% {roa_dir}",
         "assessment": "양호" if roa > 5 else ("보통" if roa > 0 else "부진")},
        {"label": "OPM-순이익률 갭",
         "value": f"{opm - net_margin:.1f}%p" if net_margin else "-",
         "assessment": "양호" if (opm - net_margin) < 5 else (
             "보통" if (opm - net_margin) < 10 else "주의") if net_margin else "-"},
        {"label": "유효법인세율",
         "value": f"{tax_rate:.1f}%" if tax_rate else "-",
         "assessment": "-"},
    ]

    return title, comment, sub_items


# ── 3. 재무 안정성 분석 ────────────────────────────────────────

def _analyze_financial_stability(metrics, years, _find, fs, ttm_data=None):
    """재무 안정성 (부채비율, 자본잠식, 유동성) 분석"""
    debt_ratios = []
    equity_trend = []

    for y in years:
        m = metrics.get(y, {})
        equity = m.get("total_equity", 0)
        liab = m.get("total_liabilities", 0)
        if equity and equity > 0:
            debt_ratios.append(liab / equity * 100)
        else:
            debt_ratios.append(999)  # 자본잠식
        equity_trend.append(equity)

    latest_dr = debt_ratios[-1]
    # debt_ratios에 999가 포함된 경우 유효한 값만으로 추세 판단
    valid_drs = [d for d in debt_ratios if d < 900]
    if len(valid_drs) >= 2:
        dr_trend = _calc_series_trend(valid_drs, invert=True)
    else:
        dr_trend = TREND_FLAT

    # 자본잠식 여부
    latest_equity = equity_trend[-1] if equity_trend else 0
    latest_capital = metrics.get(years[-1], {}).get("capital", 0)
    is_impaired = latest_capital and latest_equity and latest_equity < latest_capital

    # 자사주 매입 기반 자본잠식 감지 (수익성 강 + 자본 부족 = 바이백)
    # 수익성이 높고(OPM>15%), 영업CF 양수이면 바이백 가능성 높음
    is_buyback_driven = False
    if is_impaired or (latest_equity and latest_equity < 0):
        # 유효한 연도에서 수익성 체크
        valid_years = [y for y in years if metrics.get(y, {}).get("revenue", 0) > 0]
        if valid_years:
            latest_valid = valid_years[-1]
            m_latest = metrics.get(latest_valid, {})
            if (m_latest.get("opm", 0) > 15
                    and m_latest.get("op_cf", 0) > 0
                    and m_latest.get("net_income", 0) > 0):
                is_buyback_driven = True

    # 순차입금 추세
    net_debts = [metrics.get(y, {}).get("net_debt", 0) for y in years]
    nd_trend = _calc_series_trend(net_debts, invert=True)

    # 점수 산정
    if is_buyback_driven:
        # 자사주 매입으로 인한 음수 자본은 위험이 아님
        score = 0  # 중립 (재무구조 특수)
    elif latest_dr > 500 or (is_impaired and latest_equity < 0):
        score = -2
    elif latest_dr > 300 or is_impaired:
        score = -1
    elif latest_dr < 50 and nd_trend in (TREND_UP, TREND_FLAT):
        score = 2
    elif latest_dr < 100:
        score = 1
    elif latest_dr < 200:
        score = 0
    else:
        score = -1

    # 추세 보정 (부채비율 증가하면 감점)
    if not is_buyback_driven:
        if dr_trend == TREND_DOWN:  # 부채비율 감소 = 좋음 (invert된 상태)
            score = min(score + 1, 2)
        elif dr_trend == TREND_UP:  # 부채비율 증가 = 나쁨
            score = max(score - 1, -2)

    trend = dr_trend

    # 코멘트
    latest_m = metrics.get(years[-1], {}) if years else {}
    title, comment, sub_items = _stability_comment(
        latest_dr, dr_trend, is_impaired, net_debts, equity_trend, years,
        ttm_data, is_buyback_driven, latest_m
    )

    return {
        "category": "재무 안정성",
        "trend": trend,
        "score": score,
        "title": title,
        "comment": comment,
        "sub_items": sub_items,
    }


def _stability_comment(debt_ratio, dr_trend, is_impaired, net_debts, equity_trend,
                        years, ttm_data, is_buyback_driven=False, latest_m=None):
    """재무 안정성 코멘트 + sub_items 생성"""
    if latest_m is None:
        latest_m = {}

    if is_buyback_driven:
        title = "자사주 매입 기반 음수자본 (정상)"
        comment = (
            "대규모 자사주 매입(바이백)으로 인한 음수 자본. "
            "수익성과 현금흐름이 우수하여 재무 위험이 아닌 주주환원 전략입니다. "
        )
    elif is_impaired:
        title = "자본잠식 상태 - 재무 위험"
        comment = (
            f"자본잠식 상태. 부채비율 {debt_ratio:.0f}%. "
            "자본 확충이 시급하며 관리종목/상장폐지 위험이 있습니다. "
        )
    elif debt_ratio > 300:
        title = f"부채비율 과다 ({debt_ratio:.0f}%)"
        comment = (
            f"부채비율 {debt_ratio:.0f}%로 재무 안정성 취약. "
            "금리 인상 시 이자 부담 급증 위험. "
        )
    elif debt_ratio > 200:
        title = f"부채비율 다소 높음 ({debt_ratio:.0f}%)"
        comment = f"부채비율 {debt_ratio:.0f}%. 업종 평균 대비 확인 필요. "
    elif debt_ratio < 50:
        title = f"매우 건전한 재무구조 (부채비율 {debt_ratio:.0f}%)"
        comment = "무차입 또는 극히 낮은 부채 수준. 재무 안정성 최상. "
    elif debt_ratio < 100:
        title = f"건전한 재무구조 (부채비율 {debt_ratio:.0f}%)"
        comment = "자기자본이 부채를 초과. 안정적 재무 구조. "
    else:
        title = f"보통 수준 재무구조 (부채비율 {debt_ratio:.0f}%)"
        comment = f"부채비율 {debt_ratio:.0f}%. 일반적 수준. "

    # 순차입금 추세
    latest_nd = net_debts[-1] if net_debts else 0
    if latest_nd and latest_nd < 0:
        comment += "순현금 상태(현금>부채)로 유동성 우수. "
    elif latest_nd and latest_nd > 0 and len(net_debts) >= 2:
        if net_debts[-1] > net_debts[-2]:
            comment += "순차입금 증가 중으로 유동성 주시 필요. "

    # 자본 추세
    equity_growth_label = "-"
    if len(equity_trend) >= 3 and all(e and e > 0 for e in equity_trend):
        if all(equity_trend[i] > equity_trend[i - 1] for i in range(1, len(equity_trend))):
            comment += "자기자본이 매년 증가하고 있어 긍정적입니다. "
            equity_growth_label = "매년 증가"
        elif all(equity_trend[i] < equity_trend[i - 1] for i in range(1, len(equity_trend))):
            comment += "자기자본이 지속 감소 중이며 주의가 필요합니다. "
            equity_growth_label = "감소 추세"
        else:
            equity_growth_label = "변동"

    # 단기채/장기채 구조 분석
    short_debt_r = latest_m.get("short_debt_ratio", 0)
    long_debt_r = latest_m.get("long_debt_ratio", 0)
    debt_structure_comment = ""
    if short_debt_r > 70 and latest_m.get("debt_total", 0) > 0:
        debt_structure_comment = "단기채 비중이 높아 차환 위험에 주의가 필요합니다. "
        comment += debt_structure_comment

    # ── sub_items 생성 ──
    dr_dir = "↑" if dr_trend == TREND_UP else ("↓" if dr_trend == TREND_DOWN else "→")
    nd_eok = latest_nd / 1e8 if latest_nd else 0
    nd_label = f"순현금 {abs(nd_eok):,.0f}억" if nd_eok < 0 else f"순차입금 {nd_eok:,.0f}억"

    sub_items = [
        {"label": "부채비율",
         "value": f"{debt_ratio:.0f}% {dr_dir}",
         "assessment": "양호" if debt_ratio < 100 else (
             "보통" if debt_ratio < 200 else "주의")},
        {"label": "순차입금/순현금",
         "value": nd_label,
         "assessment": "양호" if nd_eok < 0 else ("보통" if nd_eok < 1000 else "주의")},
        {"label": "차입금 만기 구조",
         "value": f"단기 {short_debt_r:.0f}% / 장기 {long_debt_r:.0f}%" if (
             short_debt_r or long_debt_r) else "무차입",
         "assessment": "양호" if short_debt_r < 50 else (
             "보통" if short_debt_r < 70 else "주의")},
        {"label": "자본 추세",
         "value": equity_growth_label,
         "assessment": "양호" if equity_growth_label == "매년 증가" else (
             "주의" if equity_growth_label == "감소 추세" else "보통")},
    ]
    if is_buyback_driven:
        sub_items.append(
            {"label": "자사주 매입",
             "value": "바이백 진행 중",
             "assessment": "참고"})

    return title, comment, sub_items


# ── 4. 현금흐름 품질 분석 ──────────────────────────────────────

def _analyze_cashflow_quality(metrics, years, ttm_derived=None):
    """현금흐름 품질 분석"""
    op_cfs = [metrics.get(y, {}).get("op_cf", 0) for y in years]
    fcfs = [metrics.get(y, {}).get("fcf", 0) for y in years]
    net_incomes = [metrics.get(y, {}).get("net_income", 0) for y in years]

    # 영업CF 양수 비율
    opcf_positive_count = sum(1 for v in op_cfs if v and v > 0)
    opcf_positive_ratio = opcf_positive_count / len(years) if years else 0

    # FCF 양수 비율
    fcf_positive_count = sum(1 for v in fcfs if v and v > 0)
    fcf_positive_ratio = fcf_positive_count / len(years) if years else 0

    # 영업CF vs 순이익 (이익의 질)
    # 영업CF > 순이익이면 이익의 질이 높음
    earnings_quality_count = 0
    for i, y in enumerate(years):
        opcf = op_cfs[i] or 0
        ni = net_incomes[i] or 0
        if ni > 0 and opcf >= ni:
            earnings_quality_count += 1

    earnings_quality = earnings_quality_count / max(
        sum(1 for ni in net_incomes if ni and ni > 0), 1
    )

    # 추세
    opcf_trend = _calc_series_trend(op_cfs)

    # 점수
    if opcf_positive_ratio >= 0.8 and fcf_positive_ratio >= 0.6 and earnings_quality >= 0.6:
        score = 2
    elif opcf_positive_ratio >= 0.6 and fcf_positive_ratio >= 0.4:
        score = 1
    elif opcf_positive_ratio >= 0.4:
        score = 0
    elif opcf_positive_ratio < 0.3:
        score = -2
    else:
        score = -1

    if opcf_trend == TREND_DOWN:
        score = max(score - 1, -2)

    trend = opcf_trend

    # CAPEX/매출, FCF마진, CF마진 계산
    revenues = [metrics.get(y, {}).get("revenue", 0) for y in years]
    capexes = [metrics.get(y, {}).get("capex", 0) for y in years]
    debt_totals = [metrics.get(y, {}).get("debt_total", 0) for y in years]

    title, comment, sub_items = _cashflow_comment(
        op_cfs, fcfs, opcf_positive_ratio, fcf_positive_ratio,
        earnings_quality, opcf_trend, years, ttm_derived,
        revenues, capexes, net_incomes, debt_totals
    )

    return {
        "category": "현금흐름 품질",
        "trend": trend,
        "score": score,
        "title": title,
        "comment": comment,
        "sub_items": sub_items,
    }


def _cashflow_comment(op_cfs, fcfs, opcf_ratio, fcf_ratio,
                      eq, opcf_trend, years, ttm_derived,
                      revenues=None, capexes=None, net_incomes=None,
                      debt_totals=None):
    """현금흐름 코멘트 + sub_items 생성"""
    latest_opcf = op_cfs[-1] if op_cfs else 0

    if opcf_ratio >= 0.8 and fcf_ratio >= 0.6:
        title = "현금흐름 우수"
        comment = f"영업현금흐름 {len(years)}년 중 {int(opcf_ratio * len(years))}년 양수. "
    elif opcf_ratio >= 0.6:
        title = "현금흐름 양호"
        comment = "영업현금흐름 대체로 양수. "
    elif opcf_ratio >= 0.4:
        title = "현금흐름 불안정"
        comment = "영업현금흐름이 들쭉날쭉하며 안정성 부족. "
    else:
        title = "현금흐름 부진"
        comment = "영업현금흐름 적자가 빈번하여 유동성 위험. "

    if eq >= 0.8:
        comment += "이익의 질 우수 (영업CF > 순이익). "
    elif eq < 0.4:
        comment += "이익의 질 낮음 (순이익 대비 영업CF 부족). "

    if opcf_trend == TREND_UP:
        comment += "현금 창출력이 개선되고 있습니다. "
    elif opcf_trend == TREND_DOWN:
        comment += "현금 창출력이 약화 추세입니다. "

    # CF 마진, FCF 마진, CAPEX 강도 계산
    latest_rev = revenues[-1] if revenues and revenues[-1] else 0
    cf_margin = (latest_opcf / latest_rev * 100) if latest_rev > 0 else 0
    latest_fcf = fcfs[-1] if fcfs else 0
    fcf_margin = (latest_fcf / latest_rev * 100) if latest_rev > 0 else 0
    latest_capex = capexes[-1] if capexes else 0
    capex_intensity = (abs(latest_capex) / latest_rev * 100) if latest_rev > 0 else 0

    if cf_margin > 15:
        comment += f"영업CF 마진 {cf_margin:.1f}%로 현금 창출력이 탁월합니다. "
    elif cf_margin < 0:
        comment += f"영업CF 마진 {cf_margin:.1f}%로 영업에서 현금이 유출되고 있습니다. "

    # FCF vs 부채 상환력
    latest_debt = debt_totals[-1] if debt_totals else 0
    if latest_fcf > 0 and latest_debt > 0:
        debt_payoff_years = latest_debt / latest_fcf
        if debt_payoff_years < 3:
            comment += f"FCF로 전체 부채 {debt_payoff_years:.1f}년 내 상환 가능. "
        elif debt_payoff_years > 10:
            comment += f"FCF 대비 부채 규모가 커 상환에 {debt_payoff_years:.0f}년 소요 전망. "

    # ── sub_items 생성 ──
    eq_label = "우수" if eq >= 0.8 else ("양호" if eq >= 0.6 else ("보통" if eq >= 0.4 else "부진"))
    sub_items = [
        {"label": "영업CF 마진",
         "value": f"{cf_margin:.1f}%",
         "assessment": "양호" if cf_margin > 10 else ("보통" if cf_margin > 0 else "부진")},
        {"label": "FCF 마진",
         "value": f"{fcf_margin:.1f}%",
         "assessment": "양호" if fcf_margin > 5 else ("보통" if fcf_margin > 0 else "부진")},
        {"label": "이익의 질(영업CF/NI)",
         "value": eq_label,
         "assessment": "양호" if eq >= 0.6 else ("보통" if eq >= 0.4 else "주의")},
        {"label": "CAPEX/매출 비율",
         "value": f"{capex_intensity:.1f}%",
         "assessment": "참고"},
    ]
    if latest_fcf > 0 and latest_debt > 0:
        debt_payoff = latest_debt / latest_fcf
        sub_items.append(
            {"label": "부채상환 능력",
             "value": f"{debt_payoff:.1f}년",
             "assessment": "양호" if debt_payoff < 5 else (
                 "보통" if debt_payoff < 10 else "주의")})

    return title, comment, sub_items


# ── 5. 밸류에이션 분석 ─────────────────────────────────────────

def _analyze_valuation(metrics, years, stock_data, ttm_derived=None):
    """밸류에이션 적정성 분석"""
    latest = years[-1]
    m = metrics.get(latest, {})
    per = m.get("per", 0)
    pbr = m.get("pbr", 0)

    # TTM 보정
    if ttm_derived:
        ttm_per = ttm_derived.get("PER", 0)
        ttm_pbr = ttm_derived.get("PBR", 0)
        if ttm_per and ttm_per > 0:
            per = ttm_per
        if ttm_pbr and ttm_pbr > 0:
            pbr = ttm_pbr

    # PER 추세
    pers = []
    for y in years:
        p = metrics.get(y, {}).get("per", 0)
        if p and p > 0:
            pers.append(p)

    avg_per = sum(pers) / len(pers) if pers else 0

    # 점수
    if per and per > 0:
        if per < 8 and pbr and pbr < 1:
            score = 2   # 저평가
            trend = TREND_UP
        elif per < 15 and pbr and pbr < 2:
            score = 1
            trend = TREND_FLAT
        elif per < 30:
            score = 0
            trend = TREND_FLAT
        elif per < 50:
            score = -1
            trend = TREND_DOWN
        else:
            score = -1
            trend = TREND_DOWN
    elif per and per < 0:
        score = -2  # 적자
        trend = TREND_DOWN
    else:
        score = 0
        trend = TREND_FLAT

    # PFCR, PER 밴드 계산
    latest_pfcr = metrics.get(years[-1], {}).get("pfcr", 0) if years else 0
    if ttm_derived:
        ttm_pfcr = ttm_derived.get("PFCR", 0)
        if ttm_pfcr and ttm_pfcr > 0:
            latest_pfcr = ttm_pfcr

    # PBR 추세
    pbrs = []
    for y in years:
        p = metrics.get(y, {}).get("pbr", 0)
        if p and p > 0:
            pbrs.append(p)
    avg_pbr = sum(pbrs) / len(pbrs) if pbrs else 0

    per_min = min(pers) if pers else 0
    per_max = max(pers) if pers else 0

    title, comment, sub_items = _valuation_comment(
        per, pbr, avg_per, stock_data, ttm_derived,
        latest_pfcr, avg_pbr, per_min, per_max, pers, pbrs)

    return {
        "category": "밸류에이션",
        "trend": trend,
        "score": score,
        "title": title,
        "comment": comment,
        "sub_items": sub_items,
    }


def _valuation_comment(per, pbr, avg_per, stock_data, ttm_derived,
                       pfcr=0, avg_pbr=0, per_min=0, per_max=0,
                       pers=None, pbrs=None):
    """밸류에이션 코멘트 + sub_items"""
    price = stock_data.get("price", 0)

    if per and per < 0:
        title = "밸류에이션 산정 불가 (적자)"
        comment = "순이익 적자로 PER 산정 불가. 주가 판단 시 PBR, FCF 등 대안 지표 활용 필요. "
    elif per and per < 8 and pbr and 0 < pbr < 1:
        title = f"저평가 구간 (PER {per:.1f}배, PBR {pbr:.2f}배)"
        comment = (
            f"PER {per:.1f}배, PBR {pbr:.2f}배로 "
            "시장 대비 저평가 영역. 밸류 트랩 여부 확인 필요. "
        )
    elif per and per < 15:
        title = f"합리적 밸류에이션 (PER {per:.1f}배)"
        comment = f"PER {per:.1f}배"
        if pbr:
            comment += f", PBR {pbr:.2f}배"
        comment += ". 적정 수준의 밸류에이션. "
    elif per and per < 30:
        title = f"다소 높은 밸류에이션 (PER {per:.1f}배)"
        comment = f"PER {per:.1f}배로 시장 평균 대비 높은 편. 성장성 프리미엄 반영 여부 확인. "
    elif per and per >= 30:
        title = f"고평가 구간 (PER {per:.1f}배)"
        comment = f"PER {per:.1f}배로 높은 밸류에이션. 높은 성장 기대가 반영된 가격. "
    else:
        title = "밸류에이션 데이터 부족"
        comment = "PER/PBR 산정 데이터 부족. "

    # 과거 평균 대비
    if avg_per and per and per > 0 and avg_per > 0:
        if per < avg_per * 0.7:
            comment += f"과거 평균 PER({avg_per:.0f}배) 대비 할인 거래 중. "
        elif per > avg_per * 1.3:
            comment += f"과거 평균 PER({avg_per:.0f}배) 대비 프리미엄 거래 중. "
        else:
            comment += f"과거 평균 PER({avg_per:.0f}배)과 유사한 수준. "

    # PFCR 코멘트
    if pfcr and pfcr > 0:
        if pfcr < 15:
            comment += f"PFCR {pfcr:.1f}배로 현금흐름 기준 적정 수준. "
        elif pfcr > 30:
            comment += f"PFCR {pfcr:.1f}배로 현금흐름 대비 고평가. "

    # ── sub_items 생성 ──
    sub_items = []
    if per and per > 0:
        per_vs_avg = ""
        if avg_per and avg_per > 0:
            ratio = per / avg_per
            if ratio < 0.8:
                per_vs_avg = " (평균 대비 할인)"
            elif ratio > 1.2:
                per_vs_avg = " (평균 대비 프리미엄)"
        sub_items.append(
            {"label": "PER",
             "value": f"{per:.1f}배{per_vs_avg}",
             "assessment": "양호" if per < 15 else ("보통" if per < 25 else "주의")})
    if pbr and pbr > 0:
        sub_items.append(
            {"label": "PBR",
             "value": f"{pbr:.2f}배",
             "assessment": "양호" if pbr < 1.5 else ("보통" if pbr < 3 else "주의")})
    if pfcr and pfcr > 0:
        sub_items.append(
            {"label": "PFCR",
             "value": f"{pfcr:.1f}배",
             "assessment": "양호" if pfcr < 15 else ("보통" if pfcr < 30 else "주의")})
    if pers and len(pers) >= 2:
        sub_items.append(
            {"label": "PER 밴드",
             "value": f"{per_min:.0f}~{per_max:.0f}배 (평균 {avg_per:.0f})",
             "assessment": "-"})

    return title, comment, sub_items


# ── 6. 효율성 분석 ──────────────────────────────────────────────

def _analyze_efficiency(metrics, years, ttm_derived=None):
    """자본 효율성 분석 (자산회전율, ROIC 근사, CAPEX 집중도, 운전자본 효율)"""
    asset_turnovers = [metrics.get(y, {}).get("asset_turnover", 0) for y in years]
    wc_ratios = [metrics.get(y, {}).get("wc_ratio", 0) for y in years]
    capexes = [metrics.get(y, {}).get("capex", 0) for y in years]
    revenues = [metrics.get(y, {}).get("revenue", 0) for y in years]
    op_profits = [metrics.get(y, {}).get("op_profit", 0) for y in years]
    tax_rates = [metrics.get(y, {}).get("eff_tax_rate", 0) for y in years]
    equities = [metrics.get(y, {}).get("total_equity", 0) for y in years]
    debt_totals = [metrics.get(y, {}).get("debt_total", 0) for y in years]

    latest = years[-1] if years else None
    m_latest = metrics.get(latest, {}) if latest else {}

    # 자산회전율
    latest_at = m_latest.get("asset_turnover", 0)
    at_trend = _calc_series_trend([v for v in asset_turnovers if v > 0])

    # ROIC 근사 = NOPAT / Invested Capital
    # NOPAT = 영업이익 × (1 - 세율)
    # Invested Capital ≈ 자본총계 + 이자발생부채
    roics = []
    for y in years:
        m = metrics.get(y, {})
        op_p = m.get("op_profit", 0) or 0
        tax_r = m.get("eff_tax_rate", 0) or 0
        eq = m.get("total_equity", 0) or 0
        dt = m.get("debt_total", 0) or 0
        invested = eq + dt
        if invested > 0 and op_p > 0:
            nopat = op_p * (1 - tax_r / 100) if tax_r > 0 else op_p * 0.75
            roics.append(nopat / invested * 100)
        elif invested > 0 and op_p < 0:
            roics.append(op_p / invested * 100)
        else:
            roics.append(0)

    latest_roic = roics[-1] if roics else 0
    roic_trend = _calc_series_trend([r for r in roics if r != 0])

    # CAPEX 집중도 (|CAPEX|/매출)
    capex_intensities = []
    for i, y in enumerate(years):
        rev = revenues[i] or 0
        cap = capexes[i] or 0
        if rev > 0 and cap != 0:
            capex_intensities.append(abs(cap) / rev * 100)
        else:
            capex_intensities.append(0)
    latest_capex_intensity = capex_intensities[-1] if capex_intensities else 0

    # 운전자본비율 추세
    valid_wc = [w for w in wc_ratios if w != 0]
    wc_trend = _calc_series_trend(valid_wc) if len(valid_wc) >= 2 else TREND_FLAT

    # 점수 산정
    if latest_at > 1.0 and latest_roic > 10:
        score = 2
    elif latest_at > 0.5 and latest_roic > 5:
        score = 1
    elif latest_at > 0.3 and latest_roic > 0:
        score = 0
    elif latest_roic < 0:
        score = -1
    else:
        score = 0

    if at_trend == TREND_UP:
        score = min(score + 1, 2)
    elif at_trend == TREND_DOWN:
        score = max(score - 1, -2)

    # 종합 추세
    if at_trend == roic_trend:
        trend = at_trend
    else:
        trend = at_trend

    # 코멘트
    comment = ""
    if latest_at > 1.0:
        title = f"높은 자본 효율 (자산회전율 {latest_at:.2f}배)"
        comment += f"자산회전율 {latest_at:.2f}배로 자산 대비 매출 창출력이 우수합니다. "
    elif latest_at > 0.5:
        title = f"보통 자본 효율 (자산회전율 {latest_at:.2f}배)"
        comment += f"자산회전율 {latest_at:.2f}배로 일반적 수준. "
    elif latest_at > 0:
        title = f"낮은 자본 효율 (자산회전율 {latest_at:.2f}배)"
        comment += f"자산회전율 {latest_at:.2f}배로 자본 집약적 산업이거나 효율성 개선 필요. "
    else:
        title = "효율성 데이터 부족"
        comment += "자산회전율 산정 불가. "

    if latest_roic > 10:
        comment += f"ROIC {latest_roic:.1f}%로 투자자본 대비 높은 수익을 창출 중. "
    elif latest_roic > 0:
        comment += f"ROIC {latest_roic:.1f}%로 투자자본 대비 적정 수익. "
    elif latest_roic < 0:
        comment += f"ROIC {latest_roic:.1f}%로 투자자본 대비 가치 파괴 중. "

    if latest_capex_intensity > 15:
        comment += f"CAPEX/매출 {latest_capex_intensity:.1f}%로 자본 집약적 구조. 진입장벽 높지만 유연성 제한. "
    elif latest_capex_intensity > 5:
        comment += f"CAPEX/매출 {latest_capex_intensity:.1f}%로 적정 투자 수준. "

    if at_trend == TREND_UP:
        comment += "자산 효율이 개선되고 있습니다. "
    elif at_trend == TREND_DOWN:
        comment += "자산 효율이 하락 추세입니다. "

    # sub_items
    at_dir = "↑" if at_trend == TREND_UP else ("↓" if at_trend == TREND_DOWN else "→")
    roic_dir = "↑" if roic_trend == TREND_UP else ("↓" if roic_trend == TREND_DOWN else "→")
    wc_dir = "↑" if wc_trend == TREND_UP else ("↓" if wc_trend == TREND_DOWN else "→")

    sub_items = [
        {"label": "자산회전율",
         "value": f"{latest_at:.2f}배 {at_dir}",
         "assessment": "양호" if latest_at > 0.8 else ("보통" if latest_at > 0.3 else "부진")},
        {"label": "ROIC",
         "value": f"{latest_roic:.1f}% {roic_dir}",
         "assessment": "양호" if latest_roic > 10 else ("보통" if latest_roic > 0 else "부진")},
        {"label": "CAPEX/매출",
         "value": f"{latest_capex_intensity:.1f}%",
         "assessment": "참고"},
        {"label": "운전자본 효율",
         "value": f"{m_latest.get('wc_ratio', 0):.1f}% {wc_dir}",
         "assessment": "참고"},
    ]

    return {
        "category": "효율성",
        "trend": trend,
        "score": score,
        "title": title,
        "comment": comment,
        "sub_items": sub_items,
    }


# ── 종합 분류 ──────────────────────────────────────────────────

def _classify_situation(details, metrics, years, avg_score):
    """종합 점수 기반 기업 상황 분류"""
    revenue_d = details[0]   # 매출 성장성
    profit_d = details[1]    # 수익성
    stability_d = details[2] # 재무 안정성
    cashflow_d = details[3]  # 현금흐름
    valuation_d = details[4] # 밸류에이션
    # details[5] = 효율성 (있으면)

    # 위기 판별 (단, 수익성+현금흐름이 양호하면 제외 — 바이백 등 특수 상황)
    strong_fundamentals = (profit_d["score"] >= 1 and cashflow_d["score"] >= 1)
    if not strong_fundamentals:
        if (stability_d["score"] <= -2
                or (profit_d["score"] <= -2 and cashflow_d["score"] <= -1)):
            return SITUATION_CRISIS

    # 하향세 판별
    negative_count = sum(1 for d in details if d["score"] < 0)
    if negative_count >= 3 and not strong_fundamentals:
        return SITUATION_DECLINING

    # 턴어라운드 판별 (과거 실적 부진 → 회복이 있어야 진정한 턴어라운드)
    # 수익성이 항상 높았던 기업은 턴어라운드가 아님
    had_weakness = (profit_d["score"] <= 0 or revenue_d["score"] <= 0)
    if had_weakness and (
            profit_d["trend"] == TREND_RECOVERY
            or revenue_d["trend"] == TREND_RECOVERY):
        if avg_score >= -0.5:
            return SITUATION_TURNAROUND

    # 고성장 판별
    if revenue_d["score"] >= 2 and profit_d["score"] >= 1:
        return SITUATION_HIGH_GROWTH

    # 안정 성장 판별
    if (revenue_d["score"] >= 1 and profit_d["score"] >= 0
            and stability_d["score"] >= 0):
        return SITUATION_STABLE_GROWTH

    # 가치주 판별
    if (revenue_d["score"] <= 0 and profit_d["score"] >= 0
            and stability_d["score"] >= 1 and valuation_d["score"] >= 1):
        return SITUATION_VALUE

    # 성숙/정체
    if avg_score >= -0.5 and avg_score <= 0.5:
        return SITUATION_MATURE

    # 나머지: 점수 기반
    if avg_score >= 1:
        return SITUATION_STABLE_GROWTH
    elif avg_score > 0:
        return SITUATION_MATURE
    elif avg_score > -1:
        return SITUATION_DECLINING
    else:
        return SITUATION_CRISIS


def _calc_confidence(years, metrics):
    """진단 신뢰도 계산 (데이터 충분성 기반)"""
    n = len(years)
    # 데이터 포인트 확인
    data_completeness = 0
    for y in years:
        m = metrics.get(y, {})
        if m.get("revenue"):
            data_completeness += 1
        if m.get("op_profit") is not None:
            data_completeness += 1
        if m.get("total_equity"):
            data_completeness += 1
        if m.get("op_cf"):
            data_completeness += 1

    max_points = n * 4
    ratio = data_completeness / max_points if max_points else 0

    # 연도 수 가중
    if n >= 5:
        year_factor = 1.0
    elif n >= 3:
        year_factor = 0.8
    else:
        year_factor = 0.6

    return round(min(ratio * year_factor, 1.0), 2)


def _generate_summary(situation, details, metrics, years,
                      ttm_derived, stock_data):
    """종합 진단 요약문 생성"""
    label = SITUATION_LABELS.get(situation, "분석 불가")
    n = len(years)
    latest = years[-1]
    m = metrics.get(latest, {})

    revenue_d = details[0]
    profit_d = details[1]
    stability_d = details[2]
    cashflow_d = details[3]

    parts = []

    # 첫 문장: 종합 진단
    if situation == SITUATION_HIGH_GROWTH:
        parts.append(
            f"이 기업은 최근 {n}년간 매출과 수익이 빠르게 성장하는 '고성장' 단계에 있습니다."
        )
    elif situation == SITUATION_STABLE_GROWTH:
        parts.append(
            f"이 기업은 최근 {n}년간 안정적인 성장을 보이며, "
            "수익성과 재무 건전성이 양호한 '안정 성장' 기업입니다."
        )
    elif situation == SITUATION_VALUE:
        parts.append(
            f"이 기업은 성장은 다소 둔화되었으나 재무 안정성이 높고 "
            "밸류에이션이 낮아 '가치주'로 분류됩니다."
        )
    elif situation == SITUATION_TURNAROUND:
        parts.append(
            f"이 기업은 과거 부진에서 벗어나 실적이 회복되는 '턴어라운드' 국면에 있습니다."
        )
    elif situation == SITUATION_MATURE:
        parts.append(
            f"이 기업은 성장이 정체되어 있으며, 현재 '성숙/정체' 단계로 판단됩니다."
        )
    elif situation == SITUATION_DECLINING:
        parts.append(
            f"이 기업은 주요 재무지표가 악화 추세를 보이며 '하향세'에 있습니다."
        )
    elif situation == SITUATION_CRISIS:
        parts.append(
            f"이 기업은 재무 안정성과 수익성이 심각하게 훼손된 '위기' 상태입니다."
        )
    else:
        parts.append(f"최근 {n}년간 재무 데이터를 기반으로 분석했습니다.")

    # 두 번째 문장: 핵심 지표 요약
    strengths = [d["category"] for d in details if d["score"] >= 1]
    weaknesses = [d["category"] for d in details if d["score"] <= -1]

    if strengths:
        parts.append(f"강점: {', '.join(strengths)}.")
    if weaknesses:
        parts.append(f"약점: {', '.join(weaknesses)}.")

    # 세 번째 문장: 결론/전망
    if situation in (SITUATION_HIGH_GROWTH, SITUATION_STABLE_GROWTH):
        if cashflow_d["score"] >= 1:
            parts.append("현금흐름도 견조하여 성장의 질이 양호합니다.")
        else:
            parts.append("다만 현금흐름의 질은 확인이 필요합니다.")
    elif situation == SITUATION_TURNAROUND:
        parts.append("실적 개선의 지속 여부를 분기별로 모니터링하는 것이 중요합니다.")
    elif situation in (SITUATION_DECLINING, SITUATION_CRISIS):
        parts.append("투자 시 높은 리스크를 감안해야 하며, 실적 반등 신호 확인이 필요합니다.")
    elif situation == SITUATION_VALUE:
        parts.append(
            "다만 밸류 트랩(성장 없는 저평가 지속) 가능성도 고려해야 합니다."
        )

    return " ".join(parts)


# ── 유틸리티 ───────────────────────────────────────────────────

def _calc_series_trend(values, invert=False):
    """시계열 데이터의 추세 방향 판단

    Args:
        values: 시간순 데이터 리스트
        invert: True이면 값이 감소하는 것이 좋음 (부채비율 등)

    Returns:
        TREND_UP | TREND_DOWN | TREND_FLAT | TREND_VOLATILE | TREND_RECOVERY
    """
    clean = [v for v in values if v is not None and v != 0]
    if len(clean) < 2:
        return TREND_FLAT

    # 선형 회귀 간이 추세
    n = len(clean)
    x_mean = (n - 1) / 2
    y_mean = sum(clean) / n

    numerator = sum((i - x_mean) * (clean[i] - y_mean) for i in range(n))
    denominator = sum((i - x_mean) ** 2 for i in range(n))

    if denominator == 0 or y_mean == 0:
        return TREND_FLAT

    slope = numerator / denominator
    # 기울기를 평균 대비 비율로 정규화
    normalized_slope = slope / abs(y_mean) * 100

    # 변동성 체크
    if n >= 3:
        changes = [clean[i] - clean[i - 1] for i in range(1, n)]
        sign_changes = sum(
            1 for i in range(1, len(changes))
            if (changes[i] > 0) != (changes[i - 1] > 0)
        )
        if sign_changes >= len(changes) * 0.6 and len(changes) >= 2:
            return TREND_VOLATILE

    # 회복 패턴: 하락 후 반등
    if n >= 3:
        mid = n // 2
        first_half = clean[:mid]
        second_half = clean[mid:]
        if (sum(first_half) / len(first_half) < sum(second_half) / len(second_half)
                and clean[-1] > clean[-2] and clean[0] > clean[mid - 1]):
            return TREND_RECOVERY

    if invert:
        normalized_slope = -normalized_slope

    if normalized_slope > 5:
        return TREND_UP
    elif normalized_slope < -5:
        return TREND_DOWN
    else:
        return TREND_FLAT


def _apply_ttm_score_adjustment(details, metrics, years, ttm_derived):
    """TTM 데이터로 도메인 점수 보정 (5%p 이상 변화 시)"""
    if not ttm_derived or not years:
        return

    latest = years[-1]
    m = metrics.get(latest, {})

    # 수익성 도메인 (index 1)
    ttm_opm = ttm_derived.get("영업이익률(%)", 0)
    annual_opm = m.get("opm", 0)
    if ttm_opm and annual_opm:
        diff = ttm_opm - annual_opm
        if diff > 5:
            details[1]["score"] = min(details[1]["score"] + 1, 2)
        elif diff < -5:
            details[1]["score"] = max(details[1]["score"] - 1, -2)

    ttm_roe = ttm_derived.get("ROE(%)", 0)
    annual_roe = m.get("roe", 0)
    if ttm_roe and annual_roe:
        diff = ttm_roe - annual_roe
        if diff > 5:
            details[1]["score"] = min(details[1]["score"] + 1, 2)
        elif diff < -5:
            details[1]["score"] = max(details[1]["score"] - 1, -2)


def _analyze_quarterly_momentum(quarterly_derived, quarterly_keys):
    """분기 QoQ 모멘텀 분석"""
    if not quarterly_derived or not quarterly_keys or len(quarterly_keys) < 4:
        return "데이터 없음"

    # 최근 4분기의 매출/영업이익 QoQ 성장률 패턴
    recent_keys = quarterly_keys[-4:]
    revenues = []
    for qk in recent_keys:
        qd = quarterly_derived.get(qk, {})
        rev = qd.get("매출액", 0)
        revenues.append(rev if rev else 0)

    if len(revenues) < 4 or not all(r > 0 for r in revenues):
        return "데이터 부족"

    # QoQ 변화율
    qoq_changes = []
    for i in range(1, len(revenues)):
        if revenues[i - 1] > 0:
            change = (revenues[i] - revenues[i - 1]) / revenues[i - 1] * 100
            qoq_changes.append(change)

    if len(qoq_changes) < 2:
        return "데이터 부족"

    recent_2 = qoq_changes[-2:]
    if all(c > 0 for c in recent_2):
        return "양호"
    elif all(c < 0 for c in recent_2):
        return "약화"
    else:
        return "혼조"


def _build_dupont_analysis(metrics, years):
    """DuPont ROE 분해"""
    valid_years = []
    net_margins = []
    asset_turnovers = []
    equity_multipliers = []
    roes = []

    for y in years:
        m = metrics.get(y, {})
        nm = m.get("net_margin", 0)
        at = m.get("asset_turnover", 0)
        em = m.get("equity_multiplier", 0)
        roe_val = m.get("roe", 0)
        # 유효한 데이터만
        if at > 0 and em > 0:
            valid_years.append(y)
            net_margins.append(round(nm, 2))
            asset_turnovers.append(round(at, 3))
            equity_multipliers.append(round(em, 2))
            roes.append(round(roe_val, 2))

    if len(valid_years) < 2:
        return None

    # 주 변동 요인 식별
    nm_change = net_margins[-1] - net_margins[0] if len(net_margins) >= 2 else 0
    at_change = asset_turnovers[-1] - asset_turnovers[0] if len(asset_turnovers) >= 2 else 0
    em_change = equity_multipliers[-1] - equity_multipliers[0] if len(equity_multipliers) >= 2 else 0

    # 정규화 변화율
    nm_pct = abs(nm_change / net_margins[0] * 100) if net_margins[0] != 0 else 0
    at_pct = abs(at_change / asset_turnovers[0] * 100) if asset_turnovers[0] != 0 else 0
    em_pct = abs(em_change / equity_multipliers[0] * 100) if equity_multipliers[0] != 0 else 0

    max_change = max(nm_pct, at_pct, em_pct)
    if max_change == nm_pct:
        if nm_change > 0:
            main_driver = "순이익률 개선"
        else:
            main_driver = "순이익률 하락"
    elif max_change == at_pct:
        if at_change > 0:
            main_driver = "자산회전율 개선"
        else:
            main_driver = "자산회전율 하락"
    else:
        if em_change > 0:
            main_driver = "재무레버리지 확대"
        else:
            main_driver = "재무레버리지 축소"

    # 코멘트
    roe_first = roes[0]
    roe_last = roes[-1]
    if roe_last > roe_first:
        comment = f"ROE가 {roe_first:.1f}%에서 {roe_last:.1f}%로 상승했으며, 주 요인은 {main_driver}입니다."
    elif roe_last < roe_first:
        comment = f"ROE가 {roe_first:.1f}%에서 {roe_last:.1f}%로 하락했으며, 주 요인은 {main_driver}입니다."
    else:
        comment = f"ROE가 {roe_last:.1f}% 수준으로 유지되고 있으며, 구성 요소 간 변동은 {main_driver}이(가) 주도했습니다."

    return {
        "years": valid_years,
        "net_margin": net_margins,
        "asset_turnover": asset_turnovers,
        "equity_multiplier": equity_multipliers,
        "roe": roes,
        "main_driver": main_driver,
        "comment": comment,
    }


def _build_yearly_table(metrics, years):
    """연도별 핵심 지표 추이 테이블 데이터"""
    rows = []

    def _add_row(label, key, divisor=1, fmt_type="number"):
        values = []
        for y in years:
            v = metrics.get(y, {}).get(key, 0) or 0
            if divisor != 1 and v:
                v = v / divisor
            values.append(round(v, 1) if fmt_type == "pct" else (
                round(v, 2) if fmt_type == "ratio" else round(v, 0)))
        rows.append({"label": label, "values": values})

    _add_row("매출액(억원)", "revenue", divisor=1e8)
    _add_row("영업이익(억원)", "op_profit", divisor=1e8)
    _add_row("순이익(억원)", "net_income", divisor=1e8)
    _add_row("영업이익률(%)", "opm", fmt_type="pct")
    _add_row("ROE(%)", "roe", fmt_type="pct")
    _add_row("ROA(%)", "roa", fmt_type="pct")

    # 부채비율 별도 계산
    debt_ratio_values = []
    for y in years:
        m = metrics.get(y, {})
        eq = m.get("total_equity", 0) or 0
        liab = m.get("total_liabilities", 0) or 0
        if eq > 0:
            debt_ratio_values.append(round(liab / eq * 100, 0))
        else:
            debt_ratio_values.append(0)
    rows.append({"label": "부채비율(%)", "values": debt_ratio_values})

    _add_row("FCF(억원)", "fcf", divisor=1e8)
    _add_row("PER(배)", "per", fmt_type="pct")
    _add_row("PBR(배)", "pbr", fmt_type="ratio")

    return {
        "years": years,
        "rows": rows,
    }


def _build_investment_checklist(details, metrics, years, ttm_derived, key_metrics):
    """투자 체크리스트 자동 생성"""
    checklist = []

    # 매출 성장 지속 가능성
    rev_d = details[0]
    rev_status = "양호" if rev_d["score"] >= 1 else ("보통" if rev_d["score"] >= 0 else "부진")
    cagr = key_metrics.get("revenue_cagr", 0)
    checklist.append({
        "question": "매출 성장 지속 가능성은?",
        "status": rev_status,
        "detail": f"CAGR {cagr:+.1f}%, 추세 {rev_d['trend']}",
    })

    # 수익성 유지 가능
    prof_d = details[1]
    prof_status = "양호" if prof_d["score"] >= 1 else ("보통" if prof_d["score"] >= 0 else "부진")
    opm = key_metrics.get("opm_latest", 0)
    roe = key_metrics.get("roe_latest", 0)
    checklist.append({
        "question": "수익성 추세는 유지 가능한가?",
        "status": prof_status,
        "detail": f"OPM {opm:.1f}%, ROE {roe:.1f}%",
    })

    # 현금흐름으로 부채 상환 가능
    cf_d = details[3]
    cf_status = "양호" if cf_d["score"] >= 1 else ("보통" if cf_d["score"] >= 0 else "부진")
    checklist.append({
        "question": "현금흐름으로 부채 상환 가능한가?",
        "status": cf_status,
        "detail": f"FCF 추세: {key_metrics.get('fcf_trend', '-')}",
    })

    # 재무 안정성
    stab_d = details[2]
    stab_status = "양호" if stab_d["score"] >= 1 else ("주의" if stab_d["score"] <= -1 else "보통")
    dr = key_metrics.get("debt_ratio_latest", 0)
    checklist.append({
        "question": "재무구조는 안정적인가?",
        "status": stab_status,
        "detail": f"부채비율 {dr:.0f}%",
    })

    # 현재 주가 적정성
    val_d = details[4]
    if val_d["score"] >= 2:
        val_status = "저평가"
    elif val_d["score"] >= 0:
        val_status = "적정"
    else:
        val_status = "고평가"
    checklist.append({
        "question": "현재 주가는 적정한가?",
        "status": val_status,
        "detail": val_d.get("title", ""),
    })

    # 자본 효율성 (효율성 도메인 = index 5)
    if len(details) > 5:
        eff_d = details[5]
        eff_status = "양호" if eff_d["score"] >= 1 else (
            "보통" if eff_d["score"] >= 0 else "부진")
        checklist.append({
            "question": "자본 효율적으로 운용되고 있는가?",
            "status": eff_status,
            "detail": eff_d.get("title", ""),
        })

    return checklist


def _build_strengths_weaknesses(details, metrics, years, ttm_derived, key_metrics):
    """강점/약점/기회/위험 자동 추출"""
    strengths = []
    weaknesses = []
    opportunities = []
    risks = []

    cagr = key_metrics.get("revenue_cagr", 0)
    opm = key_metrics.get("opm_latest", 0)
    roe = key_metrics.get("roe_latest", 0)
    dr = key_metrics.get("debt_ratio_latest", 0)
    fcf_t = key_metrics.get("fcf_trend", "-")
    momentum = key_metrics.get("quarterly_momentum", "-")

    # 매출
    if cagr > 10:
        strengths.append(f"매출 고성장 지속 (CAGR {cagr:+.1f}%)")
    elif cagr < -3:
        weaknesses.append(f"매출 감소 추세 (CAGR {cagr:+.1f}%)")
    elif cagr < 3:
        weaknesses.append(f"매출 성장 둔화 (CAGR {cagr:+.1f}%)")

    # 수익성
    if opm > 15:
        strengths.append(f"높은 영업이익률 ({opm:.1f}%)")
    elif opm < 0:
        weaknesses.append(f"영업적자 (OPM {opm:.1f}%)")
    elif opm < 5:
        weaknesses.append(f"낮은 영업이익률 ({opm:.1f}%)")

    if roe > 15:
        strengths.append(f"높은 ROE ({roe:.1f}%)")
    elif roe < 0:
        weaknesses.append(f"음수 ROE ({roe:.1f}%)")

    # 재무 안정성
    if dr < 50:
        strengths.append(f"매우 낮은 부채비율 ({dr:.0f}%)")
    elif dr > 200:
        weaknesses.append(f"높은 부채비율 ({dr:.0f}%)")
    elif dr > 300:
        risks.append(f"부채비율 과다 ({dr:.0f}%) — 관리종목 리스크")

    # 현금흐름
    if fcf_t in ("지속 양호", "대체로 양호"):
        strengths.append(f"FCF {fcf_t}")
    elif fcf_t == "부진":
        weaknesses.append("FCF 부진 — 현금 창출력 약화")

    # 밸류에이션
    val_d = details[4]
    if val_d["score"] >= 2:
        strengths.append("저평가 구간 (PER/PBR 낮음)")
    elif val_d["score"] <= -1:
        risks.append("고평가 구간 — 주가 조정 리스크")

    # 추세 기반 기회/위험
    for d in details:
        if d["trend"] == TREND_RECOVERY:
            opportunities.append(f"{d['category']} 회복 추세")
        elif d["trend"] == TREND_DOWN and d["score"] <= -1:
            risks.append(f"{d['category']} 하락 추세 지속")

    # 분기 모멘텀
    if momentum == "양호":
        opportunities.append("최근 분기 매출 모멘텀 양호 (QoQ 연속 성장)")
    elif momentum == "약화":
        risks.append("최근 분기 매출 모멘텀 약화 (QoQ 연속 역성장)")

    # TTM 기반 기회/위험
    if ttm_derived:
        ttm_opm = ttm_derived.get("영업이익률(%)", 0)
        if ttm_opm and opm:
            if ttm_opm > opm + 3:
                opportunities.append(f"TTM 영업이익률 개선 ({opm:.1f}→{ttm_opm:.1f}%)")
            elif ttm_opm < opm - 3:
                risks.append(f"TTM 영업이익률 악화 ({opm:.1f}→{ttm_opm:.1f}%)")

    return {
        "strengths": strengths,
        "weaknesses": weaknesses,
        "opportunities": opportunities,
        "risks": risks,
    }


def _empty_result(reason: str):
    """데이터 부족 시 빈 결과"""
    return {
        "situation": SITUATION_MATURE,
        "situation_label": "분석 불가",
        "situation_color": "#95A5A6",
        "confidence": 0,
        "summary": reason,
        "details": [],
        "key_metrics": {
            "revenue_cagr": 0,
            "opm_latest": 0,
            "roe_latest": 0,
            "debt_ratio_latest": 0,
            "fcf_trend": "-",
        },
    }
