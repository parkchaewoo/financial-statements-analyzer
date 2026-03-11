"""재무제표 분석 PDF 리포트 생성기 - 메인 실행 파일

사용법:
    python generate_report.py --stock-code 051500
    python generate_report.py --stock-code 005930 --api-key YOUR_KEY
    python generate_report.py --stock-code 051500 --return-rate 10
"""

import argparse
import sys
from datetime import datetime

from config import (DEFAULT_REQUIRED_RETURN, DEFAULT_ANALYSIS_YEARS,
                    DEFAULT_MARKET_RISK_PREMIUM,
                    DEFAULT_W_BUY, DEFAULT_W_FAIR,
                    DEFAULT_INCLUDE_QUARTERLY, DEFAULT_QUARTERLY_YEARS,
                    PDF_REPORT_COMBINED, PDF_REPORT_ANNUAL,
                    PDF_REPORT_QUARTERLY, PDF_REPORT_RISK, PDF_REPORT_ALL)
from pdf_report_base import _find
from data_fetcher import DataFetcher
from calculator import (
    calc_srim,
    calc_roe_forecast,
    calc_leverage_ratio,
    calc_effective_tax_rate,
    calc_working_capital,
    calc_working_capital_ratio,
    calc_net_debt,
    calc_pfcr,
    calc_per,
    calc_pbr,
    calc_eps,
    calc_bps,
    calc_opm,
    calc_roe,
    calc_roa,
    calc_capm_coe,
)
from pdf_report import PDFReportGenerator
from pdf_annual_report import AnnualReportGenerator
from pdf_quarterly_report import QuarterlyReportGenerator
from pdf_risk_report import RiskReportGenerator
from risk_analyzer import check_listing_risk, check_us_listing_risk
from trend_analyzer import analyze_trend


def collect_data(fetcher, stock_code: str, num_years: int,
                 include_quarterly: bool = False) -> dict:
    """API에서 모든 필요 데이터 수집"""

    # 1. 기업 기본 정보
    print("1/8 기업 정보 조회 중...")
    company_info = fetcher.fetch_company_info(stock_code)
    print(f"  → {company_info['corp_name']} ({company_info['stock_code']})")

    is_intl = company_info.get("is_international", False)
    currency = company_info.get("currency", "KRW")

    # 2. 최신 이용 가능 연도 탐색
    print("2/8 최신 사업보고서 연도 탐색 중...")
    latest_year = fetcher.find_latest_available_year(stock_code)
    years = list(range(latest_year - num_years + 1, latest_year + 1))
    print(f"  → 분석 기간: {years[0]}~{years[-1]}")

    # 3. 발행주식수 (EPS 역산)
    print("3/8 발행주식수 조회 중...")
    shares = fetcher.fetch_shares_outstanding(stock_code, latest_year)
    print(f"  → 발행주식수: {shares:,}주")

    # 4. 주가/시가총액
    print("4/8 주가 데이터 조회 중...")
    stock_data = fetcher.fetch_stock_data(stock_code, shares=shares)
    if is_intl:
        from pdf_report_base import _fmt_amount
        ex_rate = stock_data.get("exchange_rate", 0)
        price_krw = stock_data.get("price_krw", 0)
        print(f"  → 주가: ${stock_data['price']:,.2f} ({price_krw:,}원), "
              f"시총: {_fmt_amount(stock_data.get('market_cap', 0), currency)} "
              f"({stock_data.get('market_cap_eok', 0):,}억원)")
        if ex_rate:
            print(f"  → 환율: 1{currency} = {ex_rate:,.0f}원")
    else:
        print(f"  → 주가: {stock_data['price']:,}원, 시총: {stock_data['market_cap_eok']:,}억원")

    # 5. 주요계정 (5개년) + finstate_all 보강
    print("5/8 주요계정 조회 중...")
    financial_summary = fetcher.fetch_financial_summary(stock_code, years)
    print(f"  → {len(financial_summary)}개 연도 데이터 수집")

    # 6. 재무상태표 세부
    print("6/8 재무상태표 세부항목 조회 중...")
    balance_sheet = fetcher.fetch_balance_sheet_detail(stock_code, years)
    print(f"  → {len(balance_sheet)}개 연도 데이터 수집")

    # 7. 현금흐름표 세부
    print("7/8 현금흐름표 세부항목 조회 중...")
    cash_flow = fetcher.fetch_cash_flow_detail(stock_code, years)
    print(f"  → {len(cash_flow)}개 연도 데이터 수집")

    # 8. 주요 주주
    print("8/8 주요 주주 조회 중...")
    shareholders = fetcher.fetch_major_shareholders(stock_code)
    print(f"  → {len(shareholders)}명 주주 데이터")

    # 추가: 연도별 종가 (PER/PBR 계산용)
    print("  추가: 연도별 종가 조회 중...")
    valuation_by_year = fetcher.fetch_valuation_by_year(stock_code, years)

    # 추가: 컨센서스
    print("  추가: 컨센서스 조회 중...")
    consensus = fetcher.fetch_consensus(stock_code)
    if consensus.get("target_price"):
        tp = consensus['target_price']
        if is_intl:
            print(f"  → 목표주가: ${tp:,.2f}")
        else:
            print(f"  → 목표주가: {tp:,}원")
    else:
        print("  → 목표주가: 데이터 없음")
    c_roe = consensus.get("consensus_roe")
    if c_roe:
        print(f"  → 컨센서스 ROE: {c_roe:.1f}% (EPS={consensus.get('consensus_eps', 0):,.0f}, BPS={consensus.get('consensus_bps', 0):,.0f})")
    else:
        print("  → 컨센서스 ROE: 산출 불가 (EPS/BPS 데이터 없음)")

    # 추가: CAPM 데이터 (베타, 무위험수익률)
    print("  추가: CAPM 데이터 조회 중...")
    beta = fetcher.fetch_beta(stock_code)
    risk_free_rate = fetcher.fetch_risk_free_rate()
    print(f"  → Beta: {beta:.2f}, 무위험수익률: {risk_free_rate:.2f}%")

    # ── 분기 데이터 수집 + TTM 계산 (선택) ──
    quarterly_data = None
    ttm = None
    if include_quarterly:
        print("\n분기 데이터 조회 중 (현재 연도 포함)...")
        try:
            quarterly_data = fetcher.fetch_quarterly_data(
                stock_code, num_years=DEFAULT_QUARTERLY_YEARS
            )
            quarters = quarterly_data.get("quarters", [])
            print(f"  → {len(quarters)}개 분기 데이터 수집 완료")
            if len(quarters) > 0:
                print(f"  → 분기 목록: {quarters[0]} ~ {quarters[-1]}")

            # TTM 계산 (연간 보고서 이후 분기가 있을 때만)
            ttm = _compute_ttm(quarterly_data, latest_year)
            if ttm:
                used = ttm.get("quarters_used", [])
                print(f"  → TTM 계산 완료: {' + '.join(used)}")
            else:
                print("  → TTM 미생성 (연간 보고서 이후 분기 데이터 없음)")
        except Exception as e:
            print(f"  → 분기 데이터 조회 실패: {e}")
            quarterly_data = None
            ttm = None

    return {
        "years": years,
        "company_info": company_info,
        "stock_data": stock_data,
        "financial_summary": financial_summary,
        "balance_sheet_detail": balance_sheet,
        "cash_flow_detail": cash_flow,
        "shareholders": shareholders,
        "valuation_by_year": valuation_by_year,
        "consensus": consensus,
        "beta": beta,
        "risk_free_rate": risk_free_rate,
        "include_quarterly": include_quarterly,
        "quarterly_data": quarterly_data,
        "ttm": ttm,
    }


def compute_derived_metrics(data: dict, required_return: float,
                            roe_source: str = "consensus",
                            manual_roe: float = None,
                            coe_source: str = "manual",
                            w_buy: float = DEFAULT_W_BUY,
                            w_fair: float = DEFAULT_W_FAIR) -> tuple[dict, dict]:
    """파생 지표 계산 (수익성, 밸류에이션, 운전자본 등)"""
    years = data["years"]
    fs = data["financial_summary"]
    bs = data["balance_sheet_detail"]
    cf = data.get("cash_flow_detail", {})
    val_year = data["valuation_by_year"]
    stock_data = data["stock_data"]
    shares = stock_data.get("shares", 0)

    derived = {}
    roe_list = []

    for y in years:
        yd = {}
        fs_y = fs.get(y, {})
        bs_y = bs.get(y, {})
        cf_y = cf.get(y, {})

        # ── 주요 계정 값 추출 ──
        revenue = _find(fs_y, "매출액", "영업수익")
        op_profit = _find(fs_y, "영업이익")
        pretax = _find(fs_y, "법인세차감전", "법인세비용차감전")
        net_income = _find(fs_y, "당기순이익")
        net_income_ctrl = _find(fs_y, "지배기업의 소유주에게 귀속되는 당기순이익",
                                "당기순이익(지배)")
        total_assets = _find(fs_y, "자산총계")
        total_equity = _find(fs_y, "자본총계")
        equity_ctrl = _find(fs_y, "지배기업의 소유주에게 귀속되는 자본",
                            "자본총계(지배)")
        tax = _find(fs_y, "법인세수익", "법인세비용")

        # ── 수익성 지표 ──
        yd["영업이익률(%)"] = calc_opm(op_profit, revenue)
        yd["ROE(%)"] = calc_roe(net_income, total_equity) if total_equity else 0
        yd["ROA(%)"] = calc_roa(net_income, total_assets) if total_assets else 0
        yd["레버리지비율"] = calc_leverage_ratio(yd["ROE(%)"], yd["ROA(%)"])
        yd["유효법인세율(%)"] = calc_effective_tax_rate(tax, pretax)

        if yd["ROE(%)"]:
            roe_list.append(yd["ROE(%)"])

        # ── 밸류에이션 ──
        year_close = val_year.get(y, {}).get("close", 0)
        year_mktcap = year_close * shares if (year_close and shares) else 0

        yd["PER"] = calc_per(year_mktcap, net_income) if (net_income and year_mktcap) else 0
        yd["PBR"] = calc_pbr(year_mktcap, total_equity) if (total_equity and year_mktcap) else 0
        yd["EPS"] = calc_eps(net_income_ctrl or net_income, shares) if shares else 0
        yd["BPS"] = calc_bps(equity_ctrl or total_equity, shares) if shares else 0

        # ── 운전자본 ──
        receivables = _find(bs_y, "매출채권")
        inventory = _find(bs_y, "재고자산")
        payables = _find(bs_y, "매입채무")

        wc = calc_working_capital(receivables, inventory, payables)
        yd["운전자본"] = wc
        yd["운전자본비율(%)"] = calc_working_capital_ratio(wc, revenue)

        # ── CAPEX 자산 비중 ──
        tangible = _find(bs_y, "유형자산")
        intangible = _find(bs_y, "무형자산")
        total_assets_bs = _find(bs_y, "자산총계") or total_assets
        if total_assets_bs:
            yd["유형자산비중(%)"] = round(tangible / total_assets_bs * 100, 2) if tangible else 0
            yd["무형자산비중(%)"] = round(intangible / total_assets_bs * 100, 2) if intangible else 0

        # ── 현금성자산 ──
        cash = _find(bs_y, "현금및현금성자산")
        short_fin = _find(bs_y, "기타유동금융자산", "단기금융상품")
        fvpl = _find(bs_y, "당기손익-공정가치측정 금융자산", "당기손익-공정가치측정금융자산")
        cash_total = cash + short_fin + fvpl
        yd["현금성자산합계"] = cash_total

        # ── 차입금 ──
        short_borrow = _find(bs_y, "단기차입금")
        current_lt_bond = _find(bs_y, "유동성장기사채")
        current_lt_loan = _find(bs_y, "유동성장기차입금")
        bonds = _find(bs_y, "비유동사채", "사채")
        lt_borrow = _find(bs_y, "장기차입금")

        short_debt = short_borrow + current_lt_bond + current_lt_loan
        long_debt = bonds + lt_borrow
        total_debt = short_debt + long_debt

        yd["이자발생부채계산"] = total_debt
        if total_debt > 0:
            yd["단기채비중(%)"] = round(short_debt / total_debt * 100, 2)
            yd["장기채비중(%)"] = round(long_debt / total_debt * 100, 2)
        else:
            yd["단기채비중(%)"] = 0
            yd["장기채비중(%)"] = 0

        # ── 순차입금 ──
        yd["순차입금"] = calc_net_debt(total_debt, cash_total)

        # 성장률 계산용 원본 값 저장
        yd["_revenue"] = revenue
        yd["_op_profit"] = op_profit
        yd["_net_income"] = net_income

        # ── 현금흐름 ──
        op_cf = _find(cf_y, "영업활동으로 인한 현금흐름", "영업활동현금흐름")
        inv_cf = _find(cf_y, "투자활동으로 인한 현금흐름", "투자활동현금흐름")
        fin_cf = _find(cf_y, "재무활동으로 인한 현금흐름", "재무활동현금흐름")

        capex_tangible = _find(cf_y, "유형자산의 취득")
        capex_intangible = _find(cf_y, "무형자산의 취득")
        capex = -(abs(capex_tangible) + abs(capex_intangible)) if (capex_tangible or capex_intangible) else 0

        fcf = op_cf + capex if (op_cf and capex) else 0

        yd["영업활동CF"] = op_cf
        yd["투자활동CF"] = inv_cf
        yd["재무활동CF"] = fin_cf
        yd["CAPEX"] = capex
        yd["FCF"] = fcf

        # PFCR
        current_mktcap = stock_data.get("market_cap", 0)
        yd["PFCR"] = calc_pfcr(current_mktcap, fcf) if fcf else 0

        derived[y] = yd

    # ── YoY 성장률 (연속 연도 비교) ──
    for i, y in enumerate(years):
        yd = derived[y]
        yd["매출성장률(%)"] = None
        yd["영업이익성장률(%)"] = None
        yd["순이익성장률(%)"] = None
        if i == 0:
            continue
        prev = derived[years[i - 1]]
        for key, raw_key in [("매출성장률(%)", "_revenue"),
                             ("영업이익성장률(%)", "_op_profit"),
                             ("순이익성장률(%)", "_net_income")]:
            cur_val = yd.get(raw_key, 0)
            prev_val = prev.get(raw_key, 0)
            if cur_val and prev_val:
                yd[key] = round((cur_val - prev_val) / abs(prev_val) * 100, 2)

    # ── S-RIM 계산 ──
    roe_hist = calc_roe_forecast(roe_list)

    consensus = data.get("consensus", {})
    consensus_roe = consensus.get("consensus_roe")

    # ROE 소스 선택
    if roe_source == "manual" and manual_roe is not None:
        roe_forecast = manual_roe
        roe_source_label = f"직접입력({manual_roe:.1f}%)"
    elif roe_source == "historical":
        roe_forecast = roe_hist
        roe_source_label = "과거 가중평균"
    else:  # "consensus" (기본값)
        if consensus_roe and consensus_roe > 0:
            roe_forecast = consensus_roe
            roe_source_label = "컨센서스"
        else:
            roe_forecast = roe_hist
            roe_source_label = "과거 가중평균(컨센서스 대체)"

    # COE 소스 선택
    beta = data.get("beta", 1.0)
    risk_free_rate = data.get("risk_free_rate", 4.0)
    capm_coe = calc_capm_coe(risk_free_rate, beta, DEFAULT_MARKET_RISK_PREMIUM)

    if coe_source == "capm":
        actual_return = capm_coe
        coe_source_label = f"CAPM({capm_coe:.1f}%)"
    else:  # "manual" (기본값)
        actual_return = required_return
        coe_source_label = f"수동입력({required_return:.1f}%)"

    # 데이터가 있는 최신 연도 사용
    latest_year = years[-1]
    for y in reversed(years):
        if fs.get(y):
            latest_year = y
            break

    equity_ctrl = _find(fs.get(latest_year, {}),
                        "지배기업의 소유주에게 귀속되는 자본",
                        "자본총계(지배)")
    if not equity_ctrl:
        equity_ctrl = _find(fs.get(latest_year, {}), "자본총계")

    srim = calc_srim(equity_ctrl, roe_forecast, actual_return, shares,
                     w_buy=w_buy, w_fair=w_fair)
    srim["roe_forecast"] = roe_forecast
    srim["roe_hist"] = roe_hist
    srim["roe_source"] = roe_source_label
    srim["consensus_roe"] = consensus_roe
    srim["coe_source"] = coe_source_label
    srim["coe_value"] = actual_return
    srim["capm_coe"] = capm_coe
    srim["beta"] = beta
    srim["risk_free_rate"] = risk_free_rate
    srim["market_risk_premium"] = DEFAULT_MARKET_RISK_PREMIUM

    # ── TTM 파생지표 ──
    ttm = data.get("ttm")
    if ttm:
        ttm_derived = _compute_ttm_derived(ttm, stock_data)
        derived["_ttm"] = ttm_derived

    # ── 분기별 파생지표 ──
    quarterly_data = data.get("quarterly_data")
    if quarterly_data:
        qd, qkeys = _compute_quarterly_derived(quarterly_data)
        derived["_quarterly"] = qd
        derived["_quarterly_keys"] = qkeys

    return derived, srim


def _compute_ttm(quarterly_data: dict, latest_annual_year: int = 0) -> dict | None:
    """연간 보고서 이후 분기만으로 TTM 계산

    연간 보고서에 이미 포함된 분기는 제외하고,
    그 이후에 나온 분기 보고서만 합산.

    예: latest_annual_year=2024, 2025 Q1~Q3 있음 → TTM = Q1+Q2+Q3
        latest_annual_year=2025, 2026 Q1 있음 → TTM = Q1

    IS/CF: 해당 분기 합산
    BS: 최신 분기 값 (시점 데이터)
    """
    if not quarterly_data or not latest_annual_year:
        return None

    quarters = quarterly_data.get("quarters", [])

    # 연간 보고서 이후 분기만 필터
    post_annual_q = [q for q in quarters if int(q[:4]) > latest_annual_year]
    if not post_annual_q:
        return None

    latest_q = post_annual_q[-1]

    # 분기→연간 필드명 매핑 (DART 분기 보고서는 다른 이름 사용)
    _QUARTERLY_TO_ANNUAL = {
        "분기순이익": "당기순이익",
        "반기순이익": "당기순이익",
        "분기순이익(손실)": "당기순이익(손실)",
        "반기순이익(손실)": "당기순이익(손실)",
        "분기총포괄이익": "총포괄이익",
        "반기총포괄이익": "총포괄이익",
        "기본주당분기순이익": "기본주당이익",
        "기본주당반기순이익": "기본주당이익",
        "희석주당분기순이익": "희석주당이익",
        "희석주당반기순이익": "희석주당이익",
    }

    def _normalize_key(key: str) -> str:
        """분기 필드명을 연간 필드명으로 변환"""
        # CIS_ 접두사 처리
        prefix = ""
        bare = key
        if key.startswith("CIS_") or key.startswith("IS_"):
            prefix = key[:key.index("_") + 1]
            bare = key[key.index("_") + 1:]
        normalized = _QUARTERLY_TO_ANNUAL.get(bare, bare)
        return prefix + normalized

    # IS: 해당 분기 합산
    raw_summary = {}
    qs = quarterly_data.get("quarterly_summary", {})
    all_is_keys = set()
    for q in post_annual_q:
        all_is_keys.update(qs.get(q, {}).keys())

    for key in all_is_keys:
        total = 0
        has_data = False
        for q in post_annual_q:
            val = qs.get(q, {}).get(key, 0)
            if val:
                total += val
                has_data = True
        if has_data:
            raw_summary[key] = total

    # 연환산 배수: 4분기 미만이면 연환산 (IS/CF는 누적값이므로)
    n_quarters = len(post_annual_q)
    ann_factor = 4 / n_quarters if n_quarters < 4 else 1

    # 원본 키 유지 + 정규화된 키 추가 (연간 필드명으로 검색 가능하도록)
    # IS 항목은 연환산 적용 (연간 데이터와 비교 가능하도록)
    ttm_summary = {}
    for key, val in raw_summary.items():
        ttm_summary[key] = int(val * ann_factor) if val else 0
    for key, val in raw_summary.items():
        norm_key = _normalize_key(key)
        if norm_key != key and norm_key not in ttm_summary and val:
            ttm_summary[norm_key] = int(val * ann_factor)

    # BS: 최신 분기 값 (시점 데이터 → 연환산 안 함)
    ttm_bs = {}
    qbs = quarterly_data.get("quarterly_bs", {})
    if latest_q in qbs:
        ttm_bs = dict(qbs[latest_q])

    # CF: 해당 분기 합산 후 연환산
    ttm_cf = {}
    qcf = quarterly_data.get("quarterly_cf", {})
    all_cf_keys = set()
    for q in post_annual_q:
        all_cf_keys.update(qcf.get(q, {}).keys())

    for key in all_cf_keys:
        total = 0
        has_data = False
        for q in post_annual_q:
            val = qcf.get(q, {}).get(key, 0)
            if val:
                total += val
                has_data = True
        if has_data:
            ttm_cf[key] = int(total * ann_factor)

    # TTM 라벨: "2025 TTM" + 연환산 표시
    ttm_year = int(post_annual_q[-1][:4])
    if n_quarters < 4:
        ttm_label = f"{ttm_year} TTM*"
    else:
        ttm_label = f"{ttm_year} TTM"

    return {
        "financial_summary": ttm_summary,
        "balance_sheet": ttm_bs,
        "cash_flow": ttm_cf,
        "quarters_used": post_annual_q,
        "ttm_label": ttm_label,
        "annualized": n_quarters < 4,
        "ann_factor": ann_factor,
    }


def _compute_ttm_derived(ttm: dict, stock_data: dict) -> dict:
    """TTM 데이터에서 파생 지표 계산

    분기 수가 4 미만일 경우 IS/CF 금액을 연환산(annualize)하여
    연간 데이터와 비교 가능한 비율을 산출한다.
    (예: 3분기 합산 → ×4/3, 1분기만 → ×4/1)
    BS(재무상태표)는 시점 데이터이므로 연환산하지 않는다.
    """
    if not ttm:
        return {}

    fs = ttm.get("financial_summary", {})
    bs_data = ttm.get("balance_sheet", {})
    cf_data = ttm.get("cash_flow", {})
    shares = stock_data.get("shares", 0)

    # 연환산 배수: 4 / 사용 분기 수
    n_quarters = len(ttm.get("quarters_used", []))
    ann_factor = 4 / n_quarters if n_quarters else 1

    td = {}

    # 주요 계정 추출 (분기 보고서: 당기→분기/반기 필드명 대응)
    # IS 항목은 합산값이므로 연환산 적용
    revenue_raw = _find(fs, "매출액", "영업수익")
    op_profit_raw = _find(fs, "영업이익")
    net_income_raw = _find(fs, "당기순이익") or _find(fs, "분기순이익") or _find(fs, "반기순이익")
    net_income_ctrl_raw = (_find(fs, "지배기업의 소유주에게 귀속되는 당기순이익",
                                 "당기순이익(지배)")
                           or _find(fs, "지배기업의 소유주지분"))

    revenue = int(revenue_raw * ann_factor) if revenue_raw else 0
    op_profit = int(op_profit_raw * ann_factor) if op_profit_raw else 0
    net_income = int(net_income_raw * ann_factor) if net_income_raw else 0
    net_income_ctrl = int(net_income_ctrl_raw * ann_factor) if net_income_ctrl_raw else 0

    # BS 항목은 시점 데이터 → 연환산 안 함
    total_assets = _find(bs_data, "자산총계")
    total_equity = _find(bs_data, "자본총계")
    equity_ctrl = _find(bs_data, "지배기업의 소유주에게 귀속되는 자본",
                        "자본총계(지배)")

    # 수익성 (연환산된 IS / 시점 BS)
    td["영업이익률(%)"] = calc_opm(op_profit, revenue)
    td["ROE(%)"] = calc_roe(net_income, total_equity) if total_equity else 0
    td["ROA(%)"] = calc_roa(net_income, total_assets) if total_assets else 0
    td["레버리지비율"] = calc_leverage_ratio(td["ROE(%)"], td["ROA(%)"])

    # 밸류에이션 (현재 시가총액 vs 연환산 이익)
    current_mktcap = stock_data.get("market_cap", 0)
    td["PER"] = calc_per(current_mktcap, net_income) if (net_income and current_mktcap) else 0
    td["PBR"] = calc_pbr(current_mktcap, total_equity) if (total_equity and current_mktcap) else 0
    td["EPS"] = calc_eps(net_income_ctrl or net_income, shares) if shares else 0
    td["BPS"] = calc_bps(equity_ctrl or total_equity, shares) if shares else 0

    # 운전자본 (BS 시점값, 비율 계산 시 연환산 매출 사용)
    receivables = _find(bs_data, "매출채권")
    inventory = _find(bs_data, "재고자산")
    payables = _find(bs_data, "매입채무")
    wc = calc_working_capital(receivables, inventory, payables)
    td["운전자본"] = wc
    td["운전자본비율(%)"] = calc_working_capital_ratio(wc, revenue)

    # CAPEX/자산 비중
    tangible = _find(bs_data, "유형자산")
    intangible = _find(bs_data, "무형자산")
    total_assets_bs = total_assets
    if total_assets_bs:
        td["유형자산비중(%)"] = round(tangible / total_assets_bs * 100, 2) if tangible else 0
        td["무형자산비중(%)"] = round(intangible / total_assets_bs * 100, 2) if intangible else 0

    # 현금성자산
    cash = _find(bs_data, "현금및현금성자산")
    short_fin = _find(bs_data, "기타유동금융자산", "단기금융상품")
    fvpl = _find(bs_data, "당기손익-공정가치측정 금융자산", "당기손익-공정가치측정금융자산")
    cash_total = cash + short_fin + fvpl
    td["현금성자산합계"] = cash_total

    # 차입금
    short_borrow = _find(bs_data, "단기차입금")
    current_lt_bond = _find(bs_data, "유동성장기사채")
    current_lt_loan = _find(bs_data, "유동성장기차입금")
    bonds = _find(bs_data, "비유동사채", "사채")
    lt_borrow = _find(bs_data, "장기차입금")

    short_debt = short_borrow + current_lt_bond + current_lt_loan
    long_debt = bonds + lt_borrow
    total_debt = short_debt + long_debt

    td["이자발생부채계산"] = total_debt
    if total_debt > 0:
        td["단기채비중(%)"] = round(short_debt / total_debt * 100, 2)
        td["장기채비중(%)"] = round(long_debt / total_debt * 100, 2)
    else:
        td["단기채비중(%)"] = 0
        td["장기채비중(%)"] = 0

    # 순차입금
    td["순차입금"] = calc_net_debt(total_debt, cash_total)

    # 현금흐름 (CF도 합산값이므로 연환산)
    op_cf_raw = _find(cf_data, "영업활동으로 인한 현금흐름", "영업활동현금흐름")
    inv_cf_raw = _find(cf_data, "투자활동으로 인한 현금흐름", "투자활동현금흐름")
    fin_cf_raw = _find(cf_data, "재무활동으로 인한 현금흐름", "재무활동현금흐름")

    capex_tangible = _find(cf_data, "유형자산의 취득")
    capex_intangible = _find(cf_data, "무형자산의 취득")
    capex_raw = -(abs(capex_tangible) + abs(capex_intangible)) if (capex_tangible or capex_intangible) else 0

    op_cf = int(op_cf_raw * ann_factor) if op_cf_raw else 0
    inv_cf = int(inv_cf_raw * ann_factor) if inv_cf_raw else 0
    fin_cf = int(fin_cf_raw * ann_factor) if fin_cf_raw else 0
    capex = int(capex_raw * ann_factor) if capex_raw else 0

    fcf = op_cf + capex if (op_cf and capex) else 0

    td["영업활동CF"] = op_cf
    td["투자활동CF"] = inv_cf
    td["재무활동CF"] = fin_cf
    td["CAPEX"] = capex
    td["FCF"] = fcf
    td["PFCR"] = calc_pfcr(current_mktcap, fcf) if fcf else 0

    # 원본 값 저장
    td["_revenue"] = revenue
    td["_op_profit"] = op_profit
    td["_net_income"] = net_income

    return td


def _compute_quarterly_derived(quarterly_data: dict) -> tuple[dict, list]:
    """분기별 파생지표 계산 + QoQ/YoY 성장률

    Returns:
        (quarterly_derived, quarterly_keys)
        quarterly_derived: {"2024Q1": {"매출액": ..., "영업이익률(%)": ..., ...}, ...}
        quarterly_keys: ["2023Q1", "2023Q2", ..., "2024Q4"] (시간순)
    """
    if not quarterly_data:
        return {}, []

    quarters = quarterly_data.get("quarters", [])
    qs = quarterly_data.get("quarterly_summary", {})
    qbs = quarterly_data.get("quarterly_bs", {})
    qcf = quarterly_data.get("quarterly_cf", {})

    qd = {}
    for q in quarters:
        fs_q = qs.get(q, {})
        bs_q = qbs.get(q, {})
        cf_q = qcf.get(q, {})

        d = {}
        revenue = _find(fs_q, "매출액", "영업수익")
        op_profit = _find(fs_q, "영업이익")
        net_income = _find(fs_q, "당기순이익") or _find(fs_q, "분기순이익") or _find(fs_q, "반기순이익")
        total_equity = _find(bs_q, "자본총계")
        total_assets = _find(bs_q, "자산총계")

        d["매출액"] = revenue
        d["영업이익"] = op_profit
        d["당기순이익"] = net_income
        d["영업이익률(%)"] = calc_opm(op_profit, revenue)
        d["ROE(%)"] = calc_roe(net_income, total_equity) if total_equity else 0
        d["ROA(%)"] = calc_roa(net_income, total_assets) if total_assets else 0

        # 현금흐름
        op_cf = _find(cf_q, "영업활동으로 인한 현금흐름", "영업활동현금흐름")
        capex_tangible = _find(cf_q, "유형자산의 취득")
        capex_intangible = _find(cf_q, "무형자산의 취득")
        capex = -(abs(capex_tangible) + abs(capex_intangible)) if (capex_tangible or capex_intangible) else 0
        fcf = op_cf + capex if (op_cf and capex) else 0

        d["영업활동CF"] = op_cf
        d["FCF"] = fcf

        # BS 항목
        d["자산총계"] = total_assets
        d["부채총계"] = _find(bs_q, "부채총계")
        d["자본총계"] = total_equity

        d["_revenue"] = revenue
        d["_op_profit"] = op_profit
        d["_net_income"] = net_income

        qd[q] = d

    # QoQ / YoY 성장률
    for i, q in enumerate(quarters):
        d = qd[q]
        d["매출성장률_QoQ(%)"] = None
        d["영업이익성장률_QoQ(%)"] = None
        d["순이익성장률_QoQ(%)"] = None
        d["매출성장률_YoY(%)"] = None
        d["영업이익성장률_YoY(%)"] = None
        d["순이익성장률_YoY(%)"] = None

        # QoQ: 직전 분기
        if i > 0:
            prev_q = quarters[i - 1]
            prev = qd[prev_q]
            for key, raw_key in [("매출성장률_QoQ(%)", "_revenue"),
                                 ("영업이익성장률_QoQ(%)", "_op_profit"),
                                 ("순이익성장률_QoQ(%)", "_net_income")]:
                cur_val = d.get(raw_key, 0)
                prev_val = prev.get(raw_key, 0)
                if cur_val and prev_val:
                    d[key] = round((cur_val - prev_val) / abs(prev_val) * 100, 2)

        # YoY: 전년 동분기
        # 분기 키 형식: "2024Q1" → 전년 "2023Q1"
        try:
            q_year = int(q[:4])
            q_num = q[4:]  # "Q1", "Q2", etc.
            yoy_q = f"{q_year - 1}{q_num}"
            if yoy_q in qd:
                prev = qd[yoy_q]
                for key, raw_key in [("매출성장률_YoY(%)", "_revenue"),
                                     ("영업이익성장률_YoY(%)", "_op_profit"),
                                     ("순이익성장률_YoY(%)", "_net_income")]:
                    cur_val = d.get(raw_key, 0)
                    prev_val = prev.get(raw_key, 0)
                    if cur_val and prev_val:
                        d[key] = round((cur_val - prev_val) / abs(prev_val) * 100, 2)
        except (ValueError, IndexError):
            pass

    return qd, quarters


def main():
    parser = argparse.ArgumentParser(
        description="재무제표 분석 PDF 리포트 생성기",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "예시:\n"
            "  python generate_report.py --stock 삼성전자\n"
            "  python generate_report.py --stock-code 005930\n"
            "  python generate_report.py --stock CJ프레시웨이 --return-rate 10\n"
            "  python generate_report.py --stock AAPL --market INTL\n"
            "  python generate_report.py --stock MSFT --market INTL --return-rate 10\n"
            "\nDART API 키 발급: https://opendart.fss.or.kr"
        ),
    )
    parser.add_argument(
        "--stock-code", default=None,
        help="종목코드 (6자리, 예: 005930)"
    )
    parser.add_argument(
        "--stock", default=None,
        help="종목명 또는 종목코드 (예: 삼성전자, 005930)"
    )
    parser.add_argument(
        "--api-key", default=None, help="DART OpenAPI 키 (미지정 시 .env에서 로드)"
    )
    parser.add_argument(
        "--return-rate",
        type=float,
        default=DEFAULT_REQUIRED_RETURN,
        help=f"S-RIM 원하는 수익률 %% (기본: {DEFAULT_REQUIRED_RETURN}%%)",
    )
    parser.add_argument(
        "--years",
        type=int,
        default=DEFAULT_ANALYSIS_YEARS,
        help=f"분석 연도 수 (기본: {DEFAULT_ANALYSIS_YEARS})",
    )
    parser.add_argument(
        "--roe-source", choices=["consensus", "historical", "manual"],
        default="consensus",
        help="S-RIM ROE 소스: consensus(컨센서스 우선), historical(과거 가중평균), manual(직접입력)",
    )
    parser.add_argument(
        "--roe-value", type=float, default=None,
        help="--roe-source manual 시 사용할 ROE(%%) 값",
    )
    parser.add_argument(
        "--coe-source", choices=["manual", "capm"], default="manual",
        help="COE(자기자본비용) 소스: manual(수동입력, 기본), capm(CAPM 자동계산)",
    )
    parser.add_argument(
        "--output", default=None, help="출력 파일 경로 (기본: {기업명}_분석리포트.pdf)"
    )
    parser.add_argument(
        "--market", choices=["KR", "INTL"], default="KR",
        help="시장: KR(한국, 기본), INTL(해외/yfinance)",
    )
    parser.add_argument(
        "--w-buy", type=float, default=DEFAULT_W_BUY,
        help=f"S-RIM 매수시작가 지속계수 W (비관적, 기본: {DEFAULT_W_BUY})",
    )
    parser.add_argument(
        "--w-fair", type=float, default=DEFAULT_W_FAIR,
        help=f"S-RIM 적정가 지속계수 W (중립적, 기본: {DEFAULT_W_FAIR})",
    )
    parser.add_argument(
        "--quarterly", action="store_true", default=DEFAULT_INCLUDE_QUARTERLY,
        help="TTM 열 및 분기별 상세 페이지를 리포트에 포함 (기본: 포함)",
    )
    parser.add_argument(
        "--trend", action="store_true", default=True,
        help="재무 추이 분석 페이지를 리포트에 포함 (기본: 포함)",
    )
    parser.add_argument(
        "--no-trend", action="store_true", default=False,
        help="재무 추이 분석 페이지를 리포트에서 제외",
    )
    parser.add_argument(
        "--report-type",
        choices=[PDF_REPORT_COMBINED, PDF_REPORT_ANNUAL,
                 PDF_REPORT_QUARTERLY, PDF_REPORT_RISK, PDF_REPORT_ALL],
        default=PDF_REPORT_COMBINED,
        help="리포트 유형: combined(통합), annual(연도별), quarterly(분기별), risk(리스크), all(3개 분리)",
    )

    args = parser.parse_args()
    # --no-trend가 지정되면 trend 비활성화
    if args.no_trend:
        args.trend = False

    # 초기화
    is_intl = args.market == "INTL"
    print("=" * 50)
    print("  재무제표 분석 PDF 리포트 생성기")
    if is_intl:
        print("  (해외 주식 모드 - yfinance)")
    print("=" * 50)

    if is_intl:
        from international_fetcher import InternationalFetcher
        fetcher = InternationalFetcher()
    else:
        try:
            fetcher = DataFetcher(api_key=args.api_key)
        except ValueError as e:
            print(f"\n오류: {e}")
            sys.exit(1)

    # 종목코드 결정
    stock_query = args.stock or args.stock_code
    if not stock_query:
        print("\n오류: --stock 또는 --stock-code를 입력하세요.")
        sys.exit(1)

    try:
        stock_code = fetcher.resolve_stock_query(stock_query)
    except ValueError as e:
        print(f"\n{e}")
        sys.exit(1)

    # 데이터 수집
    print(f"\n종목코드: {stock_code}")
    try:
        data = collect_data(fetcher, stock_code, args.years,
                           include_quarterly=args.quarterly)
    except ValueError as e:
        print(f"\n오류: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n데이터 수집 중 오류: {e}")
        sys.exit(1)

    # 파생 지표 계산
    print("\n파생 지표 계산 중...")
    derived, srim = compute_derived_metrics(
        data, args.return_rate,
        roe_source=args.roe_source,
        manual_roe=args.roe_value,
        coe_source=args.coe_source,
        w_buy=args.w_buy,
        w_fair=args.w_fair,
    )

    # 위험 분석
    print("위험 분석 중...")
    ttm_data = data.get("ttm")
    ttm_derived = derived.get("_ttm")
    if is_intl:
        risk_warnings = check_us_listing_risk(data, derived)
    else:
        risk_warnings = check_listing_risk(
            data, derived,
            ttm_data=ttm_data, ttm_derived=ttm_derived
        )
    if risk_warnings:
        print(f"  → {len(risk_warnings)}건의 위험 항목 발견")
    else:
        print("  → 위험 항목 없음")

    # 추이 분석
    trend_result = None
    if args.trend:
        print("재무 추이 분석 중...")
        try:
            trend_result = analyze_trend(
                data, derived,
                ttm_data=ttm_data, ttm_derived=ttm_derived,
                quarterly_derived=derived.get("_quarterly"),
                quarterly_keys=derived.get("_quarterly_keys"),
            )
            label = trend_result.get("situation_label", "?")
            conf = trend_result.get("confidence", 0)
            print(f"  → 종합 진단: {label} (신뢰도 {conf * 100:.0f}%)")
        except Exception as e:
            print(f"  → 추이 분석 실패: {e}")
            trend_result = None

    # 리포트 데이터 조합
    report_data = {
        **data,
        "derived": derived,
        "srim": srim,
        "required_return": args.return_rate,
        "risk_warnings": risk_warnings,
        "trend_analysis": trend_result,
    }

    # PDF 생성
    corp_name = data["company_info"]["corp_name"]
    report_type = args.report_type

    print(f"\nPDF 리포트 생성 중...")

    generated_files = []

    if report_type == PDF_REPORT_COMBINED:
        output_path = args.output or f"{corp_name}_분석리포트.pdf"
        PDFReportGenerator(output_path).generate(report_data)
        generated_files.append(output_path)

    elif report_type == PDF_REPORT_ANNUAL:
        output_path = args.output or f"{corp_name}_연도별분석.pdf"
        AnnualReportGenerator(output_path).generate(report_data)
        generated_files.append(output_path)

    elif report_type == PDF_REPORT_QUARTERLY:
        output_path = args.output or f"{corp_name}_분기별분석.pdf"
        QuarterlyReportGenerator(output_path).generate(report_data)
        generated_files.append(output_path)

    elif report_type == PDF_REPORT_RISK:
        output_path = args.output or f"{corp_name}_리스크분석.pdf"
        RiskReportGenerator(output_path).generate(report_data)
        generated_files.append(output_path)

    elif report_type == PDF_REPORT_ALL:
        base = args.output.rsplit(".", 1)[0] if args.output else corp_name
        for suffix, gen_cls in [
            ("_연도별분석.pdf", AnnualReportGenerator),
            ("_분기별분석.pdf", QuarterlyReportGenerator),
            ("_리스크분석.pdf", RiskReportGenerator),
        ]:
            path = f"{base}{suffix}"
            gen_cls(path).generate(report_data)
            generated_files.append(path)

    print(f"\n{'=' * 50}")
    for f in generated_files:
        print(f"  완료! 파일: {f}")
    print(f"{'=' * 50}")


if __name__ == "__main__":
    main()
