"""재무 지표 계산 모듈 - 엑셀 Summary 시트의 파생 지표 재현"""


def calc_capm_coe(risk_free_rate: float, beta: float,
                  market_risk_premium: float = 5.5) -> float:
    """CAPM 모델로 자기자본비용(COE) 계산

    COE = Rf + β × (Rm - Rf)

    Args:
        risk_free_rate: 무위험수익률 (%, 예: 4.2)
        beta: 베타 계수 (시장 대비 변동성)
        market_risk_premium: 시장위험프리미엄 (%, 예: 5.5)

    Returns:
        COE (%, 예: 9.7)
    """
    if beta is None or beta == 0:
        beta = 1.0
    return round(risk_free_rate + beta * market_risk_premium, 2)


def calc_srim(equity: int, roe: float, required_return: float, shares: int,
              w_buy: float = 0.5, w_fair: float = 1.0) -> dict:
    """S-RIM (초과이익 모델) 적정주가 계산 — 초과이익 지속계수(W) 기반 2단계

    공식: V = BPS + W × BPS × (ROE − COE) / (1 + COE − W)
    W=1.0 일 때: V = BPS + BPS × (ROE − COE) / COE (무한등비급수)

    Args:
        equity: 자본총계(지배) (원)
        roe: ROE (%, 예: 9.6)
        required_return: 원하는 수익률 / COE (%, 예: 8.0)
        shares: 발행주식수
        w_buy: 매수시작가 W값 (비관적, default=0.5)
        w_fair: 적정가 W값 (초과이익 영구지속, default=1.0)

    Returns:
        dict: srim_price(적정주가), buy_price(매수시작가), w_buy, w_fair
    """
    if shares == 0 or required_return == 0:
        return {"srim_price": 0, "buy_price": 0,
                "w_buy": w_buy, "w_fair": w_fair}

    roe_decimal = roe / 100
    rr_decimal = required_return / 100

    def _calc_w(w: float) -> float:
        """W 기반 S-RIM 기업가치 계산"""
        if w >= 1.0:
            # W=1: 초과이익 영구 지속 (무한등비급수)
            if rr_decimal == 0:
                return float(equity)
            return equity + equity * (roe_decimal - rr_decimal) / rr_decimal
        denom = 1 + rr_decimal - w
        if denom <= 0:
            return float(equity)
        return equity + w * equity * (roe_decimal - rr_decimal) / denom

    srim_price = int(_calc_w(w_fair) / shares)   # 적정가 (W=1.0)
    buy_price = int(_calc_w(w_buy) / shares)     # 매수시작가 (W=0.5)

    return {
        "srim_price": srim_price,
        "buy_price": buy_price,
        "w_buy": w_buy,
        "w_fair": w_fair,
    }


def calc_roe_forecast(roe_list: list[float]) -> float:
    """ROE 가중평균 예측 (최근 연도에 높은 가중치)

    엑셀 공식: (E23*1 + F23*2 + G23*3) / 6
    최근 3개년 ROE를 1:2:3 가중평균
    """
    if not roe_list:
        return 0.0

    # 최근 3개년만 사용
    recent = roe_list[-3:] if len(roe_list) >= 3 else roe_list

    if len(recent) == 3:
        return (recent[0] * 1 + recent[1] * 2 + recent[2] * 3) / 6
    elif len(recent) == 2:
        return (recent[0] * 1 + recent[1] * 2) / 3
    else:
        return recent[0]


def calc_leverage_ratio(roe: float, roa: float) -> float:
    """레버리지 비율 = ROE / ROA"""
    if roa == 0:
        return 0.0
    return round(roe / roa, 2)


def calc_effective_tax_rate(tax: int, pretax_income: int) -> float:
    """유효법인세율(%) = -법인세 / 세전이익 * 100"""
    if pretax_income == 0:
        return 0.0
    return round(abs(tax) / abs(pretax_income) * 100, 2)


def calc_working_capital(receivables: int, inventory: int, payables: int) -> int:
    """운전자본 = 매출채권 + 재고자산 - 매입채무"""
    return receivables + inventory - payables


def calc_working_capital_ratio(working_capital: int, revenue: int) -> float:
    """매출대비 운전자본비율(%) = 운전자본 / 매출액 * 100"""
    if revenue == 0:
        return 0.0
    return round(working_capital / revenue * 100, 2)


def calc_net_debt(interest_bearing_debt: int, cash_assets: int) -> int:
    """순차입금 = 이자발생부채 - 현금성자산"""
    return interest_bearing_debt - cash_assets


def calc_pfcr(market_cap: int, fcf: int) -> float:
    """PFCR (Price to Free Cash Flow Ratio) = 시가총액 / FCF"""
    if fcf == 0:
        return 0.0
    return round(market_cap / fcf, 2)


def calc_per(market_cap: int, net_income: int) -> float:
    """PER = 시가총액 / 당기순이익"""
    if net_income == 0:
        return 0.0
    return round(market_cap / net_income, 2)


def calc_pbr(market_cap: int, equity: int) -> float:
    """PBR = 시가총액 / 자본총계"""
    if equity == 0:
        return 0.0
    return round(market_cap / equity, 2)


def calc_eps(net_income: int, shares: int) -> int:
    """EPS = 당기순이익(지배) / 발행주식수"""
    if shares == 0:
        return 0
    return int(net_income / shares)


def calc_bps(equity: int, shares: int) -> int:
    """BPS = 자본총계(지배) / 발행주식수"""
    if shares == 0:
        return 0
    return int(equity / shares)


def calc_dividend_yield(dps: int, price: int) -> float:
    """예상 배당수익률(%) = DPS / 주가 * 100"""
    if price == 0:
        return 0.0
    return round(dps / price * 100, 2)


def calc_opm(operating_profit: int, revenue: int) -> float:
    """영업이익률(%) = 영업이익 / 매출액 * 100"""
    if revenue == 0:
        return 0.0
    return round(operating_profit / revenue * 100, 2)


def calc_roe(net_income: int, equity: int) -> float:
    """ROE(%) = 당기순이익 / 자본총계 * 100"""
    if equity == 0:
        return 0.0
    return round(net_income / equity * 100, 2)


def calc_roa(net_income: int, total_assets: int) -> float:
    """ROA(%) = 당기순이익 / 자산총계 * 100"""
    if total_assets == 0:
        return 0.0
    return round(net_income / total_assets * 100, 2)


def calc_debt_ratio_short_long(
    short_term_debt: int, long_term_debt: int, total_debt: int
) -> dict:
    """단기채/장기채 비중 계산"""
    if total_debt == 0:
        return {"short_ratio": 0.0, "long_ratio": 0.0}
    return {
        "short_ratio": round(short_term_debt / total_debt * 100, 2),
        "long_ratio": round(long_term_debt / total_debt * 100, 2),
    }


def to_eok(value: int) -> float:
    """원 단위를 억원 단위로 변환"""
    return round(value / 100_000_000, 1)
