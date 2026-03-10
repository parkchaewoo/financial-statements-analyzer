"""S-RIM 기반 저평가 종목 스크리너

다수의 종목을 일괄 분석하여 S-RIM 대비 저평가된 종목을 찾아줍니다.
"""

import time
from data_fetcher import DataFetcher
from config import DEFAULT_W_BUY, DEFAULT_W_FAIR
from generate_report import compute_derived_metrics, collect_data


def screen_stocks(
    fetcher: DataFetcher,
    stock_codes: list[str],
    required_return: float = 8.0,
    roe_source: str = "consensus",
    num_years: int = 5,
    coe_source: str = "manual",
    progress_callback=None,
    w_buy: float = DEFAULT_W_BUY,
    w_fair: float = DEFAULT_W_FAIR,
) -> list[dict]:
    """여러 종목의 S-RIM 분석을 수행하고 저평가 순으로 정렬

    Args:
        fetcher: DataFetcher 인스턴스
        stock_codes: 분석할 종목코드 리스트
        required_return: 목표 수익률(%)
        roe_source: ROE 소스 (consensus/historical)
        num_years: 분석 연도 수
        coe_source: COE 소스 (manual/capm)
        progress_callback: (current, total, msg) 콜백
        w_buy: 매수시작가 W값 (비관적)
        w_fair: 적정가 W값 (중립적)

    Returns:
        저평가 순으로 정렬된 분석 결과 리스트
    """
    results = []
    total = len(stock_codes)

    for i, code in enumerate(stock_codes):
        if progress_callback:
            progress_callback(i, total, f"{code} 분석 중...")

        try:
            result = _analyze_single(fetcher, code, required_return, roe_source,
                                     num_years, coe_source,
                                     w_buy=w_buy, w_fair=w_fair)
            if result:
                results.append(result)
        except Exception as e:
            if progress_callback:
                progress_callback(i, total, f"{code} 실패: {e}")

        # API 속도 제한 방지
        if i < total - 1:
            time.sleep(0.5)

    # 괴리율(할인율) 기준 내림차순 정렬 (할인율이 큰 것 = 더 저평가)
    results.sort(key=lambda x: x.get("discount_pct", 0), reverse=True)

    if progress_callback:
        progress_callback(total, total, "분석 완료")

    return results


def _analyze_single(
    fetcher: DataFetcher,
    stock_code: str,
    required_return: float,
    roe_source: str,
    num_years: int,
    coe_source: str = "manual",
    w_buy: float = DEFAULT_W_BUY,
    w_fair: float = DEFAULT_W_FAIR,
) -> dict | None:
    """단일 종목 S-RIM 간이 분석

    Raises:
        Exception: 데이터 조회 또는 계산 실패 시 (caller에서 에러 표시용)
    """
    # 기업 정보
    info = fetcher.fetch_company_info(stock_code)
    corp_name = info.get("corp_name", stock_code)

    # 최신 연도
    latest_year = fetcher.find_latest_available_year(stock_code)
    years = list(range(latest_year - num_years + 1, latest_year + 1))

    # 발행주식수
    shares = fetcher.fetch_shares_outstanding(stock_code, latest_year)
    if not shares:
        return None

    # 주가
    stock_data = fetcher.fetch_stock_data(stock_code, shares=shares)
    price = stock_data.get("price", 0)
    if not price:
        return None

    # 주요계정
    fs = fetcher.fetch_financial_summary(stock_code, years)

    # 컨센서스 (ROE 계산용)
    consensus = fetcher.fetch_consensus(stock_code)

    # 연도별 종가
    val_year = fetcher.fetch_valuation_by_year(stock_code, years)

    # CAPM용 beta / risk_free_rate
    beta = 1.0
    risk_free_rate = 4.0
    if coe_source == "capm":
        try:
            beta = fetcher.fetch_beta(stock_code)
        except Exception:
            beta = 1.0
        try:
            risk_free_rate = fetcher.fetch_risk_free_rate()
        except Exception:
            risk_free_rate = 4.0

    data = {
        "years": years,
        "company_info": info,
        "stock_data": stock_data,
        "financial_summary": fs,
        "balance_sheet_detail": {},
        "cash_flow_detail": {},
        "shareholders": [],
        "valuation_by_year": val_year,
        "consensus": consensus,
        "beta": beta,
        "risk_free_rate": risk_free_rate,
    }

    derived, srim = compute_derived_metrics(
        data, required_return, roe_source=roe_source, coe_source=coe_source,
        w_buy=w_buy, w_fair=w_fair,
    )

    srim_price = srim.get("srim_price", 0)
    if not srim_price or srim_price <= 0:
        return None

    # 할인율 (현재가가 적정가 대비 얼마나 싼지)
    discount_pct = round((srim_price - price) / srim_price * 100, 1)

    # 최신 연도 수익성
    latest = derived.get(latest_year, {})

    corp_name_kr = info.get("corp_name_kr", "")

    return {
        "stock_code": stock_code,
        "corp_name": corp_name,
        "corp_name_kr": corp_name_kr,
        "price": price,
        "market_cap_eok": stock_data.get("market_cap_eok", 0),
        "srim_price": srim_price,
        "buy_price": srim.get("buy_price", 0),
        "discount_pct": discount_pct,
        "roe_forecast": srim.get("roe_forecast", 0),
        "roe_source": srim.get("roe_source", ""),
        "roe_hist": srim.get("roe_hist", 0),
        "consensus_roe": srim.get("consensus_roe"),
        "coe_value": srim.get("coe_value", 0),
        "coe_source": srim.get("coe_source", ""),
        "opm": latest.get("영업이익률(%)", 0),
        "roe": latest.get("ROE(%)", 0),
        "per": latest.get("PER", 0),
        "pbr": latest.get("PBR", 0),
        "revenue_growth": latest.get("매출성장률(%)", None),
    }


def get_market_top_stocks(fetcher: DataFetcher, market: str = "KOSPI",
                          top_n: int = 50, progress_callback=None) -> list[str]:
    """시가총액 상위 N개 종목코드 반환

    DART corp_codes에서 상장사 목록을 가져온 후, 네이버 금융에서
    시가총액을 확인하여 상위 N개를 선별합니다.

    Args:
        fetcher: DataFetcher 인스턴스
        market: "KOSPI" 또는 "KOSDAQ"
        top_n: 상위 N개

    Returns:
        종목코드 리스트
    """
    import requests
    from bs4 import BeautifulSoup

    # 네이버 금융 시가총액 순위 페이지 스크래핑
    # KOSPI: sosok=0, KOSDAQ: sosok=1
    sosok = "1" if market == "KOSDAQ" else "0"

    codes = []
    # 상위 N개를 채울 때까지 페이지 순회 (1페이지당 50개)
    pages_needed = (top_n // 50) + 1

    for page in range(1, pages_needed + 1):
        url = (
            f"https://finance.naver.com/sise/sise_market_sum.naver"
            f"?sosok={sosok}&page={page}"
        )
        try:
            headers = {"User-Agent": "Mozilla/5.0"}
            resp = requests.get(url, headers=headers, timeout=10)
            resp.encoding = "euc-kr"
            soup = BeautifulSoup(resp.text, "html.parser")

            # 종목 링크에서 코드 추출
            table = soup.find("table", class_="type_2")
            if not table:
                break

            for a_tag in table.find_all("a", class_="tltle"):
                href = a_tag.get("href", "")
                if "code=" in href:
                    code = href.split("code=")[-1]
                    if len(code) == 6 and code.isdigit():
                        codes.append(code)

            if progress_callback:
                progress_callback(0, 1, f"{market} 종목 목록 조회 중... ({len(codes)}개)")

        except Exception:
            break

        if len(codes) >= top_n:
            break

    return codes[:top_n]


def get_us_market_top_stocks(market: str = "SP500", top_n: int = 50,
                              progress_callback=None) -> list[str]:
    """Wikipedia에서 S&P 500 / NASDAQ-100 구성종목을 조회하고 시총 순으로 정렬

    Args:
        market: "SP500" 또는 "NASDAQ100"
        top_n: 시총 상위 N개

    Returns:
        시총 내림차순 정렬된 티커 심볼 리스트
    """
    import requests
    from bs4 import BeautifulSoup
    import yfinance as yf

    req_headers = {"User-Agent": "Mozilla/5.0"}
    symbols = []

    try:
        if market == "SP500":
            url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
            resp = requests.get(url, headers=req_headers, timeout=15)
            soup = BeautifulSoup(resp.text, "html.parser")
            table = soup.find("table", {"id": "constituents"})
            if table:
                for row in table.find_all("tr")[1:]:
                    cols = row.find_all("td")
                    if cols:
                        symbol = cols[0].get_text(strip=True).replace(".", "-")
                        symbols.append(symbol)
        elif market == "NASDAQ100":
            url = "https://en.wikipedia.org/wiki/Nasdaq-100"
            resp = requests.get(url, headers=req_headers, timeout=15)
            soup = BeautifulSoup(resp.text, "html.parser")
            for table in soup.find_all("table", class_="wikitable"):
                header_row = table.find("tr")
                if header_row:
                    ths = [th.get_text(strip=True) for th in header_row.find_all("th")]
                    ticker_idx = -1
                    for idx, h in enumerate(ths):
                        if h in ("Ticker", "Symbol"):
                            ticker_idx = idx
                            break
                    if ticker_idx >= 0:
                        for row in table.find_all("tr")[1:]:
                            cols = row.find_all("td")
                            if cols and len(cols) > ticker_idx:
                                symbol = cols[ticker_idx].get_text(strip=True).replace(".", "-")
                                symbols.append(symbol)
                        break

        if progress_callback:
            progress_callback(0, 1, f"{market} 종목 목록 조회 완료 ({len(symbols)}개), 시총 정렬 중...")

        # yfinance로 시총 조회 후 내림차순 정렬 (병렬 처리)
        if symbols:
            from concurrent.futures import ThreadPoolExecutor, as_completed

            def _get_mcap(sym):
                try:
                    t = yf.Ticker(sym)
                    return sym, t.fast_info.get("marketCap", 0) or 0
                except Exception:
                    return sym, 0

            sym_mcap = []
            done_count = 0
            with ThreadPoolExecutor(max_workers=20) as executor:
                futures = {executor.submit(_get_mcap, sym): sym for sym in symbols}
                for future in as_completed(futures):
                    sym_mcap.append(future.result())
                    done_count += 1
                    if progress_callback and done_count % 50 == 0:
                        progress_callback(0, 1, f"{market} 시총 조회 중... ({done_count}/{len(symbols)})")

            sym_mcap.sort(key=lambda x: x[1], reverse=True)
            symbols = [s for s, _ in sym_mcap]

            if progress_callback:
                progress_callback(0, 1, f"{market} 시총 상위 {min(top_n, len(symbols))}개 선별 완료")

    except Exception:
        pass

    return symbols[:top_n]


def format_screener_results(results: list[dict], currency: str = "KRW") -> str:
    """스크리너 결과를 텍스트 테이블로 포맷팅"""
    if not results:
        return "분석 결과가 없습니다."

    def _fmt_price(v):
        if currency == "KRW":
            return f"{v:>10,}"
        return f"{'$' + f'{v:,.2f}':>12}"

    # 해외 주식 여부 감지 (한글명이 있으면 해외)
    has_kr = any(r.get('corp_name_kr') for r in results)

    lines = []
    lines.append("=" * 150)
    lines.append("  S-RIM 저평가 종목 스크리너 결과")
    lines.append("=" * 150)
    name_width = 28 if has_kr else 16
    lines.append(
        f"{'순위':>4} {'종목명':<{name_width}} {'코드':>8}  {'현재가':>12} {'매수시작가':>12} "
        f"{'S-RIM적정가':>12} "
        f"{'할인율':>7} {'ROE예측':>7} {'COE':>6} {'OPM':>6} {'PER':>7} {'PBR':>5}"
    )
    lines.append("-" * 150)

    undervalued = [r for r in results if r["discount_pct"] > 0]
    overvalued = [r for r in results if r["discount_pct"] <= 0]

    # 저평가 종목
    for rank, r in enumerate(undervalued, 1):
        coe_str = f"{r.get('coe_value', 0):>5.1f}%" if r.get('coe_value') else "    -"
        name = r['corp_name']
        kr = r.get('corp_name_kr', '')
        if kr:
            name = f"{name}({kr})"
        lines.append(
            f"{rank:>4} {name:<28} {r['stock_code']:>8}  "
            f"{_fmt_price(r['price'])} {_fmt_price(r.get('buy_price', 0))} "
            f"{_fmt_price(r['srim_price'])} "
            f"{r['discount_pct']:>+6.1f}% {r['roe_forecast']:>6.1f}% "
            f"{coe_str} {r['opm']:>5.1f}% {r['per']:>6.1f} {r['pbr']:>5.2f}"
        )

    if overvalued:
        lines.append("-" * 150)
        lines.append("  [고평가 종목]")
        for r in overvalued:
            name = r['corp_name']
            kr = r.get('corp_name_kr', '')
            if kr:
                name = f"{name}({kr})"
            coe_str = f"{r.get('coe_value', 0):>5.1f}%" if r.get('coe_value') else "    -"
            lines.append(
                f"   - {name:<{name_width}} {r['stock_code']:>8}  "
                f"{_fmt_price(r['price'])} {_fmt_price(r.get('buy_price', 0))} "
                f"{_fmt_price(r['srim_price'])} "
                f"{r['discount_pct']:>+6.1f}% {r['roe_forecast']:>6.1f}% "
                f"{coe_str}"
            )

    lines.append("=" * 150)
    lines.append(f"  총 {len(results)}개 종목 분석 | 저평가 {len(undervalued)}개 | 고평가 {len(overvalued)}개")
    lines.append(f"  할인율 = (S-RIM적정가 - 현재가) / S-RIM적정가 × 100 (높을수록 저평가)")
    lines.append(f"  매수시작가/적정가 = 초과이익 지속계수(W) 기반 2단계 가격")
    lines.append("=" * 150)

    return "\n".join(lines)
