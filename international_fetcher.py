"""해외 주식 데이터 수집 모듈 - yfinance 기반

DataFetcher와 동일한 인터페이스로 해외 주식 데이터를 제공합니다.
yfinance 데이터를 한국어 키 딕셔너리로 변환하여 compute_derived_metrics()와 호환됩니다.
"""

import time
from datetime import datetime, timedelta

import pandas as pd

# ── yfinance → 한국어 키 매핑 ──────────────────────────────────

IS_FIELD_MAP = {
    "Total Revenue": "매출액",
    "Operating Revenue": "영업수익",
    "Operating Income": "영업이익",
    "Pretax Income": "법인세비용차감전계속사업이익",
    "Tax Provision": "법인세비용",
    "Net Income": "당기순이익",
    "Net Income Common Stockholders": "당기순이익(지배)",
    "Basic EPS": "기본주당이익",
    "Diluted EPS": "희석주당이익",
}

BS_FIELD_MAP = {
    "Total Assets": "자산총계",
    "Stockholders Equity": "자본총계",
    "Common Stock Equity": "자본총계(지배)",
    "Total Liabilities Net Minority Interest": "부채총계",
    "Total Debt": "이자발생부채",
    "Current Debt": "단기차입금",
    "Current Debt And Capital Lease Obligation": "유동성장기부채",
    "Long Term Debt": "장기차입금",
    "Capital Stock": "자본금",
    "Accounts Receivable": "매출채권",
    "Receivables": "매출채권및기타유동채권",
    "Inventory": "재고자산",
    "Accounts Payable": "매입채무",
    "Payables": "매입채무및기타유동채무",
    "Net PPE": "유형자산",
    "Goodwill And Other Intangible Assets": "무형자산",
    "Cash And Cash Equivalents": "현금및현금성자산",
    "Other Short Term Investments": "단기금융상품",
    "Cash Cash Equivalents And Short Term Investments": "현금성자산합계_raw",
}

CF_FIELD_MAP = {
    "Operating Cash Flow": "영업활동으로 인한 현금흐름",
    "Investing Cash Flow": "투자활동으로 인한 현금흐름",
    "Financing Cash Flow": "재무활동으로 인한 현금흐름",
    "Capital Expenditure": "유형자산의 취득",
    "Purchase Of PPE": "유형자산의 취득2",
    "Free Cash Flow": "FCF_raw",
    "Depreciation And Amortization": "감가상각비",
}


def _find_year_column(df: pd.DataFrame, year: int):
    """DataFrame에서 해당 연도의 컬럼을 찾습니다."""
    if df is None or df.empty:
        return None
    for col in df.columns:
        if hasattr(col, "year") and col.year == year:
            return col
    return None


def _safe_int(val) -> int:
    """안전하게 int 변환"""
    if val is None or pd.isna(val):
        return 0
    return int(val)


def _has_korean(text: str) -> bool:
    """한국어 문자 포함 여부"""
    return any('\uAC00' <= c <= '\uD7A3' or '\u3131' <= c <= '\u318E' for c in text)


def _translate_ko_to_en(query: str) -> str | None:
    """한국어 → 영어 번역 (Google Translate 무료 엔드포인트)"""
    import requests
    try:
        url = "https://translate.googleapis.com/translate_a/single"
        params = {"client": "gtx", "sl": "ko", "tl": "en", "dt": "t", "q": query}
        resp = requests.get(url, params=params, timeout=5)
        if resp.status_code == 200:
            data = resp.json()
            return data[0][0][0]
    except Exception:
        pass
    return None


def _translate_en_to_ko(query: str) -> str | None:
    """영어 → 한국어 번역 (Google Translate 무료 엔드포인트)"""
    import requests
    try:
        url = "https://translate.googleapis.com/translate_a/single"
        params = {"client": "gtx", "sl": "en", "tl": "ko", "dt": "t", "q": query}
        resp = requests.get(url, params=params, timeout=5)
        if resp.status_code == 200:
            data = resp.json()
            translated = data[0][0][0]
            # 번역이 원본과 같으면 (고유명사라 번역 안 된 경우) None 반환
            if translated.strip().lower() == query.strip().lower():
                return None
            return translated
    except Exception:
        pass
    return None


class InternationalFetcher:
    """yfinance 기반 해외 주식 데이터 수집기

    DataFetcher와 동일한 인터페이스를 제공합니다.
    """

    def __init__(self):
        import yfinance as yf
        self.yf = yf
        self._ticker_cache: dict[str, object] = {}
        self._exchange_rate_cache: dict[str, float] = {}

    def _get_ticker(self, symbol: str):
        """Ticker 객체 캐싱"""
        if symbol not in self._ticker_cache:
            self._ticker_cache[symbol] = self.yf.Ticker(symbol)
        return self._ticker_cache[symbol]

    # ── 환율 ──────────────────────────────────────────────────

    def fetch_exchange_rate(self, from_currency: str = "USD",
                            to_currency: str = "KRW") -> float:
        """환율 조회 (캐싱)"""
        pair = f"{from_currency}{to_currency}"
        if pair in self._exchange_rate_cache:
            return self._exchange_rate_cache[pair]

        try:
            ticker = self._get_ticker(f"{pair}=X")
            info = ticker.info
            rate = info.get("regularMarketPrice") or info.get("ask", 0)
            if rate and rate > 0:
                self._exchange_rate_cache[pair] = float(rate)
                return float(rate)
        except Exception:
            pass

        # 기본 환율 (fallback)
        defaults = {"USDKRW": 1450.0, "EURKRW": 1550.0, "JPYKRW": 9.5}
        rate = defaults.get(pair, 1.0)
        self._exchange_rate_cache[pair] = rate
        return rate

    # ── CAPM 데이터 ──────────────────────────────────────────────

    def fetch_beta(self, symbol: str) -> float:
        """베타 계수 조회"""
        ticker = self._get_ticker(symbol)
        info = ticker.info
        beta = info.get("beta", None)
        if beta and isinstance(beta, (int, float)):
            return round(float(beta), 2)
        return 1.0  # 기본값

    def fetch_risk_free_rate(self) -> float:
        """무위험수익률 조회 (미국 10년 국채 수익률)"""
        try:
            ticker = self._get_ticker("^TNX")
            info = ticker.info
            rate = info.get("regularMarketPrice") or info.get("previousClose", 0)
            if rate and isinstance(rate, (int, float)) and rate > 0:
                return round(float(rate), 2)
        except Exception:
            pass
        return 4.0  # 기본 fallback

    # ── 종목 검색 ──────────────────────────────────────────────

    def search_stock(self, query: str, limit: int = 20) -> list[dict]:
        """종목 검색 (한국어 입력 지원)"""
        queries_to_try = [query]

        # 한국어면 영어로 번역하여 추가 검색
        if _has_korean(query):
            translated = _translate_ko_to_en(query)
            if translated:
                queries_to_try = [translated, query]

        for q in queries_to_try:
            try:
                results = self.yf.Search(q)
                quotes = results.quotes if hasattr(results, 'quotes') else []
                if quotes:
                    output = []
                    for item in quotes[:limit]:
                        symbol = item.get("symbol", "")
                        name = item.get("shortname") or item.get("longname") or symbol
                        exchange = item.get("exchange", "")
                        output.append({
                            "corp_name": name,
                            "stock_code": symbol,
                            "exchange": exchange,
                        })
                    return output
            except Exception:
                continue

        # fallback: 직접 ticker 확인
        symbol = query.strip().upper()
        if symbol.isascii():
            try:
                ticker = self._get_ticker(symbol)
                info = ticker.info
                if info and info.get("symbol"):
                    return [{
                        "corp_name": info.get("shortName", query),
                        "stock_code": info["symbol"],
                        "exchange": info.get("exchange", ""),
                    }]
            except Exception:
                pass
        return []

    def resolve_stock_query(self, query: str) -> str:
        """티커 심볼 확인 및 반환 (한국어 지원)"""
        stripped = query.strip()

        # 한국어 입력이면 검색으로 티커 찾기
        if _has_korean(stripped):
            results = self.search_stock(stripped, limit=5)
            if results:
                return results[0]["stock_code"]
            raise ValueError(f"'{query}'에 해당하는 해외 종목을 찾을 수 없습니다.")

        symbol = stripped.upper()
        ticker = self._get_ticker(symbol)
        info = ticker.info
        if not info or not info.get("symbol"):
            raise ValueError(f"'{query}' 티커를 찾을 수 없습니다.")
        return info["symbol"]

    # ── 기업 정보 ──────────────────────────────────────────────

    def fetch_company_info(self, symbol: str) -> dict:
        """기업 기본 정보"""
        ticker = self._get_ticker(symbol)
        info = ticker.info

        corp_name = info.get("shortName") or info.get("longName", symbol)

        # 한글 회사명 조회
        corp_name_kr = _translate_en_to_ko(corp_name) or ""

        return {
            "corp_code": symbol,
            "corp_name": corp_name,
            "corp_name_kr": corp_name_kr,
            "stock_name": info.get("shortName", symbol),
            "stock_code": symbol,
            "corp_cls": "",
            "est_dt": "",
            "sector": info.get("sector", ""),
            "industry": info.get("industry", ""),
            "currency": info.get("currency", "USD"),
            "exchange": info.get("exchange", ""),
            "is_international": True,
        }

    # ── 주가 / 시가총액 ───────────────────────────────────────

    def fetch_stock_data(self, symbol: str, shares: int = 0) -> dict:
        """현재 주가, 시가총액 (환율 포함)"""
        ticker = self._get_ticker(symbol)
        info = ticker.info

        price = info.get("currentPrice") or info.get("regularMarketPrice", 0)
        market_cap = info.get("marketCap", 0)
        shares_out = shares or info.get("sharesOutstanding", 0)
        currency = info.get("currency", "USD")

        # 환율 조회 및 원화 환산
        exchange_rate = self.fetch_exchange_rate(currency, "KRW")
        price_krw = round(price * exchange_rate) if price else 0
        market_cap_krw = round(market_cap * exchange_rate) if market_cap else 0

        return {
            "price": price,
            "price_krw": price_krw,
            "market_cap": market_cap,
            "market_cap_krw": market_cap_krw,
            "market_cap_eok": round(market_cap_krw / 1e8) if market_cap_krw else 0,
            "shares": shares_out,
            "date": datetime.now().strftime("%Y-%m-%d"),
            "currency": currency,
            "exchange_rate": exchange_rate,
        }

    def fetch_shares_outstanding(self, symbol: str, year: int) -> int:
        """발행주식수"""
        ticker = self._get_ticker(symbol)
        info = ticker.info
        shares = info.get("sharesOutstanding", 0)
        if not shares:
            # balance sheet에서 시도
            bs = ticker.balance_sheet
            col = _find_year_column(bs, year)
            if col is not None and "Ordinary Shares Number" in bs.index:
                shares = _safe_int(bs.loc["Ordinary Shares Number", col])
        return shares or 0

    # ── 연도 탐색 ──────────────────────────────────────────────

    def find_latest_available_year(self, symbol: str) -> int:
        """최신 재무제표 연도"""
        ticker = self._get_ticker(symbol)
        financials = ticker.financials
        if financials is not None and not financials.empty:
            latest_col = financials.columns[0]
            if hasattr(latest_col, "year"):
                return latest_col.year
        return datetime.now().year - 1

    # ── 재무제표 ───────────────────────────────────────────────

    def fetch_financial_summary(self, symbol: str, years: list[int]) -> dict:
        """주요계정 (손익+재무상태표) → 한국어 키 딕셔너리"""
        ticker = self._get_ticker(symbol)
        income_df = ticker.financials
        bs_df = ticker.balance_sheet
        result = {}

        for year in years:
            year_data = {}

            # 손익계산서
            is_col = _find_year_column(income_df, year)
            if is_col is not None:
                for en_name, kr_name in IS_FIELD_MAP.items():
                    if en_name in income_df.index:
                        val = _safe_int(income_df.loc[en_name, is_col])
                        year_data[kr_name] = val
                        year_data[f"IS_{kr_name}"] = val

            # 재무상태표
            bs_col = _find_year_column(bs_df, year)
            if bs_col is not None:
                for en_name, kr_name in BS_FIELD_MAP.items():
                    if en_name in bs_df.index:
                        val = _safe_int(bs_df.loc[en_name, bs_col])
                        year_data[kr_name] = val
                        year_data[f"BS_{kr_name}"] = val

            if year_data:
                result[year] = year_data

        return result

    def fetch_balance_sheet_detail(self, symbol: str, years: list[int]) -> dict:
        """재무상태표 세부항목"""
        ticker = self._get_ticker(symbol)
        bs_df = ticker.balance_sheet
        result = {}

        for year in years:
            year_data = {}
            col = _find_year_column(bs_df, year)
            if col is not None:
                for en_name, kr_name in BS_FIELD_MAP.items():
                    if en_name in bs_df.index:
                        year_data[kr_name] = _safe_int(bs_df.loc[en_name, col])
                # 추가 항목 직접 매핑
                for row_name in bs_df.index:
                    val = _safe_int(bs_df.loc[row_name, col])
                    if val:
                        year_data[f"__{row_name}"] = val
            if year_data:
                result[year] = year_data

        return result

    def fetch_cash_flow_detail(self, symbol: str, years: list[int]) -> dict:
        """현금흐름표 세부항목"""
        ticker = self._get_ticker(symbol)
        cf_df = ticker.cashflow
        result = {}

        for year in years:
            year_data = {}
            col = _find_year_column(cf_df, year)
            if col is not None:
                for en_name, kr_name in CF_FIELD_MAP.items():
                    if en_name in cf_df.index:
                        year_data[kr_name] = _safe_int(cf_df.loc[en_name, col])
            if year_data:
                result[year] = year_data

        return result

    # ── 분기별 데이터 수집 ────────────────────────────────────────

    def fetch_quarterly_data(self, symbol: str, num_years: int = 2) -> dict:
        """분기별 재무데이터 수집 (yfinance — 이미 단독 분기 데이터, 역산 불필요)

        Args:
            symbol: 티커 심볼
            num_years: 수집할 연도 수 (기본 2년 = 최대 8분기)

        Returns:
            {
                "quarters": ["2023Q1", ..., "2024Q4"],
                "quarterly_summary": {"2023Q1": {계정: 금액}, ...},
                "quarterly_bs": {"2023Q1": {계정: 금액}, ...},
                "quarterly_cf": {"2023Q1": {계정: 금액}, ...},
            }
        """
        ticker = self._get_ticker(symbol)
        qtr_is = ticker.quarterly_financials
        qtr_bs = ticker.quarterly_balance_sheet
        qtr_cf = ticker.quarterly_cashflow

        # 분기 컬럼 수집 (IS 기준으로 가용 분기 파악)
        quarter_cols = []
        if qtr_is is not None and not qtr_is.empty:
            for col in qtr_is.columns:
                if hasattr(col, "year") and hasattr(col, "quarter"):
                    quarter_cols.append(col)

        # 시간순 정렬 후 최근 num_years*4 분기만
        quarter_cols.sort()
        quarter_cols = quarter_cols[-(num_years * 4):]

        quarters = []
        quarterly_summary = {}
        quarterly_bs = {}
        quarterly_cf = {}

        for col in quarter_cols:
            qkey = f"{col.year}Q{col.quarter}"
            quarters.append(qkey)

            # IS (손익계산서)
            q_is = {}
            if qtr_is is not None and col in qtr_is.columns:
                for en_name, kr_name in IS_FIELD_MAP.items():
                    if en_name in qtr_is.index:
                        val = _safe_int(qtr_is.loc[en_name, col])
                        q_is[kr_name] = val
                        q_is[f"IS_{kr_name}"] = val
            quarterly_summary[qkey] = q_is

            # BS (재무상태표)
            q_bs = {}
            bs_col = self._find_quarter_col(qtr_bs, col)
            if bs_col is not None:
                for en_name, kr_name in BS_FIELD_MAP.items():
                    if en_name in qtr_bs.index:
                        q_bs[kr_name] = _safe_int(qtr_bs.loc[en_name, bs_col])
            quarterly_bs[qkey] = q_bs

            # CF (현금흐름표)
            q_cf = {}
            cf_col = self._find_quarter_col(qtr_cf, col)
            if cf_col is not None:
                for en_name, kr_name in CF_FIELD_MAP.items():
                    if en_name in qtr_cf.index:
                        q_cf[kr_name] = _safe_int(qtr_cf.loc[en_name, cf_col])
            quarterly_cf[qkey] = q_cf

        return {
            "quarters": quarters,
            "quarterly_summary": quarterly_summary,
            "quarterly_bs": quarterly_bs,
            "quarterly_cf": quarterly_cf,
        }

    @staticmethod
    def _find_quarter_col(df: pd.DataFrame, target_col):
        """DataFrame에서 동일 분기 컬럼 찾기"""
        if df is None or df.empty:
            return None
        for col in df.columns:
            if hasattr(col, "year") and hasattr(col, "quarter"):
                if col.year == target_col.year and col.quarter == target_col.quarter:
                    return col
        return None

    # ── 주주 현황 ──────────────────────────────────────────────

    def fetch_major_shareholders(self, symbol: str) -> list[dict]:
        """주요 주주 현황"""
        ticker = self._get_ticker(symbol)
        shareholders = []

        # major_holders: index=breakdown명, columns=["Value"]
        _MAJOR_LABELS = {
            "insidersPercentHeld": "내부자 지분",
            "institutionsPercentHeld": "기관 지분",
            "institutionsFloatPercentHeld": "유동주식 중 기관",
        }
        try:
            major = ticker.major_holders
            if major is not None and not major.empty:
                for idx_name, row in major.iterrows():
                    label = _MAJOR_LABELS.get(str(idx_name))
                    if not label:
                        continue  # institutionsCount 등 제외
                    val = row.iloc[0]
                    if isinstance(val, (int, float)) and val > 0:
                        ratio = float(val) * 100 if val < 1 else float(val)
                        shareholders.append({
                            "name": label,
                            "shares": 0,
                            "ratio": round(ratio, 2),
                        })
        except Exception:
            pass

        # institutional_holders: 상위 기관투자자
        try:
            inst = ticker.institutional_holders
            if inst is not None and not inst.empty:
                for _, row in inst.head(5).iterrows():
                    name = row.get("Holder", "")
                    pct = row.get("pctHeld", 0)
                    if pct and isinstance(pct, (int, float)) and pct > 0:
                        ratio = float(pct) * 100 if pct < 1 else float(pct)
                        shareholders.append({
                            "name": str(name)[:20],
                            "shares": int(row.get("Shares", 0)),
                            "ratio": round(ratio, 2),
                        })
        except Exception:
            pass

        return shareholders[:5]

    # ── 연도별 종가 ────────────────────────────────────────────

    def fetch_valuation_by_year(self, symbol: str, years: list[int]) -> dict:
        """연도별 연말 종가"""
        ticker = self._get_ticker(symbol)
        result = {}

        for year in years:
            try:
                start = f"{year}-12-01"
                end = f"{year + 1}-01-15"
                hist = ticker.history(start=start, end=end)
                if hist is not None and not hist.empty:
                    # 해당 연도 마지막 거래일
                    year_data = hist[hist.index.year == year]
                    if not year_data.empty:
                        close = float(year_data["Close"].iloc[-1])
                        result[year] = {"close": close}
                time.sleep(0.1)
            except Exception:
                continue

        return result

    # ── 컨센서스 ───────────────────────────────────────────────

    def fetch_consensus(self, symbol: str) -> dict:
        """애널리스트 컨센서스 데이터"""
        ticker = self._get_ticker(symbol)
        info = ticker.info
        consensus = {
            "target_price": 0,
            "opinion": 0,
            "items": [],
        }

        # 목표주가
        target = info.get("targetMeanPrice", 0)
        if target:
            consensus["target_price"] = target

        # 투자의견 (1=Strong Buy ~ 5=Sell → 5점 만점으로 변환)
        rec = info.get("recommendationMean", 0)
        if rec:
            consensus["opinion"] = round(6 - rec, 1)  # 1→5, 5→1 변환

        # Forward EPS/BPS로 ROE 계산
        forward_eps = info.get("forwardEps", 0)
        bps = info.get("bookValue", 0)
        if forward_eps and bps and bps > 0:
            consensus["consensus_roe"] = round(forward_eps / bps * 100, 2)
            consensus["consensus_eps"] = forward_eps
            consensus["consensus_bps"] = bps

        # 추가 컨센서스 항목
        items = []
        for label, key in [
            ("Forward EPS", "forwardEps"),
            ("Trailing EPS", "trailingEps"),
            ("Forward PE", "forwardPE"),
            ("PEG Ratio", "pegRatio"),
            ("Book Value", "bookValue"),
            ("Target Price", "targetMeanPrice"),
        ]:
            val = info.get(key)
            if val:
                items.append({
                    "label": label,
                    "values": [f"{val:.2f}" if isinstance(val, float) else str(val)],
                })
        consensus["items"] = items

        return consensus
