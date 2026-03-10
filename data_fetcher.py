"""DART OpenAPI 및 pykrx를 통한 재무 데이터 수집 모듈"""

import OpenDartReader
from pykrx import stock
import pandas as pd
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import time
import re

from config import (
    DART_API_KEY,
    REPORT_TYPE_ANNUAL,
    QUARTERLY_REPORT_CODES,
    UNIT_BILLION,
)


class DataFetcher:
    def __init__(self, api_key=None):
        self.api_key = api_key or DART_API_KEY
        if not self.api_key:
            raise ValueError(
                "DART API 키가 필요합니다.\n"
                ".env 파일에 DART_API_KEY를 설정하거나 --api-key 옵션을 사용하세요.\n"
                "API 키 발급: https://opendart.fss.or.kr"
            )
        self.dart = OpenDartReader(self.api_key)
        self._finstate_all_cache: dict[tuple, pd.DataFrame | None] = {}

    # ── 종목 검색 ──────────────────────────────────────────────

    def search_stock(self, query: str, limit: int = 20) -> list[dict]:
        """종목명 또는 종목코드로 상장기업 검색

        Args:
            query: 검색어 (종목명 일부 또는 종목코드)
            limit: 최대 결과 수

        Returns:
            [{"corp_name": str, "stock_code": str}, ...]
        """
        cc = self.dart.corp_codes
        # 상장기업만 (6자리 종목코드 보유)
        listed = cc[cc["stock_code"].str.match(r"^\d{6}$", na=False)]

        # 종목코드로 검색 (숫자만 입력된 경우)
        if query.isdigit():
            matches = listed[listed["stock_code"].str.contains(query, na=False)]
        else:
            # 종목명으로 검색 (대소문자 무시)
            matches = listed[listed["corp_name"].str.contains(query, case=False, na=False)]

        results = []
        for _, row in matches.head(limit).iterrows():
            results.append({
                "corp_name": row["corp_name"],
                "stock_code": row["stock_code"],
            })
        return results

    def resolve_stock_query(self, query: str) -> str:
        """검색어를 종목코드로 변환. 6자리 숫자면 그대로 반환, 아니면 검색.

        정확히 1건 매칭되면 자동 선택, 여러 건이면 ValueError.
        """
        query = query.strip()
        if len(query) == 6 and query.isdigit():
            return query

        results = self.search_stock(query)
        if not results:
            raise ValueError(f"'{query}'에 해당하는 상장기업을 찾을 수 없습니다.")
        if len(results) == 1:
            return results[0]["stock_code"]

        # 정확히 일치하는 이름 우선
        for r in results:
            if r["corp_name"] == query:
                return r["stock_code"]

        # 여러 건 → 목록 표시
        lines = [f"'{query}' 검색 결과 ({len(results)}건):"]
        for r in results[:10]:
            lines.append(f"  {r['stock_code']} - {r['corp_name']}")
        if len(results) > 10:
            lines.append(f"  ... 외 {len(results) - 10}건")
        lines.append("종목코드를 직접 입력하세요.")
        raise ValueError("\n".join(lines))

    # ── 기업 기본 정보 ──────────────────────────────────────────

    def fetch_company_info(self, stock_code: str) -> dict:
        """기업 기본 정보 조회 (DART 기업개황)"""
        corp_code = self.dart.find_corp_code(stock_code)
        if not corp_code:
            raise ValueError(f"종목코드 {stock_code}에 해당하는 기업을 찾을 수 없습니다.")

        company = self.dart.company(corp_code)
        return {
            "corp_code": corp_code,
            "corp_name": company.get("corp_name", ""),
            "stock_name": company.get("stock_name", ""),
            "stock_code": stock_code,
            "corp_cls": company.get("corp_cls", ""),
            "est_dt": company.get("est_dt", ""),
        }

    # ── 주가 / 시가총액 ─────────────────────────────────────────

    def fetch_stock_data(self, stock_code: str, shares: int = 0) -> dict:
        """현재 주가 및 시가총액 조회

        pykrx OHLCV로 주가 조회, 시가총액 = 주가 × 발행주식수
        """
        today = datetime.now().strftime("%Y%m%d")
        start = (datetime.now() - timedelta(days=14)).strftime("%Y%m%d")

        latest_price = 0
        trade_date = today

        try:
            ohlcv = stock.get_market_ohlcv_by_date(start, today, stock_code)
            if ohlcv is not None and not ohlcv.empty:
                latest_price = int(ohlcv.iloc[-1]["종가"])
                trade_date = ohlcv.index[-1].strftime("%Y%m%d")
        except Exception as e:
            print(f"  pykrx 주가 조회 실패: {e}")

        market_cap = latest_price * shares if shares else 0

        return {
            "price": latest_price,
            "market_cap": market_cap,
            "market_cap_eok": round(market_cap / UNIT_BILLION) if market_cap else 0,
            "shares": shares,
            "date": trade_date,
        }

    # ── 발행주식수 계산 ──────────────────────────────────────────

    def fetch_shares_outstanding(self, stock_code: str, year: int) -> int:
        """발행주식수 추출

        finstate_all의 EPS와 당기순이익(지배)으로 역산
        """
        df = self._fetch_finstate_all_cached(stock_code, year)
        if df is None:
            return 0

        # IS/CIS에서 EPS, 지배주주 순이익, 전체 순이익 찾기
        is_rows = df[df["sj_div"] == "IS"] if "sj_div" in df.columns else df
        cis_rows = df[df["sj_div"] == "CIS"] if "sj_div" in df.columns else pd.DataFrame()
        eps = 0
        net_income_ctrl = 0
        net_income = 0

        # IS 먼저 검색 (우선순위 높음)
        for src in [is_rows, cis_rows]:
            for _, row in src.iterrows():
                nm = row.get("account_nm", "")
                amt = self._parse_amount(row.get("thstrm_amount", ""))
                if not eps and "기본주당" in nm and "이익" in nm:
                    eps = amt
                # 지배주주 순이익: IS의 "지배기업 소유지분/소유주지분"
                if not net_income_ctrl and "지배기업" in nm and "소유" in nm:
                    # IS에서만 가져옴 (CIS 총포괄이익 제외)
                    if src is is_rows:
                        net_income_ctrl = amt
                if not net_income and nm in ("당기순이익", "당기순이익(손실)"):
                    net_income = amt

        # CIS에서 지배주주 순이익이 IS에 없을 경우 (하이닉스 등)
        if not net_income_ctrl:
            saw_net_income = False
            for _, row in cis_rows.iterrows():
                nm = row.get("account_nm", "")
                amt = self._parse_amount(row.get("thstrm_amount", ""))
                if nm in ("당기순이익", "당기순이익(손실)"):
                    saw_net_income = True
                # 당기순이익 아래의 지배기업 소유주지분만 취함
                if saw_net_income and "지배기업" in nm and "소유" in nm:
                    net_income_ctrl = amt
                    break

        # 1차: 지배주주 순이익 / EPS
        if eps and net_income_ctrl:
            return abs(net_income_ctrl) // abs(eps)

        # 2차: 전체 순이익 / EPS (비지배지분 비중 작은 경우 근사치)
        if eps and net_income:
            return abs(net_income) // abs(eps)

        # 3차 폴백: 자본금 / 액면가 (5000원 시도 후 1000원)
        bs = df[df["sj_div"] == "BS"] if "sj_div" in df.columns else df
        for _, row in bs.iterrows():
            if row.get("account_nm", "") == "자본금":
                capital = self._parse_amount(row.get("thstrm_amount", ""))
                if capital:
                    # 액면가 5,000원 시도: 결과가 합리적이면 사용
                    shares_5k = capital // 5000
                    shares_1k = capital // 1000
                    # EPS가 있으면 어느 쪽이 합리적인지 판단
                    if eps:
                        return shares_5k  # EPS는 있는데 순이익이 없는 경우
                    return shares_5k  # 대부분 액면가 5,000원

        return 0

    # ── 최신 이용 가능 연도 탐색 ─────────────────────────────────

    def find_latest_available_year(self, stock_code: str) -> int:
        """DART에서 사업보고서가 존재하는 최신 연도 반환"""
        current_year = datetime.now().year
        for year in range(current_year - 1, current_year - 3, -1):
            try:
                df = self.dart.finstate(stock_code, year, reprt_code=REPORT_TYPE_ANNUAL)
                if df is not None and isinstance(df, pd.DataFrame) and not df.empty:
                    return year
            except Exception:
                pass
            time.sleep(0.3)
        return current_year - 2  # 안전한 폴백

    # ── 주요계정 (5개년) ─────────────────────────────────────────

    def fetch_financial_summary(self, stock_code: str, years: list[int]) -> dict:
        """단일회사 주요계정 + finstate_all 보강 데이터 조회"""
        result = {}

        for year in years:
            try:
                year_data = {}

                # 1) finstate (주요계정)
                df = self.dart.finstate(stock_code, year, reprt_code=REPORT_TYPE_ANNUAL)
                if df is not None and isinstance(df, pd.DataFrame) and not df.empty:
                    if "fs_div" in df.columns:
                        cfs = df[df["fs_div"] == "CFS"]
                        if cfs.empty:
                            cfs = df[df["fs_div"] == "OFS"]
                        df = cfs

                    for _, row in df.iterrows():
                        account = row.get("account_nm", "")
                        if not account:
                            continue
                        amount = self._parse_amount(row.get("thstrm_amount", ""))
                        sj_div = row.get("sj_div", "")
                        key = f"{sj_div}_{account}" if sj_div else account
                        year_data[key] = amount
                        if account not in year_data:
                            year_data[account] = amount

                # 2) finstate_all로 보강 (지배주주 자본, EPS/BPS 등)
                df_all = self._fetch_finstate_all_cached(stock_code, year)
                if df_all is not None:
                    for _, row in df_all.iterrows():
                        nm = row.get("account_nm", "")
                        sj = row.get("sj_div", "")
                        amt = self._parse_amount(row.get("thstrm_amount", ""))
                        full_key = f"{sj}_{nm}"
                        if full_key not in year_data:
                            year_data[full_key] = amt
                        if nm not in year_data:
                            year_data[nm] = amt

                if year_data:
                    result[year] = year_data
                time.sleep(0.3)
            except Exception as e:
                print(f"  {year}년 주요계정 조회 실패: {e}")
                continue

        return result

    # ── 재무상태표 세부항목 ───────────────────────────────────────

    def fetch_balance_sheet_detail(self, stock_code: str, years: list[int]) -> dict:
        """재무상태표 세부 항목 조회"""
        result = {}

        for year in years:
            try:
                df = self._fetch_finstate_all_cached(stock_code, year)
                if df is None:
                    continue

                bs = df[df["sj_div"] == "BS"] if "sj_div" in df.columns else df

                year_data = {}
                for _, row in bs.iterrows():
                    account = row.get("account_nm", "")
                    if not account:
                        continue
                    amount = self._parse_amount(row.get("thstrm_amount", ""))
                    year_data[account] = amount

                result[year] = year_data
            except Exception as e:
                print(f"  {year}년 재무상태표 조회 실패: {e}")
                continue

        return result

    # ── 현금흐름표 세부항목 ───────────────────────────────────────

    def fetch_cash_flow_detail(self, stock_code: str, years: list[int]) -> dict:
        """현금흐름표 세부 항목 조회"""
        result = {}

        for year in years:
            try:
                df = self._fetch_finstate_all_cached(stock_code, year)
                if df is None:
                    continue

                cf = df[df["sj_div"] == "CF"] if "sj_div" in df.columns else df

                year_data = {}
                for _, row in cf.iterrows():
                    account = row.get("account_nm", "")
                    if not account:
                        continue
                    amount = self._parse_amount(row.get("thstrm_amount", ""))
                    year_data[account] = amount

                result[year] = year_data
            except Exception as e:
                print(f"  {year}년 현금흐름표 조회 실패: {e}")
                continue

        return result

    # ── 주요 주주 현황 ───────────────────────────────────────────

    def fetch_major_shareholders(self, stock_code: str) -> list[dict]:
        """주요 주주 현황 조회"""
        try:
            df = self.dart.major_shareholders(stock_code)
            if df is None or (isinstance(df, pd.DataFrame) and df.empty):
                return []

            shareholders = []
            seen = set()
            for _, row in df.iterrows():
                name = str(row.get("repror", row.get("nm", "")))
                if not name or name in seen or name == "nan":
                    continue
                seen.add(name)

                ratio = row.get("stkrt", row.get("bsis_posesn_stock_qota_rt", 0))
                try:
                    ratio = float(str(ratio).replace(",", ""))
                except (ValueError, TypeError):
                    ratio = 0.0

                shares_count = row.get("stkqy", row.get("bsis_posesn_stock_co", 0))
                try:
                    shares_count = int(str(shares_count).replace(",", ""))
                except (ValueError, TypeError):
                    shares_count = 0

                shareholders.append({
                    "name": name,
                    "shares": shares_count,
                    "ratio": ratio,
                })

                if len(shareholders) >= 5:
                    break

            return shareholders
        except Exception as e:
            print(f"  주요 주주 조회 실패: {e}")
            return []

    # ── 연도별 종가 (PER/PBR 계산용) ─────────────────────────────

    def fetch_valuation_by_year(self, stock_code: str, years: list[int]) -> dict:
        """연도별 종가 및 시가총액 조회"""
        result = {}
        for year in years:
            try:
                start = f"{year}0101"
                end = f"{year}1231"
                df = stock.get_market_ohlcv_by_date(start, end, stock_code, freq="y")
                if df is not None and not df.empty:
                    result[year] = {"close": int(df.iloc[-1]["종가"])}
                time.sleep(0.2)
            except Exception as e:
                print(f"  {year}년 종가 조회 실패: {e}")
                continue

        return result

    # ── 컨센서스 (네이버 금융 스크래핑) ──────────────────────────

    def fetch_consensus(self, stock_code: str) -> dict:
        """네이버 금융에서 컨센서스 데이터 스크래핑

        Returns:
            {
                "target_price": 목표주가,
                "opinion": 투자의견 (점수),
                "items": [
                    {"label": "매출액", "current": ..., "3m_ago": ..., "6m_ago": ...},
                    ...
                ]
            }
        """
        try:
            return self._fetch_naver_consensus(stock_code)
        except Exception as e:
            print(f"  컨센서스 조회 실패: {e}")
            return {}

    def _fetch_naver_consensus(self, stock_code: str) -> dict:
        """네이버 금융 컨센서스 페이지 스크래핑"""
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }

        result = {"items": []}

        # 1) 투자의견/목표주가 (컨센서스 개요)
        url = f"https://finance.naver.com/item/coinfo.naver?code={stock_code}"
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            resp.encoding = "euc-kr"
        except Exception:
            return result

        # 2) 네이버 증권 종목분석 API (투자의견 컨센서스)
        consensus_url = (
            f"https://navercomp.wisereport.co.kr/v2/company/c1060001.aspx"
            f"?cmp_cd={stock_code}&cn="
        )
        try:
            resp2 = requests.get(consensus_url, headers=headers, timeout=10)
            resp2.encoding = "utf-8"
            soup = BeautifulSoup(resp2.text, "html.parser")

            # 컨센서스 테이블 파싱
            tables = soup.select("table.gHead")
            for table in tables:
                rows = table.select("tr")
                for row in rows:
                    cols = row.select("td, th")
                    if len(cols) >= 2:
                        label = cols[0].get_text(strip=True)
                        if label in ("투자의견", "목표주가", "매출액", "영업이익", "순이익",
                                     "EPS", "PER", "BPS", "PBR", "ROE"):
                            values = [c.get_text(strip=True) for c in cols[1:]]
                            if label == "목표주가":
                                try:
                                    result["target_price"] = int(
                                        values[0].replace(",", "").replace("원", "")
                                    )
                                except (ValueError, IndexError):
                                    pass
                            elif label == "투자의견":
                                try:
                                    result["opinion"] = float(values[0])
                                except (ValueError, IndexError):
                                    pass
                            else:
                                item = {"label": label, "values": values}
                                result["items"].append(item)
        except Exception as e:
            print(f"  컨센서스 상세 조회 실패: {e}")

        # 3) 폴백: 네이버 증권 기본 컨센서스 정보
        if not result.get("target_price"):
            try:
                api_url = (
                    f"https://navercomp.wisereport.co.kr/v2/company/c1010001.aspx"
                    f"?cmp_cd={stock_code}"
                )
                resp3 = requests.get(api_url, headers=headers, timeout=10)
                resp3.encoding = "utf-8"
                soup3 = BeautifulSoup(resp3.text, "html.parser")

                # 목표주가
                for dt in soup3.select("dt"):
                    text = dt.get_text(strip=True)
                    if "목표주가" in text:
                        dd = dt.find_next_sibling("dd")
                        if dd:
                            price_text = dd.get_text(strip=True)
                            nums = re.findall(r"[\d,]+", price_text)
                            if nums:
                                result["target_price"] = int(nums[0].replace(",", ""))

                # 투자의견
                for dt in soup3.select("dt"):
                    text = dt.get_text(strip=True)
                    if "투자의견" in text:
                        dd = dt.find_next_sibling("dd")
                        if dd:
                            try:
                                result["opinion"] = float(dd.get_text(strip=True))
                            except ValueError:
                                pass

                # 실적 컨센서스 테이블
                tables = soup3.select("table")
                for table in tables:
                    caption = table.select_one("caption")
                    if caption and "컨센서스" in caption.get_text():
                        rows = table.select("tbody tr")
                        for row in rows:
                            cols = row.select("td")
                            th = row.select_one("th")
                            if th and len(cols) >= 1:
                                label = th.get_text(strip=True)
                                values = [c.get_text(strip=True) for c in cols]
                                result["items"].append({"label": label, "values": values})
            except Exception:
                pass

        # 4) 컨센서스 ROE 역산 (EPS / BPS * 100)
        self._calc_consensus_roe(result)

        return result

    @staticmethod
    def _calc_consensus_roe(consensus: dict):
        """컨센서스 EPS/BPS에서 ROE 역산하여 consensus에 추가"""
        eps_val = 0
        bps_val = 0
        for item in consensus.get("items", []):
            label = item.get("label", "")
            values = item.get("values", [])
            # 마지막 값 = 가장 미래 예측치
            raw = values[-1] if values else ""
            parsed = re.sub(r"[^\d.\-]", "", raw.replace(",", ""))
            if not parsed:
                continue
            try:
                num = float(parsed)
            except ValueError:
                continue
            if label == "EPS":
                eps_val = num
            elif label == "BPS":
                bps_val = num

        if eps_val and bps_val and bps_val > 0:
            consensus["consensus_roe"] = round(eps_val / bps_val * 100, 2)
            consensus["consensus_eps"] = eps_val
            consensus["consensus_bps"] = bps_val

    # ── CAPM 데이터 ──────────────────────────────────────────────

    def fetch_beta(self, stock_code: str) -> float:
        """베타 계수 조회 (네이버 금융 → pykrx 직접 계산 fallback)"""
        # 1차: 네이버 금융 스크래핑
        try:
            url = f"https://finance.naver.com/item/main.naver?code={stock_code}"
            headers = {
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
            resp = requests.get(url, headers=headers, timeout=10)
            resp.encoding = "euc-kr"
            soup = BeautifulSoup(resp.text, "html.parser")
            for em in soup.select("em"):
                text = em.get_text(strip=True)
                parent = em.parent
                if parent and "베타" in parent.get_text():
                    try:
                        val = float(text)
                        if 0 < val < 5:  # 합리적 범위 체크
                            return round(val, 2)
                    except ValueError:
                        pass
        except Exception:
            pass

        # 2차: pykrx + yfinance로 직접 계산 (1년 일별 수익률 기반, KOSPI 대비)
        try:
            end = datetime.now()
            start = end - timedelta(days=365)
            start_str = start.strftime("%Y%m%d")
            end_str = end.strftime("%Y%m%d")

            stock_prices = stock.get_market_ohlcv(start_str, end_str, stock_code)

            # KOSPI 지수는 yfinance ^KS11 사용 (pykrx index API 호환성 이슈 방지)
            import yfinance as yf
            kospi = yf.Ticker("^KS11")
            kospi_hist = kospi.history(period="1y")

            if stock_prices is not None and not stock_prices.empty and \
               kospi_hist is not None and not kospi_hist.empty:
                stock_ret = stock_prices["종가"].pct_change().dropna()
                kospi_ret = kospi_hist["Close"].pct_change().dropna()

                # 인덱스 정규화 (timezone 제거)
                stock_ret.index = stock_ret.index.tz_localize(None) if stock_ret.index.tz is None else stock_ret.index
                kospi_ret.index = kospi_ret.index.tz_localize(None)

                # 공통 날짜만 사용
                common_idx = stock_ret.index.intersection(kospi_ret.index)
                if len(common_idx) >= 60:  # 최소 60거래일
                    s = stock_ret.loc[common_idx]
                    m = kospi_ret.loc[common_idx]
                    cov = s.cov(m)
                    var = m.var()
                    if var > 0:
                        beta = cov / var
                        if 0 < beta < 5:
                            return round(beta, 2)
        except Exception:
            pass

        return 1.0  # 최종 기본값

    def fetch_risk_free_rate(self) -> float:
        """무위험수익률 조회 (한국 국고채 3년물 수익률)"""
        # 1차: 네이버 금융 — 국고채 3년물
        try:
            url = "https://finance.naver.com/marketindex/interestDailyQuote.naver?marketindexCd=IRR_GOVT03Y"
            headers = {
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
            resp = requests.get(url, headers=headers, timeout=10)
            resp.encoding = "euc-kr"
            soup = BeautifulSoup(resp.text, "html.parser")
            # 테이블: 날짜 | 수익률 | 변동폭 | 변동률 → tds[1]이 수익률
            tds = soup.select("table.tbl_exchange tbody tr td")
            if len(tds) >= 2:
                rate_text = tds[1].get_text(strip=True).replace(",", "")
                rate = float(rate_text)
                if 0 < rate < 20:
                    return round(rate, 2)
        except Exception:
            pass

        # 2차: 네이버 금융 국고채 10년물
        try:
            url = "https://finance.naver.com/marketindex/interestDailyQuote.naver?marketindexCd=IRR_GOVT10Y"
            headers = {"User-Agent": "Mozilla/5.0"}
            resp = requests.get(url, headers=headers, timeout=10)
            resp.encoding = "euc-kr"
            soup = BeautifulSoup(resp.text, "html.parser")
            tds = soup.select("table.tbl_exchange tbody tr td")
            if len(tds) >= 2:
                rate_text = tds[1].get_text(strip=True).replace(",", "")
                rate = float(rate_text)
                if 0 < rate < 20:
                    return round(rate, 2)
        except Exception:
            pass

        return 2.8  # 한국 국고채 기본 fallback

    # ── 분기별 데이터 수집 ────────────────────────────────────────

    def fetch_quarterly_data(self, stock_code: str, num_years: int = 2) -> dict:
        """분기별 재무데이터 수집 + 누적→단독 분기 역산 (de-cumulation)

        DART 분기 보고서는 IS/CF가 누적 데이터이므로 단독 분기 값으로 역산 필요.
        BS 항목은 시점 데이터이므로 그대로 사용.

        현재 연도의 분기 보고서(Q1~Q3)도 조회하여 TTM에 반영.

        Args:
            stock_code: 종목코드
            num_years: 수집할 연도 수 (기본 2년 = 최근 8분기)

        Returns:
            {
                "quarters": ["2023Q1", ..., "2025Q1"],
                "quarterly_summary": {"2023Q1": {계정: 금액}, ...},
                "quarterly_bs": {"2023Q1": {계정: 금액}, ...},
                "quarterly_cf": {"2023Q1": {계정: 금액}, ...},
            }
        """
        latest_year = self.find_latest_available_year(stock_code)
        current_year = datetime.now().year
        quarters = []
        quarterly_summary = {}
        quarterly_bs = {}
        quarterly_cf = {}

        # 수집 연도 목록: 기존 num_years + 사업보고서 이후~현재 연도 (분기 보고서)
        # 예: latest_year=2024, current_year=2026 → [2023, 2024, 2025, 2026]
        years_to_fetch = list(range(latest_year - num_years + 1, latest_year + 1))
        for y in range(latest_year + 1, current_year + 1):
            years_to_fetch.append(y)

        for year in years_to_fetch:
            # 각 분기 보고서의 누적 데이터 수집
            cum_is = {}   # {"Q1": {계정: 누적값}, "Q2": {...}, ...}
            cum_cf = {}
            bs_data = {}  # BS는 시점 데이터 → 직접 사용

            # 사업보고서 미발행 연도는 Q1~Q3만 조회
            if year > latest_year:
                codes_to_try = [(c, q) for c, q in QUARTERLY_REPORT_CODES if q != "Q4"]
            else:
                codes_to_try = QUARTERLY_REPORT_CODES

            for reprt_code, q_label in codes_to_try:
                try:
                    df = self._fetch_finstate_all_cached(stock_code, year, reprt_code)
                    if df is None:
                        continue

                    q_is = {}
                    q_bs = {}
                    q_cf = {}

                    for _, row in df.iterrows():
                        nm = row.get("account_nm", "")
                        if not nm:
                            continue
                        sj = row.get("sj_div", "")
                        amt = self._parse_amount(row.get("thstrm_amount", ""))

                        if sj in ("IS", "CIS"):
                            q_is[nm] = amt
                            q_is[f"{sj}_{nm}"] = amt
                        elif sj == "BS":
                            q_bs[nm] = amt
                        elif sj == "CF":
                            q_cf[nm] = amt

                    if q_is:
                        cum_is[q_label] = q_is
                    if q_bs:
                        bs_data[q_label] = q_bs
                    if q_cf:
                        cum_cf[q_label] = q_cf

                except Exception as e:
                    print(f"  {year}년 {q_label} 분기 데이터 조회 실패: {e}")
                    continue

            # IS/CF 역산 (누적→단독)
            standalone_is = self._decumulate_flow_items(cum_is, year)
            standalone_cf = self._decumulate_flow_items(cum_cf, year)

            # BS는 시점 데이터 → 직접 사용
            for q_label, data in bs_data.items():
                qkey = f"{year}{q_label}"
                quarterly_bs[qkey] = data

            # 결과 취합
            for qkey, data in standalone_is.items():
                if data:  # 빈 dict 제외
                    quarterly_summary[qkey] = data
                    if qkey not in quarters:
                        quarters.append(qkey)

            for qkey, data in standalone_cf.items():
                if data:
                    quarterly_cf[qkey] = data

        # 시간순 정렬
        quarters.sort()

        return {
            "quarters": quarters,
            "quarterly_summary": quarterly_summary,
            "quarterly_bs": quarterly_bs,
            "quarterly_cf": quarterly_cf,
        }

    def _decumulate_flow_items(self, cumulative: dict, year: int) -> dict:
        """IS/CF 누적 데이터를 단독 분기 값으로 역산

        Args:
            cumulative: {"Q1": {계정: 누적값}, "Q2": {H1 누적}, "Q3": {Q1-Q3 누적}, "Q4": {연간}}
            year: 해당 연도

        Returns:
            {"2024Q1": {계정: 단독값}, "2024Q2": {...}, ...}
        """
        result = {}
        q1 = cumulative.get("Q1", {})
        h1 = cumulative.get("Q2", {})      # H1 = Q1+Q2 누적
        q3_cum = cumulative.get("Q3", {})   # Q1+Q2+Q3 누적
        annual = cumulative.get("Q4", {})   # 연간 = Q1+Q2+Q3+Q4

        # Q1: 단독 = Q1 누적 그대로
        if q1:
            result[f"{year}Q1"] = dict(q1)

        # Q2: 단독 = H1 누적 - Q1 누적
        if h1 and q1:
            all_keys = set(h1) | set(q1)
            result[f"{year}Q2"] = {k: h1.get(k, 0) - q1.get(k, 0) for k in all_keys}

        # Q3: 단독 = Q3 누적 - H1 누적
        if q3_cum and h1:
            all_keys = set(q3_cum) | set(h1)
            result[f"{year}Q3"] = {k: q3_cum.get(k, 0) - h1.get(k, 0) for k in all_keys}

        # Q4: 단독 = 연간 - Q3 누적
        if annual and q3_cum:
            all_keys = set(annual) | set(q3_cum)
            result[f"{year}Q4"] = {k: annual.get(k, 0) - q3_cum.get(k, 0) for k in all_keys}

        return result

    # ── 내부 헬퍼 ────────────────────────────────────────────────

    def _fetch_finstate_all_cached(self, stock_code: str, year: int,
                                    reprt_code: str = REPORT_TYPE_ANNUAL) -> pd.DataFrame | None:
        """finstate_all 캐시 (같은 연도+보고서유형 중복 API 호출 방지)"""
        key = (stock_code, year, reprt_code)
        if key in self._finstate_all_cache:
            return self._finstate_all_cache[key]

        df = self.dart.finstate_all(stock_code, year, reprt_code=reprt_code, fs_div="CFS")
        if df is None or (isinstance(df, pd.DataFrame) and df.empty):
            df = self.dart.finstate_all(stock_code, year, reprt_code=reprt_code, fs_div="OFS")
        if df is None or (isinstance(df, pd.DataFrame) and df.empty):
            df = None

        self._finstate_all_cache[key] = df
        time.sleep(0.3)
        return df

    @staticmethod
    def _parse_amount(value) -> int:
        """금액 문자열을 정수로 변환"""
        if not value or pd.isna(value):
            return 0
        try:
            return int(str(value).replace(",", ""))
        except (ValueError, TypeError):
            return 0
