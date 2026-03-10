"""Streamlit 웹앱 - 재무제표 분석 리포트 생성기

핸드폰 브라우저로 접속하여 사용 가능.
실행: streamlit run streamlit_app.py
같은 WiFi: http://<PC_IP>:8501
"""

import os
import sys
import tempfile
from datetime import datetime

import streamlit as st

# 프로젝트 경로 설정
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import (
    DEFAULT_REQUIRED_RETURN, DEFAULT_ANALYSIS_YEARS,
    DEFAULT_W_BUY, DEFAULT_W_FAIR,
    DEFAULT_INCLUDE_QUARTERLY,
    PDF_REPORT_COMBINED, PDF_REPORT_ANNUAL,
    PDF_REPORT_QUARTERLY, PDF_REPORT_RISK, PDF_REPORT_ALL,
)

st.set_page_config(
    page_title="재무분석 리포트",
    page_icon="📊",
    layout="wide",
)


# ── DART API 키 확인 ──────────────────────────────────────────

from config import DART_API_KEY
_dart_available = bool(DART_API_KEY)


# ── 캐시된 fetcher ──────────────────────────────────────────

@st.cache_resource
def get_kr_fetcher():
    from data_fetcher import DataFetcher
    return DataFetcher()


@st.cache_resource
def get_intl_fetcher():
    from international_fetcher import InternationalFetcher
    return InternationalFetcher()


def get_fetcher(market: str):
    if market == "INTL":
        return get_intl_fetcher()
    return get_kr_fetcher()


# ── 사이드바 ────────────────────────────────────────────────

st.sidebar.title("설정")

tab_choice = st.sidebar.radio("모드", ["리포트 생성", "스크리너"], horizontal=True)

if _dart_available:
    market = st.sidebar.selectbox("시장", ["KR", "INTL"],
                                   format_func=lambda x: "한국 (DART)" if x == "KR" else "해외 (yfinance)")
else:
    market = "INTL"
    st.sidebar.selectbox("시장", ["해외 (yfinance)"], disabled=True)
    st.sidebar.caption("DART API 키 미설정 → 해외 주식만 사용 가능")

if tab_choice == "리포트 생성":
    stock_input = st.sidebar.text_input("종목 검색", placeholder="애플 / Apple / AAPL")

    # 종목 검색 기능
    if stock_input and st.sidebar.button("검색", use_container_width=True):
        try:
            fetcher = get_fetcher(market)
            results = fetcher.search_stock(stock_input, limit=10)
            if results:
                st.session_state["search_results"] = results
            else:
                st.session_state["search_results"] = []
                st.sidebar.warning("검색 결과가 없습니다.")
        except Exception as e:
            st.session_state["search_results"] = []
            st.sidebar.warning(f"검색 실패: {e}")

    # 검색 결과가 있으면 선택 UI 표시
    if st.session_state.get("search_results"):
        results = st.session_state["search_results"]
        options = [f"{r['stock_code']} - {r['corp_name']}" for r in results]
        selected = st.sidebar.selectbox("종목 선택", options)
        stock_query = selected.split(" - ")[0]  # 코드만 추출
    else:
        stock_query = stock_input  # 직접 입력한 값 사용

    num_years = st.sidebar.slider("분석 연도 수", 3, 10, DEFAULT_ANALYSIS_YEARS)
    required_return = st.sidebar.number_input(
        "목표 수익률 (%)", min_value=1.0, max_value=30.0,
        value=DEFAULT_REQUIRED_RETURN, step=0.5,
    )

    roe_source = st.sidebar.selectbox("ROE 소스", ["consensus", "historical", "manual"],
                                       format_func=lambda x: {
                                           "consensus": "컨센서스 우선",
                                           "historical": "과거 가중평균",
                                           "manual": "직접 입력",
                                       }[x])
    manual_roe = None
    if roe_source == "manual":
        manual_roe = st.sidebar.number_input("ROE (%)", min_value=0.1, max_value=100.0, value=10.0, step=0.5)

    coe_source = st.sidebar.selectbox("COE 소스", ["manual", "capm"],
                                       format_func=lambda x: {
                                           "manual": f"수동 입력 ({required_return}%)",
                                           "capm": "CAPM 자동 계산",
                                       }[x])

    col_w1, col_w2 = st.sidebar.columns(2)
    w_buy = col_w1.number_input("매수W", min_value=0.0, max_value=1.0,
                                 value=DEFAULT_W_BUY, step=0.1)
    w_fair = col_w2.number_input("적정W", min_value=0.0, max_value=1.0,
                                  value=DEFAULT_W_FAIR, step=0.1)

    include_quarterly = st.sidebar.checkbox("분기/TTM 포함", value=DEFAULT_INCLUDE_QUARTERLY)
    include_trend = st.sidebar.checkbox("추이 분석 포함", value=True)

    report_type = st.sidebar.selectbox(
        "리포트 유형",
        [PDF_REPORT_COMBINED, PDF_REPORT_ANNUAL, PDF_REPORT_QUARTERLY,
         PDF_REPORT_RISK, PDF_REPORT_ALL],
        format_func=lambda x: {
            PDF_REPORT_COMBINED: "통합 리포트",
            PDF_REPORT_ANNUAL: "연도별 분석",
            PDF_REPORT_QUARTERLY: "분기별 분석",
            PDF_REPORT_RISK: "리스크 분석",
            PDF_REPORT_ALL: "3개 분리 (전부)",
        }[x],
    )

# ── 메인 영역 ───────────────────────────────────────────────

if tab_choice == "리포트 생성":
    st.title("재무제표 분석 리포트")

    if st.sidebar.button("리포트 생성", type="primary", use_container_width=True):
        if not stock_query:
            st.error("종목명 또는 코드를 입력하세요.")
            st.stop()

        from generate_report import collect_data, compute_derived_metrics
        from risk_analyzer import check_listing_risk, check_us_listing_risk
        from trend_analyzer import analyze_trend
        from pdf_report import PDFReportGenerator
        from pdf_annual_report import AnnualReportGenerator
        from pdf_quarterly_report import QuarterlyReportGenerator
        from pdf_risk_report import RiskReportGenerator

        fetcher = get_fetcher(market)
        is_intl = market == "INTL"

        progress = st.progress(0, text="종목 조회 중...")
        status = st.empty()

        try:
            # 1. 종목코드 해석
            stock_code = fetcher.resolve_stock_query(stock_query)
            progress.progress(5, text=f"종목: {stock_code}")

            # 2. 데이터 수집
            progress.progress(10, text="데이터 수집 중...")
            data = collect_data(fetcher, stock_code, num_years,
                                include_quarterly=include_quarterly)
            corp_name = data["company_info"]["corp_name"]
            progress.progress(50, text=f"{corp_name} 데이터 수집 완료")

            # 3. 파생 지표
            progress.progress(55, text="파생 지표 계산 중...")
            derived, srim = compute_derived_metrics(
                data, required_return,
                roe_source=roe_source,
                manual_roe=manual_roe,
                coe_source=coe_source,
                w_buy=w_buy, w_fair=w_fair,
            )

            # 4. 위험 분석
            progress.progress(65, text="위험 분석 중...")
            ttm_data = data.get("ttm")
            ttm_derived = derived.get("_ttm")
            if is_intl:
                risk_warnings = check_us_listing_risk(data, derived)
            else:
                risk_warnings = check_listing_risk(
                    data, derived,
                    ttm_data=ttm_data, ttm_derived=ttm_derived,
                )

            # 5. 추이 분석
            trend_result = None
            if include_trend:
                progress.progress(70, text="추이 분석 중...")
                try:
                    trend_result = analyze_trend(
                        data, derived,
                        ttm_data=ttm_data, ttm_derived=ttm_derived,
                        quarterly_derived=derived.get("_quarterly"),
                        quarterly_keys=derived.get("_quarterly_keys"),
                    )
                except Exception:
                    trend_result = None

            # 6. 리포트 데이터 조합
            report_data = {
                **data,
                "derived": derived,
                "srim": srim,
                "required_return": required_return,
                "risk_warnings": risk_warnings,
                "trend_analysis": trend_result,
            }

            # 7. 핵심 지표 표시
            price = data["stock_data"].get("price", 0)
            srim_price = srim.get("srim_price", 0)
            buy_price = srim.get("buy_price", 0)
            discount = round((srim_price - price) / srim_price * 100, 1) if srim_price else 0

            col1, col2, col3, col4 = st.columns(4)
            currency = data["company_info"].get("currency", "KRW")
            fmt = lambda v: f"${v:,.2f}" if currency != "KRW" else f"{v:,}원"

            col1.metric("현재가", fmt(price))
            col2.metric("S-RIM 적정가", fmt(srim_price))
            col3.metric("매수시작가", fmt(buy_price))
            col4.metric("할인율", f"{discount:+.1f}%",
                         delta=f"{'저평가' if discount > 0 else '고평가'}",
                         delta_color="normal" if discount > 0 else "inverse")

            # 추가 정보
            with st.expander("상세 정보", expanded=False):
                info_cols = st.columns(3)
                info_cols[0].write(f"**ROE 예측:** {srim.get('roe_forecast', 0):.1f}% ({srim.get('roe_source', '')})")
                info_cols[0].write(f"**COE:** {srim.get('coe_value', 0):.1f}% ({srim.get('coe_source', '')})")
                info_cols[1].write(f"**W(매수/적정):** {w_buy}/{w_fair}")
                info_cols[1].write(f"**Beta:** {srim.get('beta', 1.0):.2f}")
                if risk_warnings:
                    info_cols[2].warning(f"위험 항목 {len(risk_warnings)}건 발견")
                else:
                    info_cols[2].success("위험 항목 없음")
                if trend_result:
                    info_cols[2].write(
                        f"**추이:** {trend_result.get('situation_label', '?')} "
                        f"(신뢰도 {trend_result.get('confidence', 0) * 100:.0f}%)"
                    )

            # 8. PDF 생성
            progress.progress(80, text="PDF 생성 중...")

            generated = {}  # {display_name: (file_bytes, filename)}

            if report_type == PDF_REPORT_COMBINED:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    PDFReportGenerator(tmp.name).generate(report_data)
                    fname = f"{corp_name}_분석리포트.pdf"
                    generated[fname] = (open(tmp.name, "rb").read(), fname)

            elif report_type == PDF_REPORT_ANNUAL:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    AnnualReportGenerator(tmp.name).generate(report_data)
                    fname = f"{corp_name}_연도별분석.pdf"
                    generated[fname] = (open(tmp.name, "rb").read(), fname)

            elif report_type == PDF_REPORT_QUARTERLY:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    QuarterlyReportGenerator(tmp.name).generate(report_data)
                    fname = f"{corp_name}_분기별분석.pdf"
                    generated[fname] = (open(tmp.name, "rb").read(), fname)

            elif report_type == PDF_REPORT_RISK:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    RiskReportGenerator(tmp.name).generate(report_data)
                    fname = f"{corp_name}_리스크분석.pdf"
                    generated[fname] = (open(tmp.name, "rb").read(), fname)

            elif report_type == PDF_REPORT_ALL:
                for suffix, gen_cls in [
                    ("_연도별분석.pdf", AnnualReportGenerator),
                    ("_분기별분석.pdf", QuarterlyReportGenerator),
                    ("_리스크분석.pdf", RiskReportGenerator),
                ]:
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        gen_cls(tmp.name).generate(report_data)
                        fname = f"{corp_name}{suffix}"
                        generated[fname] = (open(tmp.name, "rb").read(), fname)

            progress.progress(100, text="완료!")

            # 다운로드 버튼
            st.subheader("PDF 다운로드")
            dl_cols = st.columns(len(generated))
            for i, (display, (pdf_bytes, filename)) in enumerate(generated.items()):
                dl_cols[i].download_button(
                    label=f"📥 {display}",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True,
                )

            # 세션에 결과 저장
            st.session_state["last_result"] = {
                "corp_name": corp_name,
                "srim": srim,
                "price": price,
                "discount": discount,
                "generated": generated,
            }

        except ValueError as e:
            progress.empty()
            st.error(f"오류: {e}")
        except Exception as e:
            progress.empty()
            st.error(f"처리 중 오류 발생: {e}")

    # 이전 결과가 있으면 다운로드 버튼 유지
    elif "last_result" in st.session_state:
        result = st.session_state["last_result"]
        st.info(f"이전 분석 결과: **{result['corp_name']}** (할인율 {result['discount']:+.1f}%)")
        gen = result.get("generated", {})
        if gen:
            dl_cols = st.columns(len(gen))
            for i, (display, (pdf_bytes, filename)) in enumerate(gen.items()):
                dl_cols[i].download_button(
                    label=f"📥 {display}",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True,
                )
    else:
        st.info("왼쪽 사이드바에서 종목을 입력하고 '리포트 생성'을 클릭하세요.")


# ── 스크리너 탭 ─────────────────────────────────────────────

elif tab_choice == "스크리너":
    st.title("S-RIM 저평가 종목 스크리너")

    # 스크리너 설정
    scr_col1, scr_col2 = st.sidebar.columns(2)
    scr_return = st.sidebar.number_input(
        "목표 수익률 (%)", min_value=1.0, max_value=30.0,
        value=DEFAULT_REQUIRED_RETURN, step=0.5, key="scr_return",
    )
    scr_roe = st.sidebar.selectbox("ROE 소스", ["consensus", "historical"],
                                    format_func=lambda x: "컨센서스 우선" if x == "consensus" else "과거 가중평균",
                                    key="scr_roe")
    scr_coe = st.sidebar.selectbox("COE 소스", ["manual", "capm"],
                                    format_func=lambda x: "수동 입력" if x == "manual" else "CAPM 자동",
                                    key="scr_coe")

    scr_w_col1, scr_w_col2 = st.sidebar.columns(2)
    scr_w_buy = scr_w_col1.number_input("매수W", min_value=0.0, max_value=1.0,
                                          value=DEFAULT_W_BUY, step=0.1, key="scr_w_buy")
    scr_w_fair = scr_w_col2.number_input("적정W", min_value=0.0, max_value=1.0,
                                           value=DEFAULT_W_FAIR, step=0.1, key="scr_w_fair")

    # 종목 소스
    st.subheader("종목 선택")

    if market == "KR":
        source_options = ["KOSPI 시총 상위", "KOSDAQ 시총 상위", "직접 입력"]
    else:
        source_options = ["S&P 500 시총 상위", "NASDAQ-100 시총 상위", "직접 입력"]

    source = st.radio("종목 소스", source_options, horizontal=True)

    if "직접 입력" in source:
        stock_input = st.text_area(
            "종목코드 (콤마 또는 줄바꿈 구분)",
            placeholder="005930, 000660, 051500\n또는\nAAPL\nMSFT\nGOOGL",
            height=100,
        )
    else:
        top_n = st.slider("상위 종목 수", 10, 100, 30, step=10)

    if st.button("스크리너 실행", type="primary", use_container_width=True):
        from screener import screen_stocks, get_market_top_stocks, get_us_market_top_stocks
        from pdf_report import ScreenerPDFGenerator

        fetcher = get_fetcher(market)
        is_intl = market == "INTL"

        progress = st.progress(0, text="종목 목록 준비 중...")

        # 종목 목록 결정
        try:
            if "직접 입력" in source:
                raw = stock_input.replace(",", "\n").strip()
                codes_raw = [c.strip() for c in raw.split("\n") if c.strip()]
                stock_codes = []
                for c in codes_raw:
                    try:
                        stock_codes.append(fetcher.resolve_stock_query(c))
                    except ValueError:
                        st.warning(f"종목 '{c}' 찾을 수 없음, 건너뜀")
            elif market == "KR":
                mkt = "KOSDAQ" if "KOSDAQ" in source else "KOSPI"
                stock_codes = get_market_top_stocks(fetcher, market=mkt, top_n=top_n)
            else:
                mkt = "NASDAQ100" if "NASDAQ" in source else "SP500"
                stock_codes = get_us_market_top_stocks(market=mkt, top_n=top_n)

            if not stock_codes:
                st.error("분석할 종목이 없습니다.")
                st.stop()

            progress.progress(10, text=f"{len(stock_codes)}개 종목 분석 시작...")

            # 스크리너 실행
            def _progress_cb(current, total, msg):
                pct = 10 + int(current / max(total, 1) * 80)
                progress.progress(min(pct, 90), text=msg)

            results = screen_stocks(
                fetcher, stock_codes,
                required_return=scr_return,
                roe_source=scr_roe,
                num_years=5,
                coe_source=scr_coe,
                progress_callback=_progress_cb,
                w_buy=scr_w_buy,
                w_fair=scr_w_fair,
            )

            progress.progress(95, text="결과 정리 중...")

            if not results:
                st.warning("분석 가능한 종목이 없습니다.")
                st.stop()

            # 결과 표시
            undervalued = [r for r in results if r.get("discount_pct", 0) > 0]
            overvalued = [r for r in results if r.get("discount_pct", 0) <= 0]

            st.subheader(f"분석 결과: {len(results)}개 종목")
            met_cols = st.columns(3)
            met_cols[0].metric("총 분석", f"{len(results)}개")
            met_cols[1].metric("저평가", f"{len(undervalued)}개")
            met_cols[2].metric("고평가", f"{len(overvalued)}개")

            # 테이블
            import pandas as pd

            currency = "USD" if is_intl else "KRW"

            def _make_df(items):
                rows = []
                for r in items:
                    name = r.get("corp_name", "")
                    kr = r.get("corp_name_kr", "")
                    if kr:
                        name = f"{name} ({kr})"
                    rows.append({
                        "종목명": name,
                        "코드": r.get("stock_code", ""),
                        "현재가": r.get("price", 0),
                        "적정가": r.get("srim_price", 0),
                        "매수시작가": r.get("buy_price", 0),
                        "할인율(%)": r.get("discount_pct", 0),
                        "ROE예측(%)": r.get("roe_forecast", 0),
                        "COE(%)": r.get("coe_value", 0),
                        "OPM(%)": r.get("opm", 0),
                        "PER": r.get("per", 0),
                        "PBR": r.get("pbr", 0),
                    })
                return pd.DataFrame(rows)

            if undervalued:
                st.markdown("**저평가 종목**")
                df_under = _make_df(undervalued)
                st.dataframe(df_under, use_container_width=True, hide_index=True)

            if overvalued:
                st.markdown("**고평가 종목**")
                df_over = _make_df(overvalued)
                st.dataframe(df_over, use_container_width=True, hide_index=True)

            # PDF 다운로드
            progress.progress(98, text="PDF 생성 중...")
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                ScreenerPDFGenerator(tmp.name).generate(
                    results,
                    roe_source=scr_roe,
                    required_return=scr_return,
                    market=source,
                    currency=currency,
                    coe_source=scr_coe,
                )
                pdf_bytes = open(tmp.name, "rb").read()
                fname = f"S-RIM_스크리너_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"

            progress.progress(100, text="완료!")

            st.download_button(
                label="📥 스크리너 PDF 다운로드",
                data=pdf_bytes,
                file_name=fname,
                mime="application/pdf",
                use_container_width=True,
            )

            st.session_state["screener_results"] = results

        except Exception as e:
            progress.empty()
            st.error(f"스크리너 실행 중 오류: {e}")

    elif "screener_results" in st.session_state:
        st.info("이전 스크리너 결과가 있습니다. 다시 실행하려면 '스크리너 실행'을 클릭하세요.")
