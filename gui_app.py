"""재무제표 분석 PDF 리포트 생성기 - GUI 버전 (tkinter)"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# 프로젝트 경로 설정
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import (DEFAULT_REQUIRED_RETURN, DEFAULT_ANALYSIS_YEARS,
                    DEFAULT_W_BUY, DEFAULT_W_FAIR,
                    DEFAULT_INCLUDE_QUARTERLY,
                    PDF_REPORT_COMBINED, PDF_REPORT_ANNUAL,
                    PDF_REPORT_QUARTERLY, PDF_REPORT_RISK, PDF_REPORT_ALL)
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
)
from pdf_report import PDFReportGenerator
from pdf_annual_report import AnnualReportGenerator
from pdf_quarterly_report import QuarterlyReportGenerator
from pdf_risk_report import RiskReportGenerator
from risk_analyzer import check_listing_risk, check_us_listing_risk


class ReportApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("재무제표 분석 리포트 생성기")
        self.root.geometry("720x660")
        self.root.resizable(False, False)

        # macOS 스타일
        self.root.configure(bg="#F5F5F5")
        style = ttk.Style()
        style.theme_use("aqua" if sys.platform == "darwin" else "clam")

        self._is_running = False
        self._screener_results = None  # 스크리너 결과 저장
        self._screener_meta = {}       # 스크리너 메타 정보
        self._build_ui()

    def _build_ui(self):
        # ── 메인 프레임 ──
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        # 제목
        title_label = ttk.Label(
            main,
            text="재무제표 분석 PDF 리포트 생성기",
            font=("Helvetica", 16, "bold"),
        )
        title_label.pack(pady=(0, 10))

        # ── 탭 ──
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill="both", expand=True)

        # 탭 1: 리포트 생성
        report_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(report_tab, text=" 리포트 생성 ")

        # 탭 2: S-RIM 스크리너
        screener_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(screener_tab, text=" S-RIM 스크리너 ")

        self._build_report_tab(report_tab)
        self._build_screener_tab(screener_tab)

    def _build_report_tab(self, parent):
        """리포트 생성 탭 UI"""
        # ── 입력 영역 ──
        input_frame = ttk.LabelFrame(parent, text="설정", padding=15)
        input_frame.pack(fill="x", pady=(0, 10))

        # 시장 선택
        row0 = ttk.Frame(input_frame)
        row0.pack(fill="x", pady=3)
        ttk.Label(row0, text="시장:", width=14).pack(side="left")
        self.market_var = tk.StringVar(value="KR")
        ttk.Radiobutton(row0, text="한국 (DART)", variable=self.market_var,
                        value="KR", command=self._on_market_change).pack(side="left")
        ttk.Radiobutton(row0, text="해외 (yfinance)", variable=self.market_var,
                        value="INTL", command=self._on_market_change).pack(side="left", padx=(10, 0))

        # 종목 검색
        row1 = ttk.Frame(input_frame)
        row1.pack(fill="x", pady=3)
        ttk.Label(row1, text="종목:", width=14).pack(side="left")
        self.stock_code_var = tk.StringVar(value="051500")
        entry = ttk.Entry(row1, textvariable=self.stock_code_var, width=20)
        entry.pack(side="left", padx=(0, 5))
        ttk.Button(row1, text="검색", command=self._on_search_stock).pack(side="left", padx=(0, 5))
        self.stock_name_label = ttk.Label(row1, text="종목명 또는 코드 입력 후 검색", foreground="gray")
        self.stock_name_label.pack(side="left")

        # API 키
        self.api_key_frame = ttk.Frame(input_frame)
        self.api_key_frame.pack(fill="x", pady=3)
        ttk.Label(self.api_key_frame, text="DART API 키:", width=14).pack(side="left")
        self.api_key_var = tk.StringVar(value=self._load_api_key())
        self.api_entry = ttk.Entry(self.api_key_frame, textvariable=self.api_key_var, width=45, show="*")
        self.api_entry.pack(side="left", padx=(0, 5))
        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.api_key_frame, text="보기", variable=self.show_key_var,
            command=lambda: self.api_entry.configure(show="" if self.show_key_var.get() else "*"),
        ).pack(side="left")

        # COE 소스 선택
        row_coe = ttk.Frame(input_frame)
        row_coe.pack(fill="x", pady=3)
        ttk.Label(row_coe, text="COE 소스:", width=14).pack(side="left")
        self.coe_source_var = tk.StringVar(value="manual")
        coe_combo = ttk.Combobox(
            row_coe, textvariable=self.coe_source_var,
            values=["manual", "capm"],
            width=10, state="readonly",
        )
        coe_combo.pack(side="left", padx=(0, 5))
        self.coe_info_label = ttk.Label(row_coe, text="수동입력 / CAPM 자동계산", foreground="gray")
        self.coe_info_label.pack(side="left")
        coe_combo.bind("<<ComboboxSelected>>", self._on_coe_source_change)

        # 수익률
        row3 = ttk.Frame(input_frame)
        row3.pack(fill="x", pady=3)
        ttk.Label(row3, text="목표 수익률(%):", width=14).pack(side="left")
        self.return_rate_var = tk.StringVar(value=str(DEFAULT_REQUIRED_RETURN))
        self.return_rate_entry = ttk.Entry(row3, textvariable=self.return_rate_var, width=8)
        self.return_rate_entry.pack(side="left", padx=(0, 10))
        self.return_rate_desc = ttk.Label(row3, text="S-RIM 밸류에이션 계산에 사용 (COE=수동 시)", foreground="gray")
        self.return_rate_desc.pack(side="left")

        # ROE 소스 선택
        row3b = ttk.Frame(input_frame)
        row3b.pack(fill="x", pady=3)
        ttk.Label(row3b, text="ROE 소스:", width=14).pack(side="left")
        self.roe_source_var = tk.StringVar(value="consensus")
        roe_combo = ttk.Combobox(
            row3b, textvariable=self.roe_source_var,
            values=["consensus", "historical", "manual"],
            width=12, state="readonly",
        )
        roe_combo.pack(side="left", padx=(0, 5))
        self.manual_roe_var = tk.StringVar(value="")
        self.manual_roe_entry = ttk.Entry(row3b, textvariable=self.manual_roe_var, width=6)
        self.manual_roe_entry.pack(side="left", padx=(0, 5))
        self.manual_roe_entry.configure(state="disabled")
        ttk.Label(row3b, text="컨센서스우선 / 과거가중평균 / 직접입력(%)", foreground="gray").pack(side="left")
        roe_combo.bind("<<ComboboxSelected>>", self._on_roe_source_change)

        # 분석 연도 수
        row4 = ttk.Frame(input_frame)
        row4.pack(fill="x", pady=3)
        ttk.Label(row4, text="분석 연도 수:", width=14).pack(side="left")
        self.years_var = tk.StringVar(value=str(DEFAULT_ANALYSIS_YEARS))
        years_combo = ttk.Combobox(
            row4, textvariable=self.years_var, values=["3", "4", "5", "6", "7"],
            width=5, state="readonly",
        )
        years_combo.pack(side="left", padx=(0, 10))
        ttk.Label(row4, text="최근 N개년 데이터 분석", foreground="gray").pack(side="left")

        # 초과이익 지속계수(W)
        row4w = ttk.Frame(input_frame)
        row4w.pack(fill="x", pady=3)
        ttk.Label(row4w, text="지속계수(W):", width=14).pack(side="left")
        ttk.Label(row4w, text="매수W:", foreground="gray").pack(side="left")
        self.w_buy_var = tk.StringVar(value=str(DEFAULT_W_BUY))
        ttk.Entry(row4w, textvariable=self.w_buy_var, width=5).pack(side="left", padx=(0, 8))
        ttk.Label(row4w, text="적정W:", foreground="gray").pack(side="left")
        self.w_fair_var = tk.StringVar(value=str(DEFAULT_W_FAIR))
        ttk.Entry(row4w, textvariable=self.w_fair_var, width=5).pack(side="left", padx=(0, 8))
        ttk.Label(row4w, text="(0.0~1.0, 기본: 0.5/1.0)", foreground="gray").pack(side="left")

        # TTM/분기 포함
        row4q = ttk.Frame(input_frame)
        row4q.pack(fill="x", pady=3)
        ttk.Label(row4q, text="TTM/분기 포함:", width=14).pack(side="left")
        self.include_quarterly_var = tk.BooleanVar(value=DEFAULT_INCLUDE_QUARTERLY)
        ttk.Checkbutton(
            row4q, text="TTM 열 + 분기별 상세 페이지 추가",
            variable=self.include_quarterly_var,
        ).pack(side="left")
        ttk.Label(row4q, text="(최근 8분기, API 호출 증가)", foreground="gray").pack(side="left", padx=(5, 0))

        # 추이 분석 포함
        row4t = ttk.Frame(input_frame)
        row4t.pack(fill="x", pady=3)
        ttk.Label(row4t, text="추이 분석:", width=14).pack(side="left")
        self.include_trend_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            row4t, text="재무 추이 분석 페이지 추가",
            variable=self.include_trend_var,
        ).pack(side="left")
        ttk.Label(row4t, text="(성장성/수익성/안정성/현금흐름 종합 진단)", foreground="gray").pack(side="left", padx=(5, 0))

        # 리포트 유형 선택
        row_rt = ttk.Frame(input_frame)
        row_rt.pack(fill="x", pady=3)
        ttk.Label(row_rt, text="리포트 유형:", width=14).pack(side="left")
        self.report_type_var = tk.StringVar(value=PDF_REPORT_COMBINED)
        for val, label in [
            (PDF_REPORT_COMBINED, "통합"),
            (PDF_REPORT_ANNUAL, "연도별"),
            (PDF_REPORT_QUARTERLY, "분기별"),
            (PDF_REPORT_RISK, "리스크"),
            (PDF_REPORT_ALL, "전체(3개)"),
        ]:
            ttk.Radiobutton(
                row_rt, text=label, variable=self.report_type_var, value=val,
            ).pack(side="left", padx=(0, 6))

        # 출력 경로
        row5 = ttk.Frame(input_frame)
        row5.pack(fill="x", pady=3)
        ttk.Label(row5, text="저장 위치:", width=14).pack(side="left")
        self.output_var = tk.StringVar(value="")
        ttk.Entry(row5, textvariable=self.output_var, width=35).pack(side="left", padx=(0, 5))
        ttk.Button(row5, text="찾아보기...", command=self._browse_output).pack(side="left")

        # ── 실행 버튼 ──
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", pady=10)

        self.run_btn = ttk.Button(
            btn_frame, text="리포트 생성", command=self._on_generate,
        )
        self.run_btn.pack(side="left", padx=(0, 10))

        self.open_btn = ttk.Button(
            btn_frame, text="파일 열기", command=self._open_file, state="disabled",
        )
        self.open_btn.pack(side="left")

        # ── 프로그레스 ──
        self.progress = ttk.Progressbar(parent, mode="determinate", length=560)
        self.progress.pack(fill="x", pady=(0, 5))

        self.status_var = tk.StringVar(value="종목코드를 입력하고 '리포트 생성'을 누르세요.")
        status_label = ttk.Label(parent, textvariable=self.status_var, foreground="gray")
        status_label.pack(anchor="w")

        # ── 로그 영역 ──
        log_frame = ttk.LabelFrame(parent, text="진행 로그", padding=5)
        log_frame.pack(fill="both", expand=True, pady=(5, 0))

        self.log_text = tk.Text(log_frame, height=8, font=("Menlo", 10), state="disabled", bg="#1E1E1E", fg="#D4D4D4")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        self._last_output_path = None

    def _build_screener_tab(self, parent):
        """S-RIM 스크리너 탭 UI"""
        # 설정 영역
        setting_frame = ttk.LabelFrame(parent, text="스크리너 설정", padding=10)
        setting_frame.pack(fill="x", pady=(0, 10))

        # 종목 입력 방식 선택
        row1 = ttk.Frame(setting_frame)
        row1.pack(fill="x", pady=3)
        ttk.Label(row1, text="종목 입력:", width=14).pack(side="left")
        self.scr_mode_var = tk.StringVar(value="manual")
        ttk.Radiobutton(row1, text="직접 입력", variable=self.scr_mode_var,
                        value="manual", command=self._on_scr_mode_change).pack(side="left")
        ttk.Radiobutton(row1, text="KOSPI", variable=self.scr_mode_var,
                        value="kospi_top", command=self._on_scr_mode_change).pack(side="left", padx=(10, 0))
        ttk.Radiobutton(row1, text="KOSDAQ", variable=self.scr_mode_var,
                        value="kosdaq_top", command=self._on_scr_mode_change).pack(side="left", padx=(10, 0))
        ttk.Radiobutton(row1, text="S&P 500", variable=self.scr_mode_var,
                        value="sp500_top", command=self._on_scr_mode_change).pack(side="left", padx=(10, 0))
        ttk.Radiobutton(row1, text="NASDAQ", variable=self.scr_mode_var,
                        value="nasdaq_top", command=self._on_scr_mode_change).pack(side="left", padx=(10, 0))

        # 상위 N개
        row1b = ttk.Frame(setting_frame)
        row1b.pack(fill="x", pady=3)
        ttk.Label(row1b, text="상위 N개:", width=14).pack(side="left")
        self.scr_topn_var = tk.StringVar(value="20")
        self.scr_topn_combo = ttk.Combobox(
            row1b, textvariable=self.scr_topn_var,
            values=["10", "20", "30", "50", "100"], width=5,
        )
        self.scr_topn_combo.pack(side="left", padx=(0, 10))
        ttk.Label(row1b, text="(시총 상위 모드 — 직접 숫자 입력 가능)", foreground="gray").pack(side="left")

        # 직접 입력 (종목코드 쉼표 구분)
        row2 = ttk.Frame(setting_frame)
        row2.pack(fill="x", pady=3)
        ttk.Label(row2, text="종목 목록:", width=14).pack(side="left", anchor="n")
        self.scr_stocks_text = tk.Text(row2, height=3, width=50, font=("Menlo", 10))
        self.scr_stocks_text.pack(side="left", fill="x", expand=True)
        self.scr_stocks_text.insert("1.0", "005930, 000660, 035720, 035420, 051500")

        # ROE 소스 / 수익률
        row3 = ttk.Frame(setting_frame)
        row3.pack(fill="x", pady=3)
        ttk.Label(row3, text="ROE 소스:", width=14).pack(side="left")
        self.scr_roe_var = tk.StringVar(value="consensus")
        ttk.Combobox(
            row3, textvariable=self.scr_roe_var,
            values=["consensus", "historical"], width=12, state="readonly",
        ).pack(side="left", padx=(0, 15))
        ttk.Label(row3, text="목표 수익률(%):", width=12).pack(side="left")
        self.scr_rr_var = tk.StringVar(value="8")
        self.scr_rr_entry = ttk.Entry(row3, textvariable=self.scr_rr_var, width=5)
        self.scr_rr_entry.pack(side="left")

        # COE 소스
        row3b = ttk.Frame(setting_frame)
        row3b.pack(fill="x", pady=3)
        ttk.Label(row3b, text="COE 소스:", width=14).pack(side="left")
        self.scr_coe_source_var = tk.StringVar(value="manual")
        scr_coe_combo = ttk.Combobox(
            row3b, textvariable=self.scr_coe_source_var,
            values=["manual", "capm"], width=12, state="readonly",
        )
        scr_coe_combo.pack(side="left", padx=(0, 10))
        scr_coe_combo.bind("<<ComboboxSelected>>", self._on_scr_coe_source_change)
        self.scr_coe_desc_label = ttk.Label(
            row3b, text="수동입력: 위 목표 수익률을 COE로 사용", foreground="gray",
        )
        self.scr_coe_desc_label.pack(side="left")

        # 초과이익 지속계수(W) — 스크리너
        row3w = ttk.Frame(setting_frame)
        row3w.pack(fill="x", pady=3)
        ttk.Label(row3w, text="지속계수(W):", width=14).pack(side="left")
        ttk.Label(row3w, text="매수W:", foreground="gray").pack(side="left")
        self.scr_w_buy_var = tk.StringVar(value=str(DEFAULT_W_BUY))
        ttk.Entry(row3w, textvariable=self.scr_w_buy_var, width=5).pack(side="left", padx=(0, 8))
        ttk.Label(row3w, text="적정W:", foreground="gray").pack(side="left")
        self.scr_w_fair_var = tk.StringVar(value=str(DEFAULT_W_FAIR))
        ttk.Entry(row3w, textvariable=self.scr_w_fair_var, width=5).pack(side="left", padx=(0, 8))
        ttk.Label(row3w, text="(0.0~1.0, 기본: 0.5/1.0)", foreground="gray").pack(side="left")

        # 실행 버튼
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", pady=5)

        self.scr_run_btn = ttk.Button(
            btn_frame, text="스크리너 실행", command=self._on_screener_run,
        )
        self.scr_run_btn.pack(side="left", padx=(0, 10))

        self.scr_pdf_btn = ttk.Button(
            btn_frame, text="PDF 저장", command=self._on_screener_pdf, state="disabled",
        )
        self.scr_pdf_btn.pack(side="left", padx=(0, 10))

        self.scr_progress = ttk.Progressbar(btn_frame, mode="determinate", length=250)
        self.scr_progress.pack(side="left", fill="x", expand=True)

        self.scr_status_var = tk.StringVar(value="종목을 입력하고 '스크리너 실행'을 누르세요.")
        ttk.Label(parent, textvariable=self.scr_status_var, foreground="gray").pack(anchor="w")

        # 결과 영역
        result_frame = ttk.LabelFrame(parent, text="분석 결과", padding=5)
        result_frame.pack(fill="both", expand=True, pady=(5, 0))

        self.scr_result_text = tk.Text(
            result_frame, font=("Menlo", 9), state="disabled",
            bg="#1E1E1E", fg="#D4D4D4", wrap="none",
        )
        yscroll = ttk.Scrollbar(result_frame, orient="vertical", command=self.scr_result_text.yview)
        xscroll = ttk.Scrollbar(result_frame, orient="horizontal", command=self.scr_result_text.xview)
        self.scr_result_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.pack(side="right", fill="y")
        xscroll.pack(side="bottom", fill="x")
        self.scr_result_text.pack(fill="both", expand=True)

    def _on_market_change(self):
        """리포트 탭 시장 선택 변경"""
        market = self.market_var.get()
        if market == "INTL":
            self.stock_code_var.set("AAPL")
            self.stock_name_label.configure(text="티커 심볼 입력 (예: AAPL, MSFT)", foreground="gray")
            # API 키 프레임 비활성화
            for child in self.api_key_frame.winfo_children():
                try:
                    child.configure(state="disabled")
                except Exception:
                    pass
        else:
            self.stock_code_var.set("051500")
            self.stock_name_label.configure(text="종목명 또는 코드 입력 후 검색", foreground="gray")
            for child in self.api_key_frame.winfo_children():
                try:
                    child.configure(state="normal")
                except Exception:
                    pass
            self.api_entry.configure(show="" if self.show_key_var.get() else "*")

    def _on_scr_mode_change(self):
        mode = self.scr_mode_var.get()
        if mode == "manual":
            self.scr_stocks_text.configure(state="normal")
            self.scr_topn_combo.configure(state="disabled")
        else:
            self.scr_stocks_text.configure(state="disabled")
            self.scr_topn_combo.configure(state="readonly")

    def _on_scr_coe_source_change(self, _event=None):
        if self.scr_coe_source_var.get() == "capm":
            self.scr_rr_entry.configure(state="disabled")
            self.scr_coe_desc_label.configure(
                text="CAPM: Rf + Beta × MRP (종목별 자동계산)")
        else:
            self.scr_rr_entry.configure(state="normal")
            self.scr_coe_desc_label.configure(
                text="수동입력: 위 목표 수익률을 COE로 사용")

    def _is_intl_screener_mode(self, mode: str) -> bool:
        return mode in ("sp500_top", "nasdaq_top")

    def _on_screener_run(self):
        if self._is_running:
            return

        mode = self.scr_mode_var.get()
        is_intl = self._is_intl_screener_mode(mode)

        # 해외 모드가 아닐 때만 API 키 필요
        api_key = self.api_key_var.get().strip()
        if not is_intl and mode != "manual":
            if not api_key:
                messagebox.showwarning("입력 오류", "DART API 키가 필요합니다.\n'리포트 생성' 탭에서 입력하세요.")
                return

        try:
            rr = float(self.scr_rr_var.get())
        except ValueError:
            messagebox.showwarning("입력 오류", "수익률은 숫자로 입력하세요.")
            return

        # 상위 N개 유효성 검사
        if mode != "manual":
            try:
                top_n_val = int(self.scr_topn_var.get())
                if top_n_val < 1 or top_n_val > 500:
                    messagebox.showwarning("입력 오류", "상위 N개는 1~500 사이의 숫자를 입력하세요.")
                    return
            except ValueError:
                messagebox.showwarning("입력 오류", "상위 N개는 숫자로 입력하세요.")
                return

        mode = self.scr_mode_var.get()
        roe_src = self.scr_roe_var.get()
        coe_src = self.scr_coe_source_var.get()

        # W값 파싱 (스크리너)
        try:
            scr_w_buy = float(self.scr_w_buy_var.get())
            scr_w_fair = float(self.scr_w_fair_var.get())
            if not (0 <= scr_w_buy <= 1 and 0 <= scr_w_fair <= 1):
                raise ValueError("W값은 0.0~1.0 범위여야 합니다.")
            if not (scr_w_buy <= scr_w_fair):
                raise ValueError("W값은 매수W ≤ 적정W 순이어야 합니다.")
        except ValueError as e:
            messagebox.showwarning("입력 오류", f"지속계수(W) 오류: {e}")
            return

        self._is_running = True
        self.scr_run_btn.configure(state="disabled")
        self.scr_result_text.configure(state="normal")
        self.scr_result_text.delete("1.0", "end")
        self.scr_result_text.configure(state="disabled")
        self.scr_progress.configure(value=0)

        thread = threading.Thread(
            target=self._run_screener,
            args=(api_key, mode, roe_src, rr, coe_src, scr_w_buy, scr_w_fair),
            daemon=True,
        )
        thread.start()

    def _run_screener(self, api_key: str, mode: str, roe_src: str, rr: float,
                      coe_src: str = "manual",
                      w_buy: float = DEFAULT_W_BUY, w_fair: float = DEFAULT_W_FAIR):
        from screener import screen_stocks, get_market_top_stocks, get_us_market_top_stocks, format_screener_results

        is_intl = self._is_intl_screener_mode(mode)

        try:
            # fetcher 선택
            if is_intl:
                from international_fetcher import InternationalFetcher
                fetcher = InternationalFetcher()
            else:
                fetcher = DataFetcher(api_key=api_key)

            # 종목 리스트 결정
            if mode == "manual":
                text = self.scr_stocks_text.get("1.0", "end").strip()
                raw_codes = [s.strip() for s in text.replace("\n", ",").split(",") if s.strip()]
                # 수동 입력 시 해외 티커 감지: 알파벳 포함 여부
                has_alpha = any(not q.isdigit() for q in raw_codes if q)
                if has_alpha:
                    from international_fetcher import InternationalFetcher
                    fetcher = InternationalFetcher()
                    is_intl = True
                elif not api_key:
                    self._scr_status("DART API 키가 필요합니다.")
                    return
                stock_codes = []
                for q in raw_codes:
                    try:
                        code = fetcher.resolve_stock_query(q)
                        stock_codes.append(code)
                    except ValueError:
                        self._scr_log(f"  '{q}' 종목을 찾을 수 없습니다. 건너뜁니다.")
            elif mode == "kospi_top":
                top_n = int(self.scr_topn_var.get())
                self._scr_status(f"KOSPI 시총 상위 {top_n}개 종목 조회 중...")
                stock_codes = get_market_top_stocks(fetcher, "KOSPI", top_n)
            elif mode == "kosdaq_top":
                top_n = int(self.scr_topn_var.get())
                self._scr_status(f"KOSDAQ 시총 상위 {top_n}개 종목 조회 중...")
                stock_codes = get_market_top_stocks(fetcher, "KOSDAQ", top_n)
            elif mode == "sp500_top":
                top_n = int(self.scr_topn_var.get())
                self._scr_status(f"S&P 500 상위 {top_n}개 종목 조회 중...")
                stock_codes = get_us_market_top_stocks("SP500", top_n)
            elif mode == "nasdaq_top":
                top_n = int(self.scr_topn_var.get())
                self._scr_status(f"NASDAQ-100 상위 {top_n}개 종목 조회 중...")
                stock_codes = get_us_market_top_stocks("NASDAQ100", top_n)

            if not stock_codes:
                self._scr_status("분석할 종목이 없습니다.")
                return

            self._scr_log(f"총 {len(stock_codes)}개 종목 분석 시작...\n")

            def progress_cb(current, total, msg):
                pct = int(current / max(total, 1) * 100)
                self.root.after(0, lambda: self.scr_progress.configure(value=pct))
                self._scr_status(f"[{current}/{total}] {msg}")
                self._scr_log(f"  [{current+1}/{total}] {msg}")

            results = screen_stocks(
                fetcher, stock_codes,
                required_return=rr,
                roe_source=roe_src,
                coe_source=coe_src,
                progress_callback=progress_cb,
                w_buy=w_buy, w_fair=w_fair,
            )

            # 결과 출력
            currency = "USD" if is_intl else "KRW"
            output = format_screener_results(results, currency=currency)
            self._scr_log(f"\n{output}")

            success = len(results)
            undervalued = len([r for r in results if r["discount_pct"] > 0])
            self._scr_status(f"완료! {success}개 분석 | 저평가 {undervalued}개 발견")
            self.root.after(0, lambda: self.scr_progress.configure(value=100))

            # 결과 저장 (PDF 저장용)
            self._screener_results = results
            market_label = {
                "kospi_top": "KOSPI", "kosdaq_top": "KOSDAQ",
                "sp500_top": "S&P500", "nasdaq_top": "NASDAQ100",
            }.get(mode, "해외직접입력" if is_intl else "직접입력")
            self._screener_meta = {
                "roe_source": roe_src,
                "required_return": rr,
                "market": market_label,
                "currency": currency,
                "coe_source": coe_src,
                "w_buy": w_buy,
                "w_fair": w_fair,
            }
            if results:
                self.root.after(0, lambda: self.scr_pdf_btn.configure(state="normal"))

        except Exception as e:
            self._scr_status(f"오류: {e}")
            self._scr_log(f"\n오류: {e}")
        finally:
            self._is_running = False
            self.root.after(0, lambda: self.scr_run_btn.configure(state="normal"))

    def _on_screener_pdf(self):
        """스크리너 결과를 PDF로 저장"""
        if not self._screener_results:
            messagebox.showwarning("저장 오류", "먼저 스크리너를 실행하세요.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf")],
            initialfile=f"SRIM_스크리너_{self._screener_meta.get('market', '')}.pdf",
        )
        if not path:
            return

        try:
            from pdf_report import ScreenerPDFGenerator
            gen = ScreenerPDFGenerator(path)
            gen.generate(
                self._screener_results,
                roe_source=self._screener_meta.get("roe_source", "consensus"),
                required_return=self._screener_meta.get("required_return", 8.0),
                market=self._screener_meta.get("market", ""),
                currency=self._screener_meta.get("currency", "KRW"),
                coe_source=self._screener_meta.get("coe_source", "manual"),
            )
            messagebox.showinfo("완료", f"PDF가 저장되었습니다.\n\n{path}")
            # 자동으로 파일 열기
            if sys.platform == "darwin":
                os.system(f'open "{path}"')
            elif sys.platform == "win32":
                os.startfile(path)
        except Exception as e:
            messagebox.showerror("오류", f"PDF 저장 실패: {e}")

    def _scr_log(self, msg: str):
        def _append():
            self.scr_result_text.configure(state="normal")
            self.scr_result_text.insert("end", msg + "\n")
            self.scr_result_text.see("end")
            self.scr_result_text.configure(state="disabled")
        self.root.after(0, _append)

    def _scr_status(self, msg: str):
        self.root.after(0, lambda: self.scr_status_var.set(msg))

    def _load_api_key(self) -> str:
        """환경변수 또는 .env에서 API 키 로드"""
        from dotenv import load_dotenv
        load_dotenv()
        return os.getenv("DART_API_KEY", "")

    def _on_search_stock(self):
        """종목 검색 버튼 클릭"""
        query = self.stock_code_var.get().strip()
        if not query:
            messagebox.showwarning("입력 오류", "종목명 또는 종목코드를 입력하세요.")
            return

        is_intl = self.market_var.get() == "INTL"

        # 이미 6자리 숫자면 검색 불필요 (한국 시장)
        if not is_intl and len(query) == 6 and query.isdigit():
            self.stock_name_label.configure(text="종목코드 직접 입력됨", foreground="green")
            return

        if not is_intl:
            api_key = self.api_key_var.get().strip()
            if not api_key:
                messagebox.showwarning("입력 오류", "검색하려면 DART API 키가 필요합니다.")
                return

        # UI 피드백: "검색 중..." 표시
        self.stock_name_label.configure(text="검색 중...", foreground="gray")

        # 백그라운드 스레드에서 검색 실행 (OpenDartReader 초기화 블로킹 방지)
        thread = threading.Thread(
            target=self._run_search_stock,
            args=(query, is_intl),
            daemon=True,
        )
        thread.start()

    def _run_search_stock(self, query: str, is_intl: bool):
        """백그라운드에서 종목 검색 수행"""
        try:
            if is_intl:
                from international_fetcher import InternationalFetcher
                fetcher = InternationalFetcher()
                results = fetcher.search_stock(query)
            else:
                api_key = self.api_key_var.get().strip()
                fetcher = DataFetcher(api_key=api_key)
                results = fetcher.search_stock(query)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("오류", f"검색 중 오류: {e}"))
            self.root.after(0, lambda: self.stock_name_label.configure(
                text="", foreground="black"))
            return

        # UI 업데이트는 메인 스레드에서 수행
        self.root.after(0, lambda: self._handle_search_results(query, results))

    def _handle_search_results(self, query: str, results: list):
        """메인 스레드에서 검색 결과 처리"""
        if not results:
            self.stock_name_label.configure(text="", foreground="black")
            messagebox.showinfo("검색 결과", f"'{query}'에 해당하는 상장기업이 없습니다.")
            return

        if len(results) == 1:
            # 1건이면 자동 선택
            r = results[0]
            self.stock_code_var.set(r["stock_code"])
            self.stock_name_label.configure(
                text=f"{r['corp_name']} ({r['stock_code']})", foreground="green"
            )
            return

        # 여러 건 → 선택 다이얼로그
        self._show_search_results(results)

    def _show_search_results(self, results: list[dict]):
        """검색 결과 선택 다이얼로그"""
        dialog = tk.Toplevel(self.root)
        dialog.title("종목 검색 결과")
        dialog.geometry("350x300")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="종목을 선택하세요:", padding=10).pack(anchor="w")

        # 리스트박스
        frame = ttk.Frame(dialog, padding=(10, 0, 10, 10))
        frame.pack(fill="both", expand=True)

        listbox = tk.Listbox(frame, font=("Menlo", 11))
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        listbox.pack(fill="both", expand=True)

        for r in results:
            listbox.insert("end", f"{r['stock_code']}  {r['corp_name']}")

        def on_select(_event=None):
            sel = listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            r = results[idx]
            self.stock_code_var.set(r["stock_code"])
            self.stock_name_label.configure(
                text=f"{r['corp_name']} ({r['stock_code']})", foreground="green"
            )
            dialog.destroy()

        listbox.bind("<Double-1>", on_select)

        btn_frame = ttk.Frame(dialog, padding=10)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="선택", command=on_select).pack(side="left", padx=(0, 5))
        ttk.Button(btn_frame, text="취소", command=dialog.destroy).pack(side="left")

    def _on_roe_source_change(self, _event=None):
        if self.roe_source_var.get() == "manual":
            self.manual_roe_entry.configure(state="normal")
        else:
            self.manual_roe_entry.configure(state="disabled")

    def _on_coe_source_change(self, _event=None):
        if self.coe_source_var.get() == "capm":
            self.return_rate_entry.configure(state="disabled")
            self.coe_info_label.configure(text="CAPM: Rf + Beta x MRP (자동계산)")
        else:
            self.return_rate_entry.configure(state="normal")
            self.coe_info_label.configure(text="수동입력 / CAPM 자동계산")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf")],
            initialfile="분석리포트.pdf",
        )
        if path:
            self.output_var.set(path)

    def _log(self, msg: str):
        """로그 메시지 추가 (스레드 안전)"""
        def _append():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, _append)

    def _set_status(self, msg: str):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _set_progress(self, value: int):
        self.root.after(0, lambda: self.progress.configure(value=value))

    def _on_generate(self):
        if self._is_running:
            return

        stock_query = self.stock_code_var.get().strip()
        api_key = self.api_key_var.get().strip()
        is_intl = self.market_var.get() == "INTL"

        if not stock_query:
            messagebox.showwarning("입력 오류", "종목명 또는 종목코드를 입력하세요.")
            return
        if not is_intl and not api_key:
            messagebox.showwarning("입력 오류", "DART API 키를 입력하세요.\n발급: https://opendart.fss.or.kr")
            return

        try:
            return_rate = float(self.return_rate_var.get())
        except ValueError:
            messagebox.showwarning("입력 오류", "수익률은 숫자로 입력하세요.")
            return

        num_years = int(self.years_var.get())

        # W값 파싱
        try:
            w_buy = float(self.w_buy_var.get())
            w_fair = float(self.w_fair_var.get())
            if not (0 <= w_buy <= 1 and 0 <= w_fair <= 1):
                raise ValueError("W값은 0.0~1.0 범위여야 합니다.")
            if not (w_buy <= w_fair):
                raise ValueError("W값은 매수W ≤ 적정W 순이어야 합니다.")
        except ValueError as e:
            messagebox.showwarning("입력 오류", f"지속계수(W) 오류: {e}")
            return

        # UI 비활성화
        self._is_running = True
        self.run_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.progress.configure(value=0)

        include_quarterly = self.include_quarterly_var.get()
        include_trend = self.include_trend_var.get()

        # 백그라운드 스레드에서 실행
        thread = threading.Thread(
            target=self._run_generation,
            args=(stock_query, api_key, return_rate, num_years, is_intl,
                  w_buy, w_fair, include_quarterly, include_trend),
            daemon=True,
        )
        thread.start()

    def _run_generation(self, stock_query: str, api_key: str, return_rate: float,
                        num_years: int, is_intl: bool = False,
                        w_buy: float = DEFAULT_W_BUY,
                        w_fair: float = DEFAULT_W_FAIR,
                        include_quarterly: bool = False,
                        include_trend: bool = True):
        """백그라운드에서 리포트 생성 실행"""
        try:
            # 초기화
            self._set_status("데이터 수집 준비 중...")
            self._log("=" * 45)
            market_label = "해외" if is_intl else "한국"
            self._log(f"  재무제표 분석 PDF 리포트 생성 시작 ({market_label})")
            self._log("=" * 45)

            if is_intl:
                from international_fetcher import InternationalFetcher
                fetcher = InternationalFetcher()
            else:
                fetcher = DataFetcher(api_key=api_key)
            self._set_progress(3)

            # 0. 종목코드 확인
            self._set_status("종목 확인 중...")
            self._log(f"\n[종목] '{stock_query}' 검색 중...")
            stock_code = fetcher.resolve_stock_query(stock_query)
            self._log(f"  → 종목코드: {stock_code}")
            self._set_progress(5)

            # 1. 기업 정보
            self._set_status("1/9 기업 정보 조회 중...")
            self._log("\n[1/9] 기업 정보 조회 중...")
            company_info = fetcher.fetch_company_info(stock_code)
            self._log(f"  → {company_info['corp_name']} ({company_info['stock_code']})")
            self._set_progress(8)

            # 2. 최신 연도 탐색
            self._set_status("2/9 최신 사업보고서 탐색 중...")
            self._log("[2/9] 최신 사업보고서 연도 탐색 중...")
            latest_year = fetcher.find_latest_available_year(stock_code)
            years = list(range(latest_year - num_years + 1, latest_year + 1))
            self._log(f"  → 분석 기간: {years[0]}~{years[-1]}")
            self._set_progress(13)

            # 3. 발행주식수
            self._set_status("3/9 발행주식수 조회 중...")
            self._log("[3/9] 발행주식수 조회 중...")
            shares = fetcher.fetch_shares_outstanding(stock_code, latest_year)
            self._log(f"  → 발행주식수: {shares:,}주")
            self._set_progress(20)

            # 4. 주가
            self._set_status("4/9 주가 데이터 조회 중...")
            self._log("[4/9] 주가 데이터 조회 중...")
            stock_data = fetcher.fetch_stock_data(stock_code, shares=shares)
            if is_intl:
                ex_rate = stock_data.get("exchange_rate", 0)
                price_krw = stock_data.get("price_krw", 0)
                mktcap_eok = stock_data.get("market_cap_eok", 0)
                self._log(f"  → 주가: ${stock_data['price']:,.2f} ({price_krw:,}원)")
                self._log(f"  → 시총: ${stock_data['market_cap'] / 1e9:,.1f}B ({mktcap_eok:,}억원)")
                if ex_rate:
                    self._log(f"  → 환율: 1USD = {ex_rate:,.0f}원")
            else:
                self._log(f"  → 주가: {stock_data['price']:,}원, 시총: {stock_data['market_cap_eok']:,}억원")
            self._set_progress(28)

            # 5. 주요계정
            self._set_status("5/9 주요계정 조회 중...")
            self._log("[5/9] 주요계정 조회 중...")
            financial_summary = fetcher.fetch_financial_summary(stock_code, years)
            self._log(f"  → {len(financial_summary)}개 연도 데이터")
            self._set_progress(42)

            # 6. 재무상태표
            self._set_status("6/9 재무상태표 세부항목 조회 중...")
            self._log("[6/9] 재무상태표 세부항목 조회 중...")
            balance_sheet = fetcher.fetch_balance_sheet_detail(stock_code, years)
            self._log(f"  → {len(balance_sheet)}개 연도 데이터")
            self._set_progress(55)

            # 7. 현금흐름표
            self._set_status("7/9 현금흐름표 세부항목 조회 중...")
            self._log("[7/9] 현금흐름표 세부항목 조회 중...")
            cash_flow = fetcher.fetch_cash_flow_detail(stock_code, years)
            self._log(f"  → {len(cash_flow)}개 연도 데이터")
            self._set_progress(65)

            # 8. 주주현황
            self._set_status("8/9 주요 주주 조회 중...")
            self._log("[8/9] 주요 주주 조회 중...")
            shareholders = fetcher.fetch_major_shareholders(stock_code)
            self._log(f"  → {len(shareholders)}명 주주")
            self._set_progress(72)

            # 연도별 종가
            self._log("  추가: 연도별 종가 조회 중...")
            valuation_by_year = fetcher.fetch_valuation_by_year(stock_code, years)
            self._set_progress(78)

            # 9. 컨센서스
            self._set_status("9/9 컨센서스 조회 중...")
            self._log("[9/9] 컨센서스 조회 중...")
            consensus = fetcher.fetch_consensus(stock_code)
            if consensus.get("target_price"):
                tp = consensus['target_price']
                if is_intl:
                    self._log(f"  → 목표주가: ${tp:,.2f}")
                else:
                    self._log(f"  → 목표주가: {tp:,}원")
            else:
                self._log("  → 목표주가: 데이터 없음")
            c_roe = consensus.get("consensus_roe")
            if c_roe:
                self._log(f"  → 컨센서스 ROE: {c_roe:.1f}% (EPS={consensus.get('consensus_eps', 0):,.0f}, BPS={consensus.get('consensus_bps', 0):,.0f})")
            else:
                self._log("  → 컨센서스 ROE: 산출 불가 (EPS/BPS 데이터 없음)")
            self._set_progress(83)

            # 파생 지표 계산
            self._set_status("파생 지표 계산 중...")
            self._log("\n[계산] 파생 지표 계산 중...")

            # CAPM 데이터
            self._log("  추가: CAPM 데이터 조회 중...")
            beta = fetcher.fetch_beta(stock_code)
            risk_free_rate = fetcher.fetch_risk_free_rate()
            self._log(f"  → Beta: {beta:.2f}, 무위험수익률: {risk_free_rate:.2f}%")

            # 분기 데이터 수집 + TTM (선택)
            quarterly_data = None
            ttm = None
            if include_quarterly:
                self._log("\n[분기] 분기별 데이터 조회 중...")
                self._set_status("분기 데이터 조회 중...")
                try:
                    from config import DEFAULT_QUARTERLY_YEARS
                    quarterly_data = fetcher.fetch_quarterly_data(
                        stock_code, num_years=DEFAULT_QUARTERLY_YEARS
                    )
                    quarters = quarterly_data.get("quarters", [])
                    self._log(f"  → {len(quarters)}개 분기 데이터 수집 완료")
                    if quarters:
                        self._log(f"  → 분기 목록: {quarters[0]} ~ {quarters[-1]}")

                    # TTM 계산
                    from generate_report import _compute_ttm
                    ttm = _compute_ttm(quarterly_data)
                    if ttm:
                        self._log("  → TTM(최근 4분기) 계산 완료")
                    else:
                        self._log("  → TTM 계산 불가 (분기 데이터 부족)")
                except Exception as e:
                    self._log(f"  → 분기 데이터 조회 실패: {e}")
                    quarterly_data = None
                    ttm = None

            data = {
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

            roe_src = self.roe_source_var.get()
            manual_roe = None
            if roe_src == "manual":
                try:
                    manual_roe = float(self.manual_roe_var.get())
                except ValueError:
                    raise ValueError("ROE 직접입력 값을 숫자로 입력하세요.")
            coe_src = self.coe_source_var.get()
            derived, srim = self._compute_derived(data, return_rate, roe_src, manual_roe, coe_src,
                                                    w_buy=w_buy, w_fair=w_fair)
            sp = srim.get('srim_price', 0)
            bp = srim.get('buy_price', 0)
            if is_intl:
                self._log(f"  → S-RIM 매수가: ${bp:,.2f} | 적정가: ${sp:,.2f}")
            else:
                self._log(f"  → S-RIM 매수가: {bp:,}원 | 적정가: {sp:,}원")
            self._log(f"    (W: 매수={w_buy}, 적정={w_fair})")
            self._log(f"  → ROE 예측: {srim.get('roe_forecast', 0):.1f}% (소스: {srim.get('roe_source', '')})")
            self._log(f"  → COE: {srim.get('coe_source', '')} = {srim.get('coe_value', 0):.1f}%")
            if coe_src == "capm":
                self._log(f"    (Rf={srim.get('risk_free_rate', 0):.2f}%, "
                          f"Beta={srim.get('beta', 0):.2f}, "
                          f"MRP={srim.get('market_risk_premium', 0):.1f}%)")
            self._set_progress(88)

            # 위험 분석
            self._set_status("위험 분석 중...")
            ttm_derived = derived.get("_ttm")
            if is_intl:
                self._log("\n[위험분석] US listing compliance 체크 중...")
                risk_warnings = check_us_listing_risk(data, derived)
            else:
                self._log("\n[위험분석] 관리종목/상장폐지 조건 체크 중...")
                risk_warnings = check_listing_risk(
                    data, derived,
                    ttm_data=ttm, ttm_derived=ttm_derived
                )
            if risk_warnings:
                self._log(f"  → {len(risk_warnings)}건의 위험 항목 발견")
                for w in risk_warnings:
                    self._log(f"    ⚠ {w['title']}")
            else:
                self._log("  → 위험 항목 없음")
            self._set_progress(90)

            # 추이 분석
            trend_result = None
            if include_trend:
                self._set_status("재무 추이 분석 중...")
                self._log("\n[추이분석] 재무 추이 종합 분석 중...")
                try:
                    from trend_analyzer import analyze_trend
                    trend_result = analyze_trend(
                        data, derived,
                        ttm_data=ttm, ttm_derived=ttm_derived,
                        quarterly_derived=derived.get("_quarterly"),
                        quarterly_keys=derived.get("_quarterly_keys"),
                    )
                    label = trend_result.get("situation_label", "?")
                    conf = trend_result.get("confidence", 0)
                    self._log(f"  → 종합 진단: {label} (신뢰도 {conf * 100:.0f}%)")
                    summary = trend_result.get("summary", "")
                    if summary:
                        # 요약을 50자씩 잘라서 로그에 표시
                        for i in range(0, len(summary), 50):
                            self._log(f"    {summary[i:i+50]}")
                except Exception as e:
                    self._log(f"  → 추이 분석 실패: {e}")
                    trend_result = None
            self._set_progress(94)

            # PDF 생성
            self._set_status("PDF 리포트 생성 중...")
            self._log("\n[PDF] 리포트 생성 중...")

            corp_name = company_info["corp_name"]
            report_type = self.report_type_var.get()
            user_output = self.output_var.get().strip()
            base_dir = os.path.dirname(os.path.abspath(__file__))

            report_data = {
                **data,
                "derived": derived,
                "srim": srim,
                "required_return": return_rate,
                "risk_warnings": risk_warnings,
                "trend_analysis": trend_result,
            }

            generated_files = []

            if report_type == PDF_REPORT_COMBINED:
                output_path = user_output or os.path.join(base_dir, f"{corp_name}_분석리포트.pdf")
                PDFReportGenerator(output_path).generate(report_data)
                generated_files.append(output_path)

            elif report_type == PDF_REPORT_ANNUAL:
                output_path = user_output or os.path.join(base_dir, f"{corp_name}_연도별분석.pdf")
                AnnualReportGenerator(output_path).generate(report_data)
                generated_files.append(output_path)

            elif report_type == PDF_REPORT_QUARTERLY:
                output_path = user_output or os.path.join(base_dir, f"{corp_name}_분기별분석.pdf")
                QuarterlyReportGenerator(output_path).generate(report_data)
                generated_files.append(output_path)

            elif report_type == PDF_REPORT_RISK:
                output_path = user_output or os.path.join(base_dir, f"{corp_name}_리스크분석.pdf")
                RiskReportGenerator(output_path).generate(report_data)
                generated_files.append(output_path)

            elif report_type == PDF_REPORT_ALL:
                base = user_output.rsplit(".", 1)[0] if user_output else os.path.join(base_dir, corp_name)
                for suffix, gen_cls in [
                    ("_연도별분석.pdf", AnnualReportGenerator),
                    ("_분기별분석.pdf", QuarterlyReportGenerator),
                    ("_리스크분석.pdf", RiskReportGenerator),
                ]:
                    path = f"{base}{suffix}"
                    gen_cls(path).generate(report_data)
                    generated_files.append(path)
                    self._log(f"  → 생성: {os.path.basename(path)}")

            self._last_output_path = generated_files[0] if generated_files else ""
            self._set_progress(100)
            files_display = "\n".join(generated_files)
            self._set_status(f"완료! → {len(generated_files)}개 파일 생성")
            self._log(f"\n{'=' * 45}")
            for f in generated_files:
                self._log(f"  완료! 파일: {f}")
            self._log(f"{'=' * 45}")

            self.root.after(0, lambda: self.open_btn.configure(state="normal"))
            self.root.after(0, lambda fd=files_display: messagebox.showinfo(
                "완료", f"리포트가 생성되었습니다.\n\n{fd}"
            ))

        except Exception as e:
            err_msg = str(e)
            self._set_status(f"오류 발생: {err_msg}")
            self._log(f"\n❌ 오류: {err_msg}")
            self.root.after(0, lambda msg=err_msg: messagebox.showerror("오류", msg))

        finally:
            self._is_running = False
            self.root.after(0, lambda: self.run_btn.configure(state="normal"))

    def _compute_derived(self, data: dict, required_return: float,
                         roe_source: str = "consensus",
                         manual_roe: float = None,
                         coe_source: str = "manual",
                         w_buy: float = DEFAULT_W_BUY,
                         w_fair: float = DEFAULT_W_FAIR) -> tuple[dict, dict]:
        """파생 지표 계산 (generate_report.py의 로직 재사용)"""
        from generate_report import compute_derived_metrics
        return compute_derived_metrics(data, required_return,
                                       roe_source=roe_source,
                                       manual_roe=manual_roe,
                                       coe_source=coe_source,
                                       w_buy=w_buy, w_fair=w_fair)

    def _open_file(self):
        if self._last_output_path and os.path.exists(self._last_output_path):
            if sys.platform == "darwin":
                os.system(f'open "{self._last_output_path}"')
            elif sys.platform == "win32":
                os.startfile(self._last_output_path)
            else:
                os.system(f'xdg-open "{self._last_output_path}"')

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ReportApp()
    app.run()
