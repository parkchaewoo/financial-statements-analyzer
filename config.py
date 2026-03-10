"""설정값 관리 모듈"""

import os
from dotenv import load_dotenv

load_dotenv()

# DART OpenAPI 키
DART_API_KEY = os.getenv("DART_API_KEY", "")

# 기본 설정
DEFAULT_REQUIRED_RETURN = 8.0  # S-RIM 원하는 수익률 (%)
DEFAULT_ANALYSIS_YEARS = 5     # 분석 연도 수

# CAPM 기본값
DEFAULT_MARKET_RISK_PREMIUM = 5.5  # 시장위험프리미엄 (%)
# 무위험수익률은 실시간 조회 (미국 10년 국채: ^TNX, 한국 3년 국채: yfinance)

# S-RIM 초과이익 지속계수(W) 기본값
DEFAULT_W_BUY = 0.5    # 비관적 시나리오 (매수시작가)
DEFAULT_W_FAIR = 1.0   # 초과이익 영구 지속 (적정가)
REPORT_TYPE_ANNUAL = "11011"   # 사업보고서 (연간)

# 분기 보고서 코드 (DART)
REPORT_TYPE_Q1 = "11013"       # 1분기보고서
REPORT_TYPE_SEMI = "11012"     # 반기보고서
REPORT_TYPE_Q3 = "11014"       # 3분기보고서

# 분기 보고서 목록 (연도 내 시간순, de-cumulation용)
QUARTERLY_REPORT_CODES = [
    ("11013", "Q1"),   # Q1 누적
    ("11012", "Q2"),   # H1 누적 (Q1+Q2)
    ("11014", "Q3"),   # Q1-Q3 누적
    ("11011", "Q4"),   # 연간 (Q1-Q4)
]

# 분기 데이터 설정
DEFAULT_INCLUDE_QUARTERLY = True   # TTM/분기 포함 여부
DEFAULT_QUARTERLY_YEARS = 2        # 분기 데이터 수집 연도 수 (2년 = 8분기)

# 재무제표 구분
FS_DIV_CONSOLIDATED = "CFS"   # 연결재무제표
FS_DIV_INDIVIDUAL = "OFS"     # 별도재무제표

# 단위 (억원 변환)
UNIT_BILLION = 100_000_000    # 1억

# PDF 리포트 유형 (DART 보고서 코드와 구분)
PDF_REPORT_COMBINED = "combined"      # 기존 통합 리포트
PDF_REPORT_ANNUAL = "annual"          # 연도별 + TTM 리포트
PDF_REPORT_QUARTERLY = "quarterly"    # 분기별 리포트
PDF_REPORT_RISK = "risk"              # 리스크 리포트
PDF_REPORT_ALL = "all"                # 3개 전부
