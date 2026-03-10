"""BeeWare/Briefcase 빌드 준비 스크립트

백엔드 Python 모듈과 한글 폰트를 mobile/src/financial_report/ 에 복사.
Briefcase는 sources 디렉토리만 번들하므로 이 작업이 필수.

사용법:
    python prepare_mobile.py
    cd mobile
    briefcase create android
    briefcase build android
"""

import os
import shutil

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
MOBILE_SRC = os.path.join(PROJECT_ROOT, "mobile", "src", "financial_report")
RESOURCES_DIR = os.path.join(MOBILE_SRC, "resources")

# 번들할 백엔드 모듈
BACKEND_MODULES = [
    "config.py",
    "calculator.py",
    "data_fetcher.py",
    "international_fetcher.py",
    "generate_report.py",
    "screener.py",
    "risk_analyzer.py",
    "trend_analyzer.py",
    "pdf_report_base.py",
    "pdf_report.py",
    "pdf_annual_report.py",
    "pdf_quarterly_report.py",
    "pdf_risk_report.py",
]

# 한글 폰트 후보 경로
FONT_CANDIDATES = [
    "/Library/Fonts/NanumGothic.ttf",
    "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
    "/System/Library/Fonts/Supplemental/NotoSansGothic-Regular.ttf",
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "C:/Windows/Fonts/malgun.ttf",
]


def main():
    os.makedirs(RESOURCES_DIR, exist_ok=True)

    # 1. 백엔드 모듈 복사
    print("=== 백엔드 모듈 복사 ===")
    copied = 0
    for module in BACKEND_MODULES:
        src = os.path.join(PROJECT_ROOT, module)
        dst = os.path.join(MOBILE_SRC, module)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"  OK: {module}")
            copied += 1
        else:
            print(f"  SKIP: {module} (파일 없음)")
    print(f"  → {copied}/{len(BACKEND_MODULES)}개 모듈 복사 완료\n")

    # 2. .env 파일 복사 (DART API 키)
    env_src = os.path.join(PROJECT_ROOT, ".env")
    env_dst = os.path.join(MOBILE_SRC, ".env")
    if os.path.exists(env_src):
        shutil.copy2(env_src, env_dst)
        print("  OK: .env 복사 완료")
    else:
        print("  INFO: .env 없음 (해외 주식만 사용 가능)")

    # 3. 한글 폰트 번들
    print("\n=== 한글 폰트 번들 ===")
    font_dst = os.path.join(RESOURCES_DIR, "KoreanFont.ttf")
    font_found = False
    for fp in FONT_CANDIDATES:
        if os.path.exists(fp):
            shutil.copy(fp, font_dst)
            print(f"  OK: {fp} → resources/KoreanFont.ttf")
            font_found = True
            break
    if not font_found:
        print("  WARNING: 한글 폰트를 찾을 수 없습니다.")
        print("  수동으로 resources/KoreanFont.ttf 에 폰트를 복사하세요.")

    print("\n=== 준비 완료 ===")
    print("다음 명령으로 빌드하세요:")
    print("  cd mobile")
    print("  briefcase create android")
    print("  briefcase build android")
    print("  briefcase run android")


if __name__ == "__main__":
    main()
