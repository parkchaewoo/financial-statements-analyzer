"""PyInstaller 빌드 스크립트 - macOS .app / Windows .exe 생성"""

import PyInstaller.__main__
import sys
import os

APP_NAME = "재무분석리포트"
SCRIPT = "gui_app.py"
ICON = None  # .icns(macOS) 또는 .ico(Windows) 파일 경로

base_dir = os.path.dirname(os.path.abspath(__file__))

args = [
    os.path.join(base_dir, SCRIPT),
    "--name", APP_NAME,
    "--windowed",          # GUI 앱 (콘솔 창 없음)
    "--onefile",           # 단일 파일로 패키징
    "--noconfirm",         # 기존 빌드 덮어쓰기
    # 필요한 데이터 파일 포함
    "--hidden-import", "OpenDartReader",
    "--hidden-import", "pykrx",
    "--hidden-import", "reportlab",
    "--hidden-import", "bs4",
    "--hidden-import", "dotenv",
    "--hidden-import", "pandas",
    "--hidden-import", "requests",
    # 프로젝트 모듈 포함
    "--add-data", f"{os.path.join(base_dir, 'config.py')}{os.pathsep}.",
    "--add-data", f"{os.path.join(base_dir, 'data_fetcher.py')}{os.pathsep}.",
    "--add-data", f"{os.path.join(base_dir, 'calculator.py')}{os.pathsep}.",
    "--add-data", f"{os.path.join(base_dir, 'pdf_report.py')}{os.pathsep}.",
    "--add-data", f"{os.path.join(base_dir, 'risk_analyzer.py')}{os.pathsep}.",
    "--add-data", f"{os.path.join(base_dir, 'generate_report.py')}{os.pathsep}.",
    # .env 파일 포함 (있으면)
]

env_path = os.path.join(base_dir, ".env")
if os.path.exists(env_path):
    args.extend(["--add-data", f"{env_path}{os.pathsep}."])

if ICON and os.path.exists(ICON):
    args.extend(["--icon", ICON])

print(f"빌드 시작: {APP_NAME}")
print(f"플랫폼: {sys.platform}")
print(f"출력 경로: dist/{APP_NAME}")

PyInstaller.__main__.run(args)

print(f"\n빌드 완료!")
if sys.platform == "darwin":
    print(f"  앱: dist/{APP_NAME}.app")
    print(f"  또는: dist/{APP_NAME} (CLI 실행)")
else:
    print(f"  실행파일: dist/{APP_NAME}.exe")
