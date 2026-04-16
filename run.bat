@echo off
chcp 65001 > nul
echo 수불부 자동 생성 앱을 시작합니다...
echo 브라우저가 자동으로 열립니다. (http://localhost:8501)
echo 종료하려면 이 창에서 Ctrl+C 를 누르세요.
echo.
streamlit run app.py
pause
