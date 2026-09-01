@echo off
setlocal
cd /d %~dp0

python -m pip install -r requirements.txt
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --add-data "config.json;." ^
  --name "南方分中心舆情质检辅助工具" ^
  main.py

echo.
echo 打包完成，文件位置：
echo %cd%\dist\南方分中心舆情质检辅助工具\南方分中心舆情质检辅助工具.exe
pause
