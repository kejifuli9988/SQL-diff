@echo off
setlocal

cd /d %~dp0

if not exist dist mkdir dist

pyinstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name SQL语句比较工具 ^
  sql_diff_gui.py

echo.
echo 打包完成，exe 路径：
echo %~dp0dist\SQL语句比较工具.exe
pause
