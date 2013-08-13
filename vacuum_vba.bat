@echo off

rem
rem このバッチの説明
rem

rem 設定事項
set HOGE="変数の値"

rem このバッチが存在するフォルダをカレントに
pushd %0\..
cls

START CSCRIPT vacuum_vba.js

pause
exit
