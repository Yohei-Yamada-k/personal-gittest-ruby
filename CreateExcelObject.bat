@echo off
rem # CreateExcelObjec.bat
echo 以下のフォルダを生成します。
echo result

echo 以下のファイルを生成します。
echo MultiplicationTable.xlsx
echo MultiplicationTable.txt

mkdir result

ruby ExecutionCmd.rb

pause