@echo off
rem # CreateExcelObjec.bat
echo 以下のファイルを生成します。
echo MultiplicationTable.xlsx
echo MultiplicationTable.txt

echo resultフォルダを作成します。
mkdir result

ruby ExecutionCmd.rb

pause