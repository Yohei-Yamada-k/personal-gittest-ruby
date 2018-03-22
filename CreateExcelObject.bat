@echo off
rem # CreateExcelObjec.bat
echo 以下のファイルを生成します。
echo test.xlsx
echo text_write_test.txt

echo resultフォルダを作成します。
mkdir result

ruby ExecutionCmd.rb

pause