#! ruby -KS
# ExcelファイルとTextファイルを生成する関数の実行ファイル
#Original Author    Yamada
#変更者：Yamada
#変更日 ：2018.03.22

require 'win32ole'
require './CreateExcelModule'
require './CreateExcelClass'

# ----------------------------------------------------
# 関数の実行要求
# ----------------------------------------------------
execution = Execution.new
execution.main()
execution.main_2()
