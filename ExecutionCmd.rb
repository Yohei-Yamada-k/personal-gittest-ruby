#! ruby -KS
# 外部インターフェイスからアクセスする
#Original Author    Yamada
#変更者：Yamada
#変更日 ：2018.07.11

system "ruby ExecutionCreateExcel.rb"
$?
$? .pid
$? .exitstatus