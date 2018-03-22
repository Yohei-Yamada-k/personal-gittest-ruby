#! ruby -KS
# 外部インターフェイスからアクセスする
#Original Author    Yamada
#変更者：Yamada
#変更日 ：2018.03.22

system "ruby ExecutionCreateExcel.rb"
$?
$? .pid
$? .exitstatus