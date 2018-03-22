#! ruby -KS
# ExcelObjectに関するModule
#Original Author    Yamada
#変更者：Yamada
#変更日 ：2018.03.22

require 'win32ole'

 # ----------------------------------------------------
 # ExcelVBAの定数ロードに関するModule
 # ----------------------------------------------------
module EXCEL_CONST
end

 # ----------------------------------------------------
 # ExcelObjectに関するModule
 # ----------------------------------------------------
module Excelmodule

 # ----------------------------------------------------
 # Excelのオブジェクトを作成
 # ----------------------------------------------------
 def createExcelobject
  app = WIN32OLE.new('Excel.Application')
  
  # 上書きメッセージを表示しない
  app.displayAlerts = false
  return app
 end


 # ----------------------------------------------------
 # Excelファイルを生成する
 # ----------------------------------------------------
 def createExcelWorkbook(app, file)
  book = app.Workbooks.add
  WIN32OLE.const_load(app, EXCEL_CONST)

  # シートにワークシートの１を指定
  sheets = book.sheets(1)
    range=sheets.range('A1:J10')
    range.borders.lineStyle = EXCEL_CONST::XlContinuous

  # 九九の行列マトリクスを生成
  (1..9).each do |i|
   sheets.Cells(1, 1+i).Value = i
   sheets.Cells(1+i, 1).Value = i
  end

  # 九九の関数を生成
  range=sheets.range('B2:J10')
  range.value = '=$A2*B$1'
  
  # Excelファイルを保存する
  book.SaveAs(file)
  
  # ファイルを閉じる
  book.close
 end

 # ----------------------------------------------------
 # ExcelFileをReadする
 # ----------------------------------------------------
 def readExcelWprkbook(app, file)
 book = app.Workbooks.Open(file)
  
  # カレントディレクトリの変更
  Dir.chdir("result")
  File.open("MultiplicationTable.txt","w") do |text|
   book.ActiveSheet.UsedRange.Rows.each do |row|
    row.Columns.each do |cell|

    ary = cell.Address.to_a
    ary2 = row.Value.to_a     
     #ary = cell.Address.to_a
    text.puts ary.join(",")
    text.puts ary2.join(",")
     # text.puts cell.Address
     # text.puts cell.Value
     # text.puts '--'
    end
   end
  end
  
  # ファイルを閉じる
  book.Close
 end
end