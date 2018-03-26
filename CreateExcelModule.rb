#! ruby -KS
# ExcelObjectに関するModule
#Original Author    Yamada
#変更者：Yamada
#変更日 ：2018.03.22

require 'win32ole'
require 'bigdecimal'
require 'bigdecimal/util'  #<= to_d メソッドが使えるようになる

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
  # 罫線の範囲テーブル
    range = sheets.range('A1:J10')
  # 背景色の範囲テーブル
    range_row = sheets.range('B1:J1')
    range_cal = sheets.range('A1:A10')
    
  # 罫線を引く
    range.borders.lineStyle = EXCEL_CONST::XlContinuous
  # 背景色を塗る
    range_row.interior.themeColor = EXCEL_CONST::XlThemeColorAccent1
    range_cal.interior.themeColor = EXCEL_CONST::XlThemeColorAccent1

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
 def readExcelWorkbook(app, file)
 book = app.Workbooks.Open(file)
  
 # カレントディレクトリの変更
  Dir.chdir("result")
  
 # Excelのマトリクスを読み込むためのloop処理 
  File.open("MultiplicationTable.txt","w") do |text|
   (2..10).each do |j|
    array_col = book.sheets(1).Cells(j, 1).Value
     text.puts array_col.to_i.to_s + "TimesTable"
         
   array_row = []
     (1..10).each do |i|
      array_row = array_row.push(book.sheets(1).Cells(j, 1+i).Value)
     end
      text.puts array_row.map{|a| a.to_i }.join(",")
   end
  end
  
 # ファイルを閉じる
  book.Close
 end
end