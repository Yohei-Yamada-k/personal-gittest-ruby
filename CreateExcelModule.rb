#! ruby -KS
# ExcelObject�Ɋւ���Module
#Original Author    Yamada
#�ύX�ҁFYamada
#�ύX�� �F2018.03.22

require 'win32ole'

 # ----------------------------------------------------
 # ExcelVBA�̒萔���[�h�Ɋւ���Module
 # ----------------------------------------------------
module EXCEL_CONST
end

 # ----------------------------------------------------
 # ExcelObject�Ɋւ���Module
 # ----------------------------------------------------
module Excelmodule

 # ----------------------------------------------------
 # Excel�̃I�u�W�F�N�g���쐬
 # ----------------------------------------------------
 def createExcelobject
  app = WIN32OLE.new('Excel.Application')
  
  # �㏑�����b�Z�[�W��\�����Ȃ�
  app.displayAlerts = false
  return app
 end


 # ----------------------------------------------------
 # Excel�t�@�C���𐶐�����
 # ----------------------------------------------------
 def createExcelWorkbook(app, file)
  book = app.Workbooks.add
  WIN32OLE.const_load(app, EXCEL_CONST)

  # �V�[�g�Ƀ��[�N�V�[�g�̂P���w��
  sheets = book.sheets(1)
    range=sheets.range('A1:J10')
    range.borders.lineStyle = EXCEL_CONST::XlContinuous

  # ���̍s��}�g���N�X�𐶐�
  (1..9).each do |i|
   sheets.Cells(1, 1+i).Value = i
   sheets.Cells(1+i, 1).Value = i
  end

  # ���̊֐��𐶐�
  range=sheets.range('B2:J10')
  range.value = '=$A2*B$1'
  
  # Excel�t�@�C����ۑ�����
  book.SaveAs(file)
  
  # �t�@�C�������
  book.close
 end

 # ----------------------------------------------------
 # ExcelFile��Read����
 # ----------------------------------------------------
 def readExcelWprkbook(app, file)
 book = app.Workbooks.Open(file)
  
  # �J�����g�f�B���N�g���̕ύX
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
  
  # �t�@�C�������
  book.Close
 end
end