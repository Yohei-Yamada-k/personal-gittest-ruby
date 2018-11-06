#! ruby -KS
# ExcelOpen_Moduleを生成するClass
#Original Author    Yamada
#変更者：Yamada
#変更日 ：2018.03.22
#変更なし

require 'win32ole'
require './CreateExcelModule.rb'

class Execution
    include Excelmodule

 # ----------------------------------------------------
 # Excelファイルの生成
 # ----------------------------------------------------
    def main
        fso = WIN32OLE.new('Scripting.FileSystemObject')
        file = fso.GetAbsolutePathName("./result/MultiplicationTable.xlsx")

        app = createExcelobject
        createExcelWorkbook(app, file)
        app.quit()
    end

 # ----------------------------------------------------
 # Textファイルの生成
 # ----------------------------------------------------
    def main_2
        fso = WIN32OLE.new('Scripting.FileSystemObject')
        file = fso.GetAbsolutePathName("./result/MultiplicationTable.xlsx")

        app = createExcelobject
        readExcelWorkbook(app,file)

        app.quit()
    end
end