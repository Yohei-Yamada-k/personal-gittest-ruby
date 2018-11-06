#! ruby -KS
# ExcelOpen_Module�𐶐�����Class
#Original Author    Yamada
#�ύX�ҁFYamada
#�ύX�� �F2018.03.22
#�ύX�Ȃ�

require 'win32ole'
require './CreateExcelModule.rb'

class Execution
    include Excelmodule

 # ----------------------------------------------------
 # Excel�t�@�C���̐���
 # ----------------------------------------------------
    def main
        fso = WIN32OLE.new('Scripting.FileSystemObject')
        file = fso.GetAbsolutePathName("./result/MultiplicationTable.xlsx")

        app = createExcelobject
        createExcelWorkbook(app, file)
        app.quit()
    end

 # ----------------------------------------------------
 # Text�t�@�C���̐���
 # ----------------------------------------------------
    def main_2
        fso = WIN32OLE.new('Scripting.FileSystemObject')
        file = fso.GetAbsolutePathName("./result/MultiplicationTable.xlsx")

        app = createExcelobject
        readExcelWorkbook(app,file)

        app.quit()
    end
end