#! ruby -KS
# Excel�t�@�C����Text�t�@�C���𐶐�����֐��̎��s�t�@�C��
#Original Author    Yamada
#�ύX�ҁFYamada
#�ύX�� �F2018.03.22

require 'win32ole'
require './CreateExcelModule'
require './CreateExcelClass'

# ----------------------------------------------------
# �֐��̎��s�v��
# ----------------------------------------------------
execution = Execution.new
execution.main()
execution.main_2()
