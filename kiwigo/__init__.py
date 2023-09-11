
import os

# # 0. 本.py文件所在目录（不是文件）
# _dir_me_ = os.path.dirname(os.path.realpath(__file__))
# _rpth_ini_ = './heya.ini'
#
# # 1. 获取存放参数的.ini文件路径
# os.chdir(_dir_me_)
# _pth_ini_ = os.path.abspath(_rpth_ini_)
#
# print(os.path.abspath(_pth_ini_))




from .toolman import Toolman
from .tooler_docx import Af
from .tooler_eml import Sq

tm = Toolman()
init_dir = tm.init_dir

af = Af()
# doxs_to_xlx = af.doxs_to_xlx

sq = Sq()
doxs_to_xlx = sq.doxs_to_xlx
emls_to_doxs = sq.emls_to_doxs


