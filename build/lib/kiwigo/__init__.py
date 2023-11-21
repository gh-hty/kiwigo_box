
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
from .tooler_eml import Sq
from .easter_egg import CityLights
from .self_funs import SelfFuns


tm = Toolman()
dct_ini = tm._rd_ini()
# _rd_ini = _rd_ini
init_dir = tm.init_dir
tst = tm.tst
uncps = tm.uncps

sq = Sq()
fix_sqd = sq.fix_sqd
doxs_to_xlx = sq.doxs_to_xlx
emls_to_doxs = sq.emls_to_doxs
get_eml_once = sq.get_eml_once
show_log = sq.show_log

sf = SelfFuns()
open_dir = sf.open_dir

cl = CityLights()
you = cl.you

