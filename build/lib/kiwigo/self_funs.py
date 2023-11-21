import os
import sys
from .toolman import Toolman

tm = Toolman()
_dct_ini_ = tm._rd_ini()

class SelfFuns:
    def __init__(self):
        pass

    def open_dir(self, _para):
        if _para == 'kiwigo':
            print(_dct_ini_['_dir_init_py_'])
            os.system("explorer.exe %s" % _dct_ini_['_dir_init_py_'])


