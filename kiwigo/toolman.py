import os
import re
import shutil
import time
import inspect
import configparser


class Toolman():
    def __init__(self):
        pass

    def tst(self, p):
        print(p)

    # 删除多出的output目录
    def __cls_otpth(self, _dir_run_py, _hd_dir_ot, bf_mx=3):
        lst_pyout = []
        for _dir in os.listdir(_dir_run_py):
            if os.path.isdir(_dir) and _hd_dir_ot in _dir:
                lst_pyout.append(_dir)
        lst_pyout.sort()
        if len(lst_pyout) > bf_mx:
            lst_pyout1 = lst_pyout[:-bf_mx]
        else:
            return 0

        for _d in lst_pyout1:
            _dr = os.path.join(_dir_run_py, _d)
            try:
                shutil.rmtree(_dr)
                print('[删除目录]', _dr)
            except:
                print('[删除失败]', _dr)
                continue

    def _rd_ini(self, ):
        dct_ini = {}
        dct_ini['_dir_run_py_'] = os.getcwd()
        # print('_dir_run_py_', dct_ini['_dir_run_py_'])

        # 1. 读取.ini文件
        _dir_me_ = os.path.dirname(os.path.realpath(__file__))
        _rpth_ini_ = './kiwigo.ini'
        os.chdir(dct_ini['_dir_run_py_'])
        _pth_ini_ = os.path.abspath(_rpth_ini_)

        conf = configparser.ConfigParser()
        conf.read(_pth_ini_, encoding="utf-8-sig")   # todo: testing the ability
        d = conf._sections
        # print(d, _pth_ini_)

        # 3. 生成输入文件目录
        os.chdir(dct_ini['_dir_run_py_'])
        dct_ini['_dir_afbase_'] = os.path.abspath(d['dir_input']['_dir_afbase_'])
        dct_ini['_dir_afb_rjzx_'] = os.path.abspath(d['dir_input']['_dir_afb_rjzx_'])
        dct_ini['_dir_afb_kyzx_t_'] = os.path.abspath(d['dir_input']['_dir_afb_kyzx_t_'])
        dct_ini['_dir_afb_kyzx_u_'] = os.path.abspath(d['dir_input']['_dir_afb_kyzx_u_'])
        dct_ini['_dir_afb_hlge_'] = os.path.abspath(d['dir_input']['_dir_afb_hlge_'])
        dct_ini['_dir_eml_'] = os.path.abspath(d['dir_input']['_dir_eml_'])
        dct_ini['_dir_init_py_'] = os.path.dirname(os.path.realpath(__file__))

        # 4. 生成输出文件目录
        os.chdir(dct_ini['_dir_run_py_'])
        dct_ini['_hd_dir_out_'] = os.path.basename(d['dir_output']['_out_rpth_hd_'])
        dct_ini['_dir_out_'] = os.path.abspath(
            d['dir_output']['_out_rpth_hd_'] + time.strftime('%Y%m%d_%H%M%S', time.localtime()))
        # print(dct_ini['_dir_out_'])

        os.chdir(dct_ini['_dir_run_py_'])
        return dct_ini

    # 路径
    def init_dir(self, ):
        dct_ini = self._rd_ini()

        lst_k = list(filter(lambda x: re.match('_dir_.+', x) != None, dct_ini.keys()))
        for _k in lst_k:
            if not os.path.exists(dct_ini[_k]):
                os.mkdir(dct_ini[_k])
                print('[新建目录]', dct_ini[_k])

        # 删除多余的output目录
        self.__cls_otpth(_dir_run_py=dct_ini['_dir_run_py_'], _hd_dir_ot=dct_ini['_hd_dir_out_'])

    '''
    _in: <str>, <list>
    _rm: <str>, <list>
    del_blk: 若为True，输出时删除所有为空的元素
    '''
    @classmethod
    def rm_sig(cls, _in, _rm, del_blk=True):
        _type_in_is_list_ = True if type(_in).__name__ == 'list' else False
        _in = [_in] if not _type_in_is_list_ else _in
        _rm = [_rm] if type(_rm).__name__ == 'str' else _rm

        _lst_tmp = []
        for _i in _in:
            for _r in _rm:
                _i = _i.replace(_r, '')

            _lst_tmp.append(_i)

        # 检验元素是否符合个数
        if len(_in) != len(_lst_tmp):
            print(len(_in), len(_lst_tmp))
            print('[报错][loc]', inspect.currentframe().f_code.co_filename, inspect.currentframe().f_code.co_name)
            exit('[error] list个数删除前后不同')

        if del_blk:
            while '' in _lst_tmp:
                _lst_tmp.remove('')
        rst_ = _lst_tmp if _type_in_is_list_ else _lst_tmp[0]

        return rst_

    '''.docx\.xlsx 等文件的功能'''
    # def: 找到内层嵌套的表格
    def get_nested_tables_solu1(self, table):
        for table_row in table.rows:
            for table_cell in table_row.cells:
                return table_cell.tables[0]

    def fun_get_run_text(self, t1, _set='str', _mode='soft', max_p=10, max_r=10):
        if _set == 'str':
            str1 = ''
            for p_i in range(max_p):
                for r_i in range(max_r):
                    try:
                        str1 = str1 + t1.paragraphs[p_i].runs[r_i].text
                    except:
                        if _mode == 'soft':
                            break
                        elif _mode == 'hard':
                            continue
            return str1
        elif _set == 'list':
            _lst = []
            for p_i in range(max_p):
                for r_i in range(max_r):
                    try:
                        _lst.append(t1.paragraphs[p_i].runs[r_i].text)
                    except:
                        if _mode == 'soft':
                            break
                        elif _mode == 'hard':
                            continue
            return _lst

    # 读取docx中一行的单元格信息（主要用于获取授权单起止时间）
    def get_row_text(self, t1, _row, _set='str', _mode='soft', max_p=10, max_r=10, max_t=10):
        _lst = []
        for _t in range(max_t):
            try:
                _lst.append(
                    self.fun_get_run_text(t1=t1.cell(_row, _t), _set=_set, _mode=_mode, max_p=max_p, max_r=max_r))
            except:
                continue
        return _lst

    # 用于改写docx内table的内容（input t1 代表了哪个部分的表格）
    def fun_chg_run_text(self, t1, _txt, max_p=10, max_r=10):
        lst_ = []

        # 1. 找到最早的【表头】内容的 paragraphs[p_i], runs[r_i]下标
        for p_i in range(max_p):
            for r_i in range(max_r):
                while len(lst_) < 1:
                    try:
                        t1.paragraphs[p_i].runs[r_i].text = ''
                        lst_.extend([p_i, r_i])
                        break
                    except:
                        break

        # 2. 清除所有【表头】内容
        for p_i in range(max_p):
            for r_i in range(max_r):
                try:
                    t1.paragraphs[p_i].runs[r_i].text = ''
                except:
                    break

        # 3. 写入新的表头内容
        t1.paragraphs[lst_[0]].runs[lst_[1]].text = _txt\

    # todo: 解压工具

