'''
本程序用于；
   - 解析.eml，并提取授权单（一种固定格式的.docx文件）
   - 对提取出的授权单按照一定顺序编号（结合邮件到达时间，优先以同一封邮件多个授权单连号的形式分配授权单编号）

v1.01 更新内容：
   - 修复了已归档授权单为空时报错
   - 增加了自动修复命名不规范的归档授权单功能（尽可能修复，无法修复则需要手动修改）
v1.02 更新内容：
   - 修复了邮件读取目录的笔误
   - 修复了已识别授权单为纯连号或没有最大号内不够插空分配新单号的bug
   - 修复了docx内重写标题（授权单号）后未保存的bug
   - 增加了识别两页以上授权单转pdf转png的功能（*但是多页授权单表格是否能正常读取仍需观察）
v1.03 更新内容：
   - 修复了for循环中try报错break直接跳出循环的错误，改正为用continue跳出错误继续完成循环
   - 增加了.bat自动复制“【授权】”等字样的头信息
   - 区分科运中心、软件中心的机房授权单
   - 加入输出文件夹最大buffer，超过最大数后运行程序会自动清理
   - 增加【授权主题】替换一些无效字符（如Fw:）

todo:
   - 增加.xlsx文件全面记录信息
   - 简化变量命名
   - 进一步解决无归档授权单时报错的算法结构（虽然目前已能初步debug，可以使用set）
   - 考虑是否增加可以自动初始化工作目录的方法（比如新建导入的文件路径）
   - 系统解决.bat转义字符（如&等）

questions:
   - fitz.open()报黄

cautions:
   - 最大归档授权单号不得超过10000

@author: tyhua
'''


import eml_parser
import os
import base64
import datetime
import time
import re
import shutil
import pandas as pd
import copy

import pyperclip
import docx
from docx.enum.style import WD_STYLE_TYPE
from docx import Document
import comtypes.client
import fitz

import pprint


from .toolman import Toolman

tm = Toolman()
_dct_ini_ = tm._rd_ini()


_path_ = _dct_ini_['_dir_run_py_']

# 用于存放已归档的授权文件
_dir_afbase_ = _dct_ini_['_dir_afbase_']
_dir_afb_kyzx_ = _dct_ini_['_dir_afb_kyzx_']
_dir_afb_rjzx_ = _dct_ini_['_dir_afb_rjzx_']
_dir_eml_ = _dct_ini_['_dir_eml_']
_out_path_hd_ = _dct_ini_['_hd_dir_out_']
_out_path_ = _dct_ini_['_dir_out_']
# 用于存放识别成功的授权单文件（.docx）
_dir_docx_ = os.path.join(_out_path_, '#_0_docx_#')
_path_xlsx_ = os.path.join(_path_, 'a.xlsx')


_lst_kyzx_ = ['A3-302', 'A3-305', 'A3-402', 'A3-405', 'A7-405', 'A7-406',
              'DC1-201', 'DC1-202', 'DC1-301', 'DC2-203', 'DC2-401', ]
_lst_rjzx_ = ['A3-303']







# 用于保存邮件信息
_mail_txt_ = '#_mail_#.txt'
_bat_cmd_ = 'echo ☆ |clip'
_bat_fn_ = '##双击复制【授权主题】##.bat'

# 生成【授权主题】时，屏蔽的词
_lst_ban_wd_ = ['Fw:', 'Fw', '以此为准', '[', ']', '【', '】', '转发', '{', '}']

# 生成表信息
_tab_hdlst_ = ['申请来源', '机房授权表编号', '机房', '人员所属', '事由', '日期', '整机', '配件', '变更', '授权人数', '实际人数']
_tab_pth_ = os.path.join(_out_path_, '#_1_xlsx.xlsx')




_tab_hdlst_ = ['序号', '申请来源', '机房授权编号', '机房', '人员所属', '变更号', '进出时间', '授权人数', '实际人数']

# 变更中需要重复去除的符号
_lst_sig_ = [' ', ';', '；', '&', '和', ',', '，', '\\', '/', '+', '|', '、', '\xa0', '\t', '\n']
_lst_docx_sig_ = ['起始日期', '截止日期', '事由', '人员', '设备', '变更单号', '设备名称', '设备数量']

class Sq:
    def __init__(self):
        self.tm = Toolman()
        self.dict_syn = {}

        self.__dir_init(_dir=_dir_eml_)

        self.__cls_otpth()

    def __cls_otpth(self, bf_mx=2):
        lst_pyout = []
        for _dir in os.listdir(_path_):
            if os.path.isdir(_dir) and _out_path_hd_ in _dir:
                lst_pyout.append(_dir)
        lst_pyout.sort()
        if len(lst_pyout) > bf_mx:
            lst_pyout1 = lst_pyout[:-bf_mx]
        else:
            return 0

        for _d in lst_pyout1:
            try:
                shutil.rmtree(os.path.join(_path_, _d))
                print('[删除目录]', _d)
            except:
                print('[删除失败]', _d)
                continue

    def __dir_init(self, _dir):
        if not os.path.exists(_dir):
            os.mkdir(_dir)

    # 对以归档授权单进行处理
    @classmethod
    def _fix_sqd(cls, _dir_afb):
        lst_bad = []

        for _dir in os.listdir(_dir_afb):
            if _dir.endswith('.docx'):
                afn_i, _ = os.path.splitext(_dir)

                # 一次检验
                if not re.search('^20\d\d[0-1]\d-\d{2,3}-\w{3,5}(\()+\d{4}-\d{1,2}-\d{1,2} [0-9]{6}[\u4e00-\u9fa5]{2,3}(\))+$',
                        afn_i):
                    lst_bad.append(afn_i)

        for afn_i in lst_bad:
            a0 = afn_i

            # 若满足最宽泛的条件才有修改的可能，否则无法修改，只能看一眼手动去改
            if re.search(
                    '^[\s]*[0-9]{6}[\s]*-[\s]*[0-9]{2,3}[\s]*-[\s]*.{3,5}[\s]*(\(|[\uff08])+[\s]*\d{4}[\s]*-[\s]*\d{1,2}[\s]*-[\s]*\d{1,2}[\s]*[0-9]{6}[\s]*[\u4e00-\u9fa5]{2,3}[\s]*(\)|[\uff09])[\s]*$', afn_i):
                # [修复情况1]去除多余空格
                afn_i = afn_i.replace(' ', '')
                pointer_blk = [i.span() for i in re.finditer('[0-2]\d[0-6]\d[0-6]\d[\u4e00-\u9fa5]', afn_i)][0][0]
                l_ = list(afn_i)
                l_.insert(pointer_blk, ' ')
                afn_i = ''.join(l_)

                # [修复情况2]括号切换中/英文
                # print(afn_i)
                afn_i = afn_i.replace('（', '(')
                afn_i = afn_i.replace('）', ')')

            # 复检
            lst_bad = []
            if re.search(
                    '^20\d\d[0-1]\d-\d{2,3}-\w{3,5}(\()+\d{4}-\d{1,2}-\d{1,2} [0-9]{6}[\u4e00-\u9fa5]{2,3}(\))+$',
                    afn_i):
                os.rename(os.path.join(_dir_afb, a0 + '.docx'), os.path.join(_dir_afb, afn_i + '.docx'))
                print('修复完成：', a0, '\t->\t', afn_i)
            else:
                lst_bad.append(afn_i)

        print(os.path.split(_dir_afb)[-1], '授权单命名仍未修复：', lst_bad) if lst_bad else print(os.path.split(_dir_afb)[-1], '目录下授权单命名检验正常')

    def _chg_aid(self, _docx, dict_syn_ki):
        try:
            d0 = Document(_docx)
            t_ = d0.tables  # 获取文件中的表格集
            t0 = t_[0]
        except:
            return print('改aid失败')
        self.tm.fun_chg_run_text(t1=t0.cell(0, 1),
                                _txt='人员设备进出机房授权表（编号' + re.findall('^[0-9]{6}-[0-9]{2,3}', dict_syn_ki['name_nw'])[0] + '）')
        d0.save(_docx)

    def _idf_sqd(self, dir):
        # for f_i in os.listdir(dir):  # 仅遍历当前文件，不穿透深层文件夹
            # if f_i.endswith('.docx'):
            #     print(f_i)
        # a = os.path.join(dir, '202307-178（A3302 2023-07-26 256161 滑天扬）.docx')
        if not os.path.exists(_dir_docx_):
            os.mkdir(_dir_docx_)

        # 1. get
        docx_list = []
        dict_docx = {}

        for _dir in os.listdir(dir):
            if _dir.endswith('.docx'):
                docx_list.append(_dir)

        if not docx_list:
            dict_docx[-1] = {}   # 代表此邮件无授权单，否则数值代表其中有几个授权单

        # 2. output
        list_bad = []
        list_good = []
        for _id, af_i in enumerate(docx_list):
            afn_i, _ = os.path.splitext(af_i)

            try:
                d0 = Document(os.path.join(dir, af_i))
                t_ = d0.tables  # 获取文件中的表格集
                t0 = t_[0]
                t_in = self.tm.get_nested_tables_solu1(t0)
                dict_docx[_id] = {}
            except:
                # print('[不是授权单，跳过]', af_i)
                list_bad.append(af_i)
                continue

            # if fun_get_run_text(t1=t0.cell(8, 1).paragraphs[0]):

            dict_docx[_id]['af_name'] = afn_i
            dict_docx[_id]['进出时间'] = self.tm.fun_get_run_text(t1=t0.cell(8, 1)) + \
                                     ' ~ ' + self.tm.fun_get_run_text(t1=t0.cell(8, 6))
            dict_docx[_id]['wkdate'] = self.tm.fun_get_run_text(t1=t0.cell(8, 1))
            dict_docx[_id]['room'] = self.tm.fun_get_run_text(t1=t_in.cell(1, 2))
            # dict_docx['准入单编号'] = auth_id  # todo: 如果有的话

            if self.tm.fun_get_run_text(t1=t_in.cell(0, 1)):
                # print('入场')
                dict_docx[_id]['入出场'] = '入场'
            elif self.tm.fun_get_run_text(t1=t_in.cell(0, 3)):
                # print('出场')
                dict_docx[_id]['入出场'] = '出场'
            else:
                list_bad.append(af_i)
                break

            dict_docx[_id]['厂商'] = self.tm.fun_get_run_text(t1=t0.cell(2, 2)) + ';\n' + self.tm.fun_get_run_text(t1=t0.cell(1, 2))
            dict_docx[_id]['厂商人数'] = sum([int(i) for i in re.findall('[0-9]', self.tm.fun_get_run_text(t1=t0.cell(2, 2)))])

            str1 = self.tm.fun_get_run_text(t1=t0.cell(3, 2)) + '×' + self.tm.fun_get_run_text(t1=t0.cell(4, 2)) + ';' \
                   + self.tm.fun_get_run_text(t1=t0.cell(3, 4)) + '×' + self.tm.fun_get_run_text(t1=t0.cell(4, 4)) + ';' \
                   + self.tm.fun_get_run_text(t1=t0.cell(3, 7)) + '×' + self.tm.fun_get_run_text(t1=t0.cell(4, 7))
            str1 = str1.replace(';×', '')
            if str1.endswith('×'): str1 = str1.replace('×', '')
            if str1.endswith(';'): str1 = str1.replace(';', '')
            dict_docx[_id]['设备'] = str1

            # CHG
            dict_chg = {}
            for i in range(10):
                dict_chg[_id] = []
                dict_chg[_id].append(self.tm.fun_get_run_text(t1=t0.cell(6, i)))
                dict_chg['num_' + str(_id)] = [i.count('CHG') for i in dict_chg[_id]]

            dict_docx[_id]['变更标记（与生产管理核对）'] = dict_chg['num_' + str(_id)][0] \
                if dict_chg['num_' + str(_id)][0] > 0 else ''
            dict_docx[_id]['变更号'] = dict_chg[_id]

            # 申请具体事项
            # df0.loc[_id, '申请具体事项'] = '="text1"&CHAR(10)&"text2"'
            dict_docx[_id]['申请具体事项'] = str(dict_docx[_id]['厂商']) \
                                       + '\n' + str(dict_docx[_id]['变更号']) \
                                       + '\n' + str(self.tm.fun_get_run_text(t1=t0.cell(7, 1)))

            list_good.append(af_i)

        # print('[未识别docx]', len(list_bad), '\n[正常docx]', len(list_good), '\n[已发现(处理)docx]', len(docx_list))
        # print('[未识别docx是]', list_bad)
        # print(list_good)

        # 5. 转移授权单docx
        for af_i in list_good:
            shutil.copyfile(os.path.join(dir, af_i), os.path.join(_dir_docx_, af_i))

        return dict_docx

    def _sqd_png(self, _from_dx, _to_dir):
        _fm_dir, _fm_dx = os.path.split(_from_dx)
        _dx, _ = os.path.splitext(_fm_dx)
        dir_pdf = os.path.join(_fm_dir, _dx + '.pdf')
        pth_png = '_' + _dx

        # 1. .docx -> .pdf
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(_from_dx)
        doc.SaveAs(dir_pdf, FileFormat=17)
        doc.Close()
        word.Quit()

        # 2. .pdf -> .png
        d = fitz.open(dir_pdf)  # open document
        for _pg in range(0, d.page_count - 1):
            page_ = d[_pg]
            pix = page_.get_pixmap()  # render page to an image
            pix.save(os.path.join(_to_dir, '#+' + pth_png + str(_pg) + '+#.png'))  # store image as a PNG
        d.close()

        # 3. del
        os.remove(dir_pdf)

    def _syn_dict(self, dict_docx, dict_eml):
        dict_docx = dict_docx
        dict_eml = dict_eml
        dict_syn = self.dict_syn

        for _id in dict_docx.keys():
            d_key = str(str(dict_eml['date']) + dict_eml['sender'] + str(_id)).replace(' ', '')

            dict_syn[d_key] = {}
            dict_syn[d_key]['e_id'] = str(_id + 1) + '/' + str(len(dict_docx.keys()))
            dict_syn[d_key]['date'] = dict_eml['date']   # todo: use datetime格式
            dict_syn[d_key]['sender'] = dict_eml['sender']
            dict_syn[d_key]['dir_attach'] = dict_eml['dir_attach']
            dict_syn[d_key]['sbj'] = dict_eml['sbj']
            dict_syn[d_key]['af_name'] = dict_docx[_id]['af_name']
            dict_syn[d_key]['wkdate'] = dict_docx[_id]['wkdate']
            dict_syn[d_key]['room'] = dict_docx[_id]['room']

        self.dict_syn = dict_syn

    # 输入一个list（def内自动排序），输出不连续空出的单号
    def __out_blk(self, lst):
        lst = sorted(lst)
        lst_blk = []
        for _p, i in enumerate(lst):
            if _p == len(lst) - 1:
                break
            blk = lst[_p + 1] - lst[_p] - 1
            lst_blk.append(blk)
        return lst_blk

    def __dis_aid(self, lst_dised, lst_x, ):
        # lst_dised = [2, 3, 5, 12]   # 已经分配过的单号
        # lst_x = [3, 1, 2]   # 有几个单子待分配单号
        if not 0 in lst_dised:
            lst_dised.append(0)
        # for debug
        lst_dised.append(10000)
        # todo: use an elegant way to slove the problem, not just for debug
        lst_dised.sort()
        # print(lst_dised, type(lst_dised[0]), type(lst_dised[1]))
        lst_aid = []
        for k in lst_x:
            lst_blk = self.__out_blk(lst=lst_dised)
            for _b, b in enumerate(lst_blk):
                if k <= b:
                    aid_ = list(range(lst_dised[_b] + 1, lst_dised[_b] + k + 1))
                    lst_aid.append(aid_)
                    lst_dised[_b + 1: _b + 1] = iter(aid_)
                    break
        return lst_aid

    def _dist_afid(self, dir_afb, dict_syn, lst_k):
        # 1. 收集目录下授权文件，返回一个list
        # DC1-202 DC1-301 DC1-201 DC2-203 DC2-401
        lst_auth_id = []
        file_cannot_identify_list = []
        for _dir in os.listdir(dir_afb):
            if _dir.endswith('.docx'):
                if re.search('^[\s]*[0-9]{6}-[0-9]{2,3}-.{3,5}(\(|[\uff08])+\d{4}-\d{1,2}-\d{1,2} [0-9]{6}[\u4e00-\u9fa5]{2,3}(\)|[\uff09])+(.docx)$', _dir):
                    auth_mth_i = int(_dir.split('-')[0].replace(' ', ''))
                    auth_id_i = int(_dir.split('-')[1].replace(' ', ''))    # int
                    lst_auth_id.append(auth_id_i)
                else:
                    file_cannot_identify_list.append(_dir)
        print(os.path.split(dir_afb)[-1], '已识别到归档授权编号：', lst_auth_id)
        if file_cannot_identify_list:
            print(os.path.split(dir_afb)[-1], '以下文件无法识别为已授权文件：', file_cannot_identify_list)

        # 2. 分派
        lst_bk = []
        lst_k = sorted(lst_k)
        lst01 = []
        for ki in lst_k:
            lst01.append(ki[:16])

        lst02 = list(set(lst01))
        lst02.sort()

        for s in lst02:
            lst_bk.append(lst01.count(s))
        print(os.path.split(dir_afb)[-1], '待分配单号的单数：', lst_bk)

        # 2. 分配af单号
        # todo: 暂时无法识别邮件中自带顺序的授权单
        lst_aid = self.__dis_aid(lst_dised=lst_auth_id, lst_x=lst_bk)   # lst_dised: int
        print(os.path.split(dir_afb)[-1], '已分配授权单号：', lst_aid)

        # 3. 复写
        lst_aid = [i for p in lst_aid for i in p]
        for _i, ki in enumerate(lst_k):
            dict_syn[ki]['name_nw'] = str(dict_syn[ki]['name_nw']).replace('X', str(lst_aid[_i]).rjust(2, '0'))
            dict_syn[ki]['sbj_aid'] = re.findall('\d{6}-\d{2,3}', dict_syn[ki]['name_nw'])[0][4:]

            self._chg_aid(_docx=os.path.join(_dir_docx_, dict_syn[ki]['af_name'] + '.docx'), dict_syn_ki=dict_syn[ki])

            # 源文件改名
            os.rename(os.path.join(_dir_docx_, dict_syn[ki]['af_name'] + '.docx'),
                      os.path.join(_dir_docx_, dict_syn[ki]['name_nw'] + '.docx'))
            print('已重命名：\t', ki, '\t', dict_syn[ki]['af_name'], '\tto\t', dict_syn[ki]['name_nw'])

            # 截图授权单
            # self._sqd_png(_to_dir=os.path.join(dict_syn[ki]['dir_attach']),
            #               _from_dx=os.path.join(_dir_docx_, dict_syn[ki]['name_nw'] + '.docx'))

        # 4. 生成信息


        # todo: 生成一个表记录信息，便于出错排障

    def _name_sqd(self, dict_syn):
        for d_key in dict_syn.keys():
            set_bad = set()
            # 1. 识别施工日期
            # todo: 有没有更简洁的替代方法
            lst_ymd = ['%Y-%m-%d', '%Y%m%d', '%Y.%m.%d', '%Y年%m月%d日']
            for i in lst_ymd:
                try:
                    wkdate = datetime.datetime.strptime(dict_syn[d_key]['wkdate'], i)
                    dict_syn[d_key]['wkdate_r'] = datetime.datetime.strftime(wkdate, '%Y%m')
                    break
                except:
                    set_bad.add(dict_syn[d_key]['af_name'])

            date_d = datetime.datetime.strptime(dict_syn[d_key]['date'], '%Y%m%d%H%M%S')
            date_ymr = datetime.datetime.strftime(date_d, '%Y-%m-%d')
            date_hms = datetime.datetime.strftime(date_d, '%H%M%S')

            d_t = dict_syn[d_key]
            d_t['name_nw'] = d_t['wkdate_r'] + '-X-' + d_t['room'].replace('-', '') + '(' + date_ymr + ' ' + date_hms + d_t['sender'] + ')'

        return dict_syn

    # todo: 检查dict_syn
    def _chk_dct(self, dict_syn):
        print(dict_syn)

        lst_bad = []
        lst_len_k = []

        for ki in dict_syn.keys():
            lst_len_k.append(len(dict_syn[ki].keys()))

        m_ = max(lst_len_k)

    def _txt(self, dict_syn):
        # 1. 处理主题
        lst_sbjs = []
        d_ks = {}
        for ki in dict_syn.keys():
            d_ks[dict_syn[ki]['sbj']] = ki
            lst_sbjs.append(dict_syn[ki]['sbj'])
        # todo: 同名主题邮件会冲突（可能并不会发生，因为下载.eml时不会出现同名文件）

        # todo: 加入电信联通区分方法
        d_s = {}
        for si in lst_sbjs:
            # 若此处报错无['sbj_aid']或因未识别成功机房引起
            d_s[dict_syn[d_ks[si]]['dir_attach']] \
                = '【电信授权】' \
                  + '、'.join([dict_syn[ki]['sbj_aid'] for ki in dict_syn.keys() if dict_syn[ki]['sbj'] == si]) \
                  + '，' + si

        # 3. 生成自动复制.bat文件
        for _pth in d_s.keys():
            # 处理【授权主题】内容
            _cont = d_s[_pth].split('，')[-1]
            for i in _lst_ban_wd_:
                _cont = _cont.replace(i, '')
            _txt = d_s[_pth].split('，')[0] + '，' + _cont
            print(_txt)

            with open(os.path.join(_pth, _bat_fn_), 'a') as f:
                f.write(_bat_cmd_.replace('☆', d_s[_pth].replace('&', '^&')))
                # todo: 系统解决转义

    def _del_org_docx(self, dict_syn):
        for ki in dict_syn.keys():
            os.remove(os.path.join(dict_syn[ki]['dir_attach'], dict_syn[ki]['af_name'] + '.docx'))
            shutil.copyfile(src=os.path.join(_dir_docx_, dict_syn[ki]['name_nw'] + '.docx'),
                            dst=os.path.join(dict_syn[ki]['dir_attach'], '#' + dict_syn[ki]['name_nw'] + '.docx'))

    def emls_to_doxs(self, _pth_doxs_arc=_dir_afbase_, _pth_eml=_dir_eml_, pth_out=_out_path_):
        _dir_afb_kyzx_ = os.path.join(_pth_doxs_arc, '1.1_科运中心')
        _dir_afb_rjzx_ = os.path.join(_pth_doxs_arc, '1.2_软件中心')
        self.__dir_init(_dir=_pth_doxs_arc)
        self.__dir_init(_dir=_dir_afb_kyzx_)
        self.__dir_init(_dir=_dir_afb_rjzx_)

        dict_eml = {}
        os.mkdir(pth_out) if not os.path.exists(pth_out) else os.mkdir(pth_out + time.strftime('%H%M%S', time.localtime()))

        # 1. 遍历目录下eml文件
        for eml_i in os.listdir(_pth_eml):  # 仅遍历当前文件，不穿透深层文件夹
            if eml_i.endswith('.eml'):

        # 2. 处理eml
                with open(os.path.join(_pth_eml, eml_i), 'rb') as fhdl:
                    raw_email = fhdl.read()
                ep = eml_parser.EmlParser(include_attachment_data=True, include_raw_body=True)
                parsed_eml = ep.decode_email_bytes(raw_email)
                dict_eml['sbj'] = parsed_eml['header']['subject']
                # pprint.pprint(parsed_eml['header'])

        # 3. 保存信息
                eml_hd = parsed_eml['header']
                sender_ = ''.join([i.split(' <', 1)[0] for i in set(eml_hd['header']['from'])])
                date_ = datetime.datetime.strftime(eml_hd['date'], '%Y%m%d%H%M%S')
                print(date_, '\t', sender_, )

                # 创建目录
                fname0 = date_ + '_' + sender_ + os.path.splitext(eml_i)[0]
                f_name = os.path.join(pth_out, fname0)
                if os.path.exists(f_name):
                    f_name = f_name + time.strftime('%H%M%S', time.localtime())
                os.mkdir(f_name)
                dict_eml['dir_attach'] = f_name

        # 4. 创建txt记录正文
                # todo: 后续根据邮件内容确认是否可以只取parsed_eml['body'][0]
                for i in parsed_eml['body']:
                    with open(os.path.join(f_name, _mail_txt_), 'a') as f:
                        f.write(i['content'])

        # 3. 保存eml附件
                if parsed_eml.get('attachment'):
                    for i in parsed_eml['attachment']:
                        x = base64.b64decode(i['raw'])
                        with open(os.path.join(f_name, i['filename']), 'wb') as f:
                            f.write(x)

        # 4. 用于同步信息
                dict_eml['sender'] = sender_
                dict_eml['date'] = date_
                # self.dict_eml= copy.deepcopy(dict_eml)

        # 5. 处理授权单（.docx）
                dict_docx = self._idf_sqd(dir=f_name)
                if not -1 in dict_docx.keys():
                    self._syn_dict(dict_docx=dict_docx, dict_eml=dict_eml)

        # 6. 同步信息
        dict_syn = self._name_sqd(dict_syn=self.dict_syn)

        # todo: 7. 检查dict_syn是否均被识别
        lst_kyzx_k = [ki for ki in dict_syn.keys() if dict_syn[ki]['room'] in _lst_kyzx_]
        lst_rjzx_k = [ki for ki in dict_syn.keys() if dict_syn[ki]['room'] in _lst_rjzx_]
        self._dist_afid(dir_afb=_dir_afb_kyzx_, dict_syn=dict_syn, lst_k=lst_kyzx_k)
        self._dist_afid(dir_afb=_dir_afb_rjzx_, dict_syn=dict_syn, lst_k=lst_rjzx_k)

        self._txt(dict_syn=dict_syn)

        # 8. 删除原.docx授权文件（为简化邮件附件内容）
        self._del_org_docx(dict_syn=dict_syn)

    def doxs_to_xlx(self, ):
        self._doxs_to_xlx(_pth_dox=_dir_afb_kyzx_, _pth_xlx=_path_)
        self._doxs_to_xlx(_pth_dox=_dir_afb_rjzx_, _pth_xlx=_path_)

    def _doxs_to_xlx(self, _pth_dox, _pth_xlx=_path_):
        # 1. 获取所有.docx文件
        authfile_list = []
        for _dir in os.listdir(_pth_dox):
            if _dir.endswith('.docx'):
                authfile_list.append(_dir)

        # 2. output
        list_bad = []
        lst_chgs = []   # 用于变更去重
        df0 = pd.DataFrame([], columns=_tab_hdlst_)
        for row_i, af_i in enumerate(authfile_list):
            afn_i, _ = os.path.splitext(af_i)

            # 3. 获取【序号】、【申请来源】
            df0.loc[row_i, '序号'] = row_i + 1
            df0.loc[row_i, '申请来源'] = afn_i

            # 4. 尝试获取【授权编号】、【机房】
            try:
                auth_id = re.match('^[\s]*[0-9]{6}-[0-9]+', afn_i).group()
                room = re.findall('[a-z]{1}[0-9]{4}', afn_i, re.I)[0]
                app_date = re.findall('[0-9]{4}-[0-9]{2}-[0-9]{2}', afn_i)[0]
                apper = re.findall('[\u4e00-\u9fa5]{2,3}', afn_i)[0]
                df0.loc[row_i, '机房授权编号'] = auth_id
                df0.loc[row_i, '机房'] = room
            except:
                df0.loc[row_i, '机房授权编号'] = ''
                df0.loc[row_i, '机房'] = ''
                list_bad.append(af_i)

            # 5. 对表格内容进行继续获取
            d0 = Document(os.path.join(_pth_dox, af_i))
            t_ = d0.tables  # 获取文件中的表格集
            t0 = t_[0]

            # 6. 获取【人员所属】、【授权人数】、【设备】
            df0.loc[row_i, '人员所属'] = self.tm.fun_get_run_text(t1=t0.cell(2, 2)) + ';\n' + self.tm.fun_get_run_text(t1=t0.cell(1, 2))
            df0.loc[row_i, '授权人数'] = sum([int(i) for i in re.findall('[0-9]', self.tm.fun_get_run_text(t1=t0.cell(2, 2)))])

            str1 = self.tm.fun_get_run_text(t1=t0.cell(3, 2)) + '×' + self.tm.fun_get_run_text(t1=t0.cell(4, 2)) + ';' \
                   + self.tm.fun_get_run_text(t1=t0.cell(3, 4)) + '×' + self.tm.fun_get_run_text(t1=t0.cell(4, 4)) + ';' \
                   + self.tm.fun_get_run_text(t1=t0.cell(3, 7)) + '×' + self.tm.fun_get_run_text(t1=t0.cell(4, 7))

            str1 = str1.replace(';×', '')
            if str1.endswith('×'): str1 = str1.replace('×', '')
            if str1.endswith(';'): str1 = str1.replace(';', '')
            df0.loc[row_i, '设备'] = str1

            # 7. 获取【变更号】、【申请具体事项】
            chg_bare = self.tm.get_row_text(t1=t0, _row=6, _set='str', _mode='hard')
            chg_nosig = self.tm.rm_sig(_in=chg_bare, _rm=_lst_sig_)

            lst0 = []
            [lst0.extend(i) for i in [re.findall('CHG-\w{2,}-\w{2,}-\w{2,}-\d{8}-\d{4}', i) for i in chg_nosig]]

            df0.loc[row_i, '变更号'] = ', '.join(set(lst0))

            # 用于后续总变更去重
            lst_chg_i = list(set(lst0))
            if lst_chg_i:
                lst_chgs.extend(lst_chg_i)

            # 申请具体事项
            rsn_bare = self.tm.get_row_text(t1=t0, _row=7, _set='str', _mode='hard')

            lst_rsn_i = []
            for _r in rsn_bare:
                # 去除头尾坏字符
                for _sig in _lst_sig_:
                    _r = _r.strip(_sig)
                lst_rsn_i.append(_r)
            lst_rsn_i = self.tm.rm_sig(_in=lst_rsn_i, _rm=_lst_sig_ + _lst_docx_sig_)
            lst_rsn_i = list(set(lst_rsn_i))

            if len(lst_rsn_i) == 1:
                df0.loc[row_i, '申请具体事项'] = lst_rsn_i[0]
            elif len(lst_rsn_i) >= 2:
                df0.loc[row_i, '申请具体事项'] = lst_rsn_i[1]
            else:
                df0.loc[row_i, '申请具体事项'] = ''
                list_bad.append(af_i)

            # 8. 获取【进出时间】
            date_bare = self.tm.get_row_text(t1=t0, _row=8, _set='str', _mode='hard')
            date_bare = self.tm.rm_sig(_in=date_bare, _rm=_lst_sig_ + _lst_docx_sig_)
            lst_date_i = list(set(date_bare))


            # todo: 选最长的
            if len(lst_date_i) == 1:
                df0.loc[row_i, '进出时间'] = lst_date_i[0] + ' ~ ' + lst_date_i[0]
            elif len(lst_date_i) == 2:
                lst_date_i.sort()
                df0.loc[row_i, '进出时间'] = lst_date_i[0] + ' ~ ' + lst_date_i[1]
            else:
                df0.loc[row_i, '进出时间'] = ''
                list_bad.append(af_i)

        # 统计变更，变更去重
        set_chgs = set(lst_chgs)
        print('[总变更数（不去重）] ', len(lst_chgs))
        print('[总变更数（去重）] ', len(set_chgs))

        dct_chgs = {}
        for _chg in set_chgs:
            dct_chgs[_chg] = lst_chgs.count(_chg)
        print('[各变更次数（出现在授权单中几次）]', dct_chgs)

        df0.to_excel(_path_xlsx_, index=False)

        print('[识别失败]\t', len(list_bad), ': \t', list_bad)

# if __name__ == "__main__":
#     Sq._fix_sqd(_dir_afb=_dir_afb_kyzx_)
#     Sq._fix_sqd(_dir_afb=_dir_afb_rjzx_)
#
#     sq = Sq()
#     sq.emls_to_doxs(_pth_doxs=, pth_eml=_dir_eml_, pth_out=)
