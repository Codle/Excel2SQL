import msoffcrypto
import re
import os
from functools import reduce
import pandas as pd


def unlock_excel_to_new_file(input_file_path, output_dir='./'):
    """
    解锁上锁的 EXCEL 文件，使用魔法密码 “VelvetSweatshop” 可以解决 EXCEL 可以打开，但是命令行打开却提示上锁的情况。
    本函数读入一个上锁 EXCEL 文件，输出解锁后的文件。
    :param input_file_path:
    :param output_dir:
    :return:
    """

    wb_msoffcrypto_file = msoffcrypto.OfficeFile(open(input_file_path, 'rb'))
    wb_msoffcrypto_file.load_key(password='VelvetSweatshop')

    assert input_file_path.endswith('.xls')
    wb_unencrypted_filename = input_file_path

    with open(output_dir+os.path.basename(wb_unencrypted_filename), 'wb') as tmp_wb_unencrypted_file:
        wb_msoffcrypto_file.decrypt(tmp_wb_unencrypted_file)

    return output_dir+os.path.basename(wb_unencrypted_filename)


def get_tables(df):
    table = dict()
    cur_name = None
    head_flag = False
    for idx, row in df.iterrows():
        # 一行中为空的字符串
        null_col = pd.isnull(row)
        if row[0] == '指标 Indicators':
            # 判断表头
            if head_flag:
                pass
            else:
                head_flag = True
            years = []
            for item in row[2:]:
                years.append(int(item))
            table['years'] = years
        elif reduce(lambda x, y: x & y, null_col[1:]):
            # 忽略注释内容与空行
            pass
        elif reduce(lambda x, y: x & y, null_col[2:]):
            # 小表头
            cur_name = re.sub('[#\r\t\n\s]', '', row[0])
            table[cur_name] = dict()
        else:
            name = re.sub('[#\r\t\n\s]', '', row[0])
            table[cur_name][name] = [item for item in row[2:]]

    return table
