import xlrd
import pandas as pd
import argparse
import utils
import glob
import sqlite3
import re
import os

parser = argparse.ArgumentParser()
parser.add_argument('-m', '--model', type=str, default='all', help='使用模式，可选有： all, append 两种')
parser.add_argument('-i', '--input', type=str, default='excels', help='输入文件文件夹')
parser.add_argument('-f', '--file', type=str, help='新增文件, 仅在 append 模式下用')
args = parser.parse_args()


def save_to_sql(tables, sql_file='database.db'):
    conn = sqlite3.connect(sql_file)
    c = conn.cursor()

    for key in tables.keys():
        if key == 'years':
            pass
        else:
            # 新建 失败则证明以前建过 跳过
            try:
                c.execute(
                    """
                    CREATE TABLE {} (
                    YEAR INT NOT NULL PRIMARY KEY
                    )
                    """.format(key)
                )
            except sqlite3.OperationalError:
                pass

            for year in tables['years']:
                try:
                    c.execute("""INSERT INTO {} (YEAR) VALUES ({})""".format(key, int(year)))
                except sqlite3.IntegrityError:
                    pass

            c.execute("PRAGMA TABLE_INFO([{}])".format(key))
            cur_names = [row[1] for row in c]
            for idx, name in enumerate(tables[key].keys()):
                if name not in cur_names:
                    # print("""ALTER TABLE {} ADD "{}" CHAR""".format(key, name))
                    c.execute("""ALTER TABLE {} ADD "{}" FLOAT""".format(key, name))

                for i, value in enumerate(tables[key][name]):
                    try:
                        c.execute("""UPDATE {} SET "{}" = {} WHERE YEAR = {}""".
                                  format(key, name, value, tables['years'][i]))
                    except sqlite3.OperationalError:
                        pass

    conn.commit()
    conn.close()


def year_in_file_name(file_name):
    return int(re.search('\d{4}', file_name).group())


if __name__ == '__main__':
    if not os.path.isdir('unlock'):
        os.mkdir('unlock')

    if args.model == 'all':
        # 获取文件名称，提取名称中的年份排序 由小到大
        files = glob.glob(args.input+'/*.xls')
        files = sorted(files, key=year_in_file_name)

        for file in files:
            print(file)
            try:
                df = pd.read_excel(file)
            except xlrd.biffh.XLRDError:
                new_excel = utils.unlock_excel_to_new_file(file, 'unlock/')
                df = pd.read_excel(new_excel)
            cur_file_tables = utils.get_tables(df)

            save_to_sql(cur_file_tables)
            print("添加成功")
    elif args.model == 'append':

        try:
            df = pd.read_excel(args.file)
        except xlrd.biffh.XLRDError:
            new_excel = utils.unlock_excel_to_new_file(args.file, output_dir='./unlocked_files/')
            df = pd.read_excel(new_excel)

        cur_file_tables = utils.get_tables(df)

        save_to_sql(cur_file_tables)
        print("添加成功")
