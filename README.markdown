# Excel to SQL

特殊表格处理程序。默认的数据库使用了 SQLite 数据库。

SQLite 是一种简化的数据库。它只有一个文件，可以使用各种数据库可视化软件直接打开它，并支持完整的 SQL 语句，也可以方便的导入 SQL Server 或者 
MySQL 中。

推荐的数据库可视化软件： [SQLiteStudio](https://sqlitestudio.pl/index.rvt?act=download)。

## 安装依赖

Python 要求 3.0 以上版本。
所有的依赖库在 `requirements.txt` 文件中，使用下列命令来安装相关依赖

```commandline
pip install -r requirements.txt
```

## 使用说明

模式选择（-m）：包含两个模式：添加所有（all）和追加模式（append）。

### ALL

最简单的使用方法：

```commandline
python main.py 
```

默认会将 `excels` 内的所有 xls 文件提取数据到 `database.db` 文件。可以使用 `-i` 参数来指明输入文件夹。

```commandline
python main.py -i ~/Documents/xls_dir
```

### APPEND

使用命令：

```commandline
python main.py -m append -f new_excel.xls
```

将 new_excel.xls 文件内容加入数据库。

## Tips

如果数据库出现问题，可以直接删除 `database.db` 文件，然后重新执行命令。