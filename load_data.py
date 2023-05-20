import pymysql
from sqlalchemy import create_engine
import pandas as pd

# 读取Excel文件
file1 = '木工班组花名册-2023.03.17.xlsx'  # 以它为准
file2 = '花名册-木工班组(22人)2023.03.09进场.xlsx'
file3 = '花名册-木工班组总.xlsx'
df1 = pd.read_excel(file1, header=0,
                    names=['序号', '姓名', '身份证', '性别', '工种', '电话', '家庭住址', '合同签订时间', '进场时间',
                           '离场时间'], dtype={'银行卡号': str, '电话': str})
df2 = pd.read_excel(file2, header=0,
                    names=['序号', '姓名', '身份证', '性别', '工种', '电话', '家庭住址', '合同签订时间', '进场时间',
                           '离场时间', '银行卡号'], dtype={'银行卡号': str, '电话': str})
df3 = pd.read_excel(file3, header=0,
                    names=['序号', '姓名', '身份证', '性别', '工种', '电话', '家庭住址', '银行卡号', '建行开户行'],
                    dtype={'银行卡号': str, '电话': str})
# 合并数据
df = pd.concat([df1, df2, df3], ignore_index=True)
# 去重
df = df.drop_duplicates(subset=['姓名'])

# 修改列的顺序
df = df[['序号', '姓名', '性别', '工种', '电话', '家庭住址', '身份证', '合同签订时间', '进场时间',
         '离场时间', '银行卡号', '建行开户行']]

# 打印表格
print(df.to_string(index=False))


# MySQL连接信息
config = {
    "host": "localhost",
    "user": "root",
    "password": "root",
    "database": "工地"
}

# 创建数据库连接
engine = create_engine('mysql+pymysql://%(user)s:%(password)s@%(host)s/%(database)s?charset=utf8mb4' % config, echo=True)

# 插入数据到MySQL
df.to_sql(name='工人表', con=engine, if_exists='replace', index=False)

if __name__ == '__main__':
    print("hello")

