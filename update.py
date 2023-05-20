import pandas as pd
import pymysql

file = '更新.xlsx'  # 以它为准
df = pd.read_excel(file, header=0,
                   names=['序号', '姓名', '身份证', '性别', '工种', '电话', '家庭住址', '合同签订时间', '进场时间',
                          '银行卡号', '建行开户行'], dtype={'银行卡号': str, '电话': str})

# 连接数据库
conn = pymysql.connect(host='localhost', port=3306, user='root', password='root', db='工地')

# 获取游标
cursor = conn.cursor()

for index, row in df.iterrows():
    print(row['姓名'], row['身份证'], row['性别'], row['工种'], row['电话'], row['家庭住址'],
          row['合同签订时间'], row['进场时间'], row['银行卡号'], row['建行开户行'])
    # 构造 SQL 语句
    # 序号, 姓名, 性别, 工种, 电话, 家庭住址, 身份证, 合同签订时间, 进场时间, 离场时间, 银行卡号, 建行开户行
    sql = "INSERT INTO 工人 (姓名, 性别, 工种, 电话, 家庭住址, 身份证, 合同签订时间, 进场时间, 银行卡号, 建行开户行) " \
          "SELECT %s, %s, %s, %s, %s, %s, %s, %s, %s, %s " \
          "FROM dual WHERE NOT EXISTS (SELECT * FROM 工人 WHERE 姓名 = %s)"

    # 执行 SQL 语句
    cursor.execute(sql, (row['姓名'], row['性别'], row['工种'], row['电话'], row['家庭住址'], row['身份证'],
                         row['合同签订时间'], row['进场时间'], row['银行卡号'], row['建行开户行'], row['姓名']))

    # 提交事务
    conn.commit()

# 关闭游标和连接
cursor.close()

conn.close()
