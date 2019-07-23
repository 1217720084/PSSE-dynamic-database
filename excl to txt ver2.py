import xlrd

'''def strs(row):
    """
    :返回一行数据
    """
    try:
        values = "";
        for i in range(len(row)):
            if i == len(row) - 1:
                values = values + str(row[i]) + "\n"
            else:
                #使用“ ”逗号作为分隔符
                values = values + str(row[i]) + " " 
        return values
    except:
        raise
def xls_to_txt(xls_name,txt_name):
    """
    :excel文件转换为txt文件
    :param xls_name excel 文件名称
    :param txt_name txt   文件名称
    """
    try:
        data = xlrd.open_workbook(xls_name)
        sqlfile = open(txt_name, "a") 
        table = data.sheets()[0] # 表头
        nrows = table.nrows  # 行数
        #如果不需跳过表头，则将下一行中1改为0
        for ronum in range(1, nrows):
            row = table.row_values(ronum)
            values = strs(row) # 条用函数，将行数据拼接成字符串
            sqlfile.writelines(values) #将字符串写入新文件
        sqlfile.close() # 关闭写入的文件
    except:
        pass
if __name__ == '__main__':
    xls_name = 'f:/test.xls'
    txt_name = 'f:/test.txt'
    xls_txt(xls_name,txt_name)'''

'''def row_to_str(row_data):
    values = ""#形参，作为存放单行数据的string
    for j in range(len(current_row)):
        if j == len(current_row) - 1:
            #当到达该行最后一列时，添加“/”隔断
            values = values + str(current_row[j]) + "/" + "\n"
        else:
            #每个数据使用“ ”(空格)作为分隔符
            values = values + str(current_row[j]) + " " 
    return values
   
file.writelines(row_values + "\r")
# 关闭写入的文件
file.close()'''
    
def row_to_str(row_data):
    #row_data为形参，接受当前行数据
    values = ""#形参，作为存放单行数据的string
    for c in range(2,len(row_data)):
        if c == len(row_data) - 1:
            #当到达该行最后一列时，添加“/”隔断
            values = values + str(row_data[c]) + "/" + "\n"
        elif c == 2:
            #当达到母线列时，在类型左右添加（‘)（’）
            values = values + str(int(row_data[c])) + " "
        elif c == 3:
            #当达到类型列时，在类型左右添加（‘)（’）
            values = values +  "'" + str(row_data[c]) + "'"+" "
        else:
            #每个数据使用“ ”(空格)作为分隔符
            values = values + str(row_data[c]) + " " 
    return values

# 打开文件
all_data = xlrd.open_workbook('psse dynamic database.xls')#打开excel文件

table1 = all_data.sheet_by_name(u'generator')#通过名称获取excel表
table2 = all_data.sheet_by_name(u'exciter')
table3 = all_data.sheet_by_name(u'governor')
table4 = all_data.sheet_by_name(u'PSS')
table5 = all_data.sheet_by_name(u'compensator')

row_numb_gen = table1.nrows    # 发电机database行数
'''col_count_gen = table1.ncols    # 发电机database列数'''
row_numb_exc = table2.nrows    # 励磁调节器database行数
row_numb_gov = table3.nrows    # 调速器database行数
row_numb_pss = table4.nrows    # 稳定器database行数
row_numb_com = table5.nrows    # 补偿器database行数

'''title = table1.row_values(0)# 第一行数据(参数含义)'''
all_value = ""

for i in range(1, row_numb_gen):
    
    current_row_gen = table1.row_values(i)
    # 调用函数，将行数据拼接成字符串
    UID_gen = current_row_gen[1]
    #获得当前发电机的UID
    row_values = row_to_str(current_row_gen)
    # 将字符串写入新文件
    all_value += row_values
	
    for j in range(1, row_numb_exc):
        current_row_exc = table2.row_values(j)
        UID_exc = current_row_exc[1]
        if UID_exc == UID_gen:
            row_values = row_to_str(current_row_exc)
            all_value += row_values

        for k in range(1, row_numb_gov):
            current_row_gov = table3.row_values(k)
            UID_gov = current_row_gov[1]
            if UID_gov == UID_gen:
                row_values = row_to_str(current_row_gov)
                all_value += row_values

            for l in range(1, row_numb_pss):
                current_row_pss = table4.row_values(l)
                UID_pss = current_row_pss[1]
                if UID_pss == UID_gen:
                    row_values = row_to_str(current_row_pss)
                    all_value += row_values

                for m in range(1, row_numb_com):
                    current_row_com = table5.row_values(m)
                    UID_com = current_row_com[1]
                    if UID_com == UID_gen:
                        row_values = row_to_str(current_row_com)
                        all_value += row_values
    
file = open('test_for excel_to_txt.dyr', 'w')
file.write(all_value)
file.close()    
    