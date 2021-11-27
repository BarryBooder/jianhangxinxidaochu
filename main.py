import xlsxwriter

workbook = xlsxwriter.Workbook('jhkdx.xlsx')  

worksheet = workbook.add_worksheet()  

worksheet.write('A1', '序号')  
worksheet.write('B1', '卡尾号')  
worksheet.write('C1', '日期和时间')  
worksheet.write('D1', '金额')  
worksheet.write('E1', '卡内余额') 
worksheet.write('F1', '收支状态')  
worksheet.write('G1', '谁')  

KaHao = ""
RiQi = ""
JinE = ""
YuE = ""
ShouZhi = ""
Shui = ""


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False


with open("jhkdx.txt", "r", encoding='utf-8') as f:  # 打开文件
    row = 1
    for line in f.readlines():
        line = line.strip('\n')
        # 卡号
        KaHao = line[(line.index("您尾号") + 3):(line.index("储蓄卡") - 1)]
        # 日期
        if line.index("您尾号") != 0:
            for i in range(len(line)):
                if is_number(line[i]):
                    RiQi = line[i:(line.index("向您"))]
                    break
        else:
            RiQi = line[(line.index("储蓄卡") + 3):(line.index("分") + 1)]
        # 金额
        JinE = line[(line.index("人民币") + 3):(line.index("元") - 1)]
        # 余额
        YuE = line[(line.index("活期余额") + 4):(line.index("元。") - 1)]
        # 收支
        ShouZhi = line[(line.index("人民币") - 2):line.index("人民币")]
        # 谁
        if line.index("您尾号") != 0:
            for i in range(len(line)):
                if is_number((line[i])):
                    Shui = line[:i]
                    break
        else:
            flag = 0
            temp_i = 0
            for i in range(len(line)):
                if line[i] == "向":
                    flag = 1
                    temp_i = i
            if flag == 1:
                Shui = line[(line.index("向") + 1):temp_i + 4]
                Shui += line[temp_i + 4:line.index("人民币")]
            else:
                Shui = line[(line.index("分") + 1):line.index("人民币")]

        worksheet.write(row, 0, row)
        worksheet.write(row, 1, KaHao)
        worksheet.write(row, 2, RiQi)
        worksheet.write(row, 3, JinE)
        worksheet.write(row, 4, YuE)
        worksheet.write(row, 5, ShouZhi)
        worksheet.write(row, 6, Shui)

        row += 1
workbook.close()
