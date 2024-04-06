import openpyxl


def change_cake_first(msg_list):
    # msg_list:['产品名称', '批量', '抽 (送) 样日期', '商标', '型号规格',	'生产日期 (批号)', '样品数量',
    #                                     '验讫日期', '水分，％', '细菌总数 (CFU/g)', '大肠菌群 (CFU/g)', '时间', '审核人', '主检人']

    # 打开Excel文件
    workbook = openpyxl.load_workbook('蛋糕出厂报告（第一份）模板.xlsx')

    # 选择要修改的工作表
    sheet = workbook['Table1']  # 将'Sheet1'替换为你的工作表名称

    # 选择要修改的单元格
    sheet['B2'].value = msg_list[0]  # 填写产品名称
    sheet['B3'].value = msg_list[1]  # 填写批量
    sheet['B4'].value = msg_list[2]  # 填写抽 (送) 样日期
    sheet['B5'].value = msg_list[3]  # 填写商标
    sheet['D2'].value = msg_list[4]  # 填写型号规格
    sheet['D3'].value = msg_list[5]  # 填写生产日期 (批号)
    sheet['D4'].value = msg_list[6]  # 填写样品数量
    sheet['D5'].value = msg_list[7]  # 填写验讫日期
    sheet['B12'].value = msg_list[8]  # 填写水分，％
    sheet['B13'].value = msg_list[9]  # 填写细菌总数 (CFU/g)
    sheet['B14'].value = msg_list[10]  # 填写大肠菌群 (CFU/g)
    sheet['B18'].value = msg_list[11]  # 填写时间

    # 保存修改后的Excel文件
    workbook.save(f'{msg_list[5].strip(" ")}{msg_list[0]}蛋糕出厂报告（第一份）.xlsx')
    print("保存成功！")
