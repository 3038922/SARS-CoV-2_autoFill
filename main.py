# encoding='UTF-8'
import sys
import openpyxl  # excel操作模块
import configparser  # 导入配置.ini模块


# 生成列名字典，只是为了方便修改列宽时指定列，key:数字，从1开始；value:列名，从A开始
def get_num_colnum_dict():
    '''
    :return: 返回字典：{1:'A', 2:'B', ...... , 52:'AZ'}
    '''
    num_str_dict = {}
    A_Z = [chr(a) for a in range(ord('A'), ord('Z') + 1)]
    AA_AZ = ['A' + chr(a) for a in range(ord('A'), ord('Z') + 1)]
    A_AZ = A_Z + AA_AZ
    for i in A_AZ:
        num_str_dict[A_AZ.index(i) + 1] = i
    return num_str_dict


# 自适应列宽
def style_excel(self, target_wb, sheet_name):
    '''
    :param sheet_name:  excel中的sheet名
    :return:
    '''
    # 打开excel
    wb = target_wb
    # 选择对应的sheet
    sheet = wb[sheet_name]
    # 获取最大行数与最大列数
    max_column = sheet.max_column
    max_row = sheet.max_row

    # 将每一列，单元格列宽最大的列宽值存到字典里，key:列的序号从1开始(与字典num_str_dic中的key对应)；value:列宽的值
    max_column_dict = {}

    # 生成列名字典，只是为了方便修改列宽时指定列，key:数字，从1开始；value:列名，从A开始
    num_str_dict = get_num_colnum_dict()

    # 遍历全部列
    for i in range(1, max_column + 1):
        # 遍历每一列的全部行
        for j in range(1, max_row + 1):
            column = 0
            # 获取j行i列的值
            sheet_value = sheet.cell(row=j, column=i).value
            # 通过列表生成式生成字符列表，将当前获取到的单元格的str值的每一个字符放在一个列表中（列表中一个元素是一个字符）
            sheet_value_list = [k for k in str(sheet_value)]
            # 遍历当前单元格的字符列表
            for v in sheet_value_list:
                # 判定长度，一个数字或一个字母，单元格列宽+=1.1，其它+=2.2（长度可根据需要自行修改，经测试一个字母的列宽长度大概为1）
                if v.isdigit() == True or v.isalpha() == True:
                    column += 2.3
                else:
                    column += 1.0
            # 当前单元格列宽与字典中的对比，大于字典中的列宽值则将字典更新。如果字典没有这个key，抛出异常并将值添加到字典中
            try:
                if column > max_column_dict[i]:
                    max_column_dict[i] = column
            except Exception as e:
                max_column_dict[i] = column
    # 此时max_column_dict字典中已存有当前sheet的所有列的最大列宽值，直接遍历字典修改列宽
    for key, value in max_column_dict.items():
        sheet.column_dimensions[num_str_dict[key]].width = value


def main(argv=None):
    config = configparser.ConfigParser()  # 类实例化
    # 定义文件路径
    path = 'config.ini'
    config.read(path, encoding='UTF-8')
    inputFileName = config.get('select', 'inputFileName')
    outputFilename = config.get('select', 'outputFilename')
    学校名称 = config.get('select', 'schoolName')
    内地学生数 = config.getint('select', 'mainlandStu')
    港澳台学生数 = config.getint('select', 'hkMacaoTwStu')
    留学生 = config.getint('select', 'overseasStu')
    合计 = 内地学生数 + 港澳台学生数 + 留学生
    应到人数 = config.getint('select', 'arriveStu')
    # 读取表格
    wb = openpyxl.load_workbook(inputFileName)
    # 显示所有表单
    dict1 = {"总数": 0, "体温异常": 0, "其他异常": 0, "详情": [["班级", "姓名", "是否体温异常", "晨检异常情况描述"]]}
    dict2 = {"总数": 0, "详情": [["班级", "姓名", "文字描述"]]}
    dict3 = {"总数": 0, "详情": [["班级", "姓名", "文字描述"]]}
    dict4 = {
        "总数": 0,
        "因发热未到或请假": 0,
        "因咳嗽腹泄乏力请假人数": 0,
        "因其他原因请假人数": 0,
        "详情": [["班级", "姓名", "是否发热", "是否咳嗽乏力腹泻呕吐等", "是否有重点疫区接触史", "请假原因"]]
    }
    班级 = ""
    晨检午检异常描述 = ""
    乘坐公交车学生人数 = 0
    for cell in wb["主表"]["J"]:
        if (cell.row > 2):
            乘坐公交车学生人数 += int(cell.value)

    for cell in wb["晨检异常登记"]["C"]:
        if (cell.value != "") and (cell.value != "姓名"):
            dict1["总数"] += 1
            if (wb["晨检异常登记"].cell(cell.row, 4).value != "否") and (wb["请假学生登记"].cell(
                    cell.row, 4).value != "无"):
                dict1["体温异常"] += 1
            for it in wb["主表"]["R"]:
                if (str(it.value).find(cell.value) >= 0):
                    班级 = str(wb["主表"].cell(it.row, 6).value) + \
                        str(wb["主表"].cell(it.row, 7).value)
            晨检午检异常描述 += str(班级 + cell.value + wb["晨检异常登记"].cell(cell.row, 5).value) + "\n"
            dict1["详情"].append([
                班级, cell.value, wb["晨检异常登记"].cell(cell.row, 4).value,
                wb["晨检异常登记"].cell(cell.row, 5).value
            ])
    for cell in wb["学生接触外省人员登记"]["C"]:
        if (cell.value != "") and (cell.value != "姓名"):
            dict2["总数"] += 1
            for it in wb["主表"]["R"]:
                if (str(it.value).find(cell.value) >= 0):
                    班级 = str(wb["主表"].cell(it.row, 6).value) +\
                        str(wb["主表"].cell(it.row, 7).value)
            dict2["详情"].append([班级, cell.value, wb["学生接触外省人员登记"].cell(cell.row, 4).value])
    for cell in wb["家长接触重点疫区、境外登记"]["C"]:
        if (cell.value != "") and (cell.value != "姓名"):
            dict3["总数"] += 1
            for it in wb["主表"]["R"]:
                if (str(it.value).find(cell.value) >= 0):
                    班级 = str(wb["主表"].cell(it.row, 6).value) +\
                        str(wb["主表"].cell(it.row, 7).value)
            dict3["详情"].append([班级, cell.value, wb["家长接触重点疫区、境外登记"].cell(cell.row, 4).value])
    for cell in wb["请假学生登记"]["C"]:
        if (cell.value != "") and (cell.value != "姓名"):
            dict4["总数"] += 1
            if (wb["请假学生登记"].cell(cell.row, 4).value != "否") and (wb["请假学生登记"].cell(
                    cell.row, 4).value != "无"):
                dict4["因发热未到或请假"] += 1
            if (wb["请假学生登记"].cell(cell.row, 5).value != "否") and (wb["请假学生登记"].cell(
                    cell.row, 5).value != "无"):
                dict4["因咳嗽腹泄乏力请假人数"] += 1
            for it in wb["主表"]["R"]:
                if (str(it.value).find(cell.value) >= 0):
                    班级 = str(wb["主表"].cell(it.row, 6).value) +\
                        str(wb["主表"].cell(it.row, 7).value)
            dict4["详情"].append([
                班级, cell.value, wb["请假学生登记"].cell(cell.row,
                                                  4).value, wb["请假学生登记"].cell(cell.row, 5).value,
                wb["请假学生登记"].cell(cell.row, 6).value, wb["请假学生登记"].cell(cell.row, 7).value
            ])
        else:
            wb["请假学生登记"].delete_rows(cell.row)
    dict1["其他异常"] = dict1["总数"] - dict1["体温异常"]
    dict4["因其他原因请假人数"] = dict4["总数"] - dict4["因发热未到或请假"] - dict4["因咳嗽腹泄乏力请假人数"]
    # EXCEL 新表
    wbNEW = openpyxl.Workbook()
    wbNEW.active.title = "晨检异常"
    wbNEW.create_sheet("学生接触外省人员")
    wbNEW.create_sheet("家长接触重点疫区、境外")
    wbNEW.create_sheet("请假学生")
    wbNEW.create_sheet("综合统计")
    for cell in dict1["详情"]:
        wbNEW["晨检异常"].append(cell)
    for cell in dict2["详情"]:
        wbNEW["学生接触外省人员"].append(cell)
    for cell in dict3["详情"]:
        wbNEW["家长接触重点疫区、境外"].append(cell)
    for cell in dict4["详情"]:
        wbNEW["请假学生"].append(cell)
    wbNEW["综合统计"].append([
        "学校名称", "数据为后三列之和，即内地学生+港澳台+留学生", "内地学生数", "港澳台学生数", "留学生数", "合计中的当日校内发热学生人数",
        "应到人数（要求到校人数）", "实到人数（晨检人数）", "因发热未到或请假（体温37.3或以上）", "因其他原因未到或请假（除发热外）", "校内晨检异常人数（体温异常）",
        "校内晨检异常其他人数（如：咳嗽、腹泻等）", "当日校内因发热送检人数（体温37.3或以上）", "总计", "绿码人数", "黄码人数", "橙码人数", "未申领人数",
        "接触过省外返衢对象的(学生数)", "接触过境外返衢对象的(学生数)", "与疫情中高风险地区、境外回衢人员接触的家长人数统计", "乘坐公交车学生人数"
    ])
    wbNEW["综合统计"].append([
        学校名称, 合计, 内地学生数, 港澳台学生数, 留学生, 应到人数 - dict4["总数"], 应到人数, dict4["总数"], dict4["因发热未到或请假"],
        dict4["因咳嗽腹泄乏力请假人数"] + dict4["因其他原因请假人数"], dict1["总数"], dict1["体温异常"], dict1["其他异常"],
        晨检午检异常描述, dict2["总数"], dict3["总数"], 乘坐公交车学生人数
    ])

    # 设置自动列宽
    style_excel(self=None, target_wb=wbNEW, sheet_name="综合统计")
    wbNEW.save(outputFilename)


if __name__ == "__main__":
    sys.exit(main())
