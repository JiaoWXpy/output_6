import pandas as pd
import csv
import numpy as np
def count_list_heng(data_list):
    """
    将集合均分，每份n个元素
    :param list_collection:
    :param n:
    :return:返回的结果为评分后的每份可迭代对象
    """
    output = []
    count_0 = 0
    list3 = return_list3_heng(0, data_list)
    #print("first 3:{}".format(list3))
    list3_all = []
    # print(list_4)
    for i, data in enumerate(data_list):
        if i < 1:
            output.append(None)
            list3_all.append("".join(list3))
        else:
            if not str_in_list3_heng(data, list3):
                output.append(0)
                count_0 += 1
            else:
                output.append(count_0 + 1)
                count_0 = 0
                list3_tmp = return_list3_heng(i, data_list)
                if list3_tmp != False:
                    list3 = list3_tmp
            # print(output)
            list3_all.append("".join(list3))
    return output, list3_all


def return_list3_heng(i, data_list):
    list5 = []
    for t in data_list[i]:

        list5.append(t)
    # print(list5)
    list5.reverse()
    list3 = []
    for i in list5:
        if i not in list3:
            list3.append(i)
        if len(list3) >= 3:
            break
    if len(list3) == 3:
        return list3
    else:
        return False


def str_in_list3_heng(str_data, list_3):

    if str_data[-1] in list_3:
        return True
    else:
        return False

def count_list_xie(data_list):
    """
    将集合均分，每份n个元素
    :param list_collection:
    :param n:
    :return:返回的结果为评分后的每份可迭代对象
    """
    output = []
    count_0 = 0
    list3 = return_list3_xie(5, data_list)
    list3_all = []
    # print(list_4)
    for i, data in enumerate(data_list):
        if i < 5:
            output.append(None)
            list3_all.append(None)
        else:
            if not str_in_list3_xie(data, list3):
                output.append(0)
                count_0 += 1
            else:
                output.append(count_0 + 1)
                count_0 = 0
                list3_tmp = return_list3_xie(i + 1, data_list)
                if list3_tmp != False:
                    list3 = list3_tmp
            # print(output)
            list3_all.append("".join(list3))
    return output, list3_all


def return_list3_xie(i, data_list):
    list5 = data_list[i - 5:i]
    # print(list5)
    list5.reverse()
    list3 = []
    l = -1
    for i in list5:
        if i[l] not in list3:
            list3.append(i[l])
        if len(list3) >= 3:
            break
        l -= 1
    if len(list3) == 3:
        return list3
    else:
        return False


def str_in_list3_xie(str_data, list_3):
    in_list = 0
    if str_data[-1] in list_3:
        return True
    else:
        return False

def read_data():
    print("请输入文件，注意原始文件每个表中不能添加表头。")
    # fp = input("请输入文件名：")
    fp = '1000.xlsx'
    excel_writer = pd.ExcelWriter("output_" + fp)  # 定义writer，选择文件（文件可以不存在）
    print("正在读取原始数据...")
    excel_reader = pd.ExcelFile(fp)  # 指定文件
    sheet_names = excel_reader.sheet_names  # 读取文件的所有表单名，得到列表
    print(sheet_names)

    for sheet in sheet_names:
        print("正在处理:{}".format(sheet))
        df_data = excel_reader.parse(sheet_name=sheet, header=None, dtype=str)
        data_list = []
        for i in df_data.values:
            data_list.append(str(i[0]))

        return data_list


def read_rule():
    rule_file = "./rule/rule.csv"
    f = csv.reader(open(rule_file, 'r'))
    rule_list_str = []
    for i in f:
        rule_list_str.append(i)
    # print(rule_list_str[1:])
    rule_list = []

    for i in rule_list_str[1:]:

        temp_lst = [int(i[1]), int(i[2])]
        rule_list_1 = [int(x) for x in i[3:]]
        temp_lst.append(rule_list_1)
        rule_list.append(temp_lst)
    return rule_list

def return_detial_rule(rule_list):
    new_rule_list = []
    for i in rule_list:
        new_rule_list.append(int(i))



    return new_rule_list
def save_df(dict,fp):
    excel_writer = pd.ExcelWriter("output_" + fp)  # 定义writer，选择文件（文件可以不存在）
    save_df = pd.DataFrame.from_dict(dict)
    save_df.to_excel(excel_writer, sheet_name="Result", index=False)  # 写入指定表单
    print("正在保存输出文件...")
    excel_writer.save()  # 储存文件
    excel_writer.close()  # 关闭writer
    print("保存成功！")
if __name__ == '__main__':
    #读取规则文件
    rule_list = read_rule()
    #读取原始数据
    data_list = read_data()
    #print(data_list)
    print(return_list3_xie(6,data_list))


    #记录结果的字典
    result = {"raw":[],"count_3":[],"result":[],"rule":[]}
    for rule in rule_list:
        current_data_all = data_list[rule[0]-1:rule[1]]
        #print(len(current_data_all))
        #遍历第n条规则的数值
        list3 = return_list3_heng(0, current_data_all)#第一个三位数字
        rule_map = return_detial_rule(rule[2])#
        print(rule_map)
        rule_count = 0#规则计数
        count_0 = 0
        currnt_rule = 0
        in_list3_heng = False
        for i, data in enumerate(current_data_all):
            result["raw"].append(data)
            #print(i,data)

            #第一个数字，直接在结果增加None
            if i < 1:
                result["result"].append(None)
                result["count_3"].append("".join(list3))
                result["rule"].append("无")
            #从第二个数字开始
            else:
                #横向
                if rule_map[currnt_rule] > 0:
                    if rule_count == 0 and not in_list3_heng:
                        list3_tmp = return_list3_heng(i -1, data_list)
                        if list3_tmp != False:
                            list3 = list3_tmp
                    if not str_in_list3_heng(data, list3):
                        result["result"].append(0)
                        count_0 += 1
                        in_list3_heng = True
                    else:
                        result["result"].append(count_0 + 1)
                        count_0 = 0
                        list3_tmp = return_list3_heng(i, current_data_all)
                        if list3_tmp != False:
                            list3 = list3_tmp
                        rule_count += 1

                        if rule_count >= np.abs(rule_map[currnt_rule]):
                            rule_count = 0
                            currnt_rule += 1
                        if currnt_rule >= len(rule_map):
                            currnt_rule = 0
                        in_list3_heng = False
                    # print(output)
                    result["count_3"].append("".join(list3))
                    result["rule"].append("横")
                    in_list3_xie = False
                #斜向
                elif rule_map[currnt_rule] < 0:
                    if rule_count == 0 and not in_list3_xie:
                        list3_tmp = return_list3_xie(i, data_list)
                        if list3_tmp != False:
                            list3 = list3_tmp
                    if not str_in_list3_xie(data, list3):
                        result["result"].append(0)
                        count_0 += 1
                        in_list3_xie = True
                    else:
                        result["result"].append(count_0 + 1)
                        count_0 = 0

                        list3_tmp = return_list3_xie(i +1, current_data_all)
                        if list3_tmp != False:
                            list3 = list3_tmp
                        rule_count += 1

                        if rule_count >= np.abs(rule_map[currnt_rule]):
                            rule_count = 0
                            currnt_rule += 1
                        if currnt_rule >= len(rule_map):
                            currnt_rule = 0
                        in_list3_xie = False
                    # print(output)
                    result["count_3"].append("".join(list3))
                    result["rule"].append("斜")
                    in_list3_heng = False




        print("result count:{}".format(len(result["result"])))
        print(result["result"])
        #print(result)
        print(len(result["raw"]),len(result["count_3"]),len(result["result"]),len(result["rule"]))
    #print(result)
    save_df(result,'result.xlsx')
