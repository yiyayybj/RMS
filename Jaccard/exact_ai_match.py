import os
import time

import xlwt
import heapq

from Jaccard.preprocess import Preprocess

# 产品表做distinct后进行匹配
class Jaccard():
    def __init__(self):
        self.stopwords=[]
        self.current_path = os.path.abspath(__file__)
        self.father_path = os.path.abspath(os.path.dirname(self.current_path) + os.path.sep + ".")

    def getListMaxNumIndex(testList, topk):
        '''
             获取列表中最大的前n个数值的位置索引
        '''
        tmp = zip(range(len(testList)), testList)
        large5 = heapq.nlargest(topk, tmp, key=lambda x: x[1])
        return large5

    def getkhdescribe(self):
        a = Preprocess()
        khdescribe = a.pre_process("../data/train_test/test_data.xlsx", "客户商品英文描述")
        khdescribe01 = a.pro_list("../data/train_test/test_data.xlsx", "客户商品英文描述")
        khccp = a.pro_list("../data/train_test/test_data.xlsx","我方物料中文名称")
        khecp = a.pro_list("../data/train_test/test_data.xlsx","我方物料英文名称")
        khnumber = a.pro_list("../data/train_test/test_data.xlsx","我方物料编号")
        khspbunmber = a.pro_list("../data/train_test/test_data.xlsx","客户商品编号")
        boat_number = a.pro_list("../data/train_test/test_data.xlsx","船舶编号")
        customer_number = a.pro_list("../data/train_test/test_data.xlsx", "客户账号")

        outwuliaonumber = a.pro_list("../data/精确匹配1000.xlsx", "公共外部物料编号")
        wuliaonumber = a.pro_list("../data/精确匹配1000.xlsx", "物料编号")
        boat_number_match = a.pro_list("../data/精确匹配1000.xlsx", "船舶编号")
        customer_number_match = a.pro_list("../data/精确匹配1000.xlsx", "客户账号")

        spedescribe = a.pre_process("../data/产品分类后.xlsx", "ITEMENGNAME")
        spedescribe01 = a.pro_list("../data/产品分类后.xlsx", "ITEMENGNAME")
        # spcdescribe = a.pro_list("../data/产品表_分类后.xlsx","商品中文描述")
        number = a.pro_list("../data/产品分类后.xlsx", "ITEMID")

        flag = 0
        count = 0
        filename = xlwt.Workbook()
        # 给工作表命名，test
        sheet = filename.add_sheet("test")
        column_name = ['客户商品英文描述','实际商品中文名','实际商品英文名','实际物料代码','客户商品编号',
                       '船舶编号','客户账号','公共外部物料编号','预测物料编号','精确匹配',
                       '根据船舶编号匹配或根据客户账号匹配(1为根据船舶编号匹配，2为根据客户账号匹配)',
                       '预测商品英文名1', '预测物料代码1', '匹配相似度1',
                       '预测商品英文名2', '预测物料代码2', '匹配相似度2', '预测商品英文名3', '预测物料代码3', '匹配相似度3',
                       '预测商品英文名4', '预测物料代码4', '匹配相似度4', '预测商品英文名5', '预测物料代码5', '匹配相似度5', 'AI匹配'
                       ]
        row = 0

        for item in range(len(column_name)):
            sheet.write(row, item, column_name[item])
        start_time = time.time()

        for i in khspbunmber:
            flag = flag+1
            print(flag)
            c = 0  #判断精确匹配是否成功
            sheet.write(flag , 0, khdescribe01[flag-1])  # 客户商品英文描述
            sheet.write(flag , 1, khccp[flag-1])  # 实际商品中文名
            sheet.write(flag , 2, khecp[flag-1])  # 实际商品英文名
            sheet.write(flag , 3, khnumber[flag-1])  # 实际物料代码
            sheet.write(flag, 4, khspbunmber[flag - 1])  # 客户商品编号
            sheet.write(flag, 5, boat_number[flag - 1])  # 船舶编号
            sheet.write(flag, 6, customer_number[flag - 1])  # 客户账号

            s = 0
            yuce_number = []
            index = []
            if(i == 'nan'):
                sheet.write(flag, 7, '客户商品编号空')  # 客户商品编号
                continue

            #查看精确匹配表中 公共外部物料编号 与 客户商品编号 相同的产品，即为精确匹配预测的产品
            for k in outwuliaonumber:
                if(i == k):
                    yuce_number.append(wuliaonumber[s])  #在精确匹配表中预测物料编号
                    index.append(s)  #保存精确匹配表中预测物料的行索引
                s = s + 1

            # 对预测物料编号列表去重
            yuce_number_only = []
            for n in yuce_number:
                if n not in yuce_number_only:
                    yuce_number_only.append(n)

            if(len(yuce_number_only) == 1):
                sheet.write(flag, 7, outwuliaonumber[index[0]])  # 公共外部物料编号
                sheet.write(flag, 8, yuce_number_only[0])  # 预测物料编号
                if(khnumber[flag-1] == yuce_number_only[0]):
                    count = count+1
                    sheet.write(flag, 9, 1)  # 精确匹配
                else:
                    sheet.write(flag, 9, 0)  # 精确匹配
            elif(len(yuce_number_only) == 0):
                sheet.write(flag, 7, '没有找到')  # 公共外部物料编号
                sheet.write(flag, 8, '没有找到')  # 预测物料编号
                sheet.write(flag, 9, 0)  # 精确匹配
            else:
                flag01 = 0

                for l in index:
                    if (boat_number[flag - 1] == boat_number_match[l]):  # 根据船舶编号匹配
                        sheet.write(flag, 8, wuliaonumber[l])  # 预测物料编号
                        if(khnumber[flag-1] == wuliaonumber[l]):
                            count = count+1
                            sheet.write(flag, 9, 1)  # 精确匹配
                        else:
                            sheet.write(flag, 9, 0)  # 精确匹配
                        sheet.write(flag, 10, 1)  # 是否根据船舶编号匹配
                        flag01 = 1
                        break
                m = 0
                if(flag01 == 0): #根据客户账号去匹配
                    for l in index:
                        if(customer_number[flag-1] == customer_number_match[l]):
                            sheet.write(flag, 7, outwuliaonumber[l])  # 公共外部物料编号
                            sheet.write(flag, 8, wuliaonumber[l])  # 预测物料编号
                            if (khnumber[flag - 1] == wuliaonumber[l]):
                                sheet.write(flag, 9, 1)  # 精确匹配
                                count = count + 1
                            else:
                                sheet.write(flag, 9, 0)  # 精确匹配
                            sheet.write(flag, 10, 2)  # 是否根据客户账号匹配
                            m = 1
                            break
                if(flag01 == 0 and m == 0):
                    sheet.write(flag, 7, outwuliaonumber[index[0]])  # 公共外部物料编号
                    sheet.write(flag, 8, wuliaonumber[index[0]])  # 预测物料编号
                    if(khnumber[flag - 1] == wuliaonumber[index[0]]):
                        sheet.write(flag, 9, 1)  # 精确匹配
                        count = count + 1
                    else:
                        sheet.write(flag, 9, 0)  # 精确匹配
        print(count)
        print(flag)
        print(count / flag)

        flag = 0
        count = 0

        for kh_text in khdescribe:
            print(flag)
            kw = set(kh_text)
            simlar = []
            for item_text in spedescribe:
                item_text = set(item_text)
                temp = 0
                for i in item_text:
                    if i in kw:
                        temp = temp + 1
                fenmu = len(item_text) + len(kw) - temp  # 并集
                jaccard_coefficient = float(temp / fenmu)  # 交集
                simlar.append(jaccard_coefficient)

            top5 = Jaccard.getListMaxNumIndex(simlar, 5)
            s = 0

            j = 0
            for i in top5:
                # print(i[0])
                sheet.write(flag + 1, 11 + j * 3, spedescribe01[i[0]])  # 预测商品英文名1
                sheet.write(flag + 1, 12 + j * 3, number[i[0]])  # 预测物料代码
                sheet.write(flag + 1, 13 + j * 3, i[1])  # 匹配相似度
                j = j + 1
                number_list = number[i[0]].split(",")
                for k in number_list:
                    if (k == khnumber[flag]):
                        s = 1

            if (s == 1):
                count = count + 1
                sheet.write(flag + 1, 26, 1)  # 匹配成功
            else:
                sheet.write(flag + 1, 26, 0)  # 匹配失败
            flag = flag + 1

        end_time = time.time()
        print("耗时为{}秒".format(round(end_time - start_time, 4)))
        filename.save("../Jaccard_result/2.25_result_test_data_精确匹配_ai匹配.xls")
        print(count)
        print(flag)
        print(count/flag)

a = Jaccard()
a.getkhdescribe()

