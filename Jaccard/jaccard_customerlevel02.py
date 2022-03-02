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
        khlevel = a.pro_list("../data/train_test/test_data.xlsx", "客户质量等级")

        spedescribe = a.pre_process("../data/产品分类后.xlsx","ITEMENGNAME")
        spedescribe01 = a.pro_list("../data/产品分类后.xlsx", "ITEMENGNAME")
        # spcdescribe = a.pro_list("../data/产品分类后.xlsx","商品中文描述")
        number = a.pro_list("../data/产品分类后.xlsx", "ITEMID")

        # spedescribe = a.pre_process("../data/产品主数据1000.xlsx", "ITEMENGNAME")
        # spedescribe01 = a.pro_list("../data/产品主数据1000.xlsx", "ITEMENGNAME")
        # spcdescribe = a.pro_list("../data/产品主数据1000.xlsx", "ITEMCHINESEDES")
        number_individual = a.pro_list("../data/产品主数据1000.xlsx", "ITEMID")
        splevel = a.pro_list("../data/产品主数据1000.xlsx", "QUALEVEL")


        flag = 0
        count = 0
        filename = xlwt.Workbook()
        # 给工作表命名，test
        sheet = filename.add_sheet("test")
        column_name = ['客户商品英文描述','实际商品中文名','实际商品英文名','实际物料代码',
                       '预测商品英文名1','预测物料代码1','匹配相似度1','预测物料代码缩小范围1',
                       '预测商品英文名2','预测物料代码2','匹配相似度2','预测物料代码缩小范围2',
                       '预测商品英文名3','预测物料代码3','匹配相似度3','预测物料代码缩小范围3',
                       '预测商品英文名4','预测物料代码4','匹配相似度4','预测物料代码缩小范围4',
                       '预测商品英文名5','预测物料代码5','匹配相似度5','预测物料代码缩小范围5','缩小范围后，实际物料代码在预测代码中',
                       'AI匹配']
        row = 0

        for item in range(len(column_name)):
            sheet.write(row, item, column_name[item])
        start_time = time.time()
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

            top5 = Jaccard.getListMaxNumIndex(simlar,5)
            s = 0
            sheet.write(flag + 1, 0, khdescribe01[flag]) #客户商品英文描述
            sheet.write(flag + 1, 1, khccp[flag]) #实际商品中文名
            sheet.write(flag + 1, 2, khecp[flag]) #实际商品英文名
            sheet.write(flag + 1, 3, khnumber[flag]) #实际物料代码
            j = 0
            for i in top5:
                #print(i[0])
                yuce_number = []
                sheet.write(flag + 1, 4 + j * 4, spedescribe01[i[0]])  # 预测商品英文名1
                sheet.write(flag + 1, 5 + j * 4, number[i[0]])  # 预测物料代码
                sheet.write(flag + 1, 6 + j * 4, i[1]) #匹配相似度
                number_list = number[i[0]].split(",")
                print(number_list)
                flag01 = 0
                for k in number_list:
                    l = 0
                    for wuliao in number_individual: #在所有产品里查询
                        if (k == wuliao):
                            if (khlevel[flag] == splevel[l]):
                                yuce_number.append(wuliao)
                            break
                        l = l + 1
                    if (k == khnumber[flag]):
                        s = 1
                print(yuce_number)
                sheet.write(flag + 1, 7 + j * 4, yuce_number)  # 预测代码 客户质量等级与商品等级匹配 缩小范围
                j = j + 1
                for m in yuce_number:
                    if(m == khnumber[flag]):
                        flag01 = 1
                        # sheet.write(flag + 1, 7 + j * 4, 1)  # 是否缩小了预测物料代码范围
            if(flag01 == 0):
                sheet.write(flag + 1, 24 , 0)  # 缩小范围后，实际物料代码不在预测代码中
            else:
                sheet.write(flag + 1, 24 , 1)  # 缩小范围后，实际物料代码在预测代码中

            if(s == 1):
                count = count + 1
                sheet.write(flag+1, 25, 1) #匹配成功
            else:
                sheet.write(flag+1, 25, 0) #匹配失败
            flag = flag + 1
        end_time = time.time()
        print("耗时为{}秒".format(round(end_time - start_time, 4)))
        filename.save("../Jaccard_result/2.24_result_customerlevel02.xls")
        print(count)
        print(flag)
        print(count/flag)

a = Jaccard()
a.getkhdescribe()

