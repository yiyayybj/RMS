import os
import time

import xlwt
import heapq

from Jaccard.preprocess import Preprocess

# 在所有产品中进行匹配
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
        khdescribe = a.pre_process("../data/客户描述_test1000.xlsx", "客户商品英文描述")
        khdescribe01 = a.pro_list("../data/客户描述_test1000.xlsx", "客户商品英文描述")
        khccp = a.pro_list("../data/客户描述_test1000.xlsx","我方物料中文名称")
        khecp = a.pro_list("../data/客户描述_test1000.xlsx","我方物料英文名称")
        khnumber = a.pro_list("../data/客户描述_test1000.xlsx","我方物料编号")

        spedescribe = a.pre_process("../data/产品表_账套1000.xlsx","商品英文描述")
        spedescribe01 = a.pro_list("../data/产品表_账套1000.xlsx", "商品英文描述")
        spcdescribe = a.pro_list("../data/产品表_账套1000.xlsx","商品中文描述")
        number = a.pro_list("../data/产品表_账套1000.xlsx", "物料编码")

        flag = 0
        count = 0
        filename = xlwt.Workbook()
        # 给工作表命名，test
        sheet = filename.add_sheet("test")
        column_name = ['客户商品英文描述','实际商品中文名','实际商品英文名','实际物料代码','预测商品中文名1','预测商品英文名1','预测物料代码1','匹配相似度1',
                       '预测商品中文名2','预测商品英文名2','预测物料代码2','匹配相似度2','预测商品中文名3','预测商品英文名3','预测物料代码3','匹配相似度3',
                       '预测商品中文名4','预测商品英文名4','预测物料代码4','匹配相似度4','预测商品中文名5','预测商品英文名5','预测物料代码5',
                       '匹配相似度5','AI匹配']
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
            sheet.write(flag+1, 0, khdescribe01[flag]) #客户商品英文描述
            sheet.write(flag + 1, 1, khccp[flag]) #实际商品中文名
            sheet.write(flag + 1, 2, khecp[flag]) #实际商品英文名
            sheet.write(flag+1, 3, khnumber[flag]) #实际物料代码
            j = 0
            for i in top5:
                #print(i[0])
                if (i[0] > 286582):
                    i[0] = int(i[0] / 2)
                sheet.write(flag+1, 4+j*4, spcdescribe[i[0]])  # 预测商品中文名1
                sheet.write(flag+1, 5+j*4, spedescribe01[i[0]])  # 预测商品英文名1
                sheet.write(flag+1, 6+j*4, number[i[0]])  # 预测物料代码
                sheet.write(flag + 1, 7 + j * 4, i[1]) #匹配相似度
                j = j+1
                if(number[i[0]] == khnumber[flag]):
                    s = 1

            if(s == 1):
                count = count + 1
                sheet.write(flag+1, 24, 1) #匹配成功
            else:
                sheet.write(flag+1, 24, 0) #匹配失败
            flag = flag + 1
        end_time = time.time()
        print("耗时为{}秒".format(round(end_time - start_time, 4)))
        filename.save("../Jaccard_result/1.18_result_test1000.xls")
        print(count)
        print(flag)
        print(count/flag)

a = Jaccard()
a.getkhdescribe()
# a.getkhdescribe("产品表_账套1000.xlsx","test01.xlsx")
# a.getkhdescribe("同一客户客户描述.xlsx","同一客户客户描述.xlsx")

