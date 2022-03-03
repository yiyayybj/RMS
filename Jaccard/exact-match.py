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

    """
        精确匹配
    """

    def exact_match(self):   # 输入：客户商品编号、船舶编号、客户账号
        a = Preprocess()
        khspbunmber = a.pro_list("../data/train_test/test_data.xlsx", "客户商品编号")
        boat_number = a.pro_list("../data/train_test/test_data.xlsx", "船舶编号")
        customer_number = a.pro_list("../data/train_test/test_data.xlsx", "客户账号")

        outwuliaonumber = a.pro_list("../data/精确匹配1000.xlsx", "公共外部物料编号")
        wuliaonumber = a.pro_list("../data/精确匹配1000.xlsx", "物料编号")
        boat_number_match = a.pro_list("../data/精确匹配1000.xlsx", "船舶编号")
        customer_number_match = a.pro_list("../data/精确匹配1000.xlsx", "客户账号")

        start_time = time.time()
        yuce_number_all = []   #精确匹配预测的结果
        flag = 0
        for i in khspbunmber:
            s = 0
            flag = flag + 1
            yuce_number = []
            index = []
            if (i == 'nan'):
                yuce_number_all.append(-1)  #客户商品编号空
                continue

            # 查看精确匹配表中 公共外部物料编号 与 客户商品编号 相同的产品，即为精确匹配预测的产品
            for k in outwuliaonumber:
                if (i == k):
                    yuce_number.append(wuliaonumber[s])  # 在精确匹配表中预测物料编号
                    index.append(s)  # 保存精确匹配表中预测物料的行索引
                s = s + 1

            # 对预测物料编号列表去重
            yuce_number_only = []
            for n in yuce_number:
                if n not in yuce_number_only:
                    yuce_number_only.append(n)

            if (len(yuce_number_only) == 1):
                yuce_number_all.append(yuce_number_only[0])
            elif (len(yuce_number_only) == 0):
                yuce_number_all.append(-1)  # 没有找到预测物料代码
            else:
                flag01 = 0

                for l in index:
                    if (boat_number[flag - 1] == boat_number_match[l]):  # 根据船舶编号匹配
                        yuce_number_all.append(wuliaonumber[l])
                        flag01 = 1
                        break
                m = 0
                if (flag01 == 0):  # 根据客户账号去匹配
                    for l in index:
                        if (customer_number[flag - 1] == customer_number_match[l]):
                            yuce_number_all.append(wuliaonumber[l])
                            m = 1
                            break
                if (flag01 == 0 and m == 0): #根据船舶编号和客户账号都没有相应的匹配，返回客户账号为空对应的物料编号即为预测物料编号
                    yuce_number_all.append(wuliaonumber[index[0]])

        end_time = time.time()
        print("耗时为{}秒".format(round(end_time - start_time, 4)))
        return yuce_number_all



a = Jaccard()
a.exact_match()
