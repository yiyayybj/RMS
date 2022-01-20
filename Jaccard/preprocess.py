import os
import re

import jieba
import pandas as pd
from nltk import SnowballStemmer

c = []

class Preprocess():
    def __init__(self):
        self.stopwords=[]
        self.current_path = os.path.abspath(__file__)
        self.father_path = os.path.abspath(os.path.dirname(self.current_path) + os.path.sep + ".")

    def strQ2B(self, ustring):
        """全角转半角"""
        rstring = ""
        for uchar in ustring:
            inside_code = ord(uchar)
            if inside_code == 12288:  # 全角空格直接转换
                inside_code = 32
            elif (inside_code >= 65281 and inside_code <= 65374):  # 全角字符（除空格）根据关系转化
                inside_code -= 65248

            rstring += chr(inside_code)
        return rstring

    def pre_pro(self, x):
        stemmer = SnowballStemmer("english")
        global c
        if re.search('[\u4E00-\u9FA5]+', x):
            a = re.findall('\d+\s[\./]\s\d+|\d+[\./]\d+|\d+\w+|\w+\d+\.?\d?|[\u4E00-\u9FA5a-z\d]+',
                           ' '.join(jieba.cut(self.strQ2B(x.lower()))))
            jieba.setLogLevel(jieba.logging.INFO)
            b = []
            for i in a:
                i = stemmer.stem(i)
                if i not in b:
                    b.append(i)
            c = [s for s in b if not re.search(r'\d', s)]
        else:
            a = re.findall('\d+\s[\./]\s\d+|\d+[\./]\d+|\d+\w+|\w+\d+\.?\d?|[\u4E00-\u9FA5a-z\d]+',
                           self.strQ2B(x.lower()))
            b = []
            for i in a:
                i = stemmer.stem(i) # 词干化处理
                if i not in b:
                    b.append(i)
            c = [s for s in b if not re.search(r'\d', s)] # 去除英文中包含的数字
        # print(c)
        return c

    def pro_list(self,file,x):
        data = pd.read_excel(file)
        data[x] = data[x].astype('str')
        back = data[x].values.tolist()
        return back

    def pre_process(self,file,x):
        back = Preprocess.pro_list(self,file,x)
        a = []
        for i in back:
            a.append(Preprocess.pre_pro(self,i))
        return a

# a = Preprocess()
# a.pre_pro("NOTEBOOK SCHOOL USE A-5")
# a.pre_pro("NOTEBOOK SOFT COVER_A5 80PAGES 210*148MM_DELI_7654")
# print(type(a.pre_process("../data/客户描述_物料1000.xlsx","客户商品英文描述")))
# print(a.pre_process("../data/客户描述_物料1000.xlsx","客户商品英文描述"))