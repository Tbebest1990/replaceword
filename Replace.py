#coding=utf-8
import os
import win32com
from win32com.client import Dispatch
import numpy as np
import pandas as pd
import datetime

#读取产品要素，负值给相关变量
Excelfilepath="E:\姜林\收益凭证\五矿证券优享2号100期\产品要素.xlsx"
df=pd.read_excel(Excelfilepath,sheet_name="Sheet1",index_col=0)

print(df)
#格式化时间
RegisterDate=df.loc['RegisterDate',"values"].strftime("%Y{y}%m{m}%d{d}").format(y='年',m='月',d='日')
SubscriptionDate=df.loc['SubscriptionDate',"values"].strftime("%Y{y}%m{m}%d{d}").format(y='年',m='月',d='日')
DueDate=df.loc['DueDate',"values"].strftime("%Y{y}%m{m}%d{d}").format(y='年',m='月',d='日')
PaymentDate=df.loc['PaymentDate',"values"].strftime("%Y{y}%m{m}%d{d}").format(y='年',m='月',d='日')
YM=df.loc['PaymentDate',"values"].strftime("%Y{y}%m{m}").format(y='年',m='月')


ProductName=df.loc['ProductName',"values"]
ForShort=df.loc['ForShort',"values"]
ForShort=df.loc['ForShort',"values"]
code=df.loc['code',"values"]
#格式化百分数
rate1=df.loc['rate0',"values"]
rate0="{:.2f}{}".format(rate1*100,"%")
rate2=df.loc['rate',"values"]
rate="{:.2f}{}".format(rate2*100,"%")

MaxScale=df.loc['MaxScale',"values"]
MinScale=df.loc['MinScale',"values"]
StartVol=df.loc['StartVol',"values"]
IncreaseVol=df.loc['IncreaseVol',"values"]
MaxVol=df.loc['MaxVol',"values"]
PPMaxVol=df.loc['PPMaxVol',"values"]
TimeQuantum=df.loc['TimeQuantum',"values"]
InterestPeriod=df.loc['InterestPeriod',"values"]
redeem=df.loc['redeem',"values"]
BuyBack=df.loc['BuyBack',"values"]
transfer=df.loc['transfer',"values"]
print(PaymentDate)


# 处理Word文档的类
time1="2020-03-12"
time2="2020-09-18"
class RemoteWord:

    def __init__(self,filename=None):

        self.xlApp = win32com.client.Dispatch('Word.Application')  # 此处使用的是Dispatch，原文中使用的DispatchEx会报错

        self.xlApp.Visible = 0  # 后台运行，不显
        self.xlApp.DisplayAlerts = 0 # 不警告
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.doc = self.xlApp.Documents.Open(filename)
            else:
                self.doc = self.xlApp.Documents.Add()  # 创建新的文档
                self.doc.SaveAs(filename)
        else:
            self.doc = self.xlApp.Documents.Add()
            self.filename = ''
    def add_doc_end(self, string):
      #在文档末尾添加内容

        rangee = self.doc.Range()
        rangee.InsertAfter('\n' + string)

    def add_doc_start(self, string):
        #在文档开头添加内容
        rangee = self.doc.Range(0,0)
        rangee.InsertBefore(string + '\n')
    def insert_doc(self, insertPos,string):
        #在文档insertPos位置添加内容
        rangee = self.doc.Range(0,insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n' + string)

    def replace_doc(self, string, new_string):
        #替换文字
        self.xlApp.Selection.Find.ClearFormatting()
        self.xlApp.Selection.Find.Replacement.ClearFormatting()
        # (string--搜索文本,
        # True--区分大小写,
        # True--完全匹配的单词，并非单词中的部分（全字匹配）,
        # True--使用通配符,
        # True--同音,
        # True--查找单词的各种形式,
        # True--向文档尾部搜索,
        # 1,
        # True--带格式的文本,
        # new_string--替换文本,
        # 2--替换个数（全部替换)
        self.xlApp.Selection.Find.Execute(string, False, False, False, False, False, True, 1, True, new_string, 2)


    def replace_docs(self, string, new_string):
        """采用通配符匹配替换"""
        self.xlApp.Selection.Find.ClearFormatting()
        self.xlApp.Selection.Find.Replacement.ClearFormatting()
        self.xlApp.Selection.Find.Execute(string, False, False, True, False, False, False, 1, False, new_string, 2)

    def save(self):
        #保存文档
        self.doc.Save()


    def save_as(self, filename):
        #文档另存为

        self.doc.SaveAs(filename)



    def close(self):
        #保存文件、关闭文件

        self.save()
        self.xlApp.Documents.Close()
        self.xlApp.Quit()

if __name__ == '__main__':

    # path = 'E:\\XXX.docx'
    path = 'E:\姜林\收益凭证\五矿证券优享2号100期\风险等级评估表-模板.docx'
    doc = RemoteWord(path) # 初始化一个doc对象
    # 这里演示替换内容，其他功能自己按照上面类的功能按需使用

    doc.replace_doc('古代', '中国') # 替换文本内容
    doc.replace_doc('ProductName', ProductName)
    doc.replace_doc('ForShort', ForShort)
    doc.replace_doc('code', code)
   # doc.replace_doc('place', place)
   #doc.replace_doc('client', client)
    doc.replace_doc('rate0', rate0)
    doc.replace_doc('rate', rate)
    doc.replace_doc('MaxScale', MaxScale)
    doc.replace_doc('MinScale', MinScale)
    doc.replace_doc('StartVol', StartVol)
    doc.replace_doc('IncreaseVol', IncreaseVol)
    doc.replace_doc('MaxVol', MaxVol)
    doc.replace_doc('PPMaxVol', PPMaxVol)
    doc.replace_doc('TimeQuantum', TimeQuantum)
    doc.replace_doc('InterestPeriod', InterestPeriod)
    doc.replace_doc('SubscriptionDate', SubscriptionDate)
    doc.replace_doc('RegisterDate', RegisterDate)
    #doc.replace_doc('RegisterDate0', RegisterDate0)
    doc.replace_doc('DueDate', DueDate)
    doc.replace_doc('PaymentDate', PaymentDate)
    doc.replace_doc('redeem', redeem)
    doc.replace_doc('BuyBack', BuyBack)
    doc.replace_doc('transfer', transfer)
    doc.replace_doc('YM', YM)

    # doc.replace_doc('．', '.')  # 替换．为.
    # doc.replace_doc('\n', '')  # 去除空行
    # doc.replace_doc('o', '0')  # 替换o为0
    # # doc.replace_docs('([0-9])@[、,，]([0-9])@', '\1.\2')  使用@不能识别改用{1,}，\需要使用反斜杠转义
    # doc.replace_docs('([0-9]){1,}[、,，．]([0-9]){1,}', '\\1.\\2') # 将数字中间的，,、．替换成.
    # doc.replace_docs('([0-9]){1,}[旧]([0-9]){1,}', '\\101\\2')   # 将数字中间的“旧”替换成“01”
    doc.close()
