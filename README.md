# 读取数据

标签（空格分隔）： 运行result.py


---

在此输入正文


    from openpyxl import Workbook
    from openpyxl import load_workbook
    import re
    

处理excel文件
读取数据



    def load_dataset(filename):
        
        wb=load_workbook(filename)
        data=wb.active
        return data

因为正文主要在['H']列中所以提取其中的数据

    filename='scraped-text (1).xlsx'

    texts=load_dataset(filename)['H']

之后进行查找：

    def searchListWord(string,ls):
        if string==None:
            return False
        string=string.lower()
        for l in ls:
            if string.find(l)!=-1:
                return True
        return Fales
之后将文本信息更新：
用正则表达式：
处理文本，将不必要的信息去除：

        txt=txt.lower()#小写
        txt=re.sub(r'\d'," ",txt)#去除数字
        txt=re.findall(r'\w+(?:[\&]\w+)*',txt)#去除不相关的文本结构

统计目标单词出现的次数：

    def xlsx2txt(texts):
        #res={}
        respre={}
        inform={}
        for text in texts:
                txt=text.value
                if searchListWord(txt,targets):

                        txt=txt.lower()
                        txt=re.sub(r'\d'," ",txt)
                        txt=re.findall(r'\w+(?:[\&]\w+)*',txt)
                        txt=(' '.join(txt))
                        res={}

                        maxtar=0
                        for target in targets:
                                countertarget=txt.count(target)
                                res[target]=countertarget#记录下每个单词出现的次数
                                maxtar+=countertarget#本来想写一个每个单词的权重计算，但是姑且全设为1
                        respre[str(text.row)]=maxtar
                        inform[text]=res#每一行关键词出现的记录，但是感觉有暂时没什么用，相当于详细信息


        a=sorted(dict2list(respre),key=lambda x:x[1],reverse=True)
        #因为是python3.6所以没有item(),所以用了一个逆序排序
        
        nums=[]
        for i in range(len(a)):
                nums.append( a[i][0])
                #记录下行号
        return nums
最后将文件写到新建的excel中去：

    def writexlsx(filename,nums):
        wb=Workbook()
        sheet=wb.active
        sheet.title="result"

        i=1
        for num in nums:
                j=1
                datas=load_dataset(filename)[num]
                for data in datas:
                        #print(data.value)
                        sheet.cell(row=i,column=j,value=str( data.value))
                        j=j+1
                i=i+1
        wb.save('result.xlsx')
        





    
        
    
    
    
    
    
    



