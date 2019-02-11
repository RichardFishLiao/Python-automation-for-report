

```python
import numpy as np
import pandas as pd
```


```python
data = pd.read_excel('Banners Monthly Raw Data 2019_01.xlsx',sheet_name =5)#利用read_excel 获取文档，sheetname是相关的sheet
data["Year/Month"] = data["Year/Month"].astype('str')#强行把时间改成字符串格式
AdList = data[['Advertiser','Order']].drop_duplicates()#目的是获取所有的客户名单，因此不需要重复的客户，使用drop_duplicates()
filter = AdList.values#获取客户的值
```


```python
def cleanData(filter):
    rawData = data.loc[data["Advertiser"]==filter[0]]#将某一个客户名的数据筛选出来
    rawData = data.loc[data["Order"]==filter[1]]#将这个客户的Sales order 筛选出来
    return rawData
```


```python
#indexA是Initiative by month
def indexA():
    indexA=["Drop Initiative","Year/Month"]
    return indexA
#indexB 是Initiative by Month by Creative
def indexB():
    indexB = ["Drop Initiative","Year/Month","Creative"]
    return indexB
#indexC是 Banners by Month
def indexC():
    indexC = ["Drop Initiative","Line Item ID","Position Path","Line Item Duration","Year/Month"]
    return indexC

#indexD是 Banners by Month by Creative
def indexD():
    indexD = ["Drop Initiative","Line Item ID","Position Path","Line Item Duration","Year/Month","Creative"]
    return indexD
#四个index代表不同需求下的不同sheet

def SeriesMaker(rawData,index):
    Grouped = rawData.groupby(index)#根据index的要求使用Groupby函数制作一个数据透视表
    Group = Grouped.sum()#Groupby 过程中相同index的项采用sum方法加在一块显示
    impression = Group["Ad server impressions"]#获取用户浏览量
    clicks = Group["Ad server clicks"]#获取用户点击量
    return impression, clicks



#smallTotalMaker为了整合不同Drop Initiative的不同的total

def smallTotalMaker(impression,clicks):
    impreTotal = impression.groupby("Drop Initiative")#先把不同drop initiative 的total都计算出来
    cliTotal = clicks.groupby("Drop Initiative")#同样的方法把点击量也计算出来
    impreTotal = impreTotal.sum()#采用sum的方法，使用groupby函数之后必须用的
    cliTotal = cliTotal.sum()
    return impreTotal, cliTotal

def smallCount(impression):
    count = np.array(impression.index.labels[0])#利用Series的index 的labels进行分类计算有多少个drop_initiative，然后进行转化为数组类型
    #但是由于此时的count是[0,0,0,1,1,1,2,2,2]的样子，需要去重复，采用set的方法实现一个无序不重复的对象
    count = list(set(count))#然后再list转化为list
    return count


def combine(impression,impreTotal,e):
        Drop_impression = impression.loc[impression.index.labels[0] == e]#将不同dropInitiative的点击量还有浏览量分别取出来，这个方法会用两次
        #一次取出浏览量，一次取出点击量，懒就没有修改变量名字
        smallTotal = pd.Series([impreTotal[e]],index =["%s total"% impression.index.levels[0][e]])#将相应dropInitiative的小total取出来
        #并且把index改成 Drop Initiative total 
        Drop_impression = Drop_impression.append(smallTotal) #整合两边数据
        return Drop_impression

def tableMaker(SeriesArray):
    table = pd.DataFrame(SeriesArray)
    table = table.T#转置表格，因为原table是反的
    
    return table

def finalize(index):
    index = index
    rawTable = cleanData(i)#清洗数据，获取相应客户的信息
    impression,clicks = SeriesMaker(rawTable,index)#获取浏览量还有点击量，这两个变量都是Series 类型
    impreTotal,cliTotal = smallTotalMaker(impression,clicks)#获取小total
    count  = smallCount(impression)#获取Drop Initiative 的数量
    impreResult = pd.Series()#创建Series变量
    cliResult = pd.Series()
    for e in count:
        impreResult = impreResult.append(combine(impression,impreTotal,e))#将整合的数据赋给新的Series
        cliResult = cliResult.append(combine(clicks,cliTotal,e))

    impreRealTotal = impreTotal.sum()#获取最后的大total
    cliRealTotal = cliTotal.sum()
    impreRealTotal = pd.Series([impreRealTotal],index=["Total"])#给大total一个index
    cliRealTotal = pd.Series([cliRealTotal],index=["Total"])
    impreResult=impreResult.append(impreRealTotal)#像上面一下，把total跟Series整合到一块
    cliResult = cliResult.append(cliRealTotal)
    CTR = cliResult/impreResult#计算得到CTR
    CTR = CTR.apply(lambda x: format(x, '.2%'))#修改一下格式，变成小数点后两位的百分数
    table = tableMaker([impreResult,cliResult,CTR])#用函数把表格生成
    table.rename(columns={table.columns[0]:"Impressions Delivered",table.columns[1]:"Clicks Recorded",table.columns[2]:"CTR"},inplace=True)
    return table#输出表格
    
    



    
    

    
    

```


```python
OpenList = []
```


```python
for i in filter:
    try:#由于输出过程中很有可能因为某些文件的格式问题报错，设置异常处理
        temp = i[1].split('_')
        a=temp[1].split(' ')
        str = "_"
        b = str.join(a)
        if (b.find('/')!=-1):
            b = b.replace("/","-")
        if (b.find(':')!= -1):
            b = b.replace(":","_")
        #以上过程把文件名字改成可以输出不报错的格式
        path = "C:\\Users\\lliao\\Desktop\\Project\\19\\"+b+"_01.xlsx"#输出的路径
        writer = pd.ExcelWriter(path)#设定一个Excel的写入变量
        OpenList.append(path)
        tableA = finalize(indexA())#根据不同需求建立表格
        tableB = finalize(indexB())
        tableC = finalize(indexC())
        tableD = finalize(indexD())
        tableA.to_excel(writer,sheet_name='Initiative by Month')#修改sheetname
        tableB.to_excel(writer,sheet_name='Initiative by Month by Creative')
        tableC.to_excel(writer,sheet_name='Banners by Month')
        tableD.to_excel(writer,sheet_name='Banners by Month by Creative')
        writer.save()
    except IOError as error:
        print("这个客户错误："+ i)#有错误就报错，输出，但是不影响程序运行
        print(error)
       
```

    ['这个客户错误：UPMC-University of Pittsburgh Medical Center'
     '这个客户错误：1006695_UPMC Magee Women?€?s Hospital Q4 2018 Digital Campaign']
    [Errno 22] Invalid argument: 'C:\\Users\\lliao\\Desktop\\Project\\19\\UPMC_Magee_Women?€?s_Hospital_Q4_2018_Digital_Campaign_01.xlsx'
    


```python

```


    ---------------------------------------------------------------------------

    NameError                                 Traceback (most recent call last)

    <ipython-input-7-de20b750ea89> in <module>
    ----> 1 impression
    

    NameError: name 'impression' is not defined



```python

```


```python

```


```python

```


```python

```
