import os

import pandas as pd
from os import listdir
from ErrorFile import *
import re
from RenameColumn import *
from tqdm import tqdm
from tqdm._tqdm import trange

def CleanColumns(DFobj):
    TempList = []
    bar = tqdm(DFobj.columns)
    bar.set_description('数据清洗')
    for col in bar:
        TempList.append(str(col).strip())
    DFobj.columns = TempList

def GetFilePath():
    #get the path list of root dir
    RootDirPath = input('请输入目标数据所在文件夹的地址：')


    PathList = listdir(RootDirPath)
    ResList = []
    for path in PathList:
        temPath = RootDirPath + "\\" + path
        ResList.append(temPath)
    return ResList, RootDirPath

def FileCut(FilePath, ReqHeaderList='all'):

    f1 = pd.ExcelFile(FilePath)
    TemXLS = f1.parse()
    LineNumber = 0
    while ('Unnamed: 1' in TemXLS.columns) & ('Unnamed: 6' in TemXLS):
        #flag = input(f"表格的表头无用数据没有删掉，可能会影响后续处理，请问需要删除吗？\n平衡宝读取的表头数据为：{list(TemXLS.columns)}\n请输入(y/n):  ")
        TemXLS = f1.parse(skiprows=LineNumber)
        LineNumber+=1
    if ReqHeaderList == 'all':
        CleanColumns(TemXLS)
        return TemXLS
    else:
        CleanColumns(TemXLS)
        return TemXLS.loc[:,ReqHeaderList]


def Filter_2(DFObj, ColumnNameStr, FilterList, Replace=True):
    if Replace==True:
        try:
            FiltedDF = DFObj[~DFObj[ColumnNameStr].isin(FilterList)]
            #print(FiltedDF)
            return FiltedDF
        except:
            ErrorOne()
    else:
        FiltedDF = DFObj[DFObj[ColumnNameStr].isin(FilterList)]
        return FiltedDF
def Filter(DFObj, ColumnNameStr, FilterList, Replace=True):
    GoodStockList = pd.read_csv('res\BadStock.csv',encoding='ANSI')
    GoodStockList = GoodStockList.iloc[:,0].to_list()
    if '子库存说明' in DFObj.columns:
        DFObj = DFObj[DFObj['子库存说明'].isin(GoodStockList)]

    if Replace==True:
        try:
            FiltedDF = DFObj[~DFObj[ColumnNameStr].isin(FilterList)]
            FiltedDF.to_excel('temp_2.xlsx', index=False)
            transDF = pd.read_excel('temp_2.xlsx')
            #print(FiltedDF)
            return transDF
        except:
            ErrorOne()
    else:
        FiltedDF = DFObj[DFObj[ColumnNameStr].isin(FilterList)]
        FiltedDF.to_excel('temp_2.xlsx', index=False)
        transDF = pd.read_excel('temp_2.xlsx')

        return transDF


def Alter(DFObj, ColumnNameStr, **kwargs):
    TargetName = kwargs.get('TargeName')
    pattern = re.compile(r'([A-Z0-9a-z]*-)*[A-Z0-9a-z()+:]*(托管云)?(加盟)?(专用\))?(交流\))?(LC\))?')
    pattern3 = re.compile(r'([A-Z0-9a-z]*-)*[A-Z0-9a-z+:]*(LC)?(\((托管云)[\u4e00-\u9fa5]*\))?')
    IndexOfColumn = list(DFObj.columns).index(ColumnNameStr)
    NewStrList = []
    bar2 = trange(len(DFObj.index))
    bar2.set_description('产品型号转换')
    for i in bar2:
        # 空值填充
        AimStr = str(DFObj.iloc[i, IndexOfColumn])
        AimStr = AimStr.replace(" ", "")
        AimStr = AimStr.replace("(AK)","")
        patter2 = r'-[(][0-9A-Z()+-]*[\u4e00-\u9fa5]*[0-9A-Z()+-]*[\u4e00-\u9fa5]*[)]'
        sub = re.findall(patter2, AimStr)

        if sub:
            for s in sub:

                AimStr = AimStr.replace(s, "")
        try:
            if '托管云' in AimStr:
                NewStr = re.match(pattern3, AimStr)
            else:
                NewStr = re.match(pattern, AimStr)
        except:
            # 补全空缺产品名称
            DFObj.iloc[i, IndexOfColumn] = DFObj.iloc[i - 1, IndexOfColumn]
            AimStr = DFObj.iloc[i, IndexOfColumn]
            NewStr = re.match(pattern, AimStr)
        NewStr = str(NewStr.group())

        if NewStr != "":
            if (NewStr[-1] == "-") or (NewStr[-1] == "("):
                NewStr = NewStr[:-1]
            NewStrList.append(NewStr)
            # print(f'第 {i} 产品型号： {NewStr} 修改完成')
        else:
            NewStrList.append(DFObj.iloc[i, IndexOfColumn])
            # print(f'第 {i} 产品型号： {DFObj.iloc[i, IndexOfColumn]} 保持不变')

    ResDF = pd.DataFrame(NewStrList, index=DFObj.index, columns=[TargetName])
    return ResDF


def Group(DFObj, GroupList, RootPath):

    resDF = DFObj.groupby(GroupList).sum()

    TransferPath = RootPath + '\\' + 'Temp1.xls'
    resDF.to_excel(TransferPath)

    resDF = pd.read_excel(TransferPath)
    os.remove(TransferPath)
    return resDF

def ExtractFamilyStuffID(Col, ColumnStr, NewNameStr = '家族物料号'):

    ColIndex = list(Col.columns).index(ColumnStr)
    temp_list = []
    bar3 = trange(len(list(Col[ColumnStr])))
    bar3.set_description('提取家族料号')
    for i in bar3:
        try:
            TempStr = Col.iloc[i, ColIndex]
            FamilyStuffID = '10'+TempStr[1:9]
            temp_list.append(FamilyStuffID)
            #print(f'第 {i} 物料号： {Col.iloc[i, ColIndex]} 转换完成')
        except:
            #print(f'第 {i} 物料号： {Col.iloc[i, ColIndex]}转换失败，暂时保留原有物料号')
            temp_list.append(Col.iloc[i, ColIndex])
            continue
    temp_df = pd.DataFrame(temp_list, index=Col.index, columns=[NewNameStr])
    return temp_df

def ExtractMonth(CodeStr):
    CodeStr = str(CodeStr)
    if len(CodeStr) == 10:
        flag = CodeStr[2:4]
    elif len(CodeStr) == 12:
        flag = CodeStr[4:6]
    elif len(CodeStr) == 14:
        flag = CodeStr[6:8]
    else:
        print('订单号读取错误，请检查订单号的格式是否有变动')
        flag = 0
    return flag
def ChangeName(DFobj, ColNameStr, TargetNameStr):
    try:
        exclude_df = pd.read_csv(r'res\exclude.csv', encoding='ANSI')
    except:
        a = input('源文件读取出错')
    ColIndex = list(DFobj.columns).index(ColNameStr)
    ColIndex2 = list(DFobj.columns).index(TargetNameStr)
    bar4 = trange(len(DFobj))
    bar4.set_description('替换不可用的家族料号')
    for i in bar4:
        try:
            temp_index = exclude_df.loc[:,'物料编码'].to_list().index(DFobj.iloc[i, ColIndex])
            DFobj.iloc[i, ColIndex2] = exclude_df.iloc[temp_index,list(exclude_df.columns).index('产品型号')]
        except:
            #print('没有找到需要替换的物料编码')
            continue

if __name__=="__main__":
    path = r'C:\Users\SXF-Admin\Desktop\不良率分析\返修明细数据2018.xlsx'
    df = pd.read_excel(path)
    new_col = Alter(df, '产品', )
    new_col.name = '产品型号'
    new_df = pd.concat([df,new_col],axis=1)
    print(new_df)
    new_df.to_excel(r'C:\Users\SXF-Admin\Desktop\不良率分析\返修数据改.xlsx')










