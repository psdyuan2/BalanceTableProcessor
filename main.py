import pandas as pd

from CollectDataFromFile import *
from RenameColumn import *
from openpyxl import *
from shutil import copyfile
from sys import exit
from ErrorFile import *
from tqdm import tqdm
from tqdm._tqdm import trange
# Press the green button in the gutter to run the script.

if __name__ == '__main__':
    FilterList1 = ['韩国办','沙特阿拉伯办','西班牙办','缅甸办','香港办', '越南办', '马来西亚办', '香港办', '泰国办', '印尼办', '巴基斯坦办', '意大利办', '菲律宾办', '菲律宾办', '阿联酋办', '新加坡办','']
    try:
        welcome = input('欢迎使用平衡表制作小工具,请按照如下顺序将文件放入目标文件夹：(完成后请按回车确认)\n(a)订单预测表\n(b)未发订单\n(c)部分下单\n(d)未下单\n(e)发货\n(f)大交付报表\n(g)库存\n>>>')
        #get file path list

        PathList, RootPath = GetFilePath()
        empty_df = pd.DataFrame()
        empty_df.to_excel(RootPath+"\\" + "(h)结果.xlsx")
        res_path = RootPath+"\\" + "(h)结果.xlsx"

        writer_1 = pd.ExcelWriter(res_path, mode='w')
        print('文件读取成功')

    except:
        print('文件地址获取出错，请按照要求添加文件到根目录')
        ErrorOne()

    try:
    # create a copy aafafafafafAF
    #(1)collect data content from sale orders
    #用户选择希望作为预测依据的数据
        SaleOrderPath = PathList[0]
        ReqHeaderList = ['区域', '办事处', '订单号', '审批日期', '订单类型', '最终用户', '模块名', '数量', '标准交付周期', '计划交付日期', '超标天数', '物流']
        SaleOrderDF = FileCut(SaleOrderPath, ReqHeaderList)

        #print(SaleOrderDF)
        #FilterList1 = ['香港办','越南办','马来西亚办','香港办','泰国办','印尼办','巴基斯坦办','意大利办','菲律宾办','菲律宾办', '阿联酋办','新加坡办']
        #F1 = Filter(SaleOrderDF, '办事处', FilterList1)
        print('办事处数据被过滤')
        FilterList2 = ['渠道借测','办事处借测']
        #F2 = Filter(F1, '订单类型', FilterList2)
        F2 = SaleOrderDF[(~SaleOrderDF['办事处'].isin(FilterList1)) & (~SaleOrderDF['订单类型'].isin(FilterList2))]
        #Alter the product mode
        F2_sub = F2
        AlteredDF = Alter(F2_sub, '模块名')
        FinalTable1 = pd.concat([F2,AlteredDF], axis=1)
        FinalTable1.to_excel(excel_writer=writer_1, sheet_name='下单数据', index=False)

    except:
        print('下单数据处理失败，请确认格式是否正确')
        ErrorOne()



    #Process the second table
    WaitShipOderPath = PathList[1]
    ReqHeaderList = ['订单号', '订单设备类型', '处理人', '自动挑库', '交付备注', '重大项目', '状态', '类型', '区域', 'KA类型', '客户类型', '最终用户', '最终用户行业',
                         '产品型号', '数量', '备货数量']
    WaitShipOderDF = FileCut(WaitShipOderPath, ReqHeaderList)
    F3 = Filter(WaitShipOderDF, '区域', FilterList1)
    print(F3['区域'].unique())
    #F3 = WaitShipOderDF.query('WaitShipOrderDF['区域'] in FilterList1')]
    AlteredDF = Alter(F3, '产品型号')
    WaitShipOderDF_sub = F3['数量'] - F3['备货数量']
    WaitShipOderDF_sub.name = '未发订单数'
    AlteredDF.name = '型号处理'
    AlteredDF = pd.concat([AlteredDF, WaitShipOderDF_sub], axis=1)
    FinalTable2 = pd.concat([AlteredDF, F3], axis=1)
    FinalTable2.to_excel(excel_writer=writer_1, sheet_name='未发订单(总)',index=False)

    #print('未发订单处理失败，请确认文件的格式是否正确')
    #ErrorOne()

    try:
        #加一个当月下单当月未发
        FinalTable2_sub = pd.DataFrame(columns=list(FinalTable2.columns))
        NowMonth = GetNowTime()
        if len(str(NowMonth)) == 1:
            NowMonth = "0"+str(NowMonth)
        TemCoIndex = list(FinalTable2.columns).index('订单号')
        bar5 = trange(len(FinalTable2['订单号']))
        bar5.set_description('订单号中提取月份')
        for i in bar5:
            TimeCode = FinalTable2.iloc[i, TemCoIndex]
            if ExtractMonth(str(TimeCode)) == NowMonth:
                FinalTable2_sub = FinalTable2_sub.append(FinalTable2.iloc[i, :])
        FinalTable2_sub.to_excel(excel_writer=writer_1, sheet_name='当月下单当月未发', index=False)
    except:
        print('当月下单当月未发表格筛选失败，请确认文件格式是否正确')
        ErrorOne()



    #Process the third table
    BigProjectForecastPath = PathList[2]
    ReqHeaderList = ['预测发起时间', '大项目预测单号', '办事处', '大交付', '项目名称', '预计下单日期', '产品名称', '需求数量', '已下单数量']
    BigProjectForecastDF = FileCut(BigProjectForecastPath, ReqHeaderList)
    #BigProjectForecastDF.to_excel(r'D:\Outputs\test_1\t8.xls', index=False)
    # 项目状态筛选
    GroupList = ['预测发起时间', '大项目预测单号', '办事处', '大交付', '项目名称', '预计下单日期', '产品名称', '需求数量']
    # BigProjectForecastDF = Filter(BigProjectForecastDF, '项目状态', FilterList3)
    # 去重
    TempBFDF = Group(BigProjectForecastDF, GroupList, RootPath)
    # 求数量

    Quantity = TempBFDF['需求数量'] - TempBFDF['已下单数量']
    Quantity.name = '数量'
    TempBFDF.insert(loc=7, column='数量', value=Quantity)
    ProductMode = Alter(TempBFDF, '产品名称')
    ProductMode.name = '订单系统物料号'
    TempBFDF = pd.concat([ProductMode, TempBFDF], axis=1)

    #process the fourth table
    BigProjectForecastPath = PathList[3]
    ReqHeaderList = ['预测发起时间', '大项目预测单号', '办事处', '大交付', '项目名称', '预计下单日期', '原预计下单日期', '产品名称', '需求数量', '项目状态', '已下单数量']
    BigProjectForecastDF = FileCut(BigProjectForecastPath, ReqHeaderList)
    # 项目状态筛选
    FilterList3 = [' ', '部分下单', '取消备货', '全部下单']
    GroupList2 = ['预测发起时间', '大项目预测单号', '办事处', '大交付', '项目名称', '预计下单日期', '原预计下单日期', '产品名称', '需求数量']
    BigProjectForecastDF = Filter(BigProjectForecastDF, '项目状态', FilterList3)
    # 把项目状态这个列删掉
    BigProjectForecastDF = BigProjectForecastDF.drop('项目状态', axis=1)
    # 去重
    TempBFDF2 = Group(BigProjectForecastDF, GroupList2, RootPath)
    # 求数量
    Quantity = TempBFDF2['需求数量'] - TempBFDF2['已下单数量']
    Quantity.name = '数量'
    TempBFDF2.insert(loc=7, column='数量', value=Quantity)
    ProductMode = Alter(TempBFDF2, '产品名称')
    ProductMode.name = '订单系统物料号'
    TempBFDF2 = pd.concat([ProductMode, TempBFDF2], axis=1)
    FinalTable3 = pd.concat([TempBFDF, TempBFDF2], axis=0)
    FinalTable3.to_excel(excel_writer=writer_1, sheet_name='大项目预测', index=False)

    #process the fifth table
    ShipTableDFPath = PathList[4]
    ShipTableDF = FileCut(ShipTableDFPath)
    ShipTableDF = Filter(ShipTableDF, '设备类别', FilterList=['网安', '服务器', 'ADESK'], Replace=False)
    ShipTableDF['发运确认时间'] = pd.to_datetime(ShipTableDF['发运确认时间'])
    MonthCol = ShipTableDF['发运确认时间'].dt.month
    FamilyStuffIDCol = ExtractFamilyStuffID(ShipTableDF, '物料编码')
    FinalTable4 = pd.concat([MonthCol, FamilyStuffIDCol, ShipTableDF], axis=1)
    ChangeName(FinalTable4,'物料编码','家族物料号')
    FinalTable4.to_excel(excel_writer=writer_1, sheet_name='已出货', index=False)


    #process the sixth table
    SaleOrderPath2 = PathList[5]
    ReqHeaderList = ['区域', '办事处', '订单号', '审批日期', '订单类型', '最终用户', '模块名', '数量', '标准交付周期', '计划交付日期', '超标天数', '物流']
    SaleOrderDF = FileCut(SaleOrderPath2, ReqHeaderList)

    # print(SaleOrderDF)

    FilterList2 = ['渠道借测', '办事处借测','安服订单']
    F2 = SaleOrderDF[(~SaleOrderDF['办事处'].isin(FilterList1)) & (~SaleOrderDF['订单类型'].isin(FilterList2))]
    # Alter the product mode
    F2_sub = F2
    AlteredDF = Alter(F2_sub, '模块名')
    FinalTable5 = pd.concat([F2, AlteredDF], axis=1)
    FinalTable5.dropna(axis=1, thresh=10)
    FinalTable5.to_excel(excel_writer=writer_1, sheet_name='下单(大交付报表)', index=False)

    try:
        #Process the seventh table
        #库存的筛选
        StockTablePath = PathList[6]
        StockTableDF = FileCut(StockTablePath)
        FilterList4 = ['空删占位']
        StockTableDF = Filter(StockTableDF, '类别说明', ['空删占位'])
        FamilyStuffIDCol2 = ExtractFamilyStuffID(StockTableDF,'物料编码')
        FinalTable6 = pd.concat([FamilyStuffIDCol2, StockTableDF], axis=1)
        ChangeName(FinalTable6, '物料编码','家族物料号')
        CoIndex2 = list(FinalTable6.columns).index('物料编码')
        RawIndex = []
        for i in range(len(FinalTable6['物料编码'])):
            if str(FinalTable6.iloc[i, CoIndex2])[0] != "1":
                RawIndex.append(FinalTable6.index[i])
        print(RawIndex)
        FinalTable7 = FinalTable6.drop(index=RawIndex, axis=0)
        FinalTable7.to_excel(excel_writer=writer_1, sheet_name='库存', index=False)
    except:
        print('库存处理失败')
        ErrorOne()
    try:
        writer_1.save()
    except:
        print('文件写入失败')


Finish()

















# See PyCharm help at https://www.jetbrains.com/help/pycharm/
