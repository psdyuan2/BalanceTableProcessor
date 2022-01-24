from CollectDataFromFile import *

if __name__=="__main__":
    StockTablePath = r"D:\型号产品名称对照表\kc.xlsx"
    StockTableDF = pd.read_excel(StockTablePath)

    FamilyStuffIDCol2 = ExtractFamilyStuffID(StockTableDF,'物料编码')
    ProductCodeDF = Alter(StockTableDF, '物料说明', TargeName='产品型号')
    FinalTable6 = pd.concat([FamilyStuffIDCol2, ProductCodeDF, StockTableDF], axis=1)
    #FinalTable6.to_excel(r'D:\型号产品名称对照表\kc2.xlsx')

    SalesOrderDF = pd.read_excel(r"D:\型号产品名称对照表\djf.xlsx")
    ProductCodeDF = Alter(SalesOrderDF, '模块名', TargeName='产品型号')
    FinalTable7 = pd.concat([ProductCodeDF, SalesOrderDF.loc[:, ['区域','订单类型','模块名']]], axis=1)
    FinalTable7.to_excel('D:\型号产品名称对照表\dd.xlsx')
    print(FinalTable7)