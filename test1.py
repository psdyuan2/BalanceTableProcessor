import pandas as pd
import re
from os import listdir
from RenameColumn import *
from CollectDataFromFile import *
if __name__ == '__main__':
    def ChangeName(DFobj, ColNameStr):
        exclude_df = pd.read_csv('exclude.csv', encoding='ANSI')
        print(exclude_df)
        ColIndex = list(DFobj.columns).index(ColNameStr)
        for i in range(len(DFobj)):
            try:
                temp_index = exclude_df.loc[:,'物料编码'].to_list().index(DFobj.iloc[i, ColIndex])
                DFobj.iloc[i, ColIndex] = exclude_df.iloc[temp_index,list(exclude_df.columns).index('产品型号')]
            except:
                print('没有找到需要替换的物料编码')
                continue
    path = r'D:\平衡表\平衡表12_15\(g)库存.xlsx'
    temp_df = pd.read_excel(path)
    ChangeName(temp_df, '物料编码')
    temp_df.to_excel(r'D:\平衡表\平衡表12_15\(g)库存2.xlsx')
