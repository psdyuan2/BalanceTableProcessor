import pandas as pd
import datetime
def RenameColumn(AimDF, NewnameStr, **kwargs):
    if len(list(AimDF.columns)) ==1:
        OriName = list(AimDF.columns)[0]
        resDF = AimDF.rename({OriName:NewnameStr})
        return resDF
    else:
        "传单列对象啊喂！！！！！！"
def GetNowTime():
    return datetime.datetime.now().month



