# -*- coding: utf-8 -*-
"""
Created on Wed Sep 20 21:13:41 2023

@author: Ted
"""
import cbqzSupport
import pandas as pd

def loadKeyVerses(fnkeyverses):
    df=pd.read_excel(fnkeyverses,dtype={'Chapter':int,'Verse': int,'Club':str})
    #.rename({'Book':'BOOK','Chapter':'CHAPTER','Verse':'VERSE','Set':'SET'},inplace=True)
    
    df.columns = [col.upper() for col in df.columns]
    return df

# CBQZ
fncbqz_local=r'2023_GEPC\GEPC Local_reg.xls'
#fndistrict=r'2023_GEPC\2023-24 GEPC District Set.xls'

df=cbqzSupport.parseCbqzQuestions(fncbqz_local)
df=df.drop(columns=['XQ','XA'])

# %% Jessica questions
fnjess=r'2023_GEPC\gepc_from_Jessica.xlsx'
dfjess=pd.read_excel(fnjess)
dfjess.rename(columns={'qtype':'TYPE','book':'BOOK','chapter':'CHAPTER',
                   'VS':'VERSE','Q':'QUESTION','A':'ANSWER'},
          inplace=True)

# %% candidate questions
fnqftv=r'2023_GEPC\gepc_Q_FTV_Q2_F2.xlsx'
dfcand=pd.read_excel(fnqftv)
dfcand.rename(columns={'qtype':'TYPE','book':'BOOK','chapter':'CHAPTER',
                   'VS':'VERSE','Q':'QUESTION','A':'ANSWER'},
          inplace=True)

# %% combine
df=df.append(dfjess)
df=df.append(dfcand)

# %% key verses
fnkeyverses=r'2023_GEPC\GEPC2023-Key Verses.xlsx'
dfkv=loadKeyVerses(fnkeyverses)

# %% finalize

df=df.merge(dfkv,on=["BOOK","CHAPTER","VERSE"],how="left")
df.fillna('',inplace=True)

df['INDEX']=range(1,len(df)+1)
df['SET']='Local'

# %% export as excel
df.to_excel(r'2023_GEPC\gepc_2023.xlsx',index=False)

# %% export to JJ
#dfs=pd.read_html(fnlocal,flavor='bs4')
#df=dfs[0]

#print("load key verses")
#
#dfkv=loadKeyVerses(fnkeyverses)

#dfkv=pd.read_excel(fnkeyverses,dtype={'Chapter':int,'Verse': int,'Set':int})

# for jumpjock
dfjj=df.copy()
dfjj.rename(columns={'BOOK':'BK','CHAPTER':'CH','VERSE':'VS'},inplace=True)
#dfjj['INDEX']=range(1,len(dfjj)+1)
#dfjj['SET']='Local'
dfjj.to_json(r'2023_GEPC\gepc_2023_db.json',orient='records')