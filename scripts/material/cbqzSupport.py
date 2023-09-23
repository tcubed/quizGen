# -*- coding: utf-8 -*-
"""
Created on Mon Aug 22 20:40:35 2022

@author: Ted
"""
import os
import pandas as pd
from bs4 import BeautifulSoup

def parseCBQZquestions_old(fnxls):
    """Load the question database

    Args:
        fnxls (string): filename of Excel file
            The expected format of the Excel file is:
                Book, Chapter, Verse, Verse2, Question, Answer, Club
                -- Club is added to the database as a flag for 
                whether it is part of the 150 or 300 key verses.
                Other club labels are possible.
                -- This parses the Excel file looking for bolded key 
                words.  These are extracted internally.

    This function creates a Pandas DataFrame with the following 
    column headings:
        ['BK','CH','VS','VE','TYPE','QUESTION','ANSWER','CLUB',
          'QKEYWORDS','AKEYWORDS','FLAGS','BCV']
    Most of these are straightforward from the Excel file except the 
    following:
        QKEYWORDS, AKEYWORDS -- keywords in the question or answer 
                (comma separated)
        FLAGS -- currently, this only supports 'repeat'
        BCV   -- a string like <book>_<chapter>_<verse>
                (e.g. HEB_1_1) to help with not asking another 
                question that uses the same verse.
    """
    #https://stackoverflow.com/questions/12371787/how-do-i-find-the-formatting-for-a-subset-of-text-in-an-excel-document-cell?rq=1
    # accessing Column 'C' in this example
    COL_IDX = 5

    book = xlrd.open_workbook(fnxls, formatting_info=True)
    sht = book.sheet_by_index(0)

    hdr=[]
    for ii in range(sht.ncols):
        hdr.append(sht.cell_value(0,ii))
    #regcol=list(set(range(sht.ncols)).difference((COL_IDX,)))

    L=[]
    for row_idx in range(1,sht.nrows):
        if(row_idx>20): break

        # get non-question fields
        row={}
        for ii in range(sht.ncols):
            txt = sht.cell_value(row_idx, ii)
            if(isinstance(txt,str)):
                # remove non-breaking spaces
                txt=txt.replace(u'\xa0', u' ')
            row[hdr[ii]]=txt

        # read question cell and format list
        for COL_IDX in [5,6]:
            text_cell = sht.cell_value(row_idx, COL_IDX)
            text_cell_xf = book.xf_list[sht.cell_xf_index(row_idx, COL_IDX)]

            # skip rows where cell is empty
            if not text_cell:
                continue
            #print(text_cell)

            text_cell_runlist = sht.rich_text_runlist_map.get((row_idx, COL_IDX))
            if text_cell_runlist:
                #print(text_cell)
                #print('(cell multi style) SEGMENTS:')
                #print(text_cell_runlist)
                segments = []
                for segment_idx in range(len(text_cell_runlist)):
                    start = text_cell_runlist[segment_idx][0]
                    # the last segment starts at given 'start' and ends at the end of the string
                    end = None
                    if segment_idx != len(text_cell_runlist) - 1:
                        end = text_cell_runlist[segment_idx + 1][0]
                    segment_text = text_cell[start:end]
                    segments.append({
                        'text': segment_text,
                        'font': book.font_list[text_cell_runlist[segment_idx][1]]
                    })
                    # segments did not start at beginning, assume cell starts with text styled as the cell
                    if text_cell_runlist[0][0] != 0:
                        segments.insert(0, {
                            'text': text_cell[:text_cell_runlist[0][0]],
                            'font': book.font_list[text_cell_xf.font_index]
                        })

                boldlist=[]
                for segment in segments:
                    #if('path' in segment['text']):
                    #    print('   "%s"'%segment['text'],'italic:',segment['font'].italic,'bold:', segment['font'].bold)
                    if(segment['font'].bold):
                        #boldlist.append(segment['text'])
                        st=segment['text'].replace('.','')
                        boldlist.extend(st.split())
                keywords=','.join(boldlist)
            else:
                #print('(cell single style)',
                keywords=''

            # add question and answer keywords
            if(COL_IDX==5):
                row['QKEYWORDS']=keywords
            else:
                row['AKEYWORDS']=keywords

            # add column for flags
            row['FLAGS']=''

            # column for unique verse identifier
            row['BCV']='%s_%d_%d'%(row['BOOK'],int(row['CH']),int(row['VS']))

        L.append(row)
    
    # make dataframe
    df=pd.DataFrame(L)
    # 2019 HEBREWS,1P,2P
    #df=df[['BK','CH','VS','VE','TYPE','QUESTION','ANSWER','GROUP','QKEYWORDS','AKEYWORDS','FLAGS','BCV']]
    #df = df.astype({'CH': int, 'VS': int})
    # 2020 MATTHEW
    df=df[['BOOK','CH','VS','VE','TYPE','QUESTION','ANSWER','CLUB','SET','QKEYWORDS','AKEYWORDS','FLAGS','BCV','INDEX']]
    df = df.astype({'CH': int, 'VS': int,'INDEX':int})
    
    # 2022/ACTS needs some cleaning
    df['BOOK']=df['BOOK'].str.strip()
    
    #self.database=df
    
    # default -- all content
    content=[]
    ubk=df['BOOK'].unique()
    for bk in ubk:
        uch=df[df['BOOK']==bk]['CH'].unique()
        content.append((bk,list(uch)))
    #self.quizMakeup={'current':{'frac':1,'content':content}}
    #print('default quizMakeup')
    #print(self.quizMakeup)
    return content,df
    

    
# %%


def parseCbqzQuestions(fnlocal,return_span=False):
    # Read the HTML file
    with open(fnlocal, "r", encoding="utf-8") as file:
        html_content = file.read()
    
    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html_content, "html.parser")
    
    # Find the table in the HTML (you may need to adjust this based on your HTML structure)
    table = soup.find("table")
    
    # Extract table data into a list of lists
    data = []
    hdr=[]
    # question and answer metadata
    xq=[]
    xa=[]
    for jj,row in enumerate(table.find_all("tr")):
        row_data = []
        for ii,cell in enumerate(row.find_all(["th", "td"])):
            #print(cell)
            if(jj==0):
                hdr.append(cell.get_text().upper())
                continue
            if(hdr[ii] in ['QUESTION','ANSWER']):
                
                # process the spans
                spans = cell.find_all("span")
                
                sdk={}
                for span in spans:
                    key=span.get_text(strip=True)
                    if(key not in sdk):
                        sdk[key]=[]
                    for c in span.attrs['class']:
                        if(c not in sdk[key]):
                            sdk[key].append(c)
                if(hdr[ii]=='QUESTION'):
                    xq.append(sdk)
                else:
                    xa.append(sdk)
                if(return_span):
                    txt=cell
                else:
                    txt=cell.get_text()
            else:
                txt=cell.get_text()
                
            row_data.append(txt)
        if(jj==0):
            continue
        data.append(row_data)
    
    # Convert the list of lists into a Pandas DataFrame
    df = pd.DataFrame(data, columns=hdr)
    df=df[['BOOK','CHAPTER','VERSE','TYPE','QUESTION','ANSWER']]
    df['CHAPTER']=df['CHAPTER'].astype(int)
    df['VERSE']=df['VERSE'].astype(int)
    
    # append question and answer metadata
    df['XQ']=xq
    df['XA']=xa
    
    # if(fnkeyverses!=''):
    #     dfkv=loadKeyVerses(fnkeyverses)
    #     df=df.merge(dfkv,on=["BOOK","CHAPTER","VERSE"],how="left")
    #     df.fillna('',inplace=True)
    
    return df

def loadKeyVerses(fnkeyverses):
    df=pd.read_excel(fnkeyverses,dtype={'Chapter':int,'Verse': int,'Club':str})
    #.rename({'Book':'BOOK','Chapter':'CHAPTER','Verse':'VERSE','Set':'SET'},inplace=True)
    
    df.columns = [col.upper() for col in df.columns]
    return df
    
# Step 6: Display the DataFrame
#print(df)

# %%
if(__name__=='__main__'):
    #fnxls=r'2022_Acts/Acts_20220822b.xls'
    #content,df=loadDatabase(fnxls)
    #df.to_excel('acts_db.xlsx',index=False)
    #df.to_json('acts_db.json',orient='records')
    
    #df=pd.read_excel(r'2022_Acts/acts_db.xlsx')
    
    #df.to_json('acts_db.json',orient='records')
    
    fnlocal=r'2023_GEPC\GEPC Local_reg.xls'
    fnkeyverses=r'2023_GEPC\GEPC2023-Key Verses.xlsx'
    #fndistrict=r'2023_GEPC\2023-24 GEPC District Set.xls'
    
    
    df=parseCbqzQuestions(fnlocal)
    
    
    dfkv=loadKeyVerses(fnkeyverses)
    df=df.merge(dfkv,on=["BOOK","CHAPTER","VERSE"],how="left")
    df.fillna('',inplace=True)
    
    #dfs=pd.read_html(fnlocal,flavor='bs4')
    #df=dfs[0]

    #print("load key verses")
    #
    #dfkv=loadKeyVerses(fnkeyverses)
    
    #dfkv=pd.read_excel(fnkeyverses,dtype={'Chapter':int,'Verse': int,'Set':int})
    
    # for jumpjock
    dfjj=df.drop(columns=['XQ','XA'])
    dfjj.rename(columns={'BOOK':'BK','CHAPTER':'CH','VERSE':'VS'},inplace=True)
    dfjj['INDEX']=range(1,len(dfjj)+1)
    dfjj['SET']='Local'
    dfjj.to_json('gepc2023_db.json',orient='records')
