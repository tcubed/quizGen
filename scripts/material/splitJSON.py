# -*- coding: utf-8 -*-
"""
Split the database into multiple records.

This might be needed on some platforms where filesize is limited.

@author: Ted
"""
import json

fnjson=r'2023_GEPC\gepc_2023_db.json'
fnout=r'2023_GEPC\gepc'

recordsPerJson=800

#
# LOAD THE JSON DATABASE
#
with open(fnjson,'rt') as f:
    J=json.load(f)

#
# WRITE OUT INTO MULTIPLE JSONS
#
nblocks=len(J)//recordsPerJson+1
idx=0
dat=[]
for ii,j in enumerate(J):
    if(ii%recordsPerJson==0):
        if(ii>0):
            idx+=1
            with open('%s-part%d.json'%(fnout,idx),'wt') as f:
                json.dump(dat,f)
        dat=[]
    dat.append(j)
if(len(dat)):
    idx+=1
    with open('%s-part%d.json'%(fnout,idx),'wt') as f:
        json.dump(dat,f)
