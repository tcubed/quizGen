# -*- coding: utf-8 -*-
import quizGenerator

import importlib
import numpy as np
np.random.seed(1)
importlib.reload(quizGenerator)

# instance generator
fnxls=r'2020_Matthew/MatthewDistrict_2020.xls'
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')

# partial content
QG.quizMakeup={'past':{'frac':0.5,'content':[('Matthew',[1,2,3,4])]},
            'current':{'frac':0.5, 'content':[('Matthew',[5,6])]}
            }
#QG.quizMakeup={
#            'current':{'frac':1., 'content':[('Matthew',[1,2,3,4,5,6])]}
#            }

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,300)
QG.quizDistribution['ft']['limit']=(150,300)
for k in ['q','ft','int','cr','ma']:
    QG.quizDistribution[k]['set']=('Local',)

# gen quizzes
qdat=QG.generateQuizTables(nquiz=4,xtra=15)   

# write quizzes
QW=quizGenerator.QuizWriter()
fn='test_practice_0917.docx'
ttl='test Practice Quizzes - 09/17'
QW.save(fn,qdat,title=ttl)
    
