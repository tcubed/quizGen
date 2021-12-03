# -*- coding: utf-8 -*-
import quizGenerator

import importlib
import numpy as np
np.random.seed(1)
importlib.reload(quizGenerator)

# instance generator
#fnxls=r'2020_Matthew/MatthewDistrict_2020.xls'
fnxls=r'2021_RomansJames/RomansJames.xls'
#QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='epistle')
QG.verbose=True

# partial content
QG.quizMakeup={'past':{'frac':0.5,'content':[('Romans',[3,])]},
            'current':{'frac':0.5, 'content':[('Romans',[3,])]}
            }

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,)
QG.quizDistribution['ft']['limit']=(150,)
for k in QG.quizDistribution.keys():
    QG.quizDistribution[k]['set']=('Local',)

# gen quizzes
qdat=QG.generateQuizTables(nquiz=1,xtra=1)   

# write quizzes
QW=quizGenerator.QuizWriter()
QW.loose=QG.loose
fn='test_practice_1019.docx'
ttl='test Practice Quizzes - 1019'
QW.save(fn,qdat,title=ttl)
    
