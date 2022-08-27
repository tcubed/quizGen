# -*- coding: utf-8 -*-
import quizGenerator

import logging
import importlib
import numpy as np
np.random.seed(1)
importlib.reload(quizGenerator)

quizGenerator.logger.setLevel(logging.DEBUG)
quizGenerator.logger.handlers[0].setLevel(logging.DEBUG)

# instance generator
#fnxls=r'2020_Matthew/MatthewDistrict_2020.xls'
fnxls=r'2021_RomansJames/RomansJames.xls'
#QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='epistle')
QG.loose=True
QG.verbose=True
#QG.scramblePeriod=False

# partial content
QG.quizMakeup={'past':{'frac':0.5,'content':[('James',[1,2,3,4,5])]},
            'current':{'frac':0.5, 'content':[('Romans',[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15])]}
            }
QG.quizMakeup={'past':{'frac':1.,'content':[('Romans',[2])]},
            #'current':{'frac':0.5, 'content':[('Romans',[2])]}
            }
# QG.quizMakeup={
#             'current':{'frac':1., 'content':[('James',[1]),('Romans',[1,2])]}
#             }

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,)
QG.quizDistribution['ft']['limit']=(150,)
for k in QG.quizDistribution.keys():
    QG.quizDistribution[k]['set']=('Local',)

# gen quizzes
qdat=QG.generateQuizTables(nquiz=1,xtra=20)   

# write quizzes
QW=quizGenerator.QuizWriter()
#QW.loose=QG.loose
fn='test_practice_OLD.docx'
ttl='test Practice Quizzes - 1019'
QW.save(fn,qdat,title=ttl)
    
