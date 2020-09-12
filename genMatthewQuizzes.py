# -*- coding: utf-8 -*-
import quizGenerator

import importlib
import numpy as np
np.random.seed(1)
importlib.reload(quizGenerator)

QDAT={'AAC':{'date':'2020xxxx','datestr':'x/x/2020',
             'prefix':r'quizzes/2020/AAC/AAC',
             'A':{'past':[1,2],'current':[3,4]},
             'B':{'past':[1,2],'current':[3,4]}},
      
      'Marshfield':{'date':'2020xxxx','datestr':'x/x/2020',
             'prefix':r'quizzes/2020/Marshfield/Marshfield',
             'A':{'past':[1,2],'current':[3,4]},
             'B':{'past':[1,2],'current':[3,4]}},
      
      'NCD':{'date':'2020xxxx','datestr':'x/x/2020',
             'prefix':r'quizzes/2020/NCD/NCD',
             'A':{'past':[1,2],'current':[3,4]},
             'B':{'past':[1,2],'current':[3,4]}}}

fnxls=r'2020_Matthew/MatthewDistrict_2020.xls'
#pnaac=r'quizzes/2020/AAC/AAC'
#pnmarsh=r'quizzes/2020/Marshfield/EA'
#pnncd=r'quizzes/2020/NCD/NCD'

msg=[{'type':'p',
      'text':'This is a CM&A Bibble Quizzing packet.  '+\
             'Please review each quiz for accuracy.  '+\
             'The quiz packet should have these characteristics:'},
     {'type':'list',
      'text':['Satisfaction of question minimums and maximums for each type.  Distribution stats are shown at the end of each quiz.',
            '"A" division quizzes have 50% current and 50% past periods.  These stats are also shown at the end of each quiz.',
            '"B" division quizzes are only current content, which in some cases may lead to repeats which are flagged.'+\
            '  While we have tried to keep these in the alternative questions 16A, 16B, etc, you may need to replace as necessary.',
            'A teams are limited to Club 150 & 300 verses for Quote and Finish-This type questions.  B teams are limited to Club 150.']},
    {'type':'p',
      'text':'Please be aware of these watch-outs:'},
    {'type':'list',
      'text':['Singular keywords (e.g. ruler) may be partially bolded in plural contexts (e.g. rulers).']},
    {'type':'p',
      'text':'Please let Ted Tower know of any problems you discover.'},]

# %% AAC A practice quizzes
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')

# partial content
past=QDAT['AAC']['A']['past']
current=QDAT['AAC']['A']['current']
QG.quizMakeup={'past':{'frac':0.5,'content':[('Matthew',past)]},
            'current':{'frac':0.5, 'content':[('Matthew',current)]}
            }

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,300)
QG.quizDistribution['ft']['limit']=(150,300)
for k in ['q','ft','int','cr','ma']:
    QG.quizDistribution[k]['set']=('Local',)
    
qdat=QG.generateQuizTables(nquiz=4,xtra=10)   

# write quizzes
QW=quizGenerator.QuizWriter()
fn='%s_A_practice_%s.docx'%(QDAT['AAC']['prefix'],QDAT['AAC']['date'])
ttl='AAC A Practice Quizzes - %s'%QDAT['AAC']['datestr']
QW.save(fn,qdat,title=ttl,msg=msg)


# %% AAC B practice quizzes
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')

# partial content
past=QDAT['AAC']['B']['past']
current=QDAT['AAC']['B']['current']
QG.quizMakeup={'past':{'frac':0.5,'content':[('Matthew',past)]},
            'current':{'frac':0.5, 'content':[('Matthew',current)]}
            }

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,)
QG.quizDistribution['ft']['limit']=(150,)
for k in ['q','ft','int','cr','ma']:
    QG.quizDistribution[k]['set']=('Local',)
    
qdat=QG.generateQuizTables(nquiz=4,xtra=10)   

# write quizzes
QW=quizGenerator.QuizWriter()
fn='%s_B_practice_%s.docx'%(QDAT['AAC']['prefix'],QDAT['AAC']['date'])
ttl='AAC B Practice Quizzes - %s'%QDAT['AAC']['datestr']
QW.save(fn,qdat,title=ttl,msg=msg)

# %% Marshfield A practice quizzes
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')

# partial content
past=QDAT['Marshfield']['A']['past']
current=QDAT['Marshfield']['A']['current']
QG.quizMakeup={'past':{'frac':0.5,'content':[('Matthew',past)]},
            'current':{'frac':0.5, 'content':[('Matthew',current)]}
            }

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,300)
QG.quizDistribution['ft']['limit']=(150,300)
for k in ['q','ft','int','cr','ma']:
    QG.quizDistribution[k]['set']=('Local',)
    
qdat=QG.generateQuizTables(nquiz=4,xtra=10)   

# write quizzes
QW=quizGenerator.QuizWriter()
fn='%s_A_practice_%s.docx'%(QDAT['Marshfield']['prefix'],QDAT['Marshfield']['date'])
ttl='Marshfield A Practice Quizzes - %s'%QDAT['Marshfield']['datestr']
QW.save(fn,qdat,title=ttl,msg=msg)


# %% Marshfield B practice quizzes
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')

# partial content
past=QDAT['Marshfield']['B']['past']
current=QDAT['Marshfield']['B']['current']
QG.quizMakeup={'past':{'frac':0.5,'content':[('Matthew',past)]},
            'current':{'frac':0.5, 'content':[('Matthew',current)]}
            }

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,)
QG.quizDistribution['ft']['limit']=(150,)
for k in ['q','ft','int','cr','ma']:
    QG.quizDistribution[k]['set']=('Local',)
    
qdat=QG.generateQuizTables(nquiz=4,xtra=10)   

# write quizzes
QW=quizGenerator.QuizWriter()
fn='%s_B_practice_%s.docx'%(QDAT['Marshfield']['prefix'],QDAT['Marshfield']['date'])
ttl='Marshfield B Practice Quizzes - %s'%QDAT['Marshfield']['datestr']
QW.save(fn,qdat,title=ttl,msg=msg)

# %% NCD A practice quizzes
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')

# partial content
past=QDAT['NCD']['A']['past']
current=QDAT['NCD']['A']['current']
QG.quizMakeup={'past':{'frac':0.5,'content':[('Matthew',past)]},
            'current':{'frac':0.5, 'content':[('Matthew',current)]}
            }

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,300)
QG.quizDistribution['ft']['limit']=(150,300)
for k in ['q','ft','int','cr','ma']:
    QG.quizDistribution[k]['set']=('Local',)
    
qdat=QG.generateQuizTables(nquiz=4,xtra=10)   

# write quizzes
QW=quizGenerator.QuizWriter()
fn='%s_A_practice_%s.docx'%(QDAT['NCD']['prefix'],QDAT['NCD']['date'])
ttl='NCD A Practice Quizzes - %s'%QDAT['NCD']['datestr']
QW.save(fn,qdat,title=ttl,msg=msg)


# %% NCD B practice quizzes
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='gospel')

# partial content
past=QDAT['NCD']['B']['past']
current=QDAT['NCD']['B']['current']
QG.quizMakeup={'past':{'frac':0.5,'content':[('Matthew',past)]},
            'current':{'frac':0.5, 'content':[('Matthew',current)]}
            }

QG.quizDistribution={'int':{'range':(8,12),'types':('int',),'label':'Interrogative'},
            'ft':{'range':(2,3),'types':('ft','f2v','ftv','ftn'),'label':'Finish-The-Verse'},
            'q':{'range':(1,2),'types':('q','q2'),'label':'Quote'},
            'sit':{'range':(2,4),'types':('sit',),'label':'Situational'},
            'cr':{'range':(3,5),'types':('cr','cvr','cvrma','crma'),'label':'Chapter Reference'},
            'ma':{'range':(2,4),'types':('ma',),'label':'Multiple Answer'}}

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,)
QG.quizDistribution['ft']['limit']=(150,)
for k in ['q','ft','int','cr','ma']:
    QG.quizDistribution[k]['set']=('Local',)
    
qdat=QG.generateQuizTables(nquiz=4,xtra=10)   

ncdmsg=[{'type':'p',
      'text':'This is a custom NCD CM&A B-divison Bibble Quizzing packet.  '+\
             'Please review each quiz for accuracy.  '+\
             'The quiz packet should have these characteristics:'},
     {'type':'list',
      'text':['Satisfaction of question minimums and maximums for each type.  Distribution stats are shown at the end of each quiz.',
            '"A" division quizzes have 50% current and 50% past periods.  These stats are also shown at the end of each quiz.',
            '"B" division quizzes are only current content, which in some cases may lead to repeats which are flagged.'+\
            '  While we have tried to keep these in the alternative questions 16A, 16B, etc, you may need to replace as necessary.']},
    {'type':'p',
      'text':'Please be aware of these watch-outs:'},
    {'type':'list',
      'text':['Singular keywords (e.g. ruler) may be partially bolded in plural contexts (e.g. rulers).',
              'NCD-B uses a custom distribution (contact Philip Osterlund for details).']},
    {'type':'p',
      'text':'Please let Ted Tower know of any problems you discover.'},]

# write quizzes
QW=quizGenerator.QuizWriter()
fn='%s_B_practice_%s.docx'%(QDAT['NCD']['prefix'],QDAT['NCD']['date'])
ttl='NCD B Practice Quizzes - %s'%QDAT['NCD']['datestr']
QW.save(fn,qdat,title=ttl,msg=ncdmsg)
