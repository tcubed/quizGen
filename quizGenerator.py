"""CM&A Quiz Generator

This generator is creates quiz "packets" from a spreadsheet of 
questions (i.e. "the question database"), such as those from 
the Central Quizzing Leadership Team.

Typical usage:
    import numpy as np
    import quizgen
    np.random.seed(202002081)

    #
    # specify the quizMakeup, i.e. how much from from which books, 
    # chapters, and verses
    #
    quizMakeup={'current':{'frac':0.5,
                           'content':[('HEB',[12,13]),('1P',[1])]},
                'past':{'frac':0.5,
                        'content':[('HEB',[1,2,3,4,5,6,7,8,9,10,11])]}
                }
    nquiz=6

    # 
    # Instantiate the quiz generator with the database
    #
    fnxls='HEB1P2P_CMA_marked.xls'
    QG=quizgen.QuizGenerator(fndatabase=fnxls)

    # add custom limits for certain question types
    QG.quizDistribution['q']['limit']=(150,300)
    QG.quizDistribution['ft']['limit']=(150,300)
    #QG.verbose=True

    # set the quiz makeup and number of quizzes
    QG.setQuizMakeup(quizMakeup,nquiz=nquiz)
    # pull the questions that will be drawn from for this series 
    # of quizzes
    QG.getContent()
    # generate the tabulated questions, and some extras of each 
    # question type
    QG.generateQuizTables(xtra=10)

Additional options:
    QG.verbose = True           # more descriptive logging
    QG.scramblePeriod = False   # keep period blocks in order
    QG.quizType='custom'        # disables distributions

Custom Quizzes
    Custom quizzes can be created by modifying the QG.quizDistribution 
    property such as:
        # make the quiz type 'custom'
        QG.quizType='custom'
        # reset distribution
        # -- to specialize, make the minimums add up the number of 
        #    questions
        # -- to make this more even, make the minimums close to 1/5 
        #    of the total number of questions.
        # -- to be more representative of the question distributions 
        #    provided by CQLT, make the mins/maxes much higher
        qdist=QG.quizDistribution
        qdist['int']['range']=(10,50)
        qdist['cr']['range']=(10,50)
        qdist['ft']['range']=(10,50)
        qdist['ma']['range']=(10,50)
        qdist['q']['range']=(10,50)


Ted Tower, 2/2020
"""
import os
import pandas as pd
import numpy as np
from IPython.display import display
import xlrd
import re
import pprint

# imports from python-docx to create the Word document
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.enum.text import WD_LINE_SPACING

import logging
#logging.basicConfig(filename='quiz_generator.log',level=logging.DEBUG)
# create logger
logger = logging.getLogger('quiz_generator')
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s|%(name)s|%(levelname)s|%(funcName)s|%(message)s')
logger.handlers=[]
# create console handler and set level to debug
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
ch.setFormatter(formatter)
logger.addHandler(ch)
# file handler
fh = logging.FileHandler('quiz_generator.log')
fh.setFormatter(formatter)
logger.addHandler(fh)


def countTypes(df,quizDistribution):
    """get counts of different question types
    """
    tcount={}
    for qt,qdata in quizDistribution.items():
        # dataframe of this type of question
        if(df.shape[0]>0):
            if(qt=='sit'):
                dftype=df[df['TYPE'].str.lower().str.startswith(qt)]
            else:
                dftype=df[df['TYPE'].str.lower().isin(qdata['types'])]
            nrow=dftype.shape[0]
        else:
            nrow=0
        tcount[qt]=nrow
    return tcount

# DIST_EPISTLE={'int':{'range':(9,16),'types':('int',),'label':'Interrogative'},
#     'cr':{'range':(3,7),'types':('cr','cvr','cvrma','crma'),'label':'Chapter Reference'},
#     'ft':{'range':(3,4),'types':('ft','f2v','ftv','ftn'),'label':'Finish-The-Verse'},
#     'ma':{'range':(1,2),'types':('ma',),'label':'Multiple Answer'},
#     'q':{'range':(3,4),'types':('q','q2'),'label':'Quote'}}
# 2023+
DIST_EPISTLE={'int':{'range':(7,14),'types':('int',),'label':'Interrogative'},
    'cr':{'range':(3,5),'types':('cr','cvr','cvrma','crma'),'label':'Chapter Reference'},
    'ft':{'range':(3,5),'types':('ft','f2v','ftv','ftn'),'label':'Finish-The-Verse'},
    'ma':{'range':(2,4),'types':('ma',),'label':'Multiple Answer'},
    'q':{'range':(2,3),'types':('q','q2'),'label':'Quote'}}

# DIST_GOSPEL={'int':{'range':(8,14),'types':('int',),'label':'Interrogative'},
#     'cr':{'range':(3,6),'types':('cr','cvr','cvrma','crma'),'label':'Chapter Reference'},
#     'ft':{'range':(3,4),'types':('ft','f2v','ftv','ftn'),'label':'Finish-The-Verse'},
#     'ma':{'range':(1,2),'types':('ma',),'label':'Multiple Answer'},
#     'q':{'range':(2,3),'types':('q','q2'),'label':'Quote'},
#     'sit':{'range':(2,4),'types':('sit',),'label':'Situational'},}
DIST_GOSPEL={'int':{'range':(7,14),'types':('int',),'label':'Interrogative'},
    'cr':{'range':(3,5),'types':('cr','cvr','cvrma','crma'),'label':'Chapter Reference'},
    'ft':{'range':(3,5),'types':('ft','f2v','ftv','ftn'),'label':'Finish-The-Verse'},
    'ma':{'range':(2,4),'types':('ma',),'label':'Multiple Answer'},
    'q':{'range':(2,3),'types':('q','q2'),'label':'Quote'},
    'sit':{'range':(2,4),'types':('sit',),'label':'Situational'},}

class QuizGenerator():
    def __init__(self,
                 #fndatabase,
                 quizType='epistle',
                 quizDistribution=None):
        #print('QuizGenerator initialized.')
        #assert os.path.exists(fndatabase),"can't find database: %s"%fndatabase

        self.quizType=quizType.lower()
        self.nquiz=1
        #self.quizMakeup={'current':{'frac':0.5,'content':[('HEB',[6,7,8,9,10],[150,300])]},
        #                'past':{'frac':0.5,'content':[('HEB',[1,2,3,4,5],[150,300])]}
        #    }
        if(self.quizType=='epistle'):
            # epistle
            qdist=DIST_EPISTLE
        elif(self.quizType=='gospel'):
            # gospel
            qdist=DIST_GOSPEL
        elif (self.quizType=='custom'):
            qdist=None
        else:
            raise Exception('quizType is epistle/gospel/custom, not "%s"'%quizType)

        self.quizType=quizType.lower()
        self.quizDistribution=qdist

        self.quizMakeup=None
        self.verbose=False

        #
        # optional settings
        #
        self.scramblePeriod=1   # set to zero to have questions by period (as a check)
        self.rules2013=False
        self.loose=False
        #self.allowLoose=False

        # inits
        self.quizzes=None
        self.quizStats=None
        self.extraQuestions=None
        self.quizContent=None
        
        # data
        # this is a dict containing the entire quiz packet data
        #self.data={'type':,self.quizType,
        #           'quizzes':[],
        #           'extraQuestions':{}}
        
        # load database
        #self.loadDatabase(fndatabase)
        # df=pd.read_excel(fndatabase);
        # df=df.rename(columns={'BOOK':'BK'})
        # df[np.isnan(df['CLUB'])==False]['CLUB'].astype(int)
        # #df['FLAGS']=''
        # df.fillna('', inplace=True)
        # self.database=df
        
        # default -- all content
        content=[]
        # ubk=df['BK'].unique()
        # for bk in ubk:
        #     uch=df[df['BK']==bk]['CH'].unique()
        #     content.append((bk,list(uch)))
        self.quizMakeup={'current':{'frac':1,'content':content}}
        #print('default quizMakeup')
        #print(self.quizMakeup)

        
    #def setQuizType(self,quizType):
    #    self.quizType=quizType
    #    self.data['type']=quizType
    
    def loadDatabase(self,fndatabase):
        assert os.path.exists(fndatabase),"can't find database: %s"%fndatabase
        # fix dataframe
        def convert_to_str(value):
            if pd.isna(value) or value == '':
                return ''
            else:
                return str(int(value))
            
        df=pd.read_excel(fndatabase);
        if('FLAGS' not in df):
            df['FLAGS']=''
        df['FLAGS'].fillna('',inplace=True)
        df['CLUB'] = df['CLUB'].apply(convert_to_str)
        df['INDEX']=range(1, len(df) + 1)
        
        df=df.rename(columns={'BOOK':'BK',
                              'CHAPTER':'CH',
                              'VERSE':'VS'})
        #df[np.isnan(df['CLUB'])==False]['CLUB'].astype(int)
        #df['FLAGS']=''
        #df.fillna('', inplace=True)
        df['BCV']=df['BK']+'_'+df['CH'].astype(str)+'_'+df['VS'].astype(str)
        self.database=df
        
    def getQuizData(self):
        
        return {'type':self.quizType,
                'distribution':self.quizDistribution,
                'quizzes':self.quizzes,
                'extraQuestions':self.extraQuestions,
                'stats':self.quizStats}
    
    def __repr__(self):
        msg="""QuizGenerator instance

    Quiz Type: {quizType}
    Number of quizzes: {nquiz}
    Verbose: {verbose}
    Scramble period: {scramblePeriod}

    Quiz Distribution:
{qdist}

    Quiz Makeup:
{qmakeup}

    
        """.format(quizType=self.quizType,nquiz=self.nquiz,
                    qdist=pprint.pformat(self.quizDistribution,width=60),
                    qmakeup=pprint.pformat(self.quizMakeup,width=30),
                    verbose=self.verbose,
                    scramblePeriod=self.scramblePeriod)


        return(msg)

    #def setQuizMakeup(self,quizMakeup,nquiz=1):
    #    """Set the quiz generator's quizMakeup and nquiz properties."""
    #    self.quizMakeup=quizMakeup
    #    self.nquiz=nquiz

    

        

    def _getContent(self):
        """Get all the content for the quiz in specified range of 
        books & verses.  If a particular question type (e.g. FT) 
        has a limit assigned (e.g. 150), then the CLUB is used 
        to restrict those questions.
        """
        df=self.database
        logger.info('database: %d questions'%len(df))
        
        dfq=df.copy();
        dfq['bcvf']=df['CH']+df['VS']/1000
        
        logger.info('quizMakeup:'+str(self.quizMakeup))
        
        Q={}
        for period,v in self.quizMakeup.items():
            frames=[]
            #for bk,ch,grp in v['content']:
            
            #print('period:',period,' content:',v['content'])
                
            #for bk,ch in v['content']:
            for bcvint in v['content']:
                bk=bcvint[0][0]
                bcvstart=bcvint[0][1]+bcvint[0][2]/1000
                bcvend=bcvint[1][1]+bcvint[1][2]/1000
                
                #print('%s: book %s, ch %s'%(period,bk,str(ch)))
                #df1=df[(df['BK']==bk) & df['CH'].isin(ch)]
                df1=dfq[(dfq['BK']==bk) & (dfq['bcvf']>=bcvstart)&(dfq['bcvf']<=bcvend)]

                #if(len(grp)):
                #    df1=df1[df1['CLUB'].isin(grp)]
                F=[]
                for k,dv in self.quizDistribution.items():
                    #
                    # get type of question
                    #
                    if(k=='sit'):
                        # situational
                        #print(dir(df1['TYPE'].str.lower().str))
                        f=df1[df1['TYPE'].str.lower().str.startswith(k)]
                    else:
                        # all other kinds of questions
                        f=df1[df1['TYPE'].str.lower().isin(dv['types'])]
                    nrows1=f.shape[0]
                    logger.info('period %s: found %d %s questions'%(period,nrows1,k))
                    
                    # limit or set
                    if('limit' in dv):
                        #f=f[f['GROUP'].isin(dv['limit'])]
                        assert isinstance(dv['limit'][0],str),"limits should be strings (e.g. '150')"
                        f=f[f['CLUB'].isin(dv['limit'])]
                    if('set' in dv):
                        f=f[f['SET'].isin(dv['set'])]
                    nrows2=f.shape[0]
                    logger.info('... %d left after club,set'%nrows2)
                    
                    if(nrows2==0):
                        #raise Exception('Ack!  %d/%d %s questions in %s content.'%(nrows2,nrows1,k,period))
                        msg='Warning: "%s" has %d %s question(s); %d pass limits.  ' \
                            'This may result in having to regenerate quizzes.'%(period,nrows1,k,nrows2)
                        logger.warning(msg)
                        #print(msg)
                        
                    
                    F.append(f)
                df1=pd.concat(F)
                
                frames.append(df1)
            Q[period]=pd.concat(frames)

        self.quizContent=Q
    
    def pickQuestionType(self,dfquiz,dfremaining,qtype,
                         otherQuestionCounts=None,
                         loosenDistribution=False):
        """determine question type (if not specified)
        param:
            dfquiz - current quiz dataframe
            dfremaining - dataframe of remaining questions
            qtype       - quiz type (None for min/max; 'any' for random)
            otherQuestionCounts - counts from previous "blocks" of questions
                                  (e.g. current, past)
            loosenDistribution  - whether we allow distribution to drop
                                  (e.g. practice where we'd rather have 
                                   diversity than repeats)
        returns:
            qtpick  - picked question type
            nq      - questions in quiz
        """
        if(qtype==None):
            nq=dfquiz.shape[0]

            # assume minimum distribution has been met
            minmet=True
            # get type counts
            tcount=countTypes(dfquiz,self.quizDistribution)
            for qt,qdata in self.quizDistribution.items():
                # if questionCounts provided, add these to the current counts
                if(otherQuestionCounts!=None):
                    tcount[qt]+=otherQuestionCounts[qt]
                # if the count is less than the minimum, then the minmet flag is False
                if(tcount[qt]<qdata['range'][0]): minmet=False
            
            logger.debug('minimum met: %r'%minmet)
            
            if(loosenDistribution):
                # ignore minimums.  May need to do this if the content is
                # small enough
                minmet=True
            
            # calc prob of picking each type
            keys=list(self.quizDistribution.keys())
            n2satisfy=[]
            #typeRemaining={}
            for qt in keys:
                qdata=self.quizDistribution[qt]
                if(minmet):
                    # if minimum has been met, then prob proportional to num left before max
                    v=qdata['range'][1]-tcount[qt]
                else:
                    # if minimum is not met, then weights determined from questions remaining
                    # yet to be filled to satisfy the minimum
                    v=qdata['range'][0]-tcount[qt]
                n2satisfy.append(v)
                
            if(minmet):
                logger.debug('maxs: %s'%str([self.quizDistribution[k]['range'][1] for k in keys]))
            else:
                logger.debug('mins: %s'%str([self.quizDistribution[k]['range'][0] for k in keys]))
            typeRequirements=dict(zip(keys,n2satisfy))
            logger.debug('requirements: %s'%str(typeRequirements))
            logger.debug('current count: %s'%str(tcount))
            
            # calc weights
            n0=np.maximum(0,n2satisfy)
            weight=[x/sum(n0) for x in n0]
            msg='weight: %s'%str(dict(zip(keys,[round(x,2) for x in weight])))
            logger.debug(msg)
            
            if(np.any(np.isnan(weight))):
                print('wait!')
            
            # get the question type
            qtpick=np.random.choice(keys,p=weight)
        elif(qtype=='any'):
            # pick any (e.g. 16AB-20AB)
            df=dfremaining[dfremaining['used']==0]
            tcount=countTypes(df,self.quizDistribution)
            n=list(tcount.values())
            weight=[x/sum(n) for x in n]
            keys=list(tcount.keys())
            qtpick=np.random.choice(keys,p=weight)
            
            #a=df['TYPE'].value_counts()
            logger.info('qtype:any, count: %s'%str(tcount))
            #1/0
            nq=0
        else:
            raise Exception('pickQuestionType only valid for qtype=None or "any"')
        return qtpick,nq
    
    def pickQuestion(self,dfquiz,dfremaining,qtype=None,
                     otherBCV=None,
                     otherQuestionCounts=None,
                     loosenDistribution=False):
        """Pick a question given the current distribution and remaining 
        questions.

        Args:
            dfquiz (dataframe): the current quiz's dataframe
            dfremaining (dataframe): the remaining unpicked questions
            qtype (string): a specified question-draw type (e.g. for extra 
                questions); default: None
            otherBCV (list): a list of BCVs to exclude
            otherQuestionCounts (dict): a dict of counts of each question type
            loosenDistribution (bool): flag whether to remove minimums
        Returns:
            dfquiz
            dfremaining
        """
        # initialize question counts
        if(otherQuestionCounts==None):
            #tcount=self._countTypes(dfquiz)
            tcount=countTypes(dfquiz,self.quizDistribution)
        else:
            tcount=otherQuestionCounts

        # 
        # determine question type (if not specified)
        #
        if((qtype is None) or (qtype=='any')):
            qtpick,nq=self.pickQuestionType(dfquiz,dfremaining,qtype,
                                            otherQuestionCounts=otherQuestionCounts,
                                            loosenDistribution=loosenDistribution)
            logger.debug('Picked %s'%qtpick)
        else:
            nq=0
            qtpick=qtype
        
        #
        # get all the questions of this type
        #
        qdata=self.quizDistribution[qtpick]
        dftype=dfremaining[dfremaining['TYPE'].str.lower().isin(qdata['types'])]
        kqt=np.where(dftype.index)[0]
        repeat=False
        # if there is a question in the quiz, exclude book-chapter-verses that are already in the quiz
        if(nq):
            uv=dfquiz['BCV'].unique().tolist()
            if(otherBCV!=None):
                uv.extend(otherBCV)
                
            # find unused questions of this type and NOT same book-chapter verse
            kqt=np.where(~dftype['BCV'].isin(uv) & (dftype['used']==0))[0]
            if(len(kqt)==0):
                logger.warning('No unused %s questions left whose book-chapter-verse not already in quiz.'%qtpick)
                # if(loosenDistribution==False):
                #     logger.debug('%s: loosenCount==False.  Returning.'%qtpick)
                #     return dfquiz,dfremaining
                
                # if none, allow repeats but not same book-chapter-verses
                repeat=True
                kqt=np.where(~dftype['BCV'].isin(uv))[0]
                dft=dftype.iloc[kqt]
                # drop questions that have already been used THIS quiz
                uv=dfquiz.index.unique()
                kqt=np.where(~dft.index.isin(uv))[0]
        
                if(len(kqt)):
                    logger.info('%s questions w/repeats, but not in the same book-chapter-verse as another question: %d'%(qtpick,len(kqt)))
                else:
                    # if STILL no questions, then relax the B-C-V
                    # pick among all remaining questions of this type
                    kqt=np.where(dftype.index)[0]
                    logger.warning('%s questions w/repeats, including existing book-chapter-verse s: %d'%(qtpick,len(kqt)))
            else:
                logger.debug('found %d unused %s question whose book-chapter-verse not already in quiz'%(len(kqt),qtpick))
        
        dtype=dftype.iloc[kqt]
        # grab one question
        if(len(dtype)==0):
            msg='No %s questions survived this pick because of exclusions.'%qtpick
            logger.debug(msg)
            return dfquiz,dfremaining
        q=dtype.sample(n=1)

        
        
        q['used']=1
        if(repeat):
            try:
                q['FLAGS']+='R'
            except:
                print('acj?')
        row=q.iloc[0]
        logger.info('Picked %s from %s'%(qtpick,row['BCV']))

        # add to current quiz
        dfquiz=pd.concat([dfquiz,q])

        # set this question to 'used'
        #dfremaining.drop(q.index,inplace=True)
        dfremaining.at[q.index,'used']=1

        return dfquiz,dfremaining
    
    
    def pickQuestionBlock(self,dfremaining,nq,Q1,
                          usedVerses,
                          otherQuestionCounts=None):
        """pick a block of questions from a period"""
        logger.info('pick %d questions for this block'%nq)
        loosened=False
        initNum=len(Q1)
        iter=0
        while((len(Q1)-initNum)<nq):
            iter+=1
            if(iter>(3*nq)):
                msg='Cannot seem to generate enough 1-20 questions to meet distribution.  ' \
                    'This may be because too few chapters in one of the periods.  Consider ' \
                    'rerunning to get a different set or changing chapter ranges.'
                raise Exception(msg)
            nq_old=len(Q1)
            
            logger.info('current num questions: %d (remaining: %d)'%(nq_old,len(dfremaining)))
            #usedVerses=[]
            Q1,dfremaining=self.pickQuestion(Q1,dfremaining,
                                             otherBCV=usedVerses,
                                             otherQuestionCounts=otherQuestionCounts)
            if(len(Q1)==nq_old):
                if((iter>(2*nq)) and (self.loose==True)):
                    
                    if(loosened==False):
                        loosened=True
                        logger.warning('***: Q%d, ALLOWING LOOSENING OF DISTRIBUTION COUNTS'%nq_old)
                    Q1,dfremaining=self.pickQuestion(Q1,dfremaining,
                                                     otherBCV=usedVerses,
                                                     loosenDistribution=True,
                                                     otherQuestionCounts=otherQuestionCounts)
            
            #usedVerses=Q1['BCV'].unique().tolist()
            logger.info('Questions block: %d questions (iter: %d)'%(len(Q1),iter))
            
            #logger.debug('Questions 16AB-20AB: %d questions (iter: %d)'%(len(Q2),iter))
            if((len(Q1)-initNum)>=nq): break

        if(self.verbose):
            # show the distribution for each question type
            #for qt,cnt in self._countTypes(Q1).items():
            for qt,cnt in countTypes(Q1,self.quizDistribution).items():
                rng=self.quizDistribution[qt]['range']
                #print('%s: %d (min: %d, max: %d)'%(qt,cnt,rng[0],rng[1]))
                print('question %s: %d questions (%d remaining)'%(qt,Q1.shape[0],dfremaining.shape[0]))
        return Q1,loosened

    def genQuiz(self,C,nquestion=30):
        """generate quiz based on content

        Args:
            C (dict): content dict
            nquestion (int): number of questions to be generated
                (default: 30)
        """
        if(self.quizContent is None):
            self._getContent()
        
        Q1=pd.DataFrame()
        Q2=pd.DataFrame()
        Q3=pd.DataFrame()
        periodCounts={}
        loosened=False

        #
        # pick first 20 questions -- these go in the dateframe, Q1
        #
        #    for example
        #        for an "A" quiz, the 50% of the questions are picked from the "past" period,
        #                         then 50% of the questions from the "current" period
        usedVerses=[]
        q1counts=countTypes(pd.DataFrame(),self.quizDistribution)
        for ii,(period,dfremaining) in enumerate(C.items()):
            nq=int(20*self.quizMakeup[period]['frac'])
            logger.info('picking %d questions from "%s"'%(nq,period))
            Q1,tmp_loosened=self.pickQuestionBlock(dfremaining,nq,Q1,usedVerses)
            loosened=loosened or tmp_loosened
            periodCounts[period]=[nq]
        
        # scramble first questions
        if(self.scramblePeriod):
            Q1=Q1.sample(frac=1.0)
        # label first questions
        nq1=Q1.shape[0]
        lbl=[str(x+1) for x in range(nq1)]
        Q1['qn']=lbl


        usedVerses=Q1['BCV'].unique().tolist()
        #q1counts=self._countTypes(Q1)
        q1counts=countTypes(Q1,self.quizDistribution)
        #pprint.pprint(q1counts)
        logger.debug('Question 1-20 counts: %s'%str(q1counts))
        logger.debug('=================================')

        #
        # pick for rest of quiz (16AB, 17AB, 18AB, 19AB, 20AB) -- these go in dataframe, Q2
        #
        if(self.quizType!='custom'):
            nq2=10
        else:
            # a "normal" CMA quiz will have 30 questions (16+AB, through 20+AB)
            nq2=nquestion-20
        
        usedVerses=Q1['BCV'].unique().tolist()
        #q2counts=countTypes(Q1,self.quizDistribution)
        for ii,(period,dfremaining) in enumerate(C.items()):
            nq=int(nq2*self.quizMakeup[period]['frac']+1)
            logger.info('picking %d A,B questions (16AB-20AB) from "%s"'%(nq,period))
            Q2,tmp_loosened=self.pickQuestionBlock(dfremaining,nq,Q2,
                                                   usedVerses,
                                                   otherQuestionCounts=q1counts)
            loosened=loosened or tmp_loosened
            #q2counts=countTypes(pd.concat([Q1,Q2]),self.quizDistribution)
            periodCounts[period].append(nq)
        Q2=Q2.iloc[:nq2]

        if(self.verbose):
            # show the distribution for each question type
            #for qt,cnt in self._countTypes(Q2).items():
            for qt,cnt in countTypes(Q2,self.quizDistribution).items():
                rng=self.quizDistribution[qt]['range']
                q12cnt=q1counts[qt]+cnt
                #print('%s: %d (Q1: %d, Q2: %d, min: %d, max: %d)'%(qt,q12cnt,q1counts[qt],cnt,rng[0],rng[1]))

        # scramble second part
        if(self.scramblePeriod):
            Q2=Q2.sample(frac=1.0)
        # label second questions
        if(self.quizType!='custom'):
            lbl=['16A','16B','17A','17B','18A','18B','19A','19B','20A','20B']
        else:
            nq2=Q2.shape[0]
            lbl=[str(x+1+nq1) for x in range(nq2)]
        Q2['qn']=lbl


        #display(Q2)
        q2counts=countTypes(Q2,self.quizDistribution)
        #pprint.pprint(q1counts)
        logger.debug('Question 16AB-20AB counts: %s'%str(q2counts))
        logger.debug('=================================')
        
        
        #
        # pick a few overtime questions
        #
        logger.debug('pick overtime questions')
        if(self.quizType!='custom'):
            nq3=3
        else:
            nq3=0
        usedVerses=Q1['BCV'].unique().tolist()
        usedVerses2=Q2['BCV'].unique().tolist()
        usedVerses.extend(usedVerses2)
        #print(usedVerses)
        #q12counts=self._countTypes(Q1)
        q12counts=countTypes(Q1,self.quizDistribution)
        #for qt,cnt in self._countTypes(Q2).items():
        for qt,cnt in countTypes(Q2,self.quizDistribution).items():
            q12counts[qt]+=cnt
        
        logger.debug('Question 1-20AB counts: %s'%str(q12counts))
        logger.debug('=================================')
        
        #
        # 
        #
        numPeriods=len(self.quizMakeup)
        p=[v['frac'] for v in self.quizMakeup.values()]
        choices=list(self.quizMakeup.keys())
        choosePeriods=np.random.choice(choices,size=3,p=p)
        # all question types need to be different
        qtypes=list(self.quizDistribution.keys())
        overtimeTypes=np.random.choice(qtypes,3,replace=False)
        
        #for period,dfremaining in C.items():
        for ii,period in enumerate(choosePeriods):
            dfremaining=C[period]
            
            nq=1
            #nq=int(nq3*self.quizMakeup[period]['frac']+1)
            if(self.quizType=='custom'): nq=0

            if(self.verbose):
                msg='picking %d overtime questions from "%s"'%(nq,period)
                logger.debug(msg)
                #print('picking third %d questions from %s'%(nq,period))
            for qi in range(nq):
                #print(list(self.quizDistribution.keys()))
                #
                # for overtime, randomly pick a question type
                #
                #qt=np.random.permutation(list(self.quizDistribution.keys()))[0]
                #print(qt)
                #print(qt[0])
                qt=overtimeTypes[ii]
                Q3,dfremaining=self.pickQuestion(Q3,dfremaining,
                                                 otherBCV=usedVerses,
                                                 qtype=qt)
                logger.debug('Overtime: %d questions'%(len(Q3)))
                if(Q3.shape[0]>=nq3): break

        # scramble third part
        #Q2.sample(frac=1.0)
        # label overtime questions
        if(self.quizType!='custom'):
            #lbl=['21','21A','21B','22','22A','22B','23','23A','23B']
            lbl=['21','22','23']
        else:
            nq3=Q3.shape[0]
            lbl=[str(x+1+nq1+nq2) for x in range(nq3)]
            
        #print(lbl)
        #print(Q3)
        Q3['qn']=lbl

        #display(Q3)
        #
        # reorder questions
        #
        rng=['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16',
                '16A','16B','17','17A','17B','18','18A','18B','19','19A','19B','20','20A','20B',
                '21','22','23']
        frames=[]
        if(self.quizType!='custom'):
            for r in rng:
                df=Q1[Q1['qn']==r]
                frames.append(df)
                df=Q2[Q2['qn']==r]
                frames.append(df)
                df=Q3[Q3['qn']==r]
                frames.append(df)
        else:
            frames.append(Q1)
            frames.append(Q2)
            #frames.append(Q3)
        
        dfq=pd.concat(frames,sort=False)

        #display(dfq)
        #1/0
        #
        # uncomment this to write out sorted by verse
        #
        #dfq=dfq.sort_values(by=['CH','VS'])
        idxRepeat=dfq[dfq['FLAGS'].str.contains('R')].index
        #if(len(idxRepeat)):
        #    print('wait!')
        for ii,(period,dfremaining) in enumerate(C.items()):
            ir=idxRepeat[idxRepeat.isin(C[period].index)]
            C[period].loc[ir,'FLAGS']='R'
        
        stats={#'min':self._countTypes(Q1),
               'min':countTypes(Q1,self.quizDistribution),
               'max':q12counts,
               'period':periodCounts,
               'loose':loosened}

        return dfq,C,stats

    def genExtraQuestions(self,C,qtype,xtra):
        """pick extra questions
        """
        #logger.debug('picking extra questions')
        Q1=pd.DataFrame()
        for period,dfremaining in C.items():
            nq=int(xtra*self.quizMakeup[period]['frac'])+1
            for qi in range(nq):
                Q1,dfremaining=self.pickQuestion(Q1,dfremaining,qtype=qtype)
                if(Q1.shape[0]>=xtra): break
        
        # scramble first questions
        Q1.sample(frac=1.0)
        Q1=Q1.iloc[:xtra]
        
        nq1=Q1.shape[0]
        lbl=[str(x+1) for x in range(nq1)]
        Q1['qn']=lbl
        return (Q1,C)

    def generateQuizTables(self,nquiz=None,xtra=5,nquestion=30):
        """Generate quiz tables

        Args:
            xtra (int): number of extra questions to generated
            nquestion (int): number of questions to generate per quiz
        """
        if(nquiz is not None):
            self.nquiz=nquiz
        if(self.quizContent is None):
            self._getContent()
        
        logger.debug('nquiz: %d'%self.nquiz)
        logger.debug('quizMakeup: %s'%str(self.quizMakeup))

        # get copy of content from each period
        #    typically, C={'past':dataFrame,'current':dataFrame}
        C={}
        for period,df in self.quizContent.items():
            df['used']=0
            C[period]=df.copy()
        
        # loop through requested quizzes
        QQ=[];QQstats=[]
        for qi in range(self.nquiz):
            logger.info('GENERATE QUIZ %d'%(qi+1))
            dfq,C,stats=self.genQuiz(C,nquestion=nquestion)
            QQ.append(dfq)
            QQstats.append(stats)
        
        
        qxtra={}
        for qt,qdata in self.quizDistribution.items():
            logger.debug('Pick extra %s questions'%qt)
            Q1,C=self.genExtraQuestions(C,qt,xtra)
            qxtra[qt]=Q1

        self.quizzes=QQ
        self.quizStats=QQstats
        self.extraQuestions=qxtra
        
        logger.info('done generating quizzes, stats, extra questions!')
        return self.getQuizData()



if(__name__=='__main__'):
    print('quiz gen is not meant to be called directly')