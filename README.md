# quizGen
CM&A quiz generator

This generator is creates quiz "packets" from a spreadsheet of questions (i.e. "the question database"), such as those from the Central Quizzing Leadership Team.  The primary interface is a Jupyter notebook front-end to a Python module.



## How to use it

See the Jupyter notebook for a more thorough walk-through, but briefly:

### Annotate a copy of the database

We've added an additional column "GROUP" to the database, which allows us to assign a grouping name for drawing questions.  For example, we can label questions with "150" or "300" if we wish to restrict the generator to certain kinds of questions.  We'll show how to use this below.

<img src="/images/question_grouping.png" alt="question grouping"/>

### Create the quiz generator

First, we create an instance of the quiz generator instance that we can use to generate our quizzes.  QG is the Quiz generator instance, associated with the named quiz question file for Hebrews and 1st and 2nd Peter.
```python
# 
# Instantiate the quiz generator with the database
#
fnxls='HEB1P2P_CMA_marked.xls'
QG=quizgen.QuizGenerator(fndatabase=fnxls)
```

By default, the quiz generator is set for Epistles, and will have the following properties:
```python
Quiz Type: epistle
    Number of quizzes: 1
    Verbose: False
    Scramble period: 1

    Quiz Distribution:
{'cr': {'label': 'Chapter Reference', 'range': (3, 7),  'types': ('cr', 'cvr', 'cvrma', 'crma')},
 'ft': {'label': 'Finish-The-Verse',  'range': (3, 4),  'types': ('ft', 'f2v', 'ftv', 'ftn')},
'int': {'label': 'Interrogative',     'range': (9, 16), 'types': ('int',)},
 'ma': {'label': 'Multiple Answer',   'range': (1, 2),  'types': ('ma',)},
  'q': {'label': 'Quote',             'range': (3, 4),  'types': ('q', 'q2')}}
```
The quiz distribution shows the category label, the min and max of questions that can occur in a normal quiz, and the specific sub-types that fall in this category.

### Configuring the generator
 
First, we want to specify the quiz makeup (i.e. how much of different parts of the material to sample).  This is specified in JSON format like the following:

```python
quizMakeup={'current':{'frac':0.5,'content':[('HEB',[12,13]),('1P',[1])]},
               'past':{'frac':0.5,'content':[('HEB',[1,2,3,4,5,6,7,8,9,10,11])]}
            }
```

The interpretation of this is that there are two different blocks of verses ('current' and 'past'), with each block representing 50% of the questions.  The 'past' block, in this case, covers Hebrews chapters 1-11.  The 'current' block covers chapters 12-13 of Hebrews and chapter 1 of 1st Peter.

Below, we configure the generator using this quiz makeup, as well as the number of quizzes to generate.  We can also optionally configure the generator to only draw quote and finish-this type questions to '150' and '300' verses.  Districts can impose there own limits according to their grouping variables.
```python
# set the quiz makeup and number of quizzes
QG.setQuizMakeup(quizMakeup,nquiz=nquiz)

# add custom limits for certain question types
QG.quizDistribution['q']['limit']=(150,300)
QG.quizDistribution['ft']['limit']=(150,300)
#QG.verbose=True
```

### Generate quiz tables
This reads the quiz database, filters the content, and draws questions for all the requested quizzes.

```python
# pull the questions that will be drawn from for this series 
# of quizzes
QG.getContent()
# generate the tabulated questions, and some extras of each 
# question type
QG.generateQuizTables(xtra=10)
 ```

### Generate the quiz packet

Lastly, the quiz packet is generated, pulling in header information and the quiz tables into a Microsoft Word document (docx).
```python
#QG.genQuizPacket(fn='AAC_Practice_20191202_A.docx',title='AAC Practice - 12/2/2019')
msg={'intro':'This is a CM&A Bibble Quizzing packet.  The quiz master is encourage to review each quiz for accuracy.  '+\
             'The quiz packet should have these characteristics:',
     'list':['Satisfaction of question minimums and maximums for each type.  Stats are shown at the end of each quiz.',
            '"A" division quizzes have 50% current and 50% past periods.  These stats are also shown at the end of each quiz.',
            '"B" division quizzes are only current content, which in some cases may lead to repeats which are flagged.'+\
            '  While we have tried to keep these in the alternative questions 16A, 16B, etc, you may need to replace as necessary.',
            'Finish-This and Quote type questions are limited to the 150 and 300 key verses for A, and 150 key verses for B.']}

QG.genQuizPacket('A_practice_0323.docx',title='A Practice - 3/23/2020',msg=msg)
```
<img src="/images/quiz_packet.png" alt="question grouping"/>

### Extra questions and repeats

Extra questions of each type are necessary, and these are located in the back of the quiz packet.  In additition, there will be some cases, especially for quiz tiers that only quiz on a limited set of questions (e.g beginning of the year or some junior divisions), that many quizzes could eventually result in repeats being necessary.  These repeat questions are marked as such and highlighted in yellow.

This example shows that repeat questions didn't appear in any of the quizzes, and only started showing up in some of the extra questions for a junior division.
<img src="/images/extra_repeats.png" alt="question grouping"/>
