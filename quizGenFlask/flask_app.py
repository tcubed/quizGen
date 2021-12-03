# -*- coding: utf-8 -*-
"""
Created on Sun Oct 18 14:46:19 2020

@author: Ted
"""



from flask import Flask, redirect, render_template, request, url_for, send_file
import os, glob
import traceback
import time
import numpy as np
#from docx2pdf import convert

from quizGen import quizGenerator

print('>>>>>>flask_app<<<<<<<')
# instance generator
#fnxls=r'MatthewDistrict_2020_mod.xls'
#fnxls=r'MatthewDistrict_20201004.xls'
fnxls=r'RomansJames.xls'
fnxls=os.path.join('mysite','quizGen',fnxls)


print('os.getcwd():',os.getcwd())
print('os.listdir():',os.listdir())
print('find fnxls? (%s):%r'%(fnxls,os.path.exists(fnxls)))

print('instance quiz generator and writer...')
QG=quizGenerator.QuizGenerator(fndatabase=fnxls,quizType='epistle')
QW=quizGenerator.QuizWriter()

def configQuiz(content,division,nquiz,nextra):

    # partial content
    #QG.quizMakeup={'past':{'frac':0.5,'content':[('Matthew',[1,2])]},
    #            'current':{'frac':0.5, 'content':[('Matthew',[3,4])]}
    #            }
    #QG.quizMakeup={
    #            'current':{'frac':1., 'content':[('Matthew',[5,6])]}
    #            }
    QG.quizMakeup=content
    QG.quizContent=None
    print('configQuiz, content: %s'%str(content))

    # add custom limits for certain question types
    if(division=='A'):
        QG.quizDistribution['q']['limit']=(150,300,)
        QG.quizDistribution['ft']['limit']=(150,300,)
    elif(division=='B'):
        QG.quizDistribution['q']['limit']=(150,)
        QG.quizDistribution['ft']['limit']=(150,)

    # set to Local set for all question types
    for k in QG.quizDistribution.keys():
        QG.quizDistribution[k]['set']=('Local',)

    # gen quizzes
    qdat=QG.generateQuizTables(nquiz=nquiz,xtra=nextra)
    print(qdat.keys())

    # remove all docx and pdf
    print('remove existing docx and pdf')
    L=glob.glob('temp_*.docx')
    for li in L:
        os.remove(li)
    L=glob.glob('temp_*.pdf')
    for li in L:
        os.remove(li)

    # write quizzes
    tsec=time.time()
    fn='temp_%d.docx'%int(tsec)
    fntgt=os.path.join(os.getcwd(),fn)
    ttl='Practice Quizzes'
    print('create quiz packet %s in %s'%(fn,os.getcwd()));
    chlist=sorted(qdat['quizzes'][0]['CH'].unique())
    print('chlist: %s'%str(chlist))
    
    QW.loose=QG.loose
    QW.save(fn,qdat,title=ttl)

    if(os.path.exists(fntgt)):
        print('configQuiz: found %s'%fntgt)
    else:
        print('configQuiz: no %s?'%fntgt)

    # convert to pdf
    print('convert %s to pdf'%fn)
    os.system('abiword --to=pdf %s 2>/dev/null'%fn)

    fn='temp_%d.pdf'%int(tsec)
    fntgt=os.path.join(os.getcwd(),fn)
    if(os.path.exists(fntgt)):
        print('configQuiz: found %s'%fntgt)
    else:
        print('configQuiz: no %s?'%fntgt)

    return fn

def getRange(s):
    # split by commas
    L=s.split(',')
    print('getRange: ',L)
    Lf=[]
    for li in L:
        if(li==''):
            continue
        elif('-' in li):
            L1=li.split('-')
            Lf.extend(list(range(int(L1[0]),int(L1[1])+1)))
        else:
            Lf.append(int(li))
    return Lf

def sendToUser(fnquiz,outfn,outfmt):
    #https://www.roytuts.com/how-to-download-file-using-python-flask/
    rnm,ext=os.path.splitext(fnquiz)
    fndoc=os.path.join(os.getcwd(),'%s.docx'%rnm)
    fnpdf=os.path.join(os.getcwd(),'%s.pdf'%rnm)
    if(outfmt=='pdf'):
        if(os.path.exists(fnpdf)):
            print('send file %s as %s'%(fnpdf,"%s.pdf"%outfn))
            try:
                return send_file(fnpdf, as_attachment=True, attachment_filename="%s.pdf"%outfn)
            except Exception:
                msg=['problem sending file!']
                print(traceback.format_exc())
                return render_template("quizgen.html", comments=comments,messages=msg)
        else:
            print('cannot find %s'%fnpdf)
            msg=['cannot find %s'%fnpdf]
            return render_template("quizgen.html", comments=comments,messages=msg)
    elif(outfmt=='docx'):
        if(os.path.exists(fndoc)):
            print('send file %s as %s'%(fndoc,"%s.docx"%outfn))
            try:
                return send_file(fndoc, as_attachment=True, attachment_filename="%s.docx"%outfn)
            except Exception:
                msg=['problem sending file!']
                print(traceback.format_exc())
                return render_template("quizgen.html", comments=comments,messages=msg)
        else:
            print('cannot find %s'%fndoc)
            msg=['cannot find %s'%fndoc]
            return render_template("quizgen.html", comments=comments,messages=msg)
    else:
        msg=['not supported format']
        return render_template("quizgen.html", comments=comments,messages=msg)


app = Flask(__name__)
app.config["DEBUG"] = True

comments = []
msg=[]

print('define routes...')
# configure quiz
#fnxls='quizgen/MatthewDistrict_2020_mod.xls'
#QG=quizGenerator.QuizGenerator(fn=fnxls)



#@app.route('/')
#def hello_world():
#    return 'Hello from Flask!  Ka-boom!'

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("main_page.html", comments=comments)

    #comments.append(request.form["contents"])
    return redirect(url_for('index'))

@app.route("/quizgen", methods=["GET", "POST"])
def quizgen():
    print('>>>>>quizgen<<<<')
    if request.method == "GET":
        print('GET')
        msg=['']
        return render_template("quizgen.html", comments=comments,messages=[''])
        print('never get here')

    if request.method == "POST":
        print('POST')
        #
        # pull form content
        #
        # TODO: error-handling
        try:
            # book 1 info
            book1=request.form['book1']
            oldch1=getRange(request.form["pastChapters1"])
            newch1=getRange(request.form["currentChapters1"])
            # book2 info
            book2=request.form['book2']
            oldch2=getRange(request.form["pastChapters2"])
            newch2=getRange(request.form["currentChapters2"])
            
            # extra info
            #oldf=float(request.form["pastFraction"])
            newf=float(request.form["currentFraction"])
            nquiz=int(request.form["nquiz"])
            nextra=int(request.form["nextra"])
            division=request.form['division']
            
            # output info
            #outfn=safe_join('.',request.form['outfile'])
            outfn=request.form['outfile']
            outfmt=request.form['format']
            msg=['']
        except Exception as e:
            print(e)
            msg=['problem processing form!']
            msg.append(e)
            print(traceback.format_exc())
            return render_template("quizgen.html", comments=comments,messages=msg)
            print('never get here')

        # configure quiz
        #content={'past':{'frac':oldf,'content':[('Matthew',oldch)]},
        #         'current':{'frac':newf,'content':[('Matthew',newch)]}}
        content={'past':{'frac':1-newf,'content':[(book1,oldch1)]},
                 'current':{'frac':newf,'content':[(book1,newch1)]}}
        #content={'past':{'frac':oldf,'content':[('Matthew',oldch)]},
        #         'current':{'frac':newf,'content':[('Matthew',newch)]}}
        #if()
        #content['past']['content']
        if(len(book2)):
            content['past']['content'].append((book2,oldch2))
            content['current']['content'].append((book2,newch2))
        
        print('division: %s, nquiz: %d, nextra: %d, content: %s'%(division,nquiz,nextra,str(content)));

        #
        # generate quizzes
        #
        seed=int(time.time())
        np.random.seed(seed)
        #np.random.seed(1)
        rs=np.random.get_state()
        print('randomstate: %d'%rs[1][0])
        try:
            fnquiz=configQuiz(content,division,nquiz,nextra)
            print('done generating quiz!!  yay!')
        except Exception as e:
            print(e)
            msg=['Problem generating quizzes: "%s".  If not covered below in Troubleshooting, copy and email this to Ted Tower.'%e]
            #msg.append(e)
            print(traceback.format_exc())
            return render_template("quizgen.html", comments=comments,messages=msg)
            print('never get here')

        #
        # send to user
        #
        try:
            return sendToUser(fnquiz,outfn,outfmt)
            print('never get here!')
        except Exception:
            print('sending... nuts!')
            msg=['problem sending file!']
            print(traceback.format_exc())
            return render_template("quizgen.html", comments=comments,messages=msg)

    #
    # if not GET or POST
    #
    print('not GET or POST')
    msg=['']
    return render_template("quizgen.html", comments=comments,messages=msg)
    #msg=['']
    #comments.append(cfg)
    #comments.append(request.form)
    #return redirect(url_for('quizgen'))



    # https://stackoverflow.com/questions/27628053/uploading-and-downloading-files-with-flask
    #file_contents = request_file.stream.read().decode("utf-8")

    #result = transform(file_contents)

    #response = make_response(result)
    #response.headers["Content-Disposition"] = "attachment; filename=result.csv"