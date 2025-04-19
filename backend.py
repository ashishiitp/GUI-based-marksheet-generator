import os,shutil
from flask import Flask,request, redirect
from flask.templating import render_template
from proj1marksheet import generate_Marksheet
from proj1concise import concise_marksheet
from proj1sendemail import sendemail

app = Flask(__name__)

@app.route("/")
def hello():
    return render_template('index.html')

@app.route("/upload/files", methods=["POST"])
def upload():
    if request.method=="POST":
        for file in request.files:
            if file=='master_roll':
                if os.path.exists('./sample_input/master_roll.csv'):
                    os.remove('./sample_input/master_roll.csv')
            if file=='responses':
                if os.path.exists('./sample_input/responses.csv'):
                    os.remove('./sample_input/responses.csv')
            request.files[file].save('./sample_input/'+file+'.csv')

        return redirect('/')

@app.route("/generateMarksheet", methods=["POST"])
def hello1():
    try:
        os.mkdir(r'./my output')
    except FileExistsError:
        shutil.rmtree('./my output')
        os.mkdir(r'./my output')
        pass
    positive = int(request.form['positive'])
    negative = int(request.form['negative'])
    generate_Marksheet(positive,negative)
    return redirect('/')

@app.route("/generateConciseMarksheet",methods=["POST"])
def hello2():
    positive = int(request.form['positive'])
    negative = int(request.form['negative'])
    concise_marksheet(positive,negative)
    return redirect('/')

@app.route("/sendemail",methods=["POST"])
def hello3():
    print("Hello from sendemail")
    sendemail()
    return redirect('/')

app.run(debug=True)