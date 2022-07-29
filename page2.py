import sys
from flask import Flask, render_template, url_for, redirect, request, flash
from multiprocessing import context
import pandas as pd
from docxtpl import DocxTemplate
from flask import Flask, request
import os
from datetime import date

cli = sys.modules['flask.cli']
cli.show_server_banner = lambda *x: None
app = Flask(__name__)
app.config['DEBUG'] = True

months = ["січня", "лютого", "березня", "квітня", "травня", "червня",
          "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"]
days = ["перше", "друге", "третє", "четверте",
        "п'яте", "шосте", "сьоме", "восьме", "дев'яте", "десяте", "одинадцяте", "дванадцяте", "тринадцяте", "чотирнадцяте", "п'ятнадцяте", "шістнадцяте", "вісімнадцяте", "дев'ятнадцяте", "двадцяте", "двадцять перше", "двадцять друге", "двадцять третє", "двадцять четверте", "двадцять п'яте", "двадцять шосте", "двадцять сьоме", "двадцять восьме", "двадцять дев'яте", "тридцяте", "тридцять перше"]
days2 = ["першого", "другого", "третього", "четвертого",
         "п'ятого", "шостого", "сьомого", "восьмого", "дев'ятого", "десятого", "одинадцятого", "дванадцятого", "тринадцятого", "чотирнадцятого", "п'ятнадцятого", "шістнадцятого", "вісімнадцятого", "дев'ятнадцятого", "двадцятого", "двадцять першого", "двадцять другого", "двадцять третього", "двадцять четвертого", "двадцять п'ятого", "двадцять шостого", "двадцять сьомого", "двадцять восьмого", "двадцять дев'ятого", "тридцятого", "тридцять першого"]
years = ["двадцять другого", "двадцять третього", "двадцять четвертого", "двадцять п'ятого", "двадцять шостого", "двадцять сьомого", "двадцять восьмого", "двадцять дев'ятого", "тридцятого", "тридцять першого", "тридцять другого", "тридцять третього", "тридцять четвертого", "тридцять п'ятого", "тридцять шостого", "тридцять сьомого", "тридцять восьмого", "тридцять дев'ятого", "сорокового", "сорок першого", "сорок другого", "сорок третього", "сорок четвертого", "сорок п'ятого", "сорок шостого", "сорок сьомого", "сорок восьмого", "сорок дев'ятого", "п'ятдесятого", "п'ятдесят першого", "п'ятдесят другого", "п'ятдесят третього", "п'ятдесят четвертого", "п'ятдесят п'ятого", "п'ятдесят шостого", "п'ятдесят сьомого", "п'ятдесят восьмого", "п'ятдесят дев'ятого", "шістдесятого", "шістдесят першого", "шістдесят другого",
         "шістдесят третього", "шістдесят четвертого", "шістдесят п'ятого", "шістдесят шостого", "шістдесят сьомого", "шістдесят восьмого", "шістдесят дев'ятого", "сімдесятого", "сімдесят першого", "сімдесят другого", "сімдесят третього", "сімдесят четвертого", "сімдесят п'ятого", "сімдесят шостого", "сімдесят сьомого", "сімдесят восьмого", "сімдесят дев'ятого", "вісімдесятого", "вісімдесят першого", "вісімдесят другого", "вісімдесят третього", "вісімдесят четвертого", "вісімдесят п'ятого", "вісімдесят шостого", "вісімдесят сьомого", "вісімдесят восьмого", "вісімдесят дев'ятого", "дев'яностого", "дев'ятдесят першого", "дев'ятдесят другого", "дев'ятдесят третього", "дев'ятдесят четвертого", "дев'ятдесят п'ятого", "дев'ятдесят шостого", "дев'ятдесят сьомого", "дев'ятдесят восьмого", "дев'ятдесят дев'ятого"]
smallNumbers = [1, 2, 3, 4, 5, 6, 7, 8, 9]


@app.route('/', methods=['GET', 'POST'])
def start():
    return render_template('index.html')


@app.route('/oldCar', methods=['GET', 'POST'])
def oldCar():
    if request.method == 'POST':
        content = request.form.to_dict(flat=True)
        main(content, "old")
        return render_template('final.html', content=content)
    return render_template('oldCar.html', months=months, days=days, years=years, days2=days2, smallNumbers=smallNumbers)


@app.route('/newCar', methods=['GET', 'POST'])
def newCar():
    if request.method == 'POST':
        content = request.form.to_dict(flat=True)
        main(content, "new")
        return render_template('final.html', content=content)
    return render_template('newCar.html', months=months, days=days, years=years, days2=days2, smallNumbers=smallNumbers)


@app.route('/travel', methods=['GET', 'POST'])
def travel():
    if request.method == 'POST':
        content = request.form.to_dict(flat=True)
        main(content, "travel")
        return render_template('final.html', content=content)
    return render_template('travel.html',  months=months, days=days, years=years, smallNumbers=smallNumbers)


@app.route("/shutdown", methods=['GET'])
def shutdown():
    shutdown_func = request.environ.get('werkzeug.server.shutdown')
    if shutdown_func is None:
        raise RuntimeError('Not running werkzeug')
    shutdown_func()
    return "Shutting down..."


def start():
    app.run(host='0.0.0.0', threaded=True, port=5002)


def stop():
    from flask import requests
    resp = requests.get('http://localhost:5002/shutdown')


def main(content, folder):
    if content == None:
        flash('No info provided')
        return redirect(url_for('index'))
    if folder == "old":
        print("word_templates/Old"+content["file"]+".docx")
        tpl = DocxTemplate("word_templates/Old/"+content["file"]+".docx")
    if folder == "new":
        tpl = DocxTemplate("word_templates/New/"+content["file"]+".docx")
    if folder == "travel":
        tpl = DocxTemplate("word_templates/"+content["file"]+".docx")
    tpl.render(content)
    today = date.today()
    d1 = today.strftime("%d.%m.%Y")
    dir = "C:/Users/nikiw/Desktop/Робота/%s" % d1
    if not os.path.exists(dir):
        os.mkdir(dir)
    if not os.path.exists("%s/%s" % (dir, content["file"])):
        os.mkdir("%s/%s" % (dir, content["file"]))
    if not os.path.exists("%s/%s/%s" % (dir, content["file"], content["surname_N1"])):
        os.mkdir("%s/%s/%s" %
                 (dir, content["file"], content["surname_N1"]))
    dir1 = "%s/%s/%s" % (dir, content["file"], content["surname_N1"])
    tpl.save("%s/%s%s.docx" %
             (dir1, content["file"]+" ", content["surname_N1"]))
