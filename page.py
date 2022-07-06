from unicodedata import name
from flask import Flask, render_template, url_for, redirect, request, flash
from multiprocessing import context
import pandas as pd
from docxtpl import DocxTemplate


app = Flask(__name__)
app.config['DEBUG'] = True


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        content = request.form.to_dict(flat=True)
        print(content["file"]+".docx")
        if content == None:
            flash('No info provided')
            return redirect(url_for('index'))
        tpl = DocxTemplate(content["file"]+".docx")
        tpl.render(content)
        tpl.save("done/%s%s.docx" % (content["file"], content["surname_N"]))
        return render_template('final.html', content=content)
    return render_template('index.html')
