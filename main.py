from flask import Flask,render_template,request,redirect,send_file
import docx
from docx.shared import Cm
from fuction import working
from randomstring import rs_generator
from docx2pdf import convert
import os
import time






app = Flask(__name__)


@app.errorhandler(Exception)
def all_exception_handler(error):
   return redirect('/')

@app.route('/',methods=['GET','POST'])
def index():
 if request.method == 'POST':
  rqst= request.form['teksnya'].splitlines()
  test = docx.Document()
  sections = test.sections
  for section in sections:
      section.top_margin = Cm(int(request.form['matas']))
      section.bottom_margin = Cm(int(request.form['mbawah']))
      section.left_margin = Cm(int(request.form['mkiri']))
      section.right_margin = Cm(int(request.form['mkanan']))
      section.page_height = Cm(29.7)
      section.page_width = Cm(21)
      section.footer_distance = Cm(0)
      section.header_distance = Cm(0)


  list_font = ['Laprakv1','Laprakv2','Laprakv3','Laprakv4','Laprakv5','Laprakv6','Laprakv7','Laprakv8','Laprakv9']

  for baris in rqst:
   splitted=baris.split('**;')
   if 'tengah' in splitted:
    working(main=test,font=list_font,baris=splitted,alenia='tengah')
   elif 'kiri'in splitted:
    working(main=test,font=list_font,baris=splitted,alenia='kiri')
   elif 'kanan' in splitted:
    working(main=test,font=list_font,baris=splitted,alenia='kanan')
   elif 'rata' in splitted:
    working(main=test,font=list_font,baris=splitted,alenia='rata')
  
  name_file = rs_generator()
  test.save(f'{name_file}.docx')
  convert(f"{name_file}.docx")
  os.remove(f"{name_file}.docx")
  return send_file(f"{name_file}.pdf",as_attachment=True)
 return render_template(f'index.html')


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)