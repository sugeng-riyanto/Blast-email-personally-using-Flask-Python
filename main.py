
import pandas as pd
import smtplib
from flask import Flask, render_template, request, redirect, url_for, flash, make_response, session
from fileinput import filename
from openpyxl import load_workbook
from IPython.display import HTML  #pip install html
import os, fnmatch
from werkzeug.utils import secure_filename
#from simple_colors import *
#UPLOAD_FOLDER = os.path.join('static','uploads')
from prettytable import PrettyTable#pip install prettytable
app = Flask(__name__)
# Create Compress with default params



# Define allowed files (for this example I want only csv file)
ALLOWED_EXTENSIONS = {'xlsx'}

# Configure upload file path flask
#app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

#app = Flask('app')
your_name = "Sekolah Harapan Bangsa"
your_email = "shsmodernhill@shb.sch.id"
your_password = "jvvmdgxgdyqflcrf"
# Define secret key to enable session
app.secret_key = "knowledgefaithcharacters"
# If you are using something other than gmail
# then change the 'smtp.gmail.com' and 465 in the line below
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

#

# app name
@app.errorhandler(404)
def not_found(e):
  return render_template("404.html")

# Custom error handler for 500 errors
@app.errorhandler(500)
def internal_server_error(error):
   app.logger.error(f'Internal server error: {error}')
   return render_template('500.html')
#
@app.route('/')
def index():
  title = "Welcome"
  version = "V1.01.2024."
  return render_template("main.html", title_favicon=title, version=version)

@app.route('/clear')
def clear():
  title = "Welcome"
  
  folder_path="./"
  extension = "xlsx"  # set the file extension to delete
  for filename in os.listdir(folder_path): 
    file_path = os.path.join(folder_path, filename)
    try:
      if os.path.isfile(file_path) and filename.endswith(extension):
        os.remove(file_path)
    except Exception as e:
      print(f"Error deleting {file_path}: {e}")
  print("Deletion done")
  return render_template("main.html", title_favicon=title)

@app.route('/reminder')
def reminder():
  title = "RE 1, 2 dan 3"
  version = "V1.07.2023. "
  return render_template("payment.html", title_favicon=title, version=version)


@app.route('/announcement')
def announcement():
  title = "Announcement"
  version = "V1.07.2023."
  return render_template("announcement.html", title_favicon=title, version=version)
  
@app.route('/invoice')
def invoice():
  title = "Monthly, CCA or Annual Fee"
  version = "V1.07.2023. "
  return render_template("index.html", title_favicon=title, version=version)

@app.route('/buktibayar')
def buktibayar():
  title = "Kirim Bukti Bayar"
  version = "V1.07.2023. "
  return render_template("buktibayar.html", title_favicon=title, version=version)

@app.post('/view')
def view():
  title = "Upload and Send"
  file = request.files['file']
  # save file in local directory
  file.save(file.filename)

  # Parse the data as a Pandas DataFrame type
  email_list = pd.read_excel(file)
  #email_list = pd.read_excel("Book1.xlsx")
  all_Subject = email_list['Subject']
  all_Grade = email_list['Grade']
  all_virtual_account = email_list['virtual_account']
  all_customer_name = email_list['customer_name']
  all_customer_email = email_list['customer_email']
  all_trx_amount = email_list['trx_amount']
  all_expired_date = email_list['expired_date']
  all_expired_time = email_list['expired_time']
  all_description = email_list['description']
  all_link = email_list['link']
  # Loop through the emails
  for i in range(len(all_customer_email)):

    # Get each records name, email, subject and message

    Subject = str(all_Subject[i])
    grade = str(all_Grade[i])
    VA = str(all_virtual_account[i])
    name = str(all_customer_name[i])
    email = all_customer_email[i]
    nominal = "{:,.2f}".format(all_trx_amount[i])
    #nominal="{:,.2f}".format(nominal1)
    expired_date = str(all_expired_date[i])
    expired_time = str(all_expired_time[i])
    description = str(all_description[i])
    link = (all_link[i])
    #currency = "{:,.2f}".format(amount)
    #message = "Hello Ayah/Bunda Ananda " + name +",\nSemoga dalam kondisi sehat wal afiat.\nBerikut ini adalah tagihan berupa\n"+ message1 + " sejumlah Rp "+VA+",-.\nBerikut ini adalah No. Virtual Account (VA): "+VA+".\nJika ada pertanyaan atau konfirmasi dapat menghubungi finance kami, Ms Penna via no. WhatsApp: \nhttps://bit.ly/mspennashb\nTerima Kasih untuk kerjasama yang terus terjalin hingga saat ini.\n\nBest Regards\n\n*)Note:\nAbaikan pesan ini jika Ayah/Bunda telah melakukan pembayaran."
    message = "Kepada Yth.\nOrang Tua/Wali Murid " + name + " (Kelas " + grade + ")\n\nSalam Hormat,\nKami hendak menyampaikan info mengenai:\n" + Subject + ".\nBatas Tanggal Pembayaran: " + expired_date + "\nSebesar Rp. " + nominal + "\nPembayaran via nomor virtual account(VA) BNI/Bank yaitu " + VA + ".\n\nTerima kasih atas kerjasamanya.\n\nAdmin Sekolah\n\nCatatan:\nMohon diabaikan jika sudah melakukan pembayaran.\n\nKeterangan:\n" + description + "\nLink:\n" + link + ".\n\nJika ada pertanyaan atau hendak konfirmasi dapat menghubungi\nIbu Penna(Kasir) No. WA: https://bit.ly/mspennashb\nBapak Supatmin(Admin SMP & SMA) No. WA: https://bit.ly/wamrsupatminshb4\n"

    # Create the email to send
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}".format(your_name, your_email, name, email, Subject,message))
    try:
      server.sendmail(your_email, [email], full_email)
      print('Email to {} successfully sent!\n\n'.format(email))
      flash('{}. A.n {} dengan Email {} berhasil dikirim'.format(i+1,name,email))
    except Exception as e:
      print('Student{}. Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
      flash('{}. A.n {} dengan Email {} belum berhasil dikirim, check the format'.format(i+1,name, email))
    uploaded_df_html=email_list.to_html(classes='table table-stripped')
  return render_template('index.html',data_var=uploaded_df_html)
#,methods = ['POST', 'GET']
@app.post('/kirim')
def kirim():
  title = "Upload and Send"
  file = request.files['file']
  # save file in local directory
  file.save(file.filename)

  # Parse the data as a Pandas DataFrame type
  email_list1 = pd.read_excel(file)
  all_subject1 = email_list1['Subject']
  all_customer_email1 = email_list1['Email']
  all_virtual_account1 = email_list1['virtual_account']
  all_customer_name1 = email_list1['Nama_Siswa']
  all_grade1 = email_list1['Grade']
  all_bulanjalan = email_list1['bulan_berjalan']
  all_ket1 = email_list1['Ket_1']
  all_spplebih = email_list1['SPP_30hari']
  all_ket2 = email_list1['Ket_2']
  all_denda = email_list1['Denda']
  all_ket3 = email_list1['Ket_3']
  all_ket4 = email_list1['Ket_4']
  all_total=email_list1['Total']
  for i in range(len(all_customer_email1)):
    Subject = str(all_subject1[i])
    VA = str(all_virtual_account1[i])
    email = all_customer_email1 [i]
    name = str(all_customer_name1[i])
    grade= str(all_grade1[i])
    sppbuljal="{:,.2f}".format(all_bulanjalan[i])#
    ket1 = str(all_ket1[i])
    spplebih = "{:,.2f}".format(all_spplebih[i])
    ket2 = str(all_ket2[i])
    denda = "{:,.2f}".format(all_denda[i])
    ket3 = str(all_ket3[i])
    ket4 = str(all_ket4[i])
    total=str(all_total[i])
   
    #https://blog.finxter.com/how-to-print-bold-text-in-python/#:~:text=You%20can%20change%20your%20text,packages%20and%20modules%20in%20Python.
   
    message = "Kepada Yth.\nOrang Tua/Wali Murid " + name + " (Kelas " + grade + ")\n\nSalam Hormat,\nMenurut data dari bagian keuangan kami, bahwa Bapak/Ibu masih belum melunasi administrasi keuangan berupa SPP dengan detail info sebagai berikut:\na) SPP yang sedang berjalan("+ket1+"): Rp. "+sppbuljal+"\nb) Denda("+ket3+"): Rp. "+denda+"\nc) SPP bulan-bulan sebelumnya("+ket2+"): Rp. "+spplebih+"\nd) Keterangan: "+ket4+"\nTotal tagihan adalah:\nRp. "+total+"\n\nSelain itu, kami hendak memberikan info terkait aturan pembayaran SPP, sebagai berikut:\n1) Mengingat batas pembayaran SPP jatuh tempo pada tanggal 10(sepuluh) setiap bulannya, maka sesuai dengan peraturan sekolah yaitu terhitung mulai tanggal 11(sebelas) dan seterusnya pembayaran SPP akan dikenakan denda sebesar 5% dari besarnya SPP dan Bapak/Ibu mendapatkan "+Subject+".\n2) Selanjutnya jika sampai dengan akhir bulan belum menyelesaikan administrasi tersebut, maka dengan berat hati, siswa tersebut tidak bisa mengikuti Kegiatan Belajar dan Mengajar (KBM) mulai awal bulan sampai dengan pelunasan pembayaran tersebut di atas.\n3) Pembayaran hanya bisa dilakukan melalui transaksi pembayaran SPP (melalui Bank BNI dengan nomor Virtual Account masing-masing siswa), yaitu (VA): "+VA+" .\n4) Apabila Bapak/Ibu telah melakukan pembayaran, harap segera menghubungi Tata Usaha.\n\nDemikian Surat Pemberitahuan ini, atas perhatian dan kerjasamanya kami ucapkan terima kasih.\n\nHormat Kami\nAdmin Sekolah\n\nJika ada pertanyaan atau hendak konfirmasi dapat menghubungi\n1) Ibu Penna(Kasir) di No. WA: https://bit.ly/mspennashb\n2) Bapak Supatmin(Admin SMP & SMA) di No. WA: https://bit.ly/wamrsupatminshb4\n"

    # Create the email to send
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}".format(your_name, your_email, name, email, Subject,message))
    try:
      server.sendmail(your_email, [email], full_email)
      print('Email to {} successfully sent!\n\n'.format(email))
      flash('{}. A.n {} dengan Email {} berhasil dikirim'.format(i+1,name,email))
    except Exception as e:
      print('Student{}. Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
      flash('{}. A.n {} dengan Email {} belum berhasil dikirim, check the format'.format(i+1,name, email))
    uploaded_df_html1=email_list1.to_html(classes='table table-stripped')
  return render_template('payment.html',data_var1=uploaded_df_html1)


@app.post('/pengumuman')
def pengumuman():
  title = "Upload Pengumuman"
  file = request.files['file']
  # save file in local directory
  file.save(file.filename)

  # Parse the data as a Pandas DataFrame type
  email_list = pd.read_excel(file)
  #email_list = pd.read_excel("Book1.xlsx")
  all_Subject = email_list['Subject']
  all_Grade = email_list['Grade']
  all_customer_name = email_list['Nama_Siswa']
  all_customer_email = email_list['Email']
  all_description = email_list['Description']
  all_link = email_list['Link']
  # Loop through the emails
  for i in range(len(all_customer_email)):
    Subject = str(all_Subject[i])
    grade = str(all_Grade[i])
    name = str(all_customer_name[i])
    email = all_customer_email[i]
    description = str(all_description[i])
    link = (all_link[i])
    message ="Kepada Yth.\nOrang Tua/Wali Murid " +name+ " (Kelas " +grade+")\n\nSalam Hormat,\nKami hendak menyampaikan info mengenai "+Subject+". Detail sebagai berikut:\n"+description+".\n\nTerima kasih atas kerjasamanya.\n\nHormat Kami,\nKepala Sekolah Harapan Bangsa\n(Sisilia Juni Arianti dan Agustinus Joko Purwanto)\nLink:\n" +link+ ".\n\nJika ada pertanyaan perihal finansial atau hendak konfirmasi dapat menghubungi\nIbu Penna(Kasir) No. WA:https://bit.ly/mspennashb\nBapak Supatmin(Admin SMP & SMA) No. WA: https://bit.ly/wamrsupatminshb4\n"
    # Create the email to send
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}".format(your_name, your_email, name, email, Subject,message))
    try:
      server.sendmail(your_email, [email], full_email)
      print('Email to {} successfully sent!\n\n'.format(email))
      flash('{}. A.n {} dengan Email {} berhasil dikirim'.format(i+1,name,email))
    except Exception as e:
      print('Student{}. Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
      flash('{}. A.n {} dengan Email {} belum berhasil dikirim, check the format'.format(i+1,name, email))
    uploaded_df_html=email_list.to_html(classes='table table-stripped')
  return render_template('announcement.html',data_var=uploaded_df_html)

if __name__ == '__main__':
   app.run(host='0.0.0.0', debug=True)