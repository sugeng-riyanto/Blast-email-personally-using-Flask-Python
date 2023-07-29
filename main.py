import pandas as pd
import smtplib
from flask import Flask, render_template, request, redirect, url_for, flash,make_response,session
from fileinput import filename
from openpyxl import load_workbook
from IPython.display import HTML  #pip install html
import os, fnmatch
from werkzeug.utils import secure_filename
#UPLOAD_FOLDER = os.path.join('static','uploads')

app = Flask(__name__)


# Define allowed files (for this example I want only csv file)
ALLOWED_EXTENSIONS = {'xlsx'}

# Configure upload file path flask
#app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER





app = Flask('app')
your_name = "Sekolah Harapan Bangsa"
your_email="xxxxxxx@shb.sch.id"
your_password="xxxxxxxxxxxxxxxxx"
# Define secret key to enable session
app.secret_key="knowledgefaithcharacters"
# If you are using something other than gmail
# then change the 'smtp.gmail.com' and 465 in the line below
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

@app.route('/')
def index():
    title="Welcome"
    version="V1.07.2023. "
    return render_template("index.html",title_favicon=title,version=version)

        
@app.route('/')
def clear():
    title="Welcome"
    file = 'Book2.xlsx'
    location = "/"
    path = os.path.join(location, file)  
    os.remove(path)
    print ("The file has been removed")
    return render_template("index.html",title_favicon=title)
    


@app.post('/')
def view():
    title="Upload and Send"
    file = request.files['file']
	# save file in local directory
    file.save(file.filename)

    # Parse the data as a Pandas DataFrame type
    email_list = pd.read_excel(file)
    #email_list = pd.read_excel("Book1.xlsx")

    
    """
    Get all the: 
    Subject
    Grade
    virtual_account
    customer_name
    customer_email
    trx_amount
    expired_date
    expired_time
    description
    link
    """
    cc = ['shsmodernhill@shb.sch.id','sugeng.riyanto@shb.sch.id']
    all_Subject=email_list['Subject']
    all_Grade=email_list['Grade']
    all_virtual_account=email_list['virtual_account']
    all_customer_name=email_list['customer_name']
    all_customer_email=email_list['customer_email']
    all_trx_amount=email_list['trx_amount']
    all_expired_date=email_list['expired_date']
    all_expired_time=email_list['expired_time']
    all_description=email_list['description']
    all_link=email_list['link']
    # Loop through the emails
    for i in range(len(all_customer_email)):

        # Get each records name, email, subject and message
        
        Subject=all_Subject[i]
        grade=str(all_Grade[i])
        VA=str(all_virtual_account[i])
        name=all_customer_name[i]
        email=all_customer_email[i]
        nominal=str(all_trx_amount[i])
        expired_date=str(all_expired_date[i])
        expired_time=str(all_expired_time[i])
        description=str(all_description[i])
        link=(all_link[i])
        #message = "Hello Ayah/Bunda Ananda " + name +",\nSemoga dalam kondisi sehat wal afiat.\nBerikut ini adalah tagihan berupa\n"+ message1 + " sejumlah Rp "+VA+",-.\nBerikut ini adalah No. Virtual Account (VA): "+VA+".\nJika ada pertanyaan atau konfirmasi dapat menghubungi finance kami, Ms Penna via no. WhatsApp: \nhttps://bit.ly/mspennashb\nTerima Kasih untuk kerjasama yang terus terjalin hingga saat ini.\n\nBest Regards\n\n*)Note:\nAbaikan pesan ini jika Ayah/Bunda telah melakukan pembayaran."
        message="Kepada Yth.\nOrang Tua/Wali Murid "+name+" (Kelas "+grade+")\n\nSalam Hormat,\nKami hendak menyampaikan info mengenai:\n"+Subject+".\nBatas Tanggal Pembayaran: "+expired_date+"\nSebesar Rp. "+nominal+" ,-\nPembayaran via nomor virtual account(VA)/Bank yaitu "+VA+"\nJika ada pertanyaan atau hendak konfirmasi dapat menghubungi\nIbu Penna(Kasir) dengan nomor WA https://bit.ly/mspennashb atau Bapak Supatmin(Admin) dengan Nomor WA https://bit.ly/wamrsupatminshb4 \n\nTerima kasih atas kerjasamanya.\n\nAdmin Sekolah\n\nCatatan:\nMohon diabaikan jika sudah melakukan pembayaran.\n\nKeterangan:\n"+description+"\n\nLink:\n"+link+"."

        # Create the email to send
        full_email = ("From: {0} <{1}>\n"
                    "To: {2} <{3}>\n"
                    "Subject: {4}\n\n"
                    "{5}"
                    .format(your_name, your_email, name, email, Subject, message))

        # In the email field, you can add multiple other emails if you want
        # all of them to receive the same text
        try:
            server.sendmail(your_email, [email], full_email)
            print('Email to {} successfully sent!\n\n'.format(email))
            flash('{}. A.n {} dengan Email {} berhasil dikirim'.format(i,name,email))
            #return redirect(url_for('upload'))
        except Exception as e:
            print('Student{}. Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
            flash('{}. A.n {} dengan Email {} belum berhasil dikirim, check the format'.format(i,name, email))
    uploaded_df_html=email_list.to_html(classes='table table-stripped')
    return render_template('index.html',data_var=uploaded_df_html)
@app.route('/portfolio-details')
def portfolio_details():
    title="Detail"
    return render_template("portfolio-details.html",title_favicon=title)

app.run(host='0.0.0.0', port=8080)