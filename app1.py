from flask import Flask,render_template,request,send_file
import os
import pandas as pd
import selenium
from io import BytesIO
from PIL import Image
from io import BytesIO
import os
import base64
import pytesseract
from functools import wraps
import datetime
from selenium import webdriver as wb
from selenium.common.exceptions import NoSuchElementException as nse
import datetime

app=Flask(__name__)

ts = datetime.datetime.now()
import uuid
v=str(uuid.uuid1())


from pymongo import MongoClient
import pymongo
connection_url = "mongodb+srv://Nikhil:mlg@mca.6gfn0.mongodb.net/myFirstDatabase?retryWrites=true&w=majority"
client = pymongo.MongoClient(connection_url)
try:
  connection_url = "mongodb+srv://Nikhil:mlg@mca.6gfn0.mongodb.net/myFirstDatabase?retryWrites=true&w=majority"
  client = pymongo.MongoClient(connection_url)
  print("Connected successfully!!!")
except:
	print("vishnuu")
db = client.get_database('WebScraping')
collection = db.MCA

app.config["UPLOAD_FOLDER1"]="static/excel"

@app.route("/",methods=['GET','POST'])
def upload():
    if request.method == 'POST':
        upload_file = request.files['upload_excel']
        first_name = request.form.get("fname")
        collection = db.company_details
        if upload_file.filename != '':
            file_path = os.path.join(app.config["UPLOAD_FOLDER1"], upload_file.filename)
            upload_file.save(file_path)

            user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"
            options =wb.ChromeOptions()
            options.headless = True
            options.add_argument(f'user-agent={user_agent}')
            options.add_argument("--window-size=1920,1080*")
            options.add_argument('--ignore-certificate-errors')
            options.add_argument('--allow-running-insecure-content')
            options.add_argument("--disable-extensions")
            options.add_argument("--proxy-server='direct://'")
            options.add_argument("--proxy-bypass-list=*")
            options.add_argument("--start-maximized")
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--no-sandbox')
            driver = wb.Chrome(executable_path='chromedriver.exe',options=options)
            driver.get("https://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do")

            driver.maximize_window()
            iii=pd.read_excel(upload_file)
            link=iii['CORPORATE_IDENTIFICATION_NUMBER'].to_list()
            alldetails=[]
            i=0
            for i in range(0,len(link),1):
                cin=link[i]

                from io import BytesIO
                from PIL import Image
                try:
                    element =driver.find_element_by_xpath('//*[@id="MasterDataInputTab1"]/tbody/tr[3]/td[1]')
                except:continue
                loc = element.location
                size = element.size
                left  = 885
                top   = 455
                width = size['width']-80
                height = size['height']-30
                box = (int(left), int(top), int(left+width), int(top+height))
                screenshot = driver.get_screenshot_as_base64()
                img = Image.open(BytesIO(base64.b64decode(screenshot)))
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd =r"Tesseract-OCR\tesseract.exe"
                s= pytesseract.image_to_string(area)
                captcha=s[:s.find("\n")]
                try:
                    id=driver.find_element_by_id('companyID').send_keys(link[i])
                except nse:
                    pass
                try:
                    cap=driver.find_element_by_id('userEnteredCaptcha').send_keys(captcha)
                except nse:
                    pass
                try:
                    submit=driver.find_element_by_xpath('//*[@id="companyLLPMasterData_0"]')
                except nse:
                    pass
                try:
                    driver.execute_script("arguments[0].click();",submit)
                except nse:
                    pass

                try:
                    driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[1]/td[2]').text
                except nse:
                    driver.get('https://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do')
                    element =driver.find_element_by_xpath('//*[@id="MasterDataInputTab1"]/tbody/tr[3]/td[1]')
                    loc = element.location
                    size = element.size
                    left  = 885
                    top   = 455
                    width = size['width']-80
                    height = size['height']-30
                    box = (int(left), int(top), int(left+width), int(top+height))
                    screenshot = driver.get_screenshot_as_base64()
                    img = Image.open(BytesIO(base64.b64decode(screenshot)))
                    area = img.crop(box)
                    pytesseract.pytesseract.tesseract_cmd =r"Tesseract-OCR\tesseract.exe"
                    s= pytesseract.image_to_string(area)
                    captcha=s[:s.find("\n")]
                    try:
                        id=driver.find_element_by_id('companyID').send_keys(link[i])
                    except nse:
                        pass
                    try:
                        cap=driver.find_element_by_id('userEnteredCaptcha').send_keys(captcha)
                    except nse:
                        pass
                    try:
                        submit=driver.find_element_by_xpath('//*[@id="companyLLPMasterData_0"]')
                    except nse:
                        pass
                    try:
                        driver.execute_script("arguments[0].click();",submit)
                    except nse:
                        pass

                for k in range(1,30,1):
                    try:
                        a=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[1]/td[2]').text
                        break
                    except nse:
                        driver.get('https://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do')
                        element =driver.find_element_by_xpath('//*[@id="MasterDataInputTab1"]/tbody/tr[3]/td[1]')
                        loc = element.location
                        size = element.size
                        left  = 885
                        top   = 455
                        width = size['width']-80
                        height = size['height']-30
                        box = (int(left), int(top), int(left+width), int(top+height))
                        screenshot = driver.get_screenshot_as_base64()
                        img = Image.open(BytesIO(base64.b64decode(screenshot)))
                        area = img.crop(box)
                        pytesseract.pytesseract.tesseract_cmd =r"Tesseract-OCR\tesseract.exe"
                        s= pytesseract.image_to_string(area)
                        captcha=s[:s.find("\n")]
                        try:
                            id=driver.find_element_by_id('companyID').send_keys(link[i])
                        except nse:
                            pass
                        try:
                            cap=driver.find_element_by_id('userEnteredCaptcha').send_keys(captcha)
                        except nse:
                            pass
                        try:
                            submit=driver.find_element_by_xpath('//*[@id="companyLLPMasterData_0"]')
                        except nse:
                            pass
                        try:
                            driver.execute_script("arguments[0].click();",submit)
                        except nse:
                            pass
                try:
                    b=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[2]/td[2]').text
                except nse:
                    pass
                try:
                    c=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[3]/td[2]').text
                except nse:
                    pass
                try:
                    d=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[4]/td[2]').text
                except nse:
                    pass
                try:
                    e=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[5]/td[2]').text
                except nse:
                    pass
                try:
                    f=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[6]/td[2]').text
                except nse:
                    pass
                try:
                    g=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[7]/td[2]').text
                except nse:
                    pass
                try:
                    h=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[8]/td[2]').text
                except nse:
                    pass
                try:
                    i=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[9]/td[2]').text
                except nse:
                    pass
                try:
                    j=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[10]/td[2]').text
                except nse:
                    pass
                try:
                    k=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[11]/td[2]').text
                except nse:
                    pass
                try:
                    l=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[12]/td[2]').text
                except nse:
                    pass
                try:
                    m=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[13]/td[2]').text
                except nse:
                    pass
                try:
                    n=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[14]/td[2]').text
                except nse:
                    pass
                try:
                    o=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[15]/td[2]').text
                except nse:
                    pass
                try:
                    p=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[16]/td[2]').text
                except nse:
                    pass
                try:
                    q=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[17]/td[2]').text
                except nse:
                    pass
                try:
                    r=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[18]/td[2]').text
                except nse:
                    pass
                try:
                    s=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[19]/td[2]').text
                except nse:
                    pass
                try:
                    t=driver.find_element_by_xpath('//*[@id="resultTab1"]/tbody/tr[20]/td[2]').text
                except nse:
                    pass
                try:
                    aa=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[2]/td[2]').text
                except nse:
                    aa='-'
                try:
                    bb=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[2]/td[1]/a').text
                except nse:
                    bb='-'
                try:
                    cc=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[2]/td[3]').text
                except nse:
                    cc='-'
                try:
                    dd=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[3]/td[2]').text
                except nse:
                    dd='-'
                try:
                    ee=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[3]/td[1]/a').text
                except nse:
                    ee='-'
                try:
                    ff=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[3]/td[3]').text
                except nse:
                    ff='-'
                try:
                    gg=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[4]/td[2]').text
                except nse:
                    gg='-'
                try:
                    hh=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[4]/td[1]/a').text
                except nse:
                    hh='-'
                try:
                    ii=driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/section/form/div[6]/table/tbody/tr[4]/td[3]').text
                except nse:
                    ii='-'
                temp={'CIN':a,'Company Name':b,'ROC Code':c,'Registration Number':d,'Company Category':e,'Company SubCategory':f,
                  'Class of Company':g,'Authorised Capital(Rs)':h,'Paid up Capital(Rs)':i,
                  'Date of Incorporation':k,'Registered Address':l,'Address other than R/o':m,
                  'Email Id':n,'Whethalldetails Listed or not':o,'ACTIVE compliance':p,'Suspended at stock exchange':q,
                  'Date of last AGM':r,'Date of Balance Sheet':s,'Company Status(for efiling)':t,'Director(1) Name':aa,'DIN/PAN(1)':bb,'Begin date(1)':cc,
                     'Director(2) Name':dd,'DIN/PAN(2)':ee,'Begin date(2)':ff,'Director Name(3)':gg,'DIN/PAN(3)':hh,'Begin date(3)':ii}
                tempp={'database':v,'timestamp':ts,'upload_name':upload_file.filename,'CIN':a,'Company Name':b,'ROC Code':c,'Registration Number':d,'Company Category':e,'Company SubCategory':f,
                  'Class of Company':g,'Authorised Capital(Rs)':h,'Paid up Capital(Rs)':i,
                  'Date of Incorporation':k,'Registered Address':l,'Address other than R/o':m,
                  'Email Id':n,'Whethalldetails Listed or not':o,'ACTIVE compliance':p,'Suspended at stock exchange':q,
                  'Date of last AGM':r,'Date of Balance Sheet':s,'Company Status(for efiling)':t,'Director(1) Name':aa,'DIN/PAN(1)':bb,'Begin date(1)':cc,
                     'Director(2) Name':dd,'DIN/PAN(2)':ee,'Begin date(2)':ff,'Director Name(3)':gg,'DIN/PAN(3)':hh,'Begin date(3)':ii}
                rec_id1 = collection.insert_one(tempp)
                print("Data inserted with record ids",rec_id1)
                alldetails.append(temp)
                x = driver.get('https://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do')
                df=pd.DataFrame(alldetails)
                k=str(upload_file.filename)
                df.to_excel("data.xlsx")
            return render_template("download.html")
    return render_template("UploadExcel.html",variable=v)

@app.route('/download')
def download():
	path = "data.xlsx"
	return send_file(path, as_attachment=True,cache_timeout=0)

if __name__=='__main__':
    app.run(host="localhost", port=50001, debug=True)
