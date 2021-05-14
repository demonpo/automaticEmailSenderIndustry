from sys import platform
import shutil

from pyvirtualdisplay import Display
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from PIL import Image
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import time
import smtplib
from pathlib import Path
import os

import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore
import datetime


# Configuracion de firebase
#---------------------------------------
cred = credentials.Certificate('industryminimal-firebase-adminsdk-jbo7x-0eeb4b6e37.json')
firebase_admin.initialize_app(cred)
db = firestore.client()
#---------------------------------------

# Configuracion de correo y urls
#---------------------------------------
today = date.today()
diaActual = ' '+ today.strftime("%d/%m/%Y")
urlLogin = "https://plantabal.dashbox.app/authentication/login"
urlGeneralTableroCortadora = "https://plantabal.dashbox.app/dashboard/tableroCortadora5"
urlGeneralCronologicoCortadora = "https://plantabal.dashbox.app/dashboard/cronologicoCortadora5"
#urlGeneralHistoricoCortadora = "https://industry-minimal.herokuapp.com/dashboard/dashboard6"
urlGeneralTableroAbastecedor = "https://plantabal.dashbox.app/dashboard/tableroAbastecedor"
urlGeneralCronologicoAbastecedor = "https://plantabal.dashbox.app/dashboard/cronologicoAbastecedor"
#urlGeneralHistoricoAbastecedor = "https://industry-minimal.herokuapp.com/dashboard/dashboard6"


screenshotsPath = Path("./screenshots")
imgOutputsPath = Path("./imgOutputs")

directionProfile = "user-data-dir=C:\\Users\\PaulRCam\\AppData\\Local\\Google\\Chrome\\User Data\\Default"
windowsWebDriverPath = Path("./webDrivers/chromedriver.exe")
linuxWebDriverPath = Path("./webDrivers/chromedriver")
fechaDefaultFormato = "2021-04-21 00:00:00"
# https://industry-minimal.herokuapp.com/
# https://plantabal.dashbox.app/
#---------------------------------------

def pathToString(path):
    return str(path.absolute())

def getWebDriverPathByPlatform():
    if platform == "linux" or platform == "linux2":
        return linuxWebDriverPath
    elif platform == "win32":
        return windowsWebDriverPath

def deleteAllFilesInPath(path):
    folder = pathToString(path)
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))

def resizeImg():
    time.sleep(2)
    img = Image.open(screenshotsPath / "tableroCortadora.png")
    width, height = img.size
    left = 75
    top = height / 4.8
    right = width
    bottom = 3 * height / 4
    iml = img.crop((left, top, right, bottom))
    rgb_im = iml.convert('RGB')
    rgb_im.save(imgOutputsPath / "tablero2Cortadora.jpg")
    #iml.show()
    #-------------------------------------------------------------------
    img2 = Image.open(screenshotsPath / "cronologicoCortadora.png")
    width, height = img2.size
    left = 75
    top = height / 5.85
    right = width
    bottom = 3.5 * height / 4
    iml2 = img2.crop((left, top, right, bottom))
    rgb_im2 = iml2.convert('RGB')
    rgb_im2.save(imgOutputsPath / "cronologico2Cortadora.jpg")
    #iml2.show()
    #-------------------------------------------------------------------
    img3 = Image.open(screenshotsPath / "tableroAbastecedor.png")
    width, height = img3.size
    left = 75
    top = height / 4.8
    right = width
    bottom = 3 * height / 4
    iml3 = img3.crop((left, top, right, bottom))
    rgb_im3 = iml3.convert('RGB')
    rgb_im3.save(imgOutputsPath / "tablero2Abastecedor.jpg")
    #iml3.show()
    #-------------------------------------------------------------------
    img4 = Image.open(screenshotsPath / "cronologicoAbastecedor.png")
    width, height = img4.size
    left = 75
    top = height / 5.85
    right = width
    bottom = 3.5 * height / 4
    iml4 = img4.crop((left, top, right, bottom))
    rgb_im4 = iml4.convert('RGB')
    rgb_im4.save(imgOutputsPath / "cronologico2Abastecedor.jpg")
    #iml3.show()
#resizeImg()



def entryWeb(urlLogin, url1c, url2c, url3a, url4a):
    deleteAllFilesInPath(screenshotsPath)
    dateTimeObj = datetime.datetime.now() - timedelta(days=1)
    horas = '00:00:00'
    ano = dateTimeObj.year
    mes = dateTimeObj.month
    dia = dateTimeObj.day
    if mes < 10:
        mes = '0' + str(mes)
    else:
        mes = str(mes)
    if dia < 10:
        dia = '0' + str(dia)
    else:
        dia = str(dia)
    fechaFiltrar = str(ano) + '-' + mes + '-' + dia + ' 00:00:00'
    print(fechaFiltrar)
    options = webdriver.ChromeOptions()
    options.add_argument("--remote-debugging-port=9222")
    # options.add_argument(direction)
    display = Display(visible=0, size=(1920, 1080)).start()

    driver = webdriver.Chrome(executable_path= getWebDriverPathByPlatform(), options=options)
    #driver.execute_script("window.localStorage.setItem('Usuario','dprb96@gmail.com');")
    #driver.execute_script("window.localStorage.clear();")
    #INGRESO DE USUARIO A APLICATIVO
    #------------------------------------
    driver.get(urlLogin)
    driver.execute_script("window.localStorage.clear();")
    time.sleep(10)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#email')))\
                          .send_keys('dprb96@gmail.com')
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#contrasenia')))\
                          .send_keys('123456')
    time.sleep(2) # se espera a poner usuario y contrasena para dar click en login
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button#enviar')))\
                          .send_keys(Keys.ENTER)
    #------------------------------------
    time.sleep(5)
    driver.get(url1c)
    driver.execute_script("document.body.style_zoom='90%'")
    time.sleep(3)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button.btn btn-outline-dark'.replace(' ','.'))))\
                          .send_keys(Keys.ENTER)
    time.sleep(2)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#fecha')))\
                          .send_keys(fechaFiltrar)

    time.sleep(2)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button#aceptarFecha')))\
                          .send_keys(Keys.ENTER)                                                          
    time.sleep(2)
    driver.save_screenshot(pathToString(screenshotsPath / "tableroCortadora.png"))
    
    time.sleep(2)
    driver.get(url2c)
    time.sleep(3)  
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button.btn btn-outline-dark'.replace(' ','.'))))\
                          .send_keys(Keys.ENTER)
    time.sleep(2)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#fecha')))\
                          .send_keys(fechaFiltrar)

    time.sleep(2)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button#aceptarFecha')))\
                          .send_keys(Keys.ENTER)                                                          
    time.sleep(3)
    driver.save_screenshot(pathToString(screenshotsPath / "cronologicoCortadora.png"))
    time.sleep(2)
    

    #---------------------------------------
    #ABASTECEDOR
    #---------------------------------------
    
    driver.get(url3a)
    time.sleep(3)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button.btn btn-outline-dark'.replace(' ','.'))))\
                          .send_keys(Keys.ENTER)
    time.sleep(2)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#fecha')))\
                          .send_keys(fechaFiltrar)

    time.sleep(2)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button#aceptarFecha')))\
                          .send_keys(Keys.ENTER)                                                          
    time.sleep(2)
    driver.save_screenshot(pathToString(screenshotsPath / "tableroAbastecedor.png"))
    time.sleep(2)
    driver.get(url4a)
    time.sleep(3)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button.btn btn-outline-dark'.replace(' ','.'))))\
                          .send_keys(Keys.ENTER)
    time.sleep(2)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#fecha')))\
                          .send_keys(fechaFiltrar)

    time.sleep(2)
    WebDriverWait(driver, 5)\
                          .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button#aceptarFecha')))\
                          .send_keys(Keys.ENTER)                                                          
    time.sleep(2)
    driver.save_screenshot(pathToString(screenshotsPath / "cronologicoAbastecedor.png"))
    time.sleep(2)

    driver.execute_script("window.localStorage.clear();")
    driver.quit()


   #resizeImg()
    

    
def getDataOeeFirebaseCortadora() :
    
    # Asignacion de coleccion a variable para llamar datos
    x = datetime.datetime.now()
    b = x.day - 1
    # print(b)
    fecha = datetime.datetime(2021, x.month, b)
    fechaPosterior = datetime.datetime(2021, x.month, x.day) 
    fechaFiltro = str(fecha).split()
    fechaFiltroPosterior = str(fechaPosterior).split()
    oeeDiarioCortadora = db.collection(u'oeeDiarioCortadora').where(u'fechaOee', u'==', fechaFiltro[0])
    docs = oeeDiarioCortadora.stream()

    for doc in docs:
        # print(f'{doc.id} => {doc.to_dict()}')
        diccionarioOee = doc.to_dict()
        fecha = diccionarioOee['fechaOee']
        minutosTrabajados = diccionarioOee['minutosTrabajados']
        minutosTrabajadosReales = diccionarioOee['minutosTrabajadosReales']

    promedio = (float(minutosTrabajados) * 100) / float(minutosTrabajadosReales)
    promedioRedondeado = round(promedio, 2)
    return minutosTrabajados, promedioRedondeado

def getDataOeeFirebaseAbastecedor() :
    
    # Asignacion de coleccion a variable para llamar datos
    x = datetime.datetime.now()
    b = x.day - 1
    # print(b)
    fecha = datetime.datetime(2021, x.month, b)
    fechaPosterior = datetime.datetime(2021, x.month, x.day)
    fechaFiltro = str(fecha).split()
    fechaFiltroPosterior = str(fechaPosterior).split()
    oeeDiarioAbastecedor = db.collection(u'oeeDiarioAbastecedor').where(u'fechaOee', u'==', fechaFiltro[0])
    docs = oeeDiarioAbastecedor.stream()

    for doc in docs:
        # print(f'{doc.id} => {doc.to_dict()}')
        diccionarioOee = doc.to_dict()
        fecha = diccionarioOee['fechaOee']
        productosRealizados = diccionarioOee['productosRealizados']
        productosRealizadosReales = diccionarioOee['productosRealizadosReales']

    promedio = (float(productosRealizados) * 100) / float(productosRealizadosReales)
    promedioRedondeado = round(promedio, 2)
    return productosRealizados, promedioRedondeado
# getDataOeeFirebase()




def sendEmail():
    
    #OBTENCION DE DATOS DE FIREBASE POR MAQUINA
    #---------------------------------------
    minutosCortadora, promedioCortadora = getDataOeeFirebaseCortadora()
    productosAbastecedor, promedioAbastecedor = getDataOeeFirebaseAbastecedor()
    time.sleep(5)
    #---------------------------------------
    
    #FORMATO DE FECHA
    #---------------------------------------
    x = datetime.datetime.now()
    b = x.day - 1
    fecha = datetime.datetime(2021, x.month, x.day)
    fechaPosterior = datetime.datetime(2021, x.month, b)
    fechaFiltro = str(fecha).split()
    fechaFiltroPosterior = str(fechaPosterior).split()
    #---------------------------------------
    
    #CABECERA DE EMAIL
    #---------------------------------------
    message = MIMEMultipart('alternative')
    message['Subject'] = 'Máquinas| OEE y Razones de Paros Diarios -' + str(fechaFiltro[0]) #diaActual 2021-03-13
    message['From'] = 'wrojas@box.com.ec'
    message['To'] = 'demonpogp@gmail.com'
    #---------------------------------------

    #OBTENCION DE IMAGENES RECORTADAS
    #---------------------------------------
    tableroCortadoraImg = open(pathToString(screenshotsPath / "tableroCortadora.png"), 'rb')
    msgTImageCortadora = MIMEImage(tableroCortadoraImg.read())
    tableroCortadoraImg.close()
    cronologicoCortadoraImg = open(pathToString(screenshotsPath / "cronologicoCortadora.png"), 'rb')
    msgCImageCortadora = MIMEImage(cronologicoCortadoraImg.read())
    cronologicoCortadoraImg.close()

    tableroAbastecedorImg = open(pathToString(screenshotsPath / "tableroAbastecedor.png"), 'rb')
    msgTImageAbastecedor = MIMEImage(tableroAbastecedorImg.read())
    tableroAbastecedorImg.close()
    cronologicoAbastecedorImg = open(pathToString(screenshotsPath / "cronologicoAbastecedor.png"), 'rb')
    msgCImageAbastecedor = MIMEImage(cronologicoAbastecedorImg.read())
    cronologicoAbastecedorImg.close()
    #---------------------------------------

    #ESTRUCTURA DE MAQUINA CORTADORA PARA EMAIL 
    #---------------------------------------
    messageTableroCortadora = '<p><strong>M&aacute;quina - Cortadora 5 - '+ str(fechaFiltroPosterior[0]) + ' 07:00 - ' + str(fechaFiltro[0]) + ' 07:00</strong></p> \
    <table style="font-family: arial, sans-serif; border-collapse: collapse; width: 30%;"> \
        <tr> \
            <th style="border: 1px solid #dddddd;  text-align: left;  padding: 8px;">Turno</th> \
            <th style="border: 1px solid #dddddd;  text-align: left;  padding: 8px;">Total</th> \
            <th style="border: 1px solid #dddddd;  text-align: left;  padding: 8px;">OEE</th> \
        </tr> \
        <tr> \
            <td style="border: 1px solid #dddddd;  text-align: left;  padding: 8px; background-color: #dddddd;"> Turno 1 y 2 </td> \
            <td style="border: 1px solid #dddddd;  text-align: left;  padding: 8px; background-color: #dddddd;">' + str(minutosCortadora) + '</td> \
            <td style="border: 1px solid #dddddd;  text-align: left;  padding: 8px; background-color: #dddddd;">' + str(promedioCortadora)+ '%</td> \
        </tr> \
    </table> <img alt="Tablero" src="cid:image1">'
    messageCronologicoCortadora = '<p><strong>Cronol&oacute;gico de Producci&oacute;n - Cortadora 5 - '+ str(fechaFiltroPosterior[0]) + ' 07:00 - ' + str(fechaFiltro[0]) + '\
    21:00</strong></p> <img alt="Cronol&oacute;gico de Producci&oacute;n" src="cid:image2"> '
    #---------------------------------------
    
    #ESTRUCTURA DE MAQUINA ABASTECEDOR PARA EMAIL 
    #---------------------------------------
    messageTableroAbastecedor = '<p><strong>M&aacute;quina - Abastecedor - '+ str(fechaFiltroPosterior[0]) + ' 07:00 - ' + str(fechaFiltro[0]) + ' 07:00</strong></p> \
    <table style="font-family: arial, sans-serif; border-collapse: collapse; width: 30%;"> \
        <tr> \
            <th style="border: 1px solid #dddddd;  text-align: left;  padding: 8px;">Turno</th> \
            <th style="border: 1px solid #dddddd;  text-align: left;  padding: 8px;">Total</th> \
            <th style="border: 1px solid #dddddd;  text-align: left;  padding: 8px;">OEE</th> \
        </tr> \
        <tr> \
            <td style="border: 1px solid #dddddd;  text-align: left;  padding: 8px; background-color: #dddddd;"> Turno 1 y 2 </td> \
            <td style="border: 1px solid #dddddd;  text-align: left;  padding: 8px; background-color: #dddddd;">' + str(productosAbastecedor) + '</td> \
            <td style="border: 1px solid #dddddd;  text-align: left;  padding: 8px; background-color: #dddddd;">' + str(promedioAbastecedor)+ '%</td> \
        </tr> \
    </table> <img alt="Tablero Abastecedor" src="cid:image3">'
    messageCronologicoAbastecedor = '<p><strong>Cronol&oacute;gico de Producci&oacute;n - Abastecedor - ' + str(fechaFiltroPosterior[0]) + ' 07:00 - ' + str(fechaFiltro[0]) + '\
    21:00</strong></p> <img alt="Cronol&oacute;gico Abastecedor" src="cid:image4"> '
    #---------------------------------------
    #---------------------------------------
    message.attach(MIMEText(messageTableroCortadora + messageCronologicoCortadora + messageTableroAbastecedor + messageCronologicoAbastecedor, 'html'))
    msgTImageCortadora.add_header('Content-ID', '<image1>')
    message.attach(msgTImageCortadora)
    msgCImageCortadora.add_header('Content-ID', '<image2>')
    message.attach(msgCImageCortadora)
    msgTImageAbastecedor.add_header('Content-ID', '<image3>')
    message.attach(msgTImageAbastecedor)
    msgCImageAbastecedor.add_header('Content-ID', '<image4>')
    message.attach(msgCImageAbastecedor)
    #---------------------------------------
    
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login('testingemailscript41@gmail.com', '1q2w3e4r5t6.')
    server.sendmail('demonpogp@gmail.com', 'demonpogp@gmail.com', message.as_string())
    server.quit()
    #wrojas@box


#sendEmail()
entryWeb(urlLogin, urlGeneralTableroCortadora, urlGeneralCronologicoCortadora, urlGeneralTableroAbastecedor,urlGeneralCronologicoAbastecedor)

sendEmail()
