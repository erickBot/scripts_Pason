from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import openpyxl
import time
import os
from os import remove
from datetime import datetime, date, time, timedelta
import calendar

mylist = []
user = []
usuarios = []

def run():
    enviarEmail = False
    #nombre de archivo por abrir
    name_file = "Lista_de_Sensores_HGAS.xlsx"
    file = openpyxl.load_workbook(name_file)
    #en una lista guarda los sheet(hoja) del archivo excel
    list_sheetbyName = file.sheetnames
    #itera cada sheet(hoja) por nombre y lo almacena en variable sheet
    for name in list_sheetbyName:
        sheet = file.get_sheet_by_name(name)
        #lee celda especifica de la hoja y lo almacena en las variables porvencer y vencidos
        porVencer = sheet['R3'].value
        vencidos  = sheet['R4'].value
        #en una lista llamada mylist, guardamos los valores almacenados 
        mylist.append(porVencer)
        mylist.append(vencidos)  
        #iteramos lista para mirar si hay un valor que sea mayor que cero
        for x in mylist:
        #si hay algun valor mayor que cero dentro de la lista, entonces se debe enviar correo con reporte, caso
        #contrario no debe enviar
            if x >0:
                enviarEmail = True
    #finalmente evalua si la variable enviarEmail esta en True o false
    #inicialmente esta en false, si esta en True es porque se debe enviar el reporte por correo
    if (enviarEmail):
        #salta a esta funcion
        procesaDestinatarios()
    #cierra el archivo excel y termina el programa
    file.close()

def procesaDestinatarios():
    cadenaStr = ''
    #abrimos archivo txt, alli estan los email de destino
    file = open('users.txt', 'r')
    #leemos linea por linea
    for row in file:
        #almacena en una lista "user", los email de destino
        user.append(row)
    #cerramos archivo txt
    file.close()
    #itera lista "user" hasta la penultima variable dentro de la lista, 
    #ya que la ultima variable de la lista, es la ultima linea del archivo (users.txt) que no se considera
    for x in range(len(user) - 1):
        email = str(user[x])
        #itera cada email de destino y se guarda como string en la variable email
        #ahora itera el string, caracter por caracter de la variable email, hasta el penultimo caracter
        #ya que el ultimo caracter es (nueva linea: \n)
        for y in range(len(email) - 1):
            #concatena cada caracter para formar el email de destino y lo almacena en la variable cadenaStr
            cadenaStr = cadenaStr + email[y]
        #finalmente casa email dentro de la variable cadenaStr lo agrega a la lista usuarios
        usuarios.append(cadenaStr)
        #limpia la variable cadenaStr, para repetir el proceso, hasta completar de leer todos los email de
        #la lista (user)
        cadenaStr = ''
    #despues de guardar todos los email de destino dentro de la lista usuarios,
    #iteramos la lista usuarios para iniciar el proceso de enviar el reporte a los email de destino
    #uno por uno
    for email in usuarios:
        #como argumento de la funcion (sendEmail) esta la variable email, que contiene el email de destino
        sendEmail(email)

def sendEmail(email):
    #crea la instancia del objeto de mensaje
    msg = MIMEMultipart()
    message = 'FYI\n\nUn script de Paca Systems'
    ruta_adjunto = "Lista_de_Sensores_HGAS.xlsx"
    nombre_adjunto = "Lista_de_Sensores_HGAS.xlsx"
    #configura los parametros del mensaje
    password = "99e12438cf"
    msg['From'] ="rpsysmanager@gmail.com"
    msg['To'] = email
    msg['Subject'] = "Lista de Sensores HGAS por vencer y expirados en Talara"
    #agrega el cuerpo del mensaje
    msg.attach(MIMEText(message, 'plain'))
    # Abrimos el archivo que vamos a adjuntar
    archivo_adjunto = open(ruta_adjunto, 'rb')
    # Creamos un objeto MIME base
    adjunto_MIME = MIMEBase('application', 'octet-stream')
    # Y le cargamos el archivo adjunto
    adjunto_MIME.set_payload((archivo_adjunto).read())
    # Codificamos el objeto en BASE64
    encoders.encode_base64(adjunto_MIME)
    # Agregamos una cabecera al objeto
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    # Y finalmente lo agregamos al mensaje
    msg.attach(adjunto_MIME)
    #crear servidor
    server = smtplib.SMTP('smtp.gmail.com: 587')
    server.starttls()
    #Ingresa credenciales para enviar email
    server.login(msg['From'], password)
    #envia el mensaje al servidor
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
    print ("Envio de email exitoso a %s:" % (msg['To']))

if __name__ == "__main__":
    run()
