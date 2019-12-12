# -*- coding: utf-8 -*-

import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template

import gspread as gs
from oauth2client.service_account import ServiceAccountCredentials


def crear_email(to, sender, asnto):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = asnto

    return message


def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


EMAIL = "noreply@colegioaltamira.cl"
PASSWORD = ""
ARCHIVO = ""

print("---------------------------------------")
print("Altabot v1.0.0 - Correos Automatizados")
print("---------------------------------------")
print()

print("INFO: Iniciando..")

print("INFO: Abriendo archivo de plantilla..")
plantilla = read_template("plantilla.txt")

print("INFO: Estableciendo conexión con los servidores..")
# Conseguimos las credenciales para usar la API de Google Drive
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive',
         'https://www.googleapis.com/auth/gmail.readonly']

print("INFO: Autenticando..")
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
cliente = gs.authorize(creds)

print("INFO: Consiguiendo documentos de emails..")
archivo_emails = None
try:
    archivo_emails = cliente.open(ARCHIVO).sheet1
except gs.exceptions.SpreadsheetNotFound:
    print("ERROR: No se ha encontrado el archivo de correos")
    exit(1)

nombres_raw = archivo_emails.col_values(1)[1:]
nombres = []
apellidos = []
for nom_apellido in nombres_raw:
    bannana = nom_apellido.split(" ")
    nombres.append(bannana[0])

    if len(bannana) >= 3:
        apellidos.append(bannana[1] + " " + bannana[2])
    else:
        apellidos.append(bannana[1])

emails = archivo_emails.col_values(2)[1:]

print("INFO: Encontrados %d correos objetivos" % len(nombres))

print("INFO: Estableciendo conexión con el servidor de Gmail")
s = smtplib.SMTP('smtp.gmail.com', 587)
s.starttls()
s.login(user=EMAIL, password=PASSWORD)

print("INFO: Iniciando la creación de los correos..")
correos = {}

print("INFO: Resolviendo plantilla..")
print("INFO: Consiguiendo substituciones")
campos_plantilla = {}
col_index = 3

while archivo_emails.cell(1, col_index).value != "":
    titulo = archivo_emails.cell(1, col_index).value
    col = archivo_emails.col_values(col_index)[1:]

    campos_plantilla[titulo] = col
    col_index += 1

len_plantilla = len(campos_plantilla.items())
print("INFO: Se encontraron %d campos de substitución" % len_plantilla)

print("INFO: Generando correos..")
asunto = input("Ingrese el asunto de los correos: ")
print()

start_time = time.time()

x = 0
for nombre in nombres:
    print("INFO: Generando email para %s %s (%d/%d)" % (nombre, apellidos[x], x + 1, len(nombres)))
    msg = crear_email(asnto=asunto, to=emails[x], sender=EMAIL)
    plantilla_i = plantilla


    substituciones_i = {}
    for llave in campos_plantilla.keys():
        substituciones_i[llave] = campos_plantilla[llave][x]

    substituciones_i["NOMBRE"] = nombre
    substituciones_i["APELLIDO"] = apellidos[x]

    body = plantilla_i.substitute(**substituciones_i)

    msg.attach(MIMEText(body, 'plain'))

    print("\tINFO: Enviando..")
    s.send_message(msg)
    print("\tINFO: Enviado")
    print()
    x += 1

print()
print("INFO: Listo!")
print("INFO: Enviados %d de %d correos en %d segundos" % (x, len(nombres), time.time() - start_time))
print()

print("INFO: Cerrando conexión al servidor de email..")
s.close()

print("INFO: Terminando")
exit(0)
