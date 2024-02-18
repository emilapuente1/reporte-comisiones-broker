# -*- coding: utf-8 -*-
"""
Created on Sun Dec 24 08:50:56 2023

@author: Asesor ALIANO
"""

import pandas as pd
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime as dt
import sqlite3

# iMPORTAMOS EL EXCEL DE COMISIONES

liquidacion=False

#Ingresar el tipo de cambio al momento que se quiere facturar
tc=1176

#Ingresar el mes del que se quiere hacer el reporte
mes='Febrero'

#El excel debe contener el mes en el nombre
excelcomisiones_df = pd.read_excel(
    "Comisiones Febrero.xlsx",
    header=[0],
    engine="openpyxl"
    )
reportecomisiones = pd.DataFrame(excelcomisiones_df.groupby(["Cuenta"])["ComisionDolarizada"].sum())
reportecomisiones['facturacion'] = reportecomisiones['ComisionDolarizada']/2*tc

# Importamos el Listado de cuentas

listado_clientes_df = pd.DataFrame(pd.read_excel("Listado Cuentas.xlsx",
                                    header=[0],
                                    engine="openpyxl"
                                    ))
#Unimos los dos df y limpiamos los nan
listado_clientes_df = listado_clientes_df.rename(
            columns={
            "Descripcion":"Cuenta"
            })
df_clientes = listado_clientes_df[["Cuenta","Asesor"]]
df_comisiones_clientes= reportecomisiones.merge(df_clientes,on="Cuenta",how='right')
df_comisiones_clientes_nan= df_comisiones_clientes.dropna(subset=['facturacion'])



total_por_asesor =  pd.DataFrame(df_comisiones_clientes.groupby(["Asesor"],as_index=False)["facturacion"].sum())

def monto_Asesor_A(row):
    if row['Asesor'] == 'Both':
        return row['facturacion'] / 2
    elif row['Asesor'] == 'Asesor A':
        return row['facturacion'] * 0.75
    elif row['Asesor'] == 'Asesor B':
        return row['facturacion'] * 0.25
    elif row['Asesor'] == 'Asesor C':
        return row['facturacion'] * 0.2
    elif row['Asesor'] == 'Asesor A100':
        return row['facturacion'] * 1
    else:
        return None  # Opcional: manejar casos no especificados
    
def monto_Asesor_B(row):
        if row['Asesor'] == 'Both':
            return row['facturacion'] / 2
        elif row['Asesor'] == 'Asesor A':
            return row['facturacion'] * 0.25
        elif row['Asesor'] == 'Asesor B':
            return row['facturacion'] * 0.75
        elif row['Asesor'] == 'Asesor C':
            return row['facturacion'] * 0.2
        elif row['Asesor'] == 'Asesor B100':
            return row['facturacion'] * 1
        else:
            return None  # Opcional: manejar casos no especificados
        
def monto_Asesor_C(row):
          if row['Asesor'] == 'Asesor C':
              return row['facturacion'] * 0.6
          if row['Asesor'] == 'Asesor C100':
              return row['facturacion'] * 1
          else:
              return None  # Opcional: manejar casos no especificados

def ingresos_brutos(bruto):
    descuento = bruto * 0.03
    neto = bruto - descuento
    return neto.round(2)
    
   
total_por_asesor["Asesor A"] = total_por_asesor.apply(monto_Asesor_A, axis=1)
total_por_asesor["Asesor B"] = total_por_asesor.apply(monto_Asesor_B, axis=1)
total_por_asesor["Asesor C"] = total_por_asesor.apply(monto_Asesor_C, axis=1)

# Asesor ALIANO
total_bruto_Asesor_A= total_por_asesor['Asesor A'].sum().round(2)
total_neto_Asesor_A = ingresos_brutos(total_bruto_Asesor_A)

# Asesor BLAS
total_bruto_Asesor_B= total_por_asesor['Asesor B'].sum().round(2)
total_neto_Asesor_B = ingresos_brutos(total_bruto_Asesor_B)

#Asesor C
total_bruto_Asesor_C= total_por_asesor['Asesor C'].sum().round(2)
total_neto_Asesor_C = ingresos_brutos(total_bruto_Asesor_C)

# TRANSFERENCIA A RECIBIR DE broker
transferencia_broker= reportecomisiones['facturacion'].sum().round(2)
transferencia_neto_broker = ingresos_brutos(transferencia_broker)






# Creamos el reporte en excel con el total
writer=pd.ExcelWriter('ReporteComisiones'+mes+'.xlsx')


dashboard= {'Total a cobrar':['Asesor A','Asesor B','Asesor C'],
            'Monto': [total_neto_Asesor_A,total_neto_Asesor_B,total_neto_Asesor_C]}
dashboard_df=pd.DataFrame(dashboard).to_excel(
    writer,
    sheet_name='Dashboard',
    index=False,
    engine='xlsxwriter')         

total_por_asesorxlrx= total_por_asesor.to_excel(
    writer,
    sheet_name="Total por Asesor",
    index=False,
    engine='xlsxwriter'
    )

reporteexcel=df_comisiones_clientes_nan.to_excel(
        writer, 
        sheet_name='Listado total por Cliente',
        engine='xlsxwriter',
        index=False
        )



worksheet_dashboard = writer.sheets['Dashboard']
worksheet_totalporcliente = writer.sheets['Listado total por Cliente']
worksheet_totalporasesor = writer.sheets['Total por Asesor']




writer.close()

if liquidacion == True:
    # ENVIO DE REPORTE POR MAIL 
    subject = "Informe broker Enero"
    body = f"Hola\n\n Te cuento que el total de facturacion al dia de la fecha, teniendo como parametro el Dolar MEP a: {tc} \n\n FACTURACIÓN TOTAL: $ {transferencia_neto_broker},\n\n el total a cobrar para Asesor A es de ${total_neto_Asesor_A}\n\n el total a cobrar para Asesor B es ${total_neto_Asesor_B}\n\n el total a cobrar para Asesor C es ${total_neto_Asesor_C} \n\n Todo tiene descontado el 3% de IIGG\n\n " 
    sender_email = "Juan@gmail.com"
    receiver_email = 'Roman@gmail.com'
    password = "password"

    # Crear mail
    #MIME es un standard de Internet para envio de mails
    #la variable mi_mail reserva un espacio para la creacion de un mail MIME
    #es la misma logica de writer para Excel
    mi_mail = MIMEMultipart()
    mi_mail["From"] = sender_email
    mi_mail["To"] = receiver_email
    mi_mail["Subject"] = subject

    # agregar el cuerpo del mail al mail en si
    mi_mail.attach(MIMEText(body, "plain"))

    #quiero enviar el reporte en el mail
    reporte = 'ReporteComisiones'+mes+'.xlsx'  

    #leer el archivo de reporte desde Python
    #y prepararlo en una variable para ser agregado a la variable mi_mail
    #basicamente estamos abriendo el archivo reporte
    #y poniendolo en la variable part con una codificacion necesaria
    #para que sea aceptado como attachment en un mails
    with open(reporte, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    
    encoders.encode_base64(part)

    # Agregar headers
    # los mensajes en internet tienen encabezados que describen su contenido
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {reporte}",
    )

    # agregar el attachment a la variable creada para representar el mail
    #todavia todo esto esta en memoria
    mi_mail.attach(part)
    text = mi_mail.as_string()

    # Crear una conexion segura con protocolo SSL
    # al servidor de Gmail
    #usar las credenciales que pusimos arriba para loguearnos remoto
    # una vez ves logueados realmente enviar el mail usando sendmail()
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)

    print("El total a Facturar es:", transferencia_broker)
    print ( "El total a cobrar bruto de Asesor B:",total_bruto_Asesor_B)
    print ( "El total a cobrar neto de Asesor B:",total_neto_Asesor_B)
    print ( "El total a cobrar bruto de Asesor A:",total_bruto_Asesor_A)
    print ( "El total a cobrar neto de Asesor A:",total_neto_Asesor_A)
    print ( "El total a cobrar bruto de Asesor C:",total_bruto_Asesor_C)
    print ( "El total a cobrar neto de Asesor C:",total_neto_Asesor_C)
else:
     print("no es la liquidacion defintiva, solo de muestra")