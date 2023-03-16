*****************************************************
Codigo de generacion de reporte y envio por WhatsApp
Enrique Estevez
*****************************************************


import pywhatkit
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from datetime import datetime
import pyodbc
import xlrd
import time
import glob
import os
import runpy
from tkinter import messagebox
import pandas as pd

dia_sem = datetime.today().weekday()
hora_act = datetime.now().time().hour

#Valida dia de la semana 0=L 6=D
if (dia_sem <= 6):

    #Valida hora del dia
    if hora_act >=10 and hora_act <=23:

        dir = r'C:\Users\15052\Downloads'
        for f in os.listdir(dir):
            os.remove(os.path.join(dir, f))

        conex = pyodbc.connect( 'Driver={SQL Server};'
                                    'Server=ip;'
                                    'Database=Caribu;'
                                    'UID=user;'
                                    'PWD=pass;'
                                    'Trusted_Connection=no;')
        
        cursor = conex.cursor()
        
        # #Llamar Controlador de Chrome

        s=Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=s)
        driver.maximize_window()

        #Ingresa a liga ECRM
        driver.get("http://172.16.9.10/ecrm")
        time.sleep(10)

        #Iniciar Sesion

            #Selecciona el Frame con el que va  trabajar
        driver.find_element(By.XPATH, '//*[@id="srcMain"]')
        driver.switch_to.frame('srcMain')

            #Buscar Caja de Usuario e ingresa el valor Fijo
        driver.find_element(By.XPATH, '//*[@id="txtUsuario"]').send_keys('80008')
            
            #Buscar Caja de Password e ingresa valor Fijo
        driver.find_element(By.XPATH, '//*[@id="txtClave"]').send_keys('2695')

            #Buscar botón de Login, da click
        driver.find_element(By.XPATH, '//*[@id="NeoIngresarButton1"]').click()
        time.sleep(6)

            #Cambia Frame de trabajo
        driver.switch_to.default_content()
        driver.switch_to.frame('srcTop')

            #Selecciona la opcion "Upsell" de la barra superior
        driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/div/select/option[4]').click()
        time.sleep(2)

            #Cambia Frame de trabajo, seleccionando Menú lateral
        driver.switch_to.default_content()
        driver.switch_to.frame('srcMain')

            #Selecciona opción del menú lateral
        driver.find_element(By.XPATH, '//*[@id="NeoWebMenu1WebMenu1_8"]/td/div').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="NeoWebMenu1WebMenu1_8_1"]/td/div').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="NeoWebMenu1WebMenu1_8_1_3"]/td/div').click()
        time.sleep(4)

            #Cambia frame de trabajo
        driver.switch_to.frame(1)

            #Opciones de reporte

                #Selecciona tipo de vista
        driver.find_element(By.CSS_SELECTOR, '#Vista > option:nth-child(4)').click()

                #Check Filtro por fecha
        driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr/td/table/tbody/tr/td/div[1]/div/div[2]/span[4]/input').click()

                #Selecciona Fecha de inicio de Reporte
        driver.find_element(By.XPATH, '//*[@id="Desde_img"]').click()
        driver.find_element(By.XPATH, '//*[@id="Desde_DrpPnl_Calendar1_508"]').click()

                #Selecciona e ingresa Hora de inicio de Reporte
        driver.find_element(By.XPATH, '//*[@id="igtxtDesde_HORA"]').click() 
        driver.find_element(By.XPATH, '//*[@id="igtxtDesde_HORA"]').send_keys('00:00')

                #Selecciona Fecha de Fin de Reporte
        driver.find_element(By.XPATH, '//*[@id="Hasta_img"]').click()
        driver.find_element(By.XPATH, '//*[@id="Hasta_DrpPnl_Calendar1_508"]').click()

                #Selecciona e Ingresa Hora de Fin de Reporte
        driver.find_element(By.XPATH, '//*[@id="igtxtHasta_HORA"]').click() 
        driver.find_element(By.XPATH, '//*[@id="igtxtHasta_HORA"]').send_keys('23:59')

                #Descarga el reporte
        driver.find_element(By.XPATH, '//*[@id="btnEjecutar"]').click()
        time.sleep(10)

        driver.close()

        query_del = 'Delete from [Caribu].[dbo].[Gestiones] where [Fecha_Hora] > cast(getdate() as date)'

        cursor.execute (query_del)

        conex.commit()

        print("Carga Gestiones...")

        ges = glob.glob('C:\\Users\\15052\\Downloads\\*') # * Define el ultimo archivo descargado
        book = xlrd.open_workbook(max(ges, key=os.path.getctime))
        
        sheet = book.sheet_by_name("Sheet1")

        query = """INSERT INTO [Caribu].[dbo].[Gestiones] ([DN_Gestion],[Base_Datos],[Fecha_Hora],[Resolucion],[Planes_Gestion],
                [Nombre_Planes],[IDCampana],[Campana],[No_Usuario],[Usuario],[Motivo],[Tipo_Llamada],[Tipo_Discador],[Id_Contacto]) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""

        for r in range(1, sheet.nrows):
            
            Dato1  = sheet.cell(r,0).value
            Dato2  = sheet.cell(r,1).value
            Dato3  = sheet.cell(r,2).value
            Dato3 = xlrd.xldate_as_datetime(Dato3, 0)
            Dato4  = sheet.cell(r,3).value
            Dato5  = sheet.cell(r,4).value
            Dato6 = sheet.cell(r,5).value
            Dato7 = sheet.cell(r,6).value
            Dato8 = sheet.cell(r,7).value
            Dato9 = sheet.cell(r,8).value
            Dato10 = sheet.cell(r,9).value
            Dato11 = sheet.cell(r,10).value
            Dato12 = sheet.cell(r,11).value
            Dato13 = sheet.cell(r,12).value
            Dato14 = sheet.cell(r,13).value
                           
            values = (Dato1, Dato2, Dato3, Dato4, Dato5, Dato6, Dato7, Dato8, Dato9, Dato10, Dato11, Dato12, Dato13, Dato14)

            cursor.execute(query, values)
            conex.commit()
        
        print("...Gestiones actualizadas")
            
        dir = r'C:\Users\15052\Downloads'
        for f in os.listdir(dir):
            os.remove(os.path.join(dir, f))

        vista_out = "execute [Caribu].[dbo].[sp_Corte_Caribu]"

        df = pd.read_sql(vista_out, conex)
        
        df_frame = pd.DataFrame(df)

        df.to_excel(r'C:\Users\15052\Desktop\Aut_Caribu\BDD.xlsx', index=False, sheet_name="BDD")

        print(df)

        archivo = (r"C:\Users\15052\Desktop\Aut_Caribu\Caribu.xlsm")

        os.startfile(archivo)

        time.sleep(20)

        #Cierra Excel

        os.system('taskkill /F /IM Excel.exe')

        #Asigna grupo al que se va a mandar los mensajes

        id1 = 'FZevQWDuA8k1tZChue8iHt'
        id2 = 'HiwrsRIjX8DC424AVohxS5'

        # Busca el ultimo archivo con temrinacion OUT.jpg e IN.jpg

        img_U = glob.glob('C:\\Users\\15052\\Desktop\\Aut_Caribu\\Cortes\\*.jpg') # * Define el ultimo archivo descargado
        img_UP = max(img_U, key=os.path.getctime)

        pywhatkit.sendwhats_image(id1,img_path=img_UP,caption="Corte",wait_time=10, tab_close= False, close_time=5)

        time.sleep(5)

        pywhatkit.sendwhats_image(id2,img_path=img_UP,caption="Corte Caribu %s"%hora_act,wait_time=10, tab_close= False, close_time=5)

        time.sleep(10)

        print("Corte enviado")