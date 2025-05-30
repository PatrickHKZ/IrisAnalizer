from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import os, sys
from selenium import webdriver
import pyperclip

def update_progress(progress):
    with open(progreso, "w") as f:
        f.write(str(progress))


textoenviado = sys.argv[1]
progreso = sys.argv[2]
inicialesanalista = sys.argv[3]
rutiniciosesion = sys.argv[4]
contraseñainiciosesion = sys.argv[5]




update_progress(10)


def extraerinfo(reclamo,timeout=30):
    driver.execute_script("Cargar('/new-sistema/OIRS2/Modulos/BandejaAsignados/SolicitudesClientes_2.asp?numSolicitud="+reclamo+"','centro')")
    WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return jQuery.active == 0")
    )
    elemento = driver.find_element(By.CSS_SELECTOR, ".table.table-bordered > tbody > tr:nth-child(8) > td:nth-child(2)").text 
    print(elemento)

def onlywait(timeout=30):
    WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return jQuery.active == 0")
    )

def accederreclamo(reclamo,timeout=30):
    link=f"Cargar('/new-sistema/OIRS2/Modulos/BandejaAsignados/SolicitudesClientes_2.asp?numSolicitud={reclamo}','centro')"
    driver.execute_script(link)
    WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return jQuery.active == 0")
    )

def responder(documento,documento2,reclamo):

    ruta_archivo = os.path.abspath(f"data/pdf/{documento}.pdf")
    ruta_archivo2 = os.path.abspath(f"data/pdf/{documento2}.pdf")

    if  os.path.exists(ruta_archivo):
        if  os.path.exists(ruta_archivo2):
            onlywait()
            accederreclamo(reclamo)
            onlywait()
            
            driver.find_element(By.CSS_SELECTOR, ".btn.btn-success.btn-md.openBtncomentario").click()
            
            onlywait()
            
            # Localizar input comentario
            inputComentario=driver.find_element(By.ID, "descripcion2")

            # Localizar el input file
            inputFile = driver.find_element(By.XPATH, "//input[@type='file']")

            # Localizar input nombre archivo
            inputNameFile=driver.find_element(By.ID, "descripcion")

            # Localizar btn agregar archivo
            btnAddFile=driver.find_element(By.CSS_SELECTOR, ".btn.btn-primary")

            # Btn Derivar1
            btnderivar1=Select(driver.find_element(By.ID, "idproximoestado"))

            #Btn Cerrar
            btnCerrar=driver.find_element(By.CSS_SELECTOR, ".btn.btn-default")
        
            # Comenzar a responder
            onlywait()
            inputFile.send_keys(ruta_archivo)
                #Comentar
            driver.execute_script(f'document.getElementById("descripcion2").value="Estimado/a Cliente: Junto con saludar, y esperando que se encuentre bien, adjuntamos documentos correspondientes a Nota de Crédito y Refacturación. Por favor cerrar reclamo. {inicialesanalista}."')
                #Adjuntar archivo

            
                #Nombrar archivo
            inputNameFile.send_keys(documento)

                #Añadir Archivo
         
            btnAddFile.click()

            print("Archivo adjuntado correctamente")
            

            time.sleep(1)
            
                #Nombrar archivo
            driver.execute_script(f'document.getElementById("descripcion").value="{documento2}"')
           
            inputFile.send_keys(ruta_archivo2)
                #Añadir Archivo
         
            btnAddFile.click()
            print("Archivo adjuntado correctamente")


            

                #Derivar1
            btnderivar1.select_by_value("2")
            onlywait()
                #Derivar2
            btnderivar2=Select(driver.find_element(By.ID, "idarea"))
            btnderivar2.select_by_value("1")
            onlywait()
                #Derivar3
            btnderivar3=Select(driver.find_element(By.ID, "rutusuario"))
            btnderivar3.select_by_index(1)
            onlywait()
                #btnGuardar
            btnguardar=driver.find_element(By.CSS_SELECTOR, "input#botonguardar")
            btnguardar.click()
            onlywait()
                #Cerrar
            btnCerrar.click()
            onlywait()
        else:
            print(f"Documento {documento2} no encontrado, reclamo {reclamo} no respondido")
    else:
        print(f"Documento {documento} no encontrado, reclamo {reclamo} no respondido")

    return list






try:
    driver
except:
    driver = webdriver.Chrome(options=webdriver.ChromeOptions().add_experimental_option("detach",True))
    driver.get("https://aplicacionesweb.cenabast.cl/new-sistema/OIRS2/index.asp")
    driver.execute_script(f'document.getElementById("usuario").value="{rutiniciosesion}";document.getElementById("password").value="{contraseñainiciosesion}";document.getElementById("iniciar").click();')



respuestas=textoenviado


docyrec=[]



lineas=respuestas.split('\n')
lineasreparadas=[]



for linea in lineas:
    linea=linea.strip()
    if linea!='':
        lineasreparadas.append(linea)

n=0



for i in lineasreparadas:
    print(i)
    
    if n%3==0:
                    
        docyrec.append((lineasreparadas[n], lineasreparadas[n+1],lineasreparadas[n+2]))
    n+=1
  



print("=============") 


count=1
for rec,doc,doc2 in docyrec:
    responder(doc,doc2,rec)
    porcent=len(docyrec)
    porcent=60/porcent
    update_progress(int((porcent*(count))+10))
    count+=1



driver.quit()
update_progress(90)
time.sleep(1)
update_progress(100)


