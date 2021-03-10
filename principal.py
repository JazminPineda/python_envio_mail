
#Libreias 
import smtplib, ssl # PROTOCOLO envio correos electronicos y ssl conexion segura  
import getpass
from email import encoders #importar docts
from email.mime.base import MIMEBase #envair archivos codificados
from email.mime.text import MIMEText #para envio correo electronico 
from email.mime.multipart import MIMEMultipart
# Importo libreria 
import openpyxl



# Lee el archivo con metodo "cargar libro de trabajo", dataonly para solo leer datos no formulas, solo resultados
book = openpyxl.load_workbook('proyec1.xlsx', data_only =True) 
# fijar la hoja 
hoja = book.get_sheet_by_name('nueva') #Cerrar excel antes de ejecutar codigo y este es para indicar cual hoja se va a leer
# hoja = book.active para la hoja activa de excel
celdas  = hoja['A2':'F201']# lee un rango determinado de celdas

lista_empleados = []

for fila in celdas:
    empleado = [celda.value for  celda in fila] # for anidado simplificado comprensión de listas, se lee el VALOR con el metodo
    lista_empleados.append(empleado) #metodo añade elemento a una lista

# pedir datos para enviar correo  

username = input("Ingresar su nombre usuario: ")
password = getpass.getpass("Ingresar su nombre password: ") #ingreso contraseña segura
#destinatario = lista_empleados[empleado(3)] #input("Ingrese destinatario: ")
asunto = input("Agregar Asunto: ")

#Se crea mensaje 
mensaje = MIMEMultipart("alternative") #hay varios este es el estandar 
mensaje["Subject"] = asunto
mensaje["From"] = username
#mensaje["To"] = destinatario#


#creo variable donde guardo la ruta donde esta el adjunto, se coloca ruta 
archivo = "Captura.jpg" 
#adjunto contenido y lo lee en bites
with open(archivo, "rb") as adjunto: 
    #crea el contenido de tipo base
    contenid_adjunto = MIMEBase("application", "octec-strem")#contenido bites MIME, se conecta como una aplicac y string de octetos imagen o archivo no decifrarlo ni codificarlo 
    #adjunta archivo al contenido base
    contenid_adjunto.set_payload(adjunto.read())

# se tiene que codificar es un estandar 
encoders.encode_base64(contenid_adjunto)

# añade encabezado del adjunto
contenid_adjunto.add_header(
    "Content-Disposition", f"attachment; filename= {archivo}"
)
# se agrega contenido al mensaje 
mensaje.attach(contenid_adjunto)


#crear conexión segura 
context = ssl.create_default_context() # todo lo que ocurra en el contexto va hacer una conexion segura y se le agrega ss.

# creacion de correo electronico de forma segura
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context= context) as server: # A que servidor conectar, puerto de "SMTP" se 
# desarrolla en una aplicacion y se resumen en la variable server#  
 
    server.login(username, password) #se inicia secion con los ingresado por usuario 
    print("Incio sesión ")
    for empleado in lista_empleados:
        destinatario = empleado[3]

        # Codig de html 
        html = f"""
        <html>
        <body>
        ...<p> Hola! te tenemos una  <br>
        ... <b> NOTICIA </b> 
        Es grato informarte {empleado[2]} {empleado[1]}, que tu nuevo salario es {empleado[5]}
        </body>
        </html>
        """

        #contenido del mensaje en HTML    
        parte_html = MIMEText( html, "html")

        #Agrega contenido al mensaje 
        mensaje.attach(parte_html)
        mensaje_final = mensaje.as_string() # variable final y empaquetado  mensaje.as_string()

        #mensaje = f'Es grato informarte {empleado[2]} " "{empleado[1]} que tu nuevo salario es {empleado[5]} '
        server.sendmail(username, destinatario, mensaje_final)# se utiliza metodo enviar mail donde arma correo
        print("Mensaje enviado")

