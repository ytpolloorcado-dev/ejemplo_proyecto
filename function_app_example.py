import base64
import tempfile
import uuid
import os
from docxtpl import DocxTemplate
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import smtplib 
import azure.functions as func
import sys
from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool
import google.genai as bebe
import time
import secrets
import edge_tts
from fastapi import Depends, FastAPI, HTTPException, Header, status, UploadFile, File, Form
from fastapi.responses import FileResponse, Response
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, field_validator
import pandas as pd
from datetime import datetime
from typing import Optional
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


#Template de los contratos
ruta_contrato = os.path.join(os.path.dirname(__file__),"Templates","Contrato_xd.docx")

genai_api = bebe.Client(api_key="Example")



def enviar_correo(asunto: str, cuerpo: str, ruta_adjunto: str = None):
    """Función auxiliar para gestionar el envío seguro de correos."""
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    
    msg = MIMEMultipart()
    msg['To'] = 'example_coreoxd@gmail.com'
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo, 'plain'))

    # Si se pasa una ruta de archivo, adjuntarlo
    if ruta_adjunto and os.path.exists(ruta_adjunto):
        with open(ruta_adjunto, 'rb') as f:
            adjunto = MIMEBase('application', 'octet-stream')
            adjunto.set_payload(f.read())
            encoders.encode_base64(adjunto)
            adjunto.add_header(
                'Content-Disposition',
                f'attachment; filename="{os.path.basename(ruta_adjunto)}"'
            )
            msg.attach(adjunto)

    try:
        server = smtplib.SMTP(host=smtp_server, port=smtp_port)
        server.starttls()
        server.login('ejemplo_correo_con_os.getenv.com', 'la clave de mi aplicacion')
        server.send_message(msg)
        server.quit()
        print("Correo enviado exitosamente.")
    except Exception as err:
        print(f"Error al enviar correo: {err}")



# 1. Instancia del modelo V2 de Azure
app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

# 2. Instancia de FastAPI
fastapi_app = FastAPI()

consulta_plana = text("""
SELECT * FROM Empleados;
""")

#Desde mi punto de vista, la cadena de conexion mssql la pongo en un .env, esto es solo un ejemplo de lo que creo

conexion = create_engine(
    "mssql+pymssql://usuario_asql:contrasena_sql@servidor:puerto/base_de_datos",
    poolclass=NullPool,
    pool_pre_ping=True
)


@fastapi_app.get("/reembolsos")
def get_data():
    try:
        with conexion.connect() as conn:
            consultar = pd.read_sql(consulta_plana, conn)
            # FastAPI convierte automáticamente diccionarios y listas a JSON
            return consultar.to_dict(orient="records")
    except Exception as e:
        sys.stderr.write(f'Error al ejecutar la consulta: {e}\n')
        return {"error": f"Error al ejecutar la consulta: {str(e)}"}



@fastapi_app.get("/consulta")
def consultar():
    try:
        with conexion.connect() as conn:
            query_plana = text("SELECT * FROM Empleados")
            request_db = pd.read_sql_query(query_plana, conn)

            if request_db.empty:
                body = "No hay registros en la base de datos de empleados, por favor revisar la base de datos."
                enviar_correo("Consulta de Empleados vacía", body)

                return {
                    "status": 200, 
                    "message": "success",
                    "body": "no hay registros en la base de datos",
                    "Author": "Jose Gonzalez"
                }

            else:
                body = f"""Adjunto JSON escrito con la consulta de empleados realizada exitosamente:

                    {request_db.to_dict(orient='records')}

                    Reciban un saludo,
                    Jose Gonzalez"""
                enviar_correo("Consulta de Empleados exitosamente realizada", body)

                return {
                    "status": 200,
                    "message": "success, se ha enviado el correo exitosamente",
                    "body/value": request_db.to_dict(orient="records"),
                    "Author": "Jose Gonzalez"
                }

    except Exception as e:
        body = f"""Hubo un error que tienes que checar en tu endpoint de consulta de empleados, el error es el siguiente: 

        {str(e)}

        Reciban un saludo,
        Jose Gonzalez"""
        enviar_correo("Hubo un error al ejecutar el endpoint de consulta de empleados", body)

        return {
            "status": 500,
            "message": "error",
            "response": str(e),
            "Author": "Jose Gonzalez"
        }

@fastapi_app.get("/enviar_contrato")
def contrato_template_gmail():
    try:
        with conexion.connect() as conn:
            query_transform_contrato = text("""
            SELECT *,TRIM(Apellido)
                    + ' '
                        +(NOMBRE)
                    AS
                        Nombre_Completo,
                            FORMAT(
                                FechaIngreso
                                    ,'dd " de " MMMM " del " yyyy'
                                ,'es-ES'
                                )
                            AS
                                Fecha_Ingreso_Letra
                            ,
                        CAST(
                                YEAR(Getdate())
                                    -
                                YEAR(FechaIngreso)
                            AS
                                NVARCHAR(3)
                            )
                            +
                        ' año(s)'
                        AS
                    Antiguedad_Letra
                FROM
            Empleados;
            """
            )
            consulta_contrato = pd.read_sql_query(query_transform_contrato, conn)
            if consulta_contrato.empty:
                body = "No hay registros en la base de datos de empleados, por favor revisar la base de datos."
                enviar_correo("Consulta de Empleados vacía", body)

                return {
                    "status": 200, 
                    "message": "success",
                    "body": "no hay registros en la base de datos",
                    "Author": "Jose Gonzalez"
                }
            else:
                for _,nuevo in consulta_contrato.iterrows():
                    doc = DocxTemplate(ruta_contrato)
                    doc.render(nuevo.to_dict())
                    nombre_archivo = f"contrato_{nuevo['Nombre_Completo']}_{uuid.uuid1()}.docx"
                    carpeta_pipeline = os.path.join(tempfile.gettempdir(), "Pipeline")
                    
                    # 2. Creación garantizada de la carpeta si no existe
                    os.makedirs(carpeta_pipeline, exist_ok=True)
                    
                    # 3. Ruta final editable en Azure/Linux
                    nueva_ruta = os.path.join(carpeta_pipeline, nombre_archivo)
                    
                    # 4. Guardar archivo
                    doc.save(nueva_ruta)


                    body = f"""Se ha generado el contrato de la siguiente persona: 
                    
                    Nombre: {nuevo['Nombre_Completo']},
                    Fecha de Ingreso: {nuevo['Fecha_Ingreso_Letra']},
                    Antigüedad: {nuevo['Antiguedad_Letra']}.

                    Adjunto el contrato en formato Word.
                    Reciban un saludo,
                    Jose Gonzalez
                    """
                    enviar_correo(asunto=f"Contrato de {nuevo['Nombre_Completo']}",cuerpo=body,ruta_adjunto=nueva_ruta)
                return {
                    "Status": 200,
                    "message": "success",
                    "body": consulta_contrato.to_dict(orient="records"),
                    "Author": "Jose Gonzalez"
                }
    except Exception as e:
        return {
            "Status": 500,
            "message": f"hubo un error al ejecutar la API: {e}",
            "Author": "Jose Gonzalez"
        }

@fastapi_app.get("/endpoint_net")

def get_data(Empleado: int):
    try:
        query_parameter = text("""
                    SELECT *,TRIM(Apellido)
                            + ' '
                                +(NOMBRE)
                            AS
                                Nombre_Completo,
                                    FORMAT(
                                        FechaIngreso
                                            ,'dd " de " MMMM " del " yyyy'
                                        ,'es-ES'
                                        )
                                    AS
                                        Fecha_Ingreso_Letra
                                    ,
                                CAST(
                                        YEAR(Getdate())
                                            -
                                        YEAR(FechaIngreso)
                                    AS
                                        NVARCHAR(3)
                                    )
                                    +
                                ' año(s)'
                                AS
                            Antiguedad_Letra
                        FROM
                    Empleados
                            WHERE
                        EmpleadoID =
                        :empleado
                    """)


        with conexion.connect() as c:
            
            consulta_plana = pd.read_sql(
            query_parameter,c,params={"empleado": Empleado})

        return {
            "status":202,
            "message": "Sucess",
            "body/value":consulta_plana.to_dict(orient="records"),
            "Author": "Jose Gonzalez"
            }
    
    except Exception as error:
        return {
        "Status": 502,
        "Message":f"Hubo un error al consumir : {error}",
        "Author": "Josè GOnzalez"
        }



@fastapi_app.get("/mi_html.html", response_class=HTMLResponse)

def get_html():

    with open("mi_html.html","r") as adjuntar:

        return adjuntar.read()

@fastapi_app.post("/upload")

async def enviar_archvo(file: UploadFile = File(...)):

    archivo_memoria = await file.read()

    decodificar_a_base64 = base64.b64encode(archivo_memoria).decode("utf-8")

    nombre_archivo = file.filename
    tipo_contenido = file.content_type
    tamano_en_bytes = len(archivo_memoria)

    if nombre_archivo.lower().endswith(".pdf", ".jpg", ".jpeg"):

        sys.stdout.write(f'guardando archivo.... {nombre_archivo}')
        sys.stdout.flush() 

        return{
            "status": 200,
            "filename": nombre_archivo,
            "size_bytes": tamano_en_bytes,
            "Tipo_contenido": tipo_contenido,
            "mensaje":"se ha guardado el archivo satisfactoriamente",
            "archivo_memoria": decodificar_a_base64,
            "Author": "Jose GOnzalez" 
        }

    else:
        return{
            "status": 200,
            "mensaje": "Error, no podemos adjuntar en la carga util por que tu archivo no es .jpg, .png, o .jpeg. Por favor selecciona un archivo con  las extensiones ya mencionadas",
            "Author": "Jose Gonzalez"
        }


@fastapi_app.get("/grabacion.html", response_class=HTMLResponse)
def get_grabacion_html():
    texto_para_usuario = "Yo, empleado de Covia Holdings, confirmo que los documentos proporcionados son verídicos."
    with open("grabacion.html", "r", encoding="utf-8") as f:
        html = f.read().replace("{{ texto_grabar }}", texto_para_usuario)
    return html

@fastapi_app.post("/upload_video")
async def recibir_video(file: UploadFile = File(...)):

    temp_file_path = f"temp_{file.filename}"
    
    with open(temp_file_path,'wb') as video:

        content = await file.read()

        video.write(content)

    preparar_video = genai_api.files.upload(file=temp_file_path)

    chat = genai_api.chats.create(model="gemini-3.8-flash")

    response = chat.send_message(
        message=[
            preparar_video, 
            "Valida si la persona acepta el acuerdo, solo true si acepta false si no, si no habla del acuerdo recuerdale que tiene que decidir si acepta o no acepta el acuerdo"
        ]
    )

    genai_api.files.delete(name=preparar_video.name)

    if os.path.exists(temp_file_path):

        os.remove(temp_file_path)

    if response.text.lower().strip() == "true":

        return {
            "status": 200,
            "mensaje": "se ha procesado el video exitosamente",
            "respuesta": response.text 
        }
    elif response.text.lower().strip() == "false":
        return {
            "status": 200,
            "mensaje": "Muchas gracias por postularte, agradecemos el tiempo que invertiste y respetamos tu desacuerdo, puedes salir del sistema, todos tus datos seran eliminados."
        }
    else:
        return {
            "status": 200,
            "Mensaje": response.text
        }


#------------------------   New_response --------------------------------------


tts_cache = {}


class DatosChecada(BaseModel):
    empleado_id: int
    tipo: int
    time_zone: str
    content_bytes: bytes

    @field_validator("empleado_id")
    @classmethod
    def validar_emplado(cls, valor):
        if not valor:
            raise ValueError("No puede estar vacio el empleado")
        return valor


def verificar_tablet(x_api_key: str = Header(None)):
    token_esperado = os.getenv("TABLET_API_KEY", "12345")
    if not x_api_key or not token_esperado or not secrets.compare_digest(token_esperado, x_api_key):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Dispositivo no autorizado"
        )


@fastapi_app.get("/")
def mostrar_interfaz():
    ruta_index = os.path.join(os.path.dirname(__file__), "Templates", "index.html")
    return FileResponse(ruta_index)


@fastapi_app.get("/api/tts")
async def generar_voz(texto: str):
    texto_limpio = (texto or "").strip()[:200]
    if not texto_limpio:
        raise HTTPException(status_code=400, detail="Texto vacio")

    if texto_limpio in tts_cache:
        return Response(content=tts_cache[texto_limpio], media_type="audio/mpeg")

    comunicador = edge_tts.Communicate(texto_limpio, voice="es-MX-DaliaNeural", rate="+8%")
    bloques = bytearray()
    async for chunk in comunicador.stream():
        if chunk["type"] == "audio":
            bloques.extend(chunk["data"])

    audio_bytes = bytes(bloques)
    if audio_bytes:
        tts_cache[texto_limpio] = audio_bytes

    return Response(content=audio_bytes, media_type="audio/mpeg")


@fastapi_app.post("/api/validar", dependencies=[Depends(verificar_tablet)])
def validar():
    return {"status": 200}


@fastapi_app.post("/api/checar", dependencies=[Depends(verificar_tablet)])
async def registrar_asistencia(
    empleado_id: int = Form(...),
    clave_evento: int = Form(...),
    x_timezone: Optional[str] = Header(default="UTC"),
    evidencia_foto: UploadFile = File(...)
):
    foto_varbinary = await evidencia_foto.read()

    datos_obtenidos = DatosChecada(
        empleado_id=empleado_id,
        tipo=clave_evento,
        time_zone=x_timezone or "UTC",
        content_bytes=foto_varbinary
    )

    if clave_evento not in (1, 2):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Clave de evento invalida"
        )

    try:
        if clave_evento == 1:
            tipo_evento = "Entrada"
        else:
            tipo_evento = "Salida"

        #parte 2_cuando tenga con que cruzar: with conexion.connect() as conn:

            #parte 2_cuando tenga con que cruzar: texto_consulta = """
            #parte 2_cuando tenga con que cruzar:     SELECT * FROM EMPLEADOS_ACTIVOS;
            #parte 2_cuando tenga con que cruzar:     """

        #parte 2_cuando tenga con que cruzar:     consulta = pd.read_sql(texto_consulta,conexion)

        timezone = ZoneInfo(datos_obtenidos.time_zone)
        ahora = datetime.now(timezone).replace(microsecond=0, tzinfo=None)
        fecha_hora_local = ahora
        fecha_actual = ahora.date()
        hora_local = ahora.time()

        #insert table text()

        insert_table = text("""INSERT INTO Asistencias(
            CLA_TRAB,CLAVE_EVENTO,TIPO_EVENTO,FECHA,FECHA_HORA,HORA,FOTO)
            VALUES(:val_1,:val_2,:val_3,:val_4,:val_5,:val_6,:val_7)
        """)

        with conexion.begin() as conn:
            conn.execute(
                statement=insert_table,
                parameters={
                    "val_1": datos_obtenidos.empleado_id,
                    "val_2": datos_obtenidos.tipo,
                    "val_3": tipo_evento,
                    "val_4": fecha_actual,
                    "val_5": fecha_hora_local,
                    "val_6": hora_local,
                    "val_7": datos_obtenidos.content_bytes
                }
            )




        return {"mensaje": f"{tipo_evento} registrada a las {hora_local}"}

    except ZoneInfoNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Zona horaria no valida"
        )
    except Exception:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Ocurrio un error al conectarse a la base de datos"
        )



# 3. Mapeo ASGI para redireccionar todo el tráfico a FastAPI
@app.route(route="{*route}")
async def main(req: func.HttpRequest) -> func.HttpResponse:
    return await func.AsgiMiddleware(fastapi_app).handle_async(req)
