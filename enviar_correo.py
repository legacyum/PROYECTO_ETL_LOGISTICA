import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from dotenv import load_dotenv
from pathlib import Path
import logging
from datetime import datetime
from typing import List

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

print("Preparando el envío del reporte automatizado...")

# Cargar variables de entorno desde archivo .env
load_dotenv()

# 1. Configuración de credenciales y rutas
remitente = os.getenv('GMAIL_USER')
password = os.getenv('GMAIL_PASSWORD')
destinatarios: List[str] = os.getenv('DESTINATARIOS', '').split(',')
destinatarios = [email.strip() for email in destinatarios]

archivo_adjunto = 'reporte_ejecutivo_diario.xlsx'

# 2. Construcción del mensaje (HTML mejorado)
asunto = "📊 Reporte Ejecutivo Diario - Control de Operaciones"
cuerpo_html = f"""
<html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; border: 1px solid #ddd; border-radius: 8px; padding: 20px;">
            <div style="background-color: #366092; color: white; padding: 20px; border-radius: 8px 8px 0 0; text-align: center;">
                <h1 style="margin: 0;">📊 Reporte Ejecutivo</h1>
                <p style="margin: 5px 0 0 0; font-size: 14px;">Control de Operaciones Diario</p>
            </div>
            
            <div style="padding: 20px;">
                <p>Hola Equipo,</p>
                
                <p>Adjunto encontrarán el <strong>reporte ejecutivo diario</strong> con el estatus consolidado de los documentos operativos.</p>
                
                <div style="background-color: #f5f5f5; padding: 15px; border-left: 4px solid #366092; margin: 20px 0; border-radius: 4px;">
                    <h3 style="margin-top: 0;">📋 Contenido del Reporte:</h3>
                    <ul style="margin: 10px 0; padding-left: 20px;">
                        <li><strong>KPIs Principales</strong> - Métricas clave de operación</li>
                        <li><strong>Resumen por Estado</strong> - Pendientes vs. Cerrados</li>
                        <li><strong>Por Moneda</strong> - Análisis de divisas</li>
                        <li><strong>Por Año</strong> - Distribución temporal</li>
                        <li><strong>Top 10 Clientes</strong> - Principales generadores de volumen</li>
                        <li><strong>Data Detallada</strong> - Primeras 1,000 transacciones</li>
                    </ul>
                </div>
                
                <p><em>Este reporte ha sido extraído, limpiado y generado automáticamente mediante nuestro pipeline ETL de Python.</em></p>
                
                <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
                
                <div style="font-size: 12px; color: #666;">
                    <p><strong>Fecha y hora:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
                    <p><strong>Remitente:</strong> Sistema Automatizado de Reportes</p>
                </div>
            </div>
            
            <div style="background-color: #f5f5f5; padding: 15px; border-radius: 0 0 8px 8px; text-align: center; font-size: 12px; color: #666;">
                <p>Este es un correo automático. No responda a este mensaje.</p>
            </div>
        </div>
    </body>
</html>
"""

# 3. Validación de archivo
if not Path(archivo_adjunto).exists():
    logger.error(f"❌ Archivo no encontrado: {archivo_adjunto}")
    print(f"❌ Error: No se encontró el archivo '{archivo_adjunto}'.")
    exit(1)

# 4. Validación de credenciales
if not remitente or not password:
    logger.error("❌ Credenciales de Gmail no configuradas")
    print("❌ Error: Credenciales faltantes. Configura GMAIL_USER y GMAIL_PASSWORD en .env")
    exit(1)

# 5. Construcción del mensaje
def crear_mensaje(remitente_addr: str, destinatarios_list: List[str], asunto_text: str, cuerpo_html_text: str, archivo: str) -> MIMEMultipart:
    """Crea un mensaje MIME con adjunto."""
    mensaje = MIMEMultipart('alternative')
    mensaje['From'] = remitente_addr
    mensaje['To'] = ', '.join(destinatarios_list)
    mensaje['Subject'] = asunto_text
    
    # Agregar versión de texto plano como alternativa
    texto_plano = "Este correo requiere un cliente de correo con soporte HTML."
    mensaje.attach(MIMEText(texto_plano, 'plain'))
    
    # Agregar versión HTML
    mensaje.attach(MIMEText(cuerpo_html_text, 'html'))
    
    # Adjuntar archivo Excel
    try:
        with open(archivo, 'rb') as adjunto_file:
            parte = MIMEBase('application', 'octet-stream')
            parte.set_payload(adjunto_file.read())
            encoders.encode_base64(parte)
            parte.add_header('Content-Disposition', f'attachment; filename= {Path(archivo).name}')
            mensaje.attach(parte)
            logger.info(f"✅ Archivo adjuntado: {archivo}")
    except IOError as e:
        logger.error(f"❌ Error al adjuntar archivo: {e}")
        raise
    
    return mensaje

# 6. Envío del correo con reintentos
def enviar_correo(remitente_addr: str, password_auth: str, destinatarios_list: List[str], mensaje_email: MIMEMultipart, max_intentos: int = 3) -> bool:
    """Envía el correo con reintentos en caso de fallos."""
    servidor = None
    
    for intento in range(max_intentos):
        try:
            logger.info(f"Intento {intento + 1}/{max_intentos}: Conectando al servidor SMTP...")
            servidor = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
            servidor.starttls()
            servidor.login(remitente_addr, password_auth)
            
            servidor.sendmail(remitente_addr, destinatarios_list, mensaje_email.as_string())
            logger.info(f"✅ Correo enviado exitosamente a: {', '.join(destinatarios_list)}")
            return True
            
        except smtplib.SMTPAuthenticationError:
            logger.error("❌ Error de autenticación: Verifica que la contraseña de aplicación sea correcta")
            return False
        except smtplib.SMTPException as e:
            logger.warning(f"⚠️ Intento {intento + 1} falló: {e}")
            if intento < max_intentos - 1:
                logger.info("Reintentando en 5 segundos...")
                import time
                time.sleep(5)
        except Exception as e:
            logger.error(f"❌ Error inesperado: {e}")
            return False
        finally:
            if servidor:
                try:
                    servidor.quit()
                except:
                    pass
    
    logger.error(f"❌ No se pudo enviar el correo después de {max_intentos} intentos")
    return False

# 7. Ejecutar envío
try:
    logger.info(f"Preparando envío a: {', '.join(destinatarios)}")
    mensaje = crear_mensaje(remitente, destinatarios, asunto, cuerpo_html, archivo_adjunto)
    
    if enviar_correo(remitente, password, destinatarios, mensaje):
        print("✅ ¡Correo enviado exitosamente con el reporte adjunto!")
    else:
        print("❌ Error al enviar el correo. Revisa los logs para más detalles.")
        exit(1)
        
except Exception as e:
    logger.error(f"❌ Error fatal: {e}")
    print(f"❌ Error al enviar el correo: {e}")
    exit(1)