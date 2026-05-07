import pandas as pd
import numpy as np

# ==========================================
# FASE 1: EXTRACCIÓN (EXTRACT)
# ==========================================
print("Iniciando proceso ETL...")

# Cambia 'datos_crudos.csv' si le pusiste otro nombre a tu archivo de Kaggle
ruta_archivo = 'datos_crudos.csv'

try:
    df = pd.read_csv(ruta_archivo)
    print(f"✅ Archivo cargado correctamente. Filas iniciales: {len(df)}")
except FileNotFoundError:
    print("❌ Error: No se encontró el archivo CSV. Revisa que esté en la misma carpeta que este script.")
    exit()

# ==========================================
# FASE 2: TRANSFORMACIÓN Y LIMPIEZA (TRANSFORM)
# ==========================================

# 1. Eliminar registros duplicados (Regla básica de auditoría documental)
df_limpio = df.drop_duplicates()

# 2. Estandarizar cabeceras de columnas (minúsculas y sin espacios)
df_limpio.columns = df_limpio.columns.str.strip().str.lower()

# 3. Tratamiento de valores críticos
# Identifica cómo se llama la columna principal en tu CSV (ej. 'invoice_id', 'doc_id', 'order_id')
# Reemplaza 'invoice_id' con el nombre real de tu columna
columna_id = 'invoice_id' 

if columna_id in df_limpio.columns:
    # Eliminamos cualquier fila que no tenga un número de documento válido
    filas_antes = len(df_limpio)
    df_limpio = df_limpio.dropna(subset=[columna_id])
    filas_eliminadas = filas_antes - len(df_limpio)
    print(f"✅ Filas sin ID de documento eliminadas. ({filas_eliminadas} filas removidas)")
else:
    print(f"⚠️ Advertencia: Columna '{columna_id}' no encontrada.")

# 4. Clasificación automática de estados
# Si tienes una columna de fecha de cierre o pago (ej. 'clear_date'), podemos saber qué falta atender
columna_cierre = 'clear_date'

if columna_cierre in df_limpio.columns:
    # Si la celda está vacía (nula), el estado es 'Pendiente', si no, es 'Cerrado'
    df_limpio['estado_operativo'] = np.where(df_limpio[columna_cierre].isna(), 'Pendiente', 'Cerrado')
    pendientes = (df_limpio['estado_operativo'] == 'Pendiente').sum()
    cerrados = (df_limpio['estado_operativo'] == 'Cerrado').sum()
    print(f"✅ Columna de 'estado_operativo' generada automáticamente. (Pendientes: {pendientes}, Cerrados: {cerrados})")
else:
    print(f"⚠️ Advertencia: Columna '{columna_cierre}' no encontrada.")

# ==========================================
# FASE 3: CARGA (LOAD)
# ==========================================

# Exportar la data ya limpia y estructurada a un nuevo archivo
nombre_salida = 'datos_consolidados_limpios.csv'
df_limpio.to_csv(nombre_salida, index=False)

print("\n--- RESUMEN FINAL ---")
print(f"✅ Limpieza terminada. Filas finales: {len(df_limpio)}")
print(f"📁 El archivo '{nombre_salida}' está listo para usar en Power BI o automatizaciones.")