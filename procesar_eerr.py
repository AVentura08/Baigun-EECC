import os
import pandas as pd
from supabase import create_client, Client

# Configuración de credenciales Supabase desde Secrets / Variables de entorno
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("⚠️ Error: Las variables SUPABASE_URL y SUPABASE_KEY deben estar configuradas.")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def mapear_area(area_str):
    if not isinstance(area_str, str):
        return 'com'
    val = area_str.lower().strip()
    if 'com' in val: return 'com'
    if 'usa' in val or 'res' in val: return 'usa'
    if 'emp' in val: return 'emp'
    if 'ter' in val: return 'ter'
    if 'esp' in val: return 'esp'
    return 'com'

def cargar_ingresos():
    archivo = "Ingresos.xlsx"
    if not os.path.exists(archivo):
        print(f"⚠️ {archivo} no encontrado. Se omite.")
        return

    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            monto = float(row.get('monto_usd', row.get('monto', 0)))
            if pd.isna(monto): monto = 0.0

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('area', 'com'))),
                "monto_usd": monto,
                "concepto": str(row.get('concepto', '')) if not pd.isna(row.get('concepto')) else ''
            })
        except Exception as e:
            continue

    if registros:
        # Limpieza previa para evitar duplicados
        supabase.table("eerr_ingresos").delete().neq("id", 0).execute()
        supabase.table("eerr_ingresos").insert(registros).execute()
        print(f"✅ Ingresos cargados con éxito: {len(registros)} filas.")

def cargar_comisiones():
    archivo = "Comisiones.xlsx"
    if not os.path.exists(archivo):
        print(f"⚠️ {archivo} no encontrado. Se omite.")
        return

    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            monto = float(row.get('monto_usd', row.get('monto', 0)))
            if pd.isna(monto): monto = 0.0

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('area', 'com'))),
                "monto_usd": monto,
                "concepto": str(row.get('concepto', '')) if not pd.isna(row.get('concepto')) else ''
            })
        except Exception as e:
            continue

    if registros:
        # Limpieza previa para evitar duplicados
        supabase.table("eerr_comisiones").delete().neq("id", 0).execute()
        supabase.table("eerr_comisiones").insert(registros).execute()
        print(f"✅ Comisiones cargadas con éxito: {len(registros)} filas.")

def cargar_gastos():
    archivo = "Gastos.xlsx"
    if not os.path.exists(archivo):
        print(f"⚠️ {archivo} no encontrado. Se omite.")
        return

    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            monto_real = float(row.get('monto_real_usd', row.get('monto_real', row.get('monto', 0))))
            monto_presu = float(row.get('monto_presupuestado_usd', row.get('presupuesto', 0)))
            
            if pd.isna(monto_real): monto_real = 0.0
            if pd.isna(monto_presu): monto_presu = 0.0

            tipo = str(row.get('tipo_rubro', 'opex')).lower().strip()
            if tipo not in ['opex', 'capex']: tipo = 'opex'

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "tipo_rubro": tipo,
                "driver": str(row.get('driver', '')) if not pd.isna(row.get('driver')) else '',
                "concepto_analitico": str(row.get('concepto_analitico', '')) if not pd.isna(row.get('concepto_analitico')) else '',
                "monto_real_usd": monto_real,
                "monto_presupuestado_usd": monto_presu,
                "area_id": mapear_area(str(row.get('area', 'com')))
            })
        except Exception as e:
            continue

    if registros:
        # Limpieza previa para evitar duplicados
        supabase.table("eerr_gastos").delete().neq("id", 0).execute()
        supabase.table("eerr_gastos").insert(registros).execute()
        print(f"✅ Gastos cargados con éxito: {len(registros)} filas.")

if __name__ == "__main__":
    print("🚀 Iniciando procesamiento automático de EERR...")
    cargar_ingresos()
    cargar_comisiones()
    cargar_gastos()
    print("🎉 Proceso finalizado.")
