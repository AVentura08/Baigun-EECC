import os
import pandas as pd
from supabase import create_client, Client

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

def extraer_monto_usd(row):
    """ Busca columnas en USD o convierte si encuentra pesos y cotizacion """
    # 1. Buscar columna explicita en USD
    for col in ['monto_usd', 'monto usd', 'total usd', 'importe usd', 'usd']:
        if col in row and not pd.isna(row[col]) and float(row[col]) > 0:
            return float(row[col])
            
    # 2. Buscar monto general
    monto = 0.0
    for col in ['monto', 'total', 'importe', 'monto_real', 'monto_real_usd']:
        if col in row and not pd.isna(row[col]):
            monto = float(row[col])
            break

    # 3. Buscar tipo de cambio / cotizacion si el monto parece en pesos
    tc = 1.0
    for col in ['tc', 'cotizacion', 'tipo_de_cambio', 'tipo de cambio', 'cambio']:
        if col in row and not pd.isna(row[col]) and float(row[col]) > 0:
            tc = float(row[col])
            break

    if tc > 1.0:
        return round(monto / tc, 2)
    
    return round(monto, 2)

def cargar_ingresos():
    archivo = "Ingresos.xlsx"
    if not os.path.exists(archivo): return

    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            monto_usd = extraer_monto_usd(row)

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('area', 'com'))),
                "monto_usd": monto_usd,
                "concepto": str(row.get('concepto', '')) if not pd.isna(row.get('concepto')) else ''
            })
        except Exception:
            continue

    if registros:
        supabase.table("eerr_ingresos").delete().neq("id", 0).execute()
        supabase.table("eerr_ingresos").insert(registros).execute()
        print(f"✅ Ingresos cargados con éxito: {len(registros)} filas.")

def cargar_comisiones():
    archivo = "Comisiones.xlsx"
    if not os.path.exists(archivo): return

    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            monto_usd = extraer_monto_usd(row)

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('area', 'com'))),
                "monto_usd": monto_usd,
                "concepto": str(row.get('concepto', '')) if not pd.isna(row.get('concepto')) else ''
            })
        except Exception:
            continue

    if registros:
        supabase.table("eerr_comisiones").delete().neq("id", 0).execute()
        supabase.table("eerr_comisiones").insert(registros).execute()
        print(f"✅ Comisiones cargadas con éxito: {len(registros)} filas.")

def cargar_gastos():
    archivo = "Gastos.xlsx"
    if not os.path.exists(archivo): return

    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            monto_usd = extraer_monto_usd(row)

            tipo = str(row.get('tipo_rubro', 'opex')).lower().strip()
            if tipo not in ['opex', 'capex']: tipo = 'opex'

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "tipo_rubro": tipo,
                "driver": str(row.get('driver', '')) if not pd.isna(row.get('driver')) else '',
                "concepto_analitico": str(row.get('concepto_analitico', '')) if not pd.isna(row.get('concepto_analitico')) else '',
                "monto_real_usd": monto_usd,
                "monto_presupuestado_usd": 0.0,
                "area_id": mapear_area(str(row.get('area', 'com')))
            })
        except Exception:
            continue

    if registros:
        supabase.table("eerr_gastos").delete().neq("id", 0).execute()
        supabase.table("eerr_gastos").insert(registros).execute()
        print(f"✅ Gastos cargados con éxito: {len(registros)} filas.")

if __name__ == "__main__":
    print("🚀 Iniciando procesamiento automático de EERR...")
    cargar_ingresos()
    cargar_comisiones()
    cargar_gastos()
    print("🎉 Proceso finalizado.")
