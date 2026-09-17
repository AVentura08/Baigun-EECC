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

def calcular_monto_usd_exacto(row):
    """ Mapeo preciso leyendo las columnas reales del Excel """
    monto2 = row.get('monto2', 0)
    dolares = row.get('dolares', 0)
    monto = row.get('monto', 0)
    cotiz = row.get('cotización', row.get('cotizacion', 1))

    # Si hay valor en 'Monto2' o 'Dolares', ese es el monto en USD
    if not pd.isna(monto2) and float(monto2) > 0:
        return float(monto2)
    if not pd.isna(dolares) and float(dolares) > 0:
        return float(dolares)

    # Si el valor esta en pesos, divide por la Cotizacion de la fila
    if not pd.isna(monto) and float(monto) > 0:
        tc = float(cotiz) if (not pd.isna(cotiz) and float(cotiz) > 0) else 1.0
        return round(float(monto) / tc, 2)

    return 0.0

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
            if pd.isna(fecha): continue

            monto_usd = calcular_monto_usd_exacto(row)
            sector = str(row.get('sector', row.get('area', 'com')))

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(sector),
                "monto_usd": monto_usd,
                "concepto": str(row.get('tipo de operación', row.get('concepto', ''))) if not pd.isna(row.get('tipo de operación')) else ''
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
            if pd.isna(fecha): continue

            monto_usd = calcular_monto_usd_exacto(row)
            sector = str(row.get('sector', row.get('area', 'com')))

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(sector),
                "monto_usd": monto_usd,
                "concepto": str(row.get('tipo de operación', row.get('concepto', ''))) if not pd.isna(row.get('tipo de operación')) else ''
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
            if pd.isna(fecha): continue

            monto_usd = calcular_monto_usd_exacto(row)
            sector = str(row.get('sector', row.get('area', 'com')))

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
                "area_id": mapear_area(sector)
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
    
