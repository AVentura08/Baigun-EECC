import os
import pandas as pd
from supabase import create_client, Client

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("⚠️ Error: SUPABASE_URL y SUPABASE_KEY deben estar configuradas.")

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

def obtener_float(val):
    try:
        if pd.isna(val): return 0.0
        return float(val)
    except:
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

            dolares = obtener_float(row.get('dolares'))
            monto2 = obtener_float(row.get('monto2'))
            monto = obtener_float(row.get('monto'))
            cotiz = obtener_float(row.get('cotización', row.get('cotizacion')))

            # Seleccionar valor real en USD
            if dolares > 0:
                monto_usd = dolares
            elif monto2 > 0:
                monto_usd = monto2
            elif monto > 0 and cotiz > 0:
                monto_usd = round(monto / cotiz, 2)
            else:
                monto_usd = monto

            sector = str(row.get('sector', 'com'))

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
        print(f"✅ Ingresos procesados: {len(registros)} filas.")

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

            total_usd = obtener_float(row.get('ingreso monto total dolares'))
            monto_com = obtener_float(row.get('monto de comision'))
            cotiz = obtener_float(row.get('cotización', row.get('cotizacion')))
            moneda = str(row.get('moneda', '')).lower()

            if total_usd > 0:
                monto_usd = total_usd
            elif 'u$s' in moneda or 'dolar' in moneda or 'usd' in moneda:
                monto_usd = monto_com
            elif monto_com > 0 and cotiz > 0:
                monto_usd = round(monto_com / cotiz, 2)
            else:
                monto_usd = monto_com

            sector = str(row.get('sector', 'com'))

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(sector),
                "monto_usd": monto_usd,
                "concepto": str(row.get('ingreso', 'Comision')) if not pd.isna(row.get('ingreso')) else 'Comision'
            })
        except Exception:
            continue

    if registros:
        supabase.table("eerr_comisiones").delete().neq("id", 0).execute()
        supabase.table("eerr_comisiones").insert(registros).execute()
        print(f"✅ Comisiones procesadas: {len(registros)} filas.")

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

            monto = obtener_float(row.get('monto'))
            cotiz = obtener_float(row.get('cotización', row.get('cotizacion')))
            moneda = str(row.get('moneda', '')).lower()

            # Si es pesos, convierte por cotizacion
            if 'pes' in moneda or '$' in moneda:
                monto_usd = round(monto / cotiz, 2) if cotiz > 0 else monto
            else:
                monto_usd = monto

            sector = str(row.get('sector', 'com'))
            tipo_gasto = str(row.get('tipo de gasto', 'opex')).lower()
            tipo_rubro = 'capex' if 'capex' in tipo_gasto else 'opex'

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "tipo_rubro": tipo_rubro,
                "driver": str(row.get('categoría contable', '')) if not pd.isna(row.get('categoría contable')) else '',
                "concepto_analitico": str(row.get('nombre', '')) if not pd.isna(row.get('nombre')) else '',
                "monto_real_usd": monto_usd,
                "monto_presupuestado_usd": 0.0,
                "area_id": mapear_area(sector)
            })
        except Exception:
            continue

    if registros:
        supabase.table("eerr_gastos").delete().neq("id", 0).execute()
        supabase.table("eerr_gastos").insert(registros).execute()
        print(f"✅ Gastos procesados: {len(registros)} filas.")

if __name__ == "__main__":
    print("🚀 Iniciando procesamiento automático de EERR...")
    cargar_ingresos()
    cargar_comisiones()
    cargar_gastos()
    print("🎉 Proceso finalizado.")
