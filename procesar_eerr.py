import os
import pandas as pd
from supabase import create_client, Client

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("⚠️ Error: SUPABASE_URL y SUPABASE_KEY deben estar configuradas.")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def mapear_area(area_str):
    if not isinstance(area_str, str): return 'com'
    val = area_str.lower().strip()
    if 'com' in val: return 'com'
    if 'usa' in val or 'res' in val: return 'usa'
    if 'emp' in val: return 'emp'
    if 'ter' in val: return 'ter'
    if 'esp' in val: return 'esp'
    return 'com'

def float_val(v):
    try:
        if pd.isna(v): return 0.0
        return float(v)
    except:
        return 0.0

def cargar_ingresos():
    if not os.path.exists("Ingresos.xlsx"): return
    df = pd.read_excel("Ingresos.xlsx")
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            if pd.isna(fecha): continue

            # Prioridad estricta para Ingresos
            m2 = float_val(row.get('monto2'))
            dls = float_val(row.get('dolares'))
            monto = float_val(row.get('monto'))
            cot = float_val(row.get('cotización', row.get('cotizacion')))

            if m2 > 0: monto_usd = m2
            elif dls > 0: monto_usd = dls
            elif monto > 0 and cot > 0: monto_usd = round(monto / cot, 2)
            else: monto_usd = monto

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('sector', 'com'))),
                "monto_usd": monto_usd,
                "concepto": str(row.get('tipo de operación', 'Ingreso')) if not pd.isna(row.get('tipo de operación')) else 'Ingreso'
            })
        except: continue

    if registros:
        supabase.table("eerr_ingresos").delete().neq("id", 0).execute()
        supabase.table("eerr_ingresos").insert(registros).execute()

def cargar_comisiones():
    if not os.path.exists("Comisiones.xlsx"): return
    df = pd.read_excel("Comisiones.xlsx")
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            if pd.isna(fecha): continue

            # Prioridad estricta para Comisiones
            m_com = float_val(row.get('monto de comision'))
            tot_dls = float_val(row.get('ingreso monto total dolares'))
            cot = float_val(row.get('cotización', row.get('cotizacion')))
            mon = str(row.get('moneda', '')).lower()

            if m_com > 0:
                if 'u$s' in mon or 'dolar' in mon or 'usd' in mon:
                    monto_usd = m_com
                elif cot > 0:
                    monto_usd = round(m_com / cot, 2)
                else:
                    monto_usd = m_com
            elif tot_dls > 0:
                monto_usd = tot_dls
            else:
                monto_usd = 0.0

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('sector', 'com'))),
                "monto_usd": monto_usd,
                "concepto": str(row.get('ingreso', 'Comisión')) if not pd.isna(row.get('ingreso')) else 'Comisión'
            })
        except: continue

    if registros:
        supabase.table("eerr_comisiones").delete().neq("id", 0).execute()
        supabase.table("eerr_comisiones").insert(registros).execute()

def cargar_gastos():
    if not os.path.exists("Gastos.xlsx"): return
    df = pd.read_excel("Gastos.xlsx")
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            if pd.isna(fecha): continue

            monto = float_val(row.get('monto'))
            cot = float_val(row.get('cotización', row.get('cotizacion')))
            mon = str(row.get('moneda', '')).lower()

            if 'pes' in mon or '$' in mon:
                monto_usd = round(monto / cot, 2) if cot > 0 else monto
            else:
                monto_usd = monto

            t_gasto = str(row.get('tipo de gasto', 'opex')).lower()

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "tipo_rubro": 'capex' if 'capex' in t_gasto else 'opex',
                "driver": str(row.get('categoría contable', '')) if not pd.isna(row.get('categoría contable')) else '',
                "concepto_analitico": str(row.get('nombre', '')) if not pd.isna(row.get('nombre')) else '',
                "monto_real_usd": monto_usd,
                "monto_presupuestado_usd": 0.0,
                "area_id": mapear_area(str(row.get('sector', 'com')))
            })
        except: continue

    if registros:
        supabase.table("eerr_gastos").delete().neq("id", 0).execute()
        supabase.table("eerr_gastos").insert(registros).execute()

if __name__ == "__main__":
    cargar_ingresos()
    cargar_comisiones()
    cargar_gastos()
