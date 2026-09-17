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
    print("📖 Leyendo: Ingresos.xlsx")
    df = pd.read_excel("Ingresos.xlsx")
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            if pd.isna(fecha): continue

            # Formula exacta: Monto2 (USD directos) + (Monto Pesos / Cotizacion)
            monto2_usd = float_val(row.get('monto2'))
            monto_ars = float_val(row.get('monto'))
            cotiz = float_val(row.get('cotización', row.get('cotizacion')))

            monto_ars_convertido = (monto_ars / cotiz) if (monto_ars > 0 and cotiz > 0) else 0.0
            total_usd = round(monto2_usd + monto_ars_convertido, 2)

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('sector', 'com'))),
                "monto_usd": total_usd,
                "concepto": str(row.get('tipo de operación', 'Ingreso')) if not pd.isna(row.get('tipo de operación')) else 'Ingreso'
            })
        except: continue

    if registros:
        supabase.table("eerr_ingresos").delete().neq("id", 0).execute()
        supabase.table("eerr_ingresos").insert(registros).execute()
        print(f"✅ Ingresos procesados: {len(registros)} filas.")

def cargar_comisiones():
    if not os.path.exists("Comisiones.xlsx"): return
    print("📖 Leyendo: Comisiones.xlsx")
    df = pd.read_excel("Comisiones.xlsx")
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            if pd.isna(fecha): continue

            m_com = float_val(row.get('monto de comision'))
            tot_dls = float_val(row.get('ingreso monto total dolares'))
            cotiz = float_val(row.get('cotización', row.get('cotizacion')))
            moneda = str(row.get('moneda', '')).lower()

            if 'u$s' in moneda or 'dolar' in moneda or 'usd' in moneda:
                monto_usd = m_com if m_com > 0 else tot_dls
            elif m_com > 0 and cotiz > 0:
                monto_usd = round(m_com / cotiz, 2)
            else:
                monto_usd = m_com

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
        print(f"✅ Comisiones procesadas: {len(registros)} filas.")

def cargar_gastos():
    if not os.path.exists("Gastos.xlsx"): return
    print("📖 Leyendo: Gastos.xlsx")
    df = pd.read_excel("Gastos.xlsx")
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.get('fecha'))
            if pd.isna(fecha): continue

            monto = float_val(row.get('monto'))
            cotiz = float_val(row.get('cotización', row.get('cotizacion')))
            moneda = str(row.get('moneda', '')).lower()

            if 'pes' in moneda or '$' in moneda:
                monto_usd = round(monto / cotiz, 2) if cotiz > 0 else monto
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
        print(f"✅ Gastos procesados: {len(registros)} filas.")

if __name__ == "__main__":
    print("🚀 Iniciando procesamiento automático de EERR...")
    cargar_ingresos()
    cargar_comisiones()
    cargar_gastos()
    print("🎉 Proceso finalizado.")
