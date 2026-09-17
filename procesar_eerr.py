import os
import pandas as pd
from supabase import create_client, Client

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("⚠️ Error: SUPABASE_URL y SUPABASE_KEY deben estar configuradas.")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def mapear_area(sector_str, tipo_str=""):
    sec = str(sector_str).lower().strip() if not pd.isna(sector_str) else ""
    tipo = str(tipo_str).lower().strip() if not pd.isna(tipo_str) else ""

    # Regla específica para Residencial
    if 'res' in sec or 'usa' in sec:
        if 'emp' in tipo or 'proy' in tipo or 'pozo' in tipo:
            return 'emp'
        return 'usa'
    
    if 'com' in sec: return 'com'
    if 'emp' in sec: return 'emp'
    if 'ter' in sec: return 'ter'
    if 'esp' in sec: return 'esp'
    
    return 'com'

def float_val(v):
    try:
        if pd.isna(v): return 0.0
        if isinstance(v, (int, float)): return float(v)
        s = str(v).strip().replace('$', '').replace(' ', '')
        if not s or s in ['#¡VALOR!', 'nan', 'none', '-']: return 0.0
        if '.' in s and ',' in s: s = s.replace('.', '').replace(',', '.')
        elif ',' in s: s = s.replace(',', '.')
        return float(s)
    except:
        return 0.0

def cargar_ingresos():
    if not os.path.exists("Ingresos.xlsx"): return
    print("📖 Leyendo: Ingresos.xlsx")
    df = pd.read_excel("Ingresos.xlsx")

    registros = []
    for idx, row in df.iterrows():
        try:
            fecha_val = row.get('Fecha')
            if pd.isna(fecha_val): continue
            fecha = pd.to_datetime(fecha_val)

            monto_pesos = float_val(row.get('Monto'))
            monto_dolares = float_val(row.get('Monto.1'))
            cotizacion = float_val(row.get('Cotización'))

            monto_ars_convertido = (monto_pesos / cotizacion) if (monto_pesos > 0 and cotizacion > 0) else 0.0
            total_usd = round(monto_dolares + monto_ars_convertido, 2)

            sector = row.get('Sector', '')
            tipo_ingreso = row.get('Tipo de Ingreso', row.get('Tipo de Operación', ''))

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(sector, tipo_ingreso),
                "monto_usd": total_usd,
                "concepto": str(tipo_ingreso) if not pd.isna(tipo_ingreso) else 'Ingreso'
            })
        except Exception:
            continue

    if registros:
        supabase.table("eerr_ingresos").delete().neq("id", 0).execute()
        supabase.table("eerr_ingresos").insert(registros).execute()
        print("✅ Ingresos cargados correctamente.")

def cargar_comisiones():
    if not os.path.exists("Comisiones.xlsx"): return
    df = pd.read_excel("Comisiones.xlsx")

    registros = []
    for _, row in df.iterrows():
        try:
            fecha_val = row.get('Fecha')
            if pd.isna(fecha_val): continue
            fecha = pd.to_datetime(fecha_val)

            m_com = float_val(row.get('Monto de Comision', row.get('Monto de Comisión')))
            cotiz = float_val(row.get('Cotización', row.get('Cotizacion')))
            moneda = str(row.get('Moneda', '')).lower()

            monto_usd = m_com if ('u$s' in moneda or 'dolar' in moneda or 'usd' in moneda) else (round(m_com / cotiz, 2) if cotiz > 0 else m_com)

            sector = row.get('Sector', '')
            tipo = row.get('Tipo', row.get('Categoría Contable', ''))

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(sector, tipo),
                "monto_usd": monto_usd,
                "concepto": 'Comisión'
            })
        except: continue

    if registros:
        supabase.table("eerr_comisiones").delete().neq("id", 0).execute()
        supabase.table("eerr_comisiones").insert(registros).execute()

def cargar_gastos():
    if not os.path.exists("Gastos.xlsx"): return
    df = pd.read_excel("Gastos.xlsx")

    registros = []
    for _, row in df.iterrows():
        try:
            fecha_val = row.get('Fecha')
            if pd.isna(fecha_val): continue
            fecha = pd.to_datetime(fecha_val)

            monto = float_val(row.get('Monto'))
            cotiz = float_val(row.get('Cotización', row.get('Cotizacion')))
            moneda = str(row.get('Moneda', '')).lower()

            monto_usd = round(monto / cotiz, 2) if ('pes' in moneda or '$' in moneda) and cotiz > 0 else monto

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "tipo_rubro": 'opex',
                "driver": str(row.get('Categoría Contable', '')),
                "concepto_analitico": str(row.get('Nombre', '')),
                "monto_real_usd": monto_usd,
                "monto_presupuestado_usd": 0.0,
                "area_id": mapear_area(row.get('Sector', ''))
            })
        except: continue

    if registros:
        supabase.table("eerr_gastos").delete().neq("id", 0).execute()
        supabase.table("eerr_gastos").insert(registros).execute()

if __name__ == "__main__":
    cargar_ingresos()
    cargar_comisiones()
    cargar_gastos()
