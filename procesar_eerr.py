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
    """ Convierte correctamente números en formato latino (17.550,00 -> 17550.00) """
    try:
        if pd.isna(v): return 0.0
        if isinstance(v, (int, float)): return float(v)
        
        s = str(v).strip().replace('$', '').replace(' ', '')
        if not s or s == '#¡VALOR!' or s == 'nan': return 0.0

        # Si tiene punto y coma (ej: 17.550,00)
        if '.' in s and ',' in s:
            s = s.replace('.', '').replace(',', '.')
        # Si solo tiene coma (ej: 17550,00)
        elif ',' in s:
            s = s.replace(',', '.')
            
        return float(s)
    except:
        return 0.0

def cargar_ingresos():
    if not os.path.exists("Ingresos.xlsx"): return

    print("📖 Leyendo: Ingresos.xlsx")
    df = pd.read_excel("Ingresos.xlsx")
    df.columns = [str(c).strip().lower() for c in df.columns]

    registros = []
    suma_total_debug = 0.0

    for idx, row in df.iterrows():
        try:
            fecha_val = row.get('fecha')
            if pd.isna(fecha_val): continue

            fecha = pd.to_datetime(fecha_val)

            monto2_usd = float_val(row.get('monto2'))
            monto_ars = float_val(row.get('monto'))
            cotiz = float_val(row.get('cotización', row.get('cotizacion')))

            # Si hay monto en pesos y cotizacion, convierte a USD
            monto_ars_convertido = (monto_ars / cotiz) if (monto_ars > 0 and cotiz > 0) else 0.0
            
            # Suma de la fila
            total_usd = round(monto2_usd + monto_ars_convertido, 2)
            suma_total_debug += total_usd

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('sector', 'com'))),
                "monto_usd": total_usd,
                "concepto": str(row.get('tipo de operación', 'Ingreso')) if not pd.isna(row.get('tipo de operación')) else 'Ingreso'
            })
        except Exception as e:
            continue

    print(f"📊 DEBUG INGRESOS -> Filas procesadas: {len(registros)} | Suma calculada: ${suma_total_debug:,.2f} USD")

    if registros:
        supabase.table("eerr_ingresos").delete().neq("id", 0).execute()
        supabase.table("eerr_ingresos").insert(registros).execute()
        print("✅ Ingresos subidos correctamente a Supabase.")

def cargar_comisiones():
    if not os.path.exists("Comisiones.xlsx"): return
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

if __name__ == "__main__":
    cargar_ingresos()
    cargar_comisiones()
    cargar_gastos()
