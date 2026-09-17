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
    
    # Imprimir columnas reales detectadas
    print(f"📋 COLUMNAS DETECTADAS EN EXCEL: {list(df.columns)}")
    
    # Limpieza estricta de nombres de columnas
    df.columns = [str(c).strip().lower().replace(' ', '') for c in df.columns]

    registros = []
    suma_total_debug = 0.0

    for idx, row in df.iterrows():
        try:
            # Buscar cualquier columna que contenga fecha
            fecha_val = None
            for col in row.index:
                if 'fecha' in col:
                    fecha_val = row[col]
                    break

            if pd.isna(fecha_val): continue
            fecha = pd.to_datetime(fecha_val)

            # Extraer valores buscando coincidencias parciales en los nombres de columna
            monto2_usd = 0.0
            monto_ars = 0.0
            cotiz = 1.0

            for col in row.index:
                if col in ['monto2', 'monto_2', 'totalusd', 'dolares']:
                    val = float_val(row[col])
                    if val > 0: monto2_usd = val
                elif col in ['monto', 'pesos', 'montoars']:
                    val = float_val(row[col])
                    if val > 0: monto_ars = val
                elif col in ['cotización', 'cotizacion', 'tc', 'tipodecambio']:
                    val = float_val(row[col])
                    if val > 0: cotiz = val

            monto_ars_convertido = (monto_ars / cotiz) if (monto_ars > 0 and cotiz > 0) else 0.0
            total_usd = round(monto2_usd + monto_ars_convertido, 2)
            suma_total_debug += total_usd

            # Detectar sector / area
            sector_val = 'com'
            for col in row.index:
                if 'sector' in col or 'area' in col:
                    sector_val = str(row[col])
                    break

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(sector_val),
                "monto_usd": total_usd,
                "concepto": "Ingreso"
            })
        except Exception as e:
            continue

    print(f"📊 DEBUG REVISADO -> Filas: {len(registros)} | Suma calculada: ${suma_total_debug:,.2f} USD")

    if registros:
        supabase.table("eerr_ingresos").delete().neq("id", 0).execute()
        supabase.table("eerr_ingresos").insert(registros).execute()
        print("✅ Ingresos cargados a Supabase.")

def cargar_comisiones():
    if not os.path.exists("Comisiones.xlsx"): return
    df = pd.read_excel("Comisiones.xlsx")
    df.columns = [str(c).strip().lower().replace(' ', '') for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha_val = None
            for col in row.index:
                if 'fecha' in col:
                    fecha_val = row[col]
                    break
            if pd.isna(fecha_val): continue
            fecha = pd.to_datetime(fecha_val)

            m_com = 0.0
            cotiz = 1.0
            for col in row.index:
                if 'montodecomision' in col or 'comision' in col:
                    m_com = float_val(row[col])
                elif 'cotiz' in col:
                    cotiz = float_val(row[col])

            monto_usd = round(m_com / cotiz, 2) if (m_com > 0 and cotiz > 0) else m_com

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(str(row.get('sector', 'com'))),
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
    df.columns = [str(c).strip().lower().replace(' ', '') for c in df.columns]

    registros = []
    for _, row in df.iterrows():
        try:
            fecha_val = None
            for col in row.index:
                if 'fecha' in col:
                    fecha_val = row[col]
                    break
            if pd.isna(fecha_val): continue
            fecha = pd.to_datetime(fecha_val)

            monto = 0.0
            cotiz = 1.0
            for col in row.index:
                if 'monto' in col: monto = float_val(row[col])
                elif 'cotiz' in col: cotiz = float_val(row[col])

            moneda = str(row.get('moneda', '')).lower()
            monto_usd = round(monto / cotiz, 2) if ('pes' in moneda or '$' in moneda) and cotiz > 0 else monto

            registros.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "tipo_rubro": 'opex',
                "driver": str(row.get('categoríacontable', '')),
                "concepto_analitico": str(row.get('nombre', '')),
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
