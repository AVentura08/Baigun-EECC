import os
import glob
import pandas as pd
from supabase import create_client

# Credenciales exactas de tu proyecto Supabase
SUPABASE_URL = "https://xvuyzjwnixbvavbvmzct.supabase.co"
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "sb_publishable_LsnpSgrmFUNC3tMba2YF0g_zb_lZE2k")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

MAPEO_AREAS = {
    'Comercial': 'com',
    'Residencial': 'usa',
    'Emprendimientos': 'emp',
    'Terrenos': 'ter',
    'Especiales': 'esp'
}

def obtener_area_id(sector_texto):
    if pd.isna(sector_texto):
        return 'com'
    sector_str = str(sector_texto).strip()
    return MAPEO_AREAS.get(sector_str, 'com')

def buscar_archivo(patron):
    coincidencias = glob.glob(patron)
    return coincidencias[0] if coincidencias else None

# 1. INGRESOS
def cargar_ingresos():
    archivo = buscar_archivo("Ingresos*.xlsx")
    if not archivo:
        print("⚠️ No se encontró archivo de Ingresos.")
        return
    
    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df['Monto'] = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)
    df['Monto.1'] = pd.to_numeric(df['Monto.1'], errors='coerce').fillna(0)
    df['Cotización'] = pd.to_numeric(df['Cotización'], errors='coerce').replace(0, pd.NA).fillna(1)
    
    df['Total_USD'] = df['Monto.1'] + (df['Monto'] / df['Cotización'])
    
    registros = []
    for _, row in df.iterrows():
        if pd.isna(row.get('Fecha')):
            continue
        fecha = pd.to_datetime(row['Fecha'])
        registros.append({
            "fecha": fecha.strftime('%Y-%m-%d'),
            "mes": fecha.month,
            "anio": fecha.year,
            "area_id": obtener_area_id(row.get('Sector')),
            "monto_usd": round(float(row['Total_USD']), 2),
            "concepto": str(row['Nombre']) if pd.notna(row.get('Nombre')) else "Ingreso general"
        })
    if registros:
        supabase.table("eerr_ingresos").insert(registros).execute()
        print(f"✅ Ingresos procesados: {len(registros)} filas.")

# 2. COMISIONES
def cargar_comisiones():
    archivo = buscar_archivo("Comisiones*.xlsx")
    if not archivo:
        print("⚠️ No se encontró archivo de Comisiones.")
        return
    
    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df['Monto de Comision'] = pd.to_numeric(df['Monto de Comision'], errors='coerce').fillna(0)
    df['Cotización'] = pd.to_numeric(df['Cotización'], errors='coerce').replace(0, pd.NA).fillna(1)
    
    def calc_com_usd(row):
        moneda = str(row.get('Moneda', '')).strip()
        monto = row['Monto de Comision']
        cotiz = row['Cotización']
        if 'U$S' in moneda or 'USD' in moneda:
            return monto
        else:
            return monto / cotiz if cotiz else 0

    df['Total_USD'] = df.apply(calc_com_usd, axis=1)
    
    registros = []
    for _, row in df.iterrows():
        if pd.isna(row.get('Fecha')):
            continue
        fecha = pd.to_datetime(row['Fecha'])
        registros.append({
            "fecha": fecha.strftime('%Y-%m-%d'),
            "mes": fecha.month,
            "anio": fecha.year,
            "area_id": obtener_area_id(row.get('Sector')),
            "monto_usd": round(float(row['Total_USD']), 2),
            "concepto": str(row['Ingreso']) if pd.notna(row.get('Ingreso')) else "Comisión"
        })
    if registros:
        supabase.table("eerr_comisiones").insert(registros).execute()
        print(f"✅ Comisiones procesadas: {len(registros)} filas.")

# 3. GASTOS
def cargar_gastos():
    archivo = buscar_archivo("Gastos*.xlsx")
    if not archivo:
        print("⚠️ No se encontró archivo de Gastos.")
        return
    
    print(f"📖 Leyendo: {archivo}")
    df = pd.read_excel(archivo)
    df['Monto'] = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)
    df['Cotización'] = pd.to_numeric(df['Cotización'], errors='coerce').replace(0, pd.NA).fillna(1)
    df['Total_USD'] = df['Monto'] / df['Cotización']
    
    registros = []
    for _, row in df.iterrows():
        if pd.isna(row.get('Fecha')):
            continue
        fecha = pd.to_datetime(row['Fecha'])
        tipo_gasto = str(row.get('Tipo de gasto', '')).lower()
        tipo_rubro = 'capex' if 'capex' in tipo_gasto or 'inversion' in tipo_gasto else 'opex'
        
        registros.append({
            "fecha": fecha.strftime('%Y-%m-%d'),
            "mes": fecha.month,
            "anio": fecha.year,
            "tipo_rubro": tipo_rubro,
            "driver": "ofi",
            "concepto_analitico": str(row['Nombre']) if pd.notna(row.get('Nombre')) else "Gasto operativo",
            "monto_real_usd": round(float(row['Total_USD']), 2),
            "monto_presupuestado_usd": 0,
            "area_id": obtener_area_id(row.get('Sector'))
        })
    if registros:
        supabase.table("eerr_gastos").insert(registros).execute()
        print(f"✅ Gastos procesados: {len(registros)} filas.")

if __name__ == "__main__":
    cargar_ingresos()
    cargar_comisiones()
    cargar_gastos()
    print("🚀 ¡Proceso finalizado correctamente!")
