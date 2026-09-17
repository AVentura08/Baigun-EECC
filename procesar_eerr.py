import os
import pandas as pd
from supabase import create_client

# Carga de credenciales desde las variables seguras de GitHub
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

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

# 1. INGRESOS
def cargar_ingresos(ruta_excel):
    if not os.path.exists(ruta_excel):
        print(f"⚠️ No se encontró el archivo: {ruta_excel}")
        return
    df = pd.read_excel(ruta_excel)
    df['Monto'] = df['Monto'].fillna(0)
    df['Monto.1'] = df['Monto.1'].fillna(0)
    df['Cotización'] = df['Cotización'].replace(0, pd.NA).fillna(1)
    
    df['Total_USD'] = df['Monto.1'] + (df['Monto'] / df['Cotización'])
    
    registros = []
    for _, row in df.iterrows():
        fecha = pd.to_datetime(row['Fecha'])
        registros.append({
            "fecha": fecha.strftime('%Y-%m-%d'),
            "mes": fecha.month,
            "anio": fecha.year,
            "area_id": obtener_area_id(row['Sector']),
            "monto_usd": round(float(row['Total_USD']), 2),
            "concepto": str(row['Nombre']) if pd.notna(row['Nombre']) else "Ingreso general"
        })
    supabase.table("eerr_ingresos").insert(registros).execute()
    print(f"✅ Ingresos procesados: {len(registros)} filas.")

# 2. COMISIONES
def cargar_comisiones(ruta_excel):
    if not os.path.exists(ruta_excel):
        print(f"⚠️ No se encontró el archivo: {ruta_excel}")
        return
    df = pd.read_excel(ruta_excel)
    df['Monto de Comision'] = df['Monto de Comision'].fillna(0)
    df['Cotización'] = df['Cotización'].replace(0, pd.NA).fillna(1)
    
    def calc_com_usd(row):
        moneda = str(row['Moneda']).strip()
        if 'U$S' in moneda or 'USD' in moneda:
            return row['Monto de Comision']
        else:
            return row['Monto de Comision'] / row['Cotización']

    df['Total_USD'] = df.apply(calc_com_usd, axis=1)
    
    registros = []
    for _, row in df.iterrows():
        fecha = pd.to_datetime(row['Fecha'])
        registros.append({
            "fecha": fecha.strftime('%Y-%m-%d'),
            "mes": fecha.month,
            "anio": fecha.year,
            "area_id": obtener_area_id(row['Sector']),
            "monto_usd": round(float(row['Total_USD']), 2),
            "concepto": str(row['Ingreso']) if pd.notna(row['Ingreso']) else "Comisión"
        })
    supabase.table("eerr_comisiones").insert(registros).execute()
    print(f"✅ Comisiones procesadas: {len(registros)} filas.")

# 3. GASTOS
def cargar_gastos(ruta_excel):
    if not os.path.exists(ruta_excel):
        print(f"⚠️ No se encontró el archivo: {ruta_excel}")
        return
    df = pd.read_excel(ruta_excel)
    df['Monto'] = df['Monto'].fillna(0)
    df['Cotización'] = df['Cotización'].replace(0, pd.NA).fillna(1)
    df['Total_USD'] = df['Monto'] / df['Cotización']
    
    registros = []
    for _, row in df.iterrows():
        fecha = pd.to_datetime(row['Fecha'])
        tipo_gasto = str(row['Tipo de gasto']).lower() if pd.notna(row['Tipo de gasto']) else ''
        tipo_rubro = 'capex' if 'capex' in tipo_gasto or 'inversion' in tipo_gasto else 'opex'
        
        registros.append({
            "fecha": fecha.strftime('%Y-%m-%d'),
            "mes": fecha.month,
            "anio": fecha.year,
            "tipo_rubro": tipo_rubro,
            "driver": "ofi",
            "concepto_analitico": str(row['Nombre']) if pd.notna(row['Nombre']) else "Gasto operativo",
            "monto_real_usd": round(float(row['Total_USD']), 2),
            "monto_presupuestado_usd": 0,
            "area_id": obtener_area_id(row['Sector'])
        })
    supabase.table("eerr_gastos").insert(registros).execute()
    print(f"✅ Gastos procesados: {len(registros)} filas.")

if __name__ == "__main__":
    cargar_ingresos("Ingresos.xlsx")
    cargar_comisiones("Comisiones.xlsx")
    cargar_gastos("Gastos.xlsx")
