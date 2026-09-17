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
    if 'emp' in sec or 'emp' in tipo: return 'emp'
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

def cargar_datos_completos():
    if not os.path.exists("Ingresos.xlsx"):
        print("❌ No se encontró Ingresos.xlsx")
        return

    print("📖 Leyendo: Ingresos.xlsx")
    df_ingresos = pd.read_excel("Ingresos.xlsx")

    # Crear un diccionario indexado por el ID '#' para buscar el 'Tipo de Ingreso' instantáneamente
    mapa_ingresos_tipo = {}
    registros_ingresos = []

    for idx, row in df_ingresos.iterrows():
        try:
            fecha_val = row.get('Fecha')
            if pd.isna(fecha_val): continue
            fecha = pd.to_datetime(fecha_val)

            id_operacion = row.get('#')
            sector = row.get('Sector', '')
            tipo_ingreso = row.get('Tipo de Ingreso', row.get('Tipo de Operación', ''))

            # Guardar en el mapa de cruce
            if not pd.isna(id_operacion):
                mapa_ingresos_tipo[str(id_operacion).strip()] = {
                    "sector": sector,
                    "tipo": tipo_ingreso
                }

            monto_pesos = float_val(row.get('Monto'))
            monto_dolares = float_val(row.get('Monto.1'))
            cotizacion = float_val(row.get('Cotización'))

            monto_ars_convertido = (monto_pesos / cotizacion) if (monto_pesos > 0 and cotizacion > 0) else 0.0
            total_usd = round(monto_dolares + monto_ars_convertido, 2)

            registros_ingresos.append({
                "fecha": fecha.strftime('%Y-%m-%d'),
                "mes": int(fecha.month),
                "anio": int(fecha.year),
                "area_id": mapear_area(sector, tipo_ingreso),
                "monto_usd": total_usd,
                "concepto": str(tipo_ingreso) if not pd.isna(tipo_ingreso) else 'Ingreso'
            })
        except Exception:
            continue

    if registros_ingresos:
        supabase.table("eerr_ingresos").delete().neq("id", 0).execute()
        supabase.table("eerr_ingresos").insert(registros_ingresos).execute()
        print(f"✅ Ingresos procesados: {len(registros_ingresos)} filas.")

    # --- PROCESAR COMISIONES CON CRUCE POR ID (#) ---
    if os.path.exists("Comisiones.xlsx"):
        print("📖 Leyendo: Comisiones.xlsx")
        df_comisiones = pd.read_excel("Comisiones.xlsx")

        registros_comisiones = []
        for _, row in df_comisiones.iterrows():
            try:
                fecha_val = row.get('Fecha')
                if pd.isna(fecha_val): continue
                fecha = pd.to_datetime(fecha_val)

                # Intentar cruzar ID (# o Nº Ingreso) con la tabla de Ingresos
                id_comision = row.get('#', row.get('Ingreso', row.get('Nº Ingreso')))
                id_str = str(id_comision).strip() if not pd.isna(id_comision) else ""

                sector = row.get('Sector', '')
                tipo = row.get('Tipo', row.get('Categoría Contable', ''))

                # Si encontramos el ID en Ingresos, usamos el Tipo de Ingreso real
                if id_str in mapa_ingresos_tipo:
                    info_ingreso = mapa_ingresos_tipo[id_str]
                    sector = info_ingreso["sector"]
                    tipo = info_ingreso["tipo"]

                m_com = float_val(row.get('Monto de Comision', row.get('Monto de Comisión')))
                cotiz = float_val(row.get('Cotización', row.get('Cotizacion')))
                moneda = str(row.get('Moneda', '')).lower()

                monto_usd = m_com if ('u$s' in moneda or 'dolar' in moneda or 'usd' in moneda) else (round(m_com / cotiz, 2) if cotiz > 0 else m_com)

                registros_comisiones.append({
                    "fecha": fecha.strftime('%Y-%m-%d'),
                    "mes": int(fecha.month),
                    "anio": int(fecha.year),
                    "area_id": mapear_area(sector, tipo),
                    "monto_usd": monto_usd,
                    "concepto": 'Comisión'
                })
            except Exception: continue

        if registros_comisiones:
            supabase.table("eerr_comisiones").delete().neq("id", 0).execute()
            supabase.table("eerr_comisiones").insert(registros_comisiones).execute()
            print(f"✅ Comisiones procesadas con cruce por ID #: {len(registros_comisiones)} filas.")

    # --- PROCESAR GASTOS ---
    if os.path.exists("Gastos.xlsx"):
        print("📖 Leyendo: Gastos.xlsx")
        df_gastos = pd.read_excel("Gastos.xlsx")

        registros_gastos = []
        for _, row in df_gastos.iterrows():
            try:
                fecha_val = row.get('Fecha')
                if pd.isna(fecha_val): continue
                fecha = pd.to_datetime(fecha_val)

                monto = float_val(row.get('Monto'))
                cotiz = float_val(row.get('Cotización', row.get('Cotizacion')))
                moneda = str(row.get('Moneda', '')).lower()

                monto_usd = round(monto / cotiz, 2) if ('pes' in moneda or '$' in moneda) and cotiz > 0 else monto

                registros_gastos.append({
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
            except Exception: continue

        if registros_gastos:
            supabase.table("eerr_gastos").delete().neq("id", 0).execute()
            supabase.table("eerr_gastos").insert(registros_gastos).execute()
            print(f"✅ Gastos procesados: {len(registros_gastos)} filas.")

if __name__ == "__main__":
    cargar_datos_completos()
