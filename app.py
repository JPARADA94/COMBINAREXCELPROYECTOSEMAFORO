import io
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Combinador de Excel", page_icon="📊", layout="wide")

# ============================================================
# COLUMNAS ESPERADAS
# ============================================================
COLUMNAS_ESPERADAS = [
    "NOMBRE_CLIENTE",
    "NOMBRE_OPERACION",
    "N_MUESTRA",
    "CORRELATIVO",
    "FECHA_MUESTREO",
    "FECHA_INGRESO",
    "FECHA_RECEPCION",
    "FECHA_INFORME",
    "EDAD_COMPONENTE",
    "UNIDAD_EDAD_COMPONENTE",
    "EDAD_PRODUCTO",
    "UNIDAD_EDAD_PRODUCTO",
    "CANTIDAD_ADICIONADA",
    "UNIDAD_CANTIDAD_ADICIONADA",
    "PRODUCTO",
    "TIPO_PRODUCTO",
    "EQUIPO",
    "TIPO_EQUIPO",
    "MARCA_EQUIPO",
    "MODELO_EQUIPO",
    "COMPONENTE",
    "MARCA_COMPONENTE",
    "MODELO_COMPONENTE",
    "DESCRIPTOR_COMPONENTE",
    "ESTADO",
    "NIVEL_DE_SERVICIO",
    "ÍNDICE PQ (PQI) - 3",
    "PLATA (AG) - 19",
    "ALUMINIO (AL) - 20",
    "CROMO (CR) - 24",
    "COBRE (CU) - 25",
    "HIERRO (FE) - 26",
    "TITANIO (TI) - 38",
    "PLOMO (PB) - 35",
    "NÍQUEL (NI) - 32",
    "MOLIBDENO (MO) - 30",
    "SILICIO (SI) - 36",
    "SODIO (NA) - 31",
    "POTASIO (K) - 27",
    "VANADIO (V) - 39",
    "BORO (B) - 18",
    "BARIO (BA) - 21",
    "CALCIO (CA) - 22",
    "CADMIO (CD) - 23",
    "MAGNESIO (MG) - 28",
    "MANGANESO (MN) - 29",
    "FÓSFORO (P) - 34",
    "ZINC (ZN) - 40",
    "CÓDIGO ISO (4/6/14) - 47",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48",
    "OXIDACIÓN - 80",
    "NITRACIÓN - 82",
    "NÚMERO ÁCIDO (AN) - 43",
    "NÚMERO BÁSICO (BN) - 12",
    "NÚMERO BÁSICO (BN) - 17",
    "HOLLÍN - 79",
    "DILUCIÓN POR COMBUSTIBLE - 46",
    "AGUA (IR) - 81",
    "CONTENIDO AGUA (KARL FISCHER) - 41",
    "CONTENIDO GLICOL - 105",
    "VISCOSIDAD A 100 °C - 13",
    "VISCOSIDAD A 40 °C - 14",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51",
    "AGUA CUALITATIVA (PLANCHA) - 360",
    "AGUA LIBRE - 416",
    "ANÁLISIS ANTIOXIDANTES (AMINA) - 44",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45",
    "COBRE (CU) - 119",
    "ESPUMA SEC 1 - ESTABILIDAD - 60",
    "ESPUMA SEC 1 - TENDENCIA - 59",
    "ESTAÑO (SN) - 37",
    "ÍNDICE VISCOSIDAD - 359",
    "RPVOT - 10",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83",
    "**ULTRACENTRÍFUGA (UC) - 1",
    "ESTADO_PRODUCTO",
    "ESTADO_DESGASTE",
    "ESTADO_CONTAMINACION",
    "N_SOLICITUD",
    "CAMBIO_DE_PRODUCTO",
    "CAMBIO_DE_FILTRO",
    "TEMPERATURA_RESERVORIO",
    "UNIDAD_TEMPERATURA_RESERVORIO",
    "COMENTARIO_CLIENTE",
    "TIPO_DE_COMBUSTIBLE",
    "TIPO_DE_REFRIGERANTE",
    "USUARIO",
    "COMENTARIO_REPORTE",
    "id_muestra",
    "Archivo_Origen",
    "ESTADO_MUESTRA",
    "AGUA (IR) - 74",
    "AGUA (IR) - 74 - Estado",
    "AGUA (IR) - 81 - Estado",
    "AGUA LIBRE - 416 - Estado",
    "AGUA CUALITATIVA (PLANCHA) - 360 - Estado",
    "ALUMINIO (AL) - 20 - Estado",
    "BARIO (BA) - 21 - Estado",
    "BORO (B) - 18 - Estado",
    "CALCIO (CA) - 22 - Estado",
    "CADMIO (CD) - 23 - Estado",
    "COBRE (CU) - 25 - Estado",
    "COBRE (CU) - 119 - Estado",
    "CROMO (CR) - 24 - Estado",
    "HIERRO (FE) - 26 - Estado",
    "MAGNESIO (MG) - 28 - Estado",
    "MANGANESO (MN) - 29 - Estado",
    "MOLIBDENO (MO) - 30 - Estado",
    "NÍQUEL (NI) - 32 - Estado",
    "PLATA (AG) - 19 - Estado",
    "PLOMO (PB) - 35 - Estado",
    "POTASIO (K) - 27 - Estado",
    "SILICIO (SI) - 36 - Estado",
    "SODIO (NA) - 31 - Estado",
    "TITANIO (TI) - 38 - Estado",
    "VANADIO (V) - 39 - Estado",
    "ZINC (ZN) - 40 - Estado",
    "ESTAÑO (SN) - 37 - Estado",
    "FÓSFORO (P) - 34 - Estado",
    "CÓDIGO ISO (4/6/14) - 47 - Estado",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49 - Estado",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50 - Estado",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48 - Estado",
    "OXIDACIÓN - 80 - Estado",
    "NITRACIÓN - 82 - Estado",
    "ÍNDICE PQ (PQI) - 3 - Estado",
    "NÚMERO ÁCIDO (AN) - 43 - Estado",
    "NÚMERO BÁSICO (BN) - 12 - Estado",
    "NÚMERO BÁSICO (BN) - 17 - Estado",
    "CONTENIDO AGUA (KARL FISCHER) - 41 - Estado",
    "ANÁLISIS ANTIOXIDANTES (AMINA) - 44 - Estado",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45 - Estado",
    "HOLLÍN - 73",
    "HOLLÍN - 73 - Estado",
    "HOLLÍN - 79 - Estado",
    "DILUCIÓN POR COMBUSTIBLE - 46 - Estado",
    "VISCOSIDAD A 40 °C - 14 - Estado",
    "VISCOSIDAD A 100 °C - 13 - Estado",
    "ÍNDICE VISCOSIDAD - 359 - Estado",
    "ESPUMA SEC 1 - ESTABILIDAD - 60 - Estado",
    "ESPUMA SEC 1 - TENDENCIA - 59 - Estado",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado",
    "RESIDUO CARBÓN (MCR) - 361",
    "RESIDUO CARBÓN (MCR) - 361 - Estado",
    "PUNTO DE INFLAMACIÓN (PMA) - 61",
    "PUNTO DE INFLAMACIÓN (PMA) - 61 - Estado",
    "RPVOT - 10 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83 - Estado",
    "**ULTRACENTRÍFUGA (UC) - 1 - Estado",
]

# ============================================================
# UTILIDADES
# ============================================================
def limpiar_columna(col):
    if col is None:
        return ""
    return str(col).replace("\ufeff", "").replace("\n", " ").strip()


def normalizar_columnas(lista_columnas):
    return [limpiar_columna(c) for c in lista_columnas]


def analizar_columnas(columnas_archivo, columnas_esperadas):
    faltantes = [c for c in columnas_esperadas if c not in columnas_archivo]
    extras = [c for c in columnas_archivo if c not in columnas_esperadas]
    mismo_orden = columnas_archivo == columnas_esperadas

    diferencias_orden = []
    max_len = min(len(columnas_archivo), len(columnas_esperadas))
    for i in range(max_len):
        if columnas_archivo[i] != columnas_esperadas[i]:
            diferencias_orden.append({
                "Posición": i + 1,
                "Esperado": columnas_esperadas[i],
                "Encontrado": columnas_archivo[i],
            })
    return faltantes, extras, mismo_orden, diferencias_orden


def generar_motivo_invalidez(faltantes, extras, mismo_orden, diferencias_orden, cantidad_ok):
    motivos = []

    if not cantidad_ok:
        motivos.append("Cantidad de columnas incorrecta")

    if faltantes:
        motivos.append(f"Faltan {len(faltantes)} columna(s)")

    if extras:
        motivos.append(f"Hay {len(extras)} columna(s) adicional(es)")

    if not mismo_orden:
        motivos.append(f"Orden incorrecto en {len(diferencias_orden)} posición(es)")

    if not motivos:
        return "Archivo válido"

    return " | ".join(motivos)


def leer_excel_subido(uploaded_file, sheet_option):
    if sheet_option == "Primera hoja" or sheet_option == 0:
        return pd.read_excel(uploaded_file, sheet_name=0, dtype=object)
    return pd.read_excel(uploaded_file, sheet_name=sheet_option, dtype=object)


def validar_archivo(uploaded_file, sheet_option):
    try:
        df = leer_excel_subido(uploaded_file, sheet_option)
        columnas_archivo = normalizar_columnas(df.columns.tolist())
        columnas_esperadas = normalizar_columnas(COLUMNAS_ESPERADAS)

        faltantes, extras, mismo_orden, diferencias_orden = analizar_columnas(
            columnas_archivo, columnas_esperadas
        )

        cantidad_ok = len(columnas_archivo) == len(columnas_esperadas)
        nombres_ok = len(faltantes) == 0 and len(extras) == 0
        valido = cantidad_ok and nombres_ok and mismo_orden

        motivo = generar_motivo_invalidez(
            faltantes,
            extras,
            mismo_orden,
            diferencias_orden,
            cantidad_ok
        )

        df.columns = columnas_archivo

        return {
            "archivo": uploaded_file.name,
            "valido": valido,
            "motivo": motivo,
            "cantidad_archivo": len(columnas_archivo),
            "cantidad_esperada": len(columnas_esperadas),
            "faltantes": faltantes,
            "extras": extras,
            "mismo_orden": mismo_orden,
            "diferencias_orden": diferencias_orden,
            "error": "",
            "df": df if valido else None,
        }
    except Exception as e:
        return {
            "archivo": uploaded_file.name,
            "valido": False,
            "motivo": f"Error al leer archivo: {str(e)}",
            "cantidad_archivo": None,
            "cantidad_esperada": len(COLUMNAS_ESPERADAS),
            "faltantes": [],
            "extras": [],
            "mismo_orden": False,
            "diferencias_orden": [],
            "error": str(e),
            "df": None,
        }


def crear_excel_bytes(dataframes_por_hoja):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre_hoja, df in dataframes_por_hoja.items():
            df.to_excel(writer, index=False, sheet_name=nombre_hoja[:31])
    output.seek(0)
    return output.getvalue()


def construir_reporte(resultados):
    resumen = []
    detalle = []

    for r in resultados:
        resumen.append({
            "Archivo": r["archivo"],
            "Válido": "Sí" if r["valido"] else "No",
            "Motivo": r.get("motivo", ""),
            "Cantidad columnas archivo": r["cantidad_archivo"],
            "Cantidad columnas esperadas": r["cantidad_esperada"],
            "Mismo orden": "Sí" if r["mismo_orden"] else "No",
            "Faltantes": len(r["faltantes"]),
            "Extras": len(r["extras"]),
            "Error": r["error"],
        })

        for c in r["faltantes"]:
            detalle.append({
                "Archivo": r["archivo"],
                "Tipo": "Faltante",
                "Detalle": c
            })

        for c in r["extras"]:
            detalle.append({
                "Archivo": r["archivo"],
                "Tipo": "Extra",
                "Detalle": c
            })

        for d in r["diferencias_orden"]:
            detalle.append({
                "Archivo": r["archivo"],
                "Tipo": "Orden incorrecto",
                "Detalle": f"Posición {d['Posición']}: esperado='{d['Esperado']}' | encontrado='{d['Encontrado']}'",
            })

        if r["error"]:
            detalle.append({
                "Archivo": r["archivo"],
                "Tipo": "Error lectura",
                "Detalle": r["error"]
            })

    df_resumen = pd.DataFrame(resumen)
    df_detalle = pd.DataFrame(detalle)
    return crear_excel_bytes({"Resumen": df_resumen, "Detalle": df_detalle})


def construir_combinado(resultados_validos, agregar_fuente=True):
    dfs = []
    for r in resultados_validos:
        df = r["df"].copy()
        if agregar_fuente:
            df["__ARCHIVO_FUENTE__"] = r["archivo"]
        dfs.append(df)

    df_final = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    return crear_excel_bytes({"Combinado": df_final}), df_final


# ============================================================
# INTERFAZ
# ============================================================
st.title("📊 Combinador de Excel con validación de columnas")
st.markdown(
    "Sube varios archivos Excel, valida que tengan exactamente las mismas columnas y combina únicamente los que cumplan."
)

with st.expander("Ver reglas de validación"):
    st.write("- Se valida cantidad de columnas.")
    st.write("- Se validan nombres exactos de columnas.")
    st.write("- Se valida el orden exacto de las columnas.")
    st.write("- Si un archivo no cumple, se reporta el motivo exacto antes de combinar.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    modo = st.radio(
        "Modo de combinación",
        ["Combinar solo válidos", "Modo estricto, si uno falla no combinar"],
        horizontal=False,
    )
with col2:
    hoja_opcion = st.text_input("Hoja a leer", value="Primera hoja")
with col3:
    agregar_fuente = st.checkbox("Agregar columna con archivo fuente", value=True)

st.markdown("### Cargar archivos")
archivos = st.file_uploader(
    "Selecciona los archivos Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if archivos:
    resultados = []
    for archivo in archivos:
        sheet_option = 0 if hoja_opcion.strip().lower() == "primera hoja" else hoja_opcion
        resultados.append(validar_archivo(archivo, sheet_option))

    validos = [r for r in resultados if r["valido"]]
    invalidos = [r for r in resultados if not r["valido"]]

    st.markdown("### Resumen")
    c1, c2, c3 = st.columns(3)
    c1.metric("Archivos cargados", len(resultados))
    c2.metric("Válidos", len(validos))
    c3.metric("Inválidos", len(invalidos))

    resumen_df = pd.DataFrame([
        {
            "Archivo": r["archivo"],
            "Válido": "Sí" if r["valido"] else "No",
            "Motivo": r.get("motivo", ""),
            "Columnas archivo": r["cantidad_archivo"],
            "Columnas esperadas": r["cantidad_esperada"],
            "Mismo orden": "Sí" if r["mismo_orden"] else "No",
            "Faltantes": len(r["faltantes"]),
            "Extras": len(r["extras"]),
            "Error": r["error"],
        }
        for r in resultados
    ])
    st.dataframe(resumen_df, use_container_width=True)

    if invalidos:
        st.markdown("### Archivos con problemas")
        for r in invalidos:
            with st.expander(f"❌ {r['archivo']}"):
                st.error(f"Motivo exacto: {r.get('motivo', 'No identificado')}")
                if r["error"]:
                    st.error(r["error"])

                if r["faltantes"]:
                    st.write("**Columnas faltantes**")
                    st.dataframe(
                        pd.DataFrame({"Faltantes": r["faltantes"]}),
                        use_container_width=True
                    )

                if r["extras"]:
                    st.write("**Columnas extra**")
                    st.dataframe(
                        pd.DataFrame({"Extras": r["extras"]}),
                        use_container_width=True
                    )

                if r["diferencias_orden"]:
                    st.write("**Diferencias de orden**")
                    st.dataframe(
                        pd.DataFrame(r["diferencias_orden"]),
                        use_container_width=True
                    )
    else:
        st.success("Todos los archivos cumplen con la estructura esperada.")

    st.markdown("### Descargas")
    reporte_bytes = construir_reporte(resultados)
    nombre_reporte = f"Reporte_Validacion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        "📥 Descargar reporte de validación",
        data=reporte_bytes,
        file_name=nombre_reporte,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    bloquear = modo == "Modo estricto, si uno falla no combinar" and len(invalidos) > 0

    if bloquear:
        st.warning("No se generó el archivo combinado porque activaste modo estricto y hay archivos inválidos.")
    elif len(validos) == 0:
        st.warning("No hay archivos válidos para combinar.")
    else:
        combinado_bytes, df_final = construir_combinado(validos, agregar_fuente=agregar_fuente)
        nombre_combinado = f"Excel_Combinado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        st.download_button(
            "📥 Descargar Excel combinado",
            data=combinado_bytes,
            file_name=nombre_combinado,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.markdown("### Vista previa del combinado")
        st.dataframe(df_final.head(50), use_container_width=True)
else:
    st.info("Carga uno o más archivos Excel para comenzar.")
