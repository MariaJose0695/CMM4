import streamlit as st
import pandas as pd
import numpy as np
import io
import xml.etree.ElementTree as ET
from collections import OrderedDict

st.set_page_config(page_title="Comparación Perceptron Frontal vs Final", layout="wide")
st.title("📈 Comparación Perceptron Frontal vs Final")

archivo = st.file_uploader("📤 Carga el archivo Excel", type=["xlsx"])

def generar_xml_como_texto(df, station_name="T1XX_FLEX_Front_Mod", model_name="K_SUV") -> str:
    gauge = ET.Element("GAUGE")
    station = ET.SubElement(gauge, "STATION")
    ET.SubElement(station, "NAME").text = station_name

    model = ET.SubElement(station, "MODEL")
    ET.SubElement(model, "NAME").text = model_name

    checkpoints = {}
    for _, row in df.iterrows():
        perc_axis = row["Perc Axis"]  # Ejemplo: 3000L[X]
        offset = row["Calculated Offset"]

        if "[" in perc_axis and "]" in perc_axis:
            checkpoint = perc_axis.split("[")[0]
            axis = perc_axis.split("[")[1].replace("]", "")
            if checkpoint not in checkpoints:
                checkpoints[checkpoint] = {}
            checkpoints[checkpoint][axis] = offset

    for checkpoint, axes in checkpoints.items():
        cp_elem = ET.SubElement(model, "CHECKPOINT")
        ET.SubElement(cp_elem, "NAME").text = checkpoint

        for axis_name in ["X", "Y", "Z"]:
            axis_elem = ET.SubElement(cp_elem, "AXIS")
            ET.SubElement(axis_elem, "NAME").text = axis_name
            ET.SubElement(axis_elem, "OFFSET").text = str(round(axes.get(axis_name, 0.0), 3))

        diam_elem = ET.SubElement(cp_elem, "AXIS")
        ET.SubElement(diam_elem, "NAME").text = "Diameter"
        ET.SubElement(diam_elem, "OFFSET").text = "0"

    xml_bytes = ET.tostring(gauge, encoding="utf-8", method="xml")
    xml_string = xml_bytes.decode("utf-8").replace("\n", "").replace("\r", "")
    return xml_string

if archivo:
    try:
        perceptron_df = pd.read_excel(archivo, sheet_name="Perceptron")
        cmm_df = pd.read_excel(archivo, sheet_name="CMM")
        mapping_df = pd.read_excel(archivo, sheet_name="JSN-Mapping")
        axis_df = pd.read_excel(archivo, sheet_name="Axis-Mapping")

        st.success(" ✅ Archivo cargado correctamente.")

        if st.button("▶️ Ejecutar comparación"):
            resultados = []

            for _, axis_row in axis_df.iterrows():
                perc_axis = axis_row['PerceptronAxis']
                cmm_axis = axis_row['CMMAxis']

                valores_perc = []
                valores_cmm = []

                for _, map_row in mapping_df.iterrows():
                    jsn_perc = map_row['PerceptronJSN']
                    jsn_cmm = map_row['CMMJSN']

                    valor_perc = perceptron_df.loc[perceptron_df['JSN'] == jsn_perc, perc_axis]
                    valor_cmm = cmm_df.loc[cmm_df['JSN'] == jsn_cmm, cmm_axis]

                    if not valor_perc.empty and not valor_cmm.empty:
                        valores_perc.append(float(valor_perc))
                        valores_cmm.append(float(valor_cmm))

                if len(valores_perc) > 1:
                    perc_mean = np.mean(valores_perc)
                    cmm_mean = np.mean(valores_cmm)
                    correlacion = np.corrcoef(valores_perc, valores_cmm)[0,1]
                    desvest = np.std(np.array(valores_perc) - np.array(valores_cmm), ddof=1)
                    offset_calc = cmm_mean - perc_mean

                    resultados.append({
                        "Perc Axis": perc_axis,
                        "CMM Axis": cmm_axis,
                        "Perc Mean": round(perc_mean, 3),
                        "CMM Mean": round(cmm_mean, 3),
                        "Correlation coefficient": round(correlacion, 3),
                        "6 Sigma": round(6 * desvest, 3),
                        "Calculated Offset": round(offset_calc, 3)
                    })

            output_df = pd.DataFrame(resultados)
            st.subheader("📈 Resultados")
            st.dataframe(output_df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                output_df.to_excel(writer, sheet_name="Offset Summary", index=False)
            st.download_button(
                label="📥 Descargar resultados en Excel",
                data=buffer.getvalue(),
                file_name="Resultados_Offsets.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            xml_string = generar_xml_como_texto(output_df)
            st.text_area("📄 XML generado", xml_string, height=300)

            st.download_button(
                label="📥 Descargar XML",
                data=xml_string.encode("utf-8"),
                file_name="resultado.xml",
                mime="application/xml"
            )

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")

# Sección para convertir archivo TXT de CMM a Excel horizontal
st.subheader("📄 Convertir archivo TXT CMM a Excel horizontal")
archivo_txt = st.file_uploader("📤 Carga el archivo TXT de CMM", type=["txt"], key="txt_cmm")

puntos_deseados = [
 "DIM 1000L= UBICACIÓN DE RAN 1000L_  UNIDADES=MM"
]

if archivo_txt:
    try:
        contenido = archivo_txt.read().decode("latin-1").splitlines()
        datos = {}
        dim_actual = ""
        jsn = psn = date = time = cycle_mode = ""

        for linea in contenido:
            linea = linea.strip()
            if "JSN=" in linea: jsn = linea.split("JSN=")[1].strip()
            if "PSN=" in linea: psn = linea.split("PSN=")[1].strip()
            if "DATE=" in linea or "FECHA=" in linea: date = linea.split("=")[1].strip()
            if "TIME=" in linea: time = linea.split("TIME=")[1].strip()
            if "CYCLE MODE=" in linea: cycle_mode = linea.split("CYCLE MODE=")[1].strip()

            if linea.startswith("DIM ") and "UNIDADES=MM" in linea:
                dim_actual = linea.split("=")[0].replace("DIM", "").strip()
                continue

            partes = linea.split()
            if len(partes) >= 6 and partes[0] in ["X", "Y", "Z"]:
                try:
                    eje = partes[0]
                    desviacion = float(partes[5])
                    clave = f"{dim_actual}[{eje}]"
                    if clave in puntos_deseados:
                        datos[clave] = desviacion
                except:
                    continue

        fila = OrderedDict({
            "JSN": jsn,
            "PSN": psn,
            "Date": date,
            "Time": time,
            "CYCLE MODE": cycle_mode
        })

        for punto in puntos_deseados:
            fila[punto] = datos.get(punto, "")

        df_horizontal = pd.DataFrame([fila])
        st.success("✅ TXT convertido con puntos filtrados.")
        st.dataframe(df_horizontal)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_horizontal.to_excel(writer, index=False, sheet_name="Horizontal")

        st.download_button(
            label="📥 Descargar Excel horizontal",
            data=buffer.getvalue(),
            file_name="Horizontal_CMM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo TXT: {e}")
