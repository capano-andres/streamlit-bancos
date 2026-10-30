import streamlit as st
from frances import procesar_bbva_frances
from santander import procesar_santander_rio
from santander_prueba import procesar_santander_rio_prueba
from galicia import procesar_galicia
from icbc import procesar_icbc
from icbc_2 import procesar_icbc_formato_2
from icbc_formato_3 import procesar_icbc_formato_3
from macro import procesar_macro
from nacion import procesar_nacion
from provincia import procesar_provincia
from provincia_2 import procesar_provincia_formato_2
from supervielle import procesar_supervielle
from hipotecario import procesar_hipotecario
from hsbc import procesar_hsbc
from credicoop import procesar_credicoop
from mercadopago import procesar_mercadopago
from credicoop_2 import procesar_credicoop_formato_2
from macro_2 import procesar_macro_formato_2
from macro_3 import procesar_macro_formato_3
from macro_4 import procesar_macro_formato_4
from galicia_mas import procesar_galicia_mas
from comafi import procesar_comafi
from ciudad import procesar_ciudad
from patagonia import procesar_patagonia
from patagonia_2 import procesar_patagonia_formato_2
from patagonia_3 import procesar_patagonia_formato_3
from brubank import procesar_brubank
from sol import procesar_sol
from uala import procesar_uala
from cocos import procesar_cocos

st.set_page_config(page_title="Movimientos Bancos", page_icon="🏦")

# Lista de bancos (orden alfabético)
bancos = [
    "Banco del Sol",
    "BBVA Frances",
    "Brubank",
    "Ciudad",
    "Cocos",
    "Comafi",
    "Credicoop",
    "Credicoop (Formato 2)",
    "Galicia",
    "Galicia Más",
    "Hipotecario",
    "HSBC",
    "ICBC (Formato 1)",
    "ICBC (Formato 2)",
    "ICBC (Formato 3)",
    "Macro",
    "Macro (Formato 2)",
    "Macro (Formato 3)",
    "Macro (Formato 4)",
    "MercadoPago",
    "Nacion",
    "Patagonia",
    "Patagonia (Formato 2)",
    "Patagonia (Formato 3)",
    "Provincia",
    "Provincia (Formato 2)",
    "Santander Rio",
    "Santander Rio (Prueba)",
    "Supervielle",
    "Ualá",
]


def procesar_banco(banco_seleccionado, archivo_pdf):
    """Función principal que dirige el procesamiento según el banco seleccionado"""
    if banco_seleccionado == "BBVA Frances":
        return procesar_bbva_frances(archivo_pdf)
    elif banco_seleccionado == "Brubank":
        return procesar_brubank(archivo_pdf)
    elif banco_seleccionado == "Santander Rio":
        return procesar_santander_rio(archivo_pdf)
    elif banco_seleccionado == "Santander Rio (Prueba)":
        return procesar_santander_rio_prueba(archivo_pdf, cuits_propios=cuits_propios)
    elif banco_seleccionado == "Galicia":
        return procesar_galicia(archivo_pdf)
    elif banco_seleccionado == "Galicia Más":
        return procesar_galicia_mas(archivo_pdf)
    elif banco_seleccionado == "ICBC (Formato 1)":
        return procesar_icbc(archivo_pdf)
    elif banco_seleccionado == "ICBC (Formato 2)":
        return procesar_icbc_formato_2(archivo_pdf)
    elif banco_seleccionado == "ICBC (Formato 3)":
        return procesar_icbc_formato_3(archivo_pdf)
    elif banco_seleccionado == "Macro":
        return procesar_macro(archivo_pdf)
    elif banco_seleccionado == "Macro (Formato 2)":
        return procesar_macro_formato_2(archivo_pdf)
    elif banco_seleccionado == "Macro (Formato 3)":
        return procesar_macro_formato_3(archivo_pdf)
    elif banco_seleccionado == "Macro (Formato 4)":
        return procesar_macro_formato_4(archivo_pdf)
    elif banco_seleccionado == "Nacion":
        return procesar_nacion(archivo_pdf)
    elif banco_seleccionado == "Provincia":
        return procesar_provincia(archivo_pdf)
    elif banco_seleccionado == "Provincia (Formato 2)":
        return procesar_provincia_formato_2(archivo_pdf)
    elif banco_seleccionado == "Supervielle":
        return procesar_supervielle(archivo_pdf)
    elif banco_seleccionado == "Hipotecario":
        return procesar_hipotecario(archivo_pdf)
    elif banco_seleccionado == "HSBC":
        return procesar_hsbc(archivo_pdf)
    elif banco_seleccionado == "Credicoop":
        return procesar_credicoop(archivo_pdf)
    elif banco_seleccionado == "Credicoop (Formato 2)":
        return procesar_credicoop_formato_2(archivo_pdf)
    elif banco_seleccionado == "MercadoPago":
        return procesar_mercadopago(archivo_pdf)
    elif banco_seleccionado == "Comafi":
        return procesar_comafi(archivo_pdf)
    elif banco_seleccionado == "Ciudad":
        return procesar_ciudad(archivo_pdf)
    elif banco_seleccionado == "Patagonia":
        return procesar_patagonia(archivo_pdf)
    elif banco_seleccionado == "Patagonia (Formato 2)":
        return procesar_patagonia_formato_2(archivo_pdf)
    elif banco_seleccionado == "Patagonia (Formato 3)":
        return procesar_patagonia_formato_3(archivo_pdf)
    elif banco_seleccionado == "Banco del Sol":
        return procesar_sol(archivo_pdf)
    elif banco_seleccionado == "Ualá":
        return procesar_uala(archivo_pdf)
    elif banco_seleccionado == "Cocos":
        return procesar_cocos(archivo_pdf)
    else:
        st.info(f"Lógica para {banco_seleccionado} aún no implementada")
        return None


# Interfaz principal de Streamlit
st.title("Selector de Banco y Subida de PDF")

# Selector de banco
banco_seleccionado = st.selectbox("Selecciona un banco:", bancos)

# Input CUITs propios (solo para Santander Prueba)
cuits_propios = []
if banco_seleccionado == "Santander Rio (Prueba)":
    st.markdown("---")
    st.subheader("CUITs propios (transferencias entre bancos)")
    st.caption("Agregá los CUITs del titular, socios o dueños para identificar transferencias propias.")
    cant_cuits = st.number_input("Cantidad de CUITs", min_value=0, max_value=10, value=0, step=1)
    for i in range(int(cant_cuits)):
        col1, col2 = st.columns(2)
        with col1:
            cuit = st.text_input(f"CUIT #{i+1}", key=f"cuit_{i}", placeholder="30711511004")
        with col2:
            razon = st.text_input(f"Razón Social #{i+1}", key=f"razon_{i}", placeholder="Empresa SA")
        if cuit and cuit.strip() or razon and razon.strip():
            cuit_val = cuit.strip().replace("-", "") if cuit else ""
            razon_val = razon.strip() if razon else ""
            label = razon_val if razon_val else f"CUIT {cuit_val}"
            cuits_propios.append((cuit_val, razon_val, label))
    st.markdown("---")

# Subida de archivo PDF
archivo_pdf = st.file_uploader("Sube un archivo PDF", type=["pdf"])

if archivo_pdf is not None:
    st.success(f"Archivo '{archivo_pdf.name}' subido correctamente.")

    # Procesar el archivo según el banco seleccionado
    resultado = procesar_banco(banco_seleccionado, archivo_pdf)

    if resultado is not None:
        # Determinar el nombre del archivo según el banco
        nombre_archivo = f"{banco_seleccionado}.xlsx"

        st.download_button(
            label="Descargar archivo Excel procesado",
            data=resultado,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
