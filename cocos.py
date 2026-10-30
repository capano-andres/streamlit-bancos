import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

def clean_for_excel(text):
    """Elimina caracteres ilegales para Excel y espacios extra"""
    if not text: return ""
    text = str(text)
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    return text.strip()

def parse_numero_ar(s):
    """Parsea número en formato argentino (1.234,56) a float. Soporta sufijo - para negativos."""
    if not s or s == "0": return 0.0
    s = s.strip()
    neg = s.startswith("-")
    s = s.lstrip("-")
    s = s.replace(".", "").replace(",", ".")
    try:
        val = float(s)
        return -val if neg else val
    except:
        return 0.0

NUM_RE = re.compile(r'^-?[\d.,]+$')
DATE_RE = re.compile(r'^\d{2}-\d{2}-\d{4}$')

BOILERPLATE_MARKERS = [
    "RESUMEN DE CUENTA", "NRO DE COMITENTE", "FECHA REPORTE",
    "POSICION AL", "INCREMENTOS/DECREMENTOS",
]
BOILERPLATE_RE = re.compile(r'^DESDE\s+\d{2}-\d{2}-\d{4}')

def convertir_fecha_guion(fecha_str):
    """Convierte '01-07-2025' a '01/07/2025'"""
    try:
        dia, mes, anio = fecha_str.split("-")
        return f"{dia}/{mes}/{anio}"
    except:
        return fecha_str

def _agrupar_filas(words):
    """Agrupa palabras de una página en filas según su coordenada 'top'."""
    filas = []
    for w in words:
        placed = False
        for f in filas:
            if abs(f["top"] - w["top"]) < 2.5:
                f["words"].append(w)
                placed = True
                break
        if not placed:
            filas.append({"top": w["top"], "words": [w]})
    filas.sort(key=lambda f: f["top"])
    return filas

def procesar_cocos(archivo_pdf):
    st.info("Procesando archivo de Cocos Capital...")
    try:
        archivo_pdf.seek(0)
        with pdfplumber.open(io.BytesIO(archivo_pdf.read())) as pdf:
            texto_completo = ""
            paginas_filas = []
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    texto_completo += t + "\n"
                words = page.extract_words(x_tolerance=1.5)
                paginas_filas.append(_agrupar_filas(words))

        # ============================================================
        # 1. METADATOS
        # ============================================================
        cuenta = "Sin Especificar"
        periodo = "Sin Especificar"

        m_cuenta = re.search(r'Nro de comitente:\s*(\d+)', texto_completo)
        if m_cuenta:
            cuenta = m_cuenta.group(1)

        m_periodo = re.search(
            r'Desde\s+(\d{2}-\d{2}-\d{4})-?\s*Hasta\s+(\d{2}-\d{2}-\d{4})', texto_completo
        )
        if m_periodo:
            periodo = f"Del {convertir_fecha_guion(m_periodo.group(1))} al {convertir_fecha_guion(m_periodo.group(2))}"

        titular = f"Comitente {cuenta}"

        # ============================================================
        # 2. SALDOS DE CAJA (ARS/USD) desde las tablas POSICION AL
        # ============================================================
        saldos_ars = []
        saldos_usd = []
        for filas in paginas_filas:
            for f in filas:
                ws = sorted(f["words"], key=lambda w: w["x0"])
                if not ws:
                    continue
                primera = ws[0]
                if primera["x0"] < 40 and primera["text"] in ("ARS", "USD"):
                    numeros = [w["text"] for w in ws if NUM_RE.match(w["text"])]
                    if numeros:
                        val = parse_numero_ar(numeros[-1])
                        if primera["text"] == "ARS":
                            saldos_ars.append(val)
                        else:
                            saldos_usd.append(val)

        saldo_inicial_ars = saldos_ars[0] if saldos_ars else 0.0
        saldo_final_ars = saldos_ars[-1] if saldos_ars else 0.0
        saldo_inicial_usd = saldos_usd[0] if saldos_usd else 0.0
        saldo_final_usd = saldos_usd[-1] if saldos_usd else 0.0

        # ============================================================
        # 3. MOVIMIENTOS
        # ============================================================
        movimientos_raw = []
        en_movimientos = False
        bloque = None

        def cerrar_bloque():
            nonlocal bloque
            if bloque:
                movimientos_raw.append(bloque)
            bloque = None

        for filas in paginas_filas:
            for f in filas:
                ws = sorted(f["words"], key=lambda w: w["x0"])
                if not ws:
                    continue
                linea_texto = " ".join(w["text"] for w in ws)
                linea_upper = linea_texto.upper()

                if not en_movimientos:
                    if linea_upper.strip() == "MOVIMIENTOS":
                        en_movimientos = True
                    continue

                if linea_upper.startswith("POSICION AL"):
                    cerrar_bloque()
                    en_movimientos = False
                    continue

                if any(m in linea_upper for m in BOILERPLATE_MARKERS):
                    continue
                if BOILERPLATE_RE.match(linea_upper.strip()):
                    continue
                if "FECHA" in linea_upper and ("LIQ" in linea_upper or "CONC" in linea_upper or "COMPROBANTE" in linea_upper):
                    continue
                if linea_upper.strip() in ("LIQ", "CONC"):
                    continue

                fecha_liq_val = ws[0]["text"] if DATE_RE.match(ws[0]["text"]) else None

                if fecha_liq_val:
                    cerrar_bloque()
                    comprobante = []
                    especie = []
                    numeros = []
                    for w in ws[2:]:
                        if DATE_RE.match(w["text"]):
                            continue
                        if w["x0"] >= 420 and NUM_RE.match(w["text"]):
                            numeros.append(w["text"])
                        elif 145 <= w["x0"] < 280:
                            comprobante.append(w["text"])
                        elif 280 <= w["x0"] < 420:
                            especie.append(w["text"])
                    bloque = {
                        "fecha": convertir_fecha_guion(fecha_liq_val),
                        "comprobante": comprobante,
                        "especie": especie,
                        "numeros": numeros,
                    }
                else:
                    if bloque:
                        for w in ws:
                            if 145 <= w["x0"] < 280:
                                bloque["comprobante"].append(w["text"])
                            elif 280 <= w["x0"] < 420:
                                bloque["especie"].append(w["text"])
        cerrar_bloque()

        def fecha_sort_key(fecha_str):
            try:
                d, m, a = fecha_str.split("/")
                return (int(a), int(m), int(d))
            except:
                return (0, 0, 0)

        transactions_ars = []
        transactions_usd = []
        for mov in movimientos_raw:
            if len(mov["numeros"]) != 5:
                continue
            precio_str, cantidad_str, ars_str, usd_str, _usdext = mov["numeros"]
            comprobante_txt = " ".join(mov["comprobante"]).strip()
            especie_txt = " ".join(mov["especie"]).strip()

            # El comprobante trae "<Tipo de operación> - <número>"
            if " - " in comprobante_txt:
                tipo_op, numero_comp = comprobante_txt.rsplit(" - ", 1)
            else:
                tipo_op, numero_comp = comprobante_txt, ""

            especie_mostrar = "" if especie_txt in ("ARS", "USD") else especie_txt

            precio_val = parse_numero_ar(precio_str)
            cantidad_val = parse_numero_ar(cantidad_str)
            ars_val = parse_numero_ar(ars_str)
            usd_val = parse_numero_ar(usd_str)

            fila = {
                "Fecha": mov["fecha"],
                "Tipo": clean_for_excel(tipo_op),
                "Comprobante": clean_for_excel(numero_comp),
                "Especie": clean_for_excel(especie_mostrar),
                "Cantidad": round(cantidad_val, 2),
                "Precio": round(precio_val, 2),
            }

            if ars_val != 0:
                transactions_ars.append({**fila, "Importe": round(ars_val, 2)})
            if usd_val != 0:
                transactions_usd.append({**fila, "Importe": round(usd_val, 2)})

        transactions_ars.sort(key=lambda r: fecha_sort_key(r["Fecha"]))
        transactions_usd.sort(key=lambda r: fecha_sort_key(r["Fecha"]))

        if not movimientos_raw:
            st.info("No se encontraron movimientos. Se generará el Excel solo con los saldos.")
        else:
            st.success(f"Se encontraron {len(movimientos_raw)} movimientos.")

        # ============================================================
        # 4. GENERAR EXCEL (libro de operaciones, una hoja por moneda)
        # ============================================================
        # Este extracto no es una cuenta bancaria (no hay depósitos/extracciones
        # con saldo corrido): son compras/ventas, suscripciones/rescates de FCI
        # y cauciones. Por eso se arma como un libro de operaciones cronológico
        # en vez del formato CRÉDITOS/DÉBITOS con control de saldos de los bancos.
        output = io.BytesIO()
        wb = Workbook()
        wb.remove(wb.active)

        color_bg_main = "1E40AF"
        color_txt_main = "FFFFFF"
        color_header = "334155"

        thin_border = Border(
            left=Side(style='thin', color="A6A6A6"),
            right=Side(style='thin', color="A6A6A6"),
            top=Side(style='thin', color="A6A6A6"),
            bottom=Side(style='thin', color="A6A6A6")
        )

        fill_row_cred = PatternFill(start_color="F2F9F1", end_color="F2F9F1", fill_type="solid")
        fill_row_deb = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
        font_cred = Font(color="2E7D32")
        font_deb = Font(color="C00000")

        def crear_hoja(nombre_hoja, datos, saldo_ini, saldo_fin, formato_moneda):
            ws = wb.create_sheet(title=nombre_hoja)
            ws.sheet_view.showGridLines = False

            ws.merge_cells("A1:G1")
            tit = ws["A1"]
            tit.value = f"LIBRO DE OPERACIONES COCOS ({nombre_hoja}) - {clean_for_excel(titular)}"
            tit.font = Font(size=14, bold=True, color=color_txt_main)
            tit.fill = PatternFill(start_color=color_bg_main, end_color=color_bg_main, fill_type="solid")
            tit.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 25

            ws["A3"] = "SALDO DE CAJA INICIAL"
            ws["A3"].font = Font(bold=True, size=10, color="666666")
            ws["B3"] = saldo_ini
            ws["B3"].number_format = formato_moneda
            ws["B3"].font = Font(bold=True, size=11)
            ws["B3"].border = Border(bottom=Side(style='thin', color="DDDDDD"))

            ws["A4"] = "SALDO DE CAJA FINAL"
            ws["A4"].font = Font(bold=True, size=10, color="666666")
            ws["B4"] = saldo_fin
            ws["B4"].number_format = formato_moneda
            ws["B4"].font = Font(bold=True, size=11)
            ws["B4"].border = Border(bottom=Side(style='thin', color="DDDDDD"))

            ws["D3"] = "COMITENTE"
            ws["D3"].alignment = Alignment(horizontal='right')
            ws["D3"].font = Font(bold=True, color="666666", size=10)
            ws["E3"] = clean_for_excel(titular)
            ws["E3"].font = Font(bold=True, size=11)
            ws["E3"].alignment = Alignment(horizontal='center')
            ws.merge_cells("E3:G3")
            for c in ["E", "F", "G"]: ws[f"{c}3"].border = Border(bottom=Side(style='thin', color="DDDDDD"))

            ws["D4"] = "PERÍODO"
            ws["D4"].alignment = Alignment(horizontal='right')
            ws["D4"].font = Font(bold=True, color="666666", size=10)
            ws["E4"] = clean_for_excel(periodo)
            ws["E4"].font = Font(bold=True, size=11)
            ws["E4"].alignment = Alignment(horizontal='center')
            ws.merge_cells("E4:G4")
            for c in ["E", "F", "G"]: ws[f"{c}4"].border = Border(bottom=Side(style='thin', color="DDDDDD"))

            headers = ["Fecha", "Tipo de Operación", "Comprobante", "Especie", "Cantidad", "Precio", "Importe"]
            cols = ["A", "B", "C", "D", "E", "F", "G"]
            f_header = 6

            for col, h in zip(cols, headers):
                c = ws[f"{col}{f_header}"]
                c.value = h
                c.fill = PatternFill(start_color=color_header, end_color=color_header, fill_type="solid")
                c.font = Font(bold=True, color="FFFFFF")
                c.alignment = Alignment(horizontal='center')
                c.border = thin_border

            row = f_header + 1
            start_row = row
            if not datos:
                ws.merge_cells(f"A{row}:G{row}")
                ws[f"A{row}"] = "SIN MOVIMIENTOS"
                ws[f"A{row}"].alignment = Alignment(horizontal='center')
                ws[f"A{row}"].font = Font(italic=True, color="666666")
                ws[f"A{row}"].border = thin_border
                row += 1
            else:
                for r in datos:
                    es_credito = r["Importe"] > 0
                    fill = fill_row_cred if es_credito else fill_row_deb
                    font_importe = font_cred if es_credito else font_deb

                    valores = [r["Fecha"], r["Tipo"], r["Comprobante"], r["Especie"], r["Cantidad"], r["Precio"], r["Importe"]]
                    for col, val in zip(cols, valores):
                        cell = ws[f"{col}{row}"]
                        cell.value = val
                        cell.fill = fill
                        cell.border = thin_border
                    ws[f"A{row}"].alignment = Alignment(horizontal='center')
                    ws[f"C{row}"].alignment = Alignment(horizontal='center')
                    ws[f"E{row}"].number_format = '#,##0.00'
                    ws[f"F{row}"].number_format = '#,##0.00'
                    ws[f"G{row}"].number_format = formato_moneda
                    ws[f"G{row}"].font = font_importe
                    row += 1

                ws.merge_cells(f"A{row}:F{row}")
                ws[f"A{row}"] = "TOTAL NETO DEL PERÍODO"
                ws[f"A{row}"].font = Font(bold=True)
                ws[f"A{row}"].alignment = Alignment(horizontal='right')
                ws[f"G{row}"] = f"=SUM(G{start_row}:G{row-1})"
                ws[f"G{row}"].font = Font(bold=True)
                ws[f"G{row}"].number_format = formato_moneda
                for col in cols:
                    ws[f"{col}{row}"].border = thin_border
                row += 1

            ws.column_dimensions["A"].width = 12
            ws.column_dimensions["B"].width = 30
            ws.column_dimensions["C"].width = 14
            ws.column_dimensions["D"].width = 45
            ws.column_dimensions["E"].width = 16
            ws.column_dimensions["F"].width = 14
            ws.column_dimensions["G"].width = 18

        crear_hoja("Pesos", transactions_ars, saldo_inicial_ars, saldo_final_ars, '"$ "#,##0.00')

        if transactions_usd or saldo_inicial_usd != 0 or saldo_final_usd != 0:
            crear_hoja("Dolares", transactions_usd, saldo_inicial_usd, saldo_final_usd, '"U$S "#,##0.00')

        wb.save(output)
        output.seek(0)
        return output.getvalue()

    except Exception as e:
        import traceback
        st.error(f"Error al procesar: {e}")
        print(traceback.format_exc())
        return None
