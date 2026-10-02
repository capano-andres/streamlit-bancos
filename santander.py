import streamlit as st
import io
import PyPDF2
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule

# Regex para caracteres ilegales en Excel
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

def clean_for_excel(text):
    """Elimina caracteres ilegales para Excel y espacios extra"""
    if not text: return ""
    text = str(text)
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    return text.strip()

def procesar_santander_rio(archivo_pdf):
    """Procesa archivos PDF de Santander Rio con Estilo Dashboard Multi-Moneda"""
    st.info("Procesando archivo de Santander Rio...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)
        
        # Abrir el PDF usando PyPDF2
        reader = PyPDF2.PdfReader(io.BytesIO(archivo_pdf.read()))
        texto_completo = "".join(page.extract_text() + "\n" for page in reader.pages)
        
        lineas_raw = texto_completo.splitlines()

        # 1. Metadatos (Titular, Periodo)
        titular_global = "Sin Especificar"
        periodo_global = "Sin Especificar"
        
        # Titular: Linea anterior a "CUIT:" o "CUIL:"
        for i, l in enumerate(lineas_raw[:20]):
            if "CUIT:" in l or "CUIL:" in l:
                if i > 0:
                    titular_global = lineas_raw[i-1].strip()
                break
        
        # Periodo: "Desde: 27/01/23" ... "Hasta: 02/03/23"
        f_desde = None
        f_hasta = None
        for l in lineas_raw[:30]:
            match_d = re.search(r"Desde:\s*(\d{2}/\d{2}/\d{2,4})", l)
            if match_d: f_desde = match_d.group(1)
            match_h = re.search(r"Hasta:\s*(\d{2}/\d{2}/\d{2,4})", l)
            if match_h: f_hasta = match_h.group(1)
        
        if f_desde and f_hasta:
            periodo_global = f"Del {f_desde} al {f_hasta}"

        # --- DELIMITAR SECCIONES ---
        idx_pesos = None
        idx_dolares = None
        idx_fin_pesos = None # Fin de pesos puede ser inicio dolares o fin documento
        idx_fin_dolares = None

        for i, l in enumerate(lineas_raw):
            if "Movimientos en pesos" in l and idx_pesos is None:
                idx_pesos = i
            if "Movimientos en dólares" in l and idx_dolares is None:
                idx_dolares = i
            if ("Así usaste tu dinero este mes" in l or "Detalle impositivo" in l) and idx_fin_dolares is None and idx_dolares is not None:
                idx_fin_dolares = i
            # Si no hay dolares, el fin de pesos puede ser "Así usaste..."
            if ("Así usaste tu dinero este mes" in l or "Detalle impositivo" in l) and idx_fin_pesos is None and idx_pesos is not None and idx_dolares is None:
                idx_fin_pesos = i

        # Ajustar rangos
        lineas_pesos = []
        lineas_dolares = []

        if idx_pesos is not None:
            # Fin de pesos es idx_dolares si existe, sino idx_fin_pesos, sino fin archivo
            end_p = idx_dolares if idx_dolares is not None else (idx_fin_pesos if idx_fin_pesos is not None else len(lineas_raw))
            lineas_pesos = lineas_raw[idx_pesos+1 : end_p]
        
        if idx_dolares is not None:
            end_d = idx_fin_dolares if idx_fin_dolares is not None else len(lineas_raw)
            lineas_dolares = lineas_raw[idx_dolares+1 : end_d]

        # --- FUNCION EXTRACTION (REUTILIZADA) ---
        def extraer_datos_seccion(lineas):
            movimientos_text = []
            linea_actual = ""
            saldo_ini = 0.0
            saldo_fin = 0.0
            
            # Pre-procesado para unir líneas
            for l in lineas:
                # Filtrar encabezados repetidos de página (ej: "2 -  11")
                if re.match(r'^\s*\d+\s*-\s+\d+\s*$', l.strip()):
                    continue
                # Filtrar encabezados de tabla repetidos
                if "Cuenta Corriente" in l and ("CBU:" in l or "Nº" in l):
                    continue
                if "FechaComprobante" in l or "FechaComprobanteMovimiento" in l:
                    continue

                # Extraer saldos si aparecen en la sección
                if "Saldo Inicial" in l:
                    matches = re.findall(r"(-?)\$\s?([\d\.]+,\d{2})|(-?)U\$S\s?([\d\.]+,\d{2})", l)
                    # matches devuelve tuplas con grupos vacios, hay que filtrar
                    for m in matches:
                        # m = ('-', '1.200,00', '', '') para pesos
                        # m = ('', '', '-', '100,00') para dolares
                        val_str = m[1] if m[1] else m[3]
                        sign_str = m[0] if m[1] else m[2]
                        if val_str:
                            try:
                                num = float(val_str.replace(".", "").replace(",", "."))
                                if sign_str == "-": num *= -1
                                saldo_ini = num
                            except: pass

                if "Saldo total" in l:
                    matches = re.findall(r"(-?)\$\s?([\d\.]+,\d{2})|(-?)U\$S\s?([\d\.]+,\d{2})", l)
                    for m in matches:
                        val_str = m[1] if m[1] else m[3]
                        sign_str = m[0] if m[1] else m[2]
                        if val_str:
                            try: 
                                num = float(val_str.replace(".", "").replace(",", "."))
                                if sign_str == "-": num *= -1
                                saldo_fin = num
                            except: pass
                    continue # NO unir la linea de Saldo Total al movimiento anterior

                # Unir lineas de movimientos
                if re.match(r"\d{2}/\d{2}/\d{2}", l):
                    if linea_actual: movimientos_text.append(linea_actual.strip())
                    linea_actual = l
                else:
                    linea_actual += " " + l
            if linea_actual: movimientos_text.append(linea_actual.strip())

            # Parsear Movimientos usando diferencia de saldos
            parsed_data = []
            saldo_anterior = saldo_ini  # Arrancar con el saldo inicial
            
            for mov in movimientos_text:
                if "Movimientos en" in mov: continue 
                if "Saldo Inicial" in mov: continue # Filtrar Saldo Inicial siempre
                
                fecha = mov[:8]
                resto = mov[8:]
                
                # Limpieza basica de moneda para facilitar regex unico
                mov_clean = mov.replace("U$S", "$").replace("U$s", "$")

                # Buscamos todos los montos monetarios
                montos = re.findall(r"([+-]?\$\s*[\d\.,]+)", mov_clean)
                
                importe = 0.0
                desc = ""
                
                if len(montos) >= 2:
                    # El ÚLTIMO monto es el saldo acumulado (balance running)
                    str_saldo = montos[-1]
                    clean_saldo = str_saldo.replace("$", "").replace("+", "").replace("-", "").strip().replace(".", "").replace(",", ".")
                    signo_saldo = -1 if "-" in str_saldo else 1
                    try:
                        saldo_actual = float(clean_saldo) * signo_saldo
                    except:
                        saldo_actual = saldo_anterior
                    
                    # Importe = diferencia de saldos (positivo = crédito, negativo = débito)
                    importe = round(saldo_actual - saldo_anterior, 2)
                    saldo_anterior = saldo_actual
                    
                    # Descripción: todo lo que hay antes del penúltimo monto
                    str_imp = montos[-2]
                    idx_imp = mov_clean.rfind(str_imp) 
                    if idx_imp != -1:
                        desc = mov[:idx_imp] # Incluye fecha en los primeros 8 chars
                        desc = desc[8:].strip() # Quitar fecha
                    else:
                        desc = resto

                    # El comprobante son los numeros pegados al inicio (ej: 77367269Transferencia)
                    m_comp = re.match(r'^(\d+)', desc)
                    comprobante = m_comp.group(1) if m_comp else ""
                    desc = re.sub(r'^\d+', '', desc).strip()

                    parsed_data.append((fecha, comprobante, clean_for_excel(desc), importe))

                elif len(montos) == 1:
                    # Solo hay un monto, puede ser saldo inicial o algo raro.
                    if "Saldo Inicial" in mov: continue 
                    # Si es un movimiento sin saldo acumulado visible? Raro en este banco.
                    pass
            
            return parsed_data, saldo_ini, saldo_fin

        # Procesar
        datos_pesos, saldo_ini_pesos, saldo_fin_pesos = extraer_datos_seccion(lineas_pesos)
        datos_dolares, saldo_ini_dolares, saldo_fin_dolares = extraer_datos_seccion(lineas_dolares)
        

        # --- GENERACIÓN EXCEL MULTI-HOJA ---
        output = io.BytesIO()
        wb = Workbook()
        # Eliminar hoja default
        wb.remove(wb.active)
        
        # Estilos
        color_bg_main = "EC0000" 
        color_txt_main = "FFFFFF"
        thin_border = Border(left=Side(style='thin', color="A6A6A6"), right=Side(style='thin', color="A6A6A6"), top=Side(style='thin', color="A6A6A6"), bottom=Side(style='thin', color="A6A6A6"))
        fill_head_deb = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        fill_col_deb = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
        fill_row_deb = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
        fill_head_cred = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        fill_col_cred = PatternFill(start_color="EBF1DE", end_color="EBF1DE", fill_type="solid")
        fill_row_cred = PatternFill(start_color="F2F9F1", end_color="F2F9F1", fill_type="solid")
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        red_font = Font(color='9C0006', bold=True)

        def crear_hoja_dashboard(wb, nombre_hoja, datos, s_ini, s_fin, formato_moneda='"$ "#,##0.00'):
            ws = wb.create_sheet(title=nombre_hoja)
            ws.sheet_view.showGridLines = False
            
            df = pd.DataFrame(datos, columns=["Fecha", "Comprobante", "Descripcion", "Importe"])

            creditos = df[df["Importe"] > 0].copy()
            debitos = df[df["Importe"] < 0].copy()
            debitos["Importe"] = debitos["Importe"].abs() # Positivo para mostrar

            if df.empty:
                creditos = pd.DataFrame(columns=["Fecha", "Comprobante", "Descripcion", "Importe"])
                debitos = pd.DataFrame(columns=["Fecha", "Comprobante", "Descripcion", "Importe"])

            # Header
            ws.merge_cells("A1:G1")
            tit = ws["A1"]
            tit.value = f"REPORTE SANTANDER ({nombre_hoja}) - {clean_for_excel(titular_global)}"
            tit.font = Font(size=14, bold=True, color=color_txt_main)
            tit.fill = PatternFill(start_color=color_bg_main, end_color=color_bg_main, fill_type="solid")
            tit.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 25

            # Metadata
            ws["A3"] = "SALDO INICIAL"
            ws["A3"].font = Font(bold=True, size=10, color="666666")
            ws["B3"] = s_ini
            ws["B3"].number_format = formato_moneda
            ws["B3"].font = Font(bold=True, size=11)
            ws["B3"].border = Border(bottom=Side(style='thin', color="DDDDDD"))

            ws["A4"] = "SALDO FINAL"
            ws["A4"].font = Font(bold=True, size=10, color="666666")
            ws["B4"] = s_fin
            ws["B4"].number_format = formato_moneda
            ws["B4"].font = Font(bold=True, size=11)
            ws["B4"].border = Border(bottom=Side(style='thin', color="DDDDDD"))
            
            ws["D3"] = "TITULAR"; 
            ws.merge_cells("E3:G3"); ws["E3"] = clean_for_excel(titular_global)
            ws["E3"].alignment = Alignment(horizontal='center')

            ws["D4"] = "PERÍODO"; 
            ws.merge_cells("E4:G4"); ws["E4"] = clean_for_excel(periodo_global)
            ws["E4"].alignment = Alignment(horizontal='center')
            
            ws["D6"] = "CONTROL DE SALDOS"
            
            # Control Formula Placeholder
            ws["D7"] = 0
            ws["D7"].font = Font(bold=True, size=12); ws["D7"].border = thin_border
            ws.conditional_formatting.add('D7', CellIsRule(operator='notEqual', formula=['0'], stopIfTrue=True, fill=red_fill, font=red_font))

            # Tablas
            f_header = 10
            # Creditos
            ws.merge_cells(f"A{f_header}:D{f_header}"); ws[f"A{f_header}"] = "CRÉDITOS"
            ws[f"A{f_header}"].fill = fill_head_cred; ws[f"A{f_header}"].font = Font(bold=True, color="FFFFFF")
            ws[f"A{f_header}"].alignment = Alignment(horizontal="center", vertical="center")
            # Debitos
            ws.merge_cells(f"F{f_header}:I{f_header}"); ws[f"F{f_header}"] = "DÉBITOS"
            ws[f"F{f_header}"].fill = fill_head_deb; ws[f"F{f_header}"].font = Font(bold=True, color="FFFFFF")
            ws[f"F{f_header}"].alignment = Alignment(horizontal="center", vertical="center")

            # Subheaders
            cols_cred = ["A","B","C","D"]
            cols_deb = ["F","G","H","I"]
            headers = ["Fecha","Comprobante","Descripción","Importe"]
            for col, txt in zip(cols_cred + cols_deb, headers + headers):
                ws[f"{col}{f_header+1}"] = txt
                ws[f"{col}{f_header+1}"].border = thin_border
                ws[f"{col}{f_header+1}"].alignment = Alignment(horizontal='center')
                if col in cols_cred: ws[f"{col}{f_header+1}"].fill = fill_col_cred
                else: ws[f"{col}{f_header+1}"].fill = fill_col_deb

            # Llenar Creditos
            row = f_header + 2
            start_cred = row
            if creditos.empty:
                ws[f"A{row}"] = "SIN MOVIMIENTOS"; ws.merge_cells(f"A{row}:D{row}")
                ws[f"A{row}"].alignment = Alignment(horizontal='center'); ws[f"A{row}"].font = Font(italic=True, color="666666")
                row += 1
            else:
                for _, r in creditos.iterrows():
                    ws[f"A{row}"] = r["Fecha"]; ws[f"B{row}"] = r["Comprobante"]; ws[f"C{row}"] = r["Descripcion"]; ws[f"D{row}"] = r["Importe"]
                    ws[f"D{row}"].number_format = formato_moneda
                    for c in cols_cred: ws[f"{c}{row}"].border = thin_border; ws[f"{c}{row}"].fill = fill_row_cred
                    row += 1

            total_cred_row = row
            ws.merge_cells(f"A{total_cred_row}:C{total_cred_row}")
            ws[f"A{total_cred_row}"] = "TOTAL CRÉDITOS"
            ws[f"A{total_cred_row}"].font = Font(bold=True); ws[f"A{total_cred_row}"].alignment = Alignment(horizontal='right')
            ws[f"D{total_cred_row}"] = f"=SUM(D{start_cred}:D{total_cred_row-1})"
            ws[f"D{total_cred_row}"].number_format = formato_moneda; ws[f"D{total_cred_row}"].font = Font(bold=True)
            for c in cols_cred: ws[f"{c}{total_cred_row}"].border = thin_border

            # Llenar Debitos
            row = f_header + 2
            start_deb = row
            if debitos.empty:
                ws[f"F{row}"] = "SIN MOVIMIENTOS"; ws.merge_cells(f"F{row}:I{row}")
                ws[f"F{row}"].alignment = Alignment(horizontal='center'); ws[f"F{row}"].font = Font(italic=True, color="666666")
                row += 1
            else:
                for _, r in debitos.iterrows():
                    ws[f"F{row}"] = r["Fecha"]; ws[f"G{row}"] = r["Comprobante"]; ws[f"H{row}"] = r["Descripcion"]; ws[f"I{row}"] = r["Importe"]
                    ws[f"I{row}"].number_format = formato_moneda
                    for c in cols_deb: ws[f"{c}{row}"].border = thin_border; ws[f"{c}{row}"].fill = fill_row_deb
                    row += 1

            total_deb_row = row
            ws.merge_cells(f"F{total_deb_row}:H{total_deb_row}")
            ws[f"F{total_deb_row}"] = "TOTAL DÉBITOS"
            ws[f"F{total_deb_row}"].font = Font(bold=True); ws[f"F{total_deb_row}"].alignment = Alignment(horizontal='right')
            ws[f"I{total_deb_row}"] = f"=SUM(I{start_deb}:I{total_deb_row-1})"
            ws[f"I{total_deb_row}"].number_format = formato_moneda; ws[f"I{total_deb_row}"].font = Font(bold=True)
            for c in cols_deb: ws[f"{c}{total_deb_row}"].border = thin_border

            # Update Control Formula final
            ws["D7"] = f"=ROUND(B3+D{total_cred_row}-I{total_deb_row}-B4, 2)"
            ws["D7"].number_format = formato_moneda

            # Anchos
            ws.column_dimensions["B"].width = 14; ws.column_dimensions["G"].width = 14
            ws.column_dimensions["C"].width = 40; ws.column_dimensions["H"].width = 40
            ws.column_dimensions["D"].width = 18; ws.column_dimensions["I"].width = 18

        # Crear hoja Pesos
        crear_hoja_dashboard(wb, "Pesos", datos_pesos, saldo_ini_pesos, saldo_fin_pesos, formato_moneda='"$ "#,##0.00')
        
        # Crear hoja Dolares
        if datos_dolares or saldo_ini_dolares != 0 or saldo_fin_dolares != 0:
            crear_hoja_dashboard(wb, "Dolares", datos_dolares, saldo_ini_dolares, saldo_fin_dolares, formato_moneda='"U$S "#,##0.00')

        wb.save(output)
        output.seek(0)
        return output.getvalue()

    except Exception as e:
        import traceback
        st.error(f"Error al procesar el archivo: {str(e)}")
        print(traceback.format_exc())
        return None
