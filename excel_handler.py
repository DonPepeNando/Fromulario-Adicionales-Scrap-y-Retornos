import openpyxl
import os
import math

ACTIONS_MAP = {
    "Return Ok": "Return Ok",
    "Retorno": "Return Ok",
    "Ok": "Return Ok",
    "Scrap": "Scrap",
    "Adicional": "Adittional",
    "Adittional": "Adittional",
    "Consumibles": "Consumables",
    "Consumables": "Consumables"
}

def format_action_string(selected_action):
    """
    Construye la cadena de texto para la celda E1.
    """
    normalized_action = ACTIONS_MAP.get(selected_action, selected_action)
    options = [
        ("Return Ok", "Return Ok"),
        ("Scrap", "Scrap"),
        ("Adittional", "Adittional"),
        ("Consumables", "Consumables")
    ]
    
    parts = []
    for opt_key, opt_label in options:
        mark = "☑" if opt_key == normalized_action else "  "
        parts.append(f"{mark}{opt_label}")
    
    return "  ".join(parts)

def _clean(val):
    return str(val).replace('\t', ' ') if val else ""

def generate_excel_form(data, template_path, output_path, sheet_name="SCRAP"):
    """
    Carga la plantilla Excel (.xlsm / .xlsx), inyecta los datos del formulario y guarda el resultado.
    Si contiene más de 12 materiales (hasta 84), llena la hoja base primero y luego
    genera copias exactas idénticas (ej. SCRAP2, SCRAP3 o Hoja12, Hoja13) preservando el 100% de formatos, 
    imágenes, logotipos y cuadros de texto (incluyendo el cuadro sobre celda I1).
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"No se encontró la plantilla Excel en: {template_path}")
    
    abs_template = os.path.abspath(template_path)
    abs_output = os.path.abspath(output_path)
    
    out_dir = os.path.dirname(abs_output)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    materiales = data.get("materiales", [])
    total_items = len(materiales)
    num_sheets = max(1, math.ceil(total_items / 12)) if total_items > 0 else 1
    accion_str = format_action_string(data.get("accion", ""))

    # 1. Intentar generación mediante Excel COM nativo para conservar 100% cuadros de texto sobre celda I1, formas e imágenes
    try:
        import win32com.client
        # Use DispatchEx to ensure a completely fresh instance of Excel is spawned.
        # This prevents hanging invisible Excel processes from blocking the generation.
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        
        wb = excel.Workbooks.Open(abs_template)
        
        try:
            base_ws = wb.Sheets(sheet_name)
        except Exception:
            base_ws = wb.ActiveSheet
            
        base_name = base_ws.Name
        start_row = 8
        max_rows = 12

        base_ws.Range(f"B{start_row}:D{start_row + max_rows - 1}").WrapText = True
        base_ws.Range(f"B{start_row}:D{start_row + max_rows - 1}").ShrinkToFit = True
        base_ws.Range(f"D{start_row}:D{start_row + max_rows - 1}").NumberFormat = "@"

        # Llenar la Hoja 1 (base) con los datos del encabezado y primeros 12 materiales
        base_ws.Range("B2").Value = _clean(data.get("orden", ""))
        base_ws.Range("B3").Value = _clean(data.get("proyecto", ""))
        base_ws.Range("B4").Value = _clean(data.get("centro_costos", ""))
        base_ws.Range("B5").Value = _clean(data.get("area", ""))
        base_ws.Range("B6").Value = _clean(data.get("nombre", ""))
        base_ws.Range("I5").Value = _clean(data.get("folio", ""))
        base_ws.Range("I6").Value = _clean(data.get("fecha", ""))
        base_ws.Range("E1").Value = _clean(accion_str)

        for i in range(max_rows):
            row_idx = start_row + i
            if i < total_items:
                mat_item = materiales[i]
                base_ws.Range(f"A{row_idx}").Value = i + 1
                base_ws.Range(f"B{row_idx}").Value = _clean(mat_item.get("material", ""))
                base_ws.Range(f"D{row_idx}").Value = _clean(mat_item.get("batch", ""))
                base_ws.Range(f"E{row_idx}").Value = _clean(mat_item.get("cantidad", ""))
                base_ws.Range(f"G{row_idx}").Value = _clean(mat_item.get("descripcion", ""))
            else:
                base_ws.Range(f"B{row_idx}").Value = ""
                base_ws.Range(f"D{row_idx}").Value = ""
                base_ws.Range(f"E{row_idx}").Value = ""
                base_ws.Range(f"G{row_idx}").Value = ""

        # Si hay más de 12 materiales, duplicar la hoja base TERMINADA para conservar todo el formato
        # Usamos base_ws.Copy(base_ws) para evitar bugs de kwargs en win32com que destruyen celdas combinadas.
        for p in range(1, num_sheets):
            base_ws.Copy(base_ws)
            new_ws = wb.Worksheets(base_ws.Index - 1)
            target_name = f"{base_name}{p + 1}"
            new_ws.Name = target_name

            # Reasegurar encabezados idénticos en la copia exacta
            new_ws.Range("B2").Value = _clean(data.get("orden", ""))
            new_ws.Range("B3").Value = _clean(data.get("proyecto", ""))
            new_ws.Range("B4").Value = _clean(data.get("centro_costos", ""))
            new_ws.Range("B5").Value = _clean(data.get("area", ""))
            new_ws.Range("B6").Value = _clean(data.get("nombre", ""))
            new_ws.Range("I5").Value = _clean(data.get("folio", ""))
            new_ws.Range("I6").Value = _clean(data.get("fecha", ""))
            new_ws.Range("E1").Value = _clean(accion_str)

            start_mat_idx = p * 12
            for i in range(max_rows):
                mat_pos = start_mat_idx + i
                row_idx = start_row + i
                if mat_pos < total_items:
                    mat_item = materiales[mat_pos]
                    new_ws.Range(f"A{row_idx}").Value = mat_pos + 1
                    new_ws.Range(f"B{row_idx}").Value = _clean(mat_item.get("material", ""))
                    new_ws.Range(f"D{row_idx}").Value = _clean(mat_item.get("batch", ""))
                    new_ws.Range(f"E{row_idx}").Value = _clean(mat_item.get("cantidad", ""))
                    new_ws.Range(f"G{row_idx}").Value = _clean(mat_item.get("descripcion", ""))
                else:
                    new_ws.Range(f"B{row_idx}").Value = ""
                    new_ws.Range(f"D{row_idx}").Value = ""
                    new_ws.Range(f"E{row_idx}").Value = ""
                    new_ws.Range(f"G{row_idx}").Value = ""

        # Reordenar las hojas generadas
        if num_sheets > 1:
            base_ws.Move(wb.Worksheets(1))
            for p in range(1, num_sheets):
                target_name = f"{base_name}{p + 1}"
                try:
                    # Move(Before, After) -> Move(None, After)
                    wb.Worksheets(target_name).Move(None, wb.Worksheets(p))
                except:
                    pass

        file_format = 52 if abs_output.lower().endswith(".xlsm") else 51
        wb.SaveAs(abs_output, FileFormat=file_format)
        wb.Close(False)
        excel.Quit()
        return abs_output, "COM"
    except Exception as err:
        import traceback
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        log_path = os.path.join(desktop, "COM_ERROR_LOG.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("EXCEL COM FAILURE TRACEBACK:\n")
            traceback.print_exc(file=f)
            f.write(f"Exception: {err}\n\n")
        
        print("EXCEL COM FAILURE TRACEBACK:")
        traceback.print_exc()
        print(f"Aviso: Generación Excel COM secundaria activando fallback openpyxl: {err}")

    # Fallback con openpyxl si Excel COM no estuviera disponible
    import openpyxl
    is_xlsm = abs_template.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(abs_template, keep_vba=is_xlsm)
    
    if sheet_name in wb.sheetnames:
        base_ws = wb[sheet_name]
    else:
        base_ws = wb.active
        
    base_name = base_ws.title
    start_row = 8
    max_rows = 12

    from openpyxl.styles import Alignment
    for r_idx in range(start_row, start_row + max_rows):
        base_ws[f"B{r_idx}"].alignment = Alignment(wrap_text=True, shrink_to_fit=True)
        base_ws[f"C{r_idx}"].alignment = Alignment(wrap_text=True, shrink_to_fit=True)
        base_ws[f"D{r_idx}"].alignment = Alignment(wrap_text=True, shrink_to_fit=True)
        base_ws[f"D{r_idx}"].number_format = '@'

    base_ws['B2'] = _clean(data.get("orden", ""))
    base_ws['B3'] = _clean(data.get("proyecto", ""))
    base_ws['B4'] = _clean(data.get("centro_costos", ""))
    base_ws['B5'] = _clean(data.get("area", ""))
    base_ws['B6'] = _clean(data.get("nombre", ""))
    base_ws['I5'] = _clean(data.get("folio", ""))
    base_ws['I6'] = _clean(data.get("fecha", ""))
    base_ws['E1'] = _clean(accion_str)

    for i in range(max_rows):
        row_idx = start_row + i
        if i < total_items:
            mat_item = materiales[i]
            base_ws[f'A{row_idx}'] = i + 1
            base_ws[f'B{row_idx}'] = _clean(mat_item.get("material", ""))
            base_ws[f'D{row_idx}'] = _clean(mat_item.get("batch", ""))
            base_ws[f'E{row_idx}'] = _clean(mat_item.get("cantidad", ""))
            base_ws[f'G{row_idx}'] = _clean(mat_item.get("descripcion", ""))
        else:
            base_ws[f'B{row_idx}'] = None
            base_ws[f'D{row_idx}'] = None
            base_ws[f'E{row_idx}'] = None
            base_ws[f'G{row_idx}'] = None

    last_ws_pyxl = base_ws
    for p in range(1, num_sheets):
        target_name = f"{base_name}{p + 1}"
        new_ws = wb.copy_worksheet(last_ws_pyxl)
        new_ws.title = target_name

        new_ws['B2'] = _clean(data.get("orden", ""))
        new_ws['B3'] = _clean(data.get("proyecto", ""))
        new_ws['B4'] = _clean(data.get("centro_costos", ""))
        new_ws['B5'] = _clean(data.get("area", ""))
        new_ws['B6'] = _clean(data.get("nombre", ""))
        new_ws['I5'] = _clean(data.get("folio", ""))
        new_ws['I6'] = _clean(data.get("fecha", ""))
        new_ws['E1'] = _clean(accion_str)

        start_mat_idx = p * 12
        for i in range(max_rows):
            mat_pos = start_mat_idx + i
            row_idx = start_row + i
            if mat_pos < total_items:
                mat_item = materiales[mat_pos]
                new_ws[f'A{row_idx}'] = mat_pos + 1
                new_ws[f'B{row_idx}'] = _clean(mat_item.get("material", ""))
                new_ws[f'D{row_idx}'] = _clean(mat_item.get("batch", ""))
                new_ws[f'E{row_idx}'] = _clean(mat_item.get("cantidad", ""))
                new_ws[f'G{row_idx}'] = _clean(mat_item.get("descripcion", ""))
            else:
                new_ws[f'B{row_idx}'] = None
                new_ws[f'D{row_idx}'] = None
                new_ws[f'E{row_idx}'] = None
                new_ws[f'G{row_idx}'] = None
        last_ws_pyxl = new_ws

    # Reordenar las pestañas para que la serie de hojas generadas (SCRAP, SCRAP2, SCRAP3) queden al inicio del libro
    main_names = [f"{base_name}{p+1}" if p > 0 else base_name for p in range(num_sheets)]
    main_sheets = [s for s in wb.worksheets if s.title in main_names]
    other_sheets = [s for s in wb.worksheets if s.title not in main_names]
    wb._sheets = main_sheets + other_sheets

    wb.save(abs_output)
    return abs_output, "openpyxl"


def generate_pdf_form(data, template_path, output_path, sheet_name="SCRAP"):
    """
    Carga la plantilla Excel (.xlsm / .xlsx), inyecta los datos del formulario y exporta a PDF.
    Usa estrictamente Excel COM. Si falla, lanza una excepción.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"No se encontró la plantilla Excel en: {template_path}")
    
    abs_template = os.path.abspath(template_path)
    abs_output = os.path.abspath(output_path)
    
    out_dir = os.path.dirname(abs_output)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    materiales = data.get("materiales", [])
    total_items = len(materiales)
    import math
    num_sheets = max(1, math.ceil(total_items / 12)) if total_items > 0 else 1
    accion_str = format_action_string(data.get("accion", ""))

    import win32com.client
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = None
    try:
        wb = excel.Workbooks.Open(abs_template)
        
        try:
            base_ws = wb.Sheets(sheet_name)
        except Exception:
            base_ws = wb.ActiveSheet
            
        base_name = base_ws.Name
        start_row = 8
        max_rows = 12

        base_ws.Range(f"B{start_row}:D{start_row + max_rows - 1}").WrapText = True
        base_ws.Range(f"B{start_row}:D{start_row + max_rows - 1}").ShrinkToFit = True
        base_ws.Range(f"D{start_row}:D{start_row + max_rows - 1}").NumberFormat = "@"

        base_ws.Range("B2").Value = _clean(data.get("orden", ""))
        base_ws.Range("B3").Value = _clean(data.get("proyecto", ""))
        base_ws.Range("B4").Value = _clean(data.get("centro_costos", ""))
        base_ws.Range("B5").Value = _clean(data.get("area", ""))
        base_ws.Range("B6").Value = _clean(data.get("nombre", ""))
        base_ws.Range("I5").Value = _clean(data.get("folio", ""))
        base_ws.Range("I6").Value = _clean(data.get("fecha", ""))
        base_ws.Range("E1").Value = _clean(accion_str)

        for i in range(max_rows):
            row_idx = start_row + i
            if i < total_items:
                mat_item = materiales[i]
                base_ws.Range(f"A{row_idx}").Value = i + 1
                base_ws.Range(f"B{row_idx}").Value = _clean(mat_item.get("material", ""))
                base_ws.Range(f"D{row_idx}").Value = _clean(mat_item.get("batch", ""))
                base_ws.Range(f"E{row_idx}").Value = _clean(mat_item.get("cantidad", ""))
                base_ws.Range(f"G{row_idx}").Value = _clean(mat_item.get("descripcion", ""))
            else:
                base_ws.Range(f"B{row_idx}").Value = ""
                base_ws.Range(f"D{row_idx}").Value = ""
                base_ws.Range(f"E{row_idx}").Value = ""
                base_ws.Range(f"G{row_idx}").Value = ""

        for p in range(1, num_sheets):
            base_ws.Copy(base_ws)
            new_ws = wb.Worksheets(base_ws.Index - 1)
            target_name = f"{base_name}{p + 1}"
            new_ws.Name = target_name

            new_ws.Range("B2").Value = _clean(data.get("orden", ""))
            new_ws.Range("B3").Value = _clean(data.get("proyecto", ""))
            new_ws.Range("B4").Value = _clean(data.get("centro_costos", ""))
            new_ws.Range("B5").Value = _clean(data.get("area", ""))
            new_ws.Range("B6").Value = _clean(data.get("nombre", ""))
            new_ws.Range("I5").Value = _clean(data.get("folio", ""))
            new_ws.Range("I6").Value = _clean(data.get("fecha", ""))
            new_ws.Range("E1").Value = _clean(accion_str)

            start_mat_idx = p * 12
            for i in range(max_rows):
                mat_pos = start_mat_idx + i
                row_idx = start_row + i
                if mat_pos < total_items:
                    mat_item = materiales[mat_pos]
                    new_ws.Range(f"A{row_idx}").Value = mat_pos + 1
                    new_ws.Range(f"B{row_idx}").Value = _clean(mat_item.get("material", ""))
                    new_ws.Range(f"D{row_idx}").Value = _clean(mat_item.get("batch", ""))
                    new_ws.Range(f"E{row_idx}").Value = _clean(mat_item.get("cantidad", ""))
                    new_ws.Range(f"G{row_idx}").Value = _clean(mat_item.get("descripcion", ""))
                else:
                    new_ws.Range(f"B{row_idx}").Value = ""
                    new_ws.Range(f"D{row_idx}").Value = ""
                    new_ws.Range(f"E{row_idx}").Value = ""
                    new_ws.Range(f"G{row_idx}").Value = ""

        if num_sheets > 1:
            base_ws.Move(wb.Worksheets(1))
            for p in range(1, num_sheets):
                target_name = f"{base_name}{p + 1}"
                try:
                    wb.Worksheets(target_name).Move(None, wb.Worksheets(p))
                except:
                    pass
        
        sheet_names_to_export = [f"{base_name}{p+1}" if p > 0 else base_name for p in range(num_sheets)]
        wb.Worksheets(sheet_names_to_export).Select()

        wb.ActiveSheet.ExportAsFixedFormat(0, abs_output)
        
        return abs_output, "COM_PDF"
    except Exception as err:
        raise RuntimeError(f"Error generando PDF usando Excel COM: {err}")
    finally:
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
