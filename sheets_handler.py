import os
from datetime import datetime

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

HEADERS = [
    "Folio", "Orden", "Centro de Costos", "Área", 
    "Tipo de Formulario", "Material", "Cantidad", "Fecha", "Nombre Solicitante"
]

def sync_to_google_sheets(data, credentials_path="credentials.json", sheet_url_or_name=""):
    """
    Sincroniza los datos del formulario a una hoja de Google Sheets.
    Retorna tuple: (éxito: bool, mensaje: str)
    """
    if not GSPREAD_AVAILABLE:
        return False, "Las librerías gspread o google-auth no están disponibles."

    if not credentials_path or not os.path.exists(credentials_path):
        return False, f"Archivo de credenciales no encontrado: '{credentials_path}'."

    if not sheet_url_or_name:
        return False, "No se ha especificado el nombre o la URL de la hoja de Google Sheets en la configuración."

    try:
        creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # Abrir por URL o por nombre de hoja
        if sheet_url_or_name.startswith("http://") or sheet_url_or_name.startswith("https://"):
            spreadsheet = client.open_by_url(sheet_url_or_name)
        else:
            spreadsheet = client.open(sheet_url_or_name)
            
        worksheet = spreadsheet.sheet1  # Primera hoja por defecto
        
        # Asegurar encabezados exactos en A1:I1 (reemplaza cualquier encabezado previo)
        row1 = worksheet.row_values(1)
        if not row1 or row1[:len(HEADERS)] != HEADERS:
            try:
                worksheet.update(range_name="A1:I1", values=[HEADERS])
            except Exception:
                worksheet.update("A1:I1", [HEADERS])

        materiales = data.get("materiales", [])
        
        rows_to_append = []
        if materiales:
            for item in materiales:
                row = [
                    data.get("folio", ""),
                    data.get("orden", ""),
                    data.get("centro_costos", ""),
                    data.get("area", ""),
                    data.get("accion", ""),
                    item.get("material", ""),
                    item.get("cantidad", ""),
                    data.get("fecha", ""),
                    data.get("nombre", "")
                ]
                rows_to_append.append(row)
        else:
            row = [
                data.get("folio", ""),
                data.get("orden", ""),
                data.get("centro_costos", ""),
                data.get("area", ""),
                data.get("accion", ""),
                "",
                "",
                data.get("fecha", ""),
                data.get("nombre", "")
            ]
            rows_to_append.append(row)

        worksheet.append_rows(rows_to_append)
        return True, f"Se registraron exitosamente {len(rows_to_append)} filas en Google Sheets."

    except Exception as e:
        err_detail = str(e)
        if hasattr(e, "__cause__") and e.__cause__:
            err_detail = f"{err_detail} | {e.__cause__}"
        return False, f"Error al conectar con Google Sheets: {err_detail}"
