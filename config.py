import os
import json
import re
import time

LOCK_DIR = 'app_config.lock'

class FileLock:
    def __enter__(self):
        while True:
            try:
                os.mkdir(LOCK_DIR)
                break
            except FileExistsError:
                try:
                    if os.path.exists(LOCK_DIR) and (time.time() - os.path.getctime(LOCK_DIR)) > 10:
                        os.rmdir(LOCK_DIR)
                except Exception:
                    pass
                time.sleep(0.2)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            os.rmdir(LOCK_DIR)
        except Exception:
            pass

CONFIG_FILE = "app_config.json"
DEFAULT_LISTS_FILE = "listas_config.txt"

DEFAULT_CONFIG = {
    "template_path": "FORMATO SCRAP Y ADICIONALES.xlsm",
    "sheet_name": "SCRAP",
    "credentials_path": "credentials.json",
    "google_sheet_url_or_name": "",
    "lists_file_path": DEFAULT_LISTS_FILE,
    "generated_folios": {}  # dict: {folio_key: {"folio": str, "orden": str, "fecha": str, "accion": str, "file": str, "status": "LOCKED"|"RELEASED", "data": dict}}
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                config_data = DEFAULT_CONFIG.copy()
                config_data.update(data)
                return config_data
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()

def save_config(config_data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error al guardar configuración: {e}")

# -------------------------------------------------------------
# PARSER Y GESTOR DE LISTAS DE CONFIGURACIÓN (.txt)
# -------------------------------------------------------------
def load_lists_data(lists_file_path=None):
    if not lists_file_path:
        cfg = load_config()
        lists_file_path = cfg.get("lists_file_path", DEFAULT_LISTS_FILE)

    if not os.path.exists(lists_file_path):
        # Crear por defecto si no existe
        create_default_lists_file(lists_file_path)

    sections = {
        "FOLIO_INICIAL": "F00000",
        "CENTRO_COSTOS": [],
        "AREAS": [],
        "MATERIALES": [],
        "PROYECTOS": []
    }

    current_section = None
    try:
        with open(lists_file_path, "r", encoding="utf-8") as f:
            for line in f:
                line_str = line.strip()
                if not line_str or line_str.startswith("#") or line_str.startswith(";"):
                    continue

                if line_str.startswith("[") and line_str.endswith("]"):
                    current_section = line_str[1:-1].strip().upper()
                else:
                    if current_section == "FOLIO_INICIAL":
                        sections["FOLIO_INICIAL"] = line_str
                    elif current_section in ["CENTRO_COSTOS", "AREAS", "MATERIALES", "PROYECTOS"]:
                        if line_str not in sections[current_section]:
                            sections[current_section].append(line_str)
    except Exception as e:
        print(f"Error al leer archivo de listas ({lists_file_path}): {e}")

    return sections

def create_default_lists_file(file_path):
    content = """[FOLIO_INICIAL]
F00000

[CENTRO_COSTOS]
TP08000
TP06000
TP01000
TP02000

[AREAS]
Ensamble
Mantenimiento
Calidad
Almacén
Preproceso

[MATERIALES]
3200277 - Tíner / Solvente
3200111 - Aceite Sintético
3200999 - Tornillería M6
3200555 - Cable de Calibración
3200444 - Conector Industrial
3200333 - Etiqueta de Inspección

[PROYECTOS]
PRJ-1001
PRJ-1002
PRJ-2026-A
PRJ-2026-B
"""
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception as e:
        print(f"Error al crear archivo de listas por defecto: {e}")

def save_lists_data(sections, lists_file_path=None):
    if not lists_file_path:
        cfg = load_config()
        lists_file_path = cfg.get("lists_file_path", DEFAULT_LISTS_FILE)
    
    try:
        with open(lists_file_path, "w", encoding="utf-8") as f:
            for section, items in sections.items():
                f.write(f"[{section}]\n")
                if isinstance(items, list):
                    for item in items:
                        f.write(f"{item}\n")
                else:
                    f.write(f"{items}\n")
                f.write("\n")
        return True
    except Exception as e:
        print(f"Error al guardar archivo de listas ({lists_file_path}): {e}")
        return False

# -------------------------------------------------------------
# GENERACIÓN Y SECTORES DE FOLIOS AUTOMÁTICOS (F00000, F00001...)
# -------------------------------------------------------------
def get_next_auto_folio():
    """
    Calcula el siguiente número de folio automático en secuencia F00000, F00001, etc.
    """
    lists_data = load_lists_data()
    initial_folio = lists_data.get("FOLIO_INICIAL", "F00000").strip().upper()

    # Extraer prefijo de letras y número base (ej: 'F' y 0)
    match = re.match(r"^([A-Z]+)(\d+)$", initial_folio)
    if match:
        prefix = match.group(1)
        base_num = int(match.group(2))
        num_len = len(match.group(2))
    else:
        prefix = "F"
        base_num = 0
        num_len = 5

    cfg = load_config()
    folios = cfg.get("generated_folios", {})
    deleted = cfg.get("deleted_folios", [])

    max_seq = max(base_num - 1, cfg.get("max_generated_seq", -1))

    if not cfg.get("ignore_sequence_history", False):
        for folio_key in list(folios.keys()) + deleted:
            f_match = re.match(r"^([A-Z]+)(\d+)$", folio_key.upper())
            if f_match and f_match.group(1) == prefix:
                seq_val = int(f_match.group(2))
                if seq_val > max_seq:
                    max_seq = seq_val

    next_seq = max_seq + 1
    next_folio_str = f"{prefix}{next_seq:0{num_len}d}"
    return next_folio_str

# -------------------------------------------------------------
# CONTROL DE ESTADO DE FOLIOS (BLOQUEADO / LIBERADO)
# -------------------------------------------------------------
def register_folio(folio_str, orden_str, fecha_str, accion_str, file_path, full_data=None, synced=False):
    """
    Registra un folio como BLOQUEADO (LOCKED) junto con sus datos completos.
    """
    if not folio_str:
        return
    
    with FileLock():
        cfg = load_config()
    folios = cfg.get("generated_folios", {})
    
    folio_key = folio_str.strip().upper()
    folios[folio_key] = {
        "folio": folio_str.strip(),
        "orden": orden_str.strip(),
        "fecha": fecha_str.strip(),
        "accion": accion_str.strip(),
        "file": file_path,
        "status": "LOCKED",
        "synced": synced,
        "timestamp": os.path.getmtime(file_path) if os.path.exists(file_path) else 0,
        "data": full_data or {}
    }

    f_match = re.match(r"^([A-Z]+)(\d+)$", folio_key)
    if f_match:
        seq_val = int(f_match.group(2))
        curr_max = cfg.get("max_generated_seq", 0)
        if seq_val > curr_max:
            cfg["max_generated_seq"] = seq_val

    cfg["generated_folios"] = folios
    save_config(cfg)

def is_folio_locked(folio_str):
    """
    Verifica si un folio existe y está en estado BLOQUEADO (LOCKED).
    """
    if not folio_str:
        return False
    
    cfg = load_config()
    folios = cfg.get("generated_folios", {})
    folio_key = folio_str.strip().upper()
    
    if folio_key in folios:
        return folios[folio_key].get("status") == "LOCKED"
    return False

def set_folio_status(folio_str, new_status):
    """
    Cambia el estado de un folio a 'LOCKED' o 'RELEASED'.
    """
    with FileLock():
        cfg = load_config()
    folios = cfg.get("generated_folios", {})
    folio_key = folio_str.strip().upper()
    
    if folio_key in folios:
        folios[folio_key]["status"] = new_status
        cfg["generated_folios"] = folios
        save_config(cfg)
        return True
    return False

def get_folio_record(folio_str):
    """
    Obtiene el registro completo de un folio.
    """
    cfg = load_config()
    folios = cfg.get("generated_folios", {})
    folio_key = folio_str.strip().upper()
    return folios.get(folio_key)

def delete_folio_record(folio_str):
    """
    Elimina un folio individual de los registros sin reciclar su número.
    """
    if not folio_str:
        return False

    with FileLock():
        cfg = load_config()
    folios = cfg.get("generated_folios", {})
    deleted = cfg.get("deleted_folios", [])

    folio_key = folio_str.strip().upper()
    if folio_key in folios:
        f_match = re.match(r"^([A-Z]+)(\d+)$", folio_key)
        if f_match:
            seq_val = int(f_match.group(2))
            curr_max = cfg.get("max_generated_seq", 0)
            if not cfg.get("ignore_sequence_history", False) and seq_val > curr_max:
                cfg["max_generated_seq"] = seq_val

        del folios[folio_key]
        if folio_key not in deleted:
            deleted.append(folio_key)

        cfg["generated_folios"] = folios
        cfg["deleted_folios"] = deleted
        save_config(cfg)
        return True
    return False

def clear_all_folios_history():
    """
    Elimina todo el historial de folios conservando la secuencia de folios generados para no repetir números.
    """
    with FileLock():
        cfg = load_config()
    folios = cfg.get("generated_folios", {})
    deleted = cfg.get("deleted_folios", [])
    max_seq = cfg.get("max_generated_seq", 0)

    ignore_history = cfg.get("ignore_sequence_history", False)

    for folio_key in folios.keys():
        f_match = re.match(r"^([A-Z]+)(\d+)$", folio_key.upper())
        if f_match:
            seq_val = int(f_match.group(2))
            if not ignore_history and seq_val > max_seq:
                max_seq = seq_val
        if folio_key not in deleted:
            deleted.append(folio_key)

    cfg["generated_folios"] = {}
    cfg["deleted_folios"] = deleted
    cfg["max_generated_seq"] = max_seq
    save_config(cfg)
    return True


def get_unsynced_folios():
    """
    Retorna la lista de diccionarios con datos completos de los folios que no han sido sincronizados.
    """
    cfg = load_config()
    folios = cfg.get("generated_folios", {})
    unsynced = []
    for key, f_data in folios.items():
        if "synced" in f_data and f_data["synced"] is False:
            unsynced.append(f_data)
    return unsynced

def mark_folio_synced(folio_str):
    """
    Marca un folio específico como sincronizado.
    """
    if not folio_str:
        return False
    
    with FileLock():
        cfg = load_config()
    folios = cfg.get("generated_folios", {})
    folio_key = folio_str.strip().upper()
    
    if folio_key in folios:
        folios[folio_key]["synced"] = True
        cfg["generated_folios"] = folios
        save_config(cfg)
        return True
    return False

def get_and_reserve_next_auto_folio():
    """
    Bloquea el archivo, calcula el siguiente folio, lo registra como 'PRE_RESERVED' 
    para incrementar max_generated_seq, y libera el bloqueo.
    """
    with FileLock():
        next_folio = get_next_auto_folio()
        cfg = load_config()
        folios = cfg.get("generated_folios", {})
        folio_key = next_folio.strip().upper()
        folios[folio_key] = {
            "folio": next_folio.strip(),
            "status": "PRE_RESERVED",
            "timestamp": time.time(),
            "data": {}
        }
        match = re.match(r"^([A-Z]+)(\d+)$", folio_key)
        if match:
            seq_val = int(match.group(2))
            curr_max = cfg.get("max_generated_seq", 0)
            if seq_val > curr_max:
                cfg["max_generated_seq"] = seq_val
        cfg["generated_folios"] = folios
        save_config(cfg)
        return next_folio
