import re
import time
import os

with open('config.py', 'r', encoding='utf-8') as f:
    code = f.read()

code = code.replace('import re\n', 'import re\nimport time\n\nLOCK_DIR = \'app_config.lock\'\n\nclass FileLock:\n    def __enter__(self):\n        while True:\n            try:\n                os.mkdir(LOCK_DIR)\n                break\n            except FileExistsError:\n                try:\n                    if os.path.exists(LOCK_DIR) and (time.time() - os.path.getctime(LOCK_DIR)) > 10:\n                        os.rmdir(LOCK_DIR)\n                except Exception:\n                    pass\n                time.sleep(0.2)\n        return self\n\n    def __exit__(self, exc_type, exc_val, exc_tb):\n        try:\n            os.rmdir(LOCK_DIR)\n        except Exception:\n            pass\n')

code = code.replace(
    '    if not folio_str:\n        return\n    \n    cfg = load_config()',
    '    if not folio_str:\n        return\n    \n    with FileLock():\n        cfg = load_config()'
)

code = code.replace(
    '    if not folio_str:\n        return False\n    \n    cfg = load_config()',
    '    if not folio_str:\n        return False\n    \n    with FileLock():\n        cfg = load_config()'
)

code = code.replace(
    'def set_folio_status(folio_str, new_status):\n    \"\"\"\n    Cambia el estado de un folio a \'LOCKED\' o \'RELEASED\'.\n    \"\"\"\n    cfg = load_config()',
    'def set_folio_status(folio_str, new_status):\n    \"\"\"\n    Cambia el estado de un folio a \'LOCKED\' o \'RELEASED\'.\n    \"\"\"\n    with FileLock():\n        cfg = load_config()'
)

code = code.replace(
    'def delete_folio_record(folio_str):\n    \"\"\"\n    Elimina un folio individual de los registros sin reciclar su número.\n    \"\"\"\n    if not folio_str:\n        return False\n\n    cfg = load_config()',
    'def delete_folio_record(folio_str):\n    \"\"\"\n    Elimina un folio individual de los registros sin reciclar su número.\n    \"\"\"\n    if not folio_str:\n        return False\n\n    with FileLock():\n        cfg = load_config()'
)

code = code.replace(
    'def clear_all_folios_history():\n    \"\"\"\n    Elimina todo el historial de folios conservando la secuencia de folios generados para no repetir números.\n    \"\"\"\n    cfg = load_config()',
    'def clear_all_folios_history():\n    \"\"\"\n    Elimina todo el historial de folios conservando la secuencia de folios generados para no repetir números.\n    \"\"\"\n    with FileLock():\n        cfg = load_config()'
)

code += '''
def get_and_reserve_next_auto_folio():
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
        match = re.match(r"^([A-Z]+)(\\d+)$", folio_key)
        if match:
            seq_val = int(match.group(2))
            curr_max = cfg.get("max_generated_seq", 0)
            if seq_val > curr_max:
                cfg["max_generated_seq"] = seq_val
        cfg["generated_folios"] = folios
        save_config(cfg)
        return next_folio
'''

with open('config.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Patch applied successfully.")
