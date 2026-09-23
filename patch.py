import re
import time
import os

with open('config.py', 'r', encoding='utf-8') as f:
    code = f.read()

code = code.replace('import re\n', '''import re
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
''')

# Now add locking to the methods that write data
methods_to_lock = [
    ('def set_folio_status(folio_str, new_status):\n', '    with FileLock():\n        cfg = load_config()\n', '    cfg = load_config()\n'),
    ('def delete_folio_record(folio_str):\n', '    with FileLock():\n        cfg = load_config()\n', '    cfg = load_config()\n'),
    ('def clear_all_folios_history():\n', '    with FileLock():\n        cfg = load_config()\n', '    cfg = load_config()\n'),
    ('def mark_folio_synced(folio_str):\n', '    with FileLock():\n        cfg = load_config()\n', '    cfg = load_config()\n'),
    ('def register_folio(', '    with FileLock():\n        cfg = load_config()\n', '    cfg = load_config()\n')
]

for method, lock_code, orig_code in methods_to_lock:
    # Find the method start
    idx = code.find(method)
    if idx != -1:
        # Find the first cfg = load_config() after the method def
        idx2 = code.find(orig_code, idx)
        if idx2 != -1:
            code = code[:idx2] + lock_code + code[idx2 + len(orig_code):]

# Add get_and_reserve_next_auto_folio
code += '''
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
