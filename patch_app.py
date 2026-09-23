import re

with open('app.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Add self.is_modifying_folio = False in __init__
init_hook = 'self.calendar_popup = None\n'
if init_hook in code:
    code = code.replace(init_hook, init_hook + '        self.is_modifying_folio = False\n        self.after(2000, self._poll_auto_folio)\n')

# 2. Add _poll_auto_folio method in AppFormulario
poll_method = '''
    def _poll_auto_folio(self):
        try:
            if hasattr(self, "entry_folio") and self.entry_folio.winfo_exists() and self.entry_folio.cget("state") == "disabled":
                if hasattr(self, "main_tabview") and hasattr(self, "sub_tabview"):
                    if self.main_tabview.get() == "📝 Formato de adicionales" and self.sub_tabview.get() == "✨ Nueva Creación":
                        if not getattr(self, "is_modifying_folio", False):
                            current_val = self.entry_folio.get()
                            next_auto = config.get_next_auto_folio()
                            if current_val != next_auto:
                                self.entry_folio.configure(state="normal")
                                self.entry_folio.delete(0, "end")
                                self.entry_folio.insert(0, next_auto)
                                self.entry_folio.configure(state="disabled", fg_color="#F1F5F9")
        except Exception:
            pass
        self.after(2000, self._poll_auto_folio)
'''
# Add it before def _run_google_sheets_sync_bg
code = code.replace('    def _run_google_sheets_sync_bg(self):', poll_method + '\n    def _run_google_sheets_sync_bg(self):')


# 3. Update is_modifying_folio in _reset_form_to_initial_state and _load_folio_into_form
code = code.replace(
    '    def _reset_form_to_initial_state(self):\n',
    '    def _reset_form_to_initial_state(self):\n        self.is_modifying_folio = False\n'
)

code = code.replace(
    '    def _load_folio_into_form(self, record, folio_code):\n',
    '    def _load_folio_into_form(self, record, folio_code):\n        self.is_modifying_folio = True\n'
)

# 4. In on_generar_click, reserve the actual folio just before saving if not modifying and entry is disabled
# Find: data = self.collect_form_data()
collect_replacement = '''
        if not getattr(self, "is_modifying_folio", False) and self.entry_folio.cget("state") == "disabled":
            real_auto_folio = config.get_and_reserve_next_auto_folio()
            self.entry_folio.configure(state="normal")
            self.entry_folio.delete(0, "end")
            self.entry_folio.insert(0, real_auto_folio)
            self.entry_folio.configure(state="disabled", fg_color="#F1F5F9")

        data = self.collect_form_data()
'''
code = code.replace('        data = self.collect_form_data()\n', collect_replacement, 1)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(code)
