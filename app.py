import os
import sys
import subprocess
import threading
import calendar
from datetime import datetime
import tkinter as tk
from PIL import Image
import customtkinter as ctk
from tkinter import filedialog, messagebox

import config
import excel_handler
import sheets_handler

# Paleta de colores optimizada para máximo contraste visual (Wasion Blue Theme)
PAGE_BG       = "#F0F5FB"  # Fondo azul suave opaco
CARD_BG       = "#FFFFFF"  # Tarjetas blancas opacas
INPUT_BG      = "#FFFFFF"  # Fondo blanco para campos de texto
BORDER_COLOR  = "#BBDEFB"  # Borde azul claro
PRIMARY_NAVY  = "#0D47A1"  # Azul marino Wasion principal
ACCENT_BLUE   = "#1976D2"  # Azul rey vibrante
ACCENT_HOVER  = "#1565C0"
TEXT_MAIN     = "#0A1929"  # Texto principal oscuro para fondos claros
TEXT_ON_BLUE  = "#FFFFFF"  # Texto blanco para fondos azules
TEXT_SUB      = "#455A64"  # Texto secundario
SUCCESS_GREEN = "#2E7D32"  # Verde acción / agregar
SUCCESS_HOVER = "#1B5E20"
DANGER_RED    = "#C62828"  # Rojo eliminar
DANGER_HOVER  = "#991B1B"

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# Registro global de combolistas activos para ocultamiento automático al hacer clic fuera
ACTIVE_DROPDOWNS = []


def register_dropdown(combo_obj):
    if combo_obj not in ACTIVE_DROPDOWNS:
        ACTIVE_DROPDOWNS.append(combo_obj)


def close_all_dropdowns(except_combo=None):
    for c in ACTIVE_DROPDOWNS:
        if c != except_combo:
            c.hide_popup()


def is_inside(widget, container):
    """
    Verifica de forma 100% segura la relación jerárquica de widgets sin lanzar excepciones.
    """
    if not widget or not container:
        return False
    try:
        w_str = str(widget)
        c_str = str(container)
        return w_str == c_str or w_str.startswith(c_str + ".")
    except Exception:
        return False


class PasswordDialog(ctk.CTkToplevel):
    """
    Modal de autenticación por contraseña para la sección de Configuración (PLN2026).
    """
    def __init__(self, parent, on_success_callback, on_cancel_callback=None):
        super().__init__(parent)
        self.title("Autenticación de Administrador")
        self.geometry("420x240")
        self.resizable(False, False)
        self.on_success_callback = on_success_callback
        self.on_cancel_callback = on_cancel_callback

        self.transient(parent)
        self.grab_set()

        try:
            parent.update_idletasks()
            pw = 420
            ph = 240
            px = parent.winfo_x() + (parent.winfo_width() // 2) - (pw // 2)
            py = parent.winfo_y() + (parent.winfo_height() // 2) - (ph // 2)
            self.geometry(f"{pw}x{ph}+{max(0, px)}+{max(0, py)}")
        except Exception:
            pass

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        lbl_title = ctk.CTkLabel(
            self, text="🔒 Acceso a Configuración", font=("Segoe UI", 16, "bold"), text_color=PRIMARY_NAVY
        )
        lbl_title.pack(pady=(20, 5))

        lbl_sub = ctk.CTkLabel(
            self, text="Ingrese la contraseña de administrador para continuar:", font=("Segoe UI", 12), text_color=TEXT_MAIN
        )
        lbl_sub.pack(pady=(0, 15))

        self.entry_pass = ctk.CTkEntry(
            self, show="*", width=280, height=38, font=("Segoe UI", 14), fg_color="#FFFFFF", border_color=BORDER_COLOR, text_color=TEXT_MAIN
        )
        self.entry_pass.pack(pady=(0, 20))
        self.entry_pass.focus_set()
        self.entry_pass.bind("<Return>", lambda e: self.check_password())

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=25)

        ctk.CTkButton(
            btn_frame, text="Cancelar", fg_color="#64748B", hover_color="#475569", text_color="#FFFFFF", width=120, height=36, command=self._on_close
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame, text="Ingresar", fg_color=PRIMARY_NAVY, hover_color=ACCENT_HOVER, text_color="#FFFFFF", width=120, height=36, command=self.check_password
        ).pack(side="right")

        self.after(60, self._force_focus_entry)

    def _force_focus_entry(self):
        try:
            if hasattr(self, "entry_pass") and self.entry_pass.winfo_exists():
                self.entry_pass.focus_force()
        except Exception:
            pass

    def check_password(self):
        entered = self.entry_pass.get().strip()
        if entered == "PLN2026":
            self.destroy()
            self.on_success_callback()
        else:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta. Intente nuevamente.")
            self.entry_pass.delete(0, "end")

    def _on_close(self):
        self.destroy()
        if self.on_cancel_callback:
            self.on_cancel_callback()


class DropdownCalendarPopup:
    """
    Calendario desplegable flotante que bloquea el scroll de la página principal mientras permanece abierto.
    """
    def __init__(self, master_window, target_entry, on_date_selected_callback):
        self.master_window = master_window
        self.target_entry = target_entry
        self.on_date_selected_callback = on_date_selected_callback
        self.popup = None
        now = datetime.now()
        self.year = now.year
        self.month = now.month

    def toggle(self):
        if self.popup and self.popup.winfo_exists():
            self.hide()
        else:
            self.show()

    def show(self):
        close_all_dropdowns()

        if not self.popup or not self.popup.winfo_exists():
            self.popup = tk.Toplevel(self.master_window)
            self.popup.wm_overrideredirect(True)
            self.popup.configure(bg="#BBDEFB")

        try:
            x = self.target_entry.winfo_rootx()
            y = self.target_entry.winfo_rooty() + self.target_entry.winfo_height() + 2
            w = 320
            h = 290
            self.popup.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            return

        for child in self.popup.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass

        card = ctk.CTkFrame(self.popup, fg_color="#FFFFFF", border_color=BORDER_COLOR, border_width=1, corner_radius=10)
        card.pack(fill="both", expand=True)

        nav_frame = ctk.CTkFrame(card, fg_color=PRIMARY_NAVY, corner_radius=0, height=40)
        nav_frame.pack(fill="x", side="top")

        ctk.CTkButton(
            nav_frame, text="◀", width=34, height=28, fg_color=ACCENT_BLUE, hover_color=ACCENT_HOVER, text_color="#FFFFFF",
            command=self.prev_month
        ).pack(side="left", padx=6, pady=5)

        self.lbl_month_year = ctk.CTkLabel(
            nav_frame, text="", font=("Segoe UI", 12, "bold"), text_color="#FFFFFF"
        )
        self.lbl_month_year.pack(side="left", expand=True)

        ctk.CTkButton(
            nav_frame, text="▶", width=34, height=28, fg_color=ACCENT_BLUE, hover_color=ACCENT_HOVER, text_color="#FFFFFF",
            command=self.next_month
        ).pack(side="right", padx=6, pady=5)

        days_hdr = ctk.CTkFrame(card, fg_color="#E2E8F0", corner_radius=0)
        days_hdr.pack(fill="x", padx=6, pady=(6, 0))
        days_hdr.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)

        days = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]
        for idx, d in enumerate(days):
            lbl = ctk.CTkLabel(days_hdr, text=d, font=("Segoe UI", 10, "bold"), text_color=TEXT_MAIN)
            lbl.grid(row=0, column=idx, pady=3, sticky="ew")

        # CONTENEDOR CON SCROLL INTERNO PARA EL CALENDARIO
        self.days_scroll = ctk.CTkScrollableFrame(card, fg_color="#FFFFFF", bg_color="#FFFFFF", corner_radius=0)
        self.days_scroll.pack(fill="both", expand=True, padx=6, pady=6)
        self.days_scroll.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)

        self.render_calendar()

        # BLOQUEAR EL SCROLL DE LA PÁGINA PRINCIPAL MIENTRAS EL CALENDARIO ESTÁ ABIERTO
        self.master_window.disable_main_scroll()
        self.popup.bind_all("<MouseWheel>", self._on_popup_scroll)
        self.popup.bind_all("<Button-4>", self._on_popup_scroll)
        self.popup.bind_all("<Button-5>", self._on_popup_scroll)

    def _on_popup_scroll(self, event):
        """
        Redirige el desplazamiento de la rueda del ratón ÚNICAMENTE al calendario sin mover la página.
        """
        try:
            if event.delta:
                self.days_scroll._parent_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                self.days_scroll._parent_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.days_scroll._parent_canvas.yview_scroll(1, "units")
        except Exception:
            pass
        return "break"

    def prev_month(self):
        if self.month == 1:
            self.month = 12
            self.year -= 1
        else:
            self.month -= 1
        self.render_calendar()

    def next_month(self):
        if self.month == 12:
            self.month = 1
            self.year += 1
        else:
            self.month += 1
        self.render_calendar()

    def render_calendar(self):
        for child in self.days_scroll.winfo_children():
            child.destroy()

        month_names = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        self.lbl_month_year.configure(text=f"{month_names[self.month - 1]} {self.year}")

        weeks = calendar.monthcalendar(self.year, self.month)
        today = datetime.now()

        for r, week in enumerate(weeks):
            for c, day in enumerate(week):
                if day == 0:
                    lbl_empty = ctk.CTkLabel(self.days_scroll, text="", height=28)
                    lbl_empty.grid(row=r, column=c, padx=1, pady=1, sticky="ew")
                else:
                    is_today = (day == today.day and self.month == today.month and self.year == today.year)
                    btn_color = PRIMARY_NAVY if is_today else "#F1F5F9"
                    txt_color = "#FFFFFF" if is_today else TEXT_MAIN

                    btn_day = ctk.CTkButton(
                        self.days_scroll,
                        text=str(day),
                        height=28,
                        fg_color=btn_color,
                        hover_color=ACCENT_HOVER,
                        text_color=txt_color,
                        font=("Segoe UI", 11, "bold" if is_today else "normal"),
                        command=lambda d=day: self.select_day(d)
                    )
                    btn_day.grid(row=r, column=c, padx=1, pady=1, sticky="ew")

    def select_day(self, day):
        selected_date_str = f"{day:02d}/{self.month:02d}/{self.year}"
        self.on_date_selected_callback(selected_date_str)
        self.hide()

    def hide(self):
        if self.popup and self.popup.winfo_exists():
            p = self.popup
            self.popup = None
            try:
                p.withdraw()
                self.master_window.after(20, lambda: self._safe_destroy(p))
            except Exception:
                pass
        self.master_window.enable_main_scroll()

    def _safe_destroy(self, p):
        try:
            if p and p.winfo_exists():
                p.destroy()
        except Exception:
            pass


class SearchableComboBox(ctk.CTkFrame):
    """
    Componente desplegable ultra fluido con filtrado en tiempo real y cierre 100% seguro al hacer clic fuera.
    """
    def __init__(self, master, values=None, placeholder="", height=38):
        super().__init__(master, fg_color="transparent")
        self.values = values or []
        self.popup = None
        register_dropdown(self)

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)

        self.entry = ctk.CTkEntry(
            self,
            placeholder_text=placeholder,
            height=height,
            font=("Segoe UI", 12),
            fg_color="#FFFFFF",
            border_color=BORDER_COLOR,
            text_color=TEXT_MAIN,
            placeholder_text_color="#64748B"
        )
        self.entry.grid(row=0, column=0, sticky="ew")
        self.entry.bind("<KeyRelease>", self._on_key_release)
        self.entry.bind("<Button-1>", self._on_entry_click)

        self.btn_arrow = ctk.CTkButton(
            self,
            text="▼",
            width=34,
            height=height,
            fg_color="#E2E8F0",
            hover_color="#CBD5E1",
            text_color=TEXT_MAIN,
            font=("Segoe UI", 10, "bold"),
            command=self.toggle_popup
        )
        self.btn_arrow.grid(row=0, column=1, padx=(2, 0))

    def _on_entry_click(self, event=None):
        query = self.entry.get().strip().lower()
        if not query:
            filtered = self.values
        else:
            filtered = [v for v in self.values if query in v.lower()]
        self.after(30, lambda: self.show_popup(filtered))

    def _on_key_release(self, event):
        if event.keysym in ["Return", "Escape", "Tab", "Up", "Down"]:
            if event.keysym in ["Escape", "Return"]:
                self.hide_popup()
            return
        query = self.entry.get().strip().lower()
        if not query:
            filtered = self.values
        else:
            filtered = [v for v in self.values if query in v.lower()]
        self.show_popup(filtered)

    def toggle_popup(self):
        if self.popup and self.popup.winfo_exists():
            self.hide_popup()
        else:
            self.show_popup(self.values)

    def show_popup(self, options):
        close_all_dropdowns(except_combo=self)
        if not options:
            self.hide_popup()
            return

        if not self.popup or not self.popup.winfo_exists():
            self.popup = tk.Toplevel(self)
            self.popup.wm_overrideredirect(True)
            self.popup.configure(bg="#BBDEFB")

        try:
            x = self.entry.winfo_rootx()
            y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
            w = max(self.winfo_width(), 260)
            h = min(220, len(options) * 28 + 8)
            self.popup.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            return

        for child in self.popup.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass

        frame = tk.Frame(self.popup, bg="#FFFFFF", bd=1, relief="solid")
        frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")

        listbox = tk.Listbox(
            frame,
            yscrollcommand=scrollbar.set,
            font=("Segoe UI", 11),
            bg="#FFFFFF",
            fg="#0A1929",
            selectbackground=PRIMARY_NAVY,
            selectforeground="#FFFFFF",
            bd=0,
            activestyle="none"
        )
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        for opt in options:
            listbox.insert("end", opt)

        def on_select(evt):
            sel = listbox.curselection()
            if sel:
                val = listbox.get(sel[0])
                self.set(val)
            self.hide_popup()

        listbox.bind("<ButtonRelease-1>", on_select)

    def hide_popup(self):
        if self.popup and self.popup.winfo_exists():
            p = self.popup
            self.popup = None
            try:
                p.withdraw()
                self.after(20, lambda: self._safe_destroy(p))
            except Exception:
                pass

    def _safe_destroy(self, p):
        try:
            if p and p.winfo_exists():
                p.destroy()
        except Exception:
            pass

    def get(self):
        return self.entry.get().strip()

    def set(self, val):
        self.entry.delete(0, "end")
        self.entry.insert(0, str(val))

    def configure_values(self, new_values):
        self.values = new_values or []


class ActionSelectorFrame(ctk.CTkFrame):
    """
    Componente de selección de acción con 100% contraste de texto.
    """
    def __init__(self, master, default_option=None, command_callback=None):
        super().__init__(master, fg_color="#F1F5F9", border_color=BORDER_COLOR, border_width=1, corner_radius=10)
        self.options = ["Return Ok", "Scrap", "Adicional", "Consumibles"]
        self.current_value = default_option
        self.command_callback = command_callback
        self.buttons = {}

        self.grid_columnconfigure((0, 1, 2, 3), weight=1)

        for idx, opt in enumerate(self.options):
            btn = ctk.CTkButton(
                self,
                text=opt,
                font=("Segoe UI", 13, "bold"),
                height=42,
                corner_radius=8,
                command=lambda o=opt: self.select_option(o)
            )
            btn.grid(row=0, column=idx, padx=4, pady=4, sticky="ew")
            self.buttons[opt] = btn

        self._update_styles()

    def select_option(self, option):
        self.current_value = option
        self._update_styles()
        if self.command_callback:
            self.command_callback(option)

    def _update_styles(self):
        for opt, btn in self.buttons.items():
            if opt == self.current_value:
                btn.configure(
                    fg_color=PRIMARY_NAVY,
                    hover_color=ACCENT_HOVER,
                    text_color="#FFFFFF"
                )
            else:
                btn.configure(
                    fg_color="#FFFFFF",
                    hover_color="#E2E8F0",
                    text_color=TEXT_MAIN
                )

    def set(self, option):
        self.current_value = option
        self._update_styles()

    def get(self):
        return self.current_value


class MaterialRowFrame(ctk.CTkFrame):
    """
    Componente para cada fila de material con cajas de texto homogéneas.
    """
    def __init__(self, master, index, on_delete_callback, materials_options=None, default_mat="", default_cant="", default_batch="", default_desc=""):
        super().__init__(
            master,
            fg_color="#F8FAFC",
            border_color=BORDER_COLOR,
            border_width=1,
            corner_radius=8
        )
        self.index = index
        self.on_delete_callback = on_delete_callback
        self.materials_options = materials_options or ["Material Genérico"]

        self.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.grid_columnconfigure(4, weight=0)

        # -------------------------------------------------------------
        # COLUMNA 1: MATERIAL (Desplegable con búsqueda)
        # -------------------------------------------------------------
        self.combo_material = SearchableComboBox(
            self,
            values=self.materials_options,
            placeholder=f"Escriba o seleccione Material #{index+1}...",
            height=38
        )
        if default_mat:
            self.combo_material.set(default_mat)
        self.combo_material.grid(row=0, column=0, padx=6, pady=6, sticky="ew")

        # -------------------------------------------------------------
        # COLUMNA 2: BATCH / LOTE
        # -------------------------------------------------------------
        self.entry_batch = ctk.CTkEntry(
            self,
            placeholder_text="ej. Lote / Batch",
            height=38,
            font=("Segoe UI", 12),
            fg_color=INPUT_BG,
            border_color=BORDER_COLOR,
            text_color=TEXT_MAIN,
            placeholder_text_color="#64748B"
        )
        if default_batch:
            self.entry_batch.insert(0, default_batch)
        self.entry_batch.grid(row=0, column=1, padx=6, pady=6, sticky="ew")

        # -------------------------------------------------------------
        # COLUMNA 3: CANTIDAD REQUERIDA (Mismo formato exacto de caja que Material)
        # -------------------------------------------------------------
        self.entry_cantidad = ctk.CTkEntry(
            self,
            placeholder_text="ej. 5 Litros / 2 Pzs",
            height=38,
            font=("Segoe UI", 12),
            fg_color=INPUT_BG,
            border_color=BORDER_COLOR,
            text_color=TEXT_MAIN,
            placeholder_text_color="#64748B"
        )
        if default_cant:
            self.entry_cantidad.insert(0, default_cant)
        self.entry_cantidad.grid(row=0, column=2, padx=6, pady=6, sticky="ew")

        # -------------------------------------------------------------
        # COLUMNA 4: DESCRIPCIÓN / ANÁLISIS (Mismo formato exacto de caja que Material)
        # -------------------------------------------------------------
        self.entry_descripcion = ctk.CTkEntry(
            self,
            placeholder_text="ej. Observaciones / Análisis",
            height=38,
            font=("Segoe UI", 12),
            fg_color=INPUT_BG,
            border_color=BORDER_COLOR,
            text_color=TEXT_MAIN,
            placeholder_text_color="#64748B"
        )
        if default_desc:
            self.entry_descripcion.insert(0, default_desc)
        self.entry_descripcion.grid(row=0, column=3, padx=6, pady=6, sticky="ew")

        # BOTÓN ELIMINAR FILA
        self.btn_delete = ctk.CTkButton(
            self,
            text="✕",
            width=38,
            height=38,
            fg_color=DANGER_RED,
            hover_color=DANGER_HOVER,
            text_color=TEXT_ON_BLUE,
            font=("Segoe UI", 14, "bold"),
            command=self.delete_self
        )
        self.btn_delete.grid(row=0, column=4, padx=(4, 6), pady=6)

    def update_material_options(self, new_options):
        self.materials_options = new_options or ["Material Genérico"]
        self.combo_material.configure_values(self.materials_options)

    def delete_self(self):
        self.on_delete_callback(self)

    def get_data(self):
        return {
            "material": self.combo_material.get(),
            "cantidad": self.entry_cantidad.get().strip(),
            "batch": self.entry_batch.get().strip(),
            "descripcion": self.entry_descripcion.get().strip()
        }


class AppFormulario(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sistema de Formulario de Adicionales y Retornos - Wasion")
        self.geometry("1180x920")
        self.minsize(1040, 800)
        # Centrar la ventana como respaldo y luego maximizar de forma confiable
        window_width = 1180
        window_height = 920
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width / 2)
        center_y = int(screen_height/2 - window_height / 2)
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        # Maximizar después de que la interfaz haya cargado (evita glitch de tamaño)
        self.after(200, lambda: self.state("zoomed"))

        if hasattr(sys, '_MEIPASS'):
            self.base_dir = sys._MEIPASS
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.icon_path = os.path.join(self.base_dir, "app_icon.ico")
        if os.path.exists(self.icon_path):
            try:
                self.iconbitmap(self.icon_path)
            except Exception:
                pass

        self.configure(fg_color=PAGE_BG)
        self.app_config = config.load_config()
        self.lists_data = config.load_lists_data()
        self.material_rows = []
        self.calendar_popup = None
        self.is_modifying_folio = False
        self.after(2000, self._poll_auto_folio)

        # REGISTRO GLOBAL DE CLIC EN LA VENTANA PARA OCULTAR DESPLEGABLES DE FORMA 100% SEGURA
        self.bind_all("<Button-1>", self._on_global_window_click)

        self._create_ui()
        self._set_auto_folio()
        self._run_google_sheets_sync_bg()



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

    def _run_google_sheets_sync_bg(self):
        def sync_worker():
            try:
                auto_sync = self.app_config.get("auto_sync_google_sheets", True)
                if not auto_sync: return
                unsynced = config.get_unsynced_folios()
                if not unsynced: return
                creds_path = self.cfg_credentials.get().strip() or self.app_config.get("credentials_path")
                gsheet_name = self.cfg_gsheet.get().strip() or self.app_config.get("google_sheet_url_or_name")
                for f_data in unsynced:
                    ok, msg = sheets_handler.sync_to_google_sheets(f_data["data"], credentials_path=creds_path, sheet_url_or_name=gsheet_name)
                    if ok:
                        config.mark_folio_synced(f_data["folio"])
            except Exception as e:
                print(f"Error background sync: {e}")
                
        threading.Thread(target=sync_worker, daemon=True).start()

    def disable_main_scroll(self):
        """
        Desactiva temporalmente el desplazamiento de la página principal mientras el calendario está desplegado.
        """
        try:
            self.unbind_all("<MouseWheel>")
            self.unbind_all("<Button-4>")
            self.unbind_all("<Button-5>")
        except Exception:
            pass

    def enable_main_scroll(self):
        """
        Restaura el desplazamiento estándar de la página principal al cerrar el calendario.
        """
        try:
            if hasattr(self, "scroll_nueva") and hasattr(self.scroll_nueva, "_mouse_wheel_all"):
                self.bind_all("<MouseWheel>", self.scroll_nueva._mouse_wheel_all)
        except Exception:
            pass

    def _on_global_window_click(self, event):
        widget = event.widget
        # Ocultar comboboxes únicamente si se hace clic fuera del combo y de su popup
        for c in ACTIVE_DROPDOWNS:
            if c.popup and c.popup.winfo_exists():
                if is_inside(widget, c) or is_inside(widget, c.popup):
                    continue
                c.hide_popup()

        # Ocultar calendario desplegable si se hace clic fuera de él
        if self.calendar_popup and self.calendar_popup.popup and self.calendar_popup.popup.winfo_exists():
            if is_inside(widget, self.entry_fecha) or is_inside(widget, self.btn_cal_toggle) or is_inside(widget, self.calendar_popup.popup):
                pass
            else:
                self.calendar_popup.hide()

    def _create_ui(self):
        # Header Principal
        self.header_frame = ctk.CTkFrame(self, fg_color=PRIMARY_NAVY, corner_radius=0, height=95)
        self.header_frame.pack(fill="x", side="top")

        # Logo esquina superior izquierda
        logo_path = os.path.join(self.base_dir, "wasion_logo_cropped.png")
        if not os.path.exists(logo_path):
            logo_path = os.path.join(self.base_dir, "wasion_logo.png")

        if os.path.exists(logo_path):
            try:
                logo_img_pil = Image.open(logo_path)
                self.logo_ctk = ctk.CTkImage(light_image=logo_img_pil, dark_image=logo_img_pil, size=(160, 62))
                self.logo_card = ctk.CTkFrame(self.header_frame, fg_color="#FFFFFF", corner_radius=10)
                self.logo_card.pack(side="left", padx=(20, 15), pady=12)
                self.logo_label = ctk.CTkLabel(self.logo_card, image=self.logo_ctk, text="")
                self.logo_label.pack(padx=8, pady=4)
            except Exception as e:
                print(f"Aviso logo: {e}")

        # Título Principal (Sin subtítulo)
        self.header_title_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.header_title_frame.pack(side="left", padx=10, pady=20)

        self.title_label = ctk.CTkLabel(
            self.header_title_frame, text="Formulario de SCRAP, Adicionales y Retornos", font=("Segoe UI", 22, "bold"), text_color=TEXT_ON_BLUE
        )
        self.title_label.pack(anchor="w")

        # Tabview principal con 100% contraste de texto
        self.main_tabview = ctk.CTkTabview(
            self,
            corner_radius=12,
            fg_color=PAGE_BG,
            segmented_button_fg_color="#CBD5E1",
            segmented_button_selected_color=PRIMARY_NAVY,
            segmented_button_selected_hover_color=ACCENT_HOVER,
            segmented_button_unselected_color="#FFFFFF",
            segmented_button_unselected_hover_color="#E2E8F0",
            command=self._on_main_tab_change
        )
        self.main_tabview.pack(fill="both", expand=True, padx=20, pady=15)

        self.tab_adicionales = self.main_tabview.add("📝 Formato de adicionales")
        self.tab_config = self.main_tabview.add("⚙️ Configuración")

        self._update_tab_button_contrast(self.main_tabview)

        self._build_tab_formato_adicionales()
        self._build_tab_config()

        # Barra de estado inferior
        self.status_bar = ctk.CTkFrame(self, height=42, fg_color=PRIMARY_NAVY, corner_radius=0)
        self.status_bar.pack(fill="x", side="bottom")

        self.status_label = ctk.CTkLabel(
            self.status_bar, text="Listo para ingresar datos.", font=("Segoe UI", 12, "bold"), text_color="#90CAF9"
        )
        self.status_label.pack(side="left", padx=20, pady=8)

        # Marca de agua en la esquina inferior derecha
        self.watermark = ctk.CTkLabel(
            self.status_bar, text="Created by Fernando Carrasco", 
            font=("Segoe UI", 11, "italic"), text_color="#90CAF9"
        )
        self.watermark.pack(side="right", padx=20, pady=8)

    def _update_tab_button_contrast(self, tabview_obj):
        """
        Asegura que las pestañas inactivas tengan texto azul oscuro/negro sobre blanco y la activa texto blanco sobre azul.
        """
        try:
            selected = tabview_obj.get()
            for tab_name, btn in tabview_obj._segmented_button._buttons_dict.items():
                if tab_name == selected:
                    btn.configure(fg_color=PRIMARY_NAVY, text_color="#FFFFFF")
                else:
                    btn.configure(fg_color="#FFFFFF", text_color=TEXT_MAIN)
        except Exception:
            pass

    def _on_main_tab_change(self):
        self._update_tab_button_contrast(self.main_tabview)
        selected = self.main_tabview.get()
        if selected == "📝 Formato de adicionales":
            self.reload_lists_data()

        # SOLICITAR CONTRASEÑA CADA VEZ QUE SE ENTRA A CONFIGURACIÓN
        if selected == "⚙️ Configuración":
            self.main_tabview.set("📝 Formato de adicionales")
            self._update_tab_button_contrast(self.main_tabview)

            PasswordDialog(
                self,
                on_success_callback=self._grant_config_access,
                on_cancel_callback=lambda: self.status_label.configure(text="Acceso cancelado.", text_color="#90CAF9")
            )

    def _grant_config_access(self):
        self.main_tabview.set("⚙️ Configuración")
        self._update_tab_button_contrast(self.main_tabview)
        self.status_label.configure(text="Acceso concedido a Configuración.", text_color="#90CAF9")
        self._load_txt_file_to_editor()

    def _build_tab_formato_adicionales(self):
        self.sub_tabview = ctk.CTkTabview(
            self.tab_adicionales,
            corner_radius=10,
            fg_color=PAGE_BG,
            segmented_button_fg_color="#E2E8F0",
            segmented_button_selected_color=PRIMARY_NAVY,
            segmented_button_selected_hover_color=ACCENT_HOVER,
            segmented_button_unselected_color="#FFFFFF",
            segmented_button_unselected_hover_color="#CBD5E1",
            command=self._on_sub_tab_change
        )
        self.sub_tabview.pack(fill="both", expand=True, padx=5, pady=5)

        self.sub_tab_nueva = self.sub_tabview.add("✨ Nueva Creación")
        self.sub_tab_existente = self.sub_tabview.add("✏️ Modificar Existente")

        self._update_tab_button_contrast(self.sub_tabview)

        self._build_sub_tab_nueva()
        self._build_sub_tab_existente()

    def _on_sub_tab_change(self):
        self._update_tab_button_contrast(self.sub_tabview)
        self.reload_lists_data()

    def _build_sub_tab_nueva(self):
        # Fondo explícito PAGE_BG para evitar smearing al desplazar verticalmente
        self.scroll_nueva = ctk.CTkScrollableFrame(self.sub_tab_nueva, fg_color=PAGE_BG, bg_color=PAGE_BG)
        self.scroll_nueva.pack(fill="both", expand=True, padx=5, pady=5)

        # -------------------------------------------------------------
        # SECCIÓN 1: TIPO DE FORMULARIO
        # -------------------------------------------------------------
        self.sec1_box = ctk.CTkFrame(self.scroll_nueva, fg_color=CARD_BG, border_color=BORDER_COLOR, border_width=1, corner_radius=12)
        self.sec1_box.pack(fill="x", pady=(0, 15), padx=5)

        sec1_title = ctk.CTkLabel(self.sec1_box, text="1. Tipo de Formulario", font=("Segoe UI", 15, "bold"), text_color=PRIMARY_NAVY)
        sec1_title.pack(anchor="w", padx=20, pady=(15, 10))

        # Indicador para obligar a seleccionar acción
        self.lbl_action_hint = ctk.CTkLabel(
            self.sec1_box,
            text="⚠️ Por favor seleccione una de las 4 opciones para desplegar el formulario.",
            font=("Segoe UI", 12, "bold"),
            text_color="#D97706"
        )
        self.lbl_action_hint.pack(anchor="w", padx=20, pady=(0, 10))

        self.action_selector = ActionSelectorFrame(self.sec1_box, default_option=None, command_callback=self._on_action_selected)
        self.action_selector.pack(fill="x", padx=20, pady=(0, 18))

        # -------------------------------------------------------------
        # SECCIÓN 2: DATOS PRINCIPALES (OCULTOS HASTA ELEGIR ACCIÓN)
        # -------------------------------------------------------------
        self.sec2_box = ctk.CTkFrame(self.scroll_nueva, fg_color=CARD_BG, border_color=BORDER_COLOR, border_width=1, corner_radius=12)

        sec2_title = ctk.CTkLabel(self.sec2_box, text="2. Datos Principales", font=("Segoe UI", 15, "bold"), text_color=PRIMARY_NAVY)
        sec2_title.pack(anchor="w", padx=20, pady=(15, 10))

        grid_frame = ctk.CTkFrame(self.sec2_box, fg_color="transparent")
        grid_frame.pack(fill="x", padx=20, pady=(0, 15))
        grid_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        # Orden
        ctk.CTkLabel(grid_frame, text="Orden:", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).grid(row=0, column=0, sticky="w", padx=5, pady=4)
        self.entry_orden = ctk.CTkEntry(grid_frame, placeholder_text="ej. 6100029643", height=38, fg_color=INPUT_BG, border_color=BORDER_COLOR, text_color=TEXT_MAIN, placeholder_text_color="#64748B")
        self.entry_orden.grid(row=0, column=1, sticky="ew", padx=5, pady=4)

        # Centro de Costos - DESPLEGABLE CON BÚSQUEDA FLUIDA
        ctk.CTkLabel(grid_frame, text="Centro de Costos:", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).grid(row=0, column=2, sticky="w", padx=5, pady=4)
        cc_options = self.lists_data.get("CENTRO_COSTOS", ["TP08000"])
        self.combo_centro_costos = SearchableComboBox(grid_frame, values=cc_options, placeholder="Buscar Centro de Costos...", height=38)
        if cc_options:
            self.combo_centro_costos.set(cc_options[0])
        self.combo_centro_costos.grid(row=0, column=3, sticky="ew", padx=5, pady=4)

        # Área - DESPLEGABLE CON BÚSQUEDA FLUIDA
        ctk.CTkLabel(grid_frame, text="Área:", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).grid(row=1, column=0, sticky="w", padx=5, pady=4)
        area_options = self.lists_data.get("AREAS", ["Ensamble"])
        self.combo_area = SearchableComboBox(grid_frame, values=area_options, placeholder="Buscar Área...", height=38)
        if area_options:
            self.combo_area.set(area_options[0])
        self.combo_area.grid(row=1, column=1, sticky="ew", padx=5, pady=4)

        # Número de Proyecto - DESPLEGABLE CON BÚSQUEDA FLUIDA
        ctk.CTkLabel(grid_frame, text="Número de Proyecto:", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).grid(row=1, column=2, sticky="w", padx=5, pady=4)
        prj_options = self.lists_data.get("PROYECTOS", ["PRJ-1001"])
        self.combo_proyecto = SearchableComboBox(grid_frame, values=prj_options, placeholder="Buscar Número de Proyecto...", height=38)
        if prj_options:
            self.combo_proyecto.set(prj_options[0])
        self.combo_proyecto.grid(row=1, column=3, sticky="ew", padx=5, pady=4)

        # Nombre Solicitante
        ctk.CTkLabel(grid_frame, text="Nombre Solicitante:", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).grid(row=2, column=0, sticky="w", padx=5, pady=4)
        self.entry_nombre = ctk.CTkEntry(grid_frame, placeholder_text="ej. Paulina Razo", height=38, fg_color=INPUT_BG, border_color=BORDER_COLOR, text_color=TEXT_MAIN, placeholder_text_color="#64748B")
        self.entry_nombre.grid(row=2, column=1, sticky="ew", padx=5, pady=4)

        # Folio - AUTOMÁTICO (SOLO MODIFICABLE POR ADMIN CON CONTRASEÑA)
        ctk.CTkLabel(grid_frame, text="Folio (Automático):", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).grid(row=2, column=2, sticky="w", padx=5, pady=4)
        folio_box = ctk.CTkFrame(grid_frame, fg_color="transparent")
        folio_box.grid(row=2, column=3, sticky="ew", padx=5, pady=4)

        self.entry_folio = ctk.CTkEntry(folio_box, height=38, fg_color="#F1F5F9", border_color=BORDER_COLOR, text_color=PRIMARY_NAVY, font=("Segoe UI", 13, "bold"), state="disabled")
        self.entry_folio.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.entry_folio.bind("<Button-1>", self._on_locked_folio_click)

        self.btn_unlock_folio = ctk.CTkButton(
            folio_box, text="🔒", width=42, height=38, fg_color=PRIMARY_NAVY, hover_color=ACCENT_HOVER, text_color="#FFFFFF",
            font=("Segoe UI", 15), command=self._request_folio_unlock
        )
        self.btn_unlock_folio.pack(side="right")

        # Fecha CON CALENDARIO DESPLEGABLE CON SCROLL Y BLOQUEO DE PÁGINA
        ctk.CTkLabel(grid_frame, text="Fecha:", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).grid(row=3, column=0, sticky="w", padx=5, pady=4)
        fecha_box = ctk.CTkFrame(grid_frame, fg_color="transparent")
        fecha_box.grid(row=3, column=1, sticky="ew", padx=5, pady=4)

        self.entry_fecha = ctk.CTkEntry(fecha_box, height=38, fg_color=INPUT_BG, border_color=BORDER_COLOR, text_color=TEXT_MAIN)
        self.entry_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entry_fecha.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_cal_toggle = ctk.CTkButton(
            fecha_box, text="📅", width=42, height=38, fg_color=PRIMARY_NAVY, hover_color=ACCENT_HOVER, text_color="#FFFFFF",
            font=("Segoe UI", 15), command=self._toggle_dropdown_calendar
        )
        self.btn_cal_toggle.pack(side="right")

        # Calendario desplegable flotante con cuadrícula desplazable y bloqueo de página
        self.calendar_popup = DropdownCalendarPopup(self, self.entry_fecha, on_date_selected_callback=self._on_dropdown_date_selected)

        # -------------------------------------------------------------
        # SECCIÓN 3: TABLA DE MATERIALES (OCULTA HASTA ELEGIR ACCIÓN)
        # -------------------------------------------------------------
        self.sec3_box = ctk.CTkFrame(self.scroll_nueva, fg_color=CARD_BG, border_color=BORDER_COLOR, border_width=1, corner_radius=12)

        sec3_header = ctk.CTkFrame(self.sec3_box, fg_color="transparent")
        sec3_header.pack(fill="x", padx=20, pady=(15, 10))

        ctk.CTkLabel(sec3_header, text="3. Tabla de Materiales", font=("Segoe UI", 15, "bold"), text_color=PRIMARY_NAVY).pack(side="left")

        btn_add_material = ctk.CTkButton(
            sec3_header, text="+ Agregar Material", command=self.add_material_row,
            fg_color=SUCCESS_GREEN, hover_color=SUCCESS_HOVER, text_color=TEXT_ON_BLUE, font=("Segoe UI", 12, "bold"), height=38
        )
        btn_add_material.pack(side="right")

        # Encabezado permanente con nombres de columna claros
        tbl_guide_frame = ctk.CTkFrame(self.sec3_box, fg_color="#E2E8F0", corner_radius=6)
        tbl_guide_frame.pack(fill="x", padx=15, pady=(0, 8))

        ctk.CTkLabel(tbl_guide_frame, text="Material", font=("Segoe UI", 11, "bold"), text_color=TEXT_MAIN).pack(side="left", padx=15, pady=6, expand=True)
        ctk.CTkLabel(tbl_guide_frame, text="Batch / Lote", font=("Segoe UI", 11, "bold"), text_color=TEXT_MAIN).pack(side="left", padx=15, pady=6, expand=True)
        ctk.CTkLabel(tbl_guide_frame, text="Cantidad Requerida", font=("Segoe UI", 11, "bold"), text_color=TEXT_MAIN).pack(side="left", padx=15, pady=6, expand=True)
        ctk.CTkLabel(tbl_guide_frame, text="Descripción / Análisis", font=("Segoe UI", 11, "bold"), text_color=TEXT_MAIN).pack(side="left", padx=15, pady=6, expand=True)

        self.rows_container = ctk.CTkFrame(self.sec3_box, fg_color="transparent")
        self.rows_container.pack(fill="x", padx=15, pady=(0, 15))

        self.add_material_row()

        # Botones de Acción finales
        self.btn_box = ctk.CTkFrame(self.scroll_nueva, fg_color="transparent")

        # Botón Cancelar y Limpiar (Esquina inferior izquierda)
        self.btn_cancelar = ctk.CTkButton(
            self.btn_box, text="🧹 Cancelar y Limpiar", command=self.on_cancelar_limpiar_click,
            font=("Segoe UI", 15, "bold"), fg_color=DANGER_RED, hover_color=DANGER_HOVER, text_color="#FFFFFF", height=50, width=240
        )
        self.btn_cancelar.pack(side="left", padx=10)

        # Botón Generar Formulario Excel (Esquina inferior derecha)
        self.btn_generar = ctk.CTkButton(
            self.btn_box, text="🚀 Generar Formulario Excel", command=self.on_generar_click,
            font=("Segoe UI", 15, "bold"), fg_color=PRIMARY_NAVY, hover_color=ACCENT_HOVER, text_color=TEXT_ON_BLUE, height=50, width=280
        )
        self.btn_generar.pack(side="right", padx=10)

    def on_cancelar_limpiar_click(self):
        """
        Solicita confirmación al usuario antes de borrar toda la información ingresada y reiniciar el formulario.
        """
        confirm = messagebox.askokcancel(
            "Confirmar Cancelación",
            "⚠️ ¿Está seguro de que desea cancelar y limpiar el formulario?\n\n"
            "Se borrarán todos los datos ingresados actualmente y la pantalla regresará al inicio para una nueva captura.",
            icon="warning"
        )
        if confirm:
            self._reset_form_to_initial_state()
            self.status_label.configure(text="Formulario cancelado y limpiado.", text_color="#FFD54F")

    def _on_action_selected(self, option):
        """
        Al seleccionar una opción en la Sección 1, desdobla el resto del formulario.
        """
        if option is None:
            self.sec2_box.pack_forget()
            self.sec3_box.pack_forget()
            self.btn_box.pack_forget()
            self.lbl_action_hint.pack(anchor="w", padx=20, pady=(0, 10))
            return

        self.lbl_action_hint.pack_forget()

        if not self.sec2_box.winfo_ismapped():
            self.sec2_box.pack(fill="x", pady=(0, 15), padx=5)
            self.sec3_box.pack(fill="x", pady=(0, 15), padx=5)
            self.btn_box.pack(fill="x", pady=15, padx=5)

    def _clear_all_material_rows(self):
        """
        Elimina todas las filas de material programáticamente sin lanzar avisos emergentes.
        """
        for row in self.material_rows[:]:
            try:
                row.destroy()
            except Exception:
                pass
        self.material_rows.clear()

    def _reset_form_to_initial_state(self):
        self.is_modifying_folio = False
        """
        Restablece el formulario a su estado inicial replegando Secciones 2 y 3 y desplazando la vista al inicio superior.
        """
        self.action_selector.set(None)
        self._on_action_selected(None)

        self.entry_orden.delete(0, "end")
        self.entry_nombre.delete(0, "end")

        self._clear_all_material_rows()
        self.add_material_row()

        self._set_auto_folio()

        # REPOSICIONAR EL SCROLL OBLIGATORIAMENTE EN LA PARTE SUPERIOR DONDE ESTÁ LA SECCIÓN 1
        try:
            self.scroll_nueva._parent_canvas.yview_moveto(0.0)
        except Exception:
            pass

        self.update_idletasks()

    def _toggle_dropdown_calendar(self):
        self.calendar_popup.toggle()

    def _on_dropdown_date_selected(self, date_str):
        self.entry_fecha.delete(0, "end")
        self.entry_fecha.insert(0, date_str)

    def _on_locked_folio_click(self, event=None):
        if self.entry_folio.cget("state") == "disabled":
            self._request_folio_unlock()

    def _request_folio_unlock(self):
        if self.entry_folio.cget("state") == "normal":
            new_folio = self.entry_folio.get().strip().upper()
            self.entry_folio.configure(state="disabled", fg_color="#F1F5F9")
            self.btn_unlock_folio.configure(text="🔒")
            
            if new_folio:
                # Update FOLIO_INICIAL
                lists_data = config.load_lists_data()
                lists_data["FOLIO_INICIAL"] = new_folio
                config.save_lists_data(lists_data)
                
                # Reset max_generated_seq to the new base
                cfg = config.load_config()
                import re
                match = re.match(r"^([A-Z]+)(\d+)$", new_folio)
                if match:
                    cfg["max_generated_seq"] = int(match.group(2)) - 1
                cfg["ignore_sequence_history"] = True
                config.save_config(cfg)
                
                self.status_label.configure(text=f"Secuencia de Folio reiniciada a {new_folio}.", text_color="#90CAF9")
            else:
                self.status_label.configure(text="Folio bloqueado.", text_color="#90CAF9")
            return

        PasswordDialog(
            self,
            on_success_callback=self._unlock_folio_entry,
            on_cancel_callback=lambda: self.status_label.configure(text="Desbloqueo de Folio cancelado.", text_color="#90CAF9")
        )

    def _unlock_folio_entry(self):
        self.entry_folio.configure(state="normal", fg_color="#FFFFFF")
        self.entry_folio.focus_set()
        self.btn_unlock_folio.configure(text="🔓")
        messagebox.showinfo("Folio Desbloqueado", "El campo de Folio ha sido desbloqueado para edición por el administrador.")
        self.status_label.configure(text="Folio desbloqueado para edición de administrador.", text_color="#FFD54F")

    def _set_auto_folio(self):
        auto_folio = config.get_next_auto_folio()
        self.entry_folio.configure(state="normal")
        self.entry_folio.delete(0, "end")
        self.entry_folio.insert(0, auto_folio)
        self.entry_folio.configure(state="disabled", fg_color="#F1F5F9")
        if hasattr(self, "btn_unlock_folio"):
            self.btn_unlock_folio.configure(text="🔒")

    def _build_sub_tab_existente(self):
        box_folios = ctk.CTkFrame(self.sub_tab_existente, fg_color=CARD_BG, border_color=BORDER_COLOR, border_width=1, corner_radius=12)
        box_folios.pack(fill="both", expand=True, pady=10, padx=5)

        hdr_folios = ctk.CTkFrame(box_folios, fg_color="transparent")
        hdr_folios.pack(fill="x", padx=20, pady=(15, 10))

        ctk.CTkLabel(hdr_folios, text="✏️ Modificar Existente - Lista de Folios Generados", font=("Segoe UI", 16, "bold"), text_color=PRIMARY_NAVY).pack(side="left")

        ctk.CTkButton(
            hdr_folios, text="🔄 Actualizar Lista", width=130, height=34, command=self.refresh_folios_list,
            fg_color=ACCENT_BLUE, hover_color=ACCENT_HOVER, text_color="#FFFFFF"
        ).pack(side="right")

        ctk.CTkLabel(
            box_folios,
            text="Haga clic en '📝 Cargar para Modificar' en cualquier folio para transferir sus datos a la pestaña Nueva Creación o '🗑️' para eliminarlo.",
            font=("Segoe UI", 12), text_color=TEXT_SUB, justify="left"
        ).pack(anchor="w", padx=20, pady=(0, 10))

        # Contenedor con scroll y fondo sólido PAGE_BG para evitar smearing
        self.folios_scroll_container = ctk.CTkScrollableFrame(box_folios, fg_color=PAGE_BG, bg_color=PAGE_BG, border_color=BORDER_COLOR, border_width=1)
        self.folios_scroll_container.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # Pie de ventana con botón de borrado de historial completo (Esquina inferior izquierda)
        footer_folios = ctk.CTkFrame(box_folios, fg_color="transparent")
        footer_folios.pack(fill="x", padx=20, pady=(0, 15))

        btn_clear_history = ctk.CTkButton(
            footer_folios, text="🗑️ Borrar Historial de Folios", font=("Segoe UI", 12, "bold"),
            fg_color=DANGER_RED, hover_color=DANGER_HOVER, text_color="#FFFFFF", height=38, width=220,
            command=self._on_clear_history_click
        )
        btn_clear_history.pack(side="left")

        self.refresh_folios_list()

    def refresh_folios_list(self):
        for child in self.folios_scroll_container.winfo_children():
            child.destroy()

        cfg = config.load_config()
        folios_dict = cfg.get("generated_folios", {})

        if not folios_dict:
            ctk.CTkLabel(self.folios_scroll_container, text="No hay folios registrados hasta el momento.", font=("Segoe UI", 12, "italic"), text_color=TEXT_SUB).pack(pady=40)
            return

        sorted_folios = sorted(folios_dict.values(), key=lambda x: x.get("timestamp", 0), reverse=True)

        for item in sorted_folios:
            folio_code = item.get("folio", "N/A")
            orden = item.get("orden", "-")
            fecha = item.get("fecha", "-")
            accion = item.get("accion", "-")

            card = ctk.CTkFrame(self.folios_scroll_container, fg_color="#FFFFFF", border_color=BORDER_COLOR, border_width=1, corner_radius=8)
            card.pack(fill="x", pady=5, padx=5)

            info_frame = ctk.CTkFrame(card, fg_color="transparent")
            info_frame.pack(side="left", padx=15, pady=10)

            ctk.CTkLabel(info_frame, text=f"Folio: {folio_code}", font=("Segoe UI", 14, "bold"), text_color=PRIMARY_NAVY).pack(anchor="w")
            ctk.CTkLabel(info_frame, text=f"Orden: {orden}  |  Fecha: {fecha}  |  Acción: {accion}", font=("Segoe UI", 11), text_color=TEXT_SUB).pack(anchor="w")

            ctrl_frame = ctk.CTkFrame(card, fg_color="transparent")
            ctrl_frame.pack(side="right", padx=15, pady=10)

            # Botón Bote de Basura (Eliminar Folio Individual)
            btn_delete_folio = ctk.CTkButton(
                ctrl_frame, text="🗑️", font=("Segoe UI", 14), fg_color=DANGER_RED, hover_color=DANGER_HOVER, text_color="#FFFFFF",
                width=40, height=36, command=lambda f=folio_code: self._confirm_and_delete_single_folio(f)
            )
            btn_delete_folio.pack(side="right", padx=(8, 0))

            # Botón Cargar para Modificar
            btn_load = ctk.CTkButton(
                ctrl_frame, text="📝 Cargar para Modificar", font=("Segoe UI", 12, "bold"), fg_color=PRIMARY_NAVY, hover_color=ACCENT_HOVER,
                text_color="#FFFFFF", width=190, height=36, command=lambda f=folio_code: self._confirm_and_load_folio(f)
            )
            btn_load.pack(side="right")

    def _confirm_and_delete_single_folio(self, folio_code):
        confirm = messagebox.askokcancel(
            "Confirmar Eliminación de Folio",
            f"🗑️ ¿Está seguro de que desea eliminar el Folio '{folio_code}'?\n\n"
            f"⚠️ Atención: El número de folio '{folio_code}' se eliminará del registro activo y no volverá a ser reutilizado.",
            icon="warning"
        )
        if confirm:
            if config.delete_folio_record(folio_code):
                self.refresh_folios_list()
                self._set_auto_folio()
                self.status_label.configure(text=f"Folio '{folio_code}' eliminado correctamente.", text_color="#FFD54F")
            else:
                messagebox.showerror("Error", f"No se pudo eliminar el folio '{folio_code}'.")

    def _on_clear_history_click(self):
        """
        Abre el modal de contraseña del administrador para autorizar el borrado completo del historial.
        """
        PasswordDialog(
            self,
            on_success_callback=self._confirm_and_clear_entire_history,
            on_cancel_callback=lambda: self.status_label.configure(text="Borrado de historial cancelado.", text_color="#90CAF9")
        )

    def _confirm_and_clear_entire_history(self):
        """
        Muestra la advertencia/confirmación final antes de vaciar todo el historial de folios.
        """
        confirm = messagebox.askokcancel(
            "ADVERTENCIA: Borrar Historial Completo",
            "⚠️ ¿Está seguro de que desea borrar PERMANENTEMENTE todo el historial de folios registrados?\n\n"
            "• Todos los folios creados hasta la fecha serán eliminados de la lista.\n"
            "• Se conservará la secuencia de números para evitar duplicar folios futuros.\n"
            "• Esta acción NO se puede deshacer.\n\n"
            "¿Desea proceder con la eliminación completa?",
            icon="warning"
        )
        if confirm:
            config.clear_all_folios_history()
            self.refresh_folios_list()
            self._set_auto_folio()
            messagebox.showinfo("Historial Eliminado", "El historial de folios ha sido eliminado exitosamente por el administrador.")
            self.status_label.configure(text="Historial completo de folios eliminado.", text_color="#FFD54F")

    def _confirm_and_load_folio(self, folio_code):
        record = config.get_folio_record(folio_code)
        if not record:
            messagebox.showerror("Error", "No se encontraron los datos del folio seleccionado.")
            return

        orden = record.get("orden", "-")
        fecha = record.get("fecha", "-")

        # CONFIRMACIÓN AL USUARIO CON CAMBIO AUTOMÁTICO DE PANTALLA
        confirm = messagebox.askokcancel(
            "Confirmación de Carga de Folio",
            f"📋 ¿Deseas cargar la información del Folio '{folio_code}'?\n\n"
            f"Detalles:\n"
            f"• Orden: {orden}\n"
            f"• Fecha: {fecha}\n\n"
            f"Al hacer clic en Aceptar:\n"
            f"1. Todos los datos del folio se cargarán en el formulario.\n"
            f"2. La pantalla cambiará automáticamente a la pestaña '✨ Nueva Creación' para que puedas modificar y regenerar el archivo."
        )

        if confirm:
            self._load_folio_into_form(record, folio_code)

    def _load_folio_into_form(self, record, folio_code):
        self.is_modifying_folio = True
        data = record.get("data", {})
        self.entry_orden.delete(0, "end")
        self.entry_orden.insert(0, data.get("orden", ""))

        self.combo_centro_costos.set(data.get("centro_costos", ""))
        self.combo_area.set(data.get("area", ""))
        self.combo_proyecto.set(data.get("proyecto", ""))

        self.entry_nombre.delete(0, "end")
        self.entry_nombre.insert(0, data.get("nombre", ""))

        self.entry_folio.configure(state="normal")
        self.entry_folio.delete(0, "end")
        self.entry_folio.insert(0, folio_code)
        self.entry_folio.configure(state="disabled", fg_color="#F1F5F9")
        if hasattr(self, "btn_unlock_folio"):
            self.btn_unlock_folio.configure(text="🔒")

        self.entry_fecha.delete(0, "end")
        self.entry_fecha.insert(0, data.get("fecha", datetime.now().strftime("%d/%m/%Y")))

        # Seleccionar acción y desdoblar formulario
        accion_val = data.get("accion", "Scrap")
        self.action_selector.set(accion_val)
        self._on_action_selected(accion_val)

        # Cargar materiales limpiando sin avisos
        self._clear_all_material_rows()

        mats = data.get("materiales", [])
        if mats:
            for item in mats:
                self.add_material_row(
                    mat=item.get("material", ""), 
                    cant=item.get("cantidad", ""), 
                    batch=item.get("batch", ""),
                    desc=item.get("descripcion", "")
                )
        else:
            self.add_material_row()

        # Cambiar pantalla automáticamente a '✨ Nueva Creación'
        self.sub_tabview.set("✨ Nueva Creación")
        self._update_tab_button_contrast(self.sub_tabview)
        self.status_label.configure(text=f"Folio '{folio_code}' cargado para modificación.", text_color="#90CAF9")

    def _build_tab_config(self):
        # Fondo explícito PAGE_BG para evitar smearing
        scroll = ctk.CTkScrollableFrame(self.tab_config, fg_color=PAGE_BG, bg_color=PAGE_BG)
        scroll.pack(fill="both", expand=True, padx=5, pady=5)

        # -------------------------------------------------------------
        # SECCIÓN 1: ADMINISTRACIÓN DE ARCHIVO DE LISTAS DE CONFIGURACIÓN (.txt)
        # -------------------------------------------------------------
        box_txt = ctk.CTkFrame(scroll, fg_color=CARD_BG, border_color=BORDER_COLOR, border_width=1, corner_radius=12)
        box_txt.pack(fill="x", pady=(0, 15), padx=5)

        title_txt = ctk.CTkLabel(box_txt, text="📄 Gestión de Listas del Sistema (Archivo .txt)", font=("Segoe UI", 16, "bold"), text_color=PRIMARY_NAVY)
        title_txt.pack(anchor="w", padx=20, pady=(15, 5))

        sub_txt = ctk.CTkLabel(
            box_txt,
            text="Administre las listas de Centro de Costos, Áreas, Materiales, Proyectos y el Folio Inicial ([FOLIO_INICIAL]).",
            font=("Segoe UI", 11), text_color=TEXT_SUB
        )
        sub_txt.pack(anchor="w", padx=20, pady=(0, 10))

        f_txt_path = ctk.CTkFrame(box_txt, fg_color="transparent")
        f_txt_path.pack(fill="x", padx=20, pady=(0, 10))

        self.cfg_lists_file = ctk.CTkEntry(f_txt_path, height=38, fg_color=INPUT_BG, border_color=BORDER_COLOR, text_color=TEXT_MAIN)
        self.cfg_lists_file.insert(0, self.app_config.get("lists_file_path", config.DEFAULT_LISTS_FILE))
        self.cfg_lists_file.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(f_txt_path, text="Buscar...", width=95, command=self._browse_lists_file, fg_color=ACCENT_BLUE, hover_color=ACCENT_HOVER, text_color=TEXT_ON_BLUE).pack(side="right")

        self.txt_editor = ctk.CTkTextbox(box_txt, height=220, font=("Consolas", 12), fg_color="#F8FAFC", border_color=BORDER_COLOR, border_width=1, text_color=TEXT_MAIN)
        self.txt_editor.pack(fill="x", padx=20, pady=(0, 12))

        btn_txt_frame = ctk.CTkFrame(box_txt, fg_color="transparent")
        btn_txt_frame.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkButton(btn_txt_frame, text="🔄 Recargar .txt", command=self._load_txt_file_to_editor, fg_color="#64748B", hover_color="#475569", text_color="#FFFFFF", height=36).pack(side="left")
        ctk.CTkButton(btn_txt_frame, text="💾 Guardar Cambios en .txt", command=self._save_txt_editor_content, fg_color=SUCCESS_GREEN, hover_color=SUCCESS_HOVER, text_color="#FFFFFF", height=36).pack(side="right")

        # -------------------------------------------------------------
        # SECCIÓN 2: PARÁMETROS GENERALES (PLANTILLA Y GOOGLE SHEETS)
        # -------------------------------------------------------------
        box_gen = ctk.CTkFrame(scroll, fg_color=CARD_BG, border_color=BORDER_COLOR, border_width=1, corner_radius=12)
        box_gen.pack(fill="x", pady=10, padx=5)

        ctk.CTkLabel(box_gen, text="⚙️ Parámetros Generales de Plantilla y Conexiones", font=("Segoe UI", 16, "bold"), text_color=PRIMARY_NAVY).pack(anchor="w", padx=20, pady=15)

        ctk.CTkLabel(box_gen, text="Ruta de la Plantilla Excel (.xlsm):", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).pack(anchor="w", padx=20, pady=(10, 2))
        f_temp = ctk.CTkFrame(box_gen, fg_color="transparent")
        f_temp.pack(fill="x", padx=20, pady=(0, 10))
        self.cfg_template = ctk.CTkEntry(f_temp, height=38, fg_color=INPUT_BG, border_color=BORDER_COLOR, text_color=TEXT_MAIN)
        self.cfg_template.insert(0, self.app_config.get("template_path", "FORMATO SCRAP Y ADICIONALES.xlsm"))
        self.cfg_template.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_temp, text="Buscar...", width=95, command=self._browse_template, fg_color=ACCENT_BLUE, hover_color=ACCENT_HOVER, text_color=TEXT_ON_BLUE).pack(side="right")

        ctk.CTkLabel(box_gen, text="Nombre de la Hoja Excel:", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).pack(anchor="w", padx=20, pady=(10, 2))
        self.cfg_sheet_name = ctk.CTkEntry(box_gen, height=38, fg_color=INPUT_BG, border_color=BORDER_COLOR, text_color=TEXT_MAIN)
        self.cfg_sheet_name.insert(0, self.app_config.get("sheet_name", "SCRAP"))
        self.cfg_sheet_name.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(box_gen, text="Archivo de Credenciales Google (credentials.json):", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).pack(anchor="w", padx=20, pady=(10, 2))
        f_cred = ctk.CTkFrame(box_gen, fg_color="transparent")
        f_cred.pack(fill="x", padx=20, pady=(0, 10))
        self.cfg_credentials = ctk.CTkEntry(f_cred, height=38, fg_color=INPUT_BG, border_color=BORDER_COLOR, text_color=TEXT_MAIN)
        self.cfg_credentials.insert(0, self.app_config.get("credentials_path", "credentials.json"))
        self.cfg_credentials.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_cred, text="Buscar...", width=95, command=self._browse_credentials, fg_color=ACCENT_BLUE, hover_color=ACCENT_HOVER, text_color=TEXT_ON_BLUE).pack(side="right")

        ctk.CTkLabel(box_gen, text="Nombre o URL de la Hoja Google Sheets:", font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN).pack(anchor="w", padx=20, pady=(10, 2))
        self.cfg_gsheet = ctk.CTkEntry(box_gen, placeholder_text="ej. Registro_Formularios_Scrap o URL de la hoja", height=38, fg_color=INPUT_BG, border_color=BORDER_COLOR, text_color=TEXT_MAIN)
        self.cfg_gsheet.insert(0, self.app_config.get("google_sheet_url_or_name", ""))
        self.cfg_gsheet.pack(fill="x", padx=20, pady=(0, 10))

        # CHECKBOX DE SINCRONIZACIÓN AUTOMÁTICA EN CONFIGURACIÓN
        self.cfg_chk_sync_sheets = ctk.CTkCheckBox(
            box_gen, text="Sincronizar automáticamente con Google Sheets al generar formularios",
            font=("Segoe UI", 12, "bold"), text_color=PRIMARY_NAVY, fg_color=PRIMARY_NAVY, hover_color=ACCENT_HOVER, checkmark_color="#FFFFFF"
        )
        if self.app_config.get("auto_sync_google_sheets", True):
            self.cfg_chk_sync_sheets.select()
        else:
            self.cfg_chk_sync_sheets.deselect()
        self.cfg_chk_sync_sheets.pack(anchor="w", padx=20, pady=(5, 15))

        btn_save_cfg = ctk.CTkButton(
            box_gen, text="💾 Guardar Configuración General", command=self._save_configuration,
            font=("Segoe UI", 13, "bold"), fg_color=PRIMARY_NAVY, hover_color=ACCENT_HOVER, text_color=TEXT_ON_BLUE, height=40
        )
        btn_save_cfg.pack(anchor="e", padx=20, pady=(0, 20))

    def _browse_lists_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if filename:
            self.cfg_lists_file.delete(0, "end")
            self.cfg_lists_file.insert(0, filename)
            self._load_txt_file_to_editor()

    def _load_txt_file_to_editor(self):
        lists_path = self.cfg_lists_file.get().strip() or config.DEFAULT_LISTS_FILE
        if not os.path.exists(lists_path):
            config.create_default_lists_file(lists_path)

        try:
            with open(lists_path, "r", encoding="utf-8") as f:
                content = f.read()
            self.txt_editor.delete("1.0", "end")
            self.txt_editor.insert("1.0", content)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo .txt:\n{e}")

    def _save_txt_editor_content(self):
        lists_path = self.cfg_lists_file.get().strip() or config.DEFAULT_LISTS_FILE
        content = self.txt_editor.get("1.0", "end-1c")

        try:
            with open(lists_path, "w", encoding="utf-8") as f:
                f.write(content)

            self.app_config["lists_file_path"] = lists_path
            config.save_config(self.app_config)

            self.reload_lists_data()
            messagebox.showinfo("Éxito", "El archivo de listas .txt se ha guardado y aplicado correctamente.")
            self.status_label.configure(text="Listas actualizadas desde .txt.", text_color="#90CAF9")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo .txt:\n{e}")

    def reload_lists_data(self):
        self.lists_data = config.load_lists_data()

        cc = self.lists_data.get("CENTRO_COSTOS", ["TP08000"])
        current_cc = self.combo_centro_costos.get()
        self.combo_centro_costos.configure_values(cc)
        if cc:
            self.combo_centro_costos.set(current_cc if current_cc in cc else cc[0])

        ar = self.lists_data.get("AREAS", ["Ensamble"])
        current_ar = self.combo_area.get()
        self.combo_area.configure_values(ar)
        if ar:
            self.combo_area.set(current_ar if current_ar in ar else ar[0])

        pr = self.lists_data.get("PROYECTOS", ["PRJ-1001"])
        current_pr = self.combo_proyecto.get()
        self.combo_proyecto.configure_values(pr)
        if pr:
            self.combo_proyecto.set(current_pr if current_pr in pr else pr[0])

        mat_opts = self.lists_data.get("MATERIALES", ["Material Genérico"])
        for row in self.material_rows:
            row.update_material_options(mat_opts)

        self._set_auto_folio()

    def _browse_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsm *.xlsx")])
        if filename:
            self.cfg_template.delete(0, "end")
            self.cfg_template.insert(0, filename)

    def _browse_credentials(self):
        filename = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if filename:
            self.cfg_credentials.delete(0, "end")
            self.cfg_credentials.insert(0, filename)

    def _save_configuration(self):
        self.app_config["template_path"] = self.cfg_template.get().strip()
        self.app_config["sheet_name"] = self.cfg_sheet_name.get().strip()
        self.app_config["credentials_path"] = self.cfg_credentials.get().strip()
        self.app_config["google_sheet_url_or_name"] = self.cfg_gsheet.get().strip()
        self.app_config["auto_sync_google_sheets"] = bool(self.cfg_chk_sync_sheets.get())

        config.save_config(self.app_config)
        messagebox.showinfo("Configuración", "Los parámetros de configuración han sido guardados correctamente.")
        self.status_label.configure(text="Configuración guardada.", text_color="#90CAF9")

    def add_material_row(self, mat="", cant="", batch="", desc=""):
        if len(self.material_rows) >= 84:
            messagebox.showwarning("Límite alcanzado", "Se permite un máximo de 84 materiales por formulario.")
            return

        mat_options = self.lists_data.get("MATERIALES", ["Material Genérico"])
        idx = len(self.material_rows)
        row_frame = MaterialRowFrame(
            self.rows_container,
            index=idx,
            on_delete_callback=self.remove_material_row,
            materials_options=mat_options,
            default_mat=mat,
            default_cant=cant,
            default_batch=batch,
            default_desc=desc
        )
        row_frame.pack(fill="x", pady=4)
        self.material_rows.append(row_frame)
        self._reindex_rows()

    def remove_material_row(self, row_frame):
        if len(self.material_rows) <= 1:
            messagebox.showinfo("Información", "Debe existir al menos una fila de material.")
            return

        row_frame.destroy()
        self.material_rows.remove(row_frame)
        self._reindex_rows()

    def _reindex_rows(self):
        for idx, row in enumerate(self.material_rows):
            row.index = idx

    def collect_form_data(self):
        materiales = [r.get_data() for r in self.material_rows if r.get_data()["material"]]

        data = {
            "orden": self.entry_orden.get().strip(),
            "centro_costos": self.combo_centro_costos.get(),
            "area": self.combo_area.get(),
            "proyecto": self.combo_proyecto.get(),
            "nombre": self.entry_nombre.get().strip(),
            "folio": self.entry_folio.get().strip(),
            "fecha": self.entry_fecha.get().strip(),
            "accion": self.action_selector.get(),
            "materiales": materiales
        }
        return data

    def on_generar_click(self):

        if not getattr(self, "is_modifying_folio", False) and self.entry_folio.cget("state") == "disabled":
            real_auto_folio = config.get_and_reserve_next_auto_folio()
            self.entry_folio.configure(state="normal")
            self.entry_folio.delete(0, "end")
            self.entry_folio.insert(0, real_auto_folio)
            self.entry_folio.configure(state="disabled", fg_color="#F1F5F9")

        data = self.collect_form_data()

        # Verificar campos vacíos / incompletos
        empty_fields = []
        if not data["accion"]:
            empty_fields.append("1. Tipo de Formulario")
        if not data["orden"]:
            empty_fields.append("Orden")
        if not data["centro_costos"]:
            empty_fields.append("Centro de Costos")
        if not data["area"]:
            empty_fields.append("Área")
        if not data["proyecto"]:
            empty_fields.append("Número de Proyecto")
        if not data["nombre"]:
            empty_fields.append("Nombre Solicitante")
        if not data["folio"]:
            empty_fields.append("Folio")
        if not data["fecha"]:
            empty_fields.append("Fecha")

        if not self.material_rows:
            empty_fields.append("Tabla de Materiales (Sin materiales)")
        else:
            for idx, r in enumerate(self.material_rows):
                r_data = r.get_data()
                m_label = f"Material #{idx+1}"
                if not r_data["material"]:
                    empty_fields.append(f"Nombre de Material ({m_label})")
                if not r_data["cantidad"]:
                    empty_fields.append(f"Cantidad Requerida ({m_label})")
                if not r_data["descripcion"]:
                    empty_fields.append(f"Descripción / Análisis ({m_label})")

        if empty_fields:
            fields_str = "\n".join([f"  • {f}" for f in empty_fields[:8]])
            if len(empty_fields) > 8:
                fields_str += f"\n  • ...y {len(empty_fields) - 8} campos más."

            confirm = messagebox.askokcancel(
                "Información Faltante en Formulario",
                f"⚠️ El formulario contiene información incompleta / campos sin llenar:\n\n"
                f"{fields_str}\n\n"
                f"¿Desea continuar y generar el archivo de Excel con esta información faltante?",
                icon="warning"
            )
            if not confirm:
                self.status_label.configure(text="Generación pausada para completar campos.", text_color="#FFD54F")
                return

        template_path = self.cfg_template.get().strip() or self.app_config.get("template_path")
        sheet_name = self.cfg_sheet_name.get().strip() or self.app_config.get("sheet_name", "SCRAP")

        folio_code = data["folio"] or "SINFOLIO"
        fecha_clean = data["fecha"].replace("/", "-").replace("\\", "-")
        accion_str = data["accion"] or "FORMULARIO"

        default_filename = f"{folio_code}_{fecha_clean}.pdf"

        chosen_save_path = filedialog.asksaveasfilename(
            initialfile=default_filename,
            defaultextension=".pdf",
            filetypes=[("PDF Document", "*.pdf")],
            title="Seleccionar la ubicación para guardar el archivo modificado"
        )

        if not chosen_save_path:
            self.status_label.configure(text="Generación cancelada por el usuario.", text_color="#FFD54F")
            return

        try:
            saved_file = None
            engine_used = ""
            while True:
                try:
                    saved_file, engine_used = excel_handler.generate_pdf_form(data, template_path, chosen_save_path, sheet_name=sheet_name)
                    break
                except Exception as e:
                    ans = messagebox.askyesnocancel(
                        "Error al generar PDF",
                        f"Ocurrió un error al generar el PDF usando Excel:\n{str(e)}\n\n"
                        "¿Desea reintentar generar el PDF?\n\n"
                        "Seleccione 'Sí' para reintentar.\n"
                        "Seleccione 'No' para generar y guardar un archivo de Excel (.xlsm) en su lugar como respaldo.\n"
                        "Seleccione 'Cancelar' para abortar la generación."
                    )
                    if ans is True:
                        continue
                    elif ans is False:
                        excel_path = chosen_save_path.rsplit('.', 1)[0] + '.xlsm'
                        chosen_save_path = filedialog.asksaveasfilename(
                            initialfile=os.path.basename(excel_path),
                            defaultextension=".xlsm",
                            filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm")],
                            title="Seleccionar la ubicación para guardar el Excel"
                        )
                        if not chosen_save_path:
                            self.status_label.configure(text="Generación cancelada por el usuario.", text_color="#FFD54F")
                            return
                        saved_file, engine_used = excel_handler.generate_excel_form(data, template_path, chosen_save_path, sheet_name=sheet_name)
                        break
                    else:
                        self.status_label.configure(text="Generación cancelada tras error.", text_color="#FFD54F")
                        return

            # Manejo de estado y sincronización a Sheets con cola local
            auto_sync = self.app_config.get("auto_sync_google_sheets", True)
            synced_ok = False
            sheets_msg = "Sincronización deshabilitada."
            
            if auto_sync:
                creds_path = self.cfg_credentials.get().strip() or self.app_config.get("credentials_path")
                gsheet_name = self.cfg_gsheet.get().strip() or self.app_config.get("google_sheet_url_or_name")
                ok, sheets_msg = sheets_handler.sync_to_google_sheets(data, credentials_path=creds_path, sheet_url_or_name=gsheet_name)
                synced_ok = ok
            
            if folio_code:
                config.register_folio(folio_code, data["orden"], data["fecha"], data["accion"], saved_file, full_data=data, synced=synced_ok)
            
            # Lanzar tarea en segundo plano por si hay otros pendientes
            self._run_google_sheets_sync_bg()

            msg = f"✓ Archivo generado y registrado exitosamente en:\n{saved_file}"
            if engine_used == "openpyxl":
                msg = f"⚠️ ATENCIÓN: Se usó openpyxl (Excel falló).\nEl formato podría estar roto.\nGuardado en:\n{saved_file}"

            if auto_sync:
                if synced_ok:
                    msg += f"\n\n✓ Google Sheets: {sheets_msg}"
                    self.status_label.configure(text="Generado y registrado en Google Sheets.", text_color="#90CAF9")
                else:
                    msg += f"\n\n⚠️ Google Sheets aviso: {sheets_msg}\n(El formulario ha sido encolado para reintentarse automáticamente)."
                    self.status_label.configure(text=f"Generado. (Sheets: encolado para reintento)", text_color="#FFD54F")

            self._reset_form_to_initial_state()
            self.refresh_folios_list()

            if os.name == 'nt':
                os.startfile(saved_file)
            else:
                subprocess.run(["xdg-open", saved_file])

            messagebox.showinfo(
                "Formulario Generado",
                f"{msg}\n\nSe ha abierto automáticamente el archivo para su revisión e impresión."
            )

        except Exception as e:
            messagebox.showerror("Error al procesar", f"Ocurrió un error inesperado al generar el archivo:\n{str(e)}")
            self.status_label.configure(text=f"Error: {str(e)}", text_color="#EF5350")


if __name__ == "__main__":
    app = AppFormulario()
    app.mainloop()
