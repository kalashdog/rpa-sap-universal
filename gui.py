import os
import sys
import json
import threading
import webbrowser
import ctypes
import time
import subprocess
from datetime import datetime, timedelta
from tkinter import messagebox
import tkinter as tk

import customtkinter as ctk
import keyring
from PIL import Image, ImageTk
from core.orchestrator import run_plant
from config.settings import settings
from core.utils import get_onedrive_path

def get_asset_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# ═══════════════════════════════════════════════════════════
#  CONSTANTES GLOBAIS
# ═══════════════════════════════════════════════════════════
ctk.set_default_color_theme("blue")

APP_TITLE = "Hub Sesé • RPA Logística"

def get_app_version():
    try:
        app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        version_file = os.path.join(app_data, "HubSeseRPA", "current_version.txt")
        if os.path.exists(version_file):
            with open(version_file, "r", encoding="utf-8") as f:
                return f.read().strip()
    except Exception:
        pass
    return "Dev Build"

APP_VERSION = get_app_version()
EXPECTED_FOLDER = "SESÉ DASHBOARD"
PREFS_FILE = os.path.join(os.path.expanduser("~"), ".hub_sese_rpa_ui.json")

# Windows API Constants for Insomnia Mode
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

FONT = "Segoe UI"
FONT_MONO = "Consolas"

# ── Dimensões ──────────────────────────────────────────────
SIDEBAR_W = 350
MAIN_W = 380
WIN_W = SIDEBAR_W + MAIN_W + 36  # 36 = paddings
WIN_H = 640

SHAREPOINT_LINK = (
    "https://juliob50600527.sharepoint.com.mcas.ms/sites/"
    "SESEDASHBOARD/Shared%20Documents/Forms/AllItems.aspx"
)

# ── Paleta #0066F9 ─────────────────────────────────────────
ACCENT = "#0066F9"
ACCENT_HOVER = "#0052C7"
ACCENT_SOFT = ("#E8F1FF", "#0A2A5E")
ACCENT_TEXT = ("#0B4FCF", "#B9D4FF")

BG_APP = ("#EEF4FF", "#020817")
PANEL = ("#FFFFFF", "#0F172A")
PANEL_ALT = ("#F8FBFF", "#111827")

SIDEBAR_BG = ("#0B1220", "#030712")
SIDEBAR_MUTED = "#94A3B8"
SIDEBAR_CARD_BG = ("#111827", "#0B1220")
SIDEBAR_CARD_BORDER = "#1F2937"
SIDEBAR_BTN_HOVER = ("#1E293B", "#0F172A")

BORDER = ("#D9E6FF", "#1F2937")
TEXT = ("#0F172A", "#F8FAFC")
TEXT_MUTED = ("#64748B", "#94A3B8")

SECONDARY = ("#EAF1FF", "#1E293B")
SECONDARY_HOVER = ("#D7E6FF", "#334155")

SUCCESS_FG = ("#DCFCE7", "#14532D")
SUCCESS_TEXT = ("#166534", "#BBF7D0")
WARNING_FG = ("#FEF3C7", "#78350F")
WARNING_TEXT = ("#92400E", "#FDE68A")
WARNING_BORDER = ("#FDE68A", "#78350F")
DANGER = "#EF4444"
DANGER_HOVER = "#DC2626"
ERROR_FG = ("#FEE2E2", "#7F1D1D")
ERROR_TEXT = ("#B91C1C", "#FCA5A5")


# ═══════════════════════════════════════════════════════════
#  UTILITÁRIOS
# ═══════════════════════════════════════════════════════════
def load_prefs():
    try:
        if os.path.exists(PREFS_FILE):
            with open(PREFS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_prefs(p):
    try:
        with open(PREFS_FILE, "w", encoding="utf-8") as f:
            json.dump(p, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def theme_to_mode(label):
    return {"Sistema": "system", "Claro": "light", "Escuro": "dark"}.get(label, "dark")


def short_path(path, n=48):
    if not path:
        return "—"
    return path if len(path) <= n else f"…{path[-n:]}"


def safe_del_pwd(svc, usr):
    try:
        keyring.delete_password(svc, usr)
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════
#  CLASSE PRINCIPAL
# ═══════════════════════════════════════════════════════════
class RpaGUI(ctk.CTk):

    def __init__(self):
        self._prefs = load_prefs()
        ctk.set_appearance_mode(theme_to_mode(self._prefs.get("theme", "Escuro")))

        super().__init__()

        self.title(APP_TITLE)
        self.geometry(f"{WIN_W}x{WIN_H}")
        
        try:
            myappid = 'sese.rpa.logistica.v2' # ID único inventado por nós
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        # 2. Window Icon (Obrigatório ser .ico no CustomTkinter)
        try:
            icon_path = get_asset_path(".assets/rpaseselogo_perfect.ico") 
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

        self.resizable(False, False)
        self.configure(fg_color=BG_APP)
        self._center(WIN_W, WIN_H)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # estado
        self.stop_event = threading.Event()
        self.worker_thread = None
        self.onedrive_base = None
        self.caminho_onedrive = None
        self.env_ready = False
        self.current_pct = 0
        self.last_status = None
        self.execution_finished = False
        self.current_user = ""
        self.current_plant = ""
        self.password_visible = False
        self.timer_running = False
        self.start_time = None
        self.remember_var = tk.BooleanVar(value=True)
        self.autostart_var = tk.BooleanVar(value=self._prefs.get("autostart", False))

        self._reset_refs()
        self._build_shell()
        self._check_env(navigate=True)
        
        # Windows Insomnia Mode
        self._prevent_sleep()
        
        # Check for Unattended Start
        if "--autostart" in sys.argv:
            print("Autostart detetado. Aguardando 30s para estabilização da rede...")
            self.after(30000, self.trigger_auto_start)

    # ───────────────────────────────────────────────────────
    #  Refs dinâmicas
    # ───────────────────────────────────────────────────────
    def _reset_refs(self):
        for attr in (
            "combo_planta", "input_user", "input_pwd", "form_msg",
            "btn_start", "btn_toggle_pwd", "lbl_status", "lbl_percent",
            "lbl_timer", "progress_bar", "status_badge", "log_box",
            "btn_stop", "btn_back", "lbl_hint",
        ):
            setattr(self, attr, None)

    def _ok(self, w):
        try:
            return w is not None and w.winfo_exists()
        except Exception:
            return False

    def _running(self):
        return self.worker_thread is not None and self.worker_thread.is_alive()

    def _center(self, w, h):
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _on_close(self):
        if self._running():
            if not messagebox.askyesno(
                "Encerrar", "O robô está em execução.\nDeseja parar e fechar?"
            ):
                return
            self.stop_event.set()
        self.timer_running = False
        self._allow_sleep()
        self.destroy()

    def _prevent_sleep(self):
        """Injeta o comando de 'Tela Ativa' no Kernel do Windows"""
        try:
            if sys.platform == "win32":
                ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)
        except Exception as e:
            print(f"Erro ao bloquear suspensão: {e}")

    def _allow_sleep(self):
        """Devolve o controlo de energia ao Windows"""
        try:
            if sys.platform == "win32":
                ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        except Exception:
            pass
            
    def toggle_autostart(self):
        import win32com.client
        startup_folder = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs\Startup")
        shortcut_path = os.path.join(startup_folder, "HubSeseRPA.lnk")
        target_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)

        if self.autostart_var.get():
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target_path
                shortcut.Arguments = "--autostart"
                shortcut.WorkingDirectory = os.path.dirname(target_path)
                shortcut.IconLocation = target_path
                shortcut.save()
            except Exception as e:
                print(f"Erro ao criar atalho: {e}")
        else:
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)

        # Save preference
        self._prefs["autostart"] = self.autostart_var.get()
        save_prefs(self._prefs)

    def trigger_auto_start(self):
        import keyring
        if keyring.get_password("RPA_SESE_USER", "default"):
            # Trigger execution only if credentials exist
            self._start()

    # ═══════════════════════════════════════════════════════
    #  SHELL
    # ═══════════════════════════════════════════════════════
    def _build_shell(self):
        self.grid_columnconfigure(0, weight=0, minsize=SIDEBAR_W)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── SIDEBAR (larga) ──────────────────────────────
        self.sidebar = ctk.CTkFrame(
            self, width=SIDEBAR_W, corner_radius=0, fg_color=SIDEBAR_BG
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self._build_sidebar()

        # ── PAINEL DIREITO (compacto) ────────────────────
        self.main_panel = ctk.CTkFrame(self, fg_color="transparent")
        self.main_panel.grid(row=0, column=1, sticky="nsew", padx=(12, 18), pady=18)
        self.main_panel.grid_rowconfigure(0, weight=1)
        self.main_panel.grid_columnconfigure(0, weight=1)

        self.main_card = ctk.CTkFrame(
            self.main_panel, corner_radius=20,
            fg_color=PANEL, border_width=1, border_color=BORDER
        )
        self.main_card.grid(row=0, column=0, sticky="nsew")
        self.main_card.grid_rowconfigure(2, weight=1)
        self.main_card.grid_columnconfigure(0, weight=1)

        # Header
        hdr = ctk.CTkFrame(self.main_card, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 10))

        self.page_tag = ctk.CTkLabel(
            hdr, text="", font=(FONT, 10, "bold"), text_color=ACCENT
        )
        self.page_tag.pack(anchor="w")

        self.page_title = ctk.CTkLabel(
            hdr, text="", font=(FONT, 22, "bold"), text_color=TEXT
        )
        self.page_title.pack(anchor="w", pady=(2, 0))

        self.page_sub = ctk.CTkLabel(
            hdr, text="", font=(FONT, 11), text_color=TEXT_MUTED,
            wraplength=MAIN_W - 60, justify="left"
        )
        self.page_sub.pack(anchor="w", pady=(4, 0))

        ctk.CTkFrame(
            self.main_card, height=1, fg_color=BORDER
        ).grid(row=1, column=0, sticky="ew", padx=24)

        self.content = ctk.CTkFrame(self.main_card, fg_color="transparent")
        self.content.grid(row=2, column=0, sticky="nsew", padx=24, pady=(12, 20))

    # ── Sidebar ──────────────────────────────────────────

    def _build_sidebar(self):
        sb = self.sidebar
        wrap = SIDEBAR_W - 48  # padding 24 de cada lado

        # ─ Brand ─
        brand = ctk.CTkFrame(sb, fg_color="transparent")
        brand.pack(fill="x", padx=24, pady=(24, 16))

        try:
            logo_path = get_asset_path(".assets/Sesé_White.png")
            logo_img = ctk.CTkImage(
                light_image=Image.open(logo_path),
                dark_image=Image.open(logo_path),
                size=(160, 48) # Ajustado para ocupar melhor a largura
            )
            logo_label = ctk.CTkLabel(brand, image=logo_img, text="")
            logo_label.pack(anchor="w")
        except Exception:
            # Fallback if image fails
            badge = ctk.CTkLabel(
                brand, text="RPA SESÉ", width=48, height=48,
                corner_radius=14, fg_color=ACCENT,
                text_color="white", font=(FONT, 18, "bold")
            )
            badge.pack(anchor="w")

        ctk.CTkLabel(
            brand, text="Hub de Dashboards",
            font=(FONT, 22, "bold"), text_color="white"
        ).pack(anchor="w", pady=(12, 0))

        ctk.CTkLabel(
            brand, text="RPA SESÉ • SAP",
            font=(FONT, 12), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", pady=(2, 0))

        ctk.CTkLabel(
            brand,
            text=(
                "Painel centralizado para autenticação, "
                "execução e acompanhamento do robô de "
                "automação SAP com exportação em nuvem."
            ),
            wraplength=wrap, justify="left",
            font=(FONT, 12), text_color="#CBD5E1"
        ).pack(anchor="w", pady=(14, 0))

        # ─ Card ambiente ─
        env = ctk.CTkFrame(
            sb, corner_radius=16,
            fg_color=SIDEBAR_CARD_BG,
            border_width=1, border_color=SIDEBAR_CARD_BORDER
        )
        env.pack(fill="x", padx=24, pady=(8, 12))

        ctk.CTkLabel(
            env, text="AMBIENTE",
            font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", padx=14, pady=(14, 8))

        self.env_badge = ctk.CTkLabel(
            env, text="Verificando…",
            width=140, height=26, corner_radius=999,
            fg_color=("#1E293B", "#1E293B"),
            text_color="#E2E8F0", font=(FONT, 11, "bold")
        )
        self.env_badge.pack(anchor="w", padx=14)

        self.env_path = ctk.CTkLabel(
            env, text="Verificando OneDrive…",
            wraplength=wrap - 28, justify="left",
            font=(FONT, 11), text_color="#CBD5E1"
        )
        self.env_path.pack(anchor="w", padx=14, pady=(10, 14))

        # ─ Ações ─
        ctk.CTkLabel(
            sb, text="AÇÕES RÁPIDAS",
            font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", padx=24, pady=(4, 8))

        ctk.CTkButton(
            sb, text="🌐  Abrir SharePoint",
            height=36, corner_radius=12,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=(FONT, 12, "bold"),
            command=lambda: webbrowser.open_new_tab(SHAREPOINT_LINK)
        ).pack(fill="x", padx=24, pady=4)

        ctk.CTkButton(
            sb, text="🔄  Revalidar ambiente",
            height=36, corner_radius=12,
            fg_color="transparent", hover_color=SIDEBAR_BTN_HOVER,
            border_width=1, border_color="#334155",
            font=(FONT, 12, "bold"),
            command=lambda: self._check_env(navigate=not self._running())
        ).pack(fill="x", padx=24, pady=4)

        # ─ Tema ─
        ctk.CTkLabel(
            sb, text="TEMA",
            font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", padx=24, pady=(16, 8))

        seg = ctk.CTkSegmentedButton(
            sb, values=["Sistema", "Claro", "Escuro"],
            command=self._set_theme, corner_radius=12
        )
        seg.pack(fill="x", padx=24)
        seg.set(self._prefs.get("theme", "Escuro"))

        # ─ Footer ─
        ctk.CTkLabel(
            sb, text=f"v{APP_VERSION}  •  VINICIUS LIMA",
            font=(FONT, 10), text_color="#475569"
        ).pack(side="bottom", anchor="w", padx=24, pady=(8, 18))

    def _set_theme(self, sel):
        ctk.set_appearance_mode(theme_to_mode(sel))
        self._prefs["theme"] = sel
        save_prefs(self._prefs)

    # ═══════════════════════════════════════════════════════
    #  AMBIENTE
    # ═══════════════════════════════════════════════════════
    def _check_env(self, navigate=True):
        self.onedrive_base = get_onedrive_path()
        self.caminho_onedrive = None
        self.env_ready = False

        if self.onedrive_base:
            p = os.path.join(self.onedrive_base, EXPECTED_FOLDER)
            if os.path.exists(p):
                self.caminho_onedrive = p
                self.env_ready = True

        self._update_env()
        if navigate:
            (self._show_login if self.env_ready else self._show_setup)()

    def _update_env(self):
        if self.env_ready:
            self.env_badge.configure(
                text="✓  Pronto", fg_color=ACCENT_SOFT, text_color=ACCENT_TEXT
            )
            self.env_path.configure(
                text=f"Pasta validada:\n{short_path(self.caminho_onedrive, 48)}"
            )
        elif self.onedrive_base:
            self.env_badge.configure(
                text="⚠  Pendente", fg_color=WARNING_FG, text_color=WARNING_TEXT
            )
            self.env_path.configure(
                text=f"OneDrive encontrado, mas '{EXPECTED_FOLDER}' não existe."
            )
        else:
            self.env_badge.configure(
                text="✕  Sem OneDrive", fg_color=ERROR_FG, text_color=ERROR_TEXT
            )
            self.env_path.configure(text="Nenhum OneDrive encontrado.")

    # ═══════════════════════════════════════════════════════
    #  HELPERS VISUAIS
    # ═══════════════════════════════════════════════════════
    def _header(self, tag, title, sub):
        self.page_tag.configure(text=tag.upper())
        self.page_title.configure(text=title)
        self.page_sub.configure(text=sub)

    def _clear(self):
        for w in self.content.winfo_children():
            w.destroy()
        self._reset_refs()

    def _card(self, parent, alt=False, **kw):
        return ctk.CTkFrame(
            parent, corner_radius=14,
            fg_color=PANEL_ALT if alt else PANEL,
            border_width=1, border_color=BORDER, **kw
        )

    def _step(self, parent, n, title, desc):
        c = self._card(parent, alt=True)
        ctk.CTkLabel(
            c, text=n, width=26, height=26, corner_radius=999,
            fg_color=ACCENT, text_color="white", font=(FONT, 11, "bold")
        ).pack(anchor="w", padx=12, pady=(12, 8))
        ctk.CTkLabel(
            c, text=title, font=(FONT, 13, "bold"), text_color=TEXT
        ).pack(anchor="w", padx=12)
        ctk.CTkLabel(
            c, text=desc, wraplength=150, justify="left",
            font=(FONT, 10), text_color=TEXT_MUTED
        ).pack(anchor="w", padx=12, pady=(4, 12))
        return c

    def _mini_card(self, parent, title, value):
        c = self._card(parent, alt=True)
        ctk.CTkLabel(
            c, text=title.upper(),
            font=(FONT, 9, "bold"), text_color=TEXT_MUTED
        ).pack(anchor="w", padx=10, pady=(10, 4))
        ctk.CTkLabel(
            c, text=value or "—",
            font=(FONT, 12, "bold"), text_color=TEXT,
            wraplength=100, justify="left"
        ).pack(anchor="w", padx=10, pady=(0, 10))
        return c

    # ═══════════════════════════════════════════════════════
    #  TELA 1 — SETUP
    # ═══════════════════════════════════════════════════════
    def _show_setup(self):
        self._clear()
        self._header(
            "Configuração",
            "Prepare o ambiente",
            "Vincule o SharePoint ao OneDrive antes de usar o robô."
        )

        # Alerta
        alert = ctk.CTkFrame(
            self.content, corner_radius=14,
            fg_color=("#FFFBEB", "#1C1917"),
            border_width=1, border_color=WARNING_BORDER
        )
        alert.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(
            alert, text="⚠  Ambiente não configurado",
            font=(FONT, 13, "bold"), text_color=WARNING_TEXT
        ).pack(anchor="w", padx=14, pady=(12, 4))

        ctk.CTkLabel(
            alert,
            text=f"A pasta '{EXPECTED_FOLDER}' precisa existir no OneDrive.",
            wraplength=MAIN_W - 80, justify="left",
            font=(FONT, 11), text_color=WARNING_TEXT
        ).pack(anchor="w", padx=14, pady=(0, 12))

        # Steps 2×2
        grid = ctk.CTkFrame(self.content, fg_color="transparent")
        grid.pack(fill="x", pady=(0, 12))
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)

        self._step(grid, "1", "Abrir SharePoint",
                   "Abra a biblioteca no navegador."
                   ).grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        self._step(grid, "2", "Adicionar ao OneDrive",
                   "Clique 'Adicionar atalho'."
                   ).grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        self._step(grid, "3", "Validar nome",
                   f"Deve ser '{EXPECTED_FOLDER}'."
                   ).grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
        self._step(grid, "4", "Verificar",
                   "Volte e clique 'Já vinculei'."
                   ).grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))

        # Botões
        ctk.CTkButton(
            self.content, text="🌐   Abrir SharePoint",
            height=40, corner_radius=12,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=(FONT, 13, "bold"),
            command=lambda: webbrowser.open_new_tab(SHAREPOINT_LINK)
        ).pack(fill="x", pady=(8, 8))

        ctk.CTkButton(
            self.content, text="✓   Já vinculei, validar",
            height=40, corner_radius=12,
            fg_color=SECONDARY, hover_color=SECONDARY_HOVER,
            text_color=TEXT, font=(FONT, 13, "bold"),
            command=lambda: self._check_env(navigate=True)
        ).pack(fill="x")

    # ═══════════════════════════════════════════════════════
    #  TELA 2 — LOGIN
    # ═══════════════════════════════════════════════════════
    def _show_login(self):
        self._clear()
        self.password_visible = False

        self._header(
            "Autenticação SAP",
            "Iniciar automação",
            "Informe credenciais e selecione a planta."
        )

        # Info exportação (compacto)
        info = self._card(self.content, alt=True)
        info.pack(fill="x", pady=(0, 12))

        row_info = ctk.CTkFrame(info, fg_color="transparent")
        row_info.pack(fill="x", padx=14, pady=10)

        ctk.CTkLabel(
            row_info, text="✓  Exportação:",
            font=(FONT, 11, "bold"), text_color=ACCENT
        ).pack(side="left")

        ctk.CTkLabel(
            row_info, text=short_path(self.caminho_onedrive, 30),
            font=(FONT, 11), text_color=TEXT
        ).pack(side="left", padx=(6, 0))

        # ── Formulário ────────────────────────────────────
        form = ctk.CTkFrame(self.content, fg_color="transparent")
        form.pack(fill="x", pady=(4, 0))

        # Planta
        plants = list(settings.config.get("plants", {}).keys()) or ["Nenhuma"]

        ctk.CTkLabel(
            form, text="Planta",
            font=(FONT, 11, "bold"), text_color=TEXT_MUTED
        ).pack(anchor="w", pady=(0, 4))

        self.combo_planta = ctk.CTkComboBox(
            form, values=plants, height=36,
            corner_radius=10, border_width=1, border_color=BORDER,
            font=(FONT, 12), dropdown_font=(FONT, 12)
        )
        self.combo_planta.pack(anchor="w", fill="x", pady=(0, 10))

        last = self._prefs.get("last_plant")
        self.combo_planta.set(last if last in plants else plants[0])

        # Usuário
        ctk.CTkLabel(
            form, text="Usuário SAP",
            font=(FONT, 11, "bold"), text_color=TEXT_MUTED
        ).pack(anchor="w", pady=(0, 4))

        self.input_user = ctk.CTkEntry(
            form, placeholder_text="Ex.: FV2WL5N",
            height=36, corner_radius=10,
            border_width=1, border_color=BORDER, font=(FONT, 12)
        )
        self.input_user.pack(anchor="w", fill="x", pady=(0, 10))

        # Senha
        ctk.CTkLabel(
            form, text="Senha SAP",
            font=(FONT, 11, "bold"), text_color=TEXT_MUTED
        ).pack(anchor="w", pady=(0, 4))

        pwd_row = ctk.CTkFrame(form, fg_color="transparent")
        pwd_row.pack(anchor="w", fill="x", pady=(0, 8))

        self.input_pwd = ctk.CTkEntry(
            pwd_row, placeholder_text="Mínimo 12 caracteres",
            show="•", height=36, corner_radius=10,
            border_width=1, border_color=BORDER, font=(FONT, 12)
        )
        self.input_pwd.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.btn_toggle_pwd = ctk.CTkButton(
            pwd_row, text="Mostrar", width=72, height=36,
            corner_radius=10, fg_color=SECONDARY,
            hover_color=SECONDARY_HOVER, text_color=TEXT,
            font=(FONT, 11), command=self._toggle_pwd
        )
        self.btn_toggle_pwd.pack(side="right")

        # Pré-preencher
        saved_u = keyring.get_password("RPA_SESE_USER", "default")
        if saved_u:
            self.input_user.insert(0, saved_u)
            saved_p = keyring.get_password("RPA_SESE_PWD", saved_u)
            if saved_p:
                self.input_pwd.insert(0, saved_p)
                self.remember_var.set(True)
            else:
                self.remember_var.set(False)
        else:
            self.remember_var.set(True)

        # Remember
        ctk.CTkCheckBox(
            form, text="Lembrar credenciais",
            variable=self.remember_var, font=(FONT, 11)
        ).pack(anchor="w", pady=(0, 6))

        # Autostart
        ctk.CTkCheckBox(
            form, text="Iniciar automaticamente com o Windows",
            variable=self.autostart_var, font=(FONT, 11),
            command=self.toggle_autostart
        ).pack(anchor="w", pady=(0, 6))

        # Mensagem de feedback
        self.form_msg = ctk.CTkLabel(
            form, text="", font=(FONT, 11), text_color=TEXT_MUTED
        )
        self.form_msg.pack(anchor="w", pady=(0, 8))

        # Botão INICIAR
        self.btn_start = ctk.CTkButton(
            form, text="▶   Iniciar RPA",
            height=42, corner_radius=12,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=(FONT, 14, "bold"), command=self._start
        )
        self.btn_start.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(
            form, text="Ou pressione Enter no campo de senha.",
            font=(FONT, 10), text_color=TEXT_MUTED
        ).pack(anchor="w")

        self.input_pwd.bind("<Return>", lambda _: self._start())

        if "Nenhuma" in plants[0]:
            self.btn_start.configure(state="disabled")
            self._form_msg("Nenhuma planta configurada.", "warning")

        # Focus
        if not saved_u:
            self.input_user.focus()
        elif not self.input_pwd.get():
            self.input_pwd.focus()

    def _toggle_pwd(self):
        self.password_visible = not self.password_visible
        self.input_pwd.configure(show="" if self.password_visible else "•")
        self.btn_toggle_pwd.configure(
            text="Ocultar" if self.password_visible else "Mostrar"
        )

    def _form_msg(self, msg, kind="info"):
        if not self._ok(self.form_msg):
            return
        colors = {
            "info": TEXT_MUTED, "success": ACCENT,
            "warning": WARNING_TEXT, "error": ERROR_TEXT
        }
        self.form_msg.configure(text=msg, text_color=colors.get(kind, TEXT_MUTED))

    # ═══════════════════════════════════════════════════════
    #  TELA 3 — PROGRESSO
    # ═══════════════════════════════════════════════════════
    def _show_progress(self):
        self._clear()
        self.execution_finished = False

        self._header(
            "Execução",
            "Em andamento",
            "Acompanhe o progresso em tempo real."
        )

        # Mini-cards resumo
        row = ctk.CTkFrame(self.content, fg_color="transparent")
        row.pack(fill="x", pady=(0, 10))
        for c in range(3):
            row.grid_columnconfigure(c, weight=1)

        self._mini_card(row, "Planta", self.current_plant).grid(
            row=0, column=0, sticky="nsew", padx=(0, 3))
        self._mini_card(row, "Usuário", self.current_user).grid(
            row=0, column=1, sticky="nsew", padx=3)
        self._mini_card(row, "Export", "OneDrive").grid(
            row=0, column=2, sticky="nsew", padx=(3, 0))

        # Status
        sc = self._card(self.content, alt=True)
        sc.pack(fill="x", pady=(0, 10))

        top = ctk.CTkFrame(sc, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(12, 6))

        self.status_badge = ctk.CTkLabel(
            top, text="Em execução", width=110, height=24,
            corner_radius=999, fg_color=ACCENT_SOFT,
            text_color=ACCENT_TEXT, font=(FONT, 10, "bold")
        )
        self.status_badge.pack(side="left")

        self.lbl_timer = ctk.CTkLabel(
            top, text="⏱ 00:00:00",
            font=(FONT, 11), text_color=TEXT_MUTED
        )
        self.lbl_timer.pack(side="left", padx=(10, 0))

        self.lbl_percent = ctk.CTkLabel(
            top, text="0%",
            font=(FONT, 18, "bold"), text_color=TEXT
        )
        self.lbl_percent.pack(side="right")

        self.lbl_status = ctk.CTkLabel(
            sc, text="Inicializando SAP…",
            font=(FONT, 12), text_color=TEXT,
            wraplength=MAIN_W - 80, justify="left"
        )
        self.lbl_status.pack(anchor="w", padx=14, pady=(0, 8))

        self.progress_bar = ctk.CTkProgressBar(
            sc, height=6, corner_radius=3, progress_color=ACCENT
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=14, pady=(0, 12))

        # ── FOOTER (Botões) - Empacotado ANTES do Log para garantir espaço ──
        footer = ctk.CTkFrame(self.content, fg_color="transparent")
        footer.pack(side="bottom", fill="x", pady=(10, 0))

        self.lbl_hint = ctk.CTkLabel(
            footer, text="Processando passos da planta…",
            font=(FONT, 10), text_color=TEXT_MUTED
        )
        self.lbl_hint.pack(side="bottom", anchor="w", pady=(6, 0))

        self.btn_back = ctk.CTkButton(
            footer, text="Nova execução",
            height=36, corner_radius=12,
            fg_color=SECONDARY, hover_color=SECONDARY_HOVER,
            text_color=TEXT, state="disabled",
            font=(FONT, 12, "bold"), command=self._show_login
        )
        self.btn_back.pack(side="bottom", fill="x")

        self.btn_stop = ctk.CTkButton(
            footer, text="⏹   Parar",
            height=38, corner_radius=12,
            fg_color=DANGER, hover_color=DANGER_HOVER,
            font=(FONT, 13, "bold"), command=self._stop
        )
        self.btn_stop.pack(side="bottom", fill="x", pady=(0, 6))

        # ── LOG (Expande no espaço que sobrar no meio) ──
        lc = self._card(self.content, alt=True)
        lc.pack(side="top", fill="both", expand=True, pady=(0, 0))

        ctk.CTkLabel(
            lc, text="LOG", font=(FONT, 10, "bold"), text_color=TEXT_MUTED
        ).pack(anchor="w", padx=14, pady=(12, 6))

        self.log_box = ctk.CTkTextbox(
            lc, corner_radius=8, font=(FONT_MONO, 10),
            fg_color=("gray96", "#0d1117"),
            border_width=1, border_color=BORDER
        )
        self.log_box.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.log_box.configure(state="disabled")

    def _tick(self):
        """Atualiza o timer a cada segundo. Reprograma enquanto timer_running e janela existir."""
        if self.timer_running and self._ok(self.lbl_timer):
            s = int(time.time() - self.start_time)
            self.lbl_timer.configure(text=f"⏱ {timedelta(seconds=s)}")
            self.after(1000, self._tick)
        elif self.timer_running and not self._ok(self.lbl_timer):
            # Widget destruido (navegação de tela): para o timer de forma segura
            self.timer_running = False

    def _log(self, txt):
        if not self._ok(self.log_box):
            return
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}]  {txt}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ═══════════════════════════════════════════════════════
    #  STATUS CALLBACK
    # ═══════════════════════════════════════════════════════
    def update_status(self, text, pct):
        try:
            self.after(0, self._apply, text, pct)
        except Exception:
            pass

    def _apply(self, text, pct):
        try:
            pct = max(0, min(100, int(float(pct))))
        except Exception:
            pct = self.current_pct
        self.current_pct = pct

        if self._ok(self.lbl_status):
            self.lbl_status.configure(text=text)
        if self._ok(self.progress_bar):
            self.progress_bar.set(pct / 100.0)
        if self._ok(self.lbl_percent):
            self.lbl_percent.configure(text=f"{pct}%")

        if text and text != self.last_status:
            self._log(text)
            self.last_status = text

        low = (text or "").lower()

        if "erro" in low or "falha" in low or "exception" in low:
            # Erro transitório: atualiza badge mas NÃO trava o timer.
            # O _finish("error") só é chamado pelo _worker() quando o processo termina de verdade.
            self._badge("error", "Erro transient")
        elif "interromp" in low or "parada" in low:
            self._badge("warning", "Interrompido")
            self._finish("warning")
        elif pct >= 100:
            self._badge("success", "Ciclo Concluído")
        else:
            self._badge("info", "Em execução")

    def _badge(self, kind, text):
        if not self._ok(self.status_badge):
            return
        s = {
            "info":    (ACCENT_SOFT, ACCENT_TEXT),
            "success": (SUCCESS_FG,  SUCCESS_TEXT),
            "warning": (WARNING_FG,  WARNING_TEXT),
            "error":   (ERROR_FG,    ERROR_TEXT),
        }.get(kind, (ACCENT_SOFT, ACCENT_TEXT))
        self.status_badge.configure(text=text, fg_color=s[0], text_color=s[1])

    def _finish(self, kind):
        if self.execution_finished:
            return
        self.execution_finished = True
        self.timer_running = False

        if self._ok(self.btn_stop):
            self.btn_stop.configure(state="disabled")
        if self._ok(self.btn_back):
            labels = {
                "success": "▶  Nova execução",
                "warning": "↩  Voltar",
                "error":   "🔧  Tentar novamente"
            }
            self.btn_back.configure(
                state="normal", text=labels.get(kind, "Nova execução")
            )
        if self._ok(self.lbl_hint):
            hints = {
                "success": "Concluído com sucesso!",
                "warning": "Interrompido pelo usuário.",
                "error":   "Erro detectado — revise o log."
            }
            self.lbl_hint.configure(text=hints.get(kind, ""))

    # ═══════════════════════════════════════════════════════
    #  AÇÕES
    # ═══════════════════════════════════════════════════════
    def _start(self):
        if not self.env_ready:
            self._show_setup()
            return

        user = (self.input_user.get().strip() if self.input_user else "")
        pwd = (self.input_pwd.get() if self.input_pwd else "")
        planta = (self.combo_planta.get().strip() if self.combo_planta else "")

        if not user or not pwd:
            self._form_msg("Preencha usuário e senha.", "warning")
            return
        if len(user) != 7:
            self._form_msg("Usuário SAP: exatamente 7 caracteres.", "warning")
            return
        if len(pwd) < 12:
            self._form_msg("Senha: mínimo 12 caracteres.", "warning")
            return
        if not planta or planta == "Nenhuma":
            self._form_msg("Selecione uma planta válida.", "warning")
            return

        self._prefs["last_plant"] = planta
        save_prefs(self._prefs)

        stored = keyring.get_password("RPA_SESE_USER", "default")
        if self.remember_var.get():
            keyring.set_password("RPA_SESE_USER", "default", user)
            keyring.set_password("RPA_SESE_PWD", user, pwd)
        else:
            if stored:
                safe_del_pwd("RPA_SESE_PWD", stored)
            safe_del_pwd("RPA_SESE_USER", "default")
            safe_del_pwd("RPA_SESE_PWD", user)

        settings.dynamic_user = user
        settings.dynamic_pwd = pwd
        settings.export_base_path = self.caminho_onedrive

        self.current_user = user
        self.current_plant = planta
        self.current_pct = 0
        self.last_status = None
        self.stop_event.clear()

        self._show_progress()

        # Iniciar timer (deve ser feito APÓS _show_progress criar os widgets)
        self._log(f"Execução iniciada — planta '{self.current_plant}'.")
        
        # --- WAKE UP ETL ---
        try:
            etl_exe_path = os.path.join(self.caminho_onedrive, "002 - Filiais database", "006 - ETL", "HubSese_ETL.exe")
            if os.path.exists(etl_exe_path):
                # Assegura que matamos orquestradores em background rodando com plantas antigas
                os.system("taskkill /F /IM HubSese_ETL.exe /T 2>nul")
                
                import re
                # Converter "01-Anchieta" para "anchieta"
                clean_plant = re.sub(r'[\d\-]', '', self.current_plant).split()[0].strip().lower()
                
                # Acorda com a nova planta
                subprocess.Popen([etl_exe_path, clean_plant], creationflags=0x08000000)
                self._log(f"Processo ETL ativado em background com planta: {clean_plant}.")
            else:
                self._log("Aviso: Executável do ETL não foi encontrado neste caminho:")
                self._log(etl_exe_path)
        except Exception as e:
            self._log(f"Aviso: Não foi possível acordar o ETL: {e}")
        # -------------------
        self.start_time = time.time()
        self.timer_running = True
        self._tick()

        self.worker_thread = threading.Thread(
            target=self._worker, args=(planta,), daemon=True
        )
        self.worker_thread.start()

    def _worker(self, planta):
        try:
            run_plant(planta, self.update_status, self.stop_event)
            if self.stop_event.is_set():
                if not self.last_status or "interromp" not in self.last_status.lower():
                    self.update_status("Processo interrompido pelo usuário.", 0)
            else:
                self.update_status("Processo finalizado completamente.", 100)
                self.after(0, lambda: self._finish("success"))
                
        except Exception as e:
            self.update_status(f"Erro: {e}", self.current_pct)
            self.after(0, lambda: self._finish("error"))

    def _stop(self):
        if not self._running():
            return
        self.stop_event.set()
        if self._ok(self.lbl_status):
            self.lbl_status.configure(text="Sinalizando parada…")
        if self._ok(self.btn_stop):
            self.btn_stop.configure(state="disabled", text="⏳  Parando…")
        self._badge("warning", "A parar")
        self._log("Interrupção solicitada.")


if __name__ == "__main__":
    app = RpaGUI()
    app.mainloop()