import os
import sys
import csv
import urllib.request
import subprocess
import shutil
import tkinter as tk
from tkinter import messagebox

# ═══════════════════════════════════════════════════════════
# CONFIGURAÇÕES DO LAUNCHER
# ═══════════════════════════════════════════════════════════
# O link do seu Google Sheets publicado em CSV
CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRPbLJhf22LcVTlyJIV4U1wkJoaJ8wBChwH8MBCwK7LOhAmN2bGDoCmIMKLzZcd2kaNxpIP_38y2kZe/pub?gid=1948707045&single=true&output=csv"

# O Esconderijo local onde o robô real vai morar
APP_DATA = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
BIN_DIR = os.path.join(APP_DATA, "HubSeseRPA", "bin")
VERSION_FILE = os.path.join(APP_DATA, "HubSeseRPA", "current_version.txt")

os.makedirs(BIN_DIR, exist_ok=True)

# ═══════════════════════════════════════════════════════════
# FUNÇÕES DE APOIO
# ═══════════════════════════════════════════════════════════
def get_onedrive_updates_folder():
    possible_paths = set()
    for env in ["OneDriveCommercial", "OneDrive", "OneDriveConsumer"]:
        val = os.environ.get(env)
        if val and os.path.exists(val):
            possible_paths.add(val)
            
    for path in possible_paths:
        sese_path = os.path.join(path, "SESÉ DASHBOARD")
        if os.path.exists(sese_path):
            return os.path.join(sese_path, "002 - Filiais database", "000 - Global", ".rpa_update")
    return None

def show_msg(title, message, is_error=False):
    """Mostra um popup nativo do Windows"""
    root = tk.Tk()
    root.withdraw() # Esconde a janela principal do Tkinter
    if is_error:
        messagebox.showerror(title, message)
    else:
        messagebox.showinfo(title, message)
    root.destroy()

def get_current_version():
    if os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return "0.0.0"

def set_current_version(version):
    with open(VERSION_FILE, 'w', encoding='utf-8') as f:
        f.write(version)

# ═══════════════════════════════════════════════════════════
# O MOTOR DO LAUNCHER
# ═══════════════════════════════════════════════════════════
def main():
    try:
        # 1. Bate no Oráculo (Google Sheets)
        req = urllib.request.Request(CSV_URL)
        with urllib.request.urlopen(req, timeout=5) as response:
            lines = [l.decode('utf-8') for l in response.readlines()]
            
        reader = csv.reader(lines)
        next(reader) # Pula o cabeçalho
        cloud_data = next(reader)
        
        cloud_version = cloud_data[0].strip()
        cloud_filename = cloud_data[1].strip()
        cloud_status = cloud_data[2].strip().upper()
        cloud_message = cloud_data[3].strip()
        
        # 2. O Kill Switch (Manutenção Global)
        if cloud_status == "INATIVO" or cloud_status == "MANUTENÇÃO":
            show_msg("Sistema em Manutenção", f"O RPA está temporariamente suspenso.\n\nMotivo: {cloud_message}")
            sys.exit(0)
            
        # 3. A Lógica de Atualização
        local_version = get_current_version()
        local_exe_path = os.path.join(BIN_DIR, cloud_filename)
        
        if cloud_version != local_version or not os.path.exists(local_exe_path):
            # Precisa atualizar! Procura o arquivo no OneDrive
            updates_dir = get_onedrive_updates_folder()
            if not updates_dir:
                show_msg("Erro de Sincronização", "Pasta 'SESÉ DASHBOARD' não encontrada no OneDrive.", True)
                sys.exit(1)
                
            source_exe = os.path.join(updates_dir, cloud_filename)
            
            if not os.path.exists(source_exe):
                show_msg("Aguardando Download", f"Uma nova versão ({cloud_version}) foi detectada, mas o OneDrive ainda não terminou de baixar o arquivo.\n\nAguarde alguns instantes e tente novamente.")
                sys.exit(0)
                
            # Mostra aviso de atualização
            show_msg("Atualização Detectada", f"Baixando a versão {cloud_version} do Hub Sesé RPA...\n\nNovidades: {cloud_message}")
            
            # Copia o arquivo do OneDrive para o Esconderijo do AppData
            shutil.copy2(source_exe, local_exe_path)
            set_current_version(cloud_version)
            
        # 4. Executa o Robô Real e passa os argumentos (ex: --autostart)
        args = [local_exe_path] + sys.argv[1:]
        subprocess.Popen(args)
        
    except Exception as e:
        # Fallback offline: Se a internet cair, tenta abrir a última versão salva
        local_version = get_current_version()
        if local_version != "0.0.0":
            # Aqui você precisaria saber o nome do arquivo antigo, ou simplesmente procurar o único .exe na pasta bin.
            exes = [f for f in os.listdir(BIN_DIR) if f.endswith('.exe')]
            if exes:
                subprocess.Popen([os.path.join(BIN_DIR, exes[0])] + sys.argv[1:])
                sys.exit(0)
                
        show_msg("Erro de Conexão", f"Não foi possível contactar o servidor e não há versões locais salvas.\n\nDetalhe: {e}", True)

if __name__ == "__main__":
    main()