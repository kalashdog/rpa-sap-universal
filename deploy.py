import os
import sys
import json
import subprocess
import shutil
import winreg
import glob

def get_onedrive_updates_folder():
    """Mesma lógica blindada do Launcher para achar o OneDrive corporativo"""
    possible_paths = set()
    try:
        base_key = r"Software\Microsoft\OneDrive\Accounts"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_key) as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                with winreg.OpenKey(key, winreg.EnumKey(key, i)) as subkey:
                    try:
                        folder, _ = winreg.QueryValueEx(subkey, "UserFolder")
                        if folder and os.path.exists(folder): possible_paths.add(folder)
                    except FileNotFoundError: pass
    except Exception: pass

    for env in ["OneDriveCommercial", "OneDrive", "OneDriveConsumer"]:
        if os.environ.get(env) and os.path.exists(os.environ.get(env)):
            possible_paths.add(os.environ.get(env))
            
    for path in possible_paths:
        sese_path = os.path.join(path, "SESÉ DASHBOARD")
        if os.path.exists(sese_path):
            return os.path.join(sese_path, "002 - Filiais database", "000 - Global", ".rpa_update")
    return None

def main():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("="*60)
    print("   ASSISTENTE DE DEPLOY - HUB SESÉ RPA")
    print("="*60)
    
    # 1. Coleta de Informações Humanas
    version = input("\n[1] Digite a nova versão (ex: 2.3.4): ").strip()
    if not version:
        print("Operação cancelada.")
        return
        
    message = input("[2] O que mudou nesta versão? (Release Notes): ").strip()
    if not message:
        message = "Atualização de rotina e melhorias de estabilidade."
        
    filename = f"HubSese_v{version}.exe"
    exe_name_no_ext = filename.replace(".exe", "")
    
    # 2. Limpeza
    print("\n[3] Limpando cache antigo...")
    for d in ['build', 'dist']:
        if os.path.exists(d): shutil.rmtree(d)

    # 3. Compilação Blindada (Apenas o GUI)
    print(f"\n[4] Compilando {filename} (Isto pode levar 1 minuto)...")
    cmd = [
        "pyinstaller", "--noconfirm", "--onefile", "--windowed",
        "--icon=.assets\\rpaseselogo_perfect.ico",
        f"--name={exe_name_no_ext}",
        "--add-data=.assets;.assets",
        "--add-data=config;config",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=keyring.backends.Windows",
        "--collect-all=customtkinter",
        "gui.py"
    ]
    
    # Executa silenciosamente, só mostra erro se falhar
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print("\n❌ ERRO NA COMPILAÇÃO:")
        print(result.stderr)
        return
        
    print("    ✅ Compilação concluída com sucesso (25MB garantidos).")

    # 4. Localização do OneDrive
    updates_dir = get_onedrive_updates_folder()
    if not updates_dir or not os.path.exists(updates_dir):
        print("\n❌ ERRO: Pasta oculta do OneDrive (.rpa_update) não encontrada!")
        print("O ficheiro foi compilado na pasta 'dist', mas não foi enviado para a nuvem.")
        return

    # 5. O Deploy Físico
    print(f"\n[5] Enviando para a nuvem da SESÉ ({updates_dir})...")
    src_exe = os.path.join("dist", filename)
    dest_exe = os.path.join(updates_dir, filename)
    
    shutil.copy2(src_exe, dest_exe)
    print(f"    ✅ Ficheiro {filename} copiado.")

    # 6. Atualização do Oráculo (JSON)
    json_path = os.path.join(updates_dir, "update_info.json")
    json_data = {
        "version": version,
        "filename": filename,
        "status": "ATIVO",
        "message": message
    }
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=4, ensure_ascii=False)
    print("    ✅ Ficheiro update_info.json atualizado.")

    print("\n" + "="*60)
    print(f"  DEPLOY DA VERSÃO {version} REALIZADO COM SUCESSO!")
    print(" Os operadores já receberão esta versão no próximo clique.")
    print("\n[7] Limpando vestígios da compilação (.spec, build, dist)...")
    try:
        if os.path.exists("build"): shutil.rmtree("build")
        if os.path.exists("dist"): shutil.rmtree("dist")
        for spec_file in glob.glob("*.spec"):
            os.remove(spec_file)
        print("    ✅ Projeto limpo.")
    except Exception as e:
        print(f"    ⚠️ Aviso ao limpar ficheiros: {e}")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()