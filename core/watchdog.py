import psutil
import subprocess
import os
import logging

def _iniciar(nome, caminho, args=None):
    if not caminho or not os.path.exists(caminho):
        logging.error(f"[Watchdog] Executável {nome} não encontrado: {caminho}")
        return
    try:
        subprocess.Popen([caminho] + (args or []))
        logging.info(f"[Watchdog] {nome} iniciado com sucesso.")
    except Exception as e:
        logging.error(f"[Watchdog] Falha ao abrir {nome}: {e}")

def _resolver_caminho(candidatos):
    return next((c for c in candidatos if c and os.path.exists(c)), None)

def watchdog_infraestrutura():
    logging.info("[Watchdog] Verificando infraestrutura...")

    try:
        ativos = {p.name() for p in psutil.process_iter()}
    except Exception as e:
        logging.error(f"[Watchdog] Erro ao listar processos: {e}")
        return

    if not ativos & {"AnyDesk.exe", "SAP Logon.exe"}:
        logging.warning("[Watchdog] AnyDesk não encontrado. Tentando reabrir...")
        od = os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive")
        caminho = os.path.join(od, "SESÉ DASHBOARD", "Anchieta Dados",
                               "000 - Dashboard Dados", ".shared", "SAP Logon.exe") if od else None
        _iniciar("AnyDesk", caminho)
    if "OneDrive.exe" not in ativos:
        logging.warning("[Watchdog] OneDrive não encontrado. Reabrindo...")
        local = os.environ.get("LOCALAPPDATA")
        caminho = _resolver_caminho([
            os.path.join(local, "Microsoft", "OneDrive", "OneDrive.exe") if local else None,
            r"C:\Program Files\Microsoft OneDrive\OneDrive.exe"
        ])
        _iniciar("OneDrive", caminho, ["/background"])