"""
Orquestrador inteligente - loop continuo com estado por job.
request -> espera -> extract -> repeat.
"""
import json
import logging
import os
import time
import requests
from datetime import datetime, date

from core.connection import SAPConnection
from config.settings import settings
from transactions.request import (
    request_atend_linha, request_lt23_fifo1, request_lt23_cofre1,
    request_mb51_empurrada, request_mb51_besi3,
    request_lx03, request_lx02,
    request_lt22_imp2, request_lt22_imp3, request_lt22_zona,
    request_vl06i, request_maisewm021r, request_mb52,
    request_al11_besi3, request_pkmc, request_md04_global
)
from transactions.extract import extract_sp02_job
from core.watchdog import watchdog_infraestrutura

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

STATE_DIR = os.path.join(os.path.dirname(__file__), "..", "state")
MASTER_PLANT = "01-Anchieta"
CYCLE_WAIT = 300
TIMEOUT = 600

MONITORING_URL = "https://default27f32941ff804065bad2806c3a5798.6a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/74eb6c1250464387aa45f411abab074b/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KyWBNM3OzwPThb-GMLRTXH7K-B6te9Bd3Atbp5OwlA0"

JOB_ROUTER = {
    "ATEND_LINHA": request_atend_linha,
    "LT23_FIFO1": request_lt23_fifo1,
    "LT23_COFRE1": request_lt23_cofre1,
    "MB51_EMPURRADA": request_mb51_empurrada,
    "MB51_BESI3": request_mb51_besi3,
    "LX03_FIFO2": request_lx03,
    "LX03_ZONA": request_lx03,
    "LX03_BLOQUEADOS": request_lx03,
    "LX02_IMP": request_lx02,
    "LX02_BESI3": request_lx02,
    "LT22_IMP2": request_lt22_imp2,
    "LT22_IMP3": request_lt22_imp3,
    "LT22_ZONA": request_lt22_zona,
    "LT22_ZONA_GERAL": request_lt22_zona,
    "VL06I_FORNEC": request_vl06i,
    "MAISEWM021R_EMBALAGEM": request_maisewm021r,
    "MB52_AUTO": request_mb52,
    "AL11_BESI3": request_al11_besi3,
    "PKMC_GERAL": request_pkmc,
    "MD04_GLOBAL": request_md04_global,
}


class JobState:
    def __init__(self, plant_id):
        self._plant_id = plant_id
        os.makedirs(STATE_DIR, exist_ok=True)
        self._plant_file = os.path.join(STATE_DIR, f"{plant_id}.json")
        self._global_file = os.path.join(STATE_DIR, "global.json")
        self._plant_data = self._load(self._plant_file)
        self._global_data = self._load(self._global_file)

    def _load(self, path):
        try:
            with open(path, "r") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def _save_plant(self):
        with open(self._plant_file, "w") as f:
            json.dump(self._plant_data, f, indent=2)

    def _save_global(self):
        with open(self._global_file, "w") as f:
            json.dump(self._global_data, f, indent=2)

    def _is_global(self, job_data):
        return job_data.get("scope") == "global"

    def _get_store(self, job_data):
        return self._global_data if self._is_global(job_data) else self._plant_data

    def get(self, job_key, job_data):
        return self._get_store(job_data).get(job_key, {})

    def mark_requested(self, job_key, job_data):
        self._get_store(job_data)[job_key] = {
            "requested": datetime.now().isoformat(),
            "extracted": None,
            "date": date.today().isoformat()
        }
        if self._is_global(job_data):
            self._save_global()
        else:
            self._save_plant()

    def mark_extracted(self, job_key, job_data):
        store = self._get_store(job_data)
        if job_key in store:
            store[job_key]["extracted"] = datetime.now().isoformat()
            if self._is_global(job_data):
                self._save_global()
            else:
                self._save_plant()

    def should_request(self, job_key, job_data):
        s = self.get(job_key, job_data)
        today = date.today().isoformat()

        if not s or s.get("date") != today:
            return True

        if job_data.get("once_per_day"):
            return False

        if s.get("extracted"):
            return True

        req = s.get("requested")
        if req:
            elapsed = (datetime.now() - datetime.fromisoformat(req)).total_seconds()
            return elapsed >= TIMEOUT

        return True

    def needs_extraction(self, job_key, job_data):
        if not job_data.get("background_job", True):
            return False
        s = self.get(job_key, job_data)
        return (s.get("date") == date.today().isoformat()
                and s.get("requested")
                and not s.get("extracted"))


def report_status(plant_id, job_key, status_text, pct):
    try:
        payload = {
            "planta": plant_id,
            "job": job_key,
            "status": status_text,
            "concluidos": int(pct),
            "total": 100
        }
        requests.post(MONITORING_URL, json=payload, timeout=3)
    except Exception as e:
        logging.warning(f"Failed to report status: {e}")


def run_plant(plant_id: str):
    logging.info(f"Starting orchestrator for plant: {plant_id}")

    if plant_id not in settings.config.get("plants", {}):
        logging.error(f"Plant '{plant_id}' not found in config.")
        return

    conn = SAPConnection(plant_id)
    try:
        conn.connect()
        conn.ensure_logged_in()
    except Exception as e:
        logging.error(f"Connection failed for '{plant_id}': {e}")
        return

    jobs = settings.config.get("jobs", {})

    while True:
        watchdog_infraestrutura()
        
        if not conn.check_connection():
            logging.warning("Conexão SAP inativa. Tentando reconectar...")
            try:
                conn.connect()
                conn.ensure_logged_in()
            except Exception as e:
                logging.error(f"Falha ao reconectar: {e}")
                time.sleep(60)
                continue

        state = JobState(plant_id)
        t0 = datetime.now()
        logging.info(f"=== Cycle {t0.strftime('%H:%M:%S')} ===")
        
        cycle_has_error = False

        jobs_to_request = [
            k for k, v in jobs.items() 
            if v.get("active") 
            and plant_id in v.get("plant_params", {}) 
            and (v.get("scope") != "global" or plant_id == MASTER_PLANT) 
            and state.should_request(k, v)
        ]
        req_total = len(jobs_to_request)
        req_count = 0

        for job_key, job_data in jobs.items():
            if not job_data.get("active") or plant_id not in job_data.get("plant_params", {}):
                continue

            if job_data.get("scope") == "global" and plant_id != MASTER_PLANT:
                continue

            if not state.should_request(job_key, job_data):
                logging.info(f"[SKIP] {job_key}")
                continue

            t_code = job_data.get("transaction")
            
            pct = (req_count / req_total * 30) if req_total > 0 else 30
            report_status(plant_id, job_key, f"Solicitando relatório ({t_code})...", pct)
            
            logging.info(f"[REQUEST] {job_key} ({t_code})")
            try:
                conn.start_transaction(t_code)
                func = JOB_ROUTER.get(job_key)
                if func:
                    func(conn.session, plant_id, job_key)
                    state.mark_requested(job_key, job_data)
                    logging.info(f"[OK] {job_key}")
                    
                    pct = (req_count / req_total * 30) if req_total > 0 else 30
                    report_status(plant_id, job_key, "Solicitação concluída.", pct)
                else:
                    logging.warning(f"[WARN] No handler for '{job_key}'")
            except Exception as e:
                logging.error(f"[ERROR] {job_key}: {e}")
                report_status(plant_id, job_key, f"ERRO na solicitação: {str(e)[:50]}", pct)
                cycle_has_error = True

            req_count += 1

        pending = [(k, v) for k, v in jobs.items()
                    if v.get("active")
                    and plant_id in v.get("plant_params", {})
                    and (v.get("scope") != "global" or plant_id == MASTER_PLANT)
                    and state.needs_extraction(k, v)]

        if pending:
            logging.info("[WAIT] Aguardando 120s para jobs processarem...")
            
            for i in range(4):
                current_pct = 30 + (i * 7.5) # Moves from 30% to ~52.5%
                if not cycle_has_error:
                    report_status(plant_id, "SISTEMA", f"Processando SAP... ({(i+1)*30}s/120s)", current_pct)
                time.sleep(30)

            if not conn.check_connection():
                logging.warning("Conexão SAP inativa após espera de 120s. Tentando reconectar para extração...")
                try:
                    conn.connect()
                    conn.ensure_logged_in()
                except Exception as e:
                    logging.error(f"Falha ao reconectar durante extração: {e}")

            if conn.check_connection():
                logging.info(f"[SP02] {len(pending)} pending")
                try:
                    conn.start_transaction("SP02")
                    ext_total = len(pending)
                    ext_count = 0
                    for job_key, job_data in pending:
                        pct = 60 + (ext_count / ext_total * 40) if ext_total > 0 else 60
                        report_status(plant_id, job_key, "Extraindo spool na SP02...", pct)
                        logging.info(f"[EXTRACT] {job_key}")
                        if extract_sp02_job(conn.session, plant_id, job_key, job_data):
                            state.mark_extracted(job_key, job_data)
                            logging.info(f"[OK] {job_key} extracted")
                            
                            ext_count += 1
                            pct = 60 + (ext_count / ext_total * 40) if ext_total > 0 else 60
                            report_status(plant_id, job_key, "Extração salva com sucesso!", pct)
                        else:
                            ext_count += 1
                except Exception as e:
                    logging.error(f"[ERROR] SP02: {e}")
                    report_status(plant_id, "SP02", f"ERRO na extração: {str(e)[:50]}", 60)
                    cycle_has_error = True

        wait = int(CYCLE_WAIT - (datetime.now() - t0).total_seconds())
        if wait > 0:
            if not cycle_has_error:
                report_status(plant_id, "SISTEMA", f"Repouso. Próximo ciclo em {wait}s.", 100)
            else:
                report_status(plant_id, "SISTEMA", f"Repouso após erros. Próximo ciclo em {wait}s.", 100)
                
            logging.info(f"=== Next cycle in {wait}s ===")
            time.sleep(wait)
