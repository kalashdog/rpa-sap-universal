"""
Funções genéricas para extrair e baixar spools do SAP.
"""
import os
import logging
from datetime import datetime
import pythoncom

from config.settings import settings

def extract_sp02_job(session, plant_id: str, job_key: str, job_data: dict):
    try:
        plant_config = settings.config["plants"].get(plant_id, {})
        base_path = plant_config.get("base_path")
        plant_params = job_data.get("plant_params", {}).get(plant_id, {})
        local_extract = plant_params.get("local_extract", "")
        name_file = plant_params.get("name_file", f"{job_key}.txt").format(date=datetime.now())
        spool_name = job_data.get("spool_name", job_key)
        
        full_path = os.path.abspath(os.path.join(base_path, local_extract))
        os.makedirs(full_path, exist_ok=True)
        logging.info(f"Extraction for '{job_key}' (spool: '{spool_name}'): folder '{full_path}', file '{name_file}'")
        
        for i in range(3, 31):
            try:
                if session.findById(f"wnd[0]/usr/lbl[51,{i}]").Text == spool_name:
                    session.findById(f"wnd[0]/usr/chk[1,{i}]").Selected = True
                    session.findById(f"wnd[0]/usr/lbl[14,{i}]").SetFocus()
                    session.findById("wnd[0]").sendVKey(2)
                    session.findById("wnd[0]/tbar[1]/btn[48]").press()
                    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = full_path
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = name_file
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]/tbar[1]/btn[14]").press()
                    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    return True
                    
            except pythoncom.com_error:
                continue
                
        logging.info(f"Spool '{spool_name}' ainda nao apareceu no SP02. Tentando no proximo ciclo.")
        return False
        
    except KeyError as e:
        logging.error(f"Config ausente para extracao de '{job_key}': {e}")
        return False
    except pythoncom.com_error as e:
        logging.error(f"Erro COM SAP na extracao de '{job_key}': {e}")
        return False

