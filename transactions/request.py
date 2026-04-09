"""
Scripts de transacao SAP
Helpers genericos somente para sequencias verdadeiramente repetitivas.
"""
import os
import time
from datetime import datetime, timedelta
import pythoncom
import win32com.client

from config.settings import settings
from core.utils import get_target_export_dir


#  HELPERS 

def close_excel(filename):
    time.sleep(3)
    try:
        excel = win32com.client.GetObject(Class="Excel.Application")
        for wb in excel.Workbooks:
            if filename.lower() in wb.Name.lower():
                wb.Close(SaveChanges=False)
                break
    except Exception:
        pass

def send_to_background(session, spool_name: str, printer: str = "locl"):
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select()

        try:
            session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").Text
        except pythoncom.com_error:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

        try:
            session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = printer
        except pythoncom.com_error:
            pass

        session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").Text = spool_name
        session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").Text = spool_name
        session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").SetFocus()
        session.findById("wnd[1]/tbar[0]/btn[13]").press()

        try:
            session.findById("wnd[2]").sendVKey(0)
        except pythoncom.com_error:
            pass

        session.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]").sendVKey(3)
    except pythoncom.com_error as e:
        raise RuntimeError(f"Falha ao enviar job '{spool_name}' para background: {e}")

def export_xxl(session, path, filename, shell_id=None):
    """
    Exporta para XXL usando o truque do .tmp para evitar que o Excel abra automaticamente.
    """
    try:
        nome_base, extensao = os.path.splitext(filename)
        temp_filename = f"{nome_base}.tmp"

        if shell_id:
            shell = session.findById(shell_id)
        else:
            try:
                shell = session.findById("wnd[0]/shellcont[1]/shell")
            except pythoncom.com_error:
                shell = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")

        shell.pressToolbarContextButton("&MB_EXPORT")
        shell.selectContextMenuItem("&XXL")

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = temp_filename
        
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        try:
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
        except pythoncom.com_error:
            pass

        caminho_temp = os.path.abspath(os.path.join(path, temp_filename))
        caminho_final = os.path.abspath(os.path.join(path, filename))

        timeout = 45 
        start = time.time()

        while not os.path.exists(caminho_temp):
            if time.time() - start > timeout:
                raise TimeoutError(f"SAP demorou muito para exportar {temp_filename}")
            time.sleep(0.5)

        time.sleep(2.0)

        if os.path.exists(caminho_final):
            os.remove(caminho_final)

        os.rename(caminho_temp, caminho_final)

    except pythoncom.com_error as e:
        raise RuntimeError(f"Export XXL falhou para {filename}: {e}")
    except Exception as e:
        raise RuntimeError(f"Erro no processamento XXL de {filename}: {e}")

def _get_params(job_key, plant_id):
    cfg = settings.config["jobs"][job_key]
    plant_params = cfg["plant_params"][plant_id]
    spool = plant_params.get("spool_name") or cfg["spool_name"]
    return plant_params, spool



def request_atend_linha(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/radT1_ALLTA").Select()
    session.findById("wnd[0]/usr/ctxtT1_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params["variant"]
    session.findById("wnd[0]/usr/ctxtLISTV").SetFocus()
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool, params.get("printer", "locl"))


def request_lt23_fifo1(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtT1_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/txtT1_TANUM-LOW").Text = ""
    session.findById("wnd[0]/usr/txtT1_TANUM-HIGH").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtAENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtMLANGU-LOW").Text = ""
    session.findById("wnd[1]").sendVKey(0)

    session.findById("wnd[0]/usr/radT1_ALLTA").Select()
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


def request_lt23_cofre1(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtT1_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/txtT1_TANUM-LOW").Text = ""
    session.findById("wnd[0]/usr/txtT1_TANUM-HIGH").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtAENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtMLANGU-LOW").Text = ""
    session.findById("wnd[1]").sendVKey(0)

    session.findById("wnd[0]/usr/radT1_ALLTA").Select()
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  MB51_EMPURRADA 
def request_mb51_empurrada(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=4)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = params["werks"]
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = params["lgort"]
    session.findById("wnd[0]/usr/ctxtBWART-LOW").Text = "311"
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").SetFocus()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/radRFLAT_L").Select()
    session.findById("wnd[0]/usr/ctxtALV_DEF").Text = params["variant"]
    session.findById("wnd[0]/usr/ctxtALV_DEF").SetFocus()
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool, params.get("printer", "locl"))


#  MB51_BESI3 
def request_mb51_besi3(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=2)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = params["werks"]
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = params["lgort"]

    session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").press()
    
    bwart_list = params.get("bwart_list", ["y61", "311"])
    for idx, bwart in enumerate(bwart_list):
        session.findById(f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{idx}]").Text = bwart
        
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").SetFocus()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/radRFLAT_L").Select()
    session.findById("wnd[0]/usr/ctxtALV_DEF").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LX03 (FIFO2, ZONA, BLOQUEADOS) 
def request_lx03(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = params["lgnum"]
    if "lgtyp" in params:
        session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").Text = params["lgtyp"]
    session.findById("wnd[0]/usr/chkPMITB").Selected = True
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LX02 (IMP, BESI3) 
def request_lx02(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = params["lgnum"]
    
    if "werks" in params:
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = params["werks"]
        
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = params["variant"]
    session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus()

    send_to_background(session, spool)


#  LT22_IMP2 

def request_lt22_imp2(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = params["lgtyp"]
    session.findById("wnd[0]/usr/radT3_OFFTA").Select()
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LT22_IMP3 
def request_lt22_imp3(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = datetime.now().replace(day=1).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = params["lgtyp"]
    session.findById("wnd[0]/usr/radT3_ALLTA").Select()
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LT22_ALERTAOP 
def request_lt22_alertaop(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = params["lgtyp"]
    
    if params.get("radio") == "ALLTA":
        session.findById("wnd[0]/usr/radT3_ALLTA").Select()
    
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params.get("variant", "")
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LT22_ZONA 
def request_lt22_zona(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = params["lgtyp"]
    
    if params.get("radio") == "ALLTA":
        session.findById("wnd[0]/usr/radT3_ALLTA").Select()
    else:
        session.findById("wnd[0]/usr/radT3_OFFTA").SetFocus()
        
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params.get("variant", "")
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  VL06I_FORNEC (e VL06I_FORNEC2 via lgnum) 
def request_vl06i(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    today = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/btnBUTTON7").press()
    
    lgnum = params.get("lgnum")
    vstel = params.get("vstel", "")

    if lgnum:
        # CTB parte 2: filtra por número de depósito (lgnum) em vez de ponto de expedição
        session.findById("wnd[0]/usr/ctxtIF_VSTEL-LOW").Text = ""
        session.findById("wnd[0]/usr/ctxtIT_LGNUM-LOW").Text = lgnum
    elif isinstance(vstel, list):
        session.findById("wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH").press()
        for idx, val in enumerate(vstel):
            session.findById(f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{idx}]").Text = val
        session.findById("wnd[1]").sendVKey(8)
    else:
        session.findById("wnd[0]/usr/ctxtIF_VSTEL-LOW").Text = vstel
        
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").Text = params["date_low"]
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").Text = today
    session.findById("wnd[0]/usr/ctxtIT_WBSTK-LOW").SetFocus()
    session.findById("wnd[0]").sendVKey(2)

    try:
        session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton("EXCL")
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
    except pythoncom.com_error:
        pass

    session.findById("wnd[0]/usr/ctxtIT_WBSTK-LOW").Text = "c"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkIF_ITEM").Selected = True

    send_to_background(session, spool)


#  MAISEWM021R_EMBALAGEM 
def request_maisewm021r(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    today = datetime.now()
    first_day = today.replace(day=1).strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = params["werks"]
    session.findById("wnd[0]/usr/ctxtS_DATUM-LOW").Text = first_day
    session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").Text = today.strftime("%d.%m.%Y")
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/radP_MOV").Select()

    send_to_background(session, spool)


#  MB52_AUTO 
def request_mb52(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    send_to_background(session, spool)


#  AL11_BESI3 (Foreground) 
def request_al11_besi3(session, plant_id: str, job_key: str):
    # 1. Obter configurações da Planta do JSON
    plant_config = settings.config["plants"].get(plant_id, {})
    folder_name = plant_config.get("folder_name", plant_id)
    inner_base_path = plant_config.get("base_path", "")
    
    params = settings.config["jobs"][job_key]["plant_params"][plant_id]
    local_extract = params.get("local_extract", "")
    name_file = params.get("name_file", "besi3.txt").format(date=datetime.now())
    
    # 2. Construir o caminho absoluto perfeito
    # Passamos o 'folder_name' em vez do 'plant_id' para criar a raiz correta!
    base_plant_path = get_target_export_dir(folder_name)
    full_path = os.path.normpath(os.path.join(base_plant_path, inner_base_path, local_extract))
    os.makedirs(full_path, exist_ok=True)

    folders = ["\\\\10.135.7.23\\files\\PRD\\interfaces", "pp", "inbound", "BESI3", "5100", "Backup"]
    grid_id = "wnd[0]/usr/cntlGRID1/shellcont/shell"

    for folder in folders:
        i = 0
        col = "DIRNAME" if "\\\\" in folder else "NAME"
        while True:
            try:
                val = session.findById(grid_id).GetCellValue(i, col)
                if val == folder:
                    session.findById(grid_id).setCurrentCell(i, col)
                    session.findById(grid_id).doubleClickCurrentCell()
                    break
                i += 1
            except pythoncom.com_error:
                break

    session.findById(grid_id).setCurrentCell(-1, "MOD_DATE")
    session.findById(grid_id).selectColumn("MOD_DATE")
    session.findById("wnd[0]/tbar[1]/btn[40]").press()

    session.findById(grid_id).currentCellRow = -1
    session.findById(grid_id).selectColumn("USEABLE")
    session.findById("wnd[0]/tbar[1]/btn[29]").press()
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "x"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById(grid_id).selectedRows = "0"
    session.findById(grid_id).doubleClickCurrentCell()

    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = full_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = name_file
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


#  PKMC_GERAL (Foreground) 
def request_pkmc(session, plant_id: str, job_key: str):
    # 1. Obter configurações da Planta do JSON
    plant_config = settings.config["plants"].get(plant_id, {})
    folder_name = plant_config.get("folder_name", plant_id)
    inner_base_path = plant_config.get("base_path", "")
    
    params = settings.config["jobs"][job_key]["plant_params"][plant_id]
    local_extract = params.get("local_extract", "")
    name_file = params.get("name_file", "PKMC.XLSX").format(date=datetime.now())
    
    # 2. Construir o caminho absoluto perfeito
    # Passamos o 'folder_name' em vez do 'plant_id' para criar a raiz correta!
    base_plant_path = get_target_export_dir(folder_name)
    full_path = os.path.normpath(os.path.join(base_plant_path, inner_base_path, local_extract))
    os.makedirs(full_path, exist_ok=True)

    session.findById("wnd[0]/usr/ssubCCY_AND_SELECTION:SAPLMPK_CCY_UI:0111/subSELECTION:SAPLMPK_CCY_UI:0113/ctxtRMPKR-WERKS").Text = params["werks"]
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ssubCCY_AND_SELECTION:SAPLMPK_CCY_UI:0111/subSELECTION:SAPLMPK_CCY_UI:0113/btnESEL").press()
    session.findById("wnd[1]/usr/ctxtRANG_MAT-LOW").Text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    time.sleep(2)

    grid_pkmc = "wnd[0]/usr/ssubCCY_AND_SELECTION:SAPLMPK_CCY_UI:0111/subCCY:SAPLMPK_CCY_UI:0130/subBIGGRIDCONTAINER:SAPLMPK_CCY_UI:0135/cntlAVAILABLE_CONTROLCYCLES/shellcont/shell"
    session.findById(grid_pkmc).pressToolbarContextButton("&MB_VARIANT")
    session.findById(grid_pkmc).selectContextMenuItem("&LOAD")

    variant_grid = "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
    i = 0
    while True:
        try:
            val = session.findById(variant_grid).GetCellValue(i, "VARIANT")
            if val == params["variant"]:
                session.findById(variant_grid).selectedRows = str(i)
                session.findById(variant_grid).clickCurrentCell()
                break
            i += 1
        except pythoncom.com_error:
            break

    export_xxl(session, full_path, name_file, shell_id=grid_pkmc)
    session.findById("wnd[0]").sendVKey(3)


#  MD04_GLOBAL (Foreground) 
def request_md04_global(session, plant_id: str, job_key: str):
    # 1. Obter configurações da Planta do JSON
    plant_config = settings.config["plants"].get(plant_id, {})
    folder_name = plant_config.get("folder_name", plant_id)
    inner_base_path = plant_config.get("base_path", "")
    
    params = settings.config["jobs"][job_key]["plant_params"][plant_id]
    local_extract = params.get("local_extract", "")
    name_file = params.get("name_file", "MD04_full.XLSX").format(date=datetime.now())
    
    # 2. Construir o caminho absoluto perfeito
    # Passamos o 'folder_name' em vez do 'plant_id' para criar a raiz correta!
    base_plant_path = get_target_export_dir(folder_name)
    full_path = os.path.normpath(os.path.join(base_plant_path, inner_base_path, local_extract))
    os.makedirs(full_path, exist_ok=True)

    session.findById("wnd[0]/usr/tabsTAB300/tabpF02").Select()

    tab = "wnd[0]/usr/tabsTAB300/tabpF02/ssubINCLUDE300:SAPMM61R:0212"

    session.findById(f"{tab}/ctxtRM61R-WERKS2").Text = params["werks"]
    session.findById("wnd[0]").sendVKey(0)
    session.findById(f"{tab}/radRM61R-CLSKZ").Select()
    session.findById(f"{tab}/ctxtRM61R-CLASS").Text = params["class"]
    session.findById(f"{tab}/ctxtRM61R-KLART").Text = params["klart"]
    session.findById(f"{tab}/ctxtRM61R-KLART").SetFocus()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(0)

    tbl = "wnd[0]/usr/subVALUATION_DYNPRO:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S"
    for i in range(5):
        try:
            session.findById(f"{tbl}/ctxtRCTMS-MWERT[1,{i}]").Text = "*"
        except pythoncom.com_error:
            pass

    session.findById("wnd[0]/mbar/menu[4]/menu[0]").Select()
    session.findById("wnd[1]/usr/tabsPARAM/tabpSUCH").Select()
    session.findById("wnd[1]/usr/tabsPARAM/tabpSUCH/ssubSUB:SAPLCLPR:0110/txtRMCLPAR-DAR_MAX_HITS").Text = "0"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    export_xxl(session, full_path, name_file)

    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)
