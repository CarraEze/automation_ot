#PENDIENTE QUE MONITOREE UN SOLO ARCHIVO EN VEZ DE LA CARPETA COMPLETA
#AGREGAR NOMBRE DEL ARCHIVO A CONFIG.JSON PARA QUE SEA PERSONALIZABLE

#DataHandle
import os
import sys
import gspread
import pandas as pd
from gspread_formatting import CellFormat, TextFormat, Color, format_cell_range
from gspread_dataframe import set_with_dataframe
import time, re, json

##Observer
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler
import subprocess
import time

# Almacena la última ejecucion por archivo
last_execution = {}

# Ignora eventos repetidos dentro de X segundos
COOLDOWN = 5

## Define las columnas a extraer de todo el archivo de excel
COLUMNS_TO_EXTRACT = [
    "DATOS0",
    "DATOS1",
    "DATOS4",
    "DATOS5",
    ##"DATOS7", Traia info de la OT pero cuando estaba mal cerrada
    "FECHAHORA",
    "NOMBRE",
    "CODIGOALFA",
    "CLIENTE",
    "TECNICO",
    "ESTADO_OT"
]

## Define un diccionario para mapear los números de mes a sus nombres en español
MESES = {
    "01": "ENERO", "02": "FEBRERO", "03": "MARZO", "04": "ABRIL", "05": "MAYO", "06": "JUNIO", "07": "JULIO", "08": "AGOSTO", "09": "SEPTIEMBRE", "10": "OCTUBRE", "11": "NOVIEMBRE", "12": "DICIEMBRE"
}

## Objeto para darle formato al sheets
format = CellFormat(
    backgroundColor=Color(0.937, 0.937, 0.937),
    textFormat=TextFormat(
        fontFamily="Calibri",
        fontSize=10,
    )
)

## obtiene el path global, ya que al generar el exe, no se pueden direccionar archivos de manera comun
def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

base_path = get_base_path()
    
def get_config():
    with open(base_path + "\\" + "config.json", "r") as file:
        (worksheet_id, file_path, watch_folder) = json.load(file).values()
    if not worksheet_id or not file_path or not watch_folder:
        raise ValueError("Some info is missing in the config file")
    return (worksheet_id, file_path, watch_folder)

## Función para extraer el DataFrame del archivo de Excel
def fetch_dataframe(xlxs_path):
    try:
        data = pd.read_excel(xlxs_path, usecols=COLUMNS_TO_EXTRACT)
        return data
    except Exception:
        raise ValueError("The extracted DataFrame is empty or None.")
    
    ## Función para limpiar y transformar el DataFrame
def clean_dataframe(dataframe):
    filtered = dataframe.loc[:, [col for col in COLUMNS_TO_EXTRACT if col in dataframe.columns]] ## Extrae las columnas indicadas en COLUMNS_TO_EXTRACT
    filtered.rename(columns={'DATOS0': 'RECLAMO', 'DATOS1': 'HISTORIAL', 'DATOS5' : 'RESOLUCION', 'NOMBRE' : 'SOLICITANTE'}, inplace=True) ## Se renombran las columnas de manera más descriptiva
    filtered["RECLAMO"] = filtered["RECLAMO"].astype(str).str.replace(r"[\r\n]+", " - ", regex=True).str.upper()
    filtered["RESOLUCION"] = filtered["RESOLUCION"].astype(str).str.replace(r"[\r\n]+", " - ", regex=True)    
    filtered["FECHAHORA"] = pd.to_datetime(filtered["FECHAHORA"], errors='coerce').dt.floor('s') ## Se cambia el formato del dato y se elimina la precision de milisegundos
    filtered["DIA"] = filtered["FECHAHORA"].dt.day.astype(str).str.zfill(2) ## Se extrae el día de la fecha y se formatea a dos dígitos
    filtered["MES"] = filtered["FECHAHORA"].dt.month.astype(str).str.zfill(2).map(MESES) ## Se extrae el mes de la fecha, se formatea a dos dígitos y se mapea al nombre del mes en español
    filtered["HORA"] = filtered["FECHAHORA"].dt.time.astype(str).str.zfill(2) ## Se extrae la hora de la fecha y se formatea a dos dígitos
    filtered["ACOMPAÑANTE"] = [None]*len(filtered) ## Se crea una nueva columna vacía para el acompañante para coincidir con el formato del sheet usado actualmente
    filtered["MOTIVO"] = [None]*len(filtered) 
    filtered["SUBMOTIVO"] = [None]*len(filtered) 
    filtered["MATERIALES"] = [None]*len(filtered) 
    filtered[["INICIO", "FIN", "HISTORIAL", "REITERA_VISITA"]] = filtered["HISTORIAL"].apply(clean_history) ## Se limpia la columna de historial para extraer solo los cierres de OT
    filtered["CODIGOALFA"] = filtered["CODIGOALFA"].str.extract(r"-([^/]+)/")
    filtered = filtered[["CLIENTE", "RECLAMO", "SOLICITANTE","FECHAHORA", "DIA", "MES", "HORA", "TECNICO", "ACOMPAÑANTE","INICIO","FIN","MOTIVO", "SUBMOTIVO","HISTORIAL", "MATERIALES", "REITERA_VISITA","RESOLUCION", "DATOS4", "CODIGOALFA"]] ## Se reordena el DataFrame para coincidir con el formato del sheet usado actualmente
    filtered = filtered.iloc[::-1]
    return filtered

## Se limpia el historial completo para dejar unicamente los cierres de OT, con el formato (fecha) descripción del cierre
## Para entender el criterio de limpieza se debe observar el excel en crudo
def clean_history(historial):
    delimiter_start_date =  r"-{40}\s*INICIO DE TRABAJOS EN OT\.*\s*T.cnico:[\s\S]*?(\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}:\d{2})\s*-{40}"
    delimiter_end = r'-{40}\s*FINALIZACI.N DE TRABAJOS EN OT\.?\s*T.cnico:\s*([^\n\r]+)\s*(?:Nota:\s*)?([\s\S]*?)\s*(\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}:\d{2})\s*-{40}'
    ot_start_date = re.search(delimiter_start_date, historial)
    ot_ends = re.findall(delimiter_end, historial)
    ot_end_date = ot_ends[-1][-1] if ot_ends else None
    ots_appended = ""
    for ot in ot_ends:
        ots_appended += f"({ot[-1]}) {ot[0]} {ot[1].replace('\n', ' ').replace('\r', ' ').strip()} - "
    return pd.Series({
            "INICIO": ot_start_date.group(1) if ot_start_date else "SIN FECHA",
            "FIN": ot_end_date if ot_end_date else "SIN FECHA",
            "HISTORIAL": ots_appended,
            "REITERA_VISITA": "REITERA VISITA" if len(ot_ends)>1 else " "
        })

## Se hace un fetch del worksheet de google sheets
def fetch_worksheet(spreadsheet_name, worksheet_index=0):
    try:
        gc = gspread.service_account(filename=base_path + "\\" + 'credentials.json')
        sh = gc.open_by_key(spreadsheet_name)
        worksheet = sh.get_worksheet(worksheet_index)
        worksheet.clear()
        return worksheet
    except Exception as e:
        print(f"An error has ocurred during worksheet fetch: {e}")
        return None

if __name__ == "__main__":
    print(f"CARGANDO")
    start_time = time.time()
    (worksheet_id, file_path, watch_folder) = get_config()
    print(f"---25%---")
    worksheet = fetch_worksheet(worksheet_id)
    print(f"---50%---")
    dataframe = clean_dataframe(fetch_dataframe(file_path))
    print(f"---75%---")
    set_with_dataframe(worksheet, dataframe, include_index=False, include_column_header=True)
    print(f"---100%---")
    format_cell_range(worksheet, "A2:Z100",format)
    end_time = time.time()
    print(f"Total execution time: {end_time - start_time} seconds")
