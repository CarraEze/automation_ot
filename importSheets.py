import os
import sys
import gspread
import pandas as pd
from gspread_formatting import CellFormat, TextFormat, Color, format_cell_range
from gspread_formatting.dataframe import format_with_dataframe
from gspread_dataframe import set_with_dataframe


## obtiene el path global, ya que al generar el exe, no se pueden direccionar archivos de manera comun
def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

base_dir = get_base_path()

## Define las columnas a extraer de todo el archivo de excel
COLUMNS_TO_EXTRACT = [
    "DATOS0",
    "DATOS1",
    "DATOS4",
    "DATOS5",
    "DATOS7",
    "FECHAHORA",
    "NOMBRE",
    "CODIGOALFA",
    "CLIENTE",
    "TECNICO",
    "ESTADO_OT"
]

## Define un diccionario para mapear los números de mes a sus nombres en español
MESES = {
    "01": "ENERO",
    "02": "FEBRERO",
    "03": "MARZO",
    "04": "ABRIL",
    "05": "MAYO",
    "06": "JUNIO",
    "07": "JULIO",
    "08": "AGOSTO",
    "09": "SEPTIEMBRE",
    "10": "OCTUBRE",
    "11": "NOVIEMBRE",
    "12": "DICIEMBRE"
}

## Función para extraer el DataFrame del archivo de Excel
def extract_dataframe(file_path):
    try:
        data = pd.read_excel(file_path)
        missing_columns = [col for col in COLUMNS_TO_EXTRACT if col not in data.columns]
        if missing_columns:
            print(f"Warning: the following columns are missing from the file: {missing_columns}")
        return data
    except Exception as e:
        print(f"An error occurred while extracting data: {e}")
        return None
    
    ## Función para limpiar y transformar el DataFrame
def clean_dataframe(dataframe):
    filtered = dataframe.loc[:, [col for col in COLUMNS_TO_EXTRACT if col in dataframe.columns]] ## Extrae las columnas indicadas en COLUMNS_TO_EXTRACT
    filtered.rename(columns={'DATOS0': 'RECLAMO', 'DATOS1': 'HISTORIAL', 'DATOS5' : 'RESOLUCION', 'NOMBRE' : 'SOLICITANTE'}, inplace=True) ## Se renombran las columnas de manera más descriptiva
    filtered["FECHAHORA"] = pd.to_datetime(filtered["FECHAHORA"], errors='coerce').dt.floor('s') ## Se cambia el formato del dato y se elimina la precision de milisegundos
    filtered["DIA"] = filtered["FECHAHORA"].dt.day.astype(str).str.zfill(2) ## Se extrae el día de la fecha y se formatea a dos dígitos
    filtered["MES"] = filtered["FECHAHORA"].dt.month.astype(str).str.zfill(2).map(MESES) ## Se extrae el mes de la fecha, se formatea a dos dígitos y se mapea al nombre del mes en español
    filtered["HORA"] = filtered["FECHAHORA"].dt.time.astype(str).str.zfill(2) ## Se extrae la hora de la fecha y se formatea a dos dígitos
    filtered["ACOMPAÑANTE"] = [None]*len(filtered) ## Se crea una nueva columna vacía para el acompañante para coincidir con el formato del sheet usado actualmente
    filtered["HISTORIAL"] = filtered["HISTORIAL"].apply(clean_historial) ## Se limpia la columna de historial para extraer solo los cierres de OT
    filtered = filtered[["CLIENTE", "RECLAMO", "SOLICITANTE", "DIA", "MES", "HORA", "TECNICO", "ACOMPAÑANTE" ,"RESOLUCION", "DATOS4","DATOS7","HISTORIAL", "CODIGOALFA"]] ## Se reordena el DataFrame para coincidir con el formato del sheet usado actualmente
    return filtered

## Se limpia el historial completo para dejar unicamente los cierres de OT, con el formato (fecha) descripción del cierre
## Para entender el criterio de limpieza se debe observar el excel en crudo
def clean_historial(historial):
    delimiter = "-"*40
    paragraphs = historial.split(delimiter)
    cierres = [
        "(" +p[-20:].strip()+ ") " + p.strip()[-(len(p)-45):-19]
        for p in paragraphs
        if p.strip().startswith("FINALIZACIÓN DE TRABAJOS EN OT")
        ]
    hist_cleaned = " - ".join(cierres)
    return hist_cleaned
     
## Se hace un fetch del worksheet de google sheets
def obtener_worksheet(spreadsheet_name, worksheet_index=0):
    try:
        gc = gspread.service_account(filename='credentials.json')
        sh = gc.open(spreadsheet_name)
        worksheet = sh.get_worksheet(worksheet_index)
        worksheet.clear()
        return worksheet
    except Exception as e:
        print(f"An error has ocurred during worksheet fetch: {e}")
        return None

## Objeto para darle formato al sheets
format = CellFormat(
    backgroundColor=Color(0.937, 0.937, 0.937),
    textFormat=TextFormat(
        fontFamily="Calibri",
        fontSize=10,
    )
)

if __name__ == "__main__":
    worksheet = obtener_worksheet("test_sheetsImport")
    dataframe = clean_dataframe(extract_dataframe(base_dir+"/04-05.xlsx"))
    set_with_dataframe(worksheet, dataframe, include_index=False, include_column_header=True)
    format_cell_range(worksheet, "A2:Z100",format)
    os.system("pause")