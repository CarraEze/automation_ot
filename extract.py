import pandas as pd

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
]

def extract_data(file_path):
    """
    Extract data from a XLSX file and return it as a DataFrame.
    """
    try:
        data = pd.read_excel(file_path)
        missing_columns = [col for col in COLUMNS_TO_EXTRACT if col not in data.columns]
        if missing_columns:
            print(f"Warning: the following columns are missing from the file: {missing_columns}")
        if "already_filtered" in data.columns:
            print("Warning: the file has already been filtered. Skipping extraction.")
            return data
        filtered = data.loc[:, [col for col in COLUMNS_TO_EXTRACT if col in data.columns]]
        filtered["already_filtered"] = True
        filtered.to_excel(file_path, index=False)
        return filtered
    except Exception as e:
        print(f"An error occurred while extracting data: {e}")
        return None
    
if __name__ == "__main__":
    file_path = "04-05.xlsx"
    data = extract_data(file_path)
    if data is None:
        print("Failed to extract data.")