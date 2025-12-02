import zipfile
import pandas as pd
import os

def zip_csvs_to_excel(zip_path, excel_path):

    with zipfile.ZipFile(zip_path, 'r') as z:
        csv_files = [f for f in z.namelist() if f.lower().endswith(".csv")]

        if not csv_files:
            raise ValueError("El ZIP no contiene archivos CSV")

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:

            for csv_name in csv_files:

                with z.open(csv_name) as f:
                    # Detecta automáticamente el separador (,;tab)
                    df = pd.read_csv(
                        f,
                        sep=None,
                        engine="python",   # NECESARIO para sep automático
                        encoding="latin1", # funciona incluso si viene de Excel
                    )

                sheet_name = os.path.splitext(os.path.basename(csv_name))[0][:31]
                df.to_excel(writer, index=False, sheet_name=sheet_name)
