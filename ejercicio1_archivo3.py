import pandas as pd

import os
import numpy as np

def generateDf(file_path:str, cols_to_use:List[str])-> pd.DataFrame:
    df = pd.read_excel(os.path.join(file_path), skiprows= 2, usecols=cols_to_use)
    return df

def generateConsolidatedDfs() -> pd.DataFrame:
    xls_files = [f for f in os.listdir(os.getcwd()) if f.endswith('.xls')]

    dfs = []
    cols = ["Title","Circuit","Week\nAdm"]
    for xls in xls_files:
        df = generateDf(xls,cols)
        dfs.append(df)
    
    # consolidar en un solo df los archivos xls
    consolidated = pd.concat(dfs, ignore_index= True)
    
    # Agrupar por titulo y circuit, sumarizando
    movies_by_circuit_sorted = consolidated.groupby(["Title", "Circuit"])["Week\nAdm"].sum().reset_index().sort_values(by="Title", ascending=True)

    # totales por pelicula
    df_total_per_movie_sorted = movies_by_circuit_sorted.groupby("Title")["Week\nAdm"].sum().reset_index().sort_values(by="Title", ascending=False).sort_values(by="Week\nAdm", ascending=True)
    
    return df_total_per_movie_sorted, movies_by_circuit_sorted


def saveFile(df_total_per_movie_sorted: pd.DataFrame, movies_by_circuit_sorted: pd.DataFrame) -> None:
    
    output = []
    # Definir ancho de columnas
    col1_width = 30  # Ancho para el nombre del Circuito
    col2_width = 10  # Ancho para Week\Adm
    col3_width = 15  # Ancho para el Porcentaje
    
    for title in df_total_per_movie_sorted["Title"]:
        movie_data = movies_by_circuit_sorted[movies_by_circuit_sorted["Title"] == title]
    
        # pasar columnas a numpy para optimizar
        circuits = movie_data["Circuit"].values
        week_adm_values = movie_data["Week\nAdm"].astype(int).values
    
        total_adm = int(df_total_per_movie_sorted.loc[df_total_per_movie_sorted["Title"] == title, "Week\nAdm"].values[0])
        
        percentages = np.round((week_adm_values / total_adm) * 100, 2)
        
        # Convertir a string con el formato adecuado
        percentages = np.array([f"{p}%" for p in percentages])
        total_adm = int(total_adm)
        output.append("############################################################################################################################")
        output.append(f"Pelicula: {title}\n")
    
         # Agregar encabezados con formato uniforme
        output.append(f"{'Circuit':<{col1_width}} {'Week\\Adm':<{col2_width}} {'Porcentaje':<{col3_width}}")
        
        # Usar NumPy para generar filas de salida en un solo paso
        formatted_rows = np.array([f"{c:<{col1_width}} {w:<{col2_width}} {p:<{col3_width}}" 
                                   for c, w, p in zip(circuits, week_adm_values, percentages)])
    
        output.extend(formatted_rows.tolist())
    
        # Agregar el total con alineación
        output.append(f"{'Total':<{col1_width}} {total_adm:<{col2_width}}\n")
    
    # Guardar el archivo de salida
    filename = "archivo3_"+datetime.datetime.now().strftime("%Y%m%d%H%M%S") + ".txt"
    with open(filename, "w") as f:
        f.write("\n".join(output))

df1, df2 = generateConsolidatedDfs()

saveFile(df1,df2)