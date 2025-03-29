import pandas as pd
from pathlib import Path
import os
import numpy as np
from typing import List
import datetime


def generar_primer_archivo(nombre_carpeta):
    # Obtener la ruta de la carpeta (relativa al script)
    ruta_carpeta = Path(__file__).parent
    
    # Diccionario para almacenar la suma total de "Week Adm" por película
    total_peliculas = {}
    # Diccionario para almacenar el top 5 por archivo
    top5_por_archivo = {}
    # Diccionario para almacenar el total de "Week Adm" por archivo
    total_por_archivo = {}

    # Verificar si la carpeta existe
    # if not ruta_carpeta.exists():
    #     print(f"La carpeta '{nombre_carpeta}' no existe.")
    #     return None

    # Iterar sobre todos los archivos Excel en la carpeta
    for archivo in ruta_carpeta.glob('*.xls*'):  # Acepta .xls y .xlsx
        print(f"\nProcesando archivo: {archivo.name}")
        
        try:
            # Leer el archivo Excel
            df = pd.read_excel(
                archivo,
                sheet_name=0,  # Leer la primera hoja
                header=2,      # Los encabezados están en la fila 3 (índice 2)
                usecols=[1, 11]  # Leer solo las columnas 2 (Title) y 12 (Week Adm)
            )
            
            # Renombrar las columnas para facilitar el acceso
            df.columns = ['Title', 'Week Adm']
            
            # Verificar si las columnas necesarias están presentes
            if 'Title' not in df.columns or 'Week Adm' not in df.columns:
                print(f"El archivo {archivo.name} no tiene las columnas 'Title' o 'Week Adm'. Saltando...")
                continue
            
            # Sumar "Week Adm" por película en este archivo
            suma_por_pelicula = df.groupby('Title')['Week Adm'].sum().sort_values(ascending=False)
            
            # Calcular el total de "Week Adm" para este archivo
            total_archivo = df['Week Adm'].sum()
            total_por_archivo[archivo.name] = total_archivo
            
            # Actualizar el total global
            for pelicula, suma in suma_por_pelicula.items():
                if pelicula in total_peliculas:
                    total_peliculas[pelicula] += suma
                else:
                    total_peliculas[pelicula] = suma
            
            # Guardar el top 5 de este archivo con porcentaje respecto al archivo
            top5_por_archivo[archivo.name] = {
                pelicula: (suma, (suma / total_archivo) * 100)
                for pelicula, suma in suma_por_pelicula.head(5).items()
            }
        
        except Exception as e:
            print(f"Error al procesar el archivo {archivo.name}: {e}")
            continue

    # Obtener el top 10 global
    top10_global = pd.Series(total_peliculas).sort_values(ascending=False).head(10)

    # Calcular la suma total de todas las películas
    suma_total = sum(total_peliculas.values())

    # Crear un diccionario con los resultados
    resultado = {
        "top10_global": {
            pelicula: (suma, (suma / suma_total) * 100)
            for pelicula, suma in top10_global.items()
        },
        "top5_por_archivo": top5_por_archivo,
        "suma_total": suma_total
    }

    # Escribir los resultados en un archivo txt con formato de columnas
    with open('excercise1.txt', 'w', encoding='utf-8') as f:
        # Escribir el encabezado del top 10 global
        f.write("Top 10 Global:\n")
        f.write("{:<50} {:<20} {:<10}\n".format("Películas", "Cantidad de asistentes", "%"))
        f.write("-" * 80 + "\n")
        
        # Escribir el top 10 global con porcentajes
        for pelicula, (suma, porcentaje) in resultado["top10_global"].items():
            f.write("{:<50} {:<20.1f} ({:.2f}%)\n".format(pelicula, suma, porcentaje))
        
        # Escribir el encabezado del top 5 por archivo
        f.write("\nTop 5 por Archivo:\n")
        for archivo, top5 in resultado["top5_por_archivo"].items():
            f.write(f"{archivo}:\n")
            f.write("{:<50} {:<20} {:<10}\n".format("Película", "Cantidad de asistentes", "%"))
            f.write("-" * 80 + "\n")
            for pelicula, (suma, porcentaje) in top5.items():
                f.write("{:<50} {:<20.1f} ({:.2f}%)\n".format(pelicula, suma, porcentaje))
            f.write("\n")
        
        # Escribir la suma total
        f.write(f"\nSuma total de todas las películas: {suma_total}\n")

    # Devolver el resultado
    return resultado

def generar_segundo_archivo(nombre_carpeta):
    # Obtener la ruta de la carpeta (relativa al script)
    ruta_carpeta = Path(__file__).parent
    
    # Diccionario para almacenar la suma total de "Week Adm" por película
    Total_cine = {}
    # Diccionario para almacenar el top 5 por archivo
    top5_por_archivo = {}
    # Diccionario para almacenar el total de "Week Adm" por archivo
    total_por_archivo = {}

    # Verificar si la carpeta existe
    # if not ruta_carpeta.exists():
    #     print(f"La carpeta '{nombre_carpeta}' no existe.")
    #     return None

    # Iterar sobre todos los archivos Excel en la carpeta
    for archivo in ruta_carpeta.glob('*.xls*'):  # Acepta .xls y .xlsx
        print(f"\nProcesando archivo: {archivo.name}")
        
        try:
            # Leer el archivo Excel
            df = pd.read_excel(
                archivo,
                sheet_name=0,  # Leer la primera hoja
                header=2,      # Los encabezados están en la fila 3 (índice 2)
                usecols=[5, 11]  # Leer solo las columnas 2 (Title) y 12 (Week Adm)
            )
            
            # Renombrar las columnas para facilitar el acceso
            df.columns = ['Circuit', 'Week Adm']
            
            # Verificar si las columnas necesarias están presentes
            if 'Circuit' not in df.columns or 'Week Adm' not in df.columns:
                print(f"El archivo {archivo.name} no tiene las columnas 'Circuit' o 'Week Adm'. Saltando...")
                continue
            
            # Sumar "Week Adm" por película en este archivo
            suma_por_cine = df.groupby('Circuit')['Week Adm'].sum().sort_values(ascending=False)
            
            # Calcular el total de "Week Adm" para este archivo
            total_archivo = df['Week Adm'].sum()
            total_por_archivo[archivo.name] = total_archivo
            
            # Actualizar el total global
            for cine, suma in suma_por_cine.items():
                if cine in Total_cine:
                    Total_cine[cine] += suma
                else:
                    Total_cine[cine] = suma
            
            # Guardar el top 5 de este archivo con porcentaje respecto al archivo
            top5_por_archivo[archivo.name] = {
                cine: (suma, (suma / total_archivo) * 100)
                for cine, suma in suma_por_cine.head(5).items()
            }
        
        except Exception as e:
            print(f"Error al procesar el archivo {archivo.name}: {e}")
            continue

    # Obtener el top 10 global
    top10_global = pd.Series(Total_cine).sort_values(ascending=False).head(10)

    # Calcular la suma total de todas las películas
    suma_total = sum(Total_cine.values())

    # Crear un diccionario con los resultados
    resultado = {
        "top10_global": {
            cine: (suma, (suma / suma_total) * 100)
            for cine, suma in top10_global.items()
        },
        "top5_por_archivo": top5_por_archivo,
        "suma_total": suma_total
    }

    # Escribir los resultados en un archivo txt con formato de columnas
    with open('excercise2.txt', 'w', encoding='utf-8') as f:
        # Escribir el encabezado del top 10 global
        f.write("Top 10 Global:\n")
        f.write("{:<50} {:<20} {:<10}\n".format("Cantidad de cines", "Cantidad de asistentes", "%"))
        f.write("-" * 80 + "\n")
        
        # Escribir el top 10 global con porcentajes
        for cine, (suma, porcentaje) in resultado["top10_global"].items():
            f.write("{:<50} {:<20.1f} ({:.2f}%)\n".format(cine, suma, porcentaje))
        
        # Escribir el encabezado del top 5 por archivo
        f.write("\nTop 5 por Archivo:\n")
        for archivo, top5 in resultado["top5_por_archivo"].items():
            f.write(f"{archivo}:\n")
            f.write("{:<50} {:<20} {:<10}\n".format("Cantidad de cines", "Cantidad de asistentes", "%"))
            f.write("-" * 80 + "\n")
            for cine, (suma, porcentaje) in top5.items():
                f.write("{:<50} {:<20.1f} ({:.2f}%)\n".format(cine, suma, porcentaje))
            f.write("\n")
        
        # Escribir la suma total
        # f.write(f"\nSuma total de todas las Cines: {suma_total}\n")

    # Devolver el resultado
    return resultado

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


def guardar_tercer_archivo(df_total_per_movie_sorted: pd.DataFrame, movies_by_circuit_sorted: pd.DataFrame) -> None:
    
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
        output.append(f"{'Circuit':<{col1_width}} {'Week Adm':<{col2_width}} {'Porcentaje':<{col3_width}}")
        
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



# generar los archivos
archivo1 = generar_primer_archivo('exceles')
archivo2 = generar_segundo_archivo('exceles')
df1, df2 = generateConsolidatedDfs()
guardar_tercer_archivo(df1,df2)