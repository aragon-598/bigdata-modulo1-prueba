import os
import pandas as pd
import glob

DIRECTORIO_REPORTES = "Reportes"
archivos_excel = glob.glob(os.path.join(DIRECTORIO_REPORTES, "*.xlsx"))

peliculas_data = {}

for archivo in archivos_excel:
    thu_adm_value = 0
    weekend_adm_value = 0
    week_adm_value = 0

    df = pd.read_excel(archivo, header=2)
    
    thu_adm_cols = [col for col in df.columns if col.startswith("Thu\nAdm")]
    if not thu_adm_cols:
        print(f"Error: No se encontró la columna 'Thu Adm' en {archivo}. Columnas disponibles: {list(df.columns)}")
        continue

    thu_adm_col = thu_adm_cols[0]
    
    for _, row in df.iterrows():
        titulo = str(row.get("Title", "")).strip()

        wk = row.get("Wk", -1)

        if pd.isna(titulo) or wk < 1:
            continue
        
        if titulo not in peliculas_data:
            peliculas_data[titulo] = {
                "Dia Inicial": 0, "PrimerFdS": 0, "PrimeraSemanaAdm": 0, "SegundaSemanaAdm": 0, "Total": 0
            }

        thu_adm_value = row.get(thu_adm_col, 0) if not pd.isna(row.get(thu_adm_col, 0)) else 0
        weekend_adm_value = row.get("Weekend\nAdm", 0) if not pd.isna(row.get("Weekend\nAdm", 0)) else 0
        week_adm_value = row.get("Week\nAdm", 0) if not pd.isna(row.get("Week\nAdm", 0)) else 0

        if wk == 1:
            peliculas_data[titulo]["Dia Inicial"] += thu_adm_value
            peliculas_data[titulo]["PrimerFdS"] += weekend_adm_value
            peliculas_data[titulo]["PrimeraSemanaAdm"] += week_adm_value
        elif wk == 2:
            peliculas_data[titulo]["SegundaSemanaAdm"] += week_adm_value

        peliculas_data[titulo]["Total"] += week_adm_value

for titulo, data in peliculas_data.items():
    primera_semana = data["PrimeraSemanaAdm"]
    segunda_semana = data["SegundaSemanaAdm"]
    total = data["Total"]
    restante = total - (primera_semana + segunda_semana)

    data["Restante"] = max(restante, 0)
    data["% PrimeraSemana"] = f"{(primera_semana / total * 100):.2f}%" if total > 0 else "0%"
    data["% SegundaSemana"] = f"{(segunda_semana / total * 100):.2f}%" if total > 0 else "0%"
    data["% Restante"] = f"{(restante / total * 100):.2f}%" if total > 0 else "0%"

output_file = "archivo-4.txt"
with open(output_file, "w") as f:
    f.write(f"{'Titulo':<40} {'Dia Inicial':<12} {'PrimerFdS':<12} {'PrimeraSemanaAdm':<18} {'% PrimeraSemana':<18} {'SegundaSemanaAdm':<18} {'% SegundaSemana':<18} {'Restante':<12} {'% Restante':<12} {'Total':<12}\n")
    f.write("_" * 150 + "\n")

    for titulo, data in peliculas_data.items():
        f.write(f"{titulo:<40} {data['Dia Inicial']:<12} {data['PrimerFdS']:<12} {data['PrimeraSemanaAdm']:<18} {data['% PrimeraSemana']:<18} {data['SegundaSemanaAdm']:<18} {data['% SegundaSemana']:<18} {data['Restante']:<12} {data['% Restante']:<12} {data['Total']:<12}\n")

print(f"Archivo generado: {output_file}")
