import os
import pandas as pd
from collections import defaultdict

def find_duplicate_values_in_excel_files():
    """
    Busca los registros coincidentes en la columna especificada de todos los ficheros Excel
    en la carpeta de ejecución. Genera tres ficheros resumen:
    - mas_rep.txt: con los 15 valores más repetidos.
    - total.txt: con todos los valores repetidos.
    - unicos.txt: con los valores únicos.
    - usuarios.txt: con todos los valores de la columna especificada sin duplicados.
    """

    # Obtener la lista de todos los archivos en el directorio actual
    files = [f for f in os.listdir() if f.endswith('.xlsx') or f.endswith('.xls')]

    # Preguntar al usuario qué columna contiene los valores de interés
    column_name = input("¿Qué nombre de columna contiene los valores de interés?: ")

    # Diccionario para almacenar los valores de la columna especificada de cada archivo
    values_dict = defaultdict(set)
    # Diccionario para contar la cantidad de veces que aparece cada valor
    value_counts = defaultdict(int)
    # Diccionario para almacenar en qué archivos aparece cada valor
    file_occurrences = defaultdict(set)

    for file in files:
        try:
            # Leer el archivo Excel
            df = pd.read_excel(file, dtype=str)
            # Verificar si la columna especificada existe en el DataFrame
            if column_name not in df.columns:
                print(f"La columna '{column_name}' no existe en el archivo {file}.")
                continue
            # Obtener los valores únicos de la columna especificada
            values = df[column_name].dropna().tolist()
            values_dict[file].update(values)
            # Contar las ocurrencias de cada valor y registrar en qué archivos aparece
            for value in values:
                value_counts[value] += 1
                file_occurrences[value].add(file)
        except Exception as e:
            print(f"Error al leer {file}: {e}")

    # Diccionario para almacenar los valores duplicados y los archivos en los que aparecen
    duplicates_dict = {}

    for value, files in file_occurrences.items():
        if len(files) > 1:
            duplicates_dict[value] = files

    # Ordenar los valores por la cantidad de veces que se repiten
    most_common_values = sorted(value_counts.items(), key=lambda item: item[1], reverse=True)[:15]

    # Encontrar los valores que no se repiten en ningún archivo
    unique_values = {value for value, count in value_counts.items() if count == 1}

    # Escribir los valores más repetidos en mas_rep.txt
    with open('mas_rep.txt', 'w') as f:
        f.write("Los 15 valores más repetidos:\n\n")
        for value, count in most_common_values:
            f.write(f"Valor: {value}, Apariciones totales: {count}\n")
            f.write(f"Aparece en los archivos: {', '.join(file_occurrences[value])}\n")
            f.write("\n")

    # Escribir todos los valores repetidos en total.txt
    with open('total.txt', 'w') as f:
        if not duplicates_dict:
            f.write("No se encontraron valores duplicados entre los archivos.\n")
        else:
            for value, files in duplicates_dict.items():
                f.write(f"Valor duplicado: {value}\n")
                f.write(f"Aparece en los archivos: {', '.join(files)}\n")
                f.write(f"Total de apariciones: {value_counts[value]}\n")
                f.write("\n")

    # Escribir los valores únicos en unicos.txt
    with open('unicos.txt', 'w') as f:
        if not unique_values:
            f.write("No se encontraron valores únicos en los archivos.\n")
        else:
            f.write("Valores únicos:\n\n")
            for value in unique_values:
                f.write(f"Valor: {value}\n")
                f.write(f"Aparece en el archivo: {', '.join(file_occurrences[value])}\n")
                f.write("\n")

    print("Resultados escritos en mas_rep.txt, total.txt, unicos.txt y usuarios.txt")

# Ejecutar la función
find_duplicate_values_in_excel_files()
