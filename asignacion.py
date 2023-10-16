import pandas as pd
import os
import openpyxl
import datetime
import locale
from datetime import date
import tkinter as tk
from tkinter import messagebox
from tqdm import tqdm

# Leer el archivo de texto con los datos

print("Leyendo archivos...")
with open('Pases.txt', 'r') as archivo:
    # Leer todas las líneas del archivo y almacenarlas en una lista
    lineas = archivo.readlines()

# Eliminar los saltos de línea adicionales después de la última línea
while lineas[-1].strip() == '':
    lineas.pop()

# Inicializa una lista para almacenar los datos de cada fila
datos_filas = []

# Función para procesar una línea y sumar palabras hasta encontrar un dígito
def procesar_linea(linea):
    columnas = linea.split()
    columna_1 = columnas[0]
    columna_2 = columnas[1]
    columna_3 = ""
    columna_4 = ""

    for valor in columnas[2:]:
        if not valor.isdigit() :
            columna_3 += valor + " "
        else:
            columna_4 = valor
            break
        
    columna_3 = columna_3.strip()
    
    num_palabras_columna_3 = len(columna_3.split()) - 1
    
    columna_5 = ""
    
    # Combina palabras en la columna 5 hasta encontrar un dígito
    for valor in columnas[4 + num_palabras_columna_3:]:
        if not valor[0].isdigit() or valor == "9":  # Verifica si el primer carácter es un dígito
            columna_5 += valor + " "
        else:
            break
    
    columna_5 = columna_5.strip()

    num_palabras_columna_5 = len(columna_5.split()) - 1
    columna_5 = columna_5.strip()
    
    columna_6 = columnas[5 + num_palabras_columna_3 + num_palabras_columna_5]
    columna_7 = columnas[6 + num_palabras_columna_3 + num_palabras_columna_5]
    columna_8 = columnas[7 + num_palabras_columna_3 + num_palabras_columna_5]
    columna_9 = columnas[8 + num_palabras_columna_3 + num_palabras_columna_5]

    columna_10 = ""



    # Identifica el valor numérico en 'Saldo Total'
    # Identifica el valor numérico en 'Saldo Total'
    for valor in columna_8.split():
        if valor.replace(',', '').replace('.', '').isdigit():
            columna_8 = valor
        break

    # Combina palabras en 'Paquete'
    paquete_encontrado = False

    for valor in columnas[9 + num_palabras_columna_3 + num_palabras_columna_5:]:
        if paquete_encontrado:
            columna_10 += valor + " "
        elif not any(char.isdigit() for char in valor):
            columna_10 += valor + " "
        else:
            paquete_encontrado = True

    columna_10 = ''.join(filter(lambda char: not char.isdigit(), columna_10)).strip()
    columna_10 = columna_10.strip().lstrip()

    
    return [columna_1.replace("'", ''), 
            columna_2, 
            columna_3, 
            columna_4, 
            columna_5.replace(',', '').replace('.', '').replace("'", ''), 
            columna_6, 
            columna_7.replace(',', '').replace('.', '').replace("'", ''), 
            columna_8.replace(',', '').replace('.', '').replace("'", ''), 
            columna_9.strip().replace(',', '').replace('.', '').replace("'", ''), 
            columna_10.replace('/',"").strip()]

# Procesa cada línea y agrega los datos a la lista
for linea in lineas:
    datos_filas.append(procesar_linea(linea))


# Define los nombres de las columnas
nombres_columnas = ['Nro Legajo', 'Apellido', 'Nombre', 'Nro Sucursal', 'Sucursal', 'Fecha de Pase',
                    'Saldo FavaCard', 'Saldo FavaNet', 'Saldo Total', 'Cuenta']

# Crea el DataFrame

dfPases = pd.DataFrame(datos_filas, columns=nombres_columnas)

nombre_de_columna = 'Nro Legajo'  # Reemplaza 'nombre_de_columna' con el nombre real de la columna que deseas contar
cantidad_de_filas = dfPases[nombre_de_columna].count()


# Tu código para crear el DataFrame dfPases


# Leer el archivo de texto con los datos de MAECLI.TXT
maecli_data = []
with open('BRIA_MAECLI.TXT', 'r') as maecli_file:
    for line in maecli_file:
        
        values = line.strip().split(',')

        maecli_data.append([values[1].strip() , values[2].strip('"').strip(), values[3].strip('"').strip(), values[4].strip('"').strip(), values[5].strip('"').strip(), values[7].strip()])



# Tu código para cargar dfPases y maecli_data como listas
dfMaecli = pd.DataFrame(maecli_data, columns=['DNI', 'Nombre', 'Calle', 'Numero', 'Piso/Depto', 'CP'])


# Crear el DataFrame intermedio
df_intermedio = pd.DataFrame({
    'Nombres1': dfPases['Apellido'] + ' ' + dfPases['Nombre'],
    'Nombres2': dfMaecli['Nombre'],
    'DNI': dfMaecli['DNI'],
    'Calle': dfMaecli['Calle'],
    'Numero': dfMaecli['Numero'],
    'Piso/Depto': dfMaecli['Piso/Depto'],
    'CP': dfMaecli['CP'],
    'Nro Legajo': dfPases['Nro Legajo']
})


# Definir una función para buscar los datos correspondientes y devolverlos
def buscar_dni(row):
    nombre1 = row['Nombres1']
    nombre2 = row['Nombres2'].replace(" ", "")

    if pd.isna(nombre1):
        return None
    
    if isinstance(nombre1, str):
        nombre1 = nombre1.replace(" ", "")
    
    # Filtrar las filas que cumplen con la condición
    filtro = df_intermedio['Nombres2'].str.replace(" ", "") == nombre1
    dni_correspondiente = df_intermedio[filtro]['DNI'].values
    
    if len(dni_correspondiente) > 0:
        return dni_correspondiente[0]
    else:
        return None

    

def buscar_datos(row, columna):
    nombre1 = row['Nombres1']
    nombre2 = row['Nombres2'].replace(" ", "")

    if pd.isna(nombre1):
        return None
    
    if isinstance(nombre1, str):
        nombre1 = nombre1.replace(" ", "")

    # Filtrar las filas que cumplen con la condición
    filtro = df_intermedio['Nombres2'].str.replace(" ", "") == nombre1
    datos_correspondientes = df_intermedio[filtro][columna].values

    if len(datos_correspondientes) > 0:
        return datos_correspondientes[0]
    else:
        return None




# Columnas a buscar y crear en df_intermedio
columnas_a_actualizar = ['Calle', 'Numero', 'Piso/Depto', 'CP']

for columna in tqdm(columnas_a_actualizar, desc="OBTENIENDO DATOS"):
    df_intermedio[f'{columna}O'] = df_intermedio.apply(buscar_datos, axis=1, args=(columna,))    

# Aplicar la función de búsqueda a cada fila de df_intermedio y guardar los resultados en una nueva columna
df_intermedio['DNIO'] = df_intermedio.apply(buscar_dni, axis=1)


# df_intermedio ahora contendrá las columnas actualizadas con los datos correspondientes

dfPases['Calle'] = df_intermedio['CalleO']
dfPases['Numero'] = df_intermedio['NumeroO']
dfPases['Piso/Depto'] = df_intermedio['Piso/DeptoO']
dfPases['Cp'] = df_intermedio['CPO']
dfPases['DNI'] = df_intermedio['DNIO']


# # Leer el archivo de texto con los datos de MAETEL.TXT
maetel_data = []

valueDni = df_intermedio['DNIO']
with open('BRIA_MAETEL.TXT', 'r') as maetel_file:
    for line in maetel_file:
        
        values = line.strip().split(',')

        maetel_data.append([values[1].strip() , values[2].strip('"').strip(), values[3].strip('"').strip()])
        
# Tu código para cargar dfPases y maecli_data como listas

# Convertir maetel_data en un DataFrame
dfMaetel = pd.DataFrame(maetel_data, columns=['DNI', 'Tipo', 'Tel'])
dfMaetel['DniO'] = df_intermedio['DNIO']
# Create a custom sorting function
def custom_sort(tel_list):
    if 'Celular Principal' in tel_list:
        tel_list.remove('Celular Principal')
        tel_list.insert(0, 'Celular Principal')
    return '-'.join(tel_list)

# Group by 'DNI' and aggregate 'Tel' values using the custom sorting function
grouped = dfMaetel.groupby('DNI')['Tel'].agg(list).reset_index()
grouped['Tel'] = grouped['Tel'].apply(custom_sort)

# Dividir la columna combinada 'Col4' en 'Cel 1' y 'Cel 2'
grouped[['Cel 1', 'Cel 2']] = grouped['Tel'].str.split('-', n=1, expand=True)

# Eliminar la columna 'Col4' que ya no es necesaria
grouped.drop(columns=['Tel'], inplace=True)

# Renombrar las columnas
grouped.rename(columns={'DNI': 'Dni'}, inplace=True)


# Combina las columnas 'Cel 1' y 'Cel 2' de grouped en df_intermedio cuando coinciden los valores de 'DniO' y 'Dni'
df_intermedio['Cel 1'] = df_intermedio['DNIO'].map(grouped.set_index('Dni')['Cel 1'])
df_intermedio['Cel 2'] = df_intermedio['DNIO'].map(grouped.set_index('Dni')['Cel 2'])


# Lee el archivo de texto y crea el DataFrame
# Lee el archivo de texto y divide las líneas en elementos
with open('BRIA_MAEREL.txt', 'r') as file:
    lines = file.readlines()

# Inicializa listas para almacenar los datos de Legajo y DNI
legajo = []
dni = []

# Recorre las líneas del archivo
for line in lines:
    # Divide cada línea en elementos separados por comas
    elements = line.split(',')
    
    # Añade el primer elemento a la lista de Legajo y el cuarto elemento a la lista de DNI
    legajo.append(str(elements[0].strip('"')))
    dni.append(elements[3].strip())

# Crea un DataFrame a partir de las listas de Legajo y DNI
dfMaerel = pd.DataFrame({'Legajo': legajo, 'DNI': dni})

total = len(grouped)
pbar = tqdm(total=total, desc='ORDENANDO...')

for index, row in grouped.iterrows():
    pbar.update(1)
    dni_grouped = row['Dni']
    matching_legajo = dfMaerel[dfMaerel['DNI'] == dni_grouped]

    if not matching_legajo.empty:
        legajo_grouped = str(matching_legajo['Legajo'].iloc[0])

        matching_row_intermedio = df_intermedio[df_intermedio['Nro Legajo'] == legajo_grouped]

        if not matching_row_intermedio.empty:
            cel1_grouped = row['Cel 1']
            cel1_intermedio = df_intermedio.loc[df_intermedio['Nro Legajo'] == legajo_grouped, 'Cel 1'].iloc[0]

            if not pd.isna(cel1_grouped) and cel1_grouped != cel1_intermedio:
                if cel1_intermedio is None:
                    cel1_intermedio = ""

                # Verificar si las cadenas no están contenidas una en la otra
                if not all(part in cel1_intermedio.split('-') for part in cel1_grouped.split('-')):
                    df_intermedio.loc[df_intermedio['Nro Legajo'] == legajo_grouped, 'Cel 1'] = cel1_intermedio + '-' + cel1_grouped

            cel2_grouped = row['Cel 2']
            cel2_intermedio = df_intermedio.loc[df_intermedio['Nro Legajo'] == legajo_grouped, 'Cel 2'].iloc[0]

            if not pd.isna(cel2_grouped) and cel2_grouped != cel2_intermedio:
                if cel2_intermedio is None:
                    cel2_intermedio = ""

                # Verificar si las cadenas no están contenidas una en la otra
                if not all(part in cel2_intermedio.split('-') for part in cel2_grouped.split('-')):
                    df_intermedio.loc[df_intermedio['Nro Legajo'] == legajo_grouped, 'Cel 2'] = cel2_grouped + '-' + cel2_intermedio
pbar.close()

for index, row in df_intermedio.iterrows():
    valores_cel1 = set(str(row['Cel 1']).split('-')) if '-' in str(row['Cel 1']) else {str(row['Cel 1'])}
    if not pd.isna(row['Cel 2']):  # Verifica si cel2 no está vacío
        valores_cel2 = set(str(row['Cel 2']).split('-')) if '-' in str(row['Cel 2']) else {str(row['Cel 2'])}
        valores_cel2_nuevos = valores_cel2 - valores_cel1
        df_intermedio.at[index, 'Cel 2'] = '-'.join(valores_cel2_nuevos)

fecha_actual = date.today()
mes_actual = fecha_actual.strftime("%B")
ano_actual = fecha_actual.year

fecha_formateada = fecha_actual.strftime("%d/%m/%Y")

# Crear un diccionario con los datos
data = {
    'col1': "",
    'col2': df_intermedio['DNIO'],
    'col3': dfPases['Apellido'],
    'col4': dfPases['Nombre'],
    'col5': df_intermedio['CalleO'],
    'col6': df_intermedio['NumeroO'],
    'col7': df_intermedio['Piso/DeptoO'],
    'col8': "",
    'col9': "",
    'col10': df_intermedio['CPO'],
    'col12': dfPases['Sucursal'],
    'col13': "BUENOS AIRES",
    'col14': "ARGENTINA",
    'col15': "LAB",
    'col16': df_intermedio['CalleO'],
    'col17': df_intermedio['NumeroO'],
    'col18': df_intermedio['Piso/DeptoO'],
    'col19': "",
    'col20': "",
    'col21': df_intermedio['CPO'],
    'col22': dfPases['Sucursal'],
    'col23': "BUENOS AIRES",
    'col24': "ARGENTINA",
    'col25': "CEL 1",
    'col26': df_intermedio['Cel 1'],
    'col27': "CEL 2",
    'col28': df_intermedio['Cel 2'],
    'col29': "",
    'col30': "",
    'col31': datetime.date.today().strftime('%d/%m/%Y'),
    'col32': "TARJETAS",
    'col33': dfPases['Nro Sucursal'],
    'col34': dfPases['Sucursal'],
    'col35': dfPases['Fecha de Pase'],
    'col36': dfPases['Saldo Total'],
    'col37': dfPases['Nro Legajo']
}

# Crear el DataFrame dfFinal
dfFinal = pd.DataFrame(data)

# Restablece los índices del DataFrame
dfFinal.reset_index(drop=True, inplace=True)

def eliminar_filas_despues_de_n(df, col1, n):
    if n < 0 or n >= len(df):
        return df  # No se eliminan filas si n está fuera del rango válido
    
    df = df.iloc[:n]  # Conserva las filas hasta el índice n (incluyendo la fila n)
    return df

dfFinal = eliminar_filas_despues_de_n(dfFinal, 'col1', cantidad_de_filas)

nombre_archivo = f"Asignacion {mes_actual} {ano_actual}.xlsx"

# Guardar el DataFrame en un archivo Excel
dfFinal.to_excel(nombre_archivo, index=False, header=False)
print("ARCHIVO GENERADO CON EXITO..")

# Verifica si el archivo existe
if os.path.isfile(nombre_archivo):
    # Abre el archivo Excel
    excel_app = openpyxl.load_workbook(nombre_archivo)
    
    # Abre la primera hoja del libro de trabajo
    sheet = excel_app.active
    
    # Cierra el archivo para que pueda abrirse en Excel
    excel_app.close()
    
    # Abre el archivo en la aplicación predeterminada (Excel en este caso)
    os.system(f'start excel "{nombre_archivo}"')
else:
    print(f"El archivo '{nombre_archivo}' no se encontró.")
