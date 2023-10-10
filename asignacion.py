import pandas as pd
import datetime
import locale
from datetime import date
import tkinter as tk
from tkinter import messagebox
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
# Leer el archivo de texto con los datos

with open('Pases.txt', 'r') as archivo:
    # Leer todas las líneas del archivo y almacenarlas en una lista
    lineas = archivo.readlines()



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
print("Archivo Pases.txt procesado")


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
    'CP': dfMaecli['CP']
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

for columna in columnas_a_actualizar:
    df_intermedio[f'{columna}O'] = df_intermedio.apply(buscar_datos, axis=1, args=(columna,))    

# Aplicar la función de búsqueda a cada fila de df_intermedio y guardar los resultados en una nueva columna
df_intermedio['DNIO'] = df_intermedio.apply(buscar_dni, axis=1)


# df_intermedio ahora contendrá las columnas actualizadas con los datos correspondientes

dfPases['Calle'] = df_intermedio['CalleO']
dfPases['Numero'] = df_intermedio['NumeroO']
dfPases['Piso/Depto'] = df_intermedio['Piso/DeptoO']
dfPases['Cp'] = df_intermedio['CPO']
dfPases['DNI'] = df_intermedio['DNIO']

print("Archivo MAECLI.txt procesado")

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


# Agrupar por la columna 'Col2' y combinar los valores correspondientes de la columna 'Col4'
grouped = dfMaetel.groupby('DNI')['Tel'].apply('-'.join).reset_index()

# Dividir la columna combinada 'Col4' en 'Cel 1' y 'Cel 2'
grouped[['Cel 1', 'Cel 2']] = grouped['Tel'].str.split('-', n=1, expand=True)

# Eliminar la columna 'Col4' que ya no es necesaria
grouped.drop(columns=['Tel'], inplace=True)

# Renombrar las columnas
grouped.rename(columns={'DNI': 'Dni'}, inplace=True)


# Combina las columnas 'Cel 1' y 'Cel 2' de grouped en df_intermedio cuando coinciden los valores de 'DniO' y 'Dni'
df_intermedio['Cel 1'] = df_intermedio['DNIO'].map(grouped.set_index('Dni')['Cel 1'])
df_intermedio['Cel 2'] = df_intermedio['DNIO'].map(grouped.set_index('Dni')['Cel 2'])

print("Archivo MAETEL.txt procesado")

# Esto agrega las columnas 'Cel 1' y 'Cel 2' al DataFrame df_intermedio cuando coinciden los valores de DniO y Dni

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
    'col31': datetime.date.today(),
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

# Esto crea el DataFrame dfFinal con los datos especificados y lo guarda en un archivo Excel.

# Esto crea el DataFrame dfFinal con los datos especificados y lo guarda en un archivo Excel.
def mostrar_alerta_exito():
    messagebox.showinfo("Éxito", "Archivo creado con éxito")

# Crea una ventana en blanco
ventana = tk.Tk()
ventana.withdraw()  # Oculta la ventana principal

# Llama a la función para mostrar la alerta de éxito
mostrar_alerta_exito()

# Cierra la ventana después de mostrar la alerta
ventana.destroy()