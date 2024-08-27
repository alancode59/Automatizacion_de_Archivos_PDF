#Proceso de Automatización de Archivos PDF: Renombramiento, Creación y Eliminación de Caracteres Especiales

#Importación de librerías


import os
import unicodedata
import re
import win32com.client

#Listado de directorios
directorios = [
    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\M.V',
    'C:\\Users\\Alan\\0.- M.V para subir JULIO\\M.V',

    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\M.V',
    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\\M.V Mayo',

    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\MV',
    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\MV',

    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\MV',
    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\M.V',

    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\MV',
    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\M.V',

    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\M.D',
    'C:\\Users\\Alan\\Desktop\\0.- M.V para subir JULIO\\MEDIOS DE VERIFICACIÓN'


]


#función para quitar acentos, caracteres especiales
def normalizar_nombre(nombre):
    nombre = re.sub(r'\s+', ' ', nombre)  #reemplaza espacios dobles por uno solo
    nombre_normalizado = unicodedata.normalize('NFKD', nombre).encode('ASCII', 'ignore').decode('ASCII')
    nombre_sin_caracteres_especiales = re.sub(r'[^A-Za-z0-9_. -]+', '', nombre_normalizado)  #elimina caracteres especiales - excepto guiones y puntos
    return nombre_sin_caracteres_especiales




#función para convertir archivos a .pdf y eliminar el original
def convertir_a_pdf(ruta_archivo):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(ruta_archivo)
    ruta_pdf = os.path.splitext(ruta_archivo)[0] + ".pdf"
    doc.SaveAs(ruta_pdf, FileFormat=17)
    doc.Close()
    word.Quit()
    os.remove(ruta_archivo)  #eliminar el archivo original

for directorio in directorios:
    elementos = os.listdir(directorio)
    cambios_name = {}

    #renombrar los archivos en el directorio actual
    for elemento in elementos:
        nombre_base, extension = os.path.splitext(elemento)
        nombre_base_normalizado = normalizar_nombre(nombre_base)
        
        #reemplaza guiones medios por guiones bajos
        nombre_base_normalizado = nombre_base_normalizado.replace("-", "_")
        
        #reemplaza espacios y puntos por guiones bajos
        nuevo_nombre_base = nombre_base_normalizado.replace(" ", "_").replace(".", "_")
        
        
         #reemplaza múltiples guiones bajos por uno solo
        nuevo_nombre_base = re.sub(r'_+', '_', nuevo_nombre_base) 
        nuevo_nombre = nuevo_nombre_base + extension

        if nuevo_nombre != elemento:
            cambios_name[elemento] = nuevo_nombre
            ruta_vieja = os.path.join(directorio, elemento)
            ruta_nueva = os.path.join(directorio, nuevo_nombre)
            os.rename(ruta_vieja, ruta_nueva)
            print(f'Renombrado: {elemento} a {nuevo_nombre}')

    #actualizar la lista 
    elementos = os.listdir(directorio)

    #convertir archivos que no sean .pdf
    for elemento in elementos:
        nombre_base, extension = os.path.splitext(elemento)
        ruta_archivo = os.path.join(directorio, elemento)
        if extension.lower() != '.pdf':
            convertir_a_pdf(ruta_archivo)

print("Cambios realizados :) ")
