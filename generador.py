import pandas as pd
import os
from utils import *

# LEO EL EXCEL CON LOS DATOS TABULADOS
archivo = pd.read_excel('./Manual_Procedimientos.xlsx',
                        sheet_name="Documentos")

# HAGO UNA LISTA DE LOS DOCUMENTOS EXISTENTES EN LA CARPETA
lista_doc_existentes = os.listdir('./documentos')
# HAGO UNA LISTA VACÍA PARA GUARDAR TODOS LOS DOC A CREAR SEGUN EL EXCEL
lista_doc_excel = []
# SE CREAN DOS LISTA CON LOS CODIGOS Y LOS NOMBRES
lista= archivo['codigo'].unique()
lista_list = lista.tolist() # pasa lista que es np array a lista
lista2= archivo['nombre'].unique()
# SE HACE UN FOR PARA CREAR UNA LISTA CON EL CODIGO-NOMBRE
# DE LOS DOCUMENTOS QUE ESTAN EN EL EXCEL
for l1, l2 in zip (lista, lista2):
    lista_doc_excel.append(l1+'-'+l2+'.docx')

# CON ESTE BUCLE RECORRO LA LISTA DE DOCUMENTOS EN EL EXCEL
# Y SI NO ESTA EN LA CARPETA, ENTONCES LO CREA CON LA FUNCIÓN GENERADOR
# SI EL DOCUMENTO ESTA CREADO, SE ACTUALIZA CON LA FUNCIÓN ACTUALIZAR
for n in lista_doc_excel:
    codigo = n[:5] # SI EL CÓDIGO CAMBIA DE FORMATO LL-NN, HAY QUE REVISAR ACA
    if n not in lista_doc_existentes:
        generador_doc(codigo, archivo)
        print('Creo el archivo ', codigo)
    else:
        actualizacion(n, archivo)
        print(codigo, ' ya se actualizo')
# SE ACTUALIZA LA LISTA DE DOCUMENTOS EXISTENTES PARA AGREGAR 
# LOS HIPERVÍNCULOS A TODOS LOS DOCUMENTOS DE UNA VEZ
lista_doc_existentes = os.listdir('./documentos')
for doc_exist in lista_doc_existentes:
    document = docx.Document('./Documentos/' + doc_exist)
    agrega_hipervinculos(document, archivo, doc_exist, lista_list, lista2)

print('## FINALIZÓ DE AGREGAR HIPERVINCULOS ##')