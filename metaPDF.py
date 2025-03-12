# ProyectoA.com Aplicación Python para obtener, modificar y limpiar metadatos de un fichero PDF
# Versión 1.0

import argparse
from datetime import datetime
#import json
import os.path
from pypdf import PdfReader, PdfWriter

# Procedimiento para obtener los metadatos del fichero PDF y mostrarlos por pantalla
def ObtenerMetadatosPDF(ficheroOrigen):
    if os.path.exists(ficheroOrigen):
        reader = PdfReader(ficheroOrigen)
        
        # Obtenemos los metadatos actuales
        metadatosPDF = reader.metadata
        titulo = metadatosPDF.title
        asunto = metadatosPDF.subject
        autor = metadatosPDF.author
        creador = metadatosPDF.creator
        productor = metadatosPDF.producer
        fechaCreacion = metadatosPDF.creation_date
        fechaModificacion =metadatosPDF.creation_date
        numeroPaginas = len(reader.pages)
                
        if args.verbose:
            if titulo == None:
                print("Título: --VACÍO--")
            else:
                print("Título: " + titulo)
                
            if asunto == None:
                print("Asunto: --VACÍO--")
            else:
                print("Asunto: " + asunto)
                
            if autor == None:
                print("Autor: --VACÍO--")
            else:
                print("Autor: " + autor)
                
            if creador == None:
                print("Creador: --VACÍO--")
            else:
                print("Creador: " + creador)
                
            if productor == None:
                print("Productor: --VACÍO--")
            else:
                print("Productor: " + productor)
            print("Fecha creación: {0}".format(fechaCreacion))
            print("Fecha modificación: {0}".format(fechaModificacion))
            print("Número de páginas: {0}".format(numeroPaginas))
        else: # Si no se elige parámetro verbose, mostramos los datos en formato JSON
            metadatosJSON = '{{"titulo": "{0}","asunto": "{1}","autor": "{2}","creador": "{3}","productor": "{4}","fecha_creacion": "{5}","fecha_modificacion": "{6}","numero_paginas": "{7}"}}'.format(
                titulo, asunto, autor, creador, productor, fechaCreacion, fechaModificacion, numeroPaginas)
            # metadatosJSON = json.dumps(metadatosJSON)
            print(metadatosJSON)
        return 0
    else:
        if args.verbose:
            print("No existe el fichero PDF: {0}".format(ficheroOrigen))
        return 100
        
# Procedimiento para modificar los metadatos de un fichero PDF
# def ModificarMetadatosPDF(ficheroOrigen, ficheroSalida, titulo, autor, asunto, 
#                          productor, creador, palabrasClave, camposPersonalizados, fecha):
def ModificarMetadatosPDF(ficheroOrigen, ficheroSalida, metadatosPDF, fecha):
    if os.path.exists(ficheroOrigen):
        reader = PdfReader(ficheroOrigen)
        writer = PdfWriter()

        # Añadimos todas las páginas para generar el nuevo PDF con los nuevos metadatos
        for page in reader.pages:
            writer.add_page(page)

        # Formatear fecha
        utc_time = "+01'00'"  # UTC time optional
        fechaMod = datetime.now().strftime(f"D\072%Y%m%d%H%M%S{utc_time}")
        # fechaMod = datetime.now()

        if fecha != None:
            fechaMod = fecha

        writer.add_metadata(
            {
                "/Author": metadatosPDF["autor"],
                "/Producer": metadatosPDF["productor"],
                "/Title": metadatosPDF["titulo"],
                "/Subject": metadatosPDF["asunto"],
                "/Keywords": metadatosPDF["palabras_clave"],
                "/CreationDate": fechaMod,
                "/ModDate": fechaMod,
                "/Creator": metadatosPDF["aplicacion"],
                "/CustomField": metadatosPDF["campos_personalizados"],
            }
        )        
                    
        # Guardar el nuevo fichero PDF, si no se ha indicado fichero de salida, se reemplaza el original
        if ficheroSalida != "":
            ficheroOrigen = ficheroSalida
        try:
            # Si existe el fichero de salida (sea el original o el que se creará)
            # Y no se ha añadido el parámetro "reemplazar", no se realizarán modificaciones
            if not os.path.exists(ficheroOrigen) or (os.path.exists(ficheroOrigen) and args.reemplazar):
                with open(ficheroOrigen, "wb") as f:
                    writer.write(f)
                    if args.verbose:
                        print("Metadatos del fichero {0} modificados correctamente".format(ficheroOrigen))
                    return 0
            else:
                if args.verbose:
                    print("No se han realizado cambios, el fichero de salida existe y no se ha incluido el parámetro \"reemplazar\"")
                return 103
                
        except PermissionError:
            if args.verbose:
                print("El fichero destino está abierto o no tiene permisos para escribir en la carpeta destino")
            return 101
        except Exception as ex:
            if args.verbose:
                print("Se ha producido un error al generar el fichero PDF: {0}".format(getattr(ex, 'message', str(ex))))
            return 102            
    else:
        if args.verbose:
            print("No existe el fichero PDF origen: {0}".format(ficheroOrigen))
        return 100            

# Procedimiento para limpiar (vaciar) los metadatos de un fichero PDF
def LimpiarMetadatosPDF(ficheroOrigen, ficheroSalida):
    metadatosPDF = {'titulo': '', 'asunto': '','autor': '','productor': '','aplicacion': '','palabras_clave': '','campos_personalizados': ''}
    ModificarMetadatosPDF(ficheroOrigen, ficheroSalida, metadatosPDF, None)
    
# Conformamos los argumentos que admitirá el programa por la línea de comandos
# metadatosPDF = {'titulo': '', 'asunto': '','autor': '','productor': '','aplicacion': '','palabras_clave': '','campos_personalizados': ''}
# metadatosPDF = {}
parser = argparse.ArgumentParser()
parser.add_argument("-fo", "--fichero_PDF_Origen", type=str, required=True,
    help="Ruta y nombre del fichero PDF origen para obtener/modificar/limpiar los metadatos")
parser.add_argument("-fd", "--fichero_PDF_Destino", type=str, required=False,
    help="Ruta y nombre del fichero PDF destino. Si no se indica, se reemplazará el original (si se incluye el parámetro \"reemplazar\")")
parser.add_argument("-om", "--obtener", action="store_true", 
    help="Obtener y mostrar los metadatos del fichero PDF")
parser.add_argument("-mm", "--modificar", action="store_true",
    help="Modificar los metadatos del fichero PDF")
parser.add_argument("-lm", "--limpiar", action="store_true",
    help="Limpiar (vaciar) los metadatos del fichero PDF")
parser.add_argument("-r", "--reemplazar", action="store_true", 
    help="Si existe el fichero destino se reemplazará, si no se añade este parámetro, si existe el fichero destino, no se realizarán modificaciones")
parser.add_argument("-me", "--metadatos", required=False, type=str, 
    help="Si se usa el parámetro \"modificar\" se le pasarán los metadatos a añadir en este parámetro, con el formato 'titulo=valor,asunto=valor,autor=valor,productor=valor,aplicacion=valor,palabras_clave=valor,campos_personalizados=valor'")
parser.add_argument("-v", "--verbose", required=False, action="store_true", 
    help="Se activará el modo verbose (mostrar por pantalla todos los resultados)")

args = parser.parse_args()

if args.metadatos:
    metadatosPDF = dict( pair.split('=') for pair in args.metadatos.split(',') )

# Si se pasa parámetro de optención de metadatos
if args.obtener:
    ObtenerMetadatosPDF(args.fichero_PDF_Origen)
if args.modificar:
    if args.metadatos is None:
        print ("Si usa el parámetro \"modificar\" debe añadir el parámetro \"metadatos\" para indicar los metadatos a añadir")
    else:
        if args.fichero_PDF_Destino:
            ModificarMetadatosPDF(args.fichero_PDF_Origen, args.fichero_PDF_Destino, metadatosPDF, None)
        else:
            ModificarMetadatosPDF(args.fichero_PDF_Origen, "", metadatosPDF, None)
if args.limpiar:
    if args.fichero_PDF_Destino:
        LimpiarMetadatosPDF(args.fichero_PDF_Origen, args.fichero_PDF_Destino)
    else:
        LimpiarMetadatosPDF(args.fichero_PDF_Origen, "")