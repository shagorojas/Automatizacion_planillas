
# Importamos las librerías necesarias
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage  # Renombramos Image para evitar conflicto con PIL
from PIL import Image as PILImage  # Renombramos también Image de PIL
import pandas as pd
import xlsxwriter
import pathlib
import shutil
import time
import os

# Definir rutas
ruta_master = os.path.join(str(os.path.abspath(pathlib.Path().absolute())))
ruta_parametros = os.path.join(ruta_master, "Insumo", "Parametros.xlsx")
ruta_json = os.path.join(ruta_master, "Config", "Config.json")
ruta_log = os.path.join(ruta_master, "Log", "Eventos.log")
ruta_resultado = os.path.join(ruta_master, "Resultado")

# Rutas imagenes
logo_alimentos = os.path.join(ruta_master, "util", "Logo alimentos.png")
logo_operador = os.path.join(ruta_master, "util", "Logo operador.png")
logo_secretaria = os.path.join(ruta_master, "util", "Logo secretaria faca.png")

# Cargamos los parametros
df_parametros = pd.read_excel(ruta_parametros)

# Convertir a diccionario
dict_data = dict(zip(df_parametros["Concepto"], df_parametros["Valor"]))


# Cargamos los parametros por variables
departamento = dict_data["Departamento"]
municipio = dict_data["Municipio"]
operador = dict_data["Operador"]
contrato = dict_data["Contrato No."]
codigo_dane = dict_data["Codigo dane"]
codigo_dane_completo = dict_data["Codigo dane completo"]
jornada = dict_data["Jornada"]
institucion = dict_data["Institucion"]
dane_institucion = dict_data["Codigo dane institucion"]
mes_atencion = dict_data["Mes de atencion"]
anio = dict_data["Año"]

def crear_plantilla_control(df_plantilla, nombre_df):

    # Crear un DataFrame vacío
    df = pd.DataFrame(index=range(10), columns=[chr(65 + i) for i in range(35)])

    # Guardar el DataFrame en un archivo Excel
    archivo_excel = f"Resultado\\{nombre_df}.xlsx"
    df.to_excel(archivo_excel, index=False, engine='xlsxwriter')

    # Crear una conexión con el archivo Excel y agregar las imágenes
    writer = pd.ExcelWriter(archivo_excel, engine='xlsxwriter')
    df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')

    # Acceder al objeto workbook y worksheet
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Insertar las imágenes
    worksheet.insert_image('AD1', logo_alimentos)
    worksheet.insert_image('A1', logo_operador)
    worksheet.insert_image('D1', logo_secretaria)

    # Combinar celdas de A8 a AI8
    worksheet.merge_range('A8:AI8', 'Formato - REGISTRO Y CONTROL DIARIO DE ASISTENCIA', 
                            workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'fg_color': 'black',   # Color de fondo negro
                            'font_color': 'white',  # Color de texto blanco
                            'font_size': 12        # Tamaño de fuente
                        }))

    # Combinar celdas A9 y B9 y agregar el texto "DEPARTAMENTO" en negrita
    worksheet.merge_range('A9:B9', 'DEPARTAMENTO:', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 12
                        }))

    # Definir el formato con borde inferior negro
    borde_inferior_negro = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12,
        'bottom': 1,  # Borde inferior negro
    })

    # Definir el formato con borde inferior negro
    borde_inferior = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12,
        'bottom': 1,  # Borde inferior negro
    })

    # Combinar celdas C9 y D9 y agregar el texto de la variable "departamento"
    worksheet.merge_range('C9:D9', departamento, borde_inferior_negro)

    # Combinar celdas A10 y B10 y agregar el texto "MUNICIPIO" en negrita
    worksheet.merge_range('A10:B10', 'MUNICIPIO:', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 12
                        }))

    # Combinar celdas C9 y D9 y agregar el texto de la variable "municipio"
    worksheet.merge_range('C10:D10', municipio , borde_inferior_negro)

    # Combinar celdas A11 y B11 y agregar el texto "OPERADOR" en negrita
    worksheet.merge_range('A11:B11', 'OPERADOR:', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 12
                        }))

    # Combinar celdas C11 y D11 y agregar el texto de la variable "operador"
    worksheet.merge_range('C11:D11', operador , borde_inferior_negro)

    # Combinar celdas A12 y B12 y agregar el texto "CONTRATO No" en negrita
    worksheet.merge_range('A12:B12', 'CONTRATO No:', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 12
                        }))

    # Combinar celdas C12 y D12 y agregar el texto de la variable "contrato"
    worksheet.merge_range('C12:D12', contrato , borde_inferior_negro)

    # Combinar celdas G9 y L9 y agregar el texto "MUNICIPIO" en negrita
    worksheet.merge_range('G9:L9', 'NOMBRE DE INSTITUCIÓN O CENTRO EDUCATIVO:', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 12
                        }))

    # Combinar celdas G10 y L10 y agregar el texto "MUNICIPIO" en negrita
    worksheet.merge_range('G10:L10', 'CODIGO DANE INSTITUCIÓN O CENTRO EDUCATIVO:', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 12
                        }))

    # Combinar celdas G11 y J11 y agregar el texto "MUNICIPIO" en negrita
    worksheet.merge_range('G11:J11', 'MES DE ATENCIÓN:', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 12
                        }))

    # Combinar celdas R11 y S11 y agregar el texto "MUNICIPIO" en negrita
    worksheet.merge_range('R11:S11', 'AÑO:', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 12
                        }))

    # Combinar celdas S9 y AE9 y agregar el texto de la variable "institucion"
    worksheet.merge_range('S9:AE9', institucion , borde_inferior)
    # Combinar celdas S10 y AE10 y agregar el texto de la variable "institucion"
    worksheet.merge_range('S10:AE10', dane_institucion , borde_inferior)
    # Combinar celdas K11 y P11 y agregar el texto de la variable "institucion"
    worksheet.merge_range('K11:P11', mes_atencion , borde_inferior)
    # Combinar celdas T11 y X11 y agregar el texto de la variable "institucion"
    worksheet.merge_range('T11:X11', anio , borde_inferior)


    # Escribir "CÓDIGO DANE:" en E9 y B12 en negrita, sin combinar celdas
    worksheet.write('E9', 'CÓDIGO DANE:', workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12
    }))

    worksheet.write('E10', 'CÓDIGO DANE:', workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12
    }))

    worksheet.write('E11', 'JORNADA:', workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12
    }))

    # Escribir el valor de la variables en las celdas respectivas
    worksheet.write('F9', codigo_dane, borde_inferior)
    worksheet.write('F10', codigo_dane_completo, borde_inferior)
    worksheet.write('F11', jornada, borde_inferior)

    # Combinar las celdas A14, A15 y A16, agregar el texto "N°" con el color de fondo #A6A6A6
    worksheet.merge_range('A14:A16', 'N°', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6'  # Color de fondo
                        }))
    # Combinar las celdas agregar el texto "TIPO DE DOCUMENTO" con el color de fondo #A6A6A6
    worksheet.merge_range('B14:B16', 'TIPO DE DOCUMENTO', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'text_wrap': True       # Ajustar el texto para que se ajuste dentro de la celda
                        }))
    # Combinar las celdas agregar el texto "NÚMERO DE DOCUMENTO DE IDENTIDAD" con el color de fondo #A6A6A6
    worksheet.merge_range('C14:C16', 'NÚMERO DE DOCUMENTO DE IDENTIDAD', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'text_wrap': True       # Ajustar el texto para que se ajuste dentro de la celda
                        }))
    # Combinar las celdas agregar el texto "PRIMER SIMBRE DEL TITULAR DE DERECHO" con el color de fondo #A6A6A6
    worksheet.merge_range('D14:D16', 'PRIMER SIMBRE DEL TITULAR DE DERECHO', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'text_wrap': True       # Ajustar el texto para que se ajuste dentro de la celda
                        }))
    # Combinar las celdas agregar el texto "SEGUNDO SIMBRE DEL TITULAR DE DERECHO" con el color de fondo #A6A6A6
    worksheet.merge_range('E14:E16', 'SEGUNDO SIMBRE DEL TITULAR DE DERECHO', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'text_wrap': True       # Ajustar el texto para que se ajuste dentro de la celda
                        }))
    # Combinar las celdas agregar el texto "PRIMER APELLIDO DEL TITULAR DE DERECHO" con el color de fondo #A6A6A6
    worksheet.merge_range('F14:F16', 'PRIMER APELLIDO DEL TITULAR DE DERECHO', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'text_wrap': True       # Ajustar el texto para que se ajuste dentro de la celda
                        }))
    # Combinar las celdas agregar el texto "SEGUNDO APELLIDO DEL TITULAR DE DERECHO" con el color de fondo #A6A6A6
    worksheet.merge_range('G14:G16', 'SEGUNDO APELLIDO DEL TITULAR DE DERECHO', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'text_wrap': True      # Ajustar texto dentro de la celda
                        }))
    # Combinar las celdas agregar el texto "Sexo" con el color de fondo #A6A6A6
    worksheet.merge_range('H14:H16', 'Sexo', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'rotation': 90          # Rotar el texto 90 grados hacia arriba
                        }))
    # Combinar las celdas agregar el texto "Grado Educativo" con el color de fondo #A6A6A6
    worksheet.merge_range('I14:I16', 'Grado Educativo', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'rotation': 90          # Rotar el texto 90 grados hacia arriba
                        }))

    # Combinar las celdas agregar el texto "Tipo de complemento" con el color de fondo #A6A6A6
    worksheet.merge_range('J14:J16', 'Tipo de complemento', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',   # Establecer la fuente como Arial
                            'font_size': 14,        # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo
                            'rotation': 90,          # Rotar el texto 90 grados hacia arriba
                            'text_wrap': True      # Ajustar texto dentro de la celda
                        }))


    # 1. Desactivar la cuadrícula
    worksheet.hide_gridlines(2)  # 2 es para ocultar la cuadrícula en la vista de diseño

    # 2. Rellenar toda la hoja con color blanco
    formato_blanco = workbook.add_format({'bg_color': 'white'})
    worksheet.set_column('A:Z', None, formato_blanco)  # Rellenar celdas de la A a la Z con color blanco (ajustar según el número de columnas)


    # 3. Ajustar el tamaño de las columnas de acuerdo con un ancho específico
    column_widths = {
        'A': 19,     # Ancho de la columna A
        'B': 19.91,  # Ancho de la columna B
        'C': 25,     # Ancho de la columna C
        'D': 35,     # Ancho de la columna D
        'E': 51.18,  # Ancho de la columna E
        'F': 45.55,  # Ancho de la columna F
        'G': 45.18,  # Ancho de la columna G
        'H': 6.18,   # Ancho de la columna H
        'I': 5.55,   # Ancho de la columna I
        'J': 13.55,  # Ancho de la columna J
        'K': 4.91,   # Ancho de la columna K
        'L': 4.91,   # Ancho de la columna L
        'M': 4.91,   # Ancho de la columna M
        'N': 4.91,   # Ancho de la columna N
        'O': 4.91,   # Ancho de la columna O
        'P': 4.91,   # Ancho de la columna P
        'Q': 4.91,   # Ancho de la columna Q
        'R': 4.91,   # Ancho de la columna R
        'S': 4.91,   # Ancho de la columna S
        'T': 4.91,   # Ancho de la columna T
        'U': 4.91,   # Ancho de la columna U
        'V': 4.91,   # Ancho de la columna V
        'W': 4.91,   # Ancho de la columna W
        'X': 4.91,   # Ancho de la columna X
        'Y': 4.91,   # Ancho de la columna Y
        'Z': 4.91,   # Ancho de la columna Z
        'AA': 4.91,  # Ancho de la columna AA
        'AB': 4.91,  # Ancho de la columna AB
        'AC': 4.91,  # Ancho de la columna AC
        'AD': 4.91,  # Ancho de la columna AD
        'AE': 4.91,  # Ancho de la columna AE
        'AF': 4.91,  # Ancho de la columna AF
        'AG': 4.91   # Ancho de la columna AG
    }

    # Asignar el ancho especificado a cada columna
    for col, width in column_widths.items():
        worksheet.set_column(f'{col}:{col}', width)  # Establecer el ancho para cada columna


    # Combinar las celdas K14:AI14 y agregar el texto con el color de fondo #A6A6A6
    worksheet.merge_range('K14:AI14', 'FECHA DE ENTREGA - Escriba el día hábil al cual corresponde la entrega del Complemento Alimentario', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',  # Establecer la fuente como Arial
                            'font_size': 14,       # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo gris
                            'text_wrap': True      # Ajustar texto dentro de la celda
                        }))

    # Combinar las celdas K16:AG16 y agregar el texto con el color de fondo #A6A6A6
    worksheet.merge_range('K16:AG16', 'Número de días de atención - Marque con una X el día que el Titular de Derecho recibe  el complemento alimentario', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',  # Establecer la fuente como Arial
                            'font_size': 14,       # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo gris
                            'text_wrap': True      # Ajustar texto dentro de la celda
                        }))

    # Combinar las celdas AH15:AI16 y agregar el texto con el color de fondo #A6A6A6
    worksheet.merge_range('AH15:AI16', 'Total días de consumo', 
                        workbook.add_format({
                            'bold': True,          # Negrita
                            'align': 'center',     # Alinear al centro
                            'valign': 'vcenter',   # Alinear verticalmente al centro
                            'font_name': 'Arial',  # Establecer la fuente como Arial
                            'font_size': 14,       # Tamaño de fuente 14
                            'bg_color': '#A6A6A6',  # Color de fondo gris
                            'text_wrap': True      # Ajustar texto dentro de la celda
                        }))

    # Ajustar la altura de las filas
    for fila in range(1, 8):  # Filas de 1 a 7
        worksheet.set_row(fila - 1, 12)  

    worksheet.set_row(7, 18)  # Fila 8 (índice 7 en Python)

    for fila in range(9, 14):  # Filas de 9 a 13
        worksheet.set_row(fila - 1, 30)

    worksheet.set_row(13, 42.8)  # Fila 14 (índice 13 en Python)
    worksheet.set_row(14, 68.3)  # Fila 15 (índice 14 en Python)

    for fila in range(16, 37):  # Filas de 16 a 36
        worksheet.set_row(fila - 1, 36.8)
    
    worksheet.set_row(36, 30)  # Fila 37 (índice 36 en Python)
    worksheet.set_row(37, 30)  # Fila 38 (índice 37 en Python)
    worksheet.set_row(38, 54)  # Fila 39 (índice 38 en Python)
    worksheet.set_row(39, 30)  # Fila 40 (índice 39 en Python)
    worksheet.set_row(40, 19.5)  # Fila 41 (índice 40 en Python)
    worksheet.set_row(41, 19.5)  # Fila 42 (índice 41 en Python)
    worksheet.set_row(42, 145)  # Fila 43 (índice 42 en Python)
    worksheet.set_row(43, 15)  # Fila 44 (índice 43 en Python)

    # Combinar celdas de dos en dos en la columna AH y AI, desde la fila 17 hasta la 36
    merge_format = workbook.add_format({
        'align': 'center',     # Alinear al centro
        'valign': 'vcenter',   # Alinear verticalmente al centro
        'text_wrap': True,     # Ajustar texto
    })

    for row in range(17, 37):  # Desde la fila 17 hasta la 36
        worksheet.merge_range(f'AH{row}:AI{row}', '', merge_format)

    # Lista de celdas combinadas para evitar sobrescribir su formato
    celdas_combinadas = [
        'A14:A16',  # Ejemplo de combinación
        'B14:B16',
        'C14:C16',
        'D14:D16',
        'E14:E16',
        'F14:F16',
        'G14:G16',
        'H14:H16',
        'I14:I16',
        'J14:J16',
        'K14:AI14',
        'K16:AG16',
        'AH15:AI16'
    ]

    # Definir formato con bordes
    border_format_combined = workbook.add_format({
        'border': 1,       # Borde en todas las direcciones
        'bold': True,      # Negrita (opcional, si ya lo usaste en otras celdas)
        'align': 'center',  # Alinear al centro
        'valign': 'vcenter' # Alinear verticalmente al centro
    })


    # Aplicar SOLO el formato con bordes a celdas ya combinadas
    for merge_range in celdas_combinadas:
        worksheet.conditional_format(merge_range, {'type': 'no_errors', 'format': border_format_combined})


    # Formato de borde
    border_format = workbook.add_format({'border': 1})

    # Recorremos las filas y columnas dentro del rango A14:AI36
    for row in range(14, 36):  # De fila 14 a 36
        for col in range(0, 35):  # De columna A (0) hasta AI (34)
            cell_ref = xl_rowcol_to_cell(row, col)  # Convertimos a referencia A1 (Ej: "B15")

            # Si la celda no está dentro de una combinación, aplicar el borde
            if not any(cell_ref in merge_range for merge_range in celdas_combinadas):
                worksheet.write_blank(row, col, None, border_format)

    # Reemplazar NaN por cadena vacía para evitar el error
    df_plantilla = df_plantilla.fillna('')
    # Separar el valor de 'TIPODOC' por ':' y tomar solo el primer valor
    df_plantilla["TIPODOC"] = df_plantilla["TIPODOC"].str.split(":").str[0]

    # Diccionario de mapeo
    # mapeo_tipo_complemento = {
    #     "RPS": "ALMUERZO",
    #     "JU": "ALMUERZO",
    #     "RI": "RACION INDUSTRIALIZADA",
    #     "CCT": "COMIDA CALIENTE TRANSPORTADORA"
    # }

    # Asignar el valor correspondiente usando map()
    # df_plantilla["TIPO_COMPLEMENTO"] = df_plantilla["TIPO DE RACIÓN"].map(mapeo_tipo_complemento)

    # Crear un formato con Arial 14 y bordes
    formato_arial_borde = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 14,
        'border': 1,  # Agregar borde en todas las direcciones
        'align': 'center',  # Alinear al centro
        'valign': 'vcenter' # Alinear verticalmente al centro
    })

    # Insertar valores de la columna 'NUMERO_REGISTRO' del DataFrame en la columna A, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['NUMERO_REGISTRO']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 0, doc_value, formato_arial_borde)  # Columna 'A' corresponde al índice 0

    # Insertar valores de la columna 'TIPODOC' del DataFrame en la columna B, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['TIPODOC']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 1, doc_value, formato_arial_borde)  # Columna 'B' corresponde al índice 1

    # Insertar valores de la columna 'DOC' del DataFrame en la columna C, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['DOC']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 2, doc_value, formato_arial_borde)  # Columna 'C' corresponde al índice 2
    
    # Insertar valores de la columna 'NOMBRE1' del DataFrame en la columna D, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['NOMBRE1']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 3, doc_value, formato_arial_borde)  # Columna 'D' corresponde al índice 3
    
    # Insertar valores de la columna 'NOMBRE2' del DataFrame en la columna E, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['NOMBRE2']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 4, doc_value, formato_arial_borde)  # Columna 'E' corresponde al índice 4
    
    # Insertar valores de la columna 'APELLIDO1' del DataFrame en la columna F, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['APELLIDO1']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 5, doc_value, formato_arial_borde)  # Columna 'F' corresponde al índice 5
    
    # Insertar valores de la columna 'APELLIDO2' del DataFrame en la columna G, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['APELLIDO2']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 6, doc_value, formato_arial_borde)  # Columna 'G' corresponde al índice 6
    
    # Insertar valores de la columna 'GENERO' del DataFrame en la columna H, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['GENERO']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 7, doc_value, formato_arial_borde)  # Columna 'H' corresponde al índice 7
    
    # Insertar valores de la columna 'GRADO_COD' del DataFrame en la columna I, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['GRADO_COD']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 8, doc_value, formato_arial_borde)  # Columna 'I' corresponde al índice 8
    
    # Insertar valores de la columna 'TIPO DE RACIÓN' del DataFrame en la columna J, de la fila 17 a la 36
    for idx, doc_value in enumerate(df_plantilla['TIPO DE RACIÓN']):  
        row = 17 + idx  # Iniciar en la fila 17
        worksheet.write(row - 1, 9, doc_value, formato_arial_borde)  # Columna 'J' corresponde al índice 9

    # ===========================================================
    # Codigo para el formato inferior de la planilla
    # ===========================================================
    # Combinar celdas A38 y C38 y agregar el texto en negrita
    worksheet.merge_range('A38:C38', 'NOMBRE DEL RESPONSABLE DEL OPERADOR', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_name': 'Arial',
                            'font_size': 12
                        }))

    # Combinar celdas G38 y P38 y agregar el texto en negrita
    worksheet.merge_range('G38:P38', 'NOMBRE RECTOR ESTABLECIMIENTO EDUCATIVO', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_name': 'Arial',
                            'font_size': 12
                        }))

    # Combinar celdas A40 y C40 y agregar el texto en negrita
    worksheet.merge_range('A40:C40', 'FIRMA DEL RESPONSABLE DEL OPERADOR', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_name': 'Arial',
                            'font_size': 12
                        }))

    # Combinar celdas G40 y P40 y agregar el texto en negrita
    worksheet.merge_range('G40:P40', 'FIRMA DEL RESPONSABLE DEL OPERADOR', 
                        workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_name': 'Arial',
                            'font_size': 12
                        }))

    # Formato para borde inferior simple
    border_bottom_format = workbook.add_format({'bottom': 1})

    # Lista de rangos a los que se aplicará el borde inferior (excepto fila 41)
    rangos_borde_inferior = [
        (38, 3, 5),   # D38:F38 (fila 38, columnas 3 a 5)
        (38, 16, 33), # Q38:AH38 (fila 38, columnas 16 a 33)
        (40, 3, 5),   # D40:F40 (fila 40, columnas 3 a 5)
        (40, 16, 33)  # Q40:AH40 (fila 40, columnas 16 a 33)
    ]

    # Aplicar formato de borde inferior a las filas 38 y 40
    for fila, col_inicio, col_fin in rangos_borde_inferior:
        for col in range(col_inicio, col_fin + 1):
            worksheet.write_blank(fila - 1, col, None, border_bottom_format)

    # Formato para borde inferior grueso (solo A41:AI41)
    formato_borde_inferior = workbook.add_format({'bottom': 2})

    # Aplicar borde inferior grueso solo de A41 a AI41
    for col in range(0, 35):  # A (0) hasta AI (33)
        worksheet.write_blank(40, col, None, formato_borde_inferior)  # Fila 41 (índice 40 en Python)

    # Formato para la celda combinada en la fila 43
    formato_fila_43 = workbook.add_format({
        'bold': True,
        'align': 'left',      # Alineación a la izquierda
        'valign': 'top',      # Alineación en la parte superior
        'text_wrap': True,    # Ajuste de texto en la celda
        'font_name': 'Arial',
        'font_size': 12,
        'border': 1           # Bordes en todas las direcciones
    })

    # Combinar celdas A43:AI43 y aplicar formato
    worksheet.merge_range('A43:AI43', 'Observaciones:', formato_fila_43)

    # Formato para borde inferior grueso
    formato_borde_inferior_44 = workbook.add_format({'bottom': 2})

    # Aplicar borde solo en A44:AI44, evitando la intersección con la celda combinada en A43:AI43
    for col in range(0, 35):  # Columnas A (0) hasta AI (34)
        worksheet.write_blank(43, col, None, formato_borde_inferior_44)  # Fila 44 (índice 43 en Python)

    # Formato para borde izquierdo grueso
    formato_borde_izquierdo_AJ = workbook.add_format({'left': 2})  

    # Aplicar borde en la columna AJ desde la fila 1 hasta la 44
    for fila in range(0, 44):  # Filas 1 (0 en Python) hasta 44 (43 en Python)
        worksheet.write_blank(fila, 35, None, formato_borde_izquierdo_AJ)  # AJ es la columna 35 (A=0, B=1, ..., AJ=35)

    # Guardar el archivo con las modificaciones
    writer.close()

def separar_dataframes():
    # Ruta archivo excel
    ruta_archivo = "Insumo\\Focalizacion.xlsx"

    # Leer el archivo Excel desde la celda A8 hasta AF27, tomando la fila 8 como encabezado
    df_focalizacion = pd.read_excel(ruta_archivo)

    # Crear un diccionario para almacenar los DataFrames separados
    dfs_separados = {}

    # Agrupar por 'SEDE', 'JORNADA' y 'GRADO_COD'
    for (sede, jornada, grado), df_grupo in df_focalizacion.groupby(['SEDE', 'JORNADA', 'GRADO_COD']):
        # Agregar una columna de numeración secuencial para cada grupo
        df_grupo = df_grupo.copy()  # Evitar advertencias de SettingWithCopyWarning
        df_grupo['NUMERO_REGISTRO'] = range(1, len(df_grupo) + 1)

        # Dividir el grupo en bloques de 20 registros
        num_partes = (len(df_grupo) // 20) + (1 if len(df_grupo) % 20 > 0 else 0)

        for i in range(num_partes):
            # Extraer hasta 20 registros por fragmento
            df_parte = df_grupo.iloc[i * 20:(i + 1) * 20].copy()

            # Crear un nombre concatenado con la parte correspondiente
            # nombre_df = f"{sede}_{jornada}_{grado}_parte_{i+1}"  # Se empieza en 1 en vez de 0
            nombre_df = f"{sede}_{jornada}_{grado}"  # Se empieza en 1 en vez de 0
            
            # Guardarlo en el diccionario
            dfs_separados[nombre_df] = df_parte

    cantidad_dfs = len(dfs_separados)
    print(f"Cantidad de DataFrames generados: {cantidad_dfs}\n")

    # Iterar sobre los DataFrames generados y pasarlos a la función
    for i, (nombre_df, df) in enumerate(dfs_separados.items(), start=1):
        print(f"Iteración {i}: Generando plantilla de control para: {nombre_df}")
        crear_plantilla_control(df, nombre_df)

def limpiar_carpeta_resultado():
    carpeta_resultado = "Resultado"

    # Eliminar y recrear la carpeta
    if os.path.exists(carpeta_resultado):
        shutil.rmtree(carpeta_resultado)  # Borra todo el contenido de una vez
    os.makedirs(carpeta_resultado)  # Crea la carpeta vacía nuevamente

    print("Carpeta 'Resultado' vaciada y recreada.")

def main():
    # Iniciar proceso generacion de planillas
    print("="*60 + "\n  SE INICIA PROCESO DE GENERACION DE PLANILLAS DE CONTROL\n" + "="*60)

    # Registrar el tiempo de inicio
    inicio = time.time()

    # Ejecutar la función
    separar_dataframes()

    # Registrar el tiempo de fin
    fin = time.time()

    # Calcular el tiempo total de ejecución
    duracion = fin - inicio  # Diferencia en segundos
    horas, resto = divmod(duracion, 3600)  # Convertir a horas
    minutos, segundos = divmod(resto, 60)  # Convertir a minutos y segundos

    print(f"\nTiempo total de ejecución: {int(horas):02}:{int(minutos):02}:{int(segundos):02}")
    print("Proceso terminado\n")



if __name__ == "__main__":
    # Llamar a la función para ejecutar el proceso
    main()

