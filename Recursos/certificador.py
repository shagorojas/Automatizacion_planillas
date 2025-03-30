
# Importamos las librerías necesarias
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage  # Renombramos Image para evitar conflicto con PIL
from PIL import Image as PILImage  # Renombramos también Image de PIL
import pandas as pd
import xlsxwriter
import pathlib
import json 
import os

class MunicipioNoSoportadoError(Exception):
    """Excepción personalizada para manejar municipios no soportados."""
    pass

class GeneradorCertificaciones:
    def __init__(self):
        # Definir rutas
        self.ruta_master = os.path.join(str(os.path.abspath(pathlib.Path().absolute())))
        self.ruta_json = os.path.join(self.ruta_master, "Config", "Config.json")
        self.ruta_resultado = os.path.join(self.ruta_master, "Resultado")
        self.ruta_resultado_pdf = os.path.join(self.ruta_master, "Resultado pdf")
        self.ruta_resultado_combinado = os.path.join(self.ruta_master, "Resultado excel")
        self.ruta_certificaciones = os.path.join(self.ruta_master, "Resultado certificaciones")

        # Ruta insumos
        self.ruta_archivo_focalizacion = "Insumo\\Focalizacion.xlsx"
        self.ruta_archivo_aplicacion_novedades = "Insumo\\Focalizacion_actualizada.xlsx"
        self.ruta_archivo_novedades = "Insumo\\Novedades.xlsx"
        self.ruta_parametros = "Insumo\\Parametros.xlsx"

    def leer_json(self):

        # Leemos el archivo JSON que contiene la configuración del proceso
        with open(self.ruta_json) as contenido:
            self.config = json.load(contenido)  # Almacena la configuración en el atributo config
        return self.config

    def generar_certificacion(self, var_institucion, var_dane_institucion):

        # ==================================================
        # Cargar parametros de acuerdo al municipio
        # ==================================================
        # Leer paramtro JSON
        params = self.leer_json()

        if params["municipio_proceso"] == "FUNZA":
            # Leer el archivo de parametros
            df_parametros = pd.read_excel(self.ruta_parametros, sheet_name="FUNZA")

            # Rutas imagenes
            logo_alimentos = os.path.join(self.ruta_master, "util", "Logo alimentos.png")
            logo_operador = os.path.join(self.ruta_master, "util", "Logo operador FUNZA.png")
            logo_secretaria = os.path.join(self.ruta_master, "util", "Logo secretaria funza.png")
            logo_min_educacion = os.path.join(self.ruta_master, "util", "Logo Min Educacion.png")
        else:
            raise MunicipioNoSoportadoError("\nMunicipio no soportado. Deteniendo ejecución...\n")

        # Convertir a diccionario
        dict_data = dict(zip(df_parametros["Concepto"], df_parametros["Valor"]))


        # Cargamos los parametros por variables
        departamento = dict_data["Departamento"]
        municipio = dict_data["Municipio"]
        operador = dict_data["Operador"]
        contrato = dict_data["Contrato No."]
        codigo_dane = dict_data["Codigo dane"]
        codigo_dane_completo = dict_data["Codigo dane completo"]
        mes_atencion = dict_data["Mes de atencion"]
        anio = dict_data["Año"]

        rector_institucion = dict_data[var_institucion]

        # ==================================================
        # ==================================================

        # Crear la carpeta si no existe
        if not os.path.exists(self.ruta_certificaciones):
            os.makedirs(self.ruta_certificaciones)

        # Definir el nombre del archivo basado en var_institucion
        nombre_archivo = f"{var_institucion}.xlsx"
        archivo_excel = os.path.join(self.ruta_certificaciones, nombre_archivo)

        # Crear un DataFrame vacío
        df = pd.DataFrame(index=range(10), columns=[chr(58 + i) for i in range(8)])

        # Guardar el DataFrame en un archivo Excel
        df.to_excel(archivo_excel, index=False, engine='xlsxwriter')

        # Crear una conexión con el archivo Excel y agregar las imágenes
        writer = pd.ExcelWriter(archivo_excel, engine='xlsxwriter')
        df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')

        # Acceder al objeto workbook y worksheet
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Insertar la primera imagen en A3
        worksheet.insert_image('A3', logo_alimentos, {'x_scale': 0.3, 'y_scale': 0.3, 'x_offset': 8, 'y_offset': 0})

        # Insertar la segunda imagen en A3, desplazándola un poco a la derecha
        worksheet.insert_image('A3', logo_min_educacion, {'x_scale': 0.25, 'y_scale': 0.25, 'x_offset': 115, 'y_offset': 8})

        # Insertar las imágenes
        worksheet.insert_image('C3', logo_operador, {'x_scale': 0.4, 'y_scale': 0.4})
        worksheet.insert_image('B3', logo_secretaria, {'x_scale': 0.4, 'y_scale': 0.4})

        # Combinar celdas A2 a C5
        worksheet.merge_range('A2:C5', '')

        # Crear un solo formato reutilizable
        formato_celda_unicos = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 12,
            'border': 1
        })

        # Crear un solo formato reutilizable
        formato_celda_variables = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 11,
            'font_color': '#808080',  # Color de la letra en gris
            'border': 1
        })

        # Crear un solo formato reutilizable
        formato_celda_variables_negra = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 12,
            'border': 1
        })

        # Crear un solo formato reutilizable
        formato_celda_unicos_simple = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 12,
            'border': 1
        })

        # Combinar celdas D2 a H5 y agregar el texto en negrita
        worksheet.merge_range('D2:H5', 'CERTIFICADO DE ENTREGA DE RACIONES A INSTITUCIONES EDUCATIVAS:', 
                                workbook.add_format({
                                    'bold': True,
                                    'align': 'center',
                                    'valign': 'vcenter',
                                    'font_name': 'Aptos Narrow',   # Establecer la fuente como Aptos Narrow
                                    'font_size': 16
                                }))

        # Combinar celdas de A7 a H7
        worksheet.merge_range('A7:H7', 'DATOS GENERALES', 
                                workbook.add_format({
                                'bold': True,
                                'align': 'center',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',   # Establecer la fuente como Aptos Narrow
                                'font_size': 12,        # Tamaño de fuente
                                'bg_color': '#BFBFBF'   # Color de fondo
                            }))

        # Aplicar el formato a las celdas
        worksheet.write('A8', 'OPERADOR', formato_celda_unicos)
        worksheet.write('F8', 'CONTRATO N°:', formato_celda_unicos)
        worksheet.write('A10', 'INSTITUCIÓN O CENTRO EDUCATIVO', formato_celda_unicos)
        worksheet.write('F10', 'CÓDIGO DANE', formato_celda_unicos)
        worksheet.write('A11', 'DEPARTAMENTO:', formato_celda_unicos)
        worksheet.write('F11', 'CÓDIGO DANE', formato_celda_unicos)
        worksheet.write('A12', 'MUNICIPIO', formato_celda_unicos)
        worksheet.write('F12', 'CÓDIGO DANE', formato_celda_unicos)
        worksheet.write('A13', 'FECHA DE EJECUCIÓN', formato_celda_unicos)
        worksheet.write('B13', 'Desde', formato_celda_unicos)
        worksheet.write('E13', 'Hasta', formato_celda_unicos)
        worksheet.write('A14', 'NOMBRE RECTOR:', formato_celda_unicos)

        # Combinar las celdas y escribir el texto
        worksheet.merge_range('B14:H14', rector_institucion, 
                            workbook.add_format({
                                'align': 'center',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',
                                'font_size': 12,
                                'text_wrap': True,
                                'border': 1
                            }))

        # Aplicar el formato a las celdas
        worksheet.merge_range('B8:E8', operador, formato_celda_variables_negra)
        worksheet.merge_range('G8:H8', contrato, formato_celda_variables_negra)
        worksheet.merge_range('B10:E10', var_institucion, formato_celda_variables)
        worksheet.merge_range('G10:H10',  var_dane_institucion, formato_celda_variables)
        worksheet.merge_range('G11:H11', codigo_dane, formato_celda_variables)
        worksheet.merge_range('B11:E11', departamento, formato_celda_variables)
        worksheet.merge_range('G12:H12', codigo_dane_completo, formato_celda_variables)
        worksheet.merge_range('B12:E12', municipio, formato_celda_variables)

        #############################################################################
        # Logica para ingreso de variables 
        #############################################################################

        # Combinar celdas de A17 a H17
        worksheet.merge_range('A17:H17', 'CERTIFICACIÓN', 
                                workbook.add_format({
                                'bold': True,
                                'align': 'center',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',   # Establecer la fuente como Aptos Narrow
                                'font_size': 12,        # Tamaño de fuente
                                'bg_color': '#BFBFBF'   # Color de fondo
                            }))

        # Combinar celdas de A18 a H20
        worksheet.merge_range('A18:H20', 'El suscrito Rector de la Institución Educativa citada en el encabezado, certifica que se entregaron las siguientes raciones, en las fechas señaladas y de acuerdo con la siguiente distribución:', 
                                workbook.add_format({
                                'align': 'left',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',   # Establecer la fuente como Aptos Narrow
                                'font_size': 12,         # Tamaño de fuente
                                'border': 1
                            }))

        # Definir formato para las celdas combinadas
        merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 11,
            'bg_color': '#BFBFBF',
            'text_wrap': True
        })

        # Definir los valores y los rangos a combinar
        merge_ranges = {
            'A23:A24': 'NOMBRE DEL ESTABLECIMIENTO EDUCATIVO O CENTRO EDUCATIVO',
            'B23:B24': 'TIPO RACIÓN',
            'F24:H24': 'NOVEDADES'
        }

        # Aplicar la combinación de celdas con el formato
        for rango, texto in merge_ranges.items():
            worksheet.merge_range(rango, texto, merge_format)

        # Combinar celdas de C23 a H23
        worksheet.merge_range('C23:H23', 'ENTREGADO', 
                                workbook.add_format({
                                'align': 'center',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',   # Establecer la fuente como Aptos Narrow
                                'font_size': 12,         # Tamaño de fuente
                                'bg_color': '#BFBFBF'    # Color de fondo
                            }))

        # Crear un solo formato reutilizable
        formato_celda_gris = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 11,
            'bg_color': '#BFBFBF',    # Color de fondo
            'border': 1
        })

        worksheet.write('C24', 'N° RACIONES POR DÍA', formato_celda_gris)
        worksheet.write('D24', 'N° DÍAS ATENDIDOS', formato_celda_gris)
        worksheet.write('E24', 'TOTAL RACIONES', formato_celda_gris)

        # =========================================================
        # Logica para cantidad de sedes por institucion
        # =========================================================

        # Cargar el archivo de Excel
        df_focalizacion = pd.read_excel(self.ruta_archivo_aplicacion_novedades)

        df_agrupado = df_focalizacion.groupby(["INSTITUCION", "SEDE"]).size().reset_index(name="TOTAL_REGISTROS")

        # Filtrar por institución
        df_filtrado = df_agrupado[df_agrupado["INSTITUCION"] == var_institucion]

        # Definir la fila inicial
        fila_inicio = 25  
        salto_filas = 3  # Cantidad de filas a combinar en cada iteración

        # Iterar sobre cada fila del DataFrame filtrado
        for _, row in df_filtrado.iterrows():
            texto_sede = row["SEDE"]  # Obtener el valor de la columna SEDE
            
            # Definir el rango de celdas a combinar dinámicamente
            fila_fin = fila_inicio + (salto_filas - 1)  # Determinar la fila final
            rango_celdas = f'A{fila_inicio}:A{fila_fin}'  # Construir el rango dinámico
            
            # Combinar las celdas y escribir el texto
            worksheet.merge_range(rango_celdas, texto_sede, 
                                workbook.add_format({
                                    'align': 'center',
                                    'valign': 'vcenter',
                                    'font_name': 'Aptos Narrow',
                                    'font_size': 12,
                                    'text_wrap': True,
                                    'border': 1
                                }))

            # Escribir valores en la columna B
            worksheet.write(f'B{fila_inicio}', 'RPS', formato_celda_unicos_simple)
            worksheet.write(f'B{fila_inicio + 1}', 'RI', formato_celda_unicos_simple)
            worksheet.write(f'B{fila_inicio + 2}', 'CCT', formato_celda_unicos_simple)

            # Actualizar la fila de inicio para la siguiente iteración
            # fila_inicio = fila_fin + 1 

            # =========================================================
            # Logica total raciones
            # =========================================================

            # Filtrar el DataFrame según la INSTITUCION y SEDE
            df_filtrado = df_focalizacion[
                (df_focalizacion["INSTITUCION"] == var_institucion) & 
                (df_focalizacion["SEDE"] == texto_sede)
            ].copy()  # Se usa `.copy()` para evitar modificar el original

            # Verificar si la columna FECHA_NACIMIENTO existe en el DataFrame filtrado
            if "FECHA_NACIMIENTO" in df_filtrado.columns:
                idx_fecha_nacimiento = df_filtrado.columns.get_loc("FECHA_NACIMIENTO")
                
                # Obtener las columnas que vienen después de FECHA_NACIMIENTO
                columnas_despues = df_filtrado.columns[idx_fecha_nacimiento + 1:]
                
                # Reemplazar valores no "X" con 0 y las "X" con 1 en un nuevo DataFrame para no modificar el original
                df_temp = df_filtrado[columnas_despues].applymap(lambda x: 1 if x == "X" else 0)

                # Agregar al DataFrame original las sumas de "X"
                df_filtrado["TOTAL_RACIONES"] = df_temp.sum(axis=1)

                # Agrupar por TIPO DE RACIÓN y sumar TOTAL_RACIONES
                df_resultado = df_filtrado.groupby("TIPO DE RACIÓN", as_index=False)["TOTAL_RACIONES"].sum()

                # Leer insumo novedades
                df_novedades = pd.read_excel(self.ruta_archivo_novedades, sheet_name="Novedades")

                df_novedades_filtro = df_novedades[
                    (df_novedades["SEDE"] == texto_sede) &
                    (df_novedades["TIPO_NOVEDAD"] == "Aumento raciones")
                ]

                # Verifica si el DataFrame no está vacío
                if not df_novedades_filtro.empty: 
                    # Crear una copia explícita para evitar el warning
                    df_novedades_filtro = df_novedades_filtro.copy()

                    # Reemplazar valores nulos en DETALLE
                    df_novedades_filtro["DETALLE"] = df_novedades_filtro["DETALLE"].fillna("")

                    # Filtrar solo las filas que contienen "-"
                    df_novedades_filtro = df_novedades_filtro[df_novedades_filtro["DETALLE"].str.contains("-")]

                    # Separar DETALLE en CANTIDAD y RACION (máximo 2 partes)
                    df_novedades_filtro[["CANTIDAD", "RACION"]] = df_novedades_filtro["DETALLE"].str.split("-", n=1, expand=True)

                    # Convertir CANTIDAD a número
                    df_novedades_filtro["CANTIDAD"] = pd.to_numeric(df_novedades_filtro["CANTIDAD"], errors="coerce")

                    # Obtener el valor maximo para cada tipo de racion
                    df_novedades_filtro = df_novedades_filtro.groupby("RACION", as_index=False)["CANTIDAD"].max()

                    # Renombrar columnas en df_novedades_filtro
                    df_novedades_filtro = df_novedades_filtro.rename(columns={"RACION": "TIPO DE RACIÓN", "CANTIDAD": "TOTAL_RACIONES"})

                    # Concatenar ambos DataFrames
                    df_concatenado = pd.concat([df_resultado, df_novedades_filtro], ignore_index=True)

                    # Agrupar por "TIPO DE RACIÓN" y sumar "TOTAL_RACIONES"
                    df_resultado_final = df_concatenado.groupby("TIPO DE RACIÓN", as_index=False)["TOTAL_RACIONES"].sum()

                    # Agrupar por "TIPO DE RACIÓN" y sumar "TOTAL_RACIONES"
                    df_resultado_final = df_resultado_final.groupby("TIPO DE RACIÓN", as_index=False)["TOTAL_RACIONES"].sum()  
                else:
                    df_resultado_final = df_resultado

                # Definir las filas donde se deben escribir los valores
                filas_racion = {"RPS": fila_inicio, "RI": fila_inicio + 1, "CCT": fila_inicio + 2}

                # Escribir los valores en la hoja de Excel
                for tipo_racion, fila in filas_racion.items():
                    total_raciones = df_resultado_final.loc[df_resultado_final["TIPO DE RACIÓN"] == tipo_racion, "TOTAL_RACIONES"]
                    
                    if not total_raciones.empty and total_raciones.values[0] > 0:
                        worksheet.write(f'E{fila}', total_raciones.values[0], formato_celda_unicos_simple)  

            # =========================================================
            # Logica raciones maximas por dia
            # =========================================================
            # Filtrar el DataFrame según la INSTITUCION y SEDE
            df_filtrado = df_focalizacion[
                (df_focalizacion["INSTITUCION"] == var_institucion) & 
                (df_focalizacion["SEDE"] == texto_sede)
            ].copy()

            # Verificar si la columna FECHA_NACIMIENTO existe
            if "FECHA_NACIMIENTO" in df_filtrado.columns:
                idx_fecha_nacimiento = df_filtrado.columns.get_loc("FECHA_NACIMIENTO")
                
                # Obtener las columnas que vienen después de FECHA_NACIMIENTO
                columnas_despues = df_filtrado.columns[idx_fecha_nacimiento + 1:]

                # Convertir "X" en 1 y el resto en 0
                df_filtrado[columnas_despues] = df_filtrado[columnas_despues].applymap(lambda x: 1 if x == "X" else 0)

                # Agrupar por SEDE y TIPO DE RACIÓN, sumando cada columna de columnas_despues
                df_agrupado = df_filtrado.groupby(["SEDE", "TIPO DE RACIÓN"])[columnas_despues].sum()

                # Obtener el valor máximo por fila
                df_agrupado["MAXIMO_RACIONES"] = df_agrupado.max(axis=1)

                # Resetear el índice y seleccionar solo las columnas necesarias
                df_resultado = df_agrupado.reset_index()[["TIPO DE RACIÓN", "MAXIMO_RACIONES"]]

                # Leer insumo novedades
                df_novedades = pd.read_excel(self.ruta_archivo_novedades, sheet_name="Novedades")

                df_novedades_filtro = df_novedades[
                    (df_novedades["SEDE"] == texto_sede) &
                    (df_novedades["TIPO_NOVEDAD"] == "Aumento raciones")
                ]

                # Verifica si el DataFrame no está vacío
                if not df_novedades_filtro.empty: 
                    # Crear una copia explícita para evitar el warning
                    df_novedades_filtro = df_novedades_filtro.copy()

                    # Reemplazar valores nulos en DETALLE
                    df_novedades_filtro["DETALLE"] = df_novedades_filtro["DETALLE"].fillna("")

                    # Filtrar solo las filas que contienen "-"
                    df_novedades_filtro = df_novedades_filtro[df_novedades_filtro["DETALLE"].str.contains("-")]

                    # Separar DETALLE en CANTIDAD y RACION (máximo 2 partes)
                    df_novedades_filtro[["CANTIDAD", "RACION"]] = df_novedades_filtro["DETALLE"].str.split("-", n=1, expand=True)

                    # Convertir CANTIDAD a número
                    df_novedades_filtro["CANTIDAD"] = pd.to_numeric(df_novedades_filtro["CANTIDAD"], errors="coerce")

                    # Obtener el valor maximo para cada tipo de racion
                    df_novedades_filtro = df_novedades_filtro.groupby("RACION", as_index=False)["CANTIDAD"].max()

                    # Renombrar columnas en df_novedades_filtro
                    df_novedades_filtro = df_novedades_filtro.rename(columns={"RACION": "TIPO DE RACIÓN", "CANTIDAD": "MAXIMO_RACIONES"})

                    # Concatenar ambos DataFrames
                    df_concatenado = pd.concat([df_resultado, df_novedades_filtro], ignore_index=True)

                    # Agrupar por "TIPO DE RACIÓN" y sumar "MAXIMO_RACIONES"
                    df_resultado_final = df_concatenado.groupby("TIPO DE RACIÓN", as_index=False)["MAXIMO_RACIONES"].sum()

                    # Agrupar por "TIPO DE RACIÓN" y sumar "MAXIMO_RACIONES"
                    df_resultado_final = df_resultado_final.groupby("TIPO DE RACIÓN", as_index=False)["MAXIMO_RACIONES"].sum()  
                else:
                    df_resultado_final = df_resultado
            
                # Definir las filas donde se deben escribir los valores
                filas_racion = {"RPS": fila_inicio, "RI": fila_inicio + 1, "CCT": fila_inicio + 2}

                # Escribir los valores en la hoja de Excel
                for tipo_racion, fila in filas_racion.items():
                    maximo_raciones = df_resultado_final.loc[df_resultado_final["TIPO DE RACIÓN"] == tipo_racion, "MAXIMO_RACIONES"]
                    
                    if not maximo_raciones.empty and maximo_raciones.values[0] > 0:
                        worksheet.write(f'C{fila}', int(maximo_raciones.values[0]), formato_celda_unicos_simple)  # Convertir a entero antes de escribir
            
            # =========================================================
            # Logica dias atentidos
            # =========================================================
            # Filtrar el DataFrame según la INSTITUCION y SEDE
            df_filtrado = df_focalizacion[
                (df_focalizacion["INSTITUCION"] == var_institucion) & 
                (df_focalizacion["SEDE"] == texto_sede)
            ].copy()

            # Verificar si la columna FECHA_NACIMIENTO existe
            if "FECHA_NACIMIENTO" in df_filtrado.columns:
                idx_fecha_nacimiento = df_filtrado.columns.get_loc("FECHA_NACIMIENTO")
                
                # Obtener las columnas que vienen después de FECHA_NACIMIENTO
                columnas_despues = df_filtrado.columns[idx_fecha_nacimiento + 1:]

                # Agrupar por "TIPO DE RACIÓN" y contar cuántas columnas tienen al menos un valor en cada grupo
                df_dias_racion = df_filtrado.groupby("TIPO DE RACIÓN")[columnas_despues].apply(lambda x: x.notna().any().sum()).reset_index()

                # Renombrar la columna de conteo
                df_dias_racion.rename(columns={0: "DIAS_RACION"}, inplace=True)

                # Leer insumo novedades
                df_novedades = pd.read_excel(self.ruta_archivo_novedades, sheet_name="Novedades")

                df_novedades_filtro = df_novedades[
                    (df_novedades["SEDE"] == texto_sede) &
                    (df_novedades["TIPO_NOVEDAD"] == "Aumento raciones")
                ]

                # Verifica si el DataFrame no está vacío
                if not df_novedades_filtro.empty: 
                    # Crear una copia explícita para evitar el warning
                    df_novedades_filtro = df_novedades_filtro.copy()

                    # Reemplazar valores nulos en DETALLE
                    df_novedades_filtro["DETALLE"] = df_novedades_filtro["DETALLE"].fillna("")

                    # Filtrar solo las filas que contienen "-"
                    df_novedades_filtro = df_novedades_filtro[df_novedades_filtro["DETALLE"].str.contains("-")]

                    # Separar DETALLE en CANTIDAD y RACION (máximo 2 partes)
                    df_novedades_filtro[["CANTIDAD", "RACION"]] = df_novedades_filtro["DETALLE"].str.split("-", n=1, expand=True)

                    # Convertir CANTIDAD a número
                    df_novedades_filtro["CANTIDAD"] = pd.to_numeric(df_novedades_filtro["CANTIDAD"], errors="coerce")

                    # Contar los días únicos por cada RACION
                    df_conteo_dias_racion = df_novedades_filtro.groupby("RACION")["FECHA"].nunique().reset_index()

                    # Renombrar la columna
                    df_conteo_dias_racion.rename(columns={"FECHA": "DIAS_RACION","RACION": "TIPO DE RACIÓN",}, inplace=True)

                    # Concatenar ambos DataFrames
                    df_concatenado = pd.concat([df_dias_racion, df_conteo_dias_racion], ignore_index=True)

                    # Agrupar por "TIPO DE RACIÓN" y obtener el valor máximo de "DIAS_RACION"
                    df_resultado_final = df_concatenado.groupby("TIPO DE RACIÓN", as_index=False)["DIAS_RACION"].max()
                else:
                    df_resultado_final = df_dias_racion
            
                # Definir las filas donde se deben escribir los valores
                filas_racion = {"RPS": fila_inicio, "RI": fila_inicio + 1, "CCT": fila_inicio + 2}

                # Escribir los valores en la hoja de Excel
                for tipo_racion, fila in filas_racion.items():
                    maximo_raciones = df_resultado_final.loc[df_resultado_final["TIPO DE RACIÓN"] == tipo_racion, "DIAS_RACION"]
                    
                    if not maximo_raciones.empty and maximo_raciones.values[0] > 0:
                        worksheet.write(f'D{fila}', int(maximo_raciones.values[0]), formato_celda_unicos_simple)  # Convertir a entero antes de escribir

            # Actualizar la fila de inicio para la siguiente iteración
            fila_inicio = fila_fin + 1 

        # =========================================================
        # Definir formato con bordes
        border_format = workbook.add_format({
            'border': 1,       # Borde en todas las direcciones
            'align': 'center',  # Alinear al centro
            'valign': 'vcenter' # Alinear verticalmente al centro
        })

        # Aplicar el formato a la región B25:E{fila_fin}
        worksheet.conditional_format(f'B25:E{fila_fin}', {'type': 'no_errors', 'format': border_format})
        # =========================================================
        # Formula de SUMA para la columna
        # =========================================================

        # Crear un solo formato reutilizable
        formato_celda_unicos_simple = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 12,
            'border': 1
        })

        worksheet.write(f'C{fila_inicio}', f'=SUM(C25:C{fila_inicio - 1})', formato_celda_unicos_simple)
        worksheet.write(f'D{fila_inicio}', f'=SUM(D25:D{fila_inicio - 1})', formato_celda_unicos_simple)
        worksheet.write(f'E{fila_inicio}', f'=SUM(E25:E{fila_inicio - 1})', formato_celda_unicos_simple)

        # Construir el rango dinámico
        rango_celdas = f'A{fila_inicio}:B{fila_inicio}' 

        # Combinar las celdas y escribir el texto
        worksheet.merge_range(rango_celdas, "TOTAL", 
                            workbook.add_format({
                                'bold': True,
                                'align': 'left',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',
                                'font_size': 12,
                                'text_wrap': True,
                                'border': 1
                            }))
        
        # =========================================================
        # Construir el rango dinámico
        rango_celdas = f'A{fila_inicio + 1}:H{fila_inicio + 1}' 

        # Combinar las celdas y escribir el texto
        worksheet.merge_range(rango_celdas, "RPS = Ración Preparada en Sitio\nRI: Ración Industrializada\nCCT: Comida Caliente Transporta", 
                            workbook.add_format({
                                'align': 'left',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',
                                'font_size': 10,
                                'text_wrap': True,
                                'border': 1
                            }))
        
        # Altura de la fila
        worksheet.set_row(fila_inicio, 45)  # Fila 15 (índice 14 en Python)

        # Definir formato para las celdas individuales
        cell_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 11,
            'bg_color': '#BFBFBF',
            'text_wrap': True,
            'border': 1
        })

        # Definir las celdas y sus respectivos textos
        celdas_textos = {
            f'A{fila_inicio + 3}': 'DESCRPCIÓN',
            f'B{fila_inicio + 3}': 'TOTAL RACIONES ENTREGADAS RACIÓN PREPARADA EN SITIO',
            f'C{fila_inicio + 3}': 'TOTAL RACIONES ENTREGADAS RACIÓN INDUSTRIALIZADA',
            f'D{fila_inicio + 3}': 'TOTAL RACIONES ENTREGADAS COMIDA CALIENTE TRANSPORTADA',
            f'E{fila_inicio + 3}': 'No. DE TITULARES DE DERECHO',
        }

        # Aplicar formato y texto a cada celda individualmente
        for celda, texto in celdas_textos.items():
            worksheet.write(celda, texto, cell_format)

        # =========================================================

        # Crear un solo formato reutilizable
        formato_celda_unicos_simple = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 12,
            'border': 1
        })

        worksheet.write(f'A{fila_inicio + 4}', 'POBLACIÓN MAYORITARIA', formato_celda_unicos)

        worksheet.write(f'B{fila_inicio + 4}', f'=SUMIF(B25:B{fila_inicio - 1},"RPS",E25:E{fila_inicio - 1})', formato_celda_unicos_simple)
        worksheet.write(f'C{fila_inicio + 4}', f'=SUMIF(B25:B{fila_inicio - 1},"RI",E25:E{fila_inicio - 1})', formato_celda_unicos_simple)
        worksheet.write(f'D{fila_inicio + 4}', f'=SUMIF(B25:B{fila_inicio - 1},"CCT",E25:E{fila_inicio - 1})', formato_celda_unicos_simple)
        worksheet.write(f'E{fila_inicio + 4}', f'=C{fila_inicio}', formato_celda_unicos_simple)


        worksheet.write(f'A{fila_inicio + 5}', 'GRAN TOTAL', formato_celda_unicos)

        worksheet.write(f'B{fila_inicio + 5}', f'=B{fila_inicio + 4}', formato_celda_unicos_simple)
        worksheet.write(f'C{fila_inicio + 5}', f'=C{fila_inicio + 4}', formato_celda_unicos_simple)
        worksheet.write(f'D{fila_inicio + 5}', f'=D{fila_inicio + 4}', formato_celda_unicos_simple)
        worksheet.write(f'E{fila_inicio + 5}', f'=E{fila_inicio + 4}', formato_celda_unicos_simple)

        # =========================================================
        # Construir el rango dinámico
        rango_celdas = f'A{fila_inicio + 8}:H{fila_inicio + 8}' 

        worksheet.merge_range(rango_celdas, 'OBSERVACIONES', merge_format)

        # Construir el rango dinámico
        rango_celdas = f'A{fila_inicio + 9}:H{fila_inicio + 12}' 

        # Combinar celdas rango_celdas
        worksheet.merge_range(rango_celdas, '')

        # Construir el rango dinámico
        rango_celdas = f'A{fila_inicio + 14}:H{fila_inicio + 15}' 

        # Combinar las celdas y escribir el texto
        worksheet.merge_range(rango_celdas, "La presente certificación se expide como soporte de pago y con base en el registro diario de Titulares de Derecho, que se diligencia en cada Institución Educativa atendida.", 
                            workbook.add_format({
                                'align': 'left',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',
                                'font_size': 12,
                                'text_wrap': True,
                                'border': 1
                            }))
        
        # Construir el rango dinámico
        rango_celdas = f'A{fila_inicio + 17}:H{fila_inicio + 17}' 

        # Combinar las celdas y escribir el texto
        worksheet.merge_range(rango_celdas, "PARA CONSTANCIA SE FIRMA EN:", 
                            workbook.add_format({
                                'align': 'left',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',
                                'font_size': 12,
                                'text_wrap': True,
                                'border': 1
                            }))
        
        # =========================================================
                # Crear un solo formato reutilizable
        formato_celda_unicos_inferior = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 12,
            'border': 1
        })

        worksheet.write(f'A{fila_inicio + 18}', 'FECHA:', formato_celda_unicos_inferior)
        worksheet.write(f'B{fila_inicio + 18}', 'DÍA:', formato_celda_unicos_inferior)
        worksheet.write(f'C{fila_inicio + 18}', '', formato_celda_unicos_inferior)
        worksheet.write(f'D{fila_inicio + 18}', 'DE:', formato_celda_unicos_inferior)
        worksheet.write(f'E{fila_inicio + 18}', '', formato_celda_unicos_inferior)
        worksheet.write(f'F{fila_inicio + 18}', 'AÑO', formato_celda_unicos_inferior)

        # Construir el rango dinámico
        rango_celdas = f'G{fila_inicio + 18}:H{fila_inicio + 18}' 

        # Combinar celdas rango_celdas
        worksheet.merge_range(rango_celdas, '')

        formato_firma = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Aptos Narrow',
            'font_size': 12
        })

        worksheet.write(f'A{fila_inicio + 19}', 'FIRMA DEL RECTOR', formato_firma)

        # Construir el rango dinámico
        rango_celdas = f'A{fila_inicio + 20}:H{fila_inicio + 23}' 

        # Combinar celdas rango_celdas
        worksheet.merge_range(rango_celdas, '')

        worksheet.write(f'A{fila_inicio + 24}', 'NOMBRES Y APELLIDOS DEL RECTOR', formato_celda_unicos)

        # Construir el rango dinámico
        rango_celdas = f'B{fila_inicio + 24}:H{fila_inicio + 24}' 

        # Combinar las celdas y escribir el texto
        worksheet.merge_range(rango_celdas, rector_institucion, 
                            workbook.add_format({
                                'align': 'center',
                                'valign': 'vcenter',
                                'font_name': 'Aptos Narrow',
                                'font_size': 12,
                                'text_wrap': True,
                                'border': 1
                            }))

        # =========================================================

        # Lista de celdas combinadas para evitar sobrescribir su formato
        celdas_combinadas = [
            'D2:H5',  # Ejemplo de combinación
            'A7:H7',
            'B8:E8',
            'G8:H8',
            'A17:H17',
            'A23:A24',
            'B23:B24',
            'C23:H23',
            'F24:H24',
            'A2:C5',
            f'A{fila_inicio + 8}:H{fila_inicio + 8}', # Rango de celdas combinadas
            f'A{fila_inicio + 9}:H{fila_inicio + 12}',
            f'G{fila_inicio + 18}:H{fila_inicio + 18}'

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

        # 1. Desactivar la cuadrícula
        worksheet.hide_gridlines(2)  # 2 es para ocultar la cuadrícula en la vista de diseño

        # 2. Rellenar toda la hoja con color blanco
        formato_blanco = workbook.add_format({'bg_color': 'white'})
        worksheet.set_column('A:Z', None, formato_blanco)  # Rellenar celdas de la A a la Z con color blanco (ajustar según el número de columnas)

        # 3. Ajustar el tamaño de las columnas de acuerdo con un ancho específico
        column_widths = {
            'A': 36.42,  # Ancho de la columna A
            'B': 22.75,  # Ancho de la columna B
            'C': 22.75,  # Ancho de la columna C
            'D': 22.75,  # Ancho de la columna D
            'E': 19.42,  # Ancho de la columna E
            'F': 15.17,  # Ancho de la columna F
            'G': 13.42,  # Ancho de la columna G
            'H': 14,     # Ancho de la columna H
        }

        # Asignar el ancho especificado a cada columna
        for col, width in column_widths.items():
            worksheet.set_column(f'{col}:{col}', width)  # Establecer el ancho para cada columna

        # =========================================================

        # Formato para borde izquierdo grueso
        formato_borde_izquierdo = workbook.add_format({'left': 1})  

        # Aplicar borde en la columna "I" (índice 8) desde la fila 1 hasta fila_inicio + 24
        for fila in range(0, fila_inicio + 24):  # Filas desde 1 (0 en Python) hasta fila_inicio + 24
            worksheet.write_blank(fila, 8, None, formato_borde_izquierdo)  # "I" es la columna 8 (A=0, B=1, ..., I=8

        # Formato para borde superior grueso
        formato_borde_superior = workbook.add_format({'top': 1})  # Borde más grueso

        # Aplicar el formato a la fila 25 desde A hasta H
        for col in range(0, 8):  # Columnas A (0) hasta H (7)
            worksheet.write_blank(fila_inicio + 24, col, None, formato_borde_superior)  # Aplicar borde sin sobrescribir datos

        # =========================================================

        # Guardar el archivo con las modificaciones
        writer.close()

    def separar_dataframes(self):

        # Cargar el archivo de Excel
        df_focalizacion = pd.read_excel(self.ruta_archivo_aplicacion_novedades, dtype={"DANE": str})

        # Crear un diccionario para almacenar los DataFrames separados
        dfs_separados = {}

        # Agrupar por 'INSTITUCION'
        for institucion, df_grupo in df_focalizacion.groupby(['INSTITUCION']):
            dfs_separados[institucion] = df_grupo

        for (institucion), df_grupo in dfs_separados.items():
            # Obtener el nombre de la institución
            var_institucion = df_grupo['INSTITUCION'].iloc[0]  # Tomar el primer valor de 'INSTITUCION'
            var_dane_institucion = df_grupo['DANE'].iloc[0]

            # Generar la certificación
            self.generar_certificacion(var_institucion, var_dane_institucion)

    def main(self):
        print("\nINICIA PROCESO DE GENERACION DE CERTIFICACIONES\n")
        self.separar_dataframes()

if __name__ == "__main__":
    generador = GeneradorCertificaciones()
    generador.main()

