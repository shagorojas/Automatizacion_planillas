import pathlib
import json 
import os 

class ActualizarJson:

    def __init__(self):
        ruta_master = os.path.join(str(os.path.abspath(pathlib.Path().absolute())))
        self.ruta_json = os.path.join(ruta_master, "Config", "Config.json")

    def leer_json(self):

        # Leemos el archivo JSON que contiene la configuración del proceso
        with open(self.ruta_json) as contenido:
            self.config = json.load(contenido)  # Almacena la configuración en el atributo config
        return self.config
    
    def escribir_json(self, data):
        # Escribe el JSON actualizado de vuelta al archivo
        with open(self.ruta_json, 'w') as archivo:
            json.dump(data, archivo, indent=4)

    def ejecutar(self, llave, valor):
        # Leer el archivo JSON
        params = self.leer_json()
        
        # Asigna el nuevo valor a la llave especificada
        params[llave] = valor
        
        # Imprime el JSON formateado
        json_formateado = json.dumps(params, indent=4)
        print(json_formateado)
        
        # Escribe el JSON actualizado de vuelta al archivo
        self.escribir_json(params)

# if __name__ == "__main__":
#     aj = ActualizarJson()
#     aj.ejecutar("municipio_proceso", "FUNZA")
