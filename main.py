# Importar la clase
from src.planillas_en_blanco import TemplateGenerator  
from src.actualizar_json import ActualizarJson
from src.planillas_diligenciadas import GeneradorPlantillas
from src.certificador import GeneradorCertificaciones

# Importar librerías
from PIL import ImageTk, Image, ImageDraw
from tkinter import messagebox
from PIL import Image, ImageTk  # Asegúrate de tener instalada la librería Pillow
import customtkinter as ctk
from tkinter import font
import tkinter as tk 
import threading
import os  

COLOR_BARRA_SUPERIOR = "#080D0E" 
COLOR_MENU_LATERAL = "#0A2533"  
COLOR_CUERPO_PRINCIPAL = "#E5E9F0"
COLOR_MENU_CURSOR_ENCIMA = "#2f88c5"

def leer_imagen( path, size): 
        return ImageTk.PhotoImage(Image.open(path).resize(size,  Image.ADAPTIVE))  

def centrar_ventana(ventana,aplicacion_ancho,aplicacion_largo):    
    pantall_ancho = ventana.winfo_screenwidth()
    pantall_largo = ventana.winfo_screenheight()
    x = int((pantall_ancho/2) - (aplicacion_ancho/2))
    y = int((pantall_largo/2) - (aplicacion_largo/2))
    return ventana.geometry(f"{aplicacion_ancho}x{aplicacion_largo}+{x}+{y}")

class FormularioMaestroDesign(tk.Tk):

    def __init__(self, **kwargs):
        super().__init__() 
        self.kwargs = kwargs
        # Rutas imagenes
        logo_ventas = os.path.join(os.getcwd(), "util/img/logo_ventas.png")
        logo_perfil = os.path.join(os.getcwd(), "util/img/logo_perfil.png")
        sitio_construccion = os.path.join(os.getcwd(), "util/img/sitio_construccion.png")
        # Cargar imagenes
        self.logo = leer_imagen(logo_ventas, (1060, 300))
        self.perfil = leer_imagen(logo_perfil, (100, 100))
        self.img_sitio_construccion = leer_imagen(sitio_construccion, (200, 200))
        self.config_window()
        self.paneles()
        self.controles_barra_superior()        
        self.controles_menu_lateral()
        self.controles_cuerpo()
    
    def config_window(self):
        # Rutas imagenes
        logo_banco = os.path.join(os.getcwd(), "util/img/logo_banco.ico")
        # Configuración inicial de la ventana
        self.title('Optimizamos procesos, maximizamos resultados')
        self.iconbitmap(logo_banco)
        w, h = 1024, 600        
        centrar_ventana(self, w, h)        

    def paneles(self):        
        # Crear paneles: barra superior, menú lateral y cuerpo principal
        self.barra_superior = tk.Frame(
            self, bg=COLOR_BARRA_SUPERIOR, height=50)
        self.barra_superior.pack(side=tk.TOP, fill='both')      

        self.menu_lateral = tk.Frame(self, bg=COLOR_MENU_LATERAL, width=150)
        self.menu_lateral.pack(side=tk.LEFT, fill='both', expand=False) 
        
        self.cuerpo_principal = tk.Frame(
            self, bg=COLOR_CUERPO_PRINCIPAL)
        self.cuerpo_principal.pack(side=tk.RIGHT, fill='both', expand=True)
    
    def controles_barra_superior(self):
        # Configuración de la barra superior
        font_awesome = font.Font(family='FontAwesome', size=12)

        # Etiqueta de título
        self.labelTitulo = tk.Label(self.barra_superior, text="NEXUS CODE") # "NEXUS CODE"
        self.labelTitulo.config(fg="#fff", font=(
            "Poppins", 15, "bold"), bg=COLOR_BARRA_SUPERIOR, pady=10, width=16)
        self.labelTitulo.pack(side=tk.LEFT)

        # Botón del menú lateral
        self.buttonMenuLateral = tk.Button(self.barra_superior, text="☰", font=font_awesome,
                                            command=self.toggle_panel, bd=0, bg=COLOR_BARRA_SUPERIOR, fg="white")
        self.buttonMenuLateral.pack(side=tk.LEFT)

    def controles_menu_lateral(self):
        # Configuración del menú lateral
        ancho_menu = 20
        alto_menu = 2
        font_awesome = font.Font(family='FontAwesome', size=15)
        
        # Etiqueta de perfil
        self.labelPerfil = tk.Label(
            self.menu_lateral, image=self.perfil, bg=COLOR_MENU_LATERAL)
        self.labelPerfil.pack(side=tk.TOP, pady=10)

        # Botones del menú lateral
        
        self.buttonDashBoard = tk.Button(self.menu_lateral)        
        self.buttonFunza = tk.Button(self.menu_lateral)        
        self.buttonFaca = tk.Button(self.menu_lateral)        
        self.buttonInfo = tk.Button(self.menu_lateral)        

        buttons_info = [
            ("💻 Municipio FUNZA", "\n", self.buttonFunza, self.abrir_menu_planillas_funza),
            ("💻 Municipio FACA", "\n", self.buttonFaca, self.abrir_menu_planillas_faca),
            ("📃 Documentacion", "\n", self.buttonInfo, self.abrir_panel_en_construccion) #uf570
        ]

        for text, icon, button, comando in buttons_info:
            self.configurar_boton_menu(button, text, icon, font_awesome, ancho_menu, alto_menu, comando)                    
    
    def controles_cuerpo(self):
        # Imagen en el cuerpo principal
        label = tk.Label(self.cuerpo_principal, image=self.logo,
                            bg=COLOR_CUERPO_PRINCIPAL)
        label.place(x=0, y=0, relwidth=1, relheight=1)

    def configurar_boton_menu(self, button, text, icon, font_awesome, ancho_menu, alto_menu, comando):
        button.config(text=f"  {icon}    {text}", anchor="w", font=font_awesome,
                        bd=0, bg=COLOR_MENU_LATERAL, fg="white", width=ancho_menu, height=alto_menu,
                        command = comando)
        button.pack(side=tk.TOP)
        self.bind_hover_events(button)

    def bind_hover_events(self, button):
        # Asociar eventos Enter y Leave con la función dinámica
        button.bind("<Enter>", lambda event: self.on_enter(event, button))
        button.bind("<Leave>", lambda event: self.on_leave(event, button))

    def on_enter(self, event, button):
        # Cambiar estilo al pasar el ratón por encima
        button.config(bg=COLOR_MENU_CURSOR_ENCIMA, fg='white')

    def on_leave(self, event, button):
        # Restaurar estilo al salir el ratón
        button.config(bg=COLOR_MENU_LATERAL, fg='white')

    def toggle_panel(self):
        # Alternar visibilidad del menú lateral
        if self.menu_lateral.winfo_ismapped():
            self.menu_lateral.pack_forget()
        else:
            self.menu_lateral.pack(side=tk.LEFT, fill='y')

    def abrir_panel_en_construccion(self):   
        self.limpiar_panel(self.cuerpo_principal)     
        FormularioSitioConstruccionDesign(self.cuerpo_principal,self.img_sitio_construccion)
    
    def abrir_menu_planillas_funza(self):   
        self.limpiar_panel(self.cuerpo_principal)     
        FormularioProcesoFunza(self.cuerpo_principal)

    def abrir_menu_planillas_faca(self):   
        self.limpiar_panel(self.cuerpo_principal)     
        FormularioProcesoFaca(self.cuerpo_principal)

    def limpiar_panel(self,panel):
    # Función para limpiar el contenido del panel
        for widget in panel.winfo_children():
            widget.destroy()

class VentanaAdicional:
    def __init__(self, parent):
        # Crear y configurar la ventana adicional
        self.parent = parent
        self.resultado = None  # Variable para almacenar el resultado
        
        # Crear y configurar la ventana adicional
        self.ventana = tk.Toplevel(parent)

        self.config_window()
        self.ventana.transient(parent)  # Para que la ventana se muestre sobre la principal
        self.ventana.grab_set()  # Bloquear interacción con la ventana principal
        self.ventana.focus_force()  # Forzar el foco en la ventana adicional

        # Contenedor para los botones
        button_frame = tk.Frame(self.ventana)
        button_frame.pack(pady=20)

        btn1 = tk.Button(button_frame, text="Continuar", command=self.boton1_accion, bg="#2a3138", fg="white", width=12, height=2, font=("Arial", 9, "bold"))
        btn2 = tk.Button(button_frame, text="Abandonar", command=self.boton2_accion, bg="#2a3138", fg="white", width=12, height=2, font=("Arial", 9, "bold"))

        btn1.pack(side=tk.LEFT, padx=10)
        btn2.pack(side=tk.LEFT, padx=10)

    def config_window(self):
        # Rutas imagenes
        logo_banco = os.path.join(os.getcwd(), "util/img/logo_banco.ico")
        
        # Configuración inicial de la ventana
        self.ventana.title('¿Estas seguro de ejecutar el proceso?')
        self.ventana.iconbitmap(logo_banco)
        # Dimensiones de la ventana
        w, h = 300, 100
        # Centrar la ventana
        centrar_ventana(self.ventana, w, h)
        # Configuración de estilo similar a la ventana principal
        self.ventana.configure(bg="#f1faff")

    def boton1_accion(self):
        self.resultado = 1
        self.ventana.destroy()

    def boton2_accion(self):
        self.resultado = 0
        self.ventana.destroy()

    def obtener_resultado(self):
        return self.resultado

class FormularioProcesoFunza():

    def ejecutar_generador_plantillas_vacias(self):
        try:
            # Instanciar la clase TemplateGenerator y ejecutar el método main
            salida_proceso = TemplateGenerator().main()
            # Insertar el resultado en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f"{salida_proceso}\n\n")
        except Exception as e:
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error durante la ejecución\n{e}\n\n")

    def ejecutar_generador_plantillas_diligenciadas(self):
        try:
            # Instanciar la clase TemplateGenerator y ejecutar el método main
            salida_proceso = GeneradorPlantillas().main()
            # Insertar el resultado en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f"{salida_proceso}\n\n")
        except Exception as e:
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error durante la ejecución\n{e}\n\n")

    def ejecutar_generador_plantillas_pdf(self):
        try:
            # Instanciar la clase TemplateGenerator y ejecutar el método main
            salida_proceso = GeneradorPlantillas().convertir_pdf()
            # Insertar el resultado en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f"{salida_proceso}\n\n")
        except Exception as e:
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error durante la ejecución\n{e}\n\n")

    def __init__(self, panel_principal):    

        self.panel_principal = panel_principal  # Almacenar panel_principal en un atributo

        # Crear paneles: barra superior
        self.barra_superior = tk.Frame(panel_principal)
        self.barra_superior.pack(side=tk.TOP, fill=tk.X, expand=False) 

        # Asegúrate de configurar el color de fondo en barra_superior
        self.barra_superior.config(bg=COLOR_CUERPO_PRINCIPAL)

        # Crear paneles: barra inferior
        self.barra_inferior = tk.Frame(panel_principal)
        self.barra_inferior.pack(side=tk.BOTTOM, fill='both', expand=True)  

        # Asegúrate de configurar el color de fondo en barra_superior
        self.barra_inferior.config(bg=COLOR_CUERPO_PRINCIPAL)

        # Primer Label con texto
        self.labelTitulo = tk.Label(
            self.barra_superior, text="BIENVENIDO")
        self.labelTitulo.config(fg="#222d33", font=("Roboto", 30), bg=COLOR_CUERPO_PRINCIPAL)
        self.labelTitulo.pack(side=tk.TOP, fill='both', expand=True)

        # Segundo Label con texto
        self.labelsubTitulo = tk.Label(
            self.barra_superior, text="GESTIÓN DE PLANILLAS FUNZA")
        self.labelsubTitulo.config(fg="#222d33", font=("Roboto", 30), bg=COLOR_CUERPO_PRINCIPAL)
        self.labelsubTitulo.pack(side=tk.TOP, fill='both', expand=True)

        # Añadir espacio en la parte inferior usando pady
        self.labelsubTitulo.pack(side=tk.TOP, fill='both', expand=True, pady=(0, 20))  
        # 0 para la parte superior, 20 para la inferior

        # Crear subpanel para los botones
        self.subpanel_botones = tk.Frame(self.barra_inferior, bg= COLOR_CUERPO_PRINCIPAL) # bg="#222d33"
        self.subpanel_botones.pack(side=tk.TOP, fill='x')

        # Crear botones
        self.button1 = ctk.CTkButton(
            self.subpanel_botones, text="Generar planillas en blanco", command=self.action1,
            width=150, height=40, fg_color="#2a3138", text_color="white",
            corner_radius=20, font=("Arial", 14))

        self.button2 = ctk.CTkButton(
            self.subpanel_botones, text="Generar planillas diligenciadas", command=self.action2,
            width=150, height=40, fg_color="#2a3138", text_color="white",
            corner_radius=20, font=("Arial", 14))

        self.button3 = ctk.CTkButton(
            self.subpanel_botones, text="Convertir a pdf", command=self.action3,
            width=150, height=40, fg_color="#2a3138", text_color="white",
            corner_radius=20, font=("Arial", 14))

        # Ubicar los botones en la cuadrícula
        self.button1.grid(row=0, column=0, padx=20, pady=10, sticky='ew')
        self.button2.grid(row=0, column=1, padx=20, pady=10, sticky='ew')
        self.button3.grid(row=0, column=2, padx=20, pady=10, sticky='ew')

        # Configurar las columnas para que se expandan uniformemente
        self.subpanel_botones.grid_columnconfigure(0, weight=1)
        self.subpanel_botones.grid_columnconfigure(1, weight=1)
        self.subpanel_botones.grid_columnconfigure(2, weight=1)

        # Crear subpanel para el widget de texto y colocarlo en la parte inferior
        self.subpanel_texto = tk.Frame(self.barra_inferior, bg=COLOR_CUERPO_PRINCIPAL)
        self.subpanel_texto.pack(side=tk.BOTTOM, fill='both', expand=True)

        # Crear widget Text para mostrar mensajes
        self.text_widget = tk.Text(self.subpanel_texto, width=40, height=22)
        # Espacio: 20 arriba, 10 abajo, 20 izquierda/derecha
        self.text_widget.pack(padx=20, pady=(20, 20), fill='both', expand=True) 

        # Configurar el tamaño de las columnas
        self.barra_inferior.grid_columnconfigure(0, weight=1)
        self.barra_inferior.grid_columnconfigure(1, weight=3)

        # Crear paneles: barra inferior
        self.barra_final = tk.Frame(panel_principal, bg=COLOR_CUERPO_PRINCIPAL)
        self.barra_final.pack(side=tk.BOTTOM, fill='both', expand=True)

    def action1(self):
        try:
            # Crear la ventana adicional y esperar a que se cierre
            ventana_adicional = VentanaAdicional(self.panel_principal)  # Asegúrate de pasar la ventana principal correcta
            self.panel_principal.wait_window(ventana_adicional.ventana)  # Esperar a que la ventana adicional se cierre

            # Obtener el resultado de la ventana adicional
            resultado = ventana_adicional.obtener_resultado()  

            # Usar el resultado
            if resultado == 1:
                
                # Actualizar el archivo JSON
                ActualizarJson().ejecutar("municipio_proceso", "FUNZA") 

                # Insertar el resultado en el text_widget
                self.text_widget.insert(tk.END, "Inicio generacion de planillas vacias para el municipio de FUNZA\n .....\n")
                self.text_widget.update_idletasks() 

                # Ejecutar la función en un hilo separado
                hilo = threading.Thread(target=self.ejecutar_generador_plantillas_vacias, daemon=True)
                hilo.start()

            elif resultado == 0:
                print("Proceso cancelado")
            else:
                print("No se realizó ninguna selección")

            
        except Exception as e:
            # Insertar el error en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error al ejecutar el proceso\n{e}\n\n")

    def action2(self):
        try:
            # Crear la ventana adicional y esperar a que se cierre
            ventana_adicional = VentanaAdicional(self.panel_principal)  # Asegúrate de pasar la ventana principal correcta
            self.panel_principal.wait_window(ventana_adicional.ventana)  # Esperar a que la ventana adicional se cierre

            # Obtener el resultado de la ventana adicional
            resultado = ventana_adicional.obtener_resultado()  

            # Usar el resultado
            if resultado == 1:
                
                # Actualizar el archivo JSON
                ActualizarJson().ejecutar("municipio_proceso", "FUNZA") 

                # Insertar el resultado en el text_widget
                self.text_widget.insert(tk.END, "Inicio generacion de planillas diligenciadas y certificaciones para el municipio de FUNZA\n .....\n")
                self.text_widget.update_idletasks() 

                # Ejecutar la función en un hilo separado
                hilo = threading.Thread(target=self.ejecutar_generador_plantillas_diligenciadas, daemon=True)
                hilo.start()

            elif resultado == 0:
                print("Proceso cancelado")
            else:
                print("No se realizó ninguna selección")

            
        except Exception as e:
            # Insertar el error en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error al ejecutar el proceso\n{e}\n\n")

    def action3(self):
        try:
            # Crear la ventana adicional y esperar a que se cierre
            ventana_adicional = VentanaAdicional(self.panel_principal)  # Asegúrate de pasar la ventana principal correcta
            self.panel_principal.wait_window(ventana_adicional.ventana)  # Esperar a que la ventana adicional se cierre

            # Obtener el resultado de la ventana adicional
            resultado = ventana_adicional.obtener_resultado()  

            # Usar el resultado
            if resultado == 1:
                
                # Actualizar el archivo JSON
                ActualizarJson().ejecutar("municipio_proceso", "FUNZA") 

                # Insertar el resultado en el text_widget
                self.text_widget.insert(tk.END, "Inicio generacion de planillas diligenciadas y certificaciones para el municipio de FUNZA\n .....\n")
                self.text_widget.update_idletasks() 

                # Ejecutar la función en un hilo separado
                hilo = threading.Thread(target=self.ejecutar_generador_plantillas_pdf, daemon=True)
                hilo.start()

            elif resultado == 0:
                print("Proceso cancelado")
            else:
                print("No se realizó ninguna selección")

        except Exception as e:
            # Insertar el error en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error al ejecutar el proceso\n{e}\n\n")

class FormularioProcesoFaca():

    def ejecutar_generador_plantillas_vacias(self):
        try:
            # Instanciar la clase TemplateGenerator y ejecutar el método main
            salida_proceso = TemplateGenerator().main()
            # Insertar el resultado en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f"{salida_proceso}\n\n")
        except Exception as e:
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error durante la ejecución\n{e}\n\n")

    def ejecutar_generador_plantillas_diligenciadas(self):
        try:
            # Instanciar la clase TemplateGenerator y ejecutar el método main
            salida_proceso = GeneradorPlantillas().main()
            # Insertar el resultado en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f"{salida_proceso}\n\n")
        except Exception as e:
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error durante la ejecución\n{e}\n\n")

    def ejecutar_generador_certificaciones(self):
        try:
            # Instanciar la clase TemplateGenerator y ejecutar el método main
            salida_proceso = GeneradorCertificaciones().main()
            # Insertar el resultado en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f"{salida_proceso}\n\n")
        except Exception as e:
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error durante la ejecución\n{e}\n\n")

    def ejecutar_generador_plantillas_pdf(self):
        try:
            # Instanciar la clase TemplateGenerator y ejecutar el método main
            salida_proceso = GeneradorPlantillas().convertir_pdf()
            # Insertar el resultado en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f"{salida_proceso}\n\n")
        except Exception as e:
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error durante la ejecución\n{e}\n\n")

    def __init__(self, panel_principal):    

        self.panel_principal = panel_principal

        # Crear barra superior
        self.barra_superior = tk.Frame(panel_principal, bg=COLOR_CUERPO_PRINCIPAL)
        self.barra_superior.pack(side=tk.TOP, fill=tk.X)

        # Títulos
        self.labelTitulo = tk.Label(
            self.barra_superior, text="BIENVENIDO", fg="#222d33",
            font=("Roboto", 30), bg=COLOR_CUERPO_PRINCIPAL)
        self.labelTitulo.pack(side=tk.TOP, fill='both', expand=True)

        self.labelsubTitulo = tk.Label(
            self.barra_superior, text="GESTIÓN DE PLANILLAS FACA", fg="#222d33",
            font=("Roboto", 30), bg=COLOR_CUERPO_PRINCIPAL)
        self.labelsubTitulo.pack(side=tk.TOP, fill='both', expand=True, pady=(0, 20))

        # Contenedor principal (izquierda y derecha)
        self.barra_inferior = tk.Frame(panel_principal, bg=COLOR_CUERPO_PRINCIPAL)
        self.barra_inferior.pack(side=tk.TOP, fill='both', expand=True)

        # Subpanel IZQUIERDO (botones + imagen)
        self.panel_izquierdo = tk.Frame(self.barra_inferior, bg=COLOR_CUERPO_PRINCIPAL)
        self.panel_izquierdo.pack(side=tk.LEFT, fill='both', expand=True, padx=10, pady=10)

        # Subpanel de botones
        self.subpanel_botones = tk.Frame(self.panel_izquierdo, bg=COLOR_CUERPO_PRINCIPAL)
        self.subpanel_botones.pack(side=tk.TOP, fill='x')

        # Crear botones con tamaños individuales y estilos personalizados
        self.button1 = ctk.CTkButton(
            self.subpanel_botones, text="Generar planillas en blanco", command=self.action1,
            width=235, height=40, fg_color="#2a3138", text_color="white",
            corner_radius=20, font=("Arial", 14))
        self.button1.pack(padx=20, pady=(10, 5))

        self.button2 = ctk.CTkButton(
            self.subpanel_botones, text="Generar planillas diligenciadas", command=self.action2,
            width=235, height=40, fg_color="#2a3138", text_color="white",
            corner_radius=20, font=("Arial", 14))
        self.button2.pack(padx=20, pady=5)

        self.button3 = ctk.CTkButton(
            self.subpanel_botones, text="Certificador", command=self.action3,
            width=235, height=40, fg_color="#2a3138", text_color="white",
            corner_radius=20, font=("Arial", 14))
        self.button3.pack(padx=20, pady=5)

        self.button4 = ctk.CTkButton(
            self.subpanel_botones, text="Convertir a pdf", command=self.action4,
            width=235, height=40, fg_color="#2a3138", text_color="white",
            corner_radius=20, font=("Arial", 14))
        self.button4.pack(padx=20, pady=(5, 10))

        dibujo_programador = os.path.join(os.getcwd(), "util/img/Imagen perfil.png")

        # Imagen debajo de los botones (redimensionada)
        try:
            imagen = Image.open(dibujo_programador)
            imagen = imagen.resize((200, 150), Image.Resampling.LANCZOS)
            imagen_tk = ImageTk.PhotoImage(imagen)

            self.imagen_label = tk.Label(self.panel_izquierdo, image=imagen_tk, bg=COLOR_CUERPO_PRINCIPAL)
            self.imagen_label.image = imagen_tk
            self.imagen_label.pack(side=tk.TOP, pady=20)
        except Exception as e:
            print(f"No se pudo cargar la imagen: {e}")

        # Subpanel DERECHO (Text Widget)
        self.subpanel_texto = tk.Frame(self.barra_inferior, bg=COLOR_CUERPO_PRINCIPAL)
        self.subpanel_texto.pack(side=tk.LEFT, fill='both', expand=True, padx=(0, 10))  # eliminamos pady

        self.text_widget = tk.Text(self.subpanel_texto, width=40, height=22)
        self.text_widget.pack(fill='both', expand=True, padx=20, pady=(10, 20))  # solo pequeño espacio superior

        # Panel final inferior (si lo usas)
        self.barra_final = tk.Frame(panel_principal, bg=COLOR_CUERPO_PRINCIPAL)
        self.barra_final.pack(side=tk.BOTTOM, fill='both', expand=True)

    def action1(self):
        try:
            # Crear la ventana adicional y esperar a que se cierre
            ventana_adicional = VentanaAdicional(self.panel_principal)  # Asegúrate de pasar la ventana principal correcta
            self.panel_principal.wait_window(ventana_adicional.ventana)  # Esperar a que la ventana adicional se cierre

            # Obtener el resultado de la ventana adicional
            resultado = ventana_adicional.obtener_resultado()  

            # Usar el resultado
            if resultado == 1:
                
                # Actualizar el archivo JSON
                ActualizarJson().ejecutar("municipio_proceso", "FACA") 

                # Insertar el resultado en el text_widget
                self.text_widget.insert(tk.END, " Inicio generacion de planillas\n en blanco para el municipio de FACA\n .....\n")
                self.text_widget.update_idletasks() 

                # Ejecutar la función en un hilo separado
                hilo = threading.Thread(target=self.ejecutar_generador_plantillas_vacias, daemon=True)
                hilo.start()

            elif resultado == 0:
                print("Proceso cancelado")
            else:
                print("No se realizó ninguna selección")

        except Exception as e:
            # Insertar el error en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error al ejecutar el proceso\n{e}\n\n")

    def action2(self):
        try:
            # Crear la ventana adicional y esperar a que se cierre
            ventana_adicional = VentanaAdicional(self.panel_principal)  # Asegúrate de pasar la ventana principal correcta
            self.panel_principal.wait_window(ventana_adicional.ventana)  # Esperar a que la ventana adicional se cierre

            # Obtener el resultado de la ventana adicional
            resultado = ventana_adicional.obtener_resultado()  

            # Usar el resultado
            if resultado == 1:
                
                # Actualizar el archivo JSON
                ActualizarJson().ejecutar("municipio_proceso", "FACA") 

                # Insertar el resultado en el text_widget
                self.text_widget.insert(tk.END, " Inicio generacion de planillas\n diligenciadas para el municipio de FACA\n .....\n")
                self.text_widget.update_idletasks() 

                # Ejecutar la función en un hilo separado
                hilo = threading.Thread(target=self.ejecutar_generador_plantillas_diligenciadas, daemon=True)
                hilo.start()

            elif resultado == 0:
                print("Proceso cancelado")
            else:
                print("No se realizó ninguna selección")

        except Exception as e:
            # Insertar el error en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error al ejecutar el proceso\n{e}\n\n")

    def action3(self):
        try:
            # Crear la ventana adicional y esperar a que se cierre
            ventana_adicional = VentanaAdicional(self.panel_principal)  # Asegúrate de pasar la ventana principal correcta
            self.panel_principal.wait_window(ventana_adicional.ventana)  # Esperar a que la ventana adicional se cierre

            # Obtener el resultado de la ventana adicional
            resultado = ventana_adicional.obtener_resultado()  

            # Usar el resultado
            if resultado == 1:
                
                # Actualizar el archivo JSON
                ActualizarJson().ejecutar("municipio_proceso", "FACA") 

                # Insertar el resultado en el text_widget
                self.text_widget.insert(tk.END, " Inicio generacion de certificaciones\n diligenciadas para el municipio de FACA\n .....\n")
                self.text_widget.update_idletasks() 

                # Ejecutar la función en un hilo separado 
                hilo = threading.Thread(target=self.ejecutar_generador_certificaciones, daemon=True)
                hilo.start()

            elif resultado == 0:
                print("Proceso cancelado")
            else:
                print("No se realizó ninguna selección")

            
        except Exception as e:
            # Insertar el error en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error al ejecutar el proceso\n{e}\n\n")
    
    def action4(self):
        try:
            # Crear la ventana adicional y esperar a que se cierre
            ventana_adicional = VentanaAdicional(self.panel_principal)  # Asegúrate de pasar la ventana principal correcta
            self.panel_principal.wait_window(ventana_adicional.ventana)  # Esperar a que la ventana adicional se cierre

            # Obtener el resultado de la ventana adicional
            resultado = ventana_adicional.obtener_resultado()  

            # Usar el resultado
            if resultado == 1:
                
                # Actualizar el archivo JSON
                ActualizarJson().ejecutar("municipio_proceso", "FACA") 

                # Insertar el resultado en el text_widget
                self.text_widget.insert(tk.END, " Inicio proceso de convertir\n exceles a formato pdf ej el municipio de FACA\n .....\n")
                self.text_widget.update_idletasks() 

                # Ejecutar la función en un hilo separado
                hilo = threading.Thread(target=self.ejecutar_generador_plantillas_pdf, daemon=True)
                hilo.start()

            elif resultado == 0:
                print("Proceso cancelado")
            else:
                print("No se realizó ninguna selección")

            
        except Exception as e:
            # Insertar el error en el text_widget
            self.text_widget.after(0, self.text_widget.insert, tk.END, f" Error al ejecutar el proceso\n{e}\n\n")

class FormularioSitioConstruccionDesign():

    def __init__(self, panel_principal, logo):

        # Crear paneles: barra superior
        self.barra_superior = tk.Frame( panel_principal)
        self.barra_superior.pack(side=tk.TOP, fill=tk.X, expand=False) 

        # Crear paneles: barra inferior
        self.barra_inferior = tk.Frame( panel_principal)
        self.barra_inferior.pack(side=tk.BOTTOM, fill='both', expand=True)  

        # Primer Label con texto
        self.labelTitulo = tk.Label(
            self.barra_superior, text="Página en construcción")
        self.labelTitulo.config(fg="#222d33", font=("Roboto", 30), bg=COLOR_CUERPO_PRINCIPAL)
        self.labelTitulo.pack(side=tk.TOP, fill='both', expand=True)

        # Segundo Label con la imagen
        self.label_imagen = tk.Label(self.barra_inferior, image=logo)
        self.label_imagen.place(x=0, y=0, relwidth=1, relheight=1)
        self.label_imagen.config(fg="#fff", font=("Roboto", 10), bg=COLOR_CUERPO_PRINCIPAL)

app = FormularioMaestroDesign()
app.mainloop() 