import pandas as pd
import webbrowser as web
import pyautogui as pg
import time
import os
import tkinter as tk

def mostrar_mensaje():
    mensaje_formato = entrada_mensaje.get("1.0", tk.END)

    # Cargar el archivo Excel
    data = pd.read_excel("Clientes.xlsx", sheet_name='Ventas')

    # Iterar sobre las filas y crear el mensaje deseado
    for i in range(len(data)):
        celular = str(data.loc[i, 'Celular'])  # Convertir a string para que se añada al mensaje

        # Crear mensaje personalizado dinámico
        mensaje = mensaje_formato.format(**data.loc[i].to_dict())

        # Abrir una nueva pestaña para entrar a WhatsApp Web
        web.open("https://web.whatsapp.com/send?phone=" + celular + "&text=" + mensaje) 
        time.sleep(8)           # Esperar 8 segundos a que cargue
        pg.click(1226, 984)      # Hacer click en la caja de texto
        time.sleep(1)           # Esperar 2 segundos
         
        try:
            imagen = data.loc[i, 'imagen']
            # Hacer clic en el ícono de adjuntar (ajusta estas coordenadas según sea necesario)
            pg.click(715, 986)
            time.sleep(1)
            pg.click(802, 640)
            # Adjuntar la imagen
            time.sleep(1)
            pg.write(imagen)  # Escribe la ruta de la imagen
            time.sleep(1)
            pg.press('enter')  # Confirma la selección de la imagen
            time.sleep(1)
            pg.press('enter')  # Confirma la selección de la imagen
            time.sleep(1)

        except KeyError:
            pass

        pg.press('enter')       # Enviar mensaje 
        time.sleep(1)           # Esperar 3 segundos a que se envíe el mensaje
        pg.hotkey('ctrl', 'w')  # Cerrar la pestaña
        time.sleep(1)
        time.sleep(1)

# Crear la ventana Tkinter para ingresar el mensaje
ventana_mensaje = tk.Tk()
ventana_mensaje.title("Ingresar Mensaje Personalizado")

# Crear una etiqueta y una entrada para ingresar el mensaje
etiqueta_ingreso_mensaje = tk.Label(ventana_mensaje, text="Ingresa el mensaje personalizado (usa {Cliente}, {Producto}, {descuento}, {codigo}):")
etiqueta_ingreso_mensaje.pack(pady=5)

entrada_mensaje = tk.Text(ventana_mensaje, width=60, height=40)
entrada_mensaje.pack(pady=5)

# Crear un botón para procesar el mensaje
boton_procesar = tk.Button(ventana_mensaje, text="Procesar Mensaje", command=mostrar_mensaje)
boton_procesar.pack(pady=10)

# Iniciar el bucle de eventos de la ventana de ingreso de mensajes
ventana_mensaje.mainloop()
