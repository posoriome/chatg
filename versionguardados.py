import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import shutil
import json
from datetime import datetime
from tkinter import ttk
from tkinter import messagebox
import pyodbc

# Establecer la conexión con la base de datos
server = 'LAPTOP-5J961U9N\SQLDEV'
database = 'base1'

conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes')
cursor = conn.cursor()

class FileDrop(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Arrastrar y soltar archivo JSON")
        self.geometry("850x300")

        # Campo ID
        tk.Label(self, text="ID:").grid(row=0, column=0, sticky="w", pady=5)
        self.entry_id = tk.Entry(self, state='readonly', width=10)
        self.entry_id.grid(row=0, column=1, sticky="w", pady=5)

        # Botón Asignar ID
        self.btn_asignar_id = tk.Button(self, text="Asignar ID", command=self.asignar_id, state="disabled")
        self.btn_asignar_id.grid(row=0, column=1, sticky="w", pady=5, padx=(80,0))

        # Campo Identificación
        tk.Label(self, text="Identificación:").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_identificacion = tk.Entry(self, width=30, name="identificacion")
        self.entry_identificacion.grid(row=1, column=1, sticky="w", pady=5)
        self.entry_identificacion.focus()
        self.entry_identificacion.bind("<Return>", self.cargar_valores)
        self.entry_identificacion.bind("<KeyRelease>", self.actualizar_estado_boton_asignar_id)

        # Campo Nombre
        tk.Label(self, text="Nombre:").grid(row=2, column=0, sticky="w", pady=5)
        self.combo_nombre = ttk.Combobox(self, state='readonly', width=120)
        self.combo_nombre.grid(row=2, column=1, sticky="w", pady=5)
        self.combo_nombre.bind("<<ComboboxSelected>>", self.actualizar_estado_boton_asignar_id)

        # Campo Tipo
        tk.Label(self, text="Tipo:").grid(row=3, column=0, sticky="w", pady=5)
        self.combo_tipo = ttk.Combobox(self, state='readonly')
        self.combo_tipo.grid(row=3, column=1, sticky="w", pady=5)
        self.combo_tipo.bind("<<ComboboxSelected>>", self.actualizar_estado_boton_asignar_id)

        # Campo Estado
        tk.Label(self, text="Estado:").grid(row=4, column=0, sticky="w", pady=5)
        self.combo_estado = ttk.Combobox(self, state='readonly')
        self.combo_estado.grid(row=4, column=1, sticky="w", pady=5)
        self.combo_estado.bind("<<ComboboxSelected>>", self.actualizar_estado_boton_asignar_id)

        # Campo Año
        tk.Label(self, text="Año:").grid(row=5, column=0, sticky="w", pady=5)
        self.combo_anio = ttk.Combobox(self, values=list(range(2015, 2026)), width=27)
        self.combo_anio.grid(row=5, column=1, sticky="w", pady=5)
        self.combo_anio.bind("<<ComboboxSelected>>", self.actualizar_estado_boton_asignar_id)

        # Campo Mes
        tk.Label(self, text="Mes:").grid(row=6, column=0, sticky="w", pady=5)
        self.combo_mes = ttk.Combobox(self, values=["ENERO", "FEBRERO", "MARZO", "ABRIL"], width=27)
        self.combo_mes.grid(row=6, column=1, sticky="w", pady=5)
        self.combo_mes.bind("<<ComboboxSelected>>", self.actualizar_estado_boton_asignar_id)

        # Campo Archivo
        self.label_archivo = tk.Label(self, text="Arrastra y suelta un archivo JSON aquí")
        self.label_archivo.grid(row=7, columnspan=2, pady=10)

        # Botón Guardar
        self.btn_guardar = tk.Button(self, text="Guardar", command=self.guardar_archivo)
        self.btn_guardar.grid(row=8, columnspan=2, pady=10)

        # Variables para almacenar datos
        self.archivo_capturado = None

        # Configurar la funcionalidad de arrastrar y soltar
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        print("Evento de soltar detectado.")
        archivos = event.data.strip('{}').split()  # Eliminar las llaves de la ruta del archivo
        archivos_json = []  # Lista para almacenar los archivos JSON válidos

        for archivo in archivos:
            if archivo.lower().endswith('.json'):
                archivos_json.append(archivo)
                print(f"Archivo JSON capturado: {archivo}")

        if archivos_json:
            # Procesa los archivos JSON aquí (por ejemplo, carga su contenido o realiza operaciones específicas)
            # También puedes actualizar la interfaz gráfica para mostrar todos los archivos aceptados.
            self.label_archivo.config(text=f"Archivos JSON capturados: {', '.join(archivos_json)}")
        else:
            print("Ningún archivo válido de tipo JSON encontrado.")
            self.label_archivo.config(text="No se encontraron archivos JSON válidos. Por favor, arrastra archivos JSON.")


    def guardar_archivo(self):
        print("Intentando guardar el archivo...")
        if self.archivo_capturado and os.path.isfile(self.archivo_capturado):
            print("Archivo capturado y válido.")
            # Obtener la fecha actual en el formato AAAAMMDD
            fecha_actual = datetime.now().strftime("%Y%m%d")
            # Ruta de la carpeta de destino
            ruta_destino = os.path.join("D:/", fecha_actual)
            ruta_destino2 = os.path.join("D:/originales/", fecha_actual)

            # Crear la carpeta de destino si no existe
            if not os.path.exists(ruta_destino):
                os.makedirs(ruta_destino)
            
            # Crear la carpeta de destino originales si no existe
            if not os.path.exists(ruta_destino2):
                os.makedirs(ruta_destino2)

            try:
                nombre_archivo = os.path.basename(self.archivo_capturado)  # Obtener el nombre del archivo
                ruta_final = os.path.join(ruta_destino, nombre_archivo)
                ruta_final2 = os.path.join(ruta_destino2, nombre_archivo)
                shutil.copy2(self.archivo_capturado, ruta_final)
                shutil.copy2(self.archivo_capturado, ruta_final2)
                print(f"Archivo guardado en: {ruta_final}")
                print(f"Archivo guardado en: {ruta_final2}")
                self.archivo_capturado = None  # Eliminar la referencia al archivo capturado del cajón de adjuntos
                print(f"Archivo adjunto eliminado del cajón de adjuntos")
                self.label_archivo.config(text=f"Archivo guardado y adjunto eliminado del cajón de adjuntos")
                # Limpiar los campos después de guardar
                identificacion = self.entry_identificacion.get()
                nombre = self.combo_nombre.get()
                tipo = self.combo_tipo.get()
                estado = self.combo_estado.get()
                anio = self.combo_anio.get()
                mes = self.combo_mes.get()
                id = self.entry_id.get()

                self.limpiar_campos()

                # Leer el contenido actual del archivo JSON
                with open(ruta_final, 'r') as f:
                    #datos_json = json.load(f)
                    original_json = json.load(f)

                 # Copiar el contenido original a otra variable
                datos_json = original_json.copy()

                # Añadir nuevos atributos al objeto raíz del archivo JSON
                datos_json[0].update({
                    "Identificacion": identificacion,
                    "Nombre": nombre,
                    "Tipo": tipo,
                    "Estado": estado,
                    "Ano": anio,
                    "Mes": mes,
                    "IdContrato": id
                })

                 # Escribir el diccionario de nuevo en el archivo JSON
                with open(ruta_final, 'w') as f:
                    json.dump(datos_json, f, indent=4)

                print("Atributos agregados al archivo JSON.")
                # Limpiar los campos después de guardar
                self.entry_identificacion.delete(0, tk.END)
                self.combo_nombre.set("")
                self.combo_tipo.set("")
                self.combo_estado.set("")
                self.combo_anio.set("")
                self.combo_mes.set("")
                self.entry_id.set("")


            except Exception as e:
                print(f"Error al guardar el archivo: {e}")
        else:
            print("No se ha capturado ningún archivo JSON o el archivo no es válido.")

    def cargar_valores(self, event=None):
        # Obtener el valor del campo ID
        id_filtro = self.entry_identificacion.get()

        # Realizar la consulta con el ID filtrado
        query = "SELECT estado, tipo, identificacion, nombre FROM dbo.tabla1 WHERE identificacion = ?"
        cursor.execute(query, (id_filtro,))
        rows = cursor.fetchall()

        if rows:
            # Limpiar listas desplegables
            self.combo_tipo['values'] = []
            self.combo_nombre['values'] = []
            self.combo_estado['values'] = []

            # Obtener valores únicos de cada campo
            tipos = set()
            nombres = set()
            estados = set()

            for row in rows:
                tipos.add(row.tipo)
                nombres.add(row.nombre)
                estados.add(row.estado)

            # Cargar valores únicos en las listas desplegables
            self.combo_tipo['values'] = list(tipos)
            self.combo_nombre['values'] = list(nombres)
            self.combo_estado['values'] = list(estados)

        else:
            messagebox.showinfo("Sin registros", "No existen registros asociados a esta identificación")

    def dnd_enter(self, event):
        # Ocultar todos los widgets durante el arrastre
        for widget in self.winfo_children():
            widget.grid_forget()

    def dnd_leave(self, event):
        # Volver a mostrar todos los widgets después del arrastre
        row = 0
        for widget in self.winfo_children():
            widget.grid(row=row, columnspan=2, sticky="w", pady=5)
            row += 1

    def asignar_id(self):
        # Obtener los valores de los campos
        identificacion = self.entry_identificacion.get()
        nombre = self.combo_nombre.get()
        tipo = self.combo_tipo.get()
        estado = self.combo_estado.get()

        # Realizar la consulta con los filtros
        query = "SELECT ID FROM dbo.tabla1 WHERE identificacion = ? AND nombre = ? AND tipo = ? AND estado = ?"
        cursor.execute(query, (identificacion, nombre, tipo, estado))
        rows = cursor.fetchall()

        # Verificar la cantidad de registros devueltos
        if len(rows) == 1:
            self.entry_id.config(state='normal')
            self.entry_id.delete(0, tk.END)
            self.entry_id.insert(0, rows[0][0])
            self.entry_id.config(state='readonly')
        elif len(rows) > 1:
            messagebox.showerror("Error", "Esta identificación tiene más de 1 registro coincidente")
        else:
            messagebox.showinfo("Sin registros", "No se encontraron registros para los filtros proporcionados")

    def actualizar_estado_boton_asignar_id(self, event=None):
        # Verificar si los campos están llenos y diferentes de nulos o vacíos
        if self.entry_identificacion.get() and self.combo_nombre.get() and self.combo_tipo.get() and self.combo_estado.get():
            self.btn_asignar_id.config(state="normal")
        else:
            self.btn_asignar_id.config(state="disabled")

    def limpiar_campos(self):
        self.entry_identificacion.delete(0, tk.END)
        self.combo_nombre.set("")
        self.combo_tipo.set("")
        self.combo_estado.set("")
        self.combo_anio.set("")
        self.combo_mes.set("")

if __name__ == "__main__":
    app = FileDrop()
    app.mainloop()

# Cerrar el cursor y la conexión
cursor.close()
conn.close()