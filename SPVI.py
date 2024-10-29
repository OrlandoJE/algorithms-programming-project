# Sistema de punto de venta e inventario

# Importe de librerias
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, font # Para mostrar mensajes en una ventana emergente, un cuadro de diálogo que recupera información y cambiar el tipo de letra
import openpyxl as xl # Para trabajar con archivos de Excel
from datetime import datetime, timedelta # Para obtener la fecha y hora actual del sistema
import matplotlib.pyplot as plt # Para crear gráficas

# Ruta del archivo de usuarios
archivo_usuarios = "usuarios.txt" # Archivo de texto plano

# Función para cargar Excel
def cargar_excel():
    try: # Intentar abrir el archivo
        wb = xl.load_workbook("SPVI.xlsx") # Cargar el libro
        ws_productos = wb["Productos"] # Cargar la hoja de productos
        ws_inventario = wb["Inventario"] # Cargar la hoja de inventario
        ws_ventas = wb["Ventas"] # Cargar la hoja de ventas
    except FileNotFoundError: # Si el archivo no existe, crearlo
        wb = xl.Workbook() # Crear un nuevo libro
        
        ws_productos = wb.active # Cargar la hoja activa
        ws_productos.title = "Productos" # Cambiar el nombre de la hoja
        ws_productos.append(["Código", "Descripcion", "Provedor", "Precio Unitario", "Comentarios"]) # Agregar encabezados
        
        ws_inventario = wb.create_sheet("Inventario") # Crear una nueva hoja
        ws_inventario.append(["Código", "Descripcion", "Comentarios", "Existencia"]) # Agregar encabezados
        
        ws_ventas = wb.create_sheet("Ventas") # Crear una nueva hoja
        ws_ventas.append(["Código", "Descripcion", "Comentarios", "Precio Unitario", "Cantidad", "Subtotal", "Fecha"]) # Agregar encabezados
        
        wb.save("SPVI.xlsx") # Guardar el libro
        
    return wb, ws_productos, ws_inventario, ws_ventas # Retornar el libro y las hojas

# Función para cargar los usuarios desde el archivo
def cargar_usuarios():
    try:
        with open(archivo_usuarios, "r") as archivo: # Abrir el archivo en modo lectura
            contenido = archivo.read() # Leer el contenido del archivo
            if contenido:
                return eval(contenido)  # Convertir el contenido en un diccionario, en lenguaje python
    except FileNotFoundError: # Si el archivo no existe
        pass # No hacer nada
    return {} # Retornar un diccionario

# Función para guardar los usuarios en el archivo
def guardar_usuarios(usuarios):
    with open(archivo_usuarios, "w") as file: # Abrir el archivo en modo escritura
        file.write(str(usuarios))  # Convertir el diccionario en una cadena

# Función para iniciar sesión
def iniciar_sesion():
    usuario = entry_usuario.get() # Obtener el usuario
    password = entry_contraseña.get() # Obtener la contraseña

    if usuario in usuarios_registrados: # Si el usuario está registrado
        if usuarios_registrados[usuario] == password: # Si la contraseña es correcta
            if usuario == "admin": # Si el usuario es admin
                ventana_admin() # Mostrar la ventana de administrador
            else:
                ventana_usuario() # Mostrar la ventana de usuario
                
            # Borrar el contenido de los campos de entrada
            entry_usuario.delete(0, tk.END)
            entry_contraseña.delete(0, tk.END)
        else:
            messagebox.showerror("Error de inicio de sesión", "Contraseña incorrecta.")
    else:
        messagebox.showerror("Error de inicio de sesión", "El usuario no está registrado.")

# Ventana de administrador
def ventana_admin():
    ventana_principal.withdraw() # Ocultar la ventana principal
    
    # Crear ventana de administrador
    ventana_admin = tk.Toplevel(ventana_principal) # Crear una ventana
    ventana_admin.title("Panel de Administrador") # Cambiar el título de la ventana
    ventana_admin.geometry("450x300") # Cambiar el tamaño de la ventana

    lbl_usuarios_registrados = tk.Label(ventana_admin, text="Usuarios registrados:") # Crear un label
    lbl_usuarios_registrados.pack() # Acomodar el label

    listbox_usuarios = tk.Listbox(ventana_admin) # Crear un listbox
    listbox_usuarios.pack() # Acomodar el listbox
    
    for usuario in usuarios_registrados: # Recorrer los usuarios registrados
        listbox_usuarios.insert(tk.END, usuario) # Agregar el usuario al listbox
    
    def registrar_usuario():
        nuevo_usuario = tk.simpledialog.askstring("Registrar Usuario", "Ingrese el nombre de usuario:") # Mostrar un cuadro de diálogo para ingresar el nombre de usuario
        if nuevo_usuario:
            if nuevo_usuario not in usuarios_registrados: # Si el usuario no está registrado
                password = tk.simpledialog.askstring("Registrar Usuario", "Ingrese la contraseña:") # Mostrar un cuadro de diálogo para ingresar la contraseña
                if password:
                    usuarios_registrados[nuevo_usuario] = password # Agregar el usuario y la contraseña al diccionario
                    listbox_usuarios.insert(tk.END, nuevo_usuario) # Agregar el usuario al listbox
                    guardar_usuarios(usuarios_registrados) # Guardar los usuarios en el archivo
                    messagebox.showinfo("Registro Exitoso", "El usuario ha sido registrado correctamente.")
            else:
                messagebox.showerror("Error de registro", "El usuario ya está registrado.")
    
    btn_registrar = tk.Button(ventana_admin, text="Registrar Usuario", command=registrar_usuario) # Crear un botón
    btn_registrar.pack() # Acomodar el botón
    
    def mostrar_contraseña():
        seleccion = listbox_usuarios.curselection() # Obtener el índice del elemento seleccionado
        if seleccion:
            usuario_seleccionado = listbox_usuarios.get(seleccion) # Obtener el usuario seleccionado
            contraseña = usuarios_registrados[usuario_seleccionado] # Obtener la contraseña del usuario seleccionado
            messagebox.showinfo("Contraseña", f"Contraseña del usuario {usuario_seleccionado} es: {contraseña}") # Mostrar la contraseña, la letra f es para poder usar variables dentro de la cadena

    btn_mostrar_contraseña = tk.Button(ventana_admin, text="Mostrar Contraseña", command=mostrar_contraseña)
    btn_mostrar_contraseña.pack()
    
    def modificar_contraseña():
        seleccion = listbox_usuarios.curselection()
        if seleccion:
            usuario_seleccionado = listbox_usuarios.get(seleccion)
            nueva_contraseña = tk.simpledialog.askstring("Modificar Contraseña", f"Ingrese la nueva contraseña para {usuario_seleccionado}: ")
            if nueva_contraseña:
                usuarios_registrados[usuario_seleccionado] = nueva_contraseña
                guardar_usuarios(usuarios_registrados)
                messagebox.showinfo("Contraseña Modificada", f"La contraseña de {usuario_seleccionado} ha sido modificada correctamente.")

    btn_modificar_contraseña = tk.Button(ventana_admin, text="Modificar Contraseña", command=modificar_contraseña)
    btn_modificar_contraseña.pack()

    def eliminar_usuario():
        seleccion = listbox_usuarios.curselection()
        if seleccion:
            usuario_seleccionado = listbox_usuarios.get(seleccion)
            if len(usuarios_registrados) == 1 and usuario_seleccionado == "admin":
                messagebox.showerror("Error de eliminación", "No se puede eliminar el único usuario administrador.")
            else:
                confirmacion = messagebox.askyesno("Eliminar Usuario", f"¿Está seguro de eliminar al usuario '{usuario_seleccionado}'?")
                if confirmacion:
                    del usuarios_registrados[usuario_seleccionado]
                    listbox_usuarios.delete(seleccion)
                    guardar_usuarios(usuarios_registrados)  # Guardar los usuarios
                    messagebox.showinfo("Eliminación Exitosa", "El usuario ha sido eliminado correctamente.")
    
    btn_eliminar = tk.Button(ventana_admin, text="Eliminar Usuario", command=eliminar_usuario)
    btn_eliminar.pack()
    
    def cerrar_ventana_admin():
        ventana_admin.destroy()
        ventana_principal.deiconify()  # Volver a mostrar la ventana principal
        
    def cerrar_programa():
        ventana_principal.destroy()

    ventana_admin.protocol("WM_DELETE_WINDOW", cerrar_programa)  # Configurar evento de cierre
    
    # Crear el menú
    menu_admin = tk.Menu(ventana_admin)
    ventana_admin.config(menu=menu_admin)

    # Crear el menú "Productos" 
    menu_admin.add_command(label="Productos", command=ventana_productos)
    
    # Crear el menú "Inventario" 
    menu_admin.add_command(label="Inventario", command=ventana_inventario)

    # Crear el menú "Punto de Venta"
    menu_admin.add_command(label="Punto de Venta", command=ventana_punto_venta)
    
    # Crear el menú "Reportes"
    menu_admin.add_command(label="Reportes", command=ventana_reportes)
    
    # Crear el menú "Acerca de" con sus opciones
    menu_acerca_de = tk.Menu(menu_admin, tearoff=0)
    menu_admin.add_cascade(label="Ayuda", menu=menu_acerca_de)
    menu_acerca_de.add_command(label="Acerca de", command=ventana_acerca_de)
    menu_acerca_de.add_command(label="Instrucciones de uso", command=ventana_instrucciones_de_uso)
    
    # Crear el menú "Cerrar sesión"
    menu_admin.add_command(label="Cerrar sesión", command=cerrar_ventana_admin)

# Ventana de usuario regular
def ventana_usuario():
    ventana_principal.withdraw()
    
    ventana_usuario = tk.Toplevel(ventana_principal)
    ventana_usuario.title("Panel de Usuario")
    ventana_usuario.geometry("700x350")
    
    label_bienvenida = tk.Label(ventana_usuario, text="¡Bienvenido!")
    label_bienvenida.pack(pady=10)
    
    # Crear un marco para centrar los widgets
    marco_centro = tk.Frame(ventana_usuario)
    marco_centro.pack(expand=True)
    
    # Crear el botón "Productos" con la figura del camión
    boton_productos = tk.Button(marco_centro, text="🚚\nProductos", font=font.Font(size=16), command=ventana_productos, width=8, height=3, bg="#FFE0B2")
    boton_productos.grid(row=0, column=0, padx=10, pady=10)

    # Crear el botón "Inventario" con la figura de la caja
    boton_inventario = tk.Button(marco_centro, text="📦\nInventario", font=font.Font(size=16), command=ventana_inventario, width=8, height=3, bg="#FFCCBC")
    boton_inventario.grid(row=0, column=1, padx=10, pady=10)
    
    # Crear el botón "Ventas" con la figura del símbolo de dólar
    boton_ventas = tk.Button(marco_centro, text="💲\nVentas", font=font.Font(size=16), command=ventana_punto_venta, width=8, height=3, bg="#B2DFDB")
    boton_ventas.grid(row=0, column=2, padx=10, pady=10)

    # Crear el botón "Reportes" con la figura del reporte en papel
    boton_reportes = tk.Button(marco_centro, text="📊\nReportes", font=font.Font(size=16), command=ventana_reportes, width=8, height=3, bg="#B3E5FC")
    boton_reportes.grid(row=0, column=3, padx=10, pady=10)
    
    def cerrar_programa():
        ventana_principal.destroy()
        
    def cerrar_ventana_usuario():
        ventana_usuario.destroy()
        ventana_principal.deiconify()  # Volver a mostrar la ventana principal

    ventana_usuario.protocol("WM_DELETE_WINDOW", cerrar_programa)  # Configurar evento de cierre
    
    # Crear menu usuario
    menu_usuario = tk.Menu(ventana_usuario)
    ventana_usuario.config(menu=menu_usuario)
    
    # Crear el menú "Acerca de" con sus opciones
    menu_acerca_de = tk.Menu(menu_usuario, tearoff=0)
    menu_usuario.add_cascade(label="Ayuda", menu=menu_acerca_de)
    menu_acerca_de.add_command(label="Acerca de", command=ventana_acerca_de)
    menu_acerca_de.add_command(label="Instrucciones de uso", command=ventana_instrucciones_de_uso)
    
    # Crear el menú "Cerrar sesión"
    menu_usuario.add_command(label="Cerrar sesión", command=cerrar_ventana_usuario)

# Ventana de productos
def ventana_productos():
    # Cargar los datos del archivo de Excel en el Treeview
    def cargar_datos():
        titulos_columnas = [columna.value for columna in ws_productos[1]]

        lista_productos.delete(*lista_productos.get_children())

        lista_productos["columns"] = tuple(range(len(titulos_columnas)))
        lista_productos["displaycolumns"] = tuple(range(len(titulos_columnas)))
        
        datos_originales.clear()

        for i, titulo in enumerate(titulos_columnas):
            lista_productos.heading(i, text=titulo)

        for row in ws_productos.iter_rows(min_row=2, values_only=True):
            lista_productos.insert("", tk.END, values=row)
            datos_originales.append(row)

        datos_filtrados.extend(datos_originales)
    
    # Función para filtrar los resultados en el Treeview
    def filtrar_resultados(event):
        texto_busqueda = entry_buscar_producto.get().lower()

        if texto_busqueda:
            coincidencias = [datos for datos in datos_originales if any(texto_busqueda in str(valor).lower() for valor in datos)]
            datos_filtrados.clear()
            datos_filtrados.extend(coincidencias)
        else:
            datos_filtrados.clear()
            datos_filtrados.extend(datos_originales)

        # Restaurar los elementos coincidentes en el Treeview
        lista_productos.delete(*lista_productos.get_children())
        for datos in datos_filtrados:
            lista_productos.insert("", tk.END, values=datos)
    
    # Función para abrir la ventana de registro de producto
    def abrir_ventana_registro():
        ventana_registro = tk.Toplevel(ventana_productos)
        ventana_registro.title("Registrar Producto")
        ventana_registro.geometry("300x300")

        # Variables para almacenar los datos del nuevo producto
        codigo = tk.StringVar()
        descripcion = tk.StringVar()
        proveedor = tk.StringVar()
        precio = tk.StringVar()
        comentarios = tk.StringVar()
        
        # Función para registrar el nuevo producto
        def registrar_producto():
            nuevo_codigo = codigo.get()
            nueva_descripcion = descripcion.get()

            # Verificar si el código o la descripción ya existen en la hoja de productos
            for fila in ws_productos.iter_rows(min_row=2, values_only=True):
                if fila[0] == nuevo_codigo or fila[1] == nueva_descripcion:
                    messagebox.showerror("Error de registro", "El código o la descripción ya existen.")
                    return
            
            nuevo_producto = [codigo.get(), descripcion.get(), proveedor.get(), precio.get(), comentarios.get()]
            ws_productos.append(nuevo_producto)
            wb.save("SPVI.xlsx")
            cargar_datos()
            ventana_registro.destroy()
            messagebox.showinfo("Registro Exitoso", "El producto ha sido registrado correctamente.")
            
        # Etiquetas y campos de entrada para los datos del producto
        label_codigo = tk.Label(ventana_registro, text="Código:")
        label_codigo.pack()
        entry_codigo = tk.Entry(ventana_registro, textvariable=codigo)
        entry_codigo.pack()

        label_descripcion = tk.Label(ventana_registro, text="Descripción:")
        label_descripcion.pack()
        entry_descripcion = tk.Entry(ventana_registro, textvariable=descripcion)
        entry_descripcion.pack()

        label_proveedor = tk.Label(ventana_registro, text="Proveedor:")
        label_proveedor.pack()
        entry_proveedor = tk.Entry(ventana_registro, textvariable=proveedor)
        entry_proveedor.pack()

        label_precio = tk.Label(ventana_registro, text="Precio Unitario:")
        label_precio.pack()
        entry_precio = tk.Entry(ventana_registro, textvariable=precio)
        entry_precio.pack()

        label_comentarios = tk.Label(ventana_registro, text="Comentarios:")
        label_comentarios.pack()
        entry_comentarios = tk.Entry(ventana_registro, textvariable=comentarios)
        entry_comentarios.pack()

        # Botón para registrar el producto
        btn_registrar = tk.Button(ventana_registro, text="Registrar", command=registrar_producto)
        btn_registrar.pack(pady=10)
        
    # Función para modificar el producto seleccionado
    def modificar_producto():
        # Obtener el índice del elemento seleccionado en el Treeview
        seleccion = lista_productos.focus()
        if seleccion:
            ventana_modificar = tk.Toplevel(ventana_productos)
            ventana_modificar.title("Modificar Producto")
            ventana_modificar.geometry("300x300")

            # Obtener los datos actuales del producto seleccionado
            datos_actuales = lista_productos.item(seleccion)["values"]

            # Variables para almacenar los nuevos datos del producto
            codigo = tk.StringVar(value=datos_actuales[0])
            descripcion = tk.StringVar(value=datos_actuales[1])
            proveedor = tk.StringVar(value=datos_actuales[2])
            precio = tk.StringVar(value=datos_actuales[3])
            comentarios = tk.StringVar(value=datos_actuales[4])

            # Función para guardar los cambios en el producto
            def guardar_cambios():
                nuevos_datos = [codigo.get(), descripcion.get(), proveedor.get(), precio.get(), comentarios.get()]
                indice = lista_productos.index(seleccion)
                ws_productos.delete_rows(indice+2)  # Eliminar la fila correspondiente al producto seleccionado
                ws_productos.insert_rows(indice+2)  # Insertar una nueva fila en la posición del producto
                for col, valor in enumerate(nuevos_datos, start=1):
                    ws_productos.cell(row=indice+2, column=col).value = valor
                wb.save("SPVI.xlsx")  # Guardar los cambios en el archivo de Excel
                cargar_datos()  # Volver a cargar los datos en el Treeview
                ventana_modificar.destroy()
                messagebox.showinfo("Modificación Exitosa", "El producto ha sido modificado correctamente.")

            # Etiquetas y campos de entrada para los nuevos datos del producto
            label_codigo = tk.Label(ventana_modificar, text="Código:")
            label_codigo.pack()
            entry_codigo = tk.Entry(ventana_modificar, textvariable=codigo)
            entry_codigo.pack()

            label_descripcion = tk.Label(ventana_modificar, text="Descripción:")
            label_descripcion.pack()
            entry_descripcion = tk.Entry(ventana_modificar, textvariable=descripcion)
            entry_descripcion.pack()

            label_proveedor = tk.Label(ventana_modificar, text="Proveedor:")
            label_proveedor.pack()
            entry_proveedor = tk.Entry(ventana_modificar, textvariable=proveedor)
            entry_proveedor.pack()

            label_precio = tk.Label(ventana_modificar, text="Precio Unitario:")
            label_precio.pack()
            entry_precio = tk.Entry(ventana_modificar, textvariable=precio)
            entry_precio.pack()

            label_comentarios = tk.Label(ventana_modificar, text="Comentarios:")
            label_comentarios.pack()
            entry_comentarios = tk.Entry(ventana_modificar, textvariable=comentarios)
            entry_comentarios.pack()

            # Botón para guardar los cambios
            btn_guardar = tk.Button(ventana_modificar, text="Guardar Cambios", command=guardar_cambios)
            btn_guardar.pack(pady=10)

    # Función para eliminar el producto seleccionado
    def eliminar_producto():
        seleccion = lista_productos.focus()
        if seleccion:
            confirmacion = messagebox.askyesno("Eliminar Producto", "¿Está seguro de eliminar el producto seleccionado?")
            if confirmacion:
                indice = lista_productos.index(seleccion)
                ws_productos.delete_rows(indice+2)  # Eliminar la fila correspondiente al producto seleccionado
                wb.save("SPVI.xlsx")  # Guardar los cambios en el archivo de Excel
                cargar_datos()  # Volver a cargar los datos en el Treeview
                messagebox.showinfo("Eliminación Exitosa", "El producto ha sido eliminado correctamente.")
    
    #Crear Ventana de productos
    ventana_productos = tk.Toplevel(ventana_principal)
    ventana_productos.title("Productos")
    ventana_productos.geometry("1200x350")
    
    label_buscar_producto = tk.Label(ventana_productos, text="Buscar producto")
    label_buscar_producto.grid(row=0, column=0, padx=10)
    
    entry_buscar_producto = tk.Entry(ventana_productos, width=50)
    entry_buscar_producto.grid(row=1, column=0, padx=10)
    
    label_lista_productos = tk.Label(ventana_productos, text="Lista de productos")
    label_lista_productos.grid(row=2, column=0, padx=10, pady=10, columnspan=2)
    
    lista_productos = ttk.Treeview(ventana_productos, show="headings")
    lista_productos.grid(row=3, column=0, padx=10, rowspan=3)
    
    btn_registrar_producto = tk.Button(ventana_productos, text="Registrar producto", command=abrir_ventana_registro)
    btn_registrar_producto.grid(row=2, column=2, padx=10, pady=10)
    
    btn_modificar_producto = tk.Button(ventana_productos, text="Modificar producto", command=modificar_producto)
    btn_modificar_producto.grid(row=3, column=2, padx=10, pady=10)
    
    btn_eliminar_producto = tk.Button(ventana_productos, text="Eliminar producto", command=eliminar_producto)
    btn_eliminar_producto.grid(row=4, column=2, padx=10, pady=10)
    
    btn_cerrar_productos = tk.Button(ventana_productos, text="Cerrar productos", command=ventana_productos.destroy)
    btn_cerrar_productos.grid(row=5, column=2, padx=10, pady=10)
    
    # Crear menu usuario
    menu_productos = tk.Menu(ventana_productos)
    ventana_productos.config(menu=menu_productos)
    
    # Crear el menú "Acerca de" con sus opciones
    menu_acerca_de = tk.Menu(menu_productos, tearoff=0)
    menu_productos.add_cascade(label="Ayuda", menu=menu_acerca_de)
    menu_acerca_de.add_command(label="Acerca de", command=ventana_acerca_de)
    menu_acerca_de.add_command(label="Instrucciones de uso", command=ventana_instrucciones_de_uso)
    
    # Fin creacion de ventana productos
    
    # Variables para evitar errores
    datos_originales = []
    datos_filtrados = []
    
    # Reconocer cuando se está escribiendo en el entry de búsqueda
    entry_buscar_producto.bind("<KeyRelease>", filtrar_resultados)
    
    cargar_datos()

# Ventana de inventario
def ventana_inventario():
    # Sincronizar los datos entre las hojas "Productos" e "Inventario"
    def sincronizar_datos():
        # Obtener los rangos de celdas a sincronizar
        rango_productos = ws_productos['A2':'E' + str(ws_productos.max_row)]
        rango_inventario = ws_inventario['A2':'D' + str(ws_inventario.max_row)]

        # Crear un diccionario para almacenar los datos de productos existentes en el inventario
        productos_existentes = {}

        # Recorrer el rango de productos y guardar los datos en el diccionario
        for fila in rango_inventario:
            codigo = fila[0].value
            descripcion = fila[1].value
            productos_existentes[codigo] = descripcion

        # Recorrer el rango de productos y sincronizar los datos en la hoja de inventario
        for fila in rango_productos:
            codigo = fila[0].value
            descripcion = fila[1].value
            proveedor = fila[2].value
            precio_unitario = fila[3].value

            # Si el producto ya existe en el inventario, actualizar solo los comentarios
            if codigo in productos_existentes:
                comentario = fila[4].value
                ws_inventario.cell(row=fila[0].row, column=3).value = comentario
            # Si el producto no existe en el inventario, agregar una nueva fila
            else:
                comentario = fila[4].value
                cantidad = 0  # Establecer la cantidad inicial en 0
                ws_inventario.append([codigo, descripcion, comentario, cantidad])
                
        # Eliminar productos del inventario que no existen en la hoja de productos
        productos_a_eliminar = []
        for codigo, descripcion in productos_existentes.items():
            if codigo not in [fila[0].value for fila in rango_productos]:
                productos_a_eliminar.append(codigo)

        for codigo_eliminar in productos_a_eliminar:
            for fila in rango_inventario:
                if fila[0].value == codigo_eliminar:
                    ws_inventario.delete_rows(fila[0].row, amount=1)

        # Guardar los cambios en el libro de Excel
        wb.save("SPVI.xlsx")
        
    # Cargar los datos del archivo de Excel en el Treeview
    def cargar_datos_inventario():
        titulos_columnas = [columna.value for columna in ws_inventario[1]]

        lista_inventario.delete(*lista_inventario.get_children())

        lista_inventario["columns"] = tuple(range(len(titulos_columnas)))
        lista_inventario["displaycolumns"] = tuple(range(len(titulos_columnas)))
        
        datos_originales.clear()

        for i, titulo in enumerate(titulos_columnas):
            lista_inventario.heading(i, text=titulo)

        for row in ws_inventario.iter_rows(min_row=2, values_only=True):
            lista_inventario.insert("", tk.END, values=row)
            datos_originales.append(row)

        datos_filtrados.extend(datos_originales)
    
    # Función para filtrar los resultados en el Treeview
    def filtrar_resultados(event):
        texto_busqueda = entry_buscar_inventario.get().lower()

        if texto_busqueda:
            coincidencias = [datos for datos in datos_originales if any(texto_busqueda in str(valor).lower() for valor in datos)]
            datos_filtrados.clear()
            datos_filtrados.extend(coincidencias)
        else:
            datos_filtrados.clear()
            datos_filtrados.extend(datos_originales)

        # Restaurar los elementos coincidentes en el Treeview
        lista_inventario.delete(*lista_inventario.get_children())
        for datos in datos_filtrados:
            lista_inventario.insert("", tk.END, values=datos)
            
    def actualizar_cantidad(codigo, cantidad):
        for fila_index, fila in enumerate(ws_inventario.iter_rows(min_row=2, max_row=ws_inventario.max_row, values_only=True), start=2):
            if fila[0] == codigo:
                cantidad_actual = ws_inventario.cell(row=fila_index, column=4).value
                if cantidad_actual <= 0 and cantidad < 0:
                    messagebox.showerror("Error", "No se puede restar cuando el valor actual es cero.")
                else:
                    nueva_cantidad = cantidad_actual + cantidad
                    ws_inventario.cell(row=fila_index, column=4).value = nueva_cantidad
                    wb.save("SPVI.xlsx")
                    cargar_datos_inventario()
                    
                break
        else:
            messagebox.showerror("Error", "Producto no encontrado en el inventario.")
    
    def entrada_inventario():
        seleccion = lista_inventario.selection()  # Obtener la selección actual en el Treeview
        if seleccion:
            fila_seleccionada = lista_inventario.item(seleccion[0])  # Obtener los datos de la fila seleccionada
            codigo = fila_seleccionada['values'][0]  # Obtener el valor de la primera columna (código)
            cantidad = simpledialog.askinteger("Entrada de inventario", "Ingrese la cantidad a sumar:")
            if cantidad is not None:  # Verifica si se ingresó una cantidad válida
                actualizar_cantidad(codigo, cantidad)
        else:
            messagebox.showerror("Error", "Ningún producto seleccionado.")
            return None

    def mermas():
        seleccion = lista_inventario.selection()  # Obtener la selección actual en el Treeview
        if seleccion:
            fila_seleccionada = lista_inventario.item(seleccion[0])  # Obtener los datos de la fila seleccionada
            codigo = fila_seleccionada['values'][0]  # Obtener el valor de la primera columna (código)
            cantidad = simpledialog.askinteger("Mermas", "Ingrese la cantidad a restar:")
            if cantidad is not None:  # Verifica si se ingresó una cantidad válida
                actualizar_cantidad(codigo, -cantidad)  # Resta la cantidad ingresada al inventario
        else:
            messagebox.showerror("Error", "Ningún producto seleccionado.")
            return None
    
    ventana_inventario = tk.Toplevel(ventana_principal)
    ventana_inventario.title("Inventario")
    ventana_inventario.geometry("1000x350")
    
    label_lista_inventario = tk.Label(ventana_inventario, text="Lista de inventario")
    label_lista_inventario.grid(row=0, column=0, padx=10, pady=10)
    
    entry_buscar_inventario = tk.Entry(ventana_inventario, width=50)
    entry_buscar_inventario.grid(row=1, column=0, padx=10, pady=10)
    
    lista_inventario = ttk.Treeview(ventana_inventario, show="headings")
    lista_inventario.grid(row=2, column=0, padx=10, rowspan=2)
    
    btn_entrada_inventario = tk.Button(ventana_inventario, text="Entrada de inventario", command=entrada_inventario)
    btn_entrada_inventario.grid(row=2, column=1, padx=10, pady=10)
    
    btn_mermas = tk.Button(ventana_inventario, text="Mermas", command=mermas)
    btn_mermas.grid(row=3, column=1, padx=10, pady=10)
    
    btn_cerrar_inventario = tk.Button(ventana_inventario, text="Cerrar inventario", command=ventana_inventario.destroy)
    btn_cerrar_inventario.grid(row=4, column=1, padx=10)
    
    # Crear menu usuario
    menu_productos = tk.Menu(ventana_inventario)
    ventana_inventario.config(menu=menu_productos)
    
    # Crear el menú "Acerca de" con sus opciones
    menu_acerca_de = tk.Menu(menu_productos, tearoff=0)
    menu_productos.add_cascade(label="Ayuda", menu=menu_acerca_de)
    menu_acerca_de.add_command(label="Acerca de", command=ventana_acerca_de)
    menu_acerca_de.add_command(label="Instrucciones de uso", command=ventana_instrucciones_de_uso)
    
    # Variables para evitar errores
    datos_originales = []
    datos_filtrados = []
    
    # Reconocer cuando se está escribiendo en el entry de búsqueda
    entry_buscar_inventario.bind("<KeyRelease>", filtrar_resultados)
    
    sincronizar_datos()
    cargar_datos_inventario()

# Ventana de Punto de Venta
def ventana_punto_venta():
    def  titulos_lista_compras():
        titulos_columnas = [columna.value for columna in ws_ventas[1]]

        lista_compras.delete(*lista_compras.get_children())

        lista_compras["columns"] = tuple(range(len(titulos_columnas)))
        lista_compras["displaycolumns"] = tuple(range(len(titulos_columnas)))

        for i, titulo in enumerate(titulos_columnas):
            lista_compras.heading(i, text=titulo)
            lista_compras.column(i, width=120)
            
    def agregar_producto():
        # Sincrnoizar los datos de la hoja de inventario
        def sincronizar_datos():
            # Obtener los rangos de celdas a sincronizar
            rango_productos = ws_productos['A2':'E' + str(ws_productos.max_row)]
            rango_inventario = ws_inventario['A2':'D' + str(ws_inventario.max_row)]

            # Crear un diccionario para almacenar los datos de productos existentes en el inventario
            productos_existentes = {}

            # Recorrer el rango de productos y guardar los datos en el diccionario
            for fila in rango_inventario:
                codigo = fila[0].value
                descripcion = fila[1].value
                productos_existentes[codigo] = descripcion

            # Recorrer el rango de productos y sincronizar los datos en la hoja de inventario
            for fila in rango_productos:
                codigo = fila[0].value
                descripcion = fila[1].value
                proveedor = fila[2].value
                precio_unitario = fila[3].value

                # Si el producto ya existe en el inventario, actualizar solo los comentarios
                if codigo in productos_existentes:
                    comentario = fila[4].value
                    ws_inventario.cell(row=fila[0].row, column=3).value = comentario
                # Si el producto no existe en el inventario, agregar una nueva fila
                else:
                    comentario = fila[4].value
                    cantidad = 0  # Establecer la cantidad inicial en 0
                    ws_inventario.append([codigo, descripcion, comentario, cantidad])
                    
            # Eliminar productos del inventario que no existen en la hoja de productos
            productos_a_eliminar = []
            for codigo, descripcion in productos_existentes.items():
                if codigo not in [fila[0].value for fila in rango_productos]:
                    productos_a_eliminar.append(codigo)

            for codigo_eliminar in productos_a_eliminar:
                for fila in rango_inventario:
                    if fila[0].value == codigo_eliminar:
                        ws_inventario.delete_rows(fila[0].row, amount=1)

            # Guardar los cambios en el libro de Excel
            wb.save("SPVI.xlsx")
        
        # Cargar los datos de la hoja de inventario en la lista de productos
        def cargar_datos_inventario():
            titulos_columnas = [columna.value for columna in ws_inventario[1]]

            lista_productos.delete(*lista_productos.get_children())

            lista_productos["columns"] = tuple(range(len(titulos_columnas)))
            lista_productos["displaycolumns"] = tuple(range(len(titulos_columnas)))
            
            datos_originales_ap.clear()

            for i, titulo in enumerate(titulos_columnas):
                lista_productos.heading(i, text=titulo)
                lista_productos.column(i, width=150)

            for row in ws_inventario.iter_rows(min_row=2, values_only=True):
                lista_productos.insert("", tk.END, values=row)
                datos_originales_ap.append(row)

            datos_filtrados_ap.extend(datos_originales_ap)
        
        # Función para filtrar los resultados en el Treeview
        def filtrar_resultados_pv(event):
            texto_busqueda = entry_buscar_inventario.get().lower()

            if texto_busqueda:
                coincidencias = [datos for datos in datos_originales_ap if any(texto_busqueda in str(valor).lower() for valor in datos)]
                datos_filtrados_ap.clear()
                datos_filtrados_ap.extend(coincidencias)
            else:
                datos_filtrados_ap.clear()
                datos_filtrados_ap.extend(datos_originales_ap)
                
            # Restaurar los elementos coincidentes en el Treeview
            lista_productos.delete(*lista_productos.get_children())
            for datos in datos_filtrados_ap:
                lista_productos.insert("", tk.END, values=datos)
        
        def agregar():
            seleccion = lista_productos.selection()
            if seleccion:
                cantidad = simpledialog.askinteger("Cantidad", "Ingrese la cantidad de productos a agregar", parent=ventana_agregar_producto)
                
                if cantidad == None:
                    pass
                else:
                
                    producto_seleccionado = lista_productos.item(seleccion[0])["values"]
                    
                    existencia = int(producto_seleccionado[3])
                    
                    en_compras = 0
                    
                    # Verificar si el producto ya está en la lista de compras y obtener su cantidad
                    for item in lista_compras.get_children():
                        if lista_compras.item(item)["values"][:4] == producto_seleccionado[:4]:
                            en_compras = int(lista_compras.item(item)["values"][4])
                    
                    if (cantidad + en_compras) > existencia:
                        messagebox.showerror("Error", "No hay suficientes productos en el inventario")
                    else:
                        # Verificar si el producto ya está en la lista de compras
                        producto_existente = None
                        for item in lista_compras.get_children():
                            if lista_compras.item(item)["values"][:4] == producto_seleccionado[:4]:
                                producto_existente = item
                                
                        if producto_existente:
                            cantidad_existente = int(lista_compras.item(producto_existente)["values"][4])
                            nueva_cantidad = cantidad_existente + cantidad
                            lista_compras.item(producto_existente, values=producto_seleccionado + [nueva_cantidad] + [producto_seleccionado[3] * nueva_cantidad] + [datetime.now().strftime("%d/%m/%Y")])
                            
                        else:
                            lista_compras.insert("", tk.END, values=(lista_productos.item(seleccion[0])["values"] + [cantidad] + [lista_productos.item(seleccion[0])["values"][3] * cantidad] + [datetime.now().strftime("%d/%m/%Y")]))
                            
                        label_total["text"] = "Total: $" + str(sum([float(lista_compras.item(item)["values"][5]) for item in lista_compras.get_children()]))
                            
                        ventana_agregar_producto.destroy()
            else:
                messagebox.showerror("Error", "Seleccione un producto")
        
        ventana_agregar_producto = tk.Toplevel(ventana_punto_venta)
        ventana_agregar_producto.title("Agregar producto")
        ventana_agregar_producto.geometry("700x300")
        
        busqueda = tk.Frame(ventana_agregar_producto)
        busqueda.grid(row=0, column=0, padx=10, pady=10)
        
        label_buscar = tk.Label(busqueda, text="Buscar: ")
        label_buscar.grid(row=0, column=0, pady=10)
        
        entry_buscar_inventario = tk.Entry(busqueda, width=50)
        entry_buscar_inventario.grid(row=0, column=1, pady=10)
        
        lista_productos = ttk.Treeview(ventana_agregar_producto, show="headings")
        lista_productos.grid(row=1, column=0, padx=10, rowspan=2, columnspan=2)
        
        btn_agregar_a_venta = tk.Button(ventana_agregar_producto, text="Agregar", command=agregar)
        btn_agregar_a_venta.grid(row=1, column=2, padx=10, pady=10)
        
        btn_cerrar = tk.Button(ventana_agregar_producto, text="Cerrar", command=ventana_agregar_producto.destroy)
        btn_cerrar.grid(row=2, column=2, padx=10, pady=10)
    
        sincronizar_datos()
        cargar_datos_inventario()
        
        # Evento para filtrar los resultados en el Treeview
        entry_buscar_inventario.bind("<KeyRelease>", filtrar_resultados_pv)
    
    def modificar_cantidad():
        seleccion = lista_compras.selection()
        if seleccion:
            codigo_producto = lista_compras.item(seleccion[0])["values"][0]
            cantidad = simpledialog.askinteger("Cantidad", "Ingrese la nueva cantidad", parent=ventana_punto_venta)
            
            # Buscar la fila correspondiente al producto en el inventario
            for fila in ws_inventario.iter_rows(min_row=2, max_col=4):
                if fila[0].value == codigo_producto:
                    existencia = fila[3].value
            
            if cantidad <= existencia:
                lista_compras.item(seleccion[0], values=(lista_compras.item(seleccion[0])["values"][:4] + [cantidad] + [lista_compras.item(seleccion[0])["values"][3] * cantidad] + [datetime.now().strftime("%d/%m/%Y")]))
                
                label_total["text"] = "Total: $" + str(sum([float(lista_compras.item(item)["values"][5]) for item in lista_compras.get_children()]))
            else:
                messagebox.showerror("Error", "No hay suficientes productos en el inventario")
    
    def eliminar_producto():
        seleccion = lista_compras.selection()
        if seleccion:
            lista_compras.delete(seleccion[0])
            
            label_total["text"] = "Total: $" + str(sum([float(lista_compras.item(item)["values"][5]) for item in lista_compras.get_children()]))
    
    def registrar_venta():
        if len(lista_compras.get_children()) == 0:
            messagebox.showerror("Error", "No hay productos en la lista de compras")
        else:
            # Obtener los datos de cada fila de la lista de compras
            for item in lista_compras.get_children():
                ws_ventas.append(lista_compras.item(item)["values"])
                
                datos_venta = lista_compras.item(item)["values"]
                codigo_producto = datos_venta[0]
                cantidad_vendida = datos_venta[4]

                # Buscar la fila correspondiente al producto en el inventario
                for fila in ws_inventario.iter_rows(min_row=2, max_col=4):
                    if fila[0].value == codigo_producto:
                        existencia_actual = fila[3].value
                        nueva_existencia = existencia_actual - cantidad_vendida

                        # Actualizar la existencia del producto en la hoja de inventario
                        ws_inventario.cell(row=fila[0].row, column=4).value = nueva_existencia
                        
            # Guardar el archivo de Excel
            wb.save("SPVI.xlsx")
            messagebox.showinfo("Venta registrada", "La venta se registró correctamente")
            
            lista_compras.delete(*lista_compras.get_children())
            
            label_total["text"] = "0.00"
    
    def cancelar_venta():
        lista_compras.delete(*lista_compras.get_children())
        
        messagebox.showinfo("Venta cancelada", "La venta se canceló exitosamente")
        
        label_total["text"] = "0.00"
    
    ventana_punto_venta = tk.Toplevel(ventana_principal)
    ventana_punto_venta.title("Punto de Venta")
    ventana_punto_venta.geometry("1020x350")
    
    label_productos = tk.Label(ventana_punto_venta, text="Lista de productos")
    label_productos.grid(row=0, column=0, padx=10, pady=10)
    
    lista_compras = ttk.Treeview(ventana_punto_venta, show="headings")
    lista_compras.grid(row=1, column=0, padx=10, rowspan=3, columnspan=4)
    
    btn_agregar_producto = tk.Button(ventana_punto_venta, text="Agregar producto", command=agregar_producto)
    btn_agregar_producto.grid(row=1, column=4, padx=10, pady=10)
    
    btn_modificar_cantidad = tk.Button(ventana_punto_venta, text="Modificar cantidad", command=modificar_cantidad)
    btn_modificar_cantidad.grid(row=2, column=4, padx=10, pady=10)
    
    btn_eliminar_producto = tk.Button(ventana_punto_venta, text="Eliminar producto", command=eliminar_producto)
    btn_eliminar_producto.grid(row=3, column=4, padx=10, pady=10)
    
    btn_registrar_venta = tk.Button(ventana_punto_venta, text="Registrar venta", command=registrar_venta)
    btn_registrar_venta.grid(row=4, column=0, padx=10, pady=10)
    
    btn_cancelar_venta = tk.Button(ventana_punto_venta, text="Cancelar venta", command=cancelar_venta)
    btn_cancelar_venta.grid(row=4, column=1, padx=10, pady=10)
    
    label_total_texto = tk.Label(ventana_punto_venta, text="Total:")
    label_total_texto.grid(row=4, column=2, padx=10, pady=10)
    
    label_total = tk.Label(ventana_punto_venta, text="0.00")
    label_total.grid(row=4, column=3, padx=10, pady=10)
    
    btn_cerrar_punto_venta = tk.Button(ventana_punto_venta, text="Cerrar punto de venta", command=ventana_punto_venta.destroy)
    btn_cerrar_punto_venta.grid(row=4, column=4, padx=10, pady=10)
    
    # Crear menu usuario
    menu_productos = tk.Menu(ventana_punto_venta)
    ventana_punto_venta.config(menu=menu_productos)
    
    # Crear el menú "Acerca de" con sus opciones
    menu_acerca_de = tk.Menu(menu_productos, tearoff=0)
    menu_productos.add_cascade(label="Ayuda", menu=menu_acerca_de)
    menu_acerca_de.add_command(label="Acerca de", command=ventana_acerca_de)
    menu_acerca_de.add_command(label="Instrucciones de uso", command=ventana_instrucciones_de_uso)
    
    datos_originales_ap = []
    datos_filtrados_ap = []
    
    titulos_lista_compras()

# Ventana de Reportes
def ventana_reportes():
    def reporte_mes_anterior():
        ventas_por_producto = {}
        fecha_mes_anterior = hoy.replace(month = hoy.month - 1, day = 1)
        fecha_dia_maximo_del_mes_anterior = hoy.replace(day = 1) - timedelta(days = 1)
        
        for dias in range(1, fecha_dia_maximo_del_mes_anterior.day + 1):
            fecha_mes_anterior = fecha_mes_anterior.replace(day = dias)
            
            for celda in range(2, ws_ventas.max_row + 1):
                if ws_ventas.cell(row=celda, column = 7).value == fecha_mes_anterior.strftime("%d/%m/%Y"):
                    nombre_producto = ws_ventas.cell(row=celda, column = 1).value
                    cantidad = ws_ventas.cell(row=celda, column = 5).value
                    
                    if nombre_producto in ventas_por_producto:
                        ventas_por_producto[nombre_producto] += cantidad
                    else:
                        ventas_por_producto[nombre_producto] = cantidad
                        
        if ventas_por_producto:
            eje_x = list(ventas_por_producto.keys())
            eje_y = list(ventas_por_producto.values())

            plt.bar(eje_x, eje_y)
            plt.xlabel("Producto")
            plt.ylabel("Cantidad")
            plt.title("Reporte Mes Actual")
            
        plt.show()
    def reporte_mes_actual():
        ventas_por_producto = {}
        fecha_mes_actual = hoy.replace(day = 1)
        
        for dias in range(1, hoy.day + 1):
            fecha_mes_actual = fecha_mes_actual.replace(day = dias)
            
            for celda in range(2, ws_ventas.max_row + 1):
                if ws_ventas.cell(row=celda, column = 7).value == fecha_mes_actual.strftime("%d/%m/%Y"):
                    nombre_producto = ws_ventas.cell(row=celda, column = 1).value
                    cantidad = ws_ventas.cell(row=celda, column = 5).value
                    
                    if nombre_producto in ventas_por_producto:
                        ventas_por_producto[nombre_producto] += cantidad
                    else:
                        ventas_por_producto[nombre_producto] = cantidad
        
        if ventas_por_producto:
            eje_x = list(ventas_por_producto.keys())
            eje_y = list(ventas_por_producto.values())

            plt.bar(eje_x, eje_y)
            plt.xlabel("Producto")
            plt.ylabel("Cantidad")
            plt.title("Reporte Mes Actual")

        plt.show()
    def reporte_semanal():
        ventas_por_producto = {}
        
        for dias in range(7,-1,-1):
            fecha_semana_pasada = hoy - timedelta(days=dias)
            
            for celda in range(2, ws_ventas.max_row + 1):
                if ws_ventas.cell(row=celda, column = 7).value == fecha_semana_pasada.strftime("%d/%m/%Y"):
                    nombre_producto = ws_ventas.cell(row = celda, column = 1).value
                    cantidad = ws_ventas.cell(row = celda, column = 5).value
                    
                    if nombre_producto in ventas_por_producto:
                        ventas_por_producto[nombre_producto] += cantidad
                    else:
                        ventas_por_producto[nombre_producto] = cantidad
        
        if ventas_por_producto:
            eje_x = list(ventas_por_producto.keys())
            eje_y = list(ventas_por_producto.values())

            plt.bar(eje_x, eje_y)
            plt.xlabel("Producto")
            plt.ylabel("Cantidad")
            plt.title("Reporte Semanal")

        plt.show()
        
    hoy = datetime.now()
        
    ventana_reportes = tk.Toplevel(ventana_principal)
    ventana_reportes.title("Reportes")
    ventana_reportes.geometry("700x350")
    
    frame_reportes = tk.Frame(ventana_reportes)
    frame_reportes.pack(pady=40)
    
    label_reporte_mes_anterior = tk.Label(frame_reportes, text="Reporte Mes Anterior")
    label_reporte_mes_anterior.grid(row=0, column=0, padx=10, pady=10)
    
    label_reporte_mes_actual = tk.Label(frame_reportes, text="Reporte Mes Actual")
    label_reporte_mes_actual.grid(row=0, column=1, padx=10, pady=10)
    
    label_reporte_semanal = tk.Label(frame_reportes, text="Reporte Semanal")
    label_reporte_semanal.grid(row=0, column=2, padx=10, pady=10)
    
    button_reporte_mes_anterior = tk.Button(frame_reportes, text="📊\nGenerar", command=reporte_mes_anterior, font=font.Font(size=16), width=8, height=3, bg="#B3E5FC")
    button_reporte_mes_anterior.grid(row=1, column=0, padx=10, pady=10)
    
    button_reporte_mes_actual = tk.Button(frame_reportes, text="📊\nGenerar", command=reporte_mes_actual, font=font.Font(size=16), width=8, height=3, bg="#B3E5FC")
    button_reporte_mes_actual.grid(row=1, column=1, padx=10, pady=10)
    
    button_reporte_semanal = tk.Button(frame_reportes, text="📊\nGenerar", command=reporte_semanal, font=font.Font(size=16), width=8, height=3, bg="#B3E5FC")
    button_reporte_semanal.grid(row=1, column=2, padx=10, pady=10)
    
    button_cerrar_reportes = tk.Button(ventana_reportes, text="Cerrar Reportes", command=ventana_reportes.destroy)
    button_cerrar_reportes.pack()

# Ventana Acerca de
def ventana_acerca_de():
    ventana_acerca_de = tk.Toplevel(ventana_principal)
    ventana_acerca_de.geometry("350x220")
    ventana_acerca_de.title("Acerca de")
    
    label_acerca_de = tk.Label(ventana_acerca_de, text="Sistema de Punto de Venta e Inventario")
    label_acerca_de.pack(pady=10)
    
    label_autores = tk.Label(ventana_acerca_de, text="Autores:")
    label_autores.pack(pady=5)
    
    label_orlando = tk.Label(ventana_acerca_de, text="Orlando Antonio Jiménez Esparza")
    label_orlando.pack()
    
    label_lau = tk.Label(ventana_acerca_de, text="Laura Belén Rodríguez Rodríguez")
    label_lau.pack()
    
    label_amir = tk.Label(ventana_acerca_de, text="Amir Gutiérrez Partida")
    label_amir.pack()
    
    label_iteso = tk.Label(ventana_acerca_de, text="ITESO")
    label_iteso.pack(pady=10)
    
    label_materia = tk.Label(ventana_acerca_de, text="Proyecto Integrador de Algoritmos y Programación")
    label_materia.pack()

# Ventana instrucciones de uso
def ventana_instrucciones_de_uso():
    ventana_instrucciones_de_uso = tk.Toplevel(ventana_principal)
    ventana_instrucciones_de_uso.title("Instrucciones de uso")
    ventana_instrucciones_de_uso.geometry("700x450")
    
    instrucciones = """Instrucciones de uso:
    Inicio de sesión:
    Para usar el programa, primero debe iniciar sesión con un usuario y contraseña de administrador brindados por el técnico. Con dicho usuario, podrá acceder a las funciones de administración de usuarios, productos, inventario, punto de venta y reportes. De haber dado de alta a un usuario, este también podrá acceder a las funciones de producots, inventario, punto de venta y reportes, pero no a las de administración.
    
    Alta de productos:
    Para dar de alta un producto, acceder al módulo de productos y dar clic en el botón "Registrar". Se abrirá una ventana en la que deberá ingresar los datos del producto y dar clic en el botón "Registrar". También podrá modificar los datos de un producto o eliminarlo al seleccionarlo y dar click en el botón correspondiente.
    
    Módulo de inventario:
    Una vez registrado el producto, podrá registrar tanto entradas como salidas (mermas) de inventario. Para ello, acceder al módulo de inventario, seleecionar el producto y dar clic en el botón "Entrada de inventario" o "Mermas". Se abrirá una ventana en la que deberá ingresar los datos correspondientes y dar clic en el botón "Ok".
    
    Punta de venta:
    Teniendo productos registrados y existencias en inventario, podrá realizar ventas. Para ello, acceder al módulo de punto de venta, dar click en agregar producto, seleccionar el producto y dar clic en el botón "Agregar". Se abrirá una ventana en la que deberá ingresar la cantidad de productos a vender y dar clic en el botón "Ok". Podrá agregar tantos productos como haya en existencia y podrá modificar la cantidad de productos seleccionándolo y dando click en el botón correspondiente. Para eliminar un producto, seleccionarlo y dar clic en el botón "Eliminar producto". Para registrar la venta, dar clic en el botón "Registrar venta". Para cancelar la venta, dar clic en el botón "Cancelar venta".
    """
    
    label_instrucciones = tk.Label(ventana_instrucciones_de_uso, text=instrucciones,wraplength=600, justify="left")
    label_instrucciones.pack(pady=10)

# Cargar datos de usuario predefinidos y el archivo de Excel
usuarios_registrados = cargar_usuarios()
wb, ws_productos, ws_inventario, ws_ventas = cargar_excel()

# Verificar si el usuario administrador está registrado
if "admin" not in usuarios_registrados:
    usuarios_registrados["admin"] = "admin123"
    guardar_usuarios(usuarios_registrados)  # Guardar los usuarios

# Crear la ventana principal
ventana_principal = tk.Tk()
ventana_principal.title("SPVI")
ventana_principal.geometry("400x300")

# Etiquetas y campos de entrada para el inicio de sesión
label_usuario = tk.Label(ventana_principal, text="Usuario: ")
label_usuario.pack()
entry_usuario = tk.Entry(ventana_principal)
entry_usuario.pack()

label_contraseña = tk.Label(ventana_principal, text="Contraseña:")
label_contraseña.pack()
entry_contraseña = tk.Entry(ventana_principal, show="*")
entry_contraseña.pack()

boton_iniciar_sesion = tk.Button(ventana_principal, text="Iniciar sesión", command=iniciar_sesion)
boton_iniciar_sesion.pack(pady=10)

# MenuStrip
menu_principal = tk.Menu(ventana_principal)
menu_acerca_de = tk.Menu(menu_principal, tearoff=0)
menu_acerca_de.add_command(label="Acerca de", command=ventana_acerca_de)
menu_acerca_de.add_command(label="Instrucciones de uso", command=ventana_instrucciones_de_uso)
menu_principal.add_cascade(label="Ayuda", menu=menu_acerca_de)
ventana_principal.config(menu=menu_principal)

# Ejecutar la ventana principal
ventana_principal.mainloop()