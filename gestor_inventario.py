import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. CONFIGURACIÓN DE CONEXIÓN A GOOGLE SHEETS
# ==========================================
print("⏳ Conectando con Google Sheets...")

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# autentica con el archivo JSON
credenciales = Credentials.from_service_account_file("credenciales.json", scopes=scopes)
cliente = gspread.authorize(credenciales)

# Mi hoja dee google sheet
hoja = cliente.open("Inventario-Python").sheet1

# ==========================================
# 2. FUNCIONES DE SINCRONIZACIÓN
# ==========================================
def cargar_desde_nube():
    try:
        datos = hoja.get_all_values()
        inventario_local = []
        
        if len(datos) <= 1 and (len(datos) == 0 or "Hola desde Python" in datos[0][0]):
            hoja.batch_clear(["A:C"])
            hoja.append_row(["Nombre", "Categoría", "Precio"])
            return []

        for fila in datos[1:]: 
            # Limpia los símbolos del formato moneda
            precio_limpio = str(fila[2]).replace("$", "").replace(".", "").replace(",", "").strip()
            inventario_local.append([fila[0], fila[1], int(precio_limpio)])
        return inventario_local
    except Exception as e:
        print(f"⚠️ Aviso: No se pudieron cargar datos previos. Empezando de cero. ({e})")
        hoja.batch_clear(["A:C"])
        hoja.append_row(["Nombre", "Categoría", "Precio"])
        return []

def guardar_en_nube(inventario_local):
    hoja.batch_clear(["A:C"])
    datos_a_guardar = [["Nombre", "Categoría", "Precio"]]
    
    for prod in inventario_local:
        # Envia prod[0], prod[1], prod[2] a la hoja de cálculo
        datos_a_guardar.append([prod[0], prod[1], prod[2]])
        
    # USER_ENTERED fuerza a Sheets a aplicar los formatos de la columna
    hoja.update(datos_a_guardar, value_input_option="USER_ENTERED")

# ==========================================
# 3. Menú principal del gestor de inventario
# ==========================================
# Carga los datos al arrancar
inventario = cargar_desde_nube()
print("✅ ¡Conexión exitosa! Sistema listo para usar.")

while True:
    print("\n" + "="*35)
    print(" GESTOR DE INVENTARIO E-COMMERCE ")
    print("="*35)
    print("1. Agregar producto")
    print("2. Visualizar productos")
    print("3. Buscar producto")
    print("4. Eliminar producto")
    print("5. Salir")
    print("="*35)
    
    opcion = input("Seleccioná una opción (1-5): ")

    # 1. AGREGAR PRODUCTO
    if opcion == '1':
        print("\n--- Agregar Nuevo Producto ---")

        # Validación del nombre: no vacío y no numérico
        while True:
            nombre = input("Ingresá el nombre del producto: ").strip()
            if nombre == "":
                print("❌ Error: El nombre no puede estar vacío.")
            elif nombre.isdigit():
                print("❌ Error: El nombre no puede estar compuesto solo por números.")
            else:
                break
                
        # Validación de la categoría: solo letras (y espacios)
        while True:
            categoria = input("Ingresá la categoría: ").strip()
            if categoria == "":
                print("❌ Error: La categoría no puede estar vacía.")
            # .replace(" ", "") permite que acepte categorías de dos palabras como "Ropa deportiva"
            elif not categoria.replace(" ", "").isalpha():
                print("❌ Error: La categoría solo debe contener letras.")
            else:
                break
                
        # Validación del precio: limpiar símbolos y verificar número entero
        while True:
            precio_bruto = input("Ingresá el precio (sin centavos): ").strip()
            
            # Limpiamos los símbolos y espacios
            for simbolo in "$%&! ":
                precio_bruto = precio_bruto.replace(simbolo, "")
                
            # Si después de limpiar quedó un número entero válido
            if precio_bruto.isdigit() and int(precio_bruto) > 0:
                precio = int(precio_bruto)
                break
            else:
                print("❌ Error: Ingresá un precio numérico entero mayor a 0 (por ejemplo: $1500).")
        
            
        # Agrega a la lista local
        producto = [nombre, categoria, precio]
        inventario.append(producto)
        
        # Sincroniza con Sheets
        guardar_en_nube(inventario)
        print(f"✅ ¡'{nombre}' agregado al sistema y guardado en la nube!")

    # 2. VISUALIZAR PRODUCTOS
    elif opcion == '2':
        print("\n--- Lista de Productos ---")
        if len(inventario) == 0:
            print("⚠️ El inventario está vacío.")
        else:
            print(f"{'Nº':<4} | {'Nombre':<20} | {'Categoría':<15} | {'Precio'}")
            print("-" * 55)
            for i in range(len(inventario)):
                prod = inventario[i]
                print(f"{i + 1:<4} | {prod[0]:<20} | {prod[1]:<15} | ${prod[2]}")

    # 3. BUSCAR PRODUCTO
    elif opcion == '3':
        print("\n--- Buscar Producto ---")
        if len(inventario) == 0:
            print("⚠️ El inventario está vacío.")
        else:
            termino = input("Ingresá el nombre a buscar: ").strip().lower()
            encontrados = False
            
            print(f"\nResultados para '{termino}':")
            for prod in inventario:
                if termino in prod[0].lower():
                    print(f"📌 Nombre: {prod[0]} | Categoría: {prod[1]} | Precio: ${prod[2]}")
                    encontrados = True
            
            if not encontrados:
                print("❌ No se encontraron resultados que coincidan con la búsqueda.")

    # 4. ELIMINAR PRODUCTO
    elif opcion == '4':
        print("\n--- Eliminar Producto ---")
        if len(inventario) == 0:
            print("⚠️ El inventario está vacío.")
        else:
            for i in range(len(inventario)):
                print(f"{i + 1}. {inventario[i][0]} (${inventario[i][2]})")
                
            while True:
                pos_str = input("\nIngresá el número del producto a eliminar (o '0' para cancelar): ").strip()
                if pos_str.isdigit():
                    pos = int(pos_str)
                    if pos == 0:
                        print("Operación cancelada.")
                        break
                    elif 1 <= pos <= len(inventario):
                        # Le recuerda el user el producto antes de borrarlo
                        producto_a_eliminar = inventario[pos - 1]
                        
                        # Pide confirmación
                        confirmacion = input(f"⚠️ ¿Estás seguro de que querés eliminar '{producto_a_eliminar[0]}'? (s/n): ").strip().lower()
                        
                        if confirmacion == 's':
                            eliminado = inventario.pop(pos - 1)
                            # Sincronizamos con Sheets solo si confirmó
                            guardar_en_nube(inventario)
                            print(f"✅ El producto '{eliminado[0]}' ha sido eliminado de la base de datos.")
                        else:
                            print("❌ Eliminación cancelada. El producto sigue en el inventario.")
                        break
                    else:
                        print("❌ Error: El número ingresado no corresponde a ningún producto.")
                else:
                    print("❌ Error: Debes ingresar un número.")

    # 5. SALIR
    elif opcion == '5':
        print("\nSaliendo del sistema... ¡Éxitos con el proyecto!")
        break

    else:
        print("\n❌ Opción no válida. Por favor ingresá un número del 1 al 5.")