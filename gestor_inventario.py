import csv
import sqlite3

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from colorama import init, Fore, Style
init(autoreset=True)

DB_NAME = "inventario.db"

CATEGORIAS_DISPONIBLES = (
    "Componentes de PC",
    "Periféricos",
    "Monitores",
    "Notebooks y Laptops",
    "Almacenamiento",
    "Redes y Conectividad",
    "Audio y Streaming",
    "Cables y Adaptadores",
    "Accesorios y Mobiliario",
)


def conectar_db() -> sqlite3.Connection:
    """Crea y devuelve una conexión a la base de datos SQLite."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_bd() -> None:
    """Crea la tabla 'productos' si no existe con la estructura requerida."""
    with conectar_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS productos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                descripcion TEXT,
                cantidad INTEGER NOT NULL,
                precio REAL NOT NULL,
                categoria TEXT
            )
            """
        )
        conn.commit()

def pedir_nombre(valor_actual: str | None = None) -> str:
    """Solicita y valida el nombre del producto (no vacío, no puramente numérico)."""
    while True:
        aviso = f" [{valor_actual}]" if valor_actual is not None else ""
        nombre = input(f"Ingresá el nombre del producto{aviso}: ").strip()
        if nombre == "" and valor_actual is not None:
            return valor_actual
        if nombre == "":
            print(f"{Fore.RED} Error: El nombre no puede estar vacío.{Style.RESET_ALL}")
        elif nombre.isdigit():
            print(f"{Fore.RED} Error: El nombre no puede estar compuesto solo por números.{Style.RESET_ALL}")
        else:
            return nombre


def pedir_descripcion(valor_actual: str | None = None) -> str:
    """Solicita la descripción del producto."""
    aviso = f" [{valor_actual}]" if valor_actual is not None else ""
    desc = input(f"Ingresá una breve descripción{aviso}: ").strip()
    if desc == "" and valor_actual is not None:
        return valor_actual
    return desc


def pedir_cantidad(valor_actual: int | None = None, permitir_cero: bool = True) -> int:
    """Solicita y valida una cantidad entera."""
    while True:
        aviso = f" [{valor_actual}]" if valor_actual is not None else ""
        cant_str = input(f"Ingresá la cantidad disponible{aviso}: ").strip()
        if cant_str == "" and valor_actual is not None:
            return valor_actual
        if cant_str.isdigit():
            cantidad = int(cant_str)
            if permitir_cero and cantidad >= 0:
                return cantidad
            if not permitir_cero and cantidad > 0:
                return cantidad
        print(f"{Fore.RED} Error: Ingresá un número entero válido.{Style.RESET_ALL}")


def pedir_precio(valor_actual: float | None = None) -> float:
    """Limpia símbolos comunes y valida un precio entero mayor a 0 (guardado como REAL)."""
    while True:
        aviso = f" [${int(valor_actual)}]" if valor_actual is not None else ""
        precio_bruto = input(f"Ingresá el precio sin centavos{aviso}: ").strip()
        if precio_bruto == "" and valor_actual is not None:
            return float(valor_actual)

        for simbolo in "$%&! ":
            precio_bruto = precio_bruto.replace(simbolo, "")

        if precio_bruto.isdigit() and int(precio_bruto) > 0:
            return float(int(precio_bruto))
        print(f"{Fore.RED} Error: Ingresá un valor numérico entero mayor a 0 (ej: $1500 o 1500).{Style.RESET_ALL}")


def pedir_categoria(valor_actual: str | None = None) -> str:
    """Muestra una lista numerada de categorías y devuelve la elegida sin errores de tipeo."""
    print(f"\n{Fore.CYAN}Seleccioná una categoría:")
    for indice, cat in enumerate(CATEGORIAS_DISPONIBLES, start=1):
        print(f"  {Fore.WHITE}{indice}.{Style.RESET_ALL} {cat}")
        print(f"{Fore.CYAN}" + "=" * 45)

    total = len(CATEGORIAS_DISPONIBLES)
    aviso = f" [Enter para mantener '{valor_actual}']" if valor_actual is not None else ""

    while True:
        opcion = input(f"Elegí el número de categoría (1-{total}){aviso}: ").strip()
        if opcion == "" and valor_actual is not None:
            return valor_actual
        if opcion.isdigit():
            num = int(opcion)
            if 1 <= num <= total:
                return CATEGORIAS_DISPONIBLES[num - 1]
        print(f"{Fore.RED} Error: Ingresá un número válido entre 1 y {total}.")


def pedir_id(mensaje: str = "Ingresá el ID del producto: ") -> int | None:
    """Solicita un ID numérico entero positivo o permite cancelar con 0."""
    while True:
        id_str = input(f"{mensaje}(o '0' para cancelar): ").strip()
        if id_str == "0":
            return None
        if id_str.isdigit() and int(id_str) > 0:
            return int(id_str)
        print(f"{Fore.RED} Error: Ingresá un ID numérico válido mayor a 0.{Style.RESET_ALL}")


def mostrar_tabla_productos(filas: list[sqlite3.Row]) -> None:
    """Imprime los productos centrados y alineados en consola sin importar el largo."""
    if not filas:
        print(f"{Fore.YELLOW}No hay productos para mostrar.")
        return

    encabezados = ["ID", "Nombre", "Categoría", "Cant.", "Precio", "Descripción"]
    maximos = [len(h) for h in encabezados]
    anchos_minimos = [4, 14, 16, 7, 10, 16]
    anchos_maximos = [8, 28, 26, 10, 14, 36]

    registros = []
    for fila in filas:
        precio_fmt = (
            f"${int(fila['precio'])}"
            if float(fila["precio"]).is_integer()
            else f"${fila['precio']:.2f}"
        )
        valores = [
            str(fila["id"]),
            str(fila["nombre"] or ""),
            str(fila["categoria"] or ""),
            str(fila["cantidad"]),
            precio_fmt,
            str(fila["descripcion"] or ""),
        ]
        registros.append(valores)

        for i, val in enumerate(valores):
            largo_util = min(len(val), anchos_maximos[i])
            if largo_util > maximos[i]:
                maximos[i] = largo_util

    anchos = [
        max(anchos_minimos[i], min(maximos[i] + 2, anchos_maximos[i]))
        for i in range(len(encabezados))
    ]

    linea_encabezado = " | ".join(f"{encabezados[i]:^{anchos[i]}}" for i in range(len(encabezados)))
    separador = "=" * len(linea_encabezado)

    print(f"\n{Fore.CYAN}{Style.BRIGHT}{separador}")
    print(linea_encabezado)
    print(f"{separador}{Style.RESET_ALL}")

    for reg in registros:
        fila_formateada = " | ".join(
            f"{reg[i][:anchos[i]]:^{anchos[i]}}" for i in range(len(reg))
        )
        print(fila_formateada)

    print(f"{Fore.CYAN}{separador}{Style.RESET_ALL}")


def registrar_producto() -> None:
    """Registra un nuevo producto en la tabla 'productos'."""
    print("\n--- Registrar Nuevo Producto ---")
    nombre = pedir_nombre()
    descripcion = pedir_descripcion()
    cantidad = pedir_cantidad(permitir_cero=True)
    precio = pedir_precio()
    categoria = pedir_categoria()

    with conectar_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO productos (nombre, descripcion, cantidad, precio, categoria)
            VALUES (?, ?, ?, ?, ?)
            """,
            (nombre, descripcion, cantidad, precio, categoria),
        )
        conn.commit()
        nuevo_id = cursor.lastrowid

    print(f"{Fore.GREEN} ¡Producto '{nombre}' registrado con éxito bajo el ID #{nuevo_id}!")


def visualizar_productos() -> None:
    """Muestra todos los productos registrados en la base de datos."""
    print("\n--- Catálogo de Productos ---")
    with conectar_db() as conn:
        filas = conn.execute("SELECT * FROM productos ORDER BY id ASC").fetchall()
    mostrar_tabla_productos(filas)

def actualizar_producto() -> None:
    """Actualiza los datos de un producto existente buscándolo por su ID."""
    print("\n--- Actualizar Producto por ID ---")
    prod_id = pedir_id("Ingresá el ID del producto a actualizar ")
    if prod_id is None:
        print("Operación cancelada.")
        return

    with conectar_db() as conn:
        producto = conn.execute("SELECT * FROM productos WHERE id = ?", (prod_id,)).fetchone()
        if not producto:
            print(f"{Fore.RED} No existe ningún producto con el ID #{prod_id}.{Style.RESET_ALL}")
            return

        print(f"\nEditando ID #{producto['id']} (dejá vacío para mantener el valor actual):")
        nuevo_nombre = pedir_nombre(producto["nombre"])
        nueva_desc = pedir_descripcion(producto["descripcion"] or "")
        nueva_cant = pedir_cantidad(producto["cantidad"], permitir_cero=True)
        nuevo_precio = pedir_precio(producto["precio"])
        nueva_cat = pedir_categoria(producto["categoria"] or "")

        conn.execute(
            """
            UPDATE productos
            SET nombre = ?, descripcion = ?, cantidad = ?, precio = ?, categoria = ?
            WHERE id = ?
            """,
            (nuevo_nombre, nueva_desc, nueva_cant, nuevo_precio, nueva_cat, prod_id),
        )
        conn.commit()

    print(f"{Fore.GREEN} Producto ID #{prod_id} actualizado correctamente.")


def eliminar_producto() -> None:
    """Elimina un producto por su ID previa confirmación del usuario."""
    print("\n--- Eliminar Producto por ID ---")
    prod_id = pedir_id("Ingresá el ID del producto a eliminar ")
    if prod_id is None:
        print("Operación cancelada.")
        return

    with conectar_db() as conn:
        producto = conn.execute("SELECT * FROM productos WHERE id = ?", (prod_id,)).fetchone()
        if not producto:
            print(f"{Fore.RED} No se encontró ningún producto con el ID #{prod_id}.")
            return

        confirmacion = input(
            f"{Fore.YELLOW} ¿Estás seguro de eliminar '{producto['nombre']}' (ID #{producto['id']})? (s/n): "
        ).strip().lower()

        if confirmacion == "s":
            conn.execute("DELETE FROM productos WHERE id = ?", (prod_id,))
            conn.commit()
            print(f"{Fore.GREEN} Producto ID #{prod_id} eliminado de '{DB_NAME}'.")
        else:
            print(f"{Fore.RED} Eliminación cancelada.{Style.RESET_ALL}")


def buscar_producto() -> None:
    """Busca productos por ID (obligatorio) o por nombre/categoría (opcional)."""
    print("\n--- Buscar Producto ---")
    print("1. Buscar por ID")
    print("2. Buscar por nombre")
    print("3. Buscar por categoría")
    sub = input("Elegí un criterio (1-3): ").strip()

    with conectar_db() as conn:
        if sub == "1":
            prod_id = pedir_id("Ingresá el ID a buscar ")
            if prod_id is None:
                print(f"{Fore.RED} Búsqueda cancelada.{Style.RESET_ALL}")
                return
            filas = conn.execute("SELECT * FROM productos WHERE id = ?", (prod_id,)).fetchall()
        elif sub == "2":
            termino = input("Ingresá el nombre o parte del nombre: ").strip()
            filas = conn.execute(
                "SELECT * FROM productos WHERE LOWER(nombre) LIKE LOWER(?) ORDER BY id ASC",
                (f"%{termino}%",),
            ).fetchall()
        elif sub == "3":
            termino = input("Ingresá la categoría: ").strip()
            filas = conn.execute(
                "SELECT * FROM productos WHERE LOWER(categoria) LIKE LOWER(?) ORDER BY id ASC",
                (f"%{termino}%",),
            ).fetchall()
        else:
            print(f"{Fore.RED} Opción no válida.{Style.RESET_ALL}")
            return

    mostrar_tabla_productos(filas)


def reporte_bajo_stock() -> None:
    """Muestra productos con cantidad igual o inferior al límite indicado."""
    print("\n--- Reporte de Stock Bajo o Igual al Límite ---")
    limite = pedir_cantidad(permitir_cero=True)

    with conectar_db() as conn:
        filas = conn.execute(
            "SELECT * FROM productos WHERE cantidad <= ? ORDER BY cantidad ASC, id ASC",
            (limite,),
        ).fetchall()

    if not filas:
        print(f"{Fore.GREEN} No hay productos con cantidad igual o inferior a {limite}.{Style.RESET_ALL}")
    else:
        print(f"\n{Fore.YELLOW} Productos con cantidad <= {limite}:{Style.RESET_ALL}")
        mostrar_tabla_productos(filas)

def exportar_a_excel(nombre_archivo: str = "inventario_exportado.xlsx") -> None:
    """Exporta la tabla 'productos' a un archivo Excel (.xlsx) con estilos y colores."""
    with conectar_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, nombre, descripcion, cantidad, precio, categoria FROM productos ORDER BY id ASC"
        )
        filas = cursor.fetchall()

    if not filas:
        print(f"{Fore.YELLOW} No hay productos para exportar.{Style.RESET_ALL}")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"

    # Estilos reutilizables
    fuente_encabezado = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    relleno_encabezado = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    relleno_par = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    relleno_impar = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fuente_cuerpo = Font(name="Calibri", size=11, color="000000")

    borde_fino = Side(border_style="thin", color="D9D9D9")
    bordes_celda = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)

    encabezados = ["ID", "Nombre", "Descripción", "Cantidad", "Precio", "Categoría"]
    ws.append(encabezados)

    # Formatear fila de encabezados
    for col_idx, celda in enumerate(ws[1], start=1):
        celda.font = fuente_encabezado
        celda.fill = relleno_encabezado
        celda.border = bordes_celda
        celda.alignment = Alignment(horizontal="center" if col_idx in (1, 4, 5) else "left", vertical="center")
    ws.row_dimensions[1].height = 24

    # Cargar datos y aplicar estilos por fila
    for idx_fila, fila in enumerate(filas, start=2):
        ws.append(list(fila))
        relleno_fila = relleno_par if idx_fila % 2 == 0 else relleno_impar
        ws.row_dimensions[idx_fila].height = 20

        for col_idx in range(1, 7):
            celda = ws.cell(row=idx_fila, column=col_idx)
            celda.font = fuente_cuerpo
            celda.fill = relleno_fila
            celda.border = bordes_celda

            if col_idx in (1, 4):
                celda.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 5:
                celda.alignment = Alignment(horizontal="right", vertical="center")
                celda.number_format = "$ #,##0.00"
            else:
                celda.alignment = Alignment(horizontal="left", vertical="center")

    # Ajustar ancho de columnas automáticamente
    anchos = {
        "A": 8,
        "B": 26,
        "C": 34,
        "D": 12,
        "E": 15,
        "F": 18,
    }
    for letra_col, ancho in anchos.items():
        ws.column_dimensions[letra_col].width = ancho

    # Congelar encabezados y activar filtros en Excel
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{len(filas) + 1}"

    wb.save(nombre_archivo)
    print(f"{Fore.GREEN} Inventario exportado con estilos a '{nombre_archivo}'.")

def mostrar_menu() -> None:
    """Muestra el menú principal en la terminal."""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}" + "=" * 45)
    print("GESTOR DE INVENTARIO")
    print("=" * 45 + f"{Style.RESET_ALL}")
    print(f"{Fore.WHITE}1.{Style.RESET_ALL} Registrar nuevo producto")
    print(f"{Fore.WHITE}2.{Style.RESET_ALL} Visualizar productos registrados")
    print(f"{Fore.WHITE}3.{Style.RESET_ALL} Actualizar producto (por ID)")
    print(f"{Fore.WHITE}4.{Style.RESET_ALL} Eliminar producto (por ID)")
    print(f"{Fore.WHITE}5.{Style.RESET_ALL} Buscar producto (por ID / nombre / categoría)")
    print(f"{Fore.WHITE}6.{Style.RESET_ALL} Reporte de stock (cantidad <= límite)")
    print(f"{Fore.WHITE}7.{Style.RESET_ALL} Exportar inventario a Excel (.xlsx)")
    print(f"{Fore.WHITE}8.{Style.RESET_ALL} Salir")
    print(f"{Fore.CYAN}{Style.BRIGHT}" + "=" * 45)


def main() -> None:
    """Punto de entrada principal de la aplicación."""
    inicializar_bd()
    print(f"{Fore.GREEN} Base de datos '{DB_NAME}' lista y conectada.")

    while True:
        mostrar_menu()
        opcion = input("Seleccioná una opción (1-8): ").strip()

        if opcion == "1":
            registrar_producto()
        elif opcion == "2":
            visualizar_productos()
        elif opcion == "3":
            actualizar_producto()
        elif opcion == "4":
            eliminar_producto()
        elif opcion == "5":
            buscar_producto()
        elif opcion == "6":
            reporte_bajo_stock()
        elif opcion == "7":
            exportar_a_excel()
        elif opcion == "8":
            print("\nSaliendo del sistema...")
            break
        else:
            print(f"{Fore.RED} Opción inválida. Ingresá un número del 1 al 8.{Style.RESET_ALL}")


if __name__ == "__main__":
    main()