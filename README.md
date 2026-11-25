# Gestor de Inventario - Python

Un sistema de administración de inventario por consola desarrollado 100% en Python, diseñado como backend para una tienda de productos de computación. Cuenta con validación estricta de datos, persistencia local mediante una base de datos **SQLite** y exportación de reportes estilizados a **Microsoft Excel**.

> **Nota para la evaluación:** El sistema funciona de manera completamente autónoma y local. Al ejecutarse por primera vez, inicializa automáticamente la base de datos `inventario.db` y su tabla `productos` sin requerir credenciales externas ni configuraciones adicionales.

## Características Principales

- **Persistencia Local en SQLite:** Almacenamiento relacional en `inventario.db` con clave primaria autoincremental (`id`) y esquema estructurado (`nombre`, `descripcion`, `cantidad`, `precio`, `categoria`).
- **Interfaz de Consola Mejorada:** Menú interactivo con colores diferenciados mediante `colorama` y visualización de catálogo en tablas dinámicas centradas y alineadas automáticamente.
- **Validación Estricta de Datos:**
  - **Nombres:** Alfanuméricos, bloquea cadenas vacías o puramente numéricas.
  - **Categorías:** Selección numerada sobre un listado cerrado de rubros de computación (evita errores ortográficos y duplicados).
  - **Cantidades y Precios:** Limpieza automática de símbolos (`$`, `%`, `&`, `!`) y validación estricta de valores numéricos positivos.
- **Sistema CRUD, Reportes y Exportación:**
  - `[1]` Registrar nuevo producto
  - `[2]` Visualizar productos registrados
  - `[3]` Actualizar producto mediante su ID (permite conservar valores actuales presionando `Enter`)
  - `[4]` Eliminar producto mediante su ID (incluye barrera de confirmación `s/n` para evitar borrados accidentales)
  - `[5]` Buscar producto (por ID, nombre o categoría)
  - `[6]` Reporte de stock bajo (productos con cantidad igual o inferior a un límite definido)
  - `[7]` Exportar inventario a Excel (`.xlsx`) con encabezados formateados, bordes, filtros y formato moneda
  - `[8]` Salir

## Tecnologías Utilizadas

- **Python 3**
- **sqlite3**: Módulo estándar de Python para la gestión y consulta de la base de datos local (`inventario.db`).
- **colorama**: Para mejorar la legibilidad y experiencia visual en la terminal.
- **openpyxl**: Para la generación de reportes en formato Excel (`.xlsx`) con estilos, colores y formato numérico.

## Instrucciones para probar el proyecto

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/agus25varela/Pre-entrega-de-proyecto.git](https://github.com/agus25varela/Pre-entrega-de-proyecto.git)
   cd Pre-entrega-de-proyecto
   ```

2. **Instalar dependencias necesarias:**
   ```bash
   pip install colorama openpyxl
   ```

3. **Ejecutar el programa:**
   ```bash
   python gestor_inventario.py
   ```

4. **Visualizar productos en consola:**
   ```bash
   Con el menú principal activo, ingresá la opción 2 para listar en la terminal todos los registros almacenados actualmente en inventario.db.
   ```

5. **Generar el reporte en Excel (.xlsx):**
   ```bash
   Seleccioná la opción 7 (Exportar inventario a Excel (.xlsx)) en el menú principal. El sistema creará o actualizará automáticamente el archivo inventario_exportado.xlsx en la carpeta raíz del proyecto con encabezados formateados, bordes, filtros y formato de moneda.
   ```

6. **Ubicar y abrir el archivo Excel generado:**
   ```bash
   El archivo inventario_exportado.xlsx queda guardado dentro de la carpeta del proyecto. Si estás trabajando desde Visual Studio Code, podés ubicarlo en el panel lateral izquierdo (Explorador), hacer clic derecho sobre inventario_exportado.xlsx y seleccionar Reveal in File Explorer (Revelar en el Explorador de archivos) para abrirlo directamente con Microsoft Excel, LibreOffice Calc o cualquier visor de planillas.
   ```