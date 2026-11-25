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