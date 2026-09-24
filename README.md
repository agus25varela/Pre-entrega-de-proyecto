# Gestor de Inventario - MyShop

Un sistema de administración de inventario por consola desarrollado 100% en Python, diseñado como backend para un e-commerce. Cuenta con validación estricta de datos y sincronización en tiempo real con **Google Sheets** como base de datos en la nube.

> **Nota para la evaluación:** Este repositorio incluye el archivo `credenciales.json` configurado para facilitar la prueba del proyecto sin requerir configuración adicional de la API. 

## Características Principales

- **Sincronización en la Nube:** Lectura y escritura bidireccional con Google Sheets mediante su API oficial.
- **Formato Visual Preservado:** Actualiza los datos sin romper los estilos, colores o formatos de moneda configurados en el panel web (`USER_ENTERED`).
- **Validación Estricta:**
  - Nombres: Alfanuméricos, bloquea cadenas vacías o puramente numéricas.
  - Categorías: Exclusivamente caracteres alfabéticos.
  - Precios: Limpieza automática de símbolos (`$`, `%`, `&`, `!`) y validación estricta de números enteros (sin centavos).
- **Sistema CRUD Completo:**
  - `[1]` Agregar producto
  - `[2]` Visualizar catálogo
  - `[3]` Buscar producto por nombre
  - `[4]` Eliminar producto (incluye barrera de confirmación `S/N` para evitar borrados accidentales)

## Tecnologías Utilizadas

- **Python** 
- **gspread**: Para la manipulación de celdas en la hoja de cálculo.
- **google-auth**: Para la autenticación segura mediante Cuentas de Servicio.
- **Google Cloud Platform (GCP)**: API integrada de Google Sheets.

## Instrucciones para probar el proyecto

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/agus25varela/Pre-entrega-de-proyecto.git]
   cd Pre-entrega-de-proyecto