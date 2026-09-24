import gspread
from google.oauth2.service_account import Credentials

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# 1. Autenticación
credenciales = Credentials.from_service_account_file("credenciales.json", scopes=scopes)
cliente = gspread.authorize(credenciales)

# 2. Conexión
try:
    print("⏳ Conectando a Google Sheets...")
    hoja = cliente.open("Inventario-Python").sheet1
    
    # 3. Mensaje de prueba
    hoja.update_acell('A1', '¡Hola desde Python! La conexión funciona perfecto. ✅')
    
    print("✅ ¡Conexión exitosa! Abrí tu Google Sheet en el navegador y mirá la celda A1.")
    
except Exception as e:
    print("❌ Hubo un error al intentar conectar:")
    print(e)