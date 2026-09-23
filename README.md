# Formulario de Adicionales y Retornos

Aplicación de escritorio desarrollada con Python y `customtkinter` para el manejo, validación y envío de reportes de Scrap y Adicionales (entorno Wasion).

## Características Principales
- Interfaz gráfica moderna (Wasion Blue Theme).
- Procesamiento y manipulación local de archivos Excel (`.xlsm`) manteniendo macros y formato.
- Integración con Google Sheets para el volcado de datos.
- Sistema de autenticación interno para acceso a configuración.
- Ocultamiento inteligente de listas desplegables.

## Requisitos del Sistema
- Python 3.8+ (Recomendado)
- Credenciales válidas para la API de Google Sheets (`credentials.json`), el cual **no se incluye** en este repositorio por motivos de seguridad.

## Instalación y Configuración

1. **Instalar dependencias:**
   Ejecuta el siguiente comando en la raíz del proyecto para instalar las librerías necesarias:
   ```bash
   pip install -r requirements.txt
   ```

2. **Configuración de Google Sheets:**
   Solicita el archivo `credentials.json` al administrador del proyecto y colócalo en la raíz del directorio antes de iniciar la aplicación.

## Ejecución
Para lanzar la aplicación, puedes ejecutar el script principal:
```bash
python app.py
```
O bien, utilizar el acceso directo/batch provisto:
```bash
Iniciar_Formulario.bat
```

## Compilación a Ejecutable (.exe)
El proyecto incluye archivos de configuración para ser empaquetado usando PyInstaller. El archivo `.spec` (`Formulario Adicionales.spec`) contiene las directivas necesarias para incluir los assets e íconos.

## Licencia
Código **Propietario y Privado**. Todos los derechos reservados. No se permite la copia ni distribución sin autorización.
