# Alertas Innova  

Este programa permite generar archivos Excel y correos con distintos tipos de avisos.  

Para más información sobre el código, consulta el [manual de desarrollo]().  
Si no tienes acceso, comunícate con **Christopher Vivanco** (christopher.vivanco@corfo.cl).  

---

## Prerrequisitos

Tener instalado python (v3.14)

## Cómo crear el archivo `.exe`  

> **Nota:** Si tienes Python instalado pero no puedes ejecutar comandos, usa el siguiente comando en PowerShell:  

```bash
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### Paso 1: Instalar `pyinstaller`  

```bash
py -m pip install pyinstaller
```

### Paso 2: Crear el `.exe`  

Asegúrate de estar en la misma ubicación que el archivo `app.py` en la terminal y ejecuta:  

```bash
py -m PyInstaller --onefile --windowed app.py
```

### Opcional: Agregar un ícono  

Si deseas que el `.exe` tenga un icono personalizado, usa:  

```bash
py -m PyInstaller --onefile --windowed --icon=corfo.ico app.py
```

---

## Contacto  

Sígueme en [GitHub](https://github.com/loretito)  
Visita mi [página web](https://loretonancucheo.com/) 

Gracias! 😊  
