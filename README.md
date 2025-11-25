Perfecto, con esa descripción ya se entiende bien el objetivo de la herramienta.
Te propongo este `README.md` en español, adaptado a lo que hace realmente tu exe (Azure DevOps → Excel → Word por caso de prueba):

````markdown
# Generador de Evidencias de Casos de Prueba (Azure DevOps)

Aplicación de escritorio en Python que toma *test suites* exportadas desde Azure DevOps en formato Excel (`.xlsx`) y genera automáticamente documentos de Word (`.docx`) de evidencias para cada caso de prueba.

Cada suite de prueba se convierte en una carpeta de salida, y dentro se crea un archivo de Word por caso de prueba.

> La interfaz gráfica se distribuye normalmente como un ejecutable (`.exe`) construido a partir de `gui_launcher.py`.

---

## Funcionalidad

1. El usuario exporta desde Azure DevOps uno o varios test suites a archivos Excel (`.xlsx`).
2. La aplicación lee esos archivos desde la **Carpeta origen**.
3. Para cada suite:
   - Crea una carpeta con el nombre del suite (o similar) dentro de la **Carpeta destino**.
   - Genera un documento de Word por cada caso de prueba contenido en el Excel.
4. Cada documento incluye, por ejemplo:
   - Datos del proyecto.
   - Nombre del analista QA.
   - Fecha.
   - Imágenes de encabezado y pie personalizables.

---

## Interfaz gráfica

La ventana principal (ejecutable generado a partir de `gui_launcher.py`) muestra los siguientes campos:

- **Carpeta origen**  
  Carpeta donde se encuentran los archivos Excel exportados desde Azure DevOps.  
  Cada archivo suele corresponder a un test suite.

- **Carpeta destino**  
  Carpeta raíz donde se crearán las subcarpetas de salida.  
  Para cada test suite se crea una carpeta y dentro se generan los documentos de Word de cada caso de prueba.

- **Proyecto**  
  Nombre del proyecto. Se utiliza como metadato en los documentos generados (por ejemplo, en el encabezado o en el cuerpo del Word).

- **Analista QA**  
  Nombre del analista de calidad responsable de los casos de prueba. También se inserta en los documentos generados.

- **Fecha**  
  Fecha que se va a mostrar en las evidencias.  
  Habitualmente se rellena con la fecha de ejecución de las pruebas.

- **Imagen encabezado**  
  Ruta a la imagen (por ejemplo `header.png`) que se inserta como encabezado en los documentos de Word.

- **Imagen pie**  
  Ruta a la imagen (por ejemplo `footer.png`) que se inserta como pie de página en los documentos de Word.

- **Botón “Generar”**  
  Inicia el proceso de lectura de los Excel, creación de carpetas y generación de documentos de Word.

---

## Flujo de trabajo típico

1. **Exportar suites desde Azure DevOps**
   - Desde Azure DevOps, exportar los test suites que se quieren documentar a formato Excel.
   - Guardar esos archivos en una carpeta (por ejemplo `C:\Evidencias\Entrada`).

2. **Configurar la herramienta**
   - **Carpeta origen**: seleccionar `C:\Evidencias\Entrada`.
   - **Carpeta destino**: seleccionar, por ejemplo, `C:\Evidencias\Salida`.
   - Rellenar **Proyecto**, **Analista QA** y **Fecha**.
   - Seleccionar las imágenes de `header.png` y `footer.png` (se suelen distribuir junto con el exe).

3. **Generar evidencias**
   - Pulsar **Generar**.
   - El programa leerá todos los `.xlsx` de la carpeta origen y creará:
     - Una carpeta por test suite en la carpeta destino.
     - Un `.docx` por caso de prueba dentro de la carpeta correspondiente.

4. **Revisión**
   - Abrir la carpeta destino y revisar los documentos de Word generados.
   - Compartirlos con el equipo de proyecto, auditores, etc.

---

## Estructura del repositorio

Código principal del proyecto:

- `main.py`  
  Punto de entrada desde consola (si se necesita ejecutar sin interfaz gráfica).

- `gui_launcher.py`  
  Lanza la interfaz gráfica desde código fuente. Es la base para construir el ejecutable.

- `gui_launcher.spec`  
  Archivo de configuración para herramientas de empaquetado (por ejemplo, PyInstaller) con el que se genera el `.exe`.

- `generator_core.py`  
  Lógica central de generación: conecta la carga de datos, el procesamiento y la construcción de los documentos.

- `data_loader.py`  
  Funciones para leer los archivos Excel exportados desde Azure DevOps.

- `processor.py`  
  Transformación y preparación de los datos de prueba antes de generar los documentos.

- `document_generator.py`  
  Creación de los documentos de Word a partir de los datos ya procesados y las plantillas (incluyendo encabezado y pie).

---

## Requisitos

- Python 3.10 o superior.
- `pip` instalado.

El proyecto utiliza librerías para:
- Leer y manipular archivos Excel.
- Generar documentos de Word (`.docx`).
- Construir la interfaz gráfica.

Para conocer la lista exacta de dependencias, revisar los `import` en los módulos (`data_loader.py`, `document_generator.py`, etc.) y añadirlas a un `requirements.txt`.

Ejemplo de instalación de dependencias (una vez identificado lo necesario):

```bash
pip install <paquete1> <paquete2> ...
````

---

## Ejecución desde código fuente

1. Clonar el repositorio:

   ```bash
   git clone https://github.com/Bemuh/testDocGen.git
   cd testDocGen
   ```

2. (Opcional) Crear y activar un entorno virtual:

   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # Linux / macOS
   # source venv/bin/activate
   ```

3. Instalar dependencias con `pip`.

4. Lanzar la interfaz gráfica:

   ```bash
   python gui_launcher.py
   ```

---

## Construir el ejecutable (.exe)

El proyecto incluye un archivo `.spec` que permite construir un ejecutable utilizando PyInstaller.

Ejemplo de comando (una vez instalado PyInstaller):

```bash
pip install pyinstaller
pyinstaller gui_launcher.spec
```

Esto generará una carpeta `dist/` con el ejecutable resultante (el nombre concreto dependerá de la configuración del `.spec`).

Ese `.exe` es el que se distribuye al usuario final: basta con copiar el ejecutable junto con los recursos necesarios (`header.png`, `footer.png`, etc.).

---

## Notas y limitaciones

* Los archivos de entrada deben ser los Excel exportados por Azure DevOps con el formato esperado por la herramienta.
* Si cambia el formato de exportación en Azure DevOps (columnas, nombres de campos, etc.), puede ser necesario ajustar la lógica en `data_loader.py` o `processor.py`.
* Actualmente no se incluye un archivo de licencia explícito; hasta que se añada uno, se asume uso interno.

---

## Pendiente / ideas futuras

* Añadir un `requirements.txt` con todas las dependencias.
* Documentar con más detalle el formato exacto de los Excel de entrada.
* Incluir una sección de “Solución de problemas” con los errores más comunes (por ejemplo, columnas faltantes, rutas inválidas, etc.).
* Publicar ejecutables ya construidos en la sección de *Releases* del repositorio.