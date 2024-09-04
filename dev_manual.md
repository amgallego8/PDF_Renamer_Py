Manual del Programador para la Aplicación "Renamer FULL V2"
1. Estructura de Archivos
Directorio Principal:
contratos.py: Archivo principal que contiene todo el código de la aplicación.
Entrada_pdf: Carpeta donde se colocan los archivos PDF para procesar.
Salida_pdf: Carpeta donde se almacenan los archivos renombrados.
2. Dependencias y Configuración
Bibliotecas:

pytesseract: Para realizar el OCR en los archivos PDF.
pdf2image: Para convertir las páginas PDF en imágenes.
PIL: Para el procesamiento de imágenes.
tkinter: Para la interfaz gráfica de usuario y la ventana de mensajes.
Instalación de Dependencias:

bash
Copiar código
pip install pytesseract pdf2image pillow tkinter
Configuración de Tesseract: Asegúrate de configurar correctamente la ruta del OCR:

python
Copiar código
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
3. Funcionalidades Clave
Extracción de Información:

La función extract_contract_info() convierte la primera página de cada PDF en una imagen, recorta la parte relevante y extrae el texto utilizando OCR.
Procesamiento en Paralelo:

El procesamiento de PDFs se realiza en paralelo utilizando ThreadPoolExecutor, lo que permite un manejo eficiente de hasta 200 archivos simultáneamente.
El programa divide los archivos en lotes de 50 para optimizar el uso de memoria.
Renombrado y Organización:

La función rename_and_move_pdf_file() crea una subcarpeta para cada archivo renombrado, que se basa en el número de contrato y el nombre del contratista.
Los archivos PDF se renombran y se mueven a la carpeta Contratos Ordenados.
4. Manejo de Excepciones
La aplicación incluye manejo de excepciones para evitar que el proceso se interrumpa en caso de errores.
python
Copiar código
try:
    # Procesar PDF
except Exception as e:
    print(f"Error procesando {pdf_file}: {e}")
5. Detalles de Optimización
Multithreading: El uso de ThreadPoolExecutor permite que los archivos se procesen en paralelo para reducir el tiempo total de procesamiento.
Liberación de Memoria: Las imágenes generadas por pdf2image se cierran inmediatamente después de ser procesadas para evitar el uso innecesario de memoria.
