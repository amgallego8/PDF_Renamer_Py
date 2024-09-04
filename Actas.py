import os
import re
import sys
import time
import shutil
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance
import tkinter as tk
from tkinter import messagebox, ttk
from concurrent.futures import ThreadPoolExecutor
import logging

# Configura la ruta de Tesseract OCR si es necesario
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Configurar logging para guardar en un archivo
logging.basicConfig(filename='procesos.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Coordenadas especificadas
x1, y1, x2, y2 = 200, 280, 1200, 450

# Preprocesar imagen para mejorar la precisión del OCR
def preprocess_image(image):
    enhancer = ImageEnhance.Contrast(image)
    image_enhanced = enhancer.enhance(2)  # Ajustar el nivel de contraste
    return image_enhanced

def extract_text_from_image(image):
    # Preprocesar la imagen para mejorar la precisión del OCR
    image = preprocess_image(image)
    
    # Configuración de Tesseract para aumentar la velocidad
    config = '--oem 1 --psm 6'
    text = pytesseract.image_to_string(image, lang='spa', config=config)
    return text

def extract_contract_info(pdf_file_path):
    try:
        # Reducir la resolución a 150 DPI para mejorar la eficiencia
        images = convert_from_path(pdf_file_path, first_page=1, last_page=1, dpi=150)
        if images:
            cropped_image = images[0].crop((x1, y1, x2, y2))  # Recortar la parte relevante según las coordenadas proporcionadas
            text = extract_text_from_image(cropped_image)
            
            # Mostrar el texto extraído para depuración
            print(f"Texto extraído del archivo {pdf_file_path}: {text}")
        else:
            text = ""
    except Exception as e:
        logging.error(f"Error al extraer información del archivo {pdf_file_path}: {str(e)}")
        text = ""
    
    return text

def clean_text(text):
    cleaned_text = text.replace('\n', ' ').replace('\r', '')
    return cleaned_text

def extract_specific_data(text):
    # Expresión regular para capturar el número de contrato
    contract_number = re.search(r'CONTRATO\s*(\d{3}-\d{4})', text)
    
    # Expresión regular para capturar el nombre del contratista
    contractor_name = re.search(r'CONTRATISTA\s*([A-Z\s]+)', text)
    
    # Verificar si se encontraron coincidencias y extraer la información
    contract_number = contract_number.group(1) if contract_number else "No encontrado"
    contractor_name = contractor_name.group(1).strip() if contractor_name else "No encontrado"
    
    return contract_number, contractor_name

def rename_and_move_pdf_file(original_path, contract_number, contractor_name, output_directory):
    try:
        if contract_number == "No encontrado" or contractor_name == "No encontrado":
            new_file_name = "NOMBRAR MANUALMENTE.pdf"
        else:
            contractor_name = ' '.join(contractor_name.split()[:3])
            safe_contractor_name = re.sub(r'[^\w\s-]', '', contractor_name).strip()
            new_file_name = f"ACTA DE INICIO {contract_number} {safe_contractor_name}.pdf"
        
        new_file_name = new_file_name.strip()

        # Crear la carpeta 'Actas renombradas' en la carpeta de salida si no existe
        output_folder = os.path.join(output_directory, "Actas renombradas")
        os.makedirs(output_folder, exist_ok=True)
        
        new_file_path = os.path.join(output_folder, new_file_name)
        
        # Mover directamente sin creación de archivo temporal
        shutil.move(original_path, new_file_path)
        logging.info(f"Archivo procesado: {original_path} renombrado como {contract_number} - {contractor_name}")
        print(f"Archivo renombrado a: {new_file_path}")
    except Exception as e:
        logging.error(f"Error procesando {original_path}: {str(e)}")

def open_folder(path):
    if os.name == 'nt':  # Windows
        os.startfile(path)
    elif os.name == 'posix':  # macOS, Linux
        os.system(f'open "{path}"' if sys.platform == 'darwin' else f'xdg-open "{path}"')

def update_progress(progress_var, value):
    progress_var.set(value)
    root.update_idletasks()

# Nueva función que reemplaza la lambda
def process_pdf_file(input_directory, pdf_file):
    pdf_file_path = os.path.join(input_directory, pdf_file)
    contract_text = extract_contract_info(pdf_file_path)
    return {'pdf_file': pdf_file, 'contract_text': contract_text}

# Procesamiento en paralelo
def process_pdfs_in_parallel(input_directory, output_directory, pdf_files):
    with ThreadPoolExecutor(max_workers=4) as executor:  # Ajustar el número de hilos según la capacidad del sistema
        results = list(executor.map(process_pdf_file, [input_directory]*len(pdf_files), pdf_files))
    
    return results

def main():
    global root
    root = tk.Tk()
    root.withdraw()
    
    # Directorios específicos adaptados
    input_directory = r'D:\Users\Leonel\Documentos\Renamer FULL V2\Entrada_pdf'
    output_directory = r'D:\Users\Leonel\Documentos\Renamer FULL V2\Salida_pdf'
    os.makedirs(output_directory, exist_ok=True)  # Crear el directorio de salida si no existe
    
    # Eliminar archivos no PDF de la carpeta de entrada
    non_pdf_files = [f for f in os.listdir(input_directory) if not f.endswith('.pdf')]
    for file in non_pdf_files:
        file_path = os.path.join(input_directory, file)
        os.remove(file_path)  # Eliminar archivos no PDF
        print(f"Eliminado archivo no PDF: {file_path}")
    
    # Escanear archivos PDF restantes
    pdf_files = [f for f in os.listdir(input_directory) if f.endswith('.pdf')]

    if not pdf_files:
        messagebox.showinfo("Información", "No tienes archivos PDF en la carpeta Entrada_pdf. Agrega los archivos y vuelve a ejecutar el programa.")
        open_folder(input_directory)
        sys.exit()

    # Ventana de progreso
    progress_win = tk.Toplevel(root)
    progress_win.title("Progreso")
    progress_label = ttk.Label(progress_win, text="Procesando archivos PDF...")
    progress_label.pack(pady=10)
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(progress_win, variable=progress_var, maximum=100)
    progress_bar.pack(fill=tk.X, padx=20, pady=10)
    
    # Procesamiento en paralelo de archivos PDF
    start_time = time.time()
    
    results = process_pdfs_in_parallel(input_directory, output_directory, pdf_files)
    
    for i, result in enumerate(results):
        pdf_file_path = os.path.join(input_directory, result['pdf_file'])
        cleaned_text = clean_text(result['contract_text'])
        contract_number, contractor_name = extract_specific_data(cleaned_text)
        
        rename_and_move_pdf_file(pdf_file_path, contract_number, contractor_name, output_directory)
        
        # Mostrar progreso
        progress = (i + 1) / len(pdf_files) * 100
        update_progress(progress_var, progress)
    
    end_time = time.time()
    elapsed_time = end_time - start_time
    
    # Mostrar el tiempo total
    print(f"Tiempo total de procesamiento: {elapsed_time:.2f} segundos para {len(pdf_files)} archivos.")
    progress_win.destroy()
    messagebox.showinfo("Información", f"El proceso ha concluido en {elapsed_time:.2f} segundos. Presiona OK para abrir el directorio de archivos renombrados.")
    open_folder(output_directory)
    
    # Limpieza final: liberar recursos y cerrar ventanas restantes
    try:
        root.quit()  # Cierra el root de tkinter
        root.destroy()  # Asegura que la ventana principal se cierra
    except Exception as e:
        print(f"Error en la limpieza de recursos: {e}")
    
    print("Todos los procesos han sido limpiados correctamente.")

if __name__ == "__main__":
    main()
