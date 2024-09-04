import os
import re
import sys
import time
import shutil
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import tkinter as tk
from tkinter import messagebox, ttk
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configura la ruta de Tesseract OCR si es necesario
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Definir área específica para recortar la imagen y procesar solo la parte relevante
x1, y1, x2, y2 = 50, 50, 1500, 650  # Ajustar según las necesidades del documento

def extract_text_from_image(image):
    # Convertir la imagen a escala de grises para mejorar la precisión del OCR
    image = image.convert('L')  # 'L' es el modo de escala de grises en PIL
    # Configuración eficiente del OCR
    config = '--oem 1 --psm 6'  # Modo eficiente de Tesseract
    text = pytesseract.image_to_string(image, lang='spa', config=config)
    return text

def extract_contract_info(pdf_file_path):
    # Convertir la primera página, pero reducir el DPI a 100 para mejorar la eficiencia
    images = convert_from_path(pdf_file_path, first_page=1, last_page=1, dpi=100)
    if images:
        # Recortar la parte relevante de la imagen
        cropped_image = images[0].crop((x1, y1, x2, y2))  # Ajustar según las necesidades del documento
        text = extract_text_from_image(cropped_image)
        # Liberar memoria de la imagen después de procesarla
        cropped_image.close()
    else:
        text = ""
    return text

def clean_text(text):
    cleaned_text = text.replace('\n', ' ').replace('\r', '')
    return cleaned_text

def extract_specific_data(text):
    contract_number = re.search(r'No\.\s*(\d{3}-\d{4})', text)
    contractor_name = re.search(r'CONTRATISTA:\s*([A-Z\s]+) IDENTIFICACI', text)
    
    contract_number = contract_number.group(1) if contract_number else "No encontrado"
    contractor_name = contractor_name.group(1).strip() if contractor_name else "No encontrado"
    
    return contract_number, contractor_name

def rename_and_move_pdf_file(original_path, contract_number, contractor_name, output_directory):
    if contract_number == "No encontrado" or contractor_name == "No encontrado":
        new_file_name = "NOMBRAR MANUALMENTE.pdf"
    else:
        contractor_name = ' '.join(contractor_name.split()[:3])
        safe_contractor_name = re.sub(r'[^\w\s-]', '', contractor_name).strip()
        new_file_name = f"CONTRATO {contract_number} {safe_contractor_name}.pdf"
    
    new_file_name = new_file_name.strip()
    
    # Crear la carpeta Contratos Ordenados
    main_folder_path = os.path.join(output_directory, "Contratos Ordenados")
    os.makedirs(main_folder_path, exist_ok=True)
    
    # Crear una subcarpeta dentro de Contratos Ordenados con el nombre del contrato
    contract_folder_path = os.path.join(main_folder_path, new_file_name.replace('.pdf', '').strip())
    os.makedirs(contract_folder_path, exist_ok=True)
    
    # Mover el archivo a la subcarpeta sin usar archivo temporal
    new_file_path = os.path.join(contract_folder_path, new_file_name)
    
    # Copiar el archivo directamente y eliminar el original
    try:
        shutil.copy2(original_path, new_file_path)  # Copiar el archivo al nuevo destino
        os.remove(original_path)  # Eliminar el archivo original después de copiarlo
        print(f"Archivo renombrado y movido a: {new_file_path}")
    except FileNotFoundError as e:
        print(f"Error: {e}")

def process_pdf(pdf_file, input_directory, output_directory):
    pdf_file_path = os.path.join(input_directory, pdf_file)
    try:
        contract_text = extract_contract_info(pdf_file_path)
        cleaned_text = clean_text(contract_text)
        contract_number, contractor_name = extract_specific_data(cleaned_text)
        rename_and_move_pdf_file(pdf_file_path, contract_number, contractor_name, output_directory)
    except Exception as e:
        print(f"Error procesando {pdf_file}: {e}")

def process_pdfs_in_parallel(pdf_files, input_directory, output_directory, batch_size=50):
    total_files = len(pdf_files)
    for i in range(0, total_files, batch_size):
        batch = pdf_files[i:i + batch_size]  # Dividir los archivos en lotes de 50
        with ThreadPoolExecutor() as executor:
            futures = [executor.submit(process_pdf, pdf_file, input_directory, output_directory) for pdf_file in batch]
            for future in as_completed(futures):
                future.result()  # Esperar que se complete cada tarea

def clean_non_pdf_files(input_directory):
    # Eliminar archivos que no sean .pdf
    files = os.listdir(input_directory)
    for file in files:
        if not file.endswith('.pdf'):
            os.remove(os.path.join(input_directory, file))
            print(f"Eliminado archivo no válido: {file}")

def open_folder(path):
    """Abre la carpeta especificada."""
    if os.name == 'nt':  # Windows
        os.startfile(path)
    elif os.name == 'posix':  # macOS, Linux
        os.system(f'open "{path}"' if sys.platform == 'darwin' else f'xdg-open "{path}"')

def main():
    global root
    root = tk.Tk()
    root.withdraw()
    
    # Directorios específicos adaptados
    input_directory = r'D:\Users\Leonel\Documentos\Renamer FULL V2\Entrada_pdf'
    output_directory = r'D:\Users\Leonel\Documentos\Renamer FULL V2\Salida_pdf'
    os.makedirs(output_directory, exist_ok=True)  # Crear el directorio de salida si no existe
    
    # Limpiar archivos no válidos (no PDF)
    clean_non_pdf_files(input_directory)
    
    # Escanear archivos
    pdf_files = [f for f in os.listdir(input_directory) if f.endswith('.pdf')]

    if not pdf_files:
        messagebox.showinfo("Información", "No tienes archivos en la carpeta Entrada_pdf. Agrega los archivos y vuelve a ejecutar el programa.")
        sys.exit()
    
    response = messagebox.askokcancel("Confirmación", f"Tienes {len(pdf_files)} archivos en la carpeta Entrada_pdf. Para proceder, presiona OK.")
    
    if not response:
        sys.exit()
    
    start_time = time.time()
    
    # Procesamiento en paralelo de archivos PDF utilizando multithreading
    process_pdfs_in_parallel(pdf_files, input_directory, output_directory, batch_size=50)
    
    end_time = time.time()
    elapsed_time = end_time - start_time
    
    messagebox.showinfo("Información", f"El proceso ha concluido en {elapsed_time:.2f} segundos. Presiona OK para abrir el directorio de archivos renombrados.")
    open_folder(output_directory)

if __name__ == "__main__":
    main()
