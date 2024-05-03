import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
import psutil
import pandas as pd
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed
import win32com.client as win32
import pythoncom
import time

MAX_CONCURRENT_FILES = 1
CONVERSION_TIMEOUT = 300  

MASTER_FILE_PATH = r"C:\Arquivos Lopes\CONTROLE DE VENDAS\CONTROLE DE VENDAS.xlsm"
FILTERED_EXCEL_PATH = r"X:\Repositorio\Excel_Filtrado"  

def convert_file(input_file, output_file, master_workbook, new_name):
    pythoncom.CoInitialize()  

    input_path = input_file
    filename = os.path.splitext(os.path.basename(input_file))[0]

    current_date = datetime.now().strftime('%d-%m-%Y')

    if new_name:
        output_filename = f"{filename}_{new_name}_{current_date}.pdf"
    else:
        output_filename = f"{filename}_{current_date}.pdf"

    output_dir = os.path.dirname(output_file)
    output_file = os.path.join(output_dir, output_filename)

    pdf_directory = os.path.join(output_dir, filename)
    if not os.path.exists(pdf_directory):
        os.makedirs(pdf_directory)

    output_file = os.path.join(pdf_directory, output_filename)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False  

    workbook = excel.Workbooks.Open(input_path, ReadOnly=False, Editable=True)

    try:
        start_time = time.time() 
        while True:
            try:
                workbook.ExportAsFixedFormat(0, output_file, Quality=win32.constants.xlQualityStandard, IgnorePrintAreas=True, OpenAfterPublish=False)  
                print("Arquivo convertido com sucesso para PDF:", output_file)
                return True
            except Exception as e:
                print(f"Erro ao converter o arquivo para PDF: {e}")
                return False

    finally:
        workbook.Close(True)  
        excel.DisplayAlerts = True  

def filter_and_save_excels():
    source_dir = r"C:\Arquivos Lopes\CONTROLE DE VENDAS\ENVIAR POR EMAIL\PDF\EXCEL"
    excel_files = glob.glob(os.path.join(source_dir, '*.xlsx'))
    excel_files = [file for file in excel_files if not os.path.basename(file).startswith('~$')]

    num_filtered_files = 0  

    for file in excel_files:
        input_path = file

        try:
            df = pd.read_excel(input_path)

            colunas_filtradas = [1, 5, 11]
            df_filtrado = df.iloc[:, colunas_filtradas]

            nome_arquivo_original = os.path.splitext(os.path.basename(input_path))[0]

            data_atual = datetime.now().strftime('%d-%m-%Y')
            nome_arquivo_saida = f"{nome_arquivo_original}_{data_atual}_FILTRADO.xlsx"
            caminho_saida = os.path.join(FILTERED_EXCEL_PATH, nome_arquivo_saida)

            df_filtrado.to_excel(caminho_saida, index=False)

            print("Arquivo Excel filtrado e salvo:", caminho_saida)
            num_filtered_files += 1  

        except Exception as e:
            print(f"Erro no processamento do arquivo {os.path.basename(input_path)}: {e}")

    return num_filtered_files  

def convert_xlsx_to_pdf(input_dirs, output_dirs, new_name):
    total_converted = 0  

    with ThreadPoolExecutor() as executor:
        futures = []

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False  
        master_workbook = excel.Workbooks.Open(MASTER_FILE_PATH, False, True)  

        try:
            for input_dir, output_dir in zip(input_dirs, output_dirs):
                files = [file for file in os.listdir(input_dir) if file.endswith(".xlsx")]

                for file in files:
                    input_path = os.path.join(input_dir, file)
                    filename = os.path.splitext(file)[0]
                    output_path = os.path.join(output_dir, f"{filename}.pdf")

                    if os.path.exists(input_path):  
                        future = executor.submit(convert_file, input_path, output_path, master_workbook, new_name)
                        futures.append(future)

                        if len(futures) >= MAX_CONCURRENT_FILES:
                            completed = list(as_completed(futures))
                            for future in completed:
                                if future.result():
                                    total_converted += 1
                                futures.remove(future)

            for future in as_completed(futures):
                if future.result():
                    total_converted += 1

        finally:
            try:
                master_workbook.Close(False)
                excel.Quit()
            except Exception as e:
                print(f"Erro ao fechar a planilha mãe: {e}")

            for proc in psutil.process_iter():
                if proc.name() == "EXCEL.EXE":
                    try:
                        proc.kill()
                    except Exception as e:
                        print(f"Erro ao finalizar o processo do Excel: {e}")

    return total_converted

def start_conversion():
    input_dirs = [input_folder_path.get()]
    output_dirs = [output_folder_path.get()]

    novo_nome = new_name_entry.get()

    total_converted = convert_xlsx_to_pdf(input_dirs, output_dirs, novo_nome)

    num_converted_files = total_converted

    num_filtered_files = filter_and_save_excels()

    popup_message = f"Olá, Analista. {num_converted_files} arquivos foram convertidos e {num_filtered_files} arquivos foram filtrados. Gratidão para a equipe de Desenvolvimento!!!^.^"
    messagebox.showinfo("Conversão e Filtragem Concluídas", popup_message)

def browse_input_folder():
    folder_path = filedialog.askdirectory()
    input_folder_path.set(folder_path)

def browse_output_folder():
    folder_path = filedialog.askdirectory()
    output_folder_path.set(folder_path)

root = tk.Tk()
root.title("Conversor de Excel para PDF")

input_folder_path = tk.StringVar()
output_folder_path = tk.StringVar()

tk.Label(root, text="Pasta de Entrada:").grid(row=0, column=0, sticky="w")
input_entry = tk.Entry(root, textvariable=input_folder_path, width=50)
input_entry.grid(row=0, column=1, padx=5, pady=5)
browse_input_button = tk.Button(root, text="Procurar", command=browse_input_folder)
browse_input_button.grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Pasta de Saída:").grid(row=1, column=0, sticky="w")
output_entry = tk.Entry(root, textvariable=output_folder_path, width=50)
output_entry.grid(row=1, column=1, padx=5, pady=5)
browse_output_button = tk.Button(root, text="Procurar", command=browse_output_folder)
browse_output_button.grid(row=1, column=2, padx=5, pady=5)

tk.Label(root, text="Novo Nome (opcional):").grid(row=2, column=0, sticky="w")
new_name_entry = tk.Entry(root, width=50)
new_name_entry.grid(row=2, column=1, padx=5, pady=5)

convert_button = tk.Button(root, text="Iniciar Conversão", command=start_conversion)
convert_button.grid(row=3, column=1, pady=10)

root.mainloop()
