import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import PyPDF2
import openpyxl

def extract_text_from_pdfs(pdf_folder):
    """
    Extrai texto de todos os arquivos PDF em uma pasta.
    
    Parâmetros:
    pdf_folder (str): Caminho para a pasta contendo os PDFs
    
    Retorna:
    list: Lista de dicionários com informações dos PDFs
    """
    pdf_texts = []
    
    # Verifica se a pasta existe
    if not os.path.exists(pdf_folder):
        messagebox.showerror("Erro", f"A pasta {pdf_folder} não existe.")
        return pdf_texts
    
    # Percorre todos os arquivos na pasta
    for filename in os.listdir(pdf_folder):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(pdf_folder, filename)
            
            try:
                # Abre o arquivo PDF
                with open(pdf_path, 'rb') as file:
                    # Cria um leitor de PDF
                    pdf_reader = PyPDF2.PdfReader(file)
                    
                    # Extrai texto de todas as páginas
                    full_text = ''
                    for page in pdf_reader.pages:
                        full_text += page.extract_text() + '\n\n'
                    
                    # Adiciona informações à lista
                    pdf_texts.append({
                        'Nome do Arquivo': filename,
                        'Texto Extraído': full_text.strip(),
                        'Número de Páginas': len(pdf_reader.pages)
                    })
                
            except Exception as e:
                messagebox.showwarning("Aviso", f"Erro ao processar {filename}: {e}")
    
    return pdf_texts

def select_pdf_folder():
    """
    Abre janela para selecionar pasta de PDFs
    """
    folder_selected = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
    pdf_folder_entry.delete(0, tk.END)
    pdf_folder_entry.insert(0, folder_selected)

def select_excel_save_location():
    """
    Abre janela para selecionar local para salvar Excel
    """
    file_selected = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        title="Salvar planilha Excel"
    )
    excel_save_entry.delete(0, tk.END)
    excel_save_entry.insert(0, file_selected)

def process_pdfs():
    """
    Processa PDFs e salva em Excel
    """
    # Obtém os caminhos inseridos
    pdf_folder = pdf_folder_entry.get()
    output_path = excel_save_entry.get()
    
    # Verifica se os caminhos foram selecionados
    if not pdf_folder:
        messagebox.showerror("Erro", "Selecione a pasta com PDFs")
        return
    
    if not output_path:
        messagebox.showerror("Erro", "Selecione o local para salvar o Excel")
        return
    
    # Extrai textos dos PDFs
    pdf_texts = extract_text_from_pdfs(pdf_folder)
    
    # Salva em Excel
    if pdf_texts:
        try:
            # Converte para DataFrame
            df = pd.DataFrame(pdf_texts)
            
            # Salva em Excel
            df.to_excel(output_path, index=False, engine='openpyxl')
            messagebox.showinfo("Sucesso", f"Dados salvos em {output_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo: {e}")
    else:
        messagebox.showwarning("Aviso", "Nenhum PDF encontrado para extrair.")

# Cria a janela principal
root = tk.Tk()
root.title("Extrator de Texto de PDFs")
root.geometry("500x250")

# Frame para pasta de PDFs
pdf_folder_frame = tk.Frame(root)
pdf_folder_frame.pack(padx=10, pady=10, fill=tk.X)

pdf_folder_label = tk.Label(pdf_folder_frame, text="Pasta de PDFs:")
pdf_folder_label.pack(side=tk.LEFT)

pdf_folder_entry = tk.Entry(pdf_folder_frame, width=40)
pdf_folder_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=10)

pdf_folder_button = tk.Button(pdf_folder_frame, text="Procurar", command=select_pdf_folder)
pdf_folder_button.pack(side=tk.LEFT)

# Frame para salvar Excel
excel_save_frame = tk.Frame(root)
excel_save_frame.pack(padx=10, pady=10, fill=tk.X)

excel_save_label = tk.Label(excel_save_frame, text="Salvar Excel:")
excel_save_label.pack(side=tk.LEFT)

excel_save_entry = tk.Entry(excel_save_frame, width=40)
excel_save_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=10)

excel_save_button = tk.Button(excel_save_frame, text="Procurar", command=select_excel_save_location)
excel_save_button.pack(side=tk.LEFT)

# Botão de processamento
process_button = tk.Button(root, text="Extrair Textos", command=process_pdfs, bg="green", fg="white")
process_button.pack(pady=20)

# Inicia a interface gráfica
root.mainloop()