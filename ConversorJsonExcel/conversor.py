import pandas as pd
import json
from openpyxl import load_workbook
import os
from tkinter import Tk, Label, Button, Listbox, MULTIPLE, messagebox, Scrollbar, Frame
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tqdm import tqdm  # Biblioteca para a barra de progresso no console

def flatten_json(json_obj):
    """Achatar um JSON com sub-objetos."""
    flat_dict = {}
    def _flatten(json_obj, parent_key=''):
        if isinstance(json_obj, dict):
            for k, v in json_obj.items():
                new_key = f"{parent_key}.{k}" if parent_key else k
                if isinstance(v, dict):
                    _flatten(v, new_key)
                elif isinstance(v, list):
                    flat_dict[new_key] = str(v)  # Converte a lista para string
                else:
                    flat_dict[new_key] = v
        elif isinstance(json_obj, list):
            # Trata listas como strings
            flat_dict[parent_key] = str(json_obj)
        else:
            flat_dict[parent_key] = json_obj
    _flatten(json_obj)
    return flat_dict

def read_and_process_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    # Achatar o JSON
    flat_data = [flatten_json(item) for item in data]

    # Criar o DataFrame
    df = pd.DataFrame(flat_data)

    return df

def save_to_excel(df, output_path, selected_columns):
    # Filtrar as colunas selecionadas
    if selected_columns:
        df = df[selected_columns]

    # Criar um arquivo Excel
    df.to_excel(output_path, index=False, engine='openpyxl')

    # Ajustar a largura das colunas
    wb = load_workbook(output_path)
    ws = wb.active

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Obter a letra da coluna
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 2)  # Adiciona um pouco de margem
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output_path)

def select_columns(df):
    """Cria uma interface para selecionar as colunas a serem exportadas com lista suspensa e rolagem."""
    root = Tk()
    root.title("Seleção de Colunas")

    # Criar o frame principal
    frame = Frame(root)
    frame.pack(pady=10, padx=10, fill='both', expand=True)

    # Adicionar um rótulo
    Label(frame, text="Selecione as colunas para exportar:").pack()

    # Criar uma lista suspensa com rolagem
    listbox = Listbox(frame, selectmode=MULTIPLE, height=20, width=60)
    scrollbar = Scrollbar(frame, orient='vertical', command=listbox.yview)
    listbox.config(yscrollcommand=scrollbar.set)
    scrollbar.pack(side='right', fill='y')
    listbox.pack(side='left', fill='both', expand=True)

    # Adicionar as colunas ao Listbox
    for col in df.columns:
        listbox.insert('end', col)

    def on_apply():
        # Verificar quais colunas estão selecionadas
        selected_indices = listbox.curselection()
        selected_columns = [listbox.get(i) for i in selected_indices]
        if not selected_columns:
            messagebox.showwarning("Aviso", "Nenhuma coluna selecionada.")
            return
        root.destroy()
        process_and_save(df, selected_columns)

    # Adicionar botão para aplicar a seleção
    Button(frame, text="Aplicar", command=on_apply).pack(pady=10)

    root.mainloop()

def process_and_save(df, selected_columns):
    """Processa o JSON e salva o DataFrame em um arquivo Excel com as colunas selecionadas."""
    # Solicitar ao usuário o nome e o local para salvar o arquivo Excel
    output_file_path = asksaveasfilename(title="Escolha o local para salvar o arquivo Excel", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file_path:
        raise FileNotFoundError("Nenhum local selecionado para salvar o arquivo Excel.")
    
    try:
        # Inicializar a barra de progresso
        with tqdm(total=len(df), desc="Salvando arquivo Excel") as pbar:
            # Salvar o DataFrame em um arquivo Excel com as colunas selecionadas
            save_to_excel(df, output_file_path, selected_columns)
            # Atualizar a barra de progresso para 100%
            pbar.update(len(df))
        
        # Mensagem de conclusão
        messagebox.showinfo("Concluído", f"Arquivo salvo em: {output_file_path}")

    except Exception as e:
        print(f"Erro: {e}")

def run_process():
    root = Tk()
    root.withdraw()  # Ocultar a janela principal

    # Solicitar ao usuário o caminho para o arquivo JSON
    json_file_path = askopenfilename(title="Selecione o arquivo JSON", filetypes=[("JSON files", "*.json")])
    if not json_file_path:
        raise FileNotFoundError("Nenhum arquivo JSON selecionado.")

    try:
        # Processar o JSON e criar o DataFrame
        df = read_and_process_json(json_file_path)

        # Verificar o número de linhas processadas
        print(f"Número de linhas no DataFrame: {len(df)}")

        # Selecionar as colunas a serem exportadas
        select_columns(df)

    except Exception as e:
        print(f"Erro: {e}")

# Iniciar o processo
run_process()
