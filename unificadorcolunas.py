import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

# Define a estrutura para unificação de colunas.
# Cada dicionário especifica uma coluna primária, colunas de fallback e o nome da coluna final.
COLUNAS_MAPEADAS = [
    {
        "primaria": "Endereço Entrega",
        "fallbacks": ["Endereço"],
        "final": "Endereço Unificado"
    },
    {
        "primaria": "Número Entrega",
        "fallbacks": ["Número Número Entrega", "Número"],
        "final": "Número Unificado"
    },
    {
        "primaria": "CEP Entrega",
        "fallbacks": ["CEP"],
        "final": "CEP Unificado"
    },
    {
        "primaria": "Nome Bai. Entrega",
        "fallbacks": ["Nome do Bairro"],
        "final": "Bairro Unificado"
    }
]

def processar_planilha(arquivo_entrada, arquivo_saida):
    """
    Lê uma planilha (CSV ou Excel), unifica as colunas de endereço conforme o mapeamento
    e salva o resultado em um novo arquivo Excel.
    """
    try:
        # Tenta ler o arquivo como CSV, se falhar, tenta como Excel.
        try:
            df = pd.read_csv(arquivo_entrada)
        except Exception:
            df = pd.read_excel(arquivo_entrada)

        # Itera sobre as configurações de mapeamento para criar as colunas unificadas.
        for mapeamento in COLUNAS_MAPEADAS:
            col_primaria = mapeamento["primaria"]
            lista_fallbacks = mapeamento["fallbacks"]
            col_final = mapeamento["final"]

            # Garante que a coluna primária exista para evitar erros.
            if col_primaria not in df.columns:
                df[col_primaria] = ''
            df[col_primaria] = df[col_primaria].fillna('')

            # Inicia a coluna final com os valores da coluna primária.
            df[col_final] = df[col_primaria].astype(str).str.strip()

            # Preenche os valores vazios na coluna final usando as colunas de fallback.
            for col_fallback in lista_fallbacks:
                if col_fallback not in df.columns:
                    df[col_fallback] = ''
                df[col_fallback] = df[col_fallback].fillna('')

                # Aplica o fallback apenas onde a coluna final ainda está vazia.
                df.loc[df[col_final] == '', col_final] = df[col_fallback].astype(str).str.strip()
        
        # Salva o DataFrame resultante em um arquivo Excel.
        # O argumento index=False evita a criação de uma coluna de índice no arquivo final.
        df.to_excel(arquivo_saida, index=False)
        
        return (True, f"Processo concluído com sucesso!\n\nArquivo salvo como:\n{arquivo_saida}")

    except Exception as e:
        return (False, f"Ocorreu um erro inesperado:\n{e}")

def criar_interface():
    """Cria e configura a interface gráfica do usuário com Tkinter."""
    
    def selecionar_arquivo_entrada():
        """Abre uma caixa de diálogo para o usuário selecionar o arquivo de entrada."""
        filepath = filedialog.askopenfilename(
            title="Selecione a planilha de pedidos",
            filetypes=(("Excel/CSV", "*.xlsx *.csv"), ("Todos os arquivos", "*.*"))
        )
        if filepath:
            entry_entrada.config(state='normal')
            entry_entrada.delete(0, tk.END)
            entry_entrada.insert(0, filepath)
            entry_entrada.config(state='readonly')

    def iniciar_processamento():
        """Inicia o processo de unificação das colunas a partir do arquivo selecionado."""
        arquivo_entrada = entry_entrada.get()

        if not arquivo_entrada:
            messagebox.showerror("Atenção", "Por favor, selecione um arquivo para processar.")
            return

        # Pede ao usuário para definir o local e nome do arquivo de saída.
        arquivo_saida = filedialog.asksaveasfilename(
            title="Salvar arquivo processado como...",
            defaultextension=".xlsx",
            filetypes=(("Arquivo Excel", "*.xlsx"),),
            initialfile="planilha_com_colunas_unificadas.xlsx"
        )
        if not arquivo_saida:
            messagebox.showinfo("Cancelado", "O processo foi cancelado.")
            return

        sucesso, mensagem = processar_planilha(arquivo_entrada, arquivo_saida)

        if sucesso:
            messagebox.showinfo("Sucesso!", mensagem)
        else:
            messagebox.showerror("Erro", mensagem)

    # --- Configuração da Janela Principal ---
    root = tk.Tk()
    root.title("Unificador de Colunas")
    root.geometry("500x150")
    root.resizable(False, False)

    frame = ttk.Frame(root, padding="20")
    frame.pack(expand=True, fill="both")

    # --- Widgets da Interface ---
    ttk.Label(frame, text="Selecione a planilha para unificar as colunas:").grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
    
    entry_entrada = ttk.Entry(frame, width=50, state='readonly')
    entry_entrada.grid(row=1, column=0, sticky="ew")
    
    ttk.Button(frame, text="Selecionar Arquivo...", command=selecionar_arquivo_entrada).grid(row=1, column=1, padx=(5,0))

    processar_button = ttk.Button(frame, text="Processar e Salvar", command=iniciar_processamento, style="Accent.TButton")
    processar_button.grid(row=2, column=0, columnspan=2, pady=(20,0), ipady=5)

    # --- Estilo ---
    style = ttk.Style(root)
    style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
    
    root.mainloop()

if __name__ == "__main__":
    criar_interface()
