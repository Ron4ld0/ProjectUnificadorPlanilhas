import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

# ==============================================================================
# --- CONFIGURAÇÃO DAS COLUNAS (NOVA ESTRUTURA MAIS SEGURA) ---
# ==============================================================================
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


# ==============================================================================
# --- LÓGICA DE PROCESSAMENTO (TOTALMENTE REESCRITA E MAIS SIMPLES) ---
# ==============================================================================
def processar_planilha(arquivo_entrada, arquivo_saida):
    """
    Lê a planilha, cria as novas colunas unificadas com uma lógica mais segura
    e salva o resultado.
    """
    try:
        # Tenta ler o arquivo, seja CSV ou Excel
        try:
            df = pd.read_csv(arquivo_entrada)
        except Exception:
            df = pd.read_excel(arquivo_entrada)

        # Para cada mapeamento, cria a nova coluna final
        for mapeamento in COLUNAS_MAPEADAS:
            col_primaria = mapeamento["primaria"]
            lista_fallbacks = mapeamento["fallbacks"]
            col_final = mapeamento["final"]

            # Garante que a coluna primária existe e preenche células vazias
            if col_primaria not in df.columns:
                df[col_primaria] = ''
            df[col_primaria] = df[col_primaria].fillna('')

            # Inicia a coluna final com os valores da coluna primária
            df[col_final] = df[col_primaria].astype(str).str.strip()

            # Agora, para cada coluna de fallback na lista...
            for col_fallback in lista_fallbacks:
                # Garante que a coluna de fallback existe e preenche células vazias
                if col_fallback not in df.columns:
                    df[col_fallback] = ''
                df[col_fallback] = df[col_fallback].fillna('')

                # Onde a coluna final ainda estiver vazia, usa o valor do fallback atual
                df.loc[df[col_final] == '', col_final] = df[col_fallback].astype(str).str.strip()
        
        # --- ALTERAÇÃO AQUI: Salva o resultado em um novo arquivo XLSX ---
        # O argumento index=False impede que o pandas crie uma coluna extra com os índices das linhas.
        df.to_excel(arquivo_saida, index=False)
        
        return (True, f"Processo concluído com sucesso!\n\nArquivo salvo como:\n{arquivo_saida}")

    except Exception as e:
        return (False, f"Ocorreu um erro inesperado:\n{e}")

# ==============================================================================
# --- INTERFACE GRÁFICA (Não precisa mexer aqui) ---
# ==============================================================================
def criar_interface():
    def selecionar_arquivo_entrada():
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
        arquivo_entrada = entry_entrada.get()

        if not arquivo_entrada:
            messagebox.showerror("Atenção", "Por favor, selecione um arquivo para processar.")
            return

        # --- ALTERAÇÃO AQUI: Sugere o nome do arquivo de saída com a extensão .xlsx ---
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

    # --- Criação da Janela ---
    root = tk.Tk()
    root.title("Gerador de Colunas Unificadas")
    root.geometry("500x150")
    root.resizable(False, False)

    frame = ttk.Frame(root, padding="20")
    frame.pack(expand=True, fill="both")

    ttk.Label(frame, text="Selecione a planilha para unificar as colunas:").grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
    entry_entrada = ttk.Entry(frame, width=50, state='readonly')
    entry_entrada.grid(row=1, column=0, sticky="ew")
    ttk.Button(frame, text="Selecionar Arquivo...", command=selecionar_arquivo_entrada).grid(row=1, column=1, padx=(5,0))

    processar_button = ttk.Button(frame, text="Processar e Salvar", command=iniciar_processamento, style="Accent.TButton")
    processar_button.grid(row=2, column=0, columnspan=2, pady=(20,0), ipady=5)

    style = ttk.Style(root)
    style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
    
    root.mainloop()

if __name__ == "__main__":
    criar_interface()