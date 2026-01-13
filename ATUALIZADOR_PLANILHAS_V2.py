import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import os
import time

def atualizar_excel_duas_vezes():
    print("=== Atualizador de Planilhas V2 ===")
    
    # 1. Selecionar arquivo
    root = tk.Tk()
    root.withdraw()
    
    arquivo_excel = filedialog.askopenfilename(
        title="Selecione o arquivo Excel para atualizar",
        filetypes=[("Excel Files", "*.xlsx;*.xlsm;*.xlsb")]
    )
    
    if not arquivo_excel:
        print("Nenhum arquivo selecionado. Encerrando.")
        return

    # O win32com precisa do caminho absoluto (C:\Pasta\Arquivo.xlsx)
    caminho_absoluto = os.path.abspath(arquivo_excel)
    print(f"Arquivo selecionado: {caminho_absoluto}")

    excel_app = None
    workbook = None

    try:
        print("\nIniciando o Excel em segundo plano...")
        # Cria uma instância do Excel
        excel_app = win32.Dispatch("Excel.Application")
        
        # Deixe True se quiser ver o Excel abrindo e piscando, False para rodar invisível
        excel_app.Visible = False 
        
        # Desliga alertas (ex: "Deseja salvar?") para não travar o script
        excel_app.DisplayAlerts = False 

        print("-- Abrindo a pasta de trabalho...")
        workbook = excel_app.Workbooks.Open(caminho_absoluto)

        # --- LOOP DE ATUALIZAÇÃO (2 VEZES) ---
        for i in range(1, 3):
            print(f"\n--- [RODADA {i}/2] Executando 'Atualizar Tudo' ---")
            
            # Comando equivalente a clicar em "Dados" > "Atualizar Tudo"
            workbook.RefreshAll()
            
            print(f"--- [RODADA {i}/2] Aguardando conclusão das consultas (Power Query/SQL)...")
            
            # ESTE É O SEGREDO:
            # O Excel roda consultas em background. O script precisa esperar elas terminarem.
            # O comando abaixo trava o script até que as conexões assíncronas terminem.
            excel_app.CalculateUntilAsyncQueriesDone()
            
            # Força um recálculo geral (fórmulas)
            excel_app.Calculate()
            
            # Pequena pausa de segurança para garantir estabilidade entre atualizações
            time.sleep(2)
            print(f"--- [RODADA {i}/2] Concluída.")

        # --- FINALIZAÇÃO ---
        print("\nSalvano o arquivo...")
        workbook.Save()
        print("Arquivo salvo com sucesso.")

    except Exception as e:
        print(f"\n[ERRO CRÍTICO]: {e}")
        
    finally:
        # Garante que o Excel feche mesmo se der erro no meio
        if workbook:
            workbook.Close()
        if excel_app:
            print("Encerrando instância do Excel...")
            excel_app.Quit()
            # Limpa a referência da memória
            del excel_app 
        
        print("\nProcesso finalizado.")

if __name__ == "__main__":
    while True:
        atualizar_excel_duas_vezes()
        resposta = input("\nNova atualização (s/n)? ").strip().lower()
        if resposta != 's':
            break