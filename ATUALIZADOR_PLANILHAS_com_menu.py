import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import os
import time
import pythoncom  # Necessário para lidar com erros de COM em algumas situações

def exibir_menu():
    print("\n" + "="*60)
    print(" MENU DE ATUALIZAÇÃO ".center(60, "="))
    print("="*60+"\n")
    print("1. Atualizar APENAS Consultas SQL (Refresh All)")
    print("2. Atualizar APENAS Tabelas Dinâmicas")
    print("3. Atualizar POR COMPLETO (Queries + TDs + Gráficos)")
    print("0. Sair")
    print("="*60)
    
    while True:
        try:
            opcao = int(input("Escolha uma opção (0-3): "))
            if opcao in [0, 1, 2, 3]:
                return opcao
            else:
                print("Opção inválida. Digite um número de 0 a 3.")
        except ValueError:
            print("Entrada inválida. Digite um número.")

def atualizar_excel(opcao):
    # 1. Selecionar arquivo
    root = tk.Tk()
    root.withdraw()
    
    arquivo_excel = filedialog.askopenfilename(
        title="Selecione o arquivo Excel para atualizar",
        filetypes=[("Excel Files", "*.xlsx;*.xlsm;*.xlsb")]
    )
    
    if not arquivo_excel:
        print("Nenhum arquivo selecionado. Retornando ao menu.")
        return

    caminho_absoluto = os.path.abspath(arquivo_excel)
    print(f"\nArquivo selecionado: {caminho_absoluto}")

    excel_app = None
    workbook = None

    try:
        print("[SETUP] Iniciando App Excel isolado...")
        excel_app = win32.DispatchEx("Excel.Application")
        excel_app.Visible = False 
        excel_app.DisplayAlerts = False 
        excel_app.ScreenUpdating = False
        
        workbook = excel_app.Workbooks.Open(caminho_absoluto)

        # Desativa atualização em background para obrigar o script a esperar
        for connection in workbook.Connections:
            try:
                if connection.OLEDBConnection:
                    connection.OLEDBConnection.BackgroundQuery = False
                elif connection.ODBCConnection:
                    connection.ODBCConnection.BackgroundQuery = False
            except:
                pass

        # --- LÓGICA DE EXECUÇÃO BASEADA NA OPÇÃO ---
        
        # OPÇÃO 1 e 3: Atualiza Conexões/Queries
        if opcao in [1, 3]:
            t_dados_inicio = time.time()
            print("\n   -> Atualizando conexões de dados (RefreshAll)...")
            workbook.RefreshAll()
            excel_app.CalculateUntilAsyncQueriesDone()
            print(f"      Concluído em {time.time() - t_dados_inicio:.2f} segundos.")

        # OPÇÃO 2 e 3: Atualiza Tabelas Dinâmicas
        if opcao in [2, 3]:
            print("\n   -> Forçando atualização individual das Tabelas Dinâmicas...")
            t_TD_inicio = time.time()
            for sheet in workbook.Sheets:
                for pivot in sheet.PivotTables():
                    try:
                        pivot.PivotCache().Refresh()
                    except Exception as e:
                        print(f"      [Aviso] Falha ao atualizar dinâmica na aba {sheet.Name}: {e}")
            print(f"      Concluído em {time.time() - t_TD_inicio:.2f} segundos.")

        # OPÇÃO 3: Recálculo Total e Renderização de Gráficos
        if opcao == 3:
            t_calc_inicio = time.time()
            print("\n   -> Forçando recálculo total (Ctrl + Shift + Alt + F9)...")
            excel_app.CalculateFullRebuild()
            print(f"      Concluído em {time.time() - t_calc_inicio:.2f} segundos.")
            
            # Pausa técnica essencial para o Excel renderizar gráficos/dashboards.
            # Se fechar imediatamente, o Excel não tem tempo de montar a interface gráfica.
            print("\n   -> Aguardando carregamento e renderização dos gráficos...")
            excel_app.ScreenUpdating = True 
            time.sleep(4) 

        # --- FINALIZAÇÃO ---
        print("\n[SALVAR] Salvando arquivo...")
        t_save_inicio = time.time()
        workbook.Save()
        print(f"   -> Salvo em {time.time() - t_save_inicio:.2f} segundos.")

    except Exception as e:
        print(f"\n[ERRO CRÍTICO]: {e}")
        
    finally:
        print("\nLimpando processos...")
        
        if workbook:
            try:
                workbook.Close(SaveChanges=False) 
            except:
                pass
                
        if excel_app:
            try:
                excel_app.ScreenUpdating = True
                excel_app.Interactive = True
                excel_app.DisplayAlerts = True
                print("Encerrando instância dedicada do Excel...")
                excel_app.Quit() 
            except:
                pass
            del excel_app
        
        print("Processo finalizado.")

if __name__ == "__main__":
    pythoncom.CoInitialize()
    
    print("\n" + "="*60)
    print(" ATUALIZADOR DE PLANILHAS V3 - INTERATIVO ".center(60))
    print("="*60)
    
    while True:
        opcao_escolhida = exibir_menu()
        
        if opcao_escolhida == 0:
            print("\nEncerrando o atualizador. Até logo!")
            break
            
        atualizar_excel(opcao_escolhida)