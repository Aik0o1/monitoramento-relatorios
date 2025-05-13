import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import glob
from pathlib import Path
from datetime import datetime
import os
import threading

# Configurando o tema do CustomTkinter
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class ProcessadorIndicadoresApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuração da janela principal
        self.title("Processador de Indicadores")
        self.geometry("700x500")
        
        # Criando frame principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Frame principal
        self.frame = ctk.CTkFrame(self)
        self.frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.frame.grid_columnconfigure(0, weight=1)
        
        # Título
        self.title_label = ctk.CTkLabel(self.frame, text="Processador de Indicadores", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # Descrição
        self.description_label = ctk.CTkLabel(
            self.frame, 
            text="Selecione a pasta contendo os arquivos Excel e defina\no dia de referência para processamento.",
            font=ctk.CTkFont(size=14),
            justify="center"
        )
        self.description_label.grid(row=1, column=0, padx=20, pady=(0, 20))
        
        # Frame para seleção de pasta
        self.folder_frame = ctk.CTkFrame(self.frame)
        self.folder_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.folder_frame.grid_columnconfigure(0, weight=1)
        
        # Entrada para o caminho da pasta
        self.folder_entry = ctk.CTkEntry(self.folder_frame, placeholder_text="Caminho da pasta com arquivos Excel...")
        self.folder_entry.grid(row=0, column=0, padx=(20, 10), pady=10, sticky="ew")
        
        # Botão para selecionar pasta
        self.folder_button = ctk.CTkButton(self.folder_frame, text="Procurar", command=self.browse_folder, width=100)
        self.folder_button.grid(row=0, column=1, padx=(0, 20), pady=10)
        
        # Frame para configurações
        self.settings_frame = ctk.CTkFrame(self.frame)
        self.settings_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        
        # Label para o dia de referência
        self.ref_day_label = ctk.CTkLabel(self.settings_frame, text="Dia de referência:", anchor="w")
        self.ref_day_label.grid(row=0, column=0, padx=20, pady=10, sticky="w")
        
        # Entrada para o dia de referência
        self.ref_day_entry = ctk.CTkEntry(self.settings_frame, width=80)
        self.ref_day_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self.ref_day_entry.insert(0, "5")  # Valor padrão
        
        # Nome do arquivo de saída
        self.output_label = ctk.CTkLabel(self.settings_frame, text="Nome do arquivo de saída:", anchor="w")
        self.output_label.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        
        # Entrada para o nome do arquivo de saída
        self.output_entry = ctk.CTkEntry(self.settings_frame, width=250)
        self.output_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        self.output_entry.insert(0, "consolidado_indicadores.xlsx")  # Valor padrão
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(self.frame, mode="indeterminate")
        self.progress_bar.grid(row=4, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.progress_bar.grid_remove()  # Esconde inicialmente
        
        # Status
        self.status_label = ctk.CTkLabel(self.frame, text="", font=ctk.CTkFont(size=12))
        self.status_label.grid(row=5, column=0, padx=20, pady=(0, 20))
        
        # Botão para processar
        self.process_button = ctk.CTkButton(
            self.frame,
            text="Processar Arquivos",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=40,
            command=self.iniciar_processamento
        )
        self.process_button.grid(row=6, column=0, padx=20, pady=(10, 20))
    
    def browse_folder(self):
        """Abre o diálogo para seleção de pasta"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder_path)
    
    def iniciar_processamento(self):
        """Inicia o processamento em uma thread separada"""
        # Valida entradas
        pasta_arquivos = self.folder_entry.get()
        if not pasta_arquivos:
            messagebox.showerror("Erro", "Por favor, selecione uma pasta de arquivos.")
            return
        
        try:
            dia_referencia = int(self.ref_day_entry.get())
            if dia_referencia < 1 or dia_referencia > 31:
                raise ValueError("Dia de referência deve estar entre 1 e 31.")
        except ValueError as e:
            messagebox.showerror("Erro", f"Dia de referência inválido: {str(e)}")
            return
        
        nome_arquivo_excel = self.output_entry.get()
        if not nome_arquivo_excel.endswith('.xlsx'):
            nome_arquivo_excel += '.xlsx'
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, nome_arquivo_excel)
        
        # Prepara a interface para processamento
        self.status_label.configure(text="Processando arquivos...")
        self.progress_bar.grid()
        self.progress_bar.start()
        self.process_button.configure(state="disabled")
        
        # Inicia o processamento em uma thread separada
        thread = threading.Thread(target=self.processar_arquivos_thread, args=(pasta_arquivos, dia_referencia, nome_arquivo_excel))
        thread.daemon = True
        thread.start()
    
    def processar_arquivos_thread(self, pasta_arquivos, dia_referencia, nome_arquivo_excel):
        """Executa o processamento em uma thread separada"""
        try:
            # Processa os arquivos originais
            self.update_status("Processando arquivos...")
            df_resultado = self.processar_arquivos(pasta_arquivos, dia_referencia)
            
            if df_resultado.empty:
                self.finalizar_processamento(False, "Nenhum dado encontrado para processar.")
                return
            
            # Processa as flutuações
            self.update_status("Calculando estatísticas...")
            df_flutuacoes = self.processar_flutuacoes(df_resultado)
            
            # Cria um arquivo Excel com duas abas
            self.update_status("Gerando arquivo Excel...")
            with pd.ExcelWriter(nome_arquivo_excel, engine='openpyxl') as writer:
                df_resultado.to_excel(writer, sheet_name='DADOS', index=False)
                df_flutuacoes.to_excel(writer, sheet_name='ESTATÍSTICO', index=False)
            
            mensagem = f"Arquivo Excel gerado com sucesso!\n- Dados: {len(df_resultado)} registros\n- Estatísticas: {len(df_flutuacoes)} indicadores com flutuações"
            self.finalizar_processamento(True, mensagem)
            
        except Exception as e:
            self.finalizar_processamento(False, f"Erro durante o processamento: {str(e)}")
    
    def update_status(self, message):
        """Atualiza o status na interface gráfica a partir de thread separada"""
        self.after(0, lambda: self.status_label.configure(text=message))
    
    def finalizar_processamento(self, sucesso, mensagem):
        """Finaliza o processamento e atualiza a interface"""
        def _finalizar():
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            self.process_button.configure(state="normal")
            self.status_label.configure(text="")
            
            if sucesso:
                messagebox.showinfo("Sucesso", mensagem)
            else:
                messagebox.showerror("Erro", mensagem)
        
        self.after(0, _finalizar)
    
    # --- Funções de processamento de dados ---
    
    def converter_mes_numero(self, mes):
        """Converte nomes dos meses em números"""
        meses = {
            'JAN': '01', 'FEV': '02', 'MAR': '03', 'ABR': '04',
            'MAI': '05', 'JUN': '06', 'JUL': '07', 'AGO': '08',
            'SET': '09', 'OUT': '10', 'NOV': '11', 'DEZ': '12'
        }
        
        # Converte o mês para número
        mes_num = meses.get(mes.upper())
        if not mes_num:
            raise ValueError(f"Mês inválido: {mes}")
        
        # Retorna o período no formato numérico
        return mes_num

    def tratar_df(self, df):
        """Trata o DataFrame para extrair os dados relevantes"""
        try:
            df = df.drop(df.columns[2:14], axis=1)
            df = df.drop(df.columns[14:], axis=1)
            df = df.drop(df.index[0])
            df = df[~df.map(lambda x: 'Totais' in str(x)).any(axis=1)]
            # Transformando a primeira linha no cabeçalho
            df = df.set_axis(df.iloc[0], axis=1)
            df = df[1:]
            return df
        except Exception as e:
            raise Exception(f"Erro ao tratar DataFrame: {str(e)}")

    def determinar_valor_referencia(self, data_extracao_str, dia_referencia=5):
        """Determina se o valor é de referência baseado no dia da extração"""
        try:
            data_extracao = datetime.strptime(data_extracao_str, '%d/%m/%Y')
            return 'Sim' if data_extracao.day == dia_referencia else 'Não'
        except:
            return 'Não'

    def processar_arquivos(self, pasta_arquivos, dia_referencia=5):
        """Processa os arquivos Excel na pasta especificada"""
        arquivos = sorted(glob.glob(f"{pasta_arquivos}/*.xlsx"))
        resultados = []
        
        total_arquivos = len(arquivos)
        for i, arquivo in enumerate(arquivos):
            try:
                nome_arquivo = Path(arquivo).stem
                data_extracao = datetime.fromtimestamp(os.path.getmtime(arquivo)).strftime('%d/%m/%Y')
                
                # Atualiza status
                self.update_status(f"Processando arquivo {i+1} de {total_arquivos}: {nome_arquivo}")
                
                df = self.tratar_df(pd.read_excel(arquivo))
                
                for _, row in df.iterrows():
                    tipo_evento = row['Tipo de Evento']
                    ano = row['ANO']
                    
                    # Processa cada coluna de mês
                    for mes_col in [col for col in df.columns if col not in ['Tipo de Evento', 'ANO']]:
                        try:
                            valor = row[mes_col]
                            # Verifica se o valor deve ser incluído
                            if (pd.notna(valor) and 
                                str(valor).strip() not in ['', '-', '0', '0.0'] and
                                float(valor) != 0):
                                
                                mes_numero = self.converter_mes_numero(mes_col)
                                valor_ref = self.determinar_valor_referencia(data_extracao, dia_referencia)
                                
                                resultados.append({
                                    'ARQUIVO': nome_arquivo,
                                    'DATA EXTRAÇÃO': data_extracao,
                                    'INDICADOR': tipo_evento,
                                    'MÊS INDICADOR': mes_numero,
                                    'ANO INDICADOR': ano,
                                    'VALOR': int(valor),
                                    'VALOR DE REFERENCIA': valor_ref
                                })
                        except Exception as e:
                            print(f"Erro ao processar coluna {mes_col}: {e}")
                            continue
            except Exception as e:
                print(f"Erro ao processar arquivo {arquivo}: {e}")
                continue
        
        return pd.DataFrame(resultados)

    def processar_flutuacoes(self, df_original):
        """Processa as flutuações dos indicadores"""
        df = df_original.copy()
        
        # Agrupa por indicador, mês e ano
        grouped = df.groupby(['INDICADOR', 'MÊS INDICADOR', 'ANO INDICADOR'])
        
        resultados = []
        
        for (indicador, mes, ano), group in grouped:
            valores = group['VALOR'].unique()
            
            # Só processa se houver mais de um valor único
            if len(valores) > 1:
                flutuacoes = len(valores) - 1  # Número de mudanças entre valores distintos
                diff_max_min = max(valores) - min(valores)
                
                resultados.append({
                    'INDICADOR': indicador,
                    'MÊS INDICADOR': mes,
                    'ANO INDICADOR': ano,
                    'FLUTUAÇÕES': flutuacoes,
                    'DIF. MAX_MIN': diff_max_min
                })
        
        return pd.DataFrame(resultados)

if __name__ == "__main__":
    app = ProcessadorIndicadoresApp()
    app.mainloop()