import glob
import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from collections import Counter
from openpyxl.styles import PatternFill

def converter_mes_ano(mes, ano):
    # Dicionário para converter nomes dos meses em números
    meses = {
        'JAN': '01', 'FEV': '02', 'MAR': '03', 'ABR': '04',
        'MAI': '05', 'JUN': '06', 'JUL': '07', 'AGO': '08',
        'SET': '09', 'OUT': '10', 'NOV': '11', 'DEZ': '12'
    }
    
    # Converte o mês para número
    mes_num = meses.get(mes.upper())
    if not mes_num:
        raise ValueError(f"Mês inválido: {mes}")
    
    # Retorna o período no formato YYYY-MM
    return (f"{mes_num}-{ano}")

def tratar_df(df):
    df = df.drop(df.columns[2:14], axis=1)
    df = df.drop(df.columns[14:], axis=1)
    df = df.drop(df.index[0])
    df = df[~df.map(lambda x: 'Totais' in str(x)).any(axis=1)]
    # Transformando a primeira linha no cabeçalho
    df = df.set_axis(df.iloc[0], axis=1)
    df = df[1:]
    return df

def comparar_arquivos(pasta_arquivos):
    arquivos = sorted(glob.glob(f"{pasta_arquivos}/*.xlsx"))
    
    if len(arquivos) < 2:
        raise ValueError("É necessário pelo menos 2 arquivos excel para comparação")
    
    # Carrega todos os arquivos
    dfs = {}
    for arquivo in arquivos:
        nome = Path(arquivo).stem
        df = tratar_df(pd.read_excel(arquivo))
        dfs[nome] = df
    
    # Cria um dicionário para consolidar todos os valores
    consolidado = {}
    
    # Processa cada arquivo
    for nome_arquivo, df in dfs.items():
        for _, row in df.iterrows():
            # Processa cada mês/coluna
            for mes_col in [col for col in df.columns if col not in ['Tipo de Evento', 'ANO']]:
                try:
                    mes_ano = converter_mes_ano(mes_col, row['ANO'])
                    chave = (row['Tipo de Evento'], mes_ano)
                    
                    if chave not in consolidado:
                        consolidado[chave] = {'Tipo_Evento': row['Tipo de Evento'], 
                                              'Mes_Ano': mes_ano}
                    
                    consolidado[chave][nome_arquivo] = row[mes_col]
                except ValueError:
                    continue
    
    # Converte para DataFrame
    df_relatorio = pd.DataFrame(list(consolidado.values()))
    
    # Ordena por Mes_Ano e Tipo_Evento
    df_relatorio = df_relatorio.sort_values(['Mes_Ano', 'Tipo_Evento'])

    # Filtra apenas linhas com diferença entre os arquivos
    arquivos_nomes = list(dfs.keys())
    
    colunas_valores = arquivos_nomes
    df_diferencas = df_relatorio[df_relatorio[colunas_valores].nunique(axis=1) > 1]
    
    return df_diferencas
    
def criar_excel_destacado(df_relatorio, nome_arquivo="relatorio_diferencas.xlsx"):
    # Cria um novo workbook
    wb = Workbook()
    ws = wb.active
    
    # Define os estilos para células
    fill_diff = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Amarelo (diferença da moda)
    fill_high = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")  # Vermelho (valor mais alto)
    fill_low = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")   # Verde (valor mais baixo)
    
    # Escreve o cabeçalho
    for col_idx, col_name in enumerate(df_relatorio.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name)
    
    # Escreve os dados e aplica a formatação condicional
    for row_idx, (_, row) in enumerate(df_relatorio.iterrows(), 2):
        # Escreve os valores das colunas não numéricas primeiro
        for col_idx, col_name in enumerate(df_relatorio.columns, 1):
            if col_name in ['Tipo_Evento', 'Mes_Ano']:
                ws.cell(row=row_idx, column=col_idx, value=row[col_name])
        
        # Pega os valores das colunas numéricas (da 3ª em diante)
        valores = [row[col] for col in df_relatorio.columns[2:]]
        
        # Encontrar a moda (valor mais frequente)
        counter = Counter(valores)
        most_common = counter.most_common(1)
        
        if most_common[0][1] > 1:  # Se há moda (valor repetido)
            moda = most_common[0][0]
            # Destacar células diferentes da moda
            for col_idx, col_name in enumerate(df_relatorio.columns[2:], 3):
                cell = ws.cell(row=row_idx, column=col_idx, value=row[col_name])
                if cell.value != moda:
                    cell.fill = fill_diff
        else:
            # Se não há moda, encontrar o valor mais extremo
            max_val = max(valores)
            min_val = min(valores)
            
            # Destacar o máximo e mínimo
            for col_idx, col_name in enumerate(df_relatorio.columns[2:], 3):
                cell = ws.cell(row=row_idx, column=col_idx, value=row[col_name])
                if cell.value == max_val:
                    cell.fill = fill_high
                elif cell.value == min_val:
                    cell.fill = fill_low
    
    # Ajusta a largura das colunas automaticamente
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                value = str(cell.value) if cell.value is not None else ""
                if len(value) > max_length:
                    max_length = len(value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Congela a primeira linha (cabeçalho)
    ws.freeze_panes = 'A2'
    
    # Salva o arquivo
    wb.save(nome_arquivo)

# Exemplo de uso
if __name__ == "__main__":
    pasta_arquivos = "./historico-sem-mei" 

    try:
        df_relatorio = comparar_arquivos(pasta_arquivos)
        print("\nRelatório de Diferenças:")
        print(df_relatorio)
        
        # Salva o relatório em CSV
        df_relatorio.to_csv("relatorio_diferencas.csv", index=False)
        print("\nRelatório salvo em 'relatorio_diferencas.csv'")
        
        criar_excel_destacado(df_relatorio)
        print("\nExcel com células destacadas salvo em 'relatorio_diferencas.xlsx'")
    except Exception as e:
        print(f"Erro: {str(e)}")