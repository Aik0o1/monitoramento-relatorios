import glob
import pandas as pd
from pathlib import Path
from datetime import datetime
import os

def converter_mes_numero(mes):
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
    return (f"{mes_num}")

def tratar_df(df):
    df = df.drop(df.columns[2:14], axis=1)
    df = df.drop(df.columns[14:], axis=1)
    df = df.drop(df.index[0])
    df = df[~df.map(lambda x: 'Totais' in str(x)).any(axis=1)]
    # Transformando a primeira linha no cabeçalho
    df = df.set_axis(df.iloc[0], axis=1)
    df = df[1:]
    return df

def determinar_valor_referencia(data_extracao_str, dia_referencia=5):
    """Determina se o valor é de referência baseado no dia da extração"""
    try:
        data_extracao = datetime.strptime(data_extracao_str, '%d/%m/%Y')
        return 'Sim' if data_extracao.day == dia_referencia else 'Não'
    except:
        return 'Não'

def processar_arquivos(pasta_arquivos, dia_referencia=5):
    arquivos = sorted(glob.glob(f"{pasta_arquivos}/*.xlsx"))
    resultados = []
    
    for arquivo in arquivos:
        nome_arquivo = Path(arquivo).stem
        data_extracao = datetime.fromtimestamp(os.path.getmtime(arquivo)).strftime('%d/%m/%Y')
        df = tratar_df(pd.read_excel(arquivo))
        
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
                        
                        mes_numero = converter_mes_numero(mes_col)
                        valor_ref = determinar_valor_referencia(data_extracao, dia_referencia)
                        
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
                    print(f"Erro ao processar linha: {e}")
                    continue
    
    return pd.DataFrame(resultados)

# Exemplo de uso
if __name__ == "__main__":
    pasta_arquivos = "../historico-sem-mei" 

    try:
        df_resultado = processar_arquivos(pasta_arquivos)
        print("\nResultado processado:")
        print(df_resultado)
        
        # Salva o resultado em CSV
        df_resultado.to_csv("dados_consolidados.csv", index=False, sep=';', decimal=',')
        print("\nDados consolidados salvos em 'dados_consolidados.csv'")
        
    except Exception as e:
        print(f"Erro: {str(e)}")