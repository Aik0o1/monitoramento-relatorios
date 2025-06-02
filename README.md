## 📊 Processador de Indicadores BI

Este projeto realiza o processamento e análise estatística de arquivos Excel contendo indicadores mensais, gerando um relatório consolidado com as flutuações detectadas.

> Além disso, há um teste de interface gráfica utilizando customtkinter.

## 🗂️ Estrutura do Projeto

```
├── front/
│   └── main.py        # Interface gráfica (teste) com customtkinter
├── planilhas/
│   └── *.xlsx         # Arquivos Excel utilizados na análise
├── main.py            # Código principal de processamento
├── README.md          # Este arquivo
└── requirements.txt   # (Opcional) Dependências do projeto
```


## 🚀 Como Executar

1. Instale as dependências 
```
pip install pandas openpyxl
```

2. Coloque os arquivos .xlsx que deseja processar na pasta `planilhas/`.


3. Altere o dia de referência na variável dia_referencia em main.py: atualmente configurado como 5.


4. Execute o script principal:
```
python main.py
```


## ✅ Saída
O script irá gerar um arquivo Excel consolidado no formato: `monitoramento_bi_2_YYYYMMDD.xlsx`

Esse arquivo conterá duas abas:

- DADOS: Indicadores processados.
- ESTATÍSTICO: Flutuações detectadas entre os valores dos indicadores.

Exemplo de saída no terminal

```
Arquivo Excel gerado com sucesso: 'monitoramento_bi_2_20250602.xlsx'
- Aba 'DADOS': 123 registros
- Aba 'ESTATÍSTICO': 45 indicadores com flutuações
```