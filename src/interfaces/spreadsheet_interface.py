import pandas as pd
from services import spreadsheet

def read_xlsx(path: str, columns: list[str] = None, skip_rows: int = 0, sheet_name = 0) -> pd.DataFrame:
    """
    Lê um arquivo Excel, podendo escolher uma planilha específica, 
    pular linhas desnecessárias e validar colunas obrigatórias.
    
    Parâmetros:
        path (str): Caminho do arquivo Excel.
        columns (list): Lista de colunas obrigatórias.
        skip_rows (int): Número de linhas a ignorar no início do arquivo. Padrão = 0.
        sheet_name (str | int | None): Nome ou índice da planilha a ser lida. 
            - Nome da aba (ex: "Planilha2")
            - Índice (ex: 0 para a primeira, 1 para a segunda)
            - None para ler todas as planilhas.
    
    Retorna:
        pd.DataFrame: DataFrame com as colunas validadas.
    """
    return spreadsheet.read.xlsx(path=path, columns=columns, skip_rows=skip_rows, sheet_name=sheet_name)

def listar_contratos(df, columns: list[str]) -> list[dict]:
    """
    Transforma o DataFrame em uma lista de dicionários padronizada,
    garantindo que todos os valores sejam tipos nativos do Python.
    """
    return spreadsheet.control.list_contracts(df=df, expected_columns=columns)

def filter(contratos: list[dict]) -> list[dict]:
    """
    Remove todos os contratos cujo 'supplier' (fornecedor) seja 'MARQUES & BEZERRA'
    """
    return spreadsheet.control.filter(contratos)
