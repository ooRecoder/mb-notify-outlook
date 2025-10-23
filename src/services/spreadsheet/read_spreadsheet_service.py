import pandas as pd
from pathlib import Path

def xlsx(path: str, columns: list, skip_rows: int = 0, sheet_name: str | int | None = 0) -> pd.DataFrame:
    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    # Lê o arquivo Excel (planilha específica)
    df = pd.read_excel(path, sheet_name=sheet_name, skiprows=skip_rows)

    # Caso tenha lido múltiplas planilhas (sheet_name=None)
    if isinstance(df, dict):
        print(f"\nForam lidas {len(df)} planilhas: {list(df.keys())}")
        raise ValueError("Defina o parâmetro 'sheet_name' para escolher apenas uma planilha específica.")

    # Normaliza os nomes das colunas
    df.columns = [c.strip().upper() for c in df.columns]

    # Validação das colunas obrigatórias
    for col in columns:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente: {col}")

    return df
