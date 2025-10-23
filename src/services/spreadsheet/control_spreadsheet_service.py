import pandas as pd

def list_contracts(df, expected_columns) -> list[dict]:
    contracts = []

    # Garante que todas as colunas esperadas existam no DataFrame
    available_columns = set(df.columns)
    missing_columns = [c for c in expected_columns if c not in available_columns]
    if missing_columns:
        raise ValueError(f"Missing columns in DataFrame: {missing_columns}")

    # Converte as colunas de data para datetime nativo do Python
    for date_col in ["INICIO", "FINAL", "DATA DA MEDIÇÃO", "DATA VENCIMENTO"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).apply(
                lambda x: x.to_pydatetime() if pd.notna(x) else None
            )

    # Cria lista de contratos
    for _, row in df.iterrows():
        contract = {
            "status": str(row.get("STATUS", "")).strip(),
            "due_date": row.get("DATA VENCIMENTO").to_pydatetime() if pd.notna(row.get("DATA VENCIMENTO")) else None,
            "description": str(row.get("DESCRIÇÃO", "")).strip(),
            "description_detail": str(row.get("DESCRIÇÃO DETALHADA", "")).strip(),
            "asset": str(row.get("PATRIMONIO", "")).strip(),
            "supplier": str(row.get("FORNECEDOR", "")).strip(),
            "rental": str(row.get("LOCAÇÃO", "")).strip(),
            "cost_center": str(row.get("CENTRO DE CUSTO", "")).strip(),
            "total_value": float(row.get("VALOR TOTAL DA FATURA")) if pd.notna(row.get("VALOR TOTAL DA FATURA")) else None,
            "contract_number": str(row.get("NUMERO DO CONTRATO", "")).strip(),
            "contract_number_sienge": str(row.get("Nº CONTRATO SIENGE", "")).strip(),
            "observations": str(row.get("OBSERVAÇÕES", "")).strip(),
            "state": str(row.get("ESTADO", "")).strip(),
            "start_date": row.get("INICIO").to_pydatetime() if pd.notna(row.get("INICIO")) else None,
            "end_date": row.get("FINAL").to_pydatetime() if pd.notna(row.get("FINAL")) else None,
        }

        # Verifica se todos os campos esperados estão presentes dentro do contrato
        missing_fields = [key for key, value in contract.items() if value in (None, "", float('nan'))]
        if missing_fields:
            contract["missing_fields"] = missing_fields  # adiciona informação sobre campos ausentes

        contracts.append(contract)

    return contracts


def filter(contracts: list[dict]) -> list[dict]:
    filtered = []
    for c in contracts:
        supplier = str(c.get("supplier", "")).strip().upper()
        due_date = c.get("due_date")

        # Ignora contratos sem data de vencimento
        if due_date in (None, "", " "):
            continue

        # Ignora contratos com fornecedor MARQUES & BEZERRA
        if supplier == "MARQUES & BEZERRA":
            continue

        filtered.append(c)

    return filtered
