from datetime import timedelta
from interfaces import app_interfaces as interfaces

COLUNAS_ESPERADAS = [
    "DESCRIÇÃO",
    "PATRIMONIO",
    "FORNECEDOR",
    "LOCAÇÃO",
    "CENTRO DE CUSTO",
    "ESTADO",
    "INICIO",
    "FINAL",
    "DATA DA MEDIÇÃO",
    "VALOR TOTAL DA MEDIÇÃO",
    "VALOR TOTAL DA FATURA",
    "DATA VENCIMENTO",
    "TITULO EMITIDO",
    "STATUS",
    "NUMERO DO CONTRATO",
    "Nº CONTRATO SIENGE",
    "OBSERVAÇÕES",
    "DESCRIÇÃO DETALHADA",
]

FIELD_TRANSLATIONS = {
    "status": "Status",
    "due_date": "Data de Vencimento",
    "description": "Descrição",
    "description_detail": "Descrição Detalhada",
    "asset": "Patrimônio",
    "supplier": "Fornecedor",
    "rental": "Locação",
    "cost_center": "Centro de Custo",
    "total_value": "Valor Total da Fatura",
    "contract_number": "Número do Contrato",
    "contract_number_sienge": "Nº Contrato Sienge",
    "observations": "Observações",
    "state": "Estado",
    "start_date": "Data de Início",
    "end_date": "Data Final",
}

# Cache local de lembretes já existentes
reminder_cache = []

def initialize_cache():
    """Carrega todos os lembretes existentes no calendário para o cache."""
    global reminder_cache
    reminder_cache = interfaces.reminders.get(folder_type="calendar")
    print(f"🗂️ {len(reminder_cache)} lembretes carregados no cache.")

def find_existing_reminder(body_lines):
    """Procura no cache um lembrete com o mesmo patrimônio e descrição detalhada (ou corpo idêntico)."""
    # Garante que body_lines seja sempre uma string única para comparar corretamente
    if isinstance(body_lines, list):
        body_joined = "\n".join(line.strip() for line in body_lines if line.strip())
    else:
        body_joined = str(body_lines).strip()

    for reminder in reminder_cache:
        subject = reminder.get("subject", "")
        body = reminder.get("body", "").strip()

        # Normaliza quebras de linha e espaços
        body_normalized = "\n".join(line.strip() for line in body.splitlines() if line.strip())

        # Comparações
        match_body = body_normalized == body_joined
        if match_body:
            return reminder

    return None

def create_contract_reminder(contract: dict):
    due_date = contract.get("due_date")
    n_sienge = contract.get("contract_number_sienge")

    if not due_date:
        print("⚠️ Contrato sem data de vencimento — lembrete não criado.")
        return False
    if not n_sienge:
        print("⚠️ Contrato sem cadastro no Sienge — lembrete não criado.")
        return False

    start_reminder = due_date - timedelta(days=15)
    end_reminder = due_date
    subject = f"Vencimento: {n_sienge}"

    # Monta corpo do lembrete
    body_lines = []
    for key, value in contract.items():
        if key != "missing_fields" and value not in (None, ""):
            field_label = FIELD_TRANSLATIONS.get(key, key.replace("_", " ").capitalize())
            body_lines.append(f"{field_label}: {value}")
    body = "\n".join(body_lines)
    
    # Verifica cache
    existing = find_existing_reminder(body_lines)
    
    if existing:
        print(f"⏩ Lembrete existente sem alterações para {n_sienge} — ignorado.")
        return False

    return

    # Cria o lembrete
    result = interfaces.reminders.create(
        folder_type="calendar",
        subject=subject,
        body=body,
        start_time=start_reminder,
        end_time=end_reminder,
        reminder_minutes_before=15 * 24 * 60,
        reminder_set=True,
        categories="Contratos",
        is_all_day=True
    )

    # Adiciona ao cache
    reminder_cache.append({
        "subject": subject,
        "body": body,
        "start_time": start_reminder,
        "end_time": end_reminder
    })

    print(f"✅ Lembrete criado: {subject}")
    return result


if __name__ == "__main__":
    # interfaces.reminders.delete(folder_type="calendar")
    
    path = interfaces.popup.choose_path()
    if path:
        sheet = interfaces.spreadsheet.read_xlsx(
            path=path,
            columns=COLUNAS_ESPERADAS,
            skip_rows=1,
            sheet_name="OUTUBRO - 2025"
        )
        print("✅ Planilha carregada com sucesso!")

        contracts = interfaces.spreadsheet.listar_contratos(sheet, columns=COLUNAS_ESPERADAS)

        if contracts:
            filtered_contracts = interfaces.spreadsheet.filter(contracts)
            print(f"🔍 {len(filtered_contracts)} contratos encontrados após o filtro.")
            print("🚀 Inicializando cache e criação de lembretes...\n")

            initialize_cache()

            for i, contract in enumerate(filtered_contracts, start=1):
                desc = contract.get("description_detail") or contract.get("description") or "Sem descrição"
                print(f"[{i}/{len(filtered_contracts)}] Processando: {desc}")
                create_contract_reminder(contract)

            print("\n✅ Todos os lembretes foram processados!")
        else:
            print("⚠️ Nenhum contrato encontrado na planilha.")
    else:
        print("❌ Nenhum arquivo foi selecionado.")
