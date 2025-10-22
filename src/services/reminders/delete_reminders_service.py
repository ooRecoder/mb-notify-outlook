import win32com.client
from datetime import datetime, timedelta

def delete(
    folder_type: str = "calendar",
    start_date: datetime = None,
    end_date: datetime = None,
    subject_contains: str = None,
    body_contains: str = None,
    only_with_reminder: bool = False,
    include_recurring: bool = True,
) -> int:
    """
    Exclui itens do Outlook (Calendar ou Tasks) com base nos filtros fornecidos.

    Retorna o número de itens excluídos.
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    if folder_type.lower() == "calendar":
        folder = outlook.GetDefaultFolder(9)  # Calendar
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = include_recurring

        start_date = start_date or datetime.now() - timedelta(days=365 * 10)
        end_date = end_date or (datetime.now() + timedelta(days=365 * 10))

        restriction = f"[Start] >= '{start_date.strftime('%m/%d/%Y %H:%M %p')}' AND [Start] <= '{end_date.strftime('%m/%d/%Y %H:%M %p')}'"
        restricted_items = items.Restrict(restriction)

    elif folder_type.lower() == "tasks":
        folder = outlook.GetDefaultFolder(13)  # Tasks
        restricted_items = folder.Items
    else:
        raise ValueError("folder_type deve ser 'calendar' ou 'tasks'")

    # 🧩 Converter para lista real (evita problema de índice)
    all_items = [item for item in restricted_items]

    deleted_count = 0

    # ⚠️ Iterar de trás pra frente para não quebrar a coleção durante exclusão
    for item in reversed(all_items):
        try:
            if only_with_reminder and not item.ReminderSet:
                continue
            if subject_contains and subject_contains.lower() not in (item.Subject or "").lower():
                continue
            if body_contains and body_contains.lower() not in (item.Body or "").lower():
                continue

            item.Delete()
            deleted_count += 1
        except Exception as e:
            print(f"Erro ao excluir item: {e}")

    return deleted_count
