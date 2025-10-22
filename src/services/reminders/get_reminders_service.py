import win32com.client
from datetime import datetime, timedelta

def get(
    folder_type: str = "calendar",
    start_date: datetime = None,
    end_date: datetime = None,
    only_with_reminder: bool = True,
    include_recurring: bool = True,
    sort_field: str = "[Start]"
) -> list:
    # Conecta ao Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    if folder_type.lower() == "calendar":
        folder = outlook.GetDefaultFolder(9)  # 9 = Calendar
        items = folder.Items
        items.Sort(sort_field)
        items.IncludeRecurrences = include_recurring

        start_date = start_date or datetime.now()
        # 🔧 Agora, se o usuário não passar end_date, define como 5 anos à frente
        end_date = end_date or (start_date + timedelta(days=365 * 5))

        restriction = f"[Start] >= '{start_date.strftime('%m/%d/%Y %H:%M %p')}' AND [Start] <= '{end_date.strftime('%m/%d/%Y %H:%M %p')}'"
        restricted_items = items.Restrict(restriction)

        results = []
        for item in restricted_items:
            if not only_with_reminder or item.ReminderSet:
                results.append({
                    "subject": item.Subject,
                    "start": item.Start,
                    "end": item.End,
                    "reminder_set": item.ReminderSet,
                    "reminder_minutes": item.ReminderMinutesBeforeStart if item.ReminderSet else None,
                    "body": item.Body.strip() if item.Body else ""
                })
        return results

    elif folder_type.lower() == "tasks":
        folder = outlook.GetDefaultFolder(13)  # 13 = Tasks
        items = folder.Items
        items.Sort("[DueDate]")

        results = []
        for item in items:
            if not only_with_reminder or item.ReminderSet:
                results.append({
                    "subject": item.Subject,
                    "due_date": getattr(item, "DueDate", None),
                    "reminder_set": item.ReminderSet,
                    "reminder_time": getattr(item, "ReminderTime", None),
                    "body": item.Body.strip() if item.Body else ""
                })
        return results

    else:
        raise ValueError("folder_type deve ser 'calendar' ou 'tasks'")
