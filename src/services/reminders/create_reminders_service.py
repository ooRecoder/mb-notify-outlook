import win32com.client
from datetime import datetime, timedelta

def create(
    folder_type: str = "calendar",              # 'calendar' ou 'tasks'
    subject: str = "Lembrete",
    body: str = "",
    start_time: datetime = None,
    end_time: datetime = None,
    due_date: datetime = None,
    reminder_minutes_before: int = 15,          # minutos antes do início
    reminder_set: bool = True,
    categories: str = None,
    location: str = None,
    is_all_day: bool = False,
    recurrence: dict = None,                    # Ex: {"type": "daily", "interval": 1, "count": 10}
    mark_completed: bool = False
):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Escolhe a pasta correta
    if folder_type.lower() == "calendar":
        item = outlook.CreateItem(1)  # 1 = olAppointmentItem

        # Define campos básicos
        item.Subject = subject
        item.Body = body

        # Define horário
        start_time = start_time or datetime.now() + timedelta(minutes=1)
        end_time = end_time or (start_time + timedelta(hours=1))

        item.Start = start_time
        item.End = end_time
        item.AllDayEvent = is_all_day

        # Local
        if location:
            item.Location = location

        # Categoria (cor/etiqueta)
        if categories:
            item.Categories = categories

        # Lembrete
        item.ReminderSet = reminder_set
        if reminder_set:
            item.ReminderMinutesBeforeStart = reminder_minutes_before

        # Recorrência
        if recurrence:
            pattern = item.GetRecurrencePattern()
            t = recurrence.get("type", "").lower()
            if t == "daily":
                pattern.RecurrenceType = 0  # olRecursDaily
            elif t == "weekly":
                pattern.RecurrenceType = 1  # olRecursWeekly
            elif t == "monthly":
                pattern.RecurrenceType = 2  # olRecursMonthly
            elif t == "yearly":
                pattern.RecurrenceType = 5  # olRecursYearly
            else:
                raise ValueError("Tipo de recorrência inválido.")
            pattern.Interval = recurrence.get("interval", 1)
            if "count" in recurrence:
                pattern.Occurrences = recurrence["count"]

        # Salva e exibe
        item.Save()
        return f"✅ Evento '{subject}' criado no calendário ({start_time.strftime('%d/%m %H:%M')})."

    elif folder_type.lower() == "tasks":
        item = outlook.CreateItem(3)  # 3 = olTaskItem
        item.Subject = subject
        item.Body = body
        if due_date:
            item.DueDate = due_date

        item.ReminderSet = reminder_set
        if reminder_set:
            item.ReminderTime = (due_date or datetime.now() + timedelta(hours=1)) - timedelta(minutes=reminder_minutes_before)

        if mark_completed:
            item.MarkComplete()

        if categories:
            item.Categories = categories

        item.Save()
        return f"✅ Tarefa '{subject}' criada (vencimento: {due_date.strftime('%d/%m %H:%M') if due_date else 'sem data'})."

    else:
        raise ValueError("folder_type deve ser 'calendar' ou 'tasks'")
