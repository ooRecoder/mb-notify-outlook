from services import reminders
from datetime import datetime

def get(
    folder_type: str = "calendar",
    start_date: datetime = None,
    end_date: datetime = None,
    only_with_reminder: bool = True,
    include_recurring: bool = True,
    sort_field: str = "[Start]"
):
    """
    Lê lembretes do Outlook (Calendário ou Tarefas) e retorna uma lista de dicionários.

    Parâmetros:
    ----------
    folder_type : str
        'calendar' para compromissos ou 'tasks' para tarefas.
    start_date : datetime
        Data inicial do filtro. Padrão: hoje.
    end_date : datetime
        Data final do filtro. Padrão: 5 anos após o start_date (todos os futuros).
    only_with_reminder : bool
        Se True, retorna apenas itens com lembrete ativo.
    include_recurring : bool
        Se True, inclui eventos recorrentes.
    sort_field : str
        Campo usado para ordenação (ex: '[Start]', '[DueDate]').

    Retorna:
    -------
    list
        Lista de dicionários com os campos:
        - subject
        - start / due_date
        - end
        - reminder_set
        - reminder_minutes
        - body
    """
    return reminders.get(
        folder_type,
        start_date,
        end_date,
        only_with_reminder,
        include_recurring,
        sort_field
    )


def create(
    folder_type: str = "calendar",
    subject: str = "Lembrete",
    body: str = "",
    start_time: datetime = None,
    end_time: datetime = None,
    due_date: datetime = None,
    reminder_minutes_before: int = 15,
    reminder_set: bool = True,
    categories: str = None,
    location: str = None,
    is_all_day: bool = False,
    recurrence: dict = None,
    mark_completed: bool = False
):
    """
    Cria um lembrete no Outlook (Calendário ou Tarefa).

    Parâmetros:
    ----------
    folder_type : str
        'calendar' para compromissos ou 'tasks' para tarefas.

    subject : str
        Título do lembrete.

    body : str
        Descrição do lembrete.

    start_time : datetime
        Data/hora de início (usado para eventos).

    end_time : datetime
        Data/hora de término (usado para eventos).

    due_date : datetime
        Data de vencimento (usado para tarefas).

    reminder_minutes_before : int
        Quantos minutos antes o lembrete deve disparar.

    reminder_set : bool
        Define se o lembrete deve estar ativo.

    categories : str
        Categoria de cor/etiqueta (ex: 'Trabalho', 'Pessoal').

    location : str
        Localização (para compromissos).

    is_all_day : bool
        Define se é um evento de dia inteiro.

    recurrence : dict
        Define recorrência (opcional). Exemplo:
        {
            "type": "daily",      # daily, weekly, monthly, yearly
            "interval": 1,        # a cada X dias/semanas/meses
            "count": 5            # número de repetições
        }

    mark_completed : bool
        Se True, marca a tarefa como concluída após criar.

    Retorna:
    -------
    str
        Mensagem de sucesso com detalhes do lembrete criado.
    """
    return reminders.create(
        folder_type=folder_type,
        subject=subject,
        body=body,
        start_time=start_time,
        end_time=end_time,
        due_date=due_date,
        reminder_minutes_before=reminder_minutes_before,
        reminder_set=reminder_set,
        categories=categories,
        location=location,
        is_all_day=is_all_day,
        recurrence=recurrence,
        mark_completed=mark_completed
    )

def delete(
    folder_type: str = "calendar",
    start_date: datetime = None,
    end_date: datetime = None,
    subject_contains: str = None,
    body_contains: str = None,
    only_with_reminder: bool = False,
    include_recurring: bool = True,
) -> int:
    return reminders.delete(folder_type=folder_type, start_date=start_date, end_date=end_date, 
                            subject_contains=subject_contains, body_contains=body_contains,
                            only_with_reminder=only_with_reminder, include_recurring=include_recurring)