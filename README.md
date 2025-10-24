# 🧭 Visão Geral

Este projeto tem como objetivo **automatizar a leitura de planilhas Excel (`.xlsx`)**, aplicar **filtros personalizados**, e **registrar lembretes no Outlook** com base nos dados processados.

Todas as configurações do sistema (leituras, filtros, layout, etc.) são **armazenadas em um arquivo JSON**, permitindo reutilização e fácil modificação.

---

## 📋 Requisitos

* Todas as configurações devem ser salvas em um arquivo `.json`.
* A aplicação deve permitir:

  * Visualizar e editar as configurações existentes.
  * Configurar parâmetros de leitura de planilhas.
  * Definir filtros de dados.
  * Personalizar o layout da interface.

---

## ⚙️ Fluxo de Uso

1. O usuário abre a interface da aplicação.
2. A interface oferece opções para:

   * Verificar e editar configurações.
   * Escolher o arquivo `.xlsx` para leitura.
   * Selecionar um filtro a ser aplicado (opcional).
   * Registrar lembretes no Outlook.

---

## 🧠 Funcionalidades do Sistema

### 1. Configurações de Leitura (`read_config`)

* Define **como cada planilha** do arquivo `.xlsx` deve ser lida.
* Permite:

  * Especificar quais **colunas** serão lidas (filtro de colunas).
  * Nomear planilhas e mapear colunas específicas.
* Exemplo de configuração (em JSON):

  ```json
  {
    "NameConfiguration": {
      "Planilha1": {
        "skip_line": 1,
        "columns": ["Nome", "Data", "Descrição"]
      }
    }
  }
  ```

---

### 2. Aplicação de Filtros (`filter_config`)

* Opcional.
* Permite remover ou manter linhas com base em valores específicos.
* Exemplo:

  ```json
  {
    "filters": [
      { "column": "Status", "exclude": ["Cancelado", "Inativo"] }
    ]
  }
  ```

---

### 3. Configuração de Lembretes (`reminder_config`)

* Define como os lembretes serão criados no Outlook.
* Permite:

  * Definir **título** e **mensagem** usando valores da planilha.
  * Especificar **tempo de disparo** e **recorrência**.
  * Outras opções como prioridade ou categoria.

Exemplo:

```json
{
  "reminder": {
    "title": "{Nome} - {Data}",
    "message": "Lembrete automático para {Descrição}",
    "time_offset_minutes": 30,
    "recurrence": "daily"
  }
}
```

---

### 4. Registro no Outlook

* O sistema cria os lembretes conforme a configuração.
* Pode usar automação (por exemplo, via `win32com.client`) para registrar os compromissos diretamente no Outlook.

---

## 🚀 Possíveis Extensões Futuras

* Suporte a múltiplos arquivos `.xlsx` de uma vez.
* Geração de logs de execução.
* Interface web ou desktop (ex: com PyQt ou Streamlit).
* Exportação de relatórios com os lembretes criados.
