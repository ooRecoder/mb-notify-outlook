from . import read_configurations as read
from .spreadsheet.spreadsheet_configurations import SpreadsheetConfigurations


class MainConfigurations:
    def __init__(self):
        # Ler JSON
        # Inicializa as classes de configuração com informações extraidas do JSON
        configs = {} 
        
        self.spreadsheets = SpreadsheetConfigurations(configs.get("spreadsheets", {}))
        self.layout = configs.get("layout", {})
        
        pass
    def save(self):
        
        return
    pass