from .archive_spreadsheet_configurations import ArchiveConfigurations
# Arquivos spreadsheets
class SpreadsheetConfigurations():
    def __init__(self, cache: list):
        self.cache = cache
        pass
    def _exist(self, config_name):
        #Verifica se existe
        config = True 
        if config:
            return True
        return False
    def get(self, config_name: str) -> list:
        message: str = None
        success = self._exist(config_name)
        if success:
            
            config = {} # Faz a busca e retorna o valor de config
            
            return {success, config}
        message = f"Configuração {config_name} não encontrada"
        return {success, message}
    def add(self, config_name: str, config: list, extension: str ="xlsx"):
        message = None
        success = False
        if self._exist(config_name):
            return {success, message}
        return {success, message}
    def remove(self, config_name: str):
        # Valida se existe uma configuração com o mesmo nome
        # Retorna {success: bool, message: str}
        return
    def change(self, config_name: str, changes):
        # Valida se existe uma configuração com o mesmo nome
        return
    pass