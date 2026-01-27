class PhazeError(Exception):
    """Erro base para a biblioteca Phaze."""
    pass

class LocalAutomationError(PhazeError):
    """Erro durante a execucao de automacao local (timeout ou falha visual)."""
    pass

class ConnectionError(PhazeError):
    """Erro ao tentar conectar ao processo do Power BI."""
    pass