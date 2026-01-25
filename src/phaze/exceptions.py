class PowerBIError(Exception):
    """Base para todos os erros da biblioteca."""
    pass

class LocalAutomationError(PowerBIError):
    """Erros específicos da automação local (Desktop)."""
    pass