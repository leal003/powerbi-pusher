from .local_ops import Phaze
from .exceptions import PhazeError, LocalAutomationError, ConnectionError

__version__ = "1.1.4"
__all__ = ["Phaze", "PhazeError", "LocalAutomationError", "ConnectionError"]