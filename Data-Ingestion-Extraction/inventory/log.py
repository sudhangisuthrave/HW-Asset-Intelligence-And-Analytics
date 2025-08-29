"""Base class with logging capability
"""

import logging


class LoggingBase:
    """Base class with logging capability"""

    def __init__(self) -> None:
        self.logger = logging.getLogger(
            f'{__name__}.{self.__class__.__name__}',
        )
