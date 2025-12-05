import logging
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional, Callable

# Use config to get log dir if available, otherwise default relative
try:
    from .config import LOG_DIR
except ImportError:
    LOG_DIR = Path(__file__).parent.parent / 'logs'

class LogHandler:
    """
    Unified logging handler for EGS Suite.
    """
    
    def __init__(self, app_name: str, queue_callback: Optional[Callable] = None):
        """
        Args:
            app_name: Name of the application (e.g., 'unificador', 'enviador')
            queue_callback: Optional function to send logs to UI
        """
        self.logger = logging.getLogger(app_name)
        self.queue_callback = queue_callback
        self.app_name = app_name
        
        # Ensure log directory exists
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        
        self._setup_handlers()

    def _setup_handlers(self):
        # Clear existing handlers to prevent duplicates
        if self.logger.hasHandlers():
            self.logger.handlers.clear()

        # File Handler
        timestamp = datetime.now().strftime('%Y%m%d')
        self.log_file = LOG_DIR / f"{self.app_name}_{timestamp}.log"
        
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s | %(levelname)-8s | [%(name)s] | %(message)s',
            datefmt='%H:%M:%S'
        ))
        
        # Console Handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(logging.Formatter(
            '%(levelname)s: %(message)s'
        ))
        
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        self.logger.setLevel(logging.DEBUG)
        
        # Fallback: Configure root logger if it has no handlers
        # This ensures legacy code using logging.info() generic calls still works
        root_logger = logging.getLogger()
        if not root_logger.hasHandlers():
            root_logger.addHandler(file_handler)
            root_logger.addHandler(console_handler)
            root_logger.setLevel(logging.INFO)
            
        self.info(f"Initialized logging for {self.app_name}. Log file: {self.log_file}")

    def debug(self, msg: str):
        self.logger.debug(msg)

    def info(self, msg: str, to_ui: bool = False):
        self.logger.info(msg)
        if to_ui and self.queue_callback:
            self.queue_callback(('log', f"INFO: {msg}"))

    def warning(self, msg: str, to_ui: bool = True):
        self.logger.warning(msg)
        if to_ui and self.queue_callback:
            self.queue_callback(('log', f"WARN: {msg}"))

    def error(self, msg: str, to_ui: bool = True):
        self.logger.error(msg)
        if to_ui and self.queue_callback:
            self.queue_callback(('log', f"ERROR: {msg}"))

    def success(self, msg: str, to_ui: bool = True):
        self.logger.info(f"SUCCESS: {msg}")
        if to_ui and self.queue_callback:
            self.queue_callback(('log', f"SUCCESS: {msg}"))

# Singleton storage
_loggers = {}

def get_logger(app_name: str = 'root') -> LogHandler:
    if app_name not in _loggers:
        _loggers[app_name] = LogHandler(app_name)
    return _loggers[app_name]

def setup_logger(app_name: str, queue_callback: Callable) -> LogHandler:
    _loggers[app_name] = LogHandler(app_name, queue_callback)
    return _loggers[app_name]
