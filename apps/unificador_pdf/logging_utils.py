"""
Sistema de logging para o Unificador.
"""

import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, Callable

# Configurar logger
logger = logging.getLogger('unificador')


class LogHandler:
    """
    Handler de logs que envia para múltiplos destinos.
    """
    
    def __init__(self, queue_callback: Optional[Callable] = None):
        """
        Args:
            queue_callback: Função para enviar logs para a UI
        """
        self.queue_callback = queue_callback
        self._setup_file_logger()
    
    def _setup_file_logger(self):
        """Configura o logger para arquivo."""
        log_dir = Path(__file__).parent.parent / 'logs'
        log_dir.mkdir(exist_ok=True)
        
        log_file = log_dir / f"unificador_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        # Configurar formato detalhado para arquivo
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(name)s | %(message)s',
            datefmt='%H:%M:%S'
        ))
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(logging.Formatter(
            '%(levelname)s: %(message)s'
        ))
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        logger.setLevel(logging.DEBUG)
        
        self.log_file = log_file
        logger.info(f"📁 Log salvo em: {log_file}")
    
    def debug(self, msg: str):
        """Log de debug (só arquivo)."""
        logger.debug(msg)
    
    def info(self, msg: str, to_ui: bool = False):
        """Log informativo."""
        logger.info(msg)
        if to_ui and self.queue_callback:
            self.queue_callback(('log', f"ℹ️ {msg}"))
    
    def warning(self, msg: str, to_ui: bool = True):
        """Log de aviso."""
        logger.warning(msg)
        if to_ui and self.queue_callback:
            self.queue_callback(('log', f"⚠️ {msg}"))
    
    def error(self, msg: str, to_ui: bool = True):
        """Log de erro."""
        logger.error(msg)
        if to_ui and self.queue_callback:
            self.queue_callback(('log', f"✗ {msg}"))
    
    def success(self, msg: str, to_ui: bool = True):
        """Log de sucesso."""
        logger.info(f"✓ {msg}")
        if to_ui and self.queue_callback:
            self.queue_callback(('log', f"✓ {msg}"))


# Logger global (será configurado pela app)
_log_handler: Optional[LogHandler] = None


def get_logger() -> LogHandler:
    """Retorna o handler de log global."""
    global _log_handler
    if _log_handler is None:
        _log_handler = LogHandler()
    return _log_handler


def setup_logger(queue_callback: Callable) -> LogHandler:
    """Configura o logger com callback para UI."""
    global _log_handler
    _log_handler = LogHandler(queue_callback)
    return _log_handler
