"""
Sistema de logging para o Unificador com suporte a cores no console.
"""

import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, Callable

try:
    from colorama import init, Fore, Back, Style
    init(autoreset=True)  # Inicializa colorama para Windows
    COLORAMA_AVAILABLE = True
except ImportError:
    COLORAMA_AVAILABLE = False
    # Fallback: sem cores
    class Fore:
        GREEN = YELLOW = RED = CYAN = MAGENTA = BLUE = WHITE = RESET = ""
    class Style:
        BRIGHT = RESET_ALL = ""

# Configurar logger
logger = logging.getLogger('unificador')


class ColoredConsoleFormatter(logging.Formatter):
    """Formatter que adiciona cores aos logs do console."""
    
    COLORS = {
        'DEBUG': Fore.CYAN,
        'INFO': Fore.WHITE,
        'WARNING': Fore.YELLOW,
        'ERROR': Fore.RED,
        'CRITICAL': Fore.RED + Style.BRIGHT,
    }
    
    ICONS = {
        'DEBUG': '🔍',
        'INFO': 'ℹ️ ',
        'WARNING': '⚠️ ',
        'ERROR': '❌',
        'CRITICAL': '🔥',
    }
    
    def format(self, record):
        # Adiciona cor baseada no nível
        color = self.COLORS.get(record.levelname, '')
        icon = self.ICONS.get(record.levelname, '')
        
        # Formata a mensagem
        if COLORAMA_AVAILABLE:
            return f"{color}{icon} {record.getMessage()}{Style.RESET_ALL}"
        else:
            return f"{icon} {record.getMessage()}"


class LogHandler:
    """
    Handler de logs que envia para múltiplos destinos com suporte a cores.
    """
    
    def __init__(self, queue_callback: Optional[Callable] = None):
        """
        Args:
            queue_callback: Função para enviar logs para a UI
        """
        self.queue_callback = queue_callback
        self._setup_file_logger()
    
    def _setup_file_logger(self):
        """Configura o logger para arquivo e console."""
        log_dir = Path(__file__).parent.parent / 'logs'
        log_dir.mkdir(exist_ok=True)
        
        log_file = log_dir / f"unificador_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        # Configurar formato detalhado para arquivo (sem cores)
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(name)s | %(message)s',
            datefmt='%H:%M:%S'
        ))
        
        # Console handler com cores
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(ColoredConsoleFormatter())
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        logger.setLevel(logging.DEBUG)
        
        self.log_file = log_file
        self._print_header()
    
    def _print_header(self):
        """Imprime cabeçalho colorido no console."""
        if COLORAMA_AVAILABLE:
            print(f"\n{Fore.CYAN}{'='*70}{Style.RESET_ALL}")
            print(f"{Fore.CYAN}{Style.BRIGHT}  📄 UNIFICADOR DE FATURAS E BOLETOS EGS{Style.RESET_ALL}")
            print(f"{Fore.CYAN}{'='*70}{Style.RESET_ALL}\n")
            print(f"{Fore.GREEN}📁 Log salvo em: {Fore.WHITE}{self.log_file}{Style.RESET_ALL}\n")
        else:
            print("\n" + "="*70)
            print("  📄 UNIFICADOR DE FATURAS E BOLETOS EGS")
            print("="*70 + "\n")
            print(f"📁 Log salvo em: {self.log_file}\n")
    
    def print_section(self, title: str):
        """Imprime um separador de seção."""
        if COLORAMA_AVAILABLE:
            print(f"\n{Fore.MAGENTA}{Style.BRIGHT}{'─'*70}{Style.RESET_ALL}")
            print(f"{Fore.MAGENTA}{Style.BRIGHT}  {title}{Style.RESET_ALL}")
            print(f"{Fore.MAGENTA}{Style.BRIGHT}{'─'*70}{Style.RESET_ALL}\n")
        else:
            print(f"\n{'─'*70}")
            print(f"  {title}")
            print(f"{'─'*70}\n")
    
    def print_progress(self, current: int, total: int, item_name: str = ""):
        """Imprime barra de progresso."""
        percent = int((current / total) * 100) if total > 0 else 0
        bar_length = 40
        filled = int((bar_length * current) / total) if total > 0 else 0
        bar = '█' * filled + '░' * (bar_length - filled)
        
        if COLORAMA_AVAILABLE:
            print(f"\r{Fore.CYAN}[{bar}] {Fore.WHITE}{percent}% {Fore.GREEN}({current}/{total}){Style.RESET_ALL} {item_name}", end='', flush=True)
        else:
            print(f"\r[{bar}] {percent}% ({current}/{total}) {item_name}", end='', flush=True)
    
    def print_success(self, msg: str):
        """Imprime mensagem de sucesso destacada."""
        if COLORAMA_AVAILABLE:
            print(f"\n{Fore.GREEN}{Style.BRIGHT}✓ {msg}{Style.RESET_ALL}\n")
        else:
            print(f"\n✓ {msg}\n")
    
    def print_error(self, msg: str):
        """Imprime mensagem de erro destacada."""
        if COLORAMA_AVAILABLE:
            print(f"\n{Fore.RED}{Style.BRIGHT}✗ {msg}{Style.RESET_ALL}\n")
        else:
            print(f"\n✗ {msg}\n")
    
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
