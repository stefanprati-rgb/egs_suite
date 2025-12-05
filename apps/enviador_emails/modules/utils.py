import re
import pandas as pd

def normalizar_uc(uc):
    """Normaliza o código da UC removendo caracteres não numéricos."""
    if pd.isna(uc): return None
    return re.sub(r'[^0-9]', '', str(uc))
