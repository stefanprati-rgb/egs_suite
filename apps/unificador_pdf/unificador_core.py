"""
Lógica central de pareamento (Unificação).
"""
import os
from typing import List, Dict
from .extractors.uc_extractor import extrai_uc, extrai_uc_do_texto, normalizar_uc, extrai_referencia
from .extractors.value_extractor import extrair_valor_fatura, extrair_valor_boleto
from ..logging_utils import get_logger

log = get_logger()

class Documento:
    def __init__(self, caminho, tipo, texto_extraido):
        self.caminho = caminho
        self.nome = os.path.basename(caminho)
        self.tipo = tipo # 'fatura' ou 'boleto'
        self.texto = texto_extraido
        self.uc = None
        self.uc_norm = None
        self.valor = None
        self.referencia = None
        
        self.processar()

    def processar(self):
        # 1. Tenta UC do Nome (Mais confiável se o arquivo foi renomeado corretamente)
        self.uc = extrai_uc(self.nome)
        
        # 2. Se não achou no nome, tenta no texto
        if not self.uc:
            self.uc = extrai_uc_do_texto(self.texto, self.nome)
            
        if self.uc:
            self.uc_norm = normalizar_uc(self.uc)
            
        # 3. Extrai Valores e Referência
        if self.tipo == 'fatura':
            self.valor = extrair_valor_fatura(self.texto, self.nome)
        else:
            self.valor = extrair_valor_boleto(self.texto, self.nome)
            
        self.referencia = extrai_referencia(self.texto, self.nome)

class UnificadorCore:
    def __init__(self, faturas_raw, boletos_raw):
        # faturas_raw e boletos_raw são listas de tuplas (caminho, texto)
        self.docs_faturas = [Documento(c, 'fatura', t) for c, t in faturas_raw]
        self.docs_boletos = [Documento(c, 'boleto', t) for c, t in boletos_raw]
        
    def processar_pareamento(self):
        pares = []
        nao_pareados = []
        
        # Indexa faturas por UC Normalizada
        mapa_faturas = {}
        for f in self.docs_faturas:
            if f.uc_norm:
                mapa_faturas[f.uc_norm] = f
            else:
                nao_pareados.append(f)
                
        # Tenta parear boletos
        for b in self.docs_boletos:
            if b.uc_norm and b.uc_norm in mapa_faturas:
                fatura_corresp = mapa_faturas[b.uc_norm]
                
                # Validação de Valor (Opcional: Apenas loga aviso se diferente)
                match_valor = False
                if fatura_corresp.valor and b.valor:
                    diff = abs(fatura_corresp.valor - b.valor)
                    match_valor = diff < 1.0 # Tolerância de 1 real
                    
                pares.append({
                    'fatura': fatura_corresp,
                    'boleto': b,
                    'uc': f.uc,
                    'valor_match': match_valor
                })
                
                # Remove do mapa para evitar duplicidade (1 boleto -> 1 fatura)
                del mapa_faturas[b.uc_norm]
            else:
                nao_pareados.append(b)
                
        # Adiciona faturas que sobraram
        nao_pareados.extend(mapa_faturas.values())
        
        return pares, nao_pareados