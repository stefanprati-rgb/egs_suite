"""
Lógica central de pareamento (Unificação) - VERSÃO SIMPLIFICADA.
Pareia documentos baseado APENAS na UC extraída do nome do arquivo.
"""
import os
from typing import List, Dict
from .extractors.uc_extractor import extrai_uc, normalizar_uc
from .logging_utils import get_logger

log = get_logger()

class Documento:
    def __init__(self, caminho, tipo):
        self.caminho = caminho
        self.nome = os.path.basename(caminho)
        self.tipo = tipo  # 'fatura' ou 'boleto'
        self.uc = None
        self.uc_norm = None
        
        self.processar()

    def processar(self):
        # SIMPLIFICADO: Extrai apenas UC do nome do arquivo
        self.uc = extrai_uc(self.nome)
        
        if self.uc:
            self.uc_norm = normalizar_uc(self.uc)
            log.info(f"[UC] {self.tipo.upper()}: {self.uc} ({self.nome})")
        else:
            log.error(f"[UC] ✗ Não encontrada no nome: {self.nome}")

class UnificadorCore:
    def __init__(self, faturas_paths, boletos_paths):
        """
        Inicializa o unificador.
        SIMPLIFICADO: Não extrai texto dos PDFs, apenas processa os nomes dos arquivos.
        
        Args:
            faturas_paths: Lista de caminhos para arquivos de fatura
            boletos_paths: Lista de caminhos para arquivos de boleto
        """
        log.info(f"Processando {len(faturas_paths)} faturas e {len(boletos_paths)} boletos")
        
        self.docs_faturas = [Documento(c, 'fatura') for c in faturas_paths]
        self.docs_boletos = [Documento(c, 'boleto') for c in boletos_paths]
        
    def processar_pareamento(self):
        pares = []
        nao_pareados = []
        
        # Indexa faturas por UC Normalizada
        mapa_faturas = {}
        for f in self.docs_faturas:
            if f.uc_norm:
                mapa_faturas[f.uc_norm] = f
            else:
                log.warning(f"[PAREAMENTO] Fatura sem UC: {f.nome}")
                nao_pareados.append(f)
                
        # Tenta parear boletos
        for b in self.docs_boletos:
            if not b.uc_norm:
                log.warning(f"[PAREAMENTO] Boleto sem UC: {b.nome}")
                nao_pareados.append(b)
                continue
                
            if b.uc_norm in mapa_faturas:
                fatura_corresp = mapa_faturas[b.uc_norm]
                
                log.info(f"[PAREAMENTO] ✓ PAR FORMADO | UC: {b.uc} | {fatura_corresp.nome} + {b.nome}")
                
                pares.append({
                    'fatura': fatura_corresp,
                    'boleto': b,
                    'uc': fatura_corresp.uc,
                })
                
                # Remove do mapa para evitar duplicidade (1 boleto -> 1 fatura)
                del mapa_faturas[b.uc_norm]
            else:
                log.warning(f"[PAREAMENTO] ✗ Boleto sem fatura correspondente: UC {b.uc} ({b.nome})")
                nao_pareados.append(b)
                
        # Adiciona faturas que sobraram
        for f in mapa_faturas.values():
            log.warning(f"[PAREAMENTO] ✗ Fatura sem boleto correspondente: UC {f.uc} ({f.nome})")
            nao_pareados.append(f)
        
        log.info(f"[PAREAMENTO] RESUMO: {len(pares)} pares formados | {len(nao_pareados)} documentos não pareados")
        return pares, nao_pareados