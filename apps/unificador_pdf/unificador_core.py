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
        log = get_logger()
        
        # 1. Extrai UC do Nome (Mais confiável se o arquivo foi renomeado corretamente)
        uc_nome = extrai_uc(self.nome)
        
        # 2. Extrai UC do Texto
        uc_texto = extrai_uc_do_texto(self.texto, self.nome)
        
        # 3. Validação Cruzada (Camada 2)
        if uc_nome and uc_texto:
            uc_nome_norm = normalizar_uc(uc_nome)
            uc_texto_norm = normalizar_uc(uc_texto)
            
            if uc_nome_norm == uc_texto_norm:
                log.info(f"[VALIDAÇÃO] ✓ UC do nome coincide com UC do texto: {uc_nome} ({self.nome})")
                self.uc = uc_nome  # Usa formato original do nome
                self.uc_norm = uc_nome_norm
            else:
                log.warning(f"[VALIDAÇÃO] ✗ UC divergente! Nome: {uc_nome} vs Texto: {uc_texto} ({self.nome})")
                # Prioriza UC do texto (mais confiável)
                self.uc = uc_texto
                self.uc_norm = uc_texto_norm
        elif uc_nome:
            log.info(f"[VALIDAÇÃO] UC apenas no nome: {uc_nome} ({self.nome})")
            self.uc = uc_nome
            self.uc_norm = normalizar_uc(uc_nome)
        elif uc_texto:
            log.info(f"[VALIDAÇÃO] UC apenas no texto: {uc_texto} ({self.nome})")
            self.uc = uc_texto
            self.uc_norm = normalizar_uc(uc_texto)
        else:
            log.error(f"[VALIDAÇÃO] ✗ Nenhuma UC encontrada em {self.nome}")
            
        # 4. Extrai Valores e Referência
        if self.tipo == 'fatura':
            self.valor = extrair_valor_fatura(self.texto, self.nome)
        else:
            self.valor = extrair_valor_boleto(self.texto, self.nome)
            
        self.referencia = extrai_referencia(self.texto, self.nome)
        
        # Log resumo
        log.info(f"[PROCESSADO] {self.tipo.upper()} | UC: {self.uc} | Valor: R$ {self.valor} | Ref: {self.referencia} | {self.nome}")

class UnificadorCore:
    def __init__(self, faturas_raw, boletos_raw):
        # faturas_raw e boletos_raw são listas de tuplas (caminho, texto)
        self.docs_faturas = [Documento(c, 'fatura', t) for c, t in faturas_raw]
        self.docs_boletos = [Documento(c, 'boleto', t) for c, t in boletos_raw]
        
    def processar_pareamento(self):
        log = get_logger()
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
                
                # Camada 3: Validação de Valor (Exata - sem tolerância)
                match_valor = False
                if fatura_corresp.valor and b.valor:
                    diff = abs(fatura_corresp.valor - b.valor)
                    match_valor = diff < 0.01  # Apenas tolerância para arredondamento
                    
                    if match_valor:
                        log.info(f"[PAREAMENTO] ✓ Valores coincidem: R$ {fatura_corresp.valor} == R$ {b.valor}")
                    else:
                        log.error(f"[PAREAMENTO] ✗ VALORES DIVERGENTES! Fatura: R$ {fatura_corresp.valor} vs Boleto: R$ {b.valor} (Diff: R$ {diff:.2f})")
                else:
                    log.warning(f"[PAREAMENTO] ⚠ Valor ausente - Fatura: {fatura_corresp.valor} | Boleto: {b.valor}")
                
                # Camada 4: Validação de Período
                match_periodo = False
                if fatura_corresp.referencia and b.referencia:
                    match_periodo = fatura_corresp.referencia == b.referencia
                    if match_periodo:
                        log.info(f"[PAREAMENTO] ✓ Períodos coincidem: {fatura_corresp.referencia}")
                    else:
                        log.warning(f"[PAREAMENTO] ⚠ Períodos divergentes: Fatura: {fatura_corresp.referencia} vs Boleto: {b.referencia}")
                
                # Score de confiança
                confianca = 0
                if match_valor: confianca += 50
                if match_periodo: confianca += 30
                confianca += 20  # UC match (base)
                
                log.info(f"[PAREAMENTO] ✓ PAR FORMADO | UC: {b.uc} | Confiança: {confianca}% | {fatura_corresp.nome} + {b.nome}")
                
                pares.append({
                    'fatura': fatura_corresp,
                    'boleto': b,
                    'uc': fatura_corresp.uc,
                    'valor_match': match_valor,
                    'periodo_match': match_periodo,
                    'confianca': confianca
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