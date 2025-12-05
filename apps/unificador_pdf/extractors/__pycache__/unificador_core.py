"""
Lógica central para processamento e unificação de documentos.
Define a estrutura de dados e o algoritmo de pareamento.
"""

from typing import List, Dict, Optional, Tuple
from pathlib import Path
import re

# Importa as funções de utilidade
# Assumimos que estas classes/funções existem no mesmo pacote 'unificador_pdf'
from ..logging_utils import get_logger 
from .extractors.uc_extractor import extrai_uc, extrai_referencia
from .extractors.value_extractor import extrair_valor_fatura, extrair_valor_boleto

log = get_logger()

class Documento:
    """Estrutura para armazenar informações extraídas do PDF."""
    def __init__(self, caminho: Path, tipo: str, texto: str):
        self.caminho = caminho
        self.nome = caminho.name
        self.tipo = tipo # 'fatura' ou 'boleto'
        self.texto = texto
        self.uc: Optional[str] = None
        self.valor: Optional[float] = None
        self.referencia: Optional[str] = None

    def __repr__(self):
        valor_str = f"{self.valor:.2f}" if self.valor is not None else "None"
        return (f"Documento(tipo='{self.tipo}', uc='{self.uc}', "
                f"valor={valor_str}, nome='{self.nome}')")

    @staticmethod
    def normalizar_uc(uc_str: Optional[str]) -> Optional[str]:
        """Remove caracteres não-dígitos da UC para padronização de pareamento."""
        if uc_str:
            # Mantém apenas dígitos (ex: '10/10232-7' vira '10102327')
            return re.sub(r'[^\d]', '', uc_str)
        return None

    def extrair_dados(self):
        """Extrai UC, Valor e Referência com base no tipo de documento."""
        
        # 1. Extração da UC (prioriza o nome do arquivo)
        self.uc = extrai_uc(self.nome)
        if not self.uc:
            # Tenta do texto se não encontrar no nome
            self.uc = extrai_uc_do_texto(self.texto, self.nome)
        
        # 2. Extração de Valor e Referência
        if self.tipo == 'fatura':
            self.valor = extrair_valor_fatura(self.texto, self.nome)
            self.referencia = extrai_referencia(self.texto, self.nome)
        elif self.tipo == 'boleto':
            self.valor = extrair_valor_boleto(self.texto, self.nome)
            # Para o boleto EGS, a referência é menos crítica para o pareamento inicial

        log.info(f"Dados extraídos de {self.nome}: UC={self.uc}, Valor={self.valor}, Ref={self.referencia}")

class UnificadorCore:
    """Implementa a lógica de pareamento (Unificação) dos documentos."""

    def __init__(self, documentos: List[Documento]):
        self.documentos = documentos
        self.faturas: List[Documento] = [d for d in documentos if d.tipo == 'fatura']
        self.boletos: List[Documento] = [d for d in documentos if d.tipo == 'boleto']
        self.pares: List[Dict] = []
        self.nao_pareados: List[Documento] = []

    def unificar_documentos(self) -> Tuple[List[Dict], List[Documento]]:
        """
        Executa a lógica de pareamento.
        Prioriza o pareamento pela UC normalizada (o elo mais forte).
        A validação de valor é uma checagem secundária.
        """
        log.info(f"Iniciando unificação. Faturas: {len(self.faturas)}, Boletos: {len(self.boletos)}")
        
        # 1. Agrupar Faturas por UC Normalizada
        # Usa um dicionário para garantir que cada UC tenha no máximo uma fatura como chave
        faturas_por_uc: Dict[str, Documento] = {}
        for fatura in self.faturas:
            uc_normalizada = Documento.normalizar_uc(fatura.uc)
            if uc_normalizada:
                # Trata duplicatas (mantém a primeira ou implementa lógica de desempate)
                if uc_normalizada in faturas_por_uc:
                    log.warning(f"UC duplicada encontrada: {uc_normalizada}. {fatura.nome} será ignorada em favor de {faturas_por_uc[uc_normalizada].nome}")
                else:
                    faturas_por_uc[uc_normalizada] = fatura
            else:
                self.nao_pareados.append(fatura)
                log.warning(f"Fatura não pareada (UC não encontrada): {fatura.nome}")
        
        # 2. Tentar parear Boletos com Faturas
        boletos_restantes = []
        for boleto in self.boletos:
            uc_normalizada = Documento.normalizar_uc(boleto.uc)
            
            if uc_normalizada and uc_normalizada in faturas_por_uc:
                fatura = faturas_por_uc[uc_normalizada]
                
                # Checagem de Validação de Valor:
                valor_ok = False
                if fatura.valor is not None and boleto.valor is not None:
                    # Permite uma pequena margem de erro (R$ 0.01) para evitar erros de ponto flutuante
                    if abs(fatura.valor - boleto.valor) < 0.02:
                        valor_ok = True
                    else:
                        # Loga o mismatch, mas o pareamento é mantido pela UC
                        log.warning(
                            f"Pareamento por UC '{uc_normalizada}' com valor divergente: "
                            f"Fatura R${fatura.valor:.2f} vs Boleto R${boleto.valor:.2f}"
                        )
                elif fatura.valor is None or boleto.valor is None:
                    log.warning(f"Pareamento por UC '{uc_normalizada}' com valor(es) ausente(s).")

                # Cria o par, priorizando a UC como link
                self.pares.append({
                    'uc': fatura.uc,
                    'fatura_caminho': fatura.caminho,
                    'boleto_caminho': boleto.caminho,
                    'fatura_valor': fatura.valor,
                    'boleto_valor': boleto.valor,
                    'valores_coincidem': valor_ok,
                    'fatura_referencia': fatura.referencia
                })
                
                # Remove a fatura do pool (assumindo relação 1:1)
                del faturas_por_uc[uc_normalizada]

            else:
                boletos_restantes.append(boleto)
                log.warning(f"Boleto não pareado (UC não encontrada ou Fatura já utilizada): {boleto.nome}")

        # Documentos não pareados (restantes no dicionário de faturas e boletos restantes)
        self.nao_pareados.extend(faturas_por_uc.values())
        self.nao_pareados.extend(boletos_restantes)

        log.info(f"Unificação concluída. Pares encontrados: {len(self.pares)}, Não pareados: {len(self.nao_pareados)}")
        return self.pares, self.nao_pareados