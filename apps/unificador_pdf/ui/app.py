"""
Classe principal da aplicação Unificador.
Inclui logging robusto para diagnóstico de problemas.
"""

import os
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
import threading
import queue
import zipfile

from ..config import COLORS, FONTS, WINDOW_CONFIG
from ..extractors import (
    extrai_uc, extrai_uc_do_texto, 
    extrair_valor_fatura, extrair_valor_boleto,
    analisar_texto_pdf, analisar_valores_pdf
)
from ..pdf import validar_pdf_cabecalho, abrir_pdf_seguro, extrair_texto_pdf, unir_pdfs, criar_nome_arquivo
from EGS_Suite.common.logging import setup_logger, get_logger
from .styles import configurar_estilos
from .components import (
    criar_card,
    criar_seletor_pasta,
    criar_botao_acao,
    criar_area_relatorio
)


class AppUnificador:
    """Aplicação principal do Unificador de Faturas e Boletos."""
    
    def __init__(self, root: tk.Tk):
        """
        Inicializa a aplicação.
        
        Args:
            root: Janela principal Tkinter
        """
        self.root = root
        self._configurar_janela()
        self.cores = configurar_estilos()
        
        # Variáveis de estado
        self.pasta_faturas = tk.StringVar()
        self.pasta_boletos = tk.StringVar()
        self.ordem_pdf = tk.StringVar(value="fatura_primeiro")
        self.contador_faturas = tk.StringVar(value="0 arquivos")
        self.contador_boletos = tk.StringVar(value="0 arquivos")
        
        # Estatísticas de processamento
        self.stats = {
            'faturas_total': 0,
            'faturas_sem_uc': 0,
            'faturas_sem_valor': 0,
            'faturas_erro': 0,
            'faturas_ok': 0,
            'boletos_total': 0,
            'boletos_sem_uc': 0,
            'boletos_sem_valor': 0,
            'boletos_erro': 0,
            'boletos_ok': 0,
        }
        
        # Threading
        self._worker_thread = None
        self._queue = queue.Queue()
        self._cancelar = threading.Event()
        
        # Construir interface
        self._criar_interface()
        self.root.after(100, self._processar_queue)
    
    def _configurar_janela(self):
        """Configura a janela principal."""
        self.root.title(WINDOW_CONFIG['title'])
        self.root.geometry(WINDOW_CONFIG['geometry'])
        self.root.minsize(WINDOW_CONFIG['min_width'], WINDOW_CONFIG['min_height'])
    
    def _criar_interface(self):
        """Cria toda a interface gráfica."""
        self.root.configure(bg=COLORS['bg'])
        
        # Container principal
        main = tk.Frame(self.root, bg=COLORS['bg'])
        main.pack(expand=True, fill=tk.BOTH, padx=30, pady=30)
        
        self._criar_header(main)
        self._criar_card_selecao(main)
        self._criar_card_opcoes(main)
        self._criar_botoes_acao(main)
        self._criar_area_progresso(main)
        self._criar_area_relatorio(main)
    
    def _criar_header(self, parent: tk.Widget):
        """Cria o cabeçalho da aplicação."""
        header = tk.Frame(parent, bg=COLORS['bg'])
        header.pack(fill=tk.X, pady=(0, 25))
        
        # Ícone
        tk.Label(
            header,
            text="📄",
            font=FONTS['icon'],
            bg=COLORS['bg']
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        # Textos
        text_frame = tk.Frame(header, bg=COLORS['bg'])
        text_frame.pack(side=tk.LEFT)
        
        ttk.Label(
            text_frame,
            text="Unificador de PDFs",
            style='Title.TLabel'
        ).pack(anchor=tk.W)
        
        ttk.Label(
            text_frame,
            text="Combina faturas e boletos por UC e valor (com logging detalhado)",
            style='Subtitle.TLabel'
        ).pack(anchor=tk.W)
    
    def _criar_card_selecao(self, parent: tk.Widget):
        """Cria o card de seleção de pastas."""
        card = criar_card(parent)
        card.pack(fill=tk.X, pady=(0, 15))
        
        inner = tk.Frame(card, bg=COLORS['card'])
        inner.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        
        tk.Label(
            inner,
            text="📁 Seleção de Arquivos",
            font=FONTS['heading'],
            bg=COLORS['card'],
            fg=COLORS['text']
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # Seletor de faturas
        criar_seletor_pasta(
            inner, "Faturas", self.pasta_faturas, self.contador_faturas,
            self._selecionar_pasta_faturas, "📄", self.cores
        ).pack(fill=tk.X, pady=(0, 12))
        
        # Seletor de boletos
        criar_seletor_pasta(
            inner, "Boletos", self.pasta_boletos, self.contador_boletos,
            self._selecionar_pasta_boletos, "🧾", self.cores
        ).pack(fill=tk.X)
    
    def _criar_card_opcoes(self, parent: tk.Widget):
        """Cria o card de opções."""
        card = criar_card(parent)
        card.pack(fill=tk.X, pady=(0, 15))
        
        inner = tk.Frame(card, bg=COLORS['card'])
        inner.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        
        tk.Label(
            inner,
            text="⚙️ Opções de Unificação",
            font=FONTS['heading'],
            bg=COLORS['card'],
            fg=COLORS['text']
        ).pack(anchor=tk.W, pady=(0, 10))
        
        opcoes = [
            ("fatura_primeiro", "📄 Fatura primeiro, depois Boleto"),
            ("boleto_primeiro", "🧾 Boleto primeiro, depois Fatura"),
        ]
        
        for valor, texto in opcoes:
            tk.Radiobutton(
                inner,
                text=texto,
                variable=self.ordem_pdf,
                value=valor,
                font=FONTS['small'],
                bg=COLORS['card'],
                fg=COLORS['text'],
                selectcolor=COLORS['card'],
                activebackground=COLORS['card'],
                cursor='hand2'
            ).pack(anchor=tk.W, pady=3)
    
    def _criar_botoes_acao(self, parent: tk.Widget):
        """Cria os botões de ação."""
        frame = tk.Frame(parent, bg=COLORS['bg'])
        frame.pack(fill=tk.X, pady=(0, 15))
        
        # Botão Unir
        self.botao_unir = criar_botao_acao(
            frame, "🚀 Unir PDFs com Validação de Valor",
            self._iniciar_processamento,
            COLORS['success'], COLORS['success_hover']
        )
        self.botao_unir.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))
        
        # Botão Cancelar
        self.botao_cancelar = criar_botao_acao(
            frame, "✕ Cancelar",
            self._cancelar_processamento,
            COLORS['error'], COLORS['error_hover'],
            state='disabled', width=15
        )
        self.botao_cancelar.pack(side=tk.LEFT)
    
    def _criar_area_progresso(self, parent: tk.Widget):
        """Cria a área de progresso."""
        card = criar_card(parent)
        card.pack(fill=tk.X, pady=(0, 15))
        
        inner = tk.Frame(card, bg=COLORS['card'])
        inner.pack(padx=20, pady=15, fill=tk.BOTH, expand=True)
        
        self.status_label = tk.Label(
            inner,
            text="Aguardando seleção das pastas...",
            font=FONTS['small'],
            bg=COLORS['card'],
            fg=COLORS['text_light'],
            anchor=tk.W
        )
        self.status_label.pack(fill=tk.X, pady=(0, 8))
        
        self.progresso = ttk.Progressbar(
            inner,
            orient=tk.HORIZONTAL,
            mode='determinate',
            maximum=100
        )
        self.progresso.pack(fill=tk.X)
    
    def _criar_area_relatorio(self, parent: tk.Widget):
        """Cria a área de relatório."""
        card = criar_card(parent)
        card.pack(expand=True, fill=tk.BOTH)
        
        inner = tk.Frame(card, bg=COLORS['card'])
        inner.pack(padx=20, pady=15, fill=tk.BOTH, expand=True)
        
        tk.Label(
            inner,
            text="📋 Relatório de Processamento",
            font=FONTS['body_bold'],
            bg=COLORS['card'],
            fg=COLORS['text']
        ).pack(anchor=tk.W, pady=(0, 8))
        
        self.relatorio_texto = criar_area_relatorio(inner, self.cores)
        self.relatorio_texto.pack(expand=True, fill=tk.BOTH)
    
    # --- Métodos de Seleção ---
    
    def _contar_pdfs(self, caminho: str) -> int:
        """Conta arquivos PDF em uma pasta."""
        try:
            if not caminho or not os.path.exists(caminho):
                return 0
            return len([f for f in os.listdir(caminho) if f.lower().endswith('.pdf')])
        except Exception:
            return 0
    
    def _selecionar_pasta_faturas(self):
        """Abre diálogo para selecionar pasta de faturas."""
        caminho = filedialog.askdirectory(title="Selecione a pasta com as FATURAS")
        if caminho:
            self.pasta_faturas.set(caminho)
            count = self._contar_pdfs(caminho)
            self.contador_faturas.set(f"{count} arquivo{'s' if count != 1 else ''}")
            self._atualizar_status("✓ Pasta de faturas selecionada.")
    
    def _selecionar_pasta_boletos(self):
        """Abre diálogo para selecionar pasta de boletos."""
        caminho = filedialog.askdirectory(title="Selecione a pasta com os BOLETOS")
        if caminho:
            self.pasta_boletos.set(caminho)
            count = self._contar_pdfs(caminho)
            self.contador_boletos.set(f"{count} arquivo{'s' if count != 1 else ''}")
            self._atualizar_status("✓ Pasta de boletos selecionada.")
    
    # --- Métodos de UI ---
    
    def _atualizar_status(self, texto: str):
        """Atualiza o texto de status."""
        self.status_label.config(text=texto)
    
    def _adicionar_relatorio(self, texto: str):
        """Adiciona linha ao relatório."""
        self.relatorio_texto.config(state='normal')
        self.relatorio_texto.insert(tk.END, texto + '\n')
        self.relatorio_texto.config(state='disabled')
        self.relatorio_texto.see(tk.END)
    
    def _limpar_relatorio(self):
        """Limpa o relatório."""
        self.relatorio_texto.config(state='normal')
        self.relatorio_texto.delete(1.0, tk.END)
        self.relatorio_texto.config(state='disabled')
    
    def _habilitar_botoes(self, habilitar: bool = True):
        """Habilita/desabilita botões."""
        if habilitar:
            self.botao_unir.config(state='normal', bg=COLORS['success'])
            self.botao_cancelar.config(state='disabled', bg=COLORS['disabled'])
        else:
            self.botao_unir.config(state='disabled', bg=COLORS['disabled'])
            self.botao_cancelar.config(state='normal', bg=COLORS['error'])
    
    def _resetar_stats(self):
        """Reseta as estatísticas."""
        for key in self.stats:
            self.stats[key] = 0
    
    def _mostrar_estatisticas(self):
        """Mostra estatísticas de processamento."""
        self._queue.put(('log', "\n" + "="*50))
        self._queue.put(('log', "📊 ESTATÍSTICAS DE PROCESSAMENTO"))
        self._queue.put(('log', "="*50))
        
        self._queue.put(('log', f"\n📄 FATURAS:"))
        self._queue.put(('log', f"   Total processados: {self.stats['faturas_total']}"))
        self._queue.put(('log', f"   ✓ Válidos: {self.stats['faturas_ok']}"))
        self._queue.put(('log', f"   ✗ Sem UC: {self.stats['faturas_sem_uc']}"))
        self._queue.put(('log', f"   ✗ Sem valor: {self.stats['faturas_sem_valor']}"))
        self._queue.put(('log', f"   ✗ Erros: {self.stats['faturas_erro']}"))
        
        self._queue.put(('log', f"\n🧾 BOLETOS:"))
        self._queue.put(('log', f"   Total processados: {self.stats['boletos_total']}"))
        self._queue.put(('log', f"   ✓ Válidos: {self.stats['boletos_ok']}"))
        self._queue.put(('log', f"   ✗ Sem UC: {self.stats['boletos_sem_uc']}"))
        self._queue.put(('log', f"   ✗ Sem valor: {self.stats['boletos_sem_valor']}"))
        self._queue.put(('log', f"   ✗ Erros: {self.stats['boletos_erro']}"))
        
        self._queue.put(('log', "\n" + "="*50))
    
    # --- Processamento ---
    
    def _iniciar_processamento(self):
        """Inicia o processamento dos PDFs."""
        if not self.pasta_faturas.get() or not self.pasta_boletos.get():
            messagebox.showerror("Erro", "Selecione as duas pastas.")
            return
        
        # Escolher destino
        caminho_zip = filedialog.asksaveasfilename(
            title="Salvar arquivo ZIP como...",
            defaultextension=".zip",
            filetypes=[("Arquivos ZIP", "*.zip")],
            initialfile=f"unidos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        )
        
        if not caminho_zip:
            return
        
        # Preparar UI
        self._habilitar_botoes(False)
        self.progresso['value'] = 0
        self._limpar_relatorio()
        self._resetar_stats()
        self._cancelar.clear()
        
        # Configurar logger com callback para UI
        setup_logger('unificador', self._queue.put)
        
        # Suppress noisy library logs
        logging.getLogger("pdfminer").setLevel(logging.ERROR)
        logging.getLogger("pypdf").setLevel(logging.ERROR)
        
        # Iniciar worker
        args = (
            Path(self.pasta_faturas.get()),
            Path(self.pasta_boletos.get()),
            self.ordem_pdf.get(),
            Path(caminho_zip)
        )
        self._worker_thread = threading.Thread(
            target=self._worker_processar,
            args=args,
            daemon=True
        )
        self._worker_thread.start()
    
    def _cancelar_processamento(self):
        """Solicita cancelamento do processamento."""
        if self._worker_thread and self._worker_thread.is_alive():
            self._cancelar.set()
            self._adicionar_relatorio("⌛ Solicitando cancelamento...")
    
    def _processar_queue(self):
        """Processa mensagens da fila de comunicação com worker."""
        try:
            while True:
                tipo, payload = self._queue.get_nowait()
                
                if tipo == 'status':
                    self._atualizar_status(payload)
                elif tipo == 'progress':
                    self.progresso['value'] = payload
                elif tipo == 'log':
                    self._adicionar_relatorio(payload)
                elif tipo == 'info':
                    messagebox.showinfo("Concluído", payload)
                elif tipo == 'erro':
                    messagebox.showerror("Erro", payload)
                elif tipo == 'finalizar_ui':
                    self._habilitar_botoes(True)
                    
        except queue.Empty:
            pass
        
        self.root.after(100, self._processar_queue)
    
    def _worker_processar(
        self,
        pasta_faturas: Path,
        pasta_boletos: Path,
        ordem: str,
        caminho_zip: Path
    ):
        """
        Worker thread para processar os PDFs.
        
        Args:
            pasta_faturas: Caminho da pasta de faturas
            pasta_boletos: Caminho da pasta de boletos
            ordem: 'fatura_primeiro' ou 'boleto_primeiro'
            caminho_zip: Caminho para salvar o ZIP
        """
        log = get_logger('unificador')
        
        try:
            self._queue.put(('log', f"🚀 Iniciando processamento..."))
            self._queue.put(('log', f"📁 Pasta faturas: {pasta_faturas}"))
            self._queue.put(('log', f"📁 Pasta boletos: {pasta_boletos}"))
            self._queue.put(('log', ""))
            
            self._queue.put(('status', "🔍 Lendo arquivos e extraindo UCs..."))
            
            # Processar faturas
            self._queue.put(('log', "="*50))
            self._queue.put(('log', "📄 PROCESSANDO FATURAS..."))
            self._queue.put(('log', "="*50))
            map_faturas = self._processar_pasta_faturas(pasta_faturas)
            
            # Processar boletos
            self._queue.put(('log', "\n" + "="*50))
            self._queue.put(('log', "🧾 PROCESSANDO BOLETOS..."))
            self._queue.put(('log', "="*50))
            map_boletos = self._processar_pasta_boletos(pasta_boletos)
            
            # Mostrar estatísticas
            self._mostrar_estatisticas()
            
            # Encontrar pares
            ucs_faturas = set(map_faturas.keys())
            ucs_boletos = set(map_boletos.keys())
            ucs_com_par = sorted(
                ucs_faturas.intersection(ucs_boletos),
                key=lambda x: int(x) if x.isdigit() else 0
            )
            
            # Log de UCs encontradas
            self._queue.put(('log', f"\n📋 UCs de faturas: {sorted(ucs_faturas)}"))
            self._queue.put(('log', f"📋 UCs de boletos: {sorted(ucs_boletos)}"))
            self._queue.put(('log', f"📋 UCs com par (interseção): {ucs_com_par}"))
            
            # UCs só em faturas
            ucs_so_faturas = ucs_faturas - ucs_boletos
            if ucs_so_faturas:
                self._queue.put(('log', f"⚠️ UCs só em faturas (sem boleto): {sorted(ucs_so_faturas)}"))
            
            # UCs só em boletos
            ucs_so_boletos = ucs_boletos - ucs_faturas
            if ucs_so_boletos:
                self._queue.put(('log', f"⚠️ UCs só em boletos (sem fatura): {sorted(ucs_so_boletos)}"))
            
            total = len(ucs_com_par)
            if total == 0:
                self._queue.put(('status', "❌ Nenhum par UC encontrado."))
                self._queue.put(('log', "\n❌ ERRO: Nenhum par UC encontrado!"))
                self._queue.put(('log', "   Verifique se os arquivos contêm UCs válidas."))
                self._queue.put(('log', "   Consulte o arquivo de log para mais detalhes."))
                log.info(f"Log salvo em: {log.log_file}")
                self._queue.put(('erro', f"Nenhum par (UC comum) encontrado.\n\nVerifique o log em:\n{log.log_file}"))
                self._queue.put(('finalizar_ui', None))
                return
            
            # Unir PDFs
            self._queue.put(('log', f"\n🔗 UNINDO {total} PARES..."))
            sucesso = self._unir_pares(
                ucs_com_par, map_faturas, map_boletos, ordem, caminho_zip
            )
            
            # Finalizar
            if self._cancelar.is_set():
                self._queue.put(('status', "⚠️ Processo cancelado."))
            else:
                msg = f"Concluído! {sucesso}/{total} pares unidos.\nSalvo em: {caminho_zip}\n\nLog detalhado em:\n{log.log_file}"
                self._queue.put(('status', "✅ Processo finalizado!"))
                self._queue.put(('log', f"\n✅ CONCLUÍDO: {sucesso}/{total} pares unidos"))
                self._queue.put(('log', f"📁 ZIP: {caminho_zip}"))
                self._queue.put(('log', f"📁 LOG: {log.log_file}"))
                self._queue.put(('info', msg))
            
            self._queue.put(('finalizar_ui', None))
            
        except Exception as e:
            import traceback
            log.error(f"Exceção: {traceback.format_exc()}")
            self._queue.put(('finalizar_ui', None))
            self._queue.put(('status', "❌ Erro no processamento."))
            self._queue.put(('erro', f"{str(e)}\n\nVerifique o log para detalhes."))
    
    def _processar_pasta_faturas(self, pasta: Path) -> dict:
        """Processa pasta de faturas e retorna mapa UC -> dados."""
        log = get_logger('unificador')
        resultado = {}
        
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
        self.stats['faturas_total'] = len(arquivos)
        
        self._queue.put(('log', f"Encontrados {len(arquivos)} arquivos PDF"))
        
        for idx, nome in enumerate(arquivos, 1):
            self._queue.put(('log', f"\n[{idx}/{len(arquivos)}] Processando: {nome}"))
            
            caminho = pasta / nome
            
            # Validar cabeçalho PDF
            if not validar_pdf_cabecalho(caminho):
                self._queue.put(('log', f"   ✗ Arquivo não é um PDF válido"))
                self.stats['faturas_erro'] += 1
                continue
            
            # Tentar extrair UC do nome
            uc = extrai_uc(nome)
            if uc:
                self._queue.put(('log', f"   UC do nome: {uc}"))
            else:
                self._queue.put(('log', f"   UC não encontrada no nome"))
            
            texto = ""
            
            try:
                # reader = abrir_pdf_seguro(caminho)
                texto = extrair_texto_pdf(caminho)
                
                log.info(f"   Texto extraído: {len(texto)} caracteres", to_ui=True)
                
                # Mostrar preview do texto
                preview = texto[:200].replace('\n', ' ').replace('\r', '')
                log.info(f"   Preview: {preview}...", to_ui=True)
                
                # Se não encontrou UC no nome, tentar no texto
                if uc is None:
                    uc = extrai_uc_do_texto(texto, nome)
                    if uc:
                        self._queue.put(('log', f"   ✓ UC extraída do texto: {uc}"))
                    else:
                        # Fazer análise detalhada
                        analisar_texto_pdf(texto, nome)
                        
            except Exception as e:
                self._queue.put(('log', f"   ✗ Erro ao ler: {e}"))
                log.error(f"Erro ao ler fatura {nome}: {e}")
                self.stats['faturas_erro'] += 1
                continue
            
            if not uc:
                self._queue.put(('log', f"   ✗ FALHA: Nenhuma UC encontrada"))
                self.stats['faturas_sem_uc'] += 1
                continue
            
            # Extrair valor
            valor = extrair_valor_fatura(texto, nome)
            if valor is None:
                self._queue.put(('log', f"   ✗ FALHA: Nenhum valor encontrado"))
                analisar_valores_pdf(texto, nome, "fatura")
                self.stats['faturas_sem_valor'] += 1
                continue
            
            self._queue.put(('log', f"   ✓ Valor: R$ {valor:.2f}"))
            
            # Manter o mais recente
            if uc not in resultado or caminho.stat().st_mtime > Path(resultado[uc]['caminho']).stat().st_mtime:
                resultado[uc] = {'caminho': str(caminho), 'valor': valor, 'nome': nome}
                self._queue.put(('log', f"   ✓ OK - UC {uc} = R$ {valor:.2f}"))
            
            self.stats['faturas_ok'] += 1
        
        return resultado
    
    def _processar_pasta_boletos(self, pasta: Path) -> dict:
        """Processa pasta de boletos e retorna mapa UC -> dados."""
        log = get_logger('unificador')
        resultado = {}
        
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
        self.stats['boletos_total'] = len(arquivos)
        
        self._queue.put(('log', f"Encontrados {len(arquivos)} arquivos PDF"))
        
        for idx, nome in enumerate(arquivos, 1):
            self._queue.put(('log', f"\n[{idx}/{len(arquivos)}] Processando: {nome}"))
            
            caminho = pasta / nome
            
            # Validar cabeçalho PDF
            if not validar_pdf_cabecalho(caminho):
                self._queue.put(('log', f"   ✗ Arquivo não é um PDF válido"))
                self.stats['boletos_erro'] += 1
                continue
            
            # Tentar extrair UC do nome
            uc = extrai_uc(nome)
            if uc:
                self._queue.put(('log', f"   UC do nome: {uc}"))
            else:
                self._queue.put(('log', f"   UC não encontrada no nome"))
            
            texto = ""
            
            try:
                # reader = abrir_pdf_seguro(caminho)
                texto = extrair_texto_pdf(caminho)
                
                log.info(f"   Texto extraído: {len(texto)} caracteres", to_ui=True)
                
                # Mostrar preview do texto
                preview = texto[:200].replace('\n', ' ').replace('\r', '')
                log.info(f"   Preview: {preview}...", to_ui=True)
                
                # Se não encontrou UC no nome, tentar no texto
                if uc is None:
                    uc = extrai_uc_do_texto(texto, nome)
                    if uc:
                        self._queue.put(('log', f"   ✓ UC extraída do texto: {uc}"))
                    else:
                        # Fazer análise detalhada
                        analisar_texto_pdf(texto, nome)
                        
            except Exception as e:
                self._queue.put(('log', f"   ✗ Erro ao ler: {e}"))
                log.error(f"Erro ao ler boleto {nome}: {e}")
                self.stats['boletos_erro'] += 1
                continue
            
            if not uc:
                self._queue.put(('log', f"   ✗ FALHA: Nenhuma UC encontrada"))
                self.stats['boletos_sem_uc'] += 1
                continue
            
            # Extrair valor
            valor = extrair_valor_boleto(texto, nome)
            if valor is None:
                self._queue.put(('log', f"   ✗ FALHA: Nenhum valor encontrado"))
                analisar_valores_pdf(texto, nome, "boleto")
                self.stats['boletos_sem_valor'] += 1
                continue
            
            self._queue.put(('log', f"   ✓ Valor: R$ {valor:.2f}"))
            
            # Manter o mais recente
            if uc not in resultado or caminho.stat().st_mtime > Path(resultado[uc]['caminho']).stat().st_mtime:
                resultado[uc] = {'caminho': str(caminho), 'valor': valor, 'nome': nome}
                self._queue.put(('log', f"   ✓ OK - UC {uc} = R$ {valor:.2f}"))
            
            self.stats['boletos_ok'] += 1
        
        return resultado
    
    def _unir_pares(
        self,
        ucs: list[str],
        faturas: dict,
        boletos: dict,
        ordem: str,
        caminho_zip: Path
    ) -> int:
        """Une os pares de faturas e boletos."""
        sucesso = 0
        total = len(ucs)
        
        with zipfile.ZipFile(caminho_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
            for i, uc in enumerate(ucs, start=1):
                if self._cancelar.is_set():
                    self._queue.put(('log', "⚠️ Cancelado pelo usuário."))
                    break
                
                fatura = faturas[uc]
                boleto = boletos[uc]
                
                # Validar valores
                valor_fatura = round(fatura['valor'] * 100)
                valor_boleto = round(boleto['valor'] * 100)
                
                if valor_fatura != valor_boleto:
                    self._queue.put(('log', 
                        f"⚠️ UC {uc}: Valores divergentes - "
                        f"Fatura R${fatura['valor']:.2f} ≠ Boleto R${boleto['valor']:.2f} → IGNORADO"
                    ))
                    continue
                
                self._queue.put(('status', 
                    f"✅ Processando {i}/{total}: UC {uc} (R${fatura['valor']:.2f})"
                ))
                self._queue.put(('progress', (i * 100) / max(total, 1)))
                
                try:
                    # Unir PDFs
                    pdf_bytes = unir_pdfs(
                        [fatura['caminho'], boleto['caminho']],
                        ordem_boleto_primeiro=(ordem == 'boleto_primeiro')
                    )
                    
                    # Salvar no ZIP
                    nome_final = criar_nome_arquivo(
                        uc, fatura['nome'], boleto['nome'],
                        mes="Dez", ano="2025"
                    )
                    zf.writestr(nome_final, pdf_bytes)
                    
                    self._queue.put(('log', f"✓ UC {uc}: União bem-sucedida → {nome_final}"))
                    sucesso += 1
                    
                except Exception as e:
                    self._queue.put(('log', f"✗ UC {uc}: Erro ao unir - {e}"))
        
        return sucesso
