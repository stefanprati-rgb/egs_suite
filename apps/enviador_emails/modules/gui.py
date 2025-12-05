import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import logging
import threading
from datetime import datetime
from .core import enviar_emails_worker

class EmailSenderApp:
    def __init__(self, master):
        self.master = master
        master.title("EGS - Enviador de E-mails v2.7 (Modern UI)")
        master.geometry("800x750") # Aumentado um pouco para acomodar o padding extra

        self.pasta_pdfs = tk.StringVar()
        self.caminho_processado = tk.StringVar()
        self.mes_ref = tk.StringVar(value=datetime.now().strftime("%Y-%m")) 
        self.modo_envio = tk.StringVar(value="Enviar Diretamente")
        self.is_running = False

        logging.info("Criando widgets da interface gráfica...")
        self._create_widgets()

    def _create_widgets(self):
        # Definições de Estilo
        BG_COLOR = "#F8F9FA" # Off-white moderno
        ACCENT_COLOR = "#0078D7" # Azul Outlook
        TEXT_COLOR = "#212529" # Cinza escuro
        
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configuração global
        style.configure(".", background=BG_COLOR, foreground=TEXT_COLOR, font=('Segoe UI', 10))
        
        # Frame
        style.configure("TFrame", background=BG_COLOR)
        
        # Botões Modernos
        style.configure("TButton", 
            padding=(20, 10), 
            relief="flat", 
            background=ACCENT_COLOR, 
            foreground="white", 
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0
        )
        style.map("TButton", 
            background=[('active', '#005A9E'), ('pressed', '#004578')],
            relief=[('pressed', 'sunken')]
        )
        
        # Labels e Entradas
        style.configure("TLabel", background=BG_COLOR, font=('Segoe UI', 10))
        style.configure("TEntry", fieldbackground="white", padding=5)
        
        # LabelFrame
        style.configure("TLabelFrame", background=BG_COLOR, relief='flat', borderwidth=0)
        style.configure("TLabelFrame.Label", background=BG_COLOR, foreground=ACCENT_COLOR, font=('Segoe UI', 12, 'bold'))

        # --- Layout Principal ---
        self.main_frame = ttk.Frame(self.master, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Configuração
        frame_config = ttk.LabelFrame(self.main_frame, text="1. Configuração de Arquivos e Mês", padding="15")
        frame_config.pack(fill=tk.X, pady=(0, 15))
        frame_config.columnconfigure(1, weight=1)

        def create_file_selector(parent, label_text, var, command, row):
            ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
            
            entry_frame = ttk.Frame(parent)
            entry_frame.grid(row=row+1, column=0, columnspan=3, sticky=tk.EW, pady=(0, 15))
            entry_frame.columnconfigure(0, weight=1)
            
            ttk.Entry(entry_frame, textvariable=var, state='readonly').grid(row=0, column=0, sticky=tk.EW, ipady=3)
            ttk.Button(entry_frame, text="Procurar", command=command, width=10).grid(row=0, column=1, padx=(10, 0))

        create_file_selector(frame_config, "Pasta dos PDFs Unificados:", self.pasta_pdfs, lambda: self.pasta_pdfs.set(filedialog.askdirectory(title="Selecione a pasta com todos os PDFs")), row=0)
        create_file_selector(frame_config, "Planilha Processada (Boletos):", self.caminho_processado, lambda: self.caminho_processado.set(filedialog.askopenfilename(title="Selecione a planilha com a lista de UCs para envio", filetypes=[("Planilhas", "*.xlsx *.xls *.csv")])), row=2)
        
        # Mês de Referência
        frame_mes = ttk.Frame(frame_config)
        frame_mes.grid(row=4, column=0, columnspan=3, sticky=tk.W)
        ttk.Label(frame_mes, text="Mês de Ref. (AAAA-MM):").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Entry(frame_mes, textvariable=self.mes_ref, width=15).pack(side=tk.LEFT, ipady=3)
        
        # 2. Conteúdo do E-mail
        frame_email = ttk.LabelFrame(self.main_frame, text="2. Conteúdo do E-mail e Configuração", padding="15")
        frame_email.pack(fill=tk.X, pady=(0, 15))
        frame_email.columnconfigure(1, weight=1)

        ttk.Label(frame_email, text="Assunto:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.assunto_entry = ttk.Entry(frame_email)
        self.assunto_entry.insert(0, "EGS Energia - Fatura {MES_REF}")
        self.assunto_entry.grid(row=1, column=0, columnspan=2, padx=0, pady=(0, 15), sticky=tk.EW, ipady=3)

        ttk.Label(frame_email, text="Corpo do E-mail:").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        self.corpo_text = scrolledtext.ScrolledText(frame_email, wrap=tk.WORD, height=6, font=('Segoe UI', 10), borderwidth=1, relief="solid")
        self.corpo_text.insert(tk.END, "Prezado(a) {CLIENTE_NOME},\n\nSua fatura de energia referente ao mês {MES_REF} (UC {UC_CODIGO}) está em anexo.\n\nQualquer dúvida, estamos à disposição.\n\nAtenciosamente,\nEGS Energia")
        self.corpo_text.grid(row=3, column=0, columnspan=2, padx=0, pady=(0, 15), sticky="nsew")
        
        # Modo de Envio (usando um Frame para alinhar melhor)
        frame_modo = ttk.Frame(frame_email)
        frame_modo.grid(row=4, column=0, columnspan=2, sticky=tk.EW)
        ttk.Label(frame_modo, text="Modo de Envio:").pack(side=tk.LEFT, padx=(0, 10))
        ttk.OptionMenu(frame_modo, self.modo_envio, self.modo_envio.get(), "Enviar Diretamente", "Salvar (Rascunhos)").pack(side=tk.LEFT)
        
        # 3. Ações
        frame_actions = ttk.LabelFrame(self.main_frame, text="3. Execução", padding="15")
        frame_actions.pack(fill=tk.BOTH, expand=True)
        frame_actions.columnconfigure(0, weight=1)
        frame_actions.rowconfigure(2, weight=1)

        self.btn_iniciar_massa = ttk.Button(frame_actions, text="INICIAR ENVIO EM MASSA", command=self.iniciar_envio_massa, cursor="hand2")
        self.btn_iniciar_massa.grid(row=0, column=0, pady=(0, 15), sticky=tk.EW)
        
        self.progress = ttk.Progressbar(frame_actions, orient="horizontal", length=100, mode="determinate")
        self.progress.grid(row=1, column=0, pady=(0, 15), sticky=tk.EW)

        self.status_text = scrolledtext.ScrolledText(frame_actions, height=8, state='disabled', wrap=tk.WORD, font=('Consolas', 9), bg="white", borderwidth=1, relief="solid")
        self.status_text.grid(row=2, column=0, sticky="nsew")

        # Configura as cores da janela principal
        self.master.configure(bg=BG_COLOR)

    def _toggle_controls(self, enable: bool):
        """Habilita ou desabilita controles recursivamente."""
        state = 'normal' if enable else 'disabled'
        self.btn_iniciar_massa.config(state=state)
        
        # Função recursiva para desabilitar/habilitar
        def set_state(widget):
            if isinstance(widget, (ttk.Button, ttk.Entry, scrolledtext.ScrolledText, ttk.OptionMenu)):
                if widget not in [self.status_text, self.progress]:
                    try:
                        widget.config(state=state)
                    except: pass
            
            # Chama recursivamente para os filhos
            for child in widget.winfo_children():
                set_state(child)
                
        set_state(self.main_frame)

        if enable:
             self.assunto_entry.config(state='normal')
             self.corpo_text.config(state='normal')
             # Re-habilita menus que podem ter ficado travados
             self.master.nametowidget('.!frame.!labelframe2.!frame.!menubutton').config(state='normal')

    def update_status(self, message):
        def _update():
            self.status_text.config(state='normal')
            self.status_text.insert(tk.END, message + "\n")
            self.status_text.see(tk.END)
            self.status_text.config(state='disabled')
            if "ERRO" in message or "FATAL" in message: logging.error(message)
            elif "⚠️" in message or "DICA" in message: logging.warning(message)
            else: logging.info(message)
        self.master.after(0, _update)

    def update_progress(self, current, total):
        def _update():
            self.progress['value'] = (current / total) * 100 if total > 0 else 0
        self.master.after(0, _update)

    def on_completion(self, total, sucessos, num_falhas, caminho_relatorio_falhas, caminho_relatorio_sucesso):
        self.is_running = False
        self._toggle_controls(True)
        
        # Garante que a barra de progresso pare e volte ao modo normal em caso de erro na carga
        self.progress.stop()
        self.progress.config(mode='determinate')
        self.update_progress(0, 1) # Reseta para 0%
        
        self.update_status("="*50)
        self.update_status(f"Processo Finalizado. Total de UCs válidas para envio: {total}")
        self.update_status(f"✅ Sucesso: {sucessos}")
        self.update_status(f"❌ Falhas (totais): {num_falhas}")
        self.update_status("="*50)

        final_message = f"Processo concluído!\n\n"
        final_message += f"- ✅ {sucessos} e-mails enviados/salvos com sucesso.\n"
        final_message += f"- ❌ {num_falhas} falhas registradas.\n\n"
        
        if caminho_relatorio_sucesso:
            final_message += f"RELATÓRIO DE SUCESSOS:\n{os.path.abspath(caminho_relatorio_sucesso)}\n\n"
        
        if caminho_relatorio_falhas:
            final_message += f"RELATÓRIO DE FALHAS:\n{os.path.abspath(caminho_relatorio_falhas)}\n"
        
        if num_falhas > 0:
             messagebox.showwarning("Processo Concluído", final_message)
        else:
             messagebox.showinfo("Processo Concluído", final_message)

    def iniciar_envio_massa(self):
        if not all([self.pasta_pdfs.get(), self.caminho_processado.get()]):
            messagebox.showerror("Erro", "Por favor, selecione a Pasta de PDFs e a Planilha Processada.")
            return

        self.is_running = True
        self._toggle_controls(False)
        self.progress['value'] = 0
        self.status_text.config(state='normal')
        self.status_text.delete('1.0', tk.END)
        self.status_text.config(state='disabled')

        # NOVO: Inicia a barra de progresso no modo 'indeterminate'
        self.progress.config(mode='indeterminate')
        self.progress.start(10) # O valor (10) controla a velocidade da animação

        args = (
            self.pasta_pdfs.get(), self.caminho_processado.get(),
            self.mes_ref.get(), self.assunto_entry.get(), self.corpo_text.get("1.0", tk.END),
            self.modo_envio.get(), self.update_status, self.update_progress, self.on_completion,
            self.progress # Passa a referência da barra de progresso para a thread
        )
        threading.Thread(target=enviar_emails_worker, args=args, daemon=True).start()

    def on_closing(self):
        if self.is_running and not messagebox.askyesno("Sair", "O envio está em andamento. Deseja realmente fechar?"):
            return
        self.master.destroy()
