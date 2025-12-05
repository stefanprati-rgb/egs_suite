import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from threading import Thread
from datetime import datetime
import calendar

from .config import DOMINIO_REMETENTE_VALIDO
from .file_manager import corrigir_pdfs_antigos
from .outlook_service import buscar_e_salvar_boletos

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de Boletos EGS - v18.1 (Modern UI)")
        self.root.geometry("700x680")
        
        # --- MELHORIA: Modernização do Tema e Estilos ---
        style = ttk.Style()
        try:
            style.theme_use('vista')
        except:
            style.theme_use('clam')
        
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", font=("Segoe UI", 10), background="#f0f0f0")
        style.configure("TButton", font=("Segoe UI", 10, "bold"))
        style.configure("TLabelframe", font=("Segoe UI", 10, "bold"), background="#f0f0f0")
        style.configure("TLabelframe.Label", background="#f0f0f0")
        
        self.root.configure(bg="#f0f0f0")

        main = ttk.Frame(root, padding="25")
        main.pack(fill=tk.BOTH, expand=True)
        
        header_lbl = ttk.Label(main, text="Buscar Boletos no Outlook", font=("Segoe UI", 20, "bold"), foreground="#003399")
        header_lbl.pack(pady=(0, 20))
        
        date_frame = ttk.LabelFrame(main, text="1. Filtro de Período", padding=15)
        date_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(date_frame, text="Data de Início:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Label(date_frame, text="Data de Fim:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.create_date_selectors(date_frame, 0, 1)
        self.create_date_selectors(date_frame, 1, 1)
        self.set_default_dates()
        
        self.search_button = ttk.Button(main, text="▶ Iniciar Busca de Boletos", command=self.start_search_thread)
        self.search_button.pack(pady=10, ipady=8, fill=tk.X)
        
        corr_frame = ttk.LabelFrame(main, text="2. Manutenção", padding=15)
        corr_frame.pack(fill=tk.X, pady=10)
        self.fix_button = ttk.Button(corr_frame, text="🛠️ Corrigir Arquivos Antigos (da pasta 'sem_uc')", command=self.corrigir_antigos)
        self.fix_button.pack(pady=5, ipady=5, fill=tk.X)

        self.progress_bar = ttk.Progressbar(main, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(15, 0))
        self.progress_label = ttk.Label(main, text="Aguardando início...", font=("Segoe UI", 9))
        self.progress_label.pack(fill=tk.X, pady=(2, 10))

        ttk.Label(main, text="Log de Atividades:", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W)
        self.status_text = scrolledtext.ScrolledText(main, height=12, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 9))
        self.status_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # --- Tags de Cores ---
        self.status_text.tag_config('success', foreground='#008000') # Verde Escuro
        self.status_text.tag_config('error', foreground='#CC0000', font=("Consolas", 9, "bold")) # Vermelho
        self.status_text.tag_config('warning', foreground='#FF8C00') # Laranja Escuro
        self.status_text.tag_config('info', foreground='#000000')

    def create_date_selectors(self, parent, row_num, start_col):
        setattr(self, f'day_var_{row_num}', tk.StringVar()); setattr(self, f'month_var_{row_num}', tk.StringVar()); setattr(self, f'year_var_{row_num}', tk.StringVar())
        days=[f"{d:02d}" for d in range(1,32)]; months=[f"{m:02d}" for m in range(1,13)]; years=[str(y) for y in range(datetime.now().year+1, datetime.now().year-5, -1)]
        day_cb=ttk.Combobox(parent, textvariable=getattr(self, f'day_var_{row_num}'), values=days, width=4); month_cb=ttk.Combobox(parent, textvariable=getattr(self, f'month_var_{row_num}'), values=months, width=4); year_cb=ttk.Combobox(parent, textvariable=getattr(self, f'year_var_{row_num}'), values=years, width=6)
        day_cb.grid(row=row_num, column=start_col, padx=(0,2)); month_cb.grid(row=row_num, column=start_col+1, padx=2); year_cb.grid(row=row_num, column=start_col+2, padx=2)

    def set_default_dates(self):
        today = datetime.today(); first_day = today.replace(day=1); last_day = first_day.replace(day=calendar.monthrange(today.year, today.month)[1])
        self.day_var_0.set(f"{first_day.day:02d}"); self.month_var_0.set(f"{first_day.month:02d}"); self.year_var_0.set(str(first_day.year))
        self.day_var_1.set(f"{last_day.day:02d}"); self.month_var_1.set(f"{last_day.month:02d}"); self.year_var_1.set(str(last_day.year))

    def _limpar_e_preparar_ui(self):
        self.search_button.config(state=tk.DISABLED)
        self.fix_button.config(state=tk.DISABLED)
        self.status_text.configure(state=tk.NORMAL); self.status_text.delete("1.0", tk.END); self.status_text.configure(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        self.progress_label.config(text="")

    def update_status(self, message): self.root.after(0, self._update_status_thread_safe, message)
    def _update_status_thread_safe(self, message):
        self.status_text.configure(state=tk.NORMAL)
        
        # Heurística de cores
        msg_lower = message.lower()
        tag = 'info'
        if "erro" in msg_lower or "falha" in msg_lower or "crítico" in msg_lower:
            tag = 'error'
        elif "sucesso" in msg_lower or "concluída" in msg_lower or "finalizada" in msg_lower or "✅" in message:
            tag = 'success'
        elif "aviso" in msg_lower or "atenção" in msg_lower or "⚠️" in message:
            tag = 'warning'
            
        self.status_text.insert(tk.END, message + "\n", tag)
        self.status_text.configure(state=tk.DISABLED)
        self.status_text.see(tk.END)
    
    def update_progress(self, value, max_value, text): self.root.after(0, self._update_progress_thread_safe, value, max_value, text)
    def _update_progress_thread_safe(self, value, max_value, text):
        if max_value > 0:
            self.progress_bar['maximum'] = max_value
            self.progress_bar['value'] = value
            percent = int((value / max_value) * 100)
            self.progress_label.config(text=f"{text} ({percent}%)")
        else:
            self.progress_label.config(text=text)


    def on_search_completion(self, sucesso, falha, erro=False): self.root.after(0, self._on_search_completion_thread_safe, sucesso, falha, erro)
    def _on_search_completion_thread_safe(self, sucesso, falha, erro):
        resumo = (f"\n--- FIM DA BUSCA ---\n"
                  f"✅ Boletos com UC identificada: {sucesso}\n"
                  f"⚠️ Boletos sem UC (para análise): {falha}\n")
        self._update_status_thread_safe(resumo)
        self.progress_bar['value'] = 0; self.progress_label.config(text="Concluído")
        self.search_button.config(state=tk.NORMAL); self.fix_button.config(state=tk.NORMAL)
        msg_final = (f"Busca finalizada!\n\nSucessos: {sucesso}\nPara Análise: {falha}")
        if not erro: messagebox.showinfo("Busca Concluída", msg_final)
        else: messagebox.showerror("Erro", "Ocorreu um erro crítico. Verifique o log.")
    
    def validate_date_range(self):
        try:
            start = datetime.strptime(f"{self.day_var_0.get()}/{self.month_var_0.get()}/{self.year_var_0.get()}", "%d/%m/%Y")
            end = datetime.strptime(f"{self.day_var_1.get()}/{self.month_var_1.get()}/{self.year_var_1.get()}", "%d/%m/%Y")
            if start > end:
                messagebox.showerror("Erro de Data", "A data de início não pode ser posterior à data final.")
                return False
            return True
        except ValueError:
            messagebox.showerror("Data Inválida", "Pelo menos uma das datas selecionadas é inválida (ex: dia 30 em Fevereiro)."); return False

    def corrigir_antigos(self):
        self._limpar_e_preparar_ui()
        self.search_button.config(state=tk.DISABLED)
        self.fix_button.config(state=tk.DISABLED)
        self.update_status("\n--- Iniciando correção de arquivos antigos... ---")
        t = Thread(target=self._corrigir_antigos_thread, daemon=True)
        t.start()

    def _corrigir_antigos_thread(self):
        resumo = corrigir_pdfs_antigos()
        def show_summary():
            messagebox.showinfo("Correção Concluída", f"Resultado da correção:\n\n{resumo}")
            self.update_status("\n--- Correção de arquivos antigos concluída ---\n" + resumo)
            self.search_button.config(state=tk.NORMAL); self.fix_button.config(state=tk.NORMAL)
            self.progress_label.config(text="Correção Concluída")
        self.root.after(0, show_summary)


    def start_search_thread(self):
        if not self.validate_date_range():
            return
            
        self._limpar_e_preparar_ui()
        start_date_str = f"{self.day_var_0.get()}/{self.month_var_0.get()}/{self.year_var_0.get()}"
        end_date_str = f"{self.day_var_1.get()}/{self.month_var_1.get()}/{self.year_var_1.get()}"
        self.update_status(f"Iniciando busca de boletos do domínio '{DOMINIO_REMETENTE_VALIDO}'...")
        t = Thread(target=buscar_e_salvar_boletos, args=(start_date_str, end_date_str, self.update_status, self.update_progress, self.on_search_completion), daemon=True)
        t.start()
