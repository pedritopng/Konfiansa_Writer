import sys, os
import re
import locale
import pandas as pd
import datetime
import platform
import subprocess
import time
from bs4 import BeautifulSoup
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

#precisa pip instalar xlsxwriter

try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    print("Locale pt_BR.UTF-8 não suportado. Usando o locale padrão.")
    locale.setlocale(locale.LC_ALL, '')

ALL_COLUMNS = [
    "USUÁRIO", "EMPRESA COB.",
    "PROCESSO", "NOME CLIENTE", "CPF/CNPJ", "ESPÉCIE",
    "TÍTULOS COB.", "DATA VENC.", "VALORES TIT.",
    "TOTAL DEVIDO", "SITUAÇÃO", "COBRADOR", "Nº DA PC", "DATA PGTO",
    "VALOR CRÉDITO", "TAXA", "VALORES DEP.", "TOTAL CLIENTE"
]


def resource_path(rel_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, rel_path)


def abrir_arquivo(path):
    sistema = platform.system()
    try:
        if sistema == "Windows":
            os.startfile(path)
        elif sistema == "Darwin":
            subprocess.call(["open", path])
        else:
            subprocess.call(["xdg-open", path])
    except Exception as e:
        print(f"Erro ao abrir arquivo: {e}")


def clean_and_convert_to_float(value_str):
    """Helper function to clean currency strings and convert to float."""
    if not isinstance(value_str, str):
        return 0.0
    try:
        cleaned_str = value_str.strip().replace("R$", "").replace(" ", "")
        return locale.atof(cleaned_str)
    except (ValueError, IndexError):
        return 0.0


def processar_arquivo(input_file, output_file, selected_columns, start_date_str, end_date_str, gui_instance=None):
    with open(input_file, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "html.parser")

    all_text = soup.get_text(" ", strip=True)
    blocks = all_text.split("Processo:")[1:]
    if not blocks:
        raise ValueError("O arquivo HTML não contém blocos válidos com 'Processo'.")

    codigo_empresas = {
        "1104": "LATINA METALURGICOS", "1769": "LATINA ELETROFERRAGENS",
        "2374": "LEVEL IMPLEMENTOS", "1545": "LATINA PRODUTOS ELETRICOS",
        "705571": "CONSULTH GUINDASTES", "51654": "CONSULTH ELETROFERRAGENS"
    }

    linhas = []
    grande_total_valor = 0.0
    grande_total_devido_valor = 0.0
    total_blocks = len(blocks)
    primeira_empresa_info = None

    try:
        start_date = datetime.datetime.strptime(start_date_str, "%d/%m/%Y") if start_date_str else None
        end_date = datetime.datetime.strptime(end_date_str, "%d/%m/%Y") if end_date_str else None
    except ValueError:
        raise ValueError("Formato de data inválido. Use DD/MM/AAAA.")

    for i, block in enumerate(blocks):
        enterprise = re.search(r"Cliente:\s*(.+?)(?=\s+CPF\b)", block)
        enterprise_code = re.search(r"\d+", enterprise.group(1)).group() if enterprise else None
        nome_empresa = codigo_empresas.get(enterprise_code, "CÓDIGO DESCONHECIDO")

        if i == 0 and enterprise_code:
            primeira_empresa_info = f"{enterprise_code} - {nome_empresa}"

        process = re.search(r"^\s*(\d+)", block)
        cnpj = re.search(r"CPF/CNPJ\s*:\s*([\d./-]+)", block)
        name_raw = re.search(r"Nome:\s*(.+?)(?:Endereço:|CPF/CNPJ:|Telefone:|Cobrador:|Situação:)", block)
        name = name_raw.group(1).strip() if name_raw else (enterprise.group(1).strip() if enterprise else "")
        cobrador_raw = re.search(r"Cobrador\s*:\s*(.+?)\s+Telefone:", block)
        cobrador = " ".join(cobrador_raw.group(1).split()[:2]).upper() if cobrador_raw else ""
        situation = re.search(r"Situação:\s*(.+?)\s+(?:Cliente:|CPF/CNPJ)", block)
        titles = re.findall(r"([Dd][Pp]|[Pp][Rr])\s+([\d/]+).*?(\d{2}/\d{2}/\d{4}).*?([\d.,]+)", block)
        tipos_lista = ", ".join(sorted({t[0].upper() for t in titles}))
        title_list = ", ".join([t[1] for t in titles])
        title_date = ", ".join([t[2] for t in titles])
        value_list = ", ".join([str(clean_and_convert_to_float(t[3])) for t in titles])

        total_matches = re.findall(r"Total do Devedor:.*?([\d.,]+)", block)
        total_value_str = total_matches[-1] if total_matches else ""
        total_devido_float = clean_and_convert_to_float(total_value_str)
        grande_total_devido_valor += total_devido_float

        base_data = {
            "USUÁRIO": enterprise_code, "EMPRESA COB.": nome_empresa,
            "PROCESSO": process.group(1) if process else "",
            "NOME CLIENTE": name.upper(), "CPF/CNPJ": cnpj.group(1) if cnpj else "",
            "ESPÉCIE": tipos_lista, "TÍTULOS COB.": title_list,
            "DATA VENC.": title_date, "VALORES TIT.": None,
            "TOTAL DEVIDO": total_devido_float,
            "SITUAÇÃO": situation.group(1).strip().upper() if situation else "",
            "COBRADOR": cobrador,
        }

        pc_data = []
        pc_header_pattern = r"Prestação de Contas\s+Nº da PC\s+Dt\. Pgto\."
        pc_section_search = re.search(pc_header_pattern, block, re.IGNORECASE)

        if pc_section_search:
            pc_text_block = block[pc_section_search.start():]
            pc_row_pattern = re.compile(
                r"(\d+)\s+(\d{2}/\d{2}/\d{4})\s+(R\$\s*[\d.,]+)\s+(R\$\s*[\d.,]+)\s+(R\$\s*[\d.,]+)\s+(R\$\s*[\d.,]+)")
            pc_data_raw = pc_row_pattern.findall(pc_text_block)

            if start_date or end_date:
                for entry in pc_data_raw:
                    try:
                        payment_date = datetime.datetime.strptime(entry[1], "%d/%m/%Y")
                        if (not start_date or payment_date >= start_date) and \
                                (not end_date or payment_date <= end_date):
                            pc_data.append(entry)
                    except ValueError:
                        continue
            else:
                pc_data = pc_data_raw

        total_depositado_valor = 0.0
        if pc_data:
            for pc_entry in pc_data:
                total_depositado_valor += clean_and_convert_to_float(pc_entry[5])

        grande_total_valor += total_depositado_valor

        if pc_data:
            first_pc_entry = pc_data[0]
            first_row = base_data.copy()
            first_row.update({
                "Nº DA PC": first_pc_entry[0].strip(),
                "DATA PGTO": first_pc_entry[1].strip(),
                "VALOR CRÉDITO": clean_and_convert_to_float(first_pc_entry[2]),
                "TAXA": clean_and_convert_to_float(first_pc_entry[3]),
                "VALORES DEP.": clean_and_convert_to_float(first_pc_entry[5]),
                "TOTAL CLIENTE": total_depositado_valor,
            })
            linhas.append(first_row)

            for subsequent_pc_entry in pc_data[1:]:
                empty_row = {key: "" for key in ALL_COLUMNS}
                empty_row.update({
                    "Nº DA PC": subsequent_pc_entry[0].strip(),
                    "DATA PGTO": subsequent_pc_entry[1].strip(),
                    "VALOR CRÉDITO": clean_and_convert_to_float(subsequent_pc_entry[2]),
                    "TAXA": clean_and_convert_to_float(subsequent_pc_entry[3]),
                    "VALORES DEP.": clean_and_convert_to_float(subsequent_pc_entry[5]),
                })
                linhas.append(empty_row)
        else:
            if not start_date and not end_date:
                row_data = base_data.copy()
                row_data.update(
                    {"Nº DA PC": "", "DATA PGTO": "", "VALOR CRÉDITO": None, "TAXA": None, "VALORES DEP.": None,
                     "TOTAL CLIENTE": None})
                linhas.append(row_data)

        if gui_instance:
            time.sleep(0.01)
            progress_percent = int(((i + 1) / total_blocks) * 95)
            gui_instance.update_progress(progress_percent)

    df = pd.DataFrame(linhas, columns=ALL_COLUMNS)
    df = df[selected_columns]

    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Relatório', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Relatório']

    currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
    bold_format = workbook.add_format({'bold': True})

    header = df.columns.values.tolist()
    currency_cols = ["TOTAL DEVIDO", "VALOR CRÉDITO", "TAXA", "VALORES DEP.", "TOTAL CLIENTE"]

    for i, col_name in enumerate(header):
        column_len = max(df[col_name].astype(str).map(len).max(), len(col_name))
        adjusted_width = column_len + 2

        if col_name in currency_cols:
            worksheet.set_column(i, i, adjusted_width, currency_format)
        else:
            worksheet.set_column(i, i, adjusted_width)

    if (grande_total_valor > 0 or grande_total_devido_valor > 0):
        last_row = len(df) + 2
        summary_col_label = max(0, len(selected_columns) - 2)
        summary_col_value = max(1, len(selected_columns) - 1)

        worksheet.write(last_row, summary_col_label, "EMPRESA DE COBRANÇA:", bold_format)
        worksheet.write(last_row, summary_col_value, primeira_empresa_info)
        last_row += 1

        worksheet.write(last_row, summary_col_label, "DATA DE EXTRAÇÃO:", bold_format)
        worksheet.write(last_row, summary_col_value, f"{datetime.datetime.now():%d/%m/%Y %H:%M}")
        last_row += 1

        if start_date_str or end_date_str:
            periodo_str = f"{start_date_str} a {end_date_str}" if start_date_str and end_date_str else f"A partir de {start_date_str}" if start_date_str else f"Até {end_date_str}"
            worksheet.write(last_row, summary_col_label, "PERÍODO DO FILTRO:", bold_format)
            worksheet.write(last_row, summary_col_value, periodo_str)
            last_row += 1

        worksheet.write(last_row, summary_col_label, "TOTAL ARRECADADO:", bold_format)
        worksheet.write(last_row, summary_col_value, grande_total_valor, currency_format)
        last_row += 1

        worksheet.write(last_row, summary_col_label, "TOTAL EM COBRANÇA:", bold_format)
        worksheet.write(last_row, summary_col_value, grande_total_devido_valor, currency_format)

    writer.close()


class ExtratorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PlanilhaHTML v7.0")
        try:
            ico = resource_path("iconeprograma.ico")
            self.iconbitmap(ico)
        except tk.TclError:
            print("Ícone 'iconeprograma.ico' não encontrado.")

        w, h = 400, 550
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.resizable(False, False)

        self.input_file = ""
        self.output_dir = ""
        self.filename_var = tk.StringVar()
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.default_name = f"Planilha extraída {datetime.datetime.now():%d-%m-%Y}.xlsx"
        self.filename_var.set(self.default_name)
        self.latest_path = None
        self.is_dirty = False  # New state tracker
        self.column_vars = {col: tk.BooleanVar(value=True) for col in ALL_COLUMNS}
        self.checkboxes = []

        self.create_widgets()
        self._attach_traces()

    def set_dirty(self, *args):
        """Sets the dirty flag and updates button states."""
        if not self.is_dirty:
            self.is_dirty = True
            self._update_button_states()

    def _attach_traces(self):
        """Attach trace callbacks to all relevant variables."""
        self.filename_var.trace_add('write', self.set_dirty)
        self.start_date_var.trace_add('write', self.set_dirty)
        self.end_date_var.trace_add('write', self.set_dirty)
        for var in self.column_vars.values():
            var.trace_add('write', self.set_dirty)

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill='both', expand=True)
        main_frame.columnconfigure(1, weight=1)

        s = ttk.Style()
        s.configure("Accent.TButton", font=('Segoe UI', 10, 'bold'))
        s.configure("Success.TButton", font=('Segoe UI', 9), background='#28a745', foreground='green')
        s.map("Success.TButton", foreground=[('active', 'green')], background=[('active', '#218838')])
        s.configure("Open.TButton", font=('Segoe UI', 9, 'bold'), background='#007bff', foreground='green')
        s.map("Open.TButton", foreground=[('active', 'green')], background=[('active', '#0056b3')])
        s.configure("Ready.TButton", font=('Segoe UI', 10, 'bold'), background='#28a745', foreground='green')
        s.map("Ready.TButton", foreground=[('active', 'green')], background=[('active', '#218838')])

        ttk.Label(main_frame, text="Arquivo HTML:").grid(row=0, column=0, sticky='w', pady=2)
        self.input_label = ttk.Label(main_frame, text="Nenhum arquivo selecionado.", foreground='gray', wraplength=350)
        self.input_label.grid(row=0, column=1, sticky='w')
        self.btn_select_input = ttk.Button(main_frame, text="Selecionar...", command=self.select_input)
        self.btn_select_input.grid(row=0, column=2, padx=5)

        ttk.Label(main_frame, text="Pasta de Saída:").grid(row=1, column=0, sticky='w', pady=2)
        self.output_label = ttk.Label(main_frame, text="Nenhuma pasta selecionada.", foreground='gray', wraplength=350)
        self.output_label.grid(row=1, column=1, sticky='w')
        self.btn_select_output = ttk.Button(main_frame, text="Escolher...", command=self.select_output_dir)
        self.btn_select_output.grid(row=1, column=2, padx=5)

        ttk.Label(main_frame, text="Nome do Arquivo:").grid(row=2, column=0, sticky='w', pady=2)
        ttk.Entry(main_frame, textvariable=self.filename_var).grid(row=2, column=1, columnspan=2, sticky='ew', pady=2)

        date_frame = ttk.Labelframe(main_frame, text="Filtrar por Data de Pagamento (Opcional)")
        date_frame.grid(row=3, column=0, columnspan=3, sticky='ew', pady=(10, 5))
        ttk.Label(date_frame, text="Data Inicial (DD/MM/AAAA):").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(date_frame, textvariable=self.start_date_var).grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        ttk.Label(date_frame, text="Data Final (DD/MM/AAAA):").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(date_frame, textvariable=self.end_date_var).grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        date_frame.columnconfigure(1, weight=1)

        columns_frame = ttk.Labelframe(main_frame, text="Informações a extrair:")
        columns_frame.grid(row=4, column=0, columnspan=3, sticky='ew', pady=(10, 5))

        num_cols = 3
        for i, col_name in enumerate(ALL_COLUMNS):
            row = i // num_cols
            col = i % num_cols
            cb = ttk.Checkbutton(columns_frame, text=col_name, variable=self.column_vars[col_name])
            cb.grid(row=row, column=col, sticky='w', padx=5, pady=1)
            self.checkboxes.append(cb)

        btn_frame = ttk.Frame(columns_frame)
        btn_frame.grid(row=(len(ALL_COLUMNS) // num_cols) + 1, column=0, columnspan=num_cols, pady=5)
        self.btn_select_all = ttk.Button(btn_frame, text="Selecionar Todos", command=self.select_all_columns)
        self.btn_select_all.pack(side='left', padx=5)
        self.btn_deselect_all = ttk.Button(btn_frame, text="Limpar Seleção", command=self.deselect_all_columns)
        self.btn_deselect_all.pack(side='left', padx=5)
        self.btn_summary = ttk.Button(btn_frame, text="RESUMIDO", command=self.select_summary_columns)
        self.btn_summary.pack(side='left', padx=5)

        generate_frame = ttk.Frame(main_frame)
        generate_frame.grid(row=5, column=0, columnspan=3, pady=(10, 5), sticky='ew')
        generate_frame.columnconfigure((0, 1), weight=1)

        self.btn_generate = ttk.Button(generate_frame, text="GERAR RELATÓRIO", command=self.on_generate,
                                       style="Accent.TButton", state='disabled')
        self.btn_generate.grid(row=0, column=0, sticky='ew', padx=2)

        self.btn_update = ttk.Button(generate_frame, text="ATUALIZAR RELATÓRIO", command=self.on_update,
                                     style="Accent.TButton", state='disabled')
        self.btn_update.grid(row=0, column=1, sticky='ew', padx=2)

        self.progress = ttk.Progressbar(main_frame, mode='determinate', maximum=100)
        self.progress.grid(row=6, column=0, columnspan=3, sticky='ew', pady=(5, 10))

        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=7, column=0, columnspan=3, sticky='ew')
        action_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_clear = ttk.Button(action_frame, text="Limpar", command=self.reset)
        self.btn_clear.grid(row=0, column=0, sticky='ew', padx=2)
        self.btn_open = ttk.Button(action_frame, text="Abrir Planilha", command=self.open_file, state='disabled')
        self.btn_open.grid(row=0, column=1, sticky='ew', padx=2)
        ttk.Button(action_frame, text="Sair", command=self.quit).grid(row=0, column=2, sticky='ew', padx=2)

        ttk.Label(main_frame, text="Versão 7.0 - Desenvolvido por Pedro Luz", font=('Segoe UI', 11, 'italic'),
                  foreground='purple').grid(row=8, column=0, columnspan=3, pady=(15, 0))

    def select_all_columns(self):
        for var in self.column_vars.values():
            var.set(True)
        self.set_dirty()

    def deselect_all_columns(self):
        for var in self.column_vars.values():
            var.set(False)
        self.set_dirty()

    def select_summary_columns(self):
        summary_cols = [
            "NOME CLIENTE", "CPF/CNPJ", "TOTAL DEVIDO", "SITUAÇÃO",
            "DATA PGTO", "VALORES DEP.", "TOTAL CLIENTE"
        ]
        for col, var in self.column_vars.items():
            if col in summary_cols:
                var.set(True)
            else:
                var.set(False)
        self.set_dirty()

    def _update_button_states(self):
        if self.input_file and self.output_dir:
            self.btn_generate.config(state='normal', style="Ready.TButton")
        else:
            self.btn_generate.config(state='disabled', style="Accent.TButton")

        if self.latest_path and self.is_dirty:
            self.btn_update.config(state='normal', style="Ready.TButton")
        else:
            self.btn_update.config(state='disabled', style="Accent.TButton")

    def select_input(self):
        path = filedialog.askopenfilename(title="Selecione o arquivo HTML",
                                          filetypes=[("Arquivos HTML", "*.html;*.htm")])
        if path:
            self.input_file = path
            self.input_label.config(text=os.path.basename(path), foreground='black')
            self.btn_select_input.config(text="Selecionado", style="Success.TButton")
            self._update_button_states()

    def select_output_dir(self):
        directory = filedialog.askdirectory(title="Escolha a pasta de saída")
        if directory:
            self.output_dir = directory
            self.output_label.config(text=os.path.basename(directory), foreground='black')
            self.btn_select_output.config(text="Escolhido", style="Success.TButton")
            self._update_button_states()

    def on_generate(self):
        if not self._validate_inputs():
            return

        filename = self.filename_var.get().strip()
        if not filename.lower().endswith((".xlsx", ".xls")):
            filename += ".xlsx"

        base, ext = os.path.splitext(filename)
        count = 1
        self.latest_path = os.path.join(self.output_dir, filename)
        while os.path.exists(self.latest_path):
            self.latest_path = os.path.join(self.output_dir, f"{base} ({count}){ext}")
            count += 1

        self._start_worker()

    def on_update(self):
        if not self._validate_inputs():
            return
        if not self.latest_path:
            messagebox.showerror("Erro", "Nenhum relatório foi gerado para atualizar.")
            return

        self._start_worker()

    def _validate_inputs(self):
        selected_columns = [col for col, var in self.column_vars.items() if var.get()]
        if not selected_columns:
            messagebox.showerror("Erro de Validação", "Selecione pelo menos uma coluna para exportar.")
            return False

        if not self.input_file or not self.output_dir or not self.filename_var.get().strip():
            messagebox.showerror("Erro de Validação", "Todos os campos devem ser preenchidos.")
            return False
        return True

    def _start_worker(self):
        self.btn_generate.config(state='disabled')
        self.btn_update.config(state='disabled')
        self.btn_clear.config(state='disabled')
        self.btn_open.config(state='disabled', style="TButton")
        self.progress['value'] = 0

        selected_columns = [col for col, var in self.column_vars.items() if var.get()]
        start_date = self.start_date_var.get()
        end_date = self.end_date_var.get()

        threading.Thread(target=self._worker_thread, args=(selected_columns, start_date, end_date), daemon=True).start()

    def _worker_thread(self, selected_columns, start_date, end_date):
        try:
            processar_arquivo(self.input_file, self.latest_path, selected_columns, start_date, end_date, self)
            self.update_progress(100)
            self.after(200, self._on_complete)
        except Exception as e:
            self.after(0, lambda e=e: messagebox.showerror("Erro no Processamento", f"Ocorreu um erro:\n{e}"))
        finally:
            self.after(0, self.reset_ui_after_task)

    def reset_ui_after_task(self):
        self.btn_clear.config(state='normal')
        self._update_button_states()

    def _on_complete(self):
        self.is_dirty = False  # Reset dirty flag after successful generation/update
        messagebox.showinfo("Sucesso", f"Planilha gerada com sucesso!\n\nSalvo em: {self.latest_path}")
        self.btn_open.config(state='normal', style="Open.TButton")
        self._update_button_states()

    def update_progress(self, value):
        self.after(0, lambda: self.progress.config(value=value))

    def open_file(self):
        if self.latest_path:
            abrir_arquivo(self.latest_path)

    def reset(self):
        self.input_file = ""
        self.output_dir = ""
        self.latest_path = None
        self.is_dirty = False
        self.input_label.config(text="Nenhum arquivo selecionado.", foreground='gray')
        self.output_label.config(text="Nenhuma pasta selecionada.", foreground='gray')
        self.filename_var.set(self.default_name)
        self.start_date_var.set("")
        self.end_date_var.set("")
        self.progress['value'] = 0
        self.btn_open.config(state='disabled', style="TButton")
        self.btn_select_input.config(text="Selecionar...", style="TButton")
        self.btn_select_output.config(text="Escolher...", style="TButton")
        self._update_button_states()


if __name__ == "__main__":
    app = ExtratorGUI()
    app.mainloop()