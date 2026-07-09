"""
Excel Exploder - v2
Aplicativo desktop para "explodir" colunas de um arquivo Excel
(transformar valores separados por um caractere em várias linhas).

Novidades da v2:
    - Pré-visualização (linhas originais x linhas após explode)
    - Seleção de múltiplas colunas
    - Arrastar e soltar arquivo (drag and drop)
    - Histórico de processamentos

Dependências:
    pip install pandas openpyxl tkinterdnd2

Observação sobre o drag-and-drop:
    O drag-and-drop depende da biblioteca "tkinterdnd2", que não faz parte
    do Python padrão. Se ela não estiver instalada, o app funciona
    normalmente, só que sem a área de arrastar arquivo (usa-se o botão
    "Selecionar Arquivo" normalmente).

Para gerar o executável (.exe) no Windows:
    pip install pyinstaller pandas openpyxl tkinterdnd2
    pyinstaller --onefile --windowed --name ExplodeExcel excel_exploder_v2.py
O .exe aparecerá na pasta "dist".
"""

import os
import json
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

# --- Drag and drop é opcional ---------------------------------------------
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
    BaseTk = TkinterDnD.Tk
except ImportError:
    DND_AVAILABLE = False
    BaseTk = tk.Tk

HISTORY_FILE = os.path.join(os.path.expanduser("~"), ".excel_exploder_historico.json")


# --------------------------------------------------------------------------
def load_history():
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_history_entry(entry):
    history = load_history()
    history.insert(0, entry)
    history = history[:200]  # mantém no máximo 200 registros
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def explode_columns(df, columns, separador):
    """Explode uma ou mais colunas sequencialmente."""
    result = df.copy()
    for col in columns:
        result[col] = result[col].astype(str).str.split(separador)
        result = result.explode(col)
        result[col] = result[col].str.strip()
    return result.reset_index(drop=True)


# --------------------------------------------------------------------------
class ExcelExploderApp(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Exploder")
        self.geometry("520x560")
        self.resizable(False, False)

        self.filepath = None
        self.df = None
        self.column_vars = {}  # nome_coluna -> BooleanVar

        self._build_ui()

    # ---------------------------------------------------------------- UI
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # --- Seleção de arquivo / drag and drop ---
        frame_file = ttk.LabelFrame(self, text="1. Arquivo")
        frame_file.pack(fill="x", **pad)

        drop_text = ("Arraste o arquivo Excel aqui\n(ou use o botão abaixo)"
                     if DND_AVAILABLE else
                     "Selecione o arquivo Excel abaixo\n(arrastar e soltar indisponível: instale tkinterdnd2)")

        self.drop_label = tk.Label(frame_file, text=drop_text, relief="groove",
                                    height=3, bg="#f5f5f5", fg="#555")
        self.drop_label.pack(fill="x", padx=8, pady=(8, 4))

        if DND_AVAILABLE:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_drop)

        row = ttk.Frame(frame_file)
        row.pack(fill="x", padx=8, pady=(0, 8))
        self.file_entry = ttk.Entry(row, state="readonly")
        self.file_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Selecionar Arquivo",
                   command=self.select_file).pack(side="left", padx=(8, 0))

        # --- Colunas (múltipla seleção) ---
        frame_col = ttk.LabelFrame(self, text="2. Colunas para Explode")
        frame_col.pack(fill="both", **pad)

        canvas = tk.Canvas(frame_col, height=140, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_col, orient="vertical", command=canvas.yview)
        self.columns_frame = ttk.Frame(canvas)

        self.columns_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.columns_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scrollbar.pack(side="right", fill="y", pady=8, padx=(0, 8))

        # --- Separador ---
        frame_sep = ttk.LabelFrame(self, text="3. Separador")
        frame_sep.pack(fill="x", **pad)

        self.sep_var = tk.StringVar(value=";")
        options_frame = ttk.Frame(frame_sep)
        options_frame.pack(fill="x", padx=8, pady=4)

        for label, value in [(";", ";"), (",", ","), ("|", "|"), ("Outro", "OUTRO")]:
            ttk.Radiobutton(options_frame, text=label, value=value,
                             variable=self.sep_var,
                             command=self._toggle_custom_sep).pack(side="left", padx=4)

        self.custom_sep_entry = ttk.Entry(frame_sep, width=10, state="disabled")
        self.custom_sep_entry.pack(padx=8, pady=(0, 8))

        # --- Botões: preview / processar / histórico ---
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Pré-visualizar", command=self.preview).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="PROCESSAR", command=self.process).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Histórico", command=self.show_history).pack(side="left", padx=6)

        # --- Pré-visualização ---
        frame_preview = ttk.LabelFrame(self, text="Pré-visualização")
        frame_preview.pack(fill="x", **pad)
        self.preview_var = tk.StringVar(value="Selecione colunas e separador, depois clique em Pré-visualizar.")
        ttk.Label(frame_preview, textvariable=self.preview_var, wraplength=460).pack(
            fill="x", padx=8, pady=8)

        # --- Status ---
        self.status_var = tk.StringVar(value="Selecione um arquivo para começar.")
        ttk.Label(self, textvariable=self.status_var, wraplength=480,
                  foreground="#333").pack(fill="x", padx=12, pady=4)

    def _toggle_custom_sep(self):
        if self.sep_var.get() == "OUTRO":
            self.custom_sep_entry.config(state="normal")
        else:
            self.custom_sep_entry.delete(0, tk.END)
            self.custom_sep_entry.config(state="disabled")

    # ------------------------------------------------------------ Ações
    def _on_drop(self, event):
        # event.data pode vir entre chaves se o caminho tiver espaços
        path = event.data.strip("{}")
        if not path.lower().endswith((".xlsx", ".xls")):
            messagebox.showwarning("Atenção", "Solte um arquivo .xlsx ou .xls")
            return
        self._load_file(path)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Selecione um arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if path:
            self._load_file(path)

    def _load_file(self, path):
        try:
            df = pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("Erro ao abrir arquivo", f"Não foi possível ler o arquivo.\n\n{e}")
            return

        self.filepath = path
        self.df = df

        self.file_entry.config(state="normal")
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, os.path.basename(path))
        self.file_entry.config(state="readonly")

        # Recria os checkbuttons de colunas
        for widget in self.columns_frame.winfo_children():
            widget.destroy()
        self.column_vars = {}
        for col in df.columns:
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.columns_frame, text=str(col), variable=var)
            cb.pack(anchor="w", padx=4, pady=2)
            self.column_vars[col] = var

        self.preview_var.set("Selecione colunas e separador, depois clique em Pré-visualizar.")
        self.status_var.set(f"Arquivo carregado: {len(df)} linhas, {len(df.columns)} colunas.")

    def get_separator(self):
        sep = self.sep_var.get()
        if sep == "OUTRO":
            return self.custom_sep_entry.get()
        return sep

    def get_selected_columns(self):
        return [col for col, var in self.column_vars.items() if var.get()]

    def _validate(self):
        if self.filepath is None or self.df is None:
            messagebox.showwarning("Atenção", "Selecione um arquivo antes de continuar.")
            return None, None
        columns = self.get_selected_columns()
        if not columns:
            messagebox.showwarning("Atenção", "Selecione uma coluna.")
            return None, None
        separador = self.get_separator()
        if not separador:
            messagebox.showwarning("Atenção", "Informe um separador válido.")
            return None, None
        return columns, separador

    def preview(self):
        columns, separador = self._validate()
        if columns is None:
            return
        try:
            df_final = explode_columns(self.df, columns, separador)
            self.preview_var.set(
                f"Linhas originais: {len(self.df)}\n"
                f"Linhas após explode: {len(df_final)}\n"
                f"Colunas selecionadas: {', '.join(str(c) for c in columns)}"
            )
        except Exception as e:
            messagebox.showerror("Erro na pré-visualização", str(e))

    def process(self):
        columns, separador = self._validate()
        if columns is None:
            return

        try:
            start = time.time()
            df_final = explode_columns(self.df, columns, separador)

            base, ext = os.path.splitext(self.filepath)
            output_path = f"{base}_explodido{ext}"
            df_final.to_excel(output_path, index=False)

            elapsed = time.time() - start

            self.status_var.set(
                f"✅ Processo concluído em {elapsed:.1f}s. Arquivo salvo em: {output_path}"
            )
            self.preview_var.set(
                f"Linhas originais: {len(self.df)}\n"
                f"Linhas após explode: {len(df_final)}\n"
                f"Colunas selecionadas: {', '.join(str(c) for c in columns)}"
            )

            save_history_entry({
                "data": time.strftime("%Y-%m-%d %H:%M:%S"),
                "arquivo_processado": os.path.basename(self.filepath),
                "colunas": columns,
                "separador": separador,
                "linhas_geradas": len(df_final),
                "arquivo_saida": output_path,
            })

            messagebox.showinfo("Sucesso", f"✅ Processo concluído com sucesso.\n\nArquivo salvo em:\n{output_path}")

        except PermissionError:
            messagebox.showerror("Erro", "Não foi possível salvar o arquivo.\n\nVerifique se o Excel está aberto.")
        except Exception as e:
            messagebox.showerror("Erro inesperado", str(e))

    def show_history(self):
        history = load_history()

        win = tk.Toplevel(self)
        win.title("Histórico de Processamentos")
        win.geometry("620x360")

        columns = ("data", "arquivo", "colunas", "linhas")
        tree = ttk.Treeview(win, columns=columns, show="headings")
        tree.heading("data", text="Data")
        tree.heading("arquivo", text="Arquivo processado")
        tree.heading("colunas", text="Colunas")
        tree.heading("linhas", text="Linhas geradas")
        tree.column("data", width=140)
        tree.column("arquivo", width=180)
        tree.column("colunas", width=160)
        tree.column("linhas", width=100, anchor="center")
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        if not history:
            tree.insert("", "end", values=("-", "Nenhum processamento ainda", "-", "-"))
        else:
            for entry in history:
                tree.insert("", "end", values=(
                    entry.get("data", "-"),
                    entry.get("arquivo_processado", "-"),
                    ", ".join(str(c) for c in entry.get("colunas", [])),
                    entry.get("linhas_geradas", "-"),
                ))


if __name__ == "__main__":
    app = ExcelExploderApp()
    app.mainloop()
