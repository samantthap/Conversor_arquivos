
from importar import (
    os, threading, tk, ttk, filedialog, messagebox, pd,
    win32com, PDF2DOCX_Converter, Document, tabula, pdfplumber
)

def pdf_to_docx(pdf, docx_out):
    if not PDF2DOCX_Converter:
        raise Exception("O módulo pdf2docx não está instalado.")
    conv = PDF2DOCX_Converter(pdf)
    conv.convert(docx_out)
    conv.close()


def pdf_to_excel(pdf, xlsx_out):
    # 1º tentativa: tabula
    if tabula:
        try:
            tabelas = tabula.read_pdf(pdf, pages="all", multiple_tables=True)
            if not tabelas:
                raise Exception("Nenhuma tabela encontrada com tabula.")

            with pd.ExcelWriter(xlsx_out) as writer:
                for i, df in enumerate(tabelas, start=1):
                    df.to_excel(writer, f"Tabela_{i}", index=False)
            return
        except Exception as e:
            print("Erro no tabula:", e)

    # 2º tentativa: pdfplumber
    if pdfplumber:
        try:
            with pdfplumber.open(pdf) as pdf_file:
                encontrou = 0
                with pd.ExcelWriter(xlsx_out) as writer:
                    for p, page in enumerate(pdf_file.pages, start=1):
                        tables = page.extract_tables()
                        for t, table in enumerate(tables, start=1):
                            df = pd.DataFrame(table[1:], columns=table[0])
                            encontrou += 1
                            df.to_excel(writer, f"P{p}_T{t}", index=False)

                if encontrou == 0:
                    raise Exception("Nenhuma tabela encontrada com pdfplumber.")
                return
        except Exception as e:
            raise Exception(f"Erro ao extrair tabelas: {e}")

    raise Exception("Nem tabula nem pdfplumber estão disponíveis.")


def excel_to_docx(xlsx, docx_out):
    if not Document:
        raise Exception("python-docx não está instalado.")

    planilhas = pd.read_excel(xlsx, sheet_name=None)
    doc = Document()

    for nome, df in planilhas.items():
        doc.add_heading(str(nome), level=2)

        if df.empty:
            doc.add_paragraph("(Planilha vazia)")
            continue

        tabela = doc.add_table(rows=1, cols=len(df.columns))
        hdr = tabela.rows[0].cells

        for i, col in enumerate(df.columns):
            hdr[i].text = str(col)

        for _, linha in df.iterrows():
            row_cells = tabela.add_row().cells
            for i, item in enumerate(linha):
                row_cells[i].text = "" if pd.isna(item) else str(item)

        doc.add_page_break()

    doc.save(docx_out)


def docx_to_excel(docx, xlsx_out):
    if not Document:
        raise Exception("python-docx não está instalado.")

    doc = Document(docx)
    tabelas = doc.tables

    if not tabelas:
        raise Exception("Nenhuma tabela encontrada no DOCX.")

    with pd.ExcelWriter(xlsx_out) as writer:
        for i, tbl in enumerate(tabelas, start=1):
            linhas = []
            headers = None

            for r, row in enumerate(tbl.rows):
                valores = [cell.text for cell in row.cells]
                if r == 0:
                    headers = valores
                else:
                    linhas.append(valores)

            df = pd.DataFrame(linhas, columns=headers)
            df.to_excel(writer, f"Tabela_{i}", index=False)


def docx_to_pdf_com(docx, pdf_out):
    if not win32com:
        raise Exception("pywin32 não instalado.")

    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False

    try:
        doc = word.Documents.Open(os.path.abspath(docx))
        doc.SaveAs(os.path.abspath(pdf_out), FileFormat=17)
        doc.Close()
    finally:
        word.Quit()


def excel_to_pdf_com(xlsx, pdf_out):
    if not win32com:
        raise Exception("pywin32 não instalado.")

    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False

    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx), ReadOnly=True)
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_out))
        wb.Close()
    finally:
        excel.Quit()


# -----------------------
# Funções auxiliares
# -----------------------

FORMATOS = ["PDF", "Word", "Excel"]

def gerar_nome_saida(caminho, formato_out):
    base = os.path.splitext(os.path.basename(caminho))[0]
    pasta = os.path.dirname(caminho)

    ext = {
        "PDF": ".pdf",
        "Word": ".docx",
        "Excel": ".xlsx"
    }[formato_out]

    return os.path.join(pasta, f"{base}_convertido{ext}")


# -----------------------
# Interface
# -----------------------

class ConverterApp:
    def __init__(self, root):
        self.root = root
        root.title("Conversor de Arquivos")
        root.geometry("700x260")

        frame = ttk.Frame(root, padding=12)
        frame.pack(fill="both", expand=True)

        # --- Entrada do caminho ---
        ttk.Label(frame, text="Arquivo ou pasta:").grid(row=0, column=0, sticky="w")

        self.path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.path_var, width=50).grid(row=0, column=1, columnspan=2, sticky="w")

        ttk.Button(frame, text="Selecionar arquivo", command=self.selecionar_arquivo).grid(row=0, column=3, padx=5)
        ttk.Button(frame, text="Selecionar pasta", command=self.selecionar_pasta).grid(row=0, column=4, padx=5)

        # --- Formatos ---
        ttk.Label(frame, text="Formato de entrada:").grid(row=1, column=0, sticky="w", pady=10)
        self.input_cb = ttk.Combobox(frame, values=FORMATOS, state="readonly")
        self.input_cb.grid(row=1, column=1, sticky="w")
        self.input_cb.current(0)

        ttk.Label(frame, text="Formato de saída:").grid(row=1, column=2, sticky="w")
        self.output_cb = ttk.Combobox(frame, values=FORMATOS, state="readonly")
        self.output_cb.grid(row=1, column=3, sticky="w")
        self.output_cb.current(1)

        self.batch_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Processar toda a pasta (se uma pasta selecionada)", variable=self.batch_var).grid(row=2, column=1, sticky="w")

        # Botão converter
        self.btn_convert = ttk.Button(frame, text="Converter", command=self.iniciar)
        self.btn_convert.grid(row=3, column=1, pady=10)

        # Log
        self.log = tk.Text(frame, height=6)
        self.log.grid(row=4, column=0, columnspan=5, sticky="we")

    def logar(self, texto):
        self.log.config(state="normal")
        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def selecionar_arquivo(self):
        arq = filedialog.askopenfilename()
        if arq:
            self.path_var.set(arq)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.path_var.set(pasta)

    def iniciar(self):
        caminho = self.path_var.get().strip()
        if not caminho:
            messagebox.showerror("Erro", "Selecione um arquivo ou pasta.")
            return

        in_fmt = self.input_cb.get()
        out_fmt = self.output_cb.get()

        self.btn_convert.config(state="disabled")

        threading.Thread(
            target=self.executar_conversao,
            args=(caminho, in_fmt, out_fmt, self.batch_var.get()),
            daemon=True
        ).start()

    def executar_conversao(self, caminho, in_fmt, out_fmt, em_lote):
        try:
            if os.path.isdir(caminho):
                if not em_lote:
                    self.logar("Marque a opção de lote para converter a pasta inteira.")
                    return

                ext = {"PDF": ".pdf", "Word": ".docx", "Excel": ".xlsx"}[in_fmt]

                arquivos = [
                    os.path.join(caminho, f)
                    for f in os.listdir(caminho)
                    if f.lower().endswith(ext)
                ]

                if not arquivos:
                    self.logar(f"Nenhum arquivo {ext} encontrado.")
                    return

                for arq in arquivos:
                    self.converter_unico(arq, in_fmt, out_fmt)

                self.logar("Conversão concluída.")
            else:
                self.converter_unico(caminho, in_fmt, out_fmt)
                self.logar("Conversão finalizada.")

        finally:
            self.btn_convert.config(state="normal")

    def converter_unico(self, entrada, in_fmt, out_fmt):
        self.logar(f"Convertendo: {os.path.basename(entrada)}")

        saida = gerar_nome_saida(entrada, out_fmt)

        try:
            # Mesmos formatos → só copiar
            if in_fmt == out_fmt:
                import shutil
                shutil.copy(entrada, saida)
                self.logar("Arquivo copiado (mesmo formato).")
                return

            # PDF → Word
            if in_fmt == "PDF" and out_fmt == "Word":
                pdf_to_docx(entrada, saida)

            # PDF → Excel
            elif in_fmt == "PDF" and out_fmt == "Excel":
                pdf_to_excel(entrada, saida)

            # Word → PDF
            elif in_fmt == "Word" and out_fmt == "PDF":
                docx_to_pdf_com(entrada, saida)

            # Excel → PDF
            elif in_fmt == "Excel" and out_fmt == "PDF":
                excel_to_pdf_com(entrada, saida)

            # Excel → Word
            elif in_fmt == "Excel" and out_fmt == "Word":
                excel_to_docx(entrada, saida)

            # Word → Excel
            elif in_fmt == "Word" and out_fmt == "Excel":
                docx_to_excel(entrada, saida)

            else:
                raise Exception("Conversão ainda não suportada.")

            self.logar(f"OK: {saida}")

        except Exception as e:
            if os.path.exists(saida):
                os.remove(saida)
            self.logar(f"ERRO: {e}")




def main():
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
