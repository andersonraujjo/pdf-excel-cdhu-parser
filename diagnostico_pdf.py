import pdfplumber
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import re


# ──────────────────────────────────────────────
#  MOTOR DE EXTRAÇÃO
# ──────────────────────────────────────────────

# Limites X calibrados pelo diagnóstico (página 612pt de largura)
X_ITEM_FIM    = 100   # Item:       0  → 100
X_CODIGO_FIM  = 145   # Código:   100  → 145
X_SERVICO_FIM = 400   # Serviço:  145  → 400
X_UN_FIM      = 430   # Un:       400  → 430
X_QTDE_FIM    = 490   # Qtde:     430  → 490
X_UNIT_FIM    = 535   # Unitário: 490  → 535
                      # Total:    535  → 612

COLUNAS = ["Item", "Código", "Serviço", "Un", "Qtde", "Unitário", "Total"]

# Padrões de cabeçalho/rodapé — linhas que combinam são descartadas
_RE_CABECALHO = re.compile(
    r"planilha\s+de\s+pre"
    r"|p[aá]gina\s+\d"
    r"|empreendimento"
    r"|data\s+base"
    r"|projeto\s*:"
    r"|^\s*bdi\s*:"
    r"|obs\s*:"
    r"|pre[çc]os\s+unit[aá]rios"
    r"|\bcdhu\b"
    r"|\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}",
    re.IGNORECASE,
)

# Linhas acima deste Y são cabeçalho fixo da página (logo, título, etc.)
Y_DADOS_INI = 130


def _is_item(texto):
    return bool(re.match(r"^\d{3,4}(\.\d{2})*$", texto.strip()))


def _num_br(texto):
    texto = texto.strip()
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    return texto


def _agrupar_por_y(palavras, tolerancia=5.0):
    grupos = {}
    for w in palavras:
        y = (w["top"] + w["bottom"]) / 2
        chave = next((k for k in grupos if abs(k - y) <= tolerancia), None)
        if chave is None:
            chave = round(y, 1)
        grupos.setdefault(chave, []).append(w)
    return grupos


def _e_cabecalho(tokens):
    texto = " ".join(tokens).strip()
    if not texto:
        return False
    return bool(_RE_CABECALHO.search(texto))


def extrair_linhas(pagina):
    palavras = pagina.extract_words(
        x_tolerance=4, y_tolerance=4,
        keep_blank_chars=False, use_text_flow=False,
    )
    # Descarta cabeçalho fixo da página
    palavras = [w for w in palavras if w["top"] >= Y_DADOS_INI]

    grupos = _agrupar_por_y(palavras, tolerancia=5.0)
    linhas = []

    for y_key in sorted(grupos):
        ws = sorted(grupos[y_key], key=lambda w: w["x0"])
        item_tok, cod_tok, serv_tok, un_tok, qtde_tok, unit_tok, total_tok = \
            [], [], [], [], [], [], []

        for w in ws:
            xc = (w["x0"] + w["x1"]) / 2
            if xc < X_ITEM_FIM:
                item_tok.append(w["text"])
            elif xc < X_CODIGO_FIM:
                cod_tok.append(w["text"])
            elif xc < X_SERVICO_FIM:
                serv_tok.append(w["text"])
            elif xc < X_UN_FIM:
                un_tok.append(w["text"])
            elif xc < X_QTDE_FIM:
                qtde_tok.append(w["text"])
            elif xc < X_UNIT_FIM:
                unit_tok.append(w["text"])
            else:
                total_tok.append(w["text"])

        item_str = " ".join(item_tok).strip()

        if not _is_item(item_str):
            # Descarta se for cabeçalho/rodapé
            if _e_cabecalho(serv_tok + cod_tok + item_tok):
                continue
            # Linha de continuação legítima
            texto_cont = " ".join(serv_tok + un_tok + qtde_tok).strip()
            if texto_cont:
                linhas.append({
                    "Item": "", "Código": "", "Serviço": texto_cont,
                    "Un": "", "Qtde": "",
                    "Unitário": _num_br(" ".join(unit_tok)) if unit_tok else "",
                    "Total":    _num_br(" ".join(total_tok)) if total_tok else "",
                })
            continue

        linhas.append({
            "Item":     item_str,
            "Código":   " ".join(cod_tok).strip(),
            "Serviço":  " ".join(serv_tok).strip(),
            "Un":       " ".join(un_tok).strip().upper(),
            "Qtde":     _num_br(" ".join(qtde_tok)) if qtde_tok else "",
            "Unitário": _num_br(" ".join(unit_tok)) if unit_tok else "",
            "Total":    _num_br(" ".join(total_tok)) if total_tok else "",
        })

    return linhas


def _merge_multiline(linhas):
    resultado = []
    for linha in linhas:
        if not linha["Item"] and resultado:
            ultimo = resultado[-1]
            if linha["Serviço"]:
                ultimo["Serviço"] = (ultimo["Serviço"] + " " + linha["Serviço"]).strip()
            if linha["Unitário"] and not ultimo["Unitário"]:
                ultimo["Unitário"] = linha["Unitário"]
            if linha["Total"] and not ultimo["Total"]:
                ultimo["Total"] = linha["Total"]
        else:
            resultado.append(dict(linha))
    return resultado


def converter_pdf(caminho_pdf, callback_progresso=None):
    dados = []
    with pdfplumber.open(caminho_pdf) as pdf:
        total = len(pdf.pages)
        for i, pagina in enumerate(pdf.pages):
            dados.extend(extrair_linhas(pagina))
            if callback_progresso:
                callback_progresso((i + 1) / total)

    dados = _merge_multiline(dados)
    dados = [d for d in dados if d.get("Item")]

    if not dados:
        raise ValueError("Nenhum dado válido encontrado no PDF.")

    df = pd.DataFrame(dados, columns=COLUNAS)
    for col in ("Qtde", "Unitário", "Total"):
        df[col] = pd.to_numeric(df[col], errors="coerce")

    caminho_saida = caminho_pdf.replace(".pdf", "_CDHU.xlsx")
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilha")
        ws = writer.sheets["Planilha"]
        for col_cells in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 60)

    return caminho_saida


# ──────────────────────────────────────────────
#  INTERFACE GRÁFICA
# ──────────────────────────────────────────────

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Conversor CDHU Precision - Estabilidade Total")
        self.geometry("520x360")
        self.resizable(False, False)
        self._arquivo_pdf = ""
        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self, text="CONVERSOR CDHU PRECISION",
                     font=("Roboto", 20, "bold")).grid(row=0, column=0, pady=(28, 10))
        self._btn_sel = ctk.CTkButton(self, text="Selecionar PDF",
                                      width=200, command=self._selecionar)
        self._btn_sel.grid(row=1, column=0, pady=8)
        self._lbl_arquivo = ctk.CTkLabel(self, text="Aguardando arquivo...",
                                         wraplength=420, text_color="gray70")
        self._lbl_arquivo.grid(row=2, column=0, pady=4)
        self._progress = ctk.CTkProgressBar(self, width=380)
        self._progress.set(0)
        self._lbl_status = ctk.CTkLabel(self, text="", text_color="gray60")
        self._btn_conv = ctk.CTkButton(
            self, text="Converter para Excel", width=200,
            fg_color="#28a745", hover_color="#1e7e34",
            state="disabled", command=self._iniciar_conversao,
        )
        self._btn_conv.grid(row=5, column=0, pady=24)

    def _selecionar(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if caminho:
            self._arquivo_pdf = caminho
            self._lbl_arquivo.configure(text=os.path.basename(caminho), text_color="white")
            self._btn_conv.configure(state="normal")

    def _iniciar_conversao(self):
        self._btn_conv.configure(state="disabled")
        self._btn_sel.configure(state="disabled")
        self._progress.set(0)
        self._progress.grid(row=3, column=0, pady=6)
        self._lbl_status.configure(text="Processando…")
        self._lbl_status.grid(row=4, column=0)
        threading.Thread(target=self._processar, daemon=True).start()

    def _processar(self):
        try:
            saida = converter_pdf(self._arquivo_pdf, self._atualizar_progresso)
            self.after(0, lambda: self._finalizar(True, saida))
        except Exception as exc:
            self.after(0, lambda: self._finalizar(False, str(exc)))

    def _atualizar_progresso(self, pct):
        self.after(0, lambda: self._progress.set(pct))
        self.after(0, lambda: self._lbl_status.configure(text=f"Processando… {int(pct * 100)}%"))

    def _finalizar(self, sucesso, mensagem):
        self._progress.grid_forget()
        self._lbl_status.grid_forget()
        self._btn_conv.configure(state="normal")
        self._btn_sel.configure(state="normal")
        if sucesso:
            messagebox.showinfo("Sucesso ✔", f"Excel gerado em:\n{mensagem}")
        else:
            messagebox.showerror("Erro", f"Falha no processamento:\n{mensagem}")


if __name__ == "__main__":
    app = App()
    app.mainloop()