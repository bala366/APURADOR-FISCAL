import os
import re
import sys
import sqlite3
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.units import mm
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

APP_TITLE = "Apurador Fiscal - Lucro Real V5"
APP_EXE_NAME = "Apurador Fiscal"


def app_data_dir() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
        p = base / "Apurador Fiscal"
    else:
        p = Path.home() / ".apurador_fiscal"
    p.mkdir(parents=True, exist_ok=True)
    return p

DB = app_data_dir() / "apurador_fiscal.db"


def D(v):
    try:
        return Decimal(str(v or "0").replace(".", "").replace(",", ".") if isinstance(v, str) and "," in v else str(v or "0"))
    except (InvalidOperation, ValueError):
        return Decimal("0")


def money(v):
    x = Decimal(str(v or 0))
    return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def digits(s):
    return re.sub(r"\D", "", s or "")


def fmt_cnpj(s):
    s = digits(s)
    if len(s) != 14:
        return s
    return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"


def local(tag):
    return tag.split("}", 1)[-1]


def child_text(el, name, default=""):
    if el is None:
        return default
    for c in el.iter():
        if local(c.tag) == name:
            return (c.text or default).strip()
    return default


def direct_child(el, name):
    if el is None:
        return None
    for c in list(el):
        if local(c.tag) == name:
            return c
    return None


def first_desc(el, name):
    if el is None:
        return None
    for c in el.iter():
        if local(c.tag) == name:
            return c
    return None


def sum_named(el, names):
    total = Decimal("0")
    if el is None:
        return total
    for c in el.iter():
        if local(c.tag) in names:
            try:
                total += Decimal(c.text or "0")
            except Exception:
                pass
    return total


def parse_competencia(text):
    m = re.fullmatch(r"\s*(0[1-9]|1[0-2])/(\d{4})\s*", text or "")
    if not m:
        raise ValueError("Competência deve estar em MM/AAAA, por exemplo 08/2026.")
    mes, ano = m.group(1), m.group(2)
    return f"{ano}-{mes}"


class DBX:
    def __init__(self):
        self.c = sqlite3.connect(DB)
        self.c.execute("""CREATE TABLE IF NOT EXISTS config(
            id INTEGER PRIMARY KEY CHECK(id=1), empresa TEXT, cnpj TEXT, regime TEXT,
            saldo_pis REAL, saldo_cofins REAL, saldo_icms REAL)""")
        self.c.execute("""CREATE TABLE IF NOT EXISTS docs(
            chave TEXT PRIMARY KEY, arquivo TEXT, data TEXT, tipo TEXT, emit_cnpj TEXT, dest_cnpj TEXT,
            emit_nome TEXT, dest_nome TEXT, valor REAL, pis REAL, cofins REAL, icms REAL, itens INTEGER)""")
        self.c.execute("""CREATE TABLE IF NOT EXISTS itens(
            id INTEGER PRIMARY KEY AUTOINCREMENT, chave TEXT, data TEXT, tipo TEXT,
            codigo TEXT, descricao TEXT, ncm TEXT, cfop TEXT, cst_pis TEXT, cst_cofins TEXT,
            cst_icms TEXT, valor REAL, pis REAL, cofins REAL, icms REAL)""")
        self._migrate()
        self.c.commit()
        if not self.c.execute("SELECT 1 FROM config WHERE id=1").fetchone():
            self.c.execute("INSERT INTO config VALUES(1,'','','Lucro Real',0,0,0)")
            self.c.commit()

    def _migrate(self):
        cols = {r[1] for r in self.c.execute("PRAGMA table_info(docs)")}
        if "emit_nome" not in cols:
            self.c.execute("ALTER TABLE docs ADD COLUMN emit_nome TEXT DEFAULT ''")
        if "dest_nome" not in cols:
            self.c.execute("ALTER TABLE docs ADD COLUMN dest_nome TEXT DEFAULT ''")

    def config(self):
        return self.c.execute("SELECT empresa,cnpj,regime,saldo_pis,saldo_cofins,saldo_icms FROM config WHERE id=1").fetchone()

    def save_config(self, vals):
        self.c.execute("UPDATE config SET empresa=?,cnpj=?,regime=?,saldo_pis=?,saldo_cofins=?,saldo_icms=? WHERE id=1", vals)
        self.c.commit()

    def clear_docs(self):
        self.c.execute("DELETE FROM itens")
        self.c.execute("DELETE FROM docs")
        self.c.commit()


class PDFReports:
    @staticmethod
    def _header(canvas, doc, title, subtitle):
        canvas.saveState()
        w, h = landscape(A4)
        canvas.setFont("Helvetica-Bold", 15)
        canvas.drawString(16*mm, h-14*mm, "APURADOR FISCAL - LUCRO REAL")
        canvas.setFont("Helvetica", 9)
        canvas.drawString(16*mm, h-20*mm, title)
        canvas.drawRightString(w-16*mm, h-14*mm, f"Pagina {doc.page}")
        canvas.drawRightString(w-16*mm, h-20*mm, datetime.now().strftime("Gerado em %d/%m/%Y %H:%M"))
        canvas.setStrokeColor(colors.HexColor("#1F4E78"))
        canvas.line(16*mm, h-23*mm, w-16*mm, h-23*mm)
        canvas.setFont("Helvetica-Oblique", 7)
        canvas.drawString(16*mm, 8*mm, "Relatorio gerado pelo Apurador Fiscal. Software independente, sem vinculo oficial com a SEFAZ/GO.")
        canvas.restoreState()

    @staticmethod
    def _doc(path, title, subtitle):
        return SimpleDocTemplate(
            path, pagesize=landscape(A4), rightMargin=14*mm, leftMargin=14*mm,
            topMargin=28*mm, bottomMargin=14*mm,
            title=title, author="Apurador Fiscal"
        )

    @staticmethod
    def _styles():
        s = getSampleStyleSheet()
        s.add(ParagraphStyle(name="CenterSmall", parent=s["BodyText"], fontSize=7, leading=8, alignment=TA_CENTER))
        s.add(ParagraphStyle(name="Small", parent=s["BodyText"], fontSize=8, leading=10))
        s.add(ParagraphStyle(name="H", parent=s["Heading2"], fontSize=11, leading=13, spaceAfter=5, textColor=colors.HexColor("#1F4E78")))
        return s

    @staticmethod
    def monthly(path, empresa, cnpj, regime, competencia, saldos, creditos, debitos, finais, docs):
        styles = PDFReports._styles()
        story = []
        story += [Paragraph("DADOS DA EMPRESA", styles["H"]),
                  Paragraph(f"<b>Empresa:</b> {empresa or '-'} &nbsp;&nbsp; <b>CNPJ:</b> {fmt_cnpj(cnpj)} &nbsp;&nbsp; <b>Regime:</b> {regime}", styles["Small"]),
                  Paragraph(f"<b>Competencia:</b> {competencia}", styles["Small"]), Spacer(1, 5*mm)]
        data = [["TRIBUTO", "SALDO ANTERIOR", "CREDITOS DO MES", "DEBITOS DO MES", "RESULTADO FINAL"],
                ["PIS", money(saldos[0]), money(creditos[0]), money(debitos[0]), money(finais[0])],
                ["COFINS", money(saldos[1]), money(creditos[1]), money(debitos[1]), money(finais[1])],
                ["ICMS", money(saldos[2]), money(creditos[2]), money(debitos[2]), money(finais[2])]]
        t = Table(data, colWidths=[38*mm, 45*mm, 45*mm, 45*mm, 45*mm], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9EAF7")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), .4, colors.grey),
            ("ALIGN", (1,1), (-1,-1), "RIGHT"), ("FONTSIZE", (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5), ("TOPPADDING", (0,0), (-1,-1), 5)
        ]))
        story += [Paragraph("RESUMO DA APURACAO MENSAL", styles["H"]), t, Spacer(1, 7*mm),
                  Paragraph("DOCUMENTOS DA COMPETENCIA", styles["H"])]
        rows = [["DATA", "TIPO", "CHAVE NF-e", "EMITENTE / DESTINATARIO", "VALOR", "PIS", "COFINS", "ICMS", "ITENS"]]
        for r in docs:
            nome = r[8] if r[1] == "ENTRADA" else r[9]
            rows.append([r[0], r[1], Paragraph(r[2], styles["CenterSmall"]), Paragraph(nome or "-", styles["CenterSmall"]),
                         money(r[3]), money(r[4]), money(r[5]), money(r[6]), str(r[7])])
        if len(rows) == 1:
            rows.append(["-", "-", "Sem documentos no periodo", "-", "-", "-", "-", "-", "-"])
        tab = Table(rows, colWidths=[22*mm,18*mm,73*mm,48*mm,28*mm,22*mm,24*mm,24*mm,15*mm], repeatRows=1)
        tab.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9EAF7")), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), .3, colors.grey), ("FONTSIZE", (0,0), (-1,-1), 6.5),
            ("ALIGN", (4,1), (-1,-1), "RIGHT"), ("VALIGN", (0,0), (-1,-1), "MIDDLE")
        ]))
        story.append(tab)
        doc = PDFReports._doc(path, "Apuracao Fiscal Mensal", competencia)
        doc.build(story, onFirstPage=lambda c,d: PDFReports._header(c,d,"Relatorio de Apuracao Mensal",competencia),
                  onLaterPages=lambda c,d: PDFReports._header(c,d,"Relatorio de Apuracao Mensal",competencia))

    @staticmethod
    def client(path, empresa, cnpj, regime, competencia, cliente_nome, cliente_cnpj, totals, rows, saldo_antes):
        styles = PDFReports._styles()
        tv,tp,tc,ti = totals
        saldo_depois = [saldo_antes[0]-tp, saldo_antes[1]-tc, saldo_antes[2]-ti]
        story = [Paragraph("DADOS DA EMPRESA", styles["H"]),
                 Paragraph(f"<b>Empresa:</b> {empresa or '-'} &nbsp;&nbsp; <b>CNPJ:</b> {fmt_cnpj(cnpj)} &nbsp;&nbsp; <b>Regime:</b> {regime}", styles["Small"]),
                 Paragraph(f"<b>Competencia:</b> {competencia}", styles["Small"]), Spacer(1,4*mm),
                 Paragraph("DADOS DO CLIENTE", styles["H"]),
                 Paragraph(f"<b>Cliente:</b> {cliente_nome or '-'} &nbsp;&nbsp; <b>CNPJ:</b> {fmt_cnpj(cliente_cnpj)} &nbsp;&nbsp; <b>Documentos:</b> {len(rows)} &nbsp;&nbsp; <b>Total vendido:</b> {money(tv)}", styles["Small"]), Spacer(1,5*mm)]
        summary = [["TRIBUTO","SALDO DISPONIVEL ANTES DO CLIENTE","DEBITO GERADO / CREDITO CONSUMIDO","SALDO APOS CONSUMO"],
                   ["PIS",money(saldo_antes[0]),money(tp),money(saldo_depois[0])],
                   ["COFINS",money(saldo_antes[1]),money(tc),money(saldo_depois[1])],
                   ["ICMS",money(saldo_antes[2]),money(ti),money(saldo_depois[2])],
                   ["TOTAL NOMINAL","-",money(tp+tc+ti),"-"]]
        st = Table(summary, colWidths=[40*mm,60*mm,68*mm,55*mm], repeatRows=1)
        st.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9EAF7")), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), .4, colors.grey), ("ALIGN", (1,1), (-1,-1), "RIGHT"),
            ("FONTSIZE", (0,0), (-1,-1), 8)
        ]))
        story += [Paragraph("CREDITOS CONSUMIDOS / DEBITOS GERADOS", styles["H"]), st, Spacer(1,6*mm),
                  Paragraph("DOCUMENTOS DE SAIDA PARA O CLIENTE", styles["H"])]
        data = [["DATA","CHAVE NF-e","VALOR","PIS","COFINS","ICMS","ITENS"]]
        for r in rows:
            data.append([r[0], Paragraph(r[1], styles["CenterSmall"]), money(r[2]), money(r[3]), money(r[4]), money(r[5]), str(r[6])])
        if len(data)==1:
            data.append(["-","Sem documentos no periodo","-","-","-","-","-"])
        tab=Table(data,colWidths=[24*mm,95*mm,34*mm,28*mm,30*mm,30*mm,17*mm],repeatRows=1)
        tab.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9EAF7")), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), .3, colors.grey), ("FONTSIZE", (0,0), (-1,-1), 7),
            ("ALIGN", (2,1), (-1,-1), "RIGHT"), ("VALIGN", (0,0), (-1,-1), "MIDDLE")
        ]))
        story.append(tab)
        doc=PDFReports._doc(path,"Apuracao por Cliente",cliente_nome)
        doc.build(story,onFirstPage=lambda c,d: PDFReports._header(c,d,"Relatorio de Apuracao por Cliente",cliente_nome),
                  onLaterPages=lambda c,d: PDFReports._header(c,d,"Relatorio de Apuracao por Cliente",cliente_nome))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1500x900")
        self.minsize(1200,720)
        self.db=DBX()
        self._ui()
        self.load_config()
        self.refresh()

    def _ui(self):
        self.configure(bg="#f4f5f7")
        self.last_pdf = None

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Fiscal.Treeview", rowheight=25, font=("Segoe UI", 9),
                        background="white", fieldbackground="white")
        style.configure("Fiscal.Treeview.Heading", font=("Segoe UI", 9, "bold"),
                        background="#f1f1f1", foreground="#111111")

        header = tk.Frame(self, bg="white", height=112)
        header.pack(fill="x")
        header.pack_propagate(False)

        brand = tk.Frame(header, bg="white")
        brand.pack(side="left", padx=(16, 10), pady=8)
        logo_path = Path(__file__).with_name("assets") / "goias_independente.png"
        try:
            self.logo_img = tk.PhotoImage(file=str(logo_path))
            tk.Label(brand, image=self.logo_img, bg="white").pack(side="left", padx=(0, 10))
        except Exception:
            tk.Label(brand, text="GO", bg="#b50000", fg="white",
                     font=("Segoe UI", 16, "bold"), width=4, height=2).pack(side="left", padx=(0, 10))
        tk.Label(brand, text="GOIÁS\nAPURAÇÃO FISCAL", bg="white", fg="#a60000",
                 font=("Segoe UI", 11, "bold"), justify="left").pack(side="left")

        title_box = tk.Frame(header, bg="white")
        title_box.pack(side="left", expand=True, fill="both")
        tk.Label(title_box, text="APURADOR FISCAL - LUCRO REAL V4",
                 bg="white", fg="#a60000", font=("Segoe UI", 25, "bold")).pack(pady=(14, 0))
        tk.Label(title_box, text="APURAÇÃO DE PIS, COFINS E ICMS",
                 bg="white", fg="#111111", font=("Segoe UI", 15, "bold")).pack()

        right_head = tk.Frame(header, bg="white")
        right_head.pack(side="right", padx=18, pady=16)
        self.clock_label = tk.Label(right_head, text=datetime.now().strftime("%d/%m/%Y %H:%M"),
                                    bg="white", fg="#222", font=("Segoe UI", 10, "bold"))
        self.clock_label.pack(anchor="e")
        tk.Label(right_head, text="Software independente", bg="white", fg="#555",
                 font=("Segoe UI", 9)).pack(anchor="e", pady=(5, 0))

        tk.Frame(self, bg="#b40000", height=2).pack(fill="x")

        controls = tk.Frame(self, bg="white", padx=15, pady=8)
        controls.pack(fill="x")
        self.empresa = tk.StringVar()
        self.cnpj = tk.StringVar()
        self.regime = tk.StringVar(value="Lucro Real")
        self.competencia = tk.StringVar(value=datetime.now().strftime("%m/%Y"))

        def field(parent, label, var, width):
            box = tk.Frame(parent, bg="white")
            box.pack(side="left", padx=(0, 10))
            tk.Label(box, text=label, bg="white", fg="#a60000",
                     font=("Segoe UI", 9, "bold")).pack(anchor="w")
            e = tk.Entry(box, textvariable=var, width=width, relief="solid", bd=1,
                         font=("Segoe UI", 10))
            e.pack(ipady=5)
            return e

        field(controls, "EMPRESA", self.empresa, 27)
        field(controls, "CNPJ", self.cnpj, 19)

        rbox = tk.Frame(controls, bg="white")
        rbox.pack(side="left", padx=(0, 10))
        tk.Label(rbox, text="REGIME TRIBUTÁRIO", bg="white", fg="#a60000",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Combobox(rbox, textvariable=self.regime, values=["Lucro Real"],
                     state="readonly", width=16).pack(ipady=4)

        field(controls, "COMPETÊNCIA", self.competencia, 9)

        def red_btn(parent, text, cmd, width=16):
            return tk.Button(parent, text=text, command=cmd, bg="#bd0000", fg="white",
                             activebackground="#940000", activeforeground="white",
                             relief="flat", bd=0, font=("Segoe UI", 9, "bold"),
                             cursor="hand2", width=width, height=2)

        red_btn(controls, "SALVAR CADASTRO\n/ SALDOS", self.save_config, 17).pack(side="left", padx=5, pady=3)
        red_btn(controls, "IMPORTAR\nPASTA XML", self.import_folder, 16).pack(side="left", padx=5, pady=3)
        tk.Button(controls, text="LIMPAR XML\nIMPORTADOS", command=self.clear_docs,
                  bg="#f7f7f7", fg="#222", relief="solid", bd=1,
                  font=("Segoe UI", 9, "bold"), cursor="hand2", width=16, height=2).pack(side="left", padx=5, pady=3)

        body = tk.Frame(self, bg="#f4f5f7")
        body.pack(fill="both", expand=True)

        sidebar = tk.Frame(body, bg="white", width=215)
        sidebar.pack(side="left", fill="y", padx=(8, 4), pady=6)
        sidebar.pack_propagate(False)

        self.nav_buttons = {}
        navs = [
            ("mensal", "▥  Apuração Geral Mensal"),
            ("docs", "▤  Documentos (XML)"),
            ("cliente", "👥  Apuração por Cliente"),
            ("sim", "➤  Simulador de Venda"),
            ("itens", "☷  Itens Fiscais"),
        ]
        for key, label in navs:
            b = tk.Button(sidebar, text=label, anchor="w", command=lambda k=key: self.show_page(k),
                          bg="white", fg="#111", activebackground="#f0d9d9",
                          relief="flat", bd=0, font=("Segoe UI", 10, "bold"),
                          cursor="hand2", padx=12, pady=11)
            b.pack(fill="x", padx=6, pady=2)
            self.nav_buttons[key] = b

        tk.Frame(sidebar, bg="#e2e2e2", height=1).pack(fill="x", padx=10, pady=10)
        logo_box = tk.Frame(sidebar, bg="#a90000", padx=8, pady=12)
        logo_box.pack(fill="x", padx=12, pady=4)
        tk.Label(logo_box, text="GOIÁS", bg="#a90000", fg="white",
                 font=("Segoe UI", 24, "bold")).pack()
        tk.Label(logo_box, text="APURADOR FISCAL\nSoftware independente",
                 bg="#a90000", fg="white", font=("Segoe UI", 9, "bold")).pack()
        tk.Label(sidebar, text="VERSÃO 5.0.0\nLucro Real - PIS, COFINS e ICMS\nSem vínculo oficial com a SEFAZ/GO",
                 bg="white", fg="#333", justify="left", font=("Segoe UI", 8)).pack(side="bottom", anchor="w", padx=12, pady=14)

        self.page_host = tk.Frame(body, bg="#f4f5f7")
        self.page_host.pack(side="left", fill="both", expand=True, padx=(4, 8), pady=6)
        self.pages = {}
        for k in ("mensal", "docs", "cliente", "sim", "itens"):
            f = tk.Frame(self.page_host, bg="#f4f5f7")
            f.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.pages[k] = f

        p = self.pages["mensal"]

        upper = tk.Frame(p, bg="#f4f5f7")
        upper.pack(fill="x")

        balances = tk.LabelFrame(upper, text=" SALDOS ANTERIORES (INFORMADOS PELO USUÁRIO / CONTABILIDADE) ",
                                 bg="white", fg="#a60000", font=("Segoe UI", 9, "bold"),
                                 bd=1, relief="solid", padx=10, pady=8)
        balances.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.sp = tk.StringVar(value="0"); self.sc = tk.StringVar(value="0"); self.si = tk.StringVar(value="0")
        for lab, var, col in [("PIS", self.sp, "#063b9b"), ("COFINS", self.sc, "#1a145e"), ("ICMS", self.si, "#12651c")]:
            b = tk.Frame(balances, bg="white")
            b.pack(side="left", expand=True, fill="x", padx=6)
            tk.Label(b, text=lab, bg="white", fg=col, font=("Segoe UI", 9, "bold")).pack(anchor="w")
            tk.Entry(b, textvariable=var, justify="center", font=("Segoe UI", 13, "bold"),
                     fg=col, relief="solid", bd=1).pack(fill="x", ipady=7)

        indicators = tk.LabelFrame(upper, text=" INDICADORES DA COMPETÊNCIA ",
                                   bg="white", fg="#a60000", font=("Segoe UI", 9, "bold"),
                                   bd=1, relief="solid", padx=8, pady=8)
        indicators.pack(side="left", fill="x", expand=True, padx=(5, 0))
        self.kpi = {}
        for lab in ("XMLs CARREGADOS", "ENTRADAS", "SAÍDAS", "DOCUMENTOS"):
            b = tk.Frame(indicators, bg="white", highlightbackground="#cccccc", highlightthickness=1)
            b.pack(side="left", fill="both", expand=True, padx=4)
            tk.Label(b, text=lab, bg="white", fg="#333", font=("Segoe UI", 8, "bold")).pack(pady=(7, 0))
            v = tk.Label(b, text="0", bg="white", fg="#123e85", font=("Segoe UI", 17, "bold"))
            v.pack(pady=(2, 7))
            self.kpi[lab] = v

        def section_header(parent, text):
            h = tk.Label(parent, text=text, anchor="w", bg="#b90000", fg="white",
                         font=("Segoe UI", 10, "bold"), padx=10, pady=5)
            h.pack(fill="x", pady=(8, 0))
            return h

        section_header(p, "APURAÇÃO MENSAL - SALDO ANTERIOR + CRÉDITOS DAS COMPRAS - DÉBITOS DAS VENDAS")
        summary_frame = tk.Frame(p, bg="white", highlightbackground="#cccccc", highlightthickness=1)
        summary_frame.pack(fill="x")
        headers = ["TRIBUTO", "SALDO ANTERIOR", "CRÉDITOS DAS COMPRAS", "CRÉDITO DISPONÍVEL", "DÉBITOS DAS VENDAS", "CRÉDITO CONSUMIDO", "RESULTADO FINAL"]
        for col, h in enumerate(headers):
            fg = "#a40000" if h == "DÉBITOS DAS VENDAS" else ("#12651c" if h in ("CRÉDITOS DAS COMPRAS","CRÉDITO DISPONÍVEL") else "#111")
            tk.Label(summary_frame, text=h, bg="white", fg=fg, font=("Segoe UI", 8, "bold"), relief="solid", bd=1, pady=6).grid(row=0,column=col,sticky="nsew")
            summary_frame.grid_columnconfigure(col, weight=1)
        self.apur_labels = {}
        for row, tax in enumerate(("PIS","COFINS","ICMS"), start=1):
            tk.Label(summary_frame,text=tax,bg="white",font=("Segoe UI",9,"bold"),relief="solid",bd=1,pady=6).grid(row=row,column=0,sticky="nsew")
            for col,key,color in [
                (1,"anterior","#172b8c"),(2,"compras","#12651c"),(3,"disponivel","#12651c"),
                (4,"debitos","#d00000"),(5,"consumido","#172b8c"),(6,"resultado","#12651c")]:
                lab=tk.Label(summary_frame,text="R$ 0,00",bg="white",fg=color,font=("Segoe UI",9,"bold"),relief="solid",bd=1,pady=6)
                lab.grid(row=row,column=col,sticky="nsew")
                self.apur_labels[(tax,key)] = lab
        self.resultado_status = tk.Label(p, text="", bg="#f4f5f7", fg="#333", font=("Segoe UI",9,"bold"), anchor="w")
        self.resultado_status.pack(fill="x", padx=4, pady=(4,0))
        tk.Label(p, text="Crédito consumido = parcela do crédito disponível usada para compensar os débitos. Resultado positivo = crédito a transportar; resultado negativo = imposto a recolher.",
                 bg="#f4f5f7", fg="#333", font=("Segoe UI",8)).pack(anchor="w", padx=4, pady=(2,0))

        section_header(p, "DETALHAMENTO DOS DOCUMENTOS DA COMPETÊNCIA")
        cols = ("data","tipo","chave","parte","valor","pis","cofins","icms","itens")
        self.month_tree = ttk.Treeview(p, columns=cols, show="headings", style="Fiscal.Treeview", height=10)
        for c,w in zip(cols,(82,65,320,190,95,78,82,82,52)):
            self.month_tree.heading(c, text={"chave":"CHAVE DA NF-e","parte":"EMITENTE / DESTINATÁRIO"}.get(c,c.upper()))
            self.month_tree.column(c, width=w, anchor="center")
        self.month_tree.pack(fill="both", expand=True)

        actions = tk.Frame(p, bg="#f4f5f7", pady=8)
        actions.pack(fill="x")
        tk.Button(actions, text="▣  EXPORTAR PARA EXCEL", command=self.export_month_excel,
                  bg="white", fg="#08772b", relief="solid", bd=1,
                  font=("Segoe UI", 9, "bold"), cursor="hand2", padx=16, pady=9).pack(side="left", padx=(0, 8))
        red_btn(actions, "▱  GERAR RELATÓRIO PDF", self.pdf_monthly, 20).pack(side="left", padx=4)
        tk.Button(actions, text="◉  VISUALIZAR PDF", command=self.view_last_pdf,
                  bg="white", fg="#444", relief="solid", bd=1,
                  font=("Segoe UI", 9, "bold"), cursor="hand2", padx=16, pady=9).pack(side="left", padx=8)
        tk.Button(actions, text="↻  ATUALIZAR COMPETÊNCIA", command=self.refresh,
                  bg="white", fg="#555", relief="solid", bd=1,
                  font=("Segoe UI", 9, "bold"), cursor="hand2", padx=16, pady=9).pack(side="right")
        self.summary = tk.Label(p, text="", bg="#f4f5f7", fg="#444", font=("Segoe UI", 8), justify="left")
        self.summary.pack(anchor="w")

        pd = self.pages["docs"]
        section_header(pd, "DOCUMENTOS XML DA COMPETÊNCIA")
        self.tree = ttk.Treeview(pd, columns=cols, show="headings", style="Fiscal.Treeview")
        for c,w in zip(cols,(85,65,315,210,100,85,85,85,55)):
            self.tree.heading(c,text={"chave":"CHAVE DA NF-e","parte":"EMITENTE / DESTINATÁRIO"}.get(c,c.upper()))
            self.tree.column(c,width=w,anchor="center")
        self.tree.pack(fill="both",expand=True)

        pi = self.pages["itens"]
        section_header(pi, "ITENS FISCAIS - NCM / CFOP / CST")
        icols=("data","tipo","descricao","ncm","cfop","cst_pis","cst_cofins","cst_icms","valor","pis","cofins","icms")
        self.itree=ttk.Treeview(pi,columns=icols,show="headings",style="Fiscal.Treeview")
        for c in icols:
            self.itree.heading(c,text=c.upper()); self.itree.column(c,width=90 if c!="descricao" else 250,anchor="center")
        self.itree.pack(fill="both",expand=True)

        pc = self.pages["cliente"]
        section_header(pc, "APURAÇÃO FISCAL POR CLIENTE")
        cf = tk.Frame(pc, bg="white", padx=12, pady=10)
        cf.pack(fill="x")
        self.cliente_cnpj=tk.StringVar(); self.cliente_nome=tk.StringVar()
        tk.Label(cf,text="CNPJ DO CLIENTE",bg="white",fg="#a60000",font=("Segoe UI",9,"bold")).grid(row=0,column=0,sticky="w")
        tk.Entry(cf,textvariable=self.cliente_cnpj,width=20,font=("Segoe UI",10)).grid(row=1,column=0,padx=(0,8),ipady=4)
        tk.Label(cf,text="NOME / IDENTIFICAÇÃO",bg="white",fg="#a60000",font=("Segoe UI",9,"bold")).grid(row=0,column=1,sticky="w")
        tk.Entry(cf,textvariable=self.cliente_nome,width=36,font=("Segoe UI",10)).grid(row=1,column=1,padx=(0,8),ipady=4)
        red_btn(cf,"CALCULAR CLIENTE",self.calc_cliente,16).grid(row=1,column=2,padx=4)
        red_btn(cf,"IMPORTAR XML\nDESTE CLIENTE",self.import_cliente_folder,16).grid(row=1,column=3,padx=4)
        red_btn(cf,"GERAR PDF\nDO CLIENTE",self.pdf_client,15).grid(row=1,column=4,padx=4)
        self.cliente_resumo=tk.Label(pc,text="Informe o CNPJ do cliente e clique em Calcular cliente.",
                                     bg="#f4f5f7",fg="#222",font=("Segoe UI",10),justify="left")
        self.cliente_resumo.pack(anchor="w",padx=5,pady=8)
        ccols=("data","chave","valor","pis","cofins","icms","itens")
        self.ctree=ttk.Treeview(pc,columns=ccols,show="headings",style="Fiscal.Treeview")
        for c,w in zip(ccols,(85,380,110,90,90,90,55)):
            self.ctree.heading(c,text={"chave":"CHAVE DA NF-e"}.get(c,c.upper())); self.ctree.column(c,width=w,anchor="center")
        self.ctree.pack(fill="both",expand=True)

        ps = self.pages["sim"]
        section_header(ps, "SIMULADOR DE VENDA")
        sf=tk.Frame(ps,bg="white",padx=18,pady=18); sf.pack(fill="x")
        self.rp=tk.StringVar(value="1,65"); self.rc=tk.StringVar(value="7,60"); self.ri=tk.StringVar(value="19,00")
        for i,(lab,var) in enumerate([("ALÍQUOTA PIS %",self.rp),("ALÍQUOTA COFINS %",self.rc),("ALÍQUOTA ICMS %",self.ri)]):
            tk.Label(sf,text=lab,bg="white",fg="#a60000",font=("Segoe UI",9,"bold")).grid(row=0,column=i*2,sticky="e",padx=4)
            tk.Entry(sf,textvariable=var,width=10,font=("Segoe UI",11)).grid(row=0,column=i*2+1,padx=5,ipady=4)
        red_btn(sf,"CALCULAR LIMITE",self.simulate,16).grid(row=1,column=0,pady=16,sticky="w")
        self.simout=tk.Label(sf,text="",bg="white",fg="#222",font=("Segoe UI",13),justify="left")
        self.simout.grid(row=2,column=0,columnspan=6,sticky="w")

        status = tk.Frame(self, bg="#b00000", height=28)
        status.pack(fill="x")
        self.status_label = tk.Label(status, text="Banco de dados: apurador_fiscal.db", bg="#b00000",
                                     fg="white", font=("Segoe UI", 8), anchor="w", padx=12)
        self.status_label.pack(fill="both", expand=True)

        self.show_page("mensal")

    def show_page(self, key):
        self.pages[key].tkraise()
        for k,b in self.nav_buttons.items():
            if k == key:
                b.configure(bg="#bd0000", fg="white", activebackground="#970000", activeforeground="white")
            else:
                b.configure(bg="white", fg="#111", activebackground="#f0d9d9", activeforeground="#111")

    def view_last_pdf(self):
        if not self.last_pdf or not Path(self.last_pdf).exists():
            messagebox.showinfo("PDF", "Gere um relatório PDF primeiro.")
            return
        try:
            if os.name == "nt":
                os.startfile(self.last_pdf)
            elif sys.platform == "darwin":
                os.system(f'open "{self.last_pdf}"')
            else:
                os.system(f'xdg-open "{self.last_pdf}" >/dev/null 2>&1 &')
        except Exception as e:
            messagebox.showerror("PDF", f"Não foi possível abrir o PDF:\n{e}")

    def export_month_excel(self):
        try:
            saldos, cred, deb, finais = self.calc_month()
            docs = self.month_docs()
        except ValueError as e:
            messagebox.showwarning("Competência", str(e))
            return
        fn = filedialog.asksaveasfilename(
            title="Salvar planilha da apuração",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"Apuracao_Mensal_{self.competencia.get().replace('/','-')}.xlsx"
        )
        if not fn:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Apuração Mensal"
        ws["A1"] = "APURADOR FISCAL - LUCRO REAL"
        ws["A1"].font = Font(bold=True, size=16, color="A60000")
        ws.append(["Empresa", self.empresa.get(), "CNPJ", fmt_cnpj(self.cnpj.get()), "Regime", self.regime.get(), "Competência", self.competencia.get()])
        ws.append([])
        ws.append(["TRIBUTO","SALDO ANTERIOR","CRÉDITOS DO MÊS","DÉBITOS DO MÊS","RESULTADO FINAL"])
        for cell in ws[4]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="B50000")
        for tax, a, c, d, f in zip(("PIS","COFINS","ICMS"), saldos, cred, deb, finais):
            ws.append([tax, float(a), float(c), float(d), float(f)])
        ws.append([])
        ws.append(["DATA","TIPO","CHAVE NF-e","EMITENTE / DESTINATÁRIO","VALOR","PIS","COFINS","ICMS","ITENS"])
        for cell in ws[9]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="B50000")
        for r in docs:
            parte = r[8] if r[1] == "ENTRADA" else r[9]
            ws.append([r[0],r[1],r[2],parte,float(r[3]),float(r[4]),float(r[5]),float(r[6]),r[7]])
        widths = [13,12,50,35,16,14,14,14,10]
        for i,w in enumerate(widths,1):
            ws.column_dimensions[chr(64+i)].width = w
        wb.save(fn)
        messagebox.showinfo("Excel", "Planilha gerada com sucesso.")

    def comp_prefix(self):
        return parse_competencia(self.competencia.get())

    def load_config(self):
        e,c,r,p,co,i=self.db.config(); self.empresa.set(e or ""); self.cnpj.set(c or ""); self.regime.set(r or "Lucro Real")
        self.sp.set(str(p or 0)); self.sc.set(str(co or 0)); self.si.set(str(i or 0))

    def save_config(self):
        cnpj=digits(self.cnpj.get())
        if cnpj and len(cnpj)!=14:
            messagebox.showwarning("CNPJ","Informe 14 digitos no CNPJ."); return
        self.db.save_config((self.empresa.get().strip(),cnpj,self.regime.get(),float(D(self.sp.get())),float(D(self.sc.get())),float(D(self.si.get()))))
        self.refresh(); messagebox.showinfo("Salvo","Cadastro e saldos anteriores salvos.")

    def clear_docs(self):
        if messagebox.askyesno("Confirmar","Apagar os XML importados e manter o cadastro/saldos?"):
            self.db.clear_docs(); self.refresh()

    def import_folder(self):
        folder=filedialog.askdirectory(title="Selecione a pasta com XML")
        if not folder: return
        company=digits(self.cnpj.get())
        if len(company)!=14:
            messagebox.showwarning("CNPJ","Cadastre e salve o CNPJ da empresa antes de importar."); return
        ok=skip=err=0
        for f in Path(folder).rglob("*.xml"):
            try:
                if self.parse_xml(f,company): ok+=1
                else: skip+=1
            except Exception:
                err+=1
        self.db.c.commit(); self.refresh(); messagebox.showinfo("Importacao",f"Importados: {ok}\nIgnorados/duplicados: {skip}\nCom erro: {err}")

    def parse_xml(self,f,company):
        root=ET.parse(f).getroot(); inf=first_desc(root,"infNFe")
        if inf is None: return False
        ide=first_desc(inf,"ide"); emit=first_desc(inf,"emit"); dest=first_desc(inf,"dest")
        emit_cnpj=digits(child_text(emit,"CNPJ")); dest_cnpj=digits(child_text(dest,"CNPJ"))
        if company not in (emit_cnpj,dest_cnpj): return False
        tipo="SAIDA" if emit_cnpj==company else "ENTRADA"
        chave=(inf.attrib.get("Id","").replace("NFe","") or child_text(root,"chNFe") or f.stem)
        if self.db.c.execute("SELECT 1 FROM docs WHERE chave=?",(chave,)).fetchone(): return False
        data=child_text(ide,"dhEmi")[:10] or child_text(ide,"dEmi")[:10]
        total=first_desc(inf,"ICMSTot")
        try: valor=Decimal(child_text(total,"vNF") or "0")
        except: valor=Decimal("0")
        emit_nome=child_text(emit,"xNome"); dest_nome=child_text(dest,"xNome")
        docpis=docco=docic=Decimal("0"); count=0
        for det in [x for x in inf.iter() if local(x.tag)=="det"]:
            prod=direct_child(det,"prod"); imposto=direct_child(det,"imposto")
            if prod is None: continue
            try: vprod=Decimal(child_text(prod,"vProd") or "0")
            except: vprod=Decimal("0")
            pisnode=direct_child(imposto,"PIS") if imposto is not None else None
            conode=direct_child(imposto,"COFINS") if imposto is not None else None
            icnode=direct_child(imposto,"ICMS") if imposto is not None else None
            vp=sum_named(pisnode,{"vPIS"}); vc=sum_named(conode,{"vCOFINS"}); vi=sum_named(icnode,{"vICMS"})
            self.db.c.execute("""INSERT INTO itens(chave,data,tipo,codigo,descricao,ncm,cfop,cst_pis,cst_cofins,cst_icms,valor,pis,cofins,icms)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (chave,data,tipo,child_text(prod,"cProd"),child_text(prod,"xProd"),child_text(prod,"NCM"),child_text(prod,"CFOP"),
                 child_text(pisnode,"CST"),child_text(conode,"CST"),child_text(icnode,"CST"),float(vprod),float(vp),float(vc),float(vi)))
            docpis+=vp; docco+=vc; docic+=vi; count+=1
        self.db.c.execute("""INSERT INTO docs(chave,arquivo,data,tipo,emit_cnpj,dest_cnpj,emit_nome,dest_nome,valor,pis,cofins,icms,itens)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (chave,str(f),data,tipo,emit_cnpj,dest_cnpj,emit_nome,dest_nome,float(valor),float(docpis),float(docco),float(docic),count))
        return True

    def import_cliente_folder(self):
        cliente=digits(self.cliente_cnpj.get()); empresa=digits(self.cnpj.get())
        if len(empresa)!=14 or len(cliente)!=14:
            messagebox.showwarning("CNPJ","Informe o CNPJ da empresa e do cliente com 14 digitos."); return
        folder=filedialog.askdirectory(title="Selecione a pasta com XML do cliente")
        if not folder: return
        ok=skip=err=0
        for f in Path(folder).rglob("*.xml"):
            try:
                root=ET.parse(f).getroot(); inf=first_desc(root,"infNFe")
                if inf is None: skip+=1; continue
                emit=first_desc(inf,"emit"); dest=first_desc(inf,"dest")
                if digits(child_text(emit,"CNPJ"))==empresa and digits(child_text(dest,"CNPJ"))==cliente:
                    if self.parse_xml(f,empresa): ok+=1
                    else: skip+=1
                else: skip+=1
            except Exception: err+=1
        self.db.c.commit(); self.refresh(); self.calc_cliente(); messagebox.showinfo("Importacao do cliente",f"Vendas ao cliente importadas: {ok}\nIgnorados/duplicados: {skip}\nCom erro: {err}")

    def calc_month(self):
        prefix=self.comp_prefix(); p0,c0,i0=map(D,(self.sp.get(),self.sc.get(),self.si.get()))
        rows=self.db.c.execute("SELECT tipo,pis,cofins,icms FROM docs WHERE substr(data,1,7)=?",(prefix,)).fetchall()
        cred=[Decimal("0")]*3; deb=[Decimal("0")]*3
        for typ,p,c,i in rows:
            vals=[Decimal(str(p or 0)),Decimal(str(c or 0)),Decimal(str(i or 0))]
            target=cred if typ=="ENTRADA" else deb
            for x in range(3): target[x]+=vals[x]
        final=[p0+cred[0]-deb[0],c0+cred[1]-deb[1],i0+cred[2]-deb[2]]
        return [p0,c0,i0],cred,deb,final

    def month_docs(self):
        return self.db.c.execute("""SELECT data,tipo,chave,valor,pis,cofins,icms,itens,emit_nome,dest_nome
            FROM docs WHERE substr(data,1,7)=? ORDER BY data,chave""",(self.comp_prefix(),)).fetchall()

    def calc_cliente_rows(self):
        cliente=digits(self.cliente_cnpj.get()); empresa=digits(self.cnpj.get())
        if len(cliente)!=14: raise ValueError("Informe os 14 digitos do CNPJ do cliente.")
        rows=self.db.c.execute("""SELECT data,chave,valor,pis,cofins,icms,itens,dest_nome FROM docs
            WHERE tipo='SAIDA' AND emit_cnpj=? AND dest_cnpj=? AND substr(data,1,7)=? ORDER BY data,chave""",
            (empresa,cliente,self.comp_prefix())).fetchall()
        tv=tp=tc=ti=Decimal("0")
        for r in rows:
            tv+=Decimal(str(r[2] or 0)); tp+=Decimal(str(r[3] or 0)); tc+=Decimal(str(r[4] or 0)); ti+=Decimal(str(r[5] or 0))
        return rows,(tv,tp,tc,ti)

    def calc_cliente(self):
        try: rows,totals=self.calc_cliente_rows()
        except ValueError as e: messagebox.showwarning("Cliente",str(e)); return
        for x in self.ctree.get_children(): self.ctree.delete(x)
        for r in rows: self.ctree.insert("","end",values=(r[0],r[1],money(r[2]),money(r[3]),money(r[4]),money(r[5]),r[6]))
        if rows and not self.cliente_nome.get().strip(): self.cliente_nome.set(rows[0][7] or "")
        _,_,_,saldo_atual=self.calc_month(); tv,tp,tc,ti=totals
        antes=[saldo_atual[0]+tp,saldo_atual[1]+tc,saldo_atual[2]+ti]
        depois=saldo_atual
        self.cliente_resumo.config(text=(f"Cliente: {self.cliente_nome.get().strip() or '-'} | CNPJ: {fmt_cnpj(self.cliente_cnpj.get())} | Competencia: {self.competencia.get()}\n"
            f"Vendas encontradas: {len(rows)} documento(s) | Total vendido: {money(tv)}\n\n"
            f"DEBITOS GERADOS / CREDITOS CONSUMIDOS PELAS VENDAS AO CLIENTE\nPIS: {money(tp)} | COFINS: {money(tc)} | ICMS: {money(ti)}\n"
            f"Saldo antes deste cliente: PIS {money(antes[0])} | COFINS {money(antes[1])} | ICMS {money(antes[2])}\n"
            f"Saldo depois deste cliente: PIS {money(depois[0])} | COFINS {money(depois[1])} | ICMS {money(depois[2])}"))

    def refresh(self):
        try:
            saldos,cred,deb,av=self.calc_month()
            docs=self.month_docs()
        except ValueError as e:
            messagebox.showwarning("Competência",str(e)); return

        statuses=[]
        for idx,tax in enumerate(("PIS","COFINS","ICMS")):
            anterior=saldos[idx]; compras=cred[idx]; disponivel=anterior+compras; debito=deb[idx]
            consumido=min(max(Decimal("0"),disponivel), max(Decimal("0"),debito))
            resultado=disponivel-debito
            vals={"anterior":anterior,"compras":compras,"disponivel":disponivel,"debitos":debito,"consumido":consumido,"resultado":resultado}
            for key,val in vals.items():
                lab=self.apur_labels[(tax,key)]
                lab.config(text=money(val))
                if key=="resultado":
                    lab.config(fg="#12651c" if val>=0 else "#d00000")
            if resultado < 0:
                statuses.append(f"{tax} A RECOLHER: {money(abs(resultado))}")
            else:
                statuses.append(f"{tax} - CRÉDITO A TRANSPORTAR: {money(resultado)}")
        self.resultado_status.config(text="   |   ".join(statuses), fg="#333")

        entradas=sum(1 for r in docs if r[1]=="ENTRADA"); saidas=sum(1 for r in docs if r[1]=="SAIDA")
        self.kpi["XMLs CARREGADOS"].config(text=str(len(docs)))
        self.kpi["ENTRADAS"].config(text=str(entradas), fg="#8a3c00")
        self.kpi["SAÍDAS"].config(text=str(saidas), fg="#b00000")
        self.kpi["DOCUMENTOS"].config(text=str(len(docs)), fg="#5b2b84")
        self.summary.config(text=(f"Competência {self.competencia.get()} | Documentos: {len(docs)} | Entradas: {entradas} | Saídas: {saidas}\n"
            f"Créditos identificados nos XML de entrada: PIS {money(cred[0])} | COFINS {money(cred[1])} | ICMS {money(cred[2])}. "
            "Pré-apuração: validar o direito efetivo ao crédito conforme CST, CFOP, NCM e legislação aplicável."))
        self.status_label.config(text=f"Banco de dados: apurador_fiscal.db     |     XMLs carregados: {len(docs)}     |     Entradas: {entradas}     |     Saídas: {saidas}     |     Competência: {self.competencia.get()}")
        for t in (self.tree,self.month_tree,self.itree):
            for x in t.get_children(): t.delete(x)
        for r in docs:
            parte=r[8] if r[1]=="ENTRADA" else r[9]
            vals=(r[0],r[1],r[2],parte,money(r[3]),money(r[4]),money(r[5]),money(r[6]),r[7])
            self.tree.insert("","end",values=vals); self.month_tree.insert("","end",values=vals)
        for r in self.db.c.execute("""SELECT data,tipo,descricao,ncm,cfop,cst_pis,cst_cofins,cst_icms,valor,pis,cofins,icms
            FROM itens WHERE substr(data,1,7)=? ORDER BY data,id""",(self.comp_prefix(),)):
            vals=list(r); vals[8:]=[money(x) for x in vals[8:]]; self.itree.insert("","end",values=vals)
        if len(digits(self.cliente_cnpj.get()))==14: self.calc_cliente()

    def pdf_monthly(self):
        try: saldos,cred,deb,av=self.calc_month(); docs=self.month_docs()
        except ValueError as e: messagebox.showwarning("Competencia",str(e)); return
        fn=filedialog.asksaveasfilename(title="Salvar relatorio mensal",defaultextension=".pdf",filetypes=[("PDF","*.pdf")],initialfile=f"Apuracao_Mensal_{self.competencia.get().replace('/','-')}.pdf")
        if not fn: return
        try:
            PDFReports.monthly(fn,self.empresa.get(),digits(self.cnpj.get()),self.regime.get(),self.competencia.get(),saldos,cred,deb,av,docs)
            self.last_pdf = fn
            messagebox.showinfo("PDF","Relatório mensal gerado com sucesso.")
        except Exception as e: messagebox.showerror("PDF",f"Nao foi possivel gerar o PDF:\n{e}")

    def pdf_client(self):
        try: rows,totals=self.calc_cliente_rows(); _,_,_,saldo_atual=self.calc_month()
        except ValueError as e: messagebox.showwarning("Cliente",str(e)); return
        tv,tp,tc,ti=totals
        saldo_antes=[saldo_atual[0]+tp,saldo_atual[1]+tc,saldo_atual[2]+ti]
        nome=self.cliente_nome.get().strip() or (rows[0][7] if rows else "Cliente")
        fn=filedialog.asksaveasfilename(title="Salvar relatorio do cliente",defaultextension=".pdf",filetypes=[("PDF","*.pdf")],initialfile=f"Cliente_{digits(self.cliente_cnpj.get())}_{self.competencia.get().replace('/','-')}.pdf")
        if not fn: return
        try:
            PDFReports.client(fn,self.empresa.get(),digits(self.cnpj.get()),self.regime.get(),self.competencia.get(),nome,digits(self.cliente_cnpj.get()),totals,rows,saldo_antes)
            self.last_pdf = fn
            messagebox.showinfo("PDF","Relatório do cliente gerado com sucesso.")
        except Exception as e: messagebox.showerror("PDF",f"Nao foi possivel gerar o PDF:\n{e}")

    def simulate(self):
        try: _,_,_,av=self.calc_month()
        except ValueError as e: messagebox.showwarning("Competencia",str(e)); return
        def rate(v):
            s=(v or "0").replace(".","").replace(",",".")
            try: return Decimal(s)/100
            except: return Decimal("0")
        rates=[rate(self.rp.get()),rate(self.rc.get()),rate(self.ri.get())]; names=["PIS","COFINS","ICMS"]; limits=[]; lines=[]
        for n,s,r in zip(names,av,rates):
            lim=(max(s,Decimal("0"))/r) if r>0 else Decimal("Infinity"); limits.append(lim)
            lines.append(f"{n}: {money(lim) if lim.is_finite() else 'sem limite por aliquota 0%'}")
        finite=[x for x in limits if x.is_finite()]
        if finite:
            geral=min(finite); limitador=names[limits.index(geral)]
            lines += ["",f"LIMITE ESTIMADO DO MIX INFORMADO: {money(geral)}",f"Tributo limitador: {limitador}"]
        self.simout.config(text="\n".join(lines))


if __name__ == "__main__":
    App().mainloop()
