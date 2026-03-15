from flask import Flask, jsonify, send_from_directory, request
import pandas as pd
from pathlib import Path

# PDF
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from datetime import datetime

app = Flask(__name__)

BASE_DIR = Path(__file__).parent

# ================= PLANILHAS =================

PASTA = BASE_DIR / "dashboard"
PASTA.mkdir(exist_ok=True)

ARQUIVOS = {

"se": PASTA / "SDA-SIM-E-SERD-Q00-0001-00 - Índice Data Book SE-11.09.25.xlsx",

"rmt": PASTA / "SDA-SIM-E-MTRD-Q00-0155 - 00 - Índice Data Book RMT-11.09.25.xlsx",

"sda1": PASTA / "SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 1.xlsx",

"sda2": PASTA / "SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 2.xlsx",

"sda3": PASTA / "SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 3.xlsx",

"sda4": PASTA / "SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 4.xlsx",

"sda5": PASTA / "SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 5.xlsx",

"sda6": PASTA / "SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 6.xlsx"

}

# ================= LER PLANILHA =================

def carregar(nome, caminho):

    caminho = Path(caminho)

    if not caminho.exists():
        return None

    try:

        xls = pd.ExcelFile(caminho)

        aba = next(
            (a for a in xls.sheet_names if "MC" in a.upper()),
            None
        )

        if aba is None:
            return None


        # ================= SE / RMT =================

        if nome in ["se","rmt"]:

            df = pd.read_excel(
                caminho,
                sheet_name=aba,
                usecols="B:I",
                skiprows=6
            )

            df.columns = [
            "Item",
            "Setor",
            "Documento",
            "Peso",
            "Quantidade total",
            "Postagem",
            "Percentual",
            "Observação"
            ]


        # ================= SDA =================

        else:

            df = pd.read_excel(
                caminho,
                sheet_name=aba,
                usecols="B:J",
                skiprows=6
            )

            df.columns = [
            "Item",
            "Responsavel",
            "Setor",
            "Documento",
            "Peso",
            "Quantidade total",
            "Postagem",
            "Percentual",
            "Observação"
            ]


        # ================= LIMPEZA =================

        df["Item"] = df["Item"].astype(str)

        df = df[df["Item"].str.match(r"^\d+(\.\d+)*$", na=False)]

        df = df[~df["Documento"].astype(str).str.upper().str.contains("TOTAL", na=False)]

        df = df[df["Documento"].notna()]


        # ================= NUMÉRICO =================

        df["Quantidade total"] = pd.to_numeric(
            df["Quantidade total"], errors="coerce"
        )

        df["Postagem"] = pd.to_numeric(
            df["Postagem"], errors="coerce"
        ).fillna(0)


        df = df.dropna(subset=["Quantidade total"])

        df["Quantidade total"] = df["Quantidade total"].astype(int)
        df["Postagem"] = df["Postagem"].astype(int)


        return df.reset_index(drop=True)

    except:

        return None


# ================= FRONTEND =================

@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/style.css")
def css():
    return send_from_directory(BASE_DIR, "style.css")


@app.route("/script.js")
def js():
    return send_from_directory(BASE_DIR, "script.js")


# ================= API =================

@app.route("/dados")
def dados():

    sda = request.args.get("sda","geral")

    dados_excel = {k: carregar(k,v) for k,v in ARQUIVOS.items()}


    if sda == "geral":

        dfs = [
            df for df in dados_excel.values()
            if df is not None and not df.empty
        ]

    else:

        df = dados_excel.get(sda)

        dfs = [df] if df is not None else []


    if dfs:
        df_base = pd.concat(dfs, ignore_index=True)
    else:
        df_base = pd.DataFrame(columns=["Quantidade total","Postagem"])


    total = int(df_base["Quantidade total"].sum())
    postados = int(df_base["Postagem"].sum())

    progresso = round((postados/total)*100,1) if total>0 else 0


    sdas = {}

    for nome,df in dados_excel.items():

        if df is None or df.empty:
            sdas[nome] = 0
            continue

        t = df["Quantidade total"].sum()
        p = df["Postagem"].sum()

        sdas[nome] = round((p/t)*100,1) if t>0 else 0


    tabela = []

    if not df_base.empty:

        for _,row in df_base.iterrows():

            status = "Pendente"

            if row["Postagem"] == row["Quantidade total"]:
                status = "Finalizado"

            elif row["Postagem"] > 0:
                status = "Parcial"

            if status == "Finalizado":
                continue


            tabela.append({

            "item":str(row["Item"]),
            "setor":str(row.get("Setor","")),
            "documento":str(row.get("Documento","")),
            "total":int(row["Quantidade total"]),
            "postados":int(row["Postagem"]),
            "comentario":str(row.get("Observação","")),
            "status":status

            })


    return jsonify({
    "total":total,
    "postados":postados,
    "progresso":progresso,
    "sdas":sdas,
    "tabela":tabela
    })


# ================= PDF =================

@app.route("/baixar-pdf")
def baixar_pdf():

    caminho = BASE_DIR / "relatorio_databook.pdf"

    elementos = []

    titulo = ParagraphStyle("titulo", fontSize=26, leading=30, spaceAfter=5)

    subtitulo = ParagraphStyle(
        "sub",
        fontSize=12,
        textColor=colors.HexColor("#3c4a5d"),
        spaceAfter=20
    )

    texto = ParagraphStyle("texto", fontSize=9, leading=12)

    header = ParagraphStyle(
        "header",
        fontSize=10,
        textColor=colors.white
    )

    elementos.append(Paragraph("UFV SOL DO AGRESTE - RELATÓRIO DO DATABOOK", titulo))

    elementos.append(
        Paragraph(
            f"Controle e postagem de documentos — {datetime.today().strftime('%d/%m/%Y')}",
            subtitulo
        )
    )

    elementos.append(Spacer(1,10))


    for chave, arquivo in ARQUIVOS.items():

        df = carregar(chave, arquivo)

        if df is None:
            continue

        df["Status"] = "Finalizado"
        df.loc[df["Postagem"] == 0, "Status"] = "Pendente"

        df.loc[
            (df["Postagem"] > 0) &
            (df["Postagem"] < df["Quantidade total"]),
            "Status"
        ] = "Parcial"

        df = df[df["Status"] != "Finalizado"]

        if df.empty:
            continue

        df = df.fillna("")

        elementos.append(
            Paragraph(f"<b>{chave.upper()}</b>", texto)
        )

        dados = [[

            Paragraph("Item", header),
            Paragraph("Setor", header),
            Paragraph("Documento", header),
            Paragraph("Total", header),
            Paragraph("Postados", header),
            Paragraph("Comentários", header),
            Paragraph("Status", header)

        ]]

        for _, r in df.iterrows():

            dados.append([

                Paragraph(str(r["Item"]), texto),
                Paragraph(str(r.get("Setor","")), texto),
                Paragraph(str(r.get("Documento","")), texto),
                Paragraph(str(r["Quantidade total"]), texto),
                Paragraph(str(r["Postagem"]), texto),
                Paragraph(str(r.get("Observação","")), texto),
                Paragraph(str(r["Status"]), texto)

            ])

        largura_pagina = landscape(A4)[0]
        largura_util = largura_pagina - 140

        colWidths = [

            largura_util * 0.07,
            largura_util * 0.12,
            largura_util * 0.28,
            largura_util * 0.08,
            largura_util * 0.10,
            largura_util * 0.25,
            largura_util * 0.10

        ]

        tabela = Table(dados, colWidths=colWidths, repeatRows=1)

        tabela.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#2d466f")),
            ("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#e3e8f0"))
        ]))

        elementos.append(tabela)
        elementos.append(Spacer(1,25))


    doc = SimpleDocTemplate(
        str(caminho),
        pagesize=landscape(A4),
        leftMargin=70,
        rightMargin=70,
        topMargin=60,
        bottomMargin=50
    )

    doc.build(elementos)

    return send_from_directory(BASE_DIR,"relatorio_databook.pdf",as_attachment=True)


# ================= START =================

if __name__ == "__main__":
    app.run(debug=True)
