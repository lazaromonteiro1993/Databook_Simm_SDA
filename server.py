from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
from pathlib import Path
import os

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent

# ================= LOGIN SHAREPOINT =================
SP_USER = os.getenv("SP_USER")
SP_PASS = os.getenv("SP_PASS")
SITE = "https://simmsa-my.sharepoint.com"

ctx = ClientContext(SITE).with_credentials(
    UserCredential(SP_USER, SP_PASS)
)

# ================= ARQUIVOS SHAREPOINT =================
ARQUIVOS = {

"rmt": "/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK RMT/SDA-SIM-E-MTRD-Q00-0155 - 00 - Índice Data Book RMT-11.09.25.xlsx",

"se": "/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK SE/SDA-SIM-E-SERD-Q00-0001-00 - Índice Data Book SE-11.09.25.xlsx",

"sda1": "/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 1/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 1.xlsx",

"sda2": "/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 2/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 2.xlsx",

"sda3": "/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 3/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 3.xlsx",

"sda4": "/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 4/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 4.xlsx",

"sda5": "/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 5/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 5.xlsx",

"sda6": "/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 6/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 6.xlsx"

}

# ================= BAIXAR EXCEL =================

def baixar_excel(url, destino):
    with open(destino, "wb") as f:
        ctx.web.get_file_by_server_relative_url(url).download(f).execute_query()

# ================= LER EXCEL =================

def carregar(url, chave):

    caminho = BASE_DIR / f"{chave}.xlsx"

    baixar_excel(url, caminho)

    df = pd.read_excel(
        caminho,
        usecols="B:J",
        skiprows=6
    )

    df.columns = df.columns.str.strip()

    df = df[df["Item"].astype(str).str.match(r"^\d+\..*")]

    if "Setor" in df.columns:
        df["Setor"] = df["Setor"].fillna("")

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

# ================= ROTAS =================

@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "index.html")

@app.route("/<path:arquivo>")
def arquivos(arquivo):
    return send_from_directory(BASE_DIR, arquivo)

# ================= API =================

@app.route("/dados")
def dados():

    sda = request.args.get("sda", "geral")

    dados = {}

    for chave, url in ARQUIVOS.items():
        try:
            dados[chave] = carregar(url, chave)
        except:
            pass

    if not dados:
        return jsonify({"erro": "Nenhum arquivo encontrado"})

    df_geral = pd.concat(dados.values()).reset_index(drop=True)

    # progresso por SDA
    sdas = {}

    for nome, df in dados.items():

        total = df["Quantidade total"].sum()
        postados = df["Postagem"].sum()

        porcentagem = round(
            (postados / total) * 100, 1
        ) if total > 0 else 0

        sdas[nome] = porcentagem

    # base ativa
    if sda == "geral":
        df_base = df_geral
    else:
        df_base = dados.get(sda, df_geral)

    total = int(df_base["Quantidade total"].sum())
    postados = int(df_base["Postagem"].sum())

    progresso = round(
        (postados / total) * 100, 1
    ) if total > 0 else 0

    # status
    df_status = df_base.copy()

    df_status["Status"] = "Finalizado"

    df_status.loc[df_status["Postagem"] == 0, "Status"] = "Pendente"

    df_status.loc[
        (df_status["Postagem"] > 0) &
        (df_status["Postagem"] < df_status["Quantidade total"]),
        "Status"
    ] = "Parcial"

    df_status = df_status[df_status["Status"] != "Finalizado"]

    tabela = []

    for _, r in df_status.iterrows():

        tabela.append({

            "item": str(r["Item"]),
            "setor": str(r.get("Setor", "")),
            "documento": str(r.get("Documento", "")),
            "total": int(r["Quantidade total"]),
            "postados": int(r["Postagem"]),
            "comentario": "",
            "status": str(r["Status"])

        })

    return jsonify({

        "total": total,
        "postados": postados,
        "progresso": progresso,
        "sdas": sdas,
        "tabela": tabela

    })

# ================= START =================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)