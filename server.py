from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
from io import BytesIO
from pathlib import Path
import os

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent


# ================= LOGIN SHAREPOINT =================

USERNAME = os.environ.get("SP_USER")
PASSWORD = os.environ.get("SP_PASS")

SITE = "https://simmsa-my.sharepoint.com"


def baixar_excel(link):

    try:

        ctx = ClientContext(SITE).with_credentials(
            UserCredential(USERNAME, PASSWORD)
        )

        response = BytesIO()

        ctx.web.get_file_by_server_relative_url(link).download(response).execute_query()

        response.seek(0)

        return response

    except Exception as e:

        print("Erro SharePoint:", e)

        return None


# ================= ROTAS =================

@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/<path:arquivo>")
def arquivos(arquivo):
    return send_from_directory(BASE_DIR, arquivo)


# ================= ARQUIVOS =================

ARQUIVOS = {

"se":"/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK SE/SDA-SIM-E-SERD-Q00-0001-00 - Índice Data Book SE-11.09.25.xlsx",

"rmt":"/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK RMT/SDA-SIM-E-MTRD-Q00-0155 - 00 - Índice Data Book RMT-11.09.25.xlsx",

"sda1":"/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 1/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 1.xlsx",

"sda2":"/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 2/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 2.xlsx",

"sda3":"/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 3/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 3.xlsx",

"sda4":"/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 4/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 4.xlsx",

"sda5":"/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 5/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 5.xlsx",

"sda6":"/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data book/DATABOOK UFV/SDA 6/SDA-SIM-E-PVRD-Q00-0148-00 - Índice Data Book UFV - SDA 6.xlsx"

}


# ================= LEITURA EXCEL =================

def carregar(nome):

    arquivo = baixar_excel(ARQUIVOS[nome])

    if arquivo is None:
        return None


    # ================= SE =================
    if "SERD" in nome:

        df = pd.read_excel(
            arquivo,
            sheet_name="MC (S)",
            usecols="B:J",
            skiprows=6
        )

        df.columns = df.columns.str.strip()

        df = df[df["Item"].astype(str).str.match(r"^\d+\..*")]

        df = df[df["Documento"].notna()]

        df["Setor"] = ""


    # ================= RMT =================
    elif "MTRD" in nome:

        df = pd.read_excel(
            arquivo,
            sheet_name="MC (R)",
            usecols="B:J",
            skiprows=6
        )

        df.columns = df.columns.str.strip()

        df = df[df["Item"].astype(str).str.match(r"^\d+\..*")]

        df = df[df["Documento"].notna()]

        df["Setor"] = ""


    # ================= SDA =================
    else:

        df = pd.read_excel(
            arquivo,
            usecols="B:J",
            skiprows=6
        )

        df.columns = df.columns.str.strip()

        df = df[df["Item"].astype(str).str.match(r"^\d+\..*")]

        if "Setor" in df.columns:
            df["Setor"] = df["Setor"].fillna("")


    # ================= TRATAMENTO NUMÉRICO =================

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


# ================= API =================

@app.route("/dados")
def dados():

    sda = request.args.get("sda", "geral")

    dados = {k: carregar(k) for k in ARQUIVOS}

    dados = {k: v for k, v in dados.items() if v is not None}

    if not dados:
        return jsonify({"erro": "Nenhum arquivo encontrado"})


    df_geral = pd.concat(dados.values()).reset_index(drop=True)


    sdas = {}

    for nome, df in dados.items():

        total = df["Quantidade total"].sum()

        postados = df["Postagem"].sum()

        porcentagem = round((postados / total) * 100, 1) if total > 0 else 0

        sdas[nome] = porcentagem


    if sda == "geral":
        df_base = df_geral
    else:
        df_base = dados.get(sda, df_geral)


    total = int(df_base["Quantidade total"].sum())

    postados = int(df_base["Postagem"].sum())

    progresso = round((postados / total) * 100, 1) if total > 0 else 0


    return jsonify({

        "total": total,
        "postados": postados,
        "progresso": progresso,
        "sdas": sdas,
        "tabela": []

    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
