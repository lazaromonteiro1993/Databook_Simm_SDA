from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
import requests
import os
from io import BytesIO
from pathlib import Path

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent


# ================= FRONTEND =================

@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/<path:arquivo>")
def arquivos(arquivo):
    return send_from_directory(BASE_DIR, arquivo)


# ================= LINKS SHAREPOINT =================

ARQUIVOS = {

    "se": "https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20SE/SDA-SIM-E-SERD-Q00-0001-00%20-%20%C3%8Dndice%20Data%20Book%20SE-11.09.25.xlsx",

    "rmt": "https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20RMT/SDA-SIM-E-MTRD-Q00-0155%20-%2000%20-%20%C3%8Dndice%20Data%20Book%20RMT-11.09.25.xlsx",

    "sda1": "https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%201/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%201.xlsx",

    "sda2": "https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%202/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%202.xlsx",

    "sda3": "https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%203/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%203.xlsx",

    "sda4": "https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%204/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%204.xlsx",

    "sda5": "https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%205/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%205.xlsx",

    "sda6": "https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%206/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%206.xlsx"
}


# ================= DOWNLOAD EXCEL =================

def baixar_excel(url):

    user = os.getenv("SP_USER")
    password = os.getenv("SP_PASS")

    r = requests.get(url, auth=(user, password))

    if r.status_code != 200:
        return None

    return BytesIO(r.content)


# ================= LEITURA EXCEL =================

def carregar(nome, url):

    arquivo = baixar_excel(url)

    if arquivo is None:
        return None


    # SE
    if nome == "se":

        df = pd.read_excel(
            arquivo,
            sheet_name="MC (S)",
            usecols="B:J",
            skiprows=6
        )


    # RMT
    elif nome == "rmt":

        df = pd.read_excel(
            arquivo,
            sheet_name="MC (R)",
            usecols="B:J",
            skiprows=6
        )


    # SDA
    else:

        df = pd.read_excel(
            arquivo,
            usecols="B:J",
            skiprows=6
        )


    df.columns = df.columns.str.strip()

    df = df[df["Item"].astype(str).str.match(r"^\d+\..*")]

    df = df[df["Documento"].notna()]

    if "Setor" not in df.columns:
        df["Setor"] = ""


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

    dados = {}

    for nome, url in ARQUIVOS.items():

        df = carregar(nome, url)

        if df is not None:
            dados[nome] = df


    if not dados:
        return jsonify({"erro": "Falha ao acessar SharePoint"})


    df_geral = pd.concat(dados.values()).reset_index(drop=True)


    # progresso por card
    sdas = {}

    for nome, df in dados.items():

        total = df["Quantidade total"].sum()

        postados = df["Postagem"].sum()

        porcentagem = round(
            (postados / total) * 100, 1
        ) if total > 0 else 0

        sdas[nome] = porcentagem


    if sda == "geral":
        df_base = df_geral
    else:
        df_base = dados.get(sda, df_geral)


    total = int(df_base["Quantidade total"].sum())

    postados = int(df_base["Postagem"].sum())

    progresso = round(
        (postados / total) * 100, 1
    ) if total > 0 else 0


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
            "setor": str(r["Setor"]),
            "documento": str(r["Documento"]),
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
    app.run(debug=True)
