from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
from pathlib import Path
from io import BytesIO
import os
import requests
from requests.auth import HTTPBasicAuth

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent


# ================= LOGIN =================

USER = os.getenv("SHAREPOINT_USER")
PASS = os.getenv("SHAREPOINT_PASS")


# ================= LINKS =================

ARQUIVOS = {

"se": "https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQDxJXdzK5d5T7DgixafUwcyAR-WTUSEf6ukZfRvoT4urOI",

"rmt": "https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQD-KltGair7Q4q6opKhRRsUASrSLN1jq2OPfni7DhoqJ9E",

"sda1": "https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQDcTx8ePFbjSagZq263RIxlAU1UjnZVBdf_ik61AQoWTrw",

"sda2": "https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQCEzUG4i0GZRZGOgNPz0JYlAdGfjU7ifBeVBR5ARP_IRJg",

"sda3": "https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQAkLNHcirYbTarilGQEN_KUAcbWl_Q9K4186tz_Ak7soA0",

"sda4": "https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQABItzokLRLQ4-XmGIgexo3AW5Gv1mwWWSeN7_YXNfelQM",

"sda5": "https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQCIY1QTN0ClSagd2wkdDKAhAcu83tQTDvIkU0J_ooqn-OM",

"sda6": "https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQBFDj0RBVPARKYzwKljEWy9AZcbvdc-aXptISOgiIySD7Y"
}


# ================= BAIXAR EXCEL =================

def baixar_excel(url):

    try:

        r = requests.get(
            url + "?download=1",
            auth=HTTPBasicAuth(USER, PASS)
        )

        if r.status_code != 200:
            return None

        return BytesIO(r.content)

    except:
        return None


# ================= LEITURA =================

def carregar(url):

    arquivo = baixar_excel(url)

    if arquivo is None:
        return None


    df = pd.read_excel(
        arquivo,
        usecols="B:J",
        skiprows=6
    )

    df.columns = df.columns.str.strip()

    df = df[df["Item"].astype(str).str.match(r"^\d+\..*")]

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

    dados = {k: carregar(v) for k, v in ARQUIVOS.items()}
    dados = {k: v for k, v in dados.items() if v is not None}

    if not dados:
        return jsonify({"erro": "Nenhum arquivo encontrado"})


    df_geral = pd.concat(dados.values()).reset_index(drop=True)


    if sda == "geral":
        df_base = df_geral
    else:
        df_base = dados.get(sda, df_geral)


    total = int(df_base["Quantidade total"].sum())
    postados = int(df_base["Postagem"].sum())

    progresso = round(
        (postados / total) * 100, 1
    ) if total > 0 else 0


    return jsonify({

        "total": total,
        "postados": postados,
        "progresso": progresso,
        "sdas": {},
        "tabela": []

    })


# ================= START =================

if __name__ == "__main__":
    app.run()
