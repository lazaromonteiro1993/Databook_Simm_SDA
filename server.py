from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
import requests
import os
from io import BytesIO

# ================= APP =================

app = Flask(__name__)

# ================= SHAREPOINT =================

USER = os.getenv("SHAREPOINT_USER")
PASS = os.getenv("SHAREPOINT_PASS")

# ================= LINKS =================

ARQUIVOS = {

    "se": os.getenv("LINK_SE"),
    "rmt": os.getenv("LINK_RMT"),

    "sda1": os.getenv("LINK_SDA1"),
    "sda2": os.getenv("LINK_SDA2"),
    "sda3": os.getenv("LINK_SDA3"),
    "sda4": os.getenv("LINK_SDA4"),
    "sda5": os.getenv("LINK_SDA5")

}

# ================= FRONTEND =================

@app.route("/")
def home():
    return send_from_directory(".", "index.html")


@app.route("/script.js")
def js():
    return send_from_directory(".", "script.js")


@app.route("/style.css")
def css():
    return send_from_directory(".", "style.css")


# ================= DOWNLOAD EXCEL =================

def baixar_excel(url):

    if not url:
        return None

    try:

        sessao = requests.Session()
        sessao.auth = (USER, PASS)

        r = sessao.get(url)

        if r.status_code != 200:
            print("Erro download:", url)
            return None

        return BytesIO(r.content)

    except Exception as e:

        print("Erro download:", e)
        return None


# ================= LEITURA =================

def carregar(url):

    arquivo = baixar_excel(url)

    if arquivo is None:
        return None

    try:

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

    except Exception as e:

        print("Erro leitura:", e)
        return None


# ================= API =================

@app.route("/dados")
def dados():

    sda = request.args.get("sda", "geral")

    dados = {k: carregar(v) for k, v in ARQUIVOS.items()}
    dados = {k: v for k, v in dados.items() if v is not None}

    if not dados:
        return jsonify({
            "total": 0,
            "postados": 0,
            "progresso": 0,
            "sdas": {},
            "tabela": []
        })

    df_geral = pd.concat(dados.values()).reset_index(drop=True)

    # ================= PROGRESSO POR SDA =================

    sdas = {}

    for nome, df in dados.items():

        total = df["Quantidade total"].sum()
        postados = df["Postagem"].sum()

        porcentagem = round(
            (postados / total) * 100, 1
        ) if total > 0 else 0

        sdas[nome] = porcentagem


    # ================= BASE ATIVA =================

    if sda == "geral":
        df_base = df_geral
    else:
        df_base = dados.get(sda, df_geral)


    # ================= KPI =================

    total = int(df_base["Quantidade total"].sum())
    postados = int(df_base["Postagem"].sum())

    progresso = round(
        (postados / total) * 100, 1
    ) if total > 0 else 0


    # ================= STATUS =================

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

    app.run(
        host="0.0.0.0",
        port=port
    )
