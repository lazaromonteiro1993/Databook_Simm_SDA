from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
import requests
from io import BytesIO
import os

app = Flask(__name__)


# ================= LINKS SHAREPOINT =================

ARQUIVOS = {

"se":"https://simmsa-my.sharepoint.com/:x:/g/personal/valquiria_andrade_simmsolucoes_com_br/IQBFDj0RBVPARKYzwKljEWy9AZcbvdc-aXptISOgiIySD7Y",

"rmt":"https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20RMT/SDA-SIM-E-MTRD-Q00-0155%20-%2000%20-%20%C3%8Dndice%20Data%20Book%20RMT-11.09.25.xlsx",

"sda1":"https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%201/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%201.xlsx",

"sda2":"https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%202/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%202.xlsx",

"sda3":"https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%203/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%203.xlsx",

"sda4":"https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%204/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%204.xlsx",

"sda5":"https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%205/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%205.xlsx",

"sda6":"https://simmsa-my.sharepoint.com/personal/valquiria_andrade_simmsolucoes_com_br/Documents/Data%20book/DATABOOK%20UFV/SDA%206/SDA-SIM-E-PVRD-Q00-0148-00%20-%20%C3%8Dndice%20Data%20Book%20UFV%20-%20SDA%206.xlsx"
}


# ================= DOWNLOAD EXCEL =================

def baixar_excel(url):

    try:

        if "download=1" not in url:
            if "?" in url:
                url = url + "&download=1"
            else:
                url = url + "?download=1"

        r = requests.get(url, timeout=60)

        if r.status_code != 200:
            print("Erro download:", url)
            return None

        return BytesIO(r.content)

    except Exception as e:

        print("Erro:", e)
        return None


# ================= LEITURA =================

def carregar(nome, url):

    arquivo = baixar_excel(url)

    if arquivo is None:
        return None


    if nome == "se":

        df = pd.read_excel(
            arquivo,
            sheet_name="MC (S)",
            usecols="B:J",
            skiprows=6
        )

        df["Setor"] = ""


    elif nome == "rmt":

        df = pd.read_excel(
            arquivo,
            sheet_name="MC (R)",
            usecols="B:J",
            skiprows=6
        )

        df["Setor"] = ""


    else:

        df = pd.read_excel(
            arquivo,
            usecols="B:J",
            skiprows=6
        )


    df.columns = df.columns.str.strip()

    df = df[df["Item"].astype(str).str.match(r"^\d+\..*")]

    df["Quantidade total"] = pd.to_numeric(df["Quantidade total"], errors="coerce")

    df["Postagem"] = pd.to_numeric(df["Postagem"], errors="coerce").fillna(0)

    df = df.dropna(subset=["Quantidade total"])

    return df


# ================= ROTAS =================

@app.route("/")
def home():
    return send_from_directory(".", "index.html")


@app.route("/<path:arquivo>")
def arquivos(arquivo):
    return send_from_directory(".", arquivo)


# ================= API =================

@app.route("/dados")

def dados():

    sda = request.args.get("sda","geral")

    bases = {}

    for nome, url in ARQUIVOS.items():

        df = carregar(nome, url)

        if df is not None:
            bases[nome] = df


    if not bases:
        return jsonify({"erro":"Sem dados"})


    geral = pd.concat(bases.values())


    sdas = {}

    for nome, df in bases.items():

        total = df["Quantidade total"].sum()
        postados = df["Postagem"].sum()

        porcentagem = round((postados/total)*100,1) if total>0 else 0

        sdas[nome] = porcentagem


    base = geral if sda == "geral" else bases.get(sda, geral)

    total = int(base["Quantidade total"].sum())
    postados = int(base["Postagem"].sum())

    progresso = round((postados/total)*100,1) if total>0 else 0


    tabela = []

    for _, r in base.iterrows():

        status = "Finalizado"

        if r["Postagem"] == 0:
            status = "Pendente"

        elif r["Postagem"] < r["Quantidade total"]:
            status = "Parcial"


        if status != "Finalizado":

            tabela.append({

                "item":str(r["Item"]),
                "setor":str(r.get("Setor","")),
                "documento":str(r.get("Documento","")),
                "total":int(r["Quantidade total"]),
                "postados":int(r["Postagem"]),
                "comentario":"",
                "status":status

            })


    return jsonify({

        "total":total,
        "postados":postados,
        "progresso":progresso,
        "sdas":sdas,
        "tabela":tabela

    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
