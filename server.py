from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
import os

app = Flask(__name__, static_folder=".", static_url_path="")

# caminhos dos arquivos
ARQUIVOS = {
    "se": "dados/se.xlsx",
    "rmt": "dados/rmt.xlsx",
    "sda1": "dados/sda1.xlsx",
    "sda2": "dados/sda2.xlsx",
    "sda3": "dados/sda3.xlsx",
    "sda4": "dados/sda4.xlsx",
    "sda5": "dados/sda5.xlsx",
    "sda6": "dados/sda6.xlsx"
}


def ler_planilha(path):

    if not os.path.exists(path):
        return {
            "total": 0,
            "postados": 0,
            "tabela": []
        }

    df = pd.read_excel(path)

    total = len(df)
    postados = df["POSTADOS"].sum() if "POSTADOS" in df.columns else 0

    tabela = []

    for _, r in df.iterrows():

        tabela.append({
            "item": r.get("ITEM", ""),
            "setor": r.get("SETOR", ""),
            "documento": r.get("DOCUMENTO", ""),
            "total": r.get("TOTAL", ""),
            "postados": r.get("POSTADOS", ""),
            "comentario": r.get("COMENTARIO", ""),
            "status": r.get("STATUS", "")
        })

    return {
        "total": int(total),
        "postados": int(postados),
        "tabela": tabela
    }


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/dados")
def dados():

    sda = request.args.get("sda", "geral")

    totais = 0
    postados = 0
    sdas = {}
    tabela = []

    for nome, path in ARQUIVOS.items():

        d = ler_planilha(path)

        totais += d["total"]
        postados += d["postados"]

        progresso = 0
        if d["total"] > 0:
            progresso = round((d["postados"] / d["total"]) * 100)

        sdas[nome] = progresso

        if nome == sda:
            tabela = d["tabela"]

    progresso_geral = 0
    if totais > 0:
        progresso_geral = round((postados / totais) * 100)

    return jsonify({
        "total": totais,
        "postados": postados,
        "progresso": progresso_geral,
        "sdas": sdas,
        "tabela": tabela
    })


if __name__ == "__main__":
    app.run()
