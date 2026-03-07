@app.route("/dados")
def dados():

    sda = request.args.get("sda", "geral")

    dados = {k: carregar(v) for k, v in ARQUIVOS.items()}
    dados = {k: v for k, v in dados.items() if v is not None}

    if not dados:
        return jsonify({"erro": "Nenhum arquivo encontrado"})


    # ================= BASE GERAL =================

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
            "comentario": str(r.get("Observação", "")),
            "status": str(r["Status"])

        })


    return jsonify({

        "total": total,
        "postados": postados,
        "progresso": progresso,
        "sdas": sdas,
        "tabela": tabela

    })
