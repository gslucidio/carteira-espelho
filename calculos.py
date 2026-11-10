"""
Cálculos do dashboard, feitos em Python com as mesmas regras das fórmulas da planilha:
curva exponencial em dias úteis, PL pela gestora e pela ADM, conferências e painel.
"""
from collections import defaultdict

import numpy as np
import pandas as pd

from nucleo import TIPOS_DESPESA as TIPOS


def _dias_uteis(inicio, fim, feriados):
    """Dias úteis no intervalo (inicio, fim] — o mesmo que NETWORKDAYS(inicio+1, fim) do Excel."""
    if fim <= inicio:
        return 0
    return int(np.busday_count(np.datetime64(inicio) + 1, np.datetime64(fim) + 1, holidays=feriados))


def calcular(d):
    db = d["data_base"]
    fer = [np.datetime64(f) for f, _ in d["feriados"]]

    # ---------------------------------------------------------------- títulos
    linhas = []
    for t in d["titulos"]:
        em_carteira = t["data_aquisicao"] <= db and (t["status"] == "Aberto" or (t["data_baixa"] and t["data_baixa"] > db))
        du_tot = _dias_uteis(t["data_aquisicao"], t["vencimento"], fer)
        du_dec = _dias_uteis(t["data_aquisicao"], min(db, t["vencimento"]), fer)
        curva = t["preco_pago"] * (t["valor_face"] / t["preco_pago"]) ** (du_dec / du_tot) if em_carteira and du_tot else 0.0
        atraso = (db - t["vencimento"]).days if em_carteira and db > t["vencimento"] else 0
        ganho = (t["valor_recebido"] - t["preco_pago"]) if t["status"] == "Liquidado" and t["data_baixa"] and t["data_baixa"] <= db else 0.0
        linhas.append(dict(t, em_carteira=em_carteira, curva=curva, atraso=atraso, ganho=ganho))
    tit = pd.DataFrame(linhas)
    ab = tit[tit["em_carteira"]] if len(tit) else tit

    # ---------------------------------------------------------------- caixa
    cx = pd.DataFrame(d["caixa"])
    if len(cx):
        cx_db = cx[cx["data"] <= db]
        conta = round(float(cx_db["entrada"].sum() - cx_db["saida"].sum()), 2)
        por_classe = cx_db.groupby("classificacao")[["entrada", "saida"]].sum()
        saldo_banco = float(cx_db["saldo_banco"].iloc[-1]) if len(cx_db) else 0.0
    else:
        conta, saldo_banco, por_classe = 0.0, 0.0, pd.DataFrame(columns=["entrada", "saida"])

    def classe(nome, col):
        return float(por_classe[col].get(nome, 0.0)) if len(por_classe) else 0.0

    # ---------------------------------------------------------------- carteira gestora
    dc = float(ab["curva"].sum()) if len(ab) else 0.0
    fundos = float(sum((f["quantidade"] or 0) * (f["valor_cota"] or 0) for f in d["fundos"]))
    desp_g = float(sum(x["valor_gestora"] for x in d["despesas"]))
    pl_total = dc + fundos + conta - desp_g
    classes = d.get("classes") or []
    # passivo das classes pela gestora: o PL que cada classe tem na sua própria carteira
    passivo_g = float(sum((cl["pl"] if cl["pl"] is not None else cl["passivo_principal"]) for cl in classes))
    pl = pl_total - passivo_g          # PL da classe principal (subordinada ou cota única)
    cotas = d["quantidade_cotas"] or 0
    aportes = classe("Aporte de cotista", "entrada")
    resgates = classe("Resgate de cotista", "saida")
    capital = aportes - resgates
    gestora = dict(dc=dc, fundos=fundos, conta=conta, despesas=desp_g, pl_total=pl_total, passivo_classes=passivo_g,
                   pl=pl, cotas=cotas,
                   cota=pl / cotas if cotas else 0.0, aportes=aportes, resgates=resgates, capital=capital,
                   resultado=pl_total - capital, resultado_pct=(pl_total - capital) / capital if capital else 0.0,
                   aplicado_di=classe("Aplicação no fundo DI", "saida") - classe("Resgate do fundo DI", "entrada"))
    gestora["rendimento_di"] = fundos - gestora["aplicado_di"]

    # ---------------------------------------------------------------- carteira ADM
    a = d["adm"]
    dcs = sorted(a["direitos_creditorios"], key=lambda x: -abs(x["valor_mercado"] or 0))
    adm = dict(dc=dcs[0]["valor_mercado"] if dcs else 0.0, dc_outros=sum(x["valor_mercado"] or 0 for x in dcs[1:]),
               dc_outros_nomes=", ".join(str(x["descricao"]) for x in dcs[1:]),
               fundos=sum(f["valor"] or 0 for f in d["fundos"]), conta=a["conta_corrente"] or 0.0,
               despesas=sum(x["valor"] for x in d["despesas"]), cotas=a["quantidade_cotas"] or 0,
               pl_informado=a["pl"], cota_informada=a["valor_cota"],
               passivo_classes=float(sum(cl["passivo_principal"] for cl in classes)))
    adm["pl_total"] = adm["dc"] + adm["dc_outros"] + adm["fundos"] + adm["conta"] - adm["despesas"]
    adm["pl"] = adm["pl_total"] - adm["passivo_classes"]
    adm["cota"] = adm["pl"] / adm["cotas"] if adm["cotas"] else 0.0

    comparacao = pd.DataFrame([
        ("Direitos creditórios (curva)", gestora["dc"], adm["dc"], "Estoque"),
        ("Outros direitos creditórios na ADM", 0.0, adm["dc_outros"], adm["dc_outros_nomes"] or "-"),
        ("Fundos de liquidez", gestora["fundos"], adm["fundos"], "Extrato dos fundos"),
        ("Conta corrente", gestora["conta"], adm["conta"], "Caixa"),
        ("(−) Despesas a pagar", -gestora["despesas"], -adm["despesas"], "Despesas"),
    ] + ([
        ("Patrimônio total do fundo", gestora["pl_total"], adm["pl_total"], ""),
        ("(−) Cotas das demais classes", -gestora["passivo_classes"], -adm["passivo_classes"],
         "Passivo e cotas (gestora: PL de cada classe na sua carteira)"),
        ("PL da classe principal (subordinada)", gestora["pl"], adm["pl"], ""),
    ] if classes else [("Patrimônio líquido", gestora["pl"], adm["pl"], "")]),
        columns=["Linha", "Gestora", "Administradora", "Onde investigar"])
    comparacao["Diferença"] = comparacao["Gestora"] - comparacao["Administradora"]

    # ---------------------------------------------------------------- estoque × ADM
    eadm = pd.DataFrame(d["estoque_adm"])
    if len(eadm):
        # a comparação é feita na data do estoque da ADM, que pode ser anterior à data-base
        de = d.get("estoque_adm_data") or db
        comp = []
        for t in d["titulos"]:
            em = t["data_aquisicao"] <= de and (t["status"] == "Aberto" or (t["data_baixa"] and t["data_baixa"] > de))
            if not em:
                continue
            du_tot = _dias_uteis(t["data_aquisicao"], t["vencimento"], fer)
            du_dec = _dias_uteis(t["data_aquisicao"], min(de, t["vencimento"]), fer)
            cv = t["preco_pago"] * (t["valor_face"] / t["preco_pago"]) ** (du_dec / du_tot) if du_tot else t["valor_face"]
            comp.append(dict(chave=t["chave"], valor_face=t["valor_face"], curva=cv))
        comp = pd.DataFrame(comp, columns=["chave", "valor_face", "curva"])
        g = comp.groupby("chave")[["valor_face", "curva"]].sum()
        eadm = eadm.join(g, on="chave")
        eadm[["valor_face", "curva"]] = eadm[["valor_face", "curva"]].fillna(0.0)
        eadm["dif_face"] = eadm["valor_face"] - eadm["valor_nominal"]
        eadm["dif_curva"] = eadm["curva"] - eadm["valor_presente"]
        eadm["situacao_conf"] = np.where(eadm["valor_face"] == 0, "Não está em carteira na gestora",
                                         np.where((eadm["dif_face"].abs() < 0.01) & (eadm["dif_curva"].abs() < 0.05), "OK", "Diferente"))
        chaves_adm = set(eadm["chave"])
    else:
        chaves_adm = set()
    em_data_est = set(comp["chave"]) if len(eadm) else set(ab["chave"]) if len(ab) else set()
    faltam_na_adm = ab[ab["chave"].isin(em_data_est) & ~ab["chave"].isin(chaves_adm)] if len(ab) else ab
    prob_adm = int((eadm["situacao_conf"] != "OK").sum()) if len(eadm) else 0

    # ---------------------------------------------------------------- conferências
    tem_estoque = len(eadm) > 0
    conf = [
        ("Estoque", "Face em carteira × estoque analítico ADM (na data do estoque)",
         float(comp["valor_face"].sum()) if tem_estoque else None, float(eadm["valor_nominal"].sum()) if tem_estoque else None, 0.01),
        ("Estoque", "Valor na curva × estoque analítico ADM (na data do estoque)",
         float(comp["curva"].sum()) if tem_estoque else None, float(eadm["valor_presente"].sum()) if tem_estoque else None, 1),
        ("Estoque", "Valor na curva × carteira ADM", dc, adm["dc"], 1),
        ("Estoque", "Quantidade de títulos (na data do estoque)", len(comp) if tem_estoque else len(ab),
         len(eadm) if tem_estoque else None, 0.5),
        ("Estoque", "Títulos divergentes (gestora + ADM)", len(faltam_na_adm) + prob_adm, 0, 0.5),
        ("Caixa", "Pago na compra de DC × extrato",
         float(tit.loc[tit["data_aquisicao"] <= db, "preco_pago"].sum()) if len(tit) else 0.0, classe("Aquisição de DC", "saida"), 0.01),
        ("Caixa", "Recebido em liquidações × extrato",
         float(tit.loc[tit["data_baixa"].notna() & (tit["data_baixa"] <= db), "valor_recebido"].sum()) if len(tit) else 0.0,
         classe("Recebimento de DC", "entrada"), 0.01),
        ("Caixa", "Saldo calculado × saldo do banco", conta, saldo_banco, 0.01),
        ("Caixa", "Conta corrente gestora × carteira ADM", conta, adm["conta"], 0.01),
        ("ADM", "PL da ADM: soma das linhas × informado", adm["pl"], adm["pl_informado"], 0.05),
    ]
    conferencias = pd.DataFrame(conf, columns=["Grupo", "Checagem", "Gestora", "Referência", "tol"])
    conferencias["Diferença"] = conferencias["Gestora"] - conferencias["Referência"]
    conferencias["OK"] = conferencias.apply(
        lambda r: None if pd.isna(r["Referência"]) or pd.isna(r["Gestora"]) else abs(r["Diferença"]) < r["tol"], axis=1)

    # ---------------------------------------------------------------- painel
    conc = (ab.groupby("sacado")["valor_face"].sum().sort_values(ascending=False).reset_index()
            if len(ab) else pd.DataFrame(columns=["sacado", "valor_face"]))
    tot_face = conc["valor_face"].sum()
    conc["pct_dc"] = conc["valor_face"] / tot_face if tot_face else 0.0
    conc["pct_pl"] = conc["valor_face"] / pl_total if pl_total else 0.0
    venc = (ab.groupby("vencimento").agg(face=("valor_face", "sum"), titulos=("chave", "count")).reset_index()
            if len(ab) else pd.DataFrame(columns=["vencimento", "face", "titulos"]))
    venc["dias"] = venc["vencimento"].apply(lambda v: (v - db).days) if len(venc) else []
    dc_pl = dc / pl_total if pl_total else 0.0
    inicio = d.get("inicio_classe")
    limite = inicio + pd.Timedelta(days=int(d["prazo_enquadramento"])).to_pytimedelta() if inicio else None
    if dc_pl >= d["minimo_dc"]:
        situacao = "Enquadrado"
    elif limite and db > limite:
        situacao = "DESENQUADRADO"
    else:
        situacao = "Abaixo do mínimo, ainda no prazo" if limite else "Abaixo do mínimo (informe o início da Classe)"
    vencidos = ab[ab["atraso"] > 0] if len(ab) else ab
    painel = dict(concentracao=conc, vencimentos=venc, dc_pl=dc_pl, minimo=d["minimo_dc"], inicio=inicio,
                  limite=limite, situacao=situacao, vencidos_qtd=len(vencidos),
                  vencidos_face=float(vencidos["valor_face"].sum()) if len(vencidos) else 0.0,
                  maior_atraso=int(vencidos["atraso"].max()) if len(vencidos) else 0)

    # ---------------------------------------------------------------- DRE (competência, acumulado até a data-base)
    realizado = float(tit["ganho"].sum()) if len(tit) else 0.0
    apropriado = float((ab["curva"] - ab["preco_pago"]).sum()) if len(ab) else 0.0
    outros_liq = sum(classe(k, "entrada") - classe(k, "saida") for k in ("Outros", "A classificar"))
    receitas = [("Deságio realizado em títulos liquidados", realizado),
                ("Deságio apropriado dos títulos em carteira", apropriado),
                ("Rendimento dos fundos de liquidez", gestora["rendimento_di"])]
    if abs(outros_liq) >= 0.01:
        receitas.append(("Outros movimentos / não classificados (líquido)", outros_liq))
    pagas = defaultdict(float)
    if len(cx):
        for _, m in cx[(cx["data"] <= db) & cx["classificacao"].isin(["Taxas e despesas", "Tarifa bancária"])].iterrows():
            pagas[m["tipo"] or "Outras despesas"] += m["saida"] - m["entrada"]
    a_pagar = defaultdict(float)
    for x in d["despesas"]:
        a_pagar[x.get("tipo") or "Outras despesas"] += x["valor_gestora"]
    tipos = [t for t in TIPOS if pagas.get(t) or a_pagar.get(t)] + \
        [t for t in set(pagas) | set(a_pagar) if t not in TIPOS]
    despesas_dre = pd.DataFrame([(t, pagas.get(t, 0.0), a_pagar.get(t, 0.0)) for t in tipos],
                                columns=["Tipo", "Pagas", "A pagar"])
    despesas_dre["Total"] = despesas_dre["Pagas"] + despesas_dre["A pagar"]
    tot_rec = sum(v for _, v in receitas)
    tot_desp = float(despesas_dre["Total"].sum()) if len(despesas_dre) else 0.0
    dre = dict(receitas=pd.DataFrame(receitas, columns=["Linha", "Valor"]), despesas=despesas_dre,
               total_receitas=tot_rec, total_despesas=tot_desp, resultado=tot_rec - tot_desp,
               conferencia=gestora["resultado"], diferenca=(tot_rec - tot_desp) - gestora["resultado"])

    # ---------------------------------------------------------------- fluxo de caixa mensal
    grupos = {"Aporte de cotista": "Aportes de cotistas", "Resgate de cotista": "Resgates e amortizações",
              "Aquisição de DC": "Aquisição de direitos creditórios", "Recebimento de DC": "Recebimento de direitos creditórios",
              "Aplicação no fundo DI": "Fundos de liquidez (líquido)", "Resgate do fundo DI": "Fundos de liquidez (líquido)"}
    if len(cx):
        f = cx[cx["data"] <= db].copy()
        f["Mês"] = f["data"].apply(lambda v: v.strftime("%Y-%m"))
        f["Grupo"] = [grupos.get(k) or (("Despesa: " + t) if t else "Outros / não classificados")
                      for k, t in zip(f["classificacao"], f["tipo"])]
        f["Líquido"] = f["entrada"] - f["saida"]
        ordem = list(dict.fromkeys(list(grupos.values()) + ["Despesa: " + t for t in TIPOS] + ["Outros / não classificados"]))
        fluxo = f.pivot_table(index="Grupo", columns="Mês", values="Líquido", aggfunc="sum", fill_value=0.0)
        fluxo = fluxo.reindex([g for g in ordem if g in fluxo.index])
        fluxo.loc["Variação do caixa"] = fluxo.sum()
    else:
        fluxo = pd.DataFrame()

    # ---------------------------------------------------------------- passivo por classe
    passivo = pd.DataFrame()
    if classes:
        linhas_p = []
        rp = d["adm"].get("rentabilidades", {})
        linhas_p.append(dict(classe="Subordinada (carteira principal)", indice="residual", spread=None, cotas=cotas,
                             cota_adm=a["valor_cota"], pl_adm=adm["pl"], pl_gestora=pl, passivo_principal=None,
                             aplicado=None, rent_inicial=rp.get("Carteira", {}).get("Inicial"),
                             rent_mes=rp.get("Carteira", {}).get("Mês"), cdi_mes=rp.get("(%) CDI", {}).get("Mês"),
                             dia=rp.get("Carteira", {}).get("Dia"), cdi_dia=rp.get("(%) CDI", {}).get("Dia")))
        for cl in classes:
            r_ = cl["rentabilidades"]
            linhas_p.append(dict(classe=cl["nome"], indice=cl["indice"], spread=cl["spread"], cotas=cl["cotas"],
                                 cota_adm=cl["cota"], pl_adm=cl["pl"] if cl["pl"] is not None else cl["passivo_principal"],
                                 pl_gestora=cl["pl"] if cl["pl"] is not None else cl["passivo_principal"],
                                 passivo_principal=cl["passivo_principal"], aplicado=cl["aplicado"],
                                 rent_inicial=r_.get("Carteira", {}).get("Inicial"), rent_mes=r_.get("Carteira", {}).get("Mês"),
                                 cdi_mes=r_.get("(%) CDI", {}).get("Mês"), dia=r_.get("Carteira", {}).get("Dia"),
                                 cdi_dia=r_.get("(%) CDI", {}).get("Dia")))
        passivo = pd.DataFrame(linhas_p)
        passivo["pct_pl"] = passivo["pl_gestora"] / pl_total if pl_total else 0.0
        passivo["dif_principal"] = passivo["passivo_principal"] - passivo["pl_adm"]
        # aderência ao benchmark no último dia: CDI implícito = rent. do dia ÷ %CDI do dia
        def checa(row):
            if not row["spread"] or not row["dia"] or not row["cdi_dia"]:
                return None, None
            cdi = row["dia"] / row["cdi_dia"]
            esperado = (1 + cdi) * (1 + row["spread"]) ** (1 / 252) - 1
            return esperado, abs(row["dia"] - esperado) < 0.00002
        res = passivo.apply(checa, axis=1, result_type="expand")
        passivo["dia_esperado"] = res[0]
        passivo["benchmark_ok"] = [None if e is None or pd.isna(e) else bool(ok) for e, ok in zip(res[0], res[1])]
        # índices de subordinação: assume que, quanto maior o spread, mais júnior a classe (empate: ordem inversa do nome)
        mez = passivo.iloc[1:].copy()
        mez["_ord"] = list(zip(-mez["spread"].fillna(0), [-i for i in range(len(mez))]))
        ordem = ["Subordinada (carteira principal)"] + list(mez.sort_values("_ord", ascending=True)["classe"])
        acumulado, subord = 0.0, {}
        for nome_c in ordem:
            v = float(passivo.loc[passivo["classe"] == nome_c, "pl_gestora"].iloc[0])
            subord[nome_c] = acumulado / pl_total if pl_total else 0.0   # quanto está abaixo desta classe
            acumulado += v
        passivo["subordinacao_abaixo"] = passivo["classe"].map(subord)
        passivo["ordem"] = passivo["classe"].map({n_: i for i, n_ in enumerate(ordem)})

    return dict(dre=dre, passivo=passivo, fluxo_mensal=fluxo, titulos=tit, em_carteira=ab, caixa=cx, por_classe=por_classe, gestora=gestora, adm=adm,
                comparacao=comparacao, estoque_adm=eadm, faltam_na_adm=faltam_na_adm,
                conferencias=conferencias, painel=painel)
