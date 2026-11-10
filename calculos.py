"""
Cálculos do dashboard, feitos em Python com as mesmas regras das fórmulas da planilha:
curva exponencial em dias úteis, PL pela gestora e pela ADM, conferências e painel.
"""
import numpy as np
import pandas as pd


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
    pl = dc + fundos + conta - desp_g
    cotas = d["quantidade_cotas"] or 0
    aportes = classe("Aporte de cotista", "entrada")
    resgates = classe("Resgate de cotista", "saida")
    capital = aportes - resgates
    gestora = dict(dc=dc, fundos=fundos, conta=conta, despesas=desp_g, pl=pl, cotas=cotas,
                   cota=pl / cotas if cotas else 0.0, aportes=aportes, resgates=resgates, capital=capital,
                   resultado=pl - capital, resultado_pct=(pl - capital) / capital if capital else 0.0,
                   aplicado_di=classe("Aplicação no fundo DI", "saida") - classe("Resgate do fundo DI", "entrada"))
    gestora["rendimento_di"] = fundos - gestora["aplicado_di"]

    # ---------------------------------------------------------------- carteira ADM
    a = d["adm"]
    dcs = sorted(a["direitos_creditorios"], key=lambda x: -abs(x["valor_mercado"] or 0))
    adm = dict(dc=dcs[0]["valor_mercado"] if dcs else 0.0, dc_outros=sum(x["valor_mercado"] or 0 for x in dcs[1:]),
               dc_outros_nomes=", ".join(str(x["descricao"]) for x in dcs[1:]),
               fundos=sum(f["valor"] or 0 for f in d["fundos"]), conta=a["conta_corrente"] or 0.0,
               despesas=sum(x["valor"] for x in d["despesas"]), cotas=a["quantidade_cotas"] or 0,
               pl_informado=a["pl"], cota_informada=a["valor_cota"])
    adm["pl"] = adm["dc"] + adm["dc_outros"] + adm["fundos"] + adm["conta"] - adm["despesas"]
    adm["cota"] = adm["pl"] / adm["cotas"] if adm["cotas"] else 0.0

    comparacao = pd.DataFrame([
        ("Direitos creditórios (curva)", gestora["dc"], adm["dc"], "Estoque"),
        ("Outros direitos creditórios na ADM", 0.0, adm["dc_outros"], adm["dc_outros_nomes"] or "-"),
        ("Fundos de liquidez", gestora["fundos"], adm["fundos"], "Extrato dos fundos"),
        ("Conta corrente", gestora["conta"], adm["conta"], "Caixa"),
        ("(−) Despesas a pagar", -gestora["despesas"], -adm["despesas"], "Despesas"),
        ("Patrimônio líquido", gestora["pl"], adm["pl"], ""),
    ], columns=["Linha", "Gestora", "Administradora", "Onde investigar"])
    comparacao["Diferença"] = comparacao["Gestora"] - comparacao["Administradora"]

    # ---------------------------------------------------------------- estoque × ADM
    eadm = pd.DataFrame(d["estoque_adm"])
    if len(eadm):
        g = ab.groupby("chave")[["valor_face", "curva"]].sum() if len(ab) else pd.DataFrame(columns=["valor_face", "curva"])
        eadm = eadm.join(g, on="chave")
        eadm[["valor_face", "curva"]] = eadm[["valor_face", "curva"]].fillna(0.0)
        eadm["dif_face"] = eadm["valor_face"] - eadm["valor_nominal"]
        eadm["dif_curva"] = eadm["curva"] - eadm["valor_presente"]
        eadm["situacao_conf"] = np.where(eadm["valor_face"] == 0, "Não está em carteira na gestora",
                                         np.where((eadm["dif_face"].abs() < 0.01) & (eadm["dif_curva"].abs() < 0.05), "OK", "Diferente"))
        chaves_adm = set(eadm["chave"])
    else:
        chaves_adm = set()
    faltam_na_adm = ab[~ab["chave"].isin(chaves_adm)] if len(ab) else ab
    prob_adm = int((eadm["situacao_conf"] != "OK").sum()) if len(eadm) else 0

    # ---------------------------------------------------------------- conferências
    tem_estoque = len(eadm) > 0
    conf = [
        ("Estoque", "Face em carteira × estoque analítico ADM", float(ab["valor_face"].sum()) if len(ab) else 0.0,
         float(eadm["valor_nominal"].sum()) if tem_estoque else None, 0.01),
        ("Estoque", "Valor na curva × estoque analítico ADM", dc, float(eadm["valor_presente"].sum()) if tem_estoque else None, 1),
        ("Estoque", "Valor na curva × carteira ADM", dc, adm["dc"], 1),
        ("Estoque", "Quantidade de títulos em carteira", len(ab), len(eadm) if tem_estoque else None, 0.5),
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
        lambda r: None if pd.isna(r["Referência"]) else abs(r["Diferença"]) < r["tol"], axis=1)

    # ---------------------------------------------------------------- painel
    conc = (ab.groupby("sacado")["valor_face"].sum().sort_values(ascending=False).reset_index()
            if len(ab) else pd.DataFrame(columns=["sacado", "valor_face"]))
    tot_face = conc["valor_face"].sum()
    conc["pct_dc"] = conc["valor_face"] / tot_face if tot_face else 0.0
    conc["pct_pl"] = conc["valor_face"] / pl if pl else 0.0
    venc = (ab.groupby("vencimento").agg(face=("valor_face", "sum"), titulos=("chave", "count")).reset_index()
            if len(ab) else pd.DataFrame(columns=["vencimento", "face", "titulos"]))
    venc["dias"] = venc["vencimento"].apply(lambda v: (v - db).days) if len(venc) else []
    dc_pl = dc / pl if pl else 0.0
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

    return dict(titulos=tit, em_carteira=ab, caixa=cx, por_classe=por_classe, gestora=gestora, adm=adm,
                comparacao=comparacao, estoque_adm=eadm, faltam_na_adm=faltam_na_adm,
                conferencias=conferencias, painel=painel)
