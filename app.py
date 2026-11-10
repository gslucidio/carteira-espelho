"""
Carteira Espelho — dashboard Streamlit da VIZ Gestora.

Os arquivos enviados ficam só na memória da sessão (nada é gravado em disco nem em cache
compartilhado). As correções feitas na tela vão para a planilha baixada no fim, que é a
"memória" da próxima sessão.
"""
import copy
import io

import altair as alt
import pandas as pd
import streamlit as st

import calculos
import nucleo
import planilha

st.set_page_config(page_title="Carteira Espelho - VIZ Gestora", page_icon="📒", layout="wide",
                   initial_sidebar_state="expanded")

CLASSES = planilha.CLASSES_CAIXA


# =====================================================================================
# formatação
def brl(v, casas=2):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    s = f"{abs(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-R$ {s}" if v < -0.005 else f"R$ {s}"


def num(v, casas=2):
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v):
    return f"{v * 100:.1f}%".replace(".", ",")


def data_br(v):
    return v.strftime("%d/%m/%Y") if v is not None and not pd.isna(v) else "-"


def tabela(df, moeda=(), datas=(), pcts=(), destaque=None):
    """Mostra um DataFrame com formatação brasileira, sem mudar os valores."""
    fmt = {c: brl for c in moeda} | {c: data_br for c in datas} | {c: pct for c in pcts}
    sty = df.style.format(fmt, na_rep="-")
    if destaque:
        col, func = destaque
        sty = sty.apply(lambda s: ["background-color: #fde2e1" if func(v) else "" for v in s], subset=[col])
    st.dataframe(sty, width="stretch", hide_index=True)


# =====================================================================================
# estado da sessão (correções feitas na tela)
for k in ("ed_caixa", "ed_desp", "ed_painel"):
    st.session_state.setdefault(k, {})

# =====================================================================================
# barra lateral: arquivos por categoria
st.sidebar.title("📒 Carteira Espelho")
st.sidebar.caption("VIZ Gestora · gestão de carteira dos FIDCs")

with st.sidebar.expander("📂 Arquivos", expanded=True):
    up_cnab = st.file_uploader("1 · CNABs: remessas e baixas", type=["rem", "txt", "ret"], accept_multiple_files=True,
                               help="Todos os CNABs desde o início do fundo, de remessa e de baixa.")
    up_cart = st.file_uploader("2 · Carteira diária da ADM (.xlsx)", type=["xlsx"], accept_multiple_files=True,
                               help="Composição de Carteira. Em fundos com várias classes, envie a carteira principal "
                                    "(subordinada) e a de cada mezanino/sênior, da mesma data. Vale a data mais recente.")
    up_est = st.file_uploader("3 · Estoque analítico da ADM (.csv)", type=["csv"], accept_multiple_files=True,
                              help="Estoque título a título, da mesma data da carteira.")
    up_ext = st.file_uploader("4 · Extrato da conta (.csv)", type=["csv"], accept_multiple_files=True,
                              help="Extrato da conta corrente do fundo. Pode enviar vários períodos.")
    st.markdown("---")
    up_ant = st.file_uploader("5 · Última planilha baixada (opcional)", type=["xlsx"],
                              help="A planilha que você baixou no uso anterior. Dela vêm as suas correções: "
                                   "classificações do caixa, despesas ajustadas e dados de enquadramento.")

menu = st.sidebar.radio("Navegação", ["🏠 Visão geral", "📊 DRE", "🏦 Passivo e cotas", "📦 Estoque", "💰 Caixa e despesas",
                                      "✅ Conferências", "📥 Exportar"])

# =====================================================================================
# sem arquivos indispensáveis: tela de boas-vindas
if not up_cart:
    st.title("📒 Carteira Espelho")
    st.markdown("Controle independente da carteira dos FIDCs pela gestora, comparado com a administradora.")
    st.markdown("#### Envie os arquivos pela barra lateral, cada um na sua categoria")
    c = st.columns(5)
    itens = [("1 · CNABs", "Remessas e baixas (.REM), desde o início do fundo", up_cnab, False),
             ("2 · Carteira ADM", "Composição de carteira do dia (.xlsx) — obrigatório", up_cart, True),
             ("3 · Estoque ADM", "Estoque analítico título a título (.csv)", up_est, False),
             ("4 · Extrato", "Extrato da conta corrente (.csv)", up_ext, False),
             ("5 · Planilha anterior", "Opcional: traz as suas correções do último uso", up_ant, False)]
    for col, (tit, desc, arq, obrig) in zip(c, itens):
        with col:
            ok = bool(arq)
            st.markdown(f"**{'✅' if ok else ('⛔' if obrig else '⬜')} {tit}**")
            st.caption(desc)
    st.info("Os arquivos ficam só na memória desta sessão e somem ao fechar a aba. "
            "Ao terminar, baixe a planilha: ela guarda as suas correções para a próxima vez.")
    st.stop()

# =====================================================================================
# processamento (guardado na sessão enquanto os arquivos não mudarem)
assinatura = tuple((f.name, f.size) for grupo in (up_cnab, up_cart, up_est, up_ext) for f in grupo) + \
    ((up_ant.name, up_ant.size) if up_ant else ())
if st.session_state.get("assinatura") != assinatura:
    try:
        with st.spinner("Lendo arquivos..."):
            st.session_state["base"] = nucleo.processar(up_cnab, up_cart, up_est, up_ext, up_ant)
        st.session_state["assinatura"] = assinatura
    except nucleo.ErroEntrada as e:
        st.error(str(e))
        st.stop()

dados = copy.deepcopy(st.session_state["base"])
for m in dados["caixa"]:
    if m["id"] in st.session_state["ed_caixa"]:
        m["classificacao"], m["referencia"] = st.session_state["ed_caixa"][m["id"]]
for x in dados["despesas"]:
    if x["descricao"] in st.session_state["ed_desp"]:
        x["valor_gestora"], x["observacao"] = st.session_state["ed_desp"][x["descricao"]]
ep = st.session_state["ed_painel"]
dados["inicio_classe"] = ep.get("inicio", dados["inicio_classe"])
dados["minimo_dc"] = ep.get("minimo", dados["minimo_dc"])
dados["prazo_enquadramento"] = ep.get("prazo", dados["prazo_enquadramento"])
nucleo.atribuir_tipos(dados)
r = calculos.calcular(dados)
G, A, P = r["gestora"], r["adm"], r["painel"]
a_classificar = sum(1 for m in dados["caixa"] if m["classificacao"] == "A classificar")
conf = r["conferencias"]
n_ok, n_conf = int((conf["OK"] == True).sum()), int(conf["OK"].notna().sum())  # noqa: E712


def excel_bytes():
    buf = io.BytesIO()
    planilha.montar(dados, buf)
    return buf.getvalue()


st.sidebar.markdown("---")
st.sidebar.download_button("📥 Baixar planilha", data=excel_bytes(), file_name=dados["arquivo_sugerido"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")
st.sidebar.caption("Baixe ao terminar: a planilha guarda as correções feitas aqui.")

# cabeçalho comum
st.title(f"📒 {dados['fundo']}")
st.caption(f"Data-base {data_br(dados['data_base'])} · carteira ADM: {dados['carteira_adm_arquivo']} · "
           f"estoque ADM: {dados['estoque_adm_arquivo']} · correções de: {dados['memoria'] or 'nenhuma planilha anterior'}")

# =====================================================================================
if menu == "🏠 Visão geral":
    c1, c2, c3 = st.columns(3)
    c1.markdown(f"### {'✅' if n_ok == n_conf else '⚠️'} {n_ok}/{n_conf} conferências OK")
    c2.markdown(f"### {'✅' if not dados['alertas'] else '⚠️'} {len(dados['alertas'])} alerta(s)")
    c3.markdown(f"### {'✅' if not a_classificar else '🟡'} {a_classificar} lançamento(s) a classificar")
    st.markdown("---")

    m1, m2, m3, m4 = st.columns(4)
    dif_pl = G["pl"] - A["pl"]
    rotulo_pl = "PL da subordinada (gestora)" if dados.get("classes") else "PL pela gestora"
    m1.metric(rotulo_pl, brl(G["pl"]), f"{brl(dif_pl)} vs ADM" if abs(dif_pl) >= 1 else "igual à ADM",
              delta_color="off" if abs(dif_pl) < 1 else "normal")
    if dados.get("classes"):
        st.caption(f"Patrimônio total do fundo (todas as classes): {brl(G['pl_total'])} pela gestora · "
                   f"{brl(A['pl_total'])} pela ADM")
    m2.metric("Valor da cota (gestora)", num(G["cota"], 6),
              f"{num(G['cota'] - A['cota'], 6)} vs ADM" if abs(dif_pl) >= 1 else "igual à ADM",
              delta_color="off" if abs(dif_pl) < 1 else "normal")
    if G["capital"]:
        m3.metric("Resultado acumulado", brl(G["resultado"]), pct(G["resultado_pct"]) + " do capital aportado")
    else:
        m3.metric("Resultado acumulado", "-", "classifique os aportes de cotista no Caixa", delta_color="off")
    m4.metric("Direitos creditórios / PL", pct(P["dc_pl"]), f"mínimo {pct(P['minimo'])} · {P['situacao']}",
              delta_color="off")

    st.subheader("Carteira: gestora × administradora")
    tabela(r["comparacao"][["Linha", "Gestora", "Administradora", "Diferença", "Onde investigar"]],
           moeda=("Gestora", "Administradora", "Diferença"), destaque=("Diferença", lambda v: abs(v) >= 1))
    st.caption("Diferenças acima de R$ 1,00 merecem explicação. Centavos são arredondamento.")

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Composição do PL (gestora)")
        comp = pd.DataFrame({"Item": ["Direitos creditórios", "Fundos de liquidez", "Conta corrente"],
                             "Valor": [G["dc"], G["fundos"], max(G["conta"], 0)]})
        st.altair_chart(alt.Chart(comp).mark_arc(innerRadius=60).encode(
            theta="Valor:Q", color=alt.Color("Item:N", legend=alt.Legend(orient="bottom")),
            tooltip=["Item", alt.Tooltip("Valor:Q", format=",.2f")]).properties(height=280), width="stretch")
    with col_b:
        st.subheader("Passivo")
        st.markdown(f"""
| | |
|---|---|
| Aportes dos cotistas | {brl(G['aportes'])} |
| Resgates e amortizações | {brl(G['resgates'])} |
| Capital líquido investido | {brl(G['capital'])} |
| PL pela gestora | {brl(G['pl'])} |
| **Resultado acumulado** | **{brl(G['resultado'])} ({pct(G['resultado_pct'])})** |
| Quantidade de cotas | {num(G['cotas'], 8)} |
""")

    if dados["alertas"]:
        st.subheader("Alertas")
        for a in dados["alertas"]:
            st.warning(a)
    if dados["avisos"]:
        st.subheader("Correções mantidas da planilha anterior")
        for a in dados["avisos"]:
            st.info(a)

# =====================================================================================
elif menu == "📊 DRE":
    D = r["dre"]
    st.subheader("Demonstração do resultado — acumulado desde o início")
    st.caption(f"Regime de competência, até {data_br(dados['data_base'])}: receitas apropriadas e despesas pagas ou a pagar.")
    m1, m2, m3 = st.columns(3)
    m1.metric("Receitas", brl(D["total_receitas"]))
    m2.metric("Despesas", brl(-D["total_despesas"]))
    m3.metric("Resultado", brl(D["resultado"]))

    linhas = [("RECEITAS", None, True)] + [(f"   {l}", v, False) for l, v in D["receitas"].itertuples(index=False)] + \
             [("Total de receitas", D["total_receitas"], True), ("DESPESAS", None, True)] + \
             [(f"   {t}", -tot, False) for t, tot in zip(D["despesas"]["Tipo"], D["despesas"]["Total"])] + \
             [("Total de despesas", -D["total_despesas"], True), ("RESULTADO DO PERÍODO", D["resultado"], True)]
    tab = pd.DataFrame([(l, brl(v) if v is not None else "",
                         pct(v / D["total_receitas"]) if v is not None and D["total_receitas"] else "") for l, v, _ in linhas],
                       columns=["", "Valor", "% das receitas"])
    negrito = [b for _, _, b in linhas]
    st.dataframe(tab.style.apply(lambda row: ["font-weight: bold" if negrito[row.name] else "" for _ in row], axis=1),
                 hide_index=True, width="stretch", height=(len(tab) + 1) * 35 + 3)
    if abs(D["diferenca"]) < 0.05:
        st.success(f"Confere com a carteira: o resultado do DRE é igual ao PL menos o capital líquido aportado "
                   f"({brl(D['conferencia'])}).")
    else:
        st.error(f"O DRE difere do resultado pela carteira (PL − capital) em {brl(D['diferenca'])}. "
                 "Verifique lançamentos de caixa classificados de forma incoerente.")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Despesas por tipo: pagas e a pagar**")
        tabela(D["despesas"], moeda=("Pagas", "A pagar", "Total"))
    with c2:
        dd = D["despesas"].melt(id_vars="Tipo", value_vars=["Pagas", "A pagar"], var_name="Situação", value_name="Valor")
        if len(dd):
            st.altair_chart(alt.Chart(dd).mark_bar().encode(
                x=alt.X("Valor:Q", title=None), y=alt.Y("Tipo:N", sort="-x", title=None), color="Situação:N",
                tooltip=["Tipo", "Situação", alt.Tooltip("Valor:Q", format=",.2f")]).properties(height=220), width="stretch")

    st.subheader("Fluxo de caixa por mês")
    st.caption("Regime de caixa: o que efetivamente entrou (+) e saiu (−) da conta corrente em cada mês.")
    fl = r["fluxo_mensal"]
    if len(fl):
        fl_fmt = fl.copy()
        fl_fmt.columns = [pd.Timestamp(c + "-01").strftime("%b/%y") for c in fl.columns]
        st.dataframe(fl_fmt.style.format(brl).apply(
            lambda row: ["font-weight: bold" if row.name == "Variação do caixa" else "" for _ in row], axis=1),
            width="stretch")
    st.caption("Despesas pagas por tipo dependem da classificação do caixa. Pagamentos à ID CTVM entram em "
               "'Administração e custódia', porque o extrato não separa as duas taxas.")

# =====================================================================================
elif menu == "🏦 Passivo e cotas":
    pv = r["passivo"]
    if not len(pv):
        st.info("Fundo de classe única. Se o fundo tiver mezanino ou sênior, envie também as carteiras dessas classes "
                "(da mesma data) na categoria 2.")
        st.markdown(f"**Cotas:** {num(G['cotas'], 8)} · **Cota pela gestora:** {num(G['cota'], 8)} · "
                    f"**Cota ADM:** {num(A['cota_informada'] or 0, 8)}")
    else:
        st.subheader("Composição do passivo")
        k = st.columns(len(pv))
        for col, (_, row) in zip(k, pv.sort_values("ordem").iterrows()):
            col.metric(row["classe"], brl(row["pl_gestora"]), f"{pct(row['pct_pl'])} do PL total", delta_color="off")
        graf = pv.assign(rotulo=pv["classe"]).sort_values("ordem")
        st.altair_chart(alt.Chart(graf).mark_bar().encode(
            x=alt.X("sum(pl_gestora):Q", stack="normalize", axis=alt.Axis(format="%"), title="% do patrimônio total"),
            color=alt.Color("rotulo:N", sort=list(graf["rotulo"]), title="Classe"),
            order=alt.Order("ordem:Q"),
            tooltip=["classe", alt.Tooltip("pl_gestora:Q", format=",.2f", title="PL"),
                     alt.Tooltip("pct_pl:Q", format=".1%", title="% PL")]).properties(height=90), width="stretch")

        st.subheader("Classes: valores, rentabilidade e benchmark")
        tb = pv.sort_values("ordem")[["classe", "indice", "cotas", "cota_adm", "pl_adm", "aplicado", "rent_inicial",
                                      "rent_mes", "cdi_mes", "benchmark_ok"]].copy()
        tb["cotas"] = tb["cotas"].apply(lambda v: num(v, 4))
        tb["cota_adm"] = tb["cota_adm"].apply(lambda v: num(v, 6) if v else "-")
        tb["cdi_mes"] = tb["cdi_mes"].apply(lambda v: pct(v) if v is not None and not pd.isna(v) else "-")
        tb["benchmark_ok"] = tb["benchmark_ok"].map({True: "✅ aderente", False: "❌ fora"}).fillna("-")
        tabela(tb.rename(columns={"classe": "Classe", "indice": "Benchmark", "cotas": "Cotas", "cota_adm": "Cota",
                                  "pl_adm": "PL", "aplicado": "Capital aplicado", "rent_inicial": "Rent. desde o início",
                                  "rent_mes": "Rent. no mês", "cdi_mes": "% do CDI no mês", "benchmark_ok": "Benchmark (último dia)"}),
               moeda=("PL", "Capital aplicado"), pcts=("Rent. desde o início", "Rent. no mês"))
        st.caption("Benchmark: o CDI do último dia é deduzido da própria carteira da ADM (rentabilidade do dia ÷ % do CDI) "
                   "e a rentabilidade esperada é CDI + spread do índice. 'Aderente' = diferença menor que 0,002 ponto percentual.")

        st.subheader("Índices de subordinação")
        sb = pv.sort_values("ordem", ascending=False)[["classe", "pl_gestora", "subordinacao_abaixo"]]
        tabela(sb.rename(columns={"classe": "Classe", "pl_gestora": "PL",
                                  "subordinacao_abaixo": "Classes abaixo dela (% do PL total)"}),
               moeda=("PL",), pcts=("Classes abaixo dela (% do PL total)",))
        st.caption("Ordem de prioridade assumida: subordinada por último; entre as mezaninos, a de maior spread é a mais "
                   "júnior. Confirme com o regulamento e compare com os índices mínimos exigidos.")

        st.subheader("Carteira principal × carteiras das classes")
        cmp_ = pv.iloc[1:][["classe", "passivo_principal", "pl_adm", "dif_principal"]]
        tabela(cmp_.rename(columns={"classe": "Classe", "passivo_principal": "Passivo na carteira principal",
                                    "pl_adm": "PL na carteira da classe", "dif_principal": "Diferença"}),
               moeda=("Passivo na carteira principal", "PL na carteira da classe", "Diferença"),
               destaque=("Diferença", lambda v: abs(v) >= 1))
        for cl in dados["classes"]:
            if cl["divergencias"]:
                st.error(f"{cl['nome']}: aplicações registradas de forma diferente nas duas carteiras")
                dv = pd.DataFrame(cl["divergencias"])
                dv["dif"] = dv["valor_principal"].fillna(0) - dv["valor_classe"].fillna(0)
                dv["cotas_principal"] = dv["cotas_principal"].apply(lambda v: num(v, 6) if v is not None and not pd.isna(v) else "-")
                dv["cotas_classe"] = dv["cotas_classe"].apply(lambda v: num(v, 6) if v is not None and not pd.isna(v) else "-")
                tabela(dv.rename(columns={"data": "Aplicação", "valor_aplicado": "Valor aplicado",
                                          "cotas_principal": "Cotas (principal)", "cotas_classe": "Cotas (classe)",
                                          "valor_principal": "Valor (principal)", "valor_classe": "Valor (classe)",
                                          "dif": "Diferença"}),
                       moeda=("Valor aplicado", "Valor (principal)", "Valor (classe)", "Diferença"), datas=("Aplicação",))
        st.caption("Na visão da gestora, o passivo de cada classe é o PL da carteira da própria classe. Uma diferença "
                   "aqui muda o PL da subordinada.")

# =====================================================================================
elif menu == "📦 Estoque":
    ab = r["em_carteira"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Face em carteira", brl(float(ab["valor_face"].sum()) if len(ab) else 0))
    m2.metric("Valor na curva", brl(G["dc"]))
    m3.metric("Títulos em carteira", len(ab))
    m4.metric("Vencidos e não pagos", P["vencidos_qtd"], brl(P["vencidos_face"]) if P["vencidos_qtd"] else None,
              delta_color="inverse")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Concentração por sacado")
        conc = P["concentracao"]
        if len(conc):
            st.altair_chart(alt.Chart(conc).mark_bar().encode(
                x=alt.X("pct_dc:Q", axis=alt.Axis(format="%"), title="% da carteira de DC"),
                y=alt.Y("sacado:N", sort="-x", title=None),
                tooltip=[alt.Tooltip("sacado", title="Sacado"), alt.Tooltip("valor_face:Q", format=",.2f", title="Face"),
                         alt.Tooltip("pct_dc:Q", format=".1%", title="% DC")]).properties(height=260),
                width="stretch")
            tabela(conc.rename(columns={"sacado": "Sacado", "valor_face": "Face", "pct_dc": "% DC", "pct_pl": "% PL"}),
                   moeda=("Face",), pcts=("% DC", "% PL"))
        else:
            st.info("Nenhum título em carteira.")
    with c2:
        st.subheader("Vencimentos")
        v = P["vencimentos"]
        if len(v):
            vv = v.assign(data=v["vencimento"].apply(data_br))
            st.altair_chart(alt.Chart(vv).mark_bar().encode(
                x=alt.X("data:N", sort=None, title=None), y=alt.Y("face:Q", title="Face a receber"),
                tooltip=["data", alt.Tooltip("face:Q", format=",.2f"), "titulos", "dias"]).properties(height=260),
                width="stretch")
            tabela(v.rename(columns={"vencimento": "Vencimento", "face": "Face", "titulos": "Títulos", "dias": "Dias"}),
                   moeda=("Face",), datas=("Vencimento",))

    st.subheader("Enquadramento")
    e1, e2 = st.columns([2, 1])
    with e1:
        st.progress(min(P["dc_pl"] / P["minimo"], 1.0) if P["minimo"] else 0.0,
                    text=f"DC/PL {pct(P['dc_pl'])} de um mínimo de {pct(P['minimo'])}")
        st.markdown(f"**Situação:** {P['situacao']}" +
                    (f" · data limite {data_br(P['limite'])}" if P["limite"] else ""))
    with e2:
        with st.expander("Ajustar parâmetros"):
            ini = st.date_input("Início das atividades da Classe", value=dados["inicio_classe"], format="DD/MM/YYYY")
            mini = st.number_input("Mínimo DC/PL (%)", value=float(dados["minimo_dc"]) * 100, step=5.0)
            prazo = st.number_input("Prazo para enquadramento (dias)", value=int(dados["prazo_enquadramento"]), step=30)
            novo = {"inicio": ini, "minimo": mini / 100, "prazo": int(prazo)}
            if novo != {k: ep.get(k) for k in novo} and (ini != dados["inicio_classe"] or mini / 100 != dados["minimo_dc"]
                                                         or prazo != dados["prazo_enquadramento"]):
                st.session_state["ed_painel"] = novo
                st.rerun()

    st.subheader("Títulos")
    todos = st.toggle("Mostrar também os liquidados", value=False)
    t = r["titulos"] if todos else ab
    if len(t):
        tabela(t[["lote", "numero", "sacado", "data_aquisicao", "vencimento", "valor_face", "preco_pago", "curva",
                  "status", "data_baixa", "valor_recebido", "atraso"]].rename(columns={
            "lote": "Lote", "numero": "Nº", "sacado": "Sacado", "data_aquisicao": "Aquisição", "vencimento": "Vencimento",
            "valor_face": "Face", "preco_pago": "Preço pago", "curva": "Curva", "status": "Status",
            "data_baixa": "Baixa", "valor_recebido": "Recebido", "atraso": "Atraso (dias)"}),
            moeda=("Face", "Preço pago", "Curva", "Recebido"), datas=("Aquisição", "Vencimento", "Baixa"))

    st.subheader("Estoque: gestora × administradora")
    eadm, falta = r["estoque_adm"], r["faltam_na_adm"]
    if not len(eadm):
        st.warning("Sem estoque analítico da ADM: envie o arquivo na categoria 3.")
    else:
        div = eadm[eadm["situacao_conf"] != "OK"]
        if not len(div) and not len(falta):
            st.success(f"Os {len(eadm)} títulos do estoque da ADM batem com a gestora em face e curva.")
        if len(div):
            st.error(f"{len(div)} título(s) do estoque da ADM com divergência")
            tabela(div[["numero_documento", "sacado", "valor_nominal", "valor_face", "valor_presente", "curva",
                        "situacao_conf"]].rename(columns={
                "numero_documento": "Nº", "sacado": "Sacado", "valor_nominal": "Face ADM", "valor_face": "Face gestora",
                "valor_presente": "Curva ADM", "curva": "Curva gestora", "situacao_conf": "Situação"}),
                moeda=("Face ADM", "Face gestora", "Curva ADM", "Curva gestora"))
        if len(falta):
            st.error(f"{len(falta)} título(s) em carteira na gestora que não aparecem no estoque da ADM")
            tabela(falta[["lote", "numero", "sacado", "vencimento", "valor_face"]].rename(columns={
                "lote": "Lote", "numero": "Nº", "sacado": "Sacado", "vencimento": "Vencimento", "valor_face": "Face"}),
                moeda=("Face",), datas=("Vencimento",))

# =====================================================================================
elif menu == "💰 Caixa e despesas":
    st.subheader("Caixa")
    if a_classificar:
        st.warning(f"{a_classificar} lançamento(s) a classificar. Escolha na coluna **Classificação**.")
    else:
        st.success("Todos os lançamentos estão classificados.")
    so_pend = st.toggle("Mostrar só os lançamentos a classificar", value=bool(a_classificar))
    cx = pd.DataFrame(dados["caixa"])
    if len(cx):
        vis = cx[cx["classificacao"] == "A classificar"] if so_pend else cx
        ed = st.data_editor(
            vis[["id", "data", "historico", "contraparte", "entrada", "saida", "classificacao", "referencia"]],
            column_config={
                "id": "ID", "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                "historico": "Histórico", "contraparte": "Contraparte",
                "entrada": st.column_config.NumberColumn("Entrada", format="%.2f"),
                "saida": st.column_config.NumberColumn("Saída", format="%.2f"),
                "classificacao": st.column_config.SelectboxColumn("Classificação", options=CLASSES, required=True),
                "referencia": st.column_config.TextColumn("Referência")},
            disabled=["id", "data", "historico", "contraparte", "entrada", "saida"],
            hide_index=True, width="stretch", key=f"ed_cx_{so_pend}")
        mudou = False
        for _, row in ed.iterrows():
            orig = cx.loc[cx["id"] == row["id"]].iloc[0]
            ref = row["referencia"] or ""
            if row["classificacao"] != orig["classificacao"] or ref != (orig["referencia"] or ""):
                if row["classificacao"] != "A classificar" and ref.startswith("Sugestão:"):
                    ref = ""
                st.session_state["ed_caixa"][row["id"]] = (row["classificacao"], ref)
                mudou = True
        if mudou:
            st.rerun()

        pc = r["por_classe"].reset_index().rename(columns={"classificacao": "Classificação", "entrada": "Entradas",
                                                           "saida": "Saídas"})
        pc["Líquido"] = pc["Entradas"] - pc["Saídas"]
        st.markdown("**Resumo por classificação (até a data-base)**")
        tabela(pc, moeda=("Entradas", "Saídas", "Líquido"))

    st.subheader("Despesas a pagar")
    st.caption("Quando discordar da ADM, corrija o **Valor gestora** e explique na **Observação**.")
    dp = pd.DataFrame(dados["despesas"])
    if len(dp):
        dp["diferenca"] = dp["valor_gestora"] - dp["valor"]
        ed = st.data_editor(
            dp[["descricao", "lancamento", "vencimento", "valor", "valor_gestora", "diferenca", "observacao"]],
            column_config={
                "descricao": "Descrição", "lancamento": st.column_config.DateColumn("Lançamento", format="DD/MM/YYYY"),
                "vencimento": st.column_config.DateColumn("Vencimento", format="DD/MM/YYYY"),
                "valor": st.column_config.NumberColumn("Valor ADM", format="%.2f"),
                "valor_gestora": st.column_config.NumberColumn("Valor gestora", format="%.2f", required=True),
                "diferenca": st.column_config.NumberColumn("Diferença", format="%.2f"),
                "observacao": st.column_config.TextColumn("Observação", width="large")},
            disabled=["descricao", "lancamento", "vencimento", "valor", "diferenca"],
            hide_index=True, width="stretch", key="ed_desp_tab")
        mudou = False
        for _, row in ed.iterrows():
            orig = dp.loc[dp["descricao"] == row["descricao"]].iloc[0]
            if abs(row["valor_gestora"] - orig["valor_gestora"]) > 0.001 or (row["observacao"] or "") != (orig["observacao"] or ""):
                st.session_state["ed_desp"][row["descricao"]] = (float(row["valor_gestora"]), row["observacao"] or "")
                mudou = True
        if mudou:
            st.rerun()
        st.markdown(f"**Total:** ADM {brl(A['despesas'])} · gestora {brl(G['despesas'])} · "
                    f"diferença {brl(G['despesas'] - A['despesas'])}")

# =====================================================================================
elif menu == "✅ Conferências":
    for grupo, titulo in (("Estoque", "Estoque na data-base"), ("Caixa", "Caixa × títulos"), ("ADM", "Relatório da ADM")):
        st.subheader(titulo)
        g = conf[conf["Grupo"] == grupo].copy()
        g["Situação"] = g["OK"].map({True: "✅ OK", False: "❌ Diferente"}).fillna("⬜ sem arquivo")
        inteiro = g["Checagem"].str.contains("Quantidade|divergentes")
        for c in ("Gestora", "Referência", "Diferença"):
            g[c] = [num(v, 0) if i else brl(v) for v, i in zip(g[c], inteiro)]
        st.dataframe(g[["Checagem", "Gestora", "Referência", "Diferença", "Situação"]], hide_index=True,
                     width="stretch")

    st.subheader("Lotes (pelos CNABs de remessa)")
    tit = r["titulos"]
    if len(tit):
        lot = tit.groupby("lote").agg(cedente=("cedente", "first"), data=("data_aquisicao", "min"),
                                      titulos=("chave", "count"), face=("valor_face", "sum"),
                                      preco=("preco_pago", "sum")).reset_index()
        lot["desagio"] = lot["face"] - lot["preco"]
        tabela(lot.sort_values("data").rename(columns={"lote": "Lote", "cedente": "Cedente", "data": "Aquisição",
                                                       "titulos": "Títulos", "face": "Face", "preco": "Preço pago",
                                                       "desagio": "Deságio"}),
               moeda=("Face", "Preço pago", "Deságio"), datas=("Aquisição",))
        st.caption("Confira estes totais com o quadro resumo de cada termo de cessão.")

    st.subheader("Alertas do processamento")
    if dados["alertas"]:
        for a in dados["alertas"]:
            st.warning(a)
    else:
        st.success("Nenhum alerta.")

# =====================================================================================
elif menu == "📥 Exportar":
    st.subheader("Baixar a planilha")
    st.markdown("""
A planilha traz todas as abas (Carteira, Carteira x ADM, Estoque, Estoque x ADM, Caixa, Despesas,
Conferencia, Painel e Feriados), com as contas como fórmulas para auditoria.

**Ela também é a memória da próxima sessão:** na próxima vez, envie esta planilha na categoria 5
para recuperar as classificações do caixa, as despesas ajustadas e os parâmetros de enquadramento.
""")
    st.download_button("📥 Baixar planilha", data=excel_bytes(), file_name=dados["arquivo_sugerido"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    st.markdown("**Correções feitas nesta sessão**")
    if st.session_state["ed_caixa"] or st.session_state["ed_desp"] or st.session_state["ed_painel"]:
        for i, (c, ref) in st.session_state["ed_caixa"].items():
            st.markdown(f"- Caixa, lançamento {i}: {c}" + (f" ({ref})" if ref else ""))
        for d_, (v, o) in st.session_state["ed_desp"].items():
            st.markdown(f"- Despesa {d_}: gestora {brl(v)}" + (f" — {o}" if o else ""))
        if st.session_state["ed_painel"]:
            st.markdown("- Parâmetros de enquadramento ajustados")
    else:
        st.caption("Nenhuma correção feita nesta sessão.")
