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
                               help="Composição de Carteira. Se enviar mais de uma, vale a mais recente.")
    up_est = st.file_uploader("3 · Estoque analítico da ADM (.csv)", type=["csv"], accept_multiple_files=True,
                              help="Estoque título a título, da mesma data da carteira.")
    up_ext = st.file_uploader("4 · Extrato da conta (.csv)", type=["csv"], accept_multiple_files=True,
                              help="Extrato da conta corrente do fundo. Pode enviar vários períodos.")
    st.markdown("---")
    up_ant = st.file_uploader("5 · Última planilha baixada (opcional)", type=["xlsx"],
                              help="A planilha que você baixou no uso anterior. Dela vêm as suas correções: "
                                   "classificações do caixa, despesas ajustadas e dados de enquadramento.")

menu = st.sidebar.radio("Navegação", ["🏠 Visão geral", "📦 Estoque", "💰 Caixa e despesas", "✅ Conferências", "📥 Exportar"])

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
    m1.metric("PL pela gestora", brl(G["pl"]), f"{brl(dif_pl)} vs ADM" if abs(dif_pl) >= 1 else "igual à ADM",
              delta_color="off" if abs(dif_pl) < 1 else "normal")
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
