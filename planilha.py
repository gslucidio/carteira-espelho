"""
Montagem da planilha Excel da carteira espelho.

Recebe os dados já organizados por gerar_carteira.py e escreve as abas.
Os dados entram como valores (texto azul, fundo amarelo quando editáveis);
as contas ficam como fórmulas (texto preto), para que tudo possa ser auditado no Excel.
"""
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

FONTE = "Arial"
MOEDA = "#,##0.00;(#,##0.00);-"
DATA = "dd/mm/yyyy"
PCT = "0.0%"
COTA = "#,##0.00000000"

_fino = Side(style="thin", color="BFBFBF")
BORDA = Border(top=_fino, bottom=_fino, left=_fino, right=_fino)
F_CAB = PatternFill("solid", fgColor="1F3864")
F_INP = PatternFill("solid", fgColor="FFF2CC")
F_TOT = PatternFill("solid", fgColor="D9E1F2")

CLASSES_CAIXA = ["Aporte de cotista", "Resgate de cotista", "Aplicação no fundo DI", "Resgate do fundo DI",
                 "Aquisição de DC", "Recebimento de DC", "Taxas e despesas", "Tarifa bancária",
                 "Outros", "A classificar"]


def fonte(**k):
    return Font(name=FONTE, size=k.pop("size", 10), **k)


AZUL, PRETO, VERDE, NEG = fonte(color="0000FF"), fonte(), fonte(color="008000"), fonte(bold=True)
CINZA = fonte(italic=True, color="595959")


def titulo(ws, texto, sub=None):
    ws["A1"] = texto
    ws["A1"].font = fonte(size=14, bold=True, color="1F3864")
    if sub:
        ws["A2"] = sub
        ws["A2"].font = CINZA


def cabecalho(ws, linha, colunas, col0=1):
    for i, nome in enumerate(colunas, col0):
        c = ws.cell(linha, i, nome)
        c.font = fonte(bold=True, color="FFFFFF")
        c.fill = F_CAB
        c.border = BORDA
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[linha].height = 30


def cel(ws, r, c, valor, fmt=None, estilo=PRETO, editavel=False, borda=True):
    x = ws.cell(r, c, valor)
    x.font = estilo
    if fmt:
        x.number_format = fmt
    if editavel:
        x.fill = F_INP
    if borda:
        x.border = BORDA
    return x


def larguras(ws, valores):
    for i, w in enumerate(valores, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w


def linha_total(ws, r, ncols, somar, r_ini, r_fim, rotulo="TOTAL"):
    ws.cell(r, 1, rotulo).font = NEG
    for col in range(1, ncols + 1):
        ws.cell(r, col).fill = F_TOT
    for col in somar:
        letra = ws.cell(1, col).column_letter
        c = ws.cell(r, col, f"=SUM({letra}{r_ini}:{letra}{r_fim})")
        c.font = NEG
        c.number_format = MOEDA


# =====================================================================================
def montar(d, destino):
    """d: dicionário montado em gerar_carteira.py"""
    wb = Workbook()
    wb.remove(wb.active)
    ws_cart = wb.create_sheet("Carteira")
    ws_cx = wb.create_sheet("Carteira x ADM")
    ws_dre = wb.create_sheet("DRE")
    classes = d.get("classes") or []
    ws_pas = wb.create_sheet("Passivo") if classes else None
    ws_est = wb.create_sheet("Estoque")
    ws_ea = wb.create_sheet("Estoque x ADM")
    ws_caixa = wb.create_sheet("Caixa")
    ws_desp = wb.create_sheet("Despesas")
    ws_conf = wb.create_sheet("Conferencia")
    ws_pai = wb.create_sheet("Painel")
    ws_fer = wb.create_sheet("Feriados")

    # ---------------------------------------------------------------- Feriados
    titulo(ws_fer, "Feriados nacionais", "Usados na contagem de dias úteis da curva dos títulos.")
    cabecalho(ws_fer, 4, ["Data", "Feriado"])
    for i, (data, nome) in enumerate(d["feriados"], 5):
        cel(ws_fer, i, 1, data, DATA, AZUL)
        cel(ws_fer, i, 2, nome)
    larguras(ws_fer, [14, 32])
    wb.defined_names["Feriados"] = DefinedName("Feriados", attr_text=f"Feriados!$A$5:$A${4 + len(d['feriados']) + 50}")

    # ---------------------------------------------------------------- Estoque
    titulo(ws_est, "Estoque da gestora",
           "Uma linha por título, montada a partir dos CNABs de remessa e de baixa. Colunas pretas são calculadas. Nomes de lote alterados na coluna B são mantidos.")
    cols = ["Chave", "Lote", "Nº título", "Cedente", "Sacado", "CPF/CNPJ sacado", "Data aquisição", "Vencimento",
            "Valor de face", "Preço pago", "Status", "Data da baixa", "Valor recebido", "Em carteira na data-base",
            "DU totais", "DU decorridos", "Valor na curva", "Dias em atraso", "Ganho na liquidação", "Na ADM?",
            "Arquivo de remessa", "Arquivo de baixa"]
    cabecalho(ws_est, 4, cols)
    e0 = 5
    titulos = d["titulos"]
    for i, t in enumerate(titulos):
        r = e0 + i
        dados = [t["chave"], t["lote"], t["numero"], t["cedente"], t["sacado"], t["doc_sacado"], t["data_aquisicao"],
                 t["vencimento"], t["valor_face"], t["preco_pago"], t["status"], t["data_baixa"], t["valor_recebido"]]
        for j, v in enumerate(dados, 1):
            cel(ws_est, r, j, v, DATA if j in (7, 8, 12) else MOEDA if j in (9, 10, 13) else None, AZUL)
        formulas = [
            f'=IF(G{r}>DataBase,"Não",IF(OR(K{r}="Aberto",L{r}>DataBase),"Sim","Não"))',
            f"=NETWORKDAYS(G{r}+1,H{r},Feriados)",
            f"=IF(DataBase<=G{r},0,NETWORKDAYS(G{r}+1,MIN(DataBase,H{r}),Feriados))",
            f'=IF(N{r}="Sim",J{r}*(I{r}/J{r})^(P{r}/O{r}),0)',
            f'=IF(AND(N{r}="Sim",DataBase>H{r}),DataBase-H{r},0)',
            f'=IF(AND(K{r}="Liquidado",L{r}<=DataBase),M{r}-J{r},0)',
            f"=IF(N{r}=\"Sim\",IF(COUNTIF('Estoque x ADM'!$A:$A,A{r})>0,\"OK\",\"FALTA NA ADM\"),\"\")",
        ]
        for j, f in enumerate(formulas, 14):
            cel(ws_est, r, j, f, MOEDA if j in (17, 19) else "0" if j in (15, 16, 18) else None)
        cel(ws_est, r, 21, t["arquivo_remessa"], estilo=AZUL)
        cel(ws_est, r, 22, t["arquivo_baixa"], estilo=AZUL)
    e1 = e0 + max(len(titulos), 1) - 1          # pelo menos uma linha, para as fórmulas não quebrarem
    linha_total(ws_est, e1 + 1, 22, [9, 10, 13, 17, 19], e0, e1)
    dv = DataValidation(type="list", formula1='"Aberto,Liquidado"')
    ws_est.add_data_validation(dv)
    dv.add(f"K{e0}:K{e1 + 500}")
    larguras(ws_est, [20, 10, 9, 26, 32, 20, 12, 12, 13, 13, 10, 12, 13, 11, 8, 10, 14, 9, 13, 14, 36, 36])
    ws_est.freeze_panes = "D5"
    ws_est.auto_filter.ref = f"A4:V{e1}"

    def EST(col):
        return f"Estoque!${col}${e0}:${col}${e1}"

    # ---------------------------------------------------------------- Estoque x ADM
    eadm = d["estoque_adm"]
    titulo(ws_ea, "Estoque — gestora × administradora",
           "Colunas A a K: estoque analítico da ADM. Colunas L a P: comparação com a aba Estoque, calculada na data do estoque (H3).")
    ws_ea["A3"] = "Arquivo da ADM:"
    ws_ea["A3"].font = NEG
    ws_ea["C3"] = d["estoque_adm_arquivo"]
    ws_ea["F3"] = "Data de referência:"
    ws_ea["F3"].font = NEG
    cel(ws_ea, 3, 8, d["estoque_adm_data"], DATA, AZUL, borda=False)
    cabecalho(ws_ea, 5, ["Chave", "Nº documento", "Sacado", "CPF/CNPJ sacado", "Data aquisição", "Vencimento",
                         "Valor nominal", "Valor presente", "Valor aquisição", "PDD", "Situação",
                         "Face gestora", "Curva gestora", "Dif. face", "Dif. curva", "Conferência",
                         "Apoio: preço pago", "Apoio: aquisição", "Apoio: vencimento"])
    a0 = 6
    for i, t in enumerate(eadm):
        r = a0 + i
        dados = [t["chave"], t["numero_documento"], t["sacado"], t["doc_sacado"], t["data_aquisicao"], t["vencimento"],
                 t["valor_nominal"], t["valor_presente"], t["valor_aquisicao"], t["valor_pdd"], t["situacao"]]
        for j, v in enumerate(dados, 1):
            cel(ws_ea, r, j, v, DATA if j in (5, 6) else MOEDA if 7 <= j <= 10 else None, AZUL)
        # títulos em carteira NA DATA DO ESTOQUE (H3): adquiridos até lá e ainda não baixados
        def soma(col):
            return (f'(SUMIFS({EST(col)},{EST("A")},$A{r},{EST("G")},"<="&$H$3,{EST("K")},"Aberto")'
                    f'+SUMIFS({EST(col)},{EST("A")},$A{r},{EST("G")},"<="&$H$3,{EST("K")},"Liquidado",{EST("L")},">"&$H$3))')
        cel(ws_ea, r, 12, "=" + soma("I"), MOEDA)
        cel(ws_ea, r, 17, "=" + soma("J"), MOEDA)
        cel(ws_ea, r, 18, f'=SUMIFS({EST("G")},{EST("A")},$A{r})', DATA)
        cel(ws_ea, r, 19, f'=SUMIFS({EST("H")},{EST("A")},$A{r})', DATA)
        cel(ws_ea, r, 13, f'=IF(L{r}=0,0,Q{r}*(L{r}/Q{r})^(NETWORKDAYS(R{r}+1,MIN($H$3,S{r}),Feriados)'
                          f'/NETWORKDAYS(R{r}+1,S{r},Feriados)))', MOEDA)
        cel(ws_ea, r, 14, f"=L{r}-G{r}", MOEDA)
        cel(ws_ea, r, 15, f"=M{r}-H{r}", MOEDA)
        cel(ws_ea, r, 16, f'=IF(L{r}=0,"NÃO ESTÁ EM CARTEIRA NA GESTORA",IF(AND(ABS(N{r})<0.01,ABS(O{r})<0.05),"OK","DIFERENTE"))')
    a1 = a0 + max(len(eadm), 1) - 1
    linha_total(ws_ea, a1 + 1, 16, [7, 8, 9, 10, 12, 13, 14, 15], a0, a1)
    ws_ea.cell(a1 + 3, 1, "Títulos com problema:").font = NEG
    cel(ws_ea, a1 + 3, 3, f'=COUNTIF(P{a0}:P{a1},"<>OK")', borda=False)
    larguras(ws_ea, [20, 11, 32, 20, 12, 12, 13, 13, 13, 9, 10, 13, 13, 10, 10, 30, 13, 12, 12])
    ws_ea.freeze_panes = "C6"
    r_prob_ea = a1 + 3

    # ---------------------------------------------------------------- Caixa
    cx = d["caixa"]
    titulo(ws_caixa, f"Caixa — {d['conta']}",
           "Uma linha por lançamento do extrato. Saldo = saldo anterior + entrada − saída. "
           "Classifique os lançamentos amarelos pela lista da coluna J: a escolha é mantida nas próximas gerações.")
    cabecalho(ws_caixa, 4, ["ID", "Data", "Histórico", "Contraparte", "Entrada (+)", "Saída (−)", "Saldo calculado",
                            "Saldo do banco", "Diferença", "Classificação", "Referência", "Tipo de despesa"])
    c0 = 5
    for i, m in enumerate(cx):
        r = c0 + i
        dados = [m["id"], m["data"], m["historico"], m["contraparte"], m["entrada"] or None, m["saida"] or None]
        for j, v in enumerate(dados, 1):
            cel(ws_caixa, r, j, v, DATA if j == 2 else MOEDA if j in (5, 6) else None, AZUL)
        cel(ws_caixa, r, 7, f"=E{r}-F{r}" if i == 0 else f"=G{r - 1}+E{r}-F{r}", MOEDA)
        cel(ws_caixa, r, 8, m["saldo_banco"], MOEDA, AZUL)
        cel(ws_caixa, r, 9, f"=G{r}-H{r}", MOEDA)
        cel(ws_caixa, r, 10, m["classificacao"], estilo=AZUL,
            editavel=m["classificacao"] == "A classificar")
        cel(ws_caixa, r, 11, m["referencia"], estilo=AZUL)
        cel(ws_caixa, r, 12, m.get("tipo") or None, estilo=AZUL)
    c1 = c0 + max(len(cx), 1) - 1
    linha_total(ws_caixa, c1 + 1, 12, [5, 6, 9], c0, c1)
    dvc = DataValidation(type="list", formula1='"' + ",".join(CLASSES_CAIXA) + '"')
    ws_caixa.add_data_validation(dvc)
    dvc.add(f"J{c0}:J{c1 + 500}")
    sr = c1 + 4
    ws_caixa.cell(sr, 1, "Resumo por classificação (até a data-base)").font = NEG
    cabecalho(ws_caixa, sr + 1, ["Classificação", "Entradas", "Saídas", "Líquido"])
    CB, CE, CF, CJ = (f"${l}${c0}:${l}${c1 + 500}" for l in "BEFJ")
    linha_classe = {}
    for i, k in enumerate(CLASSES_CAIXA):
        r = sr + 2 + i
        linha_classe[k] = r
        cel(ws_caixa, r, 1, k)
        cel(ws_caixa, r, 2, f'=SUMIFS({CE},{CJ},A{r},{CB},"<="&DataBase)', MOEDA)
        cel(ws_caixa, r, 3, f'=SUMIFS({CF},{CJ},A{r},{CB},"<="&DataBase)', MOEDA)
        cel(ws_caixa, r, 4, f"=B{r}-C{r}", MOEDA)
    larguras(ws_caixa, [10, 12, 34, 42, 14, 14, 14, 14, 11, 22, 46, 24])
    ws_caixa.freeze_panes = "C5"

    def CLS(classe, col):
        return f"Caixa!${col}${linha_classe[classe]}"

    # ---------------------------------------------------------------- Despesas
    desp = d["despesas"]
    titulo(ws_desp, "Despesas a pagar",
           "Valores a liquidar da carteira da ADM. Quando discordar, corrija a coluna Gestora (amarela) e explique na Observação: a correção é mantida nas próximas gerações.")
    cabecalho(ws_desp, 4, ["Descrição", "Lançamento", "Vencimento", "Valor ADM", "Valor gestora", "Diferença", "Observação",
                           "Tipo de despesa"])
    for i, x in enumerate(desp, 5):
        cel(ws_desp, i, 1, x["descricao"], estilo=AZUL)
        cel(ws_desp, i, 2, x["lancamento"], DATA, AZUL)
        cel(ws_desp, i, 3, x["vencimento"], DATA, AZUL)
        cel(ws_desp, i, 4, x["valor"], MOEDA, AZUL)
        cel(ws_desp, i, 5, x["valor_gestora"], MOEDA, AZUL, editavel=True)
        cel(ws_desp, i, 6, f"=E{i}-D{i}", MOEDA)
        cel(ws_desp, i, 7, x["observacao"]).alignment = Alignment(wrap_text=True, vertical="top")
        cel(ws_desp, i, 8, x.get("tipo"), estilo=AZUL)
    d1 = 4 + max(len(desp), 1)
    r_tot_desp = d1 + 1
    linha_total(ws_desp, r_tot_desp, 8, [4, 5, 6], 5, d1)
    larguras(ws_desp, [40, 12, 12, 13, 13, 12, 70, 24])

    # ---------------------------------------------------------------- Carteira (gestora)
    w = ws_cart
    titulo(w, f"Carteira da gestora — {d['fundo']}", "Texto azul em fundo amarelo = digitado. Texto preto = fórmula.")
    w["A4"] = "Data-base"
    w["A4"].font = NEG
    cel(w, 4, 2, d["data_base"], DATA, AZUL, editavel=True)
    w["C4"] = "Todas as abas calculam nesta data."
    w["C4"].font = CINZA
    wb.defined_names["DataBase"] = DefinedName("DataBase", attr_text="Carteira!$B$4")
    fundos = d["fundos"]
    k = 2 if classes else 0                    # linhas extras para o passivo das classes
    r_plc = 13 if classes else 11              # PL da classe principal (subordinada ou cota única)
    r_cotas, r_cota = 13 + k, 14 + k
    rf0 = 18 + k                               # tabela de fundos DI
    rf1 = rf0 + max(len(fundos), 1) - 1
    cabecalho(w, 6, ["Linha", "Valor", "% do PL"])
    linhas = [("(+) Direitos creditórios (curva)", f"=SUM({EST('Q')})"),
              ("(+) Fundos de liquidez", f"=SUM(D{rf0}:D{rf1})"),
              ("(+) Conta corrente", f'=SUMIFS(Caixa!{CE},Caixa!{CB},"<="&DataBase)-SUMIFS(Caixa!{CF},Caixa!{CB},"<="&DataBase)'),
              ("(−) Despesas a pagar", f"=-Despesas!E{r_tot_desp}")]
    for i, (rot, f) in enumerate(linhas, 7):
        cel(w, i, 1, rot)
        cel(w, i, 2, f, MOEDA, VERDE)
        cel(w, i, 3, f"=B{i}/$B$11", PCT)
    rotulo_pl = "(=) Patrimônio total do fundo" if classes else "(=) Patrimônio líquido"
    for col, v in ((1, rotulo_pl), (2, "=SUM(B7:B10)"), (3, "=B11/$B$11")):
        c = cel(w, 11, col, v, MOEDA if col == 2 else PCT if col == 3 else None, NEG)
        c.fill = F_TOT
    if classes:
        cel(w, 12, 1, "(−) Cotas das demais classes (aba Passivo)")
        cel(w, 12, 2, "=-Passivo!E{}".format(5 + len(classes)), MOEDA, VERDE)
        cel(w, 12, 3, "=B12/$B$11", PCT)
        for col, v in ((1, "(=) PL da subordinada"), (2, "=B11+B12"), (3, "=B13/$B$11")):
            c = cel(w, 13, col, v, MOEDA if col == 2 else PCT if col == 3 else None, NEG)
            c.fill = F_TOT
    w.cell(r_cotas, 1, "Quantidade de cotas" + (" da subordinada" if classes else ""))
    cel(w, r_cotas, 2, d["quantidade_cotas"], COTA, AZUL, editavel=True, borda=False)
    w.cell(r_cota, 1, "Valor da cota" + (" da subordinada" if classes else ""))
    cel(w, r_cota, 2, f"=B{r_plc}/B{r_cotas}", COTA, NEG, borda=False)
    w.cell(16 + k, 1, "Fundos de liquidez (quantidade × cota)").font = NEG
    cabecalho(w, 17 + k, ["Fundo", "Quantidade de cotas", "Valor da cota", "Valor"])
    for i, f_ in enumerate(fundos or [dict(descricao="", quantidade=0, valor_cota=0)]):
        r = rf0 + i
        cel(w, r, 1, f_["descricao"], estilo=AZUL)
        cel(w, r, 2, f_["quantidade"], COTA, AZUL, editavel=True)
        cel(w, r, 3, f_["valor_cota"], COTA, AZUL, editavel=True)
        cel(w, r, 4, f"=B{r}*C{r}", MOEDA)
    r = rf1 + 1
    w.cell(r, 1, "Aplicado líquido pelo extrato (aplicações − resgates)")
    cel(w, r, 2, f'={CLS("Aplicação no fundo DI", "C")}-{CLS("Resgate do fundo DI", "B")}', MOEDA, VERDE, borda=False)
    r_rend = r + 1
    w.cell(r + 1, 1, "Rendimento acumulado implícito nos fundos")
    cel(w, r + 1, 2, f"=B8-B{r}", MOEDA, borda=False)
    p = r + 3
    w.cell(p, 1, "Passivo — cotistas").font = NEG
    itens = [("Aportes de cotistas (extrato)", f'={CLS("Aporte de cotista", "B")}', VERDE),
             ("Resgates e amortizações pagos (extrato)", f'={CLS("Resgate de cotista", "C")}', VERDE),
             ("Capital líquido investido pelos cotistas", f"=B{p + 1}-B{p + 2}", PRETO),
             ("Resultado acumulado do fundo (PL − capital líquido)", f"=B11-B{p + 3}", NEG)]
    for i, (rot, f, est) in enumerate(itens, p + 1):
        w.cell(i, 1, rot)
        cel(w, i, 2, f, MOEDA, est, borda=False)
    w.cell(p + 5, 1, "Resultado acumulado em % do capital")
    cel(w, p + 5, 2, f"=IF(B{p + 3}=0,0,B{p + 4}/B{p + 3})", PCT, borda=False)
    larguras(w, [52, 20, 18, 16])

    r_resultado = p + 4

    # ---------------------------------------------------------------- DRE
    w = ws_dre
    titulo(w, "DRE — acumulado desde o início",
           "Regime de competência até a data-base: receitas apropriadas e despesas pagas (aba Caixa) ou a pagar (aba Despesas).")
    cabecalho(w, 4, ["Linha", "Pagas", "A pagar", "Valor"])
    CL, CJ2 = f"$L${c0}:$L${c1 + 500}", f"$J${c0}:$J${c1 + 500}"
    rr = 5
    cel(w, rr, 1, "RECEITAS", estilo=NEG)
    rec = [("Deságio realizado em títulos liquidados", f"=SUM({EST('S')})"),
           ("Deságio apropriado dos títulos em carteira", f'=SUMIFS({EST("Q")},{EST("N")},"Sim")-SUMIFS({EST("J")},{EST("N")},"Sim")'),
           ("Rendimento dos fundos de liquidez", f"=Carteira!B{r_rend}"),
           ("Outros movimentos / não classificados (líquido)",
            f'=SUMIFS(Caixa!{CE},Caixa!{CJ2},"Outros",Caixa!{CB},"<="&DataBase)-SUMIFS(Caixa!{CF},Caixa!{CJ2},"Outros",Caixa!{CB},"<="&DataBase)'
            f'+SUMIFS(Caixa!{CE},Caixa!{CJ2},"A classificar",Caixa!{CB},"<="&DataBase)-SUMIFS(Caixa!{CF},Caixa!{CJ2},"A classificar",Caixa!{CB},"<="&DataBase)')]
    r_rec0 = rr + 1
    for i, (rot, f) in enumerate(rec, r_rec0):
        cel(w, i, 1, "   " + rot)
        cel(w, i, 4, f, MOEDA)
    r_rec1 = r_rec0 + len(rec) - 1
    r_trec = r_rec1 + 1
    for col, v in ((1, "Total de receitas"), (4, f"=SUM(D{r_rec0}:D{r_rec1})")):
        c = cel(w, r_trec, col, v, MOEDA if col == 4 else None, NEG)
        c.fill = F_TOT
    rr = r_trec + 2
    cel(w, rr, 1, "DESPESAS", estilo=NEG)
    r_d0 = rr + 1
    for i, tipo in enumerate(d["tipos_despesa"], r_d0):
        cel(w, i, 1, "   " + tipo)
        cel(w, i, 2, f'=SUMIFS(Caixa!{CF},Caixa!{CL},"{tipo}",Caixa!{CB},"<="&DataBase)'
                     f'-SUMIFS(Caixa!{CE},Caixa!{CL},"{tipo}",Caixa!{CB},"<="&DataBase)', MOEDA)
        cel(w, i, 3, f'=SUMIFS(Despesas!$E$5:$E${d1},Despesas!$H$5:$H${d1},"{tipo}")', MOEDA)
        cel(w, i, 4, f"=-(B{i}+C{i})", MOEDA)
    r_d1 = r_d0 + len(d["tipos_despesa"]) - 1
    r_tdesp = r_d1 + 1
    for col, v in ((1, "Total de despesas"), (2, f"=SUM(B{r_d0}:B{r_d1})"), (3, f"=SUM(C{r_d0}:C{r_d1})"),
                   (4, f"=SUM(D{r_d0}:D{r_d1})")):
        c = cel(w, r_tdesp, col, v, MOEDA if col > 1 else None, NEG)
        c.fill = F_TOT
    r_res = r_tdesp + 2
    for col, v in ((1, "RESULTADO"), (4, f"=D{r_trec}+D{r_tdesp}")):
        c = cel(w, r_res, col, v, MOEDA if col == 4 else None, fonte(bold=True, size=11))
        c.fill = F_TOT
    cel(w, r_res + 2, 1, "Conferência: resultado pela carteira (PL − capital líquido)")
    cel(w, r_res + 2, 4, f"=Carteira!B{r_resultado}", MOEDA)
    cel(w, r_res + 3, 1, "Diferença (deve ser zero)")
    cel(w, r_res + 3, 4, f"=D{r_res}-D{r_res + 2}", MOEDA)
    w.cell(r_res + 5, 1, "Despesas pagas por tipo vêm da coluna 'Tipo de despesa' da aba Caixa; a pagar, da aba Despesas. "
                         "Pagamentos à ID CTVM ficam em 'Administração e custódia', pois o extrato não separa as duas.").font = CINZA
    larguras(w, [52, 16, 16, 18])

    # ---------------------------------------------------------------- Passivo (fundos com várias classes)
    if classes:
        w = ws_pas
        titulo(w, "Passivo — classes de cotas",
               "Valores das carteiras de cada classe enviadas pela ADM, comparados com o passivo registrado na carteira principal.")
        cabecalho(w, 4, ["Classe", "Benchmark", "Cotas", "Valor da cota", "PL na carteira da classe",
                         "Passivo na carteira principal", "Diferença", "Capital aplicado", "Rent. desde o início",
                         "Rent. no mês", "% do CDI no mês", "% do PL total"])
        for i, cl in enumerate(classes, 5):
            rt = cl["rentabilidades"]
            vals = [cl["nome"], cl["indice"], cl["cotas"], cl["cota"],
                    cl["pl"] if cl["pl"] is not None else cl["passivo_principal"], cl["passivo_principal"], None,
                    cl["aplicado"], rt.get("Carteira", {}).get("Inicial"), rt.get("Carteira", {}).get("Mês"),
                    rt.get("(%) CDI", {}).get("Mês")]
            fmts = [None, None, "#,##0.0000", "#,##0.000000", MOEDA, MOEDA, MOEDA, MOEDA, "0.00%", "0.00%", "0.0%"]
            for j, (v, fm) in enumerate(zip(vals, fmts), 1):
                if j == 7:
                    cel(w, i, 7, f"=F{i}-E{i}", MOEDA)
                else:
                    cel(w, i, j, v, fm, AZUL)
            cel(w, i, 12, f"=E{i}/Carteira!$B$11", PCT)
        tp = 5 + len(classes)
        for col, v in ((1, "Total das demais classes"), (5, f"=SUM(E5:E{tp - 1})"), (6, f"=SUM(F5:F{tp - 1})"),
                       (7, f"=SUM(G5:G{tp - 1})"), (8, f"=SUM(H5:H{tp - 1})"), (12, f"=SUM(L5:L{tp - 1})")):
            c = cel(w, tp, col, v, MOEDA if col in (5, 6, 7, 8) else PCT if col == 12 else None, NEG)
        for col in range(1, 13):
            w.cell(tp, col).fill = F_TOT
        rs = tp + 1
        rp = d["adm"].get("rentabilidades", {})
        cel(w, rs, 1, "Subordinada (carteira principal)", estilo=NEG)
        cel(w, rs, 2, "residual")
        cel(w, rs, 3, f"=Carteira!B{r_cotas}", "#,##0.0000")
        cel(w, rs, 4, f"=Carteira!B{r_cota}", "#,##0.000000")
        cel(w, rs, 5, "=Carteira!B13", MOEDA)
        cel(w, rs, 9, rp.get("Carteira", {}).get("Inicial"), "0.00%", AZUL)
        cel(w, rs, 10, rp.get("Carteira", {}).get("Mês"), "0.00%", AZUL)
        cel(w, rs, 11, rp.get("(%) CDI", {}).get("Mês"), "0.0%", AZUL)
        cel(w, rs, 12, f"=E{rs}/Carteira!$B$11", PCT)
        w.cell(rs + 2, 1, "Subordinação: % do patrimônio total que fica abaixo de cada classe").font = NEG
        cel(w, rs + 3, 1, "Subordinada / patrimônio total")
        cel(w, rs + 3, 2, f"=L{rs}", PCT)
        w.cell(rs + 4, 1, "Compare com os índices mínimos do regulamento. Rentabilidades de cada classe vêm das "
                          "carteiras da ADM (texto azul).").font = CINZA
        rdv = rs + 6
        divs = [(cl["nome"], x) for cl in classes for x in cl["divergencias"]]
        if divs:
            w.cell(rdv, 1, "Aplicações registradas de forma diferente nas duas carteiras").font = fonte(bold=True, color="C00000")
            cabecalho(w, rdv + 1, ["Classe", "Aplicação", "Valor aplicado", "Cotas (principal)", "Cotas (classe)",
                                   "Valor (principal)", "Valor (classe)", "Diferença"])
            for i, (nome, x) in enumerate(divs, rdv + 2):
                for j, (v, fm) in enumerate(zip([nome, x["data"], x["valor_aplicado"], x["cotas_principal"], x["cotas_classe"],
                                                 x["valor_principal"], x["valor_classe"]],
                                                [None, DATA, MOEDA, "#,##0.000000", "#,##0.000000", MOEDA, MOEDA]), 1):
                    cel(w, i, j, v, fm, AZUL)
                cel(w, i, 8, f"=F{i}-G{i}", MOEDA)
        larguras(w, [34, 12, 14, 14, 18, 20, 13, 16, 12, 12, 12, 12])

    # ---------------------------------------------------------------- Carteira x ADM
    w = ws_cx
    adm = d["adm"]
    titulo(w, "Carteira — gestora × administradora", f"Carteira da ADM: {d['carteira_adm_arquivo']}")
    w["A4"] = "Data-base"
    w["A4"].font = NEG
    cel(w, 4, 2, "=DataBase", DATA, borda=False)
    w["D4"] = "Posição da carteira ADM:"
    w["D4"].font = NEG
    cel(w, 4, 5, adm["data_posicao"], DATA, AZUL, borda=False)
    cabecalho(w, 6, ["Linha", "Gestora", "Administradora", "Diferença", "Onde investigar"])
    dcs = sorted(adm["direitos_creditorios"], key=lambda x: -abs(x["valor_mercado"] or 0))
    dc_principal = dcs[0]["valor_mercado"] if dcs else 0
    dc_outros = sum(x["valor_mercado"] or 0 for x in dcs[1:])
    outros_nomes = ", ".join(str(x["descricao"]) for x in dcs[1:]) or "-"
    X = [("(+) Direitos creditórios (curva)", "=Carteira!B7", dc_principal, "Aba Estoque x ADM"),
         ("(+) Outros direitos creditórios na ADM", 0, dc_outros, f"Linhas da ADM: {outros_nomes}"),
         ("(+) Fundos de liquidez", "=Carteira!B8", sum(f_["valor"] or 0 for f_ in fundos), "Extrato dos fundos"),
         ("(+) Conta corrente", "=Carteira!B9", adm["conta_corrente"], "Aba Caixa"),
         ("(−) Despesas a pagar", "=Carteira!B10", f"=-Despesas!D{r_tot_desp}", "Aba Despesas")]
    for i, (rot, g, a, onde) in enumerate(X, 7):
        cel(w, i, 1, rot)
        cel(w, i, 2, g, MOEDA, VERDE if str(g).startswith("=") else AZUL)
        cel(w, i, 3, a, MOEDA, VERDE if str(a).startswith("=") else AZUL, editavel=not str(a).startswith("="))
        cel(w, i, 4, f"=B{i}-C{i}", MOEDA)
        cel(w, i, 5, onde, estilo=fonte(color="595959"))
    for col, v in enumerate([rotulo_pl, "=SUM(B7:B11)", "=SUM(C7:C11)", "=SUM(D7:D11)", ""], 1):
        c = cel(w, 12, col, v, MOEDA if col in (2, 3, 4) else None, NEG)
        c.fill = F_TOT
    xp = 12                                     # linha do PL da classe principal nesta aba
    if classes:
        cel(w, 13, 1, "(−) Cotas das demais classes")
        cel(w, 13, 2, "=Carteira!B12", MOEDA, VERDE)
        cel(w, 13, 3, -sum(cl["passivo_principal"] for cl in classes), MOEDA, AZUL, editavel=True)
        cel(w, 13, 4, "=B13-C13", MOEDA)
        cel(w, 13, 5, "Aba Passivo (gestora = PL de cada classe na sua carteira)", estilo=fonte(color="595959"))
        for col, v in enumerate(["(=) PL da subordinada", "=B12+B13", "=C12+C13", "=B14-C14", ""], 1):
            c = cel(w, 14, col, v, MOEDA if col in (2, 3, 4) else None, NEG)
            c.fill = F_TOT
        xp = 14
    xc = xp + 2
    cel(w, xc, 1, "Quantidade de cotas" + (" da subordinada" if classes else ""))
    cel(w, xc, 2, f"=Carteira!B{r_cotas}", COTA)
    cel(w, xc, 3, adm["quantidade_cotas"], COTA, AZUL, editavel=True)
    cel(w, xc, 4, f"=B{xc}-C{xc}", COTA)
    cel(w, xc + 1, 1, "Valor da cota" + (" da subordinada" if classes else ""))
    cel(w, xc + 1, 2, f"=B{xp}/B{xc}", COTA)
    cel(w, xc + 1, 3, f"=C{xp}/C{xc}", COTA)
    cel(w, xc + 1, 4, f"=B{xc + 1}-C{xc + 1}", COTA)
    xk = xc + 4
    w.cell(xk, 1, "Checagens internas do relatório da ADM").font = NEG
    cabecalho(w, xk + 1, ["Checagem", "Informado pela ADM", "Calculado das linhas", "Diferença", "Esperado"])
    for r, rot, inf, calc, fmt in ((xk + 2, "Patrimônio líquido" + (" da subordinada" if classes else ""), adm["pl"], f"=C{xp}", MOEDA),
                                  (xk + 3, "Valor da cota", adm["valor_cota"], f"=C{xc + 1}", COTA)):
        cel(w, r, 1, rot)
        cel(w, r, 2, inf, fmt, AZUL, editavel=True)
        cel(w, r, 3, calc, fmt)
        cel(w, r, 4, f"=C{r}-B{r}", fmt)
        cel(w, r, 5, "zero")
    w.cell(xk + 5, 1, "As checagens internas só confirmam que o relatório da ADM é coerente com ele mesmo. "
                      "A comparação com a gestora é a tabela de cima.").font = CINZA
    w.cell(xk + 6, 1, "Diferenças acima de R$ 1,00 merecem explicação. Diferença de centavos é arredondamento.").font = CINZA
    larguras(w, [46, 18, 18, 16, 44])

    # ---------------------------------------------------------------- Conferencia
    w = ws_conf
    titulo(w, "Conferências")
    w["A4"] = "1. Lotes (pelos CNABs de remessa)"
    w["A4"].font = NEG
    cabecalho(w, 5, ["Lote", "Cedente", "Data aquisição", "Qtd. títulos", "Valor de face", "Preço pago", "Deságio"])
    for i, lote in enumerate(d["lotes"], 6):
        cel(w, i, 1, lote["nome"], estilo=AZUL)
        cel(w, i, 2, lote["cedente"], estilo=AZUL)
        cel(w, i, 3, lote["data"], DATA, AZUL)
        cel(w, i, 4, f'=COUNTIF({EST("B")},A{i})')
        cel(w, i, 5, f'=SUMIF({EST("B")},A{i},{EST("I")})', MOEDA)
        cel(w, i, 6, f'=SUMIF({EST("B")},A{i},{EST("J")})', MOEDA)
        cel(w, i, 7, f"=E{i}-F{i}", MOEDA)
    r = 6 + len(d["lotes"]) + 1
    w.cell(r, 1, "Confira estes totais com o quadro resumo de cada termo de cessão.").font = CINZA
    r += 2
    w.cell(r, 1, "2. Estoque na data-base: gestora × administradora").font = NEG
    cabecalho(w, r + 1, ["Item", "Gestora", "Administradora", "Diferença", "Situação"])
    itens = [("Face em carteira × estoque analítico ADM (na data do estoque)", f"=SUM('Estoque x ADM'!L{a0}:L{a1})",
              f"=SUM('Estoque x ADM'!G{a0}:G{a1})", MOEDA, 0.01),
             ("Valor na curva × estoque analítico ADM (na data do estoque)", f"=SUM('Estoque x ADM'!M{a0}:M{a1})",
              f"=SUM('Estoque x ADM'!H{a0}:H{a1})", MOEDA, 1),
             ("Valor na curva × carteira ADM", "=Carteira!B7", f"='Carteira x ADM'!C7", MOEDA, 1),
             ("Quantidade de títulos em carteira", f'=COUNTIF({EST("N")},"Sim")', f"=COUNTA('Estoque x ADM'!A{a0}:A{a1})", "0", 0.5),
             ("Títulos divergentes (gestora / ADM)", f'=COUNTIF({EST("T")},"FALTA NA ADM")', f"='Estoque x ADM'!C{r_prob_ea}", "0", None)]
    for i, (rot, g, a, fmt, tol) in enumerate(itens, r + 2):
        cel(w, i, 1, rot)
        cel(w, i, 2, g, fmt)
        cel(w, i, 3, a, fmt)
        if tol is None:
            cel(w, i, 4, f"=B{i}+C{i}", fmt)
            cel(w, i, 5, f'=IF(D{i}=0,"OK","VER ABAS Estoque E Estoque x ADM")')
        else:
            cel(w, i, 4, f"=B{i}-C{i}", fmt)
            cel(w, i, 5, f'=IF(ABS(D{i})<{tol},"OK","DIFERENTE")')
    r = r + 2 + len(itens) + 1
    if d["estoque_adm_data"] and d["estoque_adm_data"] != d["data_base"]:
        w.cell(r, 1, "Atenção: o estoque analítico da ADM não é da mesma data da data-base.").font = fonte(bold=True, color="C00000")
    r += 2
    w.cell(r, 1, "3. Caixa × títulos (até a data-base)").font = NEG
    cabecalho(w, r + 1, ["Item", "Pela aba Estoque", "Pelo extrato", "Diferença", "Situação"])
    itens = [("Pago na compra de direitos creditórios", f'=SUMIFS({EST("J")},{EST("G")},"<="&DataBase)', f'={CLS("Aquisição de DC", "C")}'),
             ("Recebido na liquidação de títulos", f'=SUMIFS({EST("M")},{EST("L")},"<="&DataBase)', f'={CLS("Recebimento de DC", "B")}'),
             ("Saldo da conta corrente (calculado × banco)", f"=Caixa!G{c1}", f"=Caixa!H{c1}"),
             ("Saldo da conta corrente (gestora × carteira ADM)", "=Carteira!B9", "='Carteira x ADM'!C10")]
    for i, (rot, g, a) in enumerate(itens, r + 2):
        cel(w, i, 1, rot)
        cel(w, i, 2, g, MOEDA)
        cel(w, i, 3, a, MOEDA)
        cel(w, i, 4, f"=B{i}-C{i}", MOEDA)
        cel(w, i, 5, f'=IF(ABS(D{i})<0.01,"OK","DIFERENTE")')
    r = r + 2 + len(itens) + 1
    w.cell(r, 1, "Se a compra ou o recebimento não bater, há título sem dinheiro correspondente (ou dinheiro sem título).").font = CINZA
    r += 3
    w.cell(r, 1, "4. Alertas do processamento").font = NEG
    cabecalho(w, r + 1, ["Alerta"])
    for i, a in enumerate(d["alertas"] or ["Nenhum alerta."], r + 2):
        c = w.cell(i, 1, a)
        c.font = fonte(color="C00000") if d["alertas"] else PRETO
    r = r + 2 + max(len(d["alertas"]), 1) + 2
    w.cell(r, 1, "5. Correções manuais mantidas da planilha anterior").font = NEG
    w.cell(r + 1, 1, f"Lidas de: {d['memoria'] or 'nenhuma planilha anterior (primeira geração)'}").font = CINZA
    cabecalho(w, r + 2, ["Correção"])
    for i, a in enumerate(d["avisos"] or ["Nenhuma divergência de valor mantida."], r + 3):
        w.cell(i, 1, a)
    larguras(w, [60, 22, 18, 14, 18, 16, 14])

    # ---------------------------------------------------------------- Painel
    w = ws_pai
    titulo(w, "Painel de acompanhamento", "Considera apenas títulos em carteira na data-base.")
    w["A4"] = "Concentração por sacado"
    w["A4"].font = NEG
    cabecalho(w, 5, ["Sacado", "Face em carteira", "% da carteira de DC", "% do PL"])
    sacados = d["sacados"] or ["(sem títulos em carteira)"]
    tot = 6 + len(sacados)
    for i, s in enumerate(sacados, 6):
        cel(w, i, 1, s, estilo=AZUL)
        cel(w, i, 2, f'=SUMIFS({EST("I")},{EST("E")},A{i},{EST("N")},"Sim")', MOEDA)
        cel(w, i, 3, f"=IF($B${tot}=0,0,B{i}/$B${tot})", PCT)
        cel(w, i, 4, f"=B{i}/Carteira!$B$11", PCT)
    for col, v, fmt in ((1, "TOTAL", None), (2, f"=SUM(B6:B{tot - 1})", MOEDA), (3, f"=SUM(C6:C{tot - 1})", PCT),
                        (4, f"=SUM(D6:D{tot - 1})", PCT)):
        c = cel(w, tot, col, v, fmt, NEG)
        c.fill = F_TOT
    r = tot + 3
    w.cell(r, 1, "Vencimentos").font = NEG
    cabecalho(w, r + 1, ["Vencimento", "Face a receber", "Qtd. títulos", "Dias até o vencimento"])
    for i, v in enumerate(d["vencimentos"], r + 2):
        cel(w, i, 1, v, DATA, AZUL)
        cel(w, i, 2, f'=SUMIFS({EST("I")},{EST("H")},A{i},{EST("N")},"Sim")', MOEDA)
        cel(w, i, 3, f'=COUNTIFS({EST("H")},A{i},{EST("N")},"Sim")')
        cel(w, i, 4, f"=A{i}-DataBase")
    r = r + 2 + len(d["vencimentos"]) + 2
    w.cell(r, 1, "Atrasos").font = NEG
    cabecalho(w, r + 1, ["Item", "Valor"])
    for i, (rot, f, fmt) in enumerate([("Títulos vencidos e não pagos (qtd.)", f'=COUNTIF({EST("R")},">0")', "0"),
                                       ("Face vencida e não paga", f'=SUMIF({EST("R")},">0",{EST("I")})', MOEDA),
                                       ("Maior atraso (dias)", f'=MAX({EST("R")})', "0")], r + 2):
        cel(w, i, 1, rot)
        cel(w, i, 2, f, fmt)
    r = r + 7
    w.cell(r, 1, "Enquadramento").font = NEG
    cabecalho(w, r + 1, ["Item", "Valor"])
    e = r + 2
    cel(w, e, 1, "Direitos creditórios / PL")
    cel(w, e, 2, "=Carteira!B7/Carteira!B11", PCT)
    cel(w, e + 1, 1, "Mínimo exigido pelo regulamento")
    cel(w, e + 1, 2, d["minimo_dc"], PCT, AZUL, editavel=True)
    cel(w, e + 2, 1, "Data de início das atividades da Classe")
    cel(w, e + 2, 2, d["inicio_classe"], DATA, AZUL, editavel=True)
    cel(w, e + 3, 1, "Prazo para enquadramento (dias)")
    cel(w, e + 3, 2, d["prazo_enquadramento"], "0", AZUL, editavel=True)
    cel(w, e + 4, 1, "Data limite para estar enquadrado")
    cel(w, e + 4, 2, f'=IF(B{e + 2}="","informar início",B{e + 2}+B{e + 3})', DATA)
    cel(w, e + 5, 1, "Situação")
    cel(w, e + 5, 2, f'=IF(B{e}>=B{e + 1},"Enquadrado",IF(B{e + 2}="","Abaixo do mínimo",'
                     f'IF(DataBase>B{e + 4},"DESENQUADRADO","Abaixo do mínimo, ainda no prazo")))')
    larguras(w, [44, 18, 18, 20])

    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == "Carteira"
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(destino)
