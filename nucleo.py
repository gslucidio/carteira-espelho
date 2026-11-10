"""
Núcleo da carteira espelho: transforma os arquivos (CNABs, carteira ADM, estoque ADM, extrato)
nos dados usados pela planilha e pelo dashboard.

Uso pelo terminal/Spyder (pastas no computador):
    python nucleo.py PASTA_DO_FUNDO
Uso pelo Streamlit: app.py chama processar() com os arquivos enviados, todos em memória.
"""
import datetime as dt
import io
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

import leitores as L
import planilha


class ErroEntrada(Exception):
    """Falta um arquivo indispensável."""

OCORRENCIA_ENTRADA = {"01"}
ALIQUOTA_IOF = 0.0038   # IOF sobre aportes em FIDC

# Regras fixas de classificação do caixa, válidas para todos os fundos.
# (trecho do nome da contraparte, classificação se ENTRADA, classificação se SAÍDA, referência)
# O trecho é procurado como palavra inteira, sem diferenciar maiúsculas. None = a regra não se aplica.
# Para acrescentar uma regra, basta incluir uma linha nesta lista.
REGRAS_CAIXA = [
    ("ID RF", "Resgate do fundo DI", "Aplicação no fundo DI", ""),
    ("VIZ", None, "Taxas e despesas", "Taxa de gestão (VIZ Gestora)"),
    ("ID CORRETORA", None, "Taxas e despesas", "Taxa de administração (ID CTVM)"),
    ("ZIELO", None, "Taxas e despesas", "Despesa de plataforma (Zielo Tecnologia)"),
]


TIPOS_DESPESA = ["Taxa de gestão", "Administração e custódia", "Tributos (IOF)", "Plataforma e serviços", "Auditoria",
                 "Tarifas bancárias", "Outras despesas"]

# palavras genéricas, ignoradas ao reconhecer o nome de um FIDC investido no extrato
GENERICAS = {"FIDC", "FIC", "FI", "NP", "RL", "FUNDO", "FUNDOS", "DE", "DO", "DA", "EM", "INVESTIMENTO", "INVESTIMENTOS",
             "DIREITOS", "CREDITORIOS", "CREDITÓRIOS", "COTAS", "MULTISSETORIAL", "LP", "RF", "SENIOR", "SÊNIOR", "MEZ",
             "MEZANINO", "SUB", "SUBORDINADA", "CLASSE", "RESPONSABILIDADE", "LIMITADA", "NAO", "NÃO", "PADRONIZADOS", "II", "III"}


def tipo_fundo(descricao):
    """Fundo da carteira: cotas de FIDC (ativo-alvo de um FIC) ou fundo de liquidez."""
    t = (descricao or "").upper()
    return "Cotas de FIDC" if re.search(r"(?<![A-Z])FIDC(?![A-Z])|DIREITOS CRED", t) else "Liquidez"


def palavras_chave(descricao):
    return {p for p in re.findall(r"[A-ZÀ-Ú0-9]{3,}", (descricao or "").upper()) if p not in GENERICAS}


def tipo_despesa(texto, classificacao=None):
    """Tipo de despesa para o DRE, a partir da descrição/contraparte/referência."""
    t = (texto or "").upper()
    if classificacao == "Tarifa bancária" or "TARIFA" in t:
        return "Tarifas bancárias"
    if re.search(r"(?<![A-Z])IOF(?![A-Z])", t):
        return "Tributos (IOF)"
    if "GEST" in t or re.search(r"(?<![A-Z])VIZ(?![A-Z])", t):
        return "Taxa de gestão"
    if any(x in t for x in ("ADM", "CUST", "ID CORRETORA", "ID CTVM")):
        return "Administração e custódia"
    if "AUDIT" in t:
        return "Auditoria"
    if any(x in t for x in ("ZIELO", "PLATAFORMA", "REGISTR", "CERC", "CARTÓRIO", "CARTORIO", "ASSINATURA")):
        return "Plataforma e serviços"
    return "Outras despesas"


def atribuir_tipos(dados):
    """Preenche o tipo de despesa nos lançamentos de caixa e nas despesas a pagar."""
    for m in dados["caixa"]:
        if m["classificacao"] in ("Taxas e despesas", "Tarifa bancária"):
            m["tipo"] = tipo_despesa(f"{m.get('referencia') or ''} {m['contraparte']} {m['historico']}", m["classificacao"])
        else:
            m["tipo"] = ""
    for x in dados["despesas"]:
        x["tipo"] = tipo_despesa(x["descricao"])
    return dados


def _nome_classe(nome_carteira, nome_fundo_principal, grupo):
    """'FIDC NX CAPITAL NP RL - MEZ II' -> 'MEZ II'; sem arquivo, usa o grupo da carteira principal."""
    if nome_carteira and nome_fundo_principal and nome_carteira.startswith(nome_fundo_principal):
        resto = nome_carteira[len(nome_fundo_principal):].strip(" -")
        if resto:
            return resto
    return (grupo or "Classe").replace("Cota ", "").replace(" Pós", "").strip()


def _spread(indice):
    """'10,00 CDI' -> 0.10 (CDI + 10% a.a.)."""
    m = re.match(r"\s*([\d.,]+)\s*CDI", indice or "")
    return float(m.group(1).replace(",", ".")) / 100 if m else None


def montar_classes(adm, carteiras_classes, alertas):
    """
    Junta o passivo de cada classe registrado na carteira principal com a carteira da própria classe.
    Devolve uma lista vazia em fundos de classe única.
    """
    data = adm["data_posicao"]
    por_codigo = {}
    for c in carteiras_classes:
        if c["data_posicao"] != data or not c.get("codigo"):
            continue
        if c["codigo"] in por_codigo:
            alertas.append(f"Duas carteiras da mesma classe ({c['codigo']}) na mesma data; usada a primeira.")
            continue
        por_codigo[c["codigo"]] = c
    passivo = defaultdict(list)
    for p in adm.get("posicoes_classes", []):
        passivo[p["titulo"]].append(p)
    classes = []
    for codigo in sorted(set(passivo) | set(por_codigo)):
        arq = por_codigo.get(codigo)
        pos_p = passivo.get(codigo, [])
        pos_c = arq["posicoes_classes"] if arq else []
        grupo = (pos_p or pos_c or [{}])[0].get("grupo")
        indice = (pos_c or pos_p or [{}])[0].get("indice", "")
        nome = _nome_classe(arq["nome_fundo"] if arq else None, adm.get("nome_fundo"), grupo)
        passivo_principal = -sum(p["valor"] for p in pos_p)
        cl = dict(codigo=codigo, nome=nome, indice=indice, spread=_spread(indice),
                  passivo_principal=passivo_principal,
                  pl=arq["pl"] if arq else None, cotas=arq["quantidade_cotas"] if arq else -sum(p["quantidade"] for p in pos_p),
                  cota=arq["valor_cota"] if arq else None, aplicado=sum(p["valor_aplicado"] for p in (pos_c or pos_p)) * (1 if pos_c else -1),
                  rentabilidades=arq["rentabilidades"] if arq else {}, arquivo=arq["arquivo"] if arq else None,
                  divergencias=[])
        if not arq:
            alertas.append(f"Classe {nome}: a carteira principal registra {brl(passivo_principal)} de passivo, mas a carteira "
                           "da classe não foi enviada. Usado o valor da carteira principal.")
        elif not pos_p:
            alertas.append(f"Classe {nome}: há carteira da classe, mas a carteira principal não registra o passivo dela.")
        else:
            if abs(cl["pl"] - passivo_principal) >= 1:
                alertas.append(f"Classe {nome}: a carteira principal registra {brl(passivo_principal)} de passivo, mas a "
                               f"carteira da classe tem PL de {brl(cl['pl'])} (diferença de {brl(passivo_principal - cl['pl'])}).")
            # compara aplicação a aplicação (pela ordem e pela data/valor aplicado)
            restantes = list(pos_c)
            for p in pos_p:
                par = next((q for q in restantes if q["data"] == p["data"] and abs(q["valor_aplicado"] + p["valor_aplicado"]) < 0.01), None)
                if par is None:   # mesma data, valor aplicado diferente: compara assim mesmo
                    mesmas = [q for q in restantes if q["data"] == p["data"]]
                    par = min(mesmas, key=lambda q: abs(q["valor_aplicado"] + p["valor_aplicado"])) if mesmas else None
                if par is None:
                    cl["divergencias"].append(dict(data=p["data"], valor_aplicado=None, valor_aplicado_principal=-p["valor_aplicado"],
                                                   cotas_principal=-p["quantidade"],
                                                   cotas_classe=None, valor_principal=-p["valor"], valor_classe=None))
                    continue
                restantes.remove(par)
                if (abs(par["quantidade"] + p["quantidade"]) > 1e-6 or abs(par["valor"] + p["valor"]) >= 0.01
                        or abs(par["valor_aplicado"] + p["valor_aplicado"]) >= 0.01):
                    cl["divergencias"].append(dict(data=p["data"], valor_aplicado=par["valor_aplicado"],
                                                   valor_aplicado_principal=-p["valor_aplicado"], cotas_principal=-p["quantidade"],
                                                   cotas_classe=par["quantidade"], valor_principal=-p["valor"], valor_classe=par["valor"]))
            for q in restantes:
                cl["divergencias"].append(dict(data=q["data"], valor_aplicado=q["valor_aplicado"], valor_aplicado_principal=None,
                                               cotas_principal=None,
                                               cotas_classe=q["quantidade"], valor_principal=None, valor_classe=q["valor"]))
            if cl["divergencias"]:
                alertas.append(f"Classe {nome}: {len(cl['divergencias'])} aplicação(ões) com quantidade de cotas ou valor diferente "
                               "entre a carteira principal e a carteira da classe (detalhe na tela Passivo e cotas).")
        classes.append(cl)
    return classes


def regra_fixa(contraparte, entrada):
    nome = (contraparte or "").upper()
    for trecho, cls_ent, cls_sai, ref in REGRAS_CAIXA:
        if re.search(r"(?<![A-Z0-9])" + re.escape(trecho.upper()) + r"(?![A-Z0-9])", nome):
            cls = cls_ent if entrada else cls_sai
            if cls:
                return cls, ref
    return None


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")   # demais ocorrências com valor pago são tratadas como baixa


# =====================================================================================
# feriados nacionais
def _pascoa(ano):
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l_ = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l_) // 451
    mes = (h + l_ - 7 * m + 114) // 31
    dia = (h + l_ - 7 * m + 114) % 31 + 1
    return dt.date(ano, mes, dia)


def feriados(ano_ini, ano_fim):
    out = []
    for a in range(ano_ini, ano_fim + 1):
        p = _pascoa(a)
        lst = [(dt.date(a, 1, 1), "Confraternização Universal"), (p - dt.timedelta(48), "Carnaval"),
               (p - dt.timedelta(47), "Carnaval"), (p - dt.timedelta(2), "Sexta-feira Santa"),
               (dt.date(a, 4, 21), "Tiradentes"), (dt.date(a, 5, 1), "Dia do Trabalho"),
               (p + dt.timedelta(60), "Corpus Christi"), (dt.date(a, 9, 7), "Independência"),
               (dt.date(a, 10, 12), "Nossa Senhora Aparecida"), (dt.date(a, 11, 2), "Finados"),
               (dt.date(a, 11, 15), "Proclamação da República"), (dt.date(a, 12, 25), "Natal")]
        if a >= 2024:
            lst.append((dt.date(a, 11, 20), "Consciência Negra"))
        out += sorted(lst)
    return out


# =====================================================================================
# memória: o que o usuário corrigiu na última planilha gerada
def _data(v):
    return v.date() if isinstance(v, dt.datetime) else v


def memoria_vazia():
    return dict(arquivo=None, despesas={}, caixa={}, contrapartes={}, lotes={}, painel={})


def ultima_planilha(pasta_saida):
    arquivos = sorted(Path(pasta_saida).glob("Carteira_Espelho_*.xlsx"), key=lambda p: p.stat().st_mtime)
    arquivos = [a for a in arquivos if not a.name.startswith("~$")]
    return arquivos[-1] if arquivos else None


def ler_memoria(fonte):
    """Lê as correções manuais de uma planilha gerada antes (caminho ou arquivo em memória)."""
    mem = memoria_vazia()
    if fonte is None:
        return mem
    try:
        wb = load_workbook(io.BytesIO(L.conteudo(fonte)))
    except Exception:  # noqa: BLE001
        return mem
    if "Caixa" not in wb.sheetnames or "Despesas" not in wb.sheetnames:
        return mem
    mem["arquivo"] = L.nome_de(fonte)

    def linhas(nome, ini):
        if nome not in wb.sheetnames:
            return []
        out = []
        for r in wb[nome].iter_rows(min_row=ini, values_only=True):
            if r[0] in (None, "") or str(r[0]).strip().upper() == "TOTAL":
                break
            out.append(r)
        return out

    for r in linhas("Despesas", 5):
        desc, v_adm, v_g, obs = r[0], r[3], r[4], r[6]
        alterado = isinstance(v_g, (int, float)) and isinstance(v_adm, (int, float)) and abs(v_g - v_adm) > 0.005
        if alterado or obs:
            mem["despesas"][str(desc).strip().upper()] = (float(v_g) if alterado else None, v_adm, obs or "")

    votos = defaultdict(Counter)
    for r in linhas("Caixa", 5):
        id_, contraparte, entrada, cls, ref = str(r[0]), r[3], r[4], r[9], r[10]
        if ref and str(ref).startswith("Sugestão:"):
            ref = ""
        if cls and cls != "A classificar":
            mem["caixa"][id_] = (cls, ref or "")
            votos[(str(contraparte or "").upper(), bool(entrada))][cls] += 1
    mem["contrapartes"] = {k: c.most_common(1)[0][0] for k, c in votos.items()}

    for r in linhas("Estoque", 5):
        if r[0] and r[1]:
            mem["lotes"][str(r[0])] = r[1]

    if "Painel" in wb.sheetnames:
        rotulos = {"Mínimo exigido pelo regulamento": "minimo", "Data de início das atividades da Classe": "inicio",
                   "Prazo para enquadramento (dias)": "prazo"}
        for r in wb["Painel"].iter_rows(values_only=True):
            if r[0] in rotulos and r[1] not in (None, ""):
                mem["painel"][rotulos[r[0]]] = _data(r[1])
    return mem


# =====================================================================================
def processar(cnabs, carteiras_adm, estoques_adm, extratos, planilha_anterior=None, nome_pasta=None):
    """
    Cada argumento é uma lista de arquivos (caminho, bytes ou arquivo em memória).
    planilha_anterior: a última planilha gerada (opcional), de onde vêm as correções manuais.
    Devolve o dicionário 'dados' usado por planilha.montar() e pelo dashboard.
    """
    alertas, avisos = [], []
    mem = ler_memoria(planilha_anterior)

    # ---------------------------------------------------------------- carteira ADM
    carteiras = []
    for f in carteiras_adm:
        try:
            c = L.ler_carteira_adm(f)
            c["arquivo"] = L.nome_de(f)
            if c["data_posicao"]:
                carteiras.append(c)
        except Exception as e:  # noqa: BLE001
            alertas.append(f"Carteira ADM {L.nome_de(f)} não foi lida: {e}")
    if not carteiras:
        raise ErroEntrada("Nenhuma carteira diária da ADM válida (.xlsx). Ela é indispensável: é dela que vêm a data-base, "
                          "os fundos de liquidez, as despesas e as cotas.")
    # carteira principal: a que tem os ativos (direitos creditórios, fundos, despesas ou conta corrente).
    # As demais da mesma data, com posições só nas próprias cotas, são carteiras de classes (mezanino, sênior).
    def principal(c):
        return bool(c["direitos_creditorios"] or c["fundos"] or c["despesas"])
    principais = [c for c in carteiras if principal(c)]
    if not principais:
        raise ErroEntrada("Nenhuma das carteiras enviadas é a carteira principal do fundo (a que tem os direitos "
                          "creditórios e as despesas). Em fundos com várias classes, envie também a da subordinada.")
    adm = max(principais, key=lambda c: c["data_posicao"])
    data_base = adm["data_posicao"]
    classes = montar_classes(adm, [c for c in carteiras if not principal(c)], alertas)

    # ---------------------------------------------------------------- CNABs
    remessas, baixas = {}, []
    duplicados = defaultdict(int)
    for f in sorted(cnabs, key=L.nome_de):
        try:
            regs = L.ler_cnab(f)
        except Exception as e:  # noqa: BLE001
            alertas.append(f"CNAB {L.nome_de(f)} não foi lido: {e}")
            continue
        nome_f = L.nome_de(f)
        for r in regs:
            k = L.chave_titulo(r["cnpj_cedente"], r["numero_documento"])
            if r["ocorrencia"] in OCORRENCIA_ENTRADA:
                if k in remessas:
                    duplicados[(remessas[k]["arquivo"], nome_f)] += 1
                else:
                    remessas[k] = r
            elif r["valor_pago"] > 0:
                baixas.append((k, r))
            else:
                alertas.append(f"{nome_f}: título {r['numero_documento']} com ocorrência {r['ocorrencia']} "
                               f"sem valor pago; ignorado.")
    for (a1, a2), n in duplicados.items():
        alertas.append(f"{n} título(s) de {a2} já estavam em {a1} (arquivo repetido ou reenviado?); mantida a primeira versão.")
    baixa_por_titulo = {}
    for k, r in baixas:
        if k not in remessas:
            alertas.append(f"Baixa do título {r['numero_documento']} ({r['arquivo']}) sem remessa correspondente.")
        elif k in baixa_por_titulo:
            alertas.append(f"Título {r['numero_documento']} tem mais de uma baixa ({baixa_por_titulo[k]['arquivo']} "
                           f"e {r['arquivo']}); considerada a primeira.")
        else:
            baixa_por_titulo[k] = r
    nomes_cedentes = {r["cnpj_cedente"]: r["cedente"] for r in remessas.values()}

    titulos = []
    for k, r in sorted(remessas.items(), key=lambda kv: (kv[1]["data_arquivo"], kv[1]["numero_termo"],
                                                          kv[1]["numero_documento"].zfill(12))):
        b = baixa_por_titulo.get(k)
        num = r["numero_documento"]
        titulos.append(dict(
            chave=k, lote=mem["lotes"].get(k, f"Termo {r['numero_termo']:02d}"),
            numero=int(num) if num.isdigit() else num, cedente=r["cedente"], sacado=r["sacado"],
            doc_sacado=r["doc_sacado"], data_aquisicao=r["data_arquivo"], vencimento=r["vencimento"],
            valor_face=r["valor_face"], preco_pago=r["valor_aquisicao"],
            status="Liquidado" if b else "Aberto", data_baixa=b["data_pagamento"] if b else None,
            valor_recebido=b["valor_pago"] if b else None, arquivo_remessa=r["arquivo"],
            arquivo_baixa=b["arquivo"] if b else None))
    if not titulos:
        if cnabs or adm["direitos_creditorios"]:
            alertas.append("Nenhum título encontrado nos CNABs de remessa.")
    lotes = {}
    for t in titulos:
        lotes.setdefault(t["lote"], dict(nome=t["lote"], cedente=t["cedente"], data=t["data_aquisicao"], preco=0.0))
        lotes[t["lote"]]["preco"] += t["preco_pago"]

    # ---------------------------------------------------------------- estoque ADM
    estoques = []
    for f in estoques_adm:
        try:
            e = L.ler_estoque_adm(f)
            if e:
                estoques.append((e[0]["data_referencia"], L.nome_de(f), e))
        except Exception as ex:  # noqa: BLE001
            alertas.append(f"Estoque ADM {L.nome_de(f)} não foi lido: {ex}")
    estoques = [x for x in estoques if x[0] and x[0] <= data_base]
    if estoques:
        est_data, est_arq, estoque_adm = max(estoques, key=lambda x: x[0])
        if est_data != data_base:
            alertas.append(f"O estoque analítico da ADM é de {est_data:%d/%m/%Y}, mas a carteira é de {data_base:%d/%m/%Y}.")
    else:
        est_data, est_arq, estoque_adm = None, "não encontrado", []
        if adm["direitos_creditorios"]:
            alertas.append("Nenhum estoque analítico da ADM até a data-base: conferência título a título vazia.")

    # ---------------------------------------------------------------- extrato
    movs = {}
    for f in extratos:
        try:
            for m in L.ler_extrato(f):
                movs[m["id"]] = m
        except Exception as ex:  # noqa: BLE001
            alertas.append(f"Extrato {L.nome_de(f)} não foi lido: {ex}")
    caixa = sorted(movs.values(), key=lambda m: (m["data"], int(m["id"]) if m["id"].isdigit() else 0))
    nomes_ced = [n.upper() for n in nomes_cedentes.values()]
    preco_por_lote = {round(l_["preco"], 2): l_["nome"] for l_ in lotes.values()}
    baixas_por_data = defaultdict(lambda: [0.0, 0])
    for b in baixa_por_titulo.values():
        for chave in (b["data_pagamento"], (b["data_pagamento"], b["cnpj_cedente"])):
            baixas_por_data[chave][0] += b["valor_pago"]
            baixas_por_data[chave][1] += 1

    def doc14(m):
        d_ = m["doc_contraparte"]
        return d_.zfill(14) if len(d_) in (13, 14) else d_

    def e_cedente(m):
        nome = (m["contraparte"] or "").upper()
        return doc14(m) in nomes_cedentes or any(nome.startswith(n) or n.startswith(nome) for n in nomes_ced if nome)

    for f_ in adm["fundos"]:
        f_["tipo"] = tipo_fundo(f"{f_['descricao']} {f_.get('grupo') or ''}" if "DIREITOS" in (f_.get("grupo") or "").upper()
                                else f_["descricao"])
    chaves_fidc = {}
    for f_ in adm["fundos"]:
        if f_["tipo"] == "Cotas de FIDC":
            for p in palavras_chave(f_["descricao"]):
                chaves_fidc[p] = f_["descricao"]

    def fidc_investido(m):
        nome = (m["contraparte"] or "").upper()
        return next((v for p, v in chaves_fidc.items() if re.search(r"(?<![A-Z0-9])" + re.escape(p) + r"(?![A-Z0-9])", nome)), None)

    a_classificar = 0
    for m in caixa:
        if m["contraparte"].startswith("CNPJ ") and doc14(m) in nomes_cedentes:
            m["contraparte"] = nomes_cedentes[doc14(m)]
        aprendido = mem["contrapartes"].get((m["contraparte"].upper(), bool(m["entrada"])))
        ref = ""
        if m["id"] in mem["caixa"]:
            cls, ref = mem["caixa"][m["id"]]
        elif m["historico"].upper().startswith("TARIFA"):
            cls = "Tarifa bancária"
        elif regra_fixa(m["contraparte"], m["entrada"]):
            cls, ref = regra_fixa(m["contraparte"], m["entrada"])
        elif fidc_investido(m):
            cls = "Resgate/amortização de cotas de FIDC" if m["entrada"] else "Aplicação em cotas de FIDC"
            ref = fidc_investido(m)
        elif e_cedente(m):
            cls = "Recebimento de DC" if m["entrada"] else "Aquisição de DC"
        else:
            # sem regra fixa: quem decide é o usuário. Se a mesma contraparte já foi classificada
            # antes, isso aparece só como sugestão na Referência.
            cls = "A classificar"
            a_classificar += 1
            if aprendido:
                ref = f"Sugestão: {aprendido} (mesma contraparte já classificada assim antes)"
        if not ref and cls == "Aquisição de DC":
            ref = preco_por_lote.get(round(m["saida"], 2), "")
        if not ref and cls == "Recebimento de DC":
            v = baixas_por_data.get(m["data"])
            if not (v and abs(v[0] - m["entrada"]) < 0.01):
                v = baixas_por_data.get((m["data"], doc14(m)))
            if v and abs(v[0] - m["entrada"]) < 0.01:
                ref = f"Baixas de {m['data']:%d/%m} ({v[1]} títulos)"
        m["classificacao"], m["referencia"] = cls, ref
    # IOF de 0,38% sobre aportes: pagamentos à administradora que batem com 0,38% de uma entrada recente
    entradas = [m for m in caixa if m["entrada"] and m["classificacao"] not in
                ("Resgate do fundo DI", "Recebimento de DC", "Resgate/amortização de cotas de FIDC")]
    for m in caixa:
        if not m["saida"] or m["classificacao"] not in ("Taxas e despesas", "A classificar"):
            continue
        if m["referencia"] and not m["referencia"].startswith(("Confirmar", "Taxa de administração (ID CTVM)", "Sugestão")):
            continue
        origem = next((e for e in entradas if 0 <= (m["data"] - e["data"]).days <= 31
                       and abs(round(e["entrada"] * ALIQUOTA_IOF, 2) - m["saida"]) < 0.015), None)
        if origem:
            if m["classificacao"] == "A classificar":
                a_classificar -= 1
            m["classificacao"] = "Taxas e despesas"
            m["referencia"] = f"IOF de 0,38% sobre o aporte de {brl(origem['entrada'])} em {origem['data']:%d/%m/%Y}"
    if a_classificar:
        alertas.append(f"{a_classificar} lançamento(s) de caixa 'A classificar' (em amarelo no Caixa). "
                       "Classifique-os; a escolha fica gravada na planilha e é mantida nas próximas vezes.")

    # ---------------------------------------------------------------- despesas
    despesas, usados = [], set()
    for x in adm["despesas"]:
        chave = x["descricao"].strip().upper()
        valor_g, obs = x["valor"], ""
        if chave in mem["despesas"]:
            usados.add(chave)
            v_g, v_adm_ant, obs = mem["despesas"][chave]
            if v_g is not None and abs(v_g - x["valor"]) > 0.005:
                valor_g = v_g
                avisos.append(f"Despesa '{x['descricao']}': mantido o valor da gestora {brl(v_g)} "
                              f"(ADM: {brl(x['valor'])}), corrigido na planilha anterior.")
        elif x["vencimento"] and x["vencimento"].year > data_base.year + 5:
            obs = f"Vencimento em {x['vencimento']:%Y}: verificar com a ADM se é resíduo a baixar."
        despesas.append(dict(x, valor_gestora=valor_g, observacao=obs))
    for chave in set(mem["despesas"]) - usados:
        if mem["despesas"][chave][0] is not None:
            alertas.append(f"A despesa '{chave}', corrigida antes pela gestora, não está mais na carteira da ADM "
                           "(a ADM corrigiu?).")

    # ---------------------------------------------------------------- enquadramento
    inicio = mem["painel"].get("inicio")
    if not inicio:
        aportes = [m["data"] for m in caixa if m["classificacao"] == "Aporte de cotista"]
        inicio = min(aportes) if aportes else None

    # ---------------------------------------------------------------- montagem
    nome_fundo = adm["nome_fundo"] or nome_pasta or "Fundo"
    anos = [t["vencimento"].year for t in titulos if t["vencimento"]] + [data_base.year]
    dados = dict(
        fundo=nome_fundo, conta=adm["conta"] or "conta corrente", data_base=data_base,
        titulos=titulos, lotes=list(lotes.values()),
        estoque_adm=estoque_adm, estoque_adm_arquivo=est_arq, estoque_adm_data=est_data,
        caixa=caixa, despesas=despesas, fundos=adm["fundos"], quantidade_cotas=adm["quantidade_cotas"],
        classes=classes,
        adm=adm, carteira_adm_arquivo=adm["arquivo"],
        sacados=sorted({t["sacado"] for t in titulos}),
        vencimentos=sorted({t["vencimento"] for t in titulos if t["vencimento"]}),
        minimo_dc=mem["painel"].get("minimo", 0.5), inicio_classe=inicio,
        prazo_enquadramento=mem["painel"].get("prazo", 180),
        feriados=feriados(min(anos) - 1, max(anos) + 1),
        alertas=alertas, avisos=avisos, memoria=mem["arquivo"],
    )
    dados["tipos_despesa"] = TIPOS_DESPESA
    atribuir_tipos(dados)
    dados["arquivo_sugerido"] = (f"Carteira_Espelho_{nome_pasta or nome_arquivo(nome_fundo)}_{data_base:%Y%m%d}.xlsx")
    return dados


def nome_arquivo(texto):
    return "".join(ch for ch in texto if ch.isalnum() or ch in " -_").strip().replace(" ", "_")


def gerar(pasta):
    """Uso pelo terminal/Spyder: lê as pastas do fundo e grava a planilha em saida/."""
    pasta = Path(pasta)
    ent = pasta / "entradas"
    for sub in ("cnab", "adm", "extrato"):
        (ent / sub).mkdir(parents=True, exist_ok=True)
    (pasta / "saida").mkdir(exist_ok=True)
    arquivos = lambda sub, exts: sorted(p for p in (ent / sub).iterdir()  # noqa: E731
                                        if p.suffix.upper() in exts and not p.name.startswith("~$"))
    try:
        dados = processar(arquivos("cnab", (".REM", ".TXT", ".RET")), arquivos("adm", (".XLSX",)),
                          arquivos("adm", (".CSV",)), arquivos("extrato", (".CSV",)),
                          ultima_planilha(pasta / "saida"), nome_pasta=pasta.name)
    except ErroEntrada as e:
        sys.exit(str(e))
    destino = pasta / "saida" / dados["arquivo_sugerido"]
    try:
        planilha.montar(dados, destino)
    except PermissionError:
        sys.exit(f"Não consegui salvar {destino.name}. Feche a planilha no Excel e rode de novo.")
    t = dados["titulos"]
    print(f"Planilha gerada: {destino}")
    print(f"  Data-base {dados['data_base']:%d/%m/%Y} | {len(t)} títulos, "
          f"{sum(1 for x in t if x['status'] == 'Aberto')} em aberto | {len(dados['caixa'])} lançamentos de caixa")
    if dados["memoria"]:
        print(f"  Correções manuais lidas de: {dados['memoria']}")
    for a in dados["avisos"]:
        print("  MANTIDO:", a)
    for a in dados["alertas"]:
        print("  ALERTA:", a)
    return destino


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
    else:
        gerar(sys.argv[1])
