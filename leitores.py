"""
Leitura dos arquivos de entrada da carteira espelho.

Cada função lê um tipo de arquivo e devolve dados simples (listas de dicionários),
sem nenhuma regra de negócio. As regras ficam em gerar_carteira.py.
"""
import csv
import datetime as dt
import io
import re
from pathlib import Path

from openpyxl import load_workbook


# ---------------------------------------------------------------- utilidades
# Todas as funções de leitura aceitam um caminho (str/Path), bytes ou um arquivo em memória
# (como o UploadedFile do Streamlit). Nada é gravado em disco.
def conteudo(fonte):
    if hasattr(fonte, "getvalue"):
        return fonte.getvalue()
    if isinstance(fonte, (bytes, bytearray)):
        return bytes(fonte)
    return Path(fonte).read_bytes()


def nome_de(fonte, nome=None):
    if nome:
        return nome
    if getattr(fonte, "name", None):
        return Path(fonte.name).name
    if isinstance(fonte, (str, Path)):
        return Path(fonte).name
    return "arquivo"


def _data_ddmmaa(s):
    s = s.strip()
    if not s or set(s) == {"0"}:
        return None
    return dt.date(2000 + int(s[4:6]), int(s[2:4]), int(s[0:2]))


def _data_br(s):
    s = (s or "").strip()
    if not s:
        return None
    return dt.datetime.strptime(s[:10], "%d/%m/%Y").date()


def _num_br(s):
    """'1.234,56' -> 1234.56"""
    s = (s or "").strip()
    if not s:
        return 0.0
    return float(s.replace(".", "").replace(",", "."))


def so_digitos(s):
    return re.sub(r"\D", "", s or "")


def formata_doc(d):
    d = so_digitos(d)
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    return d


def chave_titulo(cnpj_cedente, numero_documento):
    """Identificador único do título: CNPJ do cedente + número do documento."""
    return f"{so_digitos(cnpj_cedente)}-{str(numero_documento).strip().lstrip('0')}"


# ---------------------------------------------------------------- CNAB 444
def ler_cnab(fonte, nome=None):
    """
    Lê um arquivo CNAB 444 de FIDC (remessa ou baixa).
    Devolve uma lista com um dicionário por registro de detalhe.
    """
    nome = nome_de(fonte, nome)
    linhas = conteudo(fonte).decode("latin-1").splitlines()
    linhas = [l for l in linhas if l.strip()]
    if not linhas or linhas[0][0] != "0":
        raise ValueError(f"{nome}: não parece um CNAB (sem registro header).")
    data_arquivo = _data_ddmmaa(linhas[0][94:100])
    registros = []
    for n, l in enumerate(linhas[1:], start=2):
        if l[0] != "1":
            continue
        if len(l) < 394:
            raise ValueError(f"{nome}, linha {n}: registro com {len(l)} posições (esperado 444).")
        doc_sacado = l[220:234]
        tipo_sacado = l[218:220]
        registros.append(dict(
            arquivo=nome,
            data_arquivo=data_arquivo,
            ocorrencia=l[108:110],
            numero_documento=l[110:120].strip().lstrip("0"),
            vencimento=_data_ddmmaa(l[120:126]),
            valor_face=int(l[126:139]) / 100,
            valor_pago=int(l[82:92]) / 100,
            data_pagamento=_data_ddmmaa(l[94:100]),
            data_emissao=_data_ddmmaa(l[150:156]),
            numero_termo=int(l[161:175] or 0),
            valor_aquisicao=int(l[192:205]) / 100,
            sacado=l[234:274].strip(),
            doc_sacado=formata_doc(doc_sacado[-11:] if tipo_sacado == "01" else doc_sacado),
            cedente=l[334:380].strip(),
            cnpj_cedente=so_digitos(l[380:394]),
        ))
    return registros


# ---------------------------------------------------------------- estoque ADM (CSV)
def ler_estoque_adm(fonte):
    """Estoque analítico da administradora (CSV separado por ';')."""
    texto = conteudo(fonte).decode("utf-8-sig", errors="replace")
    linhas = list(csv.DictReader(io.StringIO(texto), delimiter=";"))
    if linhas and "NU_DOCUMENTO" not in linhas[0]:
        raise ValueError("não parece o estoque analítico da ADM (sem a coluna NU_DOCUMENTO).")
    out = []
    for r in linhas:
        out.append(dict(
            chave=chave_titulo(r["DOC_CEDENTE"], r["NU_DOCUMENTO"]),
            numero_documento=r["NU_DOCUMENTO"].strip(),
            cedente=r["NOME_CEDENTE"].strip(),
            sacado=r["NOME_SACADO"].strip(),
            doc_sacado=r["DOC_SACADO"].strip(),
            data_aquisicao=_data_br(r["DATA_AQUISICAO"]),
            vencimento=_data_br(r["DATA_VENCIMENTO_AJUSTADA"]) or _data_br(r["DATA_VENCIMENTO_ORIGINAL"]),
            valor_nominal=_num_br(r["VALOR_NOMINAL"]),
            valor_presente=_num_br(r["VALOR_PRESENTE"]),
            valor_aquisicao=_num_br(r["VALOR_AQUISICAO"]),
            valor_pdd=_num_br(r["VALOR_PDD"]),
            situacao=r["SITUACAO_RECEBIVEL"].strip(),
            data_referencia=_data_br(r["DATA_REFERENCIA"]),
        ))
    return out


# ---------------------------------------------------------------- extrato (CSV)
def ler_extrato(fonte):
    """Extrato da conta corrente do fundo (CSV da ID CTVM, separado por ';')."""
    texto = conteudo(fonte).decode("utf-8-sig", errors="replace")
    linhas = list(csv.DictReader(io.StringIO(texto), delimiter=";"))
    if linhas and "dt_mov" not in linhas[0]:
        raise ValueError("não parece o extrato da conta (sem a coluna dt_mov).")
    out = []
    for r in linhas:
        if not (r.get("id") or "").strip():
            continue
        comp = (r.get("ds_complemento") or "").strip()
        info = r.get("ds_info_adicionais") or ""
        nome = re.search(r"Nome:([^,}]+)", info)
        doc = re.search(r"CpfCnpj:([^,}]+)", info)
        if comp == "Lancamento de tarifa":
            contraparte = "Tarifa"
        elif comp.upper().startswith("DEST:") and comp[5:].strip():
            contraparte = comp[5:].strip()          # "Dest: Viz Gestora" -> "Viz Gestora"
        elif comp and not comp.upper().startswith("DEST:"):
            contraparte = comp
        elif nome and nome.group(1) != "null":
            contraparte = nome.group(1).strip()
        elif doc and doc.group(1) != "null":
            contraparte = "CNPJ " + doc.group(1).strip()
        else:
            contraparte = ""
        valor = float(r["valor"])
        credito = r["fl_debito_credito"].strip().upper() == "C"
        out.append(dict(
            id=r["id"].strip(),
            data=_data_br(r["dt_mov"]),
            historico=r["ds_historico"].strip(),
            contraparte=contraparte,
            doc_contraparte=so_digitos(doc.group(1)) if doc and doc.group(1) != "null" else "",
            entrada=valor if credito else 0.0,
            saida=0.0 if credito else valor,
            saldo_banco=float(r["vl_saldo_in"]),
        ))
    return out


# ---------------------------------------------------------------- carteira ADM (XLSX)
def _linha(ws, r):
    return [c.value for c in ws[r]]


def _acha(ws, texto, coluna=None, inicio=1):
    """Primeira linha (a partir de 'inicio') que contém uma célula começando com 'texto'."""
    alvo = texto.lower()
    for row in ws.iter_rows(min_row=inicio):
        for c in row:
            if isinstance(c.value, str) and c.value.strip().lower().startswith(alvo):
                if coluna is None or c.column == coluna:
                    return c.row, c.column
    return None, None


def _numeros_a_direita(ws, r, col):
    return [v for v in _linha(ws, r)[col:] if isinstance(v, (int, float))]


def _secao(ws, titulo):
    """Lê uma tabela da carteira: acha o título, o cabeçalho e as linhas até 'Total'."""
    r, _ = _acha(ws, titulo, coluna=2)
    if r is None:
        return []
    # cabeçalho: primeira linha depois do título que começa com "Descri"
    rh = None
    for rr in range(r + 1, r + 8):
        v = ws.cell(rr, 2).value
        if isinstance(v, str) and v.strip().lower().startswith("descri"):
            rh = rr
            break
    if rh is None:
        return []
    cab = {}
    for c in ws[rh]:
        if isinstance(c.value, str) and c.value.strip():
            cab[c.value.strip()] = c.column
    itens = []
    for rr in range(rh + 1, ws.max_row + 1):
        b = ws.cell(rr, 2).value
        if any(isinstance(c.value, str) and c.value.strip().lower().startswith("total") for c in ws[rr]):
            break
        if not b:
            continue
        item = {nome: ws.cell(rr, col).value for nome, col in cab.items()}
        if not any(isinstance(v, (int, float)) for k, v in item.items() if k != next(iter(cab))):
            continue  # linha de subtítulo, sem valores
        itens.append(item)
    return itens


def _valor(item, *nomes):
    for n in nomes:
        for k, v in item.items():
            if k.lower().startswith(n.lower()):
                return v
    return None


def _pct_br(s):
    """'(6,6362%)' -> -0.066362 ; '0,0886%' -> 0.000886 ; ' - ' -> None"""
    if isinstance(s, (int, float)):
        return float(s)
    s = (s or "").strip()
    if not s or s == "-":
        return None
    neg = s.startswith("(")
    s = s.strip("()%").replace(".", "").replace(",", ".")
    try:
        v = float(s) / 100
    except ValueError:
        return None
    return -v if neg else v


def _posicoes_renda_fixa(ws):
    """
    Posições da tabela de Renda Fixa com colunas Id/Operação/... (cotas de classes do próprio fundo).
    Lê pela ordem dos valores na linha, porque o cabeçalho e os dados ficam em colunas mescladas diferentes.
    """
    out, grupo, dentro = [], None, False
    for row in ws.iter_rows():
        v = [c.value for c in row if c.value is not None]
        if not v:
            continue
        b = v[0]
        if b == "Id":
            dentro = True
            continue
        if not dentro:
            continue
        if isinstance(b, str) and b.strip().lower().startswith("total renda"):
            dentro = False
            continue
        if isinstance(b, str):
            grupo = b.strip()
            continue
        if isinstance(b, (int, float)) and len(v) >= 15 and isinstance(v[1], (dt.datetime, dt.date)):
            out.append(dict(grupo=grupo, id=v[0], data=v[1].date() if isinstance(v[1], dt.datetime) else v[1],
                            titulo=str(v[4]), quantidade=float(v[6]), taxa=v[8], indice=str(v[9]),
                            valor_aplicado=float(v[10]), pu=float(v[11]), valor=float(v[14])))
    return out


def _rentabilidades(ws):
    r, _ = _acha(ws, "Rentabilidades")
    if r is None:
        return {}
    cab, out = None, {}
    for rr in range(r + 1, min(r + 6, ws.max_row + 1)):
        vals = [c.value for c in ws[rr] if c.value is not None]
        if not vals:
            continue
        if "Dia" in vals:
            cab = [str(x).strip() for x in vals]
            continue
        if cab and isinstance(vals[0], str) and vals[0].strip() in ("Carteira", "(%) CDI"):
            out[vals[0].strip()] = dict(zip(cab, [_pct_br(x) for x in vals[1:]]))
    return out


def ler_carteira_adm(fonte):
    """Composição de Carteira (Fechamento) da administradora, em XLSX."""
    ws = load_workbook(io.BytesIO(conteudo(fonte)), data_only=True).active
    if _acha(ws, "Composição de Carteira")[0] is None and _acha(ws, "Data Posição")[0] is None:
        raise ValueError("não parece a carteira diária da ADM (sem 'Composição de Carteira').")
    r, c = _acha(ws, "Data Posição")
    data_posicao = None
    if r:
        for v in _linha(ws, r)[c:]:
            if isinstance(v, (dt.datetime, dt.date)):
                data_posicao = v.date() if isinstance(v, dt.datetime) else v
                break
            if isinstance(v, str) and re.match(r"\d{2}/\d{2}/\d{4}", v.strip()):
                data_posicao = _data_br(v)
                break

    fundos = [dict(descricao=_valor(i, "Descri"), quantidade=_valor(i, "Qtd. Cotas"),
                   valor_cota=_valor(i, "Valor Cota"), valor=_valor(i, "Valor Líquido", "Valor Bruto"))
              for i in _secao(ws, "Cotas de Investimento") or _secao(ws, "Renda Fixa")]
    if not fundos:
        fundos = [dict(descricao=_valor(i, "Descri"), quantidade=_valor(i, "Qtd. Cotas"),
                       valor_cota=_valor(i, "Valor Cota"), valor=_valor(i, "Valor Líquido", "Valor Bruto"))
                  for i in _secao(ws, "Renda Fixa")]
    dcs = [dict(descricao=_valor(i, "Descri"), codigo=_valor(i, "Código"), quantidade=_valor(i, "Qtde"),
                valor_custo=_valor(i, "Valor Custo"), valor_mercado=_valor(i, "Valor Mercado"))
           for i in _secao(ws, "Direito Creditório")]
    despesas = []
    for i in _secao(ws, "Valores a Liquidar"):
        lanc, venc = _valor(i, "Lançamento"), _valor(i, "Vencimento")
        despesas.append(dict(
            descricao=str(_valor(i, "Descri")).strip(),
            lancamento=lanc.date() if isinstance(lanc, dt.datetime) else lanc,
            vencimento=venc.date() if isinstance(venc, dt.datetime) else venc,
            valor=-float(_valor(i, "Valor") or 0),  # positivo = valor a pagar
        ))
    cc = sum(float(_valor(i, "Valor") or 0) for i in _secao(ws, "Saldos em Conta Corrente"))

    r_resumo, _ = _acha(ws, "Resumo da Carteira")

    def numero_apos(texto, inicio=1):
        rr, cc_ = _acha(ws, texto, inicio=inicio)
        if rr is None:
            return None
        nums = _numeros_a_direita(ws, rr, cc_)
        return nums[0] if nums else None

    nome_fundo, codigo = None, None
    r, c = _acha(ws, "Carteira:")
    if r:
        for v in _linha(ws, r)[c:]:
            if isinstance(v, str) and v.strip():
                m = re.match(r"^(\d+)\s*-\s*(.*)$", v.strip())
                codigo, nome_fundo = (m.group(1), m.group(2)) if m else (None, v.strip())
                break
    contas = [str(_valor(i, "Descri")).strip() for i in _secao(ws, "Saldos em Conta Corrente")]

    posicoes = _posicoes_renda_fixa(ws)
    return dict(
        codigo=codigo,
        posicoes_classes=posicoes,
        rentabilidades=_rentabilidades(ws),
        nome_fundo=nome_fundo,
        conta=contas[0] if contas else "",
        data_posicao=data_posicao,
        fundos=fundos,
        direitos_creditorios=dcs,
        despesas=despesas,
        conta_corrente=cc,
        pl=numero_apos("PATRIMÔNIO FECHAMENTO"),
        quantidade_cotas=numero_apos("Qtd. Cotas", inicio=r_resumo or 1),
        valor_cota=numero_apos("Cota Líquida", inicio=r_resumo or 1),
    )
