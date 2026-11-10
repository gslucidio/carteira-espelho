# Carteira Espelho — VIZ Gestora

Dashboard em Streamlit para a gestão independente da carteira dos FIDCs: monta o estoque
da gestora a partir dos CNABs, compara com a carteira e o estoque da administradora,
concilia o caixa pelo extrato e gera a planilha completa (com fórmulas) para auditoria.

## Como usar

1. Abra o app e envie os arquivos pela barra lateral, cada um na sua categoria:
   1. **CNABs** — remessas e baixas (.REM), desde o início do fundo
   2. **Carteira diária da ADM** (.xlsx) — obrigatório; define a data-base. Em fundos com várias classes,
      envie a carteira principal (subordinada) e a de cada mezanino/sênior, da mesma data
   3. **Estoque analítico da ADM** (.csv)
   4. **Extrato da conta** (.csv)
   5. **Última planilha baixada** (opcional) — traz as correções do uso anterior
2. Navegue pelas telas: Visão geral, DRE, Passivo e cotas, Estoque, Caixa e despesas, Conferências.
3. Classifique os lançamentos de caixa em amarelo e corrija despesas, se discordar da ADM.
4. **Baixe a planilha ao terminar.** Ela é a memória da próxima sessão.

## Privacidade

Os arquivos ficam só na memória da sessão: o app não grava nada em disco nem usa cache
compartilhado, e tudo some ao fechar a aba. O `.gitignore` impede subir arquivos de dados
para o repositório. Mantenha o repositório privado.

## Rodar no próprio computador

    pip install -r requirements.txt
    streamlit run app.py

Também dá para gerar a planilha sem o dashboard, lendo pastas locais
(`PASTA/entradas/cnab`, `PASTA/entradas/adm`, `PASTA/entradas/extrato`):

    python nucleo.py PASTA

## Publicar no Streamlit Community Cloud

1. Suba estes arquivos num repositório do GitHub (de preferência privado).
2. Em share.streamlit.io: **Create app** → escolha o repositório → arquivo principal `app.py` → **Deploy**.

## Arquivos

| Arquivo | O que faz |
|---|---|
| `app.py` | Telas do dashboard |
| `nucleo.py` | Regras: títulos, baixas, classificação do caixa, despesas, correções da planilha anterior |
| `leitores.py` | Leitura dos layouts da ID CTVM (CNAB 444, carteira, estoque, extrato) |
| `calculos.py` | Contas do dashboard (curva, PL, conferências, painel) |
| `planilha.py` | Montagem da planilha Excel |

Regras fixas de classificação do caixa (ID RF, VIZ, ID Corretora, Zielo) ficam em `REGRAS_CAIXA`,
no início de `nucleo.py`. Para incluir outra contraparte, acrescente uma linha.

## Fundos com várias classes de cotas

A carteira principal da ADM traz os ativos e registra as demais classes como passivo (posições negativas).
O app junta cada classe com a sua própria carteira e mostra: PL e cota por classe, participação no patrimônio,
rentabilidade contra o benchmark (CDI + spread, verificado com o CDI implícito na própria carteira),
índices de subordinação e divergências aplicação a aplicação entre as duas visões.
A ordem de prioridade é assumida pelo spread (maior spread = mais júnior); confirme com o regulamento.

## FICs (fundos que aplicam em cotas de outros FIDCs)

Sem CNABs nem estoque: envie só a carteira (principal e das classes, se houver) e o extrato.
Cotas de fundos com "FIDC" ou "Direitos Creditórios" no nome entram como **cotas de FIDC investidas**
(ativo-alvo), separadas dos fundos de liquidez; o tipo pode ser trocado na tabela de fundos da aba Carteira.
Transferências para esses FIDCs são classificadas como aplicação ou resgate de cotas de FIDC, e o DRE
mostra a valorização das cotas investidas.

Pagamentos à administradora iguais a 0,38% de um aporte recebido até 31 dias antes são classificados
como IOF sobre o aporte.

## Limitações

- Layouts da ID CTVM. Outra administradora exige ajustar `leitores.py`.
- Termos de cessão não são lidos: confira os totais da tela Conferências (Lotes) com o quadro resumo de cada termo.
