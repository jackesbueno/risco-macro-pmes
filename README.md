# Risco macroeconômico para PMEs: teste de estresse com regressão e Monte Carlo

Modelo em Python que mede, em duas etapas, o quanto o crescimento da atividade de **Serviços de TI**, **Indústria** e **Comércio Varejista** depende da **Selic**, do **IPCA** e do **câmbio**, e como esse risco chega ao lucro, ao fluxo de caixa e à liquidez de pequenas e médias empresas (PMEs).

Este repositório traz o código do Trabalho de Graduação em Gestão Financeira **"Modelo econométrico integrado à simulação de Monte Carlo para análise de risco financeiro em Pequenas e Médias Empresas"**. O trabalho foi aprovado no XVII FATECLOG 2026 (Fatec São Sebastião) e pela banca da Fatec Bragança Paulista.

> O modelo é uma ferramenta de teste de estresse. Ele mede a sensibilidade histórica de cada setor às variáveis macroeconômicas e **não** prevê o faturamento de uma empresa específica.

---

## Como o modelo funciona

### Etapa 1: variáveis macro e atividade setorial

- **Dados:** 176 meses, de 01/2012 a 08/2026, todos de fontes oficiais (Banco Central e IBGE).
- **Variável-alvo:** crescimento em 12 meses do índice de volume dessazonalizado de cada setor.
- **Regressores:** Selic (% a.a.), IPCA acumulado em 12 meses e variação do câmbio em 12 meses.
- **Correlação:** Pearson, Spearman e parcial, com intervalo de confiança por bootstrap circular em blocos de 12 meses. Também calcula correlações defasadas de 0 a 12 meses.
- **Regressão:** mínimos quadrados (MQO) com erros-padrão de Newey-West (HAC, 12 defasagens), pois as taxas de 12 meses se sobrepõem e os resíduos são autocorrelacionados.
- **Validação:** validação cruzada em blocos e walk-forward com embargo de 12 meses. Há duas checagens de robustez: Selic defasada em 6 meses e amostra sem a COVID (03/2020 a 06/2021).
- **Monte Carlo:** 10.000 cenários sorteados de uma t de Student multivariada (5 g.l.), com a covariância histórica (decomposição de Cholesky) e a incerteza dos coeficientes. Os resíduos são reamostrados por bootstrap em blocos, e os cenários ficam limitados à faixa histórica das variáveis, sem extrapolação.
- **Métricas de risco:** VaR 95% e 99%, CVaR 95% e probabilidade de contração por setor.

O IBC-Br (proxy mensal do PIB) **não entra no modelo**. Ele já inclui a atividade dos próprios setores e distorceria os coeficientes, então aparece apenas como contexto descritivo.

### Etapa 2: transmissão ao caixa e ao balanço

- Três perfis contábeis **fictícios** de PMEs (DRE, balanço e indicadores), um por setor.
- O script confere as identidades contábeis e corrige o que não fecha, registrando cada correção.
- Projeta DRE, fluxo de caixa e balanço em quatro cenários: ano-base, cenário atual, estresse moderado e estresse severo. Também aplica os 10.000 cenários do Monte Carlo.
- Análise de sensibilidade: gráfico tornado (±1 desvio-padrão em uma variável por vez), correlação de Spearman e regressão padronizada sobre as simulações.

---

## Principais resultados

| Setor | R² | R² walk-forward | CVaR 95% | Prob. de contração |
|---|---|---|---|---|
| Serviços de TI | 0,26 | −0,00 | 16,5 pp | 9,8% |
| Indústria | 0,28 | −0,19 | 12,7 pp | 56,0% |
| Comércio Varejista | 0,24 | −1,51 | 10,2 pp | 33,1% |

- Serviços de TI tem o maior risco de cauda, e boa parte desse risco não é explicada pelas variáveis macro.
- A alta probabilidade de contração da Indústria vem do crescimento esperado negativo (−0,7%), e não de um risco de cauda maior.
- No estresse severo, o lucro da indústria fictícia cai de R$ 488 mil para R$ 164 mil e a liquidez corrente passa de 1,43 para 1,32, sem necessidade de financiamento. Nenhum setor fica com liquidez corrente abaixo de 1,0.
- O principal canal de transmissão é a compressão de margem, pelos custos corrigidos pela inflação e pelos juros da dívida indexada. O capital de giro varia pouco.
- O R² walk-forward é nulo ou negativo nos três setores. O modelo mede sensibilidade histórica e não serve como previsão.

---

## Como executar (Google Colab)

1. Abra um notebook novo no [Google Colab](https://colab.research.google.com/).
2. Cole o conteúdo de `pmes_tcc_colab.py` em uma célula e execute. Outra opção é enviar o arquivo e rodar `%run pmes_tcc_colab.py`.
3. Quando aparecer o botão **"Escolher arquivos"**, selecione todos os arquivos de entrada de uma vez. **Os nomes não importam:** o script reconhece cada arquivo pelo conteúdo.
4. Se faltar algum arquivo, o script avisa o que falta e reabre o upload, até 3 vezes.
5. Ao final, o relatório em PDF e um `.zip` com tabelas e figuras são baixados automaticamente.

As bibliotecas `reportlab` e `openpyxl` são instaladas automaticamente se faltarem. As demais (numpy, pandas, scipy e matplotlib) já vêm no Colab.

Para rodar fora do Colab, coloque os arquivos na pasta atual (ou defina a variável de ambiente `DATA_DIR`) e execute `python pmes_tcc_colab.py`.

---

## Dados de entrada

| Arquivo | Fonte | Formato |
|---|---|---|
| Selic diária | BCB, SGS série 11 | CSV |
| Câmbio PTAX compra | BCB, SGS série 10813 (opcional; se faltar, é baixado da API do BCB) | CSV ou XLSX |
| IPCA mensal | IBGE | CSV |
| IBC-Br | Banco Central (planilha com a aba "Atual") | XLSX |
| Serviços de TI | IBGE/SIDRA, PMS, tabela 8688 (subgrupo 2.1.2) | CSV |
| Indústria geral | IBGE/SIDRA, PIM-PF, tabela 8888 | CSV |
| Comércio varejista | IBGE/SIDRA, PMC, tabela 8880 | CSV |
| Perfis financeiros (3) | Fictícios (abas DRE, Balanço Patrimonial e Indicadores) | XLSX ou CSV |

As tabelas do SIDRA devem ser exportadas brutas, com os blocos "sem ajuste sazonal" e "com ajuste sazonal". O script não gera dados simulados no lugar de dados faltantes: se um arquivo faltar ou vier em formato inesperado, a execução para com uma mensagem clara.

---

## Saídas

Todas as saídas ficam na pasta `saida_pmes`:

- `relatorio_risco_pmes.pdf`: relatório completo das duas etapas;
- `etapa1_*.csv` e `etapa2_*.csv`: tabelas de correlação, regressão, robustez, risco, projeções e sensibilidade;
- `fig_*.png` e demais gráficos: correlações, dispersão, coeficientes, distribuição do Monte Carlo, tornado e liquidez simulada;
- `auditoria_integridade.csv`: verificações automáticas da execução;
- `execucao_metadata.json`: versões das bibliotecas, semente aleatória e hash dos arquivos de entrada, para reprodutibilidade.

A semente é fixa (`SEED = 42`), então a mesma entrada gera os mesmos resultados.

---

## Configuração

Os parâmetros principais ficam no início do script (seção "0. CONFIGURAÇÃO"):

| Parâmetro | Padrão | O que faz |
|---|---|---|
| `N_SIM` | 10.000 | Número de cenários do Monte Carlo |
| `NU` | 5 | Graus de liberdade da t de Student multivariada |
| `HAC_LAGS` | 12 | Defasagens dos erros-padrão de Newey-West |
| `EMBARGO_WF` | 12 | Meses de embargo no walk-forward |
| `LAG_SELIC_ROB` | 6 | Defasagem da Selic na checagem de robustez |
| `STRESS_MILD` | Selic +2, IPCA +1,5, câmbio +15 | Choques do estresse moderado (pp) |
| `STRESS` | Selic +4, IPCA +3, câmbio +30 | Choques do estresse severo (pp) |
| `LIMITAR_FAIXA` | True | Limita estresses e sorteios à faixa histórica |
| `SUAVIZAR_PICOS` | False | Suaviza picos de 1 mês (checagem de robustez) |

---

## Limitações

- As taxas de 12 meses se sobrepõem: os 176 meses equivalem a cerca de 15 anos independentes, o que deixa as estimativas imprecisas e sensíveis a poucos episódios (2015–2016 e 2020).
- As relações estimadas são associações históricas, não causalidade.
- Os índices setoriais do IBGE medem volume do setor inteiro, não o faturamento de PMEs. Aplicá-los a uma empresa específica é uma hipótese do modelo.
- A simulação de Monte Carlo é incondicional, centrada na média de 2012–2026, e não depende do cenário atual.
- Os perfis contábeis são **fictícios** e foram gerados com apoio de inteligência artificial, porque não há demonstrativos públicos de PMEs brasileiras. Eles servem para mostrar como o modelo responde aos cenários, não para avaliar empresas reais.
- A projeção é conservadora: não considera corte de custos fixos, de Capex ou de retiradas dos sócios, ajustes que uma empresa real faria em um estresse.

---

## Autoria

**Jaques Douglas Rizzi Bueno de Assis**, Gestão Financeira, Fatec Bragança Paulista
[LinkedIn](https://www.linkedin.com/in/jaques-rizzi)

Orientador: **Prof. Alexandre Leme Sanches**

### Como citar

ASSIS, Jaques Douglas Rizzi Bueno de; SANCHES, Alexandre Leme. **Modelo econométrico integrado à simulação de Monte Carlo para análise de risco financeiro em Pequenas e Médias Empresas**. Trabalho de Graduação (Tecnologia em Gestão Financeira) – Fatec Bragança Paulista, 2026.

---

## Licença

Consulte o arquivo `LICENSE` deste repositório.
