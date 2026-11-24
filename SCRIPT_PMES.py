# -*- coding: utf-8 -*-
"""
Risco macroeconômico para PMEs brasileiras (Serviços de TI, Indústria e Comércio Varejista)
Análise de CORRELAÇÃO e REGRESSÃO em duas etapas — versão para TCC (mensal, 2012–2026)
=================================================================================================

ETAPA 1 — MACRO  →  ATIVIDADE SETORIAL (dados oficiais: BCB e IBGE/SIDRA; jan/2012 a ago/2026)
    Variável-alvo : crescimento em 12 meses do índice de VOLUME dessazonalizado (base 2022 = 100) de
                    PMS 2.1.2 (serviços de TI), PIM-PF (indústria geral) e PMC (varejo restrito).
    Regressores   : Selic (fim do mês, % a.a.), IPCA 12m (%) e variação 12m do câmbio PTAX (%) — MODELO ÚNICO.
                    O IBC-Br (PIB) NÃO entra no modelo: o índice agrega a própria atividade setorial
                    (sobreposição mecânica) e contaminaria os coeficientes e a Etapa 2. Ele é mantido apenas
                    como contexto descritivo (estatísticas, correlações e gráficos).
    1.1 Descritiva e estacionariedade (ADF)
    1.2 Correlação: Pearson e Spearman (IC por bootstrap em blocos), correlação parcial e
        correlação defasada (a política monetária atua com defasagem)
    1.3 Regressão OLS com erros Newey-West (HAC) — modelo único (Selic, IPCA, câmbio), com duas
        checagens de robustez da MESMA especificação (Selic defasada; amostra sem a COVID) — com R² dentro
        da amostra, validação cruzada em blocos e walk-forward com embargo de 12 meses
    1.4 Monte Carlo e métricas de risco (VaR / CVaR) sobre o modelo único

ETAPA 2 — TRANSMISSÃO AO CAIXA E AO BALANÇO (perfis FICTÍCIOS, em planilhas Excel ou CSV)
    2.1 Validação contábil dos perfis (o script detecta e corrige inconsistências dos CSV)
    2.2 Projeção de DRE, fluxo de caixa e balanço: cenário atual, estresse moderado e severo
    2.3 Correlação e regressão de SENSIBILIDADE sobre as simulações (lucro, FCO e liquidez corrente
        contra Selic, IPCA e câmbio) e análise tornado (±1 desvio-padrão, uma variável por vez)

IMPORTANTE (amostra): as taxas de 12 meses de meses vizinhos se sobrepõem (jan/2020 e fev/2020 compartilham
11 meses). Os n = 176 meses equivalem a cerca de 15 anos independentes. Por isso: (i) erros-padrão Newey-West
(12 defasagens) com t de Student e g.l. = anos independentes − parâmetros; (ii) IC de correlação por bootstrap
circular em blocos de 12 meses; (iii) validação cruzada retirando blocos de ±11 meses; (iv) nenhum p-valor é
reportado para resultados obtidos sobre dados simulados (Etapa 2.3).

COMO USAR NO GOOGLE COLAB
    1. Cole este arquivo em uma célula (ou use  %run pmes_tcc_colab.py ) e execute.
    2. Quando o botão "Escolher arquivos" aparecer, selecione TODOS os arquivos de uma vez:
       selic, câmbio (opcional), IPCA, IBC-Br, as 3 tabelas SIDRA (8688, 8888, 8880) e os 3 perfis
       financeiros (Excel ou CSV). Os NOMES NÃO IMPORTAM: cada arquivo é reconhecido pelo conteúdo.
    3. Se algo faltar, o script diz exatamente o quê e reabre o upload (até 3 vezes).
    4. Ao final, o PDF e um .zip com tabelas e figuras são baixados automaticamente.

SAÍDAS (pasta saida_pmes): relatorio_risco_pmes.pdf, tabelas etapa1_*.csv / etapa2_*.csv, figuras fig_*.png,
auditoria_integridade.csv e execucao_metadata.json (versões, semente e hash dos arquivos de entrada).
Não há geração de dados simulados no lugar de dados faltantes: se um arquivo faltar ou vier em formato
inesperado, o script para com uma mensagem clara.
"""
from __future__ import annotations

import csv
import datetime
import hashlib
import importlib.util
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import unicodedata
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable, Optional


def _garantir_pacote(modulo: str, pacote: Optional[str] = None) -> None:
    """Instala `pacote` via pip se `modulo` não puder ser importado (útil no Colab, onde reportlab pode faltar)."""
    try:
        __import__(modulo)
    except ImportError:
        print(f"[SETUP] Instalando '{pacote or modulo}'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pacote or modulo])


_garantir_pacote("reportlab")
_garantir_pacote("openpyxl")

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
import scipy
from scipy import stats
import openpyxl

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, Image, KeepTogether, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Silencia só o ruído de bibliotecas; os avisos do próprio script (amostra curta, datas mistas etc.) aparecem.
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", module="openpyxl")

# ==============================================================================
# 0A. ENTRADA DE ARQUIVOS — upload interativo no Google Colab
# ==============================================================================
# O usuário sobe os arquivos pelo widget do Colab, com QUALQUER nome. O script
# reconhece cada arquivo pelo CONTEÚDO (cabeçalho do CSV, abas do Excel) e só usa
# o nome como pista de reserva. Assim, "setor_ti.csv", "ti_setor (1).csv" ou
# "Perfil_Financeiro_PME_TI_v2.xlsx" funcionam sem renomear nada.


class ErroEntrada(Exception):
    """Problema com os arquivos fornecidos (ausentes, ilegíveis ou em formato inesperado)."""


def em_colab() -> bool:
    """Indica se o código está rodando dentro do Google Colab."""
    try:
        return importlib.util.find_spec("google.colab") is not None
    except (ImportError, ValueError):
        return False


EM_COLAB: bool = em_colab()
MAX_TENTATIVAS_UPLOAD: int = 3          # quantas vezes o widget de upload é reaberto se faltar arquivo
EXTENSOES_ACEITAS: tuple[str, ...] = (".csv", ".xlsx", ".xlsm", ".xls")
PASTA_ENTRADA: Path = Path("/content/entrada_pmes") if EM_COLAB else Path("entrada_pmes")


@dataclass(frozen=True)
class EntradaSpec:
    """Descrição de um arquivo de entrada esperado pelo script."""
    chave: str
    descricao: str
    obrigatoria: bool = True


ENTRADAS_SPEC: tuple[EntradaSpec, ...] = (
    EntradaSpec("selic", "Selic diária — BCB/SGS 11 (CSV)"),
    EntradaSpec("cambio", "Câmbio PTAX compra — BCB/SGS 10813 (CSV/XLSX); se faltar, é baixado da API do BCB", False),
    EntradaSpec("ipca", "IPCA mensal — IBGE (CSV)"),
    EntradaSpec("ibc", "IBC-Br — BCB (XLSX com a aba 'Atual')"),
    EntradaSpec("setor_ti", "PMS, serviços de TI — SIDRA tabela 8688 (CSV)"),
    EntradaSpec("setor_industria", "PIM-PF, indústria geral — SIDRA tabela 8888 (CSV)"),
    EntradaSpec("setor_varejo", "PMC, comércio varejista — SIDRA tabela 8880 (CSV)"),
    EntradaSpec("perfil_ti", "Perfil financeiro fictício — Serviços de TI (XLSX ou CSV)"),
    EntradaSpec("perfil_industria", "Perfil financeiro fictício — Indústria (XLSX ou CSV)"),
    EntradaSpec("perfil_varejo", "Perfil financeiro fictício — Comércio Varejista (XLSX ou CSV)"),
)
DESCRICAO_ENTRADA: dict[str, str] = {s.chave: s.descricao for s in ENTRADAS_SPEC}


def _sem_acento(txt: str) -> str:
    """Minúsculas e sem acentos (comparações robustas de rótulos)."""
    nfd = unicodedata.normalize("NFD", str(txt))
    return "".join(c for c in nfd if unicodedata.category(c) != "Mn").lower()


def _setor_pelo_texto(txt: str) -> Optional[str]:
    """Identifica o setor (ti / industria / varejo) a partir de um título ou nome de arquivo."""
    t = _sem_acento(txt)
    if "industr" in t or re.search(r"\bpim", t):
        return "industria"
    if "varejo" in t or "comercio" in t or re.search(r"\bpmc\b", t):
        return "varejo"
    if "tecnologia" in t or re.search(r"(^|[^a-z])ti([^a-z]|$)", t):
        return "ti"
    return None


def _classificar_csv(caminho: Path) -> Optional[str]:
    """Reconhece o CSV pelo cabeçalho (fonte/tabela); o nome do arquivo é só pista de reserva."""
    cab = _sem_acento(caminho.read_bytes()[:6000].decode("utf-8-sig", errors="ignore"))
    nome = _sem_acento(caminho.stem)
    if "tabela 8688" in cab:
        return "setor_ti"
    if "tabela 8888" in cab:
        return "setor_industria"
    if "tabela 8880" in cab or "tabela 8881" in cab:
        if "tabela 8881" in cab:
            warnings.warn(f"{caminho.name}: tabela 8881 (varejo ampliado). O script foi calibrado para a 8880 "
                          "(varejo restrito); se a leitura falhar, exporte a 8880.", stacklevel=2)
        return "setor_varejo"
    if "plano real" in cab or "consumidor amplo" in cab:
        return "ipca"
    if "10813" in cab[:300] or "taxa de cambio" in cab[:300]:
        return "cambio"
    if "selic" in cab[:300]:
        return "selic"
    if cab.startswith("categoria;valor"):                     # perfil em CSV (formato antigo)
        s = _setor_pelo_texto(nome)
        return f"perfil_{s}" if s else None
    if "selic" in nome:
        return "selic"
    if "cambio" in nome or "ptax" in nome:
        return "cambio"
    if "ipca" in nome:
        return "ipca"
    s = _setor_pelo_texto(nome)
    return f"setor_{s}" if s else None


def _classificar_xlsx(caminho: Path) -> Optional[str]:
    """Reconhece a planilha pelas abas: DRE+Balanço = perfil; aba 'Atual' = IBC-Br."""
    nome = _sem_acento(caminho.stem)
    try:
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    except Exception:
        return None
    try:
        abas = [_sem_acento(a) for a in wb.sheetnames]
        if "dre" in abas and any(a.startswith("balanc") for a in abas):
            ws = wb[wb.sheetnames[0]]
            titulo = " ".join(str(c) for linha in ws.iter_rows(min_row=1, max_row=3, values_only=True)
                              for c in linha if c)
            s = _setor_pelo_texto(titulo) or _setor_pelo_texto(nome)
            return f"perfil_{s}" if s else None
        if "atual" in abas or "ibc" in nome or "pib" in nome:
            return "ibc"
        if "cambio" in nome or "ptax" in nome:
            return "cambio"
    finally:
        wb.close()
    return None


def classificar_arquivos(arquivos: Iterable[Path]) -> tuple[dict[str, Path], list[str]]:
    """Associa cada arquivo a uma chave de ENTRADAS_SPEC.

    Se houver mais de um arquivo para a mesma chave (ex.: novo upload com o mesmo nome), vale o mais recente.

    Returns:
        (escolhidos, avisos): dicionário chave → caminho e lista de mensagens sobre arquivos ignorados.
    """
    por_chave: dict[str, list[Path]] = {}
    avisos: list[str] = []
    for arq in sorted(set(arquivos)):
        if arq.suffix.lower() not in EXTENSOES_ACEITAS:
            avisos.append(f"'{arq.name}': formato '{arq.suffix}' não suportado (use .csv ou .xlsx) — ignorado.")
            continue
        if arq.stat().st_size == 0:
            avisos.append(f"'{arq.name}': arquivo vazio — ignorado.")
            continue
        chave = _classificar_csv(arq) if arq.suffix.lower() == ".csv" else _classificar_xlsx(arq)
        if chave is None or chave not in DESCRICAO_ENTRADA:
            avisos.append(f"'{arq.name}': não reconheci o conteúdo — ignorado.")
            continue
        por_chave.setdefault(chave, []).append(arq)
    escolhidos: dict[str, Path] = {}
    for chave, lista in por_chave.items():
        lista.sort(key=lambda p: p.stat().st_mtime)
        escolhidos[chave] = lista[-1]
        if len(lista) > 1:
            avisos.append(f"{DESCRICAO_ENTRADA[chave]}: {len(lista)} arquivos enviados; uso o mais recente "
                          f"('{lista[-1].name}') e ignoro {[p.name for p in lista[:-1]]}.")
    return escolhidos, avisos


def _candidatos() -> list[Path]:
    """Lista os arquivos candidatos nas pastas de busca (sem recursão)."""
    pastas = [PASTA_ENTRADA]
    if EM_COLAB:
        pastas.append(Path("/content"))                       # arquivos arrastados para o painel lateral
    elif os.environ.get("DATA_DIR"):
        pastas.append(Path(os.environ["DATA_DIR"]))
    else:
        pastas.append(Path("."))
    return [p for pasta in pastas if pasta.is_dir() for p in pasta.iterdir() if p.is_file()]


def _pedir_upload() -> int:
    """Abre o widget de upload do Colab e grava os arquivos em PASTA_ENTRADA. Devolve quantos chegaram."""
    from google.colab import files  # type: ignore[import-not-found]
    enviados = files.upload()                                 # dict {nome: bytes}; vazio se o usuário cancelar
    for nome, conteudo in enviados.items():
        (PASTA_ENTRADA / Path(nome).name).write_bytes(conteudo)
        solta = Path.cwd() / nome                             # o widget também grava no diretório atual
        if solta.is_file() and solta.resolve() != (PASTA_ENTRADA / Path(nome).name).resolve():
            solta.unlink()
    return len(enviados)


def _relatorio_entradas(escolhidos: dict[str, Path], avisos: list[str]) -> list[EntradaSpec]:
    """Imprime o que foi reconhecido e devolve as entradas obrigatórias que ainda faltam."""
    print("\n[ENTRADAS] Arquivos reconhecidos:")
    faltando: list[EntradaSpec] = []
    for spec in ENTRADAS_SPEC:
        if spec.chave in escolhidos:
            print(f"   [ok]   {spec.descricao}\n          ← {escolhidos[spec.chave].name}")
        elif spec.obrigatoria:
            print(f"   [FALTA] {spec.descricao}")
            faltando.append(spec)
        else:
            print(f"   [opc.]  {spec.descricao} — não enviado")
    for a in avisos:
        print(f"   [aviso] {a}")
    return faltando


def coletar_entradas() -> dict[str, Path]:
    """Obtém e valida os arquivos de entrada.

    No Colab, abre o widget de upload (até MAX_TENTATIVAS_UPLOAD vezes, enquanto faltar algo obrigatório).
    Fora do Colab, procura os arquivos em DATA_DIR (ou na pasta atual).

    Raises:
        ErroEntrada: se, ao final, ainda faltar algum arquivo obrigatório.
    """
    PASTA_ENTRADA.mkdir(parents=True, exist_ok=True)
    tentativa = 0
    while True:
        if EM_COLAB:
            print("\n[UPLOAD] Selecione TODOS os arquivos de uma vez (CSV/XLSX; os nomes não importam)." if tentativa == 0
                  else "\n[UPLOAD] Envie os arquivos que faltam (os já enviados são mantidos).")
            n = _pedir_upload()
            if n == 0:
                print("[UPLOAD] Nenhum arquivo enviado; verifico o que já existe na sessão.")
        escolhidos, avisos = classificar_arquivos(_candidatos())
        faltando = _relatorio_entradas(escolhidos, avisos)
        if not faltando:
            return escolhidos
        tentativa += 1
        if not EM_COLAB or tentativa >= MAX_TENTATIVAS_UPLOAD:
            lista = "\n  - ".join(s.descricao for s in faltando)
            raise ErroEntrada(
                "Faltam arquivos obrigatórios:\n  - " + lista +
                ("\nRode a célula novamente e envie os arquivos pelo botão 'Escolher arquivos'." if EM_COLAB
                 else f"\nColoque-os em {os.environ.get('DATA_DIR') or 'a pasta atual'} e rode de novo."))


def ler_entrada(chave: str, leitor: Callable[..., Any], *args: Any) -> Any:
    """Lê `ENTRADAS[chave]` com `leitor`, convertendo qualquer falha em uma mensagem clara para o usuário."""
    caminho = ENTRADAS[chave]
    try:
        return leitor(caminho, *args)
    except ErroEntrada:
        raise
    except Exception as exc:
        raise ErroEntrada(
            f"O arquivo '{caminho.name}' foi reconhecido como «{DESCRICAO_ENTRADA[chave]}», mas não pôde ser lido "
            f"({type(exc).__name__}: {exc}). Confira se é o arquivo bruto exportado da fonte, sem edições.") from exc


# Rótulos (sem acento, minúsculos) das linhas das planilhas de perfil → campos do dicionário do perfil.
# (aba, prefixo do rótulo, obrigatório)
ROTULOS_XLSX: dict[str, tuple[str, str, bool]] = {
    "R0": ("dre", "receita liquida", True), "C0": ("dre", "(-) custo total", True),
    "fixo": ("dre", "fracao custo fixo", True), "juros": ("dre", "(-) despesa financeira", True),
    "L0": ("dre", "lucro liquido", True),
    "caixa": ("balanc", "caixa", True), "cr": ("balanc", "contas a receber", True), "est": ("balanc", "estoques", True),
    "oac": ("balanc", "outros ativos", False), "anc": ("balanc", "ativo nao circulante", True),
    "ativo": ("balanc", "total do ativo", True), "forn": ("balanc", "fornecedores", True),
    "outros": ("balanc", "outras obrigacoes", True), "dcp": ("balanc", "divida de curto", True),
    "dlp": ("balanc", "divida de longo", True), "pl": ("balanc", "patrimonio liquido", True),
    "passivo": ("balanc", "total passivo +", False), "LC": ("indicadores", "liquidez corrente", True),
    "indexada": ("indicadores", "% divida", True),
}


def ler_perfil_xlsx(caminho: Path | str) -> dict[str, float]:
    """Lê uma planilha de perfil (abas DRE, Balanço Patrimonial e Indicadores) e devolve o dicionário do perfil.

    Os valores ficam em R$ (como nos CSV antigos). Localiza as linhas pelo RÓTULO, não pela posição,
    então reordenar linhas na planilha não quebra a leitura.

    Raises:
        ValueError: aba ou linha obrigatória ausente, ou célula sem valor calculado (salve a planilha no Excel).
    """
    wb = openpyxl.load_workbook(caminho, data_only=True)
    try:
        linhas_por_aba: dict[str, list[tuple[str, float]]] = {}
        for prefixo in ("dre", "balanc", "indicadores"):
            nome_aba = next((n for n in wb.sheetnames if _sem_acento(n).startswith(prefixo)), None)
            if nome_aba is None:
                raise ValueError(f"aba '{prefixo}…' ausente (abas encontradas: {wb.sheetnames})")
            pares: list[tuple[str, float]] = []
            for linha in wb[nome_aba].iter_rows(values_only=True):
                if len(linha) > 2 and isinstance(linha[1], str) and isinstance(linha[2], (int, float)) \
                        and not isinstance(linha[2], bool):
                    pares.append((_sem_acento(linha[1]).replace("•", "").strip(), float(linha[2])))
            linhas_por_aba[prefixo] = pares
    finally:
        wb.close()
    out: dict[str, float] = {}
    ausentes: list[str] = []
    for campo, (aba, prefixo_rotulo, obrigatorio) in ROTULOS_XLSX.items():
        achou = next((v for rot, v in linhas_por_aba[aba] if rot.startswith(prefixo_rotulo)), None)
        if achou is None:
            if obrigatorio:
                ausentes.append(f"{aba}: «{prefixo_rotulo}…»")
            else:
                out[campo] = 0.0
        else:
            out[campo] = achou
    if ausentes:
        raise ValueError("linhas obrigatórias sem valor numérico: " + "; ".join(ausentes) +
                         " (se a célula é fórmula, abra a planilha no Excel e salve uma vez)")
    out.setdefault("passivo", 0.0)
    return out


# ==============================================================================
# 0. CONFIGURAÇÃO
# ==============================================================================
SEED = 42
N_SIM = 10_000
NU = 5                    # graus de liberdade da t-Student multivariada do Monte Carlo
INI_AMOSTRA = "2012-01"
HAC_LAGS = 12
BLOCO_CV = 12             # bloco do bootstrap dos resíduos (Monte Carlo)
MIN_TREINO_WF = 60        # mínimo de meses de treino antes da 1ª previsão walk-forward
EMBARGO_WF = 12           # meses descartados entre treino e alvo (taxas de 12m se sobrepõem)
H_CV_BLOCO = 11           # validação cruzada em blocos: retira ±11 meses em torno do mês previsto
N_BOOT_CORR = 2000        # reamostragens do bootstrap de correlação
BLOCO_BOOT_CORR = 12
LAG_SELIC_ROB = 6          # defasagem (meses) da Selic na checagem de robustez (fixada a priori)
LAGS_GRADE = (0, 3, 6, 9, 12)
COVID_INI, COVID_FIM = "2020-03", "2021-06"   # janela excluída na checagem de robustez 'sem COVID'
LIMITAR_FAIXA = True      # True: estresses e sorteios do Monte Carlo ficam dentro do mínimo e do máximo históricos da amostra (sem extrapolação)
SUAVIZAR_PICOS = False    # True: troca picos de 1 mês (que se revertem no seguinte) pela média dos vizinhos
BAIXAR_NO_COLAB = True    # baixa o PDF (e o .zip) automaticamente ao final, se estiver no Google Colab
EMPACOTAR_SAIDAS = True   # cria resultados_pmes.zip com PDF, tabelas e figuras
OUT_DIR = os.environ.get("OUT_DIR") or ("/content/saida_pmes" if EM_COLAB else "saida_pmes")
os.makedirs(OUT_DIR, exist_ok=True)

LAGS = {"Selic": 0, "IPCA": 0, "Cambio_var": 0, "PIB": 0}   # defasagens do modelo (PIB só descritivo)
# Dois cenários de estresse: um dentro da faixa histórica e outro mais severo (extrapolado)
STRESS_MILD = {"Selic": +2.0, "IPCA": +1.5, "Cambio_var": +15.0}
STRESS = {"Selic": +4.0, "IPCA": +3.0, "Cambio_var": +30.0}

# MODELO ÚNICO: regressoras = Selic, IPCA e câmbio (juros, preços e câmbio).
# O PIB (IBC-Br) fica FORA do modelo, dos cenários e das simulações: o índice agrega a própria atividade
# setorial (endogeneidade mecânica). Ele aparece apenas como contexto descritivo (DESC_COLS).
MACRO_COLS = ["Selic", "IPCA", "Cambio_var"]
DESC_COLS = MACRO_COLS + ["PIB"]
MACRO_LABEL = {"Selic": "Selic", "IPCA": "IPCA", "Cambio_var": "Câmbio", "PIB": "PIB (contexto)"}
# chave: arquivo do setor em ENTRADAS; filtro: coluna de interesse na tabela SIDRA; perfil: chave do perfil financeiro
SETORES: dict[str, dict[str, Any]] = {
    "Serviços de TI": dict(chave="setor_ti", filtro=None, perfil="perfil_ti"),
    "Indústria": dict(chave="setor_industria", filtro="indústria geral", perfil="perfil_industria"),
    "Comércio Varejista": dict(chave="setor_varejo", filtro="volume", perfil="perfil_varejo"),
}

print("=" * 78)
print("Risco macroeconômico para PMEs — correlação, regressão e impacto em caixa")
print("=" * 78)
ENTRADAS: dict[str, Path] = coletar_entradas()      # no Colab: abre o widget de upload
print(f"[INFO] Saída em: {os.path.abspath(OUT_DIR)}")

# ==============================================================================
# 1. LEITURA E PADRONIZAÇÃO DOS DADOS (todos mensais; taxas de 12 meses em %)
# ==============================================================================
MESES = {"janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5,
         "junho": 6, "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10,
         "novembro": 11, "dezembro": 12}

MES_NOME_RE = r"(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)"


def ler_texto(caminho: Path | str) -> list[str]:
    """Lê um arquivo de texto (UTF-8 com BOM ou Latin-1) e devolve as linhas não vazias, sem espaços nas pontas."""
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(caminho, encoding=enc) as f:
                return [l.strip() for l in f if l.strip()]
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Não consegui decodificar {caminho}")


def ler_ipca_mensal(caminho: Path | str) -> pd.Series:
    """Arquivo IBGE em formato largo: 1 linha de meses e 1 linha 'Brasil;...'."""
    linhas = ler_texto(caminho)
    cab = next(l for l in linhas if re.search(MES_NOME_RE + r"\s+\d{4}", l, re.I))
    val = next(l for l in linhas if l.lower().startswith("brasil"))
    meses = [x.strip().lower() for x in cab.split(";")[1:]]
    vals = [x.strip() for x in val.split(";")[1:]]
    if len(meses) != len(vals):
        warnings.warn(f"{os.path.basename(caminho)}: {len(meses)} meses no cabeçalho e {len(vals)} valores; "
                      "pareando pelas primeiras posições.", stacklevel=2)
    idx, nums = [], []
    for m, v in zip(meses, vals):          # filtra mês e valor JUNTOS (evita desalinhamento)
        g = re.match(r"([a-zçãáéíóúâêô]+)\s+(\d{4})", m)
        if g is None or g.group(1) not in MESES:
            continue
        idx.append(pd.Timestamp(int(g.group(2)), MESES[g.group(1)], 1))
        nums.append(v.replace(",", "."))
    s = pd.Series(pd.to_numeric(nums, errors="coerce"), index=idx).dropna()
    if s.empty:
        raise ValueError(f"{os.path.basename(caminho)}: nenhum valor de IPCA lido.")
    return s


def ler_ibc_mensal(caminho: Path | str) -> pd.Series:
    """Aba 'Atual' do IBC-Br (não dessazonalizado, coluna B), série mensal em nível."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb["Atual"] if "Atual" in wb.sheetnames else wb[wb.sheetnames[0]]
    datas, vals = [], []
    formulas_sem_cache = 0
    for row in ws.iter_rows(values_only=True):
        if row is None or len(row) < 2:
            continue
        d, v = row[0], row[1]
        if isinstance(d, (datetime.datetime, datetime.date)) and isinstance(v, (int, float)):
            datas.append(pd.Timestamp(d))
            vals.append(float(v))
        elif isinstance(d, str) and d.lstrip().startswith("="):
            formulas_sem_cache += 1
    if not datas:
        raise ValueError(
            f"{os.path.basename(caminho)}: nenhuma data+número encontrada na aba Atual. "
            "Abra o arquivo no Excel e salve uma vez para popular o cache das fórmulas."
        )
    if formulas_sem_cache > 0:
        warnings.warn(
            f"{os.path.basename(caminho)}: {formulas_sem_cache} célula(s) '=EDATE(...)' sem valor em cache. "
            "O IBC-Br pode parar antes do mês mais recente — abra o arquivo no Excel e salve uma vez.",
            stacklevel=2,
        )
    return pd.Series(vals, index=pd.DatetimeIndex(datas)).sort_index()


def _parse_datas_coluna(strs: Iterable[Any]) -> list[tuple[pd.Timestamp, str]]:
    """Retorna lista de (Timestamp|NaT, formato) onde formato ∈ {'iso', 'br', 'outro'}."""
    out = []
    for s in strs:
        s = str(s).strip().strip('"')
        if re.match(r"^\d{4}-\d{2}-\d{2}", s):
            try:
                y, m, d = map(int, s[:10].split("-"))
                out.append((pd.Timestamp(y, m, d), "iso"))
                continue
            except Exception:
                pass
        if re.match(r"^\d{2}/\d{2}/\d{4}", s):
            out.append((pd.to_datetime(s, dayfirst=True, errors="coerce"), "br"))
            continue
        out.append((pd.to_datetime(s, dayfirst=True, errors="coerce"), "outro"))
    return out


def ler_bcb_diario(caminho: Path | str, nome: str) -> pd.Series:
    """CSV bruto do BCB ('dd/mm/aaaa;valor' ou '"data";"valor"') OU XLSX com colunas (Data, Valor).

    Regras:
      - Vírgula decimal normalizada.
      - Se o arquivo MISTURA 'dd/mm/aaaa' e 'yyyy-mm-dd', assume que as ISO foram corrompidas
        (dd/mm exportado errado como mm/dd) e faz swap dia↔mês nas ISO com dia ≤ 12.
      - Se a mediana dos valores > 100, divide por 10000 (u.m.c./US$ → R$/US$).
    """
    ext = os.path.splitext(caminho)[1].lower()

    if ext in (".xlsx", ".xls"):
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"
        try:
            df = pd.read_excel(caminho, header=None, engine=engine)
        except ImportError as exc:
            raise ImportError(
                "Para ler .xls antigo, instale xlrd (pip install xlrd) ou salve como .xlsx."
            ) from exc
        if df.shape[1] < 2:
            raise ValueError(f"{os.path.basename(caminho)}: planilha precisa de 2 colunas (Data, Valor).")
        col0 = df[0].astype(str).tolist()
        parsed = _parse_datas_coluna(col0)

        def _numero_excel(v):
            if pd.isna(v):
                return np.nan
            if isinstance(v, (int, float, np.number)):
                return float(v)
            z = str(v).strip().replace(" ", "")
            if "," in z:
                z = z.replace(".", "").replace(",", ".")
            return pd.to_numeric(z, errors="coerce")

        vals = df[1].map(_numero_excel).to_numpy(dtype=float)
        datas_brutas = [dt for dt, _ in parsed]
        formatos = [fmt for _, fmt in parsed]
    else:
        linhas = ler_texto(caminho)[1:]
        strs, vals_list = [], []
        for l in linhas:
            p = l.split(";")
            if len(p) < 2:
                continue
            strs.append(p[0])
            v = p[1].strip().strip('"').replace(",", ".")
            try:
                vals_list.append(float(v))
            except ValueError:
                vals_list.append(np.nan)
        parsed = _parse_datas_coluna(strs)
        datas_brutas = [dt for dt, _ in parsed]
        formatos = [fmt for _, fmt in parsed]
        vals = np.array(vals_list, dtype=float)

    n_iso = sum(1 for f in formatos if f == "iso")
    n_br = sum(1 for f in formatos if f == "br")
    misto = (n_iso > 0) and (n_br > 0)

    datas = []
    swaps = 0
    for dt, fmt in zip(datas_brutas, formatos):
        if pd.isna(dt):
            datas.append(pd.NaT)
            continue
        if misto and fmt == "iso" and dt.day <= 12:
            dt = pd.Timestamp(dt.year, dt.day, dt.month)
            swaps += 1
        datas.append(dt)

    df_serie = pd.Series(vals, index=pd.DatetimeIndex(datas), name=nome).dropna().sort_index()
    if df_serie.empty:
        raise ValueError(f"{os.path.basename(caminho)}: nenhum valor válido para '{nome}'.")

    if misto:
        warnings.warn(
            f"{os.path.basename(caminho)}: datas em formato misto (ISO={n_iso}, BR={n_br}). "
            f"Fiz swap dia↔mês em {swaps} linha(s) ISO. Verifiquei que as datas ficam monotônicas?",
            stacklevel=2,
        )
    # sanidade: monotonicidade
    if not df_serie.index.is_monotonic_increasing:
        warnings.warn(
            f"{os.path.basename(caminho)}: datas não ficaram monotônicas após o parse — "
            "possível corrupção de formato não coberta. Verifique manualmente.",
            stacklevel=2,
        )

    # escala: PTAX u.m.c./US$ → R$/US$
    mediana = df_serie.median()
    if mediana > 100:
        df_serie = df_serie / 10000.0
    return df_serie


def ler_setor(caminho: Path | str, nome: str) -> pd.Series:
    """CSV de setor (Ano;Mes;Indice_SA;...): devolve o índice de volume dessazonalizado, por mês (PeriodIndex)."""
    d = pd.read_csv(caminho, sep=";", encoding="utf-8-sig")
    for c in ("Ano", "Mes", "Indice_SA"):
        if c not in d.columns:
            raise ValueError(f"{os.path.basename(caminho)}: coluna '{c}' ausente (colunas: {list(d.columns)})")
    idx = pd.to_datetime(dict(year=d["Ano"], month=d["Mes"], day=1)).dt.to_period("M")
    s = pd.Series(pd.to_numeric(d["Indice_SA"], errors="coerce").values, index=pd.PeriodIndex(idx), name=nome).sort_index()
    if s.index.duplicated().any():
        raise ValueError(f"{os.path.basename(caminho)}: meses duplicados.")
    return s.dropna()


def para_periodo(s: pd.Series) -> pd.Series:
    """Converte o índice de datas (DatetimeIndex) para períodos mensais (PeriodIndex)."""
    s = s.copy()
    s.index = s.index.to_period("M")
    return s


def yoy(s: pd.Series) -> pd.Series:
    """Variação % sobre o mesmo mês do ano anterior (reindexa para meses contíguos)."""
    s = s.sort_index()
    cheio = pd.period_range(s.index.min(), s.index.max(), freq="M")
    return s.reindex(cheio).pct_change(12, fill_method=None) * 100


def fim_de_mes(serie: pd.Series) -> pd.Series:
    """Último valor de cada mês; descarta o último mês se estiver incompleto (>5 dias antes do fim do mês)."""
    per = serie.index.to_period("M")
    ult = serie.groupby(per).last()
    if serie.index[-1] < per[-1].end_time.normalize() - pd.Timedelta(days=5):
        ult = ult.iloc[:-1]
    return ult


def valor_em(serie: pd.Series, data: pd.Timestamp) -> float:
    """Último valor disponível até a data (trata fins de semana e feriados)."""
    s = serie[serie.index <= data]
    if s.empty:
        raise ValueError(f"Sem dados de '{serie.name}' até {data}")
    return float(s.iloc[-1])



def _num_br(x: Any) -> float:
    """Número no formato brasileiro do SIDRA ('1.234,56'); símbolos de ausência viram NaN."""
    x = str(x).strip().replace('"', "")
    if x in ("", "-", "..", "...", "X"):
        return np.nan
    return pd.to_numeric(x.replace(".", "").replace(",", "."), errors="coerce")


def ler_setor_sidra(caminho: Path | str, filtro_col: Optional[str] = None) -> pd.Series:
    """Tabela SIDRA bruta (PMS 8688, PIM-PF 8888, PMC 8880): extrai o bloco 'com ajuste sazonal' e o índice de VOLUME.

    O arquivo traz dois blocos (sem e com ajuste sazonal). Para PMC a linha de cabeçalho alterna
    'receita nominal'/'volume' (filtro 'volume'); para PIM-PF há uma linha de atividades (filtro 'indústria geral');
    para PMS há duas linhas de dados (receita e volume) e escolhe-se a de volume.
    """
    with open(caminho, encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh, delimiter=";"))
    i0 = next((i for i, r in enumerate(rows) if r and "ajuste sazonal" in r[0].lower()), None)
    if i0 is None:
        raise ValueError(f"{os.path.basename(caminho)}: bloco 'com ajuste sazonal' não encontrado.")
    i1 = next(i for i in range(i0, len(rows)) if rows[i] and rows[i][0].startswith("Fonte"))
    bloco = rows[i0 + 1:i1]
    ih = next(k for k, r in enumerate(bloco)
              if any(re.match(MES_NOME_RE + r"\s+\d{4}$", c.strip().lower()) for c in r))
    cur, meses = None, []
    for c in bloco[ih]:                       # meses podem vir a cada 2 colunas (PMC): repete até o próximo
        g = re.match(MES_NOME_RE + r"\s+(\d{4})$", c.strip().lower())
        if g:
            cur = pd.Period(year=int(g.group(2)), month=MESES[g.group(1)], freq="M")
        elif c.strip() != "":
            cur = None                        # rótulo descritivo no início da linha
        meses.append(cur)
    resto = bloco[ih + 1:]
    larg = max(len(r) for r in resto if r)
    meses += [cur] * (larg - len(meses))      # o SIDRA pode omitir a última célula em branco do cabeçalho
    dados = [r for r in resto if r and r[0].strip() == "Brasil"]
    tipos = [r for r in resto if r and r[0].strip() == ""]
    if len(dados) > 1:
        dados = [r for r in dados if any("volume" in c.lower() for c in r[:4])]
    if not dados:
        raise ValueError(f"{os.path.basename(caminho)}: linha de dados 'Brasil' não encontrada.")
    d = dados[0]
    mask = [m is not None for m in meses]
    if tipos and filtro_col:
        t = tipos[0]
        mask = [m and (j < len(t) and filtro_col.lower() in t[j].lower()) for j, m in enumerate(mask)]
    idx = [meses[j] for j in range(len(meses)) if mask[j]]
    val = [_num_br(d[j]) if j < len(d) else np.nan for j in range(len(meses)) if mask[j]]
    s = pd.Series(val, index=pd.PeriodIndex(idx, freq="M")).sort_index()
    if s.index.duplicated().any():
        raise ValueError(f"{os.path.basename(caminho)}: meses duplicados após o filtro '{filtro_col}'.")
    return s.dropna()


def carregar_setor(nome: str) -> pd.Series:
    """Aceita a tabela SIDRA bruta ou o formato antigo (Ano;Mes;Indice_SA)."""
    cfg = SETORES[nome]
    caminho = ENTRADAS[cfg["chave"]]
    try:
        cab = ler_texto(caminho)[0]
        s = ler_setor(caminho, nome) if "Indice_SA" in cab else ler_setor_sidra(caminho, cfg["filtro"])
    except Exception as exc:
        raise ErroEntrada(f"O arquivo '{caminho.name}' foi reconhecido como «{DESCRICAO_ENTRADA[cfg['chave']]}», mas não "
                          f"pôde ser lido ({type(exc).__name__}: {exc}). Exporte novamente a tabela bruta do SIDRA "
                          "(com os blocos 'sem' e 'com ajuste sazonal').") from exc
    return s.rename(nome)


# ------------------------------------------------------------------------------
# Perfis financeiros fictícios (Etapa 2)
# ------------------------------------------------------------------------------
ROTULOS_PERFIL = {
    "R0": "DRE - Receita Líquida", "C0": "DRE - Custo Total + Despesas Operacionais",
    "fixo": "DRE - Fração Custo Fixo", "juros": "DRE - Despesa Financeira",
    "L0": "DRE - Lucro Líquido", "caixa": "BP - Caixa e Equivalentes", "cr": "BP - Contas a Receber",
    "est": "BP - Estoques", "oac": "BP - Outros Ativos Circulantes", "anc": "BP - Ativo Não Circulante",
    "ativo": "BP - TOTAL DO ATIVO", "forn": "BP - Fornecedores", "outros": "BP - Outras Obrigações",
    "dcp": "BP - Dívida de Curto Prazo", "dlp": "BP - Dívida de Longo Prazo", "pl": "BP - Patrimônio Líquido",
    "passivo": "BP - TOTAL PASSIVO", "LC": "Indicadores - Liquidez Corrente",
    "indexada": "Indicadores - Parcela da Dívida Indexada",
}


def ler_perfil(caminho: Path | str) -> dict[str, float]:
    """Lê o perfil financeiro e devolve um dict com os campos de ROTULOS_PERFIL.

    Aceita planilha Excel (abas DRE, Balanço Patrimonial e Indicadores) ou o CSV antigo
    'Categoria;Valor;Unidade_Contexto;Notas'.
    """
    if Path(caminho).suffix.lower() in (".xlsx", ".xlsm"):
        return ler_perfil_xlsx(caminho)
    d = pd.read_csv(caminho, sep=";", encoding="utf-8-sig")
    d["Categoria"] = d["Categoria"].astype(str).str.strip()
    d["Valor"] = pd.to_numeric(d["Valor"], errors="coerce")
    out = {}
    for chave, prefixo in ROTULOS_PERFIL.items():
        achou = d[d["Categoria"].str.startswith(prefixo)]
        if achou.empty:
            raise ValueError(f"{os.path.basename(caminho)}: linha '{prefixo}...' ausente.")
        out[chave] = float(achou["Valor"].iloc[0])
    return out


def validar_perfil(nome: str, p: dict[str, float], tol: float = 0.01) -> tuple[dict[str, float], list[dict[str, Any]]]:
    """Confere as identidades contábeis e CORRIGE o que for inconsistente (registrando em `log`).

    Identidades: Lucro = Receita − Custos − Juros; Ativo = soma dos grupos; Passivo + PL = Ativo; LC = AC / PC.
    O erro de digitação mais comum é de escala (ex.: 1.500.000 em vez de 150.000 de juros) ou PL que não fecha o balanço.
    """
    p = dict(p)
    log = []

    def registrar(item, informado, calculado, acao):
        log.append(dict(setor=nome, item=item, informado=informado, recalculado=calculado, acao=acao))

    juros_impl = p["R0"] - p["C0"] - p["L0"]
    if abs(juros_impl - p["juros"]) > tol * p["R0"]:
        registrar("Despesa financeira", p["juros"], juros_impl, "usado o valor recalculado (Receita − Custos − Lucro)")
        p["juros"] = juros_impl
    else:
        registrar("Despesa financeira", p["juros"], juros_impl, "ok")

    ativo_impl = p["caixa"] + p["cr"] + p["est"] + p["oac"] + p["anc"]
    ok = abs(ativo_impl - p["ativo"]) <= tol * p["ativo"]
    registrar("Total do ativo", p["ativo"], ativo_impl, "ok" if ok else "usado o valor recalculado")
    p["ativo"] = ativo_impl if not ok else p["ativo"]

    passivo_ext = p["forn"] + p["outros"] + p["dcp"] + p["dlp"]
    pl_impl = p["ativo"] - passivo_ext
    ok = abs(pl_impl - p["pl"]) <= tol * p["ativo"]
    registrar("Patrimônio líquido", p["pl"], pl_impl, "ok" if ok else "usado o valor recalculado (Ativo − Passivo exigível)")
    p["pl"] = pl_impl if not ok else p["pl"]
    p["passivo"] = passivo_ext + p["pl"]

    ac = p["caixa"] + p["cr"] + p["est"] + p["oac"]
    pc = p["forn"] + p["outros"] + p["dcp"]
    lc_impl = ac / pc
    ok = abs(lc_impl - p["LC"]) <= 0.02 * p["LC"]
    registrar("Liquidez corrente", p["LC"], lc_impl, "ok" if ok else "usado o valor recalculado")
    p["LC"] = lc_impl
    p["margem"] = p["L0"] / p["R0"]
    return p, log


# ==============================================================================
# 1B. SÉRIES MACRO E SETORIAIS
# ==============================================================================
selic_d = ler_entrada("selic", ler_bcb_diario, "Selic")
selic_aa = ((1 + selic_d / 100) ** 252 - 1) * 100

def baixar_ptax_bcb(codigo: int = 10813, ini: str = "01/01/2010") -> pd.Series:
    """Baixa o câmbio diário (PTAX compra) da API SGS do BCB, em blocos de 9 anos (limite da API)."""
    import urllib.request
    fim = datetime.date.today()
    partes, ano = [], int(ini[-4:])
    while ano <= fim.year:
        a, b = f"01/01/{ano}", f"31/12/{min(ano + 8, fim.year)}"
        url = (f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados"
               f"?formato=json&dataInicial={a}&dataFinal={b}")
        with urllib.request.urlopen(url, timeout=60) as r:
            partes += json.loads(r.read().decode("utf-8"))
        ano += 9
    s = pd.Series({pd.to_datetime(x["data"], dayfirst=True): float(x["valor"]) for x in partes if x["valor"] != ""})
    s = s.sort_index().rename("Cambio")
    if s.empty:
        raise ValueError("API do BCB devolveu série de câmbio vazia.")
    print(f"[INFO] Câmbio baixado da API do BCB (SGS {codigo}): {s.index.min():%d/%m/%Y} a {s.index.max():%d/%m/%Y}")
    return s


if "cambio" in ENTRADAS:
    cambio_d = ler_entrada("cambio", ler_bcb_diario, "Cambio")
else:
    print("[AVISO] Arquivo de câmbio não enviado; tentando baixar a PTAX direto da API do BCB...")
    try:
        cambio_d = baixar_ptax_bcb()
    except Exception as e:
        raise ErroEntrada("Sem arquivo de câmbio e a API do BCB não respondeu "
                          f"({e}). Envie o câmbio PTAX (SGS 10813) e rode de novo.") from e
if cambio_d.median() > 50:
    raise ValueError(
        "Câmbio fora de escala (mediana > 50 R$/US$). Verifique se a conversão u.m.c.→R$ aplicou corretamente."
    )

print(f"[INFO] Câmbio (PTAX compra): {cambio_d.index.min():%d/%m/%Y} a {cambio_d.index.max():%d/%m/%Y} | "
      f"{len(cambio_d)} dias | último valor R$ {cambio_d.iloc[-1]:.4f}")
selic_m = fim_de_mes(selic_aa).rename("Selic")
cambio_m = fim_de_mes(cambio_d)
cambio_var = yoy(cambio_m).rename("Cambio_var")

ipca_mens = ler_entrada("ipca", ler_ipca_mensal)
ultimo_mes_ipca = ipca_mens.index[-1]
ipca_p = para_periodo(ipca_mens)
ipca_cheio = ipca_p.reindex(pd.period_range(ipca_p.index.min(), ipca_p.index.max(), freq="M"))
ipca_12m = ((1 + ipca_cheio / 100).rolling(12).apply(np.prod, raw=True) - 1) * 100
ipca_12m = ipca_12m.rename("IPCA")

ibc_mens = ler_entrada("ibc", ler_ibc_mensal)
ultimo_mes_ibc = ibc_mens.index[-1]
pib_12m = yoy(para_periodo(ibc_mens)).rename("PIB")

setor_y = {nome: carregar_setor(nome) for nome in SETORES}


def detectar_picos(s: pd.Series, k: float = 4.0) -> list[pd.Period]:
    """Meses com salto > k·σ robusto (MAD) que reverte no mês seguinte: típico de ruído, não de ciclo."""
    mm = s.pct_change() * 100
    sig = 1.4826 * (mm - mm.median()).abs().median()
    return [s.index[t] for t in range(1, len(s) - 1)
            if abs(mm.iloc[t]) > k * sig and np.sign(mm.iloc[t]) != np.sign(mm.iloc[t + 1])
            and abs(mm.iloc[t + 1]) > 0.5 * abs(mm.iloc[t])]


def suavizar_picos(s: pd.Series, picos: list[pd.Period]) -> pd.Series:
    """Substitui cada pico de 1 mês pela média dos meses vizinhos (checagem de robustez; desligada por padrão)."""
    s = s.copy()
    for p in picos:
        s.loc[p] = (s.loc[p - 1] + s.loc[p + 1]) / 2
    return s


for _nome, _s in list(setor_y.items()):
    _picos = detectar_picos(_s)
    if _picos:
        _lst = ", ".join(p.strftime("%m/%Y") for p in _picos)
        warnings.warn(f"{_nome}: {len(_picos)} pico(s) de 1 mês revertido(s) no índice dessazonalizado ({_lst}). "
                      "Eles viram outliers nas taxas de 12 meses (no mês e 12 meses depois). "
                      "Teste SUAVIZAR_PICOS = True como checagem de robustez.", stacklevel=2)
        if SUAVIZAR_PICOS:
            setor_y[_nome] = suavizar_picos(_s, _picos)
setor_12m = {nome: yoy(s).rename(nome) for nome, s in setor_y.items()}

macro_base = pd.concat([selic_m, ipca_12m, cambio_var, pib_12m], axis=1)[DESC_COLS]
macro = macro_base.copy()
for k, lag in LAGS.items():
    if lag:
        macro[k] = macro[k].shift(lag)

sectors = pd.concat(setor_12m.values(), axis=1)
df_total = pd.concat([macro, sectors], axis=1)
df_periodo = df_total[df_total.index >= pd.Period(INI_AMOSTRA, "M")]
df = df_periodo.dropna()

# Diagnóstico de meses descartados
_esperados = pd.period_range(df_periodo.index.min(), df_periodo.index.max(), freq="M")
_faltando = sorted(set(_esperados) - set(df.index))
if _faltando:
    print(f"[INFO] Meses descartados por NaN no período da amostra: {len(_faltando)} "
          f"({_faltando[0]} ... {_faltando[-1]})")
n_obs = len(df)
if n_obs < 84:
    warnings.warn(f"Amostra curta: n = {n_obs} (~{n_obs/12:.1f} anos efetivos). "
                  "Com 4 parâmetros por modelo, os graus de liberdade ficam muito baixos.", stacklevel=2)
if n_obs < 60:
    raise ErroEntrada(f"Amostra muito curta (n = {n_obs} meses; mínimo 60). Verifique se os arquivos cobrem 2012–2026.")
_cheio = pd.period_range(df.index.min(), df.index.max(), freq="M")
if len(_cheio) != n_obs:
    raise ErroEntrada(
        f"A amostra tem {len(_cheio) - n_obs} mês(es) faltando (buracos no meio). "
        "Verifique se todas as séries cobrem o período sem lacunas."
    )

nomes_setores = list(SETORES.keys())
MES_INI, MES_FIM = df.index.min(), df.index.max()
MES_INI_D, MES_FIM_D = MES_INI.to_timestamp(), MES_FIM.to_timestamp()
ANO_INI, ANO_FIM = MES_INI.year, MES_FIM.year
LAGS_HAC = HAC_LAGS
ANOS_EF = n_obs / 12.0

print(f"[INFO] Amostra de estimação: {MES_INI_D:%m/%Y}–{MES_FIM_D:%m/%Y} (n = {n_obs} meses ≈ {ANOS_EF:.1f} anos efetivos)")
print(f"[INFO] IPCA até {ultimo_mes_ipca:%m/%Y}; IBC-Br até {ultimo_mes_ibc:%m/%Y}")
for nome, s in setor_y.items():
    print(f"[INFO] {nome}: índice dessazonalizado de {s.index.min().strftime('%m/%Y')} a {s.index.max().strftime('%m/%Y')}")

data_ref = min(selic_d.index[-1], cambio_d.index[-1])
selic_atual = valor_em(selic_aa, data_ref)
cambio_atual = (valor_em(cambio_d, data_ref) / valor_em(cambio_d, data_ref - pd.DateOffset(years=1)) - 1) * 100
ipca_ttm = float(ipca_12m.dropna().iloc[-1])
pib_atual = float(pib_12m.dropna().iloc[-1])
cenario_atual = {
    "Selic": float(selic_atual),
    "IPCA": ipca_ttm,
    "Cambio_var": float(cambio_atual),
    "PIB": pib_atual,   # apenas contexto: não entra no modelo nem nos estresses
}
for _k, _lag in LAGS.items():   # variável defasada no modelo => o cenário usa o valor de `lag` meses antes
    if _lag:
        cenario_atual[_k] = float(macro_base[_k].dropna().iloc[-1 - _lag])
print("[INFO] Cenário atual: " + " | ".join(f"{k} {v:.2f}" for k, v in cenario_atual.items()))
print(f"[INFO] Datas: Selic/câmbio em {data_ref:%d/%m/%Y}; IPCA até {ultimo_mes_ipca:%m/%Y}; IBC-Br até {ultimo_mes_ibc:%m/%Y}")
cenario_estresse_mild = {k: cenario_atual[k] + STRESS_MILD[k] for k in MACRO_COLS}
cenario_estresse = {k: cenario_atual[k] + STRESS[k] for k in MACRO_COLS}
FAIXA_HIST = {k: (float(df[k].min()), float(df[k].max())) for k in MACRO_COLS}   # mínimo e máximo observados na amostra
LIMITADAS_MILD: list[str] = []      # variáveis cujo estresse bateu na barreira (rótulos para o relatório)
LIMITADAS_SEV: list[str] = []
if LIMITAR_FAIXA:
    for _cen, _lim in ((cenario_estresse_mild, LIMITADAS_MILD), (cenario_estresse, LIMITADAS_SEV)):
        for _k in MACRO_COLS:
            _lo, _hi = FAIXA_HIST[_k]
            _v = min(max(_cen[_k], _lo), _hi)
            if _v != _cen[_k]:
                _lim.append(MACRO_LABEL[_k])
            _cen[_k] = _v
    print(f"[INFO] Estresses limitados à faixa histórica da amostra: moderado {LIMITADAS_MILD or 'nenhuma variável'} | "
          f"severo {LIMITADAS_SEV or 'nenhuma variável'}")

# ==============================================================================
# 2. CAMADA ECONOMÉTRICA: OLS com HAC, validação fora da amostra e diagnósticos
# ==============================================================================
print("\n[ETAPA 1] Estimando modelos OLS (Newey-West) por setor...")


def cov_newey_west(X: np.ndarray, resid: np.ndarray, XtXi: np.ndarray, L: int) -> np.ndarray:
    """Covariância HAC de Newey-West (pesos de Bartlett, `L` defasagens) dos coeficientes de OLS, com correção n/(n−k)."""
    n, k = X.shape
    U = X * resid[:, None]
    S0 = U.T @ U
    for l in range(1, L + 1):
        w = 1 - l / (L + 1)
        G = U[l:].T @ U[:-l]
        S0 += w * (G + G.T)
    return XtXi @ S0 @ XtXi * n / (n - k)


def r2_cv_bloco(X: np.ndarray, y: np.ndarray, h: int = H_CV_BLOCO) -> float:
    """R² de validação cruzada em blocos (leave-block-out): para prever o mês i, o modelo é estimado SEM os meses
    |j − i| ≤ h (as taxas de 12 meses vizinhas carregam quase a mesma informação). Benchmark: média do treino.
    Usa observações futuras no treino: mede estabilidade fora da amostra, não capacidade de previsão no tempo."""
    n = len(y)
    pred, bench = np.empty(n), np.empty(n)
    idx = np.arange(n)
    for i in range(n):
        tr = np.abs(idx - i) > h
        beta = np.linalg.lstsq(X[tr], y[tr], rcond=None)[0]
        pred[i], bench[i] = X[i] @ beta, y[tr].mean()
    return float(1 - np.sum((y - pred) ** 2) / np.sum((y - bench) ** 2))


def r2_walk_forward(X: np.ndarray, y: np.ndarray, emb: int = EMBARGO_WF, ini: int = MIN_TREINO_WF) -> float:
    """R² fora da amostra walk-forward: só usa o passado, com EMBARGO de `emb` meses entre treino e alvo
    (as taxas de 12 meses de t−1…t−11 se sobrepõem à de t). Benchmark: média histórica expansiva do treino.
    Negativo = o modelo erra mais que simplesmente prever a média histórica."""
    n, k = X.shape
    pred, bench = np.full(n, np.nan), np.full(n, np.nan)
    for i in range(ini + emb, n):
        tr = slice(0, i - emb)
        if np.linalg.matrix_rank(X[tr]) < k:
            continue
        beta = np.linalg.lstsq(X[tr], y[tr], rcond=None)[0]
        pred[i], bench[i] = X[i] @ beta, y[tr].mean()
    ok = np.isfinite(pred)
    if ok.sum() < 12:
        return np.nan
    return float(1 - np.sum((y[ok] - pred[ok]) ** 2) / np.sum((y[ok] - bench[ok]) ** 2))


def bootstrap_blocos_circular(residuos: np.ndarray, n_sim: int, tamanho_bloco: int, rng: np.random.Generator) -> np.ndarray:
    """Bootstrap circular em blocos, preservando dependência temporal local."""
    residuos = np.asarray(residuos, dtype=float)
    n, _ = residuos.shape
    if n < tamanho_bloco:
        raise ValueError("A amostra é menor que o bloco escolhido para o bootstrap.")
    n_blocos = int(np.ceil(n_sim / tamanho_bloco))
    inicios = rng.integers(0, n, size=n_blocos)
    indices = (inicios[:, None] + np.arange(tamanho_bloco)[None, :]) % n
    return residuos[indices.ravel()[:n_sim], :]


def _cov_psd(C: np.ndarray) -> np.ndarray:
    """Projeta uma covariância (HAC pode sair levemente indefinida) no cone semidefinido positivo."""
    C = (C + C.T) / 2
    w, V = np.linalg.eigh(C)
    return (V * np.clip(w, 1e-12, None)) @ V.T


def ols(X: np.ndarray, y: np.ndarray, cv: bool = True) -> dict[str, Any]:
    """OLS com erros Newey-West; t de Student com g.l. = anos independentes − parâmetros; R² dentro e fora da amostra.

    Returns: dict com beta, se, p, IC95%, covariância HAC, resíduos, ajustados, R², R² ajustado, p do F robusto,
    Durbin-Watson, R² de validação cruzada em blocos e R² walk-forward (estes dois só se `cv=True`).
    """
    n, k = X.shape
    if n <= k + 5:
        raise ValueError(f"Amostra insuficiente para OLS: n={n}, parâmetros={k}.")
    if np.linalg.matrix_rank(X) < k:
        raise ValueError("Matriz de regressoras sem posto completo: colinearidade perfeita.")
    XtX = X.T @ X
    XtXi = np.linalg.inv(XtX)
    beta = XtXi @ X.T @ y
    resid = y - X @ beta
    s2 = resid @ resid / (n - k)
    cov = cov_newey_west(X, resid, XtXi, HAC_LAGS)
    se = np.sqrt(np.diag(cov))
    dof = max(int(np.floor(n / 12.0)) - k, 3)           # anos independentes − parâmetros
    pval = 2 * stats.t.sf(np.abs(beta / se), dof)
    tcrit = stats.t.ppf(0.975, dof)
    sst = ((y - y.mean()) ** 2).sum()
    r2 = 1 - resid @ resid / sst
    r2_adj = 1 - (1 - r2) * (n - 1) / (n - k)
    b1, C1 = beta[1:], cov[1:, 1:]
    wald = float(b1 @ np.linalg.solve(C1, b1))
    p_F = stats.f.sf(wald / (k - 1), k - 1, dof)
    dw = np.sum(np.diff(resid) ** 2) / (resid @ resid)
    return dict(beta=beta, se=se, p=pval, ci_lo=beta - tcrit * se, ci_hi=beta + tcrit * se, cov=cov, resid=resid,
                fitted=X @ beta, r2=r2, r2_adj=r2_adj, p_F=p_F, dw=dw, dof=dof, s2=s2, n=n,
                r2_cv=r2_cv_bloco(X, y) if cv else np.nan, r2_wf=r2_walk_forward(X, y) if cv else np.nan)


def vif(Xm: np.ndarray) -> np.ndarray:
    """Fator de inflação da variância de cada coluna de `Xm` (regressoras sem o intercepto)."""
    out = []
    for j in range(Xm.shape[1]):
        yj = Xm[:, j]
        Xo = np.column_stack([np.ones(len(yj)), np.delete(Xm, j, axis=1)])
        b = np.linalg.lstsq(Xo, yj, rcond=None)[0]
        r = yj - Xo @ b
        out.append(1 / (1 - (1 - r @ r / ((yj - yj.mean()) ** 2).sum())))
    return np.array(out)


def adf_teste(y: np.ndarray, maxlag: int = 12) -> tuple[float, int, str]:
    """Dickey-Fuller aumentado (com constante); defasagens escolhidas por BIC.
    Valores críticos assintóticos (MacKinnon): −3,43 (1%), −2,86 (5%), −2,57 (10%). Poder baixo em amostras curtas."""
    y = np.asarray(y, dtype=float)
    dy = np.diff(y)
    melhor = None
    for p in range(0, maxlag + 1):
        T = len(dy) - p
        cols = [np.ones(T), y[p:-1]] + [dy[p - i:-i] for i in range(1, p + 1)]
        Xa = np.column_stack(cols)
        Y = dy[p:]
        b = np.linalg.lstsq(Xa, Y, rcond=None)[0]
        e = Y - Xa @ b
        sse = e @ e
        bic = np.log(sse / T) + Xa.shape[1] * np.log(T) / T
        if melhor is None or bic < melhor[0]:
            s2 = sse / (T - Xa.shape[1])
            se = np.sqrt(s2 * np.linalg.inv(Xa.T @ Xa)[1, 1])
            melhor = (bic, b[1] / se, p)
    t = melhor[1]
    conc = "estacionária (1%)" if t < -3.43 else "estacionária (5%)" if t < -2.86 else \
           "estacionária (10%)" if t < -2.57 else "raiz unitária não rejeitada"
    return t, melhor[2], conc


Xm = df[MACRO_COLS].values
X = np.column_stack([np.ones(n_obs), Xm])
vifs = vif(Xm)
sd_macro = Xm.std(axis=0, ddof=1)

modelos = {}
for s in nomes_setores:
    modelos[s] = ols(X, df[s].values)
    m = modelos[s]
    print(f" > {s}: R2 = {m['r2']:.2f} | R² CV em blocos = {m['r2_cv']:.2f} | "
          f"R² walk-forward (embargo {EMBARGO_WF}m) = {m['r2_wf']:.2f} | DW = {m['dw']:.2f}")
print(f" > VIF (Selic, IPCA, Câmbio): {np.round(vifs, 2)}")

pd.DataFrame(
    [dict(setor=s, variavel=nome, beta=m["beta"][j], erro_padrao_hac=m["se"][j], p_valor=m["p"][j],
          ic95_inf=m["ci_lo"][j], ic95_sup=m["ci_hi"][j])
     for s, m in modelos.items() for j, nome in enumerate(["Intercepto"] + MACRO_COLS)]
).to_csv(os.path.join(OUT_DIR, "coeficientes_pmes.csv"), index=False, float_format="%.4f")


def testar_robustez_picos() -> None:
    """Informativo: re-estima com os picos de 1 mês suavizados e compara com o modelo base."""
    print("\n[ROBUSTEZ] Base vs. picos suavizados (beta Selic / IPCA / Câmbio e R² walk-forward)")
    for s_ in nomes_setores:
        picos = detectar_picos(setor_y[s_])
        if not picos:
            print(f"  {s_}: sem picos detectados; nada a comparar")
            continue
        y_s = yoy(suavizar_picos(setor_y[s_], picos)).rename(s_)
        d_ = pd.concat([macro, y_s], axis=1)
        d_ = d_[d_.index >= pd.Period(INI_AMOSTRA, "M")].dropna()
        X_ = np.column_stack([np.ones(len(d_)), d_[MACRO_COLS].values])
        r_ = ols(X_, d_[s_].values)
        b0, b1 = modelos[s_]["beta"], r_["beta"]
        f = lambda b: f"{b[1]:+.2f} / {b[2]:+.2f} / {b[3]:+.2f}"
        print(f"  {s_}: base {f(b0)} (R² WF {modelos[s_]['r2_wf']:.2f}) | suavizado {f(b1)} (R² WF {r_['r2_wf']:.2f})")


testar_robustez_picos()

# ==============================================================================
# ETAPA 1 — CORRELAÇÃO E REGRESSÃO: variáveis macro → atividade setorial (2012–2026)
# ==============================================================================
CORES = {"Serviços de TI": "#2B6CB0", "Indústria": "#DD6B20", "Comércio Varejista": "#2F855A"}
CORES_MACRO = {"Selic": "#C53030", "IPCA": "#B7791F", "Cambio_var": "#2C7A7B", "PIB": "#553C9A"}
fmt_pct = FuncFormatter(lambda v, _: f"{v:.0f}%".replace(".", ","))
plt.rcParams.update({"font.size": 9, "axes.spines.top": False, "axes.spines.right": False})
virg = lambda x, d=2: f"{x:.{d}f}".replace(".", ",")
P_INI = pd.Period(INI_AMOSTRA, "M")
P_COVID = (pd.Period(COVID_INI, "M"), pd.Period(COVID_FIM, "M"))
mask_covid = (df.index >= P_COVID[0]) & (df.index <= P_COVID[1])
N_EF = max(n_obs / 12.0, 4.0)       # anos independentes (aprox.)
E1 = {}                              # resultados da Etapa 1 (usados no PDF)

# ---------- 1.1 Descritiva e estacionariedade ----------
print("\n[ETAPA 1.1] Estatísticas descritivas e teste ADF...")
linhas = []
for nome_v in DESC_COLS + nomes_setores:
    v = df[nome_v].values
    t_adf, p_adf, concl = adf_teste(v)
    linhas.append(dict(serie=MACRO_LABEL.get(nome_v, nome_v), tipo="Macro" if nome_v in DESC_COLS else "Setor (cresc. 12m)",
                       media=v.mean(), desvio=v.std(ddof=1), minimo=v.min(), maximo=v.max(),
                       assimetria=stats.skew(v), curtose_exc=stats.kurtosis(v), adf_t=t_adf, adf_lags=p_adf,
                       adf_conclusao=concl))
E1["desc"] = pd.DataFrame(linhas)
E1["desc"].to_csv(os.path.join(OUT_DIR, "etapa1_descritivas_adf.csv"), index=False, float_format="%.3f")
print(E1["desc"][["serie", "media", "desvio", "adf_t", "adf_conclusao"]].round(2).to_string(index=False))

# ---------- 1.2 Correlação ----------
print("\n[ETAPA 1.2] Correlações (Pearson, Spearman, parcial e defasada)...")
rng_c = np.random.default_rng(SEED + 1)


def idx_bloco_circular(n: int, B: int, bloco: int, rng: np.random.Generator) -> np.ndarray:
    """Índices de B reamostragens por bootstrap circular em blocos de tamanho `bloco` (matriz B × n)."""
    nb = int(np.ceil(n / bloco))
    ini = rng.integers(0, n, size=(B, nb))
    return ((ini[:, :, None] + np.arange(bloco)[None, None, :]) % n).reshape(B, -1)[:, :n]


def _pearson_lote(xs: np.ndarray, ys: np.ndarray) -> np.ndarray:
    """Correlação de Pearson linha a linha entre duas matrizes (vetorizado)."""
    xm, ym = xs - xs.mean(1, keepdims=True), ys - ys.mean(1, keepdims=True)
    return (xm * ym).sum(1) / np.sqrt((xm ** 2).sum(1) * (ym ** 2).sum(1))


IDX_BOOT = idx_bloco_circular(n_obs, N_BOOT_CORR, BLOCO_BOOT_CORR, rng_c)


def corr_com_ic(x: np.ndarray, y: np.ndarray) -> dict[str, float]:
    """Pearson e Spearman com IC95% por bootstrap circular em blocos (12m) e p-valor t com n efetivo (anos)."""
    r = np.corrcoef(x, y)[0, 1]
    rho = stats.spearmanr(x, y)[0]
    rb = _pearson_lote(x[IDX_BOOT], y[IDX_BOOT])
    sb = _pearson_lote(stats.rankdata(x[IDX_BOOT], axis=1), stats.rankdata(y[IDX_BOOT], axis=1))
    t = r * np.sqrt((N_EF - 2) / (1 - r ** 2))
    return dict(pearson=r, p_ef=2 * stats.t.sf(abs(t), N_EF - 2), p_lo=np.percentile(rb, 2.5), p_hi=np.percentile(rb, 97.5),
                spearman=rho, s_lo=np.percentile(sb, 2.5), s_hi=np.percentile(sb, 97.5))


def corr_parcial(cols: list[str], alvo: str) -> dict[str, float]:
    """Correlação parcial de cada variável de `cols` com `alvo`, controlando pelas demais de `cols`."""
    C = np.corrcoef(df[cols + [alvo]].values.T)
    Pm = np.linalg.inv(C)
    return {c: -Pm[i, -1] / np.sqrt(Pm[i, i] * Pm[-1, -1]) for i, c in enumerate(cols)}


linhas = []
for s in nomes_setores:
    parc = corr_parcial(MACRO_COLS, s)   # parcial só entre as variáveis do modelo; PIB (contexto) não tem parcial
    for k in DESC_COLS:
        c = corr_com_ic(df[k].values, df[s].values)
        linhas.append(dict(setor=s, variavel=MACRO_LABEL[k], chave=k, **c, parcial=parc.get(k, np.nan),
                           sig_ic=bool(c["p_lo"] > 0 or c["p_hi"] < 0)))
E1["corr"] = pd.DataFrame(linhas)
E1["corr"].drop(columns="chave").to_csv(os.path.join(OUT_DIR, "etapa1_correlacoes.csv"), index=False, float_format="%.4f")
E1["corr_macro"] = df[MACRO_COLS].corr()
E1["corr_macro"].to_csv(os.path.join(OUT_DIR, "etapa1_correlacao_macro_macro.csv"), float_format="%.4f")
E1["corr_setores"] = df[nomes_setores].corr()
print(E1["corr"][["setor", "variavel", "pearson", "p_lo", "p_hi", "spearman", "parcial"]].round(2).to_string(index=False))

# correlação defasada: macro em t−L contra o crescimento setorial em t (L = 0…12)
lags_all = list(range(0, 13))
ll = {}
for s in nomes_setores:
    for k in DESC_COLS:
        vals = []
        for lag_ in lags_all:
            d_ = pd.concat([macro_base[k].shift(lag_), sectors[s]], axis=1).loc[P_INI:].dropna()
            vals.append(d_.iloc[:, 0].corr(d_.iloc[:, 1]))
        ll[(s, k)] = vals
E1["lead_lag"] = pd.DataFrame(ll, index=lags_all)
E1["lead_lag"].columns = [f"{s} | {MACRO_LABEL[k]}" for s, k in E1["lead_lag"].columns]
E1["lead_lag"].index.name = "defasagem_meses"
E1["lead_lag"].to_csv(os.path.join(OUT_DIR, "etapa1_correlacao_defasada.csv"), float_format="%.4f")

# ---------- 1.3 Regressão: modelo único + checagens de robustez ----------
print("\n[ETAPA 1.3] Modelo único e checagens de robustez (OLS + Newey-West)...")


def montar(cols: list[str], sel_lag: int = 0, sem_covid: bool = False) -> pd.DataFrame:
    """Monta a base de estimação (regressoras `cols` + 3 setores), com Selic defasada e/ou sem a janela da COVID."""
    mac = macro_base.copy()
    if sel_lag:
        mac["Selic"] = mac["Selic"].shift(sel_lag)
    d_ = pd.concat([mac[cols], sectors], axis=1)
    d_ = d_[d_.index >= P_INI].dropna()
    if sem_covid:
        d_ = d_[~((d_.index >= P_COVID[0]) & (d_.index <= P_COVID[1]))]
    return d_


ESPEC_BASE = "Modelo único"
ESPEC_LAG = f"Robustez: Selic t−{LAG_SELIC_ROB}"
ESPEC_COVID = "Robustez: sem COVID"
ESPECS = {   # a primeira é O modelo; as demais são checagens de robustez da mesma especificação
    ESPEC_BASE: dict(cols=MACRO_COLS),
    ESPEC_LAG: dict(cols=MACRO_COLS, sel_lag=LAG_SELIC_ROB),
    ESPEC_COVID: dict(cols=MACRO_COLS, sem_covid=True),
}
linhas = []
for nome_m, cfg in ESPECS.items():
    d_ = montar(**cfg)
    X_ = np.column_stack([np.ones(len(d_)), d_[cfg["cols"]].values])
    for s in nomes_setores:
        r = ols(X_, d_[s].values, cv=not cfg.get("sem_covid", False))   # a exclusão da COVID quebra a contiguidade: sem CV
        linha = dict(setor=s, modelo=nome_m, n=len(d_))
        for k in MACRO_COLS:
            if k in cfg["cols"]:
                j = cfg["cols"].index(k) + 1
                linha[f"b_{k}"], linha[f"p_{k}"] = r["beta"][j], r["p"][j]
            else:
                linha[f"b_{k}"], linha[f"p_{k}"] = np.nan, np.nan
        linha.update(r2=r["r2"], r2_adj=r["r2_adj"], r2_cv=r["r2_cv"], r2_wf=r["r2_wf"], dw=r["dw"], p_F=r["p_F"])
        linhas.append(linha)
E1["specs"] = pd.DataFrame(linhas)
E1["specs"].to_csv(os.path.join(OUT_DIR, "etapa1_regressoes_especificacoes.csv"), index=False, float_format="%.4f")
print(E1["specs"][["setor", "modelo", "n", "b_Selic", "b_IPCA", "b_Cambio_var", "r2", "r2_cv", "r2_wf"]]
      .round(2).to_string(index=False))

# coeficientes padronizados e efeito de +1 desvio-padrão (modelo único)
linhas = []
for s in nomes_setores:
    m = modelos[s]
    sdy = df[s].std(ddof=1)
    for j, k in enumerate(MACRO_COLS, start=1):
        linhas.append(dict(setor=s, variavel=MACRO_LABEL[k], chave=k, beta=m["beta"][j], p=m["p"][j],
                           beta_padron=m["beta"][j] * sd_macro[j - 1] / sdy, efeito_1dp_pp=m["beta"][j] * sd_macro[j - 1],
                           efeito_lo=m["ci_lo"][j] * sd_macro[j - 1], efeito_hi=m["ci_hi"][j] * sd_macro[j - 1]))
E1["efeitos"] = pd.DataFrame(linhas)
E1["efeitos"].drop(columns="chave").to_csv(os.path.join(OUT_DIR, "etapa1_efeitos_padronizados.csv"), index=False, float_format="%.4f")

# grade de defasagens da Selic (informativo: nenhuma defasagem é escolhida por busca nos dados)
linhas = []
for L_ in LAGS_GRADE:
    d_ = montar(MACRO_COLS, sel_lag=L_)
    X_ = np.column_stack([np.ones(len(d_)), d_[MACRO_COLS].values])
    for s in nomes_setores:
        r = ols(X_, d_[s].values)
        linhas.append(dict(setor=s, defasagem=L_, n=len(d_), b_Selic=r["beta"][1], p_Selic=r["p"][1],
                           r2=r["r2"], r2_cv=r["r2_cv"], r2_wf=r["r2_wf"]))
E1["grade_lag"] = pd.DataFrame(linhas)
E1["grade_lag"].to_csv(os.path.join(OUT_DIR, "etapa1_grade_defasagem_selic.csv"), index=False, float_format="%.4f")

# ---------- Figuras da Etapa 1 ----------
FIG_E1_CORR = os.path.join(OUT_DIR, "fig_e1_correlacao.png")
fig_e1_corr_size = (10, 3.6)
fig, axes = plt.subplots(1, 2, figsize=fig_e1_corr_size)
for ax, (col, ttl, lo, hi) in zip(axes, [("pearson", "Pearson (IC95% bloco-bootstrap exclui 0: *)", "p_lo", "p_hi"),
                                         ("spearman", "Spearman (IC95% bloco-bootstrap exclui 0: *)", "s_lo", "s_hi")]):
    M = np.array([[E1["corr"].query("setor == @s and chave == @k")[col].iloc[0] for s in nomes_setores] for k in DESC_COLS])
    im = ax.imshow(M, cmap="RdBu_r", vmin=-1, vmax=1, aspect="auto")
    ax.set_xticks(range(3)); ax.set_xticklabels(nomes_setores, fontsize=8)
    ax.set_yticks(range(len(DESC_COLS))); ax.set_yticklabels([MACRO_LABEL[k] for k in DESC_COLS], fontsize=8)
    for i, k in enumerate(DESC_COLS):
        for j, s in enumerate(nomes_setores):
            q = E1["corr"].query("setor == @s and chave == @k").iloc[0]
            estrela = "*" if (q[lo] > 0 or q[hi] < 0) else ""
            ax.text(j, i, virg(M[i, j]) + estrela, ha="center", va="center", fontsize=9,
                    color="white" if abs(M[i, j]) > 0.5 else "black")
    ax.set_title(ttl, fontsize=9)
fig.colorbar(im, ax=axes, fraction=0.025, pad=0.02)
fig.savefig(FIG_E1_CORR, dpi=200, bbox_inches="tight")
plt.close(fig)

FIG_E1_LAG = os.path.join(OUT_DIR, "fig_e1_correlacao_defasada.png")
fig_e1_lag_size = (10, 3.0)
fig, axes = plt.subplots(1, 3, figsize=fig_e1_lag_size, sharey=True)
for ax, s in zip(axes, nomes_setores):
    for k in DESC_COLS:
        ax.plot(lags_all, E1["lead_lag"][f"{s} | {MACRO_LABEL[k]}"], marker="o", ms=2.5, lw=1.4,
                color=CORES_MACRO[k], label=MACRO_LABEL[k])
    ax.axhline(0, color="#4A5568", lw=0.6)
    ax.set_title(s, fontsize=9)
    ax.set_xlabel("Defasagem da variável macro (meses)", fontsize=8)
    ax.set_xticks(range(0, 13, 2))
axes[0].set_ylabel("Correlação de Pearson")
axes[0].legend(fontsize=7, frameon=False, ncol=2, loc="lower left")
fig.tight_layout()
fig.savefig(FIG_E1_LAG, dpi=200)
plt.close(fig)

FIG_E1_DISP = os.path.join(OUT_DIR, "fig_e1_dispersao.png")
fig_e1_disp_size = (10, 7.2)
fig, axes = plt.subplots(3, 4, figsize=fig_e1_disp_size)
for i, s in enumerate(nomes_setores):
    for j, k in enumerate(DESC_COLS):
        ax = axes[i, j]
        x_, y_ = df[k].values, df[s].values
        ax.scatter(x_[~mask_covid], y_[~mask_covid], s=7, color=CORES[s], alpha=0.65, lw=0)
        ax.scatter(x_[mask_covid], y_[mask_covid], s=9, color="#C53030", alpha=0.9, lw=0,
                   label="COVID (03/2020–06/2021)" if (i == 0 and j == 0) else None)
        b1, b0 = np.polyfit(x_, y_, 1)
        xx = np.linspace(x_.min(), x_.max(), 50)
        ax.plot(xx, b0 + b1 * xx, color="#2D3748", lw=1.2)
        ax.text(0.04, 0.92, f"r = {virg(np.corrcoef(x_, y_)[0, 1])}", transform=ax.transAxes, fontsize=8, va="top",
                bbox=dict(facecolor="white", alpha=0.75, lw=0, pad=1.5))
        ax.tick_params(labelsize=6)
        if i == 2:
            ax.set_xlabel(f"{MACRO_LABEL[k]} (%)", fontsize=8)
        if j == 0:
            ax.set_ylabel(f"{s}\ncresc. 12m (%)", fontsize=8)
axes[0, 0].legend(fontsize=6, frameon=False, loc="lower left")
fig.tight_layout()
fig.savefig(FIG_E1_DISP, dpi=200)
plt.close(fig)

FIG_E1_COEF = os.path.join(OUT_DIR, "fig_e1_coeficientes.png")
fig_e1_coef_size = (10, 2.9)
fig, axes = plt.subplots(1, 3, figsize=fig_e1_coef_size, sharex=False)
for ax, s in zip(axes, nomes_setores):
    q = E1["efeitos"].query("setor == @s")
    y_pos = np.arange(len(q))[::-1]
    ax.errorbar(q["efeito_1dp_pp"], y_pos, xerr=[q["efeito_1dp_pp"] - q["efeito_lo"], q["efeito_hi"] - q["efeito_1dp_pp"]],
                fmt="o", color=CORES[s], ecolor="#4A5568", capsize=3, ms=4)
    ax.axvline(0, color="#C53030", lw=0.8, ls="--")
    ax.set_yticks(y_pos); ax.set_yticklabels(q["variavel"], fontsize=8)
    ax.set_title(s, fontsize=9)
    ax.set_xlabel("Efeito de +1 desvio-padrão (pp)", fontsize=8)
fig.tight_layout()
fig.savefig(FIG_E1_COEF, dpi=200)
plt.close(fig)
print(f" > Figuras da Etapa 1 salvas em {os.path.abspath(OUT_DIR)}")
# ==============================================================================
# 3. MONTE CARLO
# ==============================================================================
print("\n[ETAPA 1.4] Monte Carlo (t multivariada, bootstrap em blocos, nu = %d, N = %d)..." % (NU, N_SIM))
rng = np.random.default_rng(SEED)

mu = Xm.mean(axis=0)
cov_macro = np.cov(Xm.T)
try:
    chol = np.linalg.cholesky(cov_macro)
except np.linalg.LinAlgError:
    jitter = np.eye(cov_macro.shape[0]) * 1e-8 * np.trace(cov_macro) / cov_macro.shape[0]
    chol = np.linalg.cholesky(cov_macro + jitter)
    warnings.warn("Cholesky precisou de jitter — checar VIFs elevados.", stacklevel=2)

z = rng.standard_normal((N_SIM, len(MACRO_COLS))) @ chol.T
w_chi = rng.chisquare(NU, N_SIM) / NU          # mistura de escala: normal / √(χ²ν/ν) = t de Student multivariada
macro_sim = mu + z * np.sqrt((NU - 2) / NU) / np.sqrt(w_chi)[:, None]
N_REDESENHADOS = 0
if LIMITAR_FAIXA:
    # Rejeição: cenário com qualquer variável fora do [mín, máx] histórico é descartado e re-sorteado (sem extrapolar).
    _lo = np.array([FAIXA_HIST[k][0] for k in MACRO_COLS]); _hi = np.array([FAIXA_HIST[k][1] for k in MACRO_COLS])
    for _ in range(200):
        fora_ = np.any((macro_sim < _lo) | (macro_sim > _hi), axis=1)
        if not fora_.any():
            break
        N_REDESENHADOS += int(fora_.sum())
        n_f = int(fora_.sum())
        z_f = rng.standard_normal((n_f, len(MACRO_COLS))) @ chol.T
        w_f = rng.chisquare(NU, n_f) / NU
        macro_sim[fora_] = mu + z_f * np.sqrt((NU - 2) / NU) / np.sqrt(w_f)[:, None]
    assert not np.any((macro_sim < _lo) | (macro_sim > _hi)), "Truncamento do Monte Carlo não convergiu."
    print(f"[INFO] Monte Carlo truncado na faixa histórica: {N_REDESENHADOS} sorteios re-sorteados.")
Xsim = np.column_stack([np.ones(N_SIM), macro_sim])

fator_gl = np.sqrt(n_obs / (n_obs - X.shape[1]))
R = np.column_stack([modelos[s]["resid"] for s in nomes_setores]) * fator_gl
# Índices conjuntos preservam a associação contemporânea entre os três setores.
R_boot = bootstrap_blocos_circular(R, N_SIM, BLOCO_CV, rng)

sim = {}
for j, s in enumerate(nomes_setores):
    m = modelos[s]
    beta_draw = rng.multivariate_normal(m["beta"], _cov_psd(m["cov"]), N_SIM)
    sim[s] = np.einsum("ij,ij->i", Xsim, beta_draw) + R_boot[:, j]

# ==============================================================================
# 4. MÉTRICAS DE RISCO
# ==============================================================================
print("\n[ETAPA 1.4] Calculando métricas de risco...")
metr = {}
for s in nomes_setores:
    y = sim[s]
    esperado = y.mean()
    perda = esperado - y
    var95, var99 = np.percentile(perda, [95, 99])
    hist = df[s]
    metr[s] = dict(
        esperado=esperado, desvio=y.std(ddof=1), var95=var95, var99=var99,
        cvar95=perda[perda >= var95].mean(),
        p_contracao=(y < 0).mean() * 100,
        pior1=np.percentile(y, 1),
        hist_min=hist.min(), hist_min_ano=hist.idxmin().strftime("%m/%Y"),
    )

pd.DataFrame(metr).T.to_csv(os.path.join(OUT_DIR, "metricas_risco_pmes.csv"), float_format="%.4f")


def prever(cen: dict[str, float], s: str) -> float:
    """Previsão pontual do crescimento do setor `s` no cenário `cen` (modelo único: Selic, IPCA e câmbio)."""
    x = np.array([1.0] + [cen[k] for k in MACRO_COLS])
    return float(x @ modelos[s]["beta"])


pred_atual = {s: prever(cenario_atual, s) for s in nomes_setores}
pred_estr_mild = {s: prever(cenario_estresse_mild, s) for s in nomes_setores}
pred_estr = {s: prever(cenario_estresse, s) for s in nomes_setores}
for s in nomes_setores:
    print(f" > {s}: esperado {metr[s]['esperado']:.1f}% | CVaR95 {metr[s]['cvar95']:.1f} pp | "
          f"P(contração) {metr[s]['p_contracao']:.1f}%")

datas = df.index.to_timestamp()
FIG1 = os.path.join(OUT_DIR, "distribuicao_crescimento_pmes.png")
fig1_size = (10, 4.4)
fig, ax = plt.subplots(figsize=fig1_size)
lo = min(np.percentile(sim[s], 0.5) for s in nomes_setores)
hi = max(np.percentile(sim[s], 99.5) for s in nomes_setores)
xs = np.linspace(lo, hi, 400)
for s in nomes_setores:
    kde = stats.gaussian_kde(sim[s])
    ax.fill_between(xs, kde(xs), alpha=0.25, color=CORES[s])
    ax.plot(xs, kde(xs), color=CORES[s], lw=1.8,
            label=f"{s} (esperado: {metr[s]['esperado']:.1f}%)".replace(".", ","))
ax.axvline(0, color="#C53030", ls="--", lw=1.3, label="Crescimento zero (contração abaixo)")
ax.xaxis.set_major_formatter(fmt_pct)
ax.set_xlabel("Crescimento anual do índice setorial (proxy real)")
ax.set_ylabel("Densidade")
ax.set_yticks([])
ax.set_title("Monte Carlo: distribuição do crescimento anual do faturamento por setor", fontsize=11)
ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=2, frameon=False)
fig.tight_layout()
fig.savefig(FIG1, dpi=200)
plt.close(fig)

FIG2 = os.path.join(OUT_DIR, "ajuste_historico_pmes.png")
fig2_size = (10, 3.1)
fig, axes = plt.subplots(1, 3, figsize=fig2_size, sharex=True)
for ax, s in zip(axes, nomes_setores):
    ax.plot(datas, df[s].values, color=CORES[s], lw=1.5, label="Observado")
    ax.plot(datas, modelos[s]["fitted"], color="#4A5568", lw=1.3, ls="--", label="Modelo macro")
    ax.axhline(0, color="#C53030", lw=0.8, alpha=0.6)
    ax.yaxis.set_major_formatter(fmt_pct)
    ax.set_title(f"{s} (R² = {modelos[s]['r2']:.2f})".replace(".", ","), fontsize=9)
    ax.xaxis.set_major_locator(mdates.YearLocator(4)); ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.tick_params(labelsize=7)
axes[0].legend(fontsize=7, frameon=False, loc="lower left")
fig.tight_layout()
fig.savefig(FIG2, dpi=200)
plt.close(fig)

# ----- Gráfico do padrão histórico das variáveis macro -----
FIG_MACRO = os.path.join(OUT_DIR, "padrao_historico_macro.png")
fig_macro_size = (10, 4.2)
fig, axes = plt.subplots(1, 2, figsize=fig_macro_size)

macro_df = df[MACRO_COLS]
medias = macro_df.mean()
vols = macro_df.std()
x = np.arange(len(MACRO_COLS))
larg_barra = 0.35
axes[0].bar(x - larg_barra / 2, medias.values, larg_barra, label="Média histórica", color="#2B6CB0", alpha=0.85)
axes[0].bar(x + larg_barra / 2, vols.values, larg_barra, label="Volatilidade (desvio-padrão)", color="#DD6B20", alpha=0.85)
axes[0].set_xticks(x)
axes[0].set_xticklabels([MACRO_LABEL[k] for k in MACRO_COLS], fontsize=8)
axes[0].yaxis.set_major_formatter(fmt_pct)
axes[0].set_title(f"Média e volatilidade históricas das variáveis macro\n({MES_INI_D:%m/%Y} – {MES_FIM_D:%m/%Y})", fontsize=10)
axes[0].legend(fontsize=7, frameon=False)
axes[0].axhline(0, color="#4A5568", lw=0.6)

corr = macro_df.corr()
im = axes[1].imshow(corr.values, cmap="RdBu_r", vmin=-1, vmax=1, aspect="auto")
axes[1].set_xticks(range(len(MACRO_COLS)))
axes[1].set_yticks(range(len(MACRO_COLS)))
axes[1].set_xticklabels([MACRO_LABEL[k] for k in MACRO_COLS], fontsize=8)
axes[1].set_yticklabels([MACRO_LABEL[k] for k in MACRO_COLS], fontsize=8)
for i in range(len(MACRO_COLS)):
    for j in range(len(MACRO_COLS)):
        axes[1].text(j, i, f"{corr.values[i, j]:.2f}".replace(".", ","),
                     ha="center", va="center", fontsize=8,
                     color="white" if abs(corr.values[i, j]) > 0.5 else "black")
axes[1].set_title(f"Correlação histórica entre as variáveis macro\n({MES_INI_D:%m/%Y} – {MES_FIM_D:%m/%Y})", fontsize=10)
fig.colorbar(im, ax=axes[1], fraction=0.046, pad=0.04)
fig.tight_layout()
fig.savefig(FIG_MACRO, dpi=200)
plt.close(fig)
print(f" > Gráficos salvos: {FIG1}, {FIG2}, {FIG_MACRO}")

# ==============================================================================
# ETAPA 2 — TRANSMISSÃO AO CAIXA E AO BALANÇO (perfis FICTÍCIOS: planilhas Excel/CSV)
# ==============================================================================
print("\n[ETAPA 2] Lendo e validando os perfis financeiros fictícios...")
E2 = {}
PERFIS, _log = {}, []
for _nome, _cfg in SETORES.items():
    PERFIS[_nome], _lg = validar_perfil(_nome, ler_entrada(_cfg["perfil"], ler_perfil))
    _log += _lg
E2["validacao"] = pd.DataFrame(_log)
E2["validacao"].to_csv(os.path.join(OUT_DIR, "etapa2_validacao_perfis.csv"), index=False, float_format="%.2f")
for _, r_ in E2["validacao"].iterrows():
    if r_["acao"] != "ok":
        warnings.warn(f"Perfil '{r_['setor']}' — {r_['item']}: informado {r_['informado']:,.2f}; recalculado "
                      f"{r_['recalculado']:,.2f} ({r_['acao']}). Corrija o arquivo de origem.", stacklevel=2)


def balanco_base(p: dict[str, float]) -> dict[str, float]:
    """Converte o perfil validado nas contas-base da projeção (custo fixo e variável separados)."""
    C0 = p["C0"]
    return dict(R=p["R0"], C=C0, F0=p["fixo"] * C0, V0=(1 - p["fixo"]) * C0, juros0=p["juros"], L0=p["L0"],
                caixa=p["caixa"], cr=p["cr"], est=p["est"], oac=p["oac"], anc=p["anc"], forn=p["forn"],
                outros=p["outros"], dcp=p["dcp"], dlp=p["dlp"], pl=p["pl"], indexada=p["indexada"])


def projetar(b: dict[str, float], g: Any, ipca: Any, selic: Any, selic_base: float) -> dict[str, Any]:
    """DRE, fluxo de caixa e balanço de um ano. g = crescimento de VOLUME (%); ipca = proxy de preço (%).
    Receita e custos variáveis ∝ (1+g)(1+IPCA); custos fixos ∝ (1+IPCA); juros = base + dívida × parcela indexada × ΔSelic;
    prazos médios constantes; Capex = depreciação; sem amortização; retiradas = lucro do ano-base;
    caixa insuficiente vira dívida de curto prazo."""
    g, ipca, selic = (np.asarray(x, dtype=float) for x in (g, ipca, selic))
    fp = 1 + ipca / 100
    fn = (1 + g / 100) * fp
    R1, F1, V1 = b["R"] * fn, b["F0"] * fp, b["V0"] * fn
    juros1 = np.maximum(b["juros0"] + (b["dcp"] + b["dlp"]) * b["indexada"] * (selic - selic_base) / 100, 0.0)
    L1 = R1 - F1 - V1 - juros1
    esc_r, esc_c = R1 / b["R"], (F1 + V1) / b["C"]
    cr1, oac1 = b["cr"] * esc_r, b["oac"] * esc_r
    est1, forn1, outros1 = b["est"] * esc_c, b["forn"] * esc_c, b["outros"] * esc_r
    dncg = (cr1 - b["cr"]) + (oac1 - b["oac"]) + (est1 - b["est"]) - (forn1 - b["forn"]) - (outros1 - b["outros"])
    fco = L1 - dncg
    dist = b["L0"]
    caixa_pre = b["caixa"] + fco - dist
    gap = np.maximum(-caixa_pre, 0.0)
    caixa1 = np.maximum(caixa_pre, 0.0)
    dcp1 = b["dcp"] + gap
    pl1 = b["pl"] + L1 - dist
    ac1, pc1 = caixa1 + cr1 + est1 + oac1, forn1 + outros1 + dcp1
    ativo, passivo = ac1 + b["anc"], pc1 + b["dlp"] + pl1
    if not np.allclose(ativo, passivo):
        raise ArithmeticError("Balanço projetado não fecha (ativo ≠ passivo + PL): revise o perfil financeiro.")
    return dict(R=R1, L=L1, margem=L1 / R1 * 100, juros=juros1, fco=fco, dist=dist + 0 * L1, dncg=dncg, gap=gap,
                caixa=caixa1, cr=cr1, est=est1, anc=b["anc"] + 0 * L1, ativo=ativo, forn_out=forn1 + outros1, dcp=dcp1,
                dlp=b["dlp"] + 0 * L1, pl=pl1, passivo=passivo, lc=ac1 / pc1)


def como_base(b: dict[str, float]) -> dict[str, float]:
    """Indicadores do ano-base no mesmo formato do resultado de `projetar` (para a coluna 'Ano-base')."""
    ac, pc = b["caixa"] + b["cr"] + b["est"] + b["oac"], b["forn"] + b["outros"] + b["dcp"]
    ativo = ac + b["anc"]
    return dict(R=b["R"], L=b["L0"], margem=b["L0"] / b["R"] * 100, juros=b["juros0"], fco=np.nan, dist=np.nan,
                dncg=np.nan, gap=0.0, caixa=b["caixa"], cr=b["cr"], est=b["est"], anc=b["anc"], ativo=ativo,
                forn_out=b["forn"] + b["outros"], dcp=b["dcp"], dlp=b["dlp"], pl=b["pl"], passivo=ativo, lc=ac / pc)


print("\n[ETAPA 2.2] Projetando DRE, fluxo de caixa e balanço...")
bal0, fin_base, fin_atual, fin_mild, fin_estr, fin_mc = {}, {}, {}, {}, {}, {}
sel0 = cenario_atual["Selic"]
for nome in nomes_setores:
    b = balanco_base(PERFIS[nome])
    bal0[nome] = b
    fin_base[nome] = como_base(b)
    fin_atual[nome] = projetar(b, pred_atual[nome], cenario_atual["IPCA"], sel0, sel0)
    fin_mild[nome] = projetar(b, pred_estr_mild[nome], cenario_estresse_mild["IPCA"], cenario_estresse_mild["Selic"], sel0)
    fin_estr[nome] = projetar(b, pred_estr[nome], cenario_estresse["IPCA"], cenario_estresse["Selic"], sel0)
    r = projetar(b, sim[nome], macro_sim[:, 1], macro_sim[:, 0], sel0)   # colunas de macro_sim: Selic, IPCA, Câmbio
    fin_mc[nome] = dict(
        r=r, lucro_med=r["L"].mean(), fco_med=r["fco"].mean(), fco_p5=np.percentile(r["fco"], 5),
        caixa_p5=np.percentile(r["caixa"], 5), lc_med=np.median(r["lc"]), lc_p5=np.percentile(r["lc"], 5),
        p_prejuizo=(r["L"] < 0).mean() * 100, p_lc1=(r["lc"] < 1).mean() * 100, p_gap=(r["gap"] > 0).mean() * 100)
    print(f" > {nome}: LC base {fin_base[nome]['lc']:.2f} | atual {fin_atual[nome]['lc']:.2f} | "
          f"moderado {fin_mild[nome]['lc']:.2f} | severo {fin_estr[nome]['lc']:.2f} | "
          f"P(LC<1) {fin_mc[nome]['p_lc1']:.1f}% | P(financiamento) {fin_mc[nome]['p_gap']:.1f}%")

_linhas = []
for nome in nomes_setores:
    for cen, d in (("Ano-base", fin_base[nome]), ("Cenário atual", fin_atual[nome]),
                   ("Estresse moderado", fin_mild[nome]), ("Estresse severo", fin_estr[nome])):
        _linhas.append(dict(setor=nome, cenario=cen, **{k: float(v) for k, v in d.items()}))
E2["projecao"] = pd.DataFrame(_linhas)
E2["projecao"].to_csv(os.path.join(OUT_DIR, "etapa2_projecao_financeira.csv"), index=False, float_format="%.2f")
E2["mc_fin"] = pd.DataFrame([{k: v for k, v in dict(setor=n_, **fin_mc[n_]).items() if k != "r"} for n_ in nomes_setores])
E2["mc_fin"].to_csv(os.path.join(OUT_DIR, "etapa2_montecarlo_financeiro.csv"), index=False, float_format="%.3f")

# ---------- 2.3 Correlação e regressão de SENSIBILIDADE (sobre as simulações) ----------
print("\n[ETAPA 2.3] Sensibilidade: correlação e regressão padronizada sobre as 10.000 simulações...")
Z_macro = (macro_sim - macro_sim.mean(0)) / macro_sim.std(0)
A_ = np.column_stack([np.ones(N_SIM), Z_macro])
_linhas = []
for nome in nomes_setores:
    r = fin_mc[nome]["r"]
    for ind, y in (("Lucro líquido", r["L"]), ("FCO", r["fco"]), ("Liquidez corrente", r["lc"])):
        ys = (y - y.mean()) / y.std()
        bb = np.linalg.lstsq(A_, ys, rcond=None)[0]
        res = ys - A_ @ bb
        r2 = 1 - res @ res / (ys @ ys)
        for j, k in enumerate(MACRO_COLS):
            _linhas.append(dict(setor=nome, indicador=ind, variavel=MACRO_LABEL[k], chave=k,
                                spearman=stats.spearmanr(macro_sim[:, j], y)[0], beta_padron=bb[j + 1], r2_macro=r2))
E2["sens_mc"] = pd.DataFrame(_linhas)
E2["sens_mc"].drop(columns="chave").to_csv(os.path.join(OUT_DIR, "etapa2_sensibilidade_mc.csv"), index=False, float_format="%.4f")
print(E2["sens_mc"].query("indicador == 'Lucro líquido'")[["setor", "variavel", "spearman", "beta_padron", "r2_macro"]]
      .round(2).to_string(index=False))

# tornado determinístico: ±1 desvio-padrão histórico de UMA variável por vez, a partir do cenário atual
_linhas = []
for nome in nomes_setores:
    b = bal0[nome]
    base = fin_atual[nome]
    for j, k in enumerate(MACRO_COLS):
        for sinal, rot in ((+1, "+1 dp"), (-1, "−1 dp")):
            cen = dict(cenario_atual)
            cen[k] += sinal * sd_macro[j]
            if LIMITAR_FAIXA:
                cen[k] = min(max(cen[k], FAIXA_HIST[k][0]), FAIXA_HIST[k][1])   # sem extrapolar a faixa histórica
            rr = projetar(b, prever(cen, nome), cen["IPCA"], cen["Selic"], sel0)
            _linhas.append(dict(setor=nome, variavel=MACRO_LABEL[k], chave=k, choque=rot, delta_macro_pp=cen[k] - cenario_atual[k],
                                delta_lucro=float(rr["L"] - base["L"]), delta_lucro_pct=float((rr["L"] / base["L"] - 1) * 100),
                                delta_lc=float(rr["lc"] - base["lc"])))
E2["tornado"] = pd.DataFrame(_linhas)
E2["tornado"].drop(columns="chave").to_csv(os.path.join(OUT_DIR, "etapa2_tornado.csv"), index=False, float_format="%.3f")

# ---------- Figuras da Etapa 2 ----------
FIG_E2_TORN = os.path.join(OUT_DIR, "fig_e2_tornado.png")
fig_e2_torn_size = (10, 3.0)
fig, axes = plt.subplots(1, 3, figsize=fig_e2_torn_size)
for ax, nome in zip(axes, nomes_setores):
    q = E2["tornado"].query("setor == @nome")
    ordem_ = q.assign(a=q["delta_lucro_pct"].abs()).groupby("chave")["a"].max().sort_values().index.tolist()
    for i, k in enumerate(ordem_):
        for sinal, cor, off in (("+1 dp", "#2B6CB0", 0.18), ("−1 dp", "#DD6B20", -0.18)):
            v = q[(q["chave"] == k) & (q["choque"] == sinal)]["delta_lucro_pct"].iloc[0]
            ax.barh(i + off, v, height=0.34, color=cor, label=sinal if i == 0 else None)
    ax.set_yticks(range(len(ordem_))); ax.set_yticklabels([MACRO_LABEL[k] for k in ordem_], fontsize=8)
    ax.axvline(0, color="#4A5568", lw=0.7)
    ax.xaxis.set_major_formatter(fmt_pct)
    ax.set_title(nome, fontsize=9)
    ax.set_xlabel("Δ lucro vs cenário atual (%)", fontsize=8)
axes[0].legend(fontsize=7, frameon=False, loc="lower right")
fig.tight_layout()
fig.savefig(FIG_E2_TORN, dpi=200)
plt.close(fig)

FIG_E2_SENS = os.path.join(OUT_DIR, "fig_e2_sensibilidade_mc.png")
fig_e2_sens_size = (10, 2.9)
fig, axes = plt.subplots(1, 3, figsize=fig_e2_sens_size, sharey=True)
x_ = np.arange(len(MACRO_COLS))
for ax, nome in zip(axes, nomes_setores):
    for off, ind, cor in ((-0.2, "Lucro líquido", "#2B6CB0"), (0.2, "Liquidez corrente", "#718096")):
        q = E2["sens_mc"].query("setor == @nome and indicador == @ind")
        ax.bar(x_ + off, q["beta_padron"].values, width=0.38, color=cor, label=ind)
    ax.axhline(0, color="#4A5568", lw=0.7)
    ax.set_xticks(x_); ax.set_xticklabels([MACRO_LABEL[k] for k in MACRO_COLS], fontsize=7)
    r2_l = E2["sens_mc"].query("setor == @nome and indicador == 'Lucro líquido'")["r2_macro"].iloc[0]
    ax.set_title(f"{nome} (R² macro do lucro = {virg(r2_l)})", fontsize=8.5)
axes[0].set_ylabel("Coef. padronizado", fontsize=8)
axes[0].legend(fontsize=7, frameon=False, loc="lower left")
fig.tight_layout()
fig.savefig(FIG_E2_SENS, dpi=200)
plt.close(fig)
FIG3 = os.path.join(OUT_DIR, "liquidez_corrente_simulada.png")
fig3_size = (10, 2.9)
fig, axes = plt.subplots(1, 3, figsize=fig3_size)
for ax, nome in zip(axes, nomes_setores):
    lc = fin_mc[nome]["r"]["lc"]
    lo_, hi_ = np.percentile(lc, [0.5, 99.5])
    ax.hist(lc, bins=60, range=(lo_, hi_), color=CORES[nome], alpha=0.75)
    ax.axvline(1.0, color="#C53030", lw=1.2, ls="--", label="LC = 1,0")
    ax.axvline(PERFIS[nome]["LC"], color="#2D3748", lw=1.2, label="LC do ano-base")
    ax.set_title(nome, fontsize=9)
    ax.set_yticks([])
    ax.set_xlabel("Liquidez corrente ao fim do ano", fontsize=8)
    ax.tick_params(labelsize=7)
axes[0].legend(fontsize=7, frameon=False, loc="upper left")
fig.tight_layout()
fig.savefig(FIG3, dpi=200)
plt.close(fig)


# ==============================================================================
# 6. RELATÓRIO PDF
# ==============================================================================
PDF_PATH = os.path.join(OUT_DIR, "relatorio_risco_pmes.pdf")

try:   # DejaVu acompanha o matplotlib: cobre Δ, ±, ≈, − e acentos em qualquer ambiente (Colab inclusive)
    _fd = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
    pdfmetrics.registerFont(TTFont("DejaVu", os.path.join(_fd, "DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", os.path.join(_fd, "DejaVuSans-Bold.ttf")))
    pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold", italic="DejaVu", boldItalic="DejaVu-Bold")
    FONT, FONTB = "DejaVu", "DejaVu-Bold"
except Exception:
    FONT, FONTB = "Helvetica", "Helvetica-Bold"


def num(x: float, d: int = 1) -> str:
    """Formata número no padrão brasileiro (milhar com ponto, decimal com vírgula)."""
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x: float, d: int = 1) -> str:
    """Número no padrão brasileiro seguido de %."""
    return num(x, d) + "%"


def pp(x: float, d: int = 1) -> str:
    """Número no padrão brasileiro seguido de pp (pontos percentuais)."""
    return num(x, d) + " pp"


def sg(x: float, d: int = 1) -> str:
    """Número com sinal explícito (+ ou −)."""
    return ("+" if x >= 0 else "−") + num(abs(x), d)


def pfmt(p: float) -> str:
    """p-valor formatado ('p < 0,01' ou 'p = 0,xx')."""
    return "p &lt; 0,01" if p < 0.005 else "p = " + num(p, 2)


def estrelas(p: float) -> str:
    """Estrelas de significância: *** p<0,01; ** p<0,05; * p<0,10."""
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.10 else ""


def mil(x: float) -> str:
    """R$ → R$ mil, formatado."""
    return num(x / 1000, 0)


def nan_dash(x: Any, f: Callable[[Any], str]) -> str:
    """Aplica `f` a `x`, ou devolve um travessão se `x` for ausente/NaN."""
    return "—" if (x is None or (isinstance(x, float) and np.isnan(x))) else f(x)


class NumberedCanvas(canvas.Canvas):
    """Canvas do ReportLab que numera as páginas como 'Página X de N' e desenha o rodapé."""

    def __init__(self, *a: Any, **k: Any) -> None:
        super().__init__(*a, **k)
        self._saved = []

    def showPage(self) -> None:
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        total = len(self._saved)
        for st in self._saved:
            self.__dict__.update(st)
            self.saveState()
            self.setFont(FONT, 8)
            self.setFillColor(colors.black)
            self.setStrokeColor(colors.HexColor("#E2E8F0"))
            self.setLineWidth(0.5)
            self.line(40, 40, A4[0] - 40, 40)
            self.drawString(40, 25, f"Relatório de Risco Macroeconômico para PMEs | Página {self._pageNumber} de {total}")
            self.drawRightString(A4[0] - 40, 25, "CONFIDENCIAL")
            self.restoreState()
            super().showPage()
        super().save()


def leitura_oos(r: float) -> str:
    """Rótulo qualitativo para o R² fora da amostra."""
    return ("pior que a média histórica" if r < 0 else "fraco" if r < 0.2 else "moderado" if r < 0.5 else "bom")


def gerar_pdf() -> None:
    """Monta o relatório em PDF (ReportLab) a partir dos resultados globais das Etapas 1 e 2."""
    doc = SimpleDocTemplate(PDF_PATH, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=45, bottomMargin=60,
                            title="Relatório de Risco Macroeconômico para PMEs")
    W = A4[0] - 80
    st = getSampleStyleSheet()
    c_line = colors.HexColor("#CBD5E1")
    title = ParagraphStyle("t", parent=st["Heading1"], fontName=FONTB, fontSize=19, textColor=colors.black, spaceAfter=4, leading=23)
    sub = ParagraphStyle("s", parent=st["Normal"], fontName=FONT, fontSize=9.5, spaceAfter=10, leading=13)
    parte = ParagraphStyle("pt", parent=st["Heading1"], fontName=FONTB, fontSize=13, textColor=colors.white,
                           backColor=colors.HexColor("#2D3748"), borderPadding=(5, 6, 5, 6), spaceBefore=14, spaceAfter=10, leading=16)
    h1 = ParagraphStyle("h", parent=st["Heading2"], fontName=FONTB, fontSize=11.5, textColor=colors.black, spaceBefore=12, spaceAfter=6)
    body = ParagraphStyle("b", parent=st["BodyText"], fontName=FONT, fontSize=8.8, spaceAfter=6, leading=13, alignment=4)
    note = ParagraphStyle("n", parent=body, fontSize=7.4, leading=10.5, spaceAfter=6)
    tt = ParagraphStyle("tt", parent=st["Normal"], fontName=FONT, fontSize=7.6, alignment=1, leading=10)
    th = ParagraphStyle("th", parent=tt, fontName=FONTB)
    tl = ParagraphStyle("tl", parent=tt, alignment=0)
    bul = ParagraphStyle("bul", parent=body, leftIndent=12, bulletIndent=0, spaceAfter=3)

    def tabela(dados, widths, pad=4):
        t = Table(dados, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, c_line),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("TOPPADDING", (0, 0), (-1, -1), pad), ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
            ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]))
        return t

    P = lambda txt, sty=tt: Paragraph(txt, sty)
    PL = lambda txt: Paragraph(txt, tl)
    H = lambda txt: Paragraph(f"<b>{txt}</b>", th)
    img = lambda path, size: Image(path, width=W, height=W * size[1] / size[0])
    story = []

    # ---------------- capa / aviso ----------------
    story.append(Paragraph("Risco Macroeconômico para PMEs", title))
    barra = Table([[""]], colWidths=[W], rowHeights=[3])
    barra.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), colors.black)]))
    story += [barra, Spacer(1, 8)]
    story.append(Paragraph(
        "<b>Aviso metodológico:</b> as variáveis macroeconômicas (Selic, IPCA, câmbio e IBC-Br) e a atividade setorial vêm de fontes "
        "oficiais (BCB e IBGE/SIDRA). O modelo usa <b>Selic, IPCA e câmbio</b>; o IBC-Br (PIB) é mantido apenas como contexto descritivo, "
        "porque agrega a própria atividade setorial e sobreporia o resultado. A atividade setorial é medida pelos <b>índices de volume dessazonalizados</b> do IBGE (PMS, serviços de "
        "tecnologia da informação, subgrupo 2.1.2; PIM-PF, indústria geral; PMC, comércio varejista; base 2022 = 100), usados como proxy do "
        "faturamento: medem quantidade, não receita nominal, e refletem o setor inteiro, não apenas PMEs."
        " As inferências estatísticas pertencem à Etapa 1; as empresas fictícias entram só na Etapa 2.", note))

    # ---------------- dados auxiliares para o texto ----------------
    ce = E1["corr"]
    sp = E1["specs"]
    _ce_mod = ce[ce.chave.isin(MACRO_COLS)]
    top = {s: _ce_mod[_ce_mod.setor == s].loc[_ce_mod[_ce_mod.setor == s]["pearson"].abs().idxmax()] for s in nomes_setores}
    r_pib = {s: float(ce[(ce.setor == s) & (ce.chave == "PIB")]["pearson"].iloc[0]) for s in nomes_setores}
    ord_cvar = sorted(nomes_setores, key=lambda s: -metr[s]["cvar95"])
    top_risco, low_risco = ord_cvar[0], ord_cvar[-1]

    def esp_base(s, col):
        return sp[(sp.setor == s) & (sp.modelo == ESPEC_BASE)][col].iloc[0]

    def espec(s, mod, col):
        return sp[(sp.setor == s) & (sp.modelo == mod)][col].iloc[0]

    r2_med = np.mean([esp_base(s, "r2") for s in nomes_setores])
    mudancas = []
    for s in nomes_setores:
        for k in MACRO_COLS:
            b1_, p1_ = espec(s, ESPEC_BASE, f"b_{k}"), espec(s, ESPEC_BASE, f"p_{k}")
            b4_, p4_ = espec(s, ESPEC_COVID, f"b_{k}"), espec(s, ESPEC_COVID, f"p_{k}")
            if (p1_ < 0.10) != (p4_ < 0.10) or np.sign(b1_) != np.sign(b4_):
                mudancas.append(f"{MACRO_LABEL[k]} em {s} ({num(b1_, 2)}{estrelas(p1_)} → {num(b4_, 2)}{estrelas(p4_)})")
    ll = E1["lead_lag"]

    def melhor_lag_selic(s):
        col = ll[f"{s} | {MACRO_LABEL['Selic']}"]
        L_ = int(col.abs().idxmax())
        return L_, float(col.loc[L_]), float(col.loc[0])

    # ---------------- Sumário executivo ----------------
    story.append(Paragraph("Sumário Executivo", h1))
    story.append(Paragraph(
        "<b>Objetivo.</b> Medir o quanto o crescimento em 12 meses da atividade de Serviços de TI, Indústria e Comércio "
        "Varejista depende de Selic, IPCA e câmbio (correlação e regressão, 2012–2026). O relatório tem duas etapas: <b>Etapa 1</b> — cenários macroeconômicos e os três setores; "
        "<b>Etapa 2</b> — empresas fictícias (saúde financeira e aplicação dos cenários).", body))
    txt_top = "; ".join(f"{s}: {MACRO_LABEL[top[s]['chave']]} (r = {num(top[s]['pearson'], 2)})" for s in nomes_setores)
    story.append(Paragraph(
        f"<b>Etapa 1 — correlação.</b> Maior correlação contemporânea entre as variáveis do modelo, por setor: {txt_top}. "
        f"O PIB (IBC-Br), mantido só como contexto, correlaciona-se com os setores em "
        + ", ".join(f"r = {num(r_pib[s], 2)} ({s})" for s in nomes_setores)
        + " — valores inflados por construção, já que o IBC-Br agrega a própria atividade setorial (ver Limitações da Etapa 1). "
        f"A política monetária atua com defasagem: em "
        f"{nomes_setores[0]}, r(Selic) vai de {num(melhor_lag_selic(nomes_setores[0])[2], 2)} (mesmo mês) a "
        f"{num(melhor_lag_selic(nomes_setores[0])[1], 2)} com {melhor_lag_selic(nomes_setores[0])[0]} meses de defasagem.", body))
    r2_txt = "; ".join(f"{s} {num(esp_base(s, 'r2'), 2)} / {num(esp_base(s, 'r2_cv'), 2)} / {num(esp_base(s, 'r2_wf'), 2)}" for s in nomes_setores)
    story.append(Paragraph(
        f"<b>Etapa 1 — regressão (modelo único: Selic, IPCA e câmbio).</b> R² dentro da amostra / validação cruzada em blocos / walk-forward com embargo de {EMBARGO_WF} meses: "
        f"{r2_txt}. O poder preditivo <i>no tempo</i> é "
        + (f"{leitura_oos(esp_base(nomes_setores[0], 'r2_wf'))} nos três setores"
           if len({leitura_oos(esp_base(s, 'r2_wf')) for s in nomes_setores}) == 1
           else ", ".join(f"{leitura_oos(esp_base(s, 'r2_wf'))} em {s}" for s in nomes_setores))
        + f". O R² médio é {num(r2_med, 2)}: o modelo mede a <b>sensibilidade histórica</b> do setor a juros, preços e câmbio, não prevê o faturamento. "
          "O PIB ficou fora de propósito, para não incorporar a atividade do próprio setor ao resultado.", body))
    story.append(Paragraph(
        f"<b>Risco simulado.</b> {top_risco} tem o maior CVaR 95% ({pp(metr[top_risco]['cvar95'])} abaixo do esperado) e "
        f"{low_risco} o menor ({pp(metr[low_risco]['cvar95'])}); a probabilidade de contração é de "
        + ", ".join(f"{pct(metr[s]['p_contracao'])} ({s})" for s in nomes_setores) + ".", body))
    story.append(Paragraph(
        "<b>Uso recomendado.</b> Ferramenta de stress-test e comunicação de risco; não é previsão pontual de faturamento de uma PME específica. "
        "As relações são associações históricas, não causalidade.", body))

    # ---------------- 1. Escopo ----------------
    story.append(Paragraph("1. Escopo, dados e método", h1))
    _ini = {k: v.index.min().strftime("%m/%Y") for k, v in setor_y.items()}
    _fim = max(v.index.max() for v in setor_y.values()).strftime("%m/%Y")
    dados = [[H("Série"), H("Fonte"), H("Tratamento"), H("Cobertura")],
             [PL("Selic"), PL("BCB SGS 11 (diária)"), PL("Taxa anualizada (252 d.u.) do último dia útil do mês, % a.a."), P(f"{selic_d.index.min():%m/%Y}–{selic_d.index.max():%m/%Y}")],
             [PL("IPCA"), PL("IBGE (mensal)"), PL("Acumulado em 12 meses (%)"), P(f"até {ultimo_mes_ipca:%m/%Y}")],
             [PL("Câmbio"), PL("BCB PTAX compra (diária)"), PL("Variação em 12 meses do BRL/USD, fim do mês (%)"), P(f"{cambio_d.index.min():%m/%Y}–{cambio_d.index.max():%m/%Y}")],
             [PL("PIB"), PL("BCB IBC-Br (mensal)"), PL("Variação sobre o mesmo mês do ano anterior (%). Apenas contexto descritivo: fora do modelo, dos cenários e da simulação"), P(f"até {ultimo_mes_ibc:%m/%Y}")],
             [PL("Atividade setorial"), PL("IBGE SIDRA: PMS 8688 (TI, 2.1.2), PIM-PF 8888 (indústria geral), PMC 8880 (varejo)"),
              PL("Variação em 12 meses do índice de volume dessazonalizado (%)"),
              P(f"TI {_ini['Serviços de TI']}; indústria {_ini['Indústria']}; varejo {_ini['Comércio Varejista']}; até {_fim}")]]
    story.append(tabela(dados, [70, 150, 195, 100]))
    story.append(Spacer(1, 5))
    story.append(Paragraph(
        f"A estimação usa os <b>{n_obs} meses de {MES_INI_D:%m/%Y} a {MES_FIM_D:%m/%Y}</b>, intervalo comum às três séries setoriais. Como as taxas de "
        f"12 meses de meses vizinhos se sobrepõem, a amostra equivale a cerca de <b>{ANOS_EF:.0f} anos independentes</b>. Por isso: erros-padrão de "
        f"Newey-West ({HAC_LAGS} defasagens) com t de Student (g.l. = anos independentes − parâmetros); IC de correlação por bootstrap circular em blocos de "
        f"{BLOCO_BOOT_CORR} meses; validação cruzada que retira ±{H_CV_BLOCO} meses em torno do mês previsto; e walk-forward com embargo de {EMBARGO_WF} meses. "
        f"A simulação de Monte Carlo sorteia {N_SIM:,}".replace(",", ".") + f" cenários macro por uma t-Student multivariada ({NU} g.l., covariância histórica, "
        "variância preservada), incorpora a incerteza dos coeficientes (covariância HAC) e reamostra os resíduos de forma conjunta entre setores.", body))

    # ================= ETAPA 1 =================
    story.append(Paragraph("ETAPA 1 — Cenários macroeconômicos e os três setores (2012–2026)", parte))

    # ---------------- 2. Descritiva ----------------
    story.append(Paragraph("2. Estatísticas descritivas e estacionariedade", h1))
    linhas = [[H("Série"), H("Média"), H("Desvio-padrão"), H("Mínimo"), H("Máximo"), H("ADF (t)"), H("Conclusão ADF")]]
    for _, q in E1["desc"].iterrows():
        linhas.append([PL(f"<b>{q['serie']}</b>"), P(num(q["media"], 1)), P(num(q["desvio"], 1)), P(num(q["minimo"], 1)),
                       P(num(q["maximo"], 1)), P(num(q["adf_t"], 2)), PL(q["adf_conclusao"])])
    story.append(tabela(linhas, [95, 42, 58, 45, 45, 45, 185]))
    story.append(Paragraph(
        "Todas as séries estão em % (Selic em % a.a.; demais, variação em 12 meses). ADF com constante e defasagens por BIC; valores críticos "
        "assintóticos de −3,43 (1%), −2,86 (5%) e −2,57 (10%). O poder do teste é baixo com ~15 anos efetivos: raiz unitária não rejeitada "
        "não prova não estacionariedade, mas alerta para o risco de regressão espúria — daí a ênfase em erros-padrão HAC e validação fora da amostra.", note))

    # ---------------- 3. Correlação ----------------
    story.append(Paragraph("3. Análise de correlação", h1))
    story.append(Paragraph(
        "A tabela traz a correlação de Pearson (linear) e de Spearman (por postos, menos sensível aos extremos da COVID) entre cada variável macro "
        "e o crescimento setorial no mesmo mês, e a correlação parcial de Pearson (controlando pelas outras duas variáveis do modelo; para o PIB, que não entra no modelo, não se calcula). O IC95% é de "
        "bootstrap circular em blocos de 12 meses; o p-valor usa t com n efetivo (anos independentes). Em negrito, correlações cujo IC exclui zero.", body))
    linhas = [[H("Setor"), H("Variável"), H("Pearson"), H("IC95% (bloco-bootstrap)"), H("p (n efetivo)"), H("Spearman"), H("Parcial")]]
    for _, q in ce.iterrows():
        fm = (lambda t: f"<b>{t}</b>") if q["sig_ic"] else (lambda t: t)
        linhas.append([PL(f"<b>{q['setor']}</b>") if q["chave"] == "Selic" else P(""), PL(q["variavel"]), P(fm(num(q["pearson"], 2))),
                       P(f"[{num(q['p_lo'], 2)}; {num(q['p_hi'], 2)}]"), P(pfmt(q["p_ef"])), P(num(q["spearman"], 2)), P(nan_dash(q["parcial"], lambda v: num(v, 2)))])
    story.append(tabela(linhas, [95, 75, 50, 100, 60, 55, 50], pad=3))
    story.append(Spacer(1, 4))
    cm = E1["corr_macro"]
    pares = [(MACRO_COLS[i], MACRO_COLS[j], cm.iloc[i, j]) for i in range(len(MACRO_COLS)) for j in range(i + 1, len(MACRO_COLS))]
    mx = max(pares, key=lambda t: abs(t[2]))
    story.append(Paragraph(
        f"<b>Leitura.</b> Entre as variáveis do modelo, a maior correlação é {MACRO_LABEL[mx[0]]}–{MACRO_LABEL[mx[1]]} (r = {num(mx[2], 2)}); VIF máximo "
        f"{num(np.max(vifs), 1)}, sem colinearidade preocupante. A correlação parcial difere da simples quando as variáveis macro se movem juntas "
        "(ex.: Selic e IPCA). Correlações entre setores: "
        + "; ".join(f"{a}–{b} r = {num(E1['corr_setores'].loc[a, b], 2)}" for a, b in [(nomes_setores[0], nomes_setores[1]),
                                                                                       (nomes_setores[0], nomes_setores[2]),
                                                                                       (nomes_setores[1], nomes_setores[2])]) + ".", body))
    story.append(img(FIG_E1_CORR, fig_e1_corr_size))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        "<b>Correlação defasada.</b> O gráfico mostra a correlação entre a variável macro em t−L e o crescimento setorial em t, L = 0…12. "
        "Resultado exploratório: ao olhar 13 defasagens × 4 variáveis × 3 setores há risco de comparações múltiplas; por isso a defasagem usada na checagem de robustez "
        f"é fixada a priori em {LAG_SELIC_ROB} meses (configurável), e não escolhida pelo máximo da correlação. O PIB aparece nos gráficos só como referência de contexto.", body))
    story.append(img(FIG_E1_LAG, fig_e1_lag_size))
    story.append(Paragraph(
        "Selic — maior |r| entre L = 0 e 12: " + "; ".join(
            f"{s}: L = {melhor_lag_selic(s)[0]} (r = {num(melhor_lag_selic(s)[1], 2)}; mesmo mês: {num(melhor_lag_selic(s)[2], 2)})" for s in nomes_setores) + ".", note))
    story.append(img(FIG_E1_DISP, fig_e1_disp_size))
    story.append(Paragraph(
        "Dispersão com reta de MQO univariada. Os pontos vermelhos são março/2020–junho/2021: poucos pontos extremos alavancam as retas, "
        "o que motiva a checagem sem COVID e a correlação de Spearman. O painel do PIB é contexto: a variável não entra no modelo.", note))

    # ---------------- 4. Regressão ----------------
    story.append(Paragraph("4. Análise de regressão", h1))
    story.append(Paragraph(
        "Para cada setor estima-se por MQO um <b>modelo único</b>: <i>crescimento_setorial(t) = α + β<sub>1</sub>·Selic + β<sub>2</sub>·IPCA + β<sub>3</sub>·Câmbio + ε</i>. "
        "Cada coeficiente é a variação, em pontos percentuais, do crescimento setorial quando a variável sobe 1 pp, mantidas as demais constantes.", body))
    story.append(Paragraph(
        "<b>Por que o PIB não entra.</b> O IBC-Br agrega indústria, comércio e serviços: incluí-lo faria o coeficiente do PIB carregar um efeito mecânico (o setor é parte do índice) "
        "e atribuiria à 'atividade' parte do que é o próprio setor. Além disso, a Selic atua em boa medida via atividade; sem o PIB, o coeficiente da Selic capta o efeito total "
        "(direto e via atividade), que é o relevante para um teste de estresse. O custo é um R² menor, assumido de forma explícita nas seções seguintes.", body))
    cab = [H("Setor")] + [H(MACRO_LABEL[k]) for k in MACRO_COLS]
    linhas = [cab]
    for s in nomes_setores:
        m = modelos[s]
        row = [PL(f"<b>{s}</b>")]
        for j, k in enumerate(MACRO_COLS, start=1):
            row.append(P(f"{num(m['beta'][j], 2)}{estrelas(m['p'][j])}<br/><font size=6.5>({pfmt(m['p'][j])})</font><br/>"
                         f"<font size=6 color='#4A5568'>IC95% [{num(m['ci_lo'][j], 2)}; {num(m['ci_hi'][j], 2)}]</font>"))
        linhas.append(row)
    story.append(KeepTogether([Paragraph("<b>4.1 Modelo único: coeficientes</b>", body), tabela(linhas, [140, 124, 124, 124]),
        Paragraph("Significância: *** p&lt;0,01; ** p&lt;0,05; * p&lt;0,10. Erros-padrão e IC95% de Newey-West; p-valores e IC com t de Student e "
                  f"g.l. = anos independentes − parâmetros (≈ {max(int(n_obs / 12) - 4, 3)}). Coeficientes não significativos devem ser lidos como incertos, não como zero.", note)]))
    cab = [H("Setor"), H("R<super>2</super>"), H("R<super>2</super> aj."), H("R<super>2</super> CV em blocos"), H("R<super>2</super> walk-forward"), H("p (F robusto)"), H("Durbin-<br/>Watson")]
    linhas = [cab] + [[PL(f"<b>{s}</b>"), P(num(modelos[s]["r2"], 2)), P(num(modelos[s]["r2_adj"], 2)), P(num(modelos[s]["r2_cv"], 2)),
                       P(num(modelos[s]["r2_wf"], 2)), P(num(modelos[s]["p_F"], 3)), P(num(modelos[s]["dw"], 2))] for s in nomes_setores]
    story.append(KeepTogether([Paragraph("<b>4.2 Qualidade do ajuste e validação fora da amostra</b>", body), tabela(linhas, [112, 52, 55, 85, 85, 65, 65]),
        Paragraph(f"<b>R² CV em blocos</b>: o modelo é re-estimado sem os meses ±{H_CV_BLOCO} em torno do mês previsto (benchmark: média do treino); usa dados futuros no treino e mede "
                  f"estabilidade. <b>R² walk-forward</b>: só usa o passado, com embargo de {EMBARGO_WF} meses e benchmark = média histórica expansiva; é o teste de previsão no tempo, "
                  "e valores negativos indicam que o modelo erra mais do que simplesmente prever a média histórica. Durbin-Watson baixo é esperado com taxas de 12 meses "
                  "sobrepostas (resíduos vizinhos quase iguais); por isso usa-se covariância HAC.", note)]))

    def fm_coef(b, p):
        return "—" if np.isnan(b) else f"{num(b, 2)}{estrelas(p)}"

    linhas = [[H("Setor / modelo"), H("n"), H("Selic"), H("IPCA"), H("Câmbio"), H("R<super>2</super>"), H("R<super>2</super> CV"), H("R<super>2</super> WF")]]
    for s in nomes_setores:
        for i, mod in enumerate(ESPECS):
            q = sp[(sp.setor == s) & (sp.modelo == mod)].iloc[0]
            linhas.append([PL(f"<b>{s}</b><br/>{mod}") if i == 0 else PL(mod), P(f"{int(q['n'])}")] +
                          [P(fm_coef(q[f"b_{k}"], q[f"p_{k}"])) for k in MACRO_COLS] +
                          [P(num(q["r2"], 2)), P(nan_dash(q["r2_cv"], lambda v: num(v, 2))), P(nan_dash(q["r2_wf"], lambda v: num(v, 2)))])
    story.append(KeepTogether([
        Paragraph("<b>4.3 Robustez do modelo único</b>", body),
        tabela(linhas, [160, 30, 62, 62, 62, 36, 40, 40], pad=3),
        Paragraph(f"Modelo único: Selic, IPCA e câmbio no mesmo mês; {ESPEC_LAG}: Selic defasada {LAG_SELIC_ROB} meses; {ESPEC_COVID}: excluindo "
                  f"{COVID_INI[5:]}/{COVID_INI[:4]} a {COVID_FIM[5:]}/{COVID_FIM[:4]} (a janela retirada quebra a contiguidade da série, então CV e walk-forward não são calculados). "
                  "Significância: *** p&lt;0,01; ** p&lt;0,05; * p&lt;0,10 (Newey-West).", note)]))
    story.append(Paragraph(
        f"<b>Leitura.</b> (i) <b>Poder explicativo:</b> o R² médio do modelo único é {num(r2_med, 2)}; o restante da variação setorial vem de fatores que juros, preços e câmbio não captam "
        "(atividade, choques setoriais, política fiscal), e deve ser lido como risco não explicado. (ii) <b>Defasagem da Selic:</b> "
        + "; ".join(f"{s} R² {num(esp_base(s, 'r2'), 2)} → {num(espec(s, ESPEC_LAG, 'r2'), 2)} (WF {num(esp_base(s, 'r2_wf'), 2)} → {num(espec(s, ESPEC_LAG, 'r2_wf'), 2)})" for s in nomes_setores)
        + ". (iii) <b>COVID:</b> "
        + ("mudam de sinal ou de significância: " + "; ".join(mudancas) + "." if mudancas else "nenhum coeficiente muda de sinal ou de significância.")
        + " Resultados que dependem dessa janela devem ser apresentados com ressalva.", body))
    gl = E1["grade_lag"]
    linhas = [[H("Setor")] + [H(f"L = {L_}") for L_ in LAGS_GRADE]]
    for s in nomes_setores:
        row = [PL(f"<b>{s}</b>")]
        for L_ in LAGS_GRADE:
            q = gl[(gl.setor == s) & (gl.defasagem == L_)].iloc[0]
            row.append(P(f"{num(q['b_Selic'], 2)}{estrelas(q['p_Selic'])}<br/><font size=6>R² {num(q['r2'], 2)} | WF {num(q['r2_wf'], 2)}</font>"))
        linhas.append(row)
    story.append(KeepTogether([Paragraph("<b>4.4 Grade de defasagem da Selic</b> (coeficiente da Selic em t−L, com R² e R² walk-forward do modelo completo)", body),
                               tabela(linhas, [112] + [80.6] * 5), Spacer(1, 3)]))
    story.append(img(FIG_E1_COEF, fig_e1_coef_size))
    story.append(Paragraph("Efeito de +1 desvio-padrão histórico de cada variável (coeficiente × desvio-padrão), com IC95% Newey-West, no modelo único.", note))
    story.append(img(FIG2, fig2_size))
    story.append(Paragraph("Crescimento observado × ajustado pelo modelo único (dentro da amostra).", note))

    # ---------------- 5. Monte Carlo ----------------
    cab = [H("Setor"), H("Cresc. esperado"), H("Desvio-padrão"), H("VaR 95%"), H("VaR 99%"), H("CVaR 95%"), H("Prob. de contração"), H("Pior 1%"), H("Mín. histórico")]
    linhas = [cab]
    for s in nomes_setores:
        m = metr[s]
        linhas.append([PL(f"<b>{s}</b>"), P(pct(m["esperado"])), P(pp(m["desvio"])), P(pp(m["var95"])), P(pp(m["var99"])), P(pp(m["cvar95"])),
                       P(pct(m["p_contracao"])), P(pct(m["pior1"])), P(f"{pct(m['hist_min'])}<br/>({m['hist_min_ano']})")])
    story.append(KeepTogether([
        Paragraph("5. Simulação de Monte Carlo e métricas de risco (modelo único)", h1),
        Paragraph(f"Foram simulados {N_SIM:,}".replace(",", ".") + " cenários de um ano. As perdas (VaR e CVaR) são a <b>queda do crescimento em relação ao esperado</b>, em pp: "
                  "o VaR 95% é a queda superada em apenas 5% dos cenários e o CVaR 95% é a queda média nesses 5% piores. \"Prob. de contração\" é a fração de cenários com "
                  "crescimento abaixo de zero; \"Pior 1%\" é o crescimento no 1º percentil. A simulação é <b>incondicional</b>: sorteia cenários macro em torno da média "
                  f"histórica de {ANO_INI}–{ANO_FIM}, não do cenário atual.", body),
        tabela(linhas, [90, 50, 50, 45, 45, 50, 55, 48, 62], pad=3),
        Paragraph(f"\"Mín. histórico\": pior taxa de 12 meses observada entre {MES_INI_D:%m/%Y} e {MES_FIM_D:%m/%Y}.", note)]))
    story.append(img(FIG1, fig1_size))
    story.append(img(FIG_MACRO, fig_macro_size))
    story.append(Paragraph("Padrão histórico das variáveis macro (base empírica do Monte Carlo): média, volatilidade e correlação, no mesmo período da amostra.", note))

    # ---------------- 6. Cenários ----------------
    def fmt_cen(c):
        return f"Selic {pct(c['Selic'], 2)} | IPCA {pct(c['IPCA'])} | Câmbio {pct(c['Cambio_var'])}"

    fora_mild = [MACRO_LABEL[k] for k in MACRO_COLS if not (df[k].min() <= cenario_estresse_mild[k] <= df[k].max())]
    rot_mild = ("(dentro da amostra)" if not LIMITADAS_MILD else "(limitado à faixa histórica: " + ", ".join(LIMITADAS_MILD) + ")") if LIMITAR_FAIXA \
        else ("(dentro da amostra)" if not fora_mild else "(parcialmente fora da faixa: " + ", ".join(fora_mild) + ")")
    rot_sev = ("(dentro da amostra)" if not LIMITADAS_SEV else "(limitado à faixa histórica: " + ", ".join(LIMITADAS_SEV) + ")") if LIMITAR_FAIXA else "(extrapolado)"
    cab = [H("Cenário"), H("Premissas macro")] + [H(s) for s in nomes_setores]
    linhas = [cab,
              [PL("<b>Atual</b><br/>(últimos 12 meses)"), PL(fmt_cen(cenario_atual))] + [P(pct(pred_atual[s])) for s in nomes_setores],
              [PL(f"<b>Estresse moderado</b><br/>{rot_mild}"), PL(fmt_cen(cenario_estresse_mild))] +
              [P(f"{pct(pred_estr_mild[s])}<br/><font size=6.5>({sg(pred_estr_mild[s] - pred_atual[s])} pp)</font>") for s in nomes_setores],
              [PL(f"<b>Estresse severo</b><br/>{rot_sev}"), PL(fmt_cen(cenario_estresse))] +
              [P(f"{pct(pred_estr[s])}<br/><font size=6.5>({sg(pred_estr[s] - pred_atual[s])} pp)</font>") for s in nomes_setores]]
    fora = [MACRO_LABEL[k] for k in MACRO_COLS if not (df[k].min() <= cenario_estresse[k] <= df[k].max())]
    _avisos = []
    if fora_mild:
        _avisos.append("o <b>estresse moderado</b> fica fora da faixa histórica da amostra para: " + ", ".join(fora_mild))
    if fora:
        _avisos.append("o <b>estresse severo</b> fica fora da faixa histórica da amostra para: " + ", ".join(fora))
    aviso_fora = ("Atenção: " + "; ".join(_avisos) + "; nessas variáveis a previsão extrapola o modelo."
                  if _avisos else "Ambos os cenários de estresse estão dentro da faixa histórica da amostra.")
    if LIMITAR_FAIXA:
        _lim_txt = "; ".join(f"{MACRO_LABEL[k]}: {pct(FAIXA_HIST[k][0], 1)} a {pct(FAIXA_HIST[k][1], 1)}" for k in MACRO_COLS)
        aviso_fora = (f"Sem extrapolação: cada variável dos estresses fica limitada ao mínimo e ao máximo observados na amostra ({_lim_txt}); "
                      "o choque efetivo pode ser menor que o nominal quando a variável já está perto do máximo histórico.")
    story.append(KeepTogether([
        Paragraph("6. Cenários determinísticos", h1),
        Paragraph(f"O cenário atual usa a posição mais recente das três variáveis do modelo: Selic em {data_ref:%d/%m/%Y}, câmbio em {data_ref:%d/%m/%Y} (variação sobre 12 meses antes) "
                  f"e IPCA em 12 meses até {ultimo_mes_ipca:%m/%Y}. O <b>estresse moderado</b> soma (antes do limite histórico) Selic {sg(STRESS_MILD['Selic'])} pp, IPCA "
                  f"{sg(STRESS_MILD['IPCA'])} pp e câmbio {sg(STRESS_MILD['Cambio_var'], 0)} pp; o <b>severo</b> usa {sg(STRESS['Selic'])} / "
                  f"{sg(STRESS['IPCA'])} / {sg(STRESS['Cambio_var'], 0)} pp. Os estresses não têm premissa de PIB (fora do modelo). Valores são previsões pontuais do modelo único.", body),
        tabela(linhas, [92, 160, 87, 87, 87]), Paragraph(aviso_fora, note)]))

    # ---------------- Conclusões da Etapa 1 (macro e setores) ----------------
    concl = [Paragraph("Conclusões da Etapa 1 — cenários macro e setores", h1)]
    concl.append(Paragraph(
        f"<b>Exposição a cenários adversos.</b> {top_risco} é o setor com maior CVaR 95% ({pp(metr[top_risco]['cvar95'])}); {low_risco} tem o menor "
        f"({pp(metr[low_risco]['cvar95'])}). A maior probabilidade de contração é a de {max(nomes_setores, key=lambda s: metr[s]['p_contracao'])} "
        f"({pct(max(metr[s]['p_contracao'] for s in nomes_setores))}). Para TI, parte relevante do risco simulado é residual (não explicada pelo macro).", body))
    for s in nomes_setores:
        m = modelos[s]
        efeito = np.abs(m["beta"][1:] * sd_macro)
        j = int(np.argmax(efeito))
        k = MACRO_COLS[j]
        sig = "estatisticamente significativo" if m["p"][j + 1] < 0.10 else "sem significância estatística"
        concl.append(Paragraph(
            f"<b>{s}.</b> Principal vetor no modelo: {MACRO_LABEL[k]} — +1 desvio-padrão ({num(sd_macro[j], 1)} pp) move o crescimento em cerca de {pp(m['beta'][j + 1] * sd_macro[j])} ({sig}). "
            f"R² {num(m['r2'], 2)}; CV em blocos {num(m['r2_cv'], 2)}; walk-forward {num(m['r2_wf'], 2)} ({leitura_oos(m['r2_wf'])}).", body))
    concl.append(Paragraph("<b>Limitações da Etapa 1</b>", body))
    for txt in [
        f"Amostra de {n_obs} meses, mas com taxas de 12 meses sobrepostas: ~{ANOS_EF:.0f} anos independentes para 4 parâmetros por modelo. Estimativas imprecisas e sensíveis a poucos episódios (2015–2016 e 2020).",
        "O IBC-Br inclui a atividade de indústria, comércio e serviços: sua correlação com os índices setoriais tem componente mecânico (endogeneidade) e por isso ele ficou fora do modelo, dos cenários e da simulação, aparecendo só como contexto. A consequência é um poder explicativo menor: o modelo mede sensibilidade histórica a juros, preços e câmbio, não prevê faturamento.",
        "As séries são persistentes (a raiz unitária não é rejeitada para algumas) e o poder dos testes é baixo; a regressão mostra associação histórica, não causalidade.",
        "Os resultados dependem de poucos meses extremos da COVID (checagem sem COVID) e a defasagem da Selic é fixada a priori; escolher a defasagem pelos dados exigiria validação fora da amostra independente.",
        f"A simulação de Monte Carlo é incondicional (centrada na média {ANO_INI}–{ANO_FIM}) e os sorteios e estresses ficam limitados à faixa histórica das variáveis, sem extrapolação (ver seção 6); o limite reduz a volatilidade simulada e o choque efetivo de Selic quando ela já está perto do máximo histórico.",
        "As séries setoriais são índices de volume do IBGE, para o setor inteiro; extrapolar para uma PME específica é hipótese do modelo.",
    ]:
        concl.append(Paragraph(txt, bul, bulletText="•"))
    story.append(KeepTogether(concl[:2]))
    story += concl[2:]

    # ================= ETAPA 2 =================
    story.append(PageBreak())
    story.append(Paragraph("ETAPA 2 — Empresas fictícias: saúde financeira e aplicação dos cenários macro", parte))
    story.append(Paragraph("Primeiro vêm os dados das três empresas fictícias (seção 7), um perfil por setor. Em seguida, o crescimento setorial e os cenários da Etapa 1 são aplicados a esses perfis "
        "(seções 8 a 11). Os valores dos perfis (receita, margem, liquidez, balanço) são <b>fictícios</b>, estruturados com base no mercado atual para demonstrar como o modelo responde ao cenário macro.", body))

    # ---------------- 7. Perfis ----------------
    story.append(Paragraph("7. Perfis financeiros e validação contábil", h1))
    story.append(Paragraph(
        "Os perfis (planilhas de entrada) são fictícios. Antes de usá-los, o script confere as identidades contábeis (lucro = receita − custos − juros; "
        "ativo = soma dos grupos; passivo + PL = ativo; liquidez corrente = AC/PC) e <b>corrige</b> o que não fecha, registrando a correção.", body))
    linhas = [[H("Indicador")] + [H(s) for s in nomes_setores]]
    linhas += [[PL("Receita líquida (R$ mil)")] + [P(mil(PERFIS[s]["R0"])) for s in nomes_setores],
               [PL("Lucro líquido (R$ mil) / margem")] + [P(f"{mil(PERFIS[s]['L0'])} / {pct(PERFIS[s]['margem'] * 100, 1)}") for s in nomes_setores],
               [PL("Liquidez corrente")] + [P(num(PERFIS[s]["LC"], 2)) for s in nomes_setores],
               [PL("Custos fixos (% dos custos)")] + [P(pct(PERFIS[s]["fixo"] * 100, 0)) for s in nomes_setores],
               [PL("PMR / PME / PMP (dias, pelo balanço)")] + [P(f"{PERFIS[s]['cr'] / PERFIS[s]['R0'] * 365:.0f} / {PERFIS[s]['est'] / PERFIS[s]['C0'] * 365:.0f} / {PERFIS[s]['forn'] / PERFIS[s]['C0'] * 365:.0f}") for s in nomes_setores],
               [PL("Dívida total (% da receita)")] + [P(pct((PERFIS[s]["dcp"] + PERFIS[s]["dlp"]) / PERFIS[s]["R0"] * 100, 0)) for s in nomes_setores],
               [PL("Parcela da dívida indexada à Selic/CDI")] + [P(pct(PERFIS[s]["indexada"] * 100, 0)) for s in nomes_setores],
               [PL("Caixa (R$ mil)")] + [P(mil(PERFIS[s]["caixa"])) for s in nomes_setores]]
    story.append(tabela(linhas, [190, 108, 108, 108]))
    corr_ = E2["validacao"][E2["validacao"]["acao"] != "ok"]
    if len(corr_):
        txt = "; ".join(f"<b>{r_['setor']}</b> — {r_['item'].lower()}: informado R$ {mil(r_['informado'])} mil, recalculado R$ {mil(r_['recalculado'])} mil" for _, r_ in corr_.iterrows())
        story.append(Paragraph(f"<b>Correções aplicadas:</b> {txt}. Recomenda-se corrigir os arquivos de origem; o relatório já usa os valores consistentes.", note))
    else:
        story.append(Paragraph("Todos os perfis fecham as identidades contábeis (tolerância de 1%).", note))
    # ---------------- Aplicação dos cenários às empresas ----------------
    story.append(Paragraph("Aplicação dos cenários macro às empresas", h1))
    story.append(Paragraph("A partir daqui, o crescimento setorial e os cenários macro da Etapa 1 são aplicados aos perfis acima para medir o efeito sobre lucro, fluxo de caixa e liquidez. "
        "Nenhum cálculo da Etapa 1 depende dos perfis.", body))
    _ini_ref = data_ref - pd.DateOffset(years=1)
    story.append(Paragraph("Data-base da aplicação", h1))
    dados_ref = [[H("Item"), H("Data / período")],
        [PL("Data de execução do relatório"), PL(f"{datetime.datetime.now():%d/%m/%Y}")],
        [PL("Data-base do cenário atual (Selic e câmbio)"), PL(f"{data_ref:%d/%m/%Y}")],
        [PL("Período de referência do cenário atual"), PL(f"12 meses até {data_ref:%d/%m/%Y} (de {_ini_ref:%d/%m/%Y} a {data_ref:%d/%m/%Y})")],
        [PL("IPCA acumulado em 12 meses"), PL(f"até {ultimo_mes_ipca:%m/%Y}")],
        [PL("Amostra histórica de estimação do modelo"), PL(f"{MES_INI_D:%m/%Y} a {MES_FIM_D:%m/%Y} (último mês com dado setorial)")],
        [PL("Ano-base dos perfis fictícios"), PL("hipotético, sem data real; tratado como o ano imediatamente anterior ao período de referência")]]
    story.append(tabela(dados_ref, [190, 326]))
    story.append(Paragraph(
        f"<b>Como ler a data.</b> O modelo relaciona o crescimento setorial às condições macro <i>do mesmo período</i>; portanto o cenário atual descreve o que as condições dos 12 meses até {data_ref:%d/%m/%Y} "
        "implicam para o crescimento do setor, e não uma previsão do ano seguinte. Os estresses moderado e severo somam choques a essas condições. "
        f"A simulação de Monte Carlo (seção 10) é incondicional: sorteia cenários em torno da média histórica {ANO_INI}–{ANO_FIM} e não depende da data-base, exceto pela Selic atual usada como referência do custo da dívida indexada. "
        "Os resultados mudam se o relatório for executado de novo com dados mais recentes.", note))
    story.append(Paragraph(
        "<b>Mecânica da projeção.</b> Receita e custos variáveis crescem com (1 + crescimento de volume previsto) × (1 + IPCA); custos fixos, com o IPCA; juros = juros do ano-base + dívida × "
        "parcela indexada × (Selic do cenário − Selic atual); prazos médios constantes; Capex igual à depreciação, sem amortização; retiradas dos sócios no nível do lucro do ano-base. "
        "Se o caixa não cobre o fluxo negativo, a diferença vira dívida de curto prazo (necessidade de financiamento).", body))

    # ---------------- 8. Projeções ----------------
    story.append(Paragraph("8. Projeção por setor (R$ mil)", h1))
    itens = [("Receita", "R", "v"), ("Lucro líquido", "L", "v"), ("Margem líquida", "margem", "p"),
             ("Fluxo de caixa operacional (FCO)", "fco", "v"), ("Retiradas dos sócios", "dist", "v"),
             ("Var. da necessidade de capital de giro", "dncg", "v"), ("Necessidade de financiamento", "gap", "v"),
             ("Caixa", "caixa", "v"), ("Contas a receber", "cr", "v"), ("Estoques", "est", "v"), ("Ativo não circulante", "anc", "v"),
             ("<b>Total do ativo</b>", "ativo", "v"), ("Fornecedores e outras obrigações", "forn_out", "v"),
             ("Dívida de curto prazo", "dcp", "v"), ("Dívida de longo prazo", "dlp", "v"), ("Patrimônio líquido", "pl", "v"),
             ("<b>Total do passivo + PL</b>", "passivo", "v"), ("<b>Liquidez corrente</b>", "lc", "r")]

    def cel(d, k, t):
        x = float(d[k])
        if np.isnan(x):
            return "—"
        return mil(x) if t == "v" else pct(x, 1) if t == "p" else num(x, 2)

    def delta(a, b_, k, t):
        if np.isnan(float(a[k])) or np.isnan(float(b_[k])):
            return "—"
        x = float(b_[k]) - float(a[k])
        if t == "v" and abs(x) < 500:
            return "0"
        return sg(x / 1000, 0) if t == "v" else (sg(x, 1) + " pp") if t == "p" else sg(x, 2)

    for nome in nomes_setores:
        linhas = [[H(nome), H("Ano-base"), H("Atual"), H("Estresse moderado"), H("Estresse severo"), H("Severo vs atual")]]
        for rot, k, t in itens:
            linhas.append([PL(rot), P(cel(fin_base[nome], k, t)), P(cel(fin_atual[nome], k, t)), P(cel(fin_mild[nome], k, t)),
                           P(cel(fin_estr[nome], k, t)), P(delta(fin_atual[nome], fin_estr[nome], k, t))])
        story.append(KeepTogether([tabela(linhas, [165, 66, 66, 72, 72, 74], pad=2.5), Spacer(1, 8)]))
    cab_sum = [H("Setor"), H("Δ Receita (R$ mil)"), H("Δ Lucro (R$ mil)"), H("Δ Caixa (R$ mil)"), H("LC atual"), H("LC estresse"), H("Δ LC")]
    linhas_sum = [cab_sum]
    for nome in nomes_setores:
        a, e = fin_atual[nome], fin_estr[nome]
        linhas_sum.append([PL(f"<b>{nome}</b>"), P(sg((e["R"] - a["R"]) / 1000, 0)), P(sg((e["L"] - a["L"]) / 1000, 0)),
                           P(sg((e["caixa"] - a["caixa"]) / 1000, 0)), P(num(a["lc"], 2)), P(num(e["lc"], 2)), P(sg(e["lc"] - a["lc"], 2))])
    story.append(KeepTogether([Paragraph("<b>Resumo de impacto no estresse severo (vs cenário atual)</b>", body),
                               tabela(linhas_sum, [110, 70, 70, 70, 55, 65, 55]),
                               Paragraph("As retiradas dos sócios ficam no nível do lucro do ano-base e não há corte de custos fixos nem de Capex. Em um estresse real as empresas ajustam esses itens, "
                                         "o que atenuaria a queda de caixa: a tabela é, portanto, um cenário conservador de transmissão do choque macro à liquidez.", note)]))

    # ---------------- 9. Sensibilidade ----------------
    story.append(Paragraph("9. Correlação e regressão de sensibilidade", h1))
    story.append(Paragraph(
        "Esta seção responde: <i>quais variáveis macro mais movem o lucro, o FCO e a liquidez de cada perfil?</i> Usa dois instrumentos. (a) <b>Tornado determinístico</b>: choque de ±1 desvio-padrão "
        "histórico em uma variável por vez a partir do cenário atual, passando pelo modelo único e pela projeção contábil. (b) <b>Correlação de Spearman e regressão padronizada (MQO)</b> dos indicadores "
        f"financeiros contra os três choques macro nas {N_SIM:,}".replace(",", ".") + " simulações. Como os dados são simulados (e os perfis, fictícios), <b>não se reportam p-valores</b>: com N grande "
        "tudo seria 'significativo'. O R² macro é a fração da variância do indicador explicada pelas variáveis macro; o restante é risco setorial idiossincrático e incerteza dos parâmetros.", body))
    story.append(img(FIG_E2_TORN, fig_e2_torn_size))
    td = E2["tornado"]
    linhas = [[H("Setor"), H("Variável"), H("Δ macro (pp)"), H("Δ lucro (+1 dp)"), H("Δ lucro (−1 dp)"), H("Δ LC (+1 dp)"), H("Δ LC (−1 dp)")]]
    for s in nomes_setores:
        for k in MACRO_COLS:
            a = td[(td.setor == s) & (td.chave == k) & (td.choque == "+1 dp")].iloc[0]
            b_ = td[(td.setor == s) & (td.chave == k) & (td.choque == "−1 dp")].iloc[0]
            linhas.append([PL(f"<b>{s}</b>") if k == "Selic" else P(""), PL(MACRO_LABEL[k]), P(f"±{num(abs(a['delta_macro_pp']), 1)}" if abs(abs(a['delta_macro_pp']) - abs(b_['delta_macro_pp'])) < 0.05
                              else f"{sg(a['delta_macro_pp'], 1)} / {sg(b_['delta_macro_pp'], 1)}"),
                           P(sg(a["delta_lucro_pct"], 1) + "%"), P(sg(b_["delta_lucro_pct"], 1) + "%"), P(sg(a["delta_lc"], 2)), P(sg(b_["delta_lc"], 2))])
    story.append(KeepTogether([tabela(linhas, [95, 75, 60, 70, 70, 70, 70], pad=2.5),
                               Paragraph("Δ lucro em % do lucro do cenário atual; Δ LC em pontos de liquidez corrente. dp = desvio-padrão histórico da variável na amostra. Quando o choque levaria a variável além do máximo ou do mínimo históricos, ele é limitado a esse valor (Δ macro efetivo, em pp).", note)]))
    sm = E2["sens_mc"]
    linhas = [[H("Setor"), H("Variável"), H("ρ Spearman lucro"), H("Coef. pad. lucro"), H("ρ Spearman LC"), H("Coef. pad. LC")]]
    for s in nomes_setores:
        for k in MACRO_COLS:
            a = sm[(sm.setor == s) & (sm.chave == k) & (sm.indicador == "Lucro líquido")].iloc[0]
            b_ = sm[(sm.setor == s) & (sm.chave == k) & (sm.indicador == "Liquidez corrente")].iloc[0]
            linhas.append([PL(f"<b>{s}</b>") if k == "Selic" else P(""), PL(MACRO_LABEL[k]), P(num(a["spearman"], 2)), P(num(a["beta_padron"], 2)),
                           P(num(b_["spearman"], 2)), P(num(b_["beta_padron"], 2))])
    story.append(KeepTogether([Paragraph("<b>Sensibilidade nas simulações</b>", body), tabela(linhas, [100, 80, 80, 80, 80, 80], pad=2.5)]))
    linhas = [[H("R² macro")] + [H(i_) for i_ in ("Lucro líquido", "FCO", "Liquidez corrente")]]
    for s in nomes_setores:
        linhas.append([PL(f"<b>{s}</b>")] + [P(num(sm[(sm.setor == s) & (sm.indicador == i_)]["r2_macro"].iloc[0], 2)) for i_ in ("Lucro líquido", "FCO", "Liquidez corrente")])
    story.append(Spacer(1, 4))
    story.append(tabela(linhas, [150, 120, 120, 120]))
    story.append(img(FIG_E2_SENS, fig_e2_sens_size))
    top_t = {s: td[(td.setor == s)].assign(a=lambda d_: d_["delta_lucro_pct"].abs()).sort_values("a", ascending=False).iloc[0] for s in nomes_setores}
    story.append(Paragraph(
        "<b>Leitura.</b> Variável com maior efeito sobre o lucro (tornado): "
        + "; ".join(f"{s}: {MACRO_LABEL[top_t[s]['chave']]} ({sg(top_t[s]['delta_lucro_pct'], 1)}% para {top_t[s]['choque']})" for s in nomes_setores)
        + ". O R² macro do lucro é "
        + ", ".join(f"{num(sm[(sm.setor == s) & (sm.indicador == 'Lucro líquido')]['r2_macro'].iloc[0], 2)} ({s})" for s in nomes_setores)
        + ": quanto menor, mais o resultado depende de fatores que o macro não captura.", body))

    # ---------------- 10. Monte Carlo financeiro ----------------
    cab = [H("Setor"), H("Lucro esp. (R$ mil)"), H("FCO esp. (R$ mil)"), H("FCO pior 5% (R$ mil)"), H("Caixa final pior 5% (R$ mil)"), H("LC mediana"),
           H("LC pior 5%"), H("Prob. prejuízo"), H("Prob. LC &lt; 1,0"), H("Prob. financiar")]
    linhas = [cab]
    for nome in nomes_setores:
        m = fin_mc[nome]
        linhas.append([PL(f"<b>{nome}</b>"), P(mil(m["lucro_med"])), P(mil(m["fco_med"])), P(mil(m["fco_p5"])), P(mil(m["caixa_p5"])), P(num(m["lc_med"], 2)),
                       P(num(m["lc_p5"], 2)), P(pct(m["p_prejuizo"])), P(pct(m["p_lc1"])), P(pct(m["p_gap"]))])
    story.append(KeepTogether([
        Paragraph("10. Monte Carlo sobre o fluxo de caixa e o balanço", h1),
        Paragraph(f"Cada um dos {N_SIM:,}".replace(",", ".") + " cenários usa o crescimento setorial, o IPCA e a Selic sorteados juntos, aplicados ao mesmo balanço-base. Como a simulação é "
                  "<b>incondicional</b> (em torno da média histórica), o custo extra da dívida indexada é calculado como (Selic sorteada − Selic atual); como a Selic atual está acima da média "
                  "histórica, o Monte Carlo tende a mostrar alívio financeiro médio nesse quesito, coerente com uma hipótese de normalização monetária. Para o cenário atual e os estresses, "
                  "a referência é a Selic atual.", body),
        tabela(linhas, [78, 44, 44, 48, 52, 38, 38, 42, 44, 46], pad=3)]))
    story.append(img(FIG3, fig3_size))

    # ---------------- 11. Conclusões da Etapa 2 (empresas fictícias) ----------------
    concl = [Paragraph("11. Conclusões da Etapa 2 — empresas fictícias", h1)]
    menor_lc = min(nomes_setores, key=lambda n_: fin_estr[n_]["lc"])
    abaixo1 = [n_ for n_ in nomes_setores if fin_estr[n_]["lc"] < 1]
    pior_p = max(nomes_setores, key=lambda n_: fin_mc[n_]["p_lc1"])
    st_sev = ", ".join(f"{s} {num(fin_atual[s]['lc'], 2)} → {num(fin_estr[s]['lc'], 2)}" for s in nomes_setores)
    concl.append(Paragraph(
        f"<b>Caixa e balanço.</b> Liquidez corrente, cenário atual → estresse severo: {st_sev}. "
        f"No estresse severo, {menor_lc} tem a menor liquidez corrente ({num(fin_estr[menor_lc]['lc'], 2)} contra {num(fin_base[menor_lc]['lc'], 2)} no ano-base); "
        + (f"ficam abaixo de 1,0: {', '.join(abaixo1)}. " if abaixo1 else "nenhum setor fica abaixo de 1,0. ")
        + f"No Monte Carlo, a maior probabilidade de LC &lt; 1,0 é a de {pior_p} ({pct(fin_mc[pior_p]['p_lc1'])}). O canal dominante é a compressão de margem "
          "(custos fixos corrigidos pela inflação e juros da dívida indexada à Selic); o capital de giro varia pouco.", body))
    concl.append(Paragraph(
        f"<b>Data de referência.</b> Estes resultados valem para as condições macro dos 12 meses até {data_ref:%d/%m/%Y} (IPCA até {ultimo_mes_ipca:%m/%Y}), "
        f"com o modelo estimado em {MES_INI_D:%m/%Y}–{MES_FIM_D:%m/%Y}.", body))
    concl.append(Paragraph("<b>Limitações da Etapa 2</b>", body))
    for txt in [
        "O IPCA entra duas vezes na projeção: como regressor do crescimento de volume (coeficiente estimado) e como proxy de preço na receita nominal; o IPCA geral não substitui o deflator setorial.",
        "Os perfis da seção 7 são fictícios e simplificados (sem Capex líquido, sem amortização, retiradas fixas no lucro do ano-base, tributos embutidos nos custos); os resultados mostram como o modelo responde ao cenário macro, não o desempenho de uma empresa real.",
        "A tabela de impacto no estresse é conservadora: não há corte de custos fixos, de Capex ou de retiradas dos sócios, ajustes que empresas reais fariam.",
    ]:
        concl.append(Paragraph(txt, bul, bulletText="•"))
    story.append(KeepTogether(concl[:2]))
    story += concl[2:]

    # ---------------- Apêndice: auditoria de integridade ----------------
    _tit_aud = Paragraph("Apêndice — Auditoria de integridade da execução", h1)
    _txt_aud = (Paragraph("Verificações automáticas feitas pelo script sobre esta execução (também salvas em auditoria_integridade.csv; "
                           "os hashes dos arquivos de entrada estão em execucao_metadata.json). \"ATENÇÃO\" não invalida o resultado: "
                           "indica um ponto que precisa constar na discussão.", note))
    linhas = [[H("Verificação"), H("Resultado"), H("Status")]]
    linhas += [[PL(r_.verificacao), PL(r_.resultado), P(f"<b>{r_.status}</b>")] for r_ in AUDITORIA.itertuples()]
    story.append(KeepTogether([_tit_aud, _txt_aud, tabela(linhas, [190, 250, 76])]))
    doc.build(story, canvasmaker=NumberedCanvas)


# ==============================================================================
# 7. AUDITORIA DE INTEGRIDADE, METADADOS E EMPACOTAMENTO
# ==============================================================================
def auditar_execucao() -> pd.DataFrame:
    """Confere a coerência interna da execução e devolve uma tabela (verificação, resultado, status).

    Não altera nenhum resultado: serve de rastro para a banca (também sai no PDF e em auditoria_integridade.csv).
    """
    linhas: list[dict[str, str]] = []

    def reg(verificacao: str, resultado: str, passou: bool) -> None:
        linhas.append(dict(verificacao=verificacao, resultado=resultado, status="OK" if passou else "ATENÇÃO"))

    reg("Amostra mensal contígua e sem valores ausentes",
        f"n = {n_obs} meses ({MES_INI_D:%m/%Y}–{MES_FIM_D:%m/%Y})",
        bool(df.notna().all().all()) and len(_cheio) == n_obs)
    reg("Colinearidade entre as variáveis macro", f"VIF máximo = {num(float(vifs.max()), 2)} (limite 5)",
        bool(vifs.max() < 5))
    reg("Sobreposição PIB–setor evitada", "IBC-Br só como contexto descritivo; fora da regressão, dos cenários e do Monte Carlo",
        "PIB" not in MACRO_COLS and macro_sim.shape[1] == len(MACRO_COLS))
    atraso_ibc = (data_ref - ultimo_mes_ibc).days
    reg("Atualidade dos dados",
        f"Selic/câmbio até {data_ref:%d/%m/%Y}; IPCA até {ultimo_mes_ipca:%m/%Y}; IBC-Br até {ultimo_mes_ibc:%m/%Y}",
        atraso_ibc < 130)
    dif_media = float(np.max(np.abs(macro_sim.mean(0) - mu) / sd_macro))
    razao_dp = macro_sim.std(0, ddof=1) / sd_macro
    if LIMITAR_FAIXA:
        _dentro = bool(np.all((macro_sim >= np.array([FAIXA_HIST[k][0] for k in MACRO_COLS])) & (macro_sim <= np.array([FAIXA_HIST[k][1] for k in MACRO_COLS]))))
        reg("Monte Carlo dentro da faixa histórica (média e volatilidade após o truncamento)",
            f"todos os sorteios em [mín, máx]; {N_REDESENHADOS} re-sorteios; desvio máx. da média = {num(dif_media, 2)} dp; razão de volatilidade {num(float(razao_dp.min()), 2)}–{num(float(razao_dp.max()), 2)}",
            _dentro and dif_media < 0.50 and bool(np.all((razao_dp > 0.50) & (razao_dp < 1.15))))
    else:
        reg("Monte Carlo reproduz média e volatilidade históricas",
            f"desvio máx. da média = {num(dif_media, 2)} dp; razão de volatilidade {num(float(razao_dp.min()), 2)}–{num(float(razao_dp.max()), 2)}",
            dif_media < 0.10 and bool(np.all((razao_dp > 0.85) & (razao_dp < 1.15))))
    probs = [metr[s]["p_contracao"] for s in nomes_setores] + [fin_mc[s][k] for s in nomes_setores
                                                              for k in ("p_prejuizo", "p_lc1", "p_gap")]
    reg("Probabilidades simuladas dentro de [0, 100]%", f"mín. {num(min(probs), 1)}% / máx. {num(max(probs), 1)}%",
        all(0 <= p <= 100 for p in probs))
    n_corr = int((E2["validacao"]["acao"] != "ok").sum())
    reg("Perfis contábeis fecham as identidades (tolerância de 1%)",
        "sem correções" if n_corr == 0 else f"{n_corr} correção(ões) automática(s) registrada(s) na seção 7", n_corr == 0)
    reg("Balanço projetado fecha em todos os cenários", "ativo = passivo + PL (verificado em cada projeção)", True)
    fora = sorted({MACRO_LABEL[k] for cen in (cenario_estresse_mild, cenario_estresse) for k in MACRO_COLS
                   if not (df[k].min() <= cen[k] <= df[k].max())})
    _lim_aud = sorted(set(LIMITADAS_MILD) | set(LIMITADAS_SEV))
    reg("Cenários de estresse dentro da faixa histórica", ("todas as variáveis dentro da faixa" + (f" (limitadas: {', '.join(_lim_aud)})" if _lim_aud else ""))
        if not fora else "fora da faixa: " + ", ".join(fora) + " (extrapolação)", not fora)
    ruins = [s for s in nomes_setores if modelos[s]["r2_wf"] < 0]
    reg("Poder preditivo no tempo (walk-forward) acima da média histórica",
        "todos os setores" if not ruins else "abaixo da média histórica em: " + ", ".join(ruins), not ruins)
    reg("Reprodutibilidade", f"semente = {SEED}; {N_SIM:,} simulações".replace(",", "."), True)
    return pd.DataFrame(linhas)


def salvar_metadados() -> Path:
    """Grava execucao_metadata.json: versões, semente, amostra e hash (SHA-256) de cada arquivo de entrada."""
    arquivos = {}
    for chave, caminho in ENTRADAS.items():
        arquivos[chave] = dict(nome=caminho.name, bytes=caminho.stat().st_size,
                               sha256=hashlib.sha256(caminho.read_bytes()).hexdigest()[:16])
    meta = dict(
        gerado_em=datetime.datetime.now().isoformat(timespec="seconds"),
        ambiente=dict(python=platform.python_version(), numpy=np.__version__, pandas=pd.__version__,
                      scipy=scipy.__version__, matplotlib=matplotlib.__version__, colab=EM_COLAB),
        parametros=dict(semente=SEED, n_sim=N_SIM, nu_t=NU, hac_lags=HAC_LAGS, embargo_wf=EMBARGO_WF,
                        lag_selic_robustez=LAG_SELIC_ROB, ini_amostra=INI_AMOSTRA),
        amostra=dict(inicio=f"{MES_INI_D:%Y-%m}", fim=f"{MES_FIM_D:%Y-%m}", n=int(n_obs)),
        arquivos_de_entrada=arquivos,
    )
    destino = Path(OUT_DIR) / "execucao_metadata.json"
    destino.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return destino


AUDITORIA = auditar_execucao()
AUDITORIA.to_csv(os.path.join(OUT_DIR, "auditoria_integridade.csv"), index=False)
print("\n[AUDITORIA]")
print(AUDITORIA.to_string(index=False))

print("\n[PDF] Gerando relatório...")
gerar_pdf()
salvar_metadados()
print(f"[SUCESSO] Relatório gerado em: {os.path.abspath(PDF_PATH)}")

if EMPACOTAR_SAIDAS:
    ZIP_PATH = shutil.make_archive(str(Path(OUT_DIR).resolve().parent / "resultados_pmes"), "zip", OUT_DIR)
    print(f"[SUCESSO] Pacote com PDF, tabelas e figuras: {ZIP_PATH}")
if BAIXAR_NO_COLAB and EM_COLAB:
    from google.colab import files  # type: ignore[import-not-found]
    files.download(PDF_PATH)
    if EMPACOTAR_SAIDAS:
        files.download(ZIP_PATH)
