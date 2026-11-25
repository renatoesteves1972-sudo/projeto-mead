# versão 7.42 — 25/11/2026
# v7.44: cronômetro reiniciado por tema e atualizado na linha de status até o fim do processamento.
# v7.41: corrige seleção multilíngue, subtítulo independente, B1 institucional controlado e DOCX do bloco 2; remove CTA fixa.
# v7.40: refinamento editorial — preserva/extrai títulos dinâmicos, prosa corrida sem listas/fichas,
# faixa 40–55 palavras como referência editorial, títulos preservados no JSON, tipos de bomba separados e B1 comercial
# explicitamente relacionado à keyword, variado e sem CTA genérico não sustentado.
# Base v7.39: parágrafos objetivos/comercial B1, arquitetura de 15 evidências e reservas.
# versão 7.33 — 24/11/2026
# v7.34: teste editorial MEAD — progressão temática, diversidade entre parágrafos e auditoria de repetição; mantém Qwen3.5:4b e arquitetura de 1 chamada/15 parágrafos.
# v7.32: preserva resposta global/parcial nas atualizações do progresso; recupera 15 parágrafos sem marcadores por fronteiras de frases, quando há conteúdo suficiente; mantém uma chamada Ollama.
# v7.27: reforça a instrução global de keyword única; normaliza capitalização e corrige repetições posteriores com retomada contextual; mantém falha segura quando a expressão está ausente; preserva uma chamada Ollama e métricas de desempenho.
# v7.23: corrige contador, gravação independente do mapa MEAD, log compacto e carregamento MEAD único.
# v7.22: recuperação COM isolada por trecho; gravação estrutural resiliente; log compacto.
# v7.18: Ollama pode usar conhecimento geral do tema; seleção prioriza COM e completa com SEM; log resumido.
# v7.17: recupera estruturas originais v7.15; corrige leitura do MEAD flat; mantém patrimônio global/SEVEN e keyword exata; preserva 30 tags e 12 segmentos.
# v7.15: correção do gate comercial condicionado à evidência no bloco 1; retirada dos checkboxes do prompt Ollama; recuperação COM e auditoria final preservadas; diagnósticos somente no log.
# Base integral: v7.14 recuperada da Biblioteca. Alterações pontuais, sem banco global novo.
# v7.11: hash COM/SEM palavra-chave, seleção exclusiva COM, auditoria de fragmentos e alocação por bloco corrigida.
# v7.10: restaura a biblioteca atual de segmentos e a priorização temática da linhagem V100.
# v7.9: 5 chamadas Ollama (uma por bloco), 3 fragmentos por chamada, narrativa encadeada e triagem de cabeçalho corrigida.
# v7.8: triagem de evidências, diagnóstico separado, segmentos relevantes e DOCX direto do JSON no padrão 6.60.
# v7.7: arquitetura anterior com 15 chamadas Ollama individuais + gate editorial mínimo 7,5 por bloco.
# v7.5: validação factual pós-Ollama por evidência lexical/parafrástica,
# preservando controle rígido de números e evitando rejeição de paráfrases.
# v7.1: extração temática reforçada para fontes técnicas/PDFs e redução de falso descarte por nome de fabricante.
#
# Correções desta versão:
# - seleção COM prioritária e SEM complementar; geração não bloqueada por falta de hashes;
# - filtro editorial menos agressivo em linguagem técnica;
# - diagnóstico explícito de fontes/candidatos insuficientes;
# - nenhuma chamada ao Ollama com bloco incompleto.

import json
import os
import re
import hashlib
import time
import unicodedata
import tempfile
import threading
import requests
import fitz
import tkinter as tk
import random
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup
from ddgs import DDGS



ARQUIVO_BANCO = r"C:\Python\gerador-conteudo\banco\conteudo.json"

ARQUIVO_MEAD = r"C:\Python\gerador-conteudo\mead\mead.json"

ARQUIVO_PROGRESSO = r"C:\Python\gerador-conteudo\banco\progresso.json"

# Modelo único utilizado em todas as etapas de IA do gerador.
MODELO_OLLAMA = "qwen3.5:4b"


# ============================================================
# RECUPERAÇÃO DE CONTEÚDO — DENTRO DO progresso.json
# ============================================================
def _ler_progresso_seguro():
    try:
        if not os.path.exists(ARQUIVO_PROGRESSO):
            return {}
        with open(ARQUIVO_PROGRESSO, "r", encoding="utf-8") as f:
            dados = json.load(f)
        return dados if isinstance(dados, dict) else {}
    except Exception as erro:
        print("⚠️ NÃO FOI POSSÍVEL LER O PROGRESSO:", repr(erro))
        return {}


def _gravar_progresso_seguro(dados):
    diretorio = os.path.dirname(ARQUIVO_PROGRESSO)
    if diretorio:
        os.makedirs(diretorio, exist_ok=True)
    temporario = ARQUIVO_PROGRESSO + ".tmp"
    with open(temporario, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)
        f.flush()
        os.fsync(f.fileno())
    os.replace(temporario, ARQUIVO_PROGRESSO)


def _tema_recuperacao_id(tema):
    try:
        return normalizar_tema_chave(str(tema or ""))
    except Exception:
        return re.sub(r"\s+", " ", str(tema or "").strip().casefold())


def _extrair_paragrafos_resposta_global(texto, limite=15):
    """Recupera parágrafos de respostas Ollama com marcadores ausentes/inconsistentes.

    Não escreve nem completa conteúdo: apenas identifica segmentos existentes.
    A validação editorial e a keyword continuam nas etapas posteriores.
    """
    bruto = str(texto or "").replace("\r", "\n").strip()
    if not bruto:
        return []

    def limpar(valor):
        valor = re.sub(r"```(?:text|txt|markdown)?|```", " ", str(valor), flags=re.I)
        valor = re.sub(r"\[(?:/?BLOCO(?:_\d+)?|/?TITULO_BLOCO|/?SUBTITULO_PAGINA|/?PARAGRAFO_\d+)\]", "\n", valor, flags=re.I)
        valor = re.sub(r"^\s*(?:#{1,6}\s*|(?:par[aá]grafo|bloco)\s*\d+\s*[:.)-]\s*)", "", valor, flags=re.I)
        valor = re.sub(r"\s+", " ", valor).strip(" \t\n\r-*•")
        return valor

    # 1) Marcadores de parágrafo mesmo quando o marcador de fechamento foi omitido.
    aberturas = list(re.finditer(r"\[PARAGRAFO_\d+\]", bruto, flags=re.I))
    if aberturas:
        achados = []
        for i, m in enumerate(aberturas):
            fim = aberturas[i + 1].start() if i + 1 < len(aberturas) else len(bruto)
            trecho = bruto[m.end():fim]
            trecho = re.split(r"\[/PARAGRAFO_\d+\]", trecho, maxsplit=1, flags=re.I)[0]
            p = limpar(trecho)
            if len(p.split()) >= 8:
                achados.append(p)
        if achados:
            return achados[:limite]

    # 2) Itens numerados que representam parágrafos completos (1., 1), 1-).
    linhas = [x.strip() for x in bruto.split("\n") if x.strip()]
    numerados = []
    atual = []
    iniciou = False
    for linha in linhas:
        m = re.match(r"^\s*(?:\*\*|__)?(?:par[aá]grafo\s*)?(\d{1,2})\s*[.)-]\s*(.*)$", linha, flags=re.I)
        if m and 1 <= int(m.group(1)) <= 30 and m.group(2).strip():
            if atual:
                p = limpar(" ".join(atual))
                if len(p.split()) >= 12:
                    numerados.append(p)
            atual = [m.group(2).strip()]
            iniciou = True
        elif iniciou:
            # Cabeçalhos e rótulos são descartados; continuação fica no parágrafo atual.
            if re.match(r"^\s*(?:#{1,6}|\[/?(?:BLOCO|TITULO|SUBTITULO))", linha, flags=re.I):
                continue
            atual.append(linha)
    if atual:
        p = limpar(" ".join(atual))
        if len(p.split()) >= 12:
            numerados.append(p)
    if len(numerados) >= 3:
        return numerados[:limite]

    # 3) Separação por linhas em branco (formato convencional).
    blocos = [limpar(x) for x in re.split(r"\n\s*\n+", bruto) if limpar(x)]
    blocos = [x for x in blocos if len(x.split()) >= 12]
    if len(blocos) >= 3:
        return blocos[:limite]

    # 4) Resposta sem marcadores e sem quebras de linha: alguns modelos
    #    concatenam os parágrafos em um único fluxo. Recupera a estrutura
    #    somente nas fronteiras de frases existentes (não reescreve conteúdo).
    texto_sem_cabecalhos = re.sub(
        r"\[(TITULO_BLOCO|SUBTITULO_PAGINA)\].*?\[/\1\]", " ", bruto,
        flags=re.I | re.S
    )
    texto_sem_cabecalhos = re.sub(
        r"\[(?:/?BLOCO(?:_\d+)?|/?TITULO_BLOCO|/?SUBTITULO_PAGINA|/?PARAGRAFO_\d+)\]",
        " ", texto_sem_cabecalhos, flags=re.I
    )
    texto_sem_cabecalhos = re.sub(r"\s+", " ", texto_sem_cabecalhos).strip()
    frases = [x.strip() for x in re.split(
        r"(?<=[.!?…])\s+(?=[A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9])", texto_sem_cabecalhos
    ) if x.strip()]
    palavras_total = len(texto_sem_cabecalhos.split())
    if len(frases) >= limite and palavras_total >= limite * 25:
        # Distribui frases sequencialmente em 15 grupos equilibrados.
        # Não inventa frases, não corta sentenças e não altera a redação.
        grupos_frases = []
        pos_frase = 0
        for grupo_idx in range(limite):
            grupos_restantes = limite - grupo_idx
            frases_restantes = len(frases) - pos_frase
            palavras_restantes = sum(len(x.split()) for x in frases[pos_frase:])
            alvo = max(25, round(palavras_restantes / grupos_restantes))
            grupo = []
            palavras_grupo = 0
            while pos_frase < len(frases):
                frase = frases[pos_frase]
                qtd = len(frase.split())
                if grupo and palavras_grupo >= alvo and (palavras_grupo + qtd) > alvo:
                    break
                # Reserva ao menos uma frase para cada grupo que falta.
                if grupo and (len(frases) - pos_frase) <= (grupos_restantes - 1):
                    break
                grupo.append(frase)
                palavras_grupo += qtd
                pos_frase += 1
                frases_restantes -= 1
                if palavras_grupo >= alvo:
                    break
            if grupo:
                grupos_frases.append(limpar(" ".join(grupo)))
        if pos_frase == len(frases) and len(grupos_frases) == limite and all(len(x.split()) >= 20 for x in grupos_frases):
            print("⚠️ RESPOSTA SEM MARCADORES: 15 parágrafos recuperados por fronteiras de frases, sem reescrita.")
            return grupos_frases

    # 5) Alguns modelos devolvem um parágrafo por linha simples. Agrupa linhas
    #    até formar unidades com pelo menos 30 palavras, sem cortar frases.
    linhas_limpas = [limpar(x) for x in linhas]
    linhas_limpas = [x for x in linhas_limpas if x and not re.match(
        r"^(?:t[ií]tulo|subt[ií]tulo|bloco\s*\d+|introdu[cç][aã]o|conclus[aã]o)\b",
        x, flags=re.I)]
    agrupados, acumulado = [], []
    for linha in linhas_limpas:
        # linhas muito curtas normalmente são cabeçalhos/itens de lista, não parágrafos.
        if len(linha.split()) < 5 and not acumulado:
            continue
        acumulado.append(linha)
        palavras = len(" ".join(acumulado).split())
        if palavras >= 30 and re.search(r"[.!?…][\"'”’)]?$", linha):
            agrupados.append(limpar(" ".join(acumulado)))
            acumulado = []
    if acumulado:
        resto = limpar(" ".join(acumulado))
        if len(resto.split()) >= 30:
            agrupados.append(resto)
    return agrupados[:limite]



def _extrair_paragrafos_por_bloco_resposta_global(texto):
    """Recupera parágrafos sem tags individuais, preservando os limites dos 5 blocos.

    A recuperação só é aceita quando cada bloco possui exatamente três parágrafos
    separados por linhas em branco. Não redistribui nem inventa texto para completar
    blocos, evitando deslocar parágrafos de um bloco para outro.
    """
    bruto = str(texto or "").replace("\r\n", "\n").replace("\r", "\n")
    resultado = []

    for numero in range(1, 6):
        abertura = re.search(rf"\[BLOCO_{numero}\]", bruto, flags=re.I)
        if not abertura:
            return []

        proximo = re.search(rf"\[BLOCO_{numero + 1}\]", bruto[abertura.end():], flags=re.I) if numero < 5 else None
        fechamento = re.search(rf"\[/BLOCO_{numero}\]", bruto[abertura.end():], flags=re.I)

        limites = []
        if proximo:
            limites.append(abertura.end() + proximo.start())
        if fechamento:
            limites.append(abertura.end() + fechamento.start())
        fim = min(limites) if limites else len(bruto)
        trecho = bruto[abertura.end():fim]

        # Remove títulos e subtítulos completos, não apenas seus rótulos.
        trecho = re.sub(
            r"\[(TITULO_BLOCO|SUBTITULO_PAGINA)\].*?\[/\1\]",
            "\n", trecho, flags=re.I | re.S
        )
        trecho = re.sub(
            r"\[/?(?:BLOCO(?:_\d+)?|TITULO_BLOCO|SUBTITULO_PAGINA|PARAGRAFO_\d+)\]",
            "\n", trecho, flags=re.I
        )
        trecho = re.sub(r"```(?:text|txt|markdown)?|```", "", trecho, flags=re.I)

        paragrafos = []
        for parte in re.split(r"\n\s*\n+", trecho):
            parte = re.sub(r"\s+", " ", parte).strip(" \t\n\r-*•")
            if len(parte.split()) >= 12:
                paragrafos.append(parte)

        if len(paragrafos) != 3:
            print(
                f"⚠️ RECUPERAÇÃO ESTRUTURAL RECUSADA — BLOCO {numero}: "
                f"{len(paragrafos)} parágrafos identificados; eram necessários 3."
            )
            return []

        resultado.extend(paragrafos)

    if len(resultado) == 15:
        print("⚠️ 15 PARÁGRAFOS RECUPERADOS POR BLOCO — texto preservado, sem redistribuição.")
        return resultado
    return []


def salvar_bloco_recuperacao(tema, numero_bloco, dados_bloco):
    """Persiste imediatamente o bloco já validado, sem criar arquivo paralelo."""
    if not isinstance(dados_bloco, dict):
        return False
    paragrafos = dados_bloco.get("paragrafos_ollama", [])
    fragmentos = dados_bloco.get("informacoes_relevantes", [])
    if (not isinstance(paragrafos, list) or len(paragrafos) != 3
            or any(not str(x or "").strip() for x in paragrafos)
            or not isinstance(fragmentos, list) or len(fragmentos) != 3):
        return False
    progresso = _ler_progresso_seguro()
    recuperacao = progresso.get("recuperacao_conteudo", {})
    if not isinstance(recuperacao, dict):
        recuperacao = {}
    tema_id = _tema_recuperacao_id(tema)
    registro = recuperacao.get(tema_id, {})
    if not isinstance(registro, dict):
        registro = {}
    blocos = registro.get("blocos", {})
    if not isinstance(blocos, dict):
        blocos = {}
    # JSON round-trip evita referências mutáveis e garante serialização.
    blocos[f"bloco_{int(numero_bloco)}"] = json.loads(json.dumps(dados_bloco, ensure_ascii=False, default=str))
    registro.update({"tema": str(tema), "blocos": blocos, "atualizado_em": time.strftime("%d/%m/%Y %H:%M:%S")})
    recuperacao[tema_id] = registro
    progresso["recuperacao_conteudo"] = recuperacao
    try:
        _gravar_progresso_seguro(progresso)
        print(f"💾 RECUPERAÇÃO SALVA NO progresso.json — {tema} / bloco_{numero_bloco}")
        return True
    except Exception as erro:
        print("⚠️ FALHA AO PRESERVAR BLOCO NO progresso.json:", repr(erro))
        return False


def carregar_blocos_recuperacao(tema):
    progresso = _ler_progresso_seguro()
    recuperacao = progresso.get("recuperacao_conteudo", {})
    if not isinstance(recuperacao, dict):
        return {}
    registro = recuperacao.get(_tema_recuperacao_id(tema), {})
    if not isinstance(registro, dict):
        return {}
    blocos = registro.get("blocos", {})
    if not isinstance(blocos, dict):
        return {}
    validos = {}
    for numero in range(1, 6):
        chave = f"bloco_{numero}"
        bloco = blocos.get(chave)
        if not isinstance(bloco, dict):
            continue
        paras = bloco.get("paragrafos_ollama", [])
        frags = bloco.get("informacoes_relevantes", [])
        if (isinstance(paras, list) and len(paras) == 3
                and all(str(x or "").strip() for x in paras)
                and isinstance(frags, list) and len(frags) == 3):
            validos[chave] = bloco
    return validos


def limpar_blocos_recuperacao(tema):
    progresso = _ler_progresso_seguro()
    recuperacao = progresso.get("recuperacao_conteudo", {})
    if not isinstance(recuperacao, dict):
        return
    recuperacao.pop(_tema_recuperacao_id(tema), None)
    if recuperacao:
        progresso["recuperacao_conteudo"] = recuperacao
    else:
        progresso.pop("recuperacao_conteudo", None)
    try:
        _gravar_progresso_seguro(progresso)
        print("✅ RECUPERAÇÃO REMOVIDA APÓS CONFIRMAÇÃO DA GRAVAÇÃO OFICIAL:", tema)
    except Exception as erro:
        print("⚠️ NÃO FOI POSSÍVEL LIMPAR RECUPERAÇÃO:", repr(erro))


# ============================================================
# DIAGNÓSTICOS — SOMENTE NO LOG DA EXECUÇÃO
# ============================================================
def _salvar_diagnostico_ultima_execucao(tema, dados, status, motivo=""):
    """Emite diagnóstico curto no log; nunca cria JSON auxiliar."""
    try:
        dados = dados if isinstance(dados, dict) else {}
        resumo = {
            "bloco": dados.get("bloco"),
            "paragrafo": dados.get("paragrafo"),
            "ocorrencias": dados.get("ocorrencias_iniciais"),
            "nota_minima": dados.get("nota_minima"),
            "media": dados.get("media"),
            "blocos": len(dados.get("informacoes_blocos", {}))
                if isinstance(dados.get("informacoes_blocos"), dict) else None,
        }
        resumo = {k: v for k, v in resumo.items() if v is not None}
        registro = {"tema": str(tema or ""), "status": str(status or ""),
                    "motivo": str(motivo or "")[:240], "dados": resumo}
        print("DIAGNÓSTICO:", json.dumps(registro, ensure_ascii=False, separators=(",", ":"), default=str))
    except Exception as erro:
        print("⚠️ FALHA AO RESUMIR DIAGNÓSTICO:", type(erro).__name__)
    return None


janela = None

IA_PROCESSANDO = False
PROCESSAMENTO_ATIVO = False
etapa_atual = 0
inicio_geracao = 0.0
tempos_etapas = []
etapas_total = 4
PAGINAS_EM_PROCESSAMENTO = set()

from pathlib import Path

# ============================================================
# TAGS FIXAS — ALTERAÇÃO DE 10/09/2026
# ============================================================
# A base padronizada pertence ao Python.
# O Ollama NÃO cria as tags finais.
# ============================================================

TAGS_FIXAS_PRODUTOS = [
    "empresa de [tema]",
    "venda de [tema]",
    "vendemos [tema]",
    "fornecemos [tema]",
    "fornecedor de [tema]",
    "onde encontrar [tema]",
    "onde comprar [tema]",
    "orçamento de [tema]",
    "cotação de [tema]",
    "[tema] com melhor preço",
    "[tema] com preço competitivo",
    "[tema] com preço justo",
    "[tema] com qualidade",
    "[tema] resistente",
    "[tema] durável",
    "fabricante de [tema]",
    "distribuidor de [tema]",
    "[tema] industrial",
    "[tema] para aplicações industriais",
    "[tema] para empresas",
    "[tema] para projetos industriais",
    "[tema] sob medida",
    "[tema] de alto desempenho",
    "[tema] de alta eficiência",
    "[tema] com suporte técnico",
    "especificações técnicas de [tema]",
    "modelos de [tema]",
    "tipos de [tema]",
    "instalação de [tema]",
    "manutenção de [tema]",
]

TAGS_FIXAS_SERVICOS = [
    "empresa especializada em [tema]", "especialista em [tema]", "onde encontrar [tema]",
    "onde contratar [tema]", "orçamento para [tema]", "cotação para [tema]", "escopo de [tema]",
    "etapas de [tema]", "execução de [tema]", "instalação de [tema]", "inspeção de [tema]",
    "manutenção de [tema]", "segurança em [tema]", "critérios técnicos de [tema]",
    "suporte técnico em [tema]",
]


PASTA_DADOS = Path(
    r"C:\Python\gerador-conteudo\dados-brutos"
)

CATEGORIAS = [
    "definicao",
    "beneficios",
    "vantagens",
    "materia_prima",
    "aplicacoes",
    "fabricacao",
    "manutencao",
    "ativos_narrativos",
    "duvidas_frequentes"
]


# ========================================================
# SUBTÍTULOS FIXOS DOS SEGMENTOS
# ========================================================

SUBTITULOS_SEGMENTOS = [
    "Principais Aplicações de [TEMA]",
    "Principais Segmentos Atendidos por [TEMA]",
    "Aplicações de [TEMA] por Segmento",
    "Segmentos de Aplicação de [TEMA]",
    "Aplicações e Segmentos de [TEMA]",
    "Principais Setores que Utilizam [TEMA]",
    "Setores Atendidos por [TEMA]",
    "Aplicações Industriais de [TEMA]",
    "Segmentos que Utilizam [TEMA]",
    "Principais Áreas de Aplicação de [TEMA]",
]

# Controle da variação entre páginas
SUBTITULOS_SEGMENTOS_DISPONIVEIS = []



# ============================================================
# BIBLIOTECA DE SEGMENTOS
# ============================================================

SEGMENTOS_PRODUTOS = [
    # Biblioteca geral de aplicações e mercados.
    "Fábrica de [TEMA] para Saneamento",
    "Fábrica de [TEMA] para Infraestrutura",
    "[TEMA] para Indústrias",
    "[TEMA] para Obras de Drenagem",
    "[TEMA] para Infraestrutura Urbana",
    "[TEMA] para Vias Públicas",
    "[TEMA] para Áreas de Circulação",
    "[TEMA] para Condomínios",
    "[TEMA] para Empresas de Engenharia",
    "[TEMA] para Construtoras",
    "[TEMA] para Obras de Infraestrutura",
    "Fornecedor de [TEMA] para Indústrias",
    "Fornecedor de [TEMA] para Empresas de Engenharia",
    "[TEMA] para Empresas de Saneamento",
    "[TEMA] para Concessionárias",
    "[TEMA] para Prefeituras",
    "[TEMA] em Loteamentos",
    "[TEMA] em Condomínios",
    "[TEMA] em Obras de Saneamento",
    "[TEMA] para Diferentes Aplicações",
    # Opções fixas adicionais para equipamentos de bombeamento e fluidos.
    "[TEMA] para Sistemas Hidráulicos",
    "[TEMA] para Transferência de Fluidos",
    "[TEMA] para Tratamento de Água",
    "[TEMA] para Processos Industriais",
    "[TEMA] para Indústrias Químicas",
    "[TEMA] para Indústrias Alimentícias",
    "[TEMA] para Mineração",
    "[TEMA] para Agricultura e Irrigação",
    "[TEMA] para Abastecimento de Água",
    "[TEMA] para Sistemas de Refrigeração",
]

SEGMENTOS_SERVICOS = [
    "Empresa de [TEMA] para Indústrias",
    "Empresa de [TEMA] para Centros Logísticos",
    "Empresa de [TEMA] para Centros de Distribuição",
    "[TEMA] para Hospitais",
    "[TEMA] para Condomínios Empresariais",
    "[TEMA] para Edifícios Corporativos",
    "[TEMA] para Centros Comerciais",
    "[TEMA] para Obras de Retrofit",
    "[TEMA] para Indústrias Metalúrgicas",
    "[TEMA] para Indústrias Siderúrgicas",
    "[TEMA] para Indústrias Químicas",
    "[TEMA] para Indústrias Alimentícias",
    "[TEMA] para Papel e Celulose",
    "[TEMA] para Mineração",
    "[TEMA] para Montadoras",
    "[TEMA] para Empresas de Automação Industrial",
    "[TEMA] para Usinas e Setor de Energia",
    "[TEMA] para Empresas de Saneamento",
    "[TEMA] para Centros Logísticos",
    "[TEMA] para Máquinas Industriais",
]

SEGMENTOS_CORINGAS = [
    "[TEMA] para Diferentes Aplicações",
    "[TEMA] para Aplicações Industriais",
    "[TEMA] para Empresas de Engenharia",
    "[TEMA] para Empresas Especializadas",
    "[TEMA] para Obras e Projetos",
    "[TEMA] para Diferentes Necessidades",
]


# ========================================================
# GERAR SUBTÍTULO DOS SEGMENTOS
# ========================================================

def gerar_subtitulo_segmentos(tema):

    global SUBTITULOS_SEGMENTOS_DISPONIVEIS

    tema = str(
        tema or ""
    ).strip()

    if not tema:
        return ""

    # Recarrega e embaralha quando todos os modelos
    # já tiverem sido utilizados
    if not SUBTITULOS_SEGMENTOS_DISPONIVEIS:

        SUBTITULOS_SEGMENTOS_DISPONIVEIS = list(
            SUBTITULOS_SEGMENTOS
        )

        random.shuffle(
            SUBTITULOS_SEGMENTOS_DISPONIVEIS
        )

    # Retira o próximo modelo da sequência
    modelo = SUBTITULOS_SEGMENTOS_DISPONIVEIS.pop(
        0
    )

    # Substitui o marcador pelo tema original
    subtitulo = modelo.replace(
        "[TEMA]",
        tema
    )

    return str(
        subtitulo
    ).strip()
    
    

# ============================================================
# GERAR SEGMENTOS VÁLIDOS PELO PYTHON
# ============================================================
#
# A biblioteca fixa é a única origem dos segmentos.
# O Ollama não participa desta etapa.
#
# Produtos:
#   SEGMENTOS_PRODUTOS + SEGMENTOS_CORINGAS
#
# Serviços:
#   SEGMENTOS_SERVICOS + SEGMENTOS_CORINGAS
#
# O tema original é preservado, inclusive acentuação.
# ============================================================

def gerar_segmentos_validos(
    tema,
    tipo
):

    tema = str(
        tema or ""
    ).strip()

    if not tema:
        return []

    tipo_normalizado = str(
        tipo or ""
    ).strip().casefold()

    if tipo_normalizado == "servico":

        biblioteca = (
            SEGMENTOS_SERVICOS
            +
            SEGMENTOS_CORINGAS
        )

    else:

        biblioteca = (
            SEGMENTOS_PRODUTOS
            +
            SEGMENTOS_CORINGAS
        )

    segmentos_validos = []

    for segmento in biblioteca:

        segmento = str(
            segmento or ""
        ).strip()

        if not segmento:
            continue

        segmento = segmento.replace(
            "[TEMA]",
            tema
        )

        segmento = re.sub(
            r"\s+",
            " ",
            segmento
        ).strip()

        if not segmento:
            continue

        # Todo segmento precisa conter o tema.
        if tema.casefold() not in segmento.casefold():
            continue

        # Evitar duplicidade.
        if segmento.casefold() in [
            item.casefold()
            for item in segmentos_validos
        ]:
            continue

        segmentos_validos.append(
            segmento
        )

    return segmentos_validos

    
    
# ============================================================
# CARREGAR MEAD
# ============================================================

def carregar_mead():

    if not os.path.exists(ARQUIVO_MEAD):

        print("❌ Arquivo MEAD não encontrado:")
        print(ARQUIVO_MEAD)

        return {}

    try:

        with open(
            ARQUIVO_MEAD,
            "r",
            encoding="utf-8"
        ) as arquivo:

            mead = json.load(arquivo)

        if not isinstance(mead, dict):

            print("❌ MEAD inválido. Estrutura esperada: objeto JSON.")

            return {}

        # Log compacto: não despejar o JSON inteiro do MEAD.
        print("MEAD carregado: estrutura válida | chaves:", len(mead))

        return mead

    except json.JSONDecodeError as erro:

        print("❌ Erro no JSON do MEAD:")
        print(erro)

        return {}

    except Exception as erro:

        print("❌ Erro ao carregar MEAD:")
        print(erro)

        return {}


# ============================================================
# RESERVAR VARIÁVEL MEAD; CARREGAMENTO ÚNICO MAIS ABAIXO
# ============================================================

MEAD = {}



# ============================================================
# PREPARAR MEAD PARA IA
# ============================================================

def preparar_mead(mead):

    """
    Prepara somente as regras editoriais necessárias do MEAD
    para serem utilizadas pelo Ollama.

    IMPORTANTE:
    - Não envia o JSON MEAD inteiro.
    - Não envia patrimônio ou textos de pesquisa.
    - Não seleciona informações.
    - Não altera os fragmentos escolhidos pelo Python.
    - O Python continua responsável pela seleção.
    - O Ollama continua responsável pela redação.
    """

    if not isinstance(mead, dict):
        return ""

    try:

        identidade = mead.get("identidade", mead)
        principio = mead.get("principio_operacional", {})
        preparacao = mead.get("preparacao", mead)
        tags = preparacao.get("tags", mead.get("tags", {}))
        construcao = mead.get("construcao_conteudo", {})
        estrutura = mead.get(
            "estrutura_pagina_editorial",
            mead.get("estrutura_pagina", {})
        )
        seo = mead.get("regras_seo", {})
        fidelidade = mead.get("fidelidade", {})
        contexto_blocos = mead.get("contexto_blocos", {})
        lista_segmentos = mead.get("lista_segmentos", {})

        mead_ia = {

            "metodo": identidade.get(
                "metodo",
                "MEAD"
            ),

            "objetivo": identidade.get(
                "objetivo",
                ""
            ),

            "principio_operacional": {

                "regra_principal": principio.get(
                    "regra_principal",
                    ""
                ),

                "regra_de_fidelidade": principio.get(
                    "regra_de_fidelidade",
                    ""
                ),

                "regra_de_protagonismo": principio.get(
                    "regra_de_protagonismo",
                    ""
                ),

                "regra_de_diversidade": principio.get(
                    "regra_de_diversidade",
                    ""
                ),

                "regra_de_separacao": principio.get(
                    "regra_de_separacao",
                    ""
                )
            },

            # ==================================================
            # PALAVRA-CHAVE
            # ==================================================

            "palavra_chave": preparacao.get(
                "palavra_chave",
                {}
            ),

            # ==================================================
            # REGRAS DE TAGS PARA O OLLAMA
            # ==================================================

            "tags": {

                "quantidade": tags.get(
                    "quantidade",
                    30
                ),

                "regra": tags.get(
                    "regra",
                    ""
                ),

                "principio_origem": tags.get(
                    "principio_origem",
                    {}
                ),

                "grupos": tags.get(
                    "grupos",
                    []
                ),

                "ordem_prioridade": tags.get(
                    "ordem_prioridade",
                    []
                ),

                "regra_localizacao": tags.get(
                    "regra_localizacao",
                    {}
                ),

                "regra_servicos": {

                    "usar": tags.get(
                        "regra_servicos",
                        {}
                    ).get(
                        "usar",
                        True
                    ),

                    "base": tags.get(
                        "regra_servicos",
                        {}
                    ).get(
                        "base",
                        ""
                    ),

                    "variacoes_permitidas": tags.get(
                        "regra_servicos",
                        {}
                    ).get(
                        "variacoes_permitidas",
                        []
                    ),

                    "regra": tags.get(
                        "regra_servicos",
                        {}
                    ).get(
                        "regra",
                        ""
                    )
                },

                "regra_produtos": {

                    "usar": tags.get(
                        "regra_produtos",
                        {}
                    ).get(
                        "usar",
                        True
                    ),

                    "base": tags.get(
                        "regra_produtos",
                        {}
                    ).get(
                        "base",
                        ""
                    ),

                    "variacoes_permitidas": tags.get(
                        "regra_produtos",
                        {}
                    ).get(
                        "variacoes_permitidas",
                        []
                    ),

                    "regra": tags.get(
                        "regra_produtos",
                        {}
                    ).get(
                        "regra",
                        ""
                    )
                },

                "regra_completude": tags.get(
                    "regra_completude",
                    {}
                ),

                "regra_substituicao": tags.get(
                    "regra_substituicao",
                    {}
                ),

                "restricoes": tags.get(
                    "restricoes",
                    []
                )
            },

            # ==================================================
            # CONTEXTO DOS CINCO BLOCOS
            # ==================================================

            "contexto_blocos": contexto_blocos,

            # ==================================================
            # CONSTRUÇÃO DO CONTEÚDO
            # ==================================================

            "construcao_conteudo": {

                "total_blocos": construcao.get(
                    "total_blocos",
                    5
                ),

                "paragrafos_por_bloco": construcao.get(
                    "paragrafos_por_bloco",
                    3
                ),

                "total_paragrafos": construcao.get(
                    "total_paragrafos",
                    15
                ),

                "palavras_por_paragrafo":
                    construcao.get(
                        "palavras_por_paragrafo",
                        {}
                    ),

                "regra_narrativa":
                    construcao.get(
                        "regra_narrativa",
                        {}
                    )
            },

            # ==================================================
            # ESTRUTURA DA PÁGINA
            # ==================================================

            "estrutura_pagina": {

                "blocos_texto": estrutura.get(
                    "blocos_texto",
                    5
                ),

                "paragrafos_por_bloco":
                    estrutura.get(
                        "paragrafos_por_bloco",
                        3
                    ),

                "total_paragrafos":
                    estrutura.get(
                        "total_paragrafos",
                        15
                    ),

                "palavras_por_paragrafo":
                    estrutura.get(
                        "palavras_por_paragrafo",
                        {}
                    )
            },


           

            # ==================================================
            # SEO
            # ==================================================

            "regras_seo": {

                "evitar": seo.get(
                    "evitar",
                    []
                ),

                "priorizar": seo.get(
                    "priorizar",
                    []
                )
            },

            # ==================================================
            # FIDELIDADE
            # ==================================================

            "fidelidade": {

                "permitido": fidelidade.get(
                    "permitido",
                    []
                ),

                "proibido": fidelidade.get(
                    "proibido",
                    []
                )
            }
        }

        return json.dumps(
            mead_ia,
            ensure_ascii=False,
            indent=2
        )

    except Exception as erro:

        print("❌ Erro ao preparar MEAD para IA:")
        print(erro)

        return ""


# ============================================================
# PREPARAR DADOS FIXOS DA PÁGINA
# ============================================================

def preparar_dados_pagina(
        tema,
        tags=None,
        segmentos_listas=None,
        dados_pagina=None
    ):

    """
    Organiza os dados já produzidos pelo Python
    para serem enviados posteriormente para
    montar_pagina_json().

    NÃO pesquisa.
    NÃO chama Ollama.
    NÃO gera conteúdo.
    NÃO altera fragmentos.
    NÃO grava o JSON.

    IMPORTANTE:
    - Não cria mais "informacoes_adicionais".
    - Não cria "categorias".
    - Os parágrafos da página devem ser recebidos
      posteriormente pelo fluxo de montagem dos blocos.
    """

    try:

        # ----------------------------------------------------
        # TEMA
        # ----------------------------------------------------

        tema = str(
            tema or ""
        ).strip()

        if not tema:

            print(
                "❌ preparar_dados_pagina(): "
                "tema vazio."
            )

            return None

        # ----------------------------------------------------
        # DADOS DA PÁGINA
        # ----------------------------------------------------

        if not isinstance(
            dados_pagina,
            dict
        ):

            dados_pagina = {}

        # ----------------------------------------------------
        # TAGS
        # ----------------------------------------------------

        if not isinstance(
            tags,
            list
        ):

            tags = []

        tags_finais = []

        for tag in tags:

            tag = str(
                tag or ""
            ).strip()

            if not tag:
                continue

            if tag in tags_finais:
                continue

            tags_finais.append(
                tag
            )

            if len(
                tags_finais
            ) >= 30:

                break

        dados_pagina[
            "tags"
        ] = tags_finais

        # ----------------------------------------------------
        # SEGMENTOS / LISTAS
        # ----------------------------------------------------
        
        if not isinstance(
            segmentos_listas,
            dict
        ):
        
            segmentos_listas = {}
        
        segmentos_oficiais = {}
        
        # ----------------------------------------------------
        # GARANTIR OS 12 SEGMENTOS
        # ----------------------------------------------------
        
        for numero in range(
            1,
            13
        ):
        
            chave = (
                f"segmento_{numero}"
            )
        
            lista = segmentos_listas.get(
                chave,
                []
            )
        
            if not isinstance(
                lista,
                list
            ):
        
                lista = []
        
            lista_final = []
        
            for item in lista:
        
                item = str(
                    item or ""
                ).strip()
        
                if not item:
                    continue
        
                lista_final.append(
                    item
                )
        
            segmentos_oficiais[
                chave
            ] = lista_final
        
        dados_pagina[
            "segmentos_listas"
        ] = segmentos_oficiais
        
        # ----------------------------------------------------
        # GERAR SUBTÍTULO DOS SEGMENTOS PELO PYTHON
        # ----------------------------------------------------
        
        dados_pagina[
            "subtitulo_segmentos"
        ] = gerar_subtitulo_segmentos(
            tema
        )
        
        # ----------------------------------------------------
        # CAPITALIZAÇÃO EDITORIAL
        # ----------------------------------------------------
        
        # Subtítulo
        dados_pagina[
            "subtitulo"
        ] = capitalizar_texto_editorial(
            dados_pagina.get(
                "subtitulo",
                ""
            )
        )
        
        # Subtítulo dos segmentos
        dados_pagina[
            "subtitulo_segmentos"
        ] = capitalizar_texto_editorial(
            dados_pagina.get(
                "subtitulo_segmentos",
                ""
            )
        )
        
        # Segmentos
        for chave, lista in dados_pagina[
            "segmentos_listas"
        ].items():
        
            if not isinstance(
                lista,
                list
            ):
                continue
        
            dados_pagina[
                "segmentos_listas"
            ][chave] = [
                capitalizar_texto_editorial(
                    segmento
                )
                for segmento in lista
            ]

        # ----------------------------------------------------
        # IMPORTANTE
        # ----------------------------------------------------
        #
        # NÃO criar:
        #
        # dados_pagina["informacoes_adicionais"]
        #
        # NÃO criar:
        #
        # dados_pagina["categorias"]
        #
        # Essas estruturas antigas não fazem mais parte
        # da estrutura editorial final da página.
        #
        # Os textos dos 15 parágrafos devem entrar nos
        # blocos da página posteriormente.
        # ----------------------------------------------------

        dados_pagina.pop(
            "informacoes_adicionais",
            None
        )

        dados_pagina.pop(
            "categorias",
            None
        )

        # ----------------------------------------------------
        # RETORNO
        # ----------------------------------------------------

        print()
        print(
            "=========================================="
        )
        print(
            "DADOS FIXOS DA PÁGINA PREPARADOS"
        )
        print(
            "=========================================="
        )
        print(
            f"TEMA: {tema}"
        )
        print(
            f"TAGS: {len(tags_finais)}"
        )

        total_itens = 0

        for numero in range(
            1,
            13
        ):

            chave = (
                f"segmento_{numero}"
            )

            quantidade = len(
                segmentos_oficiais[
                    chave
                ]
            )

            total_itens += quantidade

            print(
                f"{chave}: "
                f"{quantidade} itens"
            )

        print(
            f"TOTAL ITENS DAS LISTAS: "
            f"{total_itens}"
        )

        print(
            "INFORMACOES_ADICIONAIS: REMOVIDO"
        )

        print(
            "CATEGORIAS: REMOVIDO"
        )

        print(
            "=========================================="
        )

        return dados_pagina

    except Exception as erro:

        print()
        print(
            "❌ ERRO ao preparar dados da página:"
        )
        print(erro)

        return None


# ============================================================
# CRIAR ESTRUTURA DO NOVO JSON DA PÁGINA
# ============================================================

def criar_estrutura_json_pagina(tema):

    tema = str(
        tema or ""
    ).strip()

    if not tema:
        return {}

    return {

        tema: {

            # ====================================================
            # IDENTIFICAÇÃO
            # ====================================================

            "tema":
                "",

            "nome_site":
                "",

            "grupo_principal_projeto":
                "",

            # ====================================================
            # SEGMENTOS / FONTES / REFERÊNCIAS
            # ====================================================

            "segmentos_textuais":
                [],

            "fontes":
                [],

            "referencias":
                [],

            "trechos_utilizados":
                [],

            # ====================================================
            # CLASSIFICAÇÃO
            # ====================================================

            "grupo":
                "",

            "tipo":
                "",

            # ====================================================
            # TAGS
            # ====================================================

            "tags":
                [],

            # ====================================================
            # CONTROLE DE REPETIÇÕES
            # ====================================================

            "controle_repeticoes": {

                "palavra_chave":
                    "",

                "meta_repeticoes":
                    60,

                "repeticoes_realizadas":
                    0,

                "repeticoes_faltantes":
                    60
            },

            # ====================================================
            # MAPA MEAD
            # ====================================================

            "mapa_mead": {

                "status":
                    "",

                "texto":
                    ""
            },

            # ====================================================
            # PÁGINA
            # ====================================================

            "pagina": {

                # ==================================================
                # IDENTIFICAÇÃO DA PÁGINA
                # ==================================================

                "tema":
                    "",

                "arquivo_origem":
                    "",

                "h1":
                    "",

                "titulo":
                    "",

                "subtitulo":
                    "",

                "subtitulo_listas":
                    "",

                "subtitulo_segmentos":
                    "",

                "descricao":
                    "",


                # ==================================================
                # BLOCO 1
                # ==================================================

                "bloco_1": {

                    "id":
                        "bloco_1",

                    "hash":
                        "",

                    "informacoes_relevantes":
                        [],

                    "titulo":
                        "",

                    "paragrafos_python":
                        [
                            "",
                            "",
                            ""
                        ],

                    "paragrafos_ollama":
                        [
                            "",
                            "",
                            ""
                        ],

                },


                # ==================================================
                # BLOCO 2
                # ==================================================

                "bloco_2": {

                    "id":
                        "bloco_2",

                    "hash":
                        "",

                    "informacoes_relevantes":
                        [],

                    "titulo":
                        "",

                    "paragrafos_python":
                        [
                            "",
                            "",
                            ""
                        ],

                    "paragrafos_ollama":
                        [
                            "",
                            "",
                            ""
                        ],

                },


                # ==================================================
                # BLOCO 3
                # ==================================================

                "bloco_3": {

                    "id":
                        "bloco_3",

                    "hash":
                        "",

                    "informacoes_relevantes":
                        [],

                    "titulo":
                        "",

                    "paragrafos_python":
                        [
                            "",
                            "",
                            ""
                        ],

                    "paragrafos_ollama":
                        [
                            "",
                            "",
                            ""
                        ],

                },


                # ==================================================
                # BLOCO 4
                # ==================================================

                "bloco_4": {

                    "id":
                        "bloco_4",

                    "hash":
                        "",

                    "informacoes_relevantes":
                        [],

                    "titulo":
                        "",

                    "paragrafos_python":
                        [
                            "",
                            "",
                            ""
                        ],

                    "paragrafos_ollama":
                        [
                            "",
                            "",
                            ""
                        ],

                },


                # ==================================================
                # BLOCO 5
                # ==================================================

                "bloco_5": {

                    "id":
                        "bloco_5",

                    "hash":
                        "",

                    "informacoes_relevantes":
                        [],

                    "titulo":
                        "",

                    "paragrafos_python":
                        [
                            "",
                            "",
                            ""
                        ],

                    "paragrafos_ollama":
                        [
                            "",
                            "",
                            ""
                        ],


                },


                # ==================================================
                # SEGMENTOS GERADOS PELO PYTHON
                # ==================================================

                "segmentos_listas": {

                    "segmento_1":
                        [],

                    "segmento_2":
                        [],

                    "segmento_3":
                        [],

                    "segmento_4":
                        [],

                    "segmento_5":
                        [],

                    "segmento_6":
                        [],

                    "segmento_7":
                        [],

                    "segmento_8":
                        [],

                    "segmento_9":
                        [],

                    "segmento_10":
                        [],

                    "segmento_11":
                        [],

                    "segmento_12":
                        []
                },


                # ==================================================
                # POSICIONAMENTO DAS LISTAS
                # ==================================================

                "posicionamento_listas": {

                    "bloco":
                        None
                },


                # ==================================================
                # IMAGENS
                # ==================================================

                "imagens": {

                    "imagem_1": {

                        "url":
                            "",

                        "arquivo":
                            "",

                        "alt":
                            "",

                        "descricao":
                            ""
                    },

                    "imagem_2": {

                        "url":
                            "",

                        "arquivo":
                            "",

                        "alt":
                            "",

                        "descricao":
                            ""
                    },

                    "imagem_3": {

                        "url":
                            "",

                        "arquivo":
                            "",

                        "alt":
                            "",

                        "descricao":
                            ""
                    },

                    "imagem_4": {

                        "url":
                            "",

                        "arquivo":
                            "",

                        "alt":
                            "",

                        "descricao":
                            ""
                    },

                    "imagem_5": {

                        "url":
                            "",

                        "arquivo":
                            "",

                        "alt":
                            "",

                        "descricao":
                            ""
                    },

                    "imagem_6": {

                        "url":
                            "",

                        "arquivo":
                            "",

                        "alt":
                            "",

                        "descricao":
                            ""
                    }
                },


                # ==================================================
                # CONTROLE DA PÁGINA
                # ==================================================

                "caracteres":
                    0,

                "status":
                    "em_construcao"
            }
        }
    }        
    
    
# ============================================================
# NORMALIZAR GRUPO PRINCIPAL DO PROJETO
# ============================================================

def normalizar_grupo_principal_projeto(
    grupo_principal_projeto
):

    grupo = str(
        grupo_principal_projeto or ""
    ).strip()

    if not grupo:
        return ""

    # ========================================================
    # CORREÇÕES CONTROLADAS DE ACENTUAÇÃO
    # ========================================================

    correcoes = {

        "equipamentos hidraulicos":
            "equipamentos hidráulicos",

        "equipamento hidraulico":
            "equipamento hidráulico",

        "equipamentos hidráulicos":
            "equipamentos hidráulicos",

        "equipamento hidráulico":
            "equipamento hidráulico",

        "valvulas industriais":
            "válvulas industriais",

        "valvulas industrial":
            "válvulas industrial",

        "componentes mecanicos":
            "componentes mecânicos",

        "componente mecanico":
            "componente mecânico",

        "protecao contra incendio":
            "proteção contra incêndio",

        "proteção contra incêndio":
            "proteção contra incêndio",

        "materiais de construcao":
            "materiais de construção",

        "materiais de construção":
            "materiais de construção",

        "construcao civil":
            "construção civil",

        "construção civil":
            "construção civil",

        "instalacoes industriais":
            "instalações industriais",

        "instalações industriais":
            "instalações industriais",

        "sistemas hidraulicos":
            "sistemas hidráulicos",

        "sistemas hidráulicos":
            "sistemas hidráulicos"
    }

    chave = grupo.casefold()

    grupo_corrigido = correcoes.get(
        chave,
        grupo
    )

    return grupo_corrigido

    
# ============================================================
# MONTAR PÁGINA FINAL NO NOVO JSON OFICIAL
# ============================================================

def montar_pagina_json(
        tema,
        grupo="",
        tipo="",
        tags=None,
        controle_repeticoes=None,
        mapa_mead=None,
        informacoes_relevantes=None,
        dados_pagina=None,
        grupo_principal_projeto=""
    ):

    """
    Monta exclusivamente a estrutura oficial do JSON.

    RESPONSABILIDADES:

    - transportar os dados já produzidos pelo Python;
    - preservar o tema original;
    - preservar tags;
    - preservar controle de repetições;
    - preservar mapa MEAD;
    - preservar grupo principal;
    - preservar h1 e titulo;
    - preservar subtitulo;
    - preservar subtitulo_segmentos;
    - preservar os 5 blocos;
    - preservar informacoes_relevantes;
    - preservar paragrafos_python;
    - preservar paragrafos_ollama;
    - preservar segmentos_listas;
    - preservar posicionamento_listas.

    ESTA FUNÇÃO NÃO:

    - pesquisa;
    - chama Ollama;
    - seleciona fontes;
    - seleciona fragmentos;
    - gera conteúdo;
    - gera títulos;
    - gera segmentos;
    - gera imagens;
    - cria descricao;
    - cria subtitulo_listas;
    - cria paragrafos;
    - calcula caracteres;
    - cria status;
    - salva o banco;
    - chama salvar_banco();
    - chama atualizar_bloco();
    - cria informacoes_adicionais;
    - cria categorias.
    """

    try:

        # ====================================================
        # NORMALIZAÇÃO BÁSICA
        # ====================================================

        tema = str(
            tema or ""
        ).strip()

        grupo = str(
            grupo or ""
        ).strip()

        tipo = str(
            tipo or ""
        ).strip()

        grupo_principal_projeto = normalizar_grupo_principal_projeto(
            grupo_principal_projeto
        )


        if not tema:

            print(
                "❌ Não foi possível montar JSON: "
                "tema vazio."
            )

            return None


        # ====================================================
        # NORMALIZAR ENTRADAS
        # ====================================================

        if not isinstance(
            tags,
            list
        ):
            tags = []


        if not isinstance(
            controle_repeticoes,
            dict
        ):
            controle_repeticoes = {}


        if not isinstance(
            mapa_mead,
            dict
        ):
            mapa_mead = {}


        if not isinstance(
            dados_pagina,
            dict
        ):
            dados_pagina = {}


        # ====================================================
        # INFORMACOES_RELEVANTES
        #
        # São os objetos dos fragmentos selecionados
        # anteriormente pelo Python.
        #
        # NÃO transformar em string.
        # ====================================================

        if isinstance(
            informacoes_relevantes,
            dict
        ):

            if isinstance(
                informacoes_relevantes.get(
                    "blocos"
                ),
                dict
            ):

                fragmentos_blocos = (
                    informacoes_relevantes[
                        "blocos"
                    ]
                )

            else:

                fragmentos_blocos = (
                    informacoes_relevantes
                )

        elif isinstance(
            informacoes_relevantes,
            list
        ):

            fragmentos_blocos = {
                "bloco_1":
                    informacoes_relevantes
            }

        else:

            fragmentos_blocos = {}


        # ====================================================
        # TAGS
        # ====================================================

        tags_finais = []

        for tag in tags:

            if tag is None:
                continue

            tag_limpa = str(
                tag
            ).strip()

            if not tag_limpa:
                continue

            if tag_limpa in tags_finais:
                continue

            tags_finais.append(
                tag_limpa
            )

            if len(
                tags_finais
            ) >= 30:
                break


        # ====================================================
        # CONTROLE DE REPETIÇÕES
        # ====================================================

        controle_padrao = {

            "palavra_chave":
                tema,

            "meta_repeticoes":
                60,

            "repeticoes_realizadas":
                0,

            "repeticoes_faltantes":
                60
        }


        controle_padrao.update(
            controle_repeticoes
        )


        # ====================================================
        # DADOS GERAIS DA PÁGINA
        # ====================================================

        arquivo_origem = str(
            dados_pagina.get(
                "arquivo_origem",
                ""
            ) or ""
        ).strip()


        # ====================================================
        # H1
        #
        # NÃO gerar aqui.
        #
        # O valor deve vir pronto da etapa anterior.
        # ====================================================

        h1 = str(
            dados_pagina.get(
                "h1",
                ""
            ) or ""
        ).strip()


        # ====================================================
        # TÍTULO
        #
        # NÃO gerar aqui.
        #
        # O valor deve vir pronto da etapa anterior.
        # ====================================================

        titulo = str(
            dados_pagina.get(
                "titulo",
                ""
            ) or ""
        ).strip()


        # ====================================================
        # SUBTÍTULO
        # ====================================================

        subtitulo = str(
            dados_pagina.get(
                "subtitulo",
                ""
            ) or ""
        ).strip()


        # ====================================================
        # SUBTÍTULO DOS SEGMENTOS
        # ====================================================

        subtitulo_segmentos = str(
            dados_pagina.get(
                "subtitulo_segmentos",
                ""
            ) or ""
        ).strip()


        # ====================================================
        # SEGMENTOS
        #
        # Produzidos anteriormente pelo Python.
        # ====================================================

        segmentos_listas = dados_pagina.get(
            "segmentos_listas",
            {}
        )


        if not isinstance(
            segmentos_listas,
            dict
        ):

            segmentos_listas = {}


        # ====================================================
        # POSICIONAMENTO DAS LISTAS
        #
        # Mantém o bloco escolhido anteriormente
        # pelo Python.
        # ====================================================

        posicionamento_listas = dados_pagina.get(
            "posicionamento_listas",
            {
                "bloco": None
            }
        )


        if not isinstance(
            posicionamento_listas,
            dict
        ):

            posicionamento_listas = {
                "bloco": None
            }


        # ====================================================
        # MONTAR OS 5 BLOCOS
        # ====================================================

        blocos_finais = {}


        for numero_bloco in range(
            1,
            6
        ):

            chave_bloco = (
                f"bloco_{numero_bloco}"
            )


            bloco_recebido = dados_pagina.get(
                chave_bloco,
                {}
            )


            if not isinstance(
                bloco_recebido,
                dict
            ):

                bloco_recebido = {}


            # =================================================
            # ID
            # =================================================

            id_bloco = bloco_recebido.get(
                "id",
                chave_bloco
            )


            if not id_bloco:

                id_bloco = chave_bloco


            # =================================================
            # HASH
            # =================================================

            hash_bloco = str(
                bloco_recebido.get(
                    "hash",
                    ""
                ) or ""
            )


            # =================================================
            # INFORMACOES_RELEVANTES
            #
            # Recuperadas dos fragmentos selecionados.
            #
            # Os objetos são preservados.
            # =================================================

            informacoes_bloco = (
                fragmentos_blocos.get(
                    chave_bloco,
                    []
                )
            )


            if not isinstance(
                informacoes_bloco,
                list
            ):

                informacoes_bloco = []


            informacoes_finais = []


            for fragmento in informacoes_bloco:

                if not isinstance(
                    fragmento,
                    dict
                ):
                    continue

                informacoes_finais.append(
                    dict(
                        fragmento
                    )
                )


            # =================================================
            # TRECHOS SELECIONADOS PELO PYTHON
            #
            # IMPORTANTE:
            # não são textos escritos pelo Python.
            #
            # São os trechos selecionados.
            # =================================================

            paragrafos_python = (
                bloco_recebido.get(
                    "paragrafos_python",
                    []
                )
            )


            if not isinstance(
                paragrafos_python,
                list
            ):

                paragrafos_python = []


            paragrafos_python = list(
                paragrafos_python[:3]
            )


            while len(
                paragrafos_python
            ) < 3:

                paragrafos_python.append(
                    ""
                )


            # =================================================
            # PARÁGRAFOS GERADOS PELO OLLAMA
            # =================================================

            paragrafos_ollama = (
                bloco_recebido.get(
                    "paragrafos_ollama",
                    []
                )
            )


            if not isinstance(
                paragrafos_ollama,
                list
            ):

                paragrafos_ollama = []


            paragrafos_ollama = list(
                paragrafos_ollama[:3]
            )


            while len(
                paragrafos_ollama
            ) < 3:

                paragrafos_ollama.append(
                    ""
                )


            # =================================================
            # BLOCO FINAL
            #
            # ESTRUTURA OFICIAL:
            #
            # id
            # hash
            # informacoes_relevantes
            # paragrafos_python
            # paragrafos_ollama
            #
            # NÃO existe:
            # - titulo
            # - paragrafos
            # =================================================

            blocos_finais[
                chave_bloco
            ] = {

                "id":
                    id_bloco,

                "hash":
                    hash_bloco,

                "informacoes_relevantes":
                    informacoes_finais,

                "paragrafos_python":
                    paragrafos_python,

                "paragrafos_ollama":
                    paragrafos_ollama
            }


        # ====================================================
        # PÁGINA FINAL
        # ====================================================

        pagina_final = {

            "tema":
                tema,

            "arquivo_origem":
                arquivo_origem,

            "h1":
                h1,

            "titulo":
                titulo,

            "subtitulo":
                subtitulo,

            "subtitulo_segmentos":
                subtitulo_segmentos,

            "bloco_1":
                blocos_finais[
                    "bloco_1"
                ],

            "bloco_2":
                blocos_finais[
                    "bloco_2"
                ],

            "bloco_3":
                blocos_finais[
                    "bloco_3"
                ],

            "bloco_4":
                blocos_finais[
                    "bloco_4"
                ],

            "bloco_5":
                blocos_finais[
                    "bloco_5"
                ],

            "segmentos_listas":
                segmentos_listas,

            "posicionamento_listas":
                posicionamento_listas
        }


        # ====================================================
        # OBJETO FINAL DO TEMA
        # ====================================================

        dados_tema_final = {

            "tema":
                tema,

            "grupo":
                grupo,

            "tipo":
                tipo,

            "tags":
                tags_finais,

            "controle_repeticoes":
                controle_padrao,

            "mapa_mead":
                mapa_mead,

            "grupo_principal_projeto":
                grupo_principal_projeto,

            "pagina":
                pagina_final
        }


        # ====================================================
        # CHECK FINAL
        # ====================================================

        print()
        print(
            "=========================================="
        )
        print(
            "JSON OFICIAL MONTADO"
        )
        print(
            "=========================================="
        )

        print(
            "TEMA:",
            tema
        )

        print(
            "GRUPO:",
            grupo
        )

        print(
            "TIPO:",
            tipo
        )

        print(
            "TAGS:",
            len(
                tags_finais
            )
        )

        print(
            "BLOCOS:",
            5
        )

        print(
            "SEGMENTOS:",
            len(
                segmentos_listas
            )
        )

        print(
            "POSICIONAMENTO:",
            posicionamento_listas
        )

        print(
            "=========================================="
        )


        return dados_tema_final


    except Exception as erro:

        print()
        print(
            "=========================================="
        )

        print(
            "❌ ERRO AO MONTAR JSON OFICIAL"
        )

        print(
            "=========================================="
        )

        print(
            repr(erro)
        )

        print(
            "=========================================="
        )

        return None

# ========================================================
# 01. CARREGAR MEAD GLOBAL
# ========================================================

MEAD = carregar_mead()

MEAD_TEXTO = preparar_mead(MEAD)

print("\nMEAD preparado para IA:")
try:
    _mead_preview = json.loads(MEAD_TEXTO) if MEAD_TEXTO else {}
    print("MÉTODO:", _mead_preview.get("metodo", "MEAD"))
    print("OBJETIVO:", bool(_mead_preview.get("objetivo")))
    print("REGRAS OPERACIONAIS:", sum(
        1 for _v in _mead_preview.get("principio_operacional", {}).values()
        if str(_v or "").strip()
    ))
    print("CONTEXTO DOS BLOCOS:", len(_mead_preview.get("contexto_blocos", {})))
    print("REGRAS SEO:", bool(_mead_preview.get("regras_seo")))
except Exception:
    print("MEAD preparado; não foi possível resumir a estrutura.")


# ========================================================
# 02. GRUPOS TEMÁTICOS
# ========================================================

GRUPOS_TEMATICOS = {

    "protecao_contra_incendio": [

        "argamassa",
        "intumescente",
        "selagem",
        "corta fogo",
        "contra incendio",
        "lã de rocha",
        "la de rocha",
        "colar intumescente",
        "fita intumescente",
        "proteção passiva",
        "protecao passiva"

    ],

    "valvulas_industriais": [

        "válvula",
        "valvula",
        "hidraulica",
        "hidráulica",
        "esfera",
        "gaveta",
        "globo",
        "retenção",
        "retencao",
        "atuador"

    ],

    "componentes_mecanicos": [

        "terminal rotular",
        "terminal rotular esférico",
        "terminal rotular esferico",
        "rótula",
        "rotula",
        "rolamento",
        "mancal",
        "articulação",
        "articulacao",
        "junta esférica",
        "junta esferica"

    ]

}




# ============================================================
# IDENTIFICAR TIPO DO TEMA
# ============================================================

def pesquisar(termo, limite=10):

    resultados = []


    dominios_bloqueados = [

        "pinterest.",
        "facebook.",
        "instagram.",
        "youtube.com",
        "tiktok."

    ]


    try:

        with DDGS() as ddgs:


            busca = ddgs.text(
                termo,
                safesearch="off",
                max_results=limite
            )


            for item in busca:


                if not isinstance(item, dict):

                    continue


                url = item.get("href")


                if not url:

                    continue


                if any(
                    dominio in url.lower()
                    for dominio in dominios_bloqueados
                ):

                    continue


                resultados.append(url)



    except Exception as e:

        print()
        print("ERRO NA PESQUISA:")
        print(e)



    return resultados


# ============================================================
# FILTRAR URLS RUINS
# ============================================================

def filtrar_urls(urls):


    bloqueados = [

        "mercadolivre",
        "amazon",
        "shopee",
        "aliexpress",
        "alibaba",
        "ebay",
        "made-in-china",

        "facebook",
        "instagram",
        "youtube",
        "pinterest",
        "tiktok"

    ]


    resultado = []


    for url in urls:


        if not url:

            continue


        url_lower = url.lower()


        if any(
            item in url_lower
            for item in bloqueados
        ):

            continue


        resultado.append(url)



    return resultado




# ============================================================
# FILTRAR URL
# ============================================================

def url_valida(url):


    if not url:

        return False


    url_lower = url.lower()



    dominios_bloqueados = [

        "oceanofpdf.com",
        "pdfcoffee.com",
        "scribd.com",

        "pinterest.com",
        "facebook.com",
        "instagram.com",

        "tiktok.com",
        "youtube.com",
        "youtu.be"

    ]



    palavras_bloqueadas = [

        "ebook",
        "torrent",
        "download",
        "pdf-epub",
        "curso",
        "apostila",
        "manual-download"

    ]



    if any(
        dominio in url_lower
        for dominio in dominios_bloqueados
    ):

        return False



    if any(
        palavra in url_lower
        for palavra in palavras_bloqueadas
    ):

        return False



    return True
    
    
# ============================================================
# IDENTIFICAR EMPRESA DENTRO DO PRÓPRIO SITE
# ============================================================
#
# A pesquisa principal encontra fontes técnicas.
#
# Esta etapa é complementar:
# - pega os domínios das fontes aprovadas;
# - procura páginas institucionais dentro do mesmo domínio;
# - tenta identificar a empresa oficialmente;
# - grava a identidade na própria fonte.
#
# Não utiliza outro domínio.
# Não utiliza pesquisa externa para descobrir a empresa.
# ============================================================

def identificar_empresa_no_site(paginas):

    if not isinstance(paginas, list):
        return paginas

    # --------------------------------------------------------
    # Palavras que indicam páginas institucionais
    # --------------------------------------------------------

    termos_institucionais = [
        "empresa",
        "quem somos",
        "sobre",
        "institucional",
        "a empresa",
        "sobre nós",
        "sobre-nos",
        "quem-somos"
    ]

    # --------------------------------------------------------
    # Domínios já processados
    # --------------------------------------------------------

    identidades_por_dominio = {}

    # --------------------------------------------------------
    # Normalizar domínio
    # --------------------------------------------------------

    def obter_dominio(url):

        try:

            from urllib.parse import urlparse

            resultado = urlparse(
                str(url or "").strip()
            )

            dominio = resultado.netloc.lower()

            dominio = re.sub(
                r"^www\.",
                "",
                dominio
            )

            return dominio.strip()

        except Exception:

            return ""

    # --------------------------------------------------------
    # Procurar links institucionais dentro da página
    # --------------------------------------------------------

    def localizar_links_institucionais(
        url,
        html
    ):

        links_encontrados = []

        try:

            soup = BeautifulSoup(
                html,
                "html.parser"
            )

            dominio_origem = obter_dominio(
                url
            )

            for a in soup.find_all(
                "a",
                href=True
            ):

                href = str(
                    a.get(
                        "href",
                        ""
                    )
                ).strip()

                texto_link = a.get_text(
                    " ",
                    strip=True
                ).casefold()

                if not href:
                    continue

                # --------------------------------------------
                # O link precisa indicar uma página institucional
                # --------------------------------------------

                eh_institucional = False

                for termo in termos_institucionais:

                    if termo in texto_link:

                        eh_institucional = True
                        break

                # Também verifica o próprio endereço
                href_lower = href.casefold()

                if not eh_institucional:

                    for termo in termos_institucionais:

                        termo_url = (
                            termo
                            .replace(" ", "-")
                        )

                        if termo_url in href_lower:

                            eh_institucional = True
                            break

                if not eh_institucional:
                    continue

                # --------------------------------------------
                # Transformar URL relativa em absoluta
                # --------------------------------------------

                from urllib.parse import urljoin

                url_final = urljoin(
                    url,
                    href
                )

                dominio_link = obter_dominio(
                    url_final
                )

                # --------------------------------------------
                # Segurança:
                # permanecer no mesmo domínio
                # --------------------------------------------

                if not dominio_link:
                    continue

                if dominio_link != dominio_origem:
                    continue

                if url_final not in links_encontrados:

                    links_encontrados.append(
                        url_final
                    )

                # Limite para não fazer buscas excessivas
                if len(links_encontrados) >= 5:
                    break

        except Exception as erro:

            print(
                "⚠️ Erro ao localizar página institucional:",
                erro
            )

        return links_encontrados

    # --------------------------------------------------------
    # Processar cada fonte
    # --------------------------------------------------------

    for pagina in paginas:

        if not isinstance(
            pagina,
            dict
        ):
            continue

        url_fonte = str(
            pagina.get(
                "url",
                ""
            )
        ).strip()

        if not url_fonte:
            continue

        dominio = obter_dominio(
            url_fonte
        )

        if not dominio:
            continue

        # ----------------------------------------------------
        # Se este domínio já foi processado,
        # reutilizar a identidade.
        # ----------------------------------------------------

        if dominio in identidades_por_dominio:

            pagina[
                "identidade_fonte"
            ] = identidades_por_dominio[
                dominio
            ]

            continue

        print()
        print(
            "========================================"
        )
        print(
            "PESQUISA INSTITUCIONAL DO SITE"
        )
        print(
            "========================================"
        )

        print(
            "DOMÍNIO:",
            dominio
        )

        identidade = {
            "nome": "",
            "dominio": dominio,
            "papel": "",
            "origem_identificacao": "",
            "confianca": "baixa"
        }

        urls_institucionais = []

        # ----------------------------------------------------
        # Primeiro tenta a própria página da fonte.
        # Ela pode conter o link "Empresa".
        # ----------------------------------------------------

        try:

            resposta = requests.get(
                url_fonte,
                timeout=20,
                headers={
                    "User-Agent":
                        "Mozilla/5.0"
                }
            )

            if resposta.ok:

                urls_institucionais = (
                    localizar_links_institucionais(
                        url_fonte,
                        resposta.text
                    )
                )

        except Exception as erro:

            print(
                "⚠️ Não foi possível analisar a fonte:",
                erro
            )

        # ----------------------------------------------------
        # Se não encontrou link, tenta caminhos conhecidos
        # dentro do MESMO domínio.
        # ----------------------------------------------------

        if not urls_institucionais:

            base_url = (
                f"https://{dominio}"
            )

            caminhos = [
                "/empresa",
                "/quem-somos",
                "/quem_somos",
                "/sobre",
                "/sobre-nos",
                "/sobre_nos",
                "/institucional",
                "/a-empresa",
                "/a_empresa"
            ]

            for caminho in caminhos:

                url_teste = (
                    base_url
                    +
                    caminho
                )

                try:

                    resposta = requests.get(
                        url_teste,
                        timeout=15,
                        headers={
                            "User-Agent":
                                "Mozilla/5.0"
                        },
                        allow_redirects=True
                    )

                    if not resposta.ok:
                        continue

                    url_real = str(
                        resposta.url or ""
                    ).strip()

                    dominio_real = obter_dominio(
                        url_real
                    )

                    if dominio_real != dominio:
                        continue

                    if len(
                        resposta.text
                    ) < 500:

                        continue

                    urls_institucionais.append(
                        url_real
                    )

                    if len(
                        urls_institucionais
                    ) >= 3:

                        break

                except Exception:

                    continue

        # ----------------------------------------------------
        # Analisar páginas institucionais encontradas
        # ----------------------------------------------------

        for url_institucional in (
            urls_institucionais
        ):

            print(
                "PÁGINA INSTITUCIONAL:",
                url_institucional
            )

            try:

                resposta = requests.get(
                    url_institucional,
                    timeout=20,
                    headers={
                        "User-Agent":
                            "Mozilla/5.0"
                    },
                    allow_redirects=True
                )

                if not resposta.ok:
                    continue

                texto_institucional = ""

                try:

                    soup = BeautifulSoup(
                        resposta.text,
                        "html.parser"
                    )

                    for elemento in soup(
                        [
                            "script",
                            "style",
                            "noscript"
                        ]
                    ):

                        elemento.decompose()

                    texto_institucional = soup.get_text(
                        " ",
                        strip=True
                    )

                except Exception:

                    texto_institucional = ""

                texto_institucional = re.sub(
                    r"\s+",
                    " ",
                    texto_institucional
                ).strip()

                if len(
                    texto_institucional
                ) < 100:

                    continue

                # ------------------------------------------------
                # Usa a função de identificação já existente.
                # ------------------------------------------------

                resultado = identificar_empresa_fonte(
                    texto_institucional,
                    url_institucional
                )

                if isinstance(
                    resultado,
                    dict
                ):

                    nome = str(
                        resultado.get(
                            "nome",
                            ""
                        )
                    ).strip()

                    dominio_resultado = str(
                        resultado.get(
                            "dominio",
                            ""
                        )
                    ).strip()

                    if nome:

                        identidade = {
                            "nome": nome,
                            "dominio":
                                dominio_resultado
                                or dominio,
                            "papel":
                                str(
                                    resultado.get(
                                        "papel",
                                        ""
                                    )
                                ).strip(),
                            "origem_identificacao":
                                "pagina_institucional",
                            "confianca":
                                "alta"
                        }

                        break

            except Exception as erro:

                print(
                    "⚠️ Erro na página institucional:",
                    erro
                )


        # ----------------------------------------------------
        # SEGUNDA TENTATIVA:
        # identificar a empresa diretamente no texto
        # da própria fonte.
        #
        # Isso cobre casos em que o site não possui uma
        # página institucional acessível ou identificável.
        # ----------------------------------------------------

        if not identidade.get(
            "nome",
            ""
        ):

            texto_fonte = str(
                pagina.get(
                    "texto",
                    ""
                )
            ).strip()

            if len(
                texto_fonte
            ) >= 100:

                resultado = identificar_empresa_fonte(
                    texto_fonte,
                    url_fonte
                )

                if isinstance(
                    resultado,
                    dict
                ):

                    nome = str(
                        resultado.get(
                            "nome",
                            ""
                        )
                    ).strip()

                    dominio_resultado = str(
                        resultado.get(
                            "dominio",
                            ""
                        )
                    ).strip()

                    if nome:

                        identidade = {
                            "nome": nome,
                            "dominio":
                                dominio_resultado
                                or dominio,
                            "papel":
                                str(
                                    resultado.get(
                                        "papel",
                                        ""
                                    )
                                ).strip(),
                            "origem_identificacao":
                                "texto_da_fonte",
                            "confianca":
                                "alta"
                        }

                        print(
                            "IDENTIDADE ENCONTRADA "
                            "NO TEXTO DA FONTE:",
                            identidade
                        )


        # ----------------------------------------------------
        # Se não encontrou nome, pelo menos preserva domínio.
        # ----------------------------------------------------

        if not identidade.get(
            "dominio"
        ):

            identidade[
                "dominio"
            ] = dominio

        # ----------------------------------------------------
        # Guardar identidade por domínio
        # ----------------------------------------------------

        identidades_por_dominio[
            dominio
        ] = identidade

        # ----------------------------------------------------
        # Anexar à fonte atual
        # ----------------------------------------------------

        pagina[
            "identidade_fonte"
        ] = identidade

        print(
            "IDENTIDADE ENCONTRADA:",
            identidade
        )

    return paginas

    

# ============================================================
# PESQUISA COMPLETA
# ============================================================

def pesquisar_completo(tema):

    urls = []


    # ========================================================
    # 01. PESQUISA PDF BRASIL
    # ========================================================

    print()
    print("==============================")
    print("PESQUISA PDF BRASIL")
    print("==============================")


    consultas_pdf_brasil = [

        f"{tema} pdf",

        f"{tema} catálogo pdf",

        f"{tema} catalogo pdf",

        f"{tema} ficha técnica pdf",

        f"{tema} ficha tecnica pdf",

        f"{tema} manual pdf"

    ]


    total_pdf_brasil = 0


    for consulta in consultas_pdf_brasil:

        resultado = pesquisar(

            consulta,

            limite=5

        )

        urls.extend(resultado)

        total_pdf_brasil += len(resultado)


    print(

        "PDF BRASIL:",

        total_pdf_brasil

    )


    # ========================================================
    # 02. PESQUISA PDF EXTERIOR
    # ========================================================

    print()
    print("==============================")
    print("PESQUISA PDF EXTERIOR")
    print("==============================")


    consultas_pdf_exterior = [

        f"{tema} datasheet pdf",

        f"{tema} catalog pdf",

        f"{tema} technical data pdf",

        f"{tema} specification pdf",

        f"{tema} specifications pdf",

        f"{tema} installation manual pdf"

    ]


    total_pdf_exterior = 0


    for consulta in consultas_pdf_exterior:

        resultado = pesquisar(

            consulta,

            limite=5

        )

        urls.extend(resultado)

        total_pdf_exterior += len(resultado)


        print(
        "PDF EXTERIOR:",
        total_pdf_exterior
    )


    # ========================================================
    # 03. PESQUISA POR CATEGORIAS
    # ========================================================

    print()
    print("==============================")
    print("PESQUISA POR CATEGORIAS")
    print("==============================")


    mapa_categorias = {

        "definicao": "definição",

        "beneficios": "benefícios",

        "vantagens": "vantagens",

        "materia_prima": "materiais",

        "aplicacoes": "aplicações",

        "fabricacao": "fabricação",

        "manutencao": "manutenção",

        "ativos_narrativos": "problemas casos",

        "duvidas_frequentes": "dúvidas frequentes"

    }


    total_categorias = 0


    for categoria in CATEGORIAS:

        termo = mapa_categorias.get(

            categoria,

            categoria.replace("_", " ")

        )

        consulta = f"{tema} {termo}"


        print()

        print(
            "CONSULTA CATEGORIA:",
            consulta
        )


        resultado = pesquisar(

            consulta,

            limite=5

        )


        urls.extend(resultado)


        total_categorias += len(resultado)


    print()

    print(
        "FONTES POR CATEGORIAS:",
        total_categorias
    )



    # ========================================================
    # 03.1 BRASIL - 10 FONTES
    # ========================================================

    print()
    print("==============================")
    print("PESQUISA BRASIL")
    print("==============================")


    brasil = pesquisar(

        f"{tema} Brasil",

        limite=10

    )


    urls.extend(

        brasil

    )


    print(

        "BRASIL:",

        len(brasil)

    )


    # =====================================
    # EXTERIOR - 10 FONTES
    # PESQUISA INTERNACIONAL INTELIGENTE
    # =====================================

    print()
    print("==============================")
    print("PESQUISA EXTERIOR")
    print("==============================")


    exterior = []


    # ========================================================
    # 04. DEFINIR CONSULTAS PELO GRUPO DO TEMA
    # ========================================================

    consultas_exterior = []


    tema_normalizado = normalizar_texto(

        tema

    )


    grupo_encontrado = None


    for grupo, palavras in GRUPOS_TEMATICOS.items():

        for palavra in palavras:

            if normalizar_texto(

                palavra

            ) in tema_normalizado:

                grupo_encontrado = grupo

                break

        if grupo_encontrado:

            break
    
    
    # ========================================================
    # 05. COMPONENTES MECÂNICOS
    # ========================================================

    if grupo_encontrado == "componentes_mecanicos":

        consultas_exterior = [

            f"{tema} manufacturer",

            f"{tema} technical catalog",

            f"{tema} engineering",

            f"{tema} rod end",

            f"{tema} spherical rod end",

            f"{tema} spherical plain bearing",

            f"{tema} technical",

            f"{tema} application"

        ]


    # ========================================================
    # 06. PROTEÇÃO CONTRA INCÊNDIO
    # ========================================================

    elif grupo_encontrado == "protecao_contra_incendio":

        consultas_exterior = [

            f"{tema} fire protection",

            f"{tema} manufacturer",

            f"{tema} technical catalog",

            f"{tema} datasheet",

            f"{tema} application",

            f"{tema} passive fire protection"

        ]


    # ========================================================
    # 07. VÁLVULAS
    # ========================================================

    elif grupo_encontrado == "valvulas_industriais":

        consultas_exterior = [

            f"{tema} industrial valve",

            f"{tema} manufacturer",

            f"{tema} technical catalog",

            f"{tema} datasheet",

            f"{tema} engineering",

            f"{tema} application"

        ]


    # ========================================================
    # 08. PADRÃO
    # ========================================================

    else:

        consultas_exterior = [

            f"{tema} manufacturer",

            f"{tema} technical",

            f"{tema} engineering",

            f"{tema} catalog",

            f"{tema} application"

        ]


    # ========================================================
    # 09. EXECUTAR BUSCAS INTERNACIONAIS
    # ========================================================

    for consulta in consultas_exterior:

        encontrados = pesquisar(

            consulta,

            limite=5

        )

        exterior.extend(

            encontrados

        )


    # ========================================================
    # 10. REMOVER DUPLICADOS DO EXTERIOR
    # ========================================================

    exterior = list(

        dict.fromkeys(exterior)

    )


    urls.extend(

        exterior[:10]

    )


    print(

        "EXTERIOR:",

        len(exterior[:10])

    )


    # ========================================================
    # 11. REMOVER DUPLICADOS GERAIS
    # ========================================================

    urls_unicas = []

    vistos = set()


    for url in urls:

        if not url_valida(url):

            print("URL BLOQUEADA:")
            print(url)

            continue


        if url not in vistos:

            vistos.add(url)

            urls_unicas.append(url)


    # ========================================================
    # 12. PRIORIZAR PDFs
    # ========================================================

    urls_unicas.sort(

        key=lambda url: (

            ".pdf" not in url.lower(),

            url.lower()

        )

    )


    print()
    print("==============================")
    print("TOTAL FONTES")
    print("==============================")

    print(

        len(urls_unicas)

    )


    urls_unicas = filtrar_urls(

        urls_unicas

    )


    print()
    print("URLS APROVADAS PARA COLETA:")

    for i, url in enumerate(

        urls_unicas,

        start=1

    ):

        print(f"{i}. {url}")

    return urls_unicas


# ============================================================
# IDENTIFICAR GRUPO
# ============================================================

def identificar_grupo_tema(tema):

    tema_normalizado = normalizar_texto(tema)

    for grupo, palavras in GRUPOS_TEMATICOS.items():

        for palavra in palavras:

            if normalizar_texto(palavra) in tema_normalizado:

                return grupo

    return "geral"
    
    

# ============================================================
# ENTENDIMENTO INICIAL DO PRODUTO — PYTHON
# ============================================================
#
# O Python não precisa perguntar ao Ollama o que pesquisar.
#
# Esta função apenas transforma:
#   tema + grupo
#
# em uma orientação operacional para a pesquisa.
#
# NÃO afirma fatos técnicos.
# NÃO inventa características.
# NÃO pesquisa.
# NÃO chama Ollama.
#
# ============================================================

def gerar_entendimento_produto(
    tema,
    grupo_principal=""
):

    print()
    print("==============================")
    print("GERANDO ENTENDIMENTO INICIAL — PYTHON")
    print("==============================")

    tema = str(
        tema or ""
    ).strip()

    grupo_principal = str(
        grupo_principal or ""
    ).strip()

    if not tema:

        print(
            "TEMA VAZIO"
        )

        return ""

    if not grupo_principal:

        grupo_principal = "geral"

    # --------------------------------------------------------
    # PALAVRAS DO TEMA
    # --------------------------------------------------------

    palavras_tema = [

        palavra.strip()

        for palavra
        in re.findall(
            r"\b[\wÀ-ÿ-]{3,}\b",
            tema
        )

    ]

    # --------------------------------------------------------
    # REMOVER DUPLICIDADES
    # --------------------------------------------------------

    palavras_tema_unicas = []

    for palavra in palavras_tema:

        if (
            palavra.casefold()
            not in [
                item.casefold()
                for item
                in palavras_tema_unicas
            ]
        ):

            palavras_tema_unicas.append(
                palavra
            )

    # --------------------------------------------------------
    # TERMOS DE PESQUISA
    #
    # São orientações de busca.
    # Não são fatos sobre o produto.
    # --------------------------------------------------------

    termos_base = [

        tema,

        f"definição de {tema}",

        f"funcionamento de {tema}",

        f"aplicações de {tema}",

        f"características de {tema}",

        f"aspectos técnicos de {tema}",

        f"materiais de {tema}",

        f"manutenção de {tema}",

        f"dimensionamento de {tema}",

        f"instalação de {tema}",

        f"vantagens de {tema}"

    ]

    # --------------------------------------------------------
    # ORIENTAÇÃO DE PESQUISA
    # --------------------------------------------------------

    entendimento = {

        "tema":
            tema,

        "grupo":
            grupo_principal,

        "palavras_tema":
            palavras_tema_unicas,

        "objetivo":
            (
                "Localizar informações técnicas "
                "diretamente relacionadas ao tema."
            ),

        "focos_pesquisa": [

            "definição",
            "funcionamento",
            "características",
            "aplicações",
            "materiais",
            "fabricação",
            "manutenção",
            "instalação",
            "dimensionamento",
            "benefícios técnicos",
            "cuidados técnicos"

        ],

        "termos_pesquisa":
            termos_base,

        "fontes_prioritarias": [

            "manual técnico",
            "catálogo técnico",
            "documentação técnica",
            "norma técnica",
            "artigo técnico",
            "documentação institucional",
            "site técnico"

        ],

        "restricoes": [

            "não inventar informações",

            "não assumir características",

            "não assumir aplicações",

            "não assumir fabricantes",

            "não assumir marcas",

            "não assumir modelos",

            "não assumir números",

            "não assumir certificações"

        ]

    }

    resultado = json.dumps(
        entendimento,
        ensure_ascii=False,
        indent=2
    )

    print()
    print(
        "ENTENDIMENTO PYTHON GERADO"
    )

    print(
        "CARACTERES:",
        len(resultado)
    )

    return resultado

# ============================================================
# CAPTURA TEXTO
# ============================================================


def limpar_texto_coletado(texto):

    cortes = [

        "Principais cidades e regiões",

        "Solicite um orçamento",

        "Entre em contato",

        "Produtos relacionados",

        "Confira Também",

        "Crime de violação de direito autoral",

        "Política de Privacidade",

        "WhatsApp",

        "Online Fale com a gente",

        "Todos os direitos reservados"
    ]

    for corte in cortes:

        if corte in texto:

            texto = texto.split(corte)[0]

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    texto = texto.strip()

    return texto[:8000]
    

# ============================================================
# COLETAR PÁGINA
# ============================================================

def coletar_pagina(url):

    try:

        headers = {

            "User-Agent": (
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/120 Safari/537.36"
            )

        }

        resposta = requests.get(

            url,
            headers=headers,
            timeout=30

        )

        if resposta.status_code != 200:

            return None

        content_type = resposta.headers.get(

            "Content-Type",
            ""

        ).lower()

        eh_pdf = (

            "pdf" in content_type

            or

            url.lower().split("?")[0].endswith(".pdf")

        )

    # ========================================================
    # 01. FUNÇÃO INTERNA PARA EXTRAIR PDF
    # ========================================================

        def extrair_pdf(bytes_pdf, origem):

            arquivo_temp = None

            try:


                arquivo_temp = tempfile.NamedTemporaryFile(

                    delete=False,
                    suffix=".pdf"

                )

                arquivo_temp.write(

                    bytes_pdf

                )

                arquivo_temp.close()

                documento = fitz.open(

                    arquivo_temp.name

                )

                texto = ""

                for pagina in documento:

                    texto += pagina.get_text("text")

                documento.close()

                os.remove(

                    arquivo_temp.name

                )

                texto = limpar_texto_coletado(

                    texto

                )

                if len(texto) < 300:

                    return None

                print()
                print("==============================")
                print("PDF UTILIZADO")
                print("==============================")
                print(origem)
                print("CARACTERES:", len(texto))

                return {
                
                    "url": origem,
                
                    "tipo": "pdf",
                
                    "qualidade": "alta",
                
                    "caracteres": len(texto),
                
                    "texto": texto
                
                }

            except Exception as erro:

                print("ERRO PDF:")
                print(erro)

                if arquivo_temp:

                    try:

                        os.remove(

                            arquivo_temp.name

                        )

                    except:

                        pass

                return None

    # ========================================================
    # 02. URL JÁ É PDF
    # ========================================================

        if eh_pdf:

            return extrair_pdf(

                resposta.content,
                url

            )

    # ========================================================
    # 03. HTML
    # ========================================================

        html = resposta.text

        soup = BeautifulSoup(

            html,
            "html.parser"

        )

    # ========================================================
    # 04. PROCURAR PDF TÉCNICO
    # ========================================================

        palavras_pdf = [

            "pdf",
            "catalog",
            "catalogo",
            "catálogo",
            "datasheet",
            "manual",
            "brochure",
            "technical",
            "specification",
            "especificacao",
            "especificação",
            "download"

        ]

        pdfs = []

        from urllib.parse import urljoin

        for link in soup.find_all(

            "a",
            href=True

        ):

            href = link["href"]

            texto_link = link.get_text(

                " ",
                strip=True

            ).lower()

            href_lower = href.lower()

            if (

                ".pdf" in href_lower

                or

                any(

                    palavra in href_lower

                    for palavra in palavras_pdf

                )

                or

                any(

                    palavra in texto_link

                    for palavra in palavras_pdf

                )

            ):

                pdfs.append(

                    urljoin(

                        url,
                        href

                    )

                )

        # remover duplicados

        pdfs = list(

            dict.fromkeys(pdfs)

        )

    # ========================================================
    # 05. TENTAR TODOS OS PDFs
    # ========================================================

        for pdf in pdfs:

            print()
            print("==============================")
            print("TESTANDO PDF ENCONTRADO")
            print("==============================")
            print(pdf)

            try:

                r = requests.get(

                    pdf,
                    headers=headers,
                    timeout=30

                )

                if r.status_code != 200:

                    continue

                resultado = extrair_pdf(

                    r.content,
                    pdf

                )

                if resultado:

                    return resultado

            except:

                pass

    # ========================================================
    # 06. REMOVER LIXO HTML
    # ========================================================

        for tag in soup([

            "script",
            "style",
            "noscript",
            "svg",
            "iframe",
            "header",
            "footer",
            "nav",
            "form"

        ]):

            tag.decompose()

        texto = soup.get_text(

            separator=" ",
            strip=True

        )

        texto = limpar_texto_coletado(

            texto

        )

        if len(texto) < 300:

            return None

        return {
        
            "url": url,
        
            "tipo": "html",
        
            "qualidade": "normal",
        
            "caracteres": len(texto),
        
            "texto": texto
        
        }

    except requests.exceptions.Timeout:

        print("TIMEOUT:")
        print(url)

        return None

    except requests.exceptions.RequestException as erro:

        print("FALHA ACESSO:")
        print(url)

        print(erro)

        return None

    except Exception as erro:

        print("ERRO COLETA:")
        print(url)

        print(erro)

        return None
        

# ============================================================
# LIMPAR REFERÊNCIAS COMERCIAIS
# ============================================================

def limpar_referencias(
    texto,
    tema="",
    tipo="html"
):


    if not texto:

        return ""



    tema_normalizado = normalizar_texto(
        tema
    )



    # ========================================================
    # 01. DEFINIR NÍVEL DE LIMPEZA
    # ========================================================

    eh_pdf = (

        tipo.lower() == "pdf"

    )



    limpeza_forte = not eh_pdf



    # =====================================
    # LIMPEZA ESPECÍFICA
    # PROTEÇÃO CONTRA INCÊNDIO
    # =====================================

    eh_construcao = any(

        termo in tema_normalizado

        for termo in [

            "firestop",
            "selagem",
            "corta fogo",
            "passagem corta fogo",
            "protecao passiva"

        ]

    )



    if eh_construcao and limpeza_forte:


        padroes = [

            r"\bCP\s*\d+\b",

            r"\bCFS\s*\d+\b",

            r"\bFFC\b",

            r"\bCKC\b",

            r"\bFS-?\d+\b",


            r"\bHilti\b",

            r"\bFirestop\b",

            r"\bPromat\b",

            r"\b3M\b",

            r"\bSika\b",

            r"\bFischer\b",

            r"\bRockwool\b",


            r"\bClasse\s+Ultimate\b",

            r"\bUltimate\b"

        ]



        for padrao in padroes:


            texto = re.sub(

                padrao,

                "",

                texto,

                flags=re.IGNORECASE

            )



    # ========================================================
    # 02. LIMPEZA COMERCIAL GERAL
    # ========================================================

    remover = [

        "entre em contato",

        "solicite orçamento",

        "fale conosco",

        "peça sua cotação",

        "comprar agora",

        "consulte disponibilidade",

        "melhor preço",

        "oferta exclusiva"

    ]



    for item in remover:


        texto = re.sub(

            item,

            "",

            texto,

            flags=re.IGNORECASE

        )



    # =====================================
    # CAMPOS COMERCIAIS
    # SOMENTE HTML
    # =====================================

    if limpeza_forte:


        texto = re.sub(

            r"(part number|codigo interno|código interno)\s*[:\-]?\s*\w+",

            "",

            texto,

            flags=re.IGNORECASE

        )



    # ========================================================
    # 03. REMOVER LINKS
    # ========================================================

    texto = re.sub(

        r"https?://\S+",

        "",

        texto

    )



    texto = re.sub(

        r"www\.\S+",

        "",

        texto

    )



    # ========================================================
    # 04. NORMALIZAR ESPAÇOS
    # ========================================================

    texto = re.sub(

        r"\s+",

        " ",

        texto

    )



    return texto.strip()



# ========================================================
# SALVAR BRUTO
# ========================================================

def salvar_bruto(
    tema,
    paginas
):

    # ============================================================
    # SALVAR / ATUALIZAR PATRIMÔNIO BRUTO
    # ============================================================
    #
    # REGRAS:
    #
    # 1. URL existente:
    #    Atualiza a fonte existente.
    #
    # 2. URL nova:
    #    Adiciona ao patrimônio.
    #
    # 3. URL antiga que não veio nesta coleta:
    #    Permanece preservada.
    #
    # 4. Nunca substituir o patrimônio inteiro
    #    pela pesquisa atual.
    #
    # 5. Se o patrimônio existente não puder ser lido:
    #    NÃO sobrescrever o arquivo.
    #
    # 6. Metadados existentes são preservados
    #    quando a nova coleta não os substituir.
    #
    # ============================================================

    try:

        PASTA_DADOS.mkdir(
            parents=True,
            exist_ok=True
        )

        # ========================================================
        # 01. ARQUIVO DO TEMA
        # ========================================================

        nome_arquivo = normalizar_nome_arquivo(
            tema
        )

        arquivo = (
            PASTA_DADOS
            / f"{nome_arquivo}.json"
        )

        # ========================================================
        # 02. CARREGAR PATRIMÔNIO EXISTENTE
        # ========================================================

        fontes_existentes = []

        if arquivo.exists():

            try:

                with open(
                    arquivo,
                    "r",
                    encoding="utf-8"
                ) as f:

                    dados_existentes = json.load(f)

                if isinstance(
                    dados_existentes,
                    dict
                ):

                    fontes_existentes = (
                        dados_existentes.get(
                            "fontes",
                            []
                        )
                    )

                    if not isinstance(
                        fontes_existentes,
                        list
                    ):

                        print()
                        print(
                            "ERRO: CAMPO 'fontes' INVÁLIDO NO PATRIMÔNIO."
                        )

                        print(
                            "PATRIMÔNIO NÃO SERÁ SOBRESCRITO."
                        )

                        return {}

                else:

                    print()
                    print(
                        "ERRO: JSON DO PATRIMÔNIO NÃO É UM OBJETO VÁLIDO."
                    )

                    print(
                        "PATRIMÔNIO NÃO SERÁ SOBRESCRITO."
                    )

                    return {}

            except Exception as erro:

                print()
                print(
                    "ERRO AO CARREGAR PATRIMÔNIO EXISTENTE:"
                )

                print(erro)

                print(
                    "PATRIMÔNIO NÃO SERÁ SOBRESCRITO."
                )

                return {}

        # ========================================================
        # 03. INDEXAR PATRIMÔNIO POR URL
        # ========================================================

        patrimonio = {}

        ordem_urls = []

        for fonte in fontes_existentes:

            if not isinstance(
                fonte,
                dict
            ):
                continue

            url = str(
                fonte.get(
                    "url",
                    ""
                )
                or ""
            ).strip()

            if not url:
                continue

            if url not in patrimonio:

                ordem_urls.append(
                    url
                )

            patrimonio[url] = fonte

        # ========================================================
        # 04. PROCESSAR NOVAS FONTES
        # ========================================================

        paginas_validas = []

        if isinstance(
            paginas,
            list
        ):

            for pagina in paginas:

                if not isinstance(
                    pagina,
                    dict
                ):
                    continue

                url = str(
                    pagina.get(
                        "url",
                        ""
                    )
                    or ""
                ).strip()

                texto = str(
                    pagina.get(
                        "texto",
                        ""
                    )
                    or ""
                ).strip()

                tipo = str(
                    pagina.get(
                        "tipo",
                        "html"
                    )
                    or "html"
                ).strip().lower()

                if not url:
                    continue

                # ------------------------------------------------
                # VALIDAR TAMANHO MÍNIMO
                # ------------------------------------------------

                if tipo == "pdf":

                    tamanho_minimo = 150

                else:

                    tamanho_minimo = 200

                if len(texto) < tamanho_minimo:

                    continue

                # ------------------------------------------------
                # MONTAR NOVA FONTE
                # ------------------------------------------------

                fonte = {
                    "url": url,
                    "tipo": tipo,
                    "texto": texto
                }
                
                identidade_fonte = pagina.get(
                    "identidade_fonte"
                )
                
                if isinstance(identidade_fonte, dict):
                    fonte["identidade_fonte"] = identidade_fonte
                    
                # ------------------------------------------------
                # PRESERVAR METADADOS DA FONTE EXISTENTE
                # ------------------------------------------------
                #
                # Se a mesma URL já existe no patrimônio,
                # preserva todos os campos antigos que
                # não vierem na nova coleta.
                #
                # Isso protege:
                #
                # - status
                # - motivo
                # - outros metadados futuros
                #
                # ------------------------------------------------

                fonte_anterior = patrimonio.get(
                    url
                )

                if isinstance(
                    fonte_anterior,
                    dict
                ):

                    for chave, valor in fonte_anterior.items():

                        if chave not in fonte:

                            fonte[chave] = valor

                # ------------------------------------------------
                # NOVA URL
                # ------------------------------------------------

                if url not in patrimonio:

                    ordem_urls.append(
                        url
                    )

                # ------------------------------------------------
                # REGISTRAR FONTE VÁLIDA
                # ------------------------------------------------

                paginas_validas.append(
                    fonte
                )

                # ------------------------------------------------
                # URL EXISTENTE OU NOVA:
                # ATUALIZAR A FONTE
                # ------------------------------------------------

                patrimonio[url] = fonte

        # ========================================================
        # 05. RECONSTRUIR PATRIMÔNIO
        # ========================================================

        fontes_finais = []

        for url in ordem_urls:

            fonte = patrimonio.get(
                url
            )

            if not isinstance(
                fonte,
                dict
            ):
                continue

            texto = str(
                fonte.get(
                    "texto",
                    ""
                )
                or ""
            ).strip()

            if not texto:

                continue

            fontes_finais.append(
                fonte
            )

        # ========================================================
        # 06. PDFS PRIMEIRO
        # ========================================================

        fontes_finais.sort(
            key=lambda item: (
                0
                if str(
                    item.get(
                        "tipo",
                        ""
                    )
                ).lower() == "pdf"
                else 1
            )
        )

        # ========================================================
        # 07. ESTATÍSTICAS
        # ========================================================

        total_pdf = 0
        total_html = 0

        for fonte in fontes_finais:

            tipo = str(
                fonte.get(
                    "tipo",
                    "html"
                )
                or "html"
            ).lower()

            if tipo == "pdf":

                total_pdf += 1

            else:

                total_html += 1

        # ========================================================
        # 08. MONTAR JSON FINAL
        # ========================================================

        dados = {

            "tema": tema,

            "data_coleta": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

            "fontes": fontes_finais,

            "quantidade": len(
                fontes_finais
            ),

            "estatisticas": {

                "pdfs": total_pdf,

                "html": total_html

            }

        }

        # ========================================================
        # 09. GRAVAR PATRIMÔNIO
        # ========================================================

        with open(
            arquivo,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                dados,
                f,
                ensure_ascii=False,
                indent=4
            )

        # ========================================================
        # 10. LOG
        # ========================================================

        print()
        print("==============================")
        print("PATRIMÔNIO BRUTO ATUALIZADO")
        print("==============================")

        print(
            "TEMA:",
            tema
        )

        print(
            "FONTES ANTERIORES:",
            len(fontes_existentes)
        )

        print(
            "FONTES RECEBIDAS:",
            len(paginas_validas)
        )

        print(
            "FONTES FINAIS:",
            len(fontes_finais)
        )

        print(
            "PDFS:",
            total_pdf
        )

        print(
            "HTML:",
            total_html
        )

        print(
            "ARQUIVO:",
            arquivo
        )

        return dados

    except Exception as erro:

        print()
        print("==============================")
        print("ERRO AO SALVAR DADOS BRUTOS")
        print("==============================")

        print(erro)

        return {}



        

# ============================================================
# REMOVER ACENTOS PARA NOMES DE ARQUIVO
# ============================================================

def remover_acentos(texto):

    texto = unicodedata.normalize(
        "NFD",
        texto
    )


    texto = "".join(

        c for c in texto

        if unicodedata.category(c) != "Mn"

    )


    return texto




# ============================================================
# NORMALIZAR NOME DE ARQUIVO
# ============================================================

def normalizar_nome_arquivo(texto):

    texto = remover_acentos(

        texto.lower()

    )


    texto = re.sub(

        r"[^a-z0-9]+",

        "_",

        texto

    )


    return texto.strip("_")




# ============================================================
# CARREGAR DADOS BRUTOS EXISTENTES
# ============================================================

def carregar_bruto(tema):


    nome = normalizar_nome_arquivo(

        tema

    )


    arquivo = Path(

        PASTA_DADOS

    ) / f"{nome}.json"



    if not arquivo.exists():


        print()

        print("==============================")

        print("BRUTO NÃO ENCONTRADO")

        print("==============================")

        print(arquivo)


        return []



    try:


        with open(

            arquivo,

            "r",

            encoding="utf-8"

        ) as f:


            dados = json.load(f)



        fontes = dados.get(

            "fontes",

            []

        )



        textos = []



        for item in fontes:


            if not isinstance(

                item,

                dict

            ):

                continue



            texto = item.get(

                "texto",

                ""

            )


            tipo = item.get(

                "tipo",

                "html"

            )



            tamanho_minimo = 200



            # PDF pode ter tabelas e pouco texto

            if tipo == "pdf":

                tamanho_minimo = 150



            if len(texto) < tamanho_minimo:

                continue



            textos.append({

                "texto": texto,
            
                "url": item.get(
                    "url",
                    ""
                ),
            
                "tipo": tipo,
            
                "identidade_fonte":
                    item.get(
                        "identidade_fonte",
                        {}
                    )
            
            })



    # ========================================================
    # 01. PDFS PRIMEIRO
    # ========================================================

        textos.sort(

            key=lambda x:

            0 if x.get("tipo") == "pdf" else 1

        )



        print()

        print("==============================")

        print("BRUTO ENCONTRADO")

        print("==============================")

        print(

            "ARQUIVO:",

            arquivo

        )

        print(

            "FONTES JSON:",

            len(fontes)

        )

        print(

            "TEXTOS CARREGADOS:",

            len(textos)

        )



        pdfs = sum(

            1

            for t in textos

            if t.get("tipo") == "pdf"

        )


        print(

            "PDFS:",

            pdfs

        )



        return textos



    except Exception as erro:


        print()

        print(

            "ERRO AO CARREGAR BRUTO:",

            erro

        )


        return []

    # ========================================================
    # 02. REGRAS PARA CONTEÚDO
    # ========================================================
{
    "versao": "4.0",

    "identidade": {

        "metodo": "MEAD",

        "objetivo":
        "Criar páginas estratégicas com autoridade, contexto técnico, diferenciação, profundidade semântica e narrativa humana."
    },


    "preparacao": {


        "palavra_chave": {


            "regra":

            "A palavra-chave principal informada pelo usuário é o protagonista absoluto da narrativa.",


            "prioridade":

            [

                "produto",

                "serviço",

                "fabricante",

                "distribuidor",

                "fornecedor"

            ],


            "proibido":

            [

                "substituir o protagonista",

                "transformar aplicação em produto",

                "transformar benefício em produto",

                "transformar tecnologia em produto",

                "transformar característica em produto"

            ]

        },



        "classificacao_intencao": {


            "produto":

            {

                "regra":

                "O protagonista permanece exatamente como o produto informado."

            },


            "servico":

            {

                "regra":

                "O protagonista permanece exatamente como o serviço informado."

            },


            "fabricante":

            {

                "regra":

                "O protagonista é o fabricante. Desenvolver capacidade técnica, engenharia, processos e estrutura."

            },


            "distribuidor":

            {

                "regra":

                "O protagonista é o distribuidor. Desenvolver fornecimento, atendimento e suporte."

            },


            "fornecedor":

            {

                "regra":

                "O protagonista é o fornecedor. Desenvolver disponibilidade, seleção e relacionamento."

            }

        }

    },



    "mapa_mead": {


        "protagonista":

        "Elemento principal exatamente alinhado ao tema informado.",


        "cenario":

        "Ambiente operacional completo onde o protagonista atua.",


        "problema":

        "Necessidade real que o protagonista resolve.",


        "solucao":

        "Como o protagonista atende esta necessidade.",


        "coadjuvantes":

        "Componentes, processos, materiais, tecnologias e conceitos que ajudam a explicar.",


        "aplicacoes":

        "Locais, equipamentos, sistemas ou situações reais de utilização.",


        "publico_alvo":

        "Quem compra, especifica, utiliza ou aplica.",


        "diferenciais":

        "Características técnicas reais que influenciam a escolha.",


        "evidencias_tecnicas":

        "Informações encontradas nas fontes que comprovam características, aplicações ou funcionamento.",


        "competencia_principal":

        "Maior demonstração de domínio técnico."

    },



    "densidade_estrategica": {


        "regra":

        "Cada campo deve possuir profundidade suficiente para orientar a criação do conteúdo.",


        "priorizar":

        [

            "contexto",

            "finalidade",

            "funcionamento",

            "consequências",

            "benefícios técnicos",

            "limitações",

            "relação entre elementos",

            "evidências técnicas"

        ]

    },



    "fontes_pesquisa": {


        "prioridade_fontes":

        [

            "datasheet técnico",

            "catálogo técnico PDF",

            "manual técnico PDF",

            "ficha técnica PDF",

            "documentação técnica fabricante",

            "artigos técnicos especializados",

            "páginas comerciais"

        ],



        "hierarquia_confiabilidade": {


            "nivel_1":

            [

                "PDF técnico",

                "catálogo fabricante",

                "manual engenharia"

            ],


            "nivel_2":

            [

                "site fabricante",

                "distribuidor especializado",

                "empresa técnica"

            ],


            "nivel_3":

            [

                "blogs técnicos",

                "portais industriais"

            ]

        },



        "usar_para":

        [

            "conhecimento técnico",

            "aplicações",

            "processos",

            "entendimento do mercado",

            "identificação de componentes",

            "variações do produto",

            "materiais utilizados"

        ],



        "nunca_usar_para":

        [

            "trocar protagonista",

            "alterar intenção comercial",

            "mudar segmento",

            "copiar textos",

            "inventar informações",

            "criar especificações sem fonte"

        ]

    },



    "prioridade_documentos": {


        "regra":

        "Documentos técnicos possuem prioridade sobre páginas comerciais quando existirem informações suficientes.",


        "preferir":

        [

            "datasheet",

            "manual",

            "catálogo",

            "ficha técnica",

            "desenho técnico"

        ],


        "extrair":

        [

            "materiais",

            "componentes",

            "aplicações",

            "modelos",

            "processos",

            "condições de uso"

        ]

    },



    "controle_tecnico": {


        "nunca_criar":

        [

            "certificações",

            "normas",

            "números",

            "percentuais",

            "testes",

            "garantias",

            "aprovações",

            "resultados quantitativos"

        ]

    },



    "validacao_informacao": {


        "regra":

        "Toda informação técnica deve possuir origem nas fontes coletadas ou ser uma conclusão lógica baseada nelas.",


        "bloquear":

        [

            "números técnicos sem fonte",

            "capacidade sem documento",

            "temperatura sem referência",

            "pressão sem referência",

            "normas sem comprovação",

            "certificações inexistentes"

        ]

    },



    "classificacao_fontes": {


        "pdf_tecnico":

        {

            "peso": 10

        },


        "fabricante":

        {

            "peso": 9

        },


        "distribuidor_tecnico":

        {

            "peso": 7

        },


        "artigo_tecnico":

        {

            "peso": 6

        },


        "pagina_comercial":

        {

            "peso": 4

        }

    },



    "gambiarra_narrativa": {


        "objetivo":

        "Criar textos humanos evitando repetição estrutural.",


        "permitir_inicio":

        [

            "empresa",

            "engenharia",

            "cliente",

            "processo",

            "cenário",

            "necessidade",

            "aplicação",

            "experiência",

            "benefício"

        ],


        "evitar":

        [

            "todos os parágrafos começarem pela palavra-chave",

            "todos os parágrafos começarem igual",

            "excesso de nossa empresa",

            "excesso de nossa equipe"

        ]

    },



    "ativos_narrativos": {


        "lista":

        [

            "experiencia",

            "conhecimento_tecnico",

            "engenharia",

            "qualidade",

            "seguranca",

            "suporte_tecnico",

            "atendimento_especializado",

            "confiabilidade",

            "responsabilidade"

        ],


        "regra":

        "Utilizar somente quando fizer sentido sem inventar histórico."

    },



    "regras_seo": {


        "evitar":

        [

            "texto genérico",

            "enchimento",

            "repetição",

            "promessas sem comprovação",

            "frases artificiais"

        ],


        "priorizar":

        [

            "autoridade",

            "clareza",

            "profundidade técnica",

            "SEO semântico",

            "naturalidade"

        ]

    },



    "auditoria_seven": {


        "executar":

        True,


        "verificar":

        [

            "protagonista correto",

            "segmento correto",

            "intenção comercial correta",

            "ausência de invenções",

            "cenário profundo",

            "problema real",

            "solução coerente",

            "diversidade narrativa",

            "repetição de tese",

            "repetição sintática",

            "naturalidade humana",

            "qualidade das fontes",

            "prioridade documental"

        ]

    }

}



# ============================================================
# AVALIAR PATRIMÔNIO EXISTENTE
# ============================================================
#
# OBJETIVO:
#
# 1. Usar primeiro o patrimônio bruto já existente.
# 2. Não usar conteúdo editorial de páginas anteriores.
# 3. Aproveitar dados brutos de temas relacionados da mesma família.
# 4. Só considerar pesquisa externa quando o patrimônio for
#    realmente insuficiente.
#
# IMPORTANTE:
# Esta função NÃO altera o patrimônio.
# Esta função NÃO grava dados.
# Esta função NÃO chama Ollama.
# Esta função NÃO chama pesquisar_completo().
#
# ============================================================

def avaliar_patrimonio_existente(
    tema,
    minimo_fontes=5,
    minimo_caracteres=12000
):

    print()
    print("==================================================")
    print("AVALIAÇÃO DO PATRIMÔNIO EXISTENTE")
    print("==================================================")
    print("TEMA:", tema)

    patrimonio = []

    fontes_vistas = set()

    # ========================================================
    # FUNÇÃO INTERNA — ADICIONAR FONTE SEM DUPLICAR
    # ========================================================

    def adicionar_fontes(
        fontes,
        tema_origem
    ):

        adicionadas = 0

        if not isinstance(
            fontes,
            list
        ):
            return 0

        for item in fontes:

            if not isinstance(
                item,
                dict
            ):
                continue

            texto = str(
                item.get(
                    "texto",
                    ""
                )
                or ""
            ).strip()

            if len(texto) < 300:
                continue

            url = str(
                item.get(
                    "url",
                    ""
                )
                or ""
            ).strip()

            # ------------------------------------------------
            # IDENTIDADE DA FONTE
            # ------------------------------------------------

            if url:

                chave = (
                    "url:"
                    + url.strip().casefold()
                )

            else:

                try:

                    chave = (
                        "texto:"
                        + gerar_hash_trecho(
                            texto
                        )
                    )

                except Exception:

                    chave = (
                        "texto:"
                        + texto[:500].casefold()
                    )

            if chave in fontes_vistas:
                continue

            fontes_vistas.add(
                chave
            )

            patrimonio.append({

                "url":
                    url,

                "tipo":
                    str(
                        item.get(
                            "tipo",
                            "texto"
                        )
                        or "texto"
                    ),

                "texto":
                    texto,

                "tema_origem":
                    str(
                        tema_origem
                        or tema
                    ).strip()

            })

            adicionadas += 1

        return adicionadas

    # ========================================================
    # 01. CARREGAR PATRIMÔNIO DO PRÓPRIO TEMA
    # ========================================================

    fontes_tema = carregar_bruto(
        tema
    )

    adicionadas_tema = adicionar_fontes(
        fontes_tema,
        tema
    )

    print()
    print(
        "PATRIMÔNIO DO TEMA:",
        adicionadas_tema
    )

    # ========================================================
    # 02. CALCULAR COBERTURA INICIAL
    # ========================================================

    caracteres = sum(
        len(
            item.get(
                "texto",
                ""
            )
        )
        for item in patrimonio
    )

    print(
        "FONTES DISPONÍVEIS:",
        len(patrimonio)
    )

    print(
        "CARACTERES DISPONÍVEIS:",
        caracteres
    )

    # ========================================================
    # 03. VERIFICAR SE JÁ É SUFICIENTE
    # ========================================================

    patrimonio_suficiente = (
        len(patrimonio) >= minimo_fontes
        and
        caracteres >= minimo_caracteres
    )

    if patrimonio_suficiente:

        print()
        print(
            "🟢 PATRIMÔNIO SUFICIENTE"
        )

        print(
            "NÃO SERÁ FEITA PESQUISA EXTERNA."
        )

        return {

            "suficiente":
                True,

            "fontes":
                patrimonio,

            "quantidade_fontes":
                len(patrimonio),

            "caracteres":
                caracteres,

            "temas_relacionados":
                []

        }

    # ========================================================
    # 04. PROCURAR PATRIMÔNIO DE TEMAS RELACIONADOS
    # ========================================================
    #
    # IMPORTANTE:
    #
    # Aqui usamos SOMENTE dados-brutos.
    #
    # NÃO usamos:
    # - conteúdo editorial;
    # - páginas prontas;
    # - mapa MEAD de outro tema.
    #
    # ========================================================

    relacionados_validos = []

    try:

        relacionados = buscar_temas_relacionados(
            tema
        )

    except Exception as erro:

        print()
        print(
            "AVISO: não foi possível buscar "
            "temas relacionados."
        )

        print(
            repr(erro)
        )

        relacionados = []

    for relacionado in relacionados:

        if len(
            relacionados_validos
        ) >= 5:
            break

        if not isinstance(
            relacionado,
            str
        ):
            continue

        relacionado = relacionado.strip()

        if not relacionado:
            continue

        if normalizar_tema_chave(
            relacionado
        ) == normalizar_tema_chave(
            tema
        ):
            continue

        # ----------------------------------------------------
        # RESPEITAR A FAMÍLIA TEMÁTICA
        # ----------------------------------------------------

        try:

            pertence = pertence_ao_grupo_principal(
                tema,
                relacionado
            )

        except Exception:

            pertence = False

        if not pertence:
            continue

        bruto_relacionado = carregar_bruto(
            relacionado
        )

        quantidade_antes = len(
            patrimonio
        )

        adicionar_fontes(
            bruto_relacionado,
            relacionado
        )

        quantidade_depois = len(
            patrimonio
        )

        novas_fontes = (
            quantidade_depois
            - quantidade_antes
        )

        if novas_fontes > 0:

            relacionados_validos.append(
                relacionado
            )

            print()
            print(
                "🟢 PATRIMÔNIO RELACIONADO "
                "APROVEITADO:"
            )

            print(
                relacionado
            )

            print(
                "NOVAS FONTES:",
                novas_fontes
            )

        # ----------------------------------------------------
        # REAVALIAR APÓS CADA TEMA RELACIONADO
        # ----------------------------------------------------

        caracteres = sum(
            len(
                item.get(
                    "texto",
                    ""
                )
            )
            for item in patrimonio
        )

        patrimonio_suficiente = (
            len(patrimonio) >= minimo_fontes
            and
            caracteres >= minimo_caracteres
        )

        if patrimonio_suficiente:
            break

    # ========================================================
    # 05. RESULTADO FINAL
    # ========================================================

    caracteres = sum(
        len(
            item.get(
                "texto",
                ""
            )
        )
        for item in patrimonio
    )

    patrimonio_suficiente = (
        len(patrimonio) >= minimo_fontes
        and
        caracteres >= minimo_caracteres
    )

    print()
    print("==================================================")
    print("RESULTADO DA AVALIAÇÃO")
    print("==================================================")

    print(
        "FONTES:",
        len(patrimonio)
    )

    print(
        "CARACTERES:",
        caracteres
    )

    print(
        "TEMAS RELACIONADOS:",
        len(relacionados_validos)
    )

    print(
        "PATRIMÔNIO SUFICIENTE:",
        patrimonio_suficiente
    )

    if patrimonio_suficiente:

        print(
            "🟢 PESQUISA EXTERNA NÃO NECESSÁRIA."
        )

    else:

        print(
            "🟡 PATRIMÔNIO INSUFICIENTE."
        )

        print(
            "PESQUISA COMPLEMENTAR SERÁ PERMITIDA."
        )

    print(
        "=================================================="
    )

    return {

        "suficiente":
            patrimonio_suficiente,

        "fontes":
            patrimonio,

        "quantidade_fontes":
            len(patrimonio),

        "caracteres":
            caracteres,

        "temas_relacionados":
            relacionados_validos
    }
    
    
# ============================================================
# GERAR MAPA MEAD — PYTHON
# ============================================================
#
# O MAPA MEAD não é mais produzido pelo Ollama.
#
# O Python já possui:
# - tema;
# - tipo;
# - grupo;
# - regras MEAD;
# - fontes;
# - seleção editorial.
#
# O mapa passa a ser uma estrutura operacional local.
#
# IMPORTANTE:
# O mapa NÃO é fonte factual para a redação.
# Os fatos autorizados continuam sendo os fragmentos
# selecionados pelo Python.
#
# ============================================================

def gerar_mapa_mead(
    tema,
    textos
):

    tema = str(
        tema or ""
    ).strip()

    if not tema:

        return ""

    if not isinstance(
        textos,
        list
    ):

        textos = []

    print()
    print("==============================")
    print("GERANDO MAPA MEAD — PYTHON")
    print("==============================")

    print(
        "TEMA:",
        tema
    )

    print(
        "FONTES:",
        len(textos)
    )

    # --------------------------------------------------------
    # TIPO
    # --------------------------------------------------------

    tipo = identificar_tipo_tema(
        tema
    )

    # --------------------------------------------------------
    # GRUPO
    # --------------------------------------------------------

    grupo = identificar_grupo_tema(
        tema
    )

    # --------------------------------------------------------
    # CONTEXTO DOS BLOCOS
    # --------------------------------------------------------

    contexto_blocos = {}

    try:

        if isinstance(
            MEAD,
            dict
        ):

            contexto_blocos = MEAD.get(
                "contexto_blocos",
                {}
            )

    except Exception:

        contexto_blocos = {}

    if not isinstance(
        contexto_blocos,
        dict
    ):

        contexto_blocos = {}

    # --------------------------------------------------------
    # MONTAR OBJETO
    # --------------------------------------------------------

    mapa = {

        "tipo":
            "MAPA_MEAD_PYTHON",

        "tema":
            tema,

        "protagonista":
            tema,

        "tipo_pagina":
            tipo,

        "grupo":
            grupo,

        "intencao_comercial":
            (
                "Organizar informações técnicas "
                "relacionadas ao tema para uma página "
                "com intenção informacional e comercial."
            ),

        "regra_factual":
            (
                "Nenhuma informação factual deve ser "
                "criada a partir deste mapa. "
                "Os fatos autorizados são somente os "
                "presentes nos fragmentos selecionados "
                "pelo Python."
            ),

        "contexto_blocos":
            contexto_blocos,

        "fontes_disponiveis":
            len(textos),

        "lacunas":
            (
                "Informações não presentes nos fragmentos "
                "selecionados devem permanecer ausentes."
            )

    }

    resultado = json.dumps(
        mapa,
        ensure_ascii=False,
        indent=2
    )

    print()
    print("==============================")
    print("MAPA MEAD PYTHON GERADO")
    print("==============================")

    print(
        "CARACTERES:",
        len(resultado)
    )

    return resultado
    

# ============================================================
# VALIDAR E RECUPERAR MAPA MEAD
# ============================================================

def validar_e_recuperar_mapa_mead(
    tema,
    mapa_mead,
    textos
):

    # ========================================================
    # 01. NORMALIZAR MAPA RECEBIDO
    # ========================================================

    if isinstance(
        mapa_mead,
        dict
    ):

        mapa_mead = mapa_mead.get(
            "texto",
            ""
        )

    elif mapa_mead is not None:

        mapa_mead = str(
            mapa_mead
        )

    else:

        mapa_mead = ""


    mapa_mead = mapa_mead.strip()


    # ========================================================
    # 02. VALIDAR MAPA RECEBIDO
    # ========================================================

    if mapa_mead:

        print()
        print("==============================")
        print("VALIDANDO MAPA MEAD PARA CONTEÚDO")
        print("==============================")

        print(
            "TEMA:",
            tema
        )

        print(
            "TAMANHO:",
            len(mapa_mead)
        )


    # ========================================================
    # 03. VALIDAR PROTAGONISTA
    # ========================================================

        protagonista_valido = validar_protagonista_mead(
            mapa_mead,
            tema
        )


        print(
            "PROTAGONISTA:",
            protagonista_valido
        )


        if not protagonista_valido:

            print()
            print("==============================")
            print("MAPA MEAD REJEITADO")
            print("==============================")
            print(
                "MOTIVO: PROTAGONISTA INCOMPATÍVEL"
            )

        else:

    # ========================================================
    # 04. VALIDAR CONTEXTO TÉCNICO
    # ========================================================

            contexto_valido = validar_contexto_tecnico_mead(
                mapa_mead,
                tema
            )


            print(
                "CONTEXTO TÉCNICO:",
                contexto_valido
            )


            if contexto_valido:

                print()
                print("==============================")
                print("MAPA MEAD VALIDADO")
                print("==============================")

                return mapa_mead


    # ========================================================
    # 05. TENTAR RECUPERAR DO BANCO
    # ========================================================

    print()
    print("==============================")
    print("TENTANDO RECUPERAR MAPA MEAD")
    print("==============================")

    mapa_recuperado = obter_mapa_mead_tema(
        tema
    )


    if mapa_recuperado:

        mapa_recuperado = str(
            mapa_recuperado
        ).strip()


    # ========================================================
    # 06. VALIDAR PROTAGONISTA RECUPERADO
    # ========================================================

        protagonista_valido = validar_protagonista_mead(
            mapa_recuperado,
            tema
        )


        print(
            "PROTAGONISTA RECUPERADO:",
            protagonista_valido
        )


        if protagonista_valido:

    # ========================================================
    # 07. VALIDAR CONTEXTO RECUPERADO
    # ========================================================

            contexto_valido = validar_contexto_tecnico_mead(
                mapa_recuperado,
                tema
            )


            print(
                "CONTEXTO RECUPERADO:",
                contexto_valido
            )


            if contexto_valido:

                print()
                print("==============================")
                print("MAPA MEAD RECUPERADO")
                print("==============================")

                print(
                    "TAMANHO:",
                    len(mapa_recuperado)
                )

                return mapa_recuperado


    # ========================================================
    # 08. MAPA NÃO DISPONÍVEL
    # ========================================================

    print()
    print("==============================")
    print("MAPA MEAD NÃO DISPONÍVEL")
    print("==============================")

    print(
        "TEMA:",
        tema
    )

    return None
    
    # ========================================================
    # 09. CONFIGURAÇÃO DE FOCOS EDITORIAIS
    # ========================================================

FOCOS_EDITORIAIS = {

    "apresentacao_contexto": {
        "nome": "Apresentação e contexto",
        "ativo": True
    },

    "funcionamento": {
        "nome": "Funcionamento e domínio técnico",
        "ativo": True
    },

    "aplicacoes": {
        "nome": "Aplicações e necessidades",
        "ativo": True
    },

    "criterios": {
        "nome": "Critérios, diferenciação e confiança",
        "ativo": True
    },

    "comercial": {
        "nome": "Decisão e contexto comercial",
        "ativo": True
    },

    "informacao_tecnica": {
        "nome": "Informação técnica",
        "ativo": True
    },

    "instalacao_manutencao": {
        "nome": "Instalação e manutenção",
        "ativo": False
    },

    "beneficios": {
        "nome": "Benefícios",
        "ativo": True
    },

    "problemas_necessidades": {
        "nome": "Problemas e necessidades",
        "ativo": False
    },

    "processo_execucao": {
        "nome": "Processo de execução",
        "ativo": False
    },

    "experiencia_autoridade": {
        "nome": "Experiência e autoridade",
        "ativo": False
    },

    "seguranca": {
        "nome": "Segurança",
        "ativo": False
    },

    "atendimento_suporte": {
        "nome": "Atendimento e suporte",
        "ativo": False
    },

    "personalizacao_projeto": {
        "nome": "Personalização e projeto",
        "ativo": False
    },

    "pos_venda": {
        "nome": "Pós-venda",
        "ativo": False
    }

}



# ============================================================
# OBTER FOCOS EDITORIAIS
# ============================================================

def obter_focos_editoriais():

    focos = {}

    for chave, dados in FOCOS_EDITORIAIS.items():

        focos[chave] = bool(
            dados.get(
                "ativo",
                False
            )
        )

    return focos



# ============================================================
# FORMATAR FOCOS EDITORIAIS
# ============================================================

def formatar_focos_editoriais(
    focos=None
):

    # --------------------------------------------------------
    # USAR CONFIGURAÇÃO PADRÃO
    # --------------------------------------------------------

    if focos is None:

        focos = obter_focos_editoriais()


    linhas = []


    linhas.append(
        "FOCOS EDITORIAIS:"
    )

    linhas.append("")


    for chave, dados in FOCOS_EDITORIAIS.items():

        selecionado = bool(
            focos.get(
                chave,
                False
            )
        )


        simbolo = (
            "☑"
            if selecionado
            else "☐"
        )


        nome = dados.get(
            "nome",
            chave
        )


        linhas.append(
            f"{simbolo} {nome}"
        )


    return "\n".join(
        linhas
    )



# ============================================================
# DEBUG — MOSTRAR CONFIGURAÇÃO DOS FOCOS
# ============================================================

def mostrar_focos_editoriais():

    focos = obter_focos_editoriais()


    print()
    print(
        "========================================"
    )
    print(
        "FOCOS EDITORIAIS"
    )
    print(
        "========================================"
    )


    print(
        formatar_focos_editoriais(
            focos
        )
    )


    print(
        "========================================"
    )


    return focos    



# ============================================================
# CONTROLAR CONTEXTO PARA IA
# ============================================================

def controlar_contexto_ia(
    textos,
    limite_total=16000
):

    contexto = ""

    if not textos:
        return contexto


    # ========================================================
    # 01. PERCORRER FONTES
    # ========================================================

    for item in textos:

        if len(contexto) >= limite_total:
            break


        # ---------------------------------
        # TEXTO DA FONTE
        # ---------------------------------

        if isinstance(item, dict):

            texto = item.get(
                "texto",
                ""
            )

        elif isinstance(item, str):

            texto = item

        else:

            continue


        if not texto:
            continue


        # ---------------------------------
        # LIMITE RESTANTE
        # ---------------------------------

        restante = (
            limite_total
            - len(contexto)
        )


        if restante <= 0:
            break


        # ---------------------------------
        # ADICIONAR TEXTO
        # ---------------------------------

        trecho = texto[:restante]


        contexto += (
            "\n\n"
            + trecho
        )


    # ========================================================
    # 02. NORMALIZAÇÃO
    # ========================================================

    contexto = contexto.strip()


    # ========================================================
    # 03. DEBUG
    # ========================================================

    print()
    print("==============================")
    print("CONTROLE CONTEXTO IA")
    print("==============================")

    print(
        "LIMITE:",
        limite_total
    )

    print(
        "CARACTERES ENVIADOS:",
        len(contexto)
    )

    print("==============================")


    return contexto
    

# ============================================================
# CONTROLAR ABERTURAS DO PROTAGONISTA
# ============================================================

def controlar_aberturas_protagonista(texto, tema, limite=3):



    print()
    print("=" * 60)
    print("CONTROLE DE ABERTURAS DO PROTAGONISTA")
    print("=" * 60)

    print("TEMA:", tema)
    print("LIMITE DE ABERTURAS DIRETAS:", limite)

    if not texto or not texto.strip():

        print("TEXTO VAZIO")
        print("CONTROLE NÃO EXECUTADO")

        return texto

    # --------------------------------------------------------
    # NORMALIZAÇÃO
    # --------------------------------------------------------

    def normalizar(texto_local):

        texto_local = texto_local.lower().strip()

        texto_local = unicodedata.normalize(
            "NFD",
            texto_local
        )

        texto_local = "".join(
            caractere
            for caractere in texto_local
            if unicodedata.category(caractere) != "Mn"
        )

        texto_local = re.sub(
            r"\s+",
            " ",
            texto_local
        )

        return texto_local

    tema_normalizado = normalizar(tema)

    # --------------------------------------------------------
    # SEPARAR TEXTO DOS MARCADORES
    # --------------------------------------------------------

    linhas = texto.splitlines()

    paragrafos = []

    bloco_atual = None

    for linha in linhas:

        linha_limpa = linha.strip()

        if not linha_limpa:
            continue

        # ----------------------------------------------------
        # IGNORA MARCADORES DE BLOCO
        # ----------------------------------------------------

        if re.match(
            r"^\*\*BLOCO\s+\d+\*\*$",
            linha_limpa,
            re.IGNORECASE
        ):

            bloco_atual = linha_limpa

            continue

        # ----------------------------------------------------
        # IGNORA LINHAS DE SEGMENTOS
        # ----------------------------------------------------

        if linha_limpa.startswith("- "):
            continue

        # ----------------------------------------------------
        # IGNORA TÍTULOS / MARCADORES
        # ----------------------------------------------------

        if (
            linha_limpa.startswith("**")
            and linha_limpa.endswith("**")
        ):
            continue

        # ----------------------------------------------------
        # PARÁGRAFO
        # ----------------------------------------------------

        paragrafos.append({
            "texto": linha_limpa,
            "bloco": bloco_atual
        })

    print()
    print("PARÁGRAFOS ANALISADOS:", len(paragrafos))

    # --------------------------------------------------------
    # IDENTIFICAR ABERTURAS DIRETAS
    # --------------------------------------------------------

    ocorrencias = []

    for indice, item in enumerate(paragrafos):

        texto_paragrafo = item["texto"]

        inicio_normalizado = normalizar(
            texto_paragrafo
        )

        # Aceita:
        #
        # A bomba centrifuga...
        # A bomba centrífuga...
        #
        # Também permite artigo masculino/feminino
        # conforme o tema.

        if (
            inicio_normalizado.startswith(
                tema_normalizado
            )
            or inicio_normalizado.startswith(
                "a " + tema_normalizado
            )
            or inicio_normalizado.startswith(
                "o " + tema_normalizado
            )
        ):

            ocorrencias.append({
                "indice": indice,
                "bloco": item["bloco"],
                "texto": texto_paragrafo
            })

    # --------------------------------------------------------
    # PRINT DAS OCORRÊNCIAS
    # --------------------------------------------------------

    print()
    print("ABERTURAS DIRETAS ENCONTRADAS:", len(ocorrencias))

    if ocorrencias:

        for numero, ocorrencia in enumerate(
            ocorrencias,
            start=1
        ):

            print()
            print(
                f"ABERTURA {numero}"
            )

            print(
                "PARÁGRAFO:",
                ocorrencia["indice"] + 1
            )

            print(
                "BLOCO:",
                ocorrencia["bloco"]
            )

            print(
                "TEXTO:",
                ocorrencia["texto"][:180]
            )

    else:

        print(
            "NENHUMA ABERTURA DIRETA ENCONTRADA"
        )

    # --------------------------------------------------------
    # VERIFICAR LIMITE
    # --------------------------------------------------------

    if len(ocorrencias) <= limite:

        print()
        print("STATUS: DENTRO DO LIMITE")
        print(
            f"ABERTURAS: {len(ocorrencias)} / {limite}"
        )

        print(
            "NENHUMA CORREÇÃO NECESSÁRIA"
        )

        print("=" * 60)

        return texto

    # --------------------------------------------------------
    # EXISTE EXCESSO
    # --------------------------------------------------------

    print()
    print("STATUS: EXCESSO DE ABERTURAS")
    print(
        f"ABERTURAS ENCONTRADAS: {len(ocorrencias)}"
    )
    print(
        f"LIMITE PERMITIDO: {limite}"
    )

    # --------------------------------------------------------
    # SELECIONAR SOMENTE OS EXCEDENTES
    # --------------------------------------------------------

    excedentes = ocorrencias[limite:]

    print()
    print(
        "PARÁGRAFOS QUE SERÃO CORRIGIDOS:",
        len(excedentes)
    )

    for numero, ocorrencia in enumerate(
        excedentes,
        start=1
    ):

        print()
        print(
            f"EXCEDENTE {numero}"
        )

        print(
            "PARÁGRAFO:",
            ocorrencia["indice"] + 1
        )

        print(
            "BLOCO:",
            ocorrencia["bloco"]
        )

        print(
            "TEXTO:",
            ocorrencia["texto"][:200]
        )

    # --------------------------------------------------------
    # MONTAR PEDIDO DE CORREÇÃO
    # --------------------------------------------------------

    instrucoes = []

    instrucoes.append(
        "VARIAÇÃO EDITORIAL DE ABERTURAS"
    )

    instrucoes.append(
        f"Palavra-chave: {tema}"
    )

    instrucoes.append(
        "O texto já está pronto e não deve ser reescrito integralmente."
    )

    instrucoes.append(
        "Corrija SOMENTE os parágrafos fornecidos."
    )

    instrucoes.append(
        f"Não inicie o parágrafo com '{tema}'."
    )

    instrucoes.append(
        "A palavra-chave pode continuar aparecendo naturalmente dentro do parágrafo."
    )

    instrucoes.append(
        "Preserve integralmente o significado técnico."
    )

    instrucoes.append(
        "Não invente informações."
    )

    instrucoes.append(
        "Não transforme o texto em propaganda."
    )

    instrucoes.append(
        "Use uma abertura natural e diferente."
    )

    instrucoes.append(
        "Pode iniciar pelo cenário, operação, necessidade, aplicação, engenharia, processo, benefício, cliente ou outro elemento coerente."
    )

    instrucoes.append(
        "Não utilize outra fórmula repetitiva."
    )

    instrucoes.append(
        "Retorne somente os parágrafos corrigidos, na mesma ordem."
    )

    instrucoes.append(
        "Separe cada parágrafo corrigido por uma linha em branco."
    )

    instrucoes.append("")
    instrucoes.append("PARÁGRAFOS PARA CORRIGIR:")

    for numero, ocorrencia in enumerate(
        excedentes,
        start=1
    ):

        instrucoes.append(
            f"[PARÁGRAFO {numero}]"
        )

        instrucoes.append(
            ocorrencia["texto"]
        )

        instrucoes.append("")

    prompt_correcao = "\n".join(
        instrucoes
    )

    print()
    print("=" * 60)
    print("ENVIANDO CORREÇÃO DE ABERTURAS PARA OLLAMA")
    print("=" * 60)

    print(
        f"MODELO: {MODELO_OLLAMA}"
    )

    print(
        "PROMPT:",
        len(prompt_correcao),
        "caracteres"
    )

    inicio_ollama = time.time()

    try:

        resposta = requests.post(

            "http://localhost:11434/api/generate",

            json={

                "model": MODELO_OLLAMA,

                "prompt": prompt_correcao,

                "stream": False,

                "think": False,

                "options": {

                    "num_predict": 300,

                    "num_ctx": 4096,

                    "temperature": 0.0

                }

            },

            timeout=(30, 180)

        )

        tempo_ollama = (
            time.time() - inicio_ollama
        )

        print()
        print(
            "OLLAMA STATUS:",
            resposta.status_code
        )

        print(
            "TEMPO:",
            round(tempo_ollama, 1),
            "segundos"
        )

        if resposta.status_code != 200:

            print(
                "ERRO: OLLAMA NÃO RETORNOU 200"
            )

            print(
                "TEXTO ORIGINAL SERÁ PRESERVADO"
            )

            print("=" * 60)

            return texto

        dados = resposta.json()

        texto_corrigido = dados.get(
            "response",
            ""
        ).strip()

        print(
            "RESPOSTA OLLAMA:",
            len(texto_corrigido),
            "caracteres"
        )

        if not texto_corrigido:

            print(
                "RESPOSTA VAZIA"
            )

            print(
                "TEXTO ORIGINAL SERÁ PRESERVADO"
            )

            print("=" * 60)

            return texto

    except Exception as erro:

        print()
        print(
            "ERRO NA CORREÇÃO:",
            erro
        )

        print(
            "TEXTO ORIGINAL SERÁ PRESERVADO"
        )

        print("=" * 60)

        return texto

    # --------------------------------------------------------
    # EXTRAIR PARÁGRAFOS CORRIGIDOS
    # --------------------------------------------------------

    corrigidos = [

        p.strip()

        for p in re.split(
            r"\n\s*\n",
            texto_corrigido
        )

        if p.strip()

    ]

    print()
    print(
        "PARÁGRAFOS CORRIGIDOS RECEBIDOS:",
        len(corrigidos)
    )

    # --------------------------------------------------------
    # SEGURANÇA
    # --------------------------------------------------------

    if len(corrigidos) != len(excedentes):

        print()
        print(
            "ATENÇÃO: QUANTIDADE DE PARÁGRAFOS DIFERENTE"
        )

        print(
            "ESPERADO:",
            len(excedentes)
        )

        print(
            "RECEBIDO:",
            len(corrigidos)
        )

        print(
            "CORREÇÃO DESCARTADA"
        )

        print(
            "TEXTO ORIGINAL SERÁ PRESERVADO"
        )

        print("=" * 60)

        return texto

    # --------------------------------------------------------
    # SUBSTITUIR SOMENTE OS PARÁGRAFOS EXCEDENTES
    # --------------------------------------------------------

    mapa_correcoes = {}

    for ocorrencia, novo_texto in zip(
        excedentes,
        corrigidos
    ):

        mapa_correcoes[
            ocorrencia["indice"]
        ] = novo_texto

    novo_texto_final = texto

    # --------------------------------------------------------
    # RECONSTRUIR TEXTO PRESERVANDO BLOCOS
    # --------------------------------------------------------

    linhas_originais = texto.splitlines()

    resultado = []

    indice_paragrafo = 0

    for linha in linhas_originais:

        linha_limpa = linha.strip()

        # ----------------------------------------------------
        # LINHAS VAZIAS
        # ----------------------------------------------------

        if not linha_limpa:

            resultado.append("")

            continue

        # ----------------------------------------------------
        # MARCADORES / TÍTULOS / SEGMENTOS
        # ----------------------------------------------------

        if (
            re.match(
                r"^\*\*BLOCO\s+\d+\*\*$",
                linha_limpa,
                re.IGNORECASE
            )
            or linha_limpa.startswith("**")
            or linha_limpa.startswith("- ")
        ):

            resultado.append(linha)

            continue

        # ----------------------------------------------------
        # PARÁGRAFO
        # ----------------------------------------------------

        if indice_paragrafo in mapa_correcoes:

            resultado.append(
                mapa_correcoes[
                    indice_paragrafo
                ]
            )

        else:

            resultado.append(linha)

        indice_paragrafo += 1

    novo_texto_final = "\n".join(
        resultado
    )

    # --------------------------------------------------------
    # NOVA CONTAGEM
    # --------------------------------------------------------

    paragrafos_novos = []

    for linha in novo_texto_final.splitlines():

        linha_limpa = linha.strip()

        if not linha_limpa:
            continue

        if re.match(
            r"^\*\*BLOCO\s+\d+\*\*$",
            linha_limpa,
            re.IGNORECASE
        ):
            continue

        if linha_limpa.startswith("- "):
            continue

        if (
            linha_limpa.startswith("**")
            and linha_limpa.endswith("**")
        ):
            continue

        paragrafos_novos.append(
            linha_limpa
        )

    novas_ocorrencias = []

    for indice, paragrafo in enumerate(
        paragrafos_novos
    ):

        inicio_normalizado = normalizar(
            paragrafo
        )

        if (
            inicio_normalizado.startswith(
                tema_normalizado
            )
            or inicio_normalizado.startswith(
                "a " + tema_normalizado
            )
            or inicio_normalizado.startswith(
                "o " + tema_normalizado
            )
        ):

            novas_ocorrencias.append(
                indice
            )

    # --------------------------------------------------------
    # RESULTADO FINAL
    # --------------------------------------------------------

    print()
    print("=" * 60)
    print("RESULTADO DO CONTROLE")
    print("=" * 60)

    print(
        "ANTES:",
        len(ocorrencias),
        "aberturas diretas"
    )

    print(
        "DEPOIS:",
        len(novas_ocorrencias),
        "aberturas diretas"
    )

    print(
        "LIMITE:",
        limite
    )

    if len(novas_ocorrencias) <= limite:

        print(
            "STATUS FINAL: APROVADO"
        )

    else:

        print(
            "STATUS FINAL: AINDA ACIMA DO LIMITE"
        )

        print(
            "ATENÇÃO: NÃO SERÁ FEITA NOVA CHAMADA AO OLLAMA"
        )

    print("=" * 60)

    return novo_texto_final



# ============================================================
# AUDITAR E CORRIGIR ABERTURAS REPETIDAS
# ============================================================

def auditar_e_corrigir_aberturas(
    conteudo,
    tema
):

    print()
    print("==============================")
    print("AUDITORIA DE ABERTURAS")
    print("==============================")



    # ========================================================
    # 01. NORMALIZAÇÃO
    # ========================================================

    tema_limpo = (
        str(tema)
        .strip()
        .lower()
    )

    # ========================================================
    # 02. EXTRAIR PARÁGRAFOS
    # ========================================================

    paragrafos = [

        p.strip()

        for p in re.split(
            r"\n\s*\n",
            conteudo
        )

        if p.strip()

    ]

    # ========================================================
    # 03. CONTAR ABERTURAS DIRETAS
    # ========================================================

    aberturas_diretas = []

    for indice, paragrafo in enumerate(paragrafos):

        inicio = paragrafo.lower()

        # Remove pequenas marcações
        inicio = re.sub(
            r"^[\*\#\-\s]+",
            "",
            inicio
        ).strip()

        # ---------------------------------
        # Verificar início com o tema
        # ---------------------------------

        if inicio.startswith(
            tema_limpo
        ):

            aberturas_diretas.append(
                indice
            )

    # ========================================================
    # 04. RESULTADO
    # ========================================================

    quantidade_aberturas = len(
        aberturas_diretas
    )

    print(
        "PALAVRA-CHAVE:",
        tema
    )

    print(
        "ABERTURAS DIRETAS:",
        quantidade_aberturas
    )

    print(
        "LIMITE PERMITIDO:",
        3
    )

    print(
        "POSIÇÕES:",
        [
            x + 1
            for x in aberturas_diretas
        ]
    )

    # ========================================================
    # 05. DENTRO DO LIMITE
    # ========================================================

    if quantidade_aberturas <= 3:

        print()
        print(
            "ABERTURAS DENTRO DO LIMITE"
        )

        print(
            "CORREÇÃO NECESSÁRIA:",
            False
        )

        return conteudo

    # ========================================================
    # 06. LIMITE EXCEDIDO
    # ========================================================

    print()
    print("==============================")
    print("ABERTURAS EXCEDERAM O LIMITE")
    print("==============================")

    print(
        "ENCONTRADAS:",
        quantidade_aberturas
    )

    print(
        "PERMITIDAS:",
        3
    )

    print(
        "INICIANDO CORREÇÃO IA..."
    )

    # ========================================================
    # 07. PROMPT DE CORREÇÃO
    # ========================================================

    prompt_correcao = f"""
Você é um editor técnico especializado em naturalidade editorial.

Revise o conteúdo abaixo.

TEMA:
{tema}

==================================================
REGRA PRINCIPAL — ABERTURA DOS PARÁGRAFOS
==================================================

Existe uma regra editorial obrigatória:

No máximo 3 dos 15 parágrafos podem começar
diretamente com a palavra-chave:

"{tema}"

Os demais parágrafos devem começar naturalmente
por outros elementos da narrativa.

A palavra-chave NÃO está proibida no início dos
parágrafos.

Ela pode aparecer no início em até 3 parágrafos.

O objetivo é apenas impedir repetição excessiva.

Exemplos de entradas naturais:

o cenário;
a operação;
a necessidade;
o processo;
a aplicação;
um critério técnico;
uma característica;
uma consequência;
uma condição operacional;
a experiência;
o contexto industrial;
o sistema;
o ambiente;
a demanda;
o funcionamento;
a utilização.

==================================================
VARIAÇÃO NARRATIVA
==================================================

Evite também repetir excessivamente a mesma
estrutura de abertura.

Não substitua uma repetição por outra.

Evite sequências repetitivas como:

"Em ambientes..."
"Em ambientes..."
"Em ambientes..."

"Durante..."
"Durante..."
"Durante..."

"Para..."
"Para..."
"Para..."

"Quando..."
"Quando..."
"Quando..."

"A empresa..."
"A empresa..."
"A empresa..."

"Nossa empresa..."
"Nossa empresa..."
"Nossa empresa..."

A construção dos parágrafos deve parecer natural
e escrita por um especialista humano.

Varie:

- sujeito;
- perspectiva;
- posição da palavra-chave;
- estrutura sintática;
- ponto de entrada da informação;
- relação entre contexto e explicação técnica.

==================================================
PRESERVAÇÃO DO CONTEÚDO
==================================================

A revisão é exclusivamente editorial.

NÃO:

- altere o significado técnico;
- remova informações técnicas;
- invente informações;
- crie informações técnicas novas;
- altere especificações;
- altere números;
- altere características técnicas;
- altere aplicações técnicas;
- altere afirmações sustentadas pelas fontes;
- substitua o protagonista;
- transforme componentes em protagonistas;
- transforme manutenção em protagonista;
- transforme instalação em protagonista;
- transforme aplicações em protagonistas.

A palavra-chave continua sendo o protagonista
semântico da página.

==================================================
ESTRUTURA OBRIGATÓRIA
==================================================

Preserve exatamente:

5 blocos;

3 parágrafos em cada bloco;

15 parágrafos no total;

12 segmentos de aplicação.

NÃO:

- crie novos parágrafos;
- remova parágrafos;
- una parágrafos;
- divida parágrafos;
- altere a ordem dos blocos;
- altere os títulos dos blocos;
- altere a quantidade de segmentos;
- altere o conteúdo dos segmentos sem necessidade.

==================================================
EXTENSÃO
==================================================

Não reduza o conteúdo.

Não transforme parágrafos longos em
parágrafos curtos.

Não remova informações apenas para
corrigir uma abertura.

Quando uma abertura precisar ser alterada,
modifique somente a construção inicial
necessária para eliminar a repetição.

Preserve o restante do parágrafo.

==================================================
CRITÉRIO DE NATURALIDADE
==================================================

Antes de finalizar, verifique internamente:

1. Existem no máximo 3 parágrafos iniciados
   diretamente por "{tema}"?

2. As demais aberturas possuem variedade?

3. Existe alguma sequência excessiva de
   estruturas iniciadas por "Em", "Durante",
   "Para", "Quando", "A empresa" ou outra
   fórmula repetitiva?

4. A palavra-chave continua sendo o
   protagonista semântico?

5. O texto continua tecnicamente equivalente
   ao conteúdo original?

6. Foram preservados exatamente:
   - 5 blocos;
   - 15 parágrafos;
   - 12 segmentos?

Somente finalize quando todas essas condições
forem atendidas.

==================================================
SAÍDA
==================================================

Retorne somente o conteúdo corrigido.

Não explique as alterações.

Não apresente comentários editoriais.

Não mostre esta instrução.

Não mostre análise.

==================================================
CONTEÚDO
==================================================

{conteudo}
"""

    # ========================================================
    # 08. CHAMADA OLLAMA
    # ========================================================

    inicio_correcao = time.time()

    try:

        resposta = requests.post(

            "http://localhost:11434/api/generate",

            json={

                "model":
                    MODELO_OLLAMA,

                "prompt":
                    prompt_correcao,

                "stream":
                    False,

                "think":
                    False,

                "options": {

                    "num_predict":
                        800,

                    "num_ctx":
                        8192,

                    "temperature":
                        0.0,

                    "top_p":
                        0.9,

                    "repeat_penalty":
                        1.08

                }

            },

            timeout=(
                30,
                900
            )

        )

    except requests.exceptions.Timeout:

        print()
        print(
            "TIMEOUT NA CORREÇÃO"
        )

        print(
            "MANTENDO CONTEÚDO ORIGINAL"
        )

        return conteudo

    except requests.exceptions.ConnectionError as e:

        print()
        print(
            "ERRO DE CONEXÃO NA CORREÇÃO"
        )

        print(
            repr(e)
        )

        print(
            "MANTENDO CONTEÚDO ORIGINAL"
        )

        return conteudo

    except Exception as e:

        print()
        print(
            "ERRO NA CORREÇÃO"
        )

        print(
            repr(e)
        )

        print(
            "MANTENDO CONTEÚDO ORIGINAL"
        )

        return conteudo

    # ========================================================
    # 09. RECEBER RESULTADO
    # ========================================================

    if resposta.status_code != 200:

        print()
        print(
            "ERRO HTTP NA CORREÇÃO:",
            resposta.status_code
        )

        print(
            "MANTENDO CONTEÚDO ORIGINAL"
        )

        return conteudo

    try:

        dados = resposta.json()

        novo_conteudo = dados.get(
            "response",
            ""
        ).strip()

    except Exception:

        novo_conteudo = ""

    # ========================================================
    # 10. VALIDAR RETORNO
    # ========================================================

    if not novo_conteudo:

        print()
        print(
            "CORREÇÃO IA RETORNOU VAZIO"
        )

        print(
            "MANTENDO CONTEÚDO ORIGINAL"
        )

        return conteudo

    tempo_correcao = (
        time.time()
        - inicio_correcao
    )

    print()
    print("==============================")
    print("CORREÇÃO IA FINALIZADA")
    print("==============================")

    print(
        "TEMPO:",
        round(
            tempo_correcao,
            1
        ),
        "segundos"
    )

    print(
        "CARACTERES ANTES:",
        len(conteudo)
    )

    print(
        "CARACTERES DEPOIS:",
        len(novo_conteudo)
    )

    # ========================================================
    # 11. AUDITORIA NOVAMENTE
    # ========================================================

    paragrafos_corrigidos = [

        p.strip()

        for p in re.split(
            r"\n\s*\n",
            novo_conteudo
        )

        if p.strip()

    ]

    aberturas_corrigidas = []

    for indice, paragrafo in enumerate(
        paragrafos_corrigidos
    ):

        inicio = paragrafo.lower()

        inicio = re.sub(
            r"^[\*\#\-\s]+",
            "",
            inicio
        ).strip()

        if inicio.startswith(
            tema_limpo
        ):

            aberturas_corrigidas.append(
                indice
            )

    quantidade_corrigida = len(
        aberturas_corrigidas
    )

    print()
    print("==============================")
    print("RESULTADO PÓS-CORREÇÃO")
    print("==============================")

    print(
        "ABERTURAS ANTES:",
        quantidade_aberturas
    )

    print(
        "ABERTURAS DEPOIS:",
        quantidade_corrigida
    )

    print(
        "LIMITE:",
        3
    )

    print(
        "CORREÇÃO FUNCIONOU:",
        quantidade_corrigida <= 3
    )

    print(
        "POSIÇÕES FINAIS:",
        [
            x + 1
            for x in aberturas_corrigidas
        ]
    )

    # ========================================================
    # 12. SEGURANÇA
    # ========================================================

    # Se a IA devolveu algo ainda pior,
    # não substituímos o conteúdo original.

    if quantidade_corrigida > 3:

        print()
        print(
            "ATENÇÃO: LIMITE AINDA EXCEDIDO"
        )

        print(
            "CONTEÚDO ORIGINAL SERÁ PRESERVADO"
        )

        return conteudo

    print()
    print(
        "CONTEÚDO CORRIGIDO ACEITO"
    )

    return novo_conteudo
    




# ============================================================
# REVISAR NATURALIDADE DO CONTEÚDO
# ============================================================

def revisar_naturalidade_conteudo(
    tema,
    conteudo
):

    print()
    print("==============================")
    print("INICIANDO REVISÃO DE NATURALIDADE")
    print("==============================")

    print(
        "TEMA:",
        tema
    )

    print(
        "CONTEÚDO ORIGINAL:",
        len(conteudo),
        "caracteres"
    )

    # ========================================================
    # 01. VALIDAR CONTEÚDO
    # ========================================================

    if not conteudo:

        print()
        print("==============================")
        print("CONTEÚDO VAZIO PARA REVISÃO")
        print("==============================")

        return ""

    # ========================================================
    # 02. PROMPT ENXUTO
    # ========================================================

    prompt = f"""
Você é um editor técnico experiente.

Faça uma revisão editorial do conteúdo abaixo.

OBJETIVO:
Melhorar a naturalidade da narrativa sem modificar
o conteúdo técnico.

TEMA PRINCIPAL:
{tema}

REGRAS:

1. "{tema}" continua sendo o protagonista absoluto
da narrativa.

2. Revise principalmente as aberturas dos parágrafos
e as transições entre ideias.

3. Evite que os parágrafos comecem repetidamente
com "{tema}" ou com a mesma estrutura sintática.

4. Varie naturalmente o ponto de entrada das frases:
contexto, situação, característica, consequência,
aplicação, necessidade, processo ou resultado.

5. Não force sinônimos para "{tema}" quando isso
prejudicar a precisão técnica.

6. Preserve integralmente os fatos e informações
técnicas existentes.

NÃO ALTERE:
- números;
- especificações;
- características;
- aplicações;
- processos;
- relações de causa e efeito;
- informações técnicas;
- sentido das afirmações.

7. Não introduza informações novas.

8. Não remova informações existentes.

9. Não resuma o conteúdo.

10. Preserve exatamente a estrutura existente:
- 5 blocos;
- 15 parágrafos;
- 12 segmentos de aplicação;
- títulos;
- ordem dos blocos.

11. Componentes, manutenção, instalação e aplicações
devem continuar como elementos de apoio.
O protagonista permanece sendo "{tema}".

12. Faça somente as alterações necessárias para que
o texto pareça escrito de forma natural por um
especialista humano.

IMPORTANTE:
Não tente reconstruir o texto.
Não altere sua estrutura.
Não transforme o conteúdo em outro texto.
Apenas refine as construções que apresentam
repetição ou artificialidade.

RETORNE SOMENTE O TEXTO REVISADO.

==================================================
CONTEÚDO
==================================================

{conteudo}
"""

    # ========================================================
    # 03. CONTROLE
    # ========================================================

    num_predict = 2600
    num_ctx = 12288

    print()
    print("==============================")
    print("CONTROLE REVISÃO")
    print("==============================")

    print(
        "PROMPT:",
        len(prompt),
        "caracteres"
    )

    print(
        "TOKENS:",
        num_predict
    )

    print(
        "CONTEXTO:",
        num_ctx
    )

    print(
        "MODELO:",
        MODELO_OLLAMA
    )

    # ========================================================
    # 04. ENVIAR PARA OLLAMA
    # ========================================================

    print()
    print("==============================")
    print("ENVIANDO REVISÃO PARA OLLAMA")
    print("==============================")

    inicio_ollama = time.time()

    try:

        resposta = requests.post(

            "http://localhost:11434/api/generate",

            json={

                "model":
                    MODELO_OLLAMA,

                "prompt":
                    prompt,

                "stream":
                    True,

                "think":
                    False,

                "options": {

                    "num_predict":
                        num_predict,

                    "num_ctx":
                        num_ctx,

                    "temperature":
                        0.1,

                    "top_p":
                        0.9,

                    "repeat_penalty":
                        1.08

                }

            },

            stream=True,

            timeout=(
                30,
                900
            )

        )

    except requests.exceptions.Timeout:

        print()
        print("==============================")
        print("TIMEOUT OLLAMA — REVISÃO")
        print("==============================")

        return ""

    except requests.exceptions.ConnectionError as e:

        print()
        print("==============================")
        print("ERRO DE CONEXÃO OLLAMA — REVISÃO")
        print("==============================")

        print(
            repr(e)
        )

        return ""

    except Exception as e:

        print()
        print("==============================")
        print("ERRO OLLAMA — REVISÃO")
        print("==============================")

        print(
            repr(e)
        )

        return ""

    print()
    print("==============================")
    print("OLLAMA RESPONDENDO — REVISÃO")
    print("==============================")

    print(
        "STATUS:",
        resposta.status_code
    )

    if resposta.status_code != 200:

        try:

            print(
                resposta.text[:1000]
            )

        except Exception:

            pass

        return ""

    # ========================================================
    # 05. RECEBER STREAM
    # ========================================================

    conteudo_revisado = ""

    ultimo_print = time.time()

    try:

        for linha in resposta.iter_lines():

            if not linha:

                continue

            try:

                dados = json.loads(
                    linha.decode(
                        "utf-8"
                    )
                )

            except Exception:

                continue

            trecho = dados.get(
                "response",
                ""
            )

            if trecho:

                conteudo_revisado += trecho

                agora = time.time()

                if agora - ultimo_print >= 30:

                    print()
                    print("==============================")
                    print("STATUS REVISÃO OLLAMA")
                    print("==============================")

                    print(
                        "TEMPO DECORRIDO:",
                        round(
                            agora - inicio_ollama,
                            1
                        ),
                        "s"
                    )

                    print(
                        "CARACTERES:",
                        len(
                            conteudo_revisado
                        )
                    )

                    ultimo_print = agora

            if dados.get(
                "done",
                False
            ):

                break

    except Exception as e:

        print()
        print("==============================")
        print("ERRO DURANTE STREAM — REVISÃO")
        print("==============================")

        print(
            repr(e)
        )

        return ""

    fim_ollama = time.time()

    # ========================================================
    # 06. LIMPAR RESULTADO
    # ========================================================

    conteudo_revisado = (
        conteudo_revisado
        .strip()
    )

    print()
    print("==============================")
    print("REVISÃO DE NATURALIDADE FINALIZADA")
    print("==============================")

    print(
        "TEMPO TOTAL:",
        round(
            fim_ollama - inicio_ollama,
            1
        ),
        "segundos"
    )

    print(
        "CONTEÚDO ORIGINAL:",
        len(conteudo),
        "caracteres"
    )

    print(
        "CONTEÚDO REVISADO:",
        len(conteudo_revisado),
        "caracteres"
    )

    # ========================================================
    # 07. VALIDAR RETORNO
    # ========================================================

    if not conteudo_revisado:

        print()
        print("==============================")
        print("REVISÃO NÃO RETORNOU CONTEÚDO")
        print("==============================")

        print(
            "CONTEÚDO ORIGINAL SERÁ PRESERVADO"
        )

        return ""

    # ========================================================
    # 08. MOSTRAR RESULTADO
    # ========================================================

    print()
    print("==============================")
    print("CONTEÚDO REVISADO RECEBIDO")
    print("==============================")

    print(
        conteudo_revisado[:3000]
    )

    print()
    print("==============================")
    print("FIM REVISÃO DE NATURALIDADE")
    print("==============================")

    return conteudo_revisado



# ============================================================
# CORRIGIR CONTEÚDO COM QWEN3
# ============================================================

def corrigir_conteudo_com_qwen3(
    tema,
    conteudo,
    mapa_mead,
    paragrafos_por_bloco,
    segmentos_encontrados
):

    print()
    print("==============================")
    print("INICIANDO CORREÇÃO COM MODELO CONFIGURADO")
    print("==============================")


    if not conteudo:

        print(
            "CONTEÚDO VAZIO"
        )

        return ""


    # ========================================================
    # 01. DIAGNÓSTICO
    # ========================================================

    problemas = []


    # ========================================================
    # 02. VERIFICAR BLOCOS
    # ========================================================

    blocos_encontrados = 0


    marcadores_blocos = [

        "**BLOCO 1**",
        "**BLOCO 2**",
        "**BLOCO 3**",
        "**BLOCO 4**",
        "**BLOCO 5**"

    ]


    for marcador in marcadores_blocos:

        if marcador in conteudo:

            blocos_encontrados += 1


    if blocos_encontrados != 5:

        problemas.append(
            "A estrutura deve possuir exatamente "
            "5 blocos."
        )


    # ========================================================
    # 03. VERIFICAR PARÁGRAFOS
    # ========================================================

    total_paragrafos = sum(
        paragrafos_por_bloco.values()
    )


    for i in range(
        1,
        6
    ):

        quantidade = paragrafos_por_bloco.get(
            i,
            0
        )


        if quantidade != 3:

            problemas.append(
                f"O BLOCO {i} possui "
                f"{quantidade} parágrafos. "
                f"Deve possuir exatamente 3."
            )


    if total_paragrafos != 15:

        problemas.append(
            "O conteúdo deve possuir "
            "exatamente 15 parágrafos."
        )


    # ========================================================
    # 04. VERIFICAR SEGMENTOS
    # ========================================================

    if segmentos_encontrados != 12:

        problemas.append(
            f"Foram encontrados "
            f"{segmentos_encontrados} segmentos. "
            f"Devem existir exatamente 12."
        )


    # ========================================================
    # 05. NENHUM PROBLEMA
    # ========================================================

    if not problemas:

        print()
        print("==============================")
        print("NENHUM PROBLEMA PARA QWEN3")
        print("==============================")

        return conteudo


    # ========================================================
    # 06. MOSTRAR PROBLEMAS
    # ========================================================

    print()
    print("==============================")
    print("PROBLEMAS ENCONTRADOS")
    print("==============================")


    for problema in problemas:

        print(
            "-",
            problema
        )


    # ========================================================
    # 07. CONTROLE DO MAPA
    # ========================================================

    mapa_resumido = str(
        mapa_mead
    )


    if len(mapa_resumido) > 5000:

        mapa_resumido = mapa_resumido[
            :5000
        ]


    # ========================================================
    # 08. PROMPT QWEN3
    # ========================================================

    prompt_correcao = f"""
Você é um editor técnico responsável por corrigir
somente problemas estruturais de uma página já escrita.

TEMA:
{tema}

O conteúdo abaixo foi produzido por outro modelo
e deve ser PRESERVADO sempre que estiver correto.

Não reescreva o conteúdo inteiro.

Não melhore estilo sem necessidade.

Não altere informações técnicas corretas.

Não mude o protagonista.

Não invente informações.

Não acrescente dados técnicos que não estejam
presentes no conteúdo ou no MAPA MEAD.

Sua função é corrigir SOMENTE os problemas
identificados.

==================================================
PROBLEMAS IDENTIFICADOS
==================================================

{chr(10).join("- " + p for p in problemas)}

==================================================
REGRAS OBRIGATÓRIAS
==================================================

A página final deve possuir:

5 blocos.

Cada bloco deve possuir exatamente
3 parágrafos.

Total de 15 parágrafos.

Depois dos blocos:

12 segmentos de aplicação.

Os segmentos devem aparecer somente na
seção:

**SEGMENTOS DE APLICAÇÃO**

==================================================
MAPA MEAD
==================================================

{mapa_resumido}

Use o MAPA somente como referência para
preservar o tema e o contexto.

==================================================
CONTEÚDO ORIGINAL
==================================================

{conteudo}

==================================================
REGRA PRINCIPAL
==================================================

Preserve o máximo possível do conteúdo original.

Corrija somente:

- blocos ausentes;
- quantidade incorreta de parágrafos;
- segmentos ausentes ou em quantidade incorreta;
- problemas necessários para recuperar a estrutura.

Se um bloco já possui 3 parágrafos corretos,
NÃO reescreva esse bloco.

Se um segmento já está correto,
NÃO substitua desnecessariamente.

Não altere o sentido técnico.

Não faça uma nova redação completa.

==================================================
FORMATO FINAL
==================================================

Entregue somente o conteúdo corrigido.

Use exatamente:

**BLOCO 1**

[3 parágrafos]

**BLOCO 2**

[3 parágrafos]

**BLOCO 3**

[3 parágrafos]

**BLOCO 4**

[3 parágrafos]

**BLOCO 5**

[3 parágrafos]

**SEGMENTOS DE APLICAÇÃO**

- segmento
- segmento
- segmento
- segmento
- segmento
- segmento
- segmento
- segmento
- segmento
- segmento
- segmento
- segmento

Não explique o que foi corrigido.
Não fale sobre inteligência artificial.
Não mostre o MAPA MEAD.
Entregue somente o conteúdo final.
"""


    # ========================================================
    # 09. DEBUG
    # ========================================================

    print()
    print("==============================")
    print("CONTROLE ENVIO OLLAMA")
    print("==============================")


    print(
        "MODELO:",
        MODELO_OLLAMA
    )


    print(
        "CARACTERES CONTEÚDO:",
        len(conteudo)
    )


    print(
        "CARACTERES MAPA:",
        len(mapa_resumido)
    )


    print(
        "CARACTERES PROBLEMAS:",
        len(
            "\n".join(
                problemas
            )
        )
    )


    print(
        "PROMPT:",
        len(prompt_correcao)
    )


    # ========================================================
    # 10. CONFIGURAÇÃO OLLAMA
    # ========================================================

    num_predict = 2600
    num_ctx = 12288


    inicio_qwen3 = time.time()


    # ========================================================
    # 11. OLLAMA
    # ========================================================

    try:

        resposta = requests.post(

            "http://localhost:11434/api/generate",

            json={

                "model":
                    MODELO_OLLAMA,

                "prompt":
                    prompt_correcao,

                "stream":
                    True,

                "think":
                    False,

                "options": {

                    "num_predict":
                        num_predict,

                    "num_ctx":
                        num_ctx,

                    "temperature":
                        0.1,

                    "top_p":
                        0.9,

                    "repeat_penalty":
                        1.08

                }

            },

            stream=True,

            timeout=(
                30,
                900
            )

        )


    except requests.exceptions.Timeout:

        print()
        print("==============================")
        print("TIMEOUT QWEN3")
        print("==============================")

        return ""


    except requests.exceptions.ConnectionError as e:

        print()
        print("==============================")
        print("ERRO DE CONEXÃO QWEN3")
        print("==============================")

        print(
            repr(e)
        )

        return ""


    except Exception as e:

        print()
        print("==============================")
        print("ERRO QWEN3")
        print("==============================")

        print(
            repr(e)
        )

        return ""


    print()
    print("==============================")
    print("QWEN3 RESPONDENDO")
    print("==============================")


    print(
        "STATUS:",
        resposta.status_code
    )


    if resposta.status_code != 200:

        try:

            print(
                resposta.text[:1000]
            )

        except Exception:

            pass

        return ""


    # ========================================================
    # 12. RECEBER STREAM
    # ========================================================

    conteudo_corrigido = ""

    ultimo_print = time.time()


    try:

        for linha in resposta.iter_lines():

            if not linha:

                continue


            try:

                dados = json.loads(
                    linha.decode(
                        "utf-8"
                    )
                )

            except Exception:

                continue


            trecho = dados.get(
                "response",
                ""
            )


            if trecho:

                conteudo_corrigido += trecho


                agora = time.time()


                if agora - ultimo_print >= 30:

                    print()
                    print("==============================")
                    print("STATUS QWEN3")
                    print("==============================")


                    print(
                        "TEMPO DECORRIDO:",
                        round(
                            agora - inicio_qwen3,
                            1
                        ),
                        "s"
                    )


                    print(
                        "CARACTERES:",
                        len(
                            conteudo_corrigido
                        )
                    )


                    ultimo_print = agora


            if dados.get(
                "done",
                False
            ):

                break


    except Exception as e:

        print()
        print("==============================")
        print("ERRO DURANTE STREAM QWEN3")
        print("==============================")


        print(
            repr(e)
        )


        return ""


    fim_qwen3 = time.time()


    conteudo_corrigido = (
        conteudo_corrigido
        .strip()
    )


    # ========================================================
    # 13. RESULTADO
    # ========================================================

    print()
    print("==============================")
    print("QWEN3 FINALIZADO")
    print("==============================")


    print(
        "TEMPO TOTAL:",
        round(
            fim_qwen3 - inicio_qwen3,
            1
        ),
        "segundos"
    )


    print(
        "CARACTERES:",
        len(
            conteudo_corrigido
        )
    )


    # ========================================================
    # 14. VERIFICAR RESPOSTA
    # ========================================================

    if not conteudo_corrigido:

        print()
        print("==============================")
        print("QWEN3 RETORNOU VAZIO")
        print("==============================")

        return ""


    # ========================================================
    # 15. VALIDAÇÃO BÁSICA DA CORREÇÃO
    # ========================================================

    blocos_corrigidos = sum(

        1
        for marcador in marcadores_blocos

        if marcador in conteudo_corrigido

    )


    segmentos_corrigidos = 0


    if "**SEGMENTOS DE APLICAÇÃO**" in conteudo_corrigido:

        trecho_segmentos = (
            conteudo_corrigido.split(
                "**SEGMENTOS DE APLICAÇÃO**",
                1
            )[1]
        )


        for linha in trecho_segmentos.splitlines():

            linha = linha.strip()


            if (
                linha.startswith("- ")
                and len(linha) > 2
            ):

                segmentos_corrigidos += 1


    print()
    print("==============================")
    print("VALIDAÇÃO QWEN3")
    print("==============================")


    print(
        "BLOCOS:",
        blocos_corrigidos,
        "/ 5"
    )


    print(
        "SEGMENTOS:",
        segmentos_corrigidos,
        "/ 12"
    )


    if blocos_corrigidos != 5:

        print()
        print(
            "QWEN3 NÃO PRODUZIU ESTRUTURA VÁLIDA"
        )

        return ""


    if segmentos_corrigidos != 12:

        print()
        print(
            "QWEN3 NÃO PRODUZIU 12 SEGMENTOS"
        )

        return ""


    # ========================================================
    # 16. CORREÇÃO APROVADA
    # ========================================================

    print()
    print("==============================")
    print("CORREÇÃO QWEN3 APROVADA")
    print("==============================")


    return conteudo_corrigido
    

# ============================================================
# EXTRAIR TRECHOS RELEVANTES
# ============================================================

def extrair_trechos_relevantes(
    texto,
    tema,
    mapa_texto,
    limite=5000
):

    if not texto:

        return ""


    # ========================================================
    # 01. NORMALIZAR
    # ========================================================

    texto = str(
        texto
    ).strip()

    if not texto:

        return ""


    # ========================================================
    # 02. PALAVRAS IMPORTANTES
    # ========================================================

    base = (
        f"{tema} "
        f"{mapa_texto}"
    ).lower()


    palavras_mapa = set(

        palavra

        for palavra in re.findall(
            r"\b[a-záàâãéêíóôõúç0-9-]{4,}\b",
            base
        )

    )


    termos_tecnicos = {

        "funcionamento",
        "operação",
        "operacao",
        "pressão",
        "pressao",
        "vazão",
        "vazao",
        "temperatura",
        "eficiência",
        "eficiencia",
        "potência",
        "potencia",
        "rotação",
        "rotacao",
        "motor",
        "rotor",
        "carcaça",
        "carcaca",
        "selo",
        "vedação",
        "vedacao",
        "instalação",
        "instalacao",
        "manutenção",
        "manutencao",
        "segurança",
        "seguranca",
        "desempenho",
        "aplicação",
        "aplicacao",
        "componente",
        "componentes",
        "material",
        "modelo",
        "tipo",
        "capacidade",
        "temperatura",
        "pressão",
        "pressao",
        "altura",
        "fluido",
        "líquido",
        "liquido",
        "tubulação",
        "tubulacao",
        "energia",
        "desgaste",
        "corrosão",
        "corrosao",
        "vibração",
        "vibracao",
        "cavitação",
        "cavitacao"

    }


    termos_genericos = {

        "clique",
        "saiba",
        "contato",
        "comprar",
        "compre",
        "oferta",
        "promoção",
        "promocao",
        "preço",
        "preco",
        "consulte",
        "empresa líder",
        "empresa lider",
        "melhor preço",
        "melhor preco"

    }


    # ========================================================
    # 03. SEPARAR PARÁGRAFOS
    # ========================================================

    paragrafos = [

        p.strip()

        for p in re.split(
            r"\n\s*\n",
            texto
        )

        if p.strip()

    ]


    if not paragrafos:

        paragrafos = [

            texto

        ]


    candidatos = []


    # ========================================================
    # 04. ANALISAR CADA PARÁGRAFO
    # ========================================================

    for indice, paragrafo in enumerate(
        paragrafos
    ):

        texto_lower = paragrafo.lower()


        palavras = set(

            re.findall(
                r"\b[a-záàâãéêíóôõúç0-9-]{4,}\b",
                texto_lower
            )

        )


        pontuacao = 0


    # ========================================================
    # 05. RELAÇÃO COM O TEMA
    # ========================================================

        tema_lower = str(
            tema
        ).lower().strip()


        if tema_lower in texto_lower:

            pontuacao += 15


    # ========================================================
    # 06. TERMOS DO MAPA
    # ========================================================

        correspondencias_mapa = (
            palavras
            &
            palavras_mapa
        )


        pontuacao += min(
            len(correspondencias_mapa) * 2,
            20
        )


    # ========================================================
    # 07. INFORMAÇÃO TÉCNICA
    # ========================================================

        correspondencias_tecnicas = (
            palavras
            &
            termos_tecnicos
        )


        pontuacao += min(
            len(correspondencias_tecnicas) * 3,
            30
        )


    # ========================================================
    # 08. NÚMEROS / UNIDADES
    # ========================================================

        if re.search(
            r"\d+\s*(mm|cm|m|kg|g|bar|psi|°c|c|rpm|kw|cv|l/min|m³/h|hz|v|a)",
            texto_lower
        ):

            pontuacao += 15


    # ========================================================
    # 09. ESTRUTURA TÉCNICA
    # ========================================================

        if ":" in paragrafo:

            pontuacao += 2


        if ";" in paragrafo:

            pontuacao += 2


    # ========================================================
    # 10. PENALIZAR CONTEÚDO GENÉRICO
    # ========================================================

        for termo in termos_genericos:

            if termo in texto_lower:

                pontuacao -= 5


    # ========================================================
    # 11. TAMANHO
    # ========================================================

        tamanho = len(
            paragrafo
        )


        if tamanho < 100:

            pontuacao -= 3


        elif tamanho > 3000:

            pontuacao -= 2


        candidatos.append({

            "indice":
                indice,

            "texto":
                paragrafo,

            "pontuacao":
                pontuacao

        })


    # ========================================================
    # 12. ORDENAR
    # ========================================================

    candidatos.sort(

        key=lambda item:
            item["pontuacao"],

        reverse=True

    )


    # ========================================================
    # 13. MONTAR SELEÇÃO
    # ========================================================

    selecionados = []

    total = 0


    for candidato in candidatos:

        trecho = candidato[
            "texto"
        ]


        tamanho = len(
            trecho
        )


        if not tamanho:

            continue


    # ========================================================
    # 14. LIMITE ABSOLUTO
    # ========================================================

        if (
            total
            +
            tamanho
            >
            limite
        ):

            restante = (
                limite
                -
                total
            )


            if restante >= 250:

                trecho = trecho[
                    :restante
                ]


                # não cortar no meio da palavra

                ultimo_espaco = trecho.rfind(
                    " "
                )


                if ultimo_espaco > 0:

                    trecho = trecho[
                        :ultimo_espaco
                    ]


                selecionados.append(
                    trecho
                )


            break


        selecionados.append(
            trecho
        )


        total += tamanho


    # ========================================================
    # 15. PRESERVAR ORDEM ORIGINAL
    # ========================================================

    textos_selecionados = []


    indices_selecionados = set()


    for trecho in selecionados:

        for candidato in candidatos:

            if (
                candidato["texto"]
                ==
                trecho
            ):

                indices_selecionados.add(
                    candidato["indice"]
                )

                break


    for candidato in sorted(
        candidatos,
        key=lambda item:
            item["indice"]
    ):

        if candidato["indice"] in indices_selecionados:

            textos_selecionados.append(
                candidato["texto"]
            )


    resultado = "\n\n".join(
        textos_selecionados
    )


    return resultado[
        :limite
    ]    


# ============================================================
# SELECIONAR GRUPO
# ============================================================

def preparar_grupo_para_ia(
    grupo,
    tema,
    mapa_texto,
    limite=5000
):

    partes = []

    restante = limite


    for fonte in grupo:

        if restante <= 0:

            break


        trecho = extrair_trechos_relevantes(

            fonte["texto"],

            tema,

            mapa_texto,

            limite=restante

        )


        if not trecho:

            continue


        bloco = f"""
[FONTE {fonte["indice"]}]
TIPO: {fonte["tipo"]}
PDF TÉCNICO: {"SIM" if fonte["eh_pdf"] else "NÃO"}
URL: {fonte["url"]}

{trecho}
"""


        if len(bloco) > restante:

            bloco = bloco[
                :restante
            ]


        partes.append(
            bloco
        )


        restante -= len(
            bloco
        )


    return "\n".join(
        partes
    )[:limite]



# ============================================================
# HASH DE TRECHO
# ============================================================

def gerar_hash_trecho(
    texto
):
    """
    Gera um identificador único para o trecho selecionado.

    O mesmo trecho, mesmo que tenha espaços diferentes,
    receberá o mesmo hash.
    """

    texto_normalizado = re.sub(
        r"\s+",
        " ",
        str(
            texto or ""
        ).strip().lower()
    )

    if not texto_normalizado:

        return ""

    return hashlib.sha256(
        texto_normalizado.encode(
            "utf-8"
        )
    ).hexdigest()


def contar_palavra_chave_literal(texto, palavra_chave):
    """Conta a expressão completa, sem aceitar flexões ou variantes."""
    texto = str(texto or "")
    palavra_chave = str(palavra_chave or "").strip()
    if not texto or not palavra_chave:
        return 0
    # Para localizar âncoras na fonte, a busca é tolerante à capitalização.
    # A saída final possui uma auditoria separada, sensível a maiúsculas.
    padrao = re.compile(
        r"(?<!\w)" + re.escape(palavra_chave) + r"(?!\w)",
        flags=re.IGNORECASE | re.UNICODE,
    )
    return len(padrao.findall(texto))

# ============================================================
# SELECIONAR CONTEÚDO EDITORIAL
# ============================================================

def selecionar_conteudo_editorial(
    categorias,
    estrutura_editorial
):
    """
    Cruza as categorias encontradas pelo Python com a estrutura
    editorial selecionada pelo usuário.

    categorias:
        Dicionário com as informações classificadas pelo Python.

    estrutura_editorial:
        Dicionário com os checkboxes selecionados em
        ESTRUTURA_EDITORIAL.

    Retorna:
        Dicionário contendo somente as categorias autorizadas
        pela estrutura editorial.
    """

    resultado = {}

    if not isinstance(categorias, dict):
        return resultado

    if not isinstance(estrutura_editorial, dict):
        return resultado

    for bloco_editorial, ativo in estrutura_editorial.items():

        # Checkbox desmarcado: não participa da seleção
        if not ativo:
            continue

        # Descobre quais categorias alimentam este bloco
        categorias_permitidas = MAPEAMENTO_EDITORIAL.get(
            bloco_editorial,
            []
        )

        if not categorias_permitidas:
            continue

        informacoes_bloco = {}

        for categoria in categorias_permitidas:

            dados_categoria = categorias.get(categoria)

            if not dados_categoria:
                continue

            # Mantém exatamente o conteúdo produzido
            # pela classificação das categorias.
            informacoes_bloco[categoria] = dados_categoria

        if informacoes_bloco:
            resultado[bloco_editorial] = informacoes_bloco

    return resultado


# ============================================================
# MAPEAMENTO EDITORIAL → CATEGORIAS DE CONTEÚDO
# ============================================================
#
# Converte as opções editoriais ativadas pelo usuário
# nas categorias utilizadas pelo mecanismo de pontuação.
#
# IMPORTANTE:
# - O Python continua responsável pela seleção.
# - O Ollama não participa desta etapa.
# - Uma categoria pode atender mais de um foco editorial.
# ============================================================

MAPEAMENTO_EDITORIAL = {

    # --------------------------------------------------------
    # APRESENTAÇÃO
    # --------------------------------------------------------
    "apresentacao": [
        "definicao",
        "ativos_narrativos",
        "duvidas_frequentes"
    ],

    # --------------------------------------------------------
    # FUNCIONAMENTO
    # --------------------------------------------------------
    "funcionamento": [
        "definicao",
        "fabricacao",
        "materia_prima"
    ],

    # --------------------------------------------------------
    # APLICAÇÕES
    # --------------------------------------------------------
    "aplicacoes": [
        "aplicacoes"
    ],

    # --------------------------------------------------------
    # CRITÉRIOS
    # --------------------------------------------------------
    "criterios": [
        "vantagens",
        "manutencao",
        "duvidas_frequentes"
    ],

    # --------------------------------------------------------
    # COMERCIAL
    #
    # Não criamos uma categoria "comercial" artificial.
    # O conteúdo comercial deve ser sustentado por
    # informações técnicas, aplicações e ativos narrativos.
    # --------------------------------------------------------
    "comercial": [
        "ativos_narrativos",
        "aplicacoes",
        "vantagens"
    ],

    # --------------------------------------------------------
    # INFORMAÇÃO TÉCNICA
    # --------------------------------------------------------
    "informacao_tecnica": [
        "definicao",
        "materia_prima",
        "fabricacao",
        "manutencao"
    ],

    # --------------------------------------------------------
    # INSTALAÇÃO / EXECUÇÃO
    # --------------------------------------------------------
    "instalacao_execucao": [
        "fabricacao",
        "manutencao",
        "duvidas_frequentes"
    ],

    # --------------------------------------------------------
    # BENEFÍCIOS
    # --------------------------------------------------------
    "beneficios": [
        "beneficios",
        "vantagens"
    ]
}


# ========================================================
# IDENTIDADE EDITORIAL DO SITE
# ========================================================

def preparar_identidade_editorial(nome_site):
    """
    Normaliza o nome do site informado na interface.

    O nome do site será usado como identidade editorial
    da página final.

    Exemplo:
        Pascal Engenharia
    """
    
    nome_site = str(nome_site or "").strip()
    
    if not nome_site:
        return ""
    
    return re.sub(r"\s+", " ", nome_site)


def fragmento_pode_usar_identidade_site(
    texto,
    nome_site
):
    """
    Verifica se o fragmento pode ser utilizado como
    contexto para uma redação institucional.

    IMPORTANTE:
    - não substitui nomes de empresas automaticamente;
    - não considera empresas de terceiros como identidade
      do conteúdo;
    - permite que o Ollama utilize o contexto técnico;
    - a identidade editorial final será o nome_site.
    """
    
    texto = str(texto or "").strip()
    nome_site = preparar_identidade_editorial(nome_site)
    
    if not texto:
        return False
    
    # Se não existe nome de site configurado,
    # não bloquear o fragmento por esse motivo.
    if not nome_site:
        return True
    
    return True


    
    
# ============================================================
# SELECIONAR INFORMAÇÕES RELEVANTES
# ============================================================

def selecionar_informacoes_relevantes(
    tema,
    textos,
    mapa_mead,
    estrutura_editorial
):

    # Inicialização defensiva no início da função. O conjunto será
    # atualizado após a seleção inicial, antes de montar as reservas COM.
    # Assim, nenhuma ramificação desta função pode consultar o nome antes
    # de sua criação no escopo local.
    hashes_reservas_ocupadas = set()

    id_execucao = f"{time.time():.6f}"

    print()
    print("##################################################")
    print("ENTRADA EM selecionar_informacoes_relevantes()")
    print("ID EXECUÇÃO:", id_execucao)
    print("TEMA:", tema)
    print("##################################################")

    print()
    print("==============================")
    print("SELECIONANDO INFORMAÇÕES RELEVANTES")
    print("==============================")

    print("TEMA:", tema)
    print("TEXTOS RECEBIDOS:", len(textos))

    # ========================================================
    # FUNÇÕES DE IDENTIFICAÇÃO
    # ========================================================

    def normalizar_texto_hash(texto):

        texto = str(
            texto or ""
        )

        texto = re.sub(
            r"\s+",
            " ",
            texto
        ).strip().lower()

        return texto

    def gerar_hash_trecho(texto):

        texto_normalizado = normalizar_texto_hash(
            texto
        )

        return hashlib.sha256(
            texto_normalizado.encode(
                "utf-8"
            )
        ).hexdigest()

    def gerar_id_trecho(hash_trecho):

        return (
            "IR_"
            + hash_trecho[:16]
        )

    # ========================================================
    # RETORNO VAZIO PADRONIZADO
    # ========================================================

    resultado_vazio = {

        "status":
            "sem_informacoes",

        "caracteres":
            0,

        "texto":
            "",

        "fontes":
            [],

        "fragmentos":
            [],

        "informacoes_relevantes":
            [],

        "blocos": {

            "bloco_1": [],
            "bloco_2": [],
            "bloco_3": [],
            "bloco_4": [],
            "bloco_5": []

        },

        "blocos_informacoes": {

            "bloco_1": {
                "informacoes_relevantes": []
            },

            "bloco_2": {
                "informacoes_relevantes": []
            },

            "bloco_3": {
                "informacoes_relevantes": []
            },

            "bloco_4": {
                "informacoes_relevantes": []
            },

            "bloco_5": {
                "informacoes_relevantes": []
            }

        }

    }

    if not textos:

        print("NENHUM TEXTO RECEBIDO")

        return resultado_vazio

    # ========================================================
    # 01. NORMALIZAR FONTES
    # ========================================================

    fontes = []

    for indice, item in enumerate(
        textos,
        start=1
    ):

        if isinstance(
            item,
            dict
        ):

            texto = (
                item.get("texto")
                or item.get("conteudo")
                or item.get("text")
                or ""
            )

            url = (
                item.get("url")
                or item.get("fonte")
                or ""
            )

            tipo = (
                item.get("tipo")
                or "texto"
            )

            # ====================================================
            # PRESERVAR IDENTIDADE DA FONTE
            # ====================================================
            #
            # A identidade já foi identificada na etapa de
            # limpeza das referências.
            #
            # Aqui apenas transportamos o objeto para a estrutura
            # interna "fontes".
            #
            # NÃO identificar novamente.
            # NÃO inferir fabricante pelo domínio.
            # ====================================================

            identidade_fonte = (
                item.get(
                    "identidade_fonte",
                    {}
                )
            )

        else:

            texto = str(
                item
            )

            url = ""

            tipo = "texto"

            identidade_fonte = {}

        texto = str(
            texto
        ).strip()

        url = str(
            url
        ).strip()

        tipo = str(
            tipo
        ).strip()

        if not texto:
            continue

        eh_pdf = (
            ".pdf" in url.lower()
            or "pdf" in tipo.lower()
        )

        fontes.append({

            "indice":
                indice,

            "url":
                url,

            "tipo":
                tipo,

            "texto":
                texto,

            "eh_pdf":
                eh_pdf,

            "identidade_fonte":
                identidade_fonte

        })

    # ========================================================
    # 02. PREPARAR ESTRUTURA EDITORIAL
    # ========================================================

    if not isinstance(
        estrutura_editorial,
        dict
    ):

        estrutura_editorial = {}

    assuntos = [

        str(chave).replace(
            "_",
            " "
        )

        for chave, valor
        in estrutura_editorial.items()

        if valor

    ]

    print()
    print("==============================")
    print("CHECKBOXES RECEBIDOS")
    print("==============================")

    print(
        "ASSUNTOS:",
        assuntos
    )


      
        
    # ========================================================
    # 03. PREPARAR MAPA MEAD
    # ========================================================

    mapa_texto = str(
        mapa_mead or ""
    ).strip()

    if len(mapa_texto) > 5000:

        mapa_texto = mapa_texto[:5000]


    # ========================================================
    # 04. PREPARAR CANDIDATOS
    # ========================================================
    #
    # REGRA:
    #
    # O Python deve entregar ao Ollama fragmentos editoriais
    # utilizáveis.
    #
    # Não basta penalizar índice, comentários, menus ou tabelas.
    # Esses fragmentos devem ser retirados ANTES da pontuação.
    #
    # Cada candidato terá entre 45 e 75 palavras.
    #
    # ========================================================

    candidatos = []

    MIN_PALAVRAS_FRAGMENTO = 35
    MAX_PALAVRAS_FRAGMENTO = 75
    
    

    
    # ========================================================
    # FILTRO DE IDENTIDADE COMERCIAL / PRODUTO
    # ========================================================
    #
    # Descarta fragmentos que tragam:
    # - nomes explícitos de empresas/fabricantes/marcas;
    # - códigos de produtos;
    # - modelos;
    # - SKU / MPN / PN / part number;
    # - referências comerciais;
    # - números de série;
    # - CNPJ;
    # - telefone/e-mail/site;
    # - linguagem de catálogo/comercial.
    #
    # O objetivo é preservar somente informação técnica
    # aproveitável para a construção editorial.
    # ========================================================

    def fragmento_eh_comercialmente_limpo(
        texto,
        identidade_fonte=None
    ):

        texto = str(
            texto or ""
        ).strip()

        if not texto:
            return False

        texto_normalizado = (
            normalizar_assunto_texto(
                texto
            )
        )

        # ----------------------------------------------------
        # 01. URL / DOMÍNIO / E-MAIL
        # ----------------------------------------------------

        if re.search(
            r"(https?://|www\.)\S+",
            texto,
            re.IGNORECASE
        ):
            return False

        if re.search(
            r"\b[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}\b",
            texto,
            re.IGNORECASE
        ):
            return False

        # ----------------------------------------------------
        # 02. CNPJ
        # ----------------------------------------------------

        if re.search(
            r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b",
            texto
        ):
            return False

        # ----------------------------------------------------
        # 03. TELEFONE
        # ----------------------------------------------------

        if re.search(
            r"(?<!\d)"
            r"(?:\+?55[\s.-]*)?"
            r"\(?\d{2}\)?[\s.-]*"
            r"\d{4,5}[\s.-]*\d{4}"
            r"(?!\d)",
            texto
        ):
            return False

        # ----------------------------------------------------
        # 04. IDENTIFICADORES COMERCIAIS EXPLÍCITOS
        # ----------------------------------------------------

        padroes_identificadores = [

            r"\bsku\b",
            r"\bmpn\b",
            r"\bpart[\s-]*number\b",
            r"\bpart[\s-]*no\.?\b",
            r"\bpart[\s-]*n[oº°]?\b",
            r"\bserial[\s-]*number\b",
            r"\bn[uú]mero[\s-]*de[\s-]*s[eé]rie\b",
            r"\bc[oó]digo[\s-]*(?:do|de)?[\s-]*produto\b",
            r"\bc[oó]digo[\s-]*comercial\b",
            r"\brefer[eê]ncia[\s-]*(?:do|de)?[\s-]*produto\b",
            r"\bref\.?[\s-]*(?:do|de)?[\s-]*produto\b",
            r"\bmodelo[\s-]*(?:do|de)?[\s-]*produto\b"

        ]

        for padrao in padroes_identificadores:

            if re.search(
                padrao,
                texto_normalizado,
                re.IGNORECASE
            ):
                return False

        # ----------------------------------------------------
        # 05. CÓDIGOS / MODELOS ALFANUMÉRICOS
        #
        # Não rejeita números técnicos comuns isoladamente.
        #
        # Rejeita combinações típicas de catálogo:
        # ABC-123
        # XYZ123
        # 123-ABC
        # AB-1234-CD
        # etc.
        # ----------------------------------------------------

        padrao_codigo = re.compile(
            r"(?<![A-Za-z0-9])"
            r"(?=[A-Za-z0-9-]{4,30}(?![A-Za-z0-9]))"
            r"(?=[A-Za-z0-9-]*[A-Za-z])"
            r"(?=[A-Za-z0-9-]*\d)"
            r"[A-Za-z]{1,8}"
            r"(?:[-_/]?[A-Za-z0-9]{1,12}){1,4}"
            r"(?![A-Za-z0-9])"
        )

        ocorrencias_codigo = padrao_codigo.findall(
            texto
        )

        # ----------------------------------------------------
        # Evitar rejeitar siglas técnicas normais.
        # ----------------------------------------------------

        codigos_tecnicos_permitidos = {
            "pvc",
            "pead",
            "cpvc",
            "ppr",
            "nbr",
            "iso",
            "ansi",
            "astm",
            "din",
            "api",
            "rpm",
            "ip",
            "dn",
            "pn",
            "mca",
            "kw",
            "cv",
            "hp"
        }

        codigos_suspeitos = []

        for codigo in ocorrencias_codigo:

            codigo_normalizado = (
                codigo.casefold()
            )

            if codigo_normalizado in (
                codigos_tecnicos_permitidos
            ):
                continue

            # Siglas técnicas simples não são códigos
            # comerciais.
            if re.fullmatch(
                r"[A-Za-z]{2,5}",
                codigo
            ):
                continue

            codigos_suspeitos.append(
                codigo
            )

        if len(codigos_suspeitos) >= 1:

            # Um código isolado pode ser uma especificação.
            # Dois ou mais aumentam muito a probabilidade
            # de catálogo/modelo.
            if len(codigos_suspeitos) >= 2:
                return False

            # Código acompanhado de contexto comercial
            # também deve ser descartado.
            contexto_codigo = [
                "modelo",
                "codigo",
                "código",
                "referencia",
                "referência",
                "serie",
                "série",
                "item",
                "produto",
                "catalogo",
                "catálogo"
            ]

            if any(
                termo in texto_normalizado
                for termo in contexto_codigo
            ):
                return False

        # ----------------------------------------------------
        # 06. CONTEXTO EXPLÍCITO DE EMPRESA / FABRICANTE
        # ----------------------------------------------------

        padroes_empresa = [

            r"\bfabricad[oa]\s+pel[ao]\b",
            r"\bproduzid[oa]\s+pel[ao]\b",
            r"\bfornecid[oa]\s+pel[ao]\b",
            r"\bcomercializad[oa]\s+pel[ao]\b",
            r"\bdistribu[ií]d[oa]\s+pel[ao]\b",
            r"\bvendid[oa]\s+pel[ao]\b",

            r"\bfabricante\s*:",
            r"\bfornecedor\s*:",
            r"\bempresa\s*:",
            r"\bmarca\s*:",
            r"\bmodelo\s*:",

            r"\bfabricante\s+[\w.-]+",
            r"\bfornecedor\s+[\w.-]+",
            r"\bmarca\s+[\w.-]+",

            r"\bltda\.?\b",
            r"\bltda\b",
            r"\bepp\b",
            r"\bs\.?a\.?\b",
            r"\bs/a\b",
            r"\binc\.?\b",
            r"\bcorp\.?\b",
            r"\bllc\b"

        ]

        ocorrencias_empresa = 0

        for padrao in padroes_empresa:

            if re.search(
                padrao,
                texto_normalizado,
                re.IGNORECASE
            ):
                ocorrencias_empresa += 1

        if ocorrencias_empresa >= 1:
            return False

        # ----------------------------------------------------
        # 07. SÍMBOLOS DE MARCA REGISTRADA
        # ----------------------------------------------------

        if (
            "®" in texto
            or
            "™" in texto
        ):
            return False

        # ----------------------------------------------------
        # 08. LINGUAGEM DE CATÁLOGO / VENDA
        # ----------------------------------------------------

        marcadores_catalogo = [

            "consulte o catalogo",
            "consulte o catálogo",
            "entre em contato",
            "fale conosco",
            "solicite um orcamento",
            "solicite um orçamento",
            "peca seu orcamento",
            "peça seu orçamento",
            "compre agora",
            "adquira agora",
            "saiba mais",
            "clique aqui",
            "disponivel para compra",
            "disponível para compra",
            "preco sob consulta",
            "preço sob consulta",
            "cotacao sob consulta",
            "cotação sob consulta"

        ]

        ocorrencias_catalogo = sum(

            1

            for marcador
            in marcadores_catalogo

            if marcador
            in texto_normalizado

        )

        if ocorrencias_catalogo >= 1:
            return False

        # ----------------------------------------------------
        # 09. CABEÇALHOS COMERCIAIS
        # ----------------------------------------------------

        cabecalhos_comerciais = [

            "por que escolher",
            "porque escolher",
            "why choose",
            "our products",
            "our product",
            "fale com nossa equipe"

        ]

        for marcador in cabecalhos_comerciais:

            if marcador in texto_normalizado:
                return False

        # ----------------------------------------------------
        # 10. NOME COMERCIAL EM INGLÊS
        #
        # Não basta o texto conter uma palavra em inglês.
        # A rejeição ocorre quando aparecem marcadores
        # comerciais claros.
        # ----------------------------------------------------

        marcadores_comerciais_ingles = [

            "manufacturer",
            "supplier",
            "manufacturer's",
            "product code",
            "model number",
            "part number",
            "serial number",
            "catalog",
            "catalogue"

        ]

        for marcador in marcadores_comerciais_ingles:

            if marcador in texto_normalizado:
                return False

        return True





# ========================================================
# FILTRO EDITORIAL FINAL DO FRAGMENTO
# ========================================================
#
# ESTA FUNÇÃO FICA DENTRO DE
# selecionar_informacoes_relevantes()
#
# ========================================================

    def fragmento_eh_editorialmente_valido(
        texto,
        identidade_fonte=None
    ):

        texto_original = str(
            texto or ""
        ).strip()

        if not texto_original:
            return False

        texto_normalizado = (
            normalizar_assunto_texto(
                texto_original
            )
        )

        # ====================================================
        # 01. IDENTIDADE COMERCIAL / PRODUTO
        # ====================================================

        if not fragmento_eh_comercialmente_limpo(
            texto_original,
            identidade_fonte
        ):
            return False

        # ====================================================
        # 01.5. CONTAMINAÇÃO EDITORIAL
        # ====================================================

        if not fragmento_eh_aproveitavel_editorialmente(
            texto_original
        ):
            return False

        # ====================================================
        # 02. CRÉDITOS ACADÊMICOS
        # ====================================================

        marcadores_academicos = [

            r"\bprof\.?\s*dr\.?\b",
            r"\bprofessor\s+doutor\b",
            r"\bprofessora\s+doutora\b",
            r"\borientador\b",
            r"\borientadora\b",
            r"\bcoorientador\b",
            r"\bcoorientadora\b",
            r"\bmonografia\b",
            r"\bdissertação\b",
            r"\bdissertacao\b",
            r"\btese\b",
            r"\btrabalho\s+de\s+conclusão\b",
            r"\btrabalho\s+de\s+conclusao\b",
            r"\bbanca\s+examinadora\b",
            r"\buniversidade\b",
            r"\bfaculdade\b",
            r"\bautor\s*:",
            r"\bautora\s*:",
            r"\bautores\s*:",
            r"\bautoras\s*:",
            r"\borientação\s*:",
            r"\borientacao\s*:",
            r"\bcurso\s+de\s+graduação\b",
            r"\bcurso\s+de\s+graduacao\b",
            r"\bdepartamento\s+de\b"
        ]

        for padrao in marcadores_academicos:

            if re.search(
                padrao,
                texto_normalizado,
                re.IGNORECASE
            ):
                return False

        # ====================================================
        # 03. REFERÊNCIAS BIBLIOGRÁFICAS
        # ====================================================

        padroes_referencia = [

            r"\bsegundo\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][\wÁÀÂÃÉÊÍÓÔÕÚÇ-]{2,}",
            r"\bconforme\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][\wÁÀÂÃÉÊÍÓÔÕÚÇ-]{2,}",
            r"\bde\s+acordo\s+com\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ]",
            r"\b[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][\wÁÀÂÃÉÊÍÓÔÕÚÇ-]{2,}\s+et\s+al\.",
            r"\b[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][\wÁÀÂÃÉÊÍÓÔÕÚÇ-]{2,}\s*\(\s*\d{4}\s*\)",
            r"\b[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][\wÁÀÂÃÉÊÍÓÔÕÚÇ-]{2,}\s*,\s*\d{4}\b"
        ]

        for padrao in padroes_referencia:

            if re.search(
                padrao,
                texto_original
            ):
                return False

        # ====================================================
        # 04. CRÉDITOS / AUTORIA DE MATERIAL
        # ====================================================

        padroes_creditos = [

            r"\bfonte\s*:",
            r"\breferência\s*:",
            r"\breferencia\s*:",
            r"\bbibliografia\s*:",
            r"\belaborado\s+por\b",
            r"\bdesenvolvido\s+por\b",
            r"\bproduzido\s+por\b",
            r"\bpesquisa\s+de\b",
            r"\bestudo\s+de\b",
            r"\bcréditos\s*:",
            r"\bcreditos\s*:"

        ]

        for padrao in padroes_creditos:

            if re.search(
                padrao,
                texto_normalizado,
                re.IGNORECASE
            ):
                return False

        # ====================================================
        # 05. CABEÇALHOS DE CATÁLOGO / DATASHEET
        # ====================================================

        marcadores_datasheet = [

            "especificacoes tecnicas",
            "especificações técnicas",
            "especificacion tecnica",
            "especificaciones tecnicas",
            "dados tecnicos",
            "dados técnicos",
            "caracteristicas tecnicas",
            "características técnicas",
            "ficha tecnica",
            "ficha técnica",
            "datasheet",
            "modelo:",
            "model:",
            "model number:",
            "codigo:",
            "código:",
            "part number:",
            "item:",
            "referencia:",
            "referência:"
        ]

        marcadores_datasheet_encontrados = 0

        for marcador in marcadores_datasheet:

            if (
                normalizar_assunto_texto(
                    marcador
                )
                in texto_normalizado
            ):
                marcadores_datasheet_encontrados += 1

        # ====================================================
        # 06. ESTRUTURA DE FICHA TÉCNICA
        # ====================================================

        quantidade_unidades = len(
            re.findall(
                r"\b\d+(?:[.,]\d+)?\s*"
                r"(?:mm|cm|m|kg|g|bar|psi|rpm|"
                r"kw|cv|hp|v|a|hz|°c|ºc|"
                r"mca|l/min|m³/h|m3/h)\b",
                texto_normalizado,
                re.IGNORECASE
            )
        )

        quantidade_dois_pontos = texto_original.count(":")

        quantidade_numeros = len(
            re.findall(
                r"\b\d+(?:[.,]\d+)?\b",
                texto_original
            )
        )

        if (
            (
                quantidade_unidades >= 3
                and
                quantidade_numeros >= 4
                and
                quantidade_dois_pontos >= 2
            )
            or
            (
                marcadores_datasheet_encontrados >= 2
                and
                quantidade_numeros >= 3
            )
        ):
            return False

        # ====================================================
        # 07. TABELA / LISTA DE ESPECIFICAÇÕES
        # ====================================================

        linhas_com_campos = len(
            re.findall(
                r"(?:^|\s)"
                r"[A-Za-zÁÀÂÃÉÊÍÓÔÕÚÇáàâãéêíóôõúç"
                r" _-]{3,40}"
                r"\s*:\s*"
                r"[^\n]{1,80}",
                texto_original,
                re.MULTILINE
            )
        )

        if (
            linhas_com_campos >= 4
            and
            quantidade_numeros >= 4
        ):
            return False

        # ====================================================
        # 08. TEXTO COMERCIAL / PROMOCIONAL
        # ====================================================

        marcadores_promocionais = [

            "bom fornecedor",
            "melhor fornecedor",
            "fornecedor confiavel",
            "fornecedor confiável",
            "empresa expert",
            "somos especialistas",
            "somos especialista",
            "entre em contato",
            "fale conosco",
            "solicite um orçamento",
            "solicite um orcamento",
            "compre agora",
            "adquira agora",
            "saiba mais",
            "clique aqui"
        ]

        for marcador in marcadores_promocionais:

            if (
                normalizar_assunto_texto(
                    marcador
                )
                in texto_normalizado
            ):
                return False

        # ====================================================
        # 09. TEXTO DE PROCEDIMENTO / AVISO BRUTO
        # ====================================================

        marcadores_manual = [

            "utilize epi",
            "utilize epis",
            "use epi",
            "use epis",
            "atenção:",
            "atencao:",
            "advertência:",
            "advertencia:",
            "perigo:",
            "risco:",
            "não utilize",
            "nao utilize",
            "desligue o equipamento",
            "desligue a bomba",
        ]

        ocorrencias_manual = 0

        for marcador in marcadores_manual:

            if (
                normalizar_assunto_texto(
                    marcador
                )
                in texto_normalizado
            ):
                ocorrencias_manual += 1

        if ocorrencias_manual >= 1:
            return False

        # ====================================================
        # 10. IDIOMA ESTRANGEIRO / CATÁLOGO BRUTO
        # ====================================================

        # Fontes em outros idiomas não são rejeitadas apenas pelo idioma.
        # A aderência temática e os filtros de identidade, catálogo e
        # contaminação continuam ativos; o prompt orienta a interpretação
        # da evidência e a redação final exclusivamente em PT-BR.

        # ====================================================
        # 11. NOME DE PESSOA COMO CRÉDITO
        # ====================================================

        padroes_autoria = [

            r"\bautor(?:a|es|as)?\s*[:\-]",
            r"\bpor\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+"
            r"\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+",
            r"\belaborado\s+por\b",
            r"\bdesenvolvido\s+por\b"
        ]

        for padrao in padroes_autoria:

            if re.search(
                padrao,
                texto_original
            ):
                return False

        # ====================================================
        # 12. APROVADO
        # ====================================================

        return True



    # ========================================================
    # CONSTRUIR FRAGMENTOS
    # ========================================================
    #
    # ARQUITETURA:
    #
    # texto_original
    #       ↓
    # cópia usada somente para análise
    #       ↓
    # identificação das frases / palavras
    #       ↓
    # posição inicial e final no texto original
    #       ↓
    # candidato["texto"] = trecho original
    #
    # IMPORTANTE:
    # O Python pode analisar, contar, pontuar e filtrar.
    # Porém, o texto salvo no candidato deve continuar sendo
    # o trecho existente na fonte original.
    #
    # Não usar:
    #     " ".join(...)
    #
    # para reconstruir o conteúdo final.
    # ========================================================

    
    # ========================================================
    # CONTROLE DO FILTRO DE QUALIDADE
    # ========================================================
    
    estatisticas_qualidade = {
        "avaliados": 0,
        "qualidade_estrutural": 0,
        "selecionados": 0
    }

    # Diagnóstico detalhado por fonte.
    # Não altera a seleção; serve para identificar em qual etapa
    # cada fonte está sendo descartada.
    diagnostico_fontes = {}
    
    
    for fonte in fontes:

        candidatos_antes_fonte = len(candidatos)

        diagnostico_fonte = {
            "segmentos": 0,
            "avaliados": 0,
            "tema_aprovado": 0,
            "tema_rejeitado": 0,
            "editorial_rejeitado": 0,
            "fallback_aprovado": 0,
            "candidatos_salvos": 0,
            "motivos": {}
        }

        diagnostico_fontes[fonte.get("indice", "")] = diagnostico_fonte

        # ----------------------------------------------------
        # TEXTO ORIGINAL
        # ----------------------------------------------------
        #
        # Este texto não será reconstruído nem normalizado.
        # Ele será utilizado posteriormente para extrair
        # exatamente o trecho selecionado.
        # ----------------------------------------------------

        texto_original = str(
            fonte.get("texto", "")
        )

        if not texto_original.strip():
            continue


        # ----------------------------------------------------
        # TEXTO PARA ANÁLISE
        # ----------------------------------------------------
        #
        # Esta cópia pode ser modificada livremente.
        # Ela serve somente para localizar frases e palavras.
        #
        # O conteúdo final do candidato NÃO será obtido dela.
        # ----------------------------------------------------

        texto_para_analise = re.sub(
            r"\s+",
            " ",
            texto_original
        ).strip()

        if not texto_para_analise:
            continue


        # ----------------------------------------------------
        # LOCALIZAR FRASES
        # ----------------------------------------------------
        #
        # Mantemos também a posição da frase dentro do texto
        # de análise.
        #
        # A posição será convertida para o texto original
        # através da mesma sequência de palavras.
        # ----------------------------------------------------

        frases_analise = []

        for correspondencia in re.finditer(
            r".*?(?<=[.!?])(?=\s|$)",
            texto_para_analise,
            re.DOTALL
        ):

            inicio_frase = (
                correspondencia.start()
            )

            fim_frase = (
                correspondencia.end()
            )

            frase = texto_para_analise[
                inicio_frase:fim_frase
            ].strip()

            if not frase:
                continue

            frases_analise.append({
                "inicio": inicio_frase,
                "fim": fim_frase,
                "texto": frase
            })


        diagnostico_fonte["segmentos"] = len(frases_analise)

        # ----------------------------------------------------
        # TEXTO SEM PONTUAÇÃO FINAL
        # ----------------------------------------------------

        if not frases_analise:

            frases_analise = [{
                "inicio": 0,
                "fim": len(
                    texto_para_analise
                ),
                "texto": texto_para_analise
            }]

        else:

            ultimo_fim = frases_analise[-1]["fim"]

            restante = texto_para_analise[
                ultimo_fim:
            ].strip()

            if restante:

                inicio_restante = (
                    texto_para_analise.find(
                        restante,
                        ultimo_fim
                    )
                )

                if inicio_restante >= 0:

                    frases_analise.append({
                        "inicio":
                            inicio_restante,

                        "fim":
                            inicio_restante
                            + len(restante),

                        "texto":
                            restante
                    })

        diagnostico_fonte["segmentos"] = len(frases_analise)

        # ----------------------------------------------------
        # MAPEAR PALAVRAS DA ANÁLISE
        # PARA O TEXTO ORIGINAL
        # ----------------------------------------------------
        #
        # A análise pode ter espaços normalizados.
        # Por isso não podemos simplesmente utilizar os
        # mesmos índices de caracteres.
        #
        # Criamos uma correspondência sequencial entre as
        # palavras da cópia de análise e as palavras existentes
        # no texto original.
        # ----------------------------------------------------

        palavras_original = list(
            re.finditer(
                r"\S+",
                texto_original
            )
        )

        palavras_analise = list(
            re.finditer(
                r"\S+",
                texto_para_analise
            )
        )

        if not palavras_original:
            continue

        if not palavras_analise:
            continue


        # ----------------------------------------------------
        # CORRESPONDÊNCIA DAS PALAVRAS
        # ----------------------------------------------------
        #
        # Normalizamos somente para comparação.
        # O texto original permanece intacto.
        # ----------------------------------------------------

        quantidade_palavras = min(
            len(palavras_original),
            len(palavras_analise)
        )

        mapa_palavras = []

        for indice_palavra in range(
            quantidade_palavras
        ):

            palavra_analise = (
                palavras_analise[
                    indice_palavra
                ].group()
            )

            palavra_original = (
                palavras_original[
                    indice_palavra
                ].group()
            )

            mapa_palavras.append({
                "analise_inicio":
                    palavras_analise[
                        indice_palavra
                    ].start(),

                "analise_fim":
                    palavras_analise[
                        indice_palavra
                    ].end(),

                "original_inicio":
                    palavras_original[
                        indice_palavra
                    ].start(),

                "original_fim":
                    palavras_original[
                        indice_palavra
                    ].end(),

                "palavra_analise":
                    palavra_analise,

                "palavra_original":
                    palavra_original
            })


        
    # ========================================================
        # PERTENCIMENTO REAL AO TEMA
        # ========================================================
        #
        # Um fragmento somente pode participar da seleção se
        # realmente tratar do tema pesquisado.
        #
        # A pontuação editorial NÃO pode compensar a ausência
        # do tema.
        #
        # Exemplo:
        #
        # tema:
        #     bomba centrífuga
        #
        # fragmento:
        #     bomba submersível solar apresenta...
        #
        # Mesmo contendo "bomba", "aplicação", "desempenho",
        # etc., o fragmento deve ser rejeitado.
        #
        # ========================================================

        def fragmento_pertence_ao_tema(
            texto,
            tema
        ):

            texto = str(
                texto or ""
            ).strip()

            tema = str(
                tema or ""
            ).strip()

            if not texto or not tema:
                return False


            # ----------------------------------------------------
            # NORMALIZAÇÃO
            # ----------------------------------------------------

            texto_normalizado = (
                normalizar_assunto_texto(
                    texto
                )
            )

            tema_normalizado = (
                normalizar_assunto_texto(
                    tema
                )
            )


            # ----------------------------------------------------
            # PALAVRAS SIGNIFICATIVAS DO TEMA
            # ----------------------------------------------------

            palavras_tema = re.findall(
                r"\b[a-z0-9]{3,}\b",
                tema_normalizado
            )


            # Remove termos gramaticais.
            palavras_ignoradas = {

                "de",
                "da",
                "das",
                "do",
                "dos",
                "em",
                "na",
                "nas",
                "no",
                "nos",
                "para",
                "por",
                "com",
                "sem",
                "e",
                "a",
                "o",
                "as",
                "os"

            }


            palavras_tema = [

                palavra

                for palavra
                in palavras_tema

                if palavra
                not in palavras_ignoradas

            ]


            if not palavras_tema:
                return False


            # ----------------------------------------------------
            # PALAVRAS DO FRAGMENTO
            # ----------------------------------------------------

            palavras_fragmento = set(

                re.findall(
                    r"\b[a-z0-9]{3,}\b",
                    texto_normalizado
                )

            )

            # Equivalências lexicais multilíngues conservadoras para que
            # fontes estrangeiras relevantes não sejam descartadas antes
            # de o Ollama poder interpretá-las e redigir em PT-BR.
            equivalencias_tema = {
                "chinelo": {"flip", "flop", "flipflop", "sandals", "sandal"},
                "borracha": {"rubber", "latex", "elastomer"},
                "bomba": {"pump", "pumps"},
                "centrifuga": {"centrifugal"},
                "agua": {"water"},
                "valvula": {"valve", "valves"},
                "tubulacao": {"pipe", "piping", "pipeline"},
                "vedacao": {"seal", "sealing", "gasket"},
                "rolamento": {"bearing", "bearings"},
                "motor": {"motor", "engine"},
                "compressor": {"compressor", "compressors"},
                "filtro": {"filter", "filters"},
                "mangueira": {"hose", "hoses"},
                "conexao": {"fitting", "fittings", "connection"},
                "aco": {"steel"},
                "inox": {"stainless"},
                "plastico": {"plastic"},
                "industrial": {"industrial"},
            }
            for palavra_tema in palavras_tema:
                radical_tema = palavra_tema.rstrip("s")
                sinonimos = set(equivalencias_tema.get(palavra_tema, set()))
                sinonimos.update(equivalencias_tema.get(radical_tema, set()))
                # Só considera representada a palavra correspondente; não
                # permite que um único sinônimo valide todo o tema composto.
                if sinonimos.intersection(palavras_fragmento):
                    palavras_fragmento.add(palavra_tema)


            # ----------------------------------------------------
            # COMPARAÇÃO POR RADICAL SIMPLES
            #
            # Permite:
            #
            # bomba       → bombas
            # centrifug   → centrifuga
            # centrifug   → centrifugas
            #
            # Sem exigir que o texto tenha exatamente a mesma
            # flexão usada na palavra-chave.
            # ----------------------------------------------------

            for palavra_tema in palavras_tema:

                radical = palavra_tema

                if len(radical) >= 5:

                    # Remove terminações flexionais simples.
                    if radical.endswith("es"):
                        radical = radical[:-2]

                    elif radical.endswith("s"):
                        radical = radical[:-1]

                    elif radical.endswith("a"):
                        radical = radical[:-1]

                    elif radical.endswith("o"):
                        radical = radical[:-1]


                encontrou = False


                for palavra_fragmento in palavras_fragmento:

                    if palavra_fragmento.startswith(
                        radical
                    ):

                        encontrou = True

                        break


                # ------------------------------------------------
                # TODAS as palavras significativas do tema
                # precisam estar representadas no fragmento.
                # ------------------------------------------------

                if not encontrou:

                    return False


            return True
        
        
        # ----------------------------------------------------
        # CRIAR CANDIDATO A PARTIR DA POSIÇÃO ORIGINAL
        # ----------------------------------------------------

        def adicionar_candidato_por_palavras(
            indice_inicio,
            indice_fim
        ):

            nonlocal estatisticas_qualidade

            if (
                indice_inicio < 0
                or
                indice_fim <= indice_inicio
                or
                indice_inicio >= len(mapa_palavras)
                or
                indice_fim > len(mapa_palavras)
            ):
                return


            # ------------------------------------------------
            # POSIÇÃO REAL NO TEXTO ORIGINAL
            # ------------------------------------------------

            inicio_original = (
                mapa_palavras[
                    indice_inicio
                ]["original_inicio"]
            )

            fim_original = (
                mapa_palavras[
                    indice_fim - 1
                ]["original_fim"]
            )


            # ------------------------------------------------
            # EXTRAIR DO TEXTO ORIGINAL
            # ------------------------------------------------
            #
            # Aqui está a correção principal.
            #
            # O candidato não é reconstruído.
            # Ele é recortado diretamente da fonte original.
            # ------------------------------------------------

            texto_fragmento = (
                texto_original[
                    inicio_original:fim_original
                ]
            ).strip()

            if not texto_fragmento:
                return


            quantidade = len(
                re.findall(
                    r"\S+",
                    texto_fragmento
                )
            )

            if (
                quantidade
                <
                MIN_PALAVRAS_FRAGMENTO
                or
                quantidade
                >
                MAX_PALAVRAS_FRAGMENTO
            ):
                return
            
            
        
            # ------------------------------------------------
            # ------------------------------------------------
            # CONTABILIZAR TRECHO PARA AUDITORIA
            # ------------------------------------------------

            estatisticas_qualidade[
                "avaliados"
            ] += 1
            diagnostico_fonte["avaliados"] += 1

            # IDENTIFICAR IDENTIDADE DA FONTE
            # ------------------------------------------------
        
            identidade_fonte = fonte.get(
                "identidade_fonte",
                {}
            )
        
            if not isinstance(identidade_fonte, dict):
                identidade_fonte = {}
        
            if not identidade_fonte:
                identidade_fonte = identificar_empresa_fonte(
                    texto_fragmento,
                    fonte.get("url", "")
                )
            
            

            # ------------------------------------------------
            # ADERÊNCIA TEMÁTICA OBRIGATÓRIA
            # ------------------------------------------------
            # Todo candidato, inclusive os gerados pela extração
            # normal, precisa demonstrar que trata do tema.
            # Antes isso era verificado apenas no fallback, o que
            # tornava a extração inconsistente entre as fontes.
            # ------------------------------------------------

            if not fragmento_pertence_ao_tema(
                texto_fragmento,
                tema
            ):
                diagnostico_fonte["tema_rejeitado"] += 1
                return

            diagnostico_fonte["tema_aprovado"] += 1

            # ------------------------------------------------
            # FILTRO EDITORIAL
            # ------------------------------------------------
            #
            # O fragmento precisa ser tecnicamente aproveitável,
            # mas não podemos eliminar todo o patrimônio antes
            # da etapa de seleção.
            #
            # A seleção posterior ainda aplica:
            # - qualidade estrutural;
            # - diversidade;
            # - repetição;
            # - pontuação editorial.
            #
            # ------------------------------------------------
        
            editorialmente_valido = (
                fragmento_eh_editorialmente_valido(
                    texto_fragmento,
                    identidade_fonte
                )
            )
        
            if not editorialmente_valido:
        
                estatisticas_qualidade[
                    "qualidade_estrutural"
                ] += 1
                diagnostico_fonte["editorial_rejeitado"] += 1
        
                # ------------------------------------------------
                # FALLBACK CONTROLADO
                # ------------------------------------------------
                #
                # Se o filtro editorial rejeitar o trecho, fazemos
                # uma segunda verificação mais simples.
                #
                # O objetivo NÃO é aceitar lixo.
                #
                # O trecho ainda precisa:
                # - ter tamanho correto;
                # - pertencer ao tema;
                # - não conter URL;
                # - não conter telefone;
                # - não conter e-mail;
                # - não conter código comercial evidente.
                #
                # ------------------------------------------------
        
                texto_normalizado_fallback = (
                    normalizar_assunto_texto(
                        texto_fragmento
                    )
                )
        
                if not texto_normalizado_fallback:
                    return
        
                # -----------------------------------------------
                # URL
                # -----------------------------------------------
        
                if re.search(
                    r"(https?://|www\.)\S+",
                    texto_fragmento,
                    re.IGNORECASE
                ):
                    return
        
                # -----------------------------------------------
                # E-MAIL
                # -----------------------------------------------
        
                if re.search(
                    r"\b[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}\b",
                    texto_fragmento,
                    re.IGNORECASE
                ):
                    return
        
                # -----------------------------------------------
                # TELEFONE
                # -----------------------------------------------
        
                if re.search(
                    r"(?<!\d)"
                    r"(?:\+?55[\s.-]*)?"
                    r"\(?\d{2}\)?[\s.-]*"
                    r"\d{4,5}[\s.-]*\d{4}"
                    r"(?!\d)",
                    texto_fragmento
                ):
                    return
        
                # -----------------------------------------------
                # CNPJ
                # -----------------------------------------------
        
                if re.search(
                    r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b",
                    texto_fragmento
                ):
                    return
        
                # -----------------------------------------------
                # CÓDIGOS COMERCIAIS EVIDENTES
                # -----------------------------------------------
        
                codigo_comercial = re.search(
                    r"\b(?:SKU|MPN|"
                    r"PART[\s-]*NUMBER|"
                    r"SERIAL[\s-]*NUMBER|"
                    r"MODELO\s*[:\-]|"
                    r"C[ÓO]DIGO\s+(?:DO\s+)?PRODUTO|"
                    r"REFER[ÊE]NCIA\s+(?:DO\s+)?PRODUTO)"
                    r"\b",
                    texto_fragmento,
                    re.IGNORECASE
                )
        
                if codigo_comercial:
                    return
        
                # -----------------------------------------------
                # NOME DA EMPRESA DA FONTE
                # -----------------------------------------------
                # Em PDFs, manuais e datasheets o fabricante pode
                # aparecer no cabeçalho/rodapé de todos os trechos.
                # Isso, sozinho, não torna o conteúdo promocional.
                # A versão anterior descartava o trecho inteiro por
                # essa coincidência e podia eliminar quase todo o
                # patrimônio técnico.
                #
                # Portanto, o nome da empresa não é mais motivo de
                # rejeição nesta validação mínima. Os filtros de CTA,
                # contato, URL, telefone, CNPJ e códigos comerciais
                # continuam ativos.
                # -----------------------------------------------

                # -----------------------------------------------
                # FALLBACK APROVADO
                # -----------------------------------------------
        
                # Fallback registrado em contador por fonte; não imprimir
                # uma mensagem para cada candidato (log excessivo).


        
            # ------------------------------------------------
            # SALVAR CANDIDATO
            # ------------------------------------------------

            diagnostico_fonte["candidatos_salvos"] += 1
        
            candidatos.append({
        
                "texto":
                    texto_fragmento,
        
                "fonte":
                    fonte["indice"],
        
                "url":
                    fonte["url"],
        
                "tipo":
                    fonte["tipo"],
        
                "pdf":
                    fonte["eh_pdf"],
        
                "palavras":
                    quantidade,
        
                "identidade_fonte":
                    identidade_fonte
        
            })


        # ----------------------------------------------------
        # CONVERTER FRASES EM INTERVALOS DE PALAVRAS
        # ----------------------------------------------------

        intervalos_frases = []

        for frase in frases_analise:

            palavras_da_frase = [

                indice

                for indice, palavra
                in enumerate(mapa_palavras)

                if (
                    palavra["analise_inicio"]
                    >=
                    frase["inicio"]

                    and

                    palavra["analise_fim"]
                    <=
                    frase["fim"]
                )

            ]

            if not palavras_da_frase:
                continue

            intervalos_frases.append({

                "inicio":
                    palavras_da_frase[0],

                "fim":
                    palavras_da_frase[-1] + 1

            })


        # ----------------------------------------------------
        # ACUMULAR FRASES
        # ----------------------------------------------------

        acumulado_inicio = None
        acumulado_fim = None
        palavras_acumuladas = 0


        for intervalo in intervalos_frases:

            inicio = intervalo["inicio"]
            fim = intervalo["fim"]

            quantidade = (
                fim - inicio
            )


            # ------------------------------------------------
            # FRASE PEQUENA / MÉDIA
            # ------------------------------------------------

            if quantidade <= MAX_PALAVRAS_FRAGMENTO:

                if acumulado_inicio is None:

                    acumulado_inicio = inicio
                    acumulado_fim = fim
                    palavras_acumuladas = quantidade

                elif (
                    palavras_acumuladas
                    + quantidade
                    <= MAX_PALAVRAS_FRAGMENTO
                ):

                    acumulado_fim = fim

                    palavras_acumuladas += (
                        quantidade
                    )

                else:

                    if (
                        MIN_PALAVRAS_FRAGMENTO
                        <=
                        palavras_acumuladas
                        <=
                        MAX_PALAVRAS_FRAGMENTO
                    ):

                        adicionar_candidato_por_palavras(
                            acumulado_inicio,
                            acumulado_fim
                        )

                    acumulado_inicio = inicio
                    acumulado_fim = fim
                    palavras_acumuladas = quantidade


                # --------------------------------------------
                # CANDIDATO PRONTO
                # --------------------------------------------

                if (
                    MIN_PALAVRAS_FRAGMENTO
                    <=
                    palavras_acumuladas
                    <=
                    MAX_PALAVRAS_FRAGMENTO
                ):

                    adicionar_candidato_por_palavras(
                        acumulado_inicio,
                        acumulado_fim
                    )

                    acumulado_inicio = None
                    acumulado_fim = None
                    palavras_acumuladas = 0


            # ------------------------------------------------
            # FRASE GRANDE
            # ------------------------------------------------

            else:

                # --------------------------------------------
                # SALVAR ACUMULADO ANTERIOR
                # --------------------------------------------

                if (
                    acumulado_inicio is not None
                    and
                    MIN_PALAVRAS_FRAGMENTO
                    <=
                    palavras_acumuladas
                    <=
                    MAX_PALAVRAS_FRAGMENTO
                ):

                    adicionar_candidato_por_palavras(
                        acumulado_inicio,
                        acumulado_fim
                    )


                acumulado_inicio = None
                acumulado_fim = None
                palavras_acumuladas = 0


                # --------------------------------------------
                # DIVIDIR FRASE GRANDE
                # --------------------------------------------

                inicio_parte = inicio

                while inicio_parte < fim:

                    fim_parte = min(
                        inicio_parte
                        + MAX_PALAVRAS_FRAGMENTO,
                        fim
                    )

                    quantidade_parte = (
                        fim_parte
                        -
                        inicio_parte
                    )

                    # ----------------------------------------
                    # PARTE COMPLETA
                    # ----------------------------------------

                    if (
                        MIN_PALAVRAS_FRAGMENTO
                        <=
                        quantidade_parte
                        <=
                        MAX_PALAVRAS_FRAGMENTO
                    ):

                        adicionar_candidato_por_palavras(
                            inicio_parte,
                            fim_parte
                        )

                    # ----------------------------------------
                    # PARTE FINAL PEQUENA
                    # ----------------------------------------
                    #
                    # Não descartamos imediatamente.
                    # Guardamos para tentar combinar com a
                    # próxima frase.
                    # ----------------------------------------

                    elif (
                        quantidade_parte
                        <
                        MIN_PALAVRAS_FRAGMENTO
                    ):

                        acumulado_inicio = (
                            inicio_parte
                        )

                        acumulado_fim = (
                            fim_parte
                        )

                        palavras_acumuladas = (
                            quantidade_parte
                        )

                    inicio_parte = fim_parte


        # ----------------------------------------------------
        # ÚLTIMO ACUMULADO
        # ----------------------------------------------------

        if (
            acumulado_inicio is not None
            and
            MIN_PALAVRAS_FRAGMENTO
            <=
            palavras_acumuladas
            <=
            MAX_PALAVRAS_FRAGMENTO
        ):

            adicionar_candidato_por_palavras(
                acumulado_inicio,
                acumulado_fim
            )




            # ----------------------------------------------------
            # FALLBACK DE EXTRAÇÃO POR JANELA TEMÁTICA
            # ----------------------------------------------------
            #
            # Alguns PDFs e páginas chegam sem pontuação adequada
            # para a segmentação por frases. Nesses casos, o texto
            # pode ser tecnicamente válido, mas não gerar nenhum
            # intervalo entre 35 e 75 palavras.
            #
            # Se a fonte não produziu nenhum candidato pelo método
            # normal, procuramos uma janela CONTÍGUA no texto original
            # ao redor dos termos do tema e reutilizamos a mesma
            # rotina de validação do candidato.
            #
            # O conteúdo continua sendo recortado da fonte original;
            # não há reescrita nem criação de texto.
            # ----------------------------------------------------

            # ----------------------------------------------------
            # DIAGNÓSTICO: FALLBACK SOMENTE QUANDO A FONTE NÃO
            # GEROU NENHUM CANDIDATO PELO MÉTODO NORMAL.
            # ----------------------------------------------------

            candidatos_antes_fallback = len(candidatos)

            # Se a fonte produziu poucos candidatos, tentamos uma
            # segunda extração independente da segmentação por frases.
            # Isso é importante para PDFs, OCR, catálogos e páginas
            # com pontuação/estrutura irregular.
            if len(candidatos) - candidatos_antes_fonte < 3:

                tema_normalizado_fallback = normalizar_assunto_texto(
                    tema
                )

                termos_tema_fallback = [
                    termo
                    for termo in re.findall(
                        r"\b[a-z0-9]{3,}\b",
                        tema_normalizado_fallback
                    )
                    if len(termo) >= 4
                ]

                print(
                    "FALLBACK TEMÁTICO FONTE",
                    fonte.get("indice", ""),
                    "| TERMOS:",
                    termos_tema_fallback
                )

                if termos_tema_fallback and mapa_palavras:

                    # Localiza cada termo no mapa de palavras.
                    ocorrencias_por_termo = {}

                    for termo in termos_tema_fallback:
                        ocorrencias = []

                        for indice_palavra, item_palavra in enumerate(mapa_palavras):
                            palavra_normalizada = normalizar_assunto_texto(
                                item_palavra.get("palavra_analise", "")
                            )

                            if palavra_normalizada.startswith(termo):
                                ocorrencias.append(indice_palavra)

                        ocorrencias_por_termo[termo] = ocorrencias

                    termos_encontrados_no_texto = [
                        termo
                        for termo, ocorrencias in ocorrencias_por_termo.items()
                        if ocorrencias
                    ]

                    print(
                        "TERMOS ENCONTRADOS FONTE",
                        fonte.get("indice", ""),
                        ":",
                        termos_encontrados_no_texto
                    )

                    # Para tema composto, procura pares/grupos de ocorrências
                    # próximos. Para tema simples, usa cada ocorrência.
                    janelas = []

                    if len(termos_tema_fallback) == 1:
                        bases = [(i, i) for i in ocorrencias_por_termo.get(termos_tema_fallback[0], [])]
                    else:
                        bases = []
                        primeira = ocorrencias_por_termo.get(termos_tema_fallback[0], [])

                        for indice_primeiro in primeira:
                            grupo = [indice_primeiro]

                            for termo in termos_tema_fallback[1:]:
                                ocorrencias = ocorrencias_por_termo.get(termo, [])
                                if not ocorrencias:
                                    grupo = []
                                    break

                                proximo = min(
                                    ocorrencias,
                                    key=lambda x: abs(x - indice_primeiro)
                                )

                                if abs(proximo - indice_primeiro) <= 90:
                                    grupo.append(proximo)
                                else:
                                    grupo = []
                                    break

                            if grupo:
                                bases.append((min(grupo), max(grupo)))

                    for inicio_tema, fim_tema in bases:

                        # Janela contígua sempre recortada da fonte original.
                        inicio_janela = max(
                            0,
                            inicio_tema - 22
                        )

                        fim_janela = min(
                            len(mapa_palavras),
                            fim_tema + 48
                        )

                        # Garante o mínimo sem ultrapassar o máximo.
                        if fim_janela - inicio_janela < MIN_PALAVRAS_FRAGMENTO:
                            fim_janela = min(
                                len(mapa_palavras),
                                inicio_janela + MAX_PALAVRAS_FRAGMENTO
                            )

                        if fim_janela - inicio_janela > MAX_PALAVRAS_FRAGMENTO:
                            fim_janela = inicio_janela + MAX_PALAVRAS_FRAGMENTO

                        if (
                            MIN_PALAVRAS_FRAGMENTO
                            <= fim_janela - inicio_janela
                            <= MAX_PALAVRAS_FRAGMENTO
                        ):
                            chave_janela = (inicio_janela, fim_janela)
                            if chave_janela not in janelas:
                                janelas.append(chave_janela)

                        if len(janelas) >= 5:
                            break

                    for inicio_janela, fim_janela in janelas:
                        adicionar_candidato_por_palavras(
                            inicio_janela,
                            fim_janela
                        )

            # ----------------------------------------------------
            # FALLBACK V7.4 — RELEVÂNCIA CONTEXTUAL DA FONTE
            # ----------------------------------------------------
            #
            # Alguns textos técnicos não repetem todos os termos do tema
            # dentro de cada janela de 35–75 palavras. Exigir que
            # "bomba" e "centrifuga" apareçam simultaneamente na mesma
            # janela pode eliminar conteúdo factual legítimo.
            #
            # Primeiro verificamos se a FONTE, como um todo, demonstra
            # relação inequívoca com o tema. Depois extraímos janelas
            # ao redor dos termos encontrados, mantendo o texto original.
            # Isso NÃO autoriza uma fonte irrelevante: a fonte precisa
            # conter a expressão composta ou todos os termos significativos.
            # ----------------------------------------------------

            if len(candidatos) - candidatos_antes_fonte < 3 and mapa_palavras:

                tema_normalizado_contextual = normalizar_assunto_texto(tema)
                termos_contextuais = [
                    termo for termo in re.findall(
                        r"\b[a-z0-9]{3,}\b",
                        tema_normalizado_contextual
                    )
                    if termo not in {
                        "de", "da", "das", "do", "dos", "em", "na", "nas",
                        "no", "nos", "para", "por", "com", "sem", "e", "a",
                        "o", "as", "os"
                    }
                ]

                texto_fonte_normalizado = normalizar_assunto_texto(texto_original)
                frase_tema_exata = " ".join(termos_contextuais)

                ocorrencias_contextuais = []
                for indice_palavra, item_palavra in enumerate(mapa_palavras):
                    palavra_norm = normalizar_assunto_texto(
                        item_palavra.get("palavra_analise", "")
                    )
                    if any(
                        palavra_norm.startswith(termo)
                        for termo in termos_contextuais
                    ):
                        ocorrencias_contextuais.append(indice_palavra)

                fonte_tem_relacao_contextual = False

                if len(termos_contextuais) == 1:
                    fonte_tem_relacao_contextual = bool(ocorrencias_contextuais)
                elif frase_tema_exata and frase_tema_exata in texto_fonte_normalizado:
                    fonte_tem_relacao_contextual = True
                else:
                    termos_presentes_fonte = {
                        termo
                        for termo in termos_contextuais
                        if any(
                            normalizar_assunto_texto(
                                item_palavra.get("palavra_analise", "")
                            ).startswith(termo)
                            for item_palavra in mapa_palavras
                        )
                    }
                    fonte_tem_relacao_contextual = (
                        len(termos_presentes_fonte) == len(termos_contextuais)
                    )

                if fonte_tem_relacao_contextual and ocorrencias_contextuais:

                    print(
                        "FALLBACK CONTEXTUAL V7.4 FONTE",
                        fonte.get("indice", ""),
                        "| RELAÇÃO COM TEMA: SIM",
                        "| OCORRÊNCIAS:",
                        len(ocorrencias_contextuais)
                    )

                    janelas_contextuais = []

                    for centro in ocorrencias_contextuais:
                        inicio_janela = max(0, centro - 18)
                        fim_janela = min(
                            len(mapa_palavras),
                            inicio_janela + MAX_PALAVRAS_FRAGMENTO
                        )

                        if fim_janela - inicio_janela < MIN_PALAVRAS_FRAGMENTO:
                            inicio_janela = max(
                                0,
                                fim_janela - MIN_PALAVRAS_FRAGMENTO
                            )

                        chave = (inicio_janela, fim_janela)
                        if chave not in janelas_contextuais:
                            janelas_contextuais.append(chave)

                        if len(janelas_contextuais) >= 8:
                            break

                    antes_contextual = len(candidatos)
                    for inicio_janela, fim_janela in janelas_contextuais:
                        adicionar_candidato_por_palavras(
                            inicio_janela,
                            fim_janela
                        )

                    adicionados_contextual = len(candidatos) - antes_contextual
                    diagnostico_fonte["fallback_aprovado"] += max(0, adicionados_contextual)

                else:
                    print(
                        "FALLBACK CONTEXTUAL V7.4 FONTE",
                        fonte.get("indice", ""),
                        "| RELAÇÃO COM TEMA: NÃO"
                    )

            # ----------------------------------------------------
            # DIAGNÓSTICO POR FONTE
            # ----------------------------------------------------

            candidatos_depois_fallback = len(candidatos)


        # ========================================================
    # 05. REMOVER DUPLICADOS
    # ========================================================

    candidatos_unicos = []

    fragmentos_vistos = set()

    for candidato in candidatos:

        texto = candidato[
            "texto"
        ].strip()

        chave = normalizar_texto_hash(
            texto
        )

        if not chave:
            continue

        if chave in fragmentos_vistos:
            continue

        fragmentos_vistos.add(
            chave
        )

        # ----------------------------------------------------
        # ID + HASH DO TRECHO
        # ----------------------------------------------------

        hash_trecho = gerar_hash_trecho(
            texto
        )

        id_trecho = gerar_id_trecho(
            hash_trecho
        )

        candidato["id"] = id_trecho

        candidato["hash"] = hash_trecho

        ocorrencias_chave = contar_palavra_chave_literal(
            texto, tema
        )
        candidato["palavra_chave_status"] = (
            "COM" if ocorrencias_chave >= 1 else "SEM"
        )
        candidato["palavra_chave_ocorrencias_fonte"] = ocorrencias_chave

        candidatos_unicos.append(
            candidato
        )

    candidatos = candidatos_unicos

    print()
    print("==============================")
    print("DIAGNÓSTICO DOS CANDIDATOS")
    print("==============================")
    
    print(
        "CANDIDATOS ÚNICOS:",
        len(candidatos)
    )
    
    print(
        "FONTES RECEBIDAS:",
        len(fontes)
    )
    
    print(
        "TRECHOS AVALIADOS:",
        estatisticas_qualidade["avaliados"]
    )
    
    print(
        "REJEITADOS PELO FILTRO ESTRUTURAL:",
        estatisticas_qualidade["qualidade_estrutural"]
    )
    
    print(
        "CANDIDATOS APROVEITADOS:",
        len(candidatos)
    )

    candidatos_com_chave = sum(
        1 for c in candidatos
        if contar_palavra_chave_literal(c.get("texto", ""), tema) >= 1
    )
    candidatos_sem_chave = len(candidatos) - candidatos_com_chave
    print("HASHES COM PALAVRA-CHAVE LITERAL:", candidatos_com_chave)
    print("HASHES SEM PALAVRA-CHAVE LITERAL:", candidatos_sem_chave)

    # Diagnóstico compacto da distribuição antes da seleção.
    candidatos_por_fonte_diagnostico = {}
    for candidato in candidatos:
        fonte_id = candidato.get("fonte", "")
        candidatos_por_fonte_diagnostico[fonte_id] = (
            candidatos_por_fonte_diagnostico.get(fonte_id, 0) + 1
        )

    print("FONTES COM CANDIDATOS:", len(candidatos_por_fonte_diagnostico),
          "| CANDIDATOS ÚNICOS:", len(candidatos))
    
    print("==============================")
    
    
    # ========================================================
    # 06. PRIORIZAR PELOS ASSUNTOS DOS CHECKBOXES
    #     + CONTEXTO EDITORIAL DOS 5 BLOCOS DO MEAD
    # ========================================================

    if not isinstance(
        estrutura_editorial,
        dict
    ):

        estrutura_editorial = {}

    # --------------------------------------------------------
    # CHECKBOXES ATIVOS
    # --------------------------------------------------------

    assuntos = [

        str(chave).replace(
            "_",
            " "
        )

        for chave, valor
        in estrutura_editorial.items()

        if valor

    ]
    
    # ============================================================
    # CATEGORIAS — ORIENTAÇÃO ADICIONAL PARA SELEÇÃO
    # As categorias apenas ajudam a pontuar os candidatos.
    # Não alteram a quantidade nem a lógica de seleção dos 15
    # fragmentos.
    # ============================================================

    categorias_editoriais_ativas = []

    try:
        if isinstance(estrutura_editorial, dict):
            for bloco_editorial, ativo in estrutura_editorial.items():

                if not ativo:
                    continue

                categorias_do_bloco = MAPEAMENTO_EDITORIAL.get(
                    bloco_editorial,
                    []
                )

                for categoria in categorias_do_bloco:
                    if categoria not in categorias_editoriais_ativas:
                        categorias_editoriais_ativas.append(categoria)

    except Exception:
        categorias_editoriais_ativas = []

    print(
        "CATEGORIAS EDITORIAIS ATIVAS:",
        categorias_editoriais_ativas
    )

    # Termos associados às categorias.
    # Servem exclusivamente como sinal adicional de relevância.
    termos_categorias = {

        "definicao": [
            "definição", "definicao", "conceito",
            "descrição", "descricao", "característica",
            "caracteristicas"
        ],

        "beneficios": [
            "benefício", "beneficio", "benefícios",
            "beneficios", "desempenho", "confiabilidade",
            "durabilidade", "eficiência", "eficiencia",
            "economia", "produtividade"
        ],

        "vantagens": [
            "vantagem", "vantagens", "diferencial",
            "diferenciais", "eficiência", "eficiencia",
            "desempenho", "confiabilidade"
        ],

        "materia_prima": [
            "material", "materiais", "matéria-prima",
            "materia-prima", "composição", "composicao",
            "aço", "aco", "ferro", "alumínio", "aluminio"
        ],

        "aplicacoes": [
            "aplicação", "aplicacao", "aplicações",
            "aplicacoes", "utilização", "utilizacao",
            "uso", "empregado", "empregada"
        ],

        "fabricacao": [
            "fabricação", "fabricacao", "produção",
            "producao", "processo", "montagem",
            "construção", "construcao"
        ],

        "manutencao": [
            "manutenção", "manutencao", "inspeção",
            "inspecao", "reparo", "ajuste",
            "lubrificação", "lubrificacao"
        ],

        "ativos_narrativos": [
            "empresa", "fabricante", "fornecedor",
            "produto", "solução", "solucao",
            "serviço", "servico", "suporte",
            "atendimento"
        ],

        "duvidas_frequentes": [
            "dúvida", "duvida", "dúvidas", "duvidas",
            "pergunta", "perguntas", "problema",
            "problemas", "como", "quando", "por que"
        ]
    }

    assuntos_normalizados = []

    for assunto in assuntos:

        assunto_normalizado = (
            normalizar_assunto_texto(
                assunto
            ).strip()
        )

        if assunto_normalizado:

            assuntos_normalizados.append(
                assunto_normalizado
            )

    # ========================================================
    # TERMOS DOS CHECKBOXES
    # ========================================================

    termos_assuntos = {

        "apresentacao": [

            "contexto",
            "finalidade",
            "importância",
            "importancia",
            "necessidade",
            "conceito",
            "definição",
            "definicao",
            "introdução",
            "introducao",
            "característica",
            "caracteristicas"

        ],

        "funcionamento": [

            "funcionamento",
            "funciona",
            "operação",
            "operacao",
            "processo",
            "mecanismo",
            "acionamento",
            "desempenho",
            "movimento",
            "pressão",
            "pressao",
            "vazão",
            "vazao"

        ],

        "aplicacoes": [

            "aplicação",
            "aplicacao",
            "aplicações",
            "aplicacoes",
            "utilização",
            "utilizacao",
            "uso",
            "empregado",
            "empregada",
            "atende",
            "atendimento",
            "sistema",
            "processo"

        ],

        "criterios": [

            "critério",
            "criterio",
            "critérios",
            "criterios",
            "seleção",
            "selecao",
            "dimensionamento",
            "especificação",
            "especificacao",
            "escolha",
            "avaliação",
            "avaliacao"

        ],

        "comercial": [

            "empresa",
            "fabricante",
            "fornecedor",
            "produto",
            "solução",
            "solucao",
            "serviço",
            "servico",
            "atendimento",
            "suporte",
            "equipe",
            "experiência",
            "experiencia"

        ],

        "informacao_tecnica": [

            "informação técnica",
            "informacao tecnica",
            "especificação",
            "especificacao",
            "característica técnica",
            "caracteristica tecnica",
            "material",
            "dimensão",
            "dimensao",
            "capacidade",
            "potência",
            "potencia",
            "pressão",
            "pressao",
            "vazão",
            "vazao",
            "temperatura",
            "rendimento",
            "eficiência",
            "eficiencia"

        ],

        "instalacao_execucao": [

            "instalação",
            "instalacao",
            "montagem",
            "execução",
            "execucao",
            "implantação",
            "implantacao",
            "operação",
            "operacao",
            "manutenção",
            "manutencao",
            "segurança",
            "seguranca",
            "ajuste",
            "inspeção",
            "inspecao"

        ],

        "beneficios": [

            "benefício",
            "beneficio",
            "benefícios",
            "beneficios",
            "vantagem",
            "vantagens",
            "redução",
            "reducao",
            "economia",
            "desempenho",
            "confiabilidade",
            "durabilidade",
            "eficiência",
            "eficiencia",
            "produtividade"

        ]

    }

    # ========================================================
    # 06.1 CONTEXTO EDITORIAL DO MEAD
    # ========================================================

    contexto_blocos_mead = {}

    try:

        if isinstance(
            MEAD,
            dict
        ):

            contexto_blocos_mead = (
                MEAD.get(
                    "contexto_blocos",
                    {}
                )
            )

    except Exception:

        contexto_blocos_mead = {}

    if not isinstance(
        contexto_blocos_mead,
        dict
    ):

        contexto_blocos_mead = {}

    # --------------------------------------------------------
    # OBJETIVOS DOS CINCO BLOCOS
    #
    # Estes termos traduzem os objetivos editoriais do MEAD
    # em critérios de seleção.
    #
    # O texto original do MEAD continua sendo a autoridade.
    # Estes termos apenas permitem ao Python localizar
    # fragmentos compatíveis com aquela finalidade.
    # --------------------------------------------------------

    termos_blocos = {

        # ========================================================
        # BLOCO 1
        # O QUE É / CONTEXTO / FUNCIONAMENTO
        # ========================================================
    
        "bloco_1": [
    
            "contextualização",
            "contextualizacao",
            "contexto",
    
            "conceito",
            "definição",
            "definicao",
    
            "descrição",
            "descricao",
    
            "finalidade",
    
            "funcionamento",
            "funciona",
    
            "princípio",
            "principio",
    
            "mecanismo",
    
            "operação",
            "operacao",
    
            "movimento",
            "movimentação",
            "movimentacao",
    
            "processo",
    
            "pressão",
            "pressao",
    
            "vazão",
            "vazao",
    
            "fluido",
            "liquido",
            "líquido"
    
        ],
    
    
        # ========================================================
        # BLOCO 2
        # CARACTERÍSTICAS / APLICAÇÕES
        # ========================================================
    
        "bloco_2": [
    
            "característica",
            "caracteristicas",
    
            "características",
            "caracteristicas",
    
            "especificação",
            "especificacao",
    
            "capacidade",
    
            "vazão",
            "vazao",
    
            "pressão",
            "pressao",
    
            "temperatura",
    
            "potência",
            "potencia",
    
            "rotação",
            "rotacao",
    
            "material",
            "materiais",
    
            "configuração",
            "configuracao",
    
            "aplicação",
            "aplicacao",
    
            "aplicações",
            "aplicacoes",
    
            "utilização",
            "utilizacao",
    
            "uso",
    
            "empregado",
            "empregada",
    
            "setor",
            "processo",
    
            "sistema",
    
            "desempenho"
    
        ],
    
    
        # ========================================================
        # BLOCO 3
        # VANTAGENS / DIFERENCIAIS
        # ========================================================
    
        "bloco_3": [
    
            "vantagem",
            "vantagens",
    
            "benefício",
            "beneficio",
    
            "benefícios",
            "beneficios",
    
            "diferencial",
            "diferenciais",
    
            "eficiência",
            "eficiencia",
    
            "rendimento",
    
            "desempenho",
    
            "confiabilidade",
    
            "durabilidade",
    
            "resistência",
            "resistencia",
    
            "segurança",
            "seguranca",
    
            "produtividade",
    
            "economia",
    
            "redução",
            "reducao",
    
            "otimização",
            "otimizacao",
    
            "flexibilidade",
    
            "continuidade",
    
            "estabilidade"
    
        ],
    
    
        # ========================================================
        # BLOCO 4
        # ASPECTOS TÉCNICOS / CUIDADOS
        # ========================================================
    
        "bloco_4": [
    
            "aspecto técnico",
            "aspecto tecnico",
    
            "aspectos técnicos",
            "aspectos tecnicos",
    
            "especificação",
            "especificacao",
    
            "dimensionamento",
    
            "instalação",
            "instalacao",
    
            "montagem",
    
            "operação",
            "operacao",
    
            "manutenção",
            "manutencao",
    
            "inspeção",
            "inspecao",
    
            "reparo",
    
            "ajuste",
    
            "lubrificação",
            "lubrificacao",
    
            "vedação",
            "vedacao",
    
            "segurança",
            "seguranca",
    
            "cuidado",
            "cuidados",
    
            "pressão",
            "pressao",
    
            "vazão",
            "vazao",
    
            "temperatura",
    
            "desgaste",
    
            "corrosão",
            "corrosao",
    
            "vibração",
            "vibracao",
    
            "cavitação",
            "cavitacao",
    
            "material",
            "materiais"
    
        ],
    
    
        # ========================================================
        # BLOCO 5
        # ESCOLHA / FORNECIMENTO / SUPORTE
        # ========================================================
    
        "bloco_5": [
    
            "seleção",
            "selecao",
    
            "escolha",
    
            "dimensionamento",
    
            "especificação",
            "especificacao",
    
            "necessidade",
            "necessidades",
    
            "aplicação",
            "aplicacao",
    
            "aplicações",
            "aplicacoes",
    
            "fornecimento",
    
            "fornecedor",
    
            "fornecedores",
    
            "fabricante",
    
            "fabricantes",
    
            "distribuidor",
            "distribuidores",
    
            "suporte",
    
            "assistência",
            "assistencia",
    
            "atendimento",
    
            "atendimento especializado",
    
            "prazo",
    
            "confiabilidade",
    
            "conhecimento técnico",
            "conhecimento tecnico",
    
            "orientação",
            "orientacao",
    
            "especificações",
            "especificacoes"
    
        ]
    
    }

    # --------------------------------------------------------
    # MEAD COMO AUTORIDADE OPERACIONAL
    # --------------------------------------------------------
    termos_mead_prioritarios = {}
    palavras_ruido_mead = {"bloco","objetivo","funcao","função","contexto","deve","devem","pagina","página","conteudo","conteúdo","tecnico","técnico","tecnica","técnica","somente","apenas","quando","caso","forma","sem","com","para","sobre","entre","pelos","pelas","este","esta","esse","essa","como","uma","uns","umas","dos","das","que","não","nao","uso","usar"}
    for numero_bloco in range(1,6):
        chave_bloco=f"bloco_{numero_bloco}"; dados_mead=contexto_blocos_mead.get(chave_bloco,{})
        if not isinstance(dados_mead,dict): continue
        base_mead=" ".join(str(dados_mead.get(k,"") or "") for k in ("objetivo","funcao","função","contexto"))
        tokens=[]
        for termo in re.findall(r"\b[a-záàãâéêíóôõúç]{4,}\b",base_mead.casefold()):
            termo_n=normalizar_assunto_texto(termo)
            if termo_n and termo_n not in palavras_ruido_mead and termo_n not in tokens: tokens.append(termo_n)
        termos_mead_prioritarios[chave_bloco]=tokens
    termos_blocos_mead_fallback={
        "bloco_1":["definição","funcionamento","contexto","necessidade","importância","cenário"],
        "bloco_2":["funcionamento","aplicações","características","uso","desempenho"],
        "bloco_3":["critérios","cuidados","instalação","operação","manutenção","segurança","especificação","dimensionamento"],
        "bloco_4":["empresa","conhecimento","serviço","suporte","produto","solução","experiência","atendimento"],
        "bloco_5":["necessidade","aplicação","conhecimento","solução","seleção","fornecimento","suporte"],
    }
    for chave_bloco, termos in termos_blocos_mead_fallback.items():
        termos_blocos[chave_bloco]=(termos_mead_prioritarios.get(chave_bloco,[])*3)+termos+termos_blocos.get(chave_bloco,[])

    # --------------------------------------------------------
    # TAMBÉM LÊ O TEXTO REAL DO MEAD
    #
    # Isso mantém a seleção vinculada ao MEAD carregado,
    # sem enviar o MEAD inteiro para o Ollama.
    # --------------------------------------------------------

    # ========================================================
    # SELEÇÃO COM DIVERSIDADE DE FONTES
    # ========================================================
    # Em cada bloco, a primeira preferência é por uma fonte
    # ainda não utilizada. Só reutilizamos uma fonte quando não
    # houver alternativa válida. Isso evita que uma única página
    # domine os 15 fragmentos quando o patrimônio contém dezenas
    # de fontes relevantes.

    for numero_bloco in range(
        1,
        6
    ):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        dados_bloco_mead = (
            contexto_blocos_mead.get(
                chave_bloco,
                {}
            )
        )

        if not isinstance(
            dados_bloco_mead,
            dict
        ):

            continue

        objetivo_mead = normalizar_assunto_texto(
            dados_bloco_mead.get(
                "objetivo",
                ""
            )
        )

        funcao_mead = normalizar_assunto_texto(
            dados_bloco_mead.get(
                "funcao",
                ""
            )
        )

        texto_mead = (
            f"{objetivo_mead} "
            f"{funcao_mead}"
        ).strip()

        if texto_mead:

            termos_existentes = (
                termos_blocos.get(
                    chave_bloco,
                    []
                )
            )

            termos_blocos[
                chave_bloco
            ] = (

                termos_existentes
                +
                [
                    termo
                    for termo in texto_mead.split()
                    if len(termo) > 4
                ]

            )

    # --------------------------------------------------------
    # NORMALIZAR TERMOS DOS BLOCOS
    # --------------------------------------------------------

    termos_blocos_normalizados = {}

    for chave_bloco, termos in termos_blocos.items():

        termos_normalizados = []

        for termo in termos:

            termo_normalizado = (
                normalizar_assunto_texto(
                    termo
                ).strip()
            )

            if termo_normalizado:

                termos_normalizados.append(
                    termo_normalizado
                )

        termos_blocos_normalizados[
            chave_bloco
        ] = termos_normalizados

    # ========================================================
    # 06.2 PONTUAÇÃO DOS CHECKBOXES
    # ========================================================

    def calcular_pontuacao_checkbox(
        texto_normalizado
    ):

        pontuacao = 0

        for assunto in assuntos_normalizados:

            termos = termos_assuntos.get(
                assunto,
                []
            )

            for termo in termos:

                termo_normalizado = (
                    normalizar_assunto_texto(
                        termo
                    )
                )

                if (
                    termo_normalizado
                    and
                    termo_normalizado
                    in texto_normalizado
                ):

                    pontuacao += 1

        return pontuacao
        
    # ========================================================
    # 06.2.1 PONTUAÇÃO DOS CHECKBOXES
    # ========================================================   

    def calcular_pontuacao_categorias(texto_normalizado):
        """
        Usa as categorias editoriais ativas apenas como
        sinal adicional de relevância para os fragmentos.

        Não realiza seleção própria.
        Não altera quantidade.
        Não altera distribuição.
        """

        pontuacao = 0

        for categoria in categorias_editoriais_ativas:

            termos = termos_categorias.get(
                categoria,
                []
            )

            termos_vistos = set()

            for termo in termos:

                termo_normalizado = (
                    normalizar_assunto_texto(
                        termo
                    )
                )

                if (
                    not termo_normalizado
                    or
                    termo_normalizado
                    in termos_vistos
                ):
                    continue

                termos_vistos.add(
                    termo_normalizado
                )

                if (
                    termo_normalizado
                    in texto_normalizado
                ):
                    pontuacao += 1

        return pontuacao    


    

    # ========================================================
    # 06.3 PONTUAÇÃO ESPECÍFICA DE CADA BLOCO
    # ========================================================

    def calcular_pontuacao_bloco(
        candidato,
        chave_bloco
    ):

        texto_normalizado = (
            normalizar_assunto_texto(
                candidato.get(
                    "texto",
                    ""
                )
            )
        )

        if not texto_normalizado:

            return 0

        pontuacao = 0

        # ----------------------------------------------------
        # PRIORIDADE COMERCIAL DO BLOCO 1 (FLEXÍVEL)
        # Pontua evidências reais de fornecimento/contratação;
        # não transforma ausência de evidência em reprovação.
        # ----------------------------------------------------
        if chave_bloco == "bloco_1":
            termos_comerciais_b1 = (
                "fornecimento", "fornecedor", "fornecedores",
                "fabricacao", "fabricante", "fabricantes",
                "distribuicao", "distribuidor", "distribuidores",
                "comercializacao", "comercializa", "comercializado",
                "contratacao", "aquisicao", "compra", "venda",
                "orcamento", "cotacao", "escopo", "atendimento",
                "servico", "servicos", "produto", "produtos"
            )
            sinais_comerciais = sum(
                1 for termo in termos_comerciais_b1
                if termo in texto_normalizado
            )
            pontuacao += min(32, sinais_comerciais * 8)

        # ----------------------------------------------------
        # A. COMPATIBILIDADE COM O OBJETIVO DO BLOCO
        # ----------------------------------------------------

        termos_bloco = termos_blocos_normalizados.get(
            chave_bloco,
            []
        )

        termos_vistos = set()

        for termo in termos_bloco:

            if termo in termos_vistos:

                continue

            termos_vistos.add(
                termo
            )

            if (
                termo
                and
                termo
                in texto_normalizado
            ):

                pontuacao += 2

        # ----------------------------------------------------
        # B. COMPATIBILIDADE COM OS CHECKBOXES
        # ----------------------------------------------------

        pontuacao_checkbox = (
            calcular_pontuacao_checkbox(
                texto_normalizado
            )
        )

        pontuacao += (
            pontuacao_checkbox
        )
        
        # ----------------------------------------------------
        # B.1 COMPATIBILIDADE COM AS CATEGORIAS
        #
        # As categorias servem somente como reforço na
        # pontuação do candidato.
        #
        # NÃO fazem uma nova seleção.
        # NÃO alteram a quantidade de fragmentos.
        # NÃO alteram a distribuição dos blocos.
        # ----------------------------------------------------

        pontuacao_categorias = (
            calcular_pontuacao_categorias(
                texto_normalizado
            )
        )

        pontuacao += (
            pontuacao_categorias
        )

        # ----------------------------------------------------
        # C. PRIORIDADE TÉCNICA
        #
        # PDFs continuam tendo vantagem, mas somente como
        # critério complementar.
        # ----------------------------------------------------

        if candidato.get(
            "pdf"
        ):

            pontuacao += 1

        # ----------------------------------------------------
        # D. RELEVÂNCIA DIRETA PARA O TEMA
        # ----------------------------------------------------

        tema_normalizado = (
            normalizar_assunto_texto(
                tema
            )
        )

        palavras_tema = [
            palavra
            for palavra
            in tema_normalizado.split()
            if len(palavra) > 2
        ]

        for palavra in palavras_tema:

            if palavra in texto_normalizado:

                pontuacao += 2
                
                
        # ----------------------------------------------------
        # E. QUALIDADE EDITORIAL DO FRAGMENTO
        #
        # Mede se o trecho realmente possui conteúdo útil
        # para desenvolvimento de um parágrafo técnico.
        #
        # Não substitui:
        # - bloco
        # - checkbox
        # - categoria
        # - PDF
        # - tema
        #
        # Apenas melhora a classificação dos candidatos.
        # ----------------------------------------------------

        qualidade = 0

        texto_original = str(
            candidato.get(
                "texto",
                ""
            )
        ).strip()

        quantidade_palavras = len(
            texto_normalizado.split()
        )

        # ----------------------------------------------------
        # E.1 TAMANHO ADEQUADO
        #
        # Os fragmentos oficiais trabalham entre 50 e 60
        # palavras. Um trecho nessa faixa recebe pequeno
        # reforço por estar adequado ao uso editorial.
        # ----------------------------------------------------

        if 50 <= quantidade_palavras <= 60:

            qualidade += 2

        elif 40 <= quantidade_palavras <= 70:

            qualidade += 1

        # ----------------------------------------------------
        # E.2 FRASES COMPLETAS
        #
        # Fragmentos com pontuação e várias frases tendem
        # a possuir conteúdo explicativo real.
        # ----------------------------------------------------

        quantidade_frases = len(
            re.findall(
                r"[.!?]",
                texto_original
            )
        )

        if quantidade_frases >= 2:

            qualidade += 2

        elif quantidade_frases == 1:

            qualidade += 1

        # ----------------------------------------------------
        # E.3 DENSIDADE TÉCNICA
        #
        # Termos que normalmente indicam explicação técnica.
        # ----------------------------------------------------

        termos_tecnicos = [

            "funcionamento",
            "operacao",
            "operação",
            "aplicacao",
            "aplicação",
            "processo",
            "sistema",
            "equipamento",
            "componente",
            "pressao",
            "pressão",
            "vazao",
            "vazão",
            "temperatura",
            "desempenho",
            "eficiencia",
            "eficiência",
            "rendimento",
            "manutencao",
            "manutenção",
            "instalacao",
            "instalação",
            "dimensionamento",
            "especificacao",
            "especificação",
            "material",
            "materiais",
            "confiabilidade",
            "durabilidade",
            "seguranca",
            "segurança"
        ]

        termos_tecnicos_encontrados = 0

        for termo in termos_tecnicos:

            if termo in texto_normalizado:

                termos_tecnicos_encontrados += 1

        if termos_tecnicos_encontrados >= 4:

            qualidade += 4

        elif termos_tecnicos_encontrados >= 2:

            qualidade += 2

        elif termos_tecnicos_encontrados == 1:

            qualidade += 1

        # ----------------------------------------------------
        # E.4 ESTRUTURA EXPLICATIVA
        #
        # Reforça trechos que apresentam relação entre
        # causa, funcionamento, aplicação, resultado etc.
        # ----------------------------------------------------

        termos_explicativos = [

            "porque",
            "quando",
            "como",
            "permite",
            "possibilita",
            "utilizado",
            "utilizada",
            "responsavel",
            "responsável",
            "garante",
            "evita",
            "reduz",
            "aumenta",
            "proporciona",
            "contribui",
            "depende",
            "necessario",
            "necessário",
            "indicado",
            "indicada"
        ]

        explicativos_encontrados = 0

        for termo in termos_explicativos:

            if termo in texto_normalizado:

                explicativos_encontrados += 1

        if explicativos_encontrados >= 3:

            qualidade += 3

        elif explicativos_encontrados >= 1:

            qualidade += 1

        # ----------------------------------------------------
        # E.5 PENALIZAR ÍNDICE / SUMÁRIO
        # ----------------------------------------------------

        marcadores_indice = [

            "sumario",
            "sumário",
            "indice",
            "índice",
            "conteudo",
            "conteúdo",
            "capitulo",
            "capítulo",
            "secao",
            "seção",
            "2.1",
            "2.2",
            "2.3",
            "3.1",
            "3.2",
            "3.3"
        ]

        ocorrencias_indice = 0

        for marcador in marcadores_indice:

            if marcador in texto_normalizado:

                ocorrencias_indice += 1

        if ocorrencias_indice >= 2:

            qualidade -= 8

        elif ocorrencias_indice == 1:

            qualidade -= 4

        # ----------------------------------------------------
        # E.6 PENALIZAR MENU / NAVEGAÇÃO
        # ----------------------------------------------------

        marcadores_navegacao = [

            "home",
            "inicio",
            "menu",
            "contato",
            "login",
            "cadastro",
            "entrar",
            "politica de privacidade",
            "política de privacidade",
            "cookies",
            "siga-nos",
            "compartilhe"
        ]

        navegacao_encontrada = 0

        for marcador in marcadores_navegacao:

            if marcador in texto_normalizado:

                navegacao_encontrada += 1

        if navegacao_encontrada >= 2:

            qualidade -= 8

        elif navegacao_encontrada == 1:

            qualidade -= 3

        # ----------------------------------------------------
        # E.7 PENALIZAR COMENTÁRIOS / REDES SOCIAIS
        # ----------------------------------------------------

        marcadores_comentario = [

            "curtir",
            "comentar",
            "responder",
            "comentarios",
            "comentários",
            "seguidores",
            "linkedin",
            "facebook",
            "instagram",
            "twitter",
            "postado por"
        ]

        comentario_encontrado = 0

        for marcador in marcadores_comentario:

            if marcador in texto_normalizado:

                comentario_encontrado += 1

        if comentario_encontrado >= 2:

            qualidade -= 8

        elif comentario_encontrado == 1:

            qualidade -= 4

        # ----------------------------------------------------
        # E.8 PENALIZAR TEXTO CORROMPIDO
        # ----------------------------------------------------

        caracteres_corrompidos = (

            texto_original.count("\x00")
            +
            texto_original.count("\x03")
            +
            texto_original.count("\ufffd")

        )

        if caracteres_corrompidos > 0:

            qualidade -= 10

        # ----------------------------------------------------
        # E.9 PENALIZAR EXCESSO DE FRAGMENTAÇÃO
        #
        # Muitos separadores, números isolados e símbolos
        # indicam frequentemente índice, tabela ou conteúdo
        # extraído de PDF de forma inadequada.
        # ----------------------------------------------------

        quantidade_numeros = len(
            re.findall(
                r"\b\d+(?:[.,]\d+)?\b",
                texto_original
            )
        )

        if quantidade_numeros >= 8:

            qualidade -= 5

        elif quantidade_numeros >= 5:

            qualidade -= 2

        # ----------------------------------------------------
        # E.10 APLICAR QUALIDADE À NOTA FINAL
        # ----------------------------------------------------

        pontuacao += qualidade        

        return pontuacao



        # ========================================================
        # ADERÊNCIA TEMÁTICA PREDOMINANTE
        # ========================================================
        #
        # Regra geral do MEAD:
        #
        # Um fragmento não é considerado adequado apenas porque
        # contém os termos da palavra-chave.
        #
        # O conteúdo precisa demonstrar que o tema é um dos
        # assuntos predominantes do fragmento.
        #
        # Esta função é genérica.
        #
        # NÃO conhece produtos.
        # NÃO conhece serviços.
        # NÃO possui lista de concorrentes.
        # NÃO possui exceções por palavra-chave.
        #
        # Ela analisa somente a relação entre:
        #
        #     TEMA <-> CONTEÚDO DO FRAGMENTO
        #
        # ========================================================
    
        def fragmento_tem_aderencia_tematica(
            texto,
            tema
        ):
    
            texto = str(
                texto or ""
            ).strip()
    
            tema = str(
                tema or ""
            ).strip()
    
            if not texto or not tema:
                return False
    
    
            # ----------------------------------------------------
            # PRIMEIRO FILTRO:
            # o fragmento precisa pertencer ao tema.
            # ----------------------------------------------------
    
            if not fragmento_pertence_ao_tema(
                texto,
                tema
            ):
                return False
    
    
            # ----------------------------------------------------
            # NORMALIZAÇÃO
            # ----------------------------------------------------
    
            texto_normalizado = (
                normalizar_assunto_texto(
                    texto
                )
            )
    
            tema_normalizado = (
                normalizar_assunto_texto(
                    tema
                )
            )
    
    
            # ----------------------------------------------------
            # TERMOS SIGNIFICATIVOS DO TEMA
            # ----------------------------------------------------
    
            palavras_ignoradas = {
    
                "de",
                "da",
                "das",
                "do",
                "dos",
                "em",
                "na",
                "nas",
                "no",
                "nos",
                "para",
                "por",
                "com",
                "sem",
                "e",
                "a",
                "o",
                "as",
                "os"
    
            }
    
    
            palavras_tema = [
    
                palavra
    
                for palavra in re.findall(
                    r"\b[a-z0-9]{3,}\b",
                    tema_normalizado
                )
    
                if palavra
                not in palavras_ignoradas
    
            ]
    
    
            if not palavras_tema:
                return False
    
    
            # ----------------------------------------------------
            # DIVIDIR O FRAGMENTO EM FRASES
            # ----------------------------------------------------
    
            frases = re.split(
                r"(?<=[.!?])\s+",
                texto_normalizado
            )
    
            frases = [
    
                frase.strip()
    
                for frase in frases
    
                if frase.strip()
    
            ]
    
    
            if not frases:
                return False
    
    
            # ----------------------------------------------------
            # RADICAIS DO TEMA
            # ----------------------------------------------------
    
            radicais_tema = []
    
            for palavra in palavras_tema:
    
                radical = palavra
    
                if len(radical) >= 5:
    
                    if radical.endswith("es"):
                        radical = radical[:-2]
    
                    elif radical.endswith("s"):
                        radical = radical[:-1]
    
                    elif radical.endswith("a"):
                        radical = radical[:-1]
    
                    elif radical.endswith("o"):
                        radical = radical[:-1]
    
                radicais_tema.append(
                    radical
                )
    
    
            # ----------------------------------------------------
            # ANALISAR CADA FRASE
            # ----------------------------------------------------
    
            frases_relevantes = 0
    
            maior_aderencia_frase = 0
    
            encontrou_proximidade = False
    
    
            for frase in frases:
    
                palavras_frase = re.findall(
                    r"\b[a-z0-9]{3,}\b",
                    frase
                )
    
                if not palavras_frase:
                    continue
    
    
                encontrados = 0
    
                posicoes = []
    
    
                for radical in radicais_tema:
    
                    encontrou = False
    
                    for indice, palavra in enumerate(
                        palavras_frase
                    ):
    
                        if palavra.startswith(
                            radical
                        ):
    
                            encontrou = True
    
                            posicoes.append(
                                indice
                            )
    
                            break
    
    
                    if encontrou:
                        encontrados += 1
    
    
                # ------------------------------------------------
                # ADERÊNCIA DA FRASE
                # ------------------------------------------------
    
                if encontrados > 0:
    
                    aderencia_frase = (
                        encontrados
                        /
                        len(radicais_tema)
                    )
    
                    if aderencia_frase > maior_aderencia_frase:
    
                        maior_aderencia_frase = (
                            aderencia_frase
                        )
    
    
                # ------------------------------------------------
                # FRASE COM TODOS OS TERMOS DO TEMA
                # ------------------------------------------------
    
                if encontrados == len(
                    radicais_tema
                ):
    
                    frases_relevantes += 1
    
                    # --------------------------------------------
                    # PROXIMIDADE DOS TERMOS
                    #
                    # Quanto mais próximos os termos do tema,
                    # maior a evidência de que formam um assunto
                    # único dentro da frase.
                    # --------------------------------------------
    
                    if len(posicoes) >= 2:
    
                        distancia = (
                            max(posicoes)
                            -
                            min(posicoes)
                        )
    
                        if distancia <= 8:
    
                            encontrou_proximidade = True
    
                elif encontrados > 0:
    
                    # --------------------------------------------
                    # Frase parcialmente relacionada.
                    #
                    # Só recebe peso se houver uma quantidade
                    # relevante dos termos do tema.
                    # --------------------------------------------
    
                    if (
                        encontrados
                        /
                        len(radicais_tema)
                    ) >= 0.5:
    
                        frases_relevantes += 1
    
    
            # ----------------------------------------------------
            # PROPORÇÃO DE FRASES RELACIONADAS
            # ----------------------------------------------------
    
            proporcao_relevante = (
                frases_relevantes
                /
                len(frases)
            )
    
    
            # ----------------------------------------------------
            # REGRA PARA TEMAS COMPOSTOS
            #
            # Exemplo:
            #
            # "bomba centrífuga"
            #
            # Os dois termos precisam aparecer juntos em pelo
            # menos uma frase quando isso for possível.
            # ----------------------------------------------------
    
            if len(radicais_tema) >= 2:
    
                if not encontrou_proximidade:
    
                    return False
    
    
            # ----------------------------------------------------
            # REGRA DE PREDOMINÂNCIA
            #
            # Fragmentos muito curtos podem ter somente uma frase.
            # Nesse caso, a presença contextual é suficiente.
            #
            # Em fragmentos com várias frases, o tema precisa
            # aparecer em uma parcela relevante do conteúdo.
            # ----------------------------------------------------
    
            if len(frases) == 1:
    
                return (
                    maior_aderencia_frase >= 0.5
                )
    
    
            if len(frases) == 2:
    
                return (
                    proporcao_relevante >= 0.5
                )
    
    
            if len(frases) >= 3:
    
                return (
                    proporcao_relevante >= 0.5
                    or
                    (
                        maior_aderencia_frase >= 1.0
                        and
                        proporcao_relevante >= 0.34
                    )
                )
    
    
            return False        
    # ========================================================
    # 06.4 ORGANIZAR CANDIDATOS POR BLOCO
    # ========================================================

    candidatos_por_bloco = {

        "bloco_1": [],
        "bloco_2": [],
        "bloco_3": [],
        "bloco_4": [],
        "bloco_5": []

    }

    for candidato in candidatos:

        if not fragmento_pertence_ao_tema(
            candidato.get(
                "texto",
                ""
            ),
            tema
        ):
            continue
    
        for numero_bloco in range(
            1,
            6
        ):
    
            chave_bloco = (
                f"bloco_{numero_bloco}"
            )
    
            pontuacao = (
                calcular_pontuacao_bloco(
                    candidato,
                    chave_bloco
                )
            )
    
            candidatos_por_bloco[
                chave_bloco
            ].append({
    
                "candidato":
                    candidato,
    
                "pontuacao":
                    pontuacao
    
            })

    # ========================================================
    # 06.5 ORDENAR CADA BLOCO
    # ========================================================

    for chave_bloco in candidatos_por_bloco:

        candidatos_por_bloco[
            chave_bloco
        ] = sorted(

            candidatos_por_bloco[
                chave_bloco
            ],

            key=lambda item: (
                item["pontuacao"],
                1
                if item["candidato"].get("pdf")
                else 0
            ),

            reverse=True

        )

    print()
    print(
        "=============================="
    )
    print(
        "PRIORIZAÇÃO EDITORIAL POR BLOCO"
    )
    print(
        "=============================="
    )

    print(
        "ASSUNTOS ATIVOS:",
        assuntos
    )

    for numero_bloco in range(
        1,
        6
    ):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        print()
        print(
            chave_bloco.upper()
        )

        dados_mead = (
            contexto_blocos_mead.get(
                chave_bloco,
                {}
            )
        )

        if isinstance(
            dados_mead,
            dict
        ):

            print(
                "OBJETIVO MEAD:",
                dados_mead.get(
                    "objetivo",
                    ""
                )
            )

            print(
                "FUNÇÃO MEAD:",
                dados_mead.get(
                    "funcao",
                    ""
                )
            )

        print(
            "CANDIDATOS:",
            len(
                candidatos_por_bloco[
                    chave_bloco
                ]
            )
        )

    # ========================================================
    # 07. SELECIONAR 15 FRAGMENTOS
    #     3 PARA CADA BLOCO
    #
    # SELEÇÃO PROGRESSIVA COM DIVERSIDADE DE FONTES
    #
    # Objetivo:
    #
    # 1. respeitar a pontuação editorial;
    # 2. respeitar o objetivo de cada bloco;
    # 3. evitar repetir a mesma fonte;
    # 4. evitar repetir o mesmo assunto;
    # 5. evitar trechos muito semelhantes;
    # 6. manter preferência por conteúdo técnico;
    # 7. garantir 3 fragmentos por bloco.
    #
    # ========================================================

    fragmentos_selecionados = []

    hashes_selecionados = set()
    
    fontes_utilizadas = {}

    # --------------------------------------------------------
    # FUNÇÃO AUXILIAR:
    # NORMALIZAR PALAVRAS DO FRAGMENTO
    # --------------------------------------------------------

    def obter_palavras_conteudo(
        texto
    ):

        texto_normalizado = (
            normalizar_assunto_texto(
                texto
            )
        )

        palavras = re.findall(
            r"\b[a-z0-9]{4,}\b",
            texto_normalizado
        )

        # ----------------------------------------------------
        # REMOVER PALAVRAS MUITO GENÉRICAS
        # ----------------------------------------------------

        palavras_ignoradas = {

            "para",
            "como",
            "mais",
            "menos",
            "essa",
            "esse",
            "estas",
            "estes",
            "sobre",
            "entre",
            "tambem",
            "quando",
            "onde",
            "sendo",
            "pode",
            "podem",
            "cada",
            "pela",
            "pelo",
            "pelas",
            "pelos",
            "uma",
            "umas",
            "uns",
            "dos",
            "das",
            "com",
            "sem",
            "que",
            "por",
            "uma",
            "seus",
            "suas"

        }

        palavras = [

            palavra

            for palavra in palavras

            if palavra
            not in palavras_ignoradas

        ]

        return set(
            palavras
        )

    # --------------------------------------------------------
    # FUNÇÃO AUXILIAR:
    # MEDIR SEMELHANÇA ENTRE DOIS FRAGMENTOS
    #
    # Utilizamos interseção de palavras relevantes.
    #
    # Isso evita selecionar:
    #
    # PDF A → trecho 1
    # PDF A → trecho 2
    #
    # quando os dois trechos praticamente repetem
    # a mesma informação.
    # --------------------------------------------------------

    def calcular_semelhanca_fragmentos(
        texto_a,
        texto_b
    ):

        palavras_a = obter_palavras_conteudo(
            texto_a
        )

        palavras_b = obter_palavras_conteudo(
            texto_b
        )

        if not palavras_a or not palavras_b:

            return 0.0

        intersecao = (
            palavras_a
            &
            palavras_b
        )

        menor_conjunto = min(
            len(palavras_a),
            len(palavras_b)
        )

        if menor_conjunto <= 0:

            return 0.0

        return (
            len(intersecao)
            /
            menor_conjunto
        )

    # --------------------------------------------------------
    # FUNÇÃO AUXILIAR:
    # PENALIDADE POR REPETIÇÃO DA FONTE
    # --------------------------------------------------------

    def calcular_penalidade_fonte(
        candidato
    ):

        fonte = candidato.get(
            "fonte"
        )

        quantidade = fontes_utilizadas.get(
            fonte,
            0
        )

        if quantidade == 0:

            return 0

        if quantidade == 1:

            return 8

        if quantidade == 2:

            return 18

        return 30

    # --------------------------------------------------------
    # FUNÇÃO AUXILIAR:
    # PENALIDADE POR SEMELHANÇA COM O QUE
    # JÁ FOI SELECIONADO
    # --------------------------------------------------------

    def calcular_penalidade_semelhanca(
        candidato
    ):

        texto_candidato = candidato.get(
            "texto",
            ""
        )

        maior_semelhanca = 0.0

        for selecionado in fragmentos_selecionados:

            semelhanca = (
                calcular_semelhanca_fragmentos(
                    texto_candidato,
                    selecionado.get(
                        "texto",
                        ""
                    )
                )
            )

            if semelhanca > maior_semelhanca:

                maior_semelhanca = (
                    semelhanca
                )

        # ----------------------------------------------------
        # PENALIZAÇÃO PROGRESSIVA
        # ----------------------------------------------------

        if maior_semelhanca >= 0.70:

            return 35

        if maior_semelhanca >= 0.55:

            return 25

        if maior_semelhanca >= 0.45:

            return 15

        if maior_semelhanca >= 0.35:

            return 8

        return 0

    # --------------------------------------------------------
    # FUNÇÃO AUXILIAR:
    # PENALIDADE POR REPETIÇÃO DE ASSUNTO
    #
    # Não é uma classificação semântica pesada.
    #
    # É uma proteção simples para evitar que os 3 trechos
    # de um bloco fiquem falando essencialmente da mesma
    # coisa.
    # --------------------------------------------------------

    def calcular_penalidade_assunto(
        candidato,
        chave_bloco
    ):

        texto_candidato = (
            normalizar_assunto_texto(
                candidato.get(
                    "texto",
                    ""
                )
            )
        )

        penalidade = 0

        # ----------------------------------------------------
        # TERMOS PRINCIPAIS DO BLOCO
        # ----------------------------------------------------

        termos_bloco = set(
            termos_blocos_normalizados.get(
                chave_bloco,
                []
            )
        )

        termos_candidato = set()

        for termo in termos_bloco:

            if (
                termo
                and
                termo in texto_candidato
            ):

                termos_candidato.add(
                    termo
                )

        # ----------------------------------------------------
        # COMPARAR COM OS FRAGMENTOS DO MESMO BLOCO
        # ----------------------------------------------------

        for selecionado in fragmentos_selecionados:

            if selecionado.get(
                "bloco_mead"
            ) != chave_bloco:

                continue

            texto_selecionado = (
                normalizar_assunto_texto(
                    selecionado.get(
                        "texto",
                        ""
                    )
                )
            )

            termos_repetidos = 0

            for termo in termos_candidato:

                if termo in texto_selecionado:

                    termos_repetidos += 1

            if termos_repetidos >= 4:

                penalidade += 8

            elif termos_repetidos >= 2:

                penalidade += 4

        return penalidade

    # --------------------------------------------------------
    # FUNÇÃO AUXILIAR:
    # NOTA FINAL DE DIVERSIDADE
    # --------------------------------------------------------

    def calcular_nota_selecao(
        candidato,
        pontuacao_original,
        chave_bloco
    ):

        penalidade_fonte = (
            calcular_penalidade_fonte(
                candidato
            )
        )

        penalidade_semelhanca = (
            calcular_penalidade_semelhanca(
                candidato
            )
        )

        penalidade_assunto = (
            calcular_penalidade_assunto(
                candidato,
                chave_bloco
            )
        )

        nota_final = (
            pontuacao_original
            -
            penalidade_fonte
            -
            penalidade_semelhanca
            -
            penalidade_assunto
        )

        # ----------------------------------------------------
        # PEQUENO BÔNUS PARA FONTE NOVA
        #
        # Apenas reforça a diversidade.
        # ----------------------------------------------------

        if fontes_utilizadas.get(
            candidato.get(
                "fonte"
            ),
            0
        ) == 0:

            nota_final += 5

        # ----------------------------------------------------
        # PDF CONTINUA SENDO PRIORIDADE TÉCNICA
        # ----------------------------------------------------

        if candidato.get(
            "pdf"
        ):

            nota_final += 1

        return nota_final
        
        
    # ========================================================
    # VALIDAÇÃO ESTRUTURAL DO CANDIDATO
    # ========================================================
    #
    # A pontuação decide QUEM DEVE SER TENTADO PRIMEIRO.
    #
    # Esta função decide se o trecho pode realmente ser usado.
    #
    # NÃO cria nova pontuação.
    # NÃO substitui a pontuação editorial.
    #
    # Apenas elimina lixo estrutural evidente.
    # ========================================================

    def candidato_eh_utilizavel(
        candidato
    ):

        if not isinstance(
            candidato,
            dict
        ):
            return False

        texto = str(
            candidato.get(
                "texto",
                ""
            )
            or ""
        ).strip()

        if not texto:
            return False

        # A presença literal na fonte é apenas um indicador de prioridade.
        # Fragmentos SEM também podem sustentar conteúdo relacionado ao tema;
        # a expressão exata será exigida no texto final, não na fonte.
        ocorrencias_chave = contar_palavra_chave_literal(texto, tema)
        candidato["palavra_chave_status"] = "COM" if ocorrencias_chave >= 1 else "SEM"
        candidato["palavra_chave_ocorrencias_fonte"] = ocorrencias_chave

        # ----------------------------------------------------
        # TEMA OBRIGATÓRIO
        # ----------------------------------------------------
        #
        # A pontuação editorial jamais pode selecionar um
        # fragmento que não pertença ao tema.
        # ----------------------------------------------------

        if not fragmento_pertence_ao_tema(
            texto,
            tema
        ):

            return False

        # ----------------------------------------------------
        # IDENTIDADE COMERCIAL / PRODUTO
        # ----------------------------------------------------

        if not fragmento_eh_comercialmente_limpo(
            texto
        ):
            return False

        identidade = candidato.get("identidade_fonte", {})
        if isinstance(identidade, dict):
            nome_fonte = str(identidade.get("nome", "") or "").strip()
            confianca_fonte = str(identidade.get("confianca", "") or "").casefold()
            if nome_fonte and confianca_fonte in {"alta", "high"}:
                if re.search(r"(?<!\w)" + re.escape(nome_fonte) + r"(?!\w)", texto, re.I):
                    return False

        # O fallback de extração não pode reintroduzir títulos, listas
        # quebradas, CTAs ou resíduos que o filtro editorial rejeitou.
        if not fragmento_eh_aproveitavel_editorialmente(texto):
            return False

        ok_evidencia, _motivo_evidencia = auditar_evidencia_para_redacao(texto)
        if not ok_evidencia:
            return False

        # ----------------------------------------------------
        # QUANTIDADE DE PALAVRAS
        # ----------------------------------------------------

        palavras = texto.split()

        if len(palavras) < 35:
            return False

        # ----------------------------------------------------
        # CARACTERES CORROMPIDOS
        # ----------------------------------------------------

        if (
            "\x00" in texto
            or
            "\x03" in texto
            or
            "\ufffd" in texto
        ):
            return False

        # ----------------------------------------------------
        # NORMALIZAÇÃO LOCAL
        #
        # IMPORTANTE:
        # não depender de normalizar_assunto_texto()
        # para evitar problema de escopo.
        # ----------------------------------------------------

        texto_normalizado = (
            " ".join(
                texto.casefold().split()
            )
        )

        if not texto_normalizado:
            return False

        # ----------------------------------------------------
        # REDES SOCIAIS / COMENTÁRIOS
        # ----------------------------------------------------

        marcadores_social = [

            "comments",
            "comment",
            "others also viewed",
            "curtir",
            "comentar",
            "comentarios",
            "comentários",
            "responder",
            "seguidores",
            "linkedin",
            "facebook",
            "instagram",
            "twitter",
            "postado por"

        ]

        ocorrencias_social = sum(

            1

            for marcador
            in marcadores_social

            if marcador
            in texto_normalizado

        )

        if ocorrencias_social >= 2:
            return False

        # ----------------------------------------------------
        # NAVEGAÇÃO / MENU
        # ----------------------------------------------------

        marcadores_navegacao = [

            "home",
            "inicio",
            "menu",
            "contato",
            "login",
            "cadastro",
            "entrar",
            "cookies",
            "politica de privacidade",
            "política de privacidade",
            "siga-nos",
            "compartilhe"

        ]

        ocorrencias_navegacao = sum(

            1

            for marcador
            in marcadores_navegacao

            if marcador
            in texto_normalizado

        )

        if ocorrencias_navegacao >= 2:
            return False

        # ----------------------------------------------------
        # ÍNDICE / SUMÁRIO
        # ----------------------------------------------------

        marcadores_indice = [

            "sumario",
            "sumário",
            "indice",
            "índice",
            "capitulo",
            "capítulo",
            "referencias",
            "referências",
            "secao",
            "seção"

        ]

        ocorrencias_indice = sum(

            1

            for marcador
            in marcadores_indice

            if marcador
            in texto_normalizado

        )

        if ocorrencias_indice >= 2:
            return False

        # ----------------------------------------------------
        # ESTRUTURA DE ÍNDICE
        # ----------------------------------------------------

        ocorrencias_numeracao = len(

            re.findall(
                r"\b\d+\.\d+(?:\.\d+)*\.?\b",
                texto
            )

        )

        if ocorrencias_numeracao >= 3:
            return False

        # ----------------------------------------------------
        # LISTAS / TABELAS
        # ----------------------------------------------------

        quantidade_numeros = len(

            re.findall(
                r"\b\d+(?:[.,]\d+)?\b",
                texto
            )

        )

        marcadores_tabela = [

            "tamanho",
            "tamanhos",
            "vazao",
            "vazão",
            "pressao",
            "pressão",
            "temperatura",
            "rotacao",
            "rotação",
            "rpm",
            "peso",
            "dimensao",
            "dimensão",
            "modelo",
            "codigo",
            "código"

        ]

        ocorrencias_tabela = sum(

            1

            for marcador
            in marcadores_tabela

            if marcador
            in texto_normalizado

        )

        if (
            quantidade_numeros >= 8
            and
            ocorrencias_tabela >= 3
        ):
            return False

        # ----------------------------------------------------
        # EXCESSO DE SEPARADORES
        # ----------------------------------------------------

        separadores = len(

            re.findall(
                r"[|;]{4,}",
                texto
            )

        )

        if separadores > 0:
            return False

        # ----------------------------------------------------
        # EXCESSO DE LETRAS MAIÚSCULAS
        # ----------------------------------------------------

        letras = [

            caractere

            for caractere
            in texto

            if caractere.isalpha()

        ]

        if len(letras) >= 30:

            maiusculas = sum(

                1

                for caractere
                in letras

                if caractere.isupper()

            )

            percentual_maiusculas = (

                maiusculas
                /
                len(letras)

            )

            if percentual_maiusculas > 0.55:
                return False

        # ----------------------------------------------------
        # FRASES
        # ----------------------------------------------------

        quantidade_frases = len(

            re.findall(
                r"[.!?]",
                texto
            )

        )

        if quantidade_frases < 1:
            return False

        # ----------------------------------------------------
        # PALAVRAS MUITO CURTAS
        # ----------------------------------------------------

        palavras_reais = [

            palavra

            for palavra
            in re.findall(
                r"\b\w+\b",
                texto
            )

        ]

        if palavras_reais:

            palavras_curtas = sum(

                1

                for palavra
                in palavras_reais

                if len(palavra) <= 2

            )

            percentual_curtas = (

                palavras_curtas
                /
                len(palavras_reais)

            )

            if percentual_curtas > 0.40:
                return False

        return True      
    
        
    # ========================================================
    # FUNÇÃO CENTRAL DE ESCOLHA
    # ========================================================

    def selecionar_melhor_candidato(
        candidatos_bloco,
        chave_bloco
    ):

        melhor_candidato = None

        melhor_nota = None

        melhor_pontuacao_original = None

        for item in candidatos_bloco:

            candidato = item.get(
                "candidato"
            )

            if not isinstance(
                candidato,
                dict
            ):

                continue
                
            # ------------------------------------------------
            # DESCARTAR TRECHOS ESTRUTURALMENTE RUINS
            #
            # A pontuação continua sendo responsável pela
            # ordem de tentativa.
            #
            # Se o primeiro candidato for ruim, simplesmente
            # continuamos para o próximo.
            # ------------------------------------------------

            if not candidato_eh_utilizavel(
                candidato
            ):

                continue    

            hash_trecho = candidato.get(
                "hash",
                ""
            )

            if not hash_trecho:

                hash_trecho = gerar_hash_trecho(
                    candidato.get(
                        "texto",
                        ""
                    )
                )

                candidato["hash"] = (
                    hash_trecho
                )

            # ------------------------------------------------
            # NÃO REPETIR O MESMO FRAGMENTO
            # ------------------------------------------------

            if hash_trecho in hashes_selecionados:

                continue

            pontuacao_original = item.get(
                "pontuacao",
                0
            )

            nota_final = (
                calcular_nota_selecao(
                    candidato,
                    pontuacao_original,
                    chave_bloco
                )
            )

            # ------------------------------------------------
            # ESCOLHER O MELHOR
            # ------------------------------------------------

            if (
                melhor_candidato is None
                or
                nota_final > melhor_nota
                or
                (
                    nota_final == melhor_nota
                    and
                    pontuacao_original
                    >
                    melhor_pontuacao_original
                )
            ):

                melhor_candidato = (
                    candidato
                )

                melhor_nota = (
                    nota_final
                )

                melhor_pontuacao_original = (
                    pontuacao_original
                )

        if melhor_candidato is None:

            return None

        # ----------------------------------------------------
        # GUARDAR INFORMAÇÕES DE AUDITORIA
        # ----------------------------------------------------

        melhor_candidato[
            "pontuacao_selecao"
        ] = melhor_nota

        melhor_candidato[
            "pontuacao_editorial"
        ] = melhor_pontuacao_original

        melhor_candidato[
            "penalidade_fonte"
        ] = calcular_penalidade_fonte(
            melhor_candidato
        )

        melhor_candidato[
            "penalidade_semelhanca"
        ] = calcular_penalidade_semelhanca(
            melhor_candidato
        )

        melhor_candidato[
            "penalidade_assunto"
        ] = calcular_penalidade_assunto(
            melhor_candidato,
            chave_bloco
        )

        return melhor_candidato

    # ========================================================
    # ========================================================
    # 07. ALOCAÇÃO GLOBAL 5 x 3
    # ========================================================
    fragmentos_selecionados=[]; hashes_selecionados=set(); fontes_utilizadas={}
    disponibilidade_por_candidato={}
    for chave_bloco, lista in candidatos_por_bloco.items():
        for item in lista:
            candidato=item.get("candidato",{}); h=candidato.get("hash") or gerar_hash_trecho(candidato.get("texto","")); candidato["hash"]=h
            disponibilidade_por_candidato.setdefault(h,set()).add(chave_bloco)
    for rodada in range(3):
        for numero_bloco in range(1,6):
            chave_bloco=f"bloco_{numero_bloco}"
            if sum(1 for x in fragmentos_selecionados if x.get("bloco_mead")==chave_bloco)>=3: continue
            melhores=[]
            for item in candidatos_por_bloco.get(chave_bloco,[]):
                candidato=item.get("candidato")
                if not candidato_eh_utilizavel(candidato): continue
                if candidato.get("palavra_chave_status") != "COM": continue
                h=candidato.get("hash") or gerar_hash_trecho(candidato.get("texto",""))
                if h in hashes_selecionados: continue
                fonte=candidato.get("fonte"); score=float(item.get("pontuacao",0))
                if fontes_utilizadas.get(fonte,0)==0: score+=35
                score+=(12-min(12,len(disponibilidade_por_candidato.get(h,set()))))*1.5
                if candidato.get("pdf"): score+=2
                for escolhido in fragmentos_selecionados:
                    if escolhido.get("fonte")==fonte: score-=8
                    sem=calcular_semelhanca_fragmentos(candidato.get("texto",""),escolhido.get("texto",""))
                    if sem>=0.55: score-=15
                    elif sem>=0.40: score-=6
                melhores.append((score,candidato,item.get("pontuacao",0)))
            if not melhores: continue
            melhores.sort(key=lambda x:(x[0],x[2]),reverse=True); score,escolhido,pont=melhores[0]
            escolhido["bloco_mead"]=chave_bloco; escolhido["pontuacao_selecao"]=round(score,3); escolhido["pontuacao_editorial"]=pont
            escolhido["penalidade_fonte"]=0; escolhido["penalidade_semelhanca"]=0; escolhido["penalidade_assunto"]=0
            h=escolhido["hash"]; hashes_selecionados.add(h); fonte=escolhido.get("fonte"); fontes_utilizadas[fonte]=fontes_utilizadas.get(fonte,0)+1; fragmentos_selecionados.append(escolhido)
    if len(fragmentos_selecionados)<15:
        for numero_bloco in range(1,6):
            chave_bloco=f"bloco_{numero_bloco}"
            while sum(1 for x in fragmentos_selecionados if x.get("bloco_mead")==chave_bloco)<3:
                opcoes=[]
                for item in candidatos_por_bloco.get(chave_bloco,[]):
                    candidato=item.get("candidato")
                    if not candidato_eh_utilizavel(candidato): continue
                    # Após esgotar candidatos COM, completar com SEM sem
                    # reclassificá-los: a fonte pode sustentar fatos mesmo sem
                    # conter a expressão literal do tema.
                    status_fonte = str(candidato.get("palavra_chave_status", "SEM")).upper()
                    h=candidato.get("hash") or gerar_hash_trecho(candidato.get("texto",""))
                    if h in hashes_selecionados: continue
                    score=float(item.get("pontuacao",0))+(15 if fontes_utilizadas.get(candidato.get("fonte"),0)==0 else 0)+(30 if status_fonte == "COM" else 0); opcoes.append((score,candidato,item.get("pontuacao",0)))
                if not opcoes: break
                opcoes.sort(key=lambda x:(x[0],x[2]),reverse=True); score,escolhido,pont=opcoes[0]
                escolhido["bloco_mead"]=chave_bloco; escolhido["pontuacao_selecao"]=round(score,3); escolhido["pontuacao_editorial"]=pont
                h=escolhido["hash"]; hashes_selecionados.add(h); fonte=escolhido.get("fonte"); fontes_utilizadas[fonte]=fontes_utilizadas.get(fonte,0)+1; fragmentos_selecionados.append(escolhido)
    print("ALLOCAÇÃO GLOBAL 5 x 3"); print("CANDIDATOS TOTAIS:",len(candidatos))
    for numero_bloco in range(1,6):
        chave_bloco=f"bloco_{numero_bloco}"; print(chave_bloco,":",sum(1 for x in fragmentos_selecionados if x.get("bloco_mead")==chave_bloco))
    print("TOTAL:",len(fragmentos_selecionados),"/ 15")

    # ========================================================
    # 08. VERIFICAR QUANTIDADE
    # ========================================================

    print()
    print("==============================")
    print("FRAGMENTOS SELECIONADOS")
    print("==============================")

    print(
        "TOTAL:",
        len(fragmentos_selecionados),
        "/ 15"
    )

    if len(fragmentos_selecionados) < 15:
        faltantes = 15 - len(fragmentos_selecionados)
        print("AVISO: material pesquisado insuficiente para preencher todos os 15 apoios.")
        print("FRAGMENTOS REAIS:", len(fragmentos_selecionados), "/ 15 | POSIÇÕES COMPLEMENTADAS PELO CONHECIMENTO GERAL DO TEMA:", faltantes)
        print("O OLLAMA SERÁ EXECUTADO; a falta de hashes COM/SEM não bloqueia a redação.")


    # ========================================================
    # 09. DISTRIBUIR EM 5 BLOCOS
    # ========================================================
    #
    # Cada bloco passa a ser um objeto completo.
    #
    # Estrutura:
    #
    # bloco_1
    #   ├── id
    #   ├── hash
    #   ├── informacoes_relevantes
    #   ├── titulo
    #   └── paragrafos
    #
    # São 3 fragmentos por bloco.
    #
    # ========================================================

    blocos = {

        "bloco_1": {
            "id": "bloco_1",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "paragrafos_python": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        },

        "bloco_2": {
            "id": "bloco_2",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "paragrafos_python": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        },

        "bloco_3": {
            "id": "bloco_3",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "paragrafos_python": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        },

        "bloco_4": {
            "id": "bloco_4",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "paragrafos_python": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        },

        "bloco_5": {
            "id": "bloco_5",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "paragrafos_python": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        }

    }

    contagem_fragmentos_bloco = {f"bloco_{n}": 0 for n in range(1, 6)}

    for fragmento in fragmentos_selecionados:

        # A alocação global já definiu o bloco de cada hash.
        # Não redistribuir pela posição da lista: a seleção é
        # intercalada entre blocos e o índice sequencial misturava
        # evidências de B1, B2, B3... dentro de um mesmo bloco.
        chave_bloco = str(fragmento.get("bloco_mead", "") or "")
        if chave_bloco not in blocos:
            print("❌ FRAGMENTO SEM BLOCO MEAD VÁLIDO:", fragmento.get("id", ""))
            continue

        if contagem_fragmentos_bloco[chave_bloco] >= 3:
            print("❌ MAIS DE 3 FRAGMENTOS ALOCADOS:", chave_bloco, fragmento.get("id", ""))
            continue

        contagem_fragmentos_bloco[chave_bloco] += 1

        # ----------------------------------------------------
        # PRESERVAR O FRAGMENTO COMPLETO
        # ----------------------------------------------------

        dados_fragmento = {

            "id":
                fragmento["id"],

            "hash":
                fragmento["hash"],

            "palavra_chave_status":
                fragmento.get("palavra_chave_status", "SEM"),

            "palavra_chave_ocorrencias_fonte":
                fragmento.get("palavra_chave_ocorrencias_fonte", 0),

            "texto":
                fragmento["texto"],

            "fonte":
                fragmento["fonte"],

            "url":
                fragmento["url"],

            "tipo":
                fragmento["tipo"],

            "pdf":
                fragmento["pdf"],

            "palavras":
                fragmento["palavras"],

            "identidade_fonte":
                fragmento.get(
                    "identidade_fonte",
                    {}
                )

        }

        # ----------------------------------------------------
        # COLOCAR O FRAGMENTO NO BLOCO CORRESPONDENTE
        # ----------------------------------------------------

        blocos[
            chave_bloco
        ][
            "informacoes_relevantes"
        ].append(
            dados_fragmento
        )
        
    # ========================================================
    # 09.0.1 PRESERVAR TRECHOS SELECIONADOS PELO PYTHON
    # ========================================================
    #
    # Os 15 fragmentos já foram selecionados anteriormente.
    #
    # NÃO HÁ NOVA SELEÇÃO AQUI.
    #
    # O Python NÃO escreve parágrafos.
    # O Python NÃO resume os fragmentos.
    # O Python NÃO reescreve os fragmentos.
    # O Python apenas preserva o campo "texto"
    # dos fragmentos que já foram selecionados.
    #
    # 15 fragmentos = 15 trechos selecionados
    # 5 blocos × 3 trechos
    #
    # O Ollama receberá posteriormente esses trechos
    # para produzir os textos editoriais.
    #
    # ========================================================

    for numero_bloco in range(1, 6):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        fragmentos_bloco = (
            blocos[
                chave_bloco
            ][
                "informacoes_relevantes"
            ]
        )

        # ----------------------------------------------------
        # PRESERVAR SOMENTE OS TRECHOS SELECIONADOS
        # ----------------------------------------------------
        #
        # Cada posição recebe exatamente o conteúdo
        # do campo "texto" do fragmento selecionado.
        #
        # Nenhuma redação é feita pelo Python.
        #
        # ----------------------------------------------------

        paragrafos_python = [
            fragmento.get(
                "texto",
                ""
            )
            for fragmento in fragmentos_bloco
        ]

        # ----------------------------------------------------
        # GARANTIR AS 3 POSIÇÕES DO BLOCO
        # ----------------------------------------------------

        paragrafos_python = (
            paragrafos_python
            + [
                "",
                "",
                ""
            ]
        )[:3]

        blocos[
            chave_bloco
        ][
            "paragrafos_python"
        ] = paragrafos_python


    # ========================================================
    # RESUMO DOS TRECHOS PYTHON — LOG COMPACTO
    # ========================================================
    totais_trechos_bloco = [
        len([x for x in blocos[f"bloco_{n}"].get("paragrafos_python", []) if str(x or "").strip()])
        for n in range(1, 6)
    ]
    total_trechos_python = sum(totais_trechos_bloco)
    print("TRECHOS PYTHON:", totais_trechos_bloco, "| TOTAL:", total_trechos_python, "/ 15")

    # ========================================================
    # GERAR HASH DE CADA BLOCO
    # ========================================================

    for numero_bloco in range(
        1,
        6
    ):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        fragmentos_bloco = (
            blocos[
                chave_bloco
            ][
                "informacoes_relevantes"
            ]
        )

        texto_bloco = " ".join(

            str(
                fragmento.get(
                    "texto",
                    ""
                )
            )

            for fragmento
            in fragmentos_bloco

        ).strip()

        if texto_bloco:

            blocos[
                chave_bloco
            ][
                "hash"
            ] = gerar_hash_trecho(
                texto_bloco
            )

    # ========================================================
    # 09.1 COMPATIBILIDADE COM BLOCOS_INFORMACOES
    # ========================================================
    #
    # PRESERVAR A ESTRUTURA COMPLETA DOS BLOCOS.
    #
    # IMPORTANTE:
    # "paragrafos_python" foi criado acima e contém
    # os 15 parágrafos-base selecionados pelo Python.
    #
    # Ele precisa ser transportado para
    # "blocos_informacoes", pois essa é a estrutura
    # posteriormente utilizada por
    # gerar_conteudo_completo().
    #
    # NÃO FAZER NOVA SELEÇÃO.
    # NÃO ALTERAR OS FRAGMENTOS.
    # NÃO ALTERAR OS PARÁGRAFOS.
    #
    # ========================================================
    
    blocos_informacoes = {}

    # RESERVAS COM GLOBAIS EM MEMÓRIA (SEM DEPENDER DO BANCO GLOBAL EM DISCO).
    # A lista é compartilhada pelos cinco blocos; uma hash tentada por um
    # parágrafo deixa de ficar disponível para os demais, evitando reutilização.
    hashes_reservas_global = set()
    hashes_reservas_ocupadas = set(hashes_selecionados)
    reservas_com_global = []
    for chave_origem, lista_origem in candidatos_por_bloco.items():
        for item_reserva in lista_origem if isinstance(lista_origem, list) else []:
            candidato_reserva = item_reserva.get("candidato", {}) if isinstance(item_reserva, dict) else {}
            if not isinstance(candidato_reserva, dict):
                continue
            texto_reserva = str(candidato_reserva.get("texto", "") or "").strip()
            if not texto_reserva:
                continue
            ocorrencias_reserva = contar_palavra_chave_literal(texto_reserva, tema)
            if (str(candidato_reserva.get("palavra_chave_status", "")).upper() != "COM"
                    or ocorrencias_reserva < 1):
                continue
            hash_reserva = str(candidato_reserva.get("hash") or gerar_hash_trecho(texto_reserva))
            if not hash_reserva or hash_reserva in hashes_reservas_ocupadas or hash_reserva in hashes_reservas_global:
                continue
            if not candidato_eh_utilizavel(candidato_reserva):
                continue
            reservas_com_global.append({
                "id": candidato_reserva.get("id", ""),
                "hash": hash_reserva,
                "palavra_chave_status": "COM",
                "palavra_chave_ocorrencias_fonte": ocorrencias_reserva,
                "texto": texto_reserva,
                "fonte": candidato_reserva.get("fonte", ""),
                "url": candidato_reserva.get("url", ""),
                "tipo": candidato_reserva.get("tipo", ""),
                "pdf": candidato_reserva.get("pdf", False),
                "palavras": candidato_reserva.get("palavras", len(texto_reserva.split())),
                "identidade_fonte": candidato_reserva.get("identidade_fonte", {}),
            })
            hashes_reservas_global.add(hash_reserva)

    print("BANCO GLOBAL EM MEMÓRIA — RESERVAS COM ÚNICAS:", len(reservas_com_global))

    for numero_bloco in range(
        1,
        6
    ):
    
        chave_bloco = (
            f"bloco_{numero_bloco}"
        )
    
        dados_bloco = (
            blocos.get(
                chave_bloco,
                {}
            )
        )
    
        if not isinstance(
            dados_bloco,
            dict
        ):
    
            dados_bloco = {}
    
        # ----------------------------------------------------
        # PRESERVAR INFORMAÇÕES RELEVANTES
        # ----------------------------------------------------
    
        informacoes_relevantes_bloco = (
            dados_bloco.get(
                "informacoes_relevantes",
                []
            )
        )
    
        if not isinstance(
            informacoes_relevantes_bloco,
            list
        ):
    
            informacoes_relevantes_bloco = []
    
        # ----------------------------------------------------
        # PRESERVAR OS PARÁGRAFOS-BASE DO PYTHON
        # ----------------------------------------------------
    
        paragrafos_python = (
            dados_bloco.get(
                "paragrafos_python",
                []
            )
        )
    
        if not isinstance(
            paragrafos_python,
            list
        ):
    
            paragrafos_python = []
    
        paragrafos_python = [
    
            str(paragrafo or "").strip()
    
            for paragrafo
            in paragrafos_python
    
            if str(paragrafo or "").strip()
    
        ]
    
        # ----------------------------------------------------
        # PRESERVAR ESTRUTURA COMPLETA
        # ----------------------------------------------------
    
        # Todos os blocos consultam a mesma reserva global em memória.
        # Não há cotas exclusivas por bloco nem limite artificial de 30.
        print("RESERVAS COM GLOBAIS DISPONÍVEIS", chave_bloco, ":", len(reservas_com_global))

        blocos_informacoes[
            chave_bloco
        ] = {
    
            "id": dados_bloco.get("id", chave_bloco),
            "hash": dados_bloco.get("hash", ""),
            "informacoes_relevantes": informacoes_relevantes_bloco,
            "reservas_com": reservas_com_global,
            "titulo": dados_bloco.get("titulo", ""),
            "paragrafos_python": paragrafos_python,
            "paragrafos_ollama": dados_bloco.get("paragrafos_ollama", ["", "", ""])
        }
    
    
    # --------------------------------------------------------
    # DEBUG — CONFIRMAR TRANSPORTE DOS 15 PARÁGRAFOS
    # --------------------------------------------------------
    
    print()
    print("==============================")
    print("TRANSFERÊNCIA DOS PARÁGRAFOS-BASE")
    print("==============================")
    
    total_paragrafos_python_transferidos = 0
    
    for numero_bloco in range(
        1,
        6
    ):
    
        chave_bloco = (
            f"bloco_{numero_bloco}"
        )
    
        paragrafos_python = (
            blocos_informacoes[
                chave_bloco
            ].get(
                "paragrafos_python",
                []
            )
        )
    
        quantidade = len(
            [
                p
                for p in paragrafos_python
                if str(p).strip()
            ]
        )
    
        total_paragrafos_python_transferidos += (
            quantidade
        )
    
        print(
            chave_bloco,
            ":",
            quantidade,
            "parágrafos-base Python"
        )
    
    print()
    print(
        "TOTAL PARÁGRAFOS-BASE PYTHON TRANSFERIDOS:",
        total_paragrafos_python_transferidos
    )



    # ========================================================
    # 10. MOSTRAR DISTRIBUIÇÃO
    # ========================================================

    print()
    print("==============================")
    print("DISTRIBUIÇÃO DOS FRAGMENTOS")
    print("==============================")

    for numero in range(
        1,
        6
    ):

        chave = f"bloco_{numero}"

        quantidade = len(
            blocos[chave].get(
                "informacoes_relevantes",
                []
            )
        )

        print(
            chave,
            ":",
            quantidade,
            "fragmentos"
        )

    # ========================================================
    # 11. MONTAR TEXTO DE COMPATIBILIDADE
    # ========================================================

    partes_texto = []

    for indice, fragmento in enumerate(
        fragmentos_selecionados,
        start=1
    ):

        partes_texto.append(

            f"[FRAGMENTO {indice}]\n"
            f"[ID {fragmento['id']}]\n"
            f"[HASH {fragmento['hash']}]\n"
            f"[FONTE {fragmento['fonte']}]\n"
            f"{fragmento['texto']}"

        )

    material_selecionado = (
        "\n\n".join(
            partes_texto
        ).strip()
    )

    # ========================================================
    # 12. FONTES UTILIZADAS
    # ========================================================

    fontes_resultado = []

    fontes_vistas = set()

    for fragmento in fragmentos_selecionados:

        indice = fragmento[
            "fonte"
        ]

        if indice in fontes_vistas:

            continue

        fontes_vistas.add(
            indice
        )

        fontes_resultado.append({

            "indice":
                fragmento["fonte"],

            "url":
                fragmento["url"],

            "tipo":
                fragmento["tipo"],

            "pdf":
                fragmento["pdf"]

        })

    # ========================================================
    # 13. INFORMAÇÕES_RELEVANTES PRINCIPAIS
    # ========================================================
    #
    # ESTA É A ESTRUTURA QUE SERÁ GRAVADA NO JSON.
    #
    # Cada trecho fica identificado por:
    #
    #   id
    #   hash
    #   texto
    #   fonte
    #   url
    #   tipo
    #   pdf
    #   palavras
    #
    # O HASH É O IDENTIFICADOR PERMANENTE DO CONTEÚDO.
    #
    # ========================================================

    informacoes_relevantes = []

    for fragmento in fragmentos_selecionados:

        informacoes_relevantes.append({

            "id":
                fragmento["id"],

            "hash":
                fragmento["hash"],

            "palavra_chave_status":
                fragmento.get("palavra_chave_status", "SEM"),

            "palavra_chave_ocorrencias_fonte":
                fragmento.get("palavra_chave_ocorrencias_fonte", 0),

            "texto":
                fragmento["texto"],

            "fonte":
                fragmento["fonte"],

            "url":
                fragmento["url"],

            "tipo":
                fragmento["tipo"],

            "pdf":
                fragmento["pdf"],

            "palavras":
                fragmento["palavras"],

            "identidade_fonte":
                fragmento.get(
                    "identidade_fonte",
                    {}
                )

        })

    # ========================================================
    # 14. OBJETO FINAL
    # ========================================================

    resultado = {

        "status":
            (
                "selecionado"
                if len(
                    fragmentos_selecionados
                ) == 15
                else "selecionado_parcial"
            ),

        "caracteres":
            len(
                material_selecionado
            ),

        "texto":
            material_selecionado,

        "fontes":
            fontes_resultado,

        "fragmentos":
            fragmentos_selecionados,

        # ----------------------------------------------------
        # NOVO:
        # INFORMAÇÕES RELEVANTES COM ID + HASH
        # ----------------------------------------------------

        "informacoes_relevantes":
            informacoes_relevantes,

        "blocos":
            blocos,

        "blocos_informacoes":
            blocos_informacoes

    }

    # ========================================================
    # 15. SALVAMENTO INTERMEDIÁRIO DESATIVADO
    # ========================================================
    #
    # IMPORTANTE:
    #
    # Os fragmentos selecionados continuam existindo em memória
    # e continuam sendo utilizados pelo processamento dos 5 blocos.
    #
    # Porém, eles NÃO devem mais ser gravados separadamente
    # no banco nesta etapa.
    #
    # O salvamento oficial acontece somente depois que:
    #
    #   - os 15 fragmentos foram distribuídos;
    #   - os 5 blocos foram processados;
    #   - os 15 parágrafos Ollama foram gerados;
    #   - os segmentos foram gerados;
    #   - as tags foram geradas;
    #   - a estrutura oficial da página foi montada.
    #
    # Portanto:
    #
    #   NÃO chamar salvar_banco() aqui.
    #
    # ========================================================
    
    print()
    print("==============================")
    print("SALVAMENTO INTERMEDIÁRIO")
    print("==============================")
    print("DESATIVADO — AGUARDANDO SALVAMENTO OFICIAL")
    print("FRAGMENTOS SELECIONADOS:", len(fragmentos_selecionados))
    print("==============================")
        
        
    # ========================================================
    # 16. RESUMO COMPACTO DE IDS E HASHES
    # ========================================================
    hashes_selecionadas_log = {
        str(x.get("hash", "")) for x in fragmentos_selecionados
        if isinstance(x, dict) and x.get("hash")
    }
    print("CONTROLE DE HASHES: selecionados", len(fragmentos_selecionados),
          "| hashes únicos", len(hashes_selecionadas_log))

    # ========================================================
    # 17. DEBUG FINAL
    # ========================================================

    print()
    print("==============================")
    print("RESULTADO DA SELEÇÃO")
    print("==============================")

    print(
        "STATUS:",
        resultado["status"]
    )

    print(
        "FRAGMENTOS:",
        len(
            resultado["fragmentos"]
        )
    )

    print(
        "INFORMAÇÕES_RELEVANTES:",
        len(
            resultado[
                "informacoes_relevantes"
            ]
        )
    )

    print(
        "CARACTERES:",
        resultado["caracteres"]
    )

    print(
        "FONTES UTILIZADAS:",
        len(
            resultado["fontes"]
        )
    )

    print("DISTRIBUIÇÃO: informações detalhadas de fontes e hashes omitidas do log.")

    print()
    print("==============================")
    print("SELEÇÃO CONCLUÍDA")
    print("==============================")

    return resultado

# ========================================================
# NORMALIZAÇÃO GLOBAL PARA COMPARAÇÃO EDITORIAL
# ========================================================

def normalizar_assunto_texto(
    valor
):

    valor = str(
        valor or ""
    ).lower()

    substituicoes = str.maketrans(

        "áàãâäéèêëíìîïóòõôöúùûüç",

        "aaaaaeeeeiiiiooooouuuuc"

    )

    return valor.translate(
        substituicoes
    )
    
# ========================================================
# FILTRO DE CONTAMINAÇÃO EDITORIAL
# ========================================================

def fragmento_eh_aproveitavel_editorialmente(
    texto
):

    texto_original = str(
        texto or ""
    ).strip()

    if not texto_original:
        return False

    texto_normalizado = normalizar_assunto_texto(
        texto_original
    )

    # ========================================================
    # 01. NAVEGAÇÃO
    # ========================================================

    marcadores_navegacao = [
        "top of page",
        "skip to main content",
        "skip to content",
        "main content",
        "ir para o conteudo",
        "pular para o conteudo",
        "menu principal",
        "voltar ao topo",
        "back to top"
    ]

    for marcador in marcadores_navegacao:

        if normalizar_assunto_texto(
            marcador
        ) in texto_normalizado:

            return False

    # ========================================================
    # 02. CTA / CONTATO / CONVERSÃO
    # ========================================================

    marcadores_cta = [
        "fale conosco",
        "fale com a gente",
        "fale com nossa equipe",
        "fale com nossa equipa",
        "entre em contato",
        "entre em contacto",
        "contacte-nos",
        "contactar-nos",
        "contacte nossa equipe",
        "contacte a nossa equipa",
        "contate-nos",
        "solicite um orçamento",
        "solicite um orcamento",
        "peça um orçamento",
        "peca um orcamento",
        "saiba mais",
        "clique aqui",
        "ligue agora",
        "compre agora",
        "adquira agora",
        "baixe o pdf",
        "descarregue o pdf"
    ]

    for marcador in marcadores_cta:

        if normalizar_assunto_texto(
            marcador
        ) in texto_normalizado:

            return False

    # ========================================================
    # 03. CONTAMINAÇÃO DE NAVEGAÇÃO / CABEÇALHO
    # ========================================================

    padroes_contaminacao_cabecalho = [
        r"\|\s*skip\s+to\b",
        r"\|\s*ligue\s+agora",
        r"\|\s*buscar\b",
        r"\|\s*menu\b",

        r"^\s*introdu[cç][aã]o\s*:\s*$",
        r"^\s*aplica[cç][oõ]es\s*:\s*$",
        r"^\s*vantagens\s*:\s*$",
        r"^\s*benef[ií]cios\s*:\s*$",
        r"^\s*caracter[ií]sticas\s*:\s*$",
        r"^\s*especifica[cç][oõ]es\s*:\s*$",
        r"^\s*recursos\s*:\s*$",
        r"^\s*diferenciais\s*:\s*$",

        r"\bpor\s+que\s+este\s+guia\b",
        r"\bpor\s+que\s+este\s+artigo\b",

        r"\bno\s+pr[oó]ximo\s+post\b",
        r"\bpr[oó]ximo\s+post\b",

        r"\bassista\s+no\s+canal\b",
        r"\bcanal\s+youtube\b"
    ]

    for padrao in padroes_contaminacao_cabecalho:

        if re.search(
            padrao,
            texto_original,
            re.IGNORECASE
        ):

            return False

    # ========================================================
    # 04. MODELOS / CÓDIGOS COMERCIAIS
    # ========================================================

    padroes_modelo = [
        r"\b[A-Z]{2,}(?:/[A-Z0-9]{1,})+(?:-[A-Z0-9]+)+\b",
        r"\b[A-Z]{2,}(?:-[A-Z0-9]+){1,3}\b",
        r"\bmodelo\s+[A-Z0-9][A-Za-z0-9._/-]{2,}\b",
        r"\b(?:linha|série|serie)\s+[A-Z0-9][A-Za-z0-9._/-]{2,}\b"
    ]

    for padrao in padroes_modelo:

        if re.search(
            padrao,
            texto_original
        ):

            return False

    # ========================================================
    # 05. PRODUTO COMERCIAL ESPECÍFICO
    # ========================================================

    padroes_produto_especifico = [
        r"\b(?:a|o)\s+"
        r"[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ]"
        r"[A-Za-zÀ-ÿ-]+\s+"
        r"(?:ultra|super|premium|plus|pro|max|master)\b"
    ]

    for padrao in padroes_produto_especifico:

        if re.search(
            padrao,
            texto_original,
            re.IGNORECASE
        ):

            return False

    # ========================================================
    # 06. LISTAS EMBUTIDAS
    # ========================================================

    marcadores_lista = [
        "aplicações:",
        "aplicacoes:",
        "vantagens:",
        "benefícios:",
        "beneficios:",
        "características:",
        "caracteristicas:",
        "especificações:",
        "especificacoes:",
        "recursos:",
        "diferenciais:"
    ]

    for marcador in marcadores_lista:

        if normalizar_assunto_texto(
            marcador
        ) in texto_normalizado:

            separadores_lista = len(
                re.findall(
                    r"[;•·]",
                    texto_original
                )
            )

            ocorrencias_quebra = len(
                re.findall(
                    r"\s{2,}",
                    texto_original
                )
            )

            if (
                separadores_lista >= 1
                or
                ocorrencias_quebra >= 2
            ):

                return False

    # ========================================================
    # 07. VENDA DIRETA
    # ========================================================

    marcadores_venda = [
        "a solução perfeita",
        "solução ideal para quem busca",
        "ideal para quem busca",
        "perfeito para quem busca",
        "excelente escolha para quem busca",
        "garanta já",
        "garanta o seu",
        "não perca",
        "na nossa loja",
        "em nossa loja",
        "disponível em nossa",
        "disponivel em nossa"
    ]

    for marcador in marcadores_venda:

        if normalizar_assunto_texto(
            marcador
        ) in texto_normalizado:

            return False

    # ========================================================
    # 08. CHAMADAS EDITORIAIS EXTERNAS
    # ========================================================

    marcadores_chamada = [
        "descubra como",
        "confira as vantagens",
        "confira os principais",
        "entenda como",
        "neste artigo, vamos",
        "neste artigo vamos",
        "neste guia, vamos",
        "neste guia vamos",
        "vamos explorar",
        "veja como",
        "en este artículo",
        "en este articulo",
        "en este post",
        "en el próximo post",
        "en el proximo post",
        "contacta con nuestro equipo",
        "contacte con nuestro equipo"
    ]

    ocorrencias_chamada = 0

    for marcador in marcadores_chamada:

        if normalizar_assunto_texto(
            marcador
        ) in texto_normalizado:

            ocorrencias_chamada += 1

    if ocorrencias_chamada >= 2:

        return False

    # ========================================================
    # 09. ESPANHOL / MATERIAL ESTRANGEIRO
    # ========================================================

    marcadores_espanhol = [
        "son esenciales",
        "importancia en diversas industrias",
        "se utilizan ampliamente",
        "se puede evitar",
        "mediante el uso",
        "mantenimiento periódico",
        "mantenimiento periodico",
        "nuestro equipo",
        "presupuesto gratuito",
        "problemas mencionados",
        "rendimiento de la bomba",
        "evitar la contaminación",
        "evitar la contaminacion"
    ]

    for marcador in marcadores_espanhol:

        if normalizar_assunto_texto(
            marcador
        ) in texto_normalizado:

            return False

    # ========================================================
    # 10. FRAGMENTOS CONHECIDAMENTE CORROMPIDOS
    # ========================================================

    padroes_fragmento_corrompido = [
        r"\bevitan-se\s+inesperadas\b",
        r"\bevitar\s+inesperadas\b",
        r"\bna\s+voluta\s+do\s+fluido\s+da\s+bomba\s+entra\s+a\s+bomba\b",
        r"\bas\s+bomba\s+centr[ií]fuga\s+sanit[aá]ria\b"
    ]

    for padrao in padroes_fragmento_corrompido:

        if re.search(
            padrao,
            texto_original,
            re.IGNORECASE
        ):

            return False

    # ========================================================
    # 11. FRAGMENTO INCOMPLETO
    # ========================================================

    marcadores_inicio_invalido = [
        "de ",
        "da ",
        "do ",
        "das ",
        "dos ",
        "e ",
        "ou ",
        "como ",
        "que ",
        "para ",
        "por "
    ]

    inicio_invalido = any(
        texto_normalizado.startswith(
            marcador
        )
        for marcador
        in marcadores_inicio_invalido
    )

    if (
        inicio_invalido
        and
        len(
            texto_original.split()
        ) >= 50
    ):

        if not re.search(
            r"\b(?:é|são|pode|podem|"
            r"permite|permitem|"
            r"consiste|consistem|"
            r"funciona|funcionam|"
            r"possui|possuem|"
            r"apresenta|apresentam)\b",
            texto_normalizado,
            re.IGNORECASE
        ):

            return False

    # ========================================================
    # FRAGMENTO APROVEITÁVEL
    # ========================================================

    return True    

# ============================================================
# GERAR SEGMENTOS DA PÁGINA
# ============================================================
#
# REGRA MEAD:
#
# - A origem dos segmentos é exclusivamente a biblioteca fixa.
# - Produtos usam SEGMENTOS_PRODUTOS + SEGMENTOS_CORINGAS.
# - Serviços usam SEGMENTOS_SERVICOS + SEGMENTOS_CORINGAS.
# - O Python substitui [TEMA].
# - O Python sorteia exatamente 12 segmentos.
# - Não pode haver repetição dentro da mesma página.
# - Os checkboxes NÃO participam da criação dos segmentos.
# - O Ollama NÃO participa da criação dos segmentos.
#
# ============================================================

def gerar_segmentos_pagina(
    tema,
    tipo=None
):

    tema = str(
        tema or ""
    ).strip()

    if not tema:
        return []

    # ========================================================
    # OBTER SEGMENTOS EXCLUSIVAMENTE DA BIBLIOTECA FIXA
    # ========================================================

    segmentos_validos = gerar_segmentos_validos(
        tema,
        tipo
    )

    if not isinstance(segmentos_validos, list):
        segmentos_validos = []

    # ========================================================
    # LIMPAR DUPLICIDADES
    # ========================================================

    segmentos_unicos = []
    vistos = set()

    for segmento in segmentos_validos:
        segmento = str(segmento or "").strip()
        if not segmento:
            continue

        chave = segmento.casefold()
        if chave in vistos:
            continue

        vistos.add(chave)
        segmentos_unicos.append(segmento)

    if len(segmentos_unicos) < 12:
        print()
        print("❌ ERRO: biblioteca fixa possui menos de 12 segmentos válidos.")
        print(f"TEMA: {tema}")
        print(f"SEGMENTOS DISPONÍVEIS: {len(segmentos_unicos)}")
        return []

    # ========================================================
    # GARANTIR QUE TODOS REFERENCIAM O TEMA
    # ========================================================

    segmentos_com_tema = [
        segmento
        for segmento in segmentos_unicos
        if tema.casefold() in segmento.casefold()
    ]

    if len(segmentos_com_tema) < 12:
        print()
        print("❌ ERRO: biblioteca fixa não possui 12 segmentos com referência ao tema.")
        print(f"TEMA: {tema}")
        print(f"SEGMENTOS VÁLIDOS: {len(segmentos_com_tema)}")
        return []

    # ========================================================
    # PRIORIZAÇÃO TEMÁTICA RESTAURADA DA LINHAGEM V100
    # Para bombas e equipamentos de fluidos, prioriza aplicações
    # hidráulicas, transferência, água, saneamento e indústria.
    # A seleção usa somente rótulos cadastrados na biblioteca.
    # ========================================================

    tema_chave = re.sub(r"[^a-z0-9 ]", " ", tema.casefold())
    tema_chave = re.sub(r"\s+", " ", tema_chave).strip()

    termos_bomba = (
        "bomba", "bombas", "centrifuga", "centrífuga",
        "bombeamento", "hidraul", "hidrául", "fluido", "água"
    )
    tema_e_bomba = any(
        termo in tema_chave
        for termo in termos_bomba
    )

    if tema_e_bomba:
        termos_prioritarios = (
            "sistemas hidráulicos",
            "transferência de fluidos",
            "tratamento de água",
            "saneamento",
            "processos industriais",
            "indústrias químicas",
            "indústrias alimentícias",
            "mineração",
            "agricultura e irrigação",
            "abastecimento de água",
            "drenagem",
            "geração de energia",
            "instalações industriais",
            "manutenção industrial",
            "refrigeração",
            "operação contínua",
            "processos de produção",
            "ambientes industriais"
        )

        pesos_prioridade = {
            termo: len(termos_prioritarios) - indice
            for indice, termo in enumerate(termos_prioritarios)
        }

        # Ordenação estável: quando dois rótulos têm a mesma prioridade,
        # preserva a ordem da biblioteca oficial.
        segmentos_com_tema.sort(
            key=lambda item: max(
                (
                    peso
                    for termo, peso in pesos_prioridade.items()
                    if termo in item.casefold()
                ),
                default=0
            ),
            reverse=True
        )

    # Não sortear rótulos genéricos. Selecionar os 12 mais relevantes
    # conforme a ordem oficial e a priorização temática acima.
    segmentos_finais = segmentos_com_tema[:12]

    if len(segmentos_finais) != 12:
        print("❌ ERRO: seleção final não produziu exatamente 12 segmentos.")
        print(f"TEMA: {tema}")
        print(f"TOTAL: {len(segmentos_finais)}")
        return []

    # ========================================================
    # AUDITORIA
    # ========================================================

    print()
    print("==========================================")
    print("SEGMENTOS DA PÁGINA")
    print("==========================================")
    print(f"TEMA: {tema}")
    print(f"TIPO: {tipo or 'não informado'}")
    print(f"BIBLIOTECA FIXA DISPONÍVEL: {len(segmentos_unicos)}")
    print(f"SEGMENTOS COM TEMA: {len(segmentos_com_tema)}")
    print("SEGMENTOS SELECIONADOS: 12")

    for numero, segmento in enumerate(segmentos_finais, start=1):
        print(f"SEGMENTO_{numero}: {segmento}")

    print("==========================================")

    return segmentos_finais


# ============================================================
# FORMAS GRAMATICAIS DO TEMA
# ============================================================
#
# Função independente.
#
# NÃO fica dentro de gerar_conteudo_completo().
#
# Objetivo:
# - identificar masculino/feminino do tema;
# - fornecer as formas necessárias para segmentos e tags;
# - evitar o uso do tema sozinho;
# - manter concordância dos adjetivos.
#
# ============================================================

def obter_formas_gramaticais_tema(tema):

    import re

    tema = re.sub(
        r"\s+",
        " ",
        str(tema or "").strip()
    )

    if not tema:
        return {
            "tema": "",
            "genero": "masculino",
            "artigo": "o",
            "artigo_com_tema": "o",
            "de": "do",
            "de_com_tema": "do",
            "em": "no",
            "em_com_tema": "no",
            "especializado": "especializado",
            "tecnico": "técnico",
            "preventivo": "preventivo",
            "corretivo": "corretivo",
            "adequado": "adequado"
        }

    palavras = tema.lower().split()

    primeira_palavra = palavras[0]

    # --------------------------------------------------------
    # PALAVRAS FEMININAS CONHECIDAS
    # --------------------------------------------------------

    palavras_femininas = {
        "bomba",
        "válvula",
        "valvula",
        "máquina",
        "maquina",
        "empresa",
        "instalação",
        "instalacao",
        "manutenção",
        "manutencao",
        "assistência",
        "assistencia",
        "consultoria",
        "inspeção",
        "inspecao",
        "calibração",
        "calibracao",
        "engenharia",
        "solução",
        "solucao",
        "recuperação",
        "recuperacao",
        "adequação",
        "adequacao",
        "aplicação",
        "aplicacao",
        "operação",
        "operacao",
        "seleção",
        "selecao",
        "configuração",
        "configuracao",
        "especificação",
        "especificacao",
        "integração",
        "integracao",
        "proteção",
        "protecao",
        "produção",
        "producao",
        "transmissão",
        "transmissao",
        "gestão",
        "gestao",
        "análise",
        "analise"
    }

    # --------------------------------------------------------
    # PALAVRAS MASCULINAS CONHECIDAS
    # --------------------------------------------------------

    palavras_masculinas = {
        "equipamento",
        "sistema",
        "produto",
        "serviço",
        "servico",
        "processo",
        "motor",
        "compressor",
        "gerador",
        "painel",
        "sensor",
        "controlador",
        "acionamento",
        "dimensionamento",
        "fornecimento",
        "reparo",
        "diagnóstico",
        "diagnostico",
        "projeto",
        "orçamento",
        "orcamento",
        "suporte",
        "atendimento",
        "funcionamento",
        "desempenho",
        "benefício",
        "beneficio",
        "critério",
        "criterio",
        "modelo",
        "fabricante"
    }

    # --------------------------------------------------------
    # IDENTIFICAR GÊNERO
    # --------------------------------------------------------

    if primeira_palavra in palavras_femininas:

        genero = "feminino"

    elif primeira_palavra in palavras_masculinas:

        genero = "masculino"

    elif (
        primeira_palavra.endswith("a")
        or primeira_palavra.endswith("ção")
        or primeira_palavra.endswith("são")
        or primeira_palavra.endswith("ssão")
        or primeira_palavra.endswith("dade")
        or primeira_palavra.endswith("agem")
    ):

        genero = "feminino"

    else:

        genero = "masculino"

    # --------------------------------------------------------
    # FORMAS FEMININAS
    # --------------------------------------------------------

    if genero == "feminino":

        return {
            "tema": tema,
            "genero": "feminino",

            "artigo": "a",
            "artigo_com_tema": f"a {tema}",

            "de": "da",
            "de_com_tema": f"da {tema}",

            "em": "na",
            "em_com_tema": f"na {tema}",

            "especializado": "especializada",
            "tecnico": "técnica",
            "preventivo": "preventiva",
            "corretivo": "corretiva",
            "adequado": "adequada"
        }

    # --------------------------------------------------------
    # FORMAS MASCULINAS
    # --------------------------------------------------------

    return {
        "tema": tema,
        "genero": "masculino",

        "artigo": "o",
        "artigo_com_tema": f"o {tema}",

        "de": "do",
        "de_com_tema": f"do {tema}",

        "em": "no",
        "em_com_tema": f"no {tema}",

        "especializado": "especializado",
        "tecnico": "técnico",
        "preventivo": "preventivo",
        "corretivo": "corretivo",
        "adequado": "adequado"
    }

# ============================================================
# IDENTIFICAR TIPO DO TEMA
# ============================================================
#
# O Python determina se o tema representa:
#     "produto"
#     "servico"
#
# A decisão é baseada no próprio tema.
# Os textos pesquisados NÃO participam da classificação,
# pois uma página de produto pode conter informações sobre
# instalação, manutenção, reparo, operação etc.
#
# O Ollama NÃO participa desta decisão.
# ============================================================

def identificar_tipo_tema(
    tema
):

    tema = str(
        tema or ""
    ).strip().casefold()

    # ========================================================
    # TEMA VAZIO
    # ========================================================

    if not tema:

        print()
        print("======================================")
        print("TIPO IDENTIFICADO PELO PYTHON")
        print("======================================")
        print("TEMA VAZIO")
        print("TIPO: produto")

        return "produto"

    # ========================================================
    # TERMOS QUE INDICAM SERVIÇO
    # ========================================================

    termos_servico = [
        "manutenção",
        "manutencao",
        "conserto",
        "reparo",
        "assistência",
        "assistencia",
        "instalação",
        "instalacao",
        "montagem",
        "reforma",
        "recuperação",
        "recuperacao",
        "adequação",
        "adequacao",
        "pintura",
        "inspeção",
        "inspecao",
        "calibração",
        "calibracao",
        "consultoria",
        "diagnóstico",
        "diagnostico",
        "usinagem",
        "soldagem",
        "tratamento",
        "limpeza",
        "projeto",
        "serviço",
        "servico",
        "locação",
        "locacao"
    ]

    # ========================================================
    # DECISÃO OFICIAL
    #
    # SOMENTE O TEMA DEFINE O TIPO.
    # ========================================================

    for termo in termos_servico:

        if termo in tema:

            print()
            print("======================================")
            print("TIPO IDENTIFICADO PELO PYTHON")
            print("======================================")
            print("TEMA:", tema)
            print("TIPO: servico")
            print("CRITÉRIO:", termo)

            return "servico"

    # ========================================================
    # PADRÃO
    #
    # Se o tema não indica explicitamente um serviço,
    # ele será tratado como produto.
    # ========================================================

    print()
    print("======================================")
    print("TIPO IDENTIFICADO PELO PYTHON")
    print("======================================")
    print("TEMA:", tema)
    print("TIPO: produto")
    print("CRITÉRIO: tema não identificado como serviço")

    return "produto"

    
# ============================================================
# GERAR CONTEÚDO COMPLETO
# ============================================================





def auditar_evidencia_para_redacao(texto):
    """Triagem determinística adicional antes do Ollama. Não reescreve nem inventa fatos."""
    original = str(texto or "").strip()
    if not original:
        return False, "EVIDENCIA_VAZIA"
    if re.search(r"<\/?(?:p|div|span|h[1-6]|li|ul|ol|table|tr|td)\b|https?://|www\.", original, re.I):
        return False, "HTML_OU_URL_EMBUTIDO"
    if re.search(r"\[/?(?:BLOCO|PARAGRAFO|FRAGMENTO|TEXTO)[^]]*\]", original, re.I):
        return False, "MARCADOR_INTERNO"
    padroes_navegacao = (
        r"\bleia\s+tamb[eé]m\b", r"\bclique\s+aqui\b",
        r"\bsaiba\s+mais\b", r"\bartigos?\s+relacionados\b",
        r"\bposts?\s+relacionados\b", r"\bconfira\s+tamb[eé]m\b",
        r"\bsolicite\s+(?:agora|or[cç]amento)\b", r"\bcotar\s+agora\b"
    )
    if any(re.search(p, original, re.I) for p in padroes_navegacao):
        return False, "NAVEGACAO_OU_CTA_EMBUTIDO"
    # Dois-pontos, isoladamente, não caracterizam cabeçalho.
    # A regra antiga reprovava frases legítimas como
    # "A bomba centrífuga: seu funcionamento...".
    # Só sinalizamos título realmente separado do conteúdo
    # (linha curta terminada em dois-pontos).
    linhas_evidencia = [linha.strip() for linha in original.splitlines() if linha.strip()]
    if len(linhas_evidencia) >= 2:
        primeira_linha = linhas_evidencia[0]
        if (len(primeira_linha.split()) <= 8
                and primeira_linha.endswith((':', '：'))
                and len(' '.join(linhas_evidencia[1:]).split()) >= 8):
            return False, "CABECALHO_SEPARADO_DO_TEXTO"
    if not re.search(r"[.!?][\"'”’)]?$", original):
        return False, "EVIDENCIA_SEM_FECHAMENTO"
    if len(re.findall(r"\n+", original)) >= 4:
        return False, "QUEBRAS_EXCESSIVAS"
    if re.search(
        r"(?<![.!?])\s+(?:O que é|Como funciona|Características técnicas|Aplicações de|Vantagens de|Vantagens da|Princípios de funcionamento|Importância da|Descrição geral de)\b",
        original, re.I
    ):
        return False, "POSSIVEL_TITULO_COLADO"
    # Cabeçalhos de lista e chamadas em caixa alta frequentemente
    # aparecem concatenados ao texto extraído de páginas HTML/PDF.
    if re.search(
        r"^\s*(?:OUTROS?\s+)?(?:BENEF[IÍ]CIOS|VANTAGENS|APLICA[CÇ][OÕ]ES|CARACTER[IÍ]STICAS|DIFERENCIAIS|RECURSOS)(?:\s+EXTRAS?)?\s*:",
        original, re.I
    ):
        return False, "CABECALHO_DE_LISTA_NO_INICIO"
    if re.search(r"(?:^|[.!?]\s)[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][A-ZÁÀÂÃÉÊÍÓÔÕÚÇ _-]{7,}:\s", original):
        return False, "TITULO_EM_CAIXA_ALTA_EMBUTIDO"
    if re.search(r"\b(?:CSHI\s+Bombas|Pumps\s+Brasil|Goulds\s+Pumps)\b", original, re.I):
        return False, "IDENTIDADE_COMERCIAL_DE_TERCEIRO"
    return True, "OK"


# ============================================================
# GATE DE QUALIDADE EDITORIAL PORTADO DA V100.35
# ============================================================
def avaliar_qualidade_final_python(
    titulo,
    paragrafos,
    fragmentos,
    bloco_mead,
    tema
):
    """Gate final determinístico de qualidade editorial antes da gravação.

    A pontuação é complementar aos gates estruturais/factuais. A nota baixa
    gera alerta e não bloqueia a persistência. A fidelidade é verificada
    contra o patrimônio autorizado global da página.
    """
    titulo = str(titulo or "").strip()
    paragrafos = [str(x or "").strip() for x in (paragrafos or [])]
    fragmentos = fragmentos if isinstance(fragmentos, list) else []
    tema_n = normalizar_assunto_texto(tema or "").strip()
    titulo_n = normalizar_assunto_texto(titulo)
    texto_final = " ".join(paragrafos)
    texto_n = normalizar_assunto_texto(texto_final)
    fonte = " ".join(
        str(f.get("texto", "") or "").strip()
        for f in fragmentos if isinstance(f, dict)
    )
    fonte_n = normalizar_assunto_texto(fonte)

    falhas_criticas = []
    descontos = 0

    # ------------------------------------------------------------
    # 1. METATEXTO — FALHA CRÍTICA
    # ------------------------------------------------------------
    metatexto = [
        r"\bo python\b",
        r"\btrecho selecionado\b",
        r"\btrechos selecionados\b",
        r"\bfragmento selecionado\b",
        r"\bfragmentos selecionados\b",
        r"\bfonte selecionada\b",
        r"\bfontes selecionadas\b",
        r"\bcom base nas informacoes fornecidas\b",
        r"\bcom base no trecho\b",
        r"\bcom base nos trechos\b",
        r"\bo texto acima\b",
        r"\bo trecho acima\b",
        r"\bneste trecho\b",
        r"\bagora que voce ja sabe\b",
        r"\bcomo vimos anteriormente\b",
        r"\bneste artigo, voce\b",
        r"\bneste artigo voce\b",
        r"\bacompanhe\b",
        r"\bpalavra-chave literal\b",
        r"\bpalavra chave literal\b",
        r"\bocorre exatamente uma vez\b",
        r"\baparece exatamente uma vez\b",
        r"\bparagrafo final\b",
        r"\bconforme as regras do prompt\b",
        r"\bretorne exatamente tres paragrafos\b"
    ]
    encontrados_metatexto = [p for p in metatexto if re.search(p, texto_n, re.I)]
    if encontrados_metatexto:
        falhas_criticas.append("METATEXTO_DETECTADO")

    # ------------------------------------------------------------
    # 2. PROMOÇÃO / CTA / INSTITUCIONAL — FALHA CRÍTICA
    # ------------------------------------------------------------
    comercial = [
        r"\bentre em contato\b", r"\bfale conosco\b", r"\bsaiba mais\b",
        r"\bsolicite (?:um )?(?:orcamento|cotacao)\b",
        r"\bpeca (?:um )?(?:orcamento|cotacao)\b",
        r"\bonde comprar\b", r"\bmelhor preco\b", r"\bpreco competitivo\b",
        r"\bvendemos\b",
        r"\bpromocao\b", r"\bem promocao\b",
        r"\bconfira nossa selecao\b",
        r"\bmelhor escolha\b", r"\bescolha inteligente\b",
        r"\bsolucao definitiva\b"
    ]
    ocorrencias_comerciais = [p for p in comercial if re.search(p, texto_n, re.I)]
    if ocorrencias_comerciais:
        # O bloco 1 pode apresentar conteúdo comercial, mas somente quando
        # a própria evidência autorizada sustenta a expressão encontrada.
        # Nos demais blocos, conteúdo comercial/CTA continua proibido.
        if str(bloco_mead or "").strip().lower() != "bloco_1":
            falhas_criticas.append("CONTEUDO_COMERCIAL_OU_CTA")
        else:
            expressoes_sem_evidencia = []
            for padrao in ocorrencias_comerciais:
                for ocorrencia in re.finditer(padrao, texto_n, re.I):
                    expressao = ocorrencia.group(0).strip()
                    if expressao and expressao not in fonte_n:
                        expressoes_sem_evidencia.append(expressao)
            if expressoes_sem_evidencia:
                falhas_criticas.append("COMERCIAL_B1_SEM_EVIDENCIA:" + ",".join(sorted(set(expressoes_sem_evidencia))))

    # ------------------------------------------------------------
    # 3. INTEGRIDADE BÁSICA DO TEXTO FINAL
    # ------------------------------------------------------------
    if len(paragrafos) != 3 or any(not p for p in paragrafos):
        falhas_criticas.append("QUANTIDADE_PARAGRAFOS")

    for i, p in enumerate(paragrafos, 1):
        if not re.search(r"[.!?]$", p):
            falhas_criticas.append(f"PARAGRAFO_{i}_SEM_PONTUACAO")
        if re.search(r"<[^>]+>|https?://|www\.", p, re.I):
            falhas_criticas.append(f"PARAGRAFO_{i}_RESIDUO_ESTRUTURAL")
        if re.match(r"^[,.;:!?)]", p) or (p and p[0].islower()):
            falhas_criticas.append(f"PARAGRAFO_{i}_INICIO_TRUNCADO")

    # ------------------------------------------------------------
    # 3B. RESÍDUOS DE PÁGINA, TEXTO TRUNCADO E REDAÇÃO DE BAIXA QUALIDADE
    # ------------------------------------------------------------
    residuos_web = [
        r"\brelated posts?\b", r"\bsearch about\b", r"\bguia completo\b",
        r"\bobrigado\b", r"\beste site foi fundado\b", r"\bsolicite agora\b",
        r"\bcomo funciona\s*[-–:]\s*defini[cç][aã]o\b", r"\(port\.\)",
        r"\bleia tamb[eé]m\b", r"\bpesquise?\b"
    ]
    if any(re.search(p, texto_n, re.I) for p in residuos_web):
        falhas_criticas.append("RESIDUO_DE_PAGINA_OU_NAVEGACAO")

    defeitos_redacao = [
        r"\bantes para\b",
        r"\bde\s+depois\s+em\b", r"\brotore fechado\b", r"\bcon\.\s*(?:ela|ele|a|o|e|\Z)",
        r"\bind[uú]strias reais\b", r"\balto espectro de setores\b",
        r"\bbaixa m[eé]dia press[aã]o\b", r"\bprocessos industriais con\.?\b",
        r"\ba fornecedor de\b", r"\bliquidos invisiveis\b",
        r"\binvestimento essencial\b",
        r"\bescolhendo a bomba centrifuga adequada envolve\b"
    ]
    if any(re.search(p, texto_n, re.I) for p in defeitos_redacao):
        falhas_criticas.append("DEFEITO_DE_REDACAO_OU_TRECHO_TRUNCADO")

    for i, p in enumerate(paragrafos, 1):
        palavras_p = len(re.findall(r"\b[\wÀ-ÿ'-]+\b", p))
        # Extensão é indicador editorial, não bloqueio de gravação.
        # Não se juntam fatos sem relação apenas para atingir uma quota.
        if palavras_p < 20:
            descontos += 20
        elif palavras_p < 30:
            descontos += 12
        elif palavras_p < 40:
            descontos += 6
        if re.search(r"(?:^|[.!?]\s)(?:guia completo|vantagens da|descri[cç][aã]o geral|princ[ií]pios de design)\s*[:.]?\s*(?:$|[A-Z])", p, re.I):
            descontos += 5

    # ------------------------------------------------------------
    # 3A. PALAVRA-CHAVE LITERAL: EXATAMENTE UMA VEZ POR PARÁGRAFO
    # ------------------------------------------------------------
    # A contagem usa a expressão completa, sem diferenciar maiúsculas/minúsculas
    # e respeitando limites de palavra Unicode. Não usa busca por substring,
    # que contaria indevidamente termos maiores que apenas contêm a expressão.
    tema_literal = str(tema or "").strip()
    if tema_literal:
        padrao_tema_literal = re.compile(
            r"(?<!\w)" + re.escape(tema_literal) + r"(?!\w)",
            flags=re.UNICODE | re.IGNORECASE,
        )
        contagens_tema = [
            len(padrao_tema_literal.findall(p))
            for p in paragrafos
        ]
        for indice, quantidade in enumerate(contagens_tema, 1):
            if quantidade != 1:
                falhas_criticas.append(
                    f"PALAVRA_CHAVE_P{indice}_OCORRENCIAS:{quantidade}"
                )
    else:
        contagens_tema = []
        falhas_criticas.append("PALAVRA_CHAVE_VAZIA")

    # Barreira objetiva para erros de concordância causados pela palavra-chave.
    # Ex.: tema singular "bomba centrífuga" não pode ser antecedido por "as".
    tema_ultimo_token = (tema_n.split()[-1] if tema_n.split() else "")
    if tema_n and not tema_ultimo_token.endswith("s"):
        determinantes_plurais = r"(?:as|umas|essas|muitas|algumas|os|uns|esses|muitos|alguns)"
        padrao_concordancia = r"\b" + determinantes_plurais + r"\s+" + re.escape(tema_n) + r"\b"
        if any(re.search(padrao_concordancia, normalizar_assunto_texto(p), re.I) for p in paragrafos):
            falhas_criticas.append("PALAVRA_CHAVE_CONCORDANCIA_GRAMATICAL")

    # ------------------------------------------------------------
    # 4. AFIRMAÇÕES PROMOCIONAIS / GENERALIDADES FORTES
    # ------------------------------------------------------------
    generalidades = [
        r"\bseveramente prejudicad",
        r"\bsem essa tecnologia\b",
        r"\bessencial para qualquer\b",
        r"\bindispensavel\b",
        r"\bdesempenho superior\b",
        r"\bextremamente versatil",
        r"\bperfeitas? para\b",
        r"\bgarante(m)? sempre\b",
        r"\bsolucao economica e eficiente\b",
        r"\buma das maiores vantagens\b"
    ]
    ocorrencias_generalidade = sum(1 for p in generalidades if re.search(p, texto_n, re.I))
    descontos += min(10, ocorrencias_generalidade * 3)

    # ------------------------------------------------------------
    # 5. FIDELIDADE LEXICAL CONSERVADORA
    # ------------------------------------------------------------
    # Não exige cópia literal. Exige que uma parcela relevante dos termos
    # técnicos/concretos do resultado esteja ancorada nos trechos autorizados.
    stop = {
        "para","como","entre","sobre","essa","esse","isso","esta","este",
        "estas","estes","uma","umas","um","uns","que","com","sem","por",
        "dos","das","do","da","de","em","no","na","nos","nas","ao","aos",
        "e","ou","se","mais","menos","muito","muitos","muitas","pode","podem",
        "ser","sao","são","tem","têm","uma","tambem","também","quando","onde",
        "assim","cada","seu","sua","seus","suas","sendo","deve","devem",
        "pelo","pela","pelos","pelas","porque","como"
    }
    # FIDELIDADE E NÚMEROS SÃO AUDITADOS POR PARÁGRAFO, CONTRA A HASH
    # EXATAMENTE VINCULADA ÀQUELA POSIÇÃO. Não permitir que uma informação
    # presente em P2 autorize uma afirmação não sustentada em P1.
    coberturas_por_paragrafo = []
    numeros_novos = set()
    for _idx_fidelidade, _paragrafo_fidelidade in enumerate(paragrafos):
        _frag_fidelidade = fragmentos[_idx_fidelidade] if _idx_fidelidade < len(fragmentos) else {}
        if isinstance(_frag_fidelidade, dict):
            _fonte_fidelidade = str(_frag_fidelidade.get("texto", "") or "")
        else:
            _fonte_fidelidade = str(_frag_fidelidade or "")
        _fonte_fidelidade_n = normalizar_assunto_texto(_fonte_fidelidade)
        _saida_fidelidade_n = normalizar_assunto_texto(_paragrafo_fidelidade)
        _tokens_fonte_p = {
            x for x in re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]{3,}\b", _fonte_fidelidade_n.lower())
            if x not in stop
        }
        _tokens_saida_p = {
            x for x in re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]{3,}\b", _saida_fidelidade_n.lower())
            if x not in stop
        }
        _compartilhados_p = _tokens_saida_p & _tokens_fonte_p
        _cobertura_p = len(_compartilhados_p) / max(1, len(_tokens_saida_p))
        coberturas_por_paragrafo.append(_cobertura_p)
        if len(_tokens_saida_p) >= 45 and _cobertura_p < 0.18:
            falhas_criticas.append(f"FIDELIDADE_LEXICAL_INSUFICIENTE_P{_idx_fidelidade + 1}")
        elif _cobertura_p < 0.26:
            descontos += 2

        _numeros_fonte_p = set(re.findall(r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)", _fonte_fidelidade_n))
        _numeros_saida_p = set(re.findall(r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)", _saida_fidelidade_n))
        _novos_numeros_p = _numeros_saida_p - _numeros_fonte_p
        if _novos_numeros_p:
            numeros_novos.update(_novos_numeros_p)
            falhas_criticas.append(f"NUMEROS_OU_UNIDADES_NOVOS_P{_idx_fidelidade + 1}")

    cobertura = sum(coberturas_por_paragrafo) / max(1, len(coberturas_por_paragrafo))
    if numeros_novos:
        falhas_criticas.append("NUMEROS_OU_UNIDADES_NOVOS")

    # ------------------------------------------------------------
    # 6. COMPATIBILIDADE MEAD
    # ------------------------------------------------------------
    # Compatibilidade editorial flexível nos cinco blocos. A direção indicada
    # pelo mapa é uma preferência; não há falha crítica por ausência de conteúdo
    # comercial ou de uma categoria específica. A coerência é verificada pela
    # relação com o tema, fidelidade às fontes, integridade e qualidade textual.

    # ------------------------------------------------------------
    # 7. COERÊNCIA TEMA × CONTEÚDO
    # ------------------------------------------------------------
    tema_tokens = [x for x in re.findall(r"\b[\wÀ-ÿ]{4,}\b", tema_n) if x not in stop]
    if tema_tokens and not any(tok in texto_n for tok in tema_tokens):
        falhas_criticas.append("TEMA_AUSENTE_DO_CONTEUDO")

    # ------------------------------------------------------------
    # 8. DUPLICAÇÃO ENTRE PARÁGRAFOS
    # ------------------------------------------------------------
    conjuntos = []
    for p in paragrafos:
        conjuntos.append({x for x in re.findall(r"\b[\wÀ-ÿ]{5,}\b", normalizar_assunto_texto(p).lower()) if x not in stop})
    if len(conjuntos) == 3:
        jaccards = []
        for a, b in ((0,1),(0,2),(1,2)):
            uniao = conjuntos[a] | conjuntos[b]
            jaccards.append(len(conjuntos[a] & conjuntos[b]) / max(1, len(uniao)))
        maior_sobreposicao = max(jaccards)
        if maior_sobreposicao >= 0.82:
            falhas_criticas.append("PARAGRAFOS_REDUNDANTES")
        elif maior_sobreposicao >= 0.58:
            descontos += 10
        elif maior_sobreposicao >= 0.45:
            descontos += 4

    # ------------------------------------------------------------
    # 9. DIVERSIDADE NARRATIVA E REPETIÇÃO DE EXPRESSÕES
    # ------------------------------------------------------------
    # Auditoria conservadora: repetição de frases de 3-5 palavras entre
    # parágrafos do mesmo bloco; não confunde termo técnico isolado com cópia.
    tokens_por_paragrafo = []
    aberturas = []
    for _p in paragrafos:
        _pn = normalizar_assunto_texto(_p).lower()
        _tokens = [t for t in re.findall(r"\b[a-z0-9]{4,}\b", _pn) if t not in stop]
        tokens_por_paragrafo.append(_tokens)
        aberturas.append(_tokens[:4])
    conjuntos_ngramas = []
    for _tokens in tokens_por_paragrafo:
        _ng = set()
        for _n in (3, 4, 5):
            _ng.update(" ".join(_tokens[j:j+_n]) for j in range(max(0, len(_tokens)-_n+1)))
        conjuntos_ngramas.append(_ng)
    repeticoes_ngramas = set()
    if len(conjuntos_ngramas) == 3:
        for _i, _j in ((0,1),(0,2),(1,2)):
            repeticoes_ngramas.update(conjuntos_ngramas[_i] & conjuntos_ngramas[_j])
    repeticoes_ngramas = {x for x in repeticoes_ngramas if len(x.split()) >= 3}
    if len(repeticoes_ngramas) >= 8:
        descontos += 10
    elif len(repeticoes_ngramas) >= 4:
        descontos += 5
    elif len(repeticoes_ngramas) >= 2:
        descontos += 2
    aberturas_repetidas = 0
    if len(aberturas) == 3:
        for _i, _j in ((0,1),(0,2),(1,2)):
            if len(aberturas[_i]) >= 3 and aberturas[_i][:3] == aberturas[_j][:3]:
                aberturas_repetidas += 1
    if aberturas_repetidas:
        descontos += min(4, aberturas_repetidas * 2)
    expressoes_vazias = [
        r"\bdesign inovador\b", r"\bresistencia comprovada\b",
        r"\bmelhor escolha\b", r"\bsolucao ideal\b",
        r"\bqualidade superior\b", r"\btemperaturas extremas\b"
    ]
    generalidades_sem_contexto = [x for x in expressoes_vazias if re.search(x, texto_n, re.I)]
    if generalidades_sem_contexto:
        descontos += min(8, len(generalidades_sem_contexto) * 3)

    # ------------------------------------------------------------
    # NOTA
    # ------------------------------------------------------------
    nota = max(0.0, 10.0 - descontos / 10.0)
    if ocorrencias_generalidade:
        nota = min(nota, 8.9)
    if falhas_criticas:
        nota = min(nota, 7.9)

    # A nota é um indicador editorial, não uma barreira de persistência.
    # Falhas críticas continuam reprovando; nota abaixo de 9,0 será salva
    # com aviso de qualidade insuficiente para revisão posterior.
    aprovado = not falhas_criticas
    return aprovado, round(nota, 2), falhas_criticas, {
        "cobertura_factual_lexical": round(cobertura, 3),
        "generalidades": ocorrencias_generalidade,
        "contagens_palavra_chave_por_paragrafo": contagens_tema,
        "numeros_novos": sorted(numeros_novos),
        "ngrams_repetidos_entre_paragrafos": sorted(repeticoes_ngramas)[:20],
        "aberturas_repetidas": aberturas_repetidas,
        "generalidades_sem_contexto": generalidades_sem_contexto,
        "descontos": descontos
    }



# ============================================================
# MOTOR EDITORIAL MEAD/SEVEN — CAMADA GLOBAL DA PÁGINA
# ============================================================
# Esta camada prepara o patrimônio e o plano da página antes
# das cinco chamadas já existentes ao Ollama. Não faz pesquisa,
# não cria fatos e não acrescenta chamadas de IA.
# ============================================================

def _mead_normalizar_editorial(valor):
    texto = str(valor or "").casefold()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto).strip()


def _mead_tokens_editoriais(texto):
    stop = {
        "para", "como", "entre", "sobre", "essa", "esse", "isso",
        "esta", "este", "estas", "estes", "uma", "umas", "um", "uns",
        "que", "com", "sem", "por", "dos", "das", "do", "da", "de",
        "em", "no", "na", "nos", "nas", "ao", "aos", "e", "ou", "se",
        "mais", "menos", "muito", "muitos", "muitas", "pode", "podem",
        "ser", "sao", "tem", "tambem", "quando", "onde", "assim",
        "cada", "seu", "sua", "seus", "suas", "sendo", "deve", "devem",
        "pelo", "pela", "pelos", "pelas", "porque", "a", "o", "as", "os"
    }
    normalizado = _mead_normalizar_editorial(texto)
    return {
        t for t in re.findall(r"\b[a-z0-9][a-z0-9'-]{2,}\b", normalizado)
        if t not in stop
    }


def _mead_classificar_fragmento_editorial(texto):
    """Categorias são índices sobrepostos, nunca filtros exclusivos."""
    t = _mead_normalizar_editorial(texto)
    vocabulario = {
        "comercial": (
            "fornecedor", "fornecimento", "venda", "comprar", "compra",
            "fabricante", "distribuidor", "orcamento", "cotacao",
            "assistencia", "suporte", "empresa", "servico"
        ),
        "conceito_funcionamento": (
            "funcionamento", "funciona", "principio", "conceito",
            "operacao", "mecanismo", "processo", "acionamento"
        ),
        "caracteristicas_aplicacoes": (
            "aplicacao", "aplicacoes", "utilizado", "utilizacao",
            "caracteristica", "componente", "material", "setor",
            "industria", "equipamento", "sistema"
        ),
        "aspectos_tecnicos_cuidados": (
            "pressao", "vazao", "temperatura", "dimensao", "medida",
            "manutencao", "inspecao", "desgaste", "instalacao",
            "seguranca", "norma", "resistencia", "lubrificacao"
        ),
        "escolha_solucao_suporte": (
            "dimensionamento", "selecao", "escolha", "criterio",
            "avaliacao", "necessidade", "compatibilidade", "projeto",
            "adequacao", "diagnostico", "reparo", "solucao"
        )
    }
    categorias = []
    for categoria, termos in vocabulario.items():
        if any(re.search(r"\b" + re.escape(termo) + r"\b", t) for termo in termos):
            categorias.append(categoria)
    return categorias or ["conhecimento_geral"]


def _mead_preparar_patrimonio_global(tema, textos_relevantes, informacoes_blocos):
    """
    Monta o patrimônio da página com os 15 fragmentos selecionados e
    uma seleção adicional do pool global de reservas COM. Não cria
    reservas novas nem altera arquivos/hash persistidos.
    """
    selecionados = []
    if isinstance(textos_relevantes, dict):
        origem = textos_relevantes.get("fragmentos", [])
        if isinstance(origem, list):
            selecionados.extend(origem)

    reservas = []
    if isinstance(informacoes_blocos, dict):
        for dados in informacoes_blocos.values():
            if not isinstance(dados, dict):
                continue
            for reserva in dados.get("reservas_com", []) or []:
                if not isinstance(reserva, dict):
                    continue
                if str(reserva.get("palavra_chave_status", "")).upper() != "COM":
                    continue
                texto_reserva = str(reserva.get("texto", "") or "").strip()
                if not texto_reserva:
                    continue
                if contar_palavra_chave_literal(texto_reserva, tema) < 1:
                    continue
                reservas.append(reserva)

    # Prioriza os 15 fragmentos já aprovados para os blocos. Em seguida,
    # seleciona conhecimento adicional do pool COM por proximidade lexical
    # ao tema inteiro. Os índices são globais, não cotas por parágrafo.
    tokens_keyword = _mead_tokens_editoriais(
        _mead_perfil_keyword(tema).get("expressao_legivel", tema)
    )
    reservas_unicas = {}
    for item in reservas:
        chave = str(item.get("hash", "") or "").strip()
        if not chave:
            chave = hashlib.sha256(
                _mead_normalizar_editorial(item.get("texto", "")).encode("utf-8")
            ).hexdigest()
        reservas_unicas.setdefault(chave, item)

    selecionados_chaves = set()
    for item in selecionados:
        if not isinstance(item, dict):
            continue
        chave = str(item.get("hash", "") or "").strip()
        if chave:
            selecionados_chaves.add(chave)

    extras = []
    for chave, item in reservas_unicas.items():
        if chave in selecionados_chaves:
            continue
        tokens = _mead_tokens_editoriais(item.get("texto", ""))
        score = len(tokens & tokens_keyword)
        extras.append((score, chave, item))
    extras.sort(key=lambda x: (-x[0], x[1]))

    encontrados = list(selecionados)
    encontrados.extend(item for _, _, item in extras)

    resultado = []
    vistos = set()
    total_caracteres = 0
    max_fragmentos = 45
    max_caracteres = 16000

    for indice, item in enumerate(encontrados):
        if not isinstance(item, dict):
            continue
        texto = re.sub(r"\s+", " ", str(item.get("texto", "") or "")).strip()
        if not texto:
            continue
        hash_item = str(item.get("hash", "") or "").strip()
        chave = hash_item or hashlib.sha256(
            _mead_normalizar_editorial(texto).encode("utf-8")
        ).hexdigest()
        if chave in vistos:
            continue
        # Os fragmentos selecionados originais (primeiros itens) são preservados.
        eh_selecionado = indice < len(selecionados)
        if not eh_selecionado and (
            len(resultado) >= max_fragmentos or
            total_caracteres + len(texto) > max_caracteres
        ):
            continue
        vistos.add(chave)
        resultado.append({
            "id": str(item.get("id", "") or ""),
            "hash": chave,
            "texto": texto,
            "bloco_mead": str(item.get("bloco_mead", "") or ""),
            "categorias": _mead_classificar_fragmento_editorial(texto),
            "keyword_anchor": contar_palavra_chave_literal(texto, tema) > 0,
            "origem_pool": "selecionado" if eh_selecionado else "reserva_global_COM"
        })
        total_caracteres += len(texto)

    return resultado


def _mead_contar_keyword_exata(texto, tema):
    texto = str(texto or "")
    tema = str(tema or "").strip()
    if not texto or not tema:
        return 0
    padrao = re.compile(
        r"(?<!\w)" + re.escape(tema) + r"(?!\w)",
        flags=re.UNICODE | re.IGNORECASE,
    )
    return len(padrao.findall(texto))


def _mead_perfil_keyword(tema):
    """Decompõe a expressão para compreensão sem alterar a keyword original."""
    original = str(tema or "").strip()
    # Alguns cadastros usam prefixo numérico e underscores como identificador.
    # Isso serve apenas à leitura semântica; a keyword original nunca é alterada.
    legivel = re.sub(r"^\d+[-_]+", "", original).replace("_", " ")
    normalizada = _mead_normalizar_editorial(legivel)
    corpo = normalizada
    intencao = "informacional_ou_neutra"
    prefixos = [
        (r"^onde comprar\s+", "compra"),
        (r"^onde encontrar\s+", "localizacao_fornecedor"),
        (r"^fabricantes? de\s+", "fabricante"),
        (r"^fornecedores? de\s+", "fornecedor"),
        (r"^distribuidores? de\s+", "distribuicao"),
        (r"^empresa de\s+", "empresa_prestadora"),
        (r"^empresas de\s+", "empresa_prestadora"),
        (r"^venda de\s+", "comercial"),
        (r"^servicos? de\s+", "servico"),
    ]
    for padrao, tipo_intencao in prefixos:
        novo, n = re.subn(padrao, "", corpo, count=1)
        if n:
            corpo = novo.strip()
            intencao = tipo_intencao
            break

    # Separadores explícitos de aplicação/complemento.
    separadores = [
        r"\s+para\s+", r"\s+no setor\s+", r"\s+na industria\s+",
        r"\s+em aplicacoes?\s+", r"\s+para o setor\s+"
    ]
    nucleo = corpo
    complemento = ""
    for sep in separadores:
        m = re.search(sep, corpo)
        if m:
            nucleo = corpo[:m.start()].strip()
            complemento = corpo[m.end():].strip()
            break

    # Serviços com objeto seguido de aplicação: "reparo de pistão de
    # caminhão de lixo" -> núcleo "reparo de pistão", complemento "caminhão de lixo".
    if not complemento:
        m = re.match(
            r"^(reparo|manutencao|conserto|recuperacao|reforma|inspecao|instalacao|fabricacao|usinagem) de ([^ ]+) de (.+)$",
            corpo
        )
        if m:
            nucleo = f"{m.group(1)} de {m.group(2)}"
            complemento = m.group(3).strip()
            intencao = "servico"

    return {
        "expressao_exata": original,
        "expressao_legivel": re.sub(r"\s+", " ", legivel).strip(),
        "expressao_normalizada": normalizada,
        "intencao_provisoria": intencao,
        "nucleo_provisorio": nucleo or normalizada,
        "complementos_preservados": complemento,
        "termos_na_ordem_original": normalizada.split(),
        "regra": "Decomposição heurística para compreensão; a expressão completa permanece soberana."
    }


def _mead_referencias_horizontais(tema, limite=8):
    """Lê padrões editoriais de outras páginas para evitar redação repetida."""
    resultado = []
    caminho = str(globals().get("ARQUIVO_BANCO", "") or "")
    if not caminho or not os.path.isfile(caminho):
        return resultado
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            banco = json.load(f)
        if not isinstance(banco, dict):
            return resultado
        tema_atual = _mead_normalizar_editorial(tema)
        for chave, pagina in banco.items():
            if len(resultado) >= limite:
                break
            if _mead_normalizar_editorial(chave) == tema_atual or not isinstance(pagina, dict):
                continue
            blocos = pagina.get("blocos", {})
            titulos = []
            aberturas = []
            if isinstance(blocos, dict):
                for numero in range(1, 6):
                    b = blocos.get(f"bloco_{numero}", {})
                    if not isinstance(b, dict):
                        continue
                    titulo_b = str(b.get("titulo", "") or "").strip()
                    if titulo_b:
                        titulos.append(titulo_b)
                    paras = b.get("paragrafos_ollama", []) or []
                    if isinstance(paras, list):
                        for par in paras[:1]:
                            frase = re.split(r"(?<=[.!?])\s+", str(par or "").strip())[0]
                            if frase:
                                aberturas.append(" ".join(frase.split()[:12]))
            resultado.append({
                "pagina": str(chave),
                "titulos_existentes": titulos[:5],
                "aberturas_existentes": aberturas[:5]
            })
    except Exception as erro:
        print("AVISO MEAD — COMPARAÇÃO HORIZONTAL INDISPONÍVEL:", repr(erro))
    return resultado


def _mead_planejar_pagina_global(tema, patrimonio):
    """Plano MEAD/SEVEN: relações, lacunas e distribuição sem roteiro rígido."""
    perfil_keyword = _mead_perfil_keyword(tema)
    blocos = {}
    categorias_globais = []
    for item in patrimonio:
        categorias_globais.extend(item.get("categorias", []))
    contagem_global = {}
    for categoria in categorias_globais:
        contagem_global[categoria] = contagem_global.get(categoria, 0) + 1

    categorias_possiveis = {
        "comercial", "conceito_funcionamento", "caracteristicas_aplicacoes",
        "aspectos_tecnicos_cuidados", "escolha_solucao_suporte",
        "conhecimento_geral"
    }
    lacunas = sorted(categorias_possiveis - set(categorias_globais))

    for numero in range(1, 6):
        chave = f"bloco_{numero}"
        itens = [x for x in patrimonio if x.get("bloco_mead") == chave]
        if not itens:
            itens = patrimonio[(numero - 1) * 3:numero * 3]
        categorias = []
        for item in itens:
            categorias.extend(item.get("categorias", []))
        contagem = {}
        for categoria in categorias:
            contagem[categoria] = contagem.get(categoria, 0) + 1
        foco = sorted(contagem, key=lambda x: (-contagem[x], x))
        ids = [x.get("id") or x.get("hash", "")[:12] for x in itens]
        blocos[chave] = {
            "funcao": "definir pelo conteúdo sustentado, sem função narrativa fixa",
            "ambiente": "identificar somente quando estiver explícito nas evidências",
            "perspectiva": "escolher a perspectiva mais útil sem inventar cenário ou personagem",
            "focos_sustentados": foco[:3] or ["conhecimento_geral"],
            "relacoes_exclusivas": foco,
            "ids_evidencia": ids,
            "paragrafos": [
                {
                    "indice": indice,
                    "fato_origem": "selecionar no patrimônio global autorizado",
                    "caminho_vicinal": "desenvolver uma relação distinta e sustentada",
                    "papel_da_keyword": "uma ocorrência literal, com capitalização idêntica",
                    "risco_redundancia": "comparar com todos os outros parágrafos"
                }
                for indice in range(1, 4)
            ]
        }

    sobreposicoes = []
    grupos = []
    for item in patrimonio:
        grupos.append((item, _mead_tokens_editoriais(item.get("texto", ""))))
    for i in range(len(grupos)):
        for j in range(i + 1, len(grupos)):
            a, ta = grupos[i]
            b, tb = grupos[j]
            if not ta or not tb:
                continue
            jaccard = len(ta & tb) / max(1, len(ta | tb))
            if jaccard >= 0.52:
                sobreposicoes.append({
                    "a": a.get("id") or a.get("hash", "")[:12],
                    "b": b.get("id") or b.get("hash", "")[:12],
                    "similaridade": round(jaccard, 2)
                })

    referencias_horizontais = _mead_referencias_horizontais(tema)
    return {
        "pagina": {
            "keyword": perfil_keyword,
            "setor": "identificar somente se confirmado nos fragmentos",
            "publico": "não presumir; identificar somente se houver evidência",
            "objetivo": "responder ao tema completo com relações verdadeiras e sem redundância",
            "fatos_autorizados": [
                x.get("id") or x.get("hash", "")[:12] for x in patrimonio
            ],
            "lacunas": lacunas,
            "riscos": {
                "sobreposicoes": sobreposicoes[:20],
                "ancora_literal_ausente": not any(x.get("keyword_anchor") for x in patrimonio)
            },
            "estrutura": {"blocos": 5, "paragrafos_por_bloco": 3, "total_paragrafos": 15}
        },
        "tema_exato": str(tema or "").strip(),
        "perfil_keyword": perfil_keyword,
        "total_fragmentos": len(patrimonio),
        "keyword_anchor_disponivel": any(x.get("keyword_anchor") for x in patrimonio),
        "categorias_globais": sorted(contagem_global),
        "blocos": blocos,
        "sobreposicoes_desver": sobreposicoes[:20],
        "referencias_horizontais": referencias_horizontais
    }


def _mead_formatar_contexto_global(tema, patrimonio, plano, limite_caracteres=18000):
    partes = [
        "PATRIMÔNIO GLOBAL AUTORIZADO DA PÁGINA",
        f"TEMA / EXPRESSÃO EXATA: {tema}",
        f"PERFIL SEMÂNTICO PROVISÓRIO: núcleo={plano.get('perfil_keyword', {}).get('nucleo_provisorio', tema)}; "
        f"complementos={plano.get('perfil_keyword', {}).get('complementos_preservados', '')}; "
        f"intenção={plano.get('perfil_keyword', {}).get('intencao_provisoria', 'neutra')}.",
        "A decomposição é heurística; a expressão completa governa o significado. Não generalize nem elimine complementos.",
        "Leia o conjunto antes de redigir. Os fragmentos podem se complementar entre blocos.",
        "Use apenas fatos presentes neste patrimônio; não transforme inferências em fatos."
    ]
    total = 0
    for indice, item in enumerate(patrimonio, 1):
        trecho = item["texto"]
        linha = (
            f"\n[CONHECIMENTO_{indice} | {item.get('id') or item.get('hash','')[:12]} | "
            f"CATEGORIAS: {', '.join(item.get('categorias', []))}]\n{trecho}"
        )
        if total + len(linha) > limite_caracteres:
            restante = max(0, limite_caracteres - total)
            if restante > 250:
                partes.append(linha[:restante])
            break
        partes.append(linha)
        total += len(linha)

    partes.append("\nPLANO GLOBAL DE PÁGINA (ORIENTAÇÃO, NÃO FONTE FACTUAL):")
    dados_pagina = plano.get("pagina", {}) if isinstance(plano.get("pagina"), dict) else {}
    partes.append(
        "Objetivo editorial: "
        + str(dados_pagina.get("objetivo", "organizar fatos autorizados"))
    )
    partes.append(
        "Categorias presentes: "
        + ", ".join(plano.get("categorias_globais", []) or [])
    )
    partes.append(
        "Lacunas de evidência: "
        + ", ".join(dados_pagina.get("lacunas", []) or [])
        + ". Não preencher essas lacunas por suposição."
    )
    for chave, dados in plano.get("blocos", {}).items():
        partes.append(
            f"{chave}: focos possíveis={', '.join(dados.get('focos_sustentados', []))}; "
            f"evidências={', '.join(str(x) for x in dados.get('ids_evidencia', []))}."
        )
    if plano.get("sobreposicoes_desver"):
        partes.append(
            "ALERTAS DESVER: há fragmentos com proximidade temática. "
            "Evite repetir o mesmo fato ou explicar a mesma relação em parágrafos diferentes."
        )
    horizontais = plano.get("referencias_horizontais", []) or []
    if horizontais:
        partes.append(
            "COMPARAÇÃO HORIZONTAL: títulos e aberturas abaixo são exemplos de páginas anteriores. "
            "Use-os somente para evitar repetição de moldes; não copie redação nem fatos."
        )
        for pagina in horizontais[:8]:
            partes.append(
                f"PÁGINA ANTERIOR {pagina.get('pagina','')}: "
                f"títulos={pagina.get('titulos_existentes', [])}; "
                f"aberturas={pagina.get('aberturas_existentes', [])}"
            )
    if plano.get("keyword_anchor_disponivel"):
        partes.append(
            f"ÂNCORA: existe ao menos um fragmento autorizado que contém literalmente '{tema}'. "
            "Preserve essa expressão como referência do assunto."
        )
    else:
        partes.append(
            "ÂNCORA: nenhum fragmento selecionado contém literalmente a expressão completa. "
            "Não fabrique evidência; use a expressão exata apenas como requisito SEO e sinalize a lacuna."
        )
    return "\n".join(partes)


def _mead_auditoria_independente(blocos, tema):
    """Auditoria leve pós-redação: DESVER, ISI e Freio do Refinamento."""
    paragrafos = []
    for numero in range(1, 6):
        dados = blocos.get(f"bloco_{numero}", {}) if isinstance(blocos, dict) else {}
        if not isinstance(dados, dict):
            continue
        for indice, texto in enumerate(dados.get("paragrafos_ollama", []) or [], 1):
            texto = str(texto or "").strip()
            if texto:
                paragrafos.append((f"B{numero}P{indice}", texto))

    avisos_desver = []
    avisos_isi = []
    aberturas = {}
    for i, (id_a, texto_a) in enumerate(paragrafos):
        tokens_a = _mead_tokens_editoriais(texto_a)
        abertura = " ".join(texto_a.split()[:4]).casefold()
        if abertura:
            aberturas.setdefault(abertura, []).append(id_a)
        for id_b, texto_b in paragrafos[i + 1:]:
            tokens_b = _mead_tokens_editoriais(texto_b)
            if tokens_a and tokens_b:
                similaridade = len(tokens_a & tokens_b) / max(1, len(tokens_a | tokens_b))
                if similaridade >= 0.68:
                    avisos_desver.append({
                        "paragrafos": [id_a, id_b],
                        "similaridade": round(similaridade, 2)
                    })

    for abertura, ids in aberturas.items():
        if len(ids) > 1:
            avisos_desver.append({
                "abertura_repetida": abertura,
                "paragrafos": ids
            })

    # Sinais para revisão humana; não são reprovação automática, pois
    # algumas palavras podem ter uso literal/técnico no contexto.
    sinais_sensorialidade = (
        "respira", "dança", "pulsa", "coração", "alma", "jornada",
        "viagem", "mágica", "universo", "sabor", "aroma", "textura",
        "abraça", "conversa com", "dialoga com", "ganha vida"
    )
    for identificador, texto in paragrafos:
        encontrados = [
            termo for termo in sinais_sensorialidade
            if re.search(r"\b" + re.escape(termo) + r"\b", texto, re.I)
        ]
        if encontrados:
            avisos_isi.append({
                "paragrafo": identificador,
                "sinais": encontrados
            })

    # Freio: não sugerir reescrita quando não há falha objetiva.
    total_ocorrencias_keyword = sum(
        _mead_contar_keyword_exata(texto, tema) for _, texto in paragrafos
    )
    return {
        "total_paragrafos_auditados": len(paragrafos),
        "desver": avisos_desver[:30],
        "isi": avisos_isi[:30],
        "keyword_ocorrencias_exatas": total_ocorrencias_keyword,
        "freio_do_refinamento": (
            "PRESERVAR: sem motivo objetivo para reescrita global; "
            "corrigir somente unidades com falha confirmada."
        )
    }


MEAD_EDITORES_V2_PROMPT = """
MOTOR EDITORIAL MEAD/SEVEN — ESPECIFICAÇÃO OPERACIONAL
1. Auditar o patrimônio autorizado: fatos, relações, lacunas, restrições e origem.
2. Compreender a expressão-chave inteira: núcleo e todos os complementos. Decompor para compreender, nunca para generalizar ou apagar o sentido.
3. Mapear apenas relações sustentadas: necessidade, função, ambiente, uso, processo, decisão, pessoas, benefício e limite.
4. Planejar a página antes de escrever: objetivo, ambiente, perspectiva, relações e título de cada bloco conforme as evidências.
5. Aplicar DESVER entre parágrafos e blocos: cada informação no lugar mais útil, sem repetição conceitual.
6. Redigir com naturalidade por caminhos editoriais variados; não usar fórmulas ou aberturas mecânicas.
7. Aplicar ISI: revisar sentido literal, imagens sensoriais involuntárias e o cenário produzido pelo conjunto.
8. Validar e reparar somente a unidade com falha; preservar todo conteúdo já aprovado.
9. Aplicar o Freio do Refinamento: não reescrever texto adequado sem erro ou ganho verificável.
A keyword completa, com grafia, acentuação, ordem e capitalização idênticas ao tema, deve aparecer exatamente uma vez em cada parágrafo SEO.
A estrutura de cinco blocos e três parágrafos por bloco é contrato de saída, não roteiro semântico obrigatório.
Cada parágrafo deve ser prosa corrida; não produzir listas, tabelas, fichas técnicas,
bullets, enumerações, subtítulos ou tópicos dentro dos parágrafos SEO.
Quando houver tipos ou variantes de equipamento, não misturar categorias distintas;
no caso de bombas, separar centrífuga geral, sanitária, submersível e autoaspirante,
sem transferir características de uma para outra.
Use voz em primeira pessoa do plural somente quando representar o prestador/empresa e houver evidência para isso.
Não invente especificações, aplicações, materiais, serviços, empresas, marcas, preços, certificações, garantias, disponibilidade ou promessas.
No bloco 1, conteúdo comercial precisa se referir explicitamente à expressão-chave da página,
ser variado entre páginas e estar sustentado pelo patrimônio; não usar CTA genérico como
"Entre em contato" nem afirmar experiência, qualidade, equipe, pagamento ou região sem prova.
Tags e itens de lista só podem ser incluídos quando forem verdadeiros, compatíveis com o tema e sustentados pelo patrimônio; não preencher quotas artificiais.
"""


def gerar_conteudo_completo(
    tema,
    textos,
    mapa_mead,
    estrutura_editorial,
    arquivo_origem=None,
    dados_coleta=None,
    nome_site=""
):

    # Grupo principal vem da interface e precisa existir antes
    # de qualquer salvamento intermediário desta função.
    grupo_principal_projeto = normalizar_grupo_principal_projeto(
        entrada_grupo.get()
        if "entrada_grupo" in globals()
        else ""
    )

    inicio_geracao = time.time()

    # CONTROLE LOCAL DAS RESERVAS COM UTILIZADAS NA RECUPERAÇÃO.
    # A variável criada durante a preparação dos blocos pertence a outra
    # função e não fica visível aqui. Este conjunto precisa existir no
    # escopo de gerar_conteudo_completo e permanecer compartilhado pelos
    # cinco blocos durante toda a geração. Começa com os hashes já
    # selecionados para impedir que uma reserva substituta colida com eles.
    hashes_reservas_ocupadas = set()

    tema_base = str(
        tema or ""
    ).strip()

    # Forma editorial oficial: iniciais maiúsculas, acentos preservados.
    # A chave original (tema_base) continua intacta para banco, busca e tags.
    palavra_chave_redacao = " ".join(
        palavra[:1].upper() + palavra[1:]
        for palavra in re.sub(r"\s+", " ", tema_base).split()
    )
    palavra_chave_tag = tema_base.casefold()

    print()
    print("======================================")
    print("INICIANDO IA - CONTEÚDO COMPLETO")
    print("======================================")
    
    print(
        "TEMA:",
        tema
    )
    
    print(
        "ARQUIVO:",
        arquivo_origem
    )
    
    print(
        "TEXTOS RECEBIDOS:",
        len(textos)
    )
    
    print(
        "MAPA:",
        len(str(mapa_mead))
    )
    
    # ========================================================
    # IDENTIFICAR TIPO DO TEMA
    # ========================================================
    #
    # O Python resolve o tipo no início do processamento.
    # Essa variável será reutilizada em toda a página.
    # ========================================================

    tipo = identificar_tipo_tema(
        tema
    )
    print()
    print("======================================")
    print("TIPO OFICIAL DA PÁGINA")
    print("======================================")
    print(
        "TIPO:",
        tipo
    )


    # PRÉ-VALIDAÇÃO RÁPIDA DAS TAGS: impedir que a execução consuma
    # as 15 chamadas de IA e só depois descubra que faltam tags.
    _tags_preflight_modelos = TAGS_FIXAS_SERVICOS if tipo == "servico" else TAGS_FIXAS_PRODUTOS
    _tags_preflight = [
        re.sub(r"\s+", " ", str(x).replace("[tema]", tema_base).strip()).casefold()
        for x in _tags_preflight_modelos
        if str(x).strip()
    ]
    _tags_preflight_unicas = list(dict.fromkeys(_tags_preflight))
    print(
        "PRÉ-VALIDAÇÃO DE TAGS — CANDIDATOS:",
        len(_tags_preflight_unicas),
        "| QUANTIDADE FINAL: somente itens sustentados"
    )
    
    
    # ========================================================
    # RECUPERAR ARQUIVO_ORIGEM DA PRIMEIRA SELEÇÃO
    # ========================================================
    
    if not arquivo_origem:
    
        if isinstance(textos, dict):
    
            fragmentos = textos.get(
                "fragmentos",
                []
            )
    
            if isinstance(fragmentos, list):
    
                for fragmento in fragmentos:
    
                    if not isinstance(
                        fragmento,
                        dict
                    ):
                        continue
    
                    url = str(
                        fragmento.get(
                            "url",
                            ""
                        ) or ""
                    ).strip()
    
                    if url:
    
                        arquivo_origem = url
    
                        break
    
    
    print(
        "ARQUIVO_ORIGEM RECUPERADO:",
        arquivo_origem
    )


    # ========================================================
    # 01. IDENTIDADE DA PÁGINA
    # ========================================================
    
    if arquivo_origem:
    
        nome_arquivo = str(
            arquivo_origem
        ).strip()
    
        nome_sem_extensao = os.path.splitext(
            os.path.basename(
                arquivo_origem
            )
        )[0]
    
    else:
    
        nome_arquivo = ""
        nome_sem_extensao = tema
    
    print()
    print("======================================")
    print("IDENTIDADE DA PÁGINA")
    print("======================================")
    
    print(
        "ARQUIVO:",
        nome_arquivo
    )
    
    print(
        "IDENTIDADE:",
        normalizar_tema_chave(
            tema
        )
    )




    
    # ========================================================
    # 02. CHECKBOXES EDITORIAIS
    # ========================================================
    
    print()
    print("ESTRUTURA EDITORIAL: conferindo seleções")
    
    if not isinstance(
        estrutura_editorial,
        dict
    ):
    
        estrutura_editorial = {}
    
    assuntos = [
        str(chave).replace("_", " ")
        for chave, valor
        in estrutura_editorial.items()
        if valor
    ]
    
    print()
    print("ASSUNTOS EDITORIAIS:", len(assuntos), "selecionados")
    
    # ========================================================
    # 03. ESTRUTURA OBRIGATÓRIA
    # ========================================================
    
    total_blocos = 5
    paragrafos_por_bloco = 3
    total_paragrafos = 15
    # Quantidades estruturais obrigatórias da página MEAD.
    # Não podem ficar como None: a validação final compara os
    # totais reais com estes limites depois de gerar os itens.
    total_segmentos = 12
    total_tags = 30
    
    print()
    print("======================================")
    print("ESTRUTURA OBRIGATÓRIA")
    print("======================================")
    
    print(
        "BLOCOS:",
        total_blocos
    )
    
    print(
        "PARÁGRAFOS POR BLOCO:",
        paragrafos_por_bloco
    )
    
    print(
        "TOTAL DE PARÁGRAFOS:",
        total_paragrafos
    )
    
    print(
        "SEGMENTOS:",
        total_segmentos
    )
    
    print(
        "TAGS:",
        total_tags
    )
    
    # ========================================================
    # 04. SELECIONAR INFORMAÇÕES RELEVANTES
    # ========================================================
    
    print()
    print("======================================")
    print("SELECIONANDO INFORMAÇÕES RELEVANTES")
    print("======================================")
    
    textos_relevantes = selecionar_informacoes_relevantes(
        tema,
        textos,
        mapa_mead,
        estrutura_editorial
    )
    
    # ========================================================
    # CORREÇÃO:
    # A função selecionar_informacoes_relevantes()
    # retorna um DICIONÁRIO.
    #
    # Portanto, a validação precisa verificar o conteúdo
    # efetivamente selecionado, e não apenas se o dict existe.
    # ========================================================
    
    if not isinstance(
        textos_relevantes,
        dict
    ):
    
        print()
        print("======================================")
        print("NENHUMA INFORMAÇÃO RELEVANTE")
        print("======================================")
    
        return None
    
    contexto_inicial = str(
        textos_relevantes.get(
            "texto",
            ""
        ) or ""
    ).strip()
    
    fragmentos_iniciais = textos_relevantes.get(
        "fragmentos",
        []
    )
    
    if not contexto_inicial and not fragmentos_iniciais:
    
        print()
        print("======================================")
        print("NENHUMA INFORMAÇÃO RELEVANTE")
        print("======================================")
    
        return None
    
    # ========================================================
    # 05. PREPARAR INFORMAÇÕES
    # ========================================================

    print()
    print("======================================")
    print("PREPARANDO INFORMAÇÕES SELECIONADAS")
    print("======================================")

    if not isinstance(
        textos_relevantes,
        dict
    ):

        print()
        print("======================================")
        print("NENHUMA INFORMAÇÃO RELEVANTE")
        print("======================================")

        return None

    contexto = str(
        textos_relevantes.get(
            "texto",
            ""
        ) or ""
    ).strip()

    blocos_informacoes = (
        textos_relevantes.get(
            "blocos_informacoes",
            {}
        )
    )

    if not isinstance(
        blocos_informacoes,
        dict
    ):

        blocos_informacoes = {}

    # ========================================================
    # PRESERVAR OS BLOCOS COMPLETOS
    #
    # IMPORTANTE:
    # Não transformar os blocos em strings aqui.
    #
    # Cada bloco pode conter:
    #
    # informacoes_relevantes
    # paragrafos_python
    #
    # Os paragrafos_python precisam continuar disponíveis
    # para serem enviados posteriormente ao Ollama.
    # ========================================================

    informacoes_blocos = {}

    for numero_bloco in range(
        1,
        6
    ):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        dados_bloco = (
            blocos_informacoes.get(
                chave_bloco,
                {}
            )
        )

        if isinstance(
            dados_bloco,
            dict
        ):

            # ------------------------------------------------
            # PRESERVA O DICIONÁRIO INTEIRO
            # ------------------------------------------------

            informacoes_blocos[
                chave_bloco
            ] = dados_bloco

        elif isinstance(
            dados_bloco,
            list
        ):

            # ------------------------------------------------
            # COMPATIBILIDADE COM FORMATO ANTIGO
            # ------------------------------------------------

            informacoes_blocos[
                chave_bloco
            ] = {
                "informacoes_relevantes": dados_bloco,
                "paragrafos_python": []
            }

        else:

            texto_informacoes = str(
                dados_bloco or ""
            ).strip()

            informacoes_blocos[
                chave_bloco
            ] = {
                "informacoes_relevantes": [
                    texto_informacoes
                ] if texto_informacoes else [],
                "paragrafos_python": []
            }

    # ========================================================
    # RETOMAR BLOCOS PRESERVADOS EM progresso.json
    # ========================================================
    blocos_recuperados = carregar_blocos_recuperacao(tema)
    if blocos_recuperados:
        print("♻️ RECUPERAÇÃO ENCONTRADA:", len(blocos_recuperados), "/ 5 blocos")
        for chave_recuperada, bloco_recuperado in blocos_recuperados.items():
            informacoes_blocos[chave_recuperada] = bloco_recuperado
            print("   -", chave_recuperada, "será reutilizado sem nova chamada Ollama")
    else:
        print("RECUPERAÇÃO: nenhum bloco anterior válido para este tema.")

    # ========================================================
    # CARACTERES DO TEXTO GERAL
    # ========================================================

    caracteres_selecionados = len(
        contexto
    )

    # ========================================================
    # CALCULAR CARACTERES DAS INFORMAÇÕES DOS BLOCOS
    # SEM DESTRUIR A ESTRUTURA DOS BLOCOS
    # ========================================================

    caracteres_blocos = 0

    for numero in range(
        1,
        6
    ):

        dados_bloco = (
            informacoes_blocos.get(
                f"bloco_{numero}",
                {}
            )
        )

        if not isinstance(
            dados_bloco,
            dict
        ):

            continue

        lista_informacoes = (
            dados_bloco.get(
                "informacoes_relevantes",
                []
            )
        )

        if isinstance(
            lista_informacoes,
            list
        ):

            for item in lista_informacoes:

                if isinstance(
                    item,
                    dict
                ):

                    texto_item = str(
                        item.get(
                            "texto",
                            ""
                        ) or ""
                    ).strip()

                else:

                    texto_item = str(
                        item or ""
                    ).strip()

                caracteres_blocos += len(
                    texto_item
                )

        else:

            caracteres_blocos += len(
                str(
                    lista_informacoes or ""
                ).strip()
            )

    # ========================================================
    # CONTAGEM DOS 15 PARÁGRAFOS-BASE PYTHON
    # ========================================================

    total_paragrafos_python = 0

    print()
    print("======================================")
    print("PARÁGRAFOS-BASE PYTHON")
    print("======================================")

    for numero_bloco in range(
        1,
        6
    ):

        dados_bloco = (
            informacoes_blocos.get(
                f"bloco_{numero_bloco}",
                {}
            )
        )

        if not isinstance(
            dados_bloco,
            dict
        ):

            paragrafos_python = []

        else:

            paragrafos_python = (
                dados_bloco.get(
                    "paragrafos_python",
                    []
                )
            )

            if not isinstance(
                paragrafos_python,
                list
            ):

                paragrafos_python = []

        quantidade = len(
            paragrafos_python
        )

        total_paragrafos_python += (
            quantidade
        )

        print(
            f"BLOCO {numero_bloco} "
            f"- PARÁGRAFOS PYTHON:",
            quantidade
        )

    print()
    print(
        "TOTAL PARÁGRAFOS-BASE PYTHON:",
        total_paragrafos_python
    )

    # ========================================================
    # DADOS APÓS SELEÇÃO
    # ========================================================

    print()
    print("======================================")
    print("DADOS APÓS SELEÇÃO")
    print("======================================")

    print(
        "SELEÇÃO RECEBIDA:",
        "SIM"
        if contexto or fragmentos_iniciais
        else "NÃO"
    )

    print(
        "TIPO RECEBIDO:",
        type(
            textos_relevantes
        )
    )

    print(
        "CARACTERES TEXTO:",
        caracteres_selecionados
    )

    print(
        "CARACTERES BLOCOS:",
        caracteres_blocos
    )

    # ========================================================
    # RESUMO ÚNICO DA PREPARAÇÃO — SEM REPETIR CONTAGENS
    # ========================================================
    resumo_blocos = []
    for numero_bloco in range(1, 6):
        dados_bloco = informacoes_blocos.get(f"bloco_{numero_bloco}", {})
        itens = dados_bloco.get("informacoes_relevantes", []) if isinstance(dados_bloco, dict) else []
        resumo_blocos.append(
            f"B{numero_bloco}:{len(itens) if isinstance(itens, list) else 0} trechos"
        )
    print("PREPARAÇÃO:", " | ".join(resumo_blocos),
          "| texto:", caracteres_selecionados, "caracteres")

    # ========================================================
    # 07. CONTEXTO PARA IA
    # ========================================================
    
    print()
    print("======================================")
    print("CONTEXTO PARA IA")
    print("======================================")
    
    contexto_geracao = {}
    
    for numero_bloco in range(
        1,
        6
    ):
    
        chave_bloco = (
            f"bloco_{numero_bloco}"
        )
    
        contexto_geracao[
            chave_bloco
        ] = str(
            informacoes_blocos.get(
                chave_bloco,
                ""
            ) or ""
        ).strip()
    
    # ========================================================
    # CORREÇÃO:
    # Guardamos explicitamente o total de caracteres.
    # len(contexto_geracao) retornaria apenas 5, pois
    # contexto_geracao é um dicionário com 5 chaves.
    # ========================================================
    
    caracteres_contexto_geracao = sum(
        len(
            texto_bloco
        )
        for texto_bloco
        in contexto_geracao.values()
    )
    
    print("CONTEXTO IA:", caracteres_contexto_geracao, "caracteres | 5 blocos")
    
    # ========================================================
    # 08. NORMALIZAR MAPA
    # ========================================================
    
    if isinstance(
        mapa_mead,
        dict
    ):
    
        mapa_texto = str(
            mapa_mead
        )
    
    else:
    
        mapa_texto = str(
            mapa_mead or ""
        )
    
    mapa_texto = mapa_texto[:3000]
    
    # ========================================================
    # 09. PREPARAR MEAD
    # ========================================================
    
    contexto_mead = ""
    
    try:
    
        contexto_mead = preparar_mead(MEAD)
    
    except Exception as erro:
    
        print()
        print("======================================")
        print("ERRO AO PREPARAR MEAD PARA IA")
        print("======================================")
    
        print(
            erro
        )
    
        contexto_mead = ""
    
    contexto_mead = str(
        contexto_mead or ""
    )

    # ============================================================
    # MEAD/SEVEN — AUDITORIA, MAPA RELACIONAL E PLANO GLOBAL
    # ============================================================
    patrimonio_global_mead = _mead_preparar_patrimonio_global(
        tema, textos_relevantes, informacoes_blocos
    )
    plano_editorial_mead = _mead_planejar_pagina_global(
        tema, patrimonio_global_mead
    )
    contexto_global_pagina = _mead_formatar_contexto_global(
        tema, patrimonio_global_mead, plano_editorial_mead
    )
    print(
        "MEAD/SEVEN — PATRIMÔNIO GLOBAL:",
        len(patrimonio_global_mead),
        "| ÂNCORA LITERAL:",
        "SIM" if plano_editorial_mead.get("keyword_anchor_disponivel") else "NÃO"
    )
    print(
        "MEAD/SEVEN — ALERTAS DESVER:",
        len(plano_editorial_mead.get("sobreposicoes_desver", []))
    )
    contexto_mead_ollama_base = contexto_mead
    
    # ========================================================
    # 10. REGRA DE INDEPENDÊNCIA
    # ========================================================
    
    regra_independencia = f"""
    
    REGRA FUNDAMENTAL DESTA PÁGINA:
    
    Esta é uma página independente.
    
    TEMA:
    {tema}
    
    ARQUIVO:
    {nome_arquivo}
    
    Nunca copie ou adapte conteúdo de outra página.
    
    Use somente:
    
    o tema atual;
    as informações técnicas selecionadas;
    o mapa estratégico atual;
    as regras editoriais.
    
    O conhecimento técnico pode ser semelhante entre páginas.
    
    A REDAÇÃO NÃO PODE SER.
    
    Crie uma narrativa nova, própria e independente.
    """
    
    # ============================================================
    # IDENTIDADE FIXA + TÍTULOS EDITORIAIS POSTERIORES
    # ============================================================
    # O H1 e o title SEO identificam o tema e permanecem estáveis.
    # O subtítulo editorial e os cinco títulos de bloco serão lidos
    # da resposta global do Ollama, DEPOIS da redação dos parágrafos.
    # Não antecipar títulos genéricos que condicionem a narrativa.
    # ============================================================
    tema_titulo_exibicao = " ".join(
        palavra[:1].upper() + palavra[1:]
        for palavra in re.sub(r"\s+", " ", tema_base).split()
    )
    titulos_python = {
        "h1": tema_titulo_exibicao,
        "title": tema_titulo_exibicao,
        "subtitulo": "",
        "subtitulo_listas": "",
        "subtitulo_segmentos": "",
    }
    h1 = tema_titulo_exibicao
    titulo = tema_titulo_exibicao
    subtitulo = ""
    subtitulo_listas = ""
    subtitulo_segmentos = gerar_subtitulo_segmentos(tema_titulo_exibicao)
    print("TÍTULOS EDITORIAIS: serão definidos após os 15 parágrafos pelo Ollama.")
    
    
    # ============================================================
    # ASSOCIAR OS TÍTULOS AOS BLOCOS
    # ============================================================
    
    for numero_bloco in range(1, 6):
    
        chave_bloco = (
            f"bloco_{numero_bloco}"
        )
    
        dados_bloco = (
            informacoes_blocos.get(
                chave_bloco,
                {}
            )
        )
    
        if not isinstance(
            dados_bloco,
            dict
        ):
            dados_bloco = {}
    
        dados_bloco["titulo"] = (
            titulos_python.get(
                chave_bloco,
                ""
            )
        )
    
        informacoes_blocos[
            chave_bloco
        ] = dados_bloco
    

    # ============================================================
    # CONTROLE FINAL DO PROMPT
    # ============================================================
    #
    # IMPORTANTE:
    #
    # A arquitetura antiga utilizava:
    #
    #     contexto_fragmentos
    #     prompt
    #
    # Essas variáveis não pertencem mais à arquitetura atual.
    #
    # Agora cada bloco monta seu próprio:
    #
    #     contexto_tres_fragmentos
    #     prompt_bloco
    #
    # Portanto, neste ponto apenas registramos que a preparação
    # geral foi concluída.
    #
    # O tamanho real dos fragmentos e do prompt será informado
    # dentro do processamento individual de cada bloco.
    #
    # ============================================================
    
    print()
    print("=" * 60)
    print("CONTROLE DA ARQUITETURA OLLAMA")
    print("=" * 60)
    
    print(
        "TEMA:",
        tema
    )
    
    print(
        "ARQUITETURA:",
        "1 chamada para 15 parágrafos"
    )
    
    print(
        "BLOCOS:",
        total_blocos
    )
    
    print(
        "FRAGMENTOS POR BLOCO:",
        3
    )
    
    print(
        "CHAMADAS OLLAMA PREVISTAS:",
        1
    )
    
    print(
        "FONTES BRUTAS ENVIADAS:",
        "NÃO"
    )
    
    print(
        "MAPA MEAD ENVIADO AO OLLAMA:",
        "NÃO"
    )
    
    print(
        "MEAD EDITORIAL ENVIADO AO OLLAMA:",
        "SIM"
    )

    print(
        "CHECKBOXES DOS EDITORES ENVIADOS AO OLLAMA:",
        "NÃO"
    )
    
    print(
        "IDENTIDADE DAS FONTES ENVIADA:",
        "NÃO"
    )
    
    print(
        "CONTEXTO GLOBAL ANTIGO ENVIADO:",
        "NÃO"
    )
    
    print(
        "PROMPT GLOBAL ANTIGO:",
        "DESATIVADO"
    )
    
    print(
        "CONTEXTO POR BLOCO:",
        "SERÁ MONTADO A PARTIR DOS 3 FRAGMENTOS"
    )
    
    print("=" * 60)
    
    

    # ============================================================
    # 11. PROCESSAMENTO DOS 3 FRAGMENTOS POR BLOCO — OLLAMA
    # ============================================================
    #
    # NOVA ARQUITETURA:
    #
    # A página inteira é enviada em uma única chamada ao Ollama.
    # A resposta é separada em cinco blocos e validada em 15 posições.
    #
    # Mantemos:
    #
    # - o patrimônio dos 15 fragmentos;
    # - liberdade editorial para reorganizar e relacionar assuntos;
    # - os mesmos 5 blocos estruturais;
    # - 3 parágrafos por bloco;
    # - a mesma estrutura do JSON.
    #
    # ============================================================

    print()
    print("=" * 60)
    print("PROCESSAMENTO DA PÁGINA COMPLETA — UMA CHAMADA OLLAMA")
    print("=" * 60)

    print("CHAMADAS OLLAMA PREVISTAS: 1 PARA A PÁGINA INTEIRA")
    print("FRAGMENTOS ENVIADOS NA ÚNICA CHAMADA: 15 (3 POR BLOCO)")

    print(
        "PARÁGRAFOS ESPERADOS:",
        total_blocos * 3
    )

    # ============================================================
    # GARANTIR ESTRUTURA DOS 5 BLOCOS
    # ============================================================

    if not isinstance(
        informacoes_blocos,
        dict
    ):

        print()
        print(
            "❌ ERRO: informacoes_blocos não é um dicionário."
        )

        return None


    # Reservar todos os hashes já escolhidos para os 15 lugares.
    # Uma reserva de recuperação nunca pode reutilizar hash ocupado.
    for _dados_hash_bloco in informacoes_blocos.values():
        if not isinstance(_dados_hash_bloco, dict):
            continue
        for _frag_hash in _dados_hash_bloco.get("informacoes_relevantes", []) or []:
            if isinstance(_frag_hash, dict):
                _hash_existente = str(_frag_hash.get("hash", "") or "").strip()
                if _hash_existente:
                    hashes_reservas_ocupadas.add(_hash_existente)

    # ============================================================
    # PRÉ-VALIDAÇÃO ABSOLUTA DOS 5 BLOCOS
    # ============================================================
    #
    # O Ollama só pode começar se todos os blocos tiverem
    # exatamente 3 fragmentos. Assim o erro não aparece
    # no meio das chamadas.
    # ============================================================

    for numero_bloco in range(1, total_blocos + 1):

        chave_preflight = f"bloco_{numero_bloco}"
        bloco_preflight = informacoes_blocos.get(
            chave_preflight,
            {}
        )

        fragmentos_preflight = (
            bloco_preflight.get(
                "informacoes_relevantes",
                []
            )
            if isinstance(bloco_preflight, dict)
            else []
        )

        if not isinstance(fragmentos_preflight, list):
            fragmentos_preflight = []
            bloco_preflight["informacoes_relevantes"] = fragmentos_preflight
        while len(fragmentos_preflight) < 3:
            pos = len(fragmentos_preflight) + 1
            fragmentos_preflight.append({"id": f"CONHECIMENTO_GERAL_{chave_preflight}_{pos}", "hash": "", "palavra_chave_status": "SEM", "texto": "[SEM FRAGMENTO — usar conhecimento geral consolidado, sem inventar dados específicos.]", "fonte": "CONHECIMENTO_GERAL_DO_MODELO", "url": "", "tipo": "CONTEXTO_GERAL", "pdf": False, "palavras": 0, "identidade_fonte": {}, "contexto_geral_sem_fonte": True})

    # ============================================================
    # UMA ÚNICA CHAMADA OLLAMA PARA A PÁGINA INTEIRA (15 PARÁGRAFOS)
    # ============================================================
    # O Python mantém os 15 fragmentos e seus vínculos individuais.
    # O Ollama recebe o patrimônio completo e devolve cinco blocos.
    # As validações existentes continuam sendo executadas por bloco.
    # Recuperações que exigiriam nova chamada são desativadas nesta
    # arquitetura: se a resposta não passar, a execução informa o motivo.
    # ============================================================
    inicio_ollama_pagina = time.time()
    partes_patrimonio_pagina = []
    hashes_pagina = []
    for _numero in range(1, total_blocos + 1):
        _chave = f"bloco_{_numero}"
        _dados = informacoes_blocos.get(_chave, {})
        _frags = _dados.get("informacoes_relevantes", []) if isinstance(_dados, dict) else []
        partes_patrimonio_pagina.append(
            f"\n=== {_chave.upper()} | AGRUPAMENTO TÉCNICO; TÍTULO DEFINIDO APÓS A REDAÇÃO ==="
        )
        for _pos, _frag in enumerate(_frags[:3], 1):
            if not isinstance(_frag, dict):
                continue
            _texto = str(_frag.get("texto", "") or "").strip()
            _hash = str(_frag.get("hash", "") or "").strip()
            if _hash:
                hashes_pagina.append(_hash)
            # O Ollama recebe apenas a posição editorial e o texto autorizado.
            # IDs, hashes e status internos permanecem no Python/log/banco.
            partes_patrimonio_pagina.append(
                f"[EVIDENCIA {_chave} P{_pos}]\n{_texto}"
            )

    hashes_unicos_pagina = [h for h in hashes_pagina if h]
    if len(hashes_unicos_pagina) != len(set(hashes_unicos_pagina)):
        print("❌ HÁ HASHES REPETIDAS ENTRE OS 15 FRAGMENTOS; geração interrompida.")
        _salvar_diagnostico_ultima_execucao(tema, {"hashes": hashes_unicos_pagina}, "HASH_DUPLICADA_ENTRE_PARAGRAFOS")
        return None

    prompt_pagina_unica = f"""
Você é o redator técnico-editorial MEAD/SEVEN. Escreva UMA página completa em português brasileiro.

TEMA E EXPRESSÃO-CHAVE EXATA: {palavra_chave_redacao}
ESTRUTURA OBRIGATÓRIA: 5 blocos, 3 parágrafos em cada bloco, total de 15 parágrafos.

CONTEXTO EDITORIAL COMPACTO:
Use exclusivamente as 15 evidências específicas apresentadas mais abaixo como base factual.
Não há necessidade de repetir mapas, bancos, fontes brutas ou patrimônios paralelos.
Preserve a progressão dos cinco blocos e não complete lacunas com suposições.

DIRETRIZES DE FIDELIDADE E AUTORIA — ABORDAGEM FLEXÍVEL:
- Leia os 15 fragmentos como patrimônio informativo autorizado. Antes de redigir, planeje mentalmente a progressão dos cinco blocos e distribua os assuntos para que a página avance sem voltar repetidamente ao mesmo ponto.
- PROGRESSÃO EDITORIAL OBRIGATÓRIA, adaptada ao que as fontes realmente permitem afirmar:
  BLOCO 1 — PRIORIDADE COMERCIAL E INSTITUCIONAL: quando as evidências autorizarem, apresente a relação da nossa empresa com a palavra-chave exata, usando naturalmente expressões como “nossa empresa”, “nossos profissionais”, “nossos funcionários”, “nossa equipe”, “fornecemos” ou “atendimento”. Distribua esses aspectos entre os três parágrafos, varie a redação de página para página e relacione cada afirmação diretamente ao tema. Não invente identidade, equipe, oferta, disponibilidade, preço, orçamento, prazo ou condição comercial. Se não houver evidência comercial, use fatos autorizados sem criar uma oferta.
  BLOCO 2 — desenvolver características, componentes, configurações ou diferenças entre tipos, sem repetir a definição do bloco 1.
  BLOCO 3 — explicar critérios de escolha, dimensionamento, compatibilidade, condições de operação ou decisão técnica.
  BLOCO 4 — desenvolver aplicações e segmentos de uso concretos, explicando por que cada contexto demanda aquela solução. Não transformar nomes de empresas ou termos comerciais em aplicações.
  BLOCO 5 — abordar cuidados, manutenção, benefícios operacionais ou aspectos de contratação/fornecimento somente quando houver evidência. Fechar com orientação útil, não com propaganda genérica.
- Cada bloco deve ter um objetivo editorial próprio. Cada um dos três parágrafos deve cumprir função diferente: apresentar um aspecto específico, aprofundar outro com mecanismo/condição/consequência e acrescentar aplicação/critério/implicação prática. Não force essa sequência se as fontes não sustentarem algum ponto; reorganize sem inventar.
- Antes de escrever cada parágrafo, identifique mentalmente qual informação nova ele acrescentará. Se a ideia central já apareceu, não a reformule: escolha outro ângulo ou explique uma condição diferente sustentada pelas fontes.
- Evite repetir fórmulas como eficiência, resistência, segurança, confiabilidade, qualidade e economia sem explicar concretamente de onde vem o benefício, em quais condições ocorre e quais limites existem.
- Não use adjetivos promocionais vazios como “design inovador”, “alta eficiência”, “resistência comprovada”, “melhor escolha”, “solução ideal” ou “temperaturas extremas” sem evidência específica e contextualizada. Prefira descrever a característica verificável e sua consequência prática.
- FIDELIDADE ESTRITA: use somente informações que estejam expressamente presentes nas 15 evidências. Não use conhecimento geral para completar, corrigir ou ampliar o conteúdo. Se uma evidência for insuficiente, escreva menos; não preencha lacunas.
- NÃO MISTURE EQUIPAMENTOS: não trate versões sanitárias, submersas, autoaspirantes, bombas de deslocamento positivo ou outros tipos como se fossem o mesmo equipamento. Só compare ou relacione tipos diferentes quando a própria evidência trouxer essa comparação de forma explícita. Preserve o escopo e as condições de cada informação.
- NÃO INVENTE números, medidas, capacidades, materiais, marcas, modelos, fabricantes, normas, certificações, aplicações particulares, serviços, preços, garantias, disponibilidade, resultados ou promessas. Não invente dados comerciais nem fale em nome de uma empresa sem evidência explícita.
- IDIOMA OBRIGATÓRIO: escreva exclusivamente em português brasileiro. As evidências podem estar em inglês, espanhol ou outro idioma: interprete e traduza mentalmente os fatos relevantes antes de redigir. Não copie frases estrangeiras nem descarte um fato apenas porque a fonte está em outro idioma. Preserve nomes técnicos e unidades somente quando necessários e sustentados pela evidência.
- Não mencione fontes, fragmentos, hashes, Python, Ollama ou estas instruções.
- OBJETIVIDADE: cada parágrafo deve ter preferencialmente entre 40 e 55 palavras. Não existe mínimo obrigatório: se a evidência sustentar uma explicação completa em menos palavras, mantenha o parágrafo curto. Nunca aumente o texto apenas para atingir uma contagem.
- SEM ENROLAÇÃO: elimine chamadas genéricas, frases promocionais, introduções que apenas anunciam o assunto, conclusões óbvias e repetições da mesma ideia. Não use expressões como "confira as vantagens", "solução ideal", "alta qualidade", "excelente desempenho", "ampla gama" ou equivalentes sem informação factual específica que as comprove.
- DENSIDADE INFORMATIVA: cada frase deve acrescentar um fato autorizado, explicar uma relação presente na evidência ou esclarecer uma condição expressamente indicada. Se a frase não acrescentar informação, remova-a.
- PRECISÃO: prefira nomes técnicos e verbos concretos. Não generalize uma característica de um modelo para todos os equipamentos da categoria. Não transforme uma possibilidade em garantia ou regra universal.
- FLUIDEZ: varie naturalmente a construção das frases. Evite repetir a mesma abertura, estrutura sintática ou conclusão nos parágrafos próximos.
- AUTONOMIA: cada parágrafo deve ser compreensível por si, sem referentes vagos como "conforme mencionado", "nesse sentido" ou "essa característica" quando não estiver claro a que se referem.
- Em cada um dos 15 parágrafos, inclua a expressão exata "{palavra_chave_redacao}" uma única vez, com esta capitalização e acentuação. Não altere a expressão-chave para tentar fazê-la concordar com a frase. Construa a frase ao redor dela de forma gramaticalmente correta. Depois da única ocorrência, use pronomes ou referências naturais somente quando a concordância estiver correta; se houver risco de erro, reformule a frase.
- A expressão-chave não deve aparecer nos títulos, subtítulos ou títulos de bloco. Tags SEO são exceção e permanecem em minúsculas.
- Primeiro conclua os 15 parágrafos. SOMENTE DEPOIS, leia o conteúdo produzido e crie um título específico para cada bloco, refletindo os assuntos realmente tratados nos três parágrafos daquele bloco. Não use títulos genéricos ou uma sequência fixa de títulos.
- Depois de concluir os parágrafos, crie também um subtítulo editorial curto para a página, baseado no conjunto do conteúdo. Não repita o H1 nem a expressão-chave.
- Os títulos devem distinguir os cinco blocos e não podem prometer assuntos que não foram desenvolvidos.
- Não transforme os fragmentos em resumos individuais. Redija uma página nova, articulada e autoral, usando o patrimônio para diversificar e aprofundar os assuntos.

PATRIMÔNIO DE 15 POSIÇÕES (cada evidência está vinculada exclusivamente ao parágrafo indicado):
{chr(10).join(partes_patrimonio_pagina)}

FORMATO OBRIGATÓRIO — retorne somente os marcadores e o conteúdo:
[BLOCO_1]
[TITULO_BLOCO]Título específico[/TITULO_BLOCO]
[SUBTITULO_PAGINA]Subtítulo curto da página[/SUBTITULO_PAGINA]
[PARAGRAFO_1]Texto[/PARAGRAFO_1]
[PARAGRAFO_2]Texto[/PARAGRAFO_2]
[PARAGRAFO_3]Texto[/PARAGRAFO_3]
[/BLOCO_1]
Repita o mesmo formato para [BLOCO_2] até [BLOCO_5], sem o subtítulo de página nos demais blocos.
""".strip()

    print("OLLAMA — GERAÇÃO DA PÁGINA INTEIRA")
    print("CHAMADAS PREVISTAS: 1 | BLOCOS: 5 | PARÁGRAFOS: 15 | FRAGMENTOS: 15")
    print("AUTORIA OLLAMA: PROGRESSÃO EDITORIAL | 5 FUNÇÕES DE BLOCO | 3 ÂNGULOS DISTINTOS POR BLOCO")
    _caracteres_fragmentos_pagina = sum(
        len(str(_frag.get("texto", "") or ""))
        for _dados in informacoes_blocos.values() if isinstance(_dados, dict)
        for _frag in (_dados.get("informacoes_relevantes", []) or [])
        if isinstance(_frag, dict)
    )
    _caracteres_mead_pagina = len(str(contexto_mead_ollama_base or ""))
    _caracteres_plano_pagina = len(str(contexto_global_pagina or ""))
    _tokens_estimados_prompt = int(len(prompt_pagina_unica) / 4)
    print("PROMPT GLOBAL:", len(prompt_pagina_unica), "caracteres")
    print(
        "CARGA ENVIADA AO OLLAMA —",
        "fragmentos:", _caracteres_fragmentos_pagina, "caracteres |",
        "MEAD:", _caracteres_mead_pagina, "caracteres |",
        "plano global:", _caracteres_plano_pagina, "caracteres |",
        "prompt total:", len(prompt_pagina_unica), "caracteres |",
        "tokens estimados:", _tokens_estimados_prompt,
        "(estimativa aproximada; tokens efetivos serão lidos da resposta Ollama)"
    )
    try:
        resposta_pagina = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": MODELO_OLLAMA,
                "prompt": prompt_pagina_unica,
                "stream": True,
                "think": False,
                "keep_alive": "10m",
                "options": {
                    # Saída dimensionada para 15 parágrafos objetivos (40–65 palavras),
                    # cinco títulos e marcadores, sem incentivar preenchimento artificial.
                    "num_predict": 3800,
                    "num_ctx": 16384,
                    "temperature": 0.2,
                    "top_p": 0.9,
                    "repeat_penalty": 1.05
                }
            },
            timeout=(30, 1800),
            stream=True
        )
        resposta_pagina.raise_for_status()
        partes_resposta_stream = []
        dados_resposta_pagina = {}
        tokens_stream_estimados = 0
        ultimo_percentual_log = -1
        print("PROGRESSO OLLAMA: 0% — processando prompt e iniciando geração (estimativa por tokens).")
        for linha_stream in resposta_pagina.iter_lines(decode_unicode=True):
            if not linha_stream:
                continue
            try:
                evento_stream = json.loads(linha_stream)
            except (TypeError, ValueError):
                continue
            trecho_stream = str(evento_stream.get("response", "") or "")
            if trecho_stream:
                partes_resposta_stream.append(trecho_stream)
                # Aproximação deliberadamente identificada: ~4 caracteres/token.
                tokens_stream_estimados += max(1, int(len(trecho_stream) / 4))
                percentual_stream = min(99, int(tokens_stream_estimados / 4500 * 100))
                if percentual_stream >= ultimo_percentual_log + 5:
                    decorrido_stream = max(0.1, time.time() - inicio_ollama_pagina)
                    taxa_stream = round(tokens_stream_estimados / decorrido_stream, 2)
                    print(
                        f"PROGRESSO OLLAMA: ~{percentual_stream}% (estimativa de saída) "
                        f"| tokens aproximados (caracteres/4): {tokens_stream_estimados} "
                        f"| velocidade aproximada: {taxa_stream} tokens/s"
                    )
                    ultimo_percentual_log = percentual_stream
            if evento_stream.get("done"):
                dados_resposta_pagina = evento_stream
                break
        resposta_global_ollama = "".join(partes_resposta_stream).strip()
        if not dados_resposta_pagina:
            raise RuntimeError("Ollama encerrou o fluxo sem evento final de métricas.")
        print("PROGRESSO OLLAMA: resposta recebida — geração encerrada; percentual do fluxo era estimado.")
        def _ollama_ns_para_segundos(valor):
            try:
                return round(float(valor) / 1_000_000_000, 2)
            except (TypeError, ValueError):
                return None

        _prompt_eval_s = _ollama_ns_para_segundos(dados_resposta_pagina.get("prompt_eval_duration"))
        _eval_s = _ollama_ns_para_segundos(dados_resposta_pagina.get("eval_duration"))
        _total_s = _ollama_ns_para_segundos(dados_resposta_pagina.get("total_duration"))
        _prompt_tokens = int(dados_resposta_pagina.get("prompt_eval_count", 0) or 0)
        _output_tokens = int(dados_resposta_pagina.get("eval_count", 0) or 0)
        _prompt_rate = round(_prompt_tokens / _prompt_eval_s, 2) if _prompt_eval_s and _prompt_eval_s > 0 else None
        _output_rate = round(_output_tokens / _eval_s, 2) if _eval_s and _eval_s > 0 else None
        print(
            "COMPARAÇÃO DE TOKENS — estimados:", _tokens_estimados_prompt,
            "| entrada efetiva Ollama:", _prompt_tokens,
            "| diferença:", (_prompt_tokens - _tokens_estimados_prompt)
        )
        # Preservar a resposta bruta de cada execução, mesmo quando o Ollama
        # termina normalmente. A leitura pode falhar por variação de formato;
        # o texto original fica recuperável no progresso.json, sem JSON auxiliar.
        try:
            _progresso_resposta = _ler_progresso_seguro()
            _respostas_globais = _progresso_resposta.get("respostas_globais_ollama", {})
            if not isinstance(_respostas_globais, dict):
                _respostas_globais = {}
            _respostas_globais[_tema_recuperacao_id(tema)] = {
                "tema": str(tema),
                "resposta": resposta_global_ollama,
                "done_reason": str(dados_resposta_pagina.get("done_reason", "")),
                "prompt_tokens": _prompt_tokens,
                "output_tokens": _output_tokens,
                "salvo_em": time.strftime("%d/%m/%Y %H:%M:%S")
            }
            _progresso_resposta["respostas_globais_ollama"] = _respostas_globais
            _gravar_progresso_seguro(_progresso_resposta)
            print("💾 RESPOSTA GLOBAL PRESERVADA NO progresso.json — antes da leitura editorial")
        except Exception as _erro_preservacao_global:
            print("⚠️ NÃO FOI POSSÍVEL PRESERVAR A RESPOSTA GLOBAL:", repr(_erro_preservacao_global))

        # Se o modelo atingir o limite, também marcar a execução como parcial.
        if str(dados_resposta_pagina.get("done_reason", "")).lower() == "length":
            try:
                _progresso_parcial = _ler_progresso_seguro()
                _pendencias_parciais = _progresso_parcial.get("respostas_parciais_ollama", {})
                if not isinstance(_pendencias_parciais, dict):
                    _pendencias_parciais = {}
                _pendencias_parciais[_tema_recuperacao_id(tema)] = {
                    "tema": str(tema),
                    "resposta": resposta_global_ollama,
                    "done_reason": "length",
                    "prompt_tokens": _prompt_tokens,
                    "output_tokens": _output_tokens,
                    "salvo_em": time.strftime("%d/%m/%Y %H:%M:%S")
                }
                _progresso_parcial["respostas_parciais_ollama"] = _pendencias_parciais
                _gravar_progresso_seguro(_progresso_parcial)
                print("💾 RESPOSTA PARCIAL PRESERVADA NO progresso.json — motivo: limite de saída")
            except Exception as _erro_preservacao_resposta:
                print("⚠️ NÃO FOI POSSÍVEL PRESERVAR A RESPOSTA PARCIAL:", repr(_erro_preservacao_resposta))

        print(
            "MÉTRICAS OLLAMA:",
            "done_reason=", dados_resposta_pagina.get("done_reason", "não informado"),
            "| prompt_tokens=", _prompt_tokens,
            "| output_tokens=", _output_tokens,
            "| prompt_eval_s=", _prompt_eval_s,
            "| eval_s=", _eval_s,
            "| total_s=", _total_s,
            "| prompt_tok_s=", _prompt_rate,
            "| output_tok_s=", _output_rate
        )
    except requests.exceptions.Timeout:
        print("❌ TIMEOUT OLLAMA — geração única da página.")
        _salvar_diagnostico_ultima_execucao(tema, {"prompt_caracteres": len(prompt_pagina_unica)}, "TIMEOUT_OLLAMA_PAGINA")
        return None
    except Exception as erro:
        print("❌ ERRO NA CHAMADA ÚNICA DO OLLAMA:", repr(erro))
        _salvar_diagnostico_ultima_execucao(tema, {"prompt_caracteres": len(prompt_pagina_unica)}, "ERRO_OLLAMA_PAGINA", repr(erro))
        return None

    resposta_global_ollama = re.sub(r"```(?:text|txt)?|```", "", resposta_global_ollama, flags=re.I).strip()
    if not resposta_global_ollama:
        print("❌ OLLAMA RETORNOU RESPOSTA VAZIA PARA A PÁGINA.")
        return None
    print("RESPOSTA GLOBAL:", len(resposta_global_ollama), "caracteres | TEMPO:", round(time.time() - inicio_ollama_pagina, 2), "s")

    # ============================================================
    # PREPARAR LEITURA ALTERNATIVA QUANDO O MODELO OMITIR MARCADORES
    # ============================================================
    _paragrafos_fallback_global = []
    _titulo_fallback_global = ""
    _subtitulo_fallback_global = ""
    _aberturas_blocos_presentes = [
        bool(re.search(rf"\[BLOCO_{i}\]", resposta_global_ollama, flags=re.I))
        for i in range(1, 6)
    ]
    _marcadores_paragrafos_presentes = len(re.findall(
        r"\[PARAGRAFO_[1-3]\]", resposta_global_ollama, flags=re.I
    )) >= 15

    if not all(_aberturas_blocos_presentes) or not _marcadores_paragrafos_presentes:
        _mt = re.search(r"\[TITULO_BLOCO\](.*?)\[/TITULO_BLOCO\]", resposta_global_ollama, flags=re.I | re.S)
        _ms = re.search(r"\[SUBTITULO_PAGINA\](.*?)\[/SUBTITULO_PAGINA\]", resposta_global_ollama, flags=re.I | re.S)
        if _mt:
            _titulo_fallback_global = re.sub(r"\s+", " ", _mt.group(1)).strip()
        if _ms:
            _subtitulo_fallback_global = re.sub(r"\s+", " ", _ms.group(1)).strip()

        if all(_aberturas_blocos_presentes):
            # Há marcadores de bloco, mas faltam os marcadores de parágrafo:
            # recuperar somente quando cada bloco contém três parágrafos claros.
            _paragrafos_fallback_global = _extrair_paragrafos_por_bloco_resposta_global(
                resposta_global_ollama
            )
        elif not any(_aberturas_blocos_presentes):
            # Compatibilidade com respostas totalmente sem marcadores de bloco.
            _paragrafos_fallback_global = _extrair_paragrafos_resposta_global(
                resposta_global_ollama, limite=15
            )

        print("⚠️ ESTRUTURA INCOMPLETA — leitura alternativa segura ativada.")
        print("PARÁGRAFOS RECUPERÁVEIS:", len(_paragrafos_fallback_global), "/ 15")

    # ============================================================
    # PROCESSAR CADA BLOCO
    # ============================================================

    for numero_bloco in range(
        1,
        total_blocos + 1
    ):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        dados_bloco = (
            informacoes_blocos.get(
                chave_bloco,
                {}
            )
        )

        bloco_recuperado = blocos_recuperados.get(chave_bloco)
        if isinstance(bloco_recuperado, dict):
            paras_recuperados = bloco_recuperado.get("paragrafos_ollama", [])
            frags_recuperados = bloco_recuperado.get("informacoes_relevantes", [])
            keyword_recuperacao_ok = (
                isinstance(paras_recuperados, list)
                and len(paras_recuperados) == 3
                and all(
                    str(x or "").strip()
                    and _mead_contar_keyword_exata(str(x), palavra_chave_redacao) == 1
                    for x in paras_recuperados
                )
            )
            if (keyword_recuperacao_ok
                    and isinstance(frags_recuperados, list) and len(frags_recuperados) == 3):
                informacoes_blocos[chave_bloco] = bloco_recuperado
                print(f"♻️ {chave_bloco.upper()} RECUPERADO — 3 parágrafos validados; Ollama não será chamado.")
                continue
            elif bloco_recuperado:
                print(f"⚠️ {chave_bloco.upper()} EXISTE NA RECUPERAÇÃO, MAS NÃO PASSOU NA CONFERÊNCIA; SERÁ REGERADO.")

        if not isinstance(
            dados_bloco,
            dict
        ):

            print()
            print(
                f"❌ ERRO: {chave_bloco} "
                "não possui estrutura válida."
            )

            return None


        # ========================================================
        # INFORMAÇÕES RELEVANTES
        # ========================================================

        informacoes_relevantes = (
            dados_bloco.get(
                "informacoes_relevantes",
                []
            )
        )

        if not isinstance(
            informacoes_relevantes,
            list
        ):

            print()
            print(
                f"❌ ERRO: {chave_bloco} possui "
                "informacoes_relevantes inválidas."
            )

            return None


        # ========================================================
        # CADA BLOCO PRECISA POSSUIR EXATAMENTE 3 FRAGMENTOS
        # ========================================================

        if len(
            informacoes_relevantes
        ) != 3:

            print()
            print("=" * 60)
            print(
                f"❌ ERRO NO {chave_bloco}"
            )
            print("=" * 60)

            print(
                "FRAGMENTOS ENCONTRADOS:",
                len(
                    informacoes_relevantes
                )
            )

            print(
                "FRAGMENTOS NECESSÁRIOS:",
                3
            )

            return None


        # ========================================================
        # GARANTIR LISTA DE PARÁGRAFOS PYTHON
        # ========================================================

        paragrafos_python = (
            dados_bloco.get(
                "paragrafos_python",
                []
            )
        )

        if not isinstance(
            paragrafos_python,
            list
        ):

            paragrafos_python = []


        paragrafos_python = (
            paragrafos_python
            + [
                "",
                "",
                ""
            ]
        )[:3]


        # ========================================================
        # GARANTIR LISTA DE PARÁGRAFOS OLLAMA
        # ========================================================

        paragrafos_ollama = (
            dados_bloco.get(
                "paragrafos_ollama",
                []
            )
        )

        if not isinstance(
            paragrafos_ollama,
            list
        ):

            paragrafos_ollama = []


        paragrafos_ollama = (
            paragrafos_ollama
            + [
                "",
                "",
                ""
            ]
        )[:3]



        # ========================================================
        # TÍTULO DO BLOCO — PRESERVAR O TÍTULO EXTRAÍDO
        # ========================================================
        # O marcador [TITULO_BLOCO] já foi extraído da resposta global
        # e armazenado em dados_bloco/titulos_python. Não reinicializar
        # a variável com string vazia: isso apagava o título antes de
        # montar a estrutura final e salvar o JSON.
        titulo_bloco = str(
            dados_bloco.get("titulo", "")
            or titulos_python.get(chave_bloco, "")
            or ""
        ).strip()

        # A extração defensiva do título a partir de resultado_ollama
        # ocorre somente depois que a resposta do bloco é preparada.
        # Neste ponto, a variável ainda pode não existir.

        if titulo_bloco:
            dados_bloco["titulo"] = titulo_bloco
            titulos_python[chave_bloco] = titulo_bloco


        # ========================================================
        # HASH ORIGINAL
        # ========================================================

        hash_bloco = str(
            dados_bloco.get(
                "hash",
                ""
            )
            or ""
        ).strip()


        # ========================================================
        # VALIDAR E PREPARAR OS 3 FRAGMENTOS
        # ========================================================

        fragmentos_bloco = []

        for indice_fragmento in range(
            3
        ):

            fragmento = (
                informacoes_relevantes[
                    indice_fragmento
                ]
            )

            # ----------------------------------------------------
            # VALIDAR OBJETO
            # ----------------------------------------------------

            if not isinstance(
                fragmento,
                dict
            ):

                print()
                print(
                    f"❌ ERRO: fragmento "
                    f"{indice_fragmento + 1} do "
                    f"{chave_bloco} não é um objeto."
                )

                return None


            # ----------------------------------------------------
            # TEXTO ORIGINAL
            # ----------------------------------------------------

            texto_fragmento = str(
                fragmento.get(
                    "texto",
                    ""
                )
                or ""
            ).strip()


            if not texto_fragmento:

                print()
                print(
                    f"❌ ERRO: fragmento "
                    f"{indice_fragmento + 1} do "
                    f"{chave_bloco} está vazio."
                )

                return None


            numero_fragmento = (
                indice_fragmento + 1
            )

            id_fragmento = str(
                fragmento.get(
                    "id",
                    ""
                )
                or ""
            ).strip()


            fragmentos_bloco.append({
                "numero": numero_fragmento,
                "id": id_fragmento,
                "hash": str(fragmento.get("hash", "") or ""),
                "palavra_chave_status": str(fragmento.get("palavra_chave_status", "SEM") or "SEM"),
                "texto": texto_fragmento,
                "fonte": fragmento.get("fonte", ""),
                "url": fragmento.get("url", ""),
                "tipo": fragmento.get("tipo", ""),
                "pdf": fragmento.get("pdf", False),
                "palavras": fragmento.get("palavras", len(texto_fragmento.split())),
                "identidade_fonte": fragmento.get("identidade_fonte", {}),
                "contexto_geral_sem_fonte": bool(fragmento.get("contexto_geral_sem_fonte", False))
            })


        # ========================================================
        # LOG DOS 3 FRAGMENTOS
        # ========================================================

        print()
        print("=" * 60)
        print(
            f"OLLAMA — {chave_bloco.upper()} — 3 FRAGMENTOS"
        )
        print("=" * 60)

        print("TRECHOS:", len(fragmentos_bloco), "| palavras:",
              ", ".join(str(len(x.get("texto", "").split())) for x in fragmentos_bloco))


        # ========================================================
        # NOVA ARQUITETURA — OLLAMA RECEBE 3 TRECHOS + MEAD EDITORIAL
        # ========================================================
        #
        # CORREÇÃO v7.0:
        #
        # O MEAD já é preparado anteriormente por preparar_mead(),
        # porém a arquitetura anterior não o inseria no prompt
        # efetivamente enviado ao Ollama.
        #
        # Agora cada chamada do Ollama recebe:
        # - as regras MEAD preparadas para IA;
        # - os checkboxes editoriais ativos desta página;
        # - o bloco atual;
        # - os 3 fragmentos autorizados pelo Python.
        #
        # O MEAD NÃO autoriza o Ollama a pesquisar ou inventar
        # informações. Ele funciona como orientação editorial.
        # Os fragmentos são a base pesquisada; o modelo pode complementar
        # com conhecimento geral consolidado, sem inventar dados específicos.
        #
        # ========================================================
        #
        # Python já resolveu:
        #
        # - pesquisa;
        # - coleta;
        # - limpeza;
        # - identificação;
        # - filtros;
        # - relevância;
        # - seleção;
        # - distribuição;
        # - até 3 trechos por bloco, complementados por conhecimento geral.
        #
        # O Ollama NÃO recebe:
        #
        # - fontes brutas;
        # - mapa MEAD;
        # - MEAD;
        # - identidade da fonte;
        # - domínio;
        # - empresa;
        # - outras fontes;
        # - outros blocos;
        # - entendimento da pesquisa.
        #
        # Ele recebe os trechos disponíveis e a autorização para conhecimento geral.
        #
        # ========================================================
        
        contexto_tres_fragmentos = ""

        for fragmento in fragmentos_bloco:

            contexto_tres_fragmentos += f"""
--------------------------------------------------
TRECHO {fragmento["numero"]}
--------------------------------------------------

{fragmento["texto"]}
"""

        # ========================================================
        # MEAD EDITORIAL DA PÁGINA
        # ========================================================
        #
        # `contexto_mead` foi preparado anteriormente por
        # `preparar_mead(MEAD)`. A v6.7 preparava esse contexto,
        # mas não o colocava na chamada efetiva do Ollama.
        #
        # Além do MEAD geral, enviamos os checkboxes ativos para
        # deixar explícito o que foi escolhido pelos editores.
        # ========================================================

        # Checkboxes editoriais não fazem parte do contexto do Ollama.
        # A estrutura editorial permanece no Python para a lógica da interface,
        # mas não é serializada nem enviada ao modelo.

        contexto_mead_ollama = str(
            contexto_mead or ""
        ).strip()

        if not contexto_mead_ollama:
            contexto_mead_ollama = "MEAD não disponível para esta execução."

        # ========================================================
        # PROMPT FINAL DO BLOCO
        # ========================================================

        prompt_bloco = f"""

Você é um redator técnico.

Sua única função é REEDITAR os três trechos fornecidos
pelo Python.

Não faça pesquisa.

Não selecione informações.

Não procure outras fontes.

Não utilize conhecimento externo.

Não invente informações.

Não complete lacunas com conhecimento próprio.

Não crie fatos.

Não acrescente números.

Não acrescente características.

Não acrescente aplicações.

Não acrescente materiais.

Não acrescente normas.

Não acrescente certificações.

Não acrescente marcas.

Não acrescente empresas.

Não acrescente modelos.

Não acrescente clientes.

Não acrescente resultados.

Não acrescente informações que não estejam nos trechos.

O MEAD abaixo é uma orientação editorial da página.
Ele NÃO é uma fonte factual adicional.
Use-o para respeitar a intenção editorial, o foco e as regras
da página, mas obtenha todo conteúdo técnico factual somente
dos três trechos autorizados pelo Python.

==================================================
MEAD EDITORIAL DA PÁGINA
==================================================

{contexto_mead_ollama}

==================================================
BLOCO ATUAL
==================================================

{chave_bloco}

==================================================
REGRA DE REEDIÇÃO
==================================================

TRECHO 1 → PARÁGRAFO 1

TRECHO 2 → PARÁGRAFO 2

TRECHO 3 → PARÁGRAFO 3

Cada trecho deve gerar somente o seu próprio parágrafo.

Não misture os trechos.

Não utilize informação de outro trecho.

Não utilize informação de outro bloco.

Preserve o sentido técnico original.

Você pode corrigir:

- gramática;
- concordância;
- pontuação;
- ordem das frases;
- fluidez;
- repetição desnecessária;
- construção textual.

Você NÃO pode alterar o conteúdo factual.

O resultado deve ser natural, técnico e claro.

Não faça listas.

Não crie títulos.

Não crie subtítulos.

Não crie conclusão.

Não crie introdução fora dos três parágrafos.

Procure aproximadamente 60 a 70 palavras por parágrafo,
mas NUNCA invente conteúdo apenas para atingir tamanho.

Se o trecho possuir menos informação, produza um
parágrafo menor.

==================================================
FORMATO OBRIGATÓRIO
==================================================

[BLOCO]

[PARAGRAFO_1]
Reedição do TRECHO 1.
[/PARAGRAFO_1]

[PARAGRAFO_2]
Reedição do TRECHO 2.
[/PARAGRAFO_2]

[PARAGRAFO_3]
Reedição do TRECHO 3.
[/PARAGRAFO_3]

[/BLOCO]

==================================================
TRECHOS AUTORIZADOS PELO PYTHON
==================================================

{contexto_tres_fragmentos}

RETORNE SOMENTE O BLOCO.
"""


        # ========================================================
        # CONTROLE DO PROMPT DO BLOCO
        # ========================================================

        print()
        print(
            f"PROMPT {chave_bloco.upper()}:",
            len(prompt_bloco),
            "caracteres"
        )

        print(
            "TRECHOS ENVIADOS:",
            len(fragmentos_bloco)
        )

        print(
            "IDENTIDADE DA FONTE ENVIADA:",
            "NÃO"
        )

        print(
            "FONTES EXTERNAS ENVIADAS:",
            "NÃO"
        )

        print(
            "MAPA MEAD ENVIADO:",
            "NÃO"
        )

        print(
            "MEAD EDITORIAL ENVIADO:",
            "SIM"
            if contexto_mead_ollama
            else "NÃO"
        )

        print(
            "CHECKBOXES DOS EDITORES ENVIADOS:",
            "NÃO"
        )
        
        
        # ========================================================
        # UMA CHAMADA OLLAMA POR BLOCO — 3 PARÁGRAFOS NARRATIVOS
        # ========================================================
        # Os três fragmentos são enviados juntos. Cada parágrafo
        # continua vinculado exclusivamente ao fragmento de mesma
        # numeração para a validação factual posterior. O contexto
        # conjunto serve para dar continuidade e transições naturais,
        # não para misturar fatos entre evidências.
        # ========================================================
        inicio_ollama = time.time()

        for indice_fragmento, fragmento_individual in enumerate(fragmentos_bloco, 1):
            texto_evidencia = str(fragmento_individual.get("texto", "") or "").strip()
            if not texto_evidencia:
                motivo = "EVIDENCIA_VAZIA"
                print(f"❌ {motivo} — {chave_bloco} fragmento {indice_fragmento}")
                _salvar_diagnostico_ultima_execucao(
                    tema,
                    {"bloco": chave_bloco, "fragmento": indice_fragmento,
                     "evidencia": fragmento_individual},
                    motivo,
                    "Fragmento sem texto."
                )
                return None

            if fragmento_individual.get("contexto_geral_sem_fonte"):
                continue
            ok_evidencia, motivo_evidencia = auditar_evidencia_para_redacao(texto_evidencia)
            if not ok_evidencia:
                print(f"❌ EVIDÊNCIA REPROVADA — {chave_bloco} fragmento {indice_fragmento}: {motivo_evidencia}")
                _salvar_diagnostico_ultima_execucao(
                    tema,
                    {"bloco": chave_bloco, "fragmento": indice_fragmento,
                     "evidencia": fragmento_individual,
                     "motivo_auditoria": motivo_evidencia},
                    "EVIDENCIA_REPROVADA_PRE_OLLAMA",
                    motivo_evidencia
                )
                return None

        prompt_bloco = f"""
Você é o EDITOR MEAD/SEVEN. Sua tarefa é CRIAR a parte desta página,
não apenas trocar palavras dos fragmentos.

TEMA E EXPRESSÃO-CHAVE EXATA: {palavra_chave_redacao}
BLOCO ATUAL: {chave_bloco}
TÍTULO ATUAL (apenas referência): {titulo_bloco or 'a definir pelo plano'}

LEITURA GLOBAL OBRIGATÓRIA
Leia todo o patrimônio selecionado e o plano editorial global abaixo antes
de escrever. A página deve ter visão de conjunto, continuidade e progressão.
Você pode relacionar informações de diferentes fragmentos do patrimônio,
inclusive de outros blocos, quando a relação estiver explicitamente sustentada.
Não repita a mesma informação em blocos diferentes. O plano indica focos
possíveis; não é uma taxonomia rígida nem uma ordem narrativa obrigatória.

PATRIMÔNIO GLOBAL E PLANO DA PÁGINA
{contexto_global_pagina}

ORIENTAÇÃO MEAD EDITORES V2
{contexto_mead_ollama_base}

REGRAS DE FIDELIDADE
- Use primeiro os fatos presentes no patrimônio global autorizado.
- Quando o material for limitado, complemente com conhecimento geral e consolidado
  sobre o tema, como conceitos, princípios, funções e explicações amplamente
  reconhecidas. Não afirme que esse conhecimento veio das fontes pesquisadas.
- Não invente nem deduza dados específicos de marcas, fabricantes, empresas,
  modelos, normas, certificações, materiais, aplicações particulares, serviços,
  resultados, garantias, capacidades, preços, disponibilidade ou promessas.
- Não generalize uma característica de um modelo ou aplicação para todo
  o produto ou setor.
- Preserve todos os complementos do tema. Decompor serve para compreender,
  nunca para apagar ou ampliar indevidamente o significado.
- Não mencione fontes, fragmentos, hashes, Python, IA ou estas instruções.

PLANEJAMENTO E REDAÇÃO
- Crie um título curto e verdadeiro para este bloco, derivado do conteúdo
  efetivamente sustentado. Evite modelos fixos e títulos genéricos.
- Produza exatamente três parágrafos que façam parte de uma única página.
- Cada parágrafo deve ter uma função e uma progressão próprias.
- Use perspectivas e aberturas variadas, sem rodízio mecânico.
- Os três parágrafos não precisam seguir a ordem dos fragmentos.
- Não junte fatos sem relação apenas para aumentar o tamanho.
- Não force personagem, cenário, problema, metáfora ou história.
- A voz é primeira pessoa do plural somente quando representar o prestador
  e houver base factual para isso.
- No bloco 1, qualquer afirmação comercial deve estar explicitamente ligada
  à expressão-chave exata da página e sustentada pelo patrimônio autorizado.
- No bloco 1, distribua os pontos comerciais com naturalidade entre os
  parágrafos e varie a redação de página para página; não use frase-modelo,
  abordagem comercial genérica nem CTA como "Entre em contato".
- Não afirme experiência, qualidade, equipe, condições de pagamento,
  abrangência/região de atendimento, oferta ou serviços sem comprovação.
- Nos demais blocos, não inclua CTA nem promoção.
- Escreva somente prosa corrida nos parágrafos: não use listas, tabelas,
  fichas técnicas, bullets, tópicos, subtítulos ou enumerações internas.
- Quando o tema envolver tipos distintos de equipamento, mantenha suas
  características e aplicações separadas. Para bombas, não confunda bomba
  centrífuga geral, sanitária, submersível ou autoaspirante; só mencione
  cada tipo quando estiver sustentado e pertinente ao tema exato.

PALAVRA-CHAVE — REGRA ABSOLUTA
- Insira a expressão exata "{palavra_chave_redacao}" exatamente UMA vez em CADA parágrafo.
- Preserve a grafia, acentuação, ordem e capitalização exatamente como acima.
- Não use variante, flexão, abreviação ou palavra isolada como substituta.
- A expressão não precisa abrir a primeira frase; deve entrar naturalmente.
- Não repita a expressão no título do bloco.

EXTENSÃO E QUALIDADE
- Prefira parágrafos completos e naturais, em geral entre 40 e 55 palavras.
- A faixa de 40–55 palavras é somente referência editorial: nunca é filtro,
  motivo de reprovação ou licença para inventar/esticar conteúdo.
- Se a evidência disponível exigir menos palavras, seja preciso e conciso.
- Português brasileiro, gramática correta, pontuação e transições naturais.
- Evite repetição lexical e conceitual entre os três parágrafos e com os demais blocos.

AUDITORIA ISI E DESVER ANTES DE RESPONDER
- Leia literalmente cada frase para evitar imagens sensoriais involuntárias.
- Observe o ambiente e a cena que o conjunto dos parágrafos produz.
- Confira se os três parágrafos acrescentam ângulos diferentes.
- Se um fato já foi tratado em outro ponto da página, escolha outro aspecto
  sustentado ou aprofunde uma relação diferente, sem repetir a explicação.

TRECHOS DE MAIOR PROXIMIDADE DO BLOCO ATUAL
{contexto_tres_fragmentos}

FORMATO OBRIGATÓRIO
[BLOCO]
[TITULO_BLOCO]
Título curto e específico, derivado das evidências.
[/TITULO_BLOCO]
{("[SUBTITULO_PAGINA]\nSubtítulo curto, sem repetir o H1, ancorado no conjunto de fatos.\n[/SUBTITULO_PAGINA]" if chave_bloco == "bloco_1" else "")}
[PARAGRAFO_1]
Primeiro parágrafo da página.
[/PARAGRAFO_1]
[PARAGRAFO_2]
Segundo parágrafo da página.
[/PARAGRAFO_2]
[PARAGRAFO_3]
Terceiro parágrafo da página.
[/PARAGRAFO_3]
[/BLOCO]

Retorne somente o bloco com os marcadores.
"""

        print()
        print("=" * 58)
        print(f"LEITURA DA RESPOSTA GLOBAL — {chave_bloco.upper()} — 3 FRAGMENTOS / 3 PARÁGRAFOS")
        print("FRAGMENTOS ENVIADOS:", len(fragmentos_bloco))
        print("PROMPT:", len(prompt_bloco), "caracteres")
        print("PALAVRA-CHAVE EXIGIDA UMA VEZ POR PARÁGRAFO:", palavra_chave_redacao)
        print("NARRATIVA ENTRE OS 3 PARÁGRAFOS: ATIVA")

        marcador_bloco_inicio = f"[BLOCO_{numero_bloco}]"
        marcador_bloco_fim = f"[/BLOCO_{numero_bloco}]"
        pos_inicio_bloco = resposta_global_ollama.lower().find(marcador_bloco_inicio.lower())
        pos_fim_bloco = resposta_global_ollama.lower().find(marcador_bloco_fim.lower(), pos_inicio_bloco + len(marcador_bloco_inicio)) if pos_inicio_bloco >= 0 else -1
        _conteudo_bloco_tem_tags_paragrafo = False
        if pos_inicio_bloco >= 0:
            _fim_teste_bloco = pos_fim_bloco if pos_fim_bloco >= 0 else len(resposta_global_ollama)
            _trecho_teste_bloco = resposta_global_ollama[
                pos_inicio_bloco + len(marcador_bloco_inicio):_fim_teste_bloco
            ]
            _conteudo_bloco_tem_tags_paragrafo = bool(re.search(
                r"\[PARAGRAFO_1\]", _trecho_teste_bloco, flags=re.I
            ))

        _precisa_fallback_bloco = (
            pos_inicio_bloco < 0
            or pos_fim_bloco < 0
            or not _conteudo_bloco_tem_tags_paragrafo
        )

        if _precisa_fallback_bloco:
            if not _paragrafos_fallback_global:
                print(
                    f"❌ ESTRUTURA INCOMPLETA EM {chave_bloco.upper()} "
                    "E NÃO HÁ TRÊS PARÁGRAFOS RECUPERÁVEIS POR BLOCO."
                )
                _salvar_diagnostico_ultima_execucao(
                    tema,
                    {"bloco": chave_bloco, "paragrafos_recuperaveis": len(_paragrafos_fallback_global)},
                    "BLOCO_AUSENTE_RESPOSTA_GLOBAL"
                )
                return None

            _inicio_fallback = (numero_bloco - 1) * 3
            _grupo_fallback = _paragrafos_fallback_global[_inicio_fallback:_inicio_fallback + 3]
            if len(_grupo_fallback) != 3:
                print(f"❌ RESPOSTA INCOMPLETA: {chave_bloco} não possui três parágrafos recuperáveis.")
                return None

            _partes_sinteticas = []
            if numero_bloco == 1 and _titulo_fallback_global:
                _partes_sinteticas.append(f"[TITULO_BLOCO]{_titulo_fallback_global}[/TITULO_BLOCO]")
            if numero_bloco == 1 and _subtitulo_fallback_global:
                _partes_sinteticas.append(f"[SUBTITULO_PAGINA]{_subtitulo_fallback_global}[/SUBTITULO_PAGINA]")
            for _idx_fallback, _texto_fallback_par in enumerate(_grupo_fallback, 1):
                _partes_sinteticas.append(f"[PARAGRAFO_{_idx_fallback}]{_texto_fallback_par}[/PARAGRAFO_{_idx_fallback}]")
            resultado_ollama = "\n".join(_partes_sinteticas)
            print(f"⚠️ {chave_bloco.upper()}: bloco reconstruído a partir de texto simples ({len(_grupo_fallback)} parágrafos).")
        else:
            resultado_ollama = resposta_global_ollama[
                pos_inicio_bloco + len(marcador_bloco_inicio):pos_fim_bloco
            ].strip()
        resultado_ollama = re.sub(r"^```(?:text|txt)?\s*|\s*```$", "", resultado_ollama, flags=re.I).strip()

        if not resultado_ollama:
            print(f"❌ OLLAMA RETORNOU VAZIO — {chave_bloco}")
            _salvar_diagnostico_ultima_execucao(
                tema, {"bloco": chave_bloco, "prompt_caracteres": len(prompt_bloco)},
                "OLLAMA_RETORNOU_VAZIO", "A resposta do bloco veio vazia."
            )
            return None

        fim_ollama = time.time()
        dados_ollama = {"response": resultado_ollama}
        print("TEMPO DA CHAMADA GLOBAL (total da página):", round(fim_ollama - inicio_ollama_pagina, 2), "segundos")
        print("CHAMADAS OLLAMA NESTE BLOCO: 0 (resposta global reutilizada)")
        print("CARACTERES RETORNADOS:", len(resultado_ollama))
        print("PARÁGRAFOS ESPERADOS NESTA RESPOSTA: 3")

        # EXTRAIR OS 3 PARÁGRAFOS
        # ========================================================

        resultado_ollama = str(
            resultado_ollama or ""
        ).strip()


        # --------------------------------------------------------
        # REMOVER CERCAS DE CÓDIGO, CASO EXISTAM
        # --------------------------------------------------------

        resultado_ollama = re.sub(
            r"```(?:text|txt)?",
            "",
            resultado_ollama,
            flags=re.IGNORECASE
        )

        resultado_ollama = re.sub(
            r"```",
            "",
            resultado_ollama
        )

        resultado_ollama = resultado_ollama.strip()


        # ========================================================
        # TÍTULO DO BLOCO GERADO A PARTIR DO PLANO E DAS EVIDÊNCIAS
        # ========================================================
        match_titulo_bloco = re.search(
            r"\[TITULO_BLOCO\](.*?)\[/TITULO_BLOCO\]",
            resultado_ollama,
            flags=re.IGNORECASE | re.DOTALL
        )
        if match_titulo_bloco:
            titulo_gerado = re.sub(
                r"\s+", " ", match_titulo_bloco.group(1)
            ).strip(" \t\r\n:;.-")
            if titulo_gerado:
                # Sincroniza a variável local com o título extraído.
                # Sem isso, uma gravação posterior reutiliza titulo_bloco vazio
                # e sobrescreve o título correto no JSON final.
                titulo_bloco = titulo_gerado
                titulos_python[chave_bloco] = titulo_gerado
                dados_bloco["titulo"] = titulo_gerado
                informacoes_blocos[chave_bloco] = dados_bloco
                print("TÍTULO MEAD GERADO:", titulo_gerado)
        else:
            # Fallback estritamente extractivo: deriva o título do conteúdo
            # efetivamente redigido no próprio bloco, sem acrescentar fatos.
            _textos_titulo = []
            for _par_titulo in (resultado_ollama or "",):
                for _idx_titulo in range(1, 4):
                    _m_par_titulo = re.search(
                        rf"\[PARAGRAFO_{_idx_titulo}\](.*?)\[/PARAGRAFO_{_idx_titulo}\]",
                        _par_titulo, flags=re.I | re.S
                    )
                    if _m_par_titulo:
                        _textos_titulo.append(_m_par_titulo.group(1).strip())
            _titulo_fallback = ""
            _ignorar_inicio_titulo = (
                "no entanto", "além disso", "alem disso", "dessa forma",
                "nesse contexto", "em resumo", "por outro lado",
                "confira", "neste artigo", "como escolher"
            )
            for _texto_titulo in _textos_titulo:
                _frase_titulo = re.split(r"(?<=[.!?])\s+", _texto_titulo, maxsplit=1)[0]
                _frase_titulo = re.sub(r"^\s*[^A-Za-zÀ-ÿ0-9]+", "", _frase_titulo).strip(" .,:;–—-")
                _frase_norm = normalizar_assunto_texto(_frase_titulo)
                if any(_frase_norm.startswith(x) for x in _ignorar_inicio_titulo):
                    _frase_titulo = re.sub(r"^[^,;:.]+[,;:.]\s*", "", _frase_titulo).strip()
                _palavras_titulo = _frase_titulo.split()
                if len(_palavras_titulo) >= 4:
                    _titulo_fallback = " ".join(_palavras_titulo[:9]).strip(" .,:;–—-")
                    break
            if _titulo_fallback:
                _titulo_fallback = _titulo_fallback[0].upper() + _titulo_fallback[1:]
                titulo_bloco = _titulo_fallback
                titulos_python[chave_bloco] = _titulo_fallback
                dados_bloco["titulo"] = _titulo_fallback
                informacoes_blocos[chave_bloco] = dados_bloco
                print("TÍTULO EXTRACTIVO DE RESERVA:", chave_bloco, _titulo_fallback)
            else:
                print("⚠️ OLLAMA NÃO RETORNOU TÍTULO E NÃO HÁ FRASE ADEQUADA:", chave_bloco)

        if chave_bloco == "bloco_1":
            match_subtitulo = re.search(
                r"\[SUBTITULO_PAGINA\](.*?)\[/SUBTITULO_PAGINA\]",
                resultado_ollama,
                flags=re.IGNORECASE | re.DOTALL
            )
            if match_subtitulo:
                subtitulo = re.sub(
                    r"\s+", " ", match_subtitulo.group(1)
                ).strip(" \t\r\n:;.-")
                print("SUBTÍTULO MEAD GERADO:", subtitulo)

        # ========================================================
        # EXTRAÇÃO DIRETA DOS 3 PARÁGRAFOS
        # ========================================================

        paragrafos_extraidos = []


        for indice_paragrafo in range(
            1,
            4
        ):

            marcador_abertura = (
                f"[PARAGRAFO_{indice_paragrafo}]"
            )

            marcador_fechamento = (
                f"[/PARAGRAFO_{indice_paragrafo}]"
            )


            # ----------------------------------------------------
            # LOCALIZAR ABERTURA
            # ----------------------------------------------------

            inicio = resultado_ollama.lower().find(
                marcador_abertura.lower()
            )


            if inicio == -1:

                # Alguns modelos omitem a abertura [PARAGRAFO_N], mas
                # mantêm o fechamento [/PARAGRAFO_N]. Recupera o início
                # pela fronteira estrutural anterior, sem absorver título
                # ou subtítulo no texto do parágrafo.
                if indice_paragrafo == 1:
                    fronteiras = [
                        "[/SUBTITULO_PAGINA]",
                        "[/TITULO_BLOCO]",
                    ]
                    posicoes_fronteira = [
                        resultado_ollama.lower().find(x.lower()) + len(x)
                        for x in fronteiras
                        if resultado_ollama.lower().find(x.lower()) >= 0
                    ]
                    inicio_conteudo_recuperado = max(posicoes_fronteira) if posicoes_fronteira else 0
                else:
                    fechamento_anterior = f"[/PARAGRAFO_{indice_paragrafo - 1}]"
                    pos_fechamento_anterior = resultado_ollama.lower().find(fechamento_anterior.lower())
                    inicio_conteudo_recuperado = (
                        pos_fechamento_anterior + len(fechamento_anterior)
                        if pos_fechamento_anterior >= 0 else -1
                    )

                if inicio_conteudo_recuperado < 0:
                    print("❌ ABERTURA AUSENTE E NÃO FOI POSSÍVEL DETERMINAR O INÍCIO:", marcador_abertura)
                    print("RESPOSTA:", resultado_ollama[:3000])
                    return None

                print("⚠️ MARCADOR DE ABERTURA AUSENTE:", marcador_abertura, "| início recuperado pela fronteira anterior")
                # Mantém compatibilidade com o cálculo de inicio_conteudo abaixo.
                inicio = inicio_conteudo_recuperado - len(marcador_abertura)


            # ----------------------------------------------------
            # POSIÇÃO APÓS O MARCADOR
            # ----------------------------------------------------

            inicio_conteudo = (
                inicio
                + len(
                    marcador_abertura
                )
            )



            # ----------------------------------------------------
            # LOCALIZAR FECHAMENTO
            # ----------------------------------------------------
            
            fim = resultado_ollama.lower().find(
                marcador_fechamento.lower(),
                inicio_conteudo
            )
            
            # ----------------------------------------------------
            # COMPATIBILIDADE:
            # se o Ollama não retornar o fechamento individual
            # do parágrafo, utilizar o próximo marcador como
            # limite do parágrafo atual.
            # ----------------------------------------------------
            
            if fim == -1:

                if indice_paragrafo < 3:
                    proximo_marcador = f"[PARAGRAFO_{indice_paragrafo + 1}]"
                    fim = resultado_ollama.lower().find(
                        proximo_marcador.lower(), inicio_conteudo
                    )
                else:
                    proximo_marcador = "FIM_DA_RESPOSTA_DO_BLOCO"
                    fim = len(resultado_ollama)
            
            
            if fim == -1:
            
                print()
                print(
                    "❌ FECHAMENTO DO PARÁGRAFO NÃO ENCONTRADO:"
                )
            
                print(
                    indice_paragrafo
                )
            
                print()
                print(
                    "MARCADORES PROCURADOS:"
                )
            
                print(
                    marcador_fechamento
                )
            
                print(
                    "OU LIMITE ALTERNATIVO:",
                    proximo_marcador
                    if 'proximo_marcador' in locals()
                    else ""
                )
            
                print()
                print(
                    "RESPOSTA RECEBIDA:"
                )
            
                print(
                    resultado_ollama[:3000]
                )
            
                return None


            # ----------------------------------------------------
            # EXTRAIR TEXTO
            # ----------------------------------------------------

            texto_paragrafo = (
                resultado_ollama[
                    inicio_conteudo:fim
                ]
            )


            texto_paragrafo = str(
                texto_paragrafo or ""
            ).strip()


            if not texto_paragrafo:

                print()
                print(
                    f"❌ PARÁGRAFO {indice_paragrafo} "
                    "RETORNOU VAZIO."
                )

                return None


            # ----------------------------------------------------
            # LIMPEZA MÍNIMA
            # ----------------------------------------------------

            texto_paragrafo = re.sub(
                r"\s+",
                " ",
                texto_paragrafo
            ).strip()


            if not texto_paragrafo:

                print()
                print(
                    f"❌ PARÁGRAFO {indice_paragrafo} "
                    "VAZIO APÓS LIMPEZA."
                )

                return None


            paragrafos_extraidos.append(
                texto_paragrafo
            )

        # ========================================================
        # GARANTIR EXATAMENTE 3 PARÁGRAFOS
        # ========================================================

        if len(
            paragrafos_extraidos
        ) != 3:

            print()
            print(
                "❌ QUANTIDADE INCORRETA DE PARÁGRAFOS:"
            )

            print(
                len(
                    paragrafos_extraidos
                )
            )

            return None


        # ========================================================
        # NORMALIZAÇÃO DA PALAVRA-CHAVE — SEM NOVA CHAMADA OLLAMA
        # ========================================================
        # Uma ocorrência em caixa diferente é apenas normalizada.
        # Se o modelo repetir a expressão, preservamos a primeira e
        # substituímos as referências seguintes por uma retomada neutra,
        # sem inserir uma nova palavra-chave nem descartar o bloco inteiro.
        # Se a expressão estiver ausente, o parágrafo continua reprovado:
        # não inventamos nem inserimos uma âncora que o modelo omitiu.
        # ========================================================
        reservas_com = dados_bloco.get("reservas_com", []) if isinstance(dados_bloco, dict) else []
        if not isinstance(reservas_com, list):
            reservas_com = []
        hashes_reserva_tentadas = set()

        padrao_keyword_ci = re.compile(
            r"(?<!\w)" + re.escape(str(palavra_chave_redacao or "").strip()) + r"(?!\w)",
            flags=re.IGNORECASE | re.UNICODE
        ) if str(palavra_chave_redacao or "").strip() else None

        if tipo == "produto":
            referencia_secundaria_keyword = "o equipamento"
        elif tipo == "servico":
            referencia_secundaria_keyword = "o serviço"
        else:
            referencia_secundaria_keyword = "esse tema"

        for indice_keyword, paragrafo_keyword in enumerate(list(paragrafos_extraidos)):
            ocorrencias_ci = len(padrao_keyword_ci.findall(paragrafo_keyword)) if padrao_keyword_ci else 0
            if ocorrencias_ci >= 1:
                contador_keyword = {"n": 0}

                def _normalizar_uma_keyword(match_keyword):
                    contador_keyword["n"] += 1
                    if contador_keyword["n"] == 1:
                        return palavra_chave_redacao
                    return referencia_secundaria_keyword

                texto_normalizado_keyword = padrao_keyword_ci.sub(
                    _normalizar_uma_keyword,
                    paragrafo_keyword
                )
                paragrafos_extraidos[indice_keyword] = texto_normalizado_keyword

                if ocorrencias_ci == 1 and texto_normalizado_keyword != paragrafo_keyword:
                    print(
                        "✅ CAPITALIZAÇÃO DA PALAVRA-CHAVE NORMALIZADA:",
                        chave_bloco,
                        "PARÁGRAFO", indice_keyword + 1,
                        "→", palavra_chave_redacao
                    )
                elif ocorrencias_ci > 1:
                    print(
                        "⚠️ REPETIÇÃO DA PALAVRA-CHAVE CORRIGIDA:",
                        chave_bloco,
                        "PARÁGRAFO", indice_keyword + 1,
                        "OCORRÊNCIAS ORIGINAIS:", ocorrencias_ci,
                        "REFERÊNCIAS POSTERIORES SUBSTITUÍDAS POR:",
                        referencia_secundaria_keyword
                    )
                continue

            print(
                "⚠️ PALAVRA-CHAVE AUSENTE:",
                chave_bloco,
                "PARÁGRAFO", indice_keyword + 1,
                "— procurando reserva COM já existente, sem inserir a expressão no texto."
            )
            reserva_substituta = None
            for reserva_keyword in reservas_com:
                if not isinstance(reserva_keyword, dict):
                    continue
                texto_reserva_keyword = re.sub(
                    r"\s+", " ", str(reserva_keyword.get("texto", "") or "")
                ).strip()
                hash_reserva_keyword = str(reserva_keyword.get("hash", "") or "").strip()
                ocorrencias_reserva = len(padrao_keyword_ci.findall(texto_reserva_keyword)) if padrao_keyword_ci else 0
                if (
                    not texto_reserva_keyword
                    or str(reserva_keyword.get("palavra_chave_status", "")).upper() != "COM"
                    or ocorrencias_reserva != 1
                    or len(texto_reserva_keyword.split()) < 25
                    or not re.search(r"[.!?]$", texto_reserva_keyword)
                    or (hash_reserva_keyword and hash_reserva_keyword in hashes_reservas_ocupadas)
                ):
                    continue
                reserva_substituta = reserva_keyword
                break

            if reserva_substituta is None:
                print("❌ NENHUMA RESERVA COM ADEQUADA; bloco preservado para recuperação, sem perder os demais dados.")
                dados_bloco["paragrafos_ollama"] = list(paragrafos_extraidos)
                dados_bloco["paragrafos_python"] = list(paragrafos_python)
                dados_bloco["informacoes_relevantes"] = list(informacoes_relevantes)
                titulo_bloco = str(
                    titulos_python.get(chave_bloco, "")
                    or dados_bloco.get("titulo", "")
                    or titulo_bloco
                    or ""
                ).strip()
                dados_bloco["titulo"] = titulo_bloco
                salvar_bloco_recuperacao(tema, numero_bloco, dados_bloco)
                _salvar_diagnostico_ultima_execucao(
                    tema,
                    {"bloco": chave_bloco, "paragrafo": indice_keyword + 1,
                     "palavra_chave_exigida": palavra_chave_redacao,
                     "paragrafos_ollama": paragrafos_extraidos},
                    "PALAVRA_CHAVE_AUSENTE_SEM_RESERVA_COM",
                    "Não foi encontrada reserva COM adequada para substituir a posição sem inventar texto."
                )
                return None

            texto_reserva_substituta = str(reserva_substituta.get("texto", "") or "").strip()
            contador_reserva = {"n": 0}
            def _normalizar_keyword_reserva(match_reserva):
                contador_reserva["n"] += 1
                return palavra_chave_redacao
            texto_reserva_substituta = padrao_keyword_ci.sub(
                _normalizar_keyword_reserva, texto_reserva_substituta
            )
            paragrafos_extraidos[indice_keyword] = texto_reserva_substituta
            hash_reserva_substituta = str(reserva_substituta.get("hash", "") or gerar_hash_trecho(texto_reserva_substituta))
            hashes_reservas_ocupadas.add(hash_reserva_substituta)
            fragmentos_bloco[indice_keyword] = {
                "numero": indice_keyword + 1,
                "id": reserva_substituta.get("id", ""),
                "hash": hash_reserva_substituta,
                "palavra_chave_status": "COM",
                "texto": str(reserva_substituta.get("texto", "") or "").strip(),
                "fonte": reserva_substituta.get("fonte", ""),
                "url": reserva_substituta.get("url", ""),
                "tipo": reserva_substituta.get("tipo", ""),
                "pdf": reserva_substituta.get("pdf", False),
                "palavras": reserva_substituta.get("palavras", len(texto_reserva_substituta.split())),
                "identidade_fonte": reserva_substituta.get("identidade_fonte", {}),
                "contexto_geral_sem_fonte": False
            }
            informacoes_relevantes[indice_keyword] = dict(reserva_substituta)
            paragrafos_python[indice_keyword] = fragmentos_bloco[indice_keyword]["texto"]
            try:
                reservas_com.remove(reserva_substituta)
            except (ValueError, AttributeError):
                pass
            print(
                "✅ PARÁGRAFO RECUPERADO COM RESERVA COM:",
                chave_bloco, "P", indice_keyword + 1,
                "| HASH:", hash_reserva_substituta,
                "| PALAVRA-CHAVE: 1/1",
                "| TEXTO ORIGINAL DA FONTE PRESERVADO"
            )

        # Recontagem obrigatória com capitalização e acentuação exatas.
        for indice_recontagem, paragrafo_recontagem in enumerate(paragrafos_extraidos, 1):
            quantidade_recontagem = _mead_contar_keyword_exata(paragrafo_recontagem, palavra_chave_redacao)
            if quantidade_recontagem != 1:
                print(
                    "❌ GATE FINAL DA PALAVRA-CHAVE FALHOU:",
                    chave_bloco, "P", indice_recontagem,
                    "OCORRÊNCIAS:", quantidade_recontagem
                )
                return None

        # ========================================================
        # VALIDAÇÃO FACTUAL PÓS-OLLAMA — v7.5
        # ========================================================
        #
        # A versão anterior exigia sobreposição lexical quase
        # literal. Isso rejeitava paráfrases tecnicamente corretas.
        #
        # Agora a validação usa sinais complementares:
        # 1) cobertura por radicais leves;
        # 2) termos relevantes compartilhados;
        # 3) pequenas sequências de palavras;
        # 4) preservação obrigatória de números/unidades;
        # 5) bloqueio de parágrafo sem evidência suficiente.
        #
        # O MEAD e o contexto editorial NÃO entram como evidência.
        # Cada parágrafo é validado somente contra seu fragmento
        # autorizado correspondente.
        # ========================================================

        palavras_ruido_validacao = {
            "para", "com", "sem", "sobre", "entre", "como", "uma", "umas",
            "um", "uns", "dos", "das", "que", "por", "pelo", "pela",
            "aos", "nas", "nos", "nas", "e", "ou", "de", "da", "do",
            "em", "no", "na", "ao", "se", "ser", "sao", "são", "foi",
            "tem", "ter", "pode", "podem", "mais", "tambem", "também",
            "isso", "esse", "essa", "este", "esta", "esses", "essas",
            "estes", "estas", "quando", "onde", "assim", "cada", "qual",
            "forma", "tipo", "modo", "parte", "caso", "seu", "sua",
            "seus", "suas", "outro", "outra", "outros", "outras"
        }

        def _tokens_validacao(texto):
            normalizado = normalizar_assunto_texto(str(texto or ""))
            return [
                t for t in re.findall(r"\b[a-z0-9]{3,}\b", normalizado)
                if t not in palavras_ruido_validacao
            ]

        def _radical_leve(token):
            t = str(token or "")
            if len(t) <= 5:
                return t
            # Redução conservadora de flexões frequentes em português.
            sufixos = (
                "mente", "ções", "ção", "sões", "são", "mente",
                "amentos", "imentos", "amento", "imento", "idades",
                "idade", "ismos", "ismo", "istas", "ista", "mente",
                "ados", "idos", "adas", "idas", "ando", "endo", "indo",
                "ados", "idos", "es", "as", "os", "s"
            )
            for sufixo in sufixos:
                if t.endswith(sufixo) and len(t) - len(sufixo) >= 4:
                    return t[:-len(sufixo)]
            return t

        def _n_gramas(tokens, n=2):
            return set(
                " ".join(tokens[i:i+n])
                for i in range(max(0, len(tokens)-n+1))
            )

        def _avaliar_fidelidade_paragrafo(texto_saida, texto_fonte):
            tokens_fonte = _tokens_validacao(texto_fonte)
            tokens_saida = _tokens_validacao(texto_saida)
            radicais_fonte = {_radical_leve(t) for t in tokens_fonte}
            radicais_saida = [_radical_leve(t) for t in tokens_saida]
            compartilhados = [t for t in radicais_saida if t in radicais_fonte]
            cobertura_radical = len(compartilhados) / max(1, len(radicais_saida))
            termos_fonte = set(tokens_fonte)
            termos_saida = set(tokens_saida)
            termos_compartilhados = termos_fonte & termos_saida
            cobertura_termos = len(termos_compartilhados) / max(1, len(termos_saida))
            bigramas_fonte = _n_gramas(tokens_fonte, 2)
            bigramas_saida = _n_gramas(tokens_saida, 2)
            sobreposicao_bigrama = len(bigramas_fonte & bigramas_saida) / max(1, len(bigramas_saida))
            score = (cobertura_radical * 0.60) + (cobertura_termos * 0.25) + (sobreposicao_bigrama * 0.15)
            numeros_fonte = set(re.findall(
                r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)",
                normalizar_assunto_texto(texto_fonte)
            ))
            numeros_saida = set(re.findall(
                r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)",
                normalizar_assunto_texto(texto_saida)
            ))
            return {
                "score": score,
                "radical": cobertura_radical,
                "termos": cobertura_termos,
                "bigramas": sobreposicao_bigrama,
                "termos_compartilhados": termos_compartilhados,
                "tokens_saida": tokens_saida,
                "numeros_novos": numeros_saida - numeros_fonte,
            }

        # O texto inicial é validado normalmente. Se falhar, somente o
        # parágrafo afetado é refeito, primeiro com sua própria fonte COM e,
        # depois, com reservas COM. Os demais parágrafos aprovados ficam intactos.
        hashes_factual_tentadas = set(hashes_reserva_tentadas)
        for indice_paragrafo, resultado_paragrafo in enumerate(list(paragrafos_extraidos)):
            fragmento_atual = fragmentos_bloco[indice_paragrafo]
            texto_fonte_atual = str(fragmento_atual.get("texto", "") or "")
            avaliacao = _avaliar_fidelidade_paragrafo(resultado_paragrafo, texto_fonte_atual)
            falha_numerica = bool(avaliacao["numeros_novos"])
            # O conteúdo pode legitimamente desenvolver conhecimento geral do tema;
            # similaridade lexical com a fonte deixa de ser requisito de aprovação.
            falha_evidencia = False

            if falha_numerica or falha_evidencia:
                motivo = "número/unidade não autorizado" if falha_numerica else "evidência insuficiente"
                print("⚠️ RECUPERAÇÃO FACTUAL PONTUAL:", chave_bloco, "P", indice_paragrafo + 1, "MOTIVO:", motivo)
                if falha_numerica:
                    print("NÚMEROS NOVOS:", sorted(avaliacao["numeros_novos"]))
                else:
                    print("SCORE INICIAL:", round(avaliacao["score"], 3))

                candidatos = []
                hash_atual_factual = ""
                # A fonte atual só pode ser reutilizada se estiver marcada COM
                # e contiver de fato a expressão literal.
                if (str(fragmento_atual.get("palavra_chave_status", "")).upper() == "COM"
                        and contar_palavra_chave_literal(texto_fonte_atual, tema) >= 1):
                    hash_atual_factual = str(fragmento_atual.get("hash", "") or gerar_hash_trecho(texto_fonte_atual))
                    candidatos.append({
                        "id": fragmento_atual.get("id", ""),
                        "hash": hash_atual_factual,
                        "texto": texto_fonte_atual,
                        "fonte": fragmento_atual.get("fonte", ""),
                        "url": fragmento_atual.get("url", ""),
                        "tipo": fragmento_atual.get("tipo", ""),
                        "pdf": fragmento_atual.get("pdf", False),
                        "palavras": fragmento_atual.get("palavras", len(texto_fonte_atual.split())),
                        "identidade_fonte": fragmento_atual.get("identidade_fonte", {}),
                        "palavra_chave_status": "COM",
                    })
                for reserva in reservas_com:
                    if not isinstance(reserva, dict):
                        continue
                    texto_reserva = str(reserva.get("texto", "") or "").strip()
                    hash_reserva = str(reserva.get("hash", "") or gerar_hash_trecho(texto_reserva))
                    if (not texto_reserva or not hash_reserva
                            or str(reserva.get("palavra_chave_status", "")).upper() != "COM"
                            or contar_palavra_chave_literal(texto_reserva, tema) < 1):
                        continue
                    if hash_reserva not in hashes_factual_tentadas and hash_reserva != hash_atual_factual:
                        candidatos.append(reserva)

                recuperado_factual = False
                for candidato in candidatos:
                    texto_candidato = str(candidato.get("texto", "") or "").strip()
                    hash_candidato = str(candidato.get("hash", "") or gerar_hash_trecho(texto_candidato))
                    if not hash_candidato or hash_candidato in hashes_factual_tentadas:
                        continue
                    hashes_factual_tentadas.add(hash_candidato)
                    hashes_reservas_ocupadas.add(hash_candidato)
                    try:
                        reservas_com.remove(candidato)
                    except (ValueError, AttributeError):
                        pass
                    prompt_factual = f"""
Reescreva somente UM parágrafo técnico em português brasileiro. Use os fatos do fragmento como base e complemente apenas com conhecimento geral consolidado sobre o tema quando necessário; não invente dados específicos.

TEMA / EXPRESSÃO EXATA OBRIGATÓRIA: {palavra_chave_redacao}
A expressão exata deve aparecer exatamente UMA vez.

REGRAS OBRIGATÓRIAS:
- Não acrescente números, quantidades, medidas, unidades, percentuais, datas ou especificações que não estejam escritos no fragmento.
- Não transforme números de títulos, listas, páginas ou enumerações em dados técnicos.
- Não invente características, materiais, marcas, modelos, aplicações, resultados ou garantias.
- Preserve os fatos do fragmento. Quando ele não trouxer informação suficiente,
  use somente conhecimento geral consolidado sobre o tema, sem criar especificações.
- Produza um único parágrafo completo, natural, com pelo menos 25 palavras e pontuação final.
- Não mencione fonte, fragmento, pesquisa, Python ou estas instruções.

FRAGMENTO AUTORIZADO:
{texto_candidato}

Retorne somente o parágrafo.
""".strip()
                    print("⚠️ RESERVA FACTUAL NÃO REESCRITA: a página já consumiu sua única chamada Ollama.")
                    continue

                    texto_tentativa = re.sub(r"^```(?:text|txt)?\s*|\s*```$", "", texto_tentativa, flags=re.I).strip()
                    texto_tentativa = re.sub(r"\s+", " ", texto_tentativa).strip()
                    if (not texto_tentativa
                            or _mead_contar_keyword_exata(texto_tentativa, palavra_chave_redacao) != 1
                            or not re.search(r"[.!?]$", texto_tentativa)):
                        print("⚠️ TENTATIVA FACTUAL DESCARTADA: estrutura/palavra-chave inválida.", candidato.get("id", ""))
                        continue

                    avaliacao_tentativa = _avaliar_fidelidade_paragrafo(texto_tentativa, texto_candidato)
                    falha_num_tentativa = bool(avaliacao_tentativa["numeros_novos"])
                    falha_evid_tentativa = False
                    if falha_num_tentativa or falha_evid_tentativa:
                        print("⚠️ TENTATIVA FACTUAL REPROVADA:", candidato.get("id", ""),
                              "NÚMEROS:", sorted(avaliacao_tentativa["numeros_novos"]),
                              "SCORE:", round(avaliacao_tentativa["score"], 3))
                        continue

                    paragrafos_extraidos[indice_paragrafo] = texto_tentativa
                    fragmentos_bloco[indice_paragrafo].update({
                        "id": candidato.get("id", ""),
                        "hash": hash_candidato,
                        "palavra_chave_status": "COM",
                        "texto": texto_candidato,
                        "fonte": candidato.get("fonte", ""),
                        "url": candidato.get("url", ""),
                        "tipo": candidato.get("tipo", ""),
                        "pdf": candidato.get("pdf", False),
                        "palavras": candidato.get("palavras", len(texto_candidato.split())),
                        "identidade_fonte": candidato.get("identidade_fonte", {})
                    })
                    informacoes_relevantes[indice_paragrafo] = dict(candidato)
                    paragrafos_python[indice_paragrafo] = texto_candidato
                    avaliacao = avaliacao_tentativa
                    print("✅ RECUPERAÇÃO FACTUAL APROVADA:", candidato.get("id", ""),
                          "SCORE:", round(avaliacao["score"], 3),
                          "PALAVRA-CHAVE: 1/1", "NÚMEROS NOVOS: 0")
                    recuperado_factual = True
                    break

                if not recuperado_factual:
                    print("❌ RECUPERAÇÃO FACTUAL ESGOTADA:", chave_bloco, "P", indice_paragrafo + 1)
                    _salvar_diagnostico_ultima_execucao(
                        tema,
                        {
                            "bloco": chave_bloco,
                            "paragrafo": indice_paragrafo + 1,
                            "motivo": motivo,
                            "numeros_novos": sorted(avaliacao["numeros_novos"]),
                            "score": round(avaliacao["score"], 4),
                            "paragrafos_ollama": paragrafos_extraidos,
                            "informacoes_blocos": informacoes_blocos
                        },
                        "RECUPERACAO_FACTUAL_SEM_RESERVA_COM",
                        "O parágrafo falhou na validação factual e nenhuma fonte COM disponível produziu substituição aprovada."
                    )
                    return None

            # Validação final da posição, inclusive após eventual substituição.
            texto_final_validado = paragrafos_extraidos[indice_paragrafo]
            if _mead_contar_keyword_exata(texto_final_validado, palavra_chave_redacao) != 1:
                print("❌ GATE COM FINAL FALHOU APÓS VALIDAÇÃO FACTUAL:", chave_bloco, "P", indice_paragrafo + 1)
                return None
            avaliacao_final = _avaliar_fidelidade_paragrafo(
                texto_final_validado,
                fragmentos_bloco[indice_paragrafo].get("texto", "")
            )
            # Similaridade lexical com a fonte não é obrigatória: conhecimento
            # geral autorizado pode ampliar a explicação. Números novos seguem
            # bloqueados para evitar especificações não comprovadas.
            if avaliacao_final["numeros_novos"]:
                print("❌ VALIDAÇÃO FACTUAL FINAL FALHOU:", chave_bloco, "P", indice_paragrafo + 1)
                return None
            print(
                "🟢 VALIDAÇÃO FACTUAL:", chave_bloco,
                "PARÁGRAFO", indice_paragrafo + 1,
                "SCORE", round(avaliacao_final["score"], 3),
                "TERMOS", len(avaliacao_final["termos_compartilhados"]),
                "NÚMEROS NOVOS", 0
            )

        # ========================================================
        # GUARDAR OS 3 RESULTADOS
        # ========================================================
        
        for indice_paragrafo in range(
            3
        ):
        
            resultado_paragrafo = (
                paragrafos_extraidos[
                    indice_paragrafo
                ]
            )
        
            paragrafos_ollama[
                indice_paragrafo
            ] = resultado_paragrafo


        # ========================================================
        # GUARDAR ESTRUTURA DO BLOCO
        # ========================================================

        dados_bloco[
            "paragrafos_python"
        ] = paragrafos_python


        dados_bloco[
            "paragrafos_ollama"
        ] = paragrafos_ollama


        # Reconfere o título antes de persistir, priorizando o título extraído.
        titulo_bloco = str(
            titulos_python.get(chave_bloco, "")
            or dados_bloco.get("titulo", "")
            or titulo_bloco
            or ""
        ).strip()

        dados_bloco[
            "titulo"
        ] = titulo_bloco


        dados_bloco[
            "hash"
        ] = hash_bloco
        

        # ========================================================
        # GUARDAR ESTRUTURA DO BLOCO EM MEMÓRIA
        # ========================================================
        #
        # O bloco NÃO é salvo no JSON neste momento.
        #
        # Todos os 5 blocos serão mantidos em memória e somente
        # depois da validação completa da página serão enviados
        # ao salvar_banco() na gravação oficial.
        #
        # ========================================================

        dados_bloco[
            "paragrafos_python"
        ] = paragrafos_python


        dados_bloco[
            "paragrafos_ollama"
        ] = paragrafos_ollama


        # Reconfere o título antes de persistir, priorizando o título extraído.
        titulo_bloco = str(
            titulos_python.get(chave_bloco, "")
            or dados_bloco.get("titulo", "")
            or titulo_bloco
            or ""
        ).strip()

        dados_bloco[
            "titulo"
        ] = titulo_bloco


        dados_bloco[
            "hash"
        ] = hash_bloco


        # ========================================================
        # LOG DE CONFIRMAÇÃO
        # ========================================================

        print()
        print(
            "✅ BLOCO PROCESSADO"
        )

        print(
            "BLOCO:",
            numero_bloco
        )

        print(
            "FRAGMENTOS ENVIADOS:",
            3
        )

        print(
            "PARÁGRAFOS RECEBIDOS:",
            len(
                [
                    item
                    for item in paragrafos_ollama
                    if str(item or "").strip()
                ]
            )
        )

        print("PALAVRAS P1/P2/P3:", "/".join(str(len(x.split())) for x in paragrafos_ollama[:3]))

        print(
            "TEMPO TOTAL DO BLOCO:",
            round(
                fim_ollama - inicio_ollama,
                2
            ),
            "segundos"
        )

        print("=" * 60)


        # ========================================================
        # ATUALIZAR O DICIONÁRIO DO BLOCO
        # ========================================================

        informacoes_blocos[
            chave_bloco
        ] = dados_bloco

        # Persistir o bloco validado antes de avançar para o próximo.
        # Se a etapa posterior falhar, a próxima execução o reutiliza.
        salvar_bloco_recuperacao(
            tema,
            numero_bloco,
            dados_bloco
        )
        percentual_etapas = 20 + (numero_bloco * 12)
        print(
            f"PROGRESSO DA PÁGINA: {percentual_etapas}% — "
            f"bloco {numero_bloco}/5 validado e preservado em progresso.json."
        )


        # ========================================================
        # LOG DE CONFIRMAÇÃO
        # ========================================================

        print()
        print(
            "✅ BLOCO PROCESSADO E MANTIDO EM MEMÓRIA — AINDA NÃO GRAVADO NO JSON OFICIAL"
        )

        print(
            "BLOCO:",
            numero_bloco
        )

        print(
            "FRAGMENTOS ENVIADOS:",
            3
        )

        print(
            "PARÁGRAFOS RECEBIDOS:",
            len(
                [
                    item
                    for item in paragrafos_ollama
                    if str(item or "").strip()
                ]
            )
        )

        print(
            "PALAVRAS PARÁGRAFO 1:",
            len(
                paragrafos_ollama[0].split()
            )
        )

        print(
            "PALAVRAS PARÁGRAFO 2:",
            len(
                paragrafos_ollama[1].split()
            )
        )

        print(
            "PALAVRAS PARÁGRAFO 3:",
            len(
                paragrafos_ollama[2].split()
            )
        )

        print(
            "TEMPO TOTAL DO BLOCO:",
            round(
                fim_ollama - inicio_ollama,
                2
            ),
            "segundos"
        )

        print("=" * 60)


        # ========================================================
        # ATUALIZAR O DICIONÁRIO DO BLOCO
        # ========================================================

        informacoes_blocos[
            chave_bloco
        ] = dados_bloco


    # ============================================================
    # TÍTULOS DINÂMICOS: TODOS OS PARÁGRAFOS JÁ FORAM PROCESSADOS
    # ============================================================
    titulos_dinamicos_ok = all(
        str(informacoes_blocos.get(f"bloco_{n}", {}).get("titulo", "") or "").strip()
        for n in range(1, 6)
    )
    print(
        "TÍTULOS DINÂMICOS:",
        "5/5 recebidos após a redação" if titulos_dinamicos_ok
        else "algum título ausente; será registrado e tratado na validação"
    )
    if titulos_dinamicos_ok:
        for n in range(1, 6):
            print(f"TÍTULO BLOCO {n}:", informacoes_blocos[f"bloco_{n}"].get("titulo", ""))

    print("PROGRESSO DA PÁGINA: 80% — redação concluída; iniciando validações.")

    # ============================================================
    # VALIDAR OS 5 BLOCOS
    # ============================================================

    print()
    print("=" * 60)
    print("VALIDAÇÃO FINAL DO OLLAMA")
    print("=" * 60)


    total_paragrafos_ollama = 0


    for numero_bloco in range(
        1,
        total_blocos + 1
    ):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        dados_bloco = (
            informacoes_blocos.get(
                chave_bloco,
                {}
            )
        )

        paragrafos_ollama = (
            dados_bloco.get(
                "paragrafos_ollama",
                []
            )
            if isinstance(
                dados_bloco,
                dict
            )
            else []
        )

        quantidade = len([
            item
            for item in paragrafos_ollama
            if str(item or "").strip()
        ])


        print(
            chave_bloco + ":",
            quantidade,
            "/ 3"
        )


        total_paragrafos_ollama += (
            quantidade
        )


    print()
    print(
        "TOTAL PARÁGRAFOS OLLAMA:",
        total_paragrafos_ollama,
        "/",
        total_blocos * 3
    )


    if total_paragrafos_ollama != (
        total_blocos * 3
    ):

        print()
        print(
            "❌ A PÁGINA NÃO ESTÁ COMPLETA."
        )

        print(
            "O processamento será interrompido."
        )

        return None


    print()
    print(
        "✅ OS 15 PARÁGRAFOS FORAM PROCESSADOS."
    )

    print("=" * 60)


    # ============================================================
    # AUDITORIA INDEPENDENTE MEAD — DESVER / ISI / FREIO
    # ============================================================
    auditoria_independente_mead = _mead_auditoria_independente(
        informacoes_blocos, tema
    )
    print(
        "AUDITORIA MEAD — DESVER:",
        len(auditoria_independente_mead.get("desver", [])),
        "| ISI:",
        len(auditoria_independente_mead.get("isi", [])),
        "| KEYWORD:",
        auditoria_independente_mead.get("keyword_ocorrencias_exatas", 0),
        "/ 15"
    )
    print(
        "FREIO DO REFINAMENTO:",
        auditoria_independente_mead.get("freio_do_refinamento", "")
    )
    if auditoria_independente_mead.get("desver"):
        print(
            "ALERTAS DESVER:",
            json.dumps(auditoria_independente_mead["desver"], ensure_ascii=False)[:5000]
        )
    if auditoria_independente_mead.get("isi"):
        print(
            "ALERTAS ISI PARA REVISÃO:",
            json.dumps(auditoria_independente_mead["isi"], ensure_ascii=False)[:3000]
        )

    # ============================================================
    # GATE DE QUALIDADE — NOTA MÍNIMA 7,5 POR BLOCO
    # ============================================================
    qualidade_blocos = []
    qualidade_reprovada = False
    for numero_bloco_gate in range(1, 6):
        chave_gate = f"bloco_{numero_bloco_gate}"
        bloco_gate = informacoes_blocos.get(chave_gate, {})
        paragrafos_gate = bloco_gate.get("paragrafos_ollama", []) if isinstance(bloco_gate, dict) else []
        fragmentos_gate = patrimonio_global_mead
        if not isinstance(fragmentos_gate, list):
            fragmentos_gate = []
        aprovado_gate, nota_gate, falhas_gate, detalhes_gate = avaliar_qualidade_final_python(
            str(bloco_gate.get("titulo", "") or chave_gate),
            paragrafos_gate,
            fragmentos_gate,
            chave_gate,
            tema
        )
        registro_gate = {
            "bloco": chave_gate,
            "aprovado": bool(aprovado_gate and nota_gate >= 7.5),
            "nota": nota_gate,
            "falhas": falhas_gate,
            "detalhes": detalhes_gate
        }
        qualidade_blocos.append(registro_gate)
        print(f"INDICADOR EDITORIAL {chave_gate}: {nota_gate:.2f}/10 | {'APROVADO' if registro_gate['aprovado'] else 'REVISAR'}")
        if falhas_gate:
            print("FALHAS:", "; ".join(falhas_gate))
        if not registro_gate["aprovado"]:
            qualidade_reprovada = True

    nota_minima_pagina = min((x["nota"] for x in qualidade_blocos), default=0.0)
    media_pagina = sum(x["nota"] for x in qualidade_blocos) / max(1, len(qualidade_blocos))
    print(f"NOTA MÍNIMA ENTRE BLOCOS: {nota_minima_pagina:.2f}/10 | MÉDIA: {media_pagina:.2f}/10")
    if qualidade_reprovada or nota_minima_pagina < 7.5:
        # A auditoria editorial continua ativa, mas não impede a persistência.
        # A prioridade desta etapa é gravar os 15 parágrafos para permitir
        # inspeção e correção posterior. A validação estrutural do JSON
        # permanece obrigatória dentro de salvar_banco().
        print("⚠️ QUALIDADE EDITORIAL ABAIXO DO ALVO — CONTINUANDO PARA TENTAR GRAVAR O JSON.")
        print("RESUMO DE QUALIDADE:", " | ".join(
            f"{x.get('bloco')}: nota {x.get('nota', 0):.2f}; "
            f"{'OK' if x.get('aprovado') else 'revisar'}"
            for x in qualidade_blocos
        ))
        _salvar_diagnostico_ultima_execucao(
            tema,
            {
                "tipo": "PAGINA_15_PARAGRAFOS_INDIVIDUAIS",
                "qualidade_blocos": qualidade_blocos,
                "informacoes_blocos": informacoes_blocos,
                "nota_minima": nota_minima_pagina,
                "media": media_pagina,
                "salvamento": "PERMITIDO_COM_ALERTAS; estrutura continua validada em salvar_banco",
            },
            "SALVAMENTO_PERMITIDO_COM_ALERTAS_DE_QUALIDADE",
            " | ".join(
                f"{x.get('bloco')}: {', '.join(x.get('falhas', []) or [])}"
                for x in qualidade_blocos if not x.get("aprovado")
            ) or f"Nota mínima {nota_minima_pagina:.2f} abaixo de 7,5"
        )

    # ============================================================
    # MONTAR LISTA FINAL DOS BLOCOS
    # ============================================================

    blocos = []


    for numero_bloco in range(
        1,
        total_blocos + 1
    ):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        dados_bloco = (
            informacoes_blocos.get(
                chave_bloco,
                {}
            )
        )

        if not isinstance(
            dados_bloco,
            dict
        ):

            continue


        blocos.append({

            "numero":
                numero_bloco,

            "id":
                chave_bloco,

            "titulo":
                str(
                    dados_bloco.get(
                        "titulo",
                        titulos_python.get(
                            chave_bloco,
                            ""
                        )
                    )
                    or ""
                ).strip(),

            "hash":
                str(
                    dados_bloco.get(
                        "hash",
                        ""
                    )
                    or ""
                ).strip(),

            "informacoes_relevantes":
                dados_bloco.get(
                    "informacoes_relevantes",
                    []
                ),

            "paragrafos_python":
                dados_bloco.get(
                    "paragrafos_python",
                    []
                ),

            "paragrafos_ollama":
                dados_bloco.get(
                    "paragrafos_ollama",
                    []
                )
        })


    # ============================================================
    # CONTROLE FINAL
    # ============================================================
    
    print()
    print("=" * 60)
    print("ESTRUTURA FINAL DOS BLOCOS")
    print("=" * 60)
    
    print(
        "BLOCOS:",
        len(blocos)
    )
    
    print(
        "PARÁGRAFOS:",
        sum(
            len(
                [
                    p
                    for p in bloco.get(
                        "paragrafos_ollama",
                        []
                    )
                    if str(p or "").strip()
                ]
            )
            for bloco in blocos
        )
    )
    
    print("=" * 60)


    # ========================================================
    # LIMPEZA PADRONIZADA DAS LISTAS
    # ========================================================
    #
    # A limpeza precisa existir antes da validação final
    # e antes de qualquer gravação no JSON.
    #
    # ========================================================

    def remover_duplicados(
        lista
    ):

        resultado = []
        vistos = set()

        for item in lista:

            item_limpo = str(
                item
            ).strip()

            chave = item_limpo.casefold()

            if not item_limpo:
                continue

            if chave in vistos:
                continue

            vistos.add(
                chave
            )

            resultado.append(
                item_limpo
            )

        return resultado
        

    # ========================================================
    # 17. CRIAR SEGMENTOS COM PYTHON
    # ========================================================
    #
    # Os segmentos NÃO são gerados pelos checkboxes.
    #
    # O Python seleciona exatamente 12 opções da biblioteca fixa,
    # priorizando aplicações compatíveis com o tema.
    #
    # Cada segmento obrigatoriamente referencia o tema,
    # que representa o produto ou serviço.
    #
    # ========================================================
    
    lista_segmentos = gerar_segmentos_pagina(
        tema,
        tipo
    )

    lista_segmentos = remover_duplicados(
        lista_segmentos
    )

    if len(lista_segmentos) != 12:

        print()
        print(
            "❌ FALHA NA GERAÇÃO DOS SEGMENTOS."
        )

        print(
            "A página não será considerada válida "
            "sem exatamente 12 segmentos."
        )

        print(
            "SEGMENTOS VÁLIDOS:",
            len(lista_segmentos)
        )

        return None



    # ========================================================
    # 18. CRIAR TAGS COM PYTHON
    # ========================================================
    #
    # As tags NÃO são mais geradas pelo Ollama.
    #
    # O Python monta exatamente 30 tags a partir do tema,
    # mantendo o padrão PRODUTO/SERVIÇO.
    #
    # REGRA OBRIGATÓRIA:
    # A palavra-chave NUNCA pode aparecer sozinha como tag.
    #
    # Exemplo:
    #
    # ❌ "bomba centrifuga"
    #
    # ✅ "bomba centrifuga industrial"
    # ✅ "manutenção de bomba centrifuga"
    # ✅ "aplicações de bomba centrifuga"
    #
    # ========================================================

    # --------------------------------------------------------
    # BASE FIXA DE 10/09/2026
    # A base padronizada pertence ao Python.
    # Os placeholders [tema] são substituídos pelo tema real.
    # --------------------------------------------------------

    tags_fixas = (
        TAGS_FIXAS_SERVICOS
        if tipo == "servico"
        else TAGS_FIXAS_PRODUTOS
    )

    tags_base_fixas = [
        str(tag).replace(
            "[tema]",
            palavra_chave_tag
        ).casefold()
        for tag in tags_fixas
    ]

    # --------------------------------------------------------
    # Variações adicionais do Python
    # Servem apenas para completar as 30 tags quando necessário.
    # --------------------------------------------------------

    if tipo == "servico":

        tags_base_adicionais = [
            f"{palavra_chave_tag} industrial",
            f"{palavra_chave_tag} técnica",
            f"{palavra_chave_tag} especializada",
            f"{palavra_chave_tag} preventiva",
            f"{palavra_chave_tag} corretiva",
            f"{palavra_chave_tag} profissional",
            f"{palavra_chave_tag} em equipamentos",
            f"{palavra_chave_tag} em sistemas",
            f"{palavra_chave_tag} em processos",
            f"serviços de {palavra_chave_tag}",
            f"assistência em {palavra_chave_tag}",
            f"diagnóstico de {palavra_chave_tag}",
            f"inspeção de {palavra_chave_tag}",
            f"suporte técnico {palavra_chave_tag}",
            f"consultoria em {palavra_chave_tag}"
        ]

    else:

        tags_base_adicionais = [
            f"{palavra_chave_tag} industrial",
            f"{palavra_chave_tag} profissional",
            f"{palavra_chave_tag} técnica",
            f"{palavra_chave_tag} especializada",
            f"{palavra_chave_tag} para indústria",
            f"{palavra_chave_tag} para sistemas",
            f"{palavra_chave_tag} para processos",
            f"{palavra_chave_tag} hidráulica",
            f"{palavra_chave_tag} industrial hidráulica",
            f"aplicações de {palavra_chave_tag}",
            f"uso de {palavra_chave_tag}",
            f"soluções com {palavra_chave_tag}",
            f"equipamento {palavra_chave_tag}",
            f"sistema com {palavra_chave_tag}",
            f"assistência técnica {palavra_chave_tag}"
        ]

    # As tags fixas vêm primeiro e não são substituídas por outras.
    tags_base = tags_base_fixas + tags_base_adicionais

    # ========================================================
    # LIMPAR E VALIDAR TAGS
    # ========================================================

    lista_tags = []

    tema_normalizado = re.sub(
        r"\s+",
        " ",
        str(tema_base).strip()
    ).casefold()

    for tag in tags_base:

        tag = re.sub(
            r"\s+",
            " ",
            str(tag).strip()
        ).casefold()

        if not tag:
            continue

        # ----------------------------------------------------
        # REGRA FUNDAMENTAL:
        # A tag não pode ser exatamente a palavra-chave.
        # ----------------------------------------------------

        if tag.casefold() == tema_normalizado:

            print(
                "⚠️ TAG REJEITADA: "
                "palavra-chave isolada ->",
                tag
            )

            continue

        # ----------------------------------------------------
        # EVITAR DUPLICIDADES
        # ----------------------------------------------------

        if tag.casefold() in [
            item.casefold()
            for item in lista_tags
        ]:

            continue

        lista_tags.append(
            tag
        )

    # ========================================================
    # GARANTIA ESTRUTURAL
    # ========================================================

    lista_tags = remover_duplicados(
        lista_tags
    )

    # ========================================================
    # GARANTIA FINAL:
    # NENHUMA TAG PODE SER O TEMA PURO
    # ========================================================

    lista_tags = [
        tag
        for tag in lista_tags
        if str(tag).strip().casefold()
        != tema_normalizado
    ]

    # --------------------------------------------------------
    # COMPLEMENTO CONTROLADO — SOMENTE PYTHON
    #
    # A palavra-chave pura é proibida, mas isso não pode fazer
    # a página terminar com 29 tags. Quando a base fixa não
    # atingir 30, acrescentamos variações semânticas controladas
    # do próprio tema, sem Ollama e sem tags aleatórias.
    # --------------------------------------------------------

    if len(lista_tags) < 30:

        if tipo == "servico":
            tags_complementares = [
                f"aplicação de {tema_base}",
                f"execução de {tema_base}",
                f"planejamento de {tema_base}",
                f"avaliação de {tema_base}",
                f"controle de {tema_base}",
                f"qualidade em {tema_base}",
                f"soluções técnicas em {tema_base}",
                f"atendimento técnico de {tema_base}",
            ]
        else:
            tags_complementares = [
                f"aplicação de {tema_base}",
                f"seleção de {tema_base}",
                f"dimensionamento de {tema_base}",
                f"instalação de {tema_base}",
                f"operação de {tema_base}",
                f"manutenção de {tema_base}",
                f"desempenho de {tema_base}",
                f"eficiência de {tema_base}",
                f"especificação de {tema_base}",
                f"fornecimento de {tema_base}",
                f"soluções técnicas com {tema_base}",
                f"características de {tema_base}",
            ]

        for tag in tags_complementares:

            tag = re.sub(
                r"\s+",
                " ",
                str(tag).strip()
            )

            if not tag:
                continue

            if tag.casefold() == tema_normalizado:
                continue

            if tag.casefold() in [
                item.casefold()
                for item in lista_tags
            ]:
                continue

            lista_tags.append(tag)

            if len(lista_tags) >= 30:
                break

    # --------------------------------------------------------
    # A quantidade continua sendo obrigatória. Se mesmo a base
    # fixa + complemento controlado não alcançar 30, abortamos.
    # --------------------------------------------------------

    if len(lista_tags) < 30:

        print()
        print(
            "❌ ERRO: quantidade insuficiente de "
            "tags válidas."
        )

        print(
            "TAGS VÁLIDAS:",
            len(lista_tags)
        )

        print(
            "MÍNIMO NECESSÁRIO:",
            30
        )

        return None

    lista_tags = lista_tags[:30]

    # ========================================================
    # VALIDAÇÃO ABSOLUTA
    # ========================================================

    for tag in lista_tags:

        if str(tag).strip().casefold() == tema_normalizado:

            print()
            print(
                "❌ ERRO GRAVE: palavra-chave isolada "
                "detectada nas tags:"
            )

            print(tag)

            return None

    # ========================================================
    # LOG
    # ========================================================

    print()
    print("======================================")
    print("SEGMENTOS E TAGS CRIADOS PELO PYTHON")
    print("======================================")

    print(
        "TIPO:",
        "SERVIÇO" if tipo == "servico" else "PRODUTO"
    )

    print(
        "SEGMENTOS:",
        len(lista_segmentos),
        "/",
        total_segmentos
    )

    print(
        "TAGS:",
        len(lista_tags),
        "/",
        total_tags
    )

    print()
    print("TAGS FINAIS:")

    for numero, tag in enumerate(
        lista_tags,
        start=1
    ):

        print(
            f"TAG_{numero}:",
            tag
        )




    
    # ========================================================
    # 20. VALIDAR ESTRUTURA REAL
    # ========================================================
    
    # ============================================================
    # RECUPERAÇÃO DO SUBTÍTULO — NÃO DESCARTAR PÁGINA COMPLETA
    # ============================================================
    # O subtítulo pode ser omitido pelo modelo. Recupera a informação a
    # partir dos títulos já gerados, sem chamada adicional nem fatos novos.
    if not str(subtitulo or "").strip():
        subtitulo = "Informações e critérios relacionados ao tema"
        print("⚠️ SUBTÍTULO AUSENTE — APLICADO SUBTÍTULO CURTO E INDEPENDENTE.")

    total_blocos_real = len(
        blocos
    )
    
    total_paragrafos_real = sum(
        len(
            bloco.get(
                "paragrafos_ollama",
                []
            )
        )
        for bloco in blocos
    )
    
    total_segmentos_real = len(
        lista_segmentos
    )
    
    total_tags_real = len(
        lista_tags
    )
    
    print()
    print("======================================")
    print("VALIDAÇÃO REAL DA ESTRUTURA")
    print("======================================")
    
    print(
        "TÍTULO:",
        "OK" if titulo else "FALTANDO"
    )
    
    print(
        "SUBTÍTULO:",
        "OK" if subtitulo else "FALTANDO"
    )
    
    print(
        "BLOCOS:",
        total_blocos_real,
        "/",
        total_blocos
    )
    
    print(
        "PARÁGRAFOS:",
        total_paragrafos_real,
        "/",
        total_paragrafos
    )
    
    print(
        "SEGMENTOS:",
        total_segmentos_real,
        "/",
        total_segmentos
    )
    
    print(
        "TAGS:",
        total_tags_real,
        "/",
        total_tags
    )
    
    # Limites oficiais fixos da página MEAD. Evita que None recebido de
    # uma etapa auxiliar reprove uma estrutura que contém 12 segmentos/30 tags.
    total_blocos = 5 if not isinstance(total_blocos, int) else total_blocos
    paragrafos_por_bloco = 3 if not isinstance(paragrafos_por_bloco, int) else paragrafos_por_bloco
    total_paragrafos = 15 if not isinstance(total_paragrafos, int) else total_paragrafos
    total_segmentos = 12 if not isinstance(total_segmentos, int) or total_segmentos <= 0 else total_segmentos
    total_tags = 30 if not isinstance(total_tags, int) or total_tags <= 0 else total_tags

    estrutura_valida = True
    
    if not titulo:
        estrutura_valida = False
    
    if not subtitulo:
        estrutura_valida = False
    
    if total_blocos_real != total_blocos:
        estrutura_valida = False
    
    for bloco in blocos:
    
        if len(
            bloco.get(
                "paragrafos_ollama",
                []
            )
        ) != paragrafos_por_bloco:
    
            estrutura_valida = False
    
    if total_paragrafos_real != total_paragrafos:
        estrutura_valida = False
    
    if total_segmentos_real != total_segmentos:
        estrutura_valida = False
    
    if total_tags_real != total_tags:
        estrutura_valida = False
    
    if not estrutura_valida:
    
        print()
        print("======================================")
        print("ESTRUTURA INVÁLIDA")
        print("======================================")
    
        print(
            "A página NÃO será salva como estrutura final."
        )
    
        print(
            "O conteúdo bruto será descartado para evitar"
        )
    
        print(
            "gravar uma estrutura falsa no banco."
        )
    
        return None
        
        
        
    # ============================================================
    # CAPITALIZAÇÃO EDITORIAL
    # ============================================================
    # Aplicada SOMENTE em:
    # - subtitulo
    # - subtitulo_segmentos
    # - segmentos_listas
    #
    # NÃO aplicar em:
    # - tema
    # - tags
    # - H1
    # - titulo
    # - parágrafos
    # - fontes
    # - referências
    # ============================================================
    
    PALAVRAS_EDITORIAL_MINUSCULAS = {
        "a", "as", "o", "os",
        "de", "da", "das", "do", "dos",
        "em", "na", "nas", "no", "nos",
        "para", "por",
        "e",
        "com", "sem",
        "sobre", "entre",
        "ao", "aos",
        "à", "às",
        "num", "numa",
        "nuns", "numas",
        "dum", "duma",
        "duns", "dumas",
        "até",
        "como"
    }
    
    
    def capitalizar_texto_editorial(texto):
    
        texto = str(
            texto or ""
        ).strip()
    
        if not texto:
            return ""
    
        palavras = texto.split()
    
        resultado = []
    
        for indice, palavra in enumerate(palavras):
    
            if not palavra:
                continue
    
            palavra_normalizada = palavra.casefold()
    
            if (
                indice > 0
                and palavra_normalizada
                in PALAVRAS_EDITORIAL_MINUSCULAS
            ):
    
                resultado.append(
                    palavra_normalizada
                )
    
            else:
    
                resultado.append(
                    palavra[:1].upper()
                    +
                    palavra[1:].lower()
                )
    
        return " ".join(
            resultado
        )    
        

    # ========================================================
    # CAPITALIZAÇÃO EDITORIAL FINAL
    # ========================================================
    
    subtitulo = capitalizar_texto_editorial(
        subtitulo
    )
    
    subtitulo_segmentos = capitalizar_texto_editorial(
        subtitulo_segmentos
    )
    
    lista_segmentos = [
        capitalizar_texto_editorial(
            segmento
        )
        for segmento in lista_segmentos
    ]

    for bloco in blocos:

        if isinstance(bloco, dict):

            bloco["titulo"] = capitalizar_texto_editorial(
                bloco.get("titulo", "")
            )
    

        
    # ========================================================
    # 21. MONTAR CONTEÚDO FINAL
    # ========================================================
    
    partes_conteudo = []
    
    partes_conteudo.append(
        "TÍTULO PRINCIPAL: "
        + titulo
    )
    
    partes_conteudo.append(
        "SUBTÍTULO: "
        + subtitulo
    )
    
    partes_conteudo.append(
        "### Conteúdo"
    )
    
    for bloco in blocos:
    
        partes_conteudo.append(
            "#### Bloco "
            + str(
                bloco["numero"]
            )
            + ": "
            + bloco["titulo"]
        )
    
        for paragrafo in bloco[
            "paragrafos_ollama"
        ]:
    
            partes_conteudo.append(
                paragrafo
            )
    
    partes_conteudo.append(
        "### Segmentos"
    )
    
    for indice, segmento in enumerate(
        lista_segmentos,
        start=1
    ):
    
        partes_conteudo.append(
            f"{indice}. {segmento}"
        )
    
    partes_conteudo.append(
        "### Tags"
    )
    
    for indice, tag in enumerate(
        lista_tags,
        start=1
    ):
    
        partes_conteudo.append(
            f"{indice}. {tag}"
        )
    
    conteudo = "\n\n".join(
        partes_conteudo
    ).strip()


    # ========================================================
    # 22. ESTRUTURA REAL
    # ========================================================

    estrutura_real = {

        "blocos":
            total_blocos_real,

        "paragrafos_por_bloco":
            paragrafos_por_bloco,

        "total_paragrafos":
            total_paragrafos_real,

        "segmentos":
            total_segmentos_real,

        "tags":
            total_tags_real,

        "independente":
            True
    }


    # ========================================================
    # TESTE — CONFERIR O RESULTADO REAL DO OLLAMA
    # ========================================================
    
    print()
    print("======================================")
    print("TESTE: BLOCOS ANTES DO SALVAR_BANCO")
    print("ESTRUTURA PRONTA PARA GRAVAÇÃO:",
          len(blocos), "blocos |", total_paragrafos_real, "parágrafos |",
          total_segmentos_real, "segmentos |", total_tags_real, "tags")

    # ========================================================
    # 23. ENVIAR ESTRUTURA OFICIAL PARA SALVAR_BANCO
    # ========================================================
    
    print()
    print("======================================")
    print("SALVANDO ESTRUTURA OFICIAL NO BANCO")
    print("PROGRESSO DA PÁGINA: 95% — gravando conteudo.json e conferindo leitura.")
    print("======================================")
    
    resultado_salvamento = salvar_banco(
        tema,
        "pagina",
        conteudo,
    
        informacoes_adicionais={
    
            "nome_site":
                nome_site,
    
            "grupo_principal_projeto":
                normalizar_grupo_principal_projeto(
                    entrada_grupo.get()
                    if "entrada_grupo" in globals()
                    else ""
                ),
    
            "segmentos_textuais":
                lista_segmentos,
    
            "referencias":
                list(
                    dict.fromkeys(
                        [
                            str(
                                item.get(
                                    "url",
                                    ""
                                )
                            ).strip()
    
                            for item in dados_coleta
    
                            if isinstance(
                                item,
                                dict
                            )
    
                            and str(
                                item.get(
                                    "url",
                                    ""
                                )
                            ).strip()
                        ]
                    )
                ),
    
            "arquivo_origem":
                nome_arquivo
        },
    
        blocos=blocos,
    
        segmentos=lista_segmentos,
    
        tags=lista_tags,
    
        grupo_principal_projeto=
            normalizar_grupo_principal_projeto(
                entrada_grupo.get()
                if "entrada_grupo" in globals()
                else ""
            ),
    
        tipo=tipo,
        
        h1=h1,
        
        titulo=titulo,
        
        subtitulo=subtitulo,
        
        subtitulo_segmentos=subtitulo_segmentos,
    )


    if resultado_salvamento is not True:
        print("❌ FALHA AO SALVAR A PÁGINA NO BANCO.")
        return None

    # A cópia de recuperação só é removida depois que salvar_banco
    # confirmou a gravação e a leitura de conferência no disco.
    limpar_blocos_recuperacao(tema)

    # ========================================================
    # 24. SALVAR MAPA MEAD
    # ========================================================

    salvar_banco(
        tema,
        "mapa_mead",
        mapa_texto
    )

    # ========================================================
    # 25. RESULTADO
    # ========================================================

    print()
    print("======================================")
    print("PÁGINA GERADA E SALVA")
    print("PROGRESSO DA PÁGINA: 100% — gravação oficial confirmada.")
    print("======================================")

    print(
        "TEMA:",
        tema
    )

    print(
        "ARQUIVO:",
        nome_arquivo
    )

    print(
        "CARACTERES:",
        len(conteudo)
    )

    print(
        "BLOCOS:",
        total_blocos_real
    )

    print(
        "PARÁGRAFOS:",
        total_paragrafos_real
    )

    print(
        "SEGMENTOS:",
        total_segmentos_real
    )

    print(
        "TAGS:",
        total_tags_real
    )

    tempo_total = time.time() - inicio_geracao

    print(
        "TEMPO:",
        round(
            tempo_total,
            1
        ),
        "segundos"
    )

    return conteudo



# ============================================================
# LIMPAR REFERÊNCIAS COMERCIAIS
# ============================================================

def limpar_referencias_comerciais(texto, tema):

    if not texto:
        return texto


    tema_lower = tema.lower()


    # =====================================
    # REMOVER MARCAS SOMENTE PARA SELAGEM
    # CORTA FOGO
    # =====================================

    if any(
        palavra in tema_lower
        for palavra in [

            "firestop",
            "selagem",
            "corta fogo",
            "passagem corta fogo"

        ]
    ):

        remover = [

            "CP 636",
            "CKC",
            "Firestop",
            "Hilti",
            "3M",
            "Promat",
            "Noneifire",
            "Promaseal",
            "Tecbor",
            "FFC",
            "FCR"

        ]


        for item in remover:

            texto = re.sub(
                item,
                "",
                texto,
                flags=re.IGNORECASE
            )



    # ========================================================
    # 01. NORMALIZA TEXTO
    # ========================================================

    texto = texto.replace(
        "corta-fogo",
        "corta fogo"
    )


    texto = texto.replace(
        "Corta-Fogo",
        "Corta Fogo"
    )


    texto = re.sub(
        r"\s+",
        " ",
        texto
    )


    return texto.strip()
    
    # ========================================================
    # 02. AUDITORIA DE NATURALIDADE
    # ========================================================

    print()
    print("==============================")
    print("INICIANDO AUDITORIA DE NATURALIDADE")
    print("==============================")

    conteudo = auditar_e_corrigir_aberturas(
        conteudo,
        tema
    )

    print()
    print("==============================")
    print("AUDITORIA DE NATURALIDADE FINALIZADA")
    print("==============================")    

# ============================================================
# IDENTIFICAR IDENTIDADE DA FONTE
# ============================================================

def identificar_empresa_fonte(
    texto,
    url
):

    """
    Identifica de forma conservadora uma empresa ou marca
    explicitamente mencionada no conteúdo da fonte.

    REGRAS:

    - Não inventa nomes.
    - Não considera automaticamente o domínio como empresa.
    - O domínio é preservado apenas como informação da fonte.
    - Prioriza identificações explícitas no texto.
    - Aceita construções como:
        "fabricado pela Empresa X"
        "produzido pela Empresa X"
        "distribuído pela Empresa X"
        "fornecedor: Empresa X"
        "fabricante: Empresa X"
        "razão social: Empresa X"
        "A Pumps Brasil é..."
        "Pumps Brasil é..."
        "As bombas Sanitárias BOMBINOX são..."
        "BOMBINOX, uma empresa..."

    Retorno:

    {
        "nome": "",
        "dominio": "",
        "papel": "",
        "origem_identificacao": "",
        "confianca": "baixa"
    }
    """

    texto = str(
        texto or ""
    ).strip()

    url = str(
        url or ""
    ).strip()

    resultado = {

        "nome": "",
        "dominio": "",
        "papel": "",
        "origem_identificacao": "",
        "confianca": "baixa"

    }

    if not texto and not url:

        return resultado

    # ========================================================
    # 01. IDENTIFICAR DOMÍNIO
    # ========================================================

    dominio = ""

    try:

        from urllib.parse import urlparse

        dominio = (
            urlparse(url)
            .netloc
            .lower()
            .strip()
        )

        if dominio.startswith("www."):

            dominio = dominio[4:]

    except Exception:

        dominio = ""

    resultado[
        "dominio"
    ] = dominio

    # ========================================================
    # 02. NORMALIZAR TEXTO
    # ========================================================

    texto_analise = re.sub(
        r"\s+",
        " ",
        texto
    ).strip()

    if not texto_analise:

        return resultado

    # ========================================================
    # 03. IDENTIFICAÇÕES EXPLÍCITAS
    # ========================================================
    #
    # São os casos de maior confiança.
    # ========================================================

    padroes_explicitos = [

        # ----------------------------------------------------
        # fabricado por / fabricado pela
        # ----------------------------------------------------

        (
            r"(?:fabricado|fabricada|fabricados|fabricadas)"
            r"\s+(?:pela|pelo|por)\s+"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\- ]{2,100}?)"
            r"(?=\s+(?:e|é|são|possui|atua|oferece|"
            r"fabrica|produz|fornece|distribui)|[.,;:])",

            "fabricante"
        ),

        # ----------------------------------------------------
        # produzido por / produzido pela
        # ----------------------------------------------------

        (
            r"(?:produzido|produzida|produzidos|produzidas)"
            r"\s+(?:pela|pelo|por)\s+"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\- ]{2,100}?)"
            r"(?=\s+(?:e|é|são|possui|atua|oferece|"
            r"fabrica|produz|fornece|distribui)|[.,;:])",

            "fabricante"
        ),

        # ----------------------------------------------------
        # distribuído por / distribuído pela
        # ----------------------------------------------------

        (
            r"(?:distribuído|distribuída|distribuídos|distribuídas)"
            r"\s+(?:pela|pelo|por)\s+"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\- ]{2,100}?)"
            r"(?=\s+(?:e|é|são|possui|atua|oferece|"
            r"fabrica|produz|fornece|distribui)|[.,;:])",

            "distribuidor"
        ),

        # ----------------------------------------------------
        # fornecido por / fornecido pela
        # ----------------------------------------------------

        (
            r"(?:fornecido|fornecida|fornecidos|fornecidas)"
            r"\s+(?:pela|pelo|por)\s+"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\- ]{2,100}?)"
            r"(?=\s+(?:e|é|são|possui|atua|oferece|"
            r"fabrica|produz|fornece|distribui)|[.,;:])",

            "fornecedor"
        ),

        # ----------------------------------------------------
        # fabricante / fornecedor / empresa: X
        # ----------------------------------------------------

        (
            r"(?:empresa|fabricante|fornecedor)"
            r"\s*:\s*"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\- ]{2,100}?)"
            r"(?=\s+(?:e|é|são|atua|oferece|fabrica|produz|"
            r"fornece|distribui)|[.,;:])",

            "empresa"
        ),

        # ----------------------------------------------------
        # razão social: X
        # ----------------------------------------------------

        (
            r"razão\s+social"
            r"\s*:\s*"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\- ]{2,120}?)"
            r"(?=\s+(?:e|é|são|atua|oferece)|[.,;:])",

            "empresa"
        )

    ]

    # ========================================================
    # 04. EXECUTAR IDENTIFICAÇÕES EXPLÍCITAS
    # ========================================================

    for padrao, papel in padroes_explicitos:

        try:

            correspondencia = re.search(
                padrao,
                texto_analise,
                flags=re.IGNORECASE
            )

        except Exception:

            correspondencia = None

        if not correspondencia:

            continue

        nome = (
            correspondencia.group(1)
            .strip()
            .strip(" .,;:-")
        )

        if not nome:

            continue

        if len(nome) > 120:

            continue

        resultado[
            "nome"
        ] = nome

        resultado[
            "papel"
        ] = papel

        resultado[
            "origem_identificacao"
        ] = "texto"

        resultado[
            "confianca"
        ] = "alta"

        return resultado

    # ========================================================
    # 05. EMPRESA APRESENTADA COMO SUJEITO
    # ========================================================
    #
    # Exemplos:
    #
    #   "A Pumps Brasil é capacitada..."
    #   "A Pumps Brasil é especializada..."
    #   "Pumps Brasil é fabricante..."
    #
    # O nome precisa estar associado a uma construção
    # característica de apresentação empresarial.
    # ========================================================

    padroes_apresentacao = [

        # ----------------------------------------------------
        # "a Pumps Brasil é..."
        # "o Grupo X oferece..."
        # ----------------------------------------------------

        (
            r"\b(?:a|o)\s+"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]+"
            r"(?:\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]+){0,5})"
            r"\s*,?\s+"
            r"(?:é|são|atua|oferece|fornece|fabrica|produz|"
            r"distribui|comercializa|especializada|especializado)"
            r"\b"
        ),

        # ----------------------------------------------------
        # "Pumps Brasil é..."
        # ----------------------------------------------------

        (
            r"\b"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]+"
            r"(?:\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]+){0,5})"
            r"\s+"
            r"(?:é|são|atua|oferece|fornece|fabrica|produz|"
            r"distribui|comercializa|especializada|especializado)"
            r"\b"
        )

    ]

    palavras_invalidas = {

        "empresa",
        "fabricante",
        "fornecedor",
        "produto",
        "produtos",
        "bomba",
        "bombas",
        "equipamento",
        "equipamentos",
        "sistema",
        "sistemas",
        "indústria",
        "industria",
        "indústrias",
        "industrias",
        "mercado",
        "setor",
        "segmento",
        "solução",
        "solucao",
        "soluções",
        "solucoes"

    }

    for padrao in padroes_apresentacao:

        try:

            correspondencia = re.search(
                padrao,
                texto_analise
            )

        except Exception:

            correspondencia = None

        if not correspondencia:

            continue

        nome = (
            correspondencia.group(1)
            .strip()
            .strip(" .,;:-")
        )

        if not nome:

            continue

        if len(nome) > 100:

            continue

        if nome.casefold() in palavras_invalidas:

            continue

        resultado[
            "nome"
        ] = nome

        resultado[
            "papel"
        ] = "empresa"
        
        resultado[
            "origem_identificacao"
        ] = "texto"

        resultado[
            "confianca"
        ] = "alta"

        return resultado

    # ========================================================
    # 06. MARCA EM CAIXA ALTA
    # ========================================================
    #
    # Exemplos:
    #
    #   "As bombas Sanitárias BOMBINOX são..."
    #   "BOMBINOX são..."
    #   "BOMBINOX, uma empresa..."
    #
    # Não considerar qualquer palavra em caixa alta
    # isoladamente como empresa.
    # ========================================================

    padroes_marca = [

        # ----------------------------------------------------
        # "BOMBINOX são..."
        # ----------------------------------------------------

        (
            r"\b"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ]{3,}"
            r"[A-Z0-9ÁÀÃÂÉÊÍÓÔÕÚÇ&.\-]*)"
            r"\s+"
            r"(?:é|são|atua|oferece|fabrica|produz|"
            r"fornece|distribui|comercializa)"
            r"\b"
        ),

        # ----------------------------------------------------
        # "BOMBINOX, uma empresa..."
        # ----------------------------------------------------

        (
            r"\b"
            r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ]{3,}"
            r"[A-Z0-9ÁÀÃÂÉÊÍÓÔÕÚÇ&.\-]*)"
            r"\s*,?\s+"
            r"(?:uma|um)\s+"
            r"(?:empresa|fabricante|fornecedor|marca)"
            r"\b"
        )

    ]

    for padrao in padroes_marca:

        try:

            correspondencia = re.search(
                padrao,
                texto_analise
            )

        except Exception:

            correspondencia = None

        if not correspondencia:

            continue

        nome = (
            correspondencia.group(1)
            .strip()
            .strip(" .,;:-")
        )

        if not nome:

            continue

        if len(nome) < 3:

            continue

        resultado[
            "nome"
        ] = nome

        resultado[
            "papel"
        ] = "empresa"

        resultado[
            "origem_identificacao"
        ] = "texto"

        resultado[
            "confianca"
        ] = "media"

        return resultado

    # ========================================================
    # 07. NENHUMA EMPRESA IDENTIFICADA
    # ========================================================
    #
    # Não inventar nada.
    # O domínio permanece preservado.
    # ========================================================

    return resultado
    

# ============================================================
# LIMPAR LISTA DE REFERÊNCIAS
# ============================================================

def limpar_lista_referencias(
    textos,
    tema
):

    textos_limpos = []


    for item in textos:


        if not item:
            continue



    # ========================================================
    # 01. NORMALIZAR FORMATO
    # ========================================================

        if isinstance(item, dict):

            texto = item.get(
                "texto",
                ""
            )

            url = item.get(
                "url",
                ""
            )

            tipo = item.get(
                "tipo",
                "texto"
            )


        else:

            texto = str(item)

            url = ""

            tipo = "texto"



        if not texto:

            continue



        texto_upper = texto.upper()



    # ========================================================
    # 02. REMOVER MAPAS MEAD ANTIGOS
    # ========================================================

        if "### MAPA_MEAD" in texto_upper:

            print(
                "MAPA MEAD ANTIGO REMOVIDO"
            )

            continue



        if "PROTAGONISTA:" in texto_upper:

            print(
                "REFERÊNCIA COM PROTAGONISTA REMOVIDA"
            )

            continue



    # ========================================================
    # 03. LIMPEZA COMERCIAL
    # ========================================================

        texto = limpar_referencias_comerciais(
            texto,
            tema
        )
        
        
    # ========================================================
    # 04. LIMPEZA BÁSICA PARA IA
    # ========================================================
        
        texto = texto.strip()
        
        texto = " ".join(
            texto.split()
        )



        if len(texto) < 300:

            print(
                "TEXTO DESCARTADO - PEQUENO:",
                len(texto)
            )

            continue



    # ========================================================
    # 05. IDENTIFICAR IDENTIDADE DA FONTE
    # ========================================================

        identidade_fonte = identificar_empresa_fonte(
            texto,
            url
        )


    # ========================================================
    # 06. SALVAR TEXTO LIMPO
    # ========================================================

        textos_limpos.append(
            {
                "url": url,
                "tipo": tipo,
                "texto": texto,
                "identidade_fonte": identidade_fonte
            }
        )



    print()
    print("==============================")
    print("TEXTOS DISPONÍVEIS")
    print("==============================")
    print(len(textos))



    print()
    print("==============================")
    print("TEXTOS APÓS LIMPEZA")
    print("==============================")
    print(len(textos_limpos))



    for i, item in enumerate(textos_limpos):

        print(
            f"{i+1}: {len(item.get('texto',''))} caracteres | "
            f"{item.get('url','')[:80]}"
        )



    return textos_limpos
    

# ============================================================
# EXTRAIR MAPA MEAD DA RESPOSTA
# ============================================================

def extrair_mapa_mead(conteudo):

    if not conteudo:
        return ""

    inicio = conteudo.find("### MAPA_MEAD")

    if inicio == -1:
        return ""

    secoes = [
        "### DEFINICAO",
        "### BENEFICIOS",
        "### VANTAGENS",
        "### MATERIA_PRIMA",
        "### APLICACOES",
        "### FABRICACAO",
        "### MANUTENCAO",
        "### ATIVOS_NARRATIVOS",
        "### DUVIDAS_FREQUENTES"
    ]

    fim = len(conteudo)

    for secao in secoes:

        pos = conteudo.find(secao)

        if pos > inicio:
            fim = pos
            break

    return conteudo[inicio:fim].strip()
    
    
    

# ============================================================
# VALIDAR MAPA MEAD
# ============================================================

def validar_mapa_mead(texto):

    if not texto:
        return False


    texto_upper = texto.upper()


    # remover acentos
    texto_upper = (
        texto_upper
        .replace("Á","A")
        .replace("À","A")
        .replace("Ã","A")
        .replace("Â","A")
        .replace("É","E")
        .replace("Ê","E")
        .replace("Í","I")
        .replace("Ó","O")
        .replace("Ô","O")
        .replace("Õ","O")
        .replace("Ú","U")
        .replace("Ç","C")
    )


    obrigatorios = [

        "PROTAGONISTA",
        "CENARIO",
        "PROBLEMA",
        "SOLUCAO"

    ]


    for item in obrigatorios:

        if item not in texto_upper:

            print()
            print("==============================")
            print("MAPA MEAD FALTANDO:")
            print(item)
            print("==============================")

            return False


    return True
    
    

# ============================================================
# LER BANCO DE DADOS
# ============================================================

def carregar_banco():

    try:

        if not os.path.exists(
            ARQUIVO_BANCO
        ):

            return {}

        with open(
            ARQUIVO_BANCO,
            "r",
            encoding="utf-8"
        ) as f:

            return json.load(f)

    except Exception as e:

        print()
        print("ERRO AO LER BANCO:")
        print(e)

        return {}




# ============================================================
# BUSCAR MAPA MEAD DO PRÓPRIO TEMA
# ============================================================

def obter_mapa_mead_tema(tema):


    banco = carregar_banco()


    tema_busca = remover_acentos(
        tema.lower()
    )

    tema_busca = " ".join(
        tema_busca.split()
    )


    for chave, dados in banco.items():


        chave_normalizada = remover_acentos(
            chave.lower()
        )

        chave_normalizada = " ".join(
            chave_normalizada.split()
        )


        if chave_normalizada == tema_busca:


            print()
            print("==============================")
            print("MAPA ENCONTRADO NO BANCO")
            print("==============================")
            print(chave)


            if not isinstance(
                dados,
                dict
            ):

                return ""


            mapa = dados.get(
                "mapa_mead",
                {}
            )


            if not isinstance(
                mapa,
                dict
            ):

                return ""


            texto = mapa.get(
                "texto",
                ""
            )


            if not texto:

                return ""


            if not validar_protagonista_mead(
                texto,
                tema
            ):

                print()
                print("==============================")
                print("MAPA MEAD INVÁLIDO")
                print("==============================")

                return ""


            return texto



    print()
    print("==============================")
    print("MAPA NÃO ENCONTRADO")
    print("==============================")

    return ""
    

# ============================================================
# VERIFICAR TEMA NO BANCO
# ============================================================

def verificar_tema_banco(tema):

    banco = carregar_banco()


    if tema in banco:

        dados = banco[tema]


        resultado = {

            "existe": True,

            "mapa_mead": "",

            "conteudo": ""

        }


        mapa = dados.get(
            "mapa_mead",
            {}
        )


        if isinstance(
            mapa,
            dict
        ):

            resultado["mapa_mead"] = mapa.get(
                "texto",
                ""
            )


        categorias = dados.get(
            "categorias",
            {}
        )


        if isinstance(
            categorias,
            dict
        ):

            resultado["conteudo"] = categorias.get(
                "conteudo_completo",
                ""
            )


        return resultado



    return {

        "existe": False,

        "mapa_mead": "",

        "conteudo": ""

    }    


# ============================================================
# ALIMENTAR A IA COM O MAPA MEAD
# ============================================================

def obter_textos_banco():

    banco = carregar_banco()

    textos = []


    for tema, dados in banco.items():


        if not isinstance(
            dados,
            dict
        ):
            continue



    # ========================================================
    # 01. CARREGAR MAPA MEAD
    # ========================================================

        mapa = dados.get(
            "mapa_mead",
            {}
        )


        if isinstance(
            mapa,
            dict
        ):

            texto_mapa = mapa.get(
                "texto",
                ""
            )


            if texto_mapa:

                textos.append(
                    texto_mapa
                )



    # ========================================================
    # 02. CARREGAR CATEGORIAS
    # ========================================================

        categorias = dados.get(
            "categorias",
            {}
        )


        if isinstance(
            categorias,
            dict
        ):


            for categoria, conteudo in categorias.items():


                if isinstance(
                    conteudo,
                    str
                ):

                    textos.append(
                        conteudo
                    )



    return textos


def obter_contexto_banco(
    limite=5000
):

    textos = obter_textos_banco()

    contexto = "\n\n".join(
        item.get("texto", "")
        for item in textos
        if isinstance(item, dict)
    )
    
    print()
    print("==============================")
    print("DEBUG CONTEXTO MAPA")
    print("==============================")
    
    print(
        "ITENS RECEBIDOS:",
        len(textos)
    )
    
    print(
        "CARACTERES GERADOS:",
        len(contexto)
    )
    
    print(
        contexto[:500]
    )

    return contexto[:limite]
    
    
    
# ============================================================
# NORMALIZAR IDENTIDADE DO TEMA
# ============================================================

def normalizar_tema_chave(tema):

    """
    Cria uma identidade técnica normalizada do tema.

    IMPORTANTE:
    - NÃO altera a acentuação do tema.
    - Preserva exatamente os caracteres originais.
    - Converte para minúsculas apenas para comparação.
    - Remove espaços duplicados.
    - Deve ser usada apenas para identificação,
      comparação e controle de duplicidade.

    A remoção de acentos NÃO deve ocorrer aqui.
    Ela deve existir somente na rotina responsável
    pela criação do nome do arquivo/slug.
    """

    if tema is None:
        return ""

    texto = str(
        tema
    ).strip()

    if not texto:
        return ""

    texto = " ".join(
        texto.split()
    )

    texto = texto.casefold()

    return texto



# ============================================================
# VALIDAR ESTRUTURA DA PÁGINA
# ============================================================

def validar_estrutura_pagina(
    conteudo
):

    print()
    print("==============================")
    print("VALIDANDO ESTRUTURA DA PÁGINA")
    print("==============================")

    texto = str(
        conteudo or ""
    ).strip()

    if not texto:

        print(
            "CONTEÚDO VAZIO"
        )

        return False

    # ========================================================
    # 01. CONTAR PARÁGRAFOS
    # ========================================================

    paragrafos = []

    for bloco in texto.split("\n\n"):

        bloco = bloco.strip()

        if not bloco:
            continue

        # Ignorar títulos de bloco
        if bloco.startswith("#"):
            continue

        if bloco.startswith("BLOCO"):
            continue

        if bloco.startswith("Segmentos"):
            continue

        paragrafos.append(
            bloco
        )

    # ========================================================
    # 02. CONTAR SEGMENTOS
    # ========================================================

    segmentos = []

    encontrou_segmentos = False

    for linha in texto.splitlines():

        linha_limpa = linha.strip()

        if (
            "segmentos atendidos" in
            linha_limpa.lower()
            or
            linha_limpa.lower() == "segmentos"
        ):

            encontrou_segmentos = True
            continue

        if encontrou_segmentos:

            if not linha_limpa:
                continue

            if (
                linha_limpa.startswith("-")
                or
                linha_limpa.startswith("•")
                or
                linha_limpa[:2].isdigit()
            ):

                segmentos.append(
                    linha_limpa
                )

    # ========================================================
    # 03. RESULTADO
    # ========================================================

    total_paragrafos = len(
        paragrafos
    )

    total_segmentos = len(
        segmentos
    )

    print(
        "PARÁGRAFOS ENCONTRADOS:",
        total_paragrafos
    )

    print(
        "SEGMENTOS ENCONTRADOS:",
        total_segmentos
    )

    # ========================================================
    # 04. REGRA OBRIGATÓRIA
    # ========================================================

    if total_paragrafos < 15:

        print()
        print(
            "❌ PÁGINA REPROVADA"
        )

        print(
            "Motivo: menos de 15 parágrafos."
        )

        return False

    print()
    print(
        "✅ ESTRUTURA APROVADA"
    )

    print(
        "15 PARÁGRAFOS; SEGMENTOS SOMENTE QUANDO SUSTENTADOS"
    )

    return True



# ========================================================
# 05. GERAR TÍTULOS
# ========================================================

def gerar_titulos(
    tema,
    blocos_informacoes,
    mapa_mead
):

    print()
    print("==============================")
    print("GERANDO TÍTULOS PELO PYTHON")
    print("==============================")

    tema = str(
        tema or ""
    ).strip()

    if not tema:
        return {}


    # ========================================================
    # 01. H1 / TÍTULO
    # ========================================================
    
    # Formata o tema para exibição, preservando a acentuação
    # original e colocando a inicial de cada palavra em maiúscula.
    tema_titulo = " ".join(
        palavra[:1].upper() + palavra[1:]
        for palavra in tema.split()
    )
    
    h1 = tema_titulo
    title = tema_titulo



    # ========================================================
    # 03. SUBTÍTULO DE IMPACTO
    #
    # Criado exclusivamente pelo Python.
    # O Ollama NÃO participa.
    # ========================================================

    subtitulo = (
        f"Soluções técnicas para aplicações, "
        f"desempenho e confiabilidade em {tema_titulo}"
    )

    # ========================================================
    # 04. SUBTÍTULO DOS SEGMENTOS
    #
    # Segundo subtítulo.
    # Também criado exclusivamente pelo Python.
    # ========================================================

    subtitulo_segmentos = gerar_subtitulo_segmentos(
        tema_titulo
    )

    # ========================================================
    # 05. TÍTULOS DOS 5 BLOCOS
    # ========================================================

    bloco_1 = (
        f"O que é {tema_titulo} e como funciona"
    )

    bloco_2 = (
        f"Características e aplicações de {tema_titulo}"
    )

    bloco_3 = (
        f"Vantagens e diferenciais de {tema_titulo}"
    )

    bloco_4 = (
        f"Aspectos técnicos e cuidados com {tema_titulo}"
    )

    bloco_5 = (
        f"Escolha, fornecimento e suporte para {tema_titulo}"
    )

    # ========================================================
    # 06. RESULTADO
    # ========================================================

    resultado = {

        "h1":
            h1,

        "title":
            title,

        "subtitulo":
            subtitulo,

        "subtitulo_segmentos":
            subtitulo_segmentos,

        "bloco_1":
            bloco_1,

        "bloco_2":
            bloco_2,

        "bloco_3":
            bloco_3,

        "bloco_4":
            bloco_4,

        "bloco_5":
            bloco_5
    }

    print()
    print("==============================")
    print("TÍTULOS GERADOS PELO PYTHON")
    print("==============================")

    print(
        "H1:",
        h1
    )

    print(
        "TITLE:",
        title
    )

    print(
        "SUBTÍTULO IMPACTO:",
        subtitulo
    )

    print(
        "SUBTÍTULO SEGMENTOS:",
        subtitulo_segmentos
    )

    for numero in range(1, 6):

        print(
            f"BLOCO {numero}:",
            resultado.get(
                f"bloco_{numero}",
                ""
            )
        )

    return resultado

    
# ============================================================
# SALVAR BANCO
# ============================================================

def salvar_banco(
    tema,
    categoria,
    texto,
    informacoes_adicionais=None,
    categorias=None,
    blocos=None,
    segmentos=None,
    tags=None,
    trechos_utilizados=None,
    grupo_principal_projeto=None,
    tipo=None,
    h1=None,
    titulo=None,
    subtitulo=None,
    subtitulo_segmentos=None
):

    """
    Persiste a página no formato oficial do conteudo.json.

    REGRAS:

    - Uma única estrutura oficial por tema.
    - Uma nova execução do mesmo tema substitui a versão anterior.
    - Os cinco blocos são preservados entre chamadas.
    - Python mantém os trechos selecionados em
      informacoes_relevantes e paragrafos_python.
    - Ollama mantém o texto editorial em paragrafos_ollama.
    - Não existe mais o campo legado "paragrafos".
    - Não existe mais imagens no JSON oficial.
    - Não existe mais caracteres no JSON oficial.
    - Não existe mais status no JSON oficial.
    - Não existe mais subtitulo_listas.
    - informacoes_relevantes nunca é criada no nível global.
    - Cada bloco possui:
        id
        hash
        informacoes_relevantes
        titulo
        paragrafos_python
        paragrafos_ollama
    - Mantém:
        5 blocos
        15 posições de parágrafos
        segmentos_listas
        posicionamento_listas
        até 30 tags
    """

    global PAGINAS_EM_PROCESSAMENTO

    # ========================================================
    # PRESERVAR DIFERENÇA ENTRE:
    #
    # None  = dado não enviado nesta chamada
    # valor = dado efetivamente enviado
    # ========================================================

    recebeu_informacoes_adicionais = (
        informacoes_adicionais is not None
    )

    recebeu_categorias = (
        categorias is not None
    )

    recebeu_blocos = (
        blocos is not None
    )

    recebeu_segmentos = (
        segmentos is not None
    )

    recebeu_tags = (
        tags is not None
    )

    recebeu_trechos_utilizados = (
        trechos_utilizados is not None
    )

    # ========================================================
    # NORMALIZAÇÃO
    # ========================================================

    if not isinstance(
        informacoes_adicionais,
        dict
    ):
        informacoes_adicionais = {}

    if not isinstance(
        categorias,
        dict
    ):
        categorias = {}

    # ========================================================
    # NORMALIZAR BLOCOS
    # ========================================================

    if isinstance(
        blocos,
        list
    ):

        blocos_normalizados = {}

        for bloco in blocos:

            if not isinstance(
                bloco,
                dict
            ):
                continue

            try:
                numero_bloco = int(
                    bloco.get(
                        "numero",
                        0
                    )
                )
            except Exception:
                numero_bloco = 0

            if numero_bloco < 1 or numero_bloco > 5:
                continue

            blocos_normalizados[
                f"bloco_{numero_bloco}"
            ] = bloco

        blocos = blocos_normalizados

    elif not isinstance(
        blocos,
        dict
    ):

        blocos = {}

    if not isinstance(
        segmentos,
        list
    ):
        segmentos = []

    if not isinstance(
        tags,
        list
    ):
        tags = []

    if not isinstance(
        trechos_utilizados,
        list
    ):
        trechos_utilizados = []

    # ========================================================
    # VALIDAR TEMA
    # ========================================================

    if not tema:

        print(
            "ERRO: tema vazio ao salvar banco."
        )

        return False

    tema_original = str(
        tema
    ).strip()

    if not tema_original:

        print(
            "ERRO: tema vazio ao salvar banco."
        )

        return False

    try:

        tema_normalizado = normalizar_tema_chave(
            tema_original
        )

    except Exception:

        tema_normalizado = (
            tema_original
            .strip()
            .lower()
        )

    # ========================================================
    # GARANTIR DIRETÓRIO
    # ========================================================

    try:

        diretorio = os.path.dirname(
            ARQUIVO_BANCO
        )

        if diretorio:

            os.makedirs(
                diretorio,
                exist_ok=True
            )

    except Exception as erro:

        print(
            "AVISO: não foi possível preparar diretório:",
            repr(erro)
        )

    # ========================================================
    # CARREGAR BANCO
    # ========================================================

    # Se o arquivo oficial já existe, nunca tratar falha de leitura como
    # banco vazio: isso poderia apagar os demais temas em uma gravação nova.
    if os.path.exists(ARQUIVO_BANCO):
        try:
            with open(ARQUIVO_BANCO, "r", encoding="utf-8") as arquivo_banco_existente:
                banco_existente_verificado = json.load(arquivo_banco_existente)
            if not isinstance(banco_existente_verificado, dict):
                print("❌ BANCO OFICIAL INVÁLIDO: a raiz do JSON não é um objeto. Gravação cancelada para preservar o arquivo.")
                return False
        except Exception as erro_leitura_banco:
            print("❌ NÃO FOI POSSÍVEL LER O BANCO OFICIAL EXISTENTE. GRAVAÇÃO CANCELADA PARA NÃO SOBRESCREVER DADOS.")
            print("DETALHE:", repr(erro_leitura_banco))
            return False

    banco = carregar_banco()

    if not isinstance(banco, dict):
        print("❌ BANCO OFICIAL NÃO DISPONÍVEL. GRAVAÇÃO CANCELADA.")
        return False

    # ========================================================
    # LOCALIZAR TEMA EXISTENTE
    # ========================================================

    chave_existente = None

    for chave in list(
        banco.keys()
    ):

        if not isinstance(
            chave,
            str
        ):
            continue

        try:

            chave_normalizada = normalizar_tema_chave(
                chave
            )

        except Exception:

            chave_normalizada = (
                chave
                .strip()
                .lower()
            )

        if chave_normalizada == tema_normalizado:

            chave_existente = chave
            break

    # ========================================================
    # PRIMEIRA GRAVAÇÃO DA EXECUÇÃO
    # ========================================================

    primeira_gravacao = (
        tema_normalizado
        not in PAGINAS_EM_PROCESSAMENTO
    )

    if primeira_gravacao:

        # ----------------------------------------------------
        # Remover versão anterior do mesmo tema.
        # ----------------------------------------------------

        for chave in list(
            banco.keys()
        ):

            if not isinstance(
                chave,
                str
            ):
                continue

            try:

                chave_normalizada = normalizar_tema_chave(
                    chave
                )

            except Exception:

                chave_normalizada = (
                    chave
                    .strip()
                    .lower()
                )

            if chave_normalizada == tema_normalizado:

                del banco[
                    chave
                ]

        chave_existente = tema_original

        estrutura_inicial = criar_estrutura_json_pagina(
            tema_original
        )

        dados_iniciais = estrutura_inicial.get(
            tema_original,
            {}
        )

        if not isinstance(
            dados_iniciais,
            dict
        ):
            dados_iniciais = {}

        banco[
            chave_existente
        ] = dados_iniciais

        PAGINAS_EM_PROCESSAMENTO.add(
            tema_normalizado
        )

    elif chave_existente is None:

        chave_existente = tema_original

        estrutura_inicial = criar_estrutura_json_pagina(
            tema_original
        )

        dados_iniciais = estrutura_inicial.get(
            tema_original,
            {}
        )

        if not isinstance(
            dados_iniciais,
            dict
        ):
            dados_iniciais = {}

        banco[
            chave_existente
        ] = dados_iniciais

    # ========================================================
    # RECUPERAR DADOS DO TEMA
    # ========================================================

    dados_tema = banco.get(
        chave_existente
    )

    if not isinstance(
        dados_tema,
        dict
    ):

        estrutura_inicial = criar_estrutura_json_pagina(
            tema_original
        )

        dados_tema = estrutura_inicial.get(
            tema_original,
            {}
        )

        if not isinstance(
            dados_tema,
            dict
        ):
            dados_tema = {}

        banco[
            chave_existente
        ] = dados_tema

    # ========================================================
    # RECUPERAR PÁGINA
    # ========================================================

    pagina = dados_tema.get(
        "pagina"
    )

    if not isinstance(
        pagina,
        dict
    ):
        pagina = {}

    # ========================================================
    # FUNÇÕES AUXILIARES
    # ========================================================

    def normalizar_lista_local(valor):

        if isinstance(
            valor,
            list
        ):

            resultado = []

            for item in valor:

                item_limpo = str(
                    item or ""
                ).strip()

                if item_limpo:

                    resultado.append(
                        item_limpo
                    )

            return resultado

        if valor is None:
            return []

        resultado = []

        for item in re.split(
            r"[,;\n]+",
            str(valor)
        ):

            item_limpo = item.strip()

            if item_limpo:

                resultado.append(
                    item_limpo
                )

        return resultado

    # ========================================================
    # EXTRAIR TEXTO EDITORIAL
    #
    # IMPORTANTE:
    #
    # Nunca mais procurar "paragrafos".
    #
    # O texto editorial oficial é:
    # paragrafos_ollama
    # ========================================================

    def extrair_paragrafos_bloco(bloco):

        if not isinstance(
            bloco,
            dict
        ):
            return []

        paragrafos_ollama = bloco.get(
            "paragrafos_ollama",
            []
        )

        if not isinstance(
            paragrafos_ollama,
            list
        ):
            return []

        resultado = []

        for item in paragrafos_ollama[:3]:

            if item is None:
                continue

            texto_paragrafo = str(
                item
            )

            if texto_paragrafo.strip():

                resultado.append(
                    texto_paragrafo
                )

        return resultado[:3]

    # ========================================================
    # CRIAR BLOCO VAZIO
    # ========================================================

    def criar_bloco_vazio(numero):

        return {

            "id":
                f"bloco_{numero}",

            "hash":
                "",

            "informacoes_relevantes":
                [],

            "titulo":
                "",

            "paragrafos_python":
                [
                    "",
                    "",
                    ""
                ],

            "paragrafos_ollama":
                [
                    "",
                    "",
                    ""
                ]
        }

    # ========================================================
    # ATUALIZAR BLOCO
    # ========================================================

    def atualizar_bloco(
        numero,
        dados_bloco
    ):

        if numero < 1 or numero > 5:
            return

        chave_bloco = (
            f"bloco_{numero}"
        )

        alvo = pagina.get(
            chave_bloco
        )

        if not isinstance(
            alvo,
            dict
        ):

            alvo = criar_bloco_vazio(
                numero
            )

        if not isinstance(
            dados_bloco,
            dict
        ):

            dados_bloco = {}

        # ====================================================
        # ID
        # ====================================================

        id_bloco = str(
            dados_bloco.get(
                "id",
                alvo.get(
                    "id",
                    f"bloco_{numero}"
                )
            )
            or f"bloco_{numero}"
        ).strip()

        alvo[
            "id"
        ] = id_bloco

        # ====================================================
        # INFORMAÇÕES RELEVANTES
        #
        # LISTA DE OBJETOS.
        #
        # NÃO transformar em texto.
        # ====================================================

        if "informacoes_relevantes" in dados_bloco:

            info_recebida = (
                dados_bloco.get(
                    "informacoes_relevantes"
                )
            )

            if isinstance(
                info_recebida,
                list
            ):

                info_recebida = [
                    item
                    for item in info_recebida
                    if isinstance(
                        item,
                        dict
                    )
                ]

                if info_recebida:

                    alvo[
                        "informacoes_relevantes"
                    ] = info_recebida

        # ====================================================
        # TÍTULO
        # ====================================================

        if "titulo" in dados_bloco:

            titulo_recebido = str(
                dados_bloco.get(
                    "titulo",
                    ""
                )
                or ""
            ).strip()

            if titulo_recebido:

                alvo[
                    "titulo"
                ] = titulo_recebido

        # ====================================================
        # PARÁGRAFOS PYTHON
        #
        # Estes são os trechos selecionados pelo Python.
        #
        # Não escrever, resumir ou alterar o conteúdo aqui.
        # ====================================================

        if "paragrafos_python" in dados_bloco:

            paragrafos_python = (
                dados_bloco.get(
                    "paragrafos_python"
                )
            )

            if isinstance(
                paragrafos_python,
                list
            ):

                alvo[
                    "paragrafos_python"
                ] = (
                    list(
                        paragrafos_python[:3]
                    )
                    +
                    [
                        "",
                        "",
                        ""
                    ]
                )[:3]

        # ====================================================
        # PARÁGRAFOS OLLAMA
        # ====================================================

        if "paragrafos_ollama" in dados_bloco:

            paragrafos_ollama = (
                dados_bloco.get(
                    "paragrafos_ollama"
                )
            )

            if isinstance(
                paragrafos_ollama,
                list
            ):

                alvo[
                    "paragrafos_ollama"
                ] = (
                    [
                        str(
                            item or ""
                        ).strip()
                        for item
                        in paragrafos_ollama[:3]
                    ]
                    +
                    [
                        "",
                        "",
                        ""
                    ]
                )[:3]

        # ====================================================
        # COMPATIBILIDADE CONTROLADA
        #
        # Caso uma chamada antiga de conteudo_completo ainda
        # entregue "paragrafos", não criamos esse campo.
        #
        # Apenas transportamos o conteúdo para
        # paragrafos_ollama.
        # ====================================================

        if (
            "paragrafos" in dados_bloco
            and
            "paragrafos_ollama" not in dados_bloco
        ):

            paragrafos_antigos = (
                dados_bloco.get(
                    "paragrafos"
                )
            )

            if isinstance(
                paragrafos_antigos,
                list
            ):

                alvo[
                    "paragrafos_ollama"
                ] = (
                    [
                        str(
                            item or ""
                        ).strip()
                        for item
                        in paragrafos_antigos[:3]
                    ]
                    +
                    [
                        "",
                        "",
                        ""
                    ]
                )[:3]

        # ====================================================
        # GARANTIR ESTRUTURA DOS CAMPOS
        # ====================================================

        paragrafos_python = alvo.get(
            "paragrafos_python",
            []
        )

        if not isinstance(
            paragrafos_python,
            list
        ):
            paragrafos_python = []

        alvo[
            "paragrafos_python"
        ] = (
            list(
                paragrafos_python[:3]
            )
            +
            [
                "",
                "",
                ""
            ]
        )[:3]

        paragrafos_ollama = alvo.get(
            "paragrafos_ollama",
            []
        )

        if not isinstance(
            paragrafos_ollama,
            list
        ):
            paragrafos_ollama = []

        alvo[
            "paragrafos_ollama"
        ] = (
            [
                str(
                    item or ""
                ).strip()
                for item
                in paragrafos_ollama[:3]
            ]
            +
            [
                "",
                "",
                ""
            ]
        )[:3]

        # ====================================================
        # HASH
        #
        # O hash utiliza:
        # informacoes_relevantes
        # titulo
        # paragrafos_ollama
        #
        # Nunca "paragrafos".
        # ====================================================

        info_hash = json.dumps(
            alvo.get(
                "informacoes_relevantes",
                []
            ),
            ensure_ascii=False,
            sort_keys=True
        )

        conteudo_hash = (
            info_hash
            + "|"
            + str(
                alvo.get(
                    "titulo",
                    ""
                )
                or ""
            ).strip()
            + "|"
            + "|".join(
                str(
                    paragrafo or ""
                ).strip()
                for paragrafo
                in alvo.get(
                    "paragrafos_ollama",
                    []
                )[:3]
            )
        )

        alvo[
            "hash"
        ] = hashlib.sha256(
            conteudo_hash.encode(
                "utf-8"
            )
        ).hexdigest()

        # ====================================================
        # PROTEÇÃO CONTRA CAMPOS LEGADOS
        # ====================================================
        
        alvo.pop(
            "paragrafos",
            None
        )
        
        alvo.pop(
            "imagens",
            None
        )
        
        alvo.pop(
            "caracteres",
            None
        )
        
        alvo.pop(
            "status",
            None
        )
        
        pagina[
            chave_bloco
        ] = alvo

    # ========================================================
    # METADADOS PRINCIPAIS
    # ========================================================

    dados_tema[
        "tema"
    ] = tema_original

    try:

        grupo_identificado = identificar_grupo_tema(
            tema_original
        )

    except Exception:

        grupo_identificado = ""

    dados_tema[
        "grupo"
    ] = (
        grupo_identificado
        or dados_tema.get(
            "grupo",
            ""
        )
    )

    # ========================================================
    # TIPO
    # ========================================================

    if tipo is not None:

        tipo_recebido = str(
            tipo or ""
        ).strip()

        if tipo_recebido:

            dados_tema[
                "tipo"
            ] = tipo_recebido

    # ========================================================
    # TAGS
    # ========================================================

    if recebeu_tags:

        tags_normalizadas = []
        vistos_tags = set()

        for tag in normalizar_lista_local(
            tags
        ):

            try:

                chave_tag = normalizar_tema_chave(
                    tag
                )

            except Exception:

                chave_tag = tag.lower()

            if chave_tag in vistos_tags:
                continue

            vistos_tags.add(
                chave_tag
            )

            tags_normalizadas.append(
                tag
            )

        dados_tema[
            "tags"
        ] = tags_normalizadas[:30]

    else:

        if not isinstance(
            dados_tema.get(
                "tags"
            ),
            list
        ):

            dados_tema[
                "tags"
            ] = []

    # ========================================================
    # SEGMENTOS TEXTUAIS
    # ========================================================

    if recebeu_informacoes_adicionais:

        segmentos_textuais = (
            informacoes_adicionais.get(
                "segmentos_textuais",
                []
            )
        )

        segmentos_textuais = (
            normalizar_lista_local(
                segmentos_textuais
            )
        )

    elif recebeu_segmentos:

        segmentos_textuais = (
            normalizar_lista_local(
                segmentos
            )
        )

    else:

        informacoes_existentes = (
            dados_tema.get(
                "informacoes_adicionais",
                {}
            )
        )

        if isinstance(
            informacoes_existentes,
            dict
        ):

            segmentos_textuais = (
                informacoes_existentes.get(
                    "segmentos_textuais",
                    []
                )
            )

        else:

            segmentos_textuais = []

        if isinstance(
            segmentos_textuais,
            dict
        ):

            segmentos_textuais = [
                item
                for lista
                in segmentos_textuais.values()
                if isinstance(
                    lista,
                    list
                )
                for item
                in lista
                if item
            ]

        segmentos_textuais = (
            normalizar_lista_local(
                segmentos_textuais
            )
        )

    # ========================================================
    # REMOVER DUPLICADOS DOS SEGMENTOS
    # ========================================================

    segmentos_textuais_unicos = []
    vistos_segmentos = set()

    for segmento in segmentos_textuais:

        try:

            chave_segmento = normalizar_tema_chave(
                segmento
            )

        except Exception:

            chave_segmento = segmento.lower()

        if chave_segmento in vistos_segmentos:
            continue

        vistos_segmentos.add(
            chave_segmento
        )

        segmentos_textuais_unicos.append(
            segmento
        )

    segmentos_textuais = (
        segmentos_textuais_unicos[:12]
    )

    # ========================================================
    # FONTES
    # ========================================================

    if recebeu_informacoes_adicionais:

        fontes = (
            informacoes_adicionais.get(
                "fontes",
                []
            )
        )

        if not isinstance(
            fontes,
            list
        ):
            fontes = []

        fontes = [
            fonte
            for fonte in fontes
            if fonte
        ]

    else:

        informacoes_existentes = (
            dados_tema.get(
                "informacoes_adicionais",
                {}
            )
        )

        if isinstance(
            informacoes_existentes,
            dict
        ):

            fontes = (
                informacoes_existentes.get(
                    "fontes",
                    []
                )
            )

        else:

            fontes = []

    # ========================================================
    # REFERÊNCIAS
    # ========================================================

    if recebeu_informacoes_adicionais:

        referencias = (
            informacoes_adicionais.get(
                "referencias",
                []
            )
        )

        if not isinstance(
            referencias,
            list
        ):
            referencias = []

        referencias = [
            referencia
            for referencia in referencias
            if referencia
        ]

    else:

        informacoes_existentes = (
            dados_tema.get(
                "informacoes_adicionais",
                {}
            )
        )

        if isinstance(
            informacoes_existentes,
            dict
        ):

            referencias = (
                informacoes_existentes.get(
                    "referencias",
                    []
                )
            )

        else:

            referencias = []

    # ========================================================
    # NOME DO SITE
    # ========================================================

    if recebeu_informacoes_adicionais:

        nome_site = str(
            informacoes_adicionais.get(
                "nome_site",
                ""
            )
            or ""
        ).strip()

    else:

        informacoes_existentes = (
            dados_tema.get(
                "informacoes_adicionais",
                {}
            )
        )

        if isinstance(
            informacoes_existentes,
            dict
        ):

            nome_site = str(
                informacoes_existentes.get(
                    "nome_site",
                    ""
                )
                or ""
            ).strip()

        else:

            nome_site = ""

    # ========================================================
    # GRUPO PRINCIPAL DO PROJETO
    # ========================================================

    if grupo_principal_projeto is None:

        grupo_principal_projeto = str(
            dados_tema.get(
                "grupo_principal_projeto",
                ""
            )
            or ""
        ).strip()

    else:

        grupo_principal_projeto = str(
            grupo_principal_projeto
            or ""
        ).strip()

    # ========================================================
    # TRECHOS UTILIZADOS
    # ========================================================

    dados_tema_existente = (
        dados_tema.get(
            "informacoes_adicionais",
            {}
        )
    )

    if not isinstance(
        dados_tema_existente,
        dict
    ):

        dados_tema_existente = {}

    trechos_existentes = (
        dados_tema_existente.get(
            "trechos_utilizados",
            []
        )
    )

    if not isinstance(
        trechos_existentes,
        list
    ):

        trechos_existentes = []

    hashes_existentes = set()

    for existente in trechos_existentes:

        if not isinstance(
            existente,
            dict
        ):
            continue

        hash_existente = str(
            existente.get(
                "hash",
                ""
            )
            or ""
        ).strip()

        if hash_existente:

            hashes_existentes.add(
                hash_existente
            )

    if recebeu_trechos_utilizados:

        for item in trechos_utilizados:

            if isinstance(
                item,
                dict
            ):

                trecho = str(
                    item.get(
                        "trecho",
                        ""
                    )
                    or ""
                ).strip()

                hash_trecho = str(
                    item.get(
                        "hash",
                        ""
                    )
                    or ""
                ).strip()

                if not hash_trecho and trecho:

                    try:

                        hash_trecho = gerar_hash_trecho(
                            trecho
                        )

                    except Exception:

                        hash_trecho = ""

                if not hash_trecho:
                    continue

                if hash_trecho in hashes_existentes:
                    continue

                trechos_existentes.append({

                    "id":
                        str(
                            item.get(
                                "id",
                                hash_trecho
                            )
                            or hash_trecho
                        ),

                    "hash":
                        hash_trecho,

                    "tema_origem":
                        str(
                            item.get(
                                "tema_origem",
                                tema_original
                            )
                            or tema_original
                        ).strip(),

                    "fonte":
                        str(
                            item.get(
                                "fonte",
                                ""
                            )
                            or ""
                        ).strip()
                })

                hashes_existentes.add(
                    hash_trecho
                )

            else:

                trecho = str(
                    item or ""
                ).strip()

                if not trecho:
                    continue

                try:

                    hash_trecho = gerar_hash_trecho(
                        trecho
                    )

                except Exception:

                    hash_trecho = ""

                if not hash_trecho:
                    continue

                if hash_trecho in hashes_existentes:
                    continue

                trechos_existentes.append({

                    "id":
                        hash_trecho,

                    "hash":
                        hash_trecho,

                    "tema_origem":
                        tema_original,

                    "fonte":
                        ""
                })

                hashes_existentes.add(
                    hash_trecho
                )

    # ========================================================
    # GARANTIR BLOCOS RECEBIDOS
    # ========================================================

    if recebeu_blocos:

        for numero in range(
            1,
            6
        ):

            chave_bloco = (
                f"bloco_{numero}"
            )

            if chave_bloco not in blocos:
                continue

            dados_bloco = (
                blocos.get(
                    chave_bloco
                )
            )

            if isinstance(
                dados_bloco,
                dict
            ):

                dados_bloco_final = dict(
                    dados_bloco
                )

            else:

                dados_bloco_final = {
                    "informacoes_relevantes":
                        dados_bloco
                }

            atualizar_bloco(
                numero,
                dados_bloco_final
            )

    # ========================================================
    # CATEGORIA: PÁGINA
    # ========================================================

    if categoria == "pagina":

        if isinstance(
            texto,
            dict
        ):

            pagina_recebida = texto

            pagina[
                "tema"
            ] = str(
                pagina_recebida.get(
                    "tema",
                    tema_original
                )
                or tema_original
            ).strip()

            pagina[
                "arquivo_origem"
            ] = str(
                pagina_recebida.get(
                    "arquivo_origem",
                    pagina.get(
                        "arquivo_origem",
                        ""
                    )
                )
                or ""
            ).strip()

            # =================================================
            # SUBTÍTULO
            # =================================================

            if subtitulo is not None:

                subtitulo_recebido = str(
                    subtitulo or ""
                ).strip()

                if subtitulo_recebido:

                    pagina[
                        "subtitulo"
                    ] = subtitulo_recebido

            elif pagina_recebida.get(
                "subtitulo"
            ):

                pagina[
                    "subtitulo"
                ] = str(
                    pagina_recebida.get(
                        "subtitulo"
                    )
                    or ""
                ).strip()

            # =================================================
            # SUBTÍTULO DOS SEGMENTOS
            # =================================================

            if subtitulo_segmentos is not None:

                subtitulo_segmentos_recebido = str(
                    subtitulo_segmentos or ""
                ).strip()

                if subtitulo_segmentos_recebido:

                    pagina[
                        "subtitulo_segmentos"
                    ] = subtitulo_segmentos_recebido

            elif pagina_recebida.get(
                "subtitulo_segmentos"
            ):

                pagina[
                    "subtitulo_segmentos"
                ] = str(
                    pagina_recebida.get(
                        "subtitulo_segmentos"
                    )
                    or ""
                ).strip()

            # =================================================
            # H1 E TÍTULO
            # =================================================

            h1_recebido = str(
                pagina_recebida.get(
                    "h1",
                    ""
                )
                or ""
            ).strip()

            titulo_recebido = str(
                pagina_recebida.get(
                    "titulo",
                    ""
                )
                or ""
            ).strip()

            if h1_recebido:

                pagina[
                    "h1"
                ] = h1_recebido

            elif not pagina.get(
                "h1",
                ""
            ):

                pagina[
                    "h1"
                ] = tema_original

            if titulo_recebido:

                pagina[
                    "titulo"
                ] = titulo_recebido

            elif not pagina.get(
                "titulo",
                ""
            ):

                pagina[
                    "titulo"
                ] = tema_original

            # =================================================
            # BLOCOS RECEBIDOS DENTRO DA PÁGINA
            # =================================================

            blocos_recebidos = (
                pagina_recebida.get(
                    "blocos",
                    []
                )
            )

            if isinstance(
                blocos_recebidos,
                list
            ):

                for bloco in blocos_recebidos:

                    if not isinstance(
                        bloco,
                        dict
                    ):
                        continue

                    try:

                        numero = int(
                            bloco.get(
                                "numero",
                                0
                            )
                        )

                    except Exception:

                        numero = 0

                    atualizar_bloco(
                        numero,
                        bloco
                    )

            # =================================================
            # SEGMENTOS RECEBIDOS
            # =================================================

            segmentos_recebidos = (
                pagina_recebida.get(
                    "segmentos",
                    []
                )
            )

            if (
                isinstance(
                    segmentos_recebidos,
                    list
                )
                and
                segmentos_recebidos
            ):

                segmentos = (
                    segmentos_recebidos
                )

                recebeu_segmentos = True


        # =================================================
        # PRESERVAR METADADOS RECEBIDOS PELO PYTHON
        #
        # Esses valores são gerados por gerar_titulos()
        # e precisam ser preservados mesmo quando
        # "texto" recebido por salvar_banco() for uma string.
        # =================================================

        if h1 is not None:

            h1_recebido = str(
                h1 or ""
            ).strip()

            if h1_recebido:

                pagina[
                    "h1"
                ] = h1_recebido


        if titulo is not None:

            titulo_recebido = str(
                titulo or ""
            ).strip()

            if titulo_recebido:

                pagina[
                    "titulo"
                ] = titulo_recebido


        if subtitulo is not None:

            subtitulo_recebido = str(
                subtitulo or ""
            ).strip()

            if subtitulo_recebido:

                pagina[
                    "subtitulo"
                ] = subtitulo_recebido


        if subtitulo_segmentos is not None:

            subtitulo_segmentos_recebido = str(
                subtitulo_segmentos or ""
            ).strip()

            if subtitulo_segmentos_recebido:

                pagina[
                    "subtitulo_segmentos"
                ] = subtitulo_segmentos_recebido


    # ========================================================
    # CATEGORIA: CONTEÚDO COMPLETO
    # ========================================================

    elif categoria == "conteudo_completo":

        conteudo = str(
            texto or ""
        ).strip()

        # ----------------------------------------------------
        # TÍTULO
        # ----------------------------------------------------

        padroes_titulo = [

            r"(?im)^\s*(?:#\s*)?T[ÍI]TULO\s*:\s*(.+)$",

            r"(?im)^\s*##\s*T[ÍI]TULO\s*:?\s*(.+)$",

            r"(?im)^\s*TITLE\s*:\s*(.+)$"
        ]

        for padrao in padroes_titulo:

            resultado = re.search(
                padrao,
                conteudo
            )

            if resultado:

                pagina[
                    "titulo"
                ] = resultado.group(
                    1
                ).strip()

                break
                
        # ====================================================
        # PRESERVAR H1 E TÍTULO GERADOS PELO PYTHON
        # ====================================================

        if h1 is not None:

            h1_recebido = str(
                h1 or ""
            ).strip()

            if h1_recebido:

                pagina[
                    "h1"
                ] = h1_recebido

        if titulo is not None:

            titulo_recebido = str(
                titulo or ""
            ).strip()

            if titulo_recebido:

                pagina[
                    "titulo"
                ] = titulo_recebido        

        # ----------------------------------------------------
        # SUBTÍTULO
        # ----------------------------------------------------

        padroes_subtitulo = [

            r"(?im)^\s*(?:#\s*)?SUBT[ÍI]TULO\s*:\s*(.+)$",

            r"(?im)^\s*##\s*SUBT[ÍI]TULO\s*:?\s*(.+)$",

            r"(?im)^\s*SUBTITLE\s*:\s*(.+)$"
        ]

        for padrao in padroes_subtitulo:

            resultado = re.search(
                padrao,
                conteudo
            )

            if resultado:

                pagina[
                    "subtitulo"
                ] = resultado.group(
                    1
                ).strip()

                break

        # ====================================================
        # PRESERVAR SUBTÍTULO
        # ====================================================

        if subtitulo is not None:

            subtitulo_recebido = str(
                subtitulo or ""
            ).strip()

            if subtitulo_recebido:

                pagina[
                    "subtitulo"
                ] = subtitulo_recebido

        # ====================================================
        # PRESERVAR SUBTÍTULO DOS SEGMENTOS
        # ====================================================

        if subtitulo_segmentos is not None:

            subtitulo_segmentos_recebido = str(
                subtitulo_segmentos or ""
            ).strip()

            if subtitulo_segmentos_recebido:

                pagina[
                    "subtitulo_segmentos"
                ] = subtitulo_segmentos_recebido

        # ----------------------------------------------------
        # BLOCOS 1 A 5
        # ----------------------------------------------------

        padrao_bloco = re.compile(
            r"(?im)"
            r"^\s*"
            r"(?:#+\s*)?"
            r"BLOCO\s*"
            r"([1-5])"
            r"\s*:?"
            r"(.*)$"
        )

        ocorrencias = list(
            padrao_bloco.finditer(
                conteudo
            )
        )

        for idx, match in enumerate(
            ocorrencias
        ):

            inicio = match.end()

            if idx + 1 < len(
                ocorrencias
            ):

                fim = ocorrencias[
                    idx + 1
                ].start()

            else:

                fim = len(
                    conteudo
                )

            conteudo_bloco = conteudo[
                inicio:fim
            ].strip()

            titulo_bloco = (
                match.group(
                    2
                ).strip()
            )

            paragrafos_ollama = []

            partes = re.split(
                r"\n\s*\n+",
                conteudo_bloco
            )

            for parte in partes:

                parte_limpa = parte.strip()

                if not parte_limpa:
                    continue

                if re.match(
                    r"(?i)^\s*(?:TITULO|SUBTITULO|TAGS?|SEGMENTOS?)\s*:",
                    parte_limpa
                ):
                    continue

                paragrafos_ollama.append(
                    parte_limpa
                )

            atualizar_bloco(
                int(
                    match.group(
                        1
                    )
                ),
                {
                    "titulo":
                        titulo_bloco,

                    "paragrafos_ollama":
                        paragrafos_ollama[:3]
                }
            )

        # ----------------------------------------------------
        # SEGMENTOS
        # ----------------------------------------------------

        padroes_segmentos = [

            r"(?is)"
            r"(?:^|\n)"
            r"\s*(?:#+\s*)?"
            r"SEGMENTOS?\s*:?\s*"
            r"\n?"
            r"(.*?)(?="
            r"\n\s*(?:#+\s*)?"
            r"(?:TAGS?|BLOCO|FIM|$)"
            r")",

            r"(?is)"
            r"\[\s*SEGMENTOS?\s*\]"
            r"\s*(.*?)(?="
            r"\[\s*TAGS?\s*\]|$)"
        ]

        bloco_segmentos = None

        for padrao in padroes_segmentos:

            resultado = re.search(
                padrao,
                conteudo
            )

            if resultado:

                bloco_segmentos = resultado.group(
                    1
                ).strip()

                break

        if bloco_segmentos:

            linhas_segmentos = (
                bloco_segmentos.splitlines()
            )

            segmentos_extraidos = []

            for linha in linhas_segmentos:

                linha = linha.strip()

                if not linha:
                    continue

                linha = re.sub(
                    r"^\s*[-•*]\s*",
                    "",
                    linha
                )

                linha = re.sub(
                    r"^\s*\d+[\.\)\-:]\s*",
                    "",
                    linha
                )

                linha = linha.strip()

                if linha:

                    segmentos_extraidos.append(
                        linha
                    )

            segmentos_unicos = []
            vistos_segmentos = set()

            for segmento in segmentos_extraidos:

                try:

                    chave_segmento = normalizar_tema_chave(
                        segmento
                    )

                except Exception:

                    chave_segmento = (
                        segmento.lower()
                    )

                if chave_segmento in vistos_segmentos:
                    continue

                vistos_segmentos.add(
                    chave_segmento
                )

                segmentos_unicos.append(
                    segmento
                )

            segmentos = (
                segmentos_unicos[:12]
            )

            segmentos_textuais = (
                segmentos
            )

        # ----------------------------------------------------
        # TAGS
        # ----------------------------------------------------

        padroes_tags = [

            r"(?is)"
            r"(?:^|\n)"
            r"\s*(?:#+\s*)?"
            r"TAGS?\s*:?\s*"
            r"\n?"
            r"(.*?)(?="
            r"\n\s*(?:#+\s*)?"
            r"(?:SEGMENTOS?|BLOCO|FIM|$)"
            r")",

            r"(?is)"
            r"\[\s*TAGS?\s*\]"
            r"\s*(.*?)(?="
            r"\[\s*SEGMENTOS?\s*\]|$)"
        ]

        bloco_tags = None

        for padrao in padroes_tags:

            resultado = re.search(
                padrao,
                conteudo
            )

            if resultado:

                bloco_tags = resultado.group(
                    1
                ).strip()

                break

        if bloco_tags:

            tags_extraidas = []

            for parte in re.split(
                r"[,;\n]+",
                bloco_tags
            ):

                tag = re.sub(
                    r"^\s*[-•*]\s*",
                    "",
                    parte
                ).strip()

                tag = re.sub(
                    r"^\s*\d+[\.\)\-:]\s*",
                    "",
                    tag
                ).strip()

                if tag:

                    tags_extraidas.append(
                        tag
                    )

            tags = tags_extraidas
            recebeu_tags = True

    # ========================================================
    # CATEGORIA: ARQUIVO DE ORIGEM
    # ========================================================

    elif categoria == "arquivo_origem":

        pagina[
            "arquivo_origem"
        ] = str(
            texto or ""
        ).strip()

    # ========================================================
    # CATEGORIA: TÍTULO
    # ========================================================

    elif categoria == "titulo":

        pagina[
            "titulo"
        ] = str(
            texto or ""
        ).strip()

    # ========================================================
    # CATEGORIA: SUBTÍTULO
    # ========================================================

    elif categoria == "subtitulo":

        pagina[
            "subtitulo"
        ] = str(
            texto or ""
        ).strip()

    # ========================================================
    # CATEGORIA: TAGS
    # ========================================================

    elif categoria == "tags":

        tags = normalizar_lista_local(
            texto
        )

        recebeu_tags = True

    # ========================================================
    # CATEGORIA: SEGMENTOS
    # ========================================================

    elif categoria == "segmentos":

        segmentos = normalizar_lista_local(
            texto
        )

        recebeu_segmentos = True

    # ========================================================
    # CATEGORIA: BLOCOS
    # ========================================================

    elif categoria == "blocos":

        if isinstance(
            texto,
            list
        ):

            for bloco in texto:

                if not isinstance(
                    bloco,
                    dict
                ):
                    continue

                try:

                    numero = int(
                        bloco.get(
                            "numero",
                            0
                        )
                    )

                except Exception:

                    numero = 0

                atualizar_bloco(
                    numero,
                    bloco
                )

    # ========================================================
    # CATEGORIA: MAPA MEAD
    # ========================================================

    elif categoria == "mapa_mead":

        valor = str(
            texto or ""
        ).strip()

        dados_tema[
            "mapa_mead"
        ] = {

            "status":
                "gerado"
                if valor
                else "vazio",

            "texto":
                valor
        }

    # ========================================================
    # CATEGORIA: ESTRUTURA
    # ========================================================

    elif categoria == "estrutura":

        pass

    # ========================================================
    # NORMALIZAR TAGS FINAIS
    # ========================================================

    tags_atuais = dados_tema.get(
        "tags",
        []
    )

    if not isinstance(
        tags_atuais,
        list
    ):
        tags_atuais = []

    tags_finais = []
    vistos_tags = set()

    for tag in normalizar_lista_local(
        tags_atuais
    ):

        try:

            chave_tag = normalizar_tema_chave(
                tag
            )

        except Exception:

            chave_tag = tag.lower()

        if chave_tag in vistos_tags:
            continue

        vistos_tags.add(
            chave_tag
        )

        tags_finais.append(
            tag
        )

    dados_tema[
        "tags"
    ] = tags_finais[:30]

    # ========================================================
    # NORMALIZAR SEGMENTOS
    # ========================================================

    segmentos_finais = []
    vistos_segmentos = set()

    for segmento in normalizar_lista_local(
        segmentos_textuais
    ):

        try:

            chave_segmento = normalizar_tema_chave(
                segmento
            )

        except Exception:

            chave_segmento = segmento.lower()

        if chave_segmento in vistos_segmentos:
            continue

        vistos_segmentos.add(
            chave_segmento
        )

        segmentos_finais.append(
            segmento
        )

    segmentos_finais = (
        segmentos_finais[:12]
    )

    # ========================================================
    # GARANTIR SEGMENTOS_LISTAS
    # ========================================================

    segmentos_listas_existentes = pagina.get(
        "segmentos_listas",
        {}
    )

    if not isinstance(
        segmentos_listas_existentes,
        dict
    ):

        segmentos_listas_existentes = {}

    segmentos_listas_finais = {}

    for numero in range(
        1,
        13
    ):

        chave_segmento = (
            f"segmento_{numero}"
        )

        lista_existente = (
            segmentos_listas_existentes.get(
                chave_segmento,
                []
            )
        )

        if not isinstance(
            lista_existente,
            list
        ):

            lista_existente = []

        segmentos_listas_finais[
            chave_segmento
        ] = list(
            lista_existente
        )

    # ========================================================
    # GRAVAR SEGMENTOS RECEBIDOS PELO PYTHON
    # ========================================================

    if segmentos_finais:

        for numero, segmento in enumerate(
            segmentos_finais,
            start=1
        ):

            if numero > 12:
                break

            segmentos_listas_finais[
                f"segmento_{numero}"
            ] = [
                segmento
            ]

    pagina[
        "segmentos_listas"
    ] = segmentos_listas_finais

    # ========================================================
    # POSICIONAMENTO DAS LISTAS
    # ========================================================

    posicionamento = pagina.get(
        "posicionamento_listas",
        {}
    )

    if not isinstance(
        posicionamento,
        dict
    ):

        posicionamento = {}

    pagina[
        "posicionamento_listas"
    ] = {

        "bloco":
            posicionamento.get(
                "bloco",
                None
            )
    }

    # ========================================================
    # DADOS BÁSICOS DA PÁGINA
    # ========================================================

    pagina[
        "tema"
    ] = str(
        pagina.get(
            "tema",
            tema_original
        )
        or tema_original
    ).strip()

    pagina[
        "arquivo_origem"
    ] = str(
        pagina.get(
            "arquivo_origem",
            ""
        )
        or ""
    ).strip()

    pagina[
        "h1"
    ] = str(
        pagina.get(
            "h1",
            tema_original
        )
        or tema_original
    ).strip()

    pagina[
        "titulo"
    ] = str(
        pagina.get(
            "titulo",
            tema_original
        )
        or tema_original
    ).strip()

    pagina[
        "subtitulo"
    ] = str(
        pagina.get(
            "subtitulo",
            ""
        )
        or ""
    ).strip()

    pagina[
        "subtitulo_segmentos"
    ] = str(
        pagina.get(
            "subtitulo_segmentos",
            ""
        )
        or ""
    ).strip()
    
    
    # ========================================================
    # NOME DO SITE
    # ========================================================
    
    pagina[
        "nome_site"
    ] = str(
        nome_site or ""
    ).strip()



    # ========================================================
    # GARANTIR 5 BLOCOS
    # ========================================================

    for numero in range(
        1,
        6
    ):

        chave_bloco = (
            f"bloco_{numero}"
        )

        bloco = pagina.get(
            chave_bloco
        )

        if not isinstance(
            bloco,
            dict
        ):

            bloco = criar_bloco_vazio(
                numero
            )

        bloco[
            "id"
        ] = str(
            bloco.get(
                "id",
                f"bloco_{numero}"
            )
            or f"bloco_{numero}"
        ).strip()

        # ----------------------------------------------------
        # INFORMAÇÕES RELEVANTES
        # ----------------------------------------------------

        info_bloco = bloco.get(
            "informacoes_relevantes",
            []
        )

        if not isinstance(
            info_bloco,
            list
        ):

            info_bloco = []

        bloco[
            "informacoes_relevantes"
        ] = [
            item
            for item in info_bloco
            if isinstance(
                item,
                dict
            )
        ]

        # ----------------------------------------------------
        # TÍTULO
        # ----------------------------------------------------

        bloco[
            "titulo"
        ] = str(
            bloco.get(
                "titulo",
                ""
            )
            or ""
        ).strip()

        # ----------------------------------------------------
        # PARÁGRAFOS PYTHON
        #
        # Não transformar em conteúdo editorial.
        # ----------------------------------------------------

        paragrafos_python = bloco.get(
            "paragrafos_python",
            []
        )

        if not isinstance(
            paragrafos_python,
            list
        ):

            paragrafos_python = []

        bloco[
            "paragrafos_python"
        ] = (
            list(
                paragrafos_python[:3]
            )
            +
            [
                "",
                "",
                ""
            ]
        )[:3]

        # ----------------------------------------------------
        # PARÁGRAFOS OLLAMA
        # ----------------------------------------------------

        paragrafos_ollama = bloco.get(
            "paragrafos_ollama",
            []
        )

        if not isinstance(
            paragrafos_ollama,
            list
        ):

            paragrafos_ollama = []

        bloco[
            "paragrafos_ollama"
        ] = (
            [
                str(
                    item or ""
                ).strip()
                for item
                in paragrafos_ollama[:3]
            ]
            +
            [
                "",
                "",
                ""
            ]
        )[:3]

        # ----------------------------------------------------
        # HASH
        # ----------------------------------------------------

        info_hash = json.dumps(
            bloco.get(
                "informacoes_relevantes",
                []
            ),
            ensure_ascii=False,
            sort_keys=True
        )

        conteudo_hash = (
            info_hash
            + "|"
            + bloco.get(
                "titulo",
                ""
            )
            + "|"
            + "|".join(
                bloco.get(
                    "paragrafos_ollama",
                    []
                )[:3]
            )
        )

        bloco[
            "hash"
        ] = hashlib.sha256(
            conteudo_hash.encode(
                "utf-8"
            )
        ).hexdigest()

    # ----------------------------------------------------
    # PROTEÇÃO CONTRA CAMPOS ANTIGOS
    # ----------------------------------------------------
    
    bloco.pop(
        "paragrafos",
        None
    )
    
    bloco.pop(
        "imagens",
        None
    )
    
    bloco.pop(
        "caracteres",
        None
    )
    
    bloco.pop(
        "status",
        None
    )
    
    pagina[
        chave_bloco
    ] = bloco

    # ========================================================
    # CONTAGEM DE REPETIÇÕES
    #
    # SOMENTE TEXTO EDITORIAL DO OLLAMA.
    #
    # Os fragmentos selecionados pelo Python NÃO entram
    # nessa contagem.
    # ========================================================

    texto_para_contagem = []

    for numero in range(
        1,
        6
    ):

        bloco = pagina.get(
            f"bloco_{numero}",
            {}
        )

        if not isinstance(
            bloco,
            dict
        ):
            continue

        titulo_bloco = str(
            bloco.get(
                "titulo",
                ""
            )
            or ""
        ).strip()

        if titulo_bloco:

            texto_para_contagem.append(
                titulo_bloco
            )

        paragrafos_ollama = (
            bloco.get(
                "paragrafos_ollama",
                []
            )
        )

        if isinstance(
            paragrafos_ollama,
            list
        ):

            texto_para_contagem.extend(
                str(
                    paragrafo or ""
                ).strip()
                for paragrafo
                in paragrafos_ollama[:3]
                if str(
                    paragrafo or ""
                ).strip()
            )

    texto_para_contagem_final = " ".join(
        texto_para_contagem
    )

    try:

        palavra_chave = (
            tema_original.strip()
        )

        if palavra_chave:

            padrao_chave = re.escape(
                palavra_chave
            )

            repeticoes = len(
                re.findall(
                    rf"(?<!\w){padrao_chave}(?!\w)",
                    texto_para_contagem_final,
                    flags=re.IGNORECASE
                )
            )

        else:

            repeticoes = 0

    except Exception:

        repeticoes = 0

    dados_tema[
        "controle_repeticoes"
    ] = {

        "palavra_chave":
            tema_original,

        "meta_repeticoes":
            60,

        "repeticoes_realizadas":
            repeticoes,

        "repeticoes_faltantes":
            max(
                0,
                60 - repeticoes
            )
    }

    # ========================================================
    # PROTEÇÃO FINAL DOS BLOCOS
    # ========================================================
    
    for numero in range(
        1,
        6
    ):
    
        chave_bloco = (
            f"bloco_{numero}"
        )
    
        bloco = pagina.get(
            chave_bloco
        )
    
        if not isinstance(
            bloco,
            dict
        ):
            bloco = criar_bloco_vazio(
                numero
            )
    
        # Nunca deixar campos legados chegarem ao JSON.
    
        bloco.pop(
            "paragrafos",
            None
        )
    
        bloco.pop(
            "imagens",
            None
        )
    
        bloco.pop(
            "caracteres",
            None
        )
    
        bloco.pop(
            "status",
            None
        )
    
        pagina[
            chave_bloco
        ] = bloco

    # ========================================================
    # REMOVER CAMPOS LEGADOS DA PÁGINA
    # ========================================================

    pagina.pop(
        "informacoes_relevantes",
        None
    )

    pagina.pop(
        "imagens",
        None
    )

    pagina.pop(
        "caracteres",
        None
    )

    pagina.pop(
        "status",
        None
    )

    pagina.pop(
        "subtitulo_listas",
        None
    )

    # ========================================================
    # MONTAR PÁGINA FINAL
    # ========================================================

    pagina_final = {
    
        "tema":
            pagina.get(
                "tema",
                tema_original
            ),
    
        "arquivo_origem":
            pagina.get(
                "arquivo_origem",
                ""
            ),
    
        "nome_site":
            pagina.get(
                "nome_site",
                ""
            ),
    
        "h1":
            pagina.get(
                "h1",
                tema_original
            ),

        "subtitulo":
            pagina.get(
                "subtitulo",
                ""
            ),

        "subtitulo_segmentos":
            pagina.get(
                "subtitulo_segmentos",
                ""
            )
    }

    # ========================================================
    # 5 BLOCOS
    # ========================================================

    for numero in range(
        1,
        6
    ):

        chave_bloco = (
            f"bloco_{numero}"
        )

        pagina_final[
            chave_bloco
        ] = pagina[
            chave_bloco
        ]

    # ========================================================
    # SEGMENTOS_LISTAS
    # ========================================================

    pagina_final[
        "segmentos_listas"
    ] = pagina.get(
        "segmentos_listas",
        {}
    )

    # ========================================================
    # POSICIONAMENTO
    # ========================================================

    pagina_final[
        "posicionamento_listas"
    ] = pagina.get(
        "posicionamento_listas",
        {
            "bloco": None
        }
    )

    # ========================================================
    # DADOS FINAIS DO TEMA
    # ========================================================

    dados_tema_final = {

        "tema":
            tema_original,

        "grupo":
            dados_tema.get(
                "grupo",
                ""
            ),

        "tipo":
            dados_tema.get(
                "tipo",
                ""
            ),

        "tags":
            dados_tema.get(
                "tags",
                []
            )[:30],

        "controle_repeticoes":
            dados_tema.get(
                "controle_repeticoes",
                {}
            ),

        "mapa_mead":
            dados_tema.get(
                "mapa_mead",
                {
                    "status": "",
                    "texto": ""
                }
            ),

        "grupo_principal_projeto":
            grupo_principal_projeto,

        "pagina":
            pagina_final
    }

    # ========================================================
    # PROTEÇÃO FINAL ABSOLUTA
    # ========================================================

    dados_tema_final.pop(
        "informacoes_relevantes",
        None
    )

    dados_tema_final.pop(
        "imagens",
        None
    )

    dados_tema_final.pop(
        "caracteres",
        None
    )

    dados_tema_final.pop(
        "status",
        None
    )

    pagina_final.pop(
        "informacoes_relevantes",
        None
    )

    pagina_final.pop(
        "imagens",
        None
    )

    pagina_final.pop(
        "caracteres",
        None
    )

    pagina_final.pop(
        "status",
        None
    )

    pagina_final.pop(
        "subtitulo_listas",
        None
    )

    # ========================================================
    # SALVAR NA CHAVE DO TEMA
    # ========================================================

    banco[
        chave_existente
    ] = dados_tema_final

    # ========================================================
    # REMOVER CHAVES DUPLICADAS DO MESMO TEMA
    # ========================================================

    for chave in list(
        banco.keys()
    ):

        if chave == chave_existente:
            continue

        if not isinstance(
            chave,
            str
        ):
            continue

        try:

            chave_normalizada = normalizar_tema_chave(
                chave
            )

        except Exception:

            chave_normalizada = (
                chave
                .strip()
                .lower()
            )

        if chave_normalizada == tema_normalizado:

            del banco[
                chave
            ]

    # ========================================================
    # SALVAR JSON
    # ========================================================

    try:

        diretorio = os.path.dirname(
            ARQUIVO_BANCO
        )

        if diretorio:

            os.makedirs(
                diretorio,
                exist_ok=True
            )

        # ====================================================
        # VALIDAÇÃO OBRIGATÓRIA ANTES DA GRAVAÇÃO
        # ====================================================

        erros_json = []

        # O mapa MEAD é um registro preliminar independente. Ele deve ser
        # gravado antes da geração dos cinco blocos; não pode ser submetido
        # à validação de uma página que ainda nem foi produzida.
        if categoria == "pagina":
            if not isinstance(pagina_final, dict):
                erros_json.append("pagina_final inválida")

            for numero in range(1, 6):
                chave_bloco = f"bloco_{numero}"
                bloco_validacao = pagina_final.get(chave_bloco) if isinstance(pagina_final, dict) else None
                if not isinstance(bloco_validacao, dict):
                    erros_json.append(f"{chave_bloco} ausente ou inválido")
                    continue
                info_validacao = bloco_validacao.get("informacoes_relevantes", [])
                if not isinstance(info_validacao, list) or len(info_validacao) != 3:
                    erros_json.append(f"{chave_bloco}: esperado 3 fragmentos")
                para_validacao = bloco_validacao.get("paragrafos_ollama", [])
                if not isinstance(para_validacao, list) or len(para_validacao) != 3 or any(not str(x or "").strip() for x in para_validacao[:3]):
                    erros_json.append(f"{chave_bloco}: esperado 3 parágrafos Ollama não vazios")

        if erros_json:
            print("\n❌ JSON NÃO GRAVADO — ESTRUTURA INCOMPLETA")
            for erro_json in erros_json:
                print("-", erro_json)
            return False

        # ====================================================
        # GRAVAÇÃO ATÔMICA E CONFERÊNCIA REAL NO DISCO
        # ====================================================

        arquivo_temporario = None
        try:
            fd, arquivo_temporario = tempfile.mkstemp(
                prefix="conteudo-", suffix=".tmp", dir=diretorio or None, text=True
            )
            with os.fdopen(fd, "w", encoding="utf-8") as arquivo:
                json.dump(banco, arquivo, ensure_ascii=False, indent=4)
                arquivo.flush()
                os.fsync(arquivo.fileno())
            os.replace(arquivo_temporario, ARQUIVO_BANCO)
            arquivo_temporario = None
        finally:
            if arquivo_temporario and os.path.exists(arquivo_temporario):
                try:
                    os.remove(arquivo_temporario)
                except Exception:
                    pass

        with open(ARQUIVO_BANCO, "r", encoding="utf-8") as arquivo_verificacao:
            banco_gravado = json.load(arquivo_verificacao)

        tema_gravado = banco_gravado.get(chave_existente)
        pagina_gravada = tema_gravado.get("pagina", {}) if isinstance(tema_gravado, dict) else {}
        erros_pos_gravacao = []

        # A conferência 5x3 só se aplica à gravação da página. O mapa MEAD
        # é salvo antes da geração e deve ser confirmado pela leitura do JSON,
        # sem exigir que os parágrafos já existam.
        if categoria == "pagina":
            for numero in range(1, 6):
                chave_bloco = f"bloco_{numero}"
                bloco_gravado = pagina_gravada.get(chave_bloco)
                if not isinstance(bloco_gravado, dict):
                    erros_pos_gravacao.append(f"{chave_bloco} não foi gravado")
                    continue
                info_gravada = bloco_gravado.get("informacoes_relevantes", [])
                para_gravados = bloco_gravado.get("paragrafos_ollama", [])
                if not isinstance(info_gravada, list) or len(info_gravada) != 3:
                    erros_pos_gravacao.append(f"{chave_bloco}: fragmentos incompletos")
                if not isinstance(para_gravados, list) or len(para_gravados) != 3 or any(not str(x or "").strip() for x in para_gravados[:3]):
                    erros_pos_gravacao.append(f"{chave_bloco}: parágrafos incompletos")

        if erros_pos_gravacao:
            print("\n❌ GRAVAÇÃO REJEITADA NA CONFERÊNCIA DO DISCO")
            for erro_pos in erros_pos_gravacao:
                print("-", erro_pos)
            return False

        if categoria == "pagina":
            print("\n✅ JSON VALIDADO NO DISCO — 5 blocos / 15 fragmentos / 15 parágrafos")
        else:
            mapa_gravado = tema_gravado.get("mapa_mead", {}) if isinstance(tema_gravado, dict) else {}
            if categoria == "mapa_mead" and str(mapa_gravado.get("texto", "")).strip() != str(texto or "").strip():
                print("❌ MAPA MEAD NÃO CONFERE NA LEITURA DO DISCO")
                return False
            print("✅ JSON VALIDADO NO DISCO — categoria:", categoria)

        # ====================================================
        # CONFERÊNCIA FINAL
        # ====================================================

        total_blocos = 0
        total_paragrafos_ollama = 0
        total_fragmentos = 0

        for numero in range(
            1,
            6
        ):

            bloco = pagina_final.get(
                f"bloco_{numero}",
                {}
            )

            if not isinstance(
                bloco,
                dict
            ):
                continue

            total_blocos += 1

            info_bloco = bloco.get(
                "informacoes_relevantes",
                []
            )

            if isinstance(
                info_bloco,
                list
            ):

                total_fragmentos += len(
                    [
                        item
                        for item
                        in info_bloco
                        if isinstance(
                            item,
                            dict
                        )
                    ]
                )

            paragrafos_ollama = bloco.get(
                "paragrafos_ollama",
                []
            )

            if isinstance(
                paragrafos_ollama,
                list
            ):

                total_paragrafos_ollama += len(
                    [
                        paragrafo
                        for paragrafo
                        in paragrafos_ollama[:3]
                        if str(
                            paragrafo or ""
                        ).strip()
                    ]
                )

        print()
        print(
            "=============================================="
        )
        print(
            "BANCO JSON OFICIAL SALVO"
        )
        print(
            "=============================================="
        )
        print(
            "TEMA:",
            tema_original
        )
        print(
            "CATEGORIA:",
            categoria
        )
        print(
            "ARQUIVO:",
            ARQUIVO_BANCO
        )
        print(
            "BLOCOS:",
            total_blocos
        )
        print(
            "PARÁGRAFOS OLLAMA:",
            total_paragrafos_ollama
        )
        print(
            "FRAGMENTOS SELECIONADOS:",
            total_fragmentos
        )
        print(
            "SEGMENTOS:",
            len(
                segmentos_finais
            )
        )
        print(
            "TAGS:",
            len(
                dados_tema_final.get(
                    "tags",
                    []
                )
            )
        )
        print(
            "=============================================="
        )

        # ----------------------------------------------------
        # CHECK DOS 5 BLOCOS
        # ----------------------------------------------------

        print(
            "\nCHECK DOS 5 BLOCOS:"
        )

        for numero in range(
            1,
            6
        ):

            bloco = pagina_final.get(
                f"bloco_{numero}",
                {}
            )

            if not isinstance(
                bloco,
                dict
            ):

                print(
                    f"🔴 bloco_{numero}: inválido"
                )

                continue

            info = bloco.get(
                "informacoes_relevantes",
                []
            )

            if not isinstance(
                info,
                list
            ):

                info = []

            paragrafos_ollama = bloco.get(
                "paragrafos_ollama",
                []
            )

            quantidade_paragrafos = (
                len(
                    [
                        paragrafo
                        for paragrafo
                        in paragrafos_ollama[:3]
                        if str(
                            paragrafo or ""
                        ).strip()
                    ]
                )
                if isinstance(
                    paragrafos_ollama,
                    list
                )
                else 0
            )

            print(
                f"{'🟢' if info else '🔴'} "
                f"bloco_{numero}: "
                f"{len(info)} fragmentos selecionados | "
                f"{quantidade_paragrafos} parágrafos Ollama"
            )

        # ----------------------------------------------------
        # CHECK GLOBAL
        # ----------------------------------------------------

        print(
            "\nCHECK FINAL DE INTEGRIDADE:"
        )

        print(
            "informacoes_relevantes global:",
            "🔴 EXISTE"
            if "informacoes_relevantes"
            in dados_tema_final
            else "🟢 NÃO EXISTE"
        )

        print(
            "paragrafos legado:",
            "🔴 EXISTE"
            if any(
                "paragrafos"
                in pagina_final.get(
                    f"bloco_{numero}",
                    {}
                )
                for numero in range(
                    1,
                    6
                )
                if isinstance(
                    pagina_final.get(
                        f"bloco_{numero}",
                        {}
                    ),
                    dict
                )
            )
            else "🟢 NÃO EXISTE"
        )

        print(
            "imagens:",
            "🔴 EXISTE"
            if "imagens" in pagina_final
            else "🟢 NÃO EXISTE"
        )

        print(
            "caracteres:",
            "🔴 EXISTE"
            if "caracteres" in pagina_final
            else "🟢 NÃO EXISTE"
        )

        print(
            "status:",
            "🔴 EXISTE"
            if "status" in pagina_final
            else "🟢 NÃO EXISTE"
        )

        print(
            "subtitulo_listas:",
            "🔴 EXISTE"
            if "subtitulo_listas"
            in pagina_final
            else "🟢 NÃO EXISTE"
        )

        print(
            "5 blocos:",
            "🟢"
            if all(
                isinstance(
                    pagina_final.get(
                        f"bloco_{numero}"
                    ),
                    dict
                )
                for numero in range(
                    1,
                    6
                )
            )
            else "🔴"
        )

        print(
            "15 posições de parágrafos:",
            "🟢"
            if all(
                len(
                    pagina_final.get(
                        f"bloco_{numero}",
                        {}
                    ).get(
                        "paragrafos_ollama",
                        []
                    )
                ) == 3
                for numero in range(
                    1,
                    6
                )
                if isinstance(
                    pagina_final.get(
                        f"bloco_{numero}",
                        {}
                    ),
                    dict
                )
            )
            else "🔴"
        )

        print(
            "12 segmentos:",
            "🟢"
            if len(
                segmentos_finais
            ) == 12
            else f"⚠️ ({len(segmentos_finais)})"
        )

        print(
            "30 tags:",
            "🟢"
            if len(
                dados_tema_final.get(
                    "tags",
                    []
                )
            ) == 30
            else f"⚠️ ({len(dados_tema_final.get('tags', []))})"
        )

        print(
            "=============================================="
        )

        return True

    except Exception as erro:

        print()
        print(
            "=============================================="
        )
        print(
            "ERRO GRAVANDO BANCO"
        )
        print(
            "=============================================="
        )
        print(
            repr(erro)
        )
        print(
            "=============================================="
        )

        return False



# ============================================================
# INTERFACE
# ============================================================

def selecionar_banco():

    arquivo = filedialog.askopenfilename(
        filetypes=[
            ("JSON", "*.json")
        ]
    )

    if arquivo:

        entrada_banco.delete(
            0,
            tk.END
        )

        entrada_banco.insert(
            0,
            arquivo
        )


def selecionar_modelo():

    arquivo = filedialog.askopenfilename(
        filetypes=[
            ("DOCX", "*.docx")
        ]
    )

    if arquivo:

        entrada_modelo.delete(
            0,
            tk.END
        )

        entrada_modelo.insert(
            0,
            arquivo
        )



# ============================================================
# NOME DO SITE
# ============================================================

def atualizar_banco_automatico(event=None):

    nome = entrada_site.get().strip()

    if not nome:
        return

    nome_arquivo = (
        nome.lower()
        .replace(" ", "_")
        .replace("-", "_")
    )

    caminho = (
        rf"C:\Python\gerador-conteudo\banco\{nome_arquivo}.json"
    )

    entrada_banco.delete(0, tk.END)
    entrada_banco.insert(0, caminho)
    

# ============================================================
# GERAR MATERIAL INTERFACE
# ============================================================

def atualizar_progresso(valor, texto):

    # Enquanto um tema estiver ativo, nenhuma atualização pontual de
    # progresso pode substituir o cronômetro visível pelo tempo antigo
    # ou por uma mensagem genérica de "Processamento geral".
    try:
        if (
            globals().get("PROCESSAMENTO_ATIVO", False)
            and globals().get("INICIO_TEMA_TEMPO_ATIVO") is not None
            and globals().get("TEMA_TEMPO_ATIVO")
        ):
            inicio_ativo = globals().get("INICIO_TEMA_TEMPO_ATIVO")
            tema_ativo = globals().get("TEMA_TEMPO_ATIVO")
            tempo_ativo = max(0, time.time() - inicio_ativo)
            texto = (
                f"Tema: {tema_ativo} | Gerando conteúdo com IA | "
                f"Tempo: {formatar_tempo(tempo_ativo)}"
            )
    except Exception:
        pass

    janela.after(
        0,
        lambda: barra_progresso.config(
            value=valor
        )
    )

    janela.after(
        0,
        lambda t=texto: status_progresso.config(
            text=t
        )
    )

def iniciar_geracao():

    global IA_PROCESSANDO

    IA_PROCESSANDO = True

    thread = threading.Thread(
        target=gerar_material_interface,
        daemon=True
    )

    thread.start() 
    
    
 

# ============================================================
# INTERFACE - TEMPO ESTIMADO
# ============================================================

def formatar_tempo(segundos):

    minutos = int(segundos // 60)

    segundos = int(segundos % 60)

    return f"{minutos:02d}:{segundos:02d}"



# ============================================================
# TEMPO EM TEMPO REAL — CRONÔMETRO INDIVIDUAL POR TEMA
# ============================================================

INICIO_TEMA_TEMPO_ATIVO = None
TEMA_TEMPO_ATIVO = ""


def atualizar_tempo_execucao(inicio, tema):

    """Atualiza continuamente a linha de status do tema atual.

    O cronômetro não depende de IA_PROCESSANDO, pois deve continuar
    durante validações, montagem e gravação do conteúdo.
    """

    while (
        PROCESSAMENTO_ATIVO
        and INICIO_TEMA_TEMPO_ATIVO == inicio
        and TEMA_TEMPO_ATIVO == tema
    ):

        tempo = max(0, time.time() - inicio)

        texto = (
            f"Tema: {tema} | Gerando conteúdo com IA | "
            f"Tempo: {formatar_tempo(tempo)}"
        )

        try:
            janela.after(
                0,
                lambda t=texto: status_progresso.config(text=t)
            )
        except Exception:
            pass

        time.sleep(0.5)
        
        

# ============================================================
# COMPATIBILIDADE TEMPO TOTAL
# ============================================================

def atualizar_tempo_total(inicio, tema):

    print("FUNÇÃO atualizar_tempo_total CARREGADA")

    return atualizar_tempo_execucao(
        inicio,
        tema
    )





# ============================================================
# INTERFACE - SALVAR PROGRESSO SEGURO
# ============================================================

def salvar_progresso(
    tema,
    indice_tema,
    categoria,
    indice_categoria,
    status="processando",
    temas=None
):

    arquivo = ARQUIVO_PROGRESSO


    # Preservar a área de recuperação dos blocos ao atualizar
    # os campos normais de andamento.
    progresso_anterior = _ler_progresso_seguro()

    dados = {

        "temas": temas or [],

        "tema_atual": tema,

        "indice_tema": indice_tema,

        "categoria_atual": categoria,

        "indice_categoria": indice_categoria,

        "status": status,

        "ultima_atualizacao": time.strftime(
            "%d/%m/%Y %H:%M:%S"
        )

    }


    # Preservar todas as áreas de recuperação. Antes, esta rotina copiava
    # somente recuperacao_conteudo e apagava respostas_globais_ollama na
    # próxima atualização normal do progresso — exatamente o que impedia
    # recuperar a resposta bruta após a mensagem de preservação.
    for chave_recuperacao in (
        "recuperacao_conteudo",
        "respostas_globais_ollama",
        "respostas_parciais_ollama",
    ):
        valor_recuperacao = progresso_anterior.get(chave_recuperacao)
        if isinstance(valor_recuperacao, dict) and valor_recuperacao:
            dados[chave_recuperacao] = valor_recuperacao

    try:
    
        os.makedirs(
            os.path.dirname(arquivo),
            exist_ok=True
        )


        arquivo_temp = arquivo + ".tmp"


        with open(
            arquivo_temp,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                dados,
                f,
                ensure_ascii=False,
                indent=4
            )


        os.replace(
            arquivo_temp,
            arquivo
        )


        print()
        print("==============================")
        print("PROGRESSO SALVO")
        print("==============================")
        print({k: v for k, v in dados.items() if k != "recuperacao_conteudo"})
        recuperacao_log = dados.get("recuperacao_conteudo", {})
        if isinstance(recuperacao_log, dict):
            tema_id_log = _tema_recuperacao_id(tema)
            registro_log = recuperacao_log.get(tema_id_log, {})
            blocos_log = registro_log.get("blocos", {}) if isinstance(registro_log, dict) else {}
            print("BLOCOS PRESERVADOS PARA RECUPERAÇÃO:", len(blocos_log) if isinstance(blocos_log, dict) else 0, "/ 5")


    except Exception as e:

        print()
        print("ERRO AO SALVAR PROGRESSO:")
        print(e)


# ============================================================
# LIMPAR PROGRESSO ANTIGO
# ============================================================

def limpar_progresso():

    arquivo = ARQUIVO_PROGRESSO


    if os.path.exists(arquivo):

        os.remove(arquivo)

        print()
        print("==============================")
        print("PROGRESSO ANTIGO REMOVIDO")
        print("==============================")


    else:

        print()
        print("==============================")
        print("NENHUM PROGRESSO ANTIGO")
        print("==============================")
        

# ============================================================
# VERIFICAR NOVO CONJUNTO DE TEMAS
# ============================================================

def verificar_novo_conjunto(temas_atuais):

    arquivo = ARQUIVO_PROGRESSO


    if not os.path.exists(arquivo):

        print()
        print("==============================")
        print("PRIMEIRO PROCESSAMENTO")
        print("==============================")

        return True


    try:

        with open(
            arquivo,
            "r",
            encoding="utf-8"
        ) as f:

            progresso = json.load(f)


        temas_salvos = progresso.get(
            "temas",
            []
        )


        if temas_salvos != temas_atuais:

            print()
            print("==============================")
            print("NOVO CONJUNTO DE TEMAS")
            print("==============================")


            limpar_progresso()


            print()
            print("==============================")
            print("INICIANDO NOVO PROCESSAMENTO")
            print("==============================")


            return True



        print()
        print("==============================")
        print("MESMO CONJUNTO IDENTIFICADO")
        print("RETOMANDO PROCESSAMENTO")
        print("==============================")


        return False



    except Exception as e:

        print()
        print("ERRO AO VERIFICAR CONJUNTO")
        print(e)

        return True


# ============================================================
# INTERFACE - SALVAR PROGRESSO RETORNAR DE ONDE PAROU
# ============================================================

def carregar_progresso():

    arquivo = r"C:\Python\gerador-conteudo\banco\progresso.json"


    if not os.path.exists(arquivo):

        return None


    with open(
        arquivo,
        "r",
        encoding="utf-8"
    ) as f:

        return json.load(f)
        

# ============================================================
# VALIDAR CONTEÚDO RELEVANTE
# ============================================================

def validar_conteudo_relevante(
    texto,
    tema
):

    if not texto:
        return False


    # ========================================================
    # 01. NORMALIZA TEXTO
    # ========================================================

    def normalizar(valor):

        valor = valor.lower()

        valor = unicodedata.normalize(
            "NFKD",
            valor
        ).encode(
            "ASCII",
            "ignore"
        ).decode(
            "utf-8"
        )

        return valor



    texto_limpo = normalizar(texto)

    tema_limpo = normalizar(tema)



    # ========================================================
    # 02. PALAVRAS PRINCIPAIS DO TEMA
    # ========================================================
    
    tema_limpo = (
        tema_limpo
        .replace("_", " ")
        .replace("-", " ")
    )
    
    
    palavras_tema = tema_limpo.split()
    
    
    encontrados_tema = 0
    
    
    for palavra in palavras_tema:
    
        if palavra in texto_limpo:
    
            encontrados_tema += 1



    # ========================================================
    # 03. VALIDACAO FLEXIVEL DO TEMA
    # ========================================================

    minimo_necessario = 1


    if len(palavras_tema) >= 3:

        minimo_necessario = 2



    if encontrados_tema < minimo_necessario:


        print(
            "CONTEUDO REJEITADO - TEMA AUSENTE:",
            tema
        )


        return False

    print(
        "CONTEUDO RELEVANTE APROVADO:",
        tema
    )


    return True


# ============================================================
# VALIDAR CONTEÚDO TÉCNICO — PYTHON
# ============================================================
#
# O Ollama NÃO participa da aprovação das fontes.
#
# A decisão é feita por regras locais:
#
# - tamanho;
# - aderência ao tema;
# - densidade técnica;
# - ausência de lixo estrutural;
# - ausência de navegação;
# - ausência de propaganda;
# - ausência de contato/comercial;
# - presença de conteúdo explicativo.
#
# A seleção rigorosa continuará sendo feita posteriormente
# por selecionar_informacoes_relevantes().
#
# ============================================================

def validar_conteudo_tecnico(
    texto,
    tema
):

    texto = str(
        texto or ""
    ).strip()

    tema = str(
        tema or ""
    ).strip()

    if not texto:

        print(
            "CONTEÚDO TÉCNICO REJEITADO — VAZIO"
        )

        return False

    if len(texto) < 500:

        print(
            "CONTEÚDO TÉCNICO REJEITADO — MUITO CURTO"
        )

        return False

    # --------------------------------------------------------
    # NORMALIZAÇÃO
    # --------------------------------------------------------

    texto_normalizado = (
        unicodedata.normalize(
            "NFD",
            texto.casefold()
        )
    )

    texto_normalizado = "".join(

        caractere

        for caractere
        in texto_normalizado

        if unicodedata.category(
            caractere
        ) != "Mn"

    )

    tema_normalizado = (
        unicodedata.normalize(
            "NFD",
            tema.casefold()
        )
    )

    tema_normalizado = "".join(

        caractere

        for caractere
        in tema_normalizado

        if unicodedata.category(
            caractere
        ) != "Mn"

    )

    # --------------------------------------------------------
    # 01. TEMA
    # --------------------------------------------------------

    palavras_tema = [

        palavra

        for palavra
        in re.findall(
            r"\b[a-z0-9]{3,}\b",
            tema_normalizado
        )

    ]

    palavras_tema = [

        palavra

        for palavra
        in palavras_tema

        if palavra not in {

            "para",
            "com",
            "sem",
            "por",
            "das",
            "dos",
            "uma",
            "uns",
            "uma"

        }

    ]

    encontrados = 0

    for palavra in palavras_tema:

        if palavra in texto_normalizado:

            encontrados += 1

    if palavras_tema:

        proporcao_tema = (
            encontrados
            /
            len(palavras_tema)
        )

    else:

        proporcao_tema = 0

    if (
        len(palavras_tema) >= 2
        and
        encontrados == 0
    ):

        print(
            "CONTEÚDO TÉCNICO REJEITADO — TEMA AUSENTE"
        )

        return False

    if (
        len(palavras_tema) >= 3
        and
        proporcao_tema < 0.50
    ):

        print(
            "CONTEÚDO TÉCNICO REJEITADO — BAIXA ADERÊNCIA"
        )

        return False

    # --------------------------------------------------------
    # 02. LIXO DE NAVEGAÇÃO
    # --------------------------------------------------------

    marcadores_navegacao = [

        "skip to main content",
        "home",
        "menu",
        "login",
        "cadastro",
        "entrar",
        "cookies",
        "política de privacidade",
        "politica de privacidade",
        "siga-nos",
        "compartilhe",
        "linkedin",
        "facebook",
        "instagram",
        "twitter"

    ]

    ocorrencias_navegacao = sum(

        1

        for marcador
        in marcadores_navegacao

        if marcador
        in texto_normalizado

    )

    if ocorrencias_navegacao >= 4:

        print(
            "CONTEÚDO TÉCNICO REJEITADO — NAVEGAÇÃO"
        )

        return False

    # --------------------------------------------------------
    # 03. PROPAGANDA / COMERCIAL
    # --------------------------------------------------------

    marcadores_comerciais = [

        "entre em contato",
        "fale conosco",
        "solicite um orçamento",
        "solicite um orcamento",
        "compre agora",
        "adquira agora",
        "saiba mais",
        "clique aqui",
        "nossos produtos",
        "nossa empresa",
        "melhor fornecedor",
        "preço sob consulta",
        "preco sob consulta"

    ]

    ocorrencias_comerciais = sum(

        1

        for marcador
        in marcadores_comerciais

        if marcador
        in texto_normalizado

    )

    if ocorrencias_comerciais >= 2:

        print(
            "CONTEÚDO TÉCNICO REJEITADO — COMERCIAL"
        )

        return False

    # --------------------------------------------------------
    # 04. URL / E-MAIL / TELEFONE
    # --------------------------------------------------------

    if re.search(
        r"https?://|www\.",
        texto,
        re.IGNORECASE
    ):

        return False

    if re.search(
        r"\b[\w.+-]+@[\w.-]+\.[a-z]{2,}\b",
        texto,
        re.IGNORECASE
    ):

        return False

    # --------------------------------------------------------
    # 05. TERMOS TÉCNICOS
    # --------------------------------------------------------

    termos_tecnicos = [

        "funcionamento",
        "processo",
        "sistema",
        "equipamento",
        "componente",
        "material",
        "aplicação",
        "aplicacoes",
        "instalação",
        "manutenção",
        "dimensionamento",
        "desempenho",
        "eficiência",
        "eficiencia",
        "rendimento",
        "pressão",
        "pressao",
        "vazão",
        "vazao",
        "temperatura",
        "operação",
        "operacao",
        "fabricação",
        "fabricacao",
        "montagem",
        "durabilidade",
        "confiabilidade",
        "segurança",
        "seguranca",
        "especificação",
        "especificacao"

    ]

    termos_encontrados = 0

    for termo in termos_tecnicos:

        if termo in texto_normalizado:

            termos_encontrados += 1

    # --------------------------------------------------------
    # 06. FRASES
    # --------------------------------------------------------

    quantidade_frases = len(

        re.findall(
            r"[.!?]",
            texto
        )

    )

    if quantidade_frases < 2:

        print(
            "CONTEÚDO TÉCNICO REJEITADO — SEM TEXTO EXPLICATIVO"
        )

        return False

    # --------------------------------------------------------
    # 07. APROVAÇÃO
    #
    # Não exigimos quantidade fixa de termos técnicos.
    # Uma fonte técnica pode explicar um conceito sem
    # utilizar exatamente essas palavras.
    # --------------------------------------------------------

    if (
        termos_encontrados >= 2
        and
        quantidade_frases >= 2
    ):

        print(
            "CONTEÚDO TÉCNICO APROVADO — PYTHON"
        )

        return True

    # --------------------------------------------------------
    # CONTEÚDO COM POUCOS TERMOS TÉCNICOS
    #
    # Se o tema aparece claramente e o texto é suficientemente
    # desenvolvido, deixamos a seleção posterior decidir.
    # --------------------------------------------------------

    if (
        proporcao_tema >= 0.50
        and
        len(texto.split()) >= 100
    ):

        print(
            "CONTEÚDO TÉCNICO APROVADO — ADERÊNCIA"
        )

        return True

    print(
        "CONTEÚDO TÉCNICO REJEITADO — BAIXA DENSIDADE"
    )

    return False
    

# ============================================================
# NORMALIZAR TEXTO PARA VALIDAÇÃO MEAD
# ============================================================

def normalizar_texto(texto):

    texto = texto.lower()


    texto = unicodedata.normalize(
        "NFD",
        texto
    )


    texto = "".join(
        c for c in texto
        if unicodedata.category(c) != "Mn"
    )


    return texto




# ============================================================
# VALIDAR PROTAGONISTA MAPA MEAD
# ============================================================

def validar_protagonista_mead(
    resposta,
    tema
):

    resposta_limpa = normalizar_texto(
        resposta
    )


    tema_limpo = normalizar_texto(
        tema
    )


    if "protagonista" not in resposta_limpa:

        return False


    inicio = resposta_limpa.find(
        "protagonista"
    )


    trecho = resposta_limpa[
        inicio:inicio+300
    ]


    palavras_tema = tema_limpo.split()


    for palavra in palavras_tema:

        if palavra not in trecho:

            return False


    return True


# ============================================================
# VALIDAR CONTEXTO TÉCNICO MAPA MEAD
# ============================================================

def validar_contexto_tecnico_mead(
    resposta,
    tema
):


    resposta_limpa = normalizar_texto(
        resposta
    )


    bloqueados = {


        "valvulas macho": [

            "thread macho",
            "rosca macho",
            "rosca femea",
            "conexao macho",
            "conexao femea",
            "roscamento",
            "encaixa em thread"

        ]

    }


    tema_normalizado = normalizar_texto(
        tema
    )


    if tema_normalizado in bloqueados:


        for termo in bloqueados[tema_normalizado]:


            if termo in resposta_limpa:


                print()
                print("==============================")
                print("MAPA MEAD BLOQUEADO")
                print("==============================")

                print(
                    "CONTEXTO INCOMPATÍVEL:",
                    termo
                )


                return False



    return True





# ============================================================
# GERAR E SALVAR MAPA MEAD
# ============================================================

def gerar_e_salvar_mapa_mead(
    tema,
    textos
):

    # =====================================
    # NORMALIZAR TEXTOS RECEBIDOS
    # PRESERVANDO METADADOS
    # =====================================
    
    textos_normalizados = []

    for item in textos:

        if isinstance(item, dict):

            textos_normalizados.append(
                {
                    "texto": item.get("texto", ""),
                    "url": item.get("url", ""),
                    "tipo": item.get("tipo", "texto")
                }
            )

        else:

            textos_normalizados.append(
                {
                    "texto": str(item),
                    "url": "",
                    "tipo": "texto"
                }
            )

    textos = textos_normalizados


    print()
    print("==============================")
    print("GERANDO MAPA MEAD")
    print("==============================")


    if len(textos) < 3:

        print()
        print("==============================")
        print("COLETA INSUFICIENTE PARA MAPA MEAD")
        print("==============================")

        return ""


    mapa_mead = gerar_mapa_mead(
        tema,
        textos
    )
    
    
    print()
    print("==============================")
    print("MAPA MEAD RECEBIDO")
    print("==============================")
    
    print("MAPA MEAD:", "válido" if mapa_mead else "vazio",
          "|", len(str(mapa_mead)), "caracteres")


    if mapa_mead:
    
    
        print()
        print("==============================")
        print("MAPA MEAD RECEBIDO PARA VALIDAÇÃO")
        print("==============================")
    
        print(
            "TAMANHO:",
            len(str(mapa_mead))
        )
    
    
    
        resultado_protagonista = validar_protagonista_mead(
            mapa_mead,
            tema
        )
    
    
        print(
            "VALIDAÇÃO PROTAGONISTA:",
            resultado_protagonista
        )
    
    
    
        if not resultado_protagonista:
        
            print()
            print("==============================")
            print("MAPA MEAD BLOQUEADO")
            print("PROTAGONISTA NÃO CONFERE")
            print("==============================")
        
            return ""
    
    
    
    
        resultado_contexto = validar_contexto_tecnico_mead(
            mapa_mead,
            tema
        )
    
    
        print(
            "VALIDAÇÃO CONTEXTO:",
            resultado_contexto
        )
    
    
    
        if not resultado_contexto:
        
            print()
            print("==============================")
            print("MAPA MEAD BLOQUEADO")
            print("CONTEXTO TÉCNICO INVÁLIDO")
            print("==============================")
        
            return ""
    
    
    
    
        print()
        print("==============================")
        print("CHAMANDO SALVAR BANCO")
        print("==============================")
    
    
    
        salvar_banco(
            tema,
            "mapa_mead",
            mapa_mead
        )
    
    
    
        print()
        print("==============================")
        print("MAPA MEAD SALVO")
        print("==============================")
    
    
        return mapa_mead




    print()
    print("==============================")
    print("MAPA MEAD INVALIDO")
    print("==============================")


    return ""



# ============================================================
# BUSCAR TEMAS RELACIONADOS NO BANCO
# ============================================================

def buscar_temas_relacionados(
    tema,
    mapa_mead=None
):


    relacionados = []


    # ========================================================
    # 01. PALAVRAS BASE DA BUSCA
    # ========================================================

    palavras_tema = set(
        tema.lower().split()
    )


    # ========================================================
    # 02. USAR PROTAGONISTA DO MAPA MEAD
    # ========================================================

    if mapa_mead:


        texto_mapa = mapa_mead.lower()


        if "protagonista:" in texto_mapa:


            try:

                protagonista = texto_mapa.split(
                    "protagonista:",
                    1
                )[1]


                protagonista = protagonista.split(
                    "\n",
                    1
                )[0].strip()


                if protagonista:


                    palavras_tema.update(
                        protagonista.split()
                    )


                    print()
                    print("==============================")
                    print("PROTAGONISTA MEAD USADO NA BUSCA")
                    print("==============================")

                    print(
                        protagonista
                    )


            except Exception:


                pass



    if not os.path.exists(
        ARQUIVO_BANCO
    ):

        return relacionados



    try:

        with open(
            ARQUIVO_BANCO,
            "r",
            encoding="utf-8"
        ) as arquivo:

            banco = json.load(
                arquivo
            )


    except:


        return relacionados



    # ========================================================
    # 03. BANCO PRECISA SER DICIONÁRIO
    # ========================================================

    if not isinstance(
        banco,
        dict
    ):


        return relacionados



    for tema_banco, dados in banco.items():


        if not isinstance(
            tema_banco,
            str
        ):

            continue



        tema_banco = tema_banco.lower().strip()



        if not tema_banco:

            continue



    # ========================================================
    # 04. IGNORAR O PRÓPRIO TEMA
    # ========================================================

        if tema.lower().strip() == tema_banco:

            continue



        palavras_banco = set(
            tema_banco.split()
        )



        intersecao = palavras_tema.intersection(
            palavras_banco
        )



    # ========================================================
    # 05. COMPATIBILIDADE MÍNIMA
    # ========================================================

        if len(intersecao) >= 1:


            relacionados.append(
                tema_banco
            )



    return relacionados
    


# ============================================================
# PESQUISA MAPA NO BANCO
# ============================================================

def obter_dados_banco(tema):

    print()
    print("==============================")
    print("BUSCANDO MAPA NO BANCO")
    print("==============================")
    print("TEMA:", tema)

    resultado = {

        "existe": False,

        "mapa_mead": "",

        "conteudo": "",

        "bruto": []

    }


    # ========================================================
    # 01. VERIFICAR EXISTÊNCIA DO BANCO
    # ========================================================

    if not os.path.exists(
        ARQUIVO_BANCO
    ):

        print()
        print("==============================")
        print("BANCO NÃO ENCONTRADO")
        print("==============================")

        return resultado


    # ========================================================
    # 02. LER BANCO
    # ========================================================

    try:

        with open(
            ARQUIVO_BANCO,
            "r",
            encoding="utf-8"
        ) as arquivo:

            banco = json.load(
                arquivo
            )


    except Exception as e:

        print()
        print("==============================")
        print("ERRO LENDO BANCO")
        print("==============================")

        print(e)

        return resultado


    # ========================================================
    # 03. NORMALIZAR TEMA PROCURADO
    # ========================================================

    tema_normalizado = normalizar_tema_chave(
        tema
    )


    # ========================================================
    # 04. FUNÇÃO INTERNA PARA EXTRAIR MAPA
    # ========================================================

    def extrair_mapa_texto(mapa):

        if not mapa:

            return ""


    # ========================================================
    # 05. MAPA JÁ É TEXTO
    # ========================================================

        if isinstance(
            mapa,
            str
        ):

            return mapa.strip()


    # ========================================================
    # 06. MAPA É DICIONÁRIO
    # ========================================================

        if isinstance(
            mapa,
            dict
        ):

            # -----------------------------
            # CAMPO TEXTO
            # -----------------------------

            texto = mapa.get(
                "texto",
                ""
            )

            if isinstance(
                texto,
                str
            ):

                if texto.strip():

                    return texto.strip()


            # -----------------------------
            # CAMPO GERADO
            # -----------------------------

            gerado = mapa.get(
                "gerado",
                ""
            )


            if isinstance(
                gerado,
                str
            ):

                if gerado.strip():

                    return gerado.strip()


            # -----------------------------
            # GERADO COMO DICIONÁRIO
            # -----------------------------

            if isinstance(
                gerado,
                dict
            ):

                texto_gerado = gerado.get(
                    "texto",
                    ""
                )

                if isinstance(
                    texto_gerado,
                    str
                ):

                    if texto_gerado.strip():

                        return texto_gerado.strip()


            # -----------------------------
            # ÚLTIMO RECURSO
            # -----------------------------

            try:

                texto_dict = str(
                    mapa
                )

                if texto_dict.strip():

                    return texto_dict.strip()

            except Exception:

                pass


        return ""



    # ========================================================
    # 07. FUNÇÃO INTERNA PARA VALIDAR MAPA
    # ========================================================
    
    def validar_mapa_do_tema(
        mapa_texto
    ):
    
        if not mapa_texto:
    
            return False
    
    
    # ========================================================
    # 08. VALIDAR PROTAGONISTA
    # ========================================================
    
        valido_protagonista = validar_protagonista_mead(
            mapa_texto,
            tema
        )
    
    
        print()
        print("==============================")
        print("VALIDAÇÃO DO MAPA MEAD")
        print("==============================")
    
    
        print(
            "TEMA:",
            tema
        )
    
    
        print(
            "PROTAGONISTA VÁLIDO:",
            valido_protagonista
        )
    
    
    # ========================================================
    # 09. PROTAGONISTA INVÁLIDO
    # ========================================================
    
        if not valido_protagonista:
    
            print()
            print("==============================")
            print("MAPA MEAD IGNORADO")
            print("==============================")
    
    
            print(
                "MOTIVO: PROTAGONISTA NÃO CONFERE"
            )
    
    
            return False
    
    
    # ========================================================
    # 10. MAPA VALIDADO
    # ========================================================
    
        print()
        print("==============================")
        print("MAPA MEAD VALIDADO")
        print("==============================")
    
    
        print(
            "TEMA:",
            tema
        )
    
    
        print(
            "MAPA:",
            len(mapa_texto),
            "caracteres"
        )
    
    
        return True




    # ========================================================
    # 11. BANCO EM FORMATO LISTA
    # ========================================================

    if isinstance(
        banco,
        list
    ):

        for item in banco:

            if not isinstance(
                item,
                dict
            ):

                continue


            tema_banco = item.get(
                "tema",
                ""
            )


            if not isinstance(
                tema_banco,
                str
            ):

                continue


            tema_banco_normalizado = normalizar_tema_chave(
                tema_banco
            )


    # ========================================================
    # 12. LOCALIZAR TEMA
    # ========================================================

            if tema_banco_normalizado != tema_normalizado:

                continue


            resultado["existe"] = True


    # ========================================================
    # 13. CONTEÚDO
    # ========================================================

            resultado["conteudo"] = item.get(
                "conteudo_completo",
                ""
            )


    # ========================================================
    # 14. BRUTO
    # ========================================================

            resultado["bruto"] = item.get(
                "bruto",
                []
            )


    # ========================================================
    # 15. MAPA MEAD
    # ========================================================

            mapa_original = item.get(
                "mapa_mead",
                ""
            )


            print()
            print("==============================")
            print("MAPA BRUTO ENCONTRADO")
            print("==============================")

            print(
                "TIPO:",
                type(mapa_original)
            )


            mapa_texto = extrair_mapa_texto(
                mapa_original
            )


            print()
            print("==============================")
            print("MAPA MEAD EXTRAÍDO")
            print("==============================")

            print(
                mapa_texto[:1000]
            )


    # ========================================================
    # 16. VALIDAR MAPA
    # ========================================================

            if mapa_texto:

                mapa_valido = validar_mapa_do_tema(
                    mapa_texto
                )


                if mapa_valido:

                    resultado["mapa_mead"] = mapa_texto

                else:

                    resultado["mapa_mead"] = ""


            else:

                print()
                print("==============================")
                print("MAPA MEAD NÃO ENCONTRADO")
                print("==============================")


                resultado["mapa_mead"] = ""


    # ========================================================
    # 17. DEBUG FINAL
    # ========================================================

            print()
            print("==============================")
            print("MAPA FINAL RETORNADO PELO BANCO")
            print("==============================")

            print(
                resultado["mapa_mead"][:1000]
            )

            print(
                "TAMANHO MAPA:",
                len(
                    resultado["mapa_mead"]
                )
            )


            break


    # ========================================================
    # 18. BANCO EM FORMATO DICIONÁRIO
    # ========================================================

    elif isinstance(
        banco,
        dict
    ):


        dados = None

        tema_encontrado = None


        # =================================
        # PRIMEIRA TENTATIVA:
        # CHAVE EXATA
        # =================================

        if tema in banco:

            dados = banco.get(
                tema
            )

            tema_encontrado = tema


        # =================================
        # SEGUNDA TENTATIVA:
        # CHAVE NORMALIZADA
        # =================================

        else:

            for chave, valor in banco.items():

                if not isinstance(
                    chave,
                    str
                ):

                    continue


                chave_normalizada = normalizar_tema_chave(
                    chave
                )


                if chave_normalizada == tema_normalizado:

                    dados = valor

                    tema_encontrado = chave

                    break


    # ========================================================
    # 19. TEMA NÃO ENCONTRADO
    # ========================================================

        if not isinstance(
            dados,
            dict
        ):

            return resultado


        resultado["existe"] = True


    # ========================================================
    # 20. CONTEÚDO
    # ========================================================

        resultado["conteudo"] = dados.get(
            "conteudo_completo",
            ""
        )


        # =================================
        # COMPATIBILIDADE COM ESTRUTURAS
        # ANTIGAS DE CATEGORIAS
        # =================================

        if not resultado["conteudo"]:

            categorias = dados.get(
                "categorias",
                {}
            )


            if isinstance(
                categorias,
                dict
            ):

                resultado["conteudo"] = categorias.get(
                    "conteudo_completo",
                    ""
                )


    # ========================================================
    # 21. BRUTO
    # ========================================================

        resultado["bruto"] = dados.get(
            "bruto",
            []
        )


    # ========================================================
    # 22. MAPA MEAD
    # ========================================================

        mapa_original = dados.get(
            "mapa_mead",
            ""
        )


        print()
        print("==============================")
        print("MAPA BRUTO ENCONTRADO")
        print("==============================")

        print(
            "TEMA ENCONTRADO:",
            tema_encontrado
        )

        print(
            "TIPO:",
            type(mapa_original)
        )


        mapa_texto = extrair_mapa_texto(
            mapa_original
        )


        print()
        print("==============================")
        print("MAPA MEAD EXTRAÍDO")
        print("==============================")

        print(
            mapa_texto[:1000]
        )


    # ========================================================
    # 23. VALIDAR MAPA DO PRÓPRIO TEMA
    # ========================================================

        if mapa_texto:

            mapa_valido = validar_mapa_do_tema(
                mapa_texto
            )


            if mapa_valido:

                resultado["mapa_mead"] = mapa_texto

            else:

                resultado["mapa_mead"] = ""


        else:

            print()
            print("==============================")
            print("MAPA MEAD NÃO ENCONTRADO")
            print("==============================")


            resultado["mapa_mead"] = ""


    # ========================================================
    # 24. FORMATO DE BANCO DESCONHECIDO
    # ========================================================

    else:

        print()
        print("==============================")
        print("FORMATO DE BANCO NÃO SUPORTADO")
        print("==============================")

        print(
            "TIPO:",
            type(banco)
        )


    # ========================================================
    # 25. RESULTADO FINAL
    # ========================================================

    print()
    print("==============================")
    print("DADOS FINAIS DO BANCO")
    print("==============================")

    print(
        "TEMA:",
        tema
    )

    print(
        "EXISTE:",
        resultado["existe"]
    )

    print(
        "MAPA MEAD:",
        bool(
            resultado["mapa_mead"]
        )
    )

    print(
        "CONTEÚDO:",
        bool(
            resultado["conteudo"]
        )
    )

    print(
        "BRUTO:",
        len(
            resultado["bruto"]
        )
        if isinstance(
            resultado["bruto"],
            list
        )
        else type(
            resultado["bruto"]
        )
    )


    return resultado

    # ========================================================
    # 26. GRUPOS PRINCIPAIS
    # ========================================================

GRUPOS_PRINCIPAIS = {


    "limpeza_industrial":[

        "lavadora",
        "lavadora industrial",
        "alta pressão",
        "hidrojateadora",
        "limpeza industrial",
        "jato de água"

    ],


    "componentes_mecanicos":[

        "terminal rotular",
        "terminal de rótula",
        "rótula",
        "rotula",
        "rótulas esféricas",
        "rolamento esférico",
        "mancal",
        "bucha esférica",
        "spherical plain bearing",
        "rod end"

    ],


    "construcao":[

        "argamassa",
        "revestimento",
        "selagem",
        "selagem corta fogo",
        "corta fogo",
        "firestop",
        "passive fire protection"

    ],


    "eletrica":[

        "painel elétrico",
        "quadro elétrico",
        "cabo",
        "disjuntor",
        "automação"

    ],


    "hidraulica":[

        "válvula",
        "valvula",
        "bomba",
        "hidráulica",
        "hidraulico",
        "tubulação"

    ]

}



    # ========================================================
    # 27. SINAIS TÉCNICOS GERAIS
    # ========================================================

SINAIS_UNIVERSAIS = [

    "modelo",
    "tipo",
    "aplicação",
    "aplicacao",
    "instalação",
    "instalacao",
    "manutenção",
    "manutencao",
    "norma",
    "certificação",
    "certificacao",
    "especificação",
    "especificacao",
    "dimensão",
    "dimensao",
    "material",
    "aço",
    "aco",
    "inox",
    "temperatura",
    "pressão",
    "pressao",
    "capacidade",
    "resistência",
    "resistencia"

]




    # ========================================================
    # 28. INDICADORES TÉCNICOS
    # ========================================================

INDICADORES_TECNICOS = [

    "medida",
    "mm",
    "cm",
    "kg",
    "mpa",
    "bar",
    "°c",
    "norma",
    "componente",
    "componentes",
    "processo",
    "industrial"

]



# ============================================================
# GRUPO PRINCIPAL
# ============================================================

def pertence_ao_grupo_principal(
    tema,
    relacionado
):

    if not tema or not relacionado:

        return False


    tema_lower = tema.lower()

    relacionado_lower = relacionado.lower()


    grupo_tema = None


    # ========================================================
    # 01. IDENTIFICAR GRUPO DO TEMA
    # ========================================================

    for grupo, palavras in GRUPOS_PRINCIPAIS.items():

        for palavra in palavras:

            if palavra in tema_lower:

                grupo_tema = grupo

                break


        if grupo_tema:

            break


    # Tema sem grupo conhecido

    if not grupo_tema:

        return False


    # ========================================================
    # 02. RELAÇÃO DIRETA
    # ========================================================

    for palavra in GRUPOS_PRINCIPAIS[grupo_tema]:

        if palavra in relacionado_lower:

            return True


    # ========================================================
    # 03. RELAÇÃO TÉCNICA FLEXÍVEL
    # ========================================================

    sinais = 0


    for sinal in SINAIS_UNIVERSAIS:

        if sinal in relacionado_lower:

            sinais += 1


    for indicador in INDICADORES_TECNICOS:

        if indicador in relacionado_lower:

            sinais += 1


    # Se tiver sinais técnicos suficientes,
    # aceita mesmo sem palavra exata do grupo

    if sinais >= 2:

        return True


    return False



# ============================================================
# NORMALIZAR TEXTO SEM ACENTOS
# ============================================================

def normalizar_sem_acento(texto):

    if not texto:
        return ""


    texto = unicodedata.normalize(
        "NFD",
        str(texto)
    )


    texto = "".join(
        caractere
        for caractere in texto
        if unicodedata.category(caractere) != "Mn"
    )


    return texto.upper()
    
    

# ============================================================
# COLETAR SITES INFORMADOS PELO USUÁRIO
# ============================================================

def coletar_sites_consulta():

    resultados = []

    try:

        if "txt_sites" not in globals():
            return resultados

        sites = txt_sites.get(
            "1.0",
            tk.END
        ).strip()

        if not sites:
            return resultados

        print()
        print("==========================================")
        print("SITES INFORMADOS PELO USUÁRIO")
        print("==========================================")

        linhas = [
            linha.strip()
            for linha in sites.splitlines()
            if linha.strip()
        ]

        urls_vistas = set()

        for url in linhas:

            if not url.startswith(
                ("http://", "https://")
            ):

                url = "https://" + url

            url = url.strip()

            if url in urls_vistas:
                continue

            urls_vistas.add(url)

            print()
            print("SITE PARA CONSULTA:")
            print(url)

            resultado = coletar_pagina(
                url
            )

            if not resultado:
                print(
                    "SITE NÃO COLETADO:",
                    url
                )
                continue

            texto = str(
                resultado.get(
                    "texto",
                    ""
                )
            ).strip()

            if not texto:
                continue

            resultados.append(
                resultado
            )

            print(
                "SITE COLETADO:",
                url
            )

            print(
                "TIPO:",
                resultado.get(
                    "tipo",
                    ""
                )
            )

            print(
                "CARACTERES:",
                len(texto)
            )

        print()
        print(
            "TOTAL DE SITES UTILIZADOS:",
            len(resultados)
        )

        print("==========================================")

        return resultados

    except Exception as erro:

        print()
        print(
            "ERRO AO COLETAR SITES INFORMADOS:"
        )
        print(erro)

        return resultados


# ============================================================
# ANALISAR SITE INFORMADO COMO REFERÊNCIA PRINCIPAL
# ============================================================

def analisar_site_referencia(
    tema,
    sites_consulta
):

    resultado = {
        "disponivel": False,
        "urls": [],
        "conteudo": "",
        "contexto": ""
    }

    try:

        if not sites_consulta:
            return resultado

        textos_site = []

        for item in sites_consulta:

            if not isinstance(
                item,
                dict
            ):
                continue

            texto = str(
                item.get(
                    "texto",
                    ""
                )
            ).strip()

            url = str(
                item.get(
                    "url",
                    ""
                )
            ).strip()

            if not texto:
                continue

            if url:
                resultado["urls"].append(
                    url
                )

            textos_site.append(
                texto
            )

        if not textos_site:
            return resultado

        conteudo_site = "\n\n".join(
            textos_site
        )

        resultado["conteudo"] = (
            conteudo_site
        )

        resultado["disponivel"] = True

        contexto_maximo = 12000

        resultado["contexto"] = (
            conteudo_site[
                :contexto_maximo
            ]
        )

        print()
        print("==========================================")
        print("ANÁLISE DO SITE COMO REFERÊNCIA PRINCIPAL")
        print("==========================================")

        print(
            "TEMA:",
            tema
        )

        print(
            "SITES:",
            len(resultado["urls"])
        )

        print(
            "CARACTERES DO SITE:",
            len(conteudo_site)
        )

        print(
            "CONTEXTO DE REFERÊNCIA:",
            len(resultado["contexto"])
        )

        print("==========================================")

        return resultado

    except Exception as erro:

        print()
        print(
            "ERRO AO ANALISAR SITE DE REFERÊNCIA:"
        )
        print(erro)

        return resultado    

# ============================================================
# GERAR MATERIAL INTERFACE
# ============================================================

def gerar_material_interface():

    global ARQUIVO_BANCO
    global IA_PROCESSANDO
    global PROCESSAMENTO_ATIVO
    global INICIO_TEMA_TEMPO_ATIVO
    global TEMA_TEMPO_ATIVO

    try:

        PROCESSAMENTO_ATIVO = True
        PAGINAS_EM_PROCESSAMENTO.clear()
        IA_PROCESSANDO = True

        progresso = carregar_progresso()

        inicio_total = time.time()
        INICIO_TEMA_TEMPO_ATIVO = None
        TEMA_TEMPO_ATIVO = ""

        if progresso:

            print()
            print("==============================")
            print("RETOMANDO PROCESSAMENTO")
            print("==============================")

            print(progresso)

        atualizar_progresso(
            0,
            "Iniciando..."
        )

        ARQUIVO_BANCO = entrada_banco.get().strip()

        if not ARQUIVO_BANCO:

            ARQUIVO_BANCO = (
                r"C:\Python\gerador-conteudo\banco\conteudo.json"
            )

            print()
            print("==============================")
            print("BANCO PADRÃO UTILIZADO")
            print("==============================")

            print(ARQUIVO_BANCO)

            print()

            if os.path.exists(ARQUIVO_BANCO):
                print("BANCO EXISTE")
            else:
                print("BANCO NÃO EXISTE")

            print()

        palavras = txt_palavras.get(
            "1.0",
            tk.END
        ).strip()

        if not palavras:

            messagebox.showerror(
                "Erro",
                "Informe ao menos uma palavra-chave."
            )

            return

        lista_palavras = [
            x.strip()
            for x in palavras.splitlines()
            if x.strip()
        ]

        novo_conjunto = verificar_novo_conjunto(
            lista_palavras
        )

        if novo_conjunto:
            progresso = None

        # ========================================================
        # 01. RETOMAR DE ONDE PAROU
        # ========================================================

        inicio_lista = 0

        if progresso:

            inicio_lista = progresso.get(
                "indice_tema",
                0
            )

            if inicio_lista >= len(lista_palavras):

                print()
                print("==============================")
                print("TODOS OS TEMAS JÁ PROCESSADOS")
                print("==============================")

                atualizar_progresso(
                    100,
                    "Todos os temas já foram processados."
                )

                return

        # ========================================================
        # 02. PROCESSAR TEMAS
        # ========================================================


        for indice_tema, tema in enumerate(
            lista_palavras[inicio_lista:],
            start=inicio_lista
        ):
        
            # ============================================================
            # CORRIGIR AUTOMATICAMENTE O TEMA
            # Mantém o tema correto para todo o processamento.
            # A remoção de acentos continua exclusiva para nomes de arquivo.
            # ============================================================
        
            tema_original = str(
                tema or ""
            ).strip()
        
            if tema_original:
        
                prompt_correcao_tema = f"""
        Corrija somente a ortografia e a acentuação deste tema técnico.
        
        REGRAS:
        - Não altere as palavras.
        - Não acrescente palavras.
        - Não retire palavras.
        - Não explique.
        - Retorne somente o tema corrigido.
        - Preserve maiúsculas/minúsculas quando apropriado.
        
        TEMA:
        {tema_original}
        """
        
                try:
        
                    resposta_tema_http = requests.post(
                        "http://localhost:11434/api/chat",
                        json={
                            "model": MODELO_OLLAMA,
                            "messages": [
                                {
                                    "role": "user",
                                    "content": prompt_correcao_tema
                                }
                            ],
                            "stream": False,
                            "options": {
                                "temperature": 0.0,
                                "num_predict": 50
                            }
                        },
                        timeout=(30, 300)
                    )

                    if resposta_tema_http.status_code != 200:
                        raise RuntimeError(
                            f"Ollama retornou HTTP {resposta_tema_http.status_code}"
                        )

                    resposta_tema = resposta_tema_http.json()
                    tema_corrigido = str(
                        resposta_tema.get("message", {}).get("content", "")
                    ).strip()

                    if tema_corrigido:
                        # Se o tema original já possui caracteres acentuados,
                        # preserva exatamente a forma digitada pelo usuário.
                        if tema_original != normalizar_tema_chave(tema_original):
                            tema = tema_original
                        else:
                            tema = tema_corrigido
                    else:
                        tema = tema_original
        
                except Exception as erro:
        
                    print(
                        "AVISO: não foi possível corrigir "
                        "a acentuação do tema:",
                        erro
                    )
        
                    tema = tema_original
        
            print()
            print("==============================")
            print("TEMA ORIGINAL:", tema_original)
            print("TEMA CORRIGIDO:", tema)
            print("==============================")
        
            # Reinicia o cronômetro sempre que uma nova palavra/tema começa.
            inicio_tema = time.time()
            INICIO_TEMA_TEMPO_ATIVO = inicio_tema
            TEMA_TEMPO_ATIVO = tema

            # Atualiza imediatamente a linha da interface para zerar o
            # cronômetro do tema novo, sem aguardar outra etapa do processo.
            atualizar_progresso(
                0,
                f"Tema: {tema} | Gerando conteúdo com IA | Tempo: 00:00"
            )

            threading.Thread(
                target=atualizar_tempo_execucao,
                args=(inicio_tema, tema),
                daemon=True
            ).start()

            print()
            print("=" * 50)
            print(
                "PROCESSANDO:",
                tema
            )
            print("=" * 50)

            # ====================================================
            # 03. INICIALIZAR DADOS DO TEMA
            # ====================================================

            textos = []
            paginas = []
            urls = []
            dados_coleta = []
            paginas_aprovadas = 0
            
            # ====================================================
            # TIPO OFICIAL DA PÁGINA
            # ====================================================
            #
            # produto ou servico
            #
            # NÃO confundir com o tipo da fonte:
            # html, pdf, etc.
            # ====================================================

            tipo = identificar_tipo_tema(
                tema
            )

            print()
            print("==============================")
            print("TIPO OFICIAL DO TEMA")
            print("==============================")

            print(
                "TEMA:",
                tema
            )

            print(
                "TIPO:",
                tipo
            )

            # ====================================================
            # 04. IDENTIFICAR GRUPO DO TEMA
            # ====================================================

            print()
            print("==============================")
            print("IDENTIFICANDO GRUPO DO TEMA")
            print("==============================")

            print(
                "TEMA:",
                tema
            )

            grupo = identificar_grupo_tema(
                tema
            )

            print()
            print("==============================")
            print("GRUPO IDENTIFICADO")
            print("==============================")

            print(grupo)

            # ====================================================
            # 05. GERAR ENTENDIMENTO DO PRODUTO
            # ====================================================

            print()
            print("==============================")
            print("GERANDO ENTENDIMENTO DO PRODUTO")
            print("==============================")

            print(
                "TEMA:",
                tema
            )

            print(
                "GRUPO:",
                grupo
            )

            entendimento = gerar_entendimento_produto(
                tema,
                grupo
            )

            # ====================================================
            # 06. VALIDAR ENTENDIMENTO
            # ====================================================

            print()
            print("==============================")
            print("ENTENDIMENTO DO PRODUTO")
            print("==============================")

            if entendimento:

                print(entendimento)

                print()
                print(
                    "ENTENDIMENTO GERADO COM SUCESSO"
                )

                print(
                    "CARACTERES:",
                    len(
                        str(
                            entendimento
                        )
                    )
                )

            else:

                print(
                    "ENTENDIMENTO NÃO GERADO"
                )

                print(
                    "O FLUXO CONTINUARÁ SEM ENTENDIMENTO."
                )

            # ====================================================
            # 07. INICIO COLETA
            # ====================================================

            print()
            print("==============================")
            print("INICIO COLETA")
            print("==============================")

            print(
                "TEMA:",
                tema
            )

            # ====================================================
            # TRY INTERNO DO TEMA
            # ====================================================

            try:

                # =================================================
                # 08. OBTER DADOS EXISTENTES DO TEMA
                # =================================================

                dados_existentes = obter_dados_banco(
                    tema
                )

                # =================================================
                # 09. INICIALIZAR MAPA MEAD
                # =================================================

                mapa_mead = None

                print()
                print("==============================")
                print("VERIFICANDO MAPA MEAD")
                print("==============================")

                # =================================================
                # 10. VERIFICAR SE EXISTE MAPA NO BANCO
                # =================================================

                if dados_existentes.get("mapa_mead"):

                    print()
                    print("==============================")
                    print("MAPA MEAD ENCONTRADO NO BANCO")
                    print("==============================")

                    mapa_mead = dados_existentes.get(
                        "mapa_mead"
                    )

                    # =============================================
                    # 11. NORMALIZAR MAPA RECEBIDO
                    # =============================================

                    if isinstance(
                        mapa_mead,
                        dict
                    ):

                        mapa_mead = mapa_mead.get(
                            "texto",
                            ""
                        )

                    elif mapa_mead is None:

                        mapa_mead = ""

                    else:

                        mapa_mead = str(
                            mapa_mead
                        )

                    mapa_mead = mapa_mead.strip()

                    print()
                    print("MAPA MEAD CARREGADO:")

                    print(
                        mapa_mead[:500]
                    )

                else:

                    # =============================================
                    # 12. MAPA NÃO EXISTE
                    # =============================================

                    print()
                    print("==============================")
                    print("MAPA MEAD NÃO ENCONTRADO")
                    print("==============================")

                    mapa_mead = None

                # =================================================
                # 13. NORMALIZAR SEM ACENTO
                # =================================================

                def normalizar_sem_acento(texto):

                    if not texto:
                        return ""

                    texto = unicodedata.normalize(
                        "NFD",
                        str(texto)
                    )

                    texto = "".join(
                        caractere
                        for caractere in texto
                        if unicodedata.category(
                            caractere
                        ) != "Mn"
                    )

                    return texto.upper()

                # =================================================
                # 14. VALIDAR MAPA EXISTENTE
                # =================================================

                mapa_validado = False

                if mapa_mead:

                    mapa_maiusculo = normalizar_sem_acento(
                        mapa_mead
                    )

                    mapa_validado = True

                # =================================================
                # 15. ESTADO FINAL DO MAPA
                # =================================================

                print()
                print("==============================")
                print("ESTADO FINAL DO MAPA MEAD")
                print("==============================")

                print(
                    "MAPA EXISTE:",
                    bool(mapa_mead)
                )

                print(
                    "TAMANHO:",
                    len(mapa_mead)
                    if mapa_mead
                    else 0
                )

                if mapa_mead:

                    print()
                    print("MAPA MEAD FINAL:")

                    print(
                        mapa_mead[:500]
                    )

                else:

                    print()
                    print(
                        "NENHUM MAPA MEAD DISPONÍVEL"
                    )

                # =================================================
                # 16. GARANTIR TEXTOS
                # =================================================

                if not textos:

                    print()
                    print("==============================")
                    print("CARREGANDO DADOS BRUTOS PARA MAPA MEAD")
                    print("==============================")

                    textos = carregar_bruto(
                        tema
                    )

                    print(
                        "TEXTOS CARREGADOS:",
                        len(textos)
                    )

                # =================================================
                # 17. ADICIONAR SITES INFORMADOS PELO USUÁRIO
                # =================================================

                print()
                print("==============================")
                print("TEXTOS ANTES DA ANÁLISE")
                print("==============================")

                print(
                    "TOTAL:",
                    len(textos)
                )

                sites_consulta = coletar_sites_consulta()

                analise_site = analisar_site_referencia(
                    tema,
                    sites_consulta
                )

                if sites_consulta:

                    textos.extend(
                        sites_consulta
                    )

                    print()
                    print("==============================")
                    print("SITES ADICIONADOS AO PATRIMÔNIO")
                    print("==============================")

                    print(
                        "SITES:",
                        len(sites_consulta)
                    )

                    print(
                        "TOTAL DE TEXTOS:",
                        len(textos)
                    )

                # =================================================
                # 18. NORMALIZAR TEXTOS PARA MEAD
                # =================================================

                textos_normalizados = []

                for item in textos:

                    if isinstance(
                        item,
                        dict
                    ):

                        texto = item.get(
                            "texto",
                            ""
                        )

                        if texto and texto.strip():

                            textos_normalizados.append({
                                "url": item.get(
                                    "url",
                                    ""
                                ),
                                "tipo": item.get(
                                    "tipo",
                                    "texto"
                                ),
                                "texto": texto
                            })

                    elif isinstance(
                        item,
                        str
                    ):

                        if item.strip():

                            textos_normalizados.append({
                                "url": "",
                                "tipo": "texto",
                                "texto": item
                            })

                print()
                print("==============================")
                print("TEXTOS NORMALIZADOS PARA MEAD")
                print("==============================")

                print(
                    "QUANTIDADE:",
                    len(textos_normalizados)
                )

                # =================================================
                # 19. LIMPAR REFERÊNCIAS
                # =================================================

                textos_para_mapa = limpar_lista_referencias(
                    textos_normalizados,
                    tema
                )

                print()
                print("==============================")
                print("TEXTOS ENVIADOS AO MAPA MEAD")
                print("==============================")

                print(
                    "QUANTIDADE:",
                    len(textos_para_mapa)
                )

                # =================================================
                # 20. GARANTIR FORMATO FINAL
                # =================================================

                textos_para_mapa_corrigidos = []

                for item in textos_para_mapa:

                    if isinstance(
                        item,
                        dict
                    ):

                        texto = item.get(
                            "texto",
                            ""
                        )

                        if texto and texto.strip():

                            textos_para_mapa_corrigidos.append({
                                "url": item.get(
                                    "url",
                                    ""
                                ),
                                "tipo": item.get(
                                    "tipo",
                                    "texto"
                                ),
                                "texto": texto
                            })

                    elif isinstance(
                        item,
                        str
                    ):

                        if item.strip():

                            textos_para_mapa_corrigidos.append({
                                "url": "",
                                "tipo": "texto",
                                "texto": item
                            })

                # =================================================
                # 21. LIMPAR MAPA MEAD
                # =================================================

                def limpar_mapa_mead(texto):

                    if not texto:
                        return ""

                    substituicoes = {

                        "ATIVOS:":
                        "ATIVOS_NARRATIVOS:",

                        "EXPERIÊNCIA DO FABRICANTE":
                        "experiência técnica",

                        "experiência do fabricante":
                        "experiência técnica",

                        "garantia de alta resistência":
                        "desempenho técnico",

                        "cumprimento das normas":
                        "conformidade técnica",

                        "fabricante":
                        "especialista do segmento",

                        "Fabricante":
                        "especialista do segmento",
                    }

                    for antigo, novo in substituicoes.items():

                        texto = texto.replace(
                            antigo,
                            novo
                        )

                    remover = [

                        "IT 09",
                        "IT-09",
                        "CBPMESP",
                        "NBR",
                        "ISO",
                        "ABNT",
                        "certificação",
                        "certificado",
                        "marca",
                        "modelo",
                        "linha",
                        "série",
                    ]

                    for palavra in remover:

                        texto = texto.replace(
                            palavra,
                            ""
                        )

                    return texto.strip()

                if mapa_mead:

                    mapa_mead = limpar_mapa_mead(
                        mapa_mead
                    )

                # =================================================
                # 22. VERIFICAR CONTEÚDO RELACIONADO
                # =================================================

                textos_encontrados = False

                # =================================================
                # 23. BUSCAR TEMAS RELACIONADOS
                # =================================================

                if not dados_existentes.get("conteudo"):

                    relacionados = buscar_temas_relacionados(
                        tema
                    )

                    relacionados_filtrados = []

                    if relacionados:

                        for item in relacionados:

                            if pertence_ao_grupo_principal(
                                tema,
                                item
                            ):

                                relacionados_filtrados.append(
                                    item
                                )

                    relacionados = relacionados_filtrados

                    if relacionados:

                        relacionados = relacionados[:5]

                        print()
                        print("==============================")
                        print("TEMAS RELACIONADOS COMPATÍVEIS")
                        print("==============================")

                        for item in relacionados:

                            print(item)

                            dados_relacionado = obter_dados_banco(
                                item
                            )

                            conteudo_relacionado = False

                            # NÃO USAR MAPAS MEAD DE OUTROS TEMAS

                            if dados_relacionado.get(
                                "conteudo"
                            ):

                                textos.append(
                                    dados_relacionado["conteudo"]
                                )

                                conteudo_relacionado = True

                            if dados_relacionado.get(
                                "bruto"
                            ):

                                textos.extend(
                                    dados_relacionado["bruto"]
                                )

                                conteudo_relacionado = True

                            if conteudo_relacionado:

                                textos_encontrados = True

                            print()
                            print("==============================")
                            print("BASE RELACIONADA ANALISADA")
                            print("==============================")

                            print(
                                "TEXTOS ACUMULADOS:",
                                len(textos)
                            )

                            print(
                                "CONTEÚDO RELACIONADO VÁLIDO:",
                                conteudo_relacionado
                            )

                    else:

                        print()
                        print("==============================")
                        print("NENHUM TEMA RELACIONADO COMPATÍVEL")
                        print("==============================")


                # =================================================
                # 24. COMPLETAR COLETA QUANDO NECESSÁRIO
                # =================================================

                avaliacao_patrimonio = avaliar_patrimonio_existente(
                    tema
                )

                textos_brutos = avaliacao_patrimonio.get(
                    "fontes",
                    []
                )

                if avaliacao_patrimonio.get(
                    "suficiente",
                    False
                ):

                    print(
                        "USANDO PATRIMÔNIO EXISTENTE"
                    )

                    # Não apagar os sites já coletados.
                    textos = (
                        list(textos_brutos)
                        + list(sites_consulta)
                    )

                else:

                    print(
                        "DADOS BRUTOS NÃO ENCONTRADOS - PESQUISANDO"
                    )

                    urls = pesquisar_completo(
                        tema
                    )

                    for url in urls:

                        print()
                        print("COLETANDO:")
                        print(url)

                        dados = coletar_pagina(
                            url
                        )

                        texto = ""

                        if dados:

                            texto = dados.get(
                                "texto",
                                ""
                            )

                        # =====================================
                        # 25. GARANTIR TIPO DA FONTE
                        # =====================================
                        #
                        # tipo = tipo oficial da página
                        #        produto / servico
                        #
                        # tipo_fonte = tipo da fonte coletada
                        #              html / pdf / etc.
                        # =====================================

                        tipo_fonte = "html"

                        if isinstance(
                            dados,
                            dict
                        ):

                            tipo_fonte = dados.get(
                                "tipo",
                                "html"
                            )

                        # =====================================
                        # 26. VALIDAR CONTEÚDO
                        # =====================================

                        if len(texto) > 500:

                            relevante = (
                                validar_conteudo_relevante(
                                    texto,
                                    tema
                                )
                            )

                            tecnico = (
                                validar_conteudo_tecnico(
                                    texto,
                                    tema
                                )
                            )

                            if relevante and tecnico:

                                textos.append(
                                    {
                                        "url": url,
                                        "tipo": tipo_fonte,
                                        "texto": texto
                                    }
                                )

                                paginas_aprovadas += 1

                                dados_coleta.append({
                                    "url": url,
                                    "tipo": tipo_fonte,
                                    "status": "aprovado",
                                    "motivo": "",
                                    "texto": texto
                                })

                            else:

                                dados_coleta.append({
                                    "url": url,
                                    "tipo": tipo_fonte,
                                    "status": "descartado",
                                    "motivo": "conteúdo não aprovado",
                                    "texto": texto
                                })

                                print(
                                    "CONTEUDO MANTIDO NO JSON COMO DESCARTADO"
                                )

                    # =========================================
                    # 27. SALVAR DADOS BRUTOS
                    # =========================================

                    print()
                    print("==============================")
                    print("SALVANDO DADOS BRUTOS")
                    print("==============================")

                    print(
                        "TEMA:",
                        tema
                    )

                    print(
                        "PÁGINAS PARA SALVAR:",
                        len(dados_coleta)
                    )

                    if dados_coleta:

                        salvar_bruto(
                            tema,
                            dados_coleta
                        )

                    else:

                        print()
                        print("==============================")
                        print("DADOS COLETA VAZIOS")
                        print("BRUTO EXISTENTE SERÁ PRESERVADO")
                        print("==============================")

                    # =========================================
                    # 28. RECARREGAR BRUTO APÓS SALVAR
                    # =========================================

                    dados_coleta = carregar_bruto(
                        tema
                    )

                    textos = []

                    for item in dados_coleta:

                        if not isinstance(
                            item,
                            dict
                        ):
                            continue

                        texto = item.get(
                            "texto",
                            ""
                        )

                        if texto and texto.strip():

                            textos.append({
                                "url": item.get(
                                    "url",
                                    ""
                                ),
                                "tipo": item.get(
                                    "tipo",
                                    ""
                                ),
                                "texto": texto,
                                "identidade_fonte": item.get(
                                    "identidade_fonte",
                                    {}
                                )
                            })

                    # Preservar os sites informados pelo usuário.

                    if sites_consulta:

                        textos.extend(
                            sites_consulta
                        )

                    print(
                        "TEXTOS PARA MAPA MEAD:",
                        len(textos)
                    )

                # =================================================
                # 29. NORMALIZAR TEXTOS PARA MAPA MEAD
                # =================================================



                textos_para_mapa = limpar_lista_referencias(
                    textos,
                    tema
                )

                print(
                    "TEXTOS APÓS LIMPEZA:",
                    len(textos_para_mapa)
                )

                # =================================================
                # 30. GARANTIR FORMATO FINAL
                # =================================================
                
                textos_para_mapa_corrigidos = []
                
                for item in textos_para_mapa:
                
                    if isinstance(
                        item,
                        str
                    ):
                
                        textos_para_mapa_corrigidos.append({
                            "url": "",
                            "tipo": "texto",
                            "texto": item
                        })
                
                    elif isinstance(
                        item,
                        dict
                    ):
                
                        texto = item.get(
                            "texto",
                            ""
                        )
                
                        if texto and texto.strip():
                
                            textos_para_mapa_corrigidos.append({
                                "url": item.get(
                                    "url",
                                    ""
                                ),
                                "tipo": item.get(
                                    "tipo",
                                    "texto"
                                ),
                                "texto": texto,
                                "identidade_fonte": item.get(
                                    "identidade_fonte",
                                    {}
                                )
                            })
                
                print(
                    "TEXTOS FINAIS PARA MAPA MEAD:",
                    len(textos_para_mapa_corrigidos)
                )

                # =================================================
                # 31. GERAR MAPA MEAD SE NECESSÁRIO
                # =================================================

                if not mapa_mead:

                    mapa_mead = gerar_e_salvar_mapa_mead(
                        tema,
                        textos_para_mapa_corrigidos
                    )

                if mapa_mead:

                    mapa_mead = limpar_mapa_mead(
                        mapa_mead
                    )

                if not mapa_mead:

                    print()
                    print("==============================")
                    print("MAPA MEAD NÃO GERADO")
                    print("==============================")

                    continue

                print()
                print("==============================")
                print("COLETA E MAPA MEAD FINALIZADOS")
                print("==============================")

                salvar_progresso(
                    tema,
                    indice_tema + 1,
                    "mapa_mead",
                    0,
                    "concluido",
                    lista_palavras
                )

                # =================================================
                # 32. VERIFICAR DADOS PARA IA
                # =================================================

                if not textos:

                    print()
                    print("==============================")
                    print("SEM DADOS PARA GERAR CONTEÚDO")
                    print("==============================")

                    salvar_progresso(
                        tema,
                        indice_tema,
                        "erro_coleta",
                        0,
                        "sem dados brutos",
                        lista_palavras
                    )

                    continue

                # =================================================
                # 33. BLOQUEIO MAPA MEAD OBRIGATÓRIO
                # =================================================

                if (
                    not mapa_mead
                    or not str(mapa_mead).strip()
                    or str(mapa_mead).strip() in [
                        "{}",
                        "[]"
                    ]
                ):

                    print()
                    print("==============================")
                    print("MAPA MEAD AUSENTE OU INVÁLIDO")
                    print("==============================")

                    print(
                        "NÃO É POSSÍVEL GERAR CONTEÚDO"
                    )

                    salvar_progresso(
                        tema,
                        indice_tema,
                        "erro_mapa_mead",
                        0,
                        "mapa mead ausente",
                        lista_palavras
                    )

                    continue

                # =================================================
                # 34. INICIAR GERAÇÃO DE CONTEÚDO
                # =================================================

                salvar_progresso(
                    tema,
                    indice_tema,
                    "conteudo_completo",
                    1,
                    "processando",
                    lista_palavras
                )

                print()
                print("==============================")
                print("CRIANDO CONTEÚDO COMPLETO")
                print("==============================")

                inicio_conteudo = time.time()

                # =================================================
                # 35. INICIAR MONITOR DE IA
                # =================================================

                IA_PROCESSANDO = True

                atualizar_progresso(
                    50,
                    f"Tema: {tema} | Gerando conteúdo com IA..."
                )

                total_textos = len(textos)

                total_caracteres = sum(
                    len(
                        item.get(
                            "texto",
                            ""
                        )
                        if isinstance(
                            item,
                            dict
                        )
                        else str(item)
                    )
                    for item in textos
                )

                media_caracteres = (
                    total_caracteres / total_textos
                    if total_textos
                    else 0
                )

                print()
                print("==============================")
                print("DADOS ENVIADOS PARA IA")
                print("==============================")

                print(
                    "TEMA:",
                    tema
                )

                print(
                    "TEXTOS:",
                    total_textos
                )

                print(
                    "CARACTERES:",
                    total_caracteres
                )

                print(
                    "MÉDIA POR TEXTO:",
                    int(media_caracteres)
                )

                if textos:

                    primeiro_texto = (

                        textos[0].get(
                            "texto",
                            ""
                        )

                        if isinstance(
                            textos[0],
                            dict
                        )

                        else str(
                            textos[0]
                        )
                    )

                    print(
                        "PRIMEIRO TEXTO:",
                        len(primeiro_texto),
                        "caracteres"
                    )

                print("MAPA MEAD ENVIADO AO CONTEÚDO:",
                      "válido" if mapa_mead else "VAZIO",
                      "|", len(str(mapa_mead)), "caracteres")

                print()
                print("==============================")
                print("CHAMANDO GERADOR DE CONTEÚDO")
                print("==============================")

                inicio_ia = time.time()

                print()
                print("==============================")
                print("INICIANDO IA - CONTEÚDO")
                print("==============================")

                print(
                    "TEMA:",
                    tema
                )

                print(
                    "TEXTOS:",
                    len(textos)
                )

                print(
                    "MAPA:",
                    len(str(mapa_mead))
                )

                print(
                    "CARACTERES DOS TEXTOS:",
                    sum(
                        len(
                            item.get(
                                "texto",
                                ""
                            )
                            if isinstance(
                                item,
                                dict
                            )
                            else str(item)
                        )
                        for item in textos
                    )
                )

                print(
                    "INÍCIO IA:",
                    time.strftime("%H:%M:%S")
                )

                print()
                print(
                    "AGUARDANDO RETORNO DO OLLAMA..."
                )
                print("==============================")

                estrutura_editorial_atual = (
                    obter_estrutura_editorial()
                )

                print("ESTRUTURA EDITORIAL:",
                      sum(1 for valor in estrutura_editorial_atual.values() if valor),
                      "/", len(estrutura_editorial_atual), "itens ativos")

                # ========================================================
                # NOME DO SITE CAPTURADO DA INTERFACE
                # ========================================================
                
                nome_site = (
                    entrada_site.get().strip()
                    if "entrada_site" in globals()
                    else ""
                )
                
                conteudo_completo = (
                    gerar_conteudo_completo(
                        tema,
                        textos,
                        mapa_mead,
                        estrutura_editorial_atual,
                        dados_coleta=dados_coleta,
                        nome_site=nome_site
                    )
                )

                IA_PROCESSANDO = False

                fim_ia = time.time()

                print()
                print("==============================")
                print("OLLAMA RETORNOU")
                print("==============================")

                print(
                    "FIM IA:",
                    time.strftime("%H:%M:%S")
                )

                print(
                    "TEMPO IA:",
                    formatar_tempo(
                        fim_ia - inicio_ia
                    )
                )

                print(
                    "TIPO:",
                    type(conteudo_completo)
                )

                print(
                    "CARACTERES:",
                    len(
                        conteudo_completo or ""
                    )
                )

                tempo_total = (
                    time.time()
                    -
                    inicio_conteudo
                )

                tempo_ia = (
                    fim_ia
                    -
                    inicio_ia
                )

                tamanho_conteudo = len(
                    conteudo_completo or ""
                )

                print()
                print("==============================")
                print("RETORNOU DO GERADOR")
                print("==============================")

                print(
                    "TIPO:",
                    type(conteudo_completo)
                )

                print(
                    "TAMANHO:",
                    tamanho_conteudo
                )

                print()
                print("==============================")
                print("ESTATÍSTICAS DA IA")
                print("==============================")

                print(
                    "TEXTOS ENVIADOS:",
                    total_textos
                )

                print(
                    "CARACTERES ENVIADOS:",
                    total_caracteres
                )

                print(
                    "MÉDIA POR TEXTO:",
                    int(media_caracteres)
                )

                print(
                    "CARACTERES GERADOS:",
                    tamanho_conteudo
                )

                print(
                    "TEMPO IA:",
                    formatar_tempo(
                        tempo_ia
                    )
                )

                print(
                    "TEMPO TOTAL:",
                    formatar_tempo(
                        tempo_total
                    )
                )

                if (
                    tempo_ia > 0
                    and tamanho_conteudo > 0
                ):

                    print(
                        "VELOCIDADE:",
                        int(
                            tamanho_conteudo
                            /
                            tempo_ia
                        ),
                        "caracteres/seg"
                    )

                atualizar_progresso(
                    90,
                    (
                        f"Tema: {tema} | "
                        f"Conteúdo gerado | "
                        f"Tempo: {formatar_tempo(tempo_total)}"
                    )
                )


                # =================================================
                # 36. SALVAR CONTEÚDO
                # =================================================

                if (
                    conteudo_completo
                    and conteudo_completo.strip()
                ):

                    salvar_progresso(
                        tema,
                        indice_tema + 1,
                        "finalizado",
                        0,
                        "concluido",
                        lista_palavras
                    )

                else:

                    IA_PROCESSANDO = False

                    print()
                    print("==============================")
                    print("FALHA AO GERAR CONTEÚDO COMPLETO")
                    print("==============================")

                    print(
                        "TEMA:",
                        tema
                    )

                    print(
                        "RETORNO:",
                        repr(conteudo_completo)
                    )

                    salvar_progresso(
                        tema,
                        indice_tema,
                        "conteudo_completo",
                        1,
                        "pendente_recuperacao",
                        lista_palavras
                    )

            except Exception as e:

                IA_PROCESSANDO = False

                print()
                print("==============================")
                print("ERRO NO TEMA")
                print("==============================")

                print(
                    repr(e)
                )

                atualizar_progresso(
                    0,
                    f"Erro no tema {tema}: {e}"
                )


        # ========================================================
        # 37. FINAL DO PROCESSAMENTO GERAL
        # ========================================================

        PROCESSAMENTO_ATIVO = False
        IA_PROCESSANDO = False

        tempo_final = (
            time.time()
            -
            inicio_total
        )

        print()
        print("==============================")
        print("PROCESSAMENTO FINALIZADO")
        print("==============================")

        print(
            "TEMPO TOTAL:",
            formatar_tempo(
                tempo_final
            )
        )

        tempo_ultimo_tema = (
            time.time() - INICIO_TEMA_TEMPO_ATIVO
            if INICIO_TEMA_TEMPO_ATIVO is not None
            else tempo_final
        )
        tema_finalizado = TEMA_TEMPO_ATIVO or "Processamento"

        atualizar_progresso(
            100,
            (
                f"Tema: {tema_finalizado} | Concluído | "
                f"Tempo: {formatar_tempo(tempo_ultimo_tema)} | "
                f"Tempo geral: {formatar_tempo(tempo_final)}"
            )
        )

    except Exception as e:

        PROCESSAMENTO_ATIVO = False
        IA_PROCESSANDO = False

        print()
        print("==============================")
        print("ERRO GERAL NO PROCESSAMENTO")
        print("==============================")

        print(
            repr(e)
        )

        atualizar_progresso(
            0,
            f"Erro geral: {e}"
        )
        
        

# ============================================================
# BOTÃO - APAGAR PROGRESSO MANUAL
# ============================================================

def apagar_progresso_manual():

    try:

        if os.path.exists(ARQUIVO_PROGRESSO):

            os.remove(
                ARQUIVO_PROGRESSO
            )

            messagebox.showinfo(
                "Progresso",
                "Progresso apagado com sucesso."
            )

        else:

            messagebox.showinfo(
                "Progresso",
                "Não existe progresso salvo."
            )


    except Exception as e:

        messagebox.showerror(
            "Erro",
            str(e)
        )
        

# ============================================================
# GERAR DOCX — LEITURA DIRETA DO JSON / PADRÃO WORD 6.60
# ============================================================

def _docx_configurar_fonte(run, tamanho=11, negrito=False):
    run.font.name = "Calibri"
    run.font.size = Pt(tamanho)
    run.bold = negrito
    # Compatibilidade Word: fonte também para leste europeu/complex scripts.
    try:
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    except Exception:
        pass

def _docx_adicionar_paragrafo(documento, texto, tamanho=11, negrito=False,
                              alinhamento=None, espaco_depois=8, cor=None):
    texto = str(texto or "").strip()
    if not texto:
        return None
    p = documento.add_paragraph()
    p.alignment = alinhamento if alinhamento is not None else WD_ALIGN_PARAGRAPH.JUSTIFY
    # Espaçamento mais confortável, mantendo o padrão visual do conversor.
    p.paragraph_format.line_spacing = 1.08
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(espaco_depois)
    run = p.add_run(texto)
    _docx_configurar_fonte(run, tamanho, negrito)
    if cor is not None:
        run.font.color.rgb = cor
    return p

def _docx_adicionar_separador(documento):
    return _docx_adicionar_paragrafo(
        documento,
        "-" * 126,
        tamanho=11,
        negrito=False,
        cor=RGBColor(0, 0, 0),
        espaco_depois=8
    )

def _docx_obter_blocos(dados_pagina):
    blocos = dados_pagina.get("blocos", {})
    if not isinstance(blocos, dict):
        blocos = {}
    # Algumas versões do banco guardam bloco_1 diretamente na página.
    if not blocos:
        blocos = {
            f"bloco_{n}": dados_pagina.get(f"bloco_{n}", {})
            for n in range(1, 6)
            if isinstance(dados_pagina.get(f"bloco_{n}"), dict)
        }
    return blocos

def _docx_obter_imagens(dados_pagina):
    imagens = dados_pagina.get("imagens", {})
    if not isinstance(imagens, dict):
        imagens = {}
    return imagens

def _docx_imagem_valor(imagens, numero):
    item = imagens.get(f"imagem_{numero}", {})
    if isinstance(item, str):
        return {"url": item, "arquivo": ""}
    if not isinstance(item, dict):
        return {"url": "", "arquivo": ""}
    return {
        "url": str(item.get("url", "") or "").strip(),
        "arquivo": str(item.get("arquivo", "") or "").strip(),
        "alt": str(item.get("alt", "") or "").strip(),
        "descricao": str(item.get("descricao", "") or "").strip(),
    }

def _docx_adicionar_imagem(documento, imagem, largura=None):
    """Insere somente a imagem local já existente no JSON/interface.
    Sem arquivo local, mantém a URL como texto para não fazer download.
    Retorna True quando a imagem foi inserida fisicamente."""
    arquivo = imagem.get("arquivo", "")
    url = imagem.get("url", "")

    if arquivo and os.path.isfile(arquivo):
        try:
            run = documento.add_paragraph().add_run()
            if largura is not None:
                run.add_picture(arquivo, width=largura)
            else:
                run.add_picture(arquivo)
            return True
        except Exception as erro:
            print("AVISO: não foi possível inserir imagem local:", arquivo, repr(erro))

    if url:
        _docx_adicionar_paragrafo(documento, url, tamanho=11)
    return False

def _docx_adicionar_par_imagens(documento, imagens, numero_1, numero_2):
    """Representa o par de imagens como duas colunas de 50%, sem bordas."""
    tabela = documento.add_table(rows=1, cols=2)
    tabela.autofit = True

    try:
        tabela.columns[0].width = Inches(3.15)
        tabela.columns[1].width = Inches(3.15)
    except Exception:
        pass

    for coluna, numero in enumerate((numero_1, numero_2)):
        celula = tabela.cell(0, coluna)
        paragrafo = celula.paragraphs[0]
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragrafo.paragraph_format.line_spacing = 1.08
        paragrafo.paragraph_format.space_before = Pt(0)
        paragrafo.paragraph_format.space_after = Pt(0)
        imagem = _docx_imagem_valor(imagens, numero)
        arquivo = imagem.get("arquivo", "")
        url = imagem.get("url", "")

        if arquivo and os.path.isfile(arquivo):
            try:
                run = paragrafo.add_run()
                run.add_picture(arquivo, width=Inches(3.0))
                continue
            except Exception as erro:
                print("AVISO: não foi possível inserir imagem local:", arquivo, repr(erro))

        if url:
            run = paragrafo.add_run(url)
            run.font.name = "Calibri"
            run.font.size = Pt(9)
        else:
            run = paragrafo.add_run(f"[Imagem {numero}]")
            run.font.name = "Calibri"
            run.font.size = Pt(11)

    # Remove bordas da tabela para reproduzir visualmente os dois span6.
    try:
        tbl_pr = tabela._tbl.tblPr
        borders = tbl_pr.first_child_found_in("w:tblBorders")
        if borders is None:
            borders = OxmlElement("w:tblBorders")
            tbl_pr.append(borders)
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            elemento = borders.find(qn(f"w:{edge}"))
            if elemento is None:
                elemento = OxmlElement(f"w:{edge}")
                borders.append(elemento)
            elemento.set(qn("w:val"), "nil")
    except Exception:
        pass

def _docx_inserir_segmentos_se_existirem(documento, dados_pagina, bloco_numero):
    segmentos = dados_pagina.get("segmentos_listas", {})
    if not isinstance(segmentos, dict):
        return False

    posicionamento = dados_pagina.get("posicionamento_listas", {})
    if not isinstance(posicionamento, dict):
        posicionamento = {}

    try:
        bloco_pos = int(posicionamento.get("bloco"))
    except Exception:
        bloco_pos = 3

    if bloco_pos not in (3, 4, 5) or bloco_pos != bloco_numero:
        return False

    listas = []
    for chave, lista in segmentos.items():
        if not isinstance(lista, list):
            continue
        for item in lista:
            item = str(item or "").strip()
            if item:
                listas.append(item)

    if not listas:
        return False

    subtitulo = str(
        dados_pagina.get("subtitulo_segmentos", "") or ""
    ).strip()
    if subtitulo:
        subtitulo = subtitulo.rstrip().rstrip(":") + ":"
        _docx_adicionar_paragrafo(documento, subtitulo, tamanho=11, negrito=False)

    for item in listas:
        p = documento.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.08
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(item)
        _docx_configurar_fonte(run, 11, False)

    return True

def _docx_adicionar_bloco(documento, bloco, numero):
    if not isinstance(bloco, dict):
        return
    titulo = str(bloco.get("titulo", "") or "").strip()
    if titulo:
        # O título do segundo bloco fica alinhado à esquerda.
        # Os parágrafos e os demais títulos mantêm o alinhamento padrão.
        alinhamento_titulo = (
            WD_ALIGN_PARAGRAPH.LEFT if numero in (1, 2) else None
        )
        _docx_adicionar_paragrafo(
            documento, titulo, tamanho=14, negrito=True,
            alinhamento=alinhamento_titulo
        )

    paragrafos = bloco.get("paragrafos_ollama", [])
    if not isinstance(paragrafos, list):
        paragrafos = []

    # O DOCX usa exclusivamente o texto editorial já salvo no JSON.
    for texto in paragrafos[:3]:
        _docx_adicionar_paragrafo(documento, texto, tamanho=11)
def _docx_definir_idioma_ptbr(documento):
    """Define Português (Brasil) como idioma de revisão do documento."""
    idioma = "pt-BR"

    def aplicar(rpr):
        if rpr is None:
            return
        lang = rpr.find(qn("w:lang"))
        if lang is None:
            lang = OxmlElement("w:lang")
            rpr.append(lang)
        lang.set(qn("w:val"), idioma)
        lang.set(qn("w:eastAsia"), idioma)
        lang.set(qn("w:bidi"), idioma)

    settings = documento.settings.element
    theme_lang = settings.find(qn("w:themeFontLang"))
    if theme_lang is None:
        theme_lang = OxmlElement("w:themeFontLang")
        settings.append(theme_lang)
    theme_lang.set(qn("w:val"), idioma)
    theme_lang.set(qn("w:eastAsia"), idioma)
    theme_lang.set(qn("w:bidi"), idioma)

    for estilo in documento.styles:
        try:
            aplicar(estilo.element.get_or_add_rPr())
        except Exception:
            pass
    for paragrafo in documento.paragraphs:
        for run in paragrafo.runs:
            aplicar(run._element.get_or_add_rPr())
    for tabela in documento.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    for run in paragrafo.runs:
                        aplicar(run._element.get_or_add_rPr())


def gerar_docx_mead(dados_pagina, pasta_destino, nome_arquivo_base=""):
    """Gera DOCX a partir do JSON já gravado; não chama Ollama nem pesquisa."""
    if not isinstance(dados_pagina, dict):
        raise ValueError("Dados da página inválidos para geração do DOCX.")
    if not pasta_destino:
        pasta_destino = os.path.join(os.path.dirname(ARQUIVO_BANCO), "docx")
    os.makedirs(pasta_destino, exist_ok=True)

    documento = Document()
    for sec in documento.sections:
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3.0)
        sec.right_margin = Cm(3.0)

    normal = documento.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.paragraph_format.line_spacing = 1.08
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    try:
        normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    except Exception:
        pass

    # A estrutura do banco mantém os metadados da página separados dos blocos.
    tema = str(dados_pagina.get("tema", "") or "").strip()
    titulo = str(dados_pagina.get("h1", "") or dados_pagina.get("titulo", "") or tema).strip()
    subtitulo = str(dados_pagina.get("subtitulo", "") or "").strip()
    tags = dados_pagina.get("tags", [])
    if isinstance(tags, list):
        tags_txt = ", ".join(str(x).strip() for x in tags if str(x).strip())
    else:
        tags_txt = str(tags or "").strip()
    if tags_txt:
        _docx_adicionar_paragrafo(documento, tags_txt, tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=8)

    # Título e subtítulo de abertura.
    _docx_adicionar_paragrafo(documento, titulo, tamanho=14, negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    if subtitulo:
        _docx_adicionar_paragrafo(documento, subtitulo, tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    _docx_adicionar_separador(documento)

    imagens = _docx_obter_imagens(dados_pagina)
    imagens_informadas = any(_docx_imagem_valor(imagens, n).get("arquivo") or _docx_imagem_valor(imagens, n).get("url") for n in range(1, 7))
    if imagens_informadas:
        _docx_adicionar_par_imagens(documento, imagens, 1, 2)
    else:
        _docx_adicionar_paragrafo(documento, "[Imagem 1]", tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, cor=RGBColor(68, 114, 196))
        _docx_adicionar_paragrafo(documento, "[Imagem 2]", tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, cor=RGBColor(68, 114, 196))
    _docx_adicionar_separador(documento)


    blocos = _docx_obter_blocos(dados_pagina)
    posicionamento_imagens = dados_pagina.get("posicionamento_imagens", {})
    if not isinstance(posicionamento_imagens, dict):
        posicionamento_imagens = {}
    try:
        pos_imagens_34 = int(posicionamento_imagens.get("imagem_3", 2) or 2)
    except Exception:
        pos_imagens_34 = 2

    for numero in range(1, 6):
        # Imagens 3 e 4 ficam na posição gravada; padrão do modelo: antes do bloco 2.
        if numero == pos_imagens_34:
        
            _docx_adicionar_separador(documento)
            _docx_adicionar_paragrafo(documento, "Otimização", tamanho=14, negrito=True, cor=RGBColor(255, 0, 0))
            if imagens_informadas:
                _docx_adicionar_par_imagens(documento, imagens, 3, 4)
            else:
                _docx_adicionar_paragrafo(documento, "[Imagem 3]", tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, cor=RGBColor(68, 114, 196))
                _docx_adicionar_paragrafo(documento, "[Imagem 4]", tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, cor=RGBColor(68, 114, 196))
            _docx_adicionar_separador(documento)

        bloco = blocos.get(f"bloco_{numero}", {})
        _docx_adicionar_bloco(documento, bloco, numero)

        if numero == 1:
            # Assinatura institucional intermediária preservada no modelo DOCX.
            _docx_adicionar_separador(documento)
            _docx_adicionar_paragrafo(
                documento,
                "Entre em Contato e Solicite seu Orçamento!",
                tamanho=14,
                negrito=True,
                alinhamento=WD_ALIGN_PARAGRAPH.CENTER
            )
            _docx_adicionar_paragrafo(
                documento,
                "Conte com nossa equipe para esclarecer suas dúvidas e orientar a contratação de serviços e equipamentos de acordo com as necessidades do seu projeto.",
                tamanho=11,
                alinhamento=WD_ALIGN_PARAGRAPH.CENTER
            )
            _docx_adicionar_paragrafo(
                documento,
                "Entre em contato pelo WhatsApp, telefone ou e-mail. Teremos satisfação em conhecer sua demanda e apresentar nossas soluções.",
                tamanho=11,
                alinhamento=WD_ALIGN_PARAGRAPH.CENTER
            )

        # No modelo recebido, a lista de segmentos aparece após o terceiro bloco.
        if numero == 3:
            dados_segmentos = dict(dados_pagina)
            posicionamento = dados_segmentos.get("posicionamento_listas", {})
            if not isinstance(posicionamento, dict):
                posicionamento = {}
            try:
                pos_segmentos = int(posicionamento.get("bloco") or 3)
            except Exception:
                pos_segmentos = 3
            dados_segmentos["posicionamento_listas"] = {"bloco": pos_segmentos}
            _docx_inserir_segmentos_se_existirem(documento, dados_segmentos, numero)
        elif numero in (4, 5):
            _docx_inserir_segmentos_se_existirem(documento, dados_pagina, numero)
        # O bloco 1 termina com a chamada comercial; não inserir separador
        # aqui, pois a seção "Otimização" vem imediatamente em seguida.
        if numero != 1:
            _docx_adicionar_separador(documento)

    nome_site = str(dados_pagina.get("nome_site", "") or "").strip()
    titulo_cta_final = f"Fale com a {nome_site}" if nome_site else "Fale com nossa equipe"
    _docx_adicionar_paragrafo(documento, titulo_cta_final, tamanho=14, negrito=True)
    _docx_adicionar_paragrafo(
        documento,
        f"Entre em contato com nossos profissionais e solicite mais informações sobre {titulo}. Estamos à disposição!",
        tamanho=11,
        alinhamento=WD_ALIGN_PARAGRAPH.LEFT
    )

    base_nome = str(nome_arquivo_base or titulo or "documento").strip()
    nome_arquivo = re.sub(r"[^a-zA-Z0-9À-ÿ_-]+", "_", base_nome.lower()).strip("_") + ".docx"
    caminho = os.path.join(pasta_destino, nome_arquivo)
    documento.core_properties.title = titulo
    documento.core_properties.subject = "Página gerada a partir do conteudo.json"
    _docx_definir_idioma_ptbr(documento)
    documento.save(caminho)
    return caminho


# ============================================================
# INTERFACE
# ============================================================
           

def gerar_docx_interface():
    """Lê o conteudo.json e monta os DOCX sem regenerar o conteúdo."""
    pasta = entrada_pasta_docx.get().strip()
    if not pasta:
        messagebox.showwarning("DOCX", "Selecione a pasta destino dos DOCX.")
        return
    caminho_banco = ARQUIVO_BANCO
    if not os.path.exists(caminho_banco):
        messagebox.showerror("DOCX", "Banco de conteúdo não encontrado.")
        return
    try:
        with open(caminho_banco, "r", encoding="utf-8") as arquivo:
            banco = json.load(arquivo)
    except Exception as erro:
        messagebox.showerror("DOCX", f"Erro ao abrir banco:\n{erro}")
        return
    if not isinstance(banco, dict) or not banco:
        messagebox.showwarning("DOCX", "Banco vazio.")
        return

    total = 0
    erros = 0
    print("\n======================================")
    print("GERAR DOCX — FONTE ÚNICA: conteudo.json")
    print("TEMAS NO BANCO:", len(banco))
    print("PASTA DESTINO:", pasta)
    print("OLLAMA / PESQUISA / REGERAÇÃO: NÃO")
    print("======================================")

    for tema, dados in banco.items():
        if not isinstance(dados, dict):
            continue
        pagina_original = dados.get("pagina")
        if not isinstance(pagina_original, dict):
            pagina_original = dados
        pagina = dict(pagina_original)
        pagina.setdefault("tema", dados.get("tema", tema))
        pagina.setdefault("tags", dados.get("tags", []))
        pagina.setdefault("nome_site", dados.get("nome_site", ""))
        try:
            caminho = gerar_docx_mead(pagina, pasta, nome_arquivo_base=tema)
            total += 1
            print("DOCX GERADO:", tema, "->", caminho)
        except Exception as erro:
            erros += 1
            print("ERRO AO GERAR DOCX:", tema, repr(erro))

    messagebox.showinfo("DOCX", f"{total} arquivo(s) DOCX gerado(s).\nErros: {erros}")
    print("DOCX GERADOS:", total, "| ERROS:", erros)


ESTRUTURA_EDITORIAL = {
    "apresentacao": True,
    "funcionamento": True,
    "aplicacoes": True,
    "criterios": True,
    "comercial": True,
    "informacao_tecnica": True,
    "instalacao_execucao": True,
    "beneficios": True
}


def obter_estrutura_editorial():

    estrutura = {
        chave: bool(valor)
        for chave, valor
        in ESTRUTURA_EDITORIAL.items()
    }

    print()
    print("==============================")
    print("CHECKBOXES EDITORIAIS")
    print("==============================")

    for chave, valor in estrutura.items():

        print(
            f"{chave}: {valor}"
        )

    print("==============================")
    print("ESTRUTURA EDITORIAL ENVIADA")
    print("==============================")

    print(
        estrutura
    )

    print("==============================")

    return estrutura


def selecionar_todos_blocos_editoriais():

    for variavel in VARIAVEIS_EDITORIAIS.values():

        variavel.set(True)

    atualizar_estrutura_editorial()


def limpar_blocos_editoriais():

    for variavel in VARIAVEIS_EDITORIAIS.values():

        variavel.set(False)

    atualizar_estrutura_editorial()


def atualizar_estrutura_editorial():

    global ESTRUTURA_EDITORIAL

    for chave, variavel in VARIAVEIS_EDITORIAIS.items():

        ESTRUTURA_EDITORIAL[chave] = bool(
            variavel.get()
        )


    # ========================================================
    # 01. NOMES DOS BLOCOS EDITORIAIS
    # ========================================================

NOMES_BLOCOS_EDITORIAIS = {

    "apresentacao":
        "Apresentação e contexto",

    "funcionamento":
        "Funcionamento / Como funciona",

    "aplicacoes":
        "Aplicações / Situações atendidas",

    "criterios":
        "Critérios e diferenciais",

    "comercial":
        "Contexto comercial",

    "informacao_tecnica":
        "Informação técnica",

    "instalacao_execucao":
        "Instalação / Execução / Processo",

    "beneficios":
        "Benefícios"
}


    # ========================================================
    # 02. VARIÁVEIS DOS CHECKBOXES
    # ========================================================

VARIAVEIS_EDITORIAIS = {}




# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

def abrir_interface():

    global janela
    global entrada_site
    global entrada_grupo
    global entrada_banco
    global entrada_modelo
    global txt_sites
    global txt_palavras
    global barra_progresso
    global status_progresso
    global entrada_pasta_docx
    global status_tempo

    janela = tk.Tk()

    janela.title(
        "Gerador SEO MEAD"
    )

    janela.geometry(
        "1200x800"
    )


    # =====================================
    # PAINEL LATERAL
    # ESTRUTURA EDITORIAL
    # =====================================

    frame_editorial = tk.LabelFrame(
        janela,
        text="ESTRUTURA EDITORIAL",
        padx=12,
        pady=12
    )

    frame_editorial.pack(
        side="right",
        fill="y",
        padx=10,
        pady=10
    )


    # ========================================================
    # 01. TÍTULO DO PAINEL
    # ========================================================

    tk.Label(
        frame_editorial,
        text="Selecione os conteúdos\ndesta página:",
        font=("Arial", 10, "bold"),
        justify="left"
    ).pack(
        anchor="w",
        pady=(0, 10)
    )


    # ========================================================
    # 02. CHECKBOXES
    # ========================================================

    for chave, nome in NOMES_BLOCOS_EDITORIAIS.items():

        variavel = tk.BooleanVar(
            value=ESTRUTURA_EDITORIAL.get(
                chave,
                True
            )
        )

        VARIAVEIS_EDITORIAIS[chave] = variavel

        tk.Checkbutton(
            frame_editorial,
            text=nome,
            variable=variavel,
            anchor="w",
            justify="left",
            command=atualizar_estrutura_editorial
        ).pack(
            anchor="w",
            fill="x",
            pady=2
        )


    # ========================================================
    # 03. SEPARADOR
    # ========================================================

    tk.Frame(
        frame_editorial,
        height=2
    ).pack(
        fill="x",
        pady=10
    )


    # ========================================================
    # 04. SELECIONAR TODOS
    # ========================================================

    tk.Button(
        frame_editorial,
        text="SELECIONAR TODOS",
        command=selecionar_todos_blocos_editoriais
    ).pack(
        fill="x",
        pady=3
    )


    # ========================================================
    # 05. LIMPAR SELEÇÃO
    # ========================================================

    tk.Button(
        frame_editorial,
        text="LIMPAR SELEÇÃO",
        command=limpar_blocos_editoriais
    ).pack(
        fill="x",
        pady=3
    )


    # ========================================================
    # 06. NOME DO SITE
    # ========================================================

    tk.Label(
        janela,
        text="Nome do Site"
    ).pack()

    entrada_site = tk.Entry(
        janela,
        width=80
    )

    entrada_site.pack()

    entrada_site.bind(
        "<KeyRelease>",
        atualizar_banco_automatico
    )


    # ========================================================
    # 07. GRUPO
    # ========================================================

    tk.Label(
        janela,
        text="Grupo Principal do Projeto"
    ).pack()

    entrada_grupo = tk.Entry(
        janela,
        width=80
    )

    entrada_grupo.pack()


    # ========================================================
    # 08. BANCO
    # ========================================================

    tk.Label(
        janela,
        text="Arquivo Banco"
    ).pack()

    frame_banco = tk.Frame(
        janela
    )

    frame_banco.pack()

    entrada_banco = tk.Entry(
        frame_banco,
        width=70
    )

    entrada_banco.pack(
        side="left"
    )

    tk.Button(
        frame_banco,
        text="Selecionar",
        command=selecionar_banco
    ).pack(
        side="left"
    )


    # ========================================================
    # 09. MODELO DOCX
    # ========================================================

    tk.Label(
        janela,
        text="Modelo DOCX"
    ).pack()

    frame_modelo = tk.Frame(
        janela
    )

    frame_modelo.pack()

    entrada_modelo = tk.Entry(
        frame_modelo,
        width=70
    )

    entrada_modelo.pack(
        side="left"
    )

    tk.Button(
        frame_modelo,
        text="Selecionar",
        command=selecionar_modelo
    ).pack(
        side="left"
    )


    # ========================================================
    # 10. SITES
    # ========================================================

    tk.Label(
        janela,
        text="Sites para Consulta"
    ).pack()

    txt_sites = tk.Text(
        janela,
        height=8,
        width=100
    )

    txt_sites.pack()


    # ========================================================
    # 11. PALAVRAS
    # ========================================================

    tk.Label(
        janela,
        text="Palavras-chave"
    ).pack()

    txt_palavras = tk.Text(
        janela,
        height=15,
        width=100
    )

    txt_palavras.pack()


    # ========================================================
    # 12. STATUS
    # ========================================================

    status_progresso = tk.Label(
        janela,
        text="Aguardando..."
    )

    status_progresso.pack(
        pady=5
    )

    status_tempo = tk.Label(
        janela,
        text=""
    )

    status_tempo.pack(
        pady=2
    )

    barra_progresso = ttk.Progressbar(
        janela,
        length=700,
        mode="determinate"
    )

    barra_progresso.pack(
        pady=5
    )


    # ========================================================
    # 13. PASTA DOCX
    # ========================================================

    tk.Label(
        janela,
        text="Pasta destino dos DOCX:"
    ).pack()

    entrada_pasta_docx = tk.Entry(
        janela,
        width=60
    )

    entrada_pasta_docx.pack(
        pady=5
    )


    def selecionar_pasta_docx():

        pasta = filedialog.askdirectory()

        if pasta:

            entrada_pasta_docx.delete(
                0,
                tk.END
            )

            entrada_pasta_docx.insert(
                0,
                pasta
            )


    tk.Button(
        janela,
        text="SELECIONAR PASTA DOCX",
        command=selecionar_pasta_docx
    ).pack(
        pady=5
    )


    # ========================================================
    # 14. BOTÕES
    # ========================================================
    
    frame_botoes = tk.Frame(
        janela
    )
    
    frame_botoes.pack(
        pady=10
    )
    
    
    # ========================================================
    # 15. APAGAR PROGRESSO
    # ========================================================
    
    tk.Button(
        frame_botoes,
        text="APAGAR PROGRESSO",
        height=2,
        command=apagar_progresso_manual
    ).pack(
        side="left",
        padx=5
    )
    
    
    # ========================================================
    # 16. GERAR MATERIAL
    # ========================================================
    
    tk.Button(
        frame_botoes,
        text="GERAR MATERIAL",
        height=2,
        command=iniciar_geracao
    ).pack(
        side="left",
        padx=5
    )
    
    
    # ========================================================
    # 17. GERAR DOCX
    # ========================================================
    
    tk.Button(
        frame_botoes,
        text="GERAR DOCX",
        height=2,
        command=gerar_docx_interface
    ).pack(
        side="left",
        padx=5
    )


    janela.mainloop()



# ============================================================
# EXECUÇÃO
# ============================================================

def executar():

    print("VERSÃO MEAD: 7.15 — AUDITORIA DE RESERVAS E JSON")
    print("ARQUIVO PYTHON EM EXECUÇÃO:", os.path.abspath(__file__))
    print("JSON OFICIAL CONFIGURADO:", ARQUIVO_BANCO)

    global etapa_atual
    global inicio_geracao
    global tempos_etapas
    global etapas_total

    inicio_geracao = time.time()

    etapa_atual = 0
    tempos_etapas = []
    etapas_total = 4
    PAGINAS_EM_PROCESSAMENTO.clear()

    # ========================================================
    # 01. TEMA
    # ========================================================

    tema = input(
        "Tema: "
    ).strip()

    if not tema:

        print()
        print("==============================")
        print("TEMA NÃO INFORMADO")
        print("==============================")

        return

    # ========================================================
    # 02. IDENTIFICAR GRUPO
    # ========================================================

    grupo = identificar_grupo_tema(
        tema
    )

    print()
    print("==============================")
    print("GRUPO IDENTIFICADO")
    print("==============================")

    print(
        grupo
    )

    # ========================================================
    # 03. GERAR ENTENDIMENTO DO PRODUTO
    # ========================================================

    print()
    print("==============================")
    print("GERANDO ENTENDIMENTO DO PRODUTO")
    print("==============================")

    print(
        "TEMA:",
        tema
    )

    print(
        "GRUPO:",
        grupo
    )

    print()
    print("CHAMANDO:")
    print(
        "gerar_entendimento_produto()"
    )

    inicio_entendimento = time.time()

    entendimento = gerar_entendimento_produto(
        tema,
        grupo
    )

    fim_entendimento = time.time()

    # ========================================================
    # 04. RESULTADO DO ENTENDIMENTO
    # ========================================================

    print()
    print("==============================")
    print("ENTENDIMENTO DO PRODUTO RETORNOU")
    print("==============================")

    print(
        "TIPO:",
        type(entendimento)
    )

    print(
        "CARACTERES:",
        len(
            str(
                entendimento or ""
            )
        )
    )

    print(
        "TEMPO:",
        formatar_tempo(
            fim_entendimento
            -
            inicio_entendimento
        )
    )

    print()

    if entendimento:

        print(
            "ENTENDIMENTO:"
        )

        print(
            entendimento
        )

    else:

        print(
            "ENTENDIMENTO VAZIO"
        )

    print()
    print("==============================")
    print("FIM ENTENDIMENTO DO PRODUTO")
    print("==============================")

    # ========================================================
    # 05. PESQUISA
    # ========================================================

    print()
    print("==============================")
    print("PESQUISANDO")
    print("==============================")

    urls = pesquisar_completo(
        tema
    )

    etapa_atual += 1

    tempo = time.time() - inicio_geracao

    tempos_etapas.append(
        tempo
    )

    media = (
        sum(tempos_etapas)
        /
        len(tempos_etapas)
    )

    restante = media * (
        etapas_total
        -
        etapa_atual
    )

    atualizar_progresso(
        25,
        f"Pesquisa concluída | Restante: {formatar_tempo(restante)}"
    )

    # ========================================================
    # 06. FONTES ENCONTRADAS
    # ========================================================

    print()
    print("==============================")
    print("FONTES ENCONTRADAS")
    print("==============================")

    print(
        len(urls)
    )

    paginas = []

    textos = []

    # ========================================================
    # 07. COLETAR FONTES
    # ========================================================

    for i, url in enumerate(
        urls,
        start=1
    ):

        print()

        print(
            f"{i}/{len(urls)}",
            url
        )

        dados = coletar_pagina(
            url
        )

    # ========================================================
    # 08. NORMALIZAR RETORNO DA COLETA
    # ========================================================

        texto = ""

        tipo = "html"

        if isinstance(
            dados,
            dict
        ):

            texto = dados.get(
                "texto",
                ""
            )

            tipo = dados.get(
                "tipo",
                "html"
            )

        elif isinstance(
            dados,
            str
        ):

            texto = dados

        if not texto:

            texto = ""

        print(
            "CARACTERES:",
            len(texto)
        )

    # ========================================================
    # 09. VALIDAR RELEVÂNCIA DA FONTE
    # ========================================================

        texto_lower = texto.lower()

        palavras_tema = [

            p

            for p in tema.lower().split()

            if len(p) > 3

        ]

        ocorrencias = 0

        for palavra in palavras_tema:

            if palavra in texto_lower:

                ocorrencias += 1

    # ========================================================
    # 10. TERMOS TÉCNICOS POR GRUPO
    # ========================================================

        TERMOS_GRUPOS = {

            "hidraulica": [

                "hidraulica",
                "hidráulica",
                "valvula",
                "válvula",
                "pressao",
                "pressão",
                "atuador",
                "fluido",
                "bomba",
                "vazao",
                "vazão",
                "cilindro"

            ],

            "blindagem": [

                "blindagem",
                "balistico",
                "balístico",
                "protecao",
                "proteção",
                "vidro blindado",
                "nivel iii",
                "nível iii"

            ],

            "construcao": [

                "concreto",
                "argamassa",
                "cimento",
                "estrutura",
                "obra",
                "fundacao",
                "fundação"

            ]

        }

        termos_grupo = TERMOS_GRUPOS.get(
            grupo,
            []
        )

        pontos_grupo = 0

        for termo in termos_grupo:

            if termo in texto_lower:

                pontos_grupo += 1

    # ========================================================
    # 11. APROVAÇÃO FINAL
    # ========================================================

        print()
        print("==============================")
        print("ANTES APROVAÇÃO FINAL")
        print("==============================")

        print(
            "TIPO:",
            tipo
        )

        print(
            "TEXTO:",
            len(texto)
        )

        print(
            "OCORRENCIAS:",
            ocorrencias
        )

        print(
            "TERMOS GRUPO:",
            termos_grupo
        )

        print(
            "PONTOS GRUPO:",
            pontos_grupo
        )

        if (

            len(texto) > 1500

            and ocorrencias >= max(
                2,
                len(palavras_tema) // 2
            )

            and (

                len(termos_grupo) == 0

                or

                pontos_grupo >= 3

            )

        ):

            print()
            print("==============================")
            print("ENTROU NA APROVAÇÃO")
            print("==============================")

            print(
                "URL:",
                url
            )

            print(
                "TIPO:",
                tipo
            )

            print(
                "CARACTERES:",
                len(texto)
            )

            paginas.append({

                "url":
                    url,

                "tipo":
                    tipo,

                "caracteres":
                    len(texto),

                "texto":
                    texto

            })

            textos.append({

                "url":
                    url,

                "tipo":
                    tipo,

                "texto":
                    texto[:30000]

            })

            print()
            print("APROVADO")

            print(
                "CARACTERES:",
                len(texto)
            )

            print(
                "PALAVRAS TEMA:",
                ocorrencias
            )

            print(
                "PONTOS GRUPO:",
                pontos_grupo
            )

        else:

            print()
            print("DESCARTADO")

            print(
                "URL:",
                url
            )

            print(
                "CARACTERES:",
                len(texto)
            )

            print(
                "PALAVRAS TEMA:",
                ocorrencias
            )

            print(
                "PONTOS GRUPO:",
                pontos_grupo
            )

    # ========================================================
    # 12. IDENTIFICAR EMPRESAS NOS SITES DAS FONTES
    # ========================================================
    
    if paginas:
    
        print()
        print("==============================")
        print("IDENTIFICANDO EMPRESAS NOS SITES")
        print("==============================")
    
        paginas = identificar_empresa_no_site(
            paginas
        )
        
    
    # ========================================================
    # 12.a SALVAR NOVO BRUTO
    # ========================================================

    from urllib.parse import urlparse

    identidades_por_dominio = {}

    for item in paginas:

        if not isinstance(item, dict):
            continue

        url = str(
            item.get("url", "")
        ).strip()

        if not url:
            continue

        try:

            dominio = urlparse(
                url
            ).netloc.lower()

            dominio = re.sub(
                r"^www\.",
                "",
                dominio
            ).strip()

        except Exception:

            dominio = ""

        if not dominio:
            continue

        identidade = item.get(
            "identidade_fonte",
            {}
        )

        if not isinstance(
            identidade,
            dict
        ):
            continue

        nome = str(
            identidade.get(
                "nome",
                ""
            )
        ).strip()

        # ----------------------------------------------------
        # NÃO SUBSTITUIR IDENTIDADE VÁLIDA POR VAZIA
        # ----------------------------------------------------

        if nome:

            identidades_por_dominio[
                dominio
            ] = identidade

        elif dominio not in identidades_por_dominio:

            identidades_por_dominio[
                dominio
            ] = identidade


    # --------------------------------------------------------
    # APLICAR IDENTIDADE AOS TEXTOS PELO DOMÍNIO
    # --------------------------------------------------------

    for item in textos:

        if not isinstance(
            item,
            dict
        ):
            continue

        url = str(
            item.get("url", "")
        ).strip()

        if not url:
            continue

        try:

            dominio = urlparse(
                url
            ).netloc.lower()

            dominio = re.sub(
                r"^www\.",
                "",
                dominio
            ).strip()

        except Exception:

            dominio = ""

        if not dominio:
            continue

        identidade = (
            identidades_por_dominio.get(
                dominio,
                {}
            )
        )

        if not isinstance(
            identidade,
            dict
        ):
            continue

        nome = str(
            identidade.get(
                "nome",
                ""
            )
        ).strip()

        if nome:

            item[
                "identidade_fonte"
            ] = identidade


    if paginas:

        print()
        print("==============================")
        print("SALVANDO NOVO BRUTO")
        print("==============================")

        salvar_bruto(
            tema,
            paginas
        )

    else:

        print()
        print("==============================")
        print("BRUTO EXISTENTE PRESERVADO")
        print("==============================")
        

    # ========================================================
    # 13. ATUALIZAR PROGRESSO
    # ========================================================

    etapa_atual += 1

    tempo = time.time() - inicio_geracao

    tempos_etapas.append(
        tempo
    )

    media = (
        sum(tempos_etapas)
        /
        len(tempos_etapas)
    )

    restante = media * (
        etapas_total
        -
        etapa_atual
    )

    atualizar_progresso(
        50,
        f"Fontes coletadas | Restante: {formatar_tempo(restante)}"
    )

    # ========================================================
    # 14. RESUMO DA COLETA
    # ========================================================

    print()
    print("==============================")
    print("TESTE DE COLETA FINALIZADO")
    print("==============================")

    print()

    print(
        "PÁGINAS COLETADAS:",
        len(paginas)
    )

    print()

    for i, item in enumerate(
        paginas,
        start=1
    ):

        print(
            f"{i} - {item['url']}"
        )

    print()

    print(
        "TEXTOS CAPTURADOS:",
        len(textos)
    )

    print()

    for i, texto in enumerate(
        textos,
        start=1
    ):

        print(
            f"{i} - {len(texto.get('texto', ''))} caracteres"
        )

    # ========================================================
    # 15. GARANTIR FONTES PARA MAPA
    # ========================================================

    if not textos:

        print()
        print("==============================")
        print("NENHUM TEXTO DISPONÍVEL")
        print("==============================")

        return

    # ========================================================
    # 16. NORMALIZAR TEXTOS PARA MAPA
    # ========================================================

    textos_para_mapa = []

    for item in textos:

        if isinstance(
            item,
            dict
        ):

            texto_item = item.get(
                "texto",
                ""
            )

            if texto_item:

                textos_para_mapa.append({

                    "url":
                        item.get(
                            "url",
                            ""
                        ),

                    "tipo":
                        item.get(
                            "tipo",
                            "texto"
                        ),

                    "texto":
                        texto_item

                })

        elif isinstance(
            item,
            str
        ):

            if item.strip():

                textos_para_mapa.append({

                    "url":
                        "",

                    "tipo":
                        "texto",

                    "texto":
                        item

                })

    
    # ========================================================
    # 17. LIMPAR REFERÊNCIAS
    # ========================================================

    try:

        textos_para_mapa = limpar_lista_referencias(
            textos_para_mapa,
            tema
        )

    except Exception as e:

        print()
        print("==============================")
        print("AVISO - LIMPEZA DO MAPA")
        print("==============================")

        print(
            repr(e)
        )

    # ========================================================
    # 18. CORRIGIR FORMATO FINAL
    # ========================================================

    textos_para_mapa_corrigidos = []

    for item in textos_para_mapa:

        if isinstance(
            item,
            dict
        ):

            texto_item = item.get(
                "texto",
                ""
            )

            if texto_item and texto_item.strip():

                textos_para_mapa_corrigidos.append({

                    "url":
                        item.get(
                            "url",
                            ""
                        ),

                    "tipo":
                        item.get(
                            "tipo",
                            "texto"
                        ),

                    "texto":
                        texto_item,

                    # ====================================================
                    # PRESERVAR IDENTIDADE DA FONTE
                    # ====================================================
                    #
                    # A identidade já foi identificada anteriormente
                    # na etapa de limpeza das referências.
                    #
                    # Aqui apenas transportamos o objeto.
                    #
                    # NÃO identificar novamente.
                    # NÃO inferir fabricante pelo domínio.
                    # ====================================================

                    "identidade_fonte":
                        item.get(
                            "identidade_fonte",
                            {
                                "nome": "",
                                "dominio": "",
                                "papel": "",
                                "origem_identificacao": "",
                                "confianca": "baixa"
                            }
                        )

                })

        elif isinstance(
            item,
            str
        ):

            if item.strip():

                textos_para_mapa_corrigidos.append({

                    "url":
                        "",

                    "tipo":
                        "texto",

                    "texto":
                        item,

                    "identidade_fonte":
                        {
                            "nome": "",
                            "dominio": "",
                            "papel": "",
                            "origem_identificacao": "",
                            "confianca": "baixa"
                        }

                })

    print()
    print("==============================")
    print("TEXTOS PARA MAPA MEAD")
    print("==============================")

    print(
        "QUANTIDADE:",
        len(
            textos_para_mapa_corrigidos
        )
    )

    # ========================================================
    # 19. GERAR MAPA MEAD
    # ========================================================

    print()
    print("==============================")
    print("GERANDO MAPA MEAD")
    print("==============================")

    mapa_mead = gerar_e_salvar_mapa_mead(
        tema,
        textos_para_mapa_corrigidos
    )

    # ========================================================
    # 20. VALIDAR MAPA
    # ========================================================

    print()
    print("==============================")
    print("RESULTADO MAPA MEAD")
    print("==============================")

    print(
        "TIPO:",
        type(mapa_mead)
    )

    print(
        "CARACTERES:",
        len(
            str(
                mapa_mead or ""
            )
        )
    )

    if mapa_mead:

        print()

        print(
            mapa_mead[:3000]
        )

    else:

        print(
            "MAPA MEAD VAZIO"
        )

        print()
        print("==============================")
        print("CONTEÚDO NÃO SERÁ GERADO")
        print("==============================")

        return

    # ========================================================
    # 21. GERAR CONTEÚDO
    # ========================================================

    print()
    print("==============================")
    print("GERANDO CONTEÚDO COM IA")
    print("==============================")

    print()
    print("==============================")
    print("GERANDO CONTEÚDO COMPLETO")
    print("==============================")

    etapa_atual += 1

    tempo = time.time() - inicio_geracao

    tempos_etapas.append(
        tempo
    )

    media = (
        sum(tempos_etapas)
        /
        len(tempos_etapas)
    )

    restante = media * (
        etapas_total
        -
        etapa_atual
    )

    atualizar_progresso(
        75,
        f"Gerando conteúdo IA | Restante: {formatar_tempo(restante)}"
    )

    # ========================================================
    # 22. CHAMAR GERADOR
    # ========================================================

    print()
    print("==============================")
    print("CHAMANDO GERADOR DE CONTEÚDO")
    print("==============================")

    print(
        "TEXTOS:",
        len(textos)
    )

    print(
        "TEMA:",
        tema
    )

    print(
        "MAPA:",
        len(
            str(
                mapa_mead
            )
        )
    )

    print(
        "ENTENDIMENTO:",
        len(
            str(
                entendimento or ""
            )
        ),
        "caracteres"
    )

    # ========================================================
    # 22.1 VALIDAR DADOS ANTES DO GERADOR
    # ========================================================

    if not textos:

        print()
        print("==============================")
        print("NENHUM TEXTO DISPONÍVEL")
        print("==============================")

        return None

    if not mapa_mead:

        print()
        print("==============================")
        print("MAPA MEAD NÃO DISPONÍVEL")
        print("==============================")

        return None

    print()
    print("==============================")
    print("DADOS ENVIADOS AO GERADOR")
    print("==============================")

    print(
        "TEXTOS BRUTOS:",
        len(textos)
    )

    print(
        "MAPA MEAD:",
        len(
            str(
                mapa_mead
            )
        ),
        "caracteres"
    )

    print(
        "ESTRUTURA EDITORIAL:",
        obter_estrutura_editorial()
    )

    # ========================================================
    # 22.2 CAPTURAR NOME DO SITE
    # ========================================================
    
    nome_site = (
        entrada_site.get().strip()
        if "entrada_site" in globals()
        else ""
    )
    
    print(
        "NOME DO SITE:",
        nome_site
    )
    
    # ========================================================
    # 22.3 CHAMAR GERADOR
    # ========================================================
    
    conteudo_completo = gerar_conteudo_completo(
        tema,
        textos,
        mapa_mead,
        obter_estrutura_editorial(),
        nome_site=nome_site
    )

    # ========================================================
    # 22.4 VALIDAR RETORNO DO GERADOR
    # ========================================================

    print()
    print("==============================")
    print("RETORNO DO GERADOR")
    print("==============================")

    if not conteudo_completo:

        print(
            "GERADOR NÃO RETORNOU CONTEÚDO"
        )

        return None

    conteudo_completo = str(
        conteudo_completo
    ).strip()

    print(
        "TIPO:",
        type(
            conteudo_completo
        ).__name__
    )

    print(
        "CARACTERES:",
        len(
            conteudo_completo
        )
    )

    # ========================================================
    # 23. RESULTADO CONTEÚDO
    # ========================================================

    print()
    print("==============================")
    print("RESULTADO CONTEÚDO")
    print("==============================")

    print(
        "TIPO:",
        type(conteudo_completo)
    )

    print(
        "CARACTERES:",
        len(
            str(
                conteudo_completo or ""
            )
        )
    )



    # ========================================================
    # 25. FINAL
    # ========================================================

    print()
    print("==============================")
    print("CONTEÚDO FINALIZADO")
    print("==============================")

    print(
        "TEMA:",
        tema
    )

    print(
        "ENTENDIMENTO GERADO:",
        bool(entendimento)
    )

    print(
        "MAPA MEAD GERADO:",
        bool(mapa_mead)
    )

    print(
        "CONTEÚDO GERADO:",
        bool(conteudo_completo)
    )

    print()
    print(
        "TEMPO TOTAL:",
        formatar_tempo(
            time.time()
            -
            inicio_geracao
        )
    )

    # ========================================================
    # 26. EXECUÇÃO DO PROGRAMA
    # ========================================================

if __name__ == "__main__":

    abrir_interface()
