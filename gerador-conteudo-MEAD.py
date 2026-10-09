# gerador-conteudo-MEAD-v9.18-Lux-v71-ok.py - 09/10/2026
# Reservatório ampliado de candidatos limpos — seleção final continua em 15
# Pré-barreira de URLs + seleção natural de trechos + bruto sem descartados


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
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup
from ddgs import DDGS



ARQUIVO_BANCO = r"C:\Python\gerador-conteudo\banco\conteudo-site.json"

ARQUIVO_MEAD = r"C:\Python\gerador-conteudo\mead\mead.json"

ARQUIVO_PROGRESSO = r"C:\Python\gerador-conteudo\banco\progresso.json"

janela = None

IA_PROCESSANDO = False
PROCESSAMENTO_ATIVO = False
etapa_atual = 0
inicio_geracao = 0.0
tempos_etapas = []
etapas_total = 4
PAGINAS_EM_PROCESSAMENTO = set()

# Entidades empresariais/modelos identificadas em outras fontes da mesma página.
# São usadas apenas como barreira de saída do Ollama; não entram no prompt editorial.
ENTIDADES_PROIBIDAS_PAGINA = set()

from pathlib import Path
from urllib.parse import urlparse

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
    "[tema] durável"
]

TAGS_FIXAS_SERVICOS = [
    "empresa especializada em [tema]",
    "especialista em [tema]",
    "realizamos [tema]",
    "executamos [tema]",
    "prestação de [tema]",
    "profissional de [tema]",
    "onde encontrar [tema]",
    "onde contratar [tema]",
    "orçamento para [tema]",
    "cotação para [tema]",
    "[tema] com atendimento especializado",
    "[tema] com suporte técnico",
    "[tema] preço",
    "manutenção de [tema]",
    "assistência técnica em [tema]"
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

        print("\nMEAD carregado:")
        print(mead)

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
# CARREGAR MEAD NA MEMÓRIA
# ============================================================

MEAD = carregar_mead()



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

        identidade = mead.get("identidade", {})
        principio = mead.get("principio_operacional", {})
        preparacao = mead.get("preparacao", {})
        tags = preparacao.get("tags", {})
        construcao = mead.get("construcao_conteudo", {})
        estrutura = mead.get("estrutura_pagina_editorial", {})
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

                    "fragmentos_autorizados":
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

                    "fragmentos_autorizados":
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

                    "fragmentos_autorizados":
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

                    "fragmentos_autorizados":
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

                    "fragmentos_autorizados":
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
    - preservar fragmentos_autorizados;
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
        # POSICIONAMENTO DAS IMAGENS 3 E 4
        #
        # As imagens 1 e 2 permanecem fixas no topo.
        # As imagens 3 e 4 recebem posições distintas entre
        # os blocos inferiores e essa decisão fica persistida
        # no JSON para o DOCX apenas reproduzir a escolha.
        # ====================================================
        posicionamento_imagens = dados_pagina.get(
            "posicionamento_imagens",
            {}
        )

        if not isinstance(posicionamento_imagens, dict):
            posicionamento_imagens = {}

        imagens_3_4_existentes = []
        for numero_imagem in (3, 4):
            item_imagem = imagens_existentes.get(
                f"imagem_{numero_imagem}",
                {}
            ) if isinstance(imagens_existentes, dict) else {}
            if isinstance(item_imagem, dict):
                if str(item_imagem.get("arquivo", "") or "").strip() or str(item_imagem.get("url", "") or "").strip():
                    imagens_3_4_existentes.append(numero_imagem)

        if len(imagens_3_4_existentes) == 2:
            # Imagem 3 e imagem 4 formam UM único conjunto editorial.
            # Portanto, as duas ocupam sempre a mesma posição e serão
            # renderizadas lado a lado no DOCX. A posição sorteada é
            # entre os blocos inferiores: depois do bloco 2, 3 ou 4.
            posicoes_disponiveis = [3, 4, 5]
            try:
                posicao_par = int(
                    posicionamento_imagens.get("imagem_3", 0) or 0
                )
            except Exception:
                posicao_par = 0

            try:
                posicao_4 = int(
                    posicionamento_imagens.get("imagem_4", 0) or 0
                )
            except Exception:
                posicao_4 = 0

            if posicao_par not in posicoes_disponiveis or posicao_4 != posicao_par:
                posicao_par = random.choice(posicoes_disponiveis)

            posicionamento_imagens = {
                "imagem_3": posicao_par,
                "imagem_4": posicao_par
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

            fragmentos_autorizados = (
                bloco_recebido.get(
                    "fragmentos_autorizados",
                    []
                )
            )


            if not isinstance(
                fragmentos_autorizados,
                list
            ):

                fragmentos_autorizados = []


            fragmentos_autorizados = list(
                fragmentos_autorizados[:3]
            )


            while len(
                fragmentos_autorizados
            ) < 3:

                fragmentos_autorizados.append(
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
            # fragmentos_autorizados
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

                "fragmentos_autorizados":
                    fragmentos_autorizados,

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
                posicionamento_listas,

            "posicionamento_imagens":
                posicionamento_imagens
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
print(MEAD_TEXTO)


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

    # ========================================================
    # v9.18-v9 — PESQUISA EXCLUSIVA EM FONTES BRASILEIRAS
    # ========================================================
    # As consultas são conduzidas em português e com sinal de Brasil.
    # A barreira .br é aplicada sobre as URLs retornadas, evitando
    # depender de `site:.br` no mecanismo do DDGS.
    #
    # A consulta também recebe termos em português quando
    # necessário, mas a garantia principal é a URL final:
    # somente domínios .br podem entrar no patrimônio.
    # ========================================================

    dominios_bloqueados = [

        "pinterest.",
        "facebook.",
        "instagram.",
        "youtube.com",
        "youtu.be",
        "tiktok.",
        "linkedin.com",
        "twitter.com",
        "x.com",
        "reddit.com",
        "threads.net",
        "telegram.me",
        "t.me"

    ]

    consulta_original = str(
        termo or ""
    ).strip()

    if not consulta_original:
        return resultados

    # --------------------------------------------------------
    # IMPORTANTE — v9.18-v9
    # --------------------------------------------------------
    # NÃO usar `site:.br` dentro da consulta do DDGS.
    # Alguns mecanismos/fontes do DDGS retornam "No results
    # found" quando recebem esse operador, fazendo a pesquisa
    # inteira parecer vazia.
    #
    # A restrição brasileira é aplicada DEPOIS do retorno, na
    # URL. Assim mantemos a origem exclusivamente brasileira sem
    # depender da sintaxe do mecanismo de busca.
    # --------------------------------------------------------

    consultas_tentativa = []

    consulta_base = consulta_original

    # Aumenta o sinal de Brasil/português sem usar operador de
    # domínio. Isso ajuda o buscador a priorizar conteúdo nacional.
    if "brasil" not in consulta_base.casefold():
        consultas_tentativa.append(
            f"{consulta_base} Brasil"
        )

    consultas_tentativa.append(
        consulta_base
    )

    # Remove duplicatas preservando a ordem.
    consultas_tentativa = list(dict.fromkeys(
        consultas_tentativa
    ))

    try:

        with DDGS() as ddgs:

            for consulta in consultas_tentativa:

                try:

                    busca = ddgs.text(
                        consulta,
                        safesearch="off",
                        max_results=max(limite, 10)
                    )

                    encontrados_nesta_tentativa = 0

                    for item in busca:

                        if not isinstance(item, dict):
                            continue

                        url = item.get("href")

                        if not url:
                            continue

                        try:
                            parsed = urlparse(str(url))
                            dominio = (
                                parsed.netloc or ""
                            ).lower().strip().split(":", 1)[0]
                        except Exception:
                            continue

                        # ------------------------------------------------
                        # BARREIRA ABSOLUTA DE ORIGEM
                        # ------------------------------------------------
                        # Só entra domínio brasileiro .br.
                        # Ex.: www.gov.br, abnt.org.br, empresa.com.br
                        # ------------------------------------------------
                        if not dominio.endswith(".br"):
                            continue

                        if any(
                            dominio == bloqueado.rstrip(".")
                            or dominio.endswith(bloqueado.rstrip("."))
                            or bloqueado in dominio
                            for bloqueado in dominios_bloqueados
                        ):
                            continue

                        url_limpa = str(url).strip()

                        if url_limpa not in resultados:
                            resultados.append(url_limpa)
                            encontrados_nesta_tentativa += 1

                        if len(resultados) >= limite:
                            return resultados[:limite]

                    # Se já encontramos fontes brasileiras, não há
                    # necessidade de ampliar a consulta.
                    if encontrados_nesta_tentativa > 0:
                        break

                except Exception as erro_tentativa:

                    # "No results found" não deve abortar a pesquisa.
                    # Tenta a próxima formulação.
                    print()
                    print("FALHA NA CONSULTA:")
                    print(consulta)
                    print("MOTIVO:", erro_tentativa)
                    if len(consultas_tentativa) > 1:
                        continue

    except Exception as e:

        print()
        print("ERRO NA PESQUISA:")
        print(e)

    return resultados[:limite]


# ============================================================

# ============================================================
# PRÉ-BARREIRA DE URLS — v9.15
# ============================================================
#
# A URL passa por esta barreira ANTES de qualquer
# coletar_pagina(url).
#
# A ausência do tema na URL NÃO reprova a fonte.
# A relevância temática continua sendo validada depois,
# sobre o conteúdo real coletado.
# ============================================================

def avaliar_url_pre_download(url, tema=None):

    try:

        url_original = str(
            url or ""
        ).strip()

        if not url_original:
            return False

        from urllib.parse import urlparse

        parsed = urlparse(
            url_original
        )

        esquema = (
            parsed.scheme or ""
        ).lower().strip()

        if esquema not in (
            "http",
            "https"
        ):
            return False

        dominio = (
            parsed.netloc or ""
        ).lower().strip()

        caminho = (
            parsed.path or ""
        ).lower().strip()

        consulta = (
            parsed.query or ""
        ).lower().strip()

        # ----------------------------------------------------
        # DOMÍNIOS JÁ BLOQUEADOS PELA v9.14
        # ----------------------------------------------------

        dominios_bloqueados = [

            "mercadolivre.",
            "amazon.",
            "shopee.",
            "aliexpress.",
            "alibaba.",
            "ebay.",
            "made-in-china.",

            "pinterest.",
            "facebook.",
            "instagram.",
            "youtube.",
            "youtu.be",
            "tiktok.",

            "oceanofpdf.com",
            "pdfcoffee.com",
            "scribd.com"

        ]

        if any(
            item in dominio
            for item in dominios_bloqueados
        ):
            return False

        # ----------------------------------------------------
        # REDES SOCIAIS
        # ----------------------------------------------------

        dominios_sociais = [

            "linkedin.com",
            "twitter.com",
            "x.com",
            "reddit.com",
            "threads.net",
            "telegram.me",
            "t.me"

        ]

        if any(
            item in dominio
            for item in dominios_sociais
        ):
            return False

        # ----------------------------------------------------
        # LOGIN / CADASTRO / AUTENTICAÇÃO
        # ----------------------------------------------------

        caminhos_autenticacao = [

            "login",
            "log-in",
            "signin",
            "sign-in",
            "sign_up",
            "signup",
            "sign-up",
            "cadastro",
            "cadastrar",
            "registrar",
            "registro",
            "register",
            "account",
            "conta",
            "minha-conta",
            "my-account",
            "auth",
            "autenticacao",
            "autenticação",
            "password",
            "senha"

        ]

        if any(
            termo in caminho
            for termo in caminhos_autenticacao
        ):
            return False

        # ----------------------------------------------------
        # CHECKOUT / CARRINHO / TRANSAÇÃO
        # ----------------------------------------------------

        caminhos_transacionais = [

            "checkout",
            "carrinho",
            "cart",
            "basket",
            "pedido",
            "orders",
            "order",
            "pagamento",
            "payment",
            "finalizar-compra",
            "finalizar_compra",
            "comprar-agora",
            "buy-now"

        ]

        if any(
            termo in caminho
            for termo in caminhos_transacionais
        ):
            return False

        # ----------------------------------------------------
        # CONTATO / ATENDIMENTO
        # ----------------------------------------------------

        caminhos_contato = [

            "contato",
            "contact",
            "fale-conosco",
            "fale_conosco",
            "sac",
            "atendimento"

        ]

        if any(
            termo in caminho
            for termo in caminhos_contato
        ):
            return False

        # ----------------------------------------------------
        # POLÍTICAS / PRIVACIDADE / COOKIES / TERMOS
        # ----------------------------------------------------

        caminhos_politicas = [

            "politica-de-privacidade",
            "politica_privacidade",
            "politica-de-cookies",
            "politica_cookies",
            "privacy-policy",
            "privacy_policy",
            "privacidade",
            "cookies",
            "termos-de-uso",
            "termos_de_uso",
            "terms-of-use",
            "terms_of_use",
            "termos-e-condicoes",
            "terms-and-conditions",
            "lgpd"

        ]

        if any(
            termo in caminho
            for termo in caminhos_politicas
        ):
            return False

        # ----------------------------------------------------
        # BUSCA INTERNA
        # ----------------------------------------------------

        caminhos_busca = [

            "/busca",
            "/buscar",
            "/search",
            "/resultado-busca",
            "/resultados-busca",
            "/search-results"

        ]

        if any(
            termo in caminho
            for termo in caminhos_busca
        ):
            return False

        parametros_busca = [

            "search=",
            "query=",
            "q=",
            "s=",
            "keyword=",
            "keywords=",
            "busca=",
            "buscar="

        ]

        if any(
            parametro in consulta
            for parametro in parametros_busca
        ):
            return False

        # ----------------------------------------------------
        # AUTOR / TAG / FEED / ARQUIVOS AUTOMÁTICOS
        # ----------------------------------------------------

        caminhos_arquivos_editoriais = [

            "/author/",
            "/autor/",
            "/tag/",
            "/tags/",
            "/feed",
            "/rss",
            "/category/",
            "/categoria/",
            "/arquivo/",
            "/archives/",
            "/wp-json/",
            "/xmlrpc.php"

        ]

        if any(
            termo in caminho
            for termo in caminhos_arquivos_editoriais
        ):
            return False

        # ----------------------------------------------------
        # ENDPOINTS SEM INTERESSE EDITORIAL
        # ----------------------------------------------------

        endpoints_tecnicos = [

            "/api/",
            "/ajax/",
            "/graphql",
            "/wp-admin/",
            "/wp-login.php"

        ]

        if any(
            termo in caminho
            for termo in endpoints_tecnicos
        ):
            return False

        # ----------------------------------------------------
        # IMAGENS / VÍDEOS / ÁUDIO / BINÁRIOS
        #
        # PDF permanece permitido porque a v9.14 pesquisa
        # e coleta PDF como fonte editorial.
        # ----------------------------------------------------

        extensao = ""

        ultimo_ponto = caminho.rfind(".")

        if ultimo_ponto >= 0:
            extensao = caminho[
                ultimo_ponto:
            ].lower()

        extensoes_nao_editoriais = [

            ".jpg",
            ".jpeg",
            ".png",
            ".gif",
            ".webp",
            ".svg",
            ".bmp",
            ".ico",
            ".tif",
            ".tiff",
            ".avif",

            ".mp4",
            ".webm",
            ".mov",
            ".avi",
            ".mkv",
            ".mpeg",
            ".mpg",
            ".m4v",

            ".mp3",
            ".wav",
            ".ogg",
            ".oga",
            ".m4a",
            ".flac",

            ".zip",
            ".rar",
            ".7z",
            ".tar",
            ".gz",
            ".exe",
            ".msi",
            ".apk",
            ".dmg",
            ".iso",

            ".doc",
            ".docx",
            ".xls",
            ".xlsx",
            ".ppt",
            ".pptx"

        ]

        if extensao in extensoes_nao_editoriais:
            return False

        # ----------------------------------------------------
        # DOMÍNIO OBRIGATÓRIO
        # ----------------------------------------------------

        if not dominio:
            return False

        # ----------------------------------------------------
        # IMPORTANTE:
        # NÃO verificar se o tema aparece na URL.
        # A validação temática ocorre no conteúdo.
        # ----------------------------------------------------

        return True

    except Exception:
        # Se a própria avaliação falhar, não permitir download.
        return False


def filtrar_urls_pre_download(
    urls,
    tema=None
):

    """
    Aplica a pré-barreira depois de filtrar_urls()
    e antes de qualquer download.
    """

    if not isinstance(
        urls,
        list
    ):
        return []

    resultado = []

    vistas = set()

    bloqueadas = 0

    for url in urls:

        url_limpa = str(
            url or ""
        ).strip()

        if not url_limpa:
            continue

        chave = url_limpa.casefold()

        if chave in vistas:
            continue

        vistas.add(chave)

        if not avaliar_url_pre_download(
            url_limpa,
            tema
        ):

            bloqueadas += 1

            print(
                "PRÉ-BARREIRA BLOQUEOU:",
                url_limpa
            )

            continue

        resultado.append(
            url_limpa
        )

    print()
    print("==============================")
    print("PRÉ-BARREIRA DE URLS — v9.15")
    print("==============================")
    print(
        "RECEBIDAS:",
        len(urls)
    )
    print(
        "APROVADAS:",
        len(resultado)
    )
    print(
        "BLOQUEADAS:",
        bloqueadas
    )
    print("==============================")

    return resultado


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

def pesquisar_completo(tema, estrutura_editorial=None):

    urls = []
    estrutura = normalizar_checkboxes_editoriais(estrutura_editorial)
    ativos = assuntos_editoriais_ativos(estrutura)

    print()
    print("==================================================")
    print("PESQUISA PADRÃO — FONTES BRASILEIRAS / PORTUGUÊS")
    print("==================================================")
    print("ASSUNTOS AUTORIZADOS:", ativos)
    print("ASSUNTOS DESAUTORIZADOS NÃO SERÃO PESQUISADOS")
    print("==================================================")

    # ========================================================
    # 01. CONSULTAS TÉCNICAS / DOCUMENTAIS — SOMENTE ATIVOS
    # ========================================================

    consultas_documentais = []
    for assunto in ativos:
        for termo in TERMOS_PESQUISA_CHECKBOX.get(assunto, []):
            consultas_documentais.append(f"{tema} {termo} pdf")
            consultas_documentais.append(f"{tema} {termo} ficha técnica")

    # Tema puro somente quando pelo menos um assunto foi autorizado.
    if ativos:
        consultas_documentais.extend([
            f"{tema} manual técnico",
            f"{tema} documentação técnica",
        ])

    # Remove duplicatas preservando ordem.
    consultas_documentais = list(dict.fromkeys(consultas_documentais))

    total_documentais = 0

    for consulta in consultas_documentais:

        resultado = pesquisar(
            consulta,
            limite=5
        )

        urls.extend(resultado)
        total_documentais += len(resultado)

    print(
        "FONTES DOCUMENTAIS BRASILEIRAS:",
        total_documentais
    )

    # ========================================================
    # 02. PESQUISA POR CHECKBOX — SEM CATEGORIAS INVENTADAS
    # ========================================================

    total_categorias = 0

    for assunto in ativos:

        termos = TERMOS_PESQUISA_CHECKBOX.get(assunto, [])

        for termo in termos:

            consulta = f"{tema} {termo}"

            print()
            print("CONSULTA AUTORIZADA:", consulta)

            resultado = pesquisar(
                consulta,
                limite=5
            )

            urls.extend(resultado)
            total_categorias += len(resultado)

    print()
    print("FONTES POR CHECKBOX AUTORIZADO:", total_categorias)

    # ========================================================
    # 03. CONSULTA GERAL BRASIL
    # ========================================================

    brasil = []
    if ativos:
        brasil = pesquisar(
            f"{tema} Brasil",
            limite=10
        )
        urls.extend(brasil)

    print(
        "FONTES BRASILEIRAS GERAIS:",
        len(brasil)
    )

    # ========================================================
    # 04. REMOVER DUPLICADOS GERAIS
    # ========================================================

    urls_unicas = []
    vistos = set()

    for url in urls:

        if not url_valida(url):
            print("URL BLOQUEADA:")
            print(url)
            continue

        try:
            dominio = (
                urlparse(str(url)).netloc
                .lower()
                .strip()
                .split(":", 1)[0]
            )
        except Exception:
            continue

        # Segunda barreira: mesmo que uma URL antiga tenha sido
        # acumulada por algum caminho legado, ela não entra.
        if not dominio.endswith(".br"):
            print("URL ESTRANGEIRA BLOQUEADA:")
            print(url)
            continue

        chave = str(url).strip()

        if chave not in vistos:
            vistos.add(chave)
            urls_unicas.append(chave)

    print()
    print("==================================================")
    print("RESULTADO DA PESQUISA BRASILEIRA")
    print("==================================================")
    print("URLs ÚNICAS:", len(urls_unicas))
    print("ORIGEM PERMITIDA: somente .br")
    print("IDIOMA/CONSULTAS: português")
    print("PESQUISA EXTERIOR: REMOVIDA")
    print("==================================================")

    return urls_unicas

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
    grupo_principal="",
    estrutura_editorial=None
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
    # TERMOS DE PESQUISA — SOMENTE CHECKBOXES ATIVOS
    # --------------------------------------------------------

    estrutura = normalizar_checkboxes_editoriais(estrutura_editorial)
    assuntos_ativos = assuntos_editoriais_ativos(estrutura)

    termos_base = []
    for assunto in assuntos_ativos:
        for termo in TERMOS_PESQUISA_CHECKBOX.get(assunto, []):
            termos_base.append(f"{termo} de {tema}")

    termos_base = list(dict.fromkeys(termos_base))

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
            assunto.replace("_", " ")
            for assunto in assuntos_ativos
        ],

        "termos_pesquisa":
            termos_base,

        "fontes_prioritarias": [

            "manual técnico",
            "datasheet",
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

        "VER TODOS OS PRODUTOS",

        "Ver todos os produtos",

        "Todos os produtos",

        "Produtos",

        "Menu",

        "Início",

        "Home",

        "Leia mais",

        "Veja também",

        "Compartilhe",

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

def coletar_pagina(
    url,
    tema=None
):

    # ========================================================
    # SEGUNDA BARREIRA — PROTEÇÃO NO PONTO DE DOWNLOAD
    # ========================================================

    if not avaliar_url_pre_download(
        url,
        tema
    ):

        print()
        print(
            "DOWNLOAD BLOQUEADO PELA PRÉ-BARREIRA:"
        )
        print(
            url
        )

        return None

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

        # SEGURANÇA: validar também a URL final após redirecionamentos.
        url_final = str(getattr(resposta, "url", "") or url).strip()
        if not avaliar_url_pre_download(url_final, tema):
            print("DOWNLOAD BLOQUEADO APÓS REDIRECIONAMENTO:", url_final)
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

                if not avaliar_url_pre_download(pdf, tema):
                    print("PDF BLOQUEADO PELA PRÉ-BARREIRA:", pdf)
                    continue

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

                # Reaplica a limpeza antes da persistência para impedir
                # que o patrimônio acumule ruído de navegação/interface.
                texto = limpar_texto_coletado(texto)

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


            # A partir da v9.16, registros explicitamente
            # descartados NÃO podem voltar ao patrimônio bruto.
            # Registros legados sem status continuam compatíveis.
            status_registro = str(
                item.get("status", "") or ""
            ).strip().casefold()

            if status_registro and status_registro != "aprovado":
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
    textos,
    estrutura_editorial=None
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
    # CONTEXTO DOS BLOCOS — FILTRADO PELOS CHECKBOXES
    # --------------------------------------------------------

    estrutura = normalizar_checkboxes_editoriais(estrutura_editorial)
    ativos = assuntos_editoriais_ativos(estrutura)
    blocos_autorizados = blocos_autorizados_pelos_checkboxes(estrutura)

    contexto_base = {}

    try:
        if isinstance(MEAD, dict):
            contexto_base = MEAD.get("contexto_blocos", {})
    except Exception:
        contexto_base = {}

    if not isinstance(contexto_base, dict):
        contexto_base = {}

    contexto_blocos = {}
    for numero_bloco in range(1, 6):
        chave = f"bloco_{numero_bloco}"
        assuntos_bloco = sorted(blocos_autorizados.get(chave, set()))
        base = contexto_base.get(chave, {})
        if not assuntos_bloco:
            contexto_blocos[chave] = {
                "ativo": False,
                "assuntos_autorizados": [],
                "objetivo": "Bloco não autorizado pelos checkboxes.",
                "funcao": "Não utilizar este bloco para criar ou exigir conteúdo.",
                "contexto": ""
            }
            continue
        # Não reaproveitar objetivo/função/contexto do MEAD legado
        # quando esses campos podem carregar assuntos desmarcados.
        # O bloco passa a ser reconstruído somente com os checkboxes ativos.
        contexto_blocos[chave] = {}
        contexto_blocos[chave]["ativo"] = True
        contexto_blocos[chave]["assuntos_autorizados"] = assuntos_bloco
        contexto_blocos[chave]["objetivo"] = (
            f"Tratar somente: {', '.join(a.replace('_', ' ') for a in assuntos_bloco)}. "
            "Não criar outros assuntos para preencher este bloco."
        )
        contexto_blocos[chave]["funcao"] = (
            "Função editorial limitada aos assuntos autorizados: "
            + ", ".join(a.replace('_', ' ') for a in assuntos_bloco)
            + "."
        )
        contexto_blocos[chave]["contexto"] = (
            "Nenhum assunto fora dos checkboxes ativos pode ser exigido, "
            "pesquisado ou usado para preencher este bloco."
        )

    print("ASSUNTOS AUTORIZADOS NO MEAD:", ativos)
    print("BLOCOS AUTORIZADOS:", {k: sorted(v) for k, v in blocos_autorizados.items()})

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

        "assuntos_autorizados":
            ativos,

        "checkboxes_editoriais":
            estrutura,

        "intencao_comercial":
            (
                "Somente o que foi autorizado pelos checkboxes "
                "pode participar da estrutura editorial. "
                "A intenção comercial não é presumida quando "
                "o checkbox comercial está desmarcado."
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
        "MODELO: qwen2.5:3b"
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

                "model": "qwen2.5:3b",

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
                    "qwen2.5:3b",

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
        "qwen2.5:3b"
    )

    # ========================================================
    # 04. ENVIAR PARA OLLAMA
    # ========================================================

    print()
    print("==============================")
    print("ENVIANDO REVISÃO PARA OLLAMA")
    print("==============================")

    inicio_ollama = time.time()

    fim_ollama = None

    try:

        resposta = requests.post(

            "http://localhost:11434/api/generate",

            json={

                "model":
                    "qwen2.5:3b",

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
    
    if fim_ollama is not None:
        print(
            "TEMPO TOTAL:",
            round(
                fim_ollama - inicio_ollama,
                1
            ),
            "segundos"
        )
    else:
        print(
            "TEMPO TOTAL: não disponível"
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
    print("INICIANDO CORREÇÃO COM QWEN3")
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
    print("CONTROLE ENVIO QWEN3")
    print("==============================")


    print(
        "MODELO:",
        "qwen3:latest"
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
    # 10. CONFIGURAÇÃO QWEN3
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
                    "qwen3:latest",

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

# ============================================================
# AUTORIDADE EDITORIAL — CHECKBOXES DO USUÁRIO
# ============================================================
# Os oito checkboxes abaixo são a única autoridade sobre quais
# assuntos podem ser pesquisados, entrar no MEAD ou ser usados
# para selecionar fragmentos.
#
# Regra: desmarcado = NÃO AUTORIZADO.
# O sistema não cria assunto adicional para preencher lacunas.
# ============================================================

ASSUNTOS_EDITORIAIS_OFICIAIS = (
    "apresentacao",
    "funcionamento",
    "aplicacoes",
    "criterios",
    "comercial",
    "informacao_tecnica",
    "instalacao_execucao",
    "beneficios",
)

MAPEAMENTO_ASSUNTO_BLOCOS = {
    "apresentacao": {"bloco_1"},
    "funcionamento": {"bloco_1", "bloco_2"},
    "aplicacoes": {"bloco_2", "bloco_5"},
    "criterios": {"bloco_3", "bloco_5"},
    "comercial": {"bloco_4", "bloco_5"},
    "informacao_tecnica": {"bloco_2", "bloco_3", "bloco_4"},
    "instalacao_execucao": {"bloco_3"},
    "beneficios": {"bloco_3", "bloco_5"},
}

MAPEAMENTO_ASSUNTO_FUNCOES = {
    "apresentacao": {"contexto"},
    "funcionamento": {"funcionamento"},
    "aplicacoes": {"aplicacao"},
    "criterios": {"tecnico", "conhecimento", "manutencao", "seguranca"},
    "comercial": {"institucional", "suporte", "solucao", "comercial"},
    "informacao_tecnica": {"tecnico", "conhecimento"},
    "instalacao_execucao": {"instalacao", "manutencao", "seguranca"},
    "beneficios": {"tecnico", "conhecimento"},
}

TERMOS_PESQUISA_CHECKBOX = {
    "apresentacao": [
        "definição", "conceito", "o que é", "contexto", "finalidade", "importância"
    ],
    "funcionamento": [
        "funcionamento", "princípio de funcionamento", "como funciona", "mecanismo", "operação"
    ],
    "aplicacoes": [
        "aplicações", "aplicação", "utilização", "uso", "onde é utilizado"
    ],
    "criterios": [
        "critérios", "seleção", "especificação", "dimensionamento", "cuidados", "manutenção"
    ],
    "comercial": [
        "fornecedor", "fornecimento", "solução", "suporte", "atendimento", "empresa"
    ],
    "informacao_tecnica": [
        "informação técnica", "características técnicas", "especificação técnica",
        "materiais", "componentes", "parâmetros técnicos"
    ],
    "instalacao_execucao": [
        "instalação", "execução", "montagem", "comissionamento", "operação", "manutenção"
    ],
    "beneficios": [
        "benefícios", "vantagens", "desempenho", "eficiência", "confiabilidade", "durabilidade"
    ],
}


def normalizar_checkboxes_editoriais(estrutura_editorial=None):
    """Retorna somente os oito checkboxes oficiais, sem criar assuntos."""
    if not isinstance(estrutura_editorial, dict):
        estrutura_editorial = globals().get("ESTRUTURA_EDITORIAL", {})

    return {
        chave: bool(estrutura_editorial.get(chave, False))
        for chave in ASSUNTOS_EDITORIAIS_OFICIAIS
    }


def assuntos_editoriais_ativos(estrutura_editorial=None):
    estrutura = normalizar_checkboxes_editoriais(estrutura_editorial)
    return [chave for chave, ativo in estrutura.items() if ativo]


def blocos_autorizados_pelos_checkboxes(estrutura_editorial=None):
    ativos = assuntos_editoriais_ativos(estrutura_editorial)
    blocos = {f"bloco_{i}": set() for i in range(1, 6)}
    for assunto in ativos:
        for bloco in MAPEAMENTO_ASSUNTO_BLOCOS.get(assunto, set()):
            blocos.setdefault(bloco, set()).add(assunto)
    return blocos


def funcoes_autorizadas_pelos_checkboxes(estrutura_editorial=None):
    ativos = assuntos_editoriais_ativos(estrutura_editorial)
    funcoes = {f"bloco_{i}": set() for i in range(1, 6)}
    blocos = blocos_autorizados_pelos_checkboxes(estrutura_editorial)
    for bloco, assuntos in blocos.items():
        for assunto in assuntos:
            funcoes[bloco].update(MAPEAMENTO_ASSUNTO_FUNCOES.get(assunto, set()))
    return funcoes


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

def barreira_fonte_brasileira_portugues(url, texto, identidade_fonte=None):
    """BARREIRA 1 — fonte brasileira + conteúdo compatível com português."""
    url = str(url or "").strip()
    texto = str(texto or "").strip()

    if not url:
        return False, "FONTE_SEM_URL"

    try:
        host = (urlparse(url).hostname or "").lower().strip(".")
    except Exception:
        host = ""

    if not host:
        return False, "DOMINIO_INVALIDO"

    # Pesquisa desta versão é brasileira: somente domínios .br.
    if not host.endswith(".br"):
        return False, "FONTE_NAO_BRASILEIRA"

    n = normalizar_assunto_texto(texto)

    # Idioma estrangeiro evidente. Espanhol exige dois marcadores para
    # evitar falso positivo; inglês basta quando há construção claramente
    # textual, não nomes técnicos isolados.
    espanhol = re.findall(
        r"\b(?:los|las|una|unos|unas|principales|incluyen|correctamente|"
        r"instalada|permanece|fugas|vibraciones|mantenimiento|desgaste|"
        r"alineacion|desalineacion|rodamientos|holgura|rendimiento|"
        r"problemas|vida util|ademas|sin embargo|para ello)\b",
        n,
        re.I
    )
    if len(espanhol) >= 2:
        return False, "IDIOMA_ESPANHOL"

    ingles = re.findall(
        r"\b(?:the|this|these|therefore|however|according|features|"
        r"maintenance|performance|installation|application)\b",
        n,
        re.I
    )
    if len(ingles) >= 3:
        return False, "IDIOMA_INGLES"

    # Artefatos de extração que tornam a fonte pouco confiável para esta
    # etapa. Não tentamos corrigir aqui; o fragmento será descartado.
    if "\ufffd" in texto or "\x00" in texto:
        return False, "ARTEFATO_EXTRACAO"

    return True, "OK"


def selecionar_informacoes_relevantes(
    tema,
    textos,
    mapa_mead,
    estrutura_editorial
):

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

        ok_fonte, motivo_fonte = barreira_fonte_brasileira_portugues(
            url, texto, identidade_fonte
        )

        if not ok_fonte:
            # Contadores da Barreira 1 são mantidos na própria função para
            # que o log diferencie coleta de aprovação.
            continue

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

    print("AUTORIDADE: somente os checkboxes acima podem gerar conteúdo.")
    if not assuntos:
        print("NENHUM CHECKBOX ATIVO — SELEÇÃO EDITORIAL BLOQUEADA.")


      
        
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
    # Não existe limite artificial de palavras para um candidato.
    # O Python trabalha com unidades editoriais naturais: parágrafos
    # e grupos de frases completos, sempre extraídos literalmente da
    # fonte original. A quantidade de palavras vira apenas um sinal
    # de qualidade/pontuação, nunca uma barreira de seleção.
    #
    # ========================================================

    candidatos = []

    print()
    print("============================================================")
    print("BARREIRA 1 — FONTES BRASILEIRAS / PORTUGUÊS")
    print("============================================================")
    print("FONTES COLETADAS:", len(textos))
    print("FONTES BRASILEIRAS ACEITAS:", len(fontes))
    print("FONTES FORA DA BARREIRA:", max(0, len(textos) - len(fontes)))
    print("REGRA DE ORIGEM: somente .br")
    print("REGRA DE IDIOMA: português predominante")
    print("============================================================")

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
            "technical specifications",
            "technical data",
            "dados tecnicos",
            "dados técnicos",
            "caracteristicas tecnicas",
            "características técnicas",
            "ficha tecnica",
            "ficha técnica",
            "technical datasheet",
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
            "nossa experiencia",
            "nossa experiência",
            "nossa equipe",
            "nossos produtos",
            "nossa empresa",
            "a empresa oferece",
            "a empresa fornece",
            "a empresa comercializa",
            "a empresa trabalha",
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

        marcadores_estrangeiros = [

            "technical specifications",
            "technical data",
            "product specifications",
            "centrifugal pump technology",
            "water pump",
            "pump model",
            "especificaciones tecnicas",
            "especificaciones técnicas",
            "detalles",
            "caracteristicas tecnicas",
            "características técnicas"
        ]

        ocorrencias_estrangeiras = 0

        for marcador in marcadores_estrangeiros:

            if (
                normalizar_assunto_texto(
                    marcador
                )
                in texto_normalizado
            ):
                ocorrencias_estrangeiras += 1

        if ocorrencias_estrangeiras >= 1:
            return False

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

            # Fragmentos somente terminam em frase completa.
            # Nunca aceitar cauda sem ponto final/interrogacao/exclamacao.
            texto_fragmento = texto_fragmento.rstrip()
            if not re.search(r"[.!?](?:[\"'”’»\)\]})]*?)\s*$", texto_fragmento):
                return

            quantidade = len(
                re.findall(
                    r"\S+",
                    texto_fragmento
                )
            )

            # A quantidade de palavras NÃO bloqueia o candidato.
            # O trecho precisa ser editorialmente válido, tematicamente
            # aderente e extraído da fonte original.
            #
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
                # REJEIÇÃO DEFINITIVA
                # ------------------------------------------------
                #
                # v8.2: o filtro editorial é uma barreira real.
                # Um trecho rejeitado aqui NÃO pode voltar para a
                # lista por uma validação mínima posterior.
                #
                # Regra: pontuação só ordena candidatos aprovados.
                # Ela nunca ressuscita um trecho rejeitado.
                # ------------------------------------------------
                return

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
        # EXTRAÇÃO EDITORIAL NATURAL — v9.16
        # ----------------------------------------------------
        #
        # O candidato deixa de ser definido por uma faixa fixa de
        # palavras. O Python procura unidades que já existem na fonte:
        #
        # 1. parágrafos reais, quando a fonte preserva separação;
        # 2. grupos de frases completas, quando não há parágrafos;
        # 3. janelas contextuais ao redor do tema, apenas como fallback.
        #
        # Em todos os casos o texto é recortado literalmente da fonte
        # original. Não há resumo, concatenação artificial ou limite
        # mínimo/máximo de palavras.
        # ----------------------------------------------------

        intervalos_frases = []

        for frase in frases_analise:

            palavras_da_frase = [
                indice
                for indice, palavra in enumerate(mapa_palavras)
                if (
                    palavra["analise_inicio"] >= frase["inicio"]
                    and
                    palavra["analise_fim"] <= frase["fim"]
                )
            ]

            if not palavras_da_frase:
                continue

            intervalos_frases.append({
                "inicio": palavras_da_frase[0],
                "fim": palavras_da_frase[-1] + 1
            })

        # 1) Parágrafos reais da fonte original.
        intervalos_paragrafos = []
        for correspondencia in re.finditer(
            r"(?s)(.*?)(?:\n\s*\n|\Z)",
            texto_original
        ):
            trecho = correspondencia.group(0).strip()
            if not trecho:
                continue

            inicio_char = correspondencia.start()
            fim_char = correspondencia.end()

            indices = [
                indice
                for indice, palavra in enumerate(palavras_original)
                if palavra.end() > inicio_char and palavra.start() < fim_char
            ]

            if indices:
                intervalo = (indices[0], indices[-1] + 1)
                if intervalo not in intervalos_paragrafos:
                    intervalos_paragrafos.append(intervalo)

        # Se a fonte veio sem quebras de parágrafo, usamos grupos de
        # frases completas. Isso evita produzir uma fonte inteira como
        # um único candidato.
        if intervalos_paragrafos:
            for inicio, fim in intervalos_paragrafos:
                adicionar_candidato_por_palavras(inicio, fim)

                # Parágrafos excessivamente longos ganham trechos naturais
                # por frases, sem cortar palavras nem impor quantidade de
                # palavras.
                frases_do_paragrafo = [
                    (f["inicio"], f["fim"])
                    for f in intervalos_frases
                    if f["inicio"] >= inicio and f["fim"] <= fim
                ]

                if len(frases_do_paragrafo) > 4:
                    for pos in range(0, len(frases_do_paragrafo), 3):
                        grupo = frases_do_paragrafo[pos:pos + 3]
                        if grupo:
                            adicionar_candidato_por_palavras(
                                grupo[0][0],
                                grupo[-1][1]
                            )
        else:
            # 2) Sem parágrafos: frases completas em grupos de até 3.
            if intervalos_frases:
                for pos in range(0, len(intervalos_frases), 2):
                    grupo = intervalos_frases[pos:pos + 3]
                    adicionar_candidato_por_palavras(
                        grupo[0]["inicio"],
                        grupo[-1]["fim"]
                    )

                # Também preserva frases individuais, permitindo que uma
                # frase técnica muito boa concorra sem ser descartada por
                # tamanho.
                for intervalo in intervalos_frases:
                    adicionar_candidato_por_palavras(
                        intervalo["inicio"],
                        intervalo["fim"]
                    )

        # 3) Fallback contextual: quando a extração natural produziu poucos
        # candidatos, criar janelas de frases ao redor das ocorrências do
        # tema. Continua sendo um recorte literal da fonte original.
        if len(candidatos) < 3 and intervalos_frases:
            termos_contextuais = [
                termo for termo in re.findall(
                    r"\b[a-z0-9]{3,}\b",
                    normalizar_assunto_texto(tema)
                )
                if termo not in {
                    "de", "da", "das", "do", "dos", "em", "na", "nas",
                    "no", "nos", "para", "por", "com", "sem", "e", "a",
                    "o", "as", "os"
                }
            ]

            ocorrencias = []
            for indice, intervalo in enumerate(intervalos_frases):
                texto_frase = texto_para_analise[
                    0:0
                ]
                inicio = intervalo["inicio"]
                fim = intervalo["fim"]
                texto_frase = " ".join(
                    mapa_palavras[i]["palavra_analise"]
                    for i in range(inicio, fim)
                )
                normalizado = normalizar_assunto_texto(texto_frase)
                if any(termo in normalizado for termo in termos_contextuais):
                    ocorrencias.append(indice)

            for indice in ocorrencias[:8]:
                inicio_frase = max(0, indice - 1)
                fim_frase = min(
                    len(intervalos_frases),
                    indice + 2
                )
                adicionar_candidato_por_palavras(
                    intervalos_frases[inicio_frase]["inicio"],
                    intervalos_frases[fim_frase - 1]["fim"]
                )

    # ========================================================
    # 04.5. RESERVATÓRIO AMPLIADO DE CANDIDATOS — v31
    # ========================================================
    # O reservatório não pode ser apenas grande; ele precisa conter
    # recortes com capacidade suficiente para formar os blocos.
    #
    # v30 mostrou que 150 candidatos ainda podiam deixar B1 com apenas
    # 58/55/54 palavras nos melhores itens. O problema era o tamanho dos
    # recortes, não a quantidade bruta.
    #
    # v31 mantém os mesmos gates e acrescenta somente ALTERNATIVAS
    # LITERAIS de 3 frases, além das janelas de 2 frases. Nenhum texto é
    # reescrito, completado ou inventado.
    #
    # A coleta é feita em rodízio entre as fontes para que uma única fonte
    # não consuma todo o reservatório antes que as demais contribuam.
    # ========================================================

    RESERVATORIO_BRUTO_ALVO = 150

    if len(candidatos) < RESERVATORIO_BRUTO_ALVO:

        candidatos_antes_reserva = len(candidatos)
        candidatos_reserva_adicionados = 0

        fontes_reserva = []
        for fonte in fontes:
            texto_original = str(fonte.get("texto", "") or "").strip()
            if not texto_original:
                continue

            frases_fonte = [
                m.group(0).strip()
                for m in re.finditer(
                    r"[^.!?]+[.!?]+(?:\s|$)",
                    texto_original,
                    flags=re.S
                )
                if m.group(0).strip()
            ]
            if len(frases_fonte) < 2:
                continue

            identidade_fonte = fonte.get("identidade_fonte", {})
            if not isinstance(identidade_fonte, dict):
                identidade_fonte = {}

            fontes_reserva.append({
                "fonte": fonte,
                "frases": frases_fonte,
                "identidade": identidade_fonte,
                "pos": 0
            })

        # Primeiro usamos janelas de 3 frases, pois elas aumentam a
        # capacidade factual sem alterar uma única palavra da fonte.
        # Depois usamos 2 frases para preencher lacunas.
        for tamanho_janela in (3, 2):
            if len(candidatos) >= RESERVATORIO_BRUTO_ALVO:
                break

            progresso = True
            while progresso and len(candidatos) < RESERVATORIO_BRUTO_ALVO:
                progresso = False

                for estado in fontes_reserva:
                    if len(candidatos) >= RESERVATORIO_BRUTO_ALVO:
                        break

                    frases_fonte = estado["frases"]
                    pos = estado["pos"]
                    limite = len(frases_fonte) - tamanho_janela + 1
                    if pos >= limite:
                        continue

                    grupo = frases_fonte[pos:pos + tamanho_janela]
                    estado["pos"] = pos + 1
                    progresso = True

                    trecho_reserva = " ".join(
                        str(frase).strip()
                        for frase in grupo
                        if str(frase).strip()
                    ).strip()

                    if not trecho_reserva:
                        continue

                    if not fragmento_pertence_ao_tema(trecho_reserva, tema):
                        continue

                    if not fragmento_eh_editorialmente_valido(
                        trecho_reserva,
                        estado["identidade"]
                    ):
                        continue

                    candidato_reserva = {
                        "texto": trecho_reserva,
                        "fonte": estado["fonte"]["indice"],
                        "url": estado["fonte"]["url"],
                        "tipo": estado["fonte"]["tipo"],
                        "pdf": estado["fonte"]["eh_pdf"],
                        "palavras": len(trecho_reserva.split()),
                        "identidade_fonte": estado["identidade"],
                        "_reservatorio_janela": tamanho_janela,
                    }

                    candidatos.append(candidato_reserva)
                    candidatos_reserva_adicionados += 1

        print()
        print("==============================")
        print("RESERVATÓRIO BRUTO AMPLIADO — v31")
        print("==============================")
        print("ALVO DO RESERVATÓRIO:", RESERVATORIO_BRUTO_ALVO)
        print("ANTES:", candidatos_antes_reserva)
        print("DEPOIS:", len(candidatos))
        print("ADICIONADOS:", candidatos_reserva_adicionados)
        print("JANELAS LITERAIS: 3 frases + 2 frases")
        print("REGRA: mesmos filtros editoriais; nenhuma reescrita")
        print("ESTRATÉGIA v31: capacidade factual + rodízio de fontes")
        print("SELEÇÃO FINAL: permanece em 15 fragmentos")

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
        "CANDIDATOS APROVEITADOS / RESERVA LIMPA:",
        len(candidatos)
    )

    # Diagnóstico compacto da distribuição antes da seleção.
    candidatos_por_fonte_diagnostico = {}
    for candidato in candidatos:
        fonte_id = candidato.get("fonte", "")
        candidatos_por_fonte_diagnostico[fonte_id] = (
            candidatos_por_fonte_diagnostico.get(fonte_id, 0) + 1
        )

    print("CANDIDATOS POR FONTE:")
    for fonte_id, quantidade in sorted(
        candidatos_por_fonte_diagnostico.items(),
        key=lambda item: item[0]
    ):
        print("  FONTE", fonte_id, ":", quantidade)

    print()
    print("DIAGNÓSTICO DETALHADO POR FONTE V7.4:")
    for fonte_id in sorted(diagnostico_fontes, key=lambda x: int(x) if str(x).isdigit() else str(x)):
        d = diagnostico_fontes[fonte_id]
        print(
            "  FONTE", fonte_id,
            "| SEGMENTOS:", d["segmentos"],
            "| AVALIADOS:", d["avaliados"],
            "| TEMA OK:", d["tema_aprovado"],
            "| TEMA REJ:", d["tema_rejeitado"],
            "| EDITORIAL REJ:", d["editorial_rejeitado"],
            "| SALVOS:", d["candidatos_salvos"]
        )
    
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
    # AUTORIDADE ÚNICA: CHECKBOXES DO USUÁRIO.
    # O MEAD global/legado não pode reintroduzir assuntos.
    # --------------------------------------------------------

    contexto_blocos_mead = {}
    blocos_check_editorial = blocos_autorizados_pelos_checkboxes(
        estrutura_editorial
    )
    funcoes_check_editorial = funcoes_autorizadas_pelos_checkboxes(
        estrutura_editorial
    )

    descricoes_assunto = {
        "apresentacao": "apresentação e contextualização",
        "funcionamento": "funcionamento",
        "aplicacoes": "aplicações",
        "criterios": "critérios",
        "comercial": "dimensão comercial autorizada",
        "informacao_tecnica": "informação técnica",
        "instalacao_execucao": "instalação e execução",
        "beneficios": "benefícios",
    }

    for numero_bloco in range(1, 6):
        chave = f"bloco_{numero_bloco}"
        assuntos_bloco = sorted(blocos_check_editorial.get(chave, set()))
        funcoes_bloco = sorted(funcoes_check_editorial.get(chave, set()))

        if not assuntos_bloco:
            contexto_blocos_mead[chave] = {
                "ativo": False,
                "assuntos_autorizados": [],
                "funcoes_autorizadas": [],
                "objetivo": "Bloco não autorizado pelos checkboxes.",
                "funcao": "Não utilizar este bloco para criar ou exigir conteúdo.",
                "contexto": "",
            }
            continue

        nomes = [descricoes_assunto.get(a, a.replace('_', ' ')) for a in assuntos_bloco]
        contexto_blocos_mead[chave] = {
            "ativo": True,
            "assuntos_autorizados": assuntos_bloco,
            "funcoes_autorizadas": funcoes_bloco,
            "objetivo": (
                "Tratar exclusivamente " + ", ".join(nomes) + ". "
                "Nenhum outro assunto pode ser criado para completar este bloco."
            ),
            "funcao": (
                "Função editorial limitada a: " + ", ".join(funcoes_bloco) + "."
            ),
            "contexto": (
                "Apenas evidências compatíveis com os checkboxes ativos podem "
                "ser usadas na seleção deste bloco."
            ),
        }

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

    # ========================================================
    # AUTORIDADE DOS CHECKBOXES SOBRE OS TERMOS DE BLOCO
    # ========================================================
    # O dicionário legado acima continua existindo por compatibilidade,
    # mas não pode autorizar assuntos desmarcados. Reconstituímos os
    # termos de cada bloco exclusivamente a partir dos checkboxes ativos.
    # ========================================================

    blocos_check = blocos_autorizados_pelos_checkboxes(estrutura_editorial)
    termos_blocos_autorizados = {f"bloco_{i}": [] for i in range(1, 6)}

    for chave_bloco, assuntos_bloco in blocos_check.items():
        termos = []
        for assunto in sorted(assuntos_bloco):
            termos.extend(TERMOS_PESQUISA_CHECKBOX.get(assunto, []))
        # IMPORTANTE: o texto descritivo do MEAD NÃO entra no vocabulário
        # de evidência. O MEAD define autoridade editorial; não deve
        # transformar palavras de instrução ("tratar", "exclusivamente",
        # "assunto", "pode", etc.) em sinais de relevância factual.
        termos_blocos_autorizados[chave_bloco] = list(dict.fromkeys(
            normalizar_assunto_texto(t) for t in termos if str(t).strip()
        ))

    termos_blocos = termos_blocos_autorizados

    print("TERMOS DE EVIDÊNCIA AUTORIZADOS PELOS CHECKBOXES:")
    for _bloco, _termos in termos_blocos.items():
        print(_bloco, "=>", _termos[:30])

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
        # Os fragmentos oficiais da v8.2 trabalham entre 70 e 100
        # palavras. A faixa central recebe o maior reforço.
        # ----------------------------------------------------

        if 70 <= quantidade_palavras <= 100:

            qualidade += 3

        elif 60 <= quantidade_palavras <= 110:

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
    # 06.3A — FUNÇÃO EDITORIAL REAL DO FRAGMENTO
    # ========================================================
    # A pontuação por palavras não é suficiente para decidir em
    # qual bloco um fragmento pode entrar. Um trecho pode conter
    # "solução", "aplicação" ou "característica" e ainda assim
    # pertencer claramente a funcionamento, manutenção ou contexto.
    #
    # Esta classificação é uma BARREIRA DE ELEGIBILIDADE:
    # pontuação não consegue ressuscitar um papel editorial errado.
    # ========================================================

    def classificar_funcao_editorial_fragmento(texto):
        texto_n = normalizar_assunto_texto(texto or "")
        if not texto_n:
            return {
                "principal": "indefinido",
                "funcoes": [],
                "pontuacoes": {}
            }

        grupos_funcao = {
            "contexto": [
                "contexto", "cenário", "cenario", "necessidade",
                "demanda", "abastecimento", "infraestrutura",
                "consumo", "fornecimento de agua", "saneamento",
                "processo industrial", "demanda industrial"
            ],
            "funcionamento": [
                "funcionamento", "funciona", "principio de funcionamento",
                "pressao", "pressão", "vazao", "vazão", "impulsor",
                "rotacao", "rotação", "centrifuga", "centrífuga",
                "eixo", "motor", "sucção", "succao", "descarga"
            ],
            "aplicacao": [
                "aplicacao", "aplicação", "utilizacao", "utilização",
                "utilizado", "utilizada", "empregada", "empregado",
                "abastecimento", "irrigacao", "irrigação",
                "drenagem", "processo industrial", "transferencia",
                "transferência", "transporte de liquido", "transporte de líquido"
            ],
            "tecnico": [
                "caracteristica", "características", "caracteristicas",
                "dimensionamento", "especificacao", "especificação",
                "material", "materiais", "temperatura", "pressao",
                "pressão", "vazao", "vazão", "rendimento", "eficiencia",
                "eficiência", "desempenho", "componente", "componentes",
                "corrosao", "corrosão", "durabilidade", "confiabilidade"
            ],
            "manutencao": [
                "manutencao", "manutenção", "manutencoes", "manutenções",
                "inspecao", "inspeção", "desgaste", "falha", "falhas",
                "lubrificacao", "lubrificação", "reparo", "reparos",
                "troca de", "limpeza", "problema", "problemas"
            ],
            "instalacao": [
                "instalacao", "instalação", "montagem", "montar",
                "instalado", "instalada", "alinhamento", "fixacao",
                "fixação", "tubulacao", "tubulação", "conexao",
                "conexão", "comissionamento"
            ],
            "seguranca": [
                "seguranca", "segurança", "risco", "riscos", "protecao",
                "proteção", "operacao segura", "operação segura",
                "procedimento de seguranca", "procedimento de segurança"
            ],
            "institucional": [
                "empresa", "fabricante", "fabricacao", "fabricação",
                "fabricamos", "experiencia", "experiência",
                "conhecimento", "especializada", "especializado",
                "atuacao", "atuação", "equipe", "engenharia"
            ],
            "suporte": [
                "suporte", "atendimento", "assistencia", "assistência",
                "pos venda", "pós venda", "orientacao", "orientação",
                "acompanhamento", "consultoria", "manutencao especializada",
                "manutenção especializada"
            ],
            "solucao": [
                "solucao", "solução", "solucoes", "soluções",
                "fornecimento", "fornecer", "fornece", "sistema completo",
                "conjunto", "configuracao", "configuração", "adequacao",
                "adequação"
            ],
            # B4 também pode ser sustentado por CONHECIMENTO TÉCNICO.
            # Isso representa a dimensão de conhecimento do MEAD sem
            # transformar um trecho técnico em alegação institucional.
            # É deliberadamente separado de "institucional" e "comercial".
            "conhecimento": [
                "conhecimento técnico", "conhecimento tecnico",
                "especificacao", "especificação",
                "dimensionamento", "dimensionar",
                "selecao", "seleção",
                "caracteristica", "características", "caracteristicas",
                "principio de funcionamento", "princípio de funcionamento",
                "componentes", "componente",
                "condicoes de operacao", "condições de operação",
                "desempenho", "rendimento",
                "criterio de selecao", "critério de seleção",
                "criterios de selecao", "critérios de seleção",
                "cuidados na operacao", "cuidados na operação"
            ],
            "comercial": [
                "preco", "preço", "promocao", "promoção", "comprar",
                "venda", "oferta", "orcamento", "orçamento", "desconto",
                "parcelamento", "entrega", "frete"
            ]
        }

        pontuacoes = {}
        for funcao, termos in grupos_funcao.items():
            pontos = 0
            for termo in termos:
                if termo in texto_n:
                    pontos += 1
            pontuacoes[funcao] = pontos

        # --------------------------------------------------------
        # SOLUÇÃO TÉCNICA — importante para páginas de PRODUTO
        #
        # Nem todo conteúdo de solução contém literalmente a palavra
        # "solução". Em material técnico, a dimensão de solução aparece
        # quando o trecho explica adequação, seleção, dimensionamento,
        # indicação de uso ou atendimento a uma necessidade.
        #
        # Isso NÃO transforma um trecho técnico qualquer em comercial:
        # exigimos pelo menos dois sinais de adequação/seleção, ou uma
        # expressão factual forte de indicação/uso.
        # --------------------------------------------------------
        sinais_adequacao = [
            "adequado para", "adequada para", "adequados para",
            "adequadas para", "indicado para", "indicada para",
            "indicados para", "indicadas para", "utilizado para",
            "utilizada para", "utilizados para", "utilizadas para",
            "pode ser utilizado", "pode ser utilizada",
            "podem ser utilizados", "podem ser utilizadas",
            "atende a", "atendem a", "atendimento de",
            "dimensionamento", "dimensionar", "seleção", "selecao",
            "escolha", "vazão necessária", "vazao necessaria",
            "pressão necessária", "pressao necessaria",
            "necessidade de bombeamento", "necessidades de bombeamento",
            "condições de operação", "condicoes de operacao",
            "aplicação específica", "aplicacao especifica"
        ]
        sinais_presentes = sum(1 for termo in sinais_adequacao if termo in texto_n)

        # "solução" explícita é válida quando o trecho continua sendo
        # técnico e não contém propaganda/comercialização.
        solucao_explicita = any(
            termo in texto_n for termo in [
                "solução técnica", "solucao tecnica",
                "solução de bombeamento", "solucao de bombeamento",
                "solução para", "solucao para",
                "soluções para", "solucoes para"
            ]
        )

        if sinais_presentes >= 2 or solucao_explicita:
            pontuacoes["solucao"] = max(
                pontuacoes.get("solucao", 0),
                2
            )

        # Uma aplicação factual pode representar uma solução para a
        # página de produto quando o trecho relaciona produto + uso +
        # adequação. Não basta mencionar somente "aplicação".
        if (
            pontuacoes.get("aplicacao", 0) >= 2
            and sinais_presentes >= 1
        ):
            pontuacoes["solucao"] = max(
                pontuacoes.get("solucao", 0),
                2
            )

        # --------------------------------------------------------
        # CONHECIMENTO TÉCNICO PARA O BLOCO 4
        # --------------------------------------------------------
        # O bloco 4 do MEAD não exige que todo trecho seja institucional.
        # Ele pode introduzir a dimensão de conhecimento, desde que o
        # conhecimento esteja efetivamente sustentado pelo fragmento.
        # Aqui usamos dois níveis:
        #   1) sinais explícitos de conhecimento/especificação; ou
        #   2) pelo menos três sinais técnicos distintos.
        #
        # Isso NÃO cria autoridade empresarial. Apenas permite que um
        # conteúdo técnico factual seja usado como base da dimensão
        # "conhecimento" do bloco 4.
        # --------------------------------------------------------
        sinais_conhecimento_explicito = [
            "especificacao", "especificação",
            "dimensionamento", "dimensionar",
            "selecao", "seleção",
            "criterio de selecao", "critério de seleção",
            "criterios de selecao", "critérios de seleção",
            "conhecimento técnico", "conhecimento tecnico",
            "principio de funcionamento", "princípio de funcionamento",
            "condicoes de operacao", "condições de operação"
        ]
        conhecimento_explicito = sum(
            1 for termo in sinais_conhecimento_explicito
            if termo in texto_n
        )

        sinais_tecnicos_distintos = [
            "pressao", "pressão", "vazao", "vazão",
            "impulsor", "eixo", "motor", "rotacao", "rotação",
            "componente", "componentes", "material", "materiais",
            "temperatura", "rendimento", "eficiencia", "eficiência",
            "desempenho", "dimensionamento", "especificacao",
            "especificação", "selecao", "seleção", "operacao",
            "operação", "manutencao", "manutenção", "instalacao",
            "instalação", "corrosao", "corrosão", "confiabilidade"
        ]
        tecnicos_distintos = sum(
            1 for termo in sinais_tecnicos_distintos
            if termo in texto_n
        )

        if conhecimento_explicito >= 1 or tecnicos_distintos >= 3:
            pontuacoes["conhecimento"] = max(
                pontuacoes.get("conhecimento", 0),
                2
            )

        # "comercial" é uma função editorial válida quando o
        # checkbox comercial está ativo. A limpeza comercial continua
        # sendo um gate separado; aqui apenas classificamos a função.
        pontuacoes_validas = dict(pontuacoes)

        ordenadas = sorted(
            pontuacoes_validas.items(),
            key=lambda x: x[1],
            reverse=True
        )

        principal = ordenadas[0][0] if ordenadas and ordenadas[0][1] > 0 else "indefinido"

        # Funções com evidência suficiente. Usamos 2 sinais para
        # evitar que uma palavra isolada classifique o trecho.
        funcoes = [k for k, v in ordenadas if v >= 2]

        return {
            "principal": principal,
            "funcoes": funcoes,
            "pontuacoes": pontuacoes,
            "comercial": pontuacoes.get("comercial", 0)
        }

    # ========================================================
    # AUTORIDADE REAL DOS CHECKBOXES SOBRE AS FUNÇÕES
    # ========================================================
    # Não existe mais uma segunda lista legada capaz de excluir uma
    # função que o usuário marcou. O mapeamento oficial é a única
    # autorização. Isso é especialmente importante para "comercial".
    funcoes_permitidas_bloco = funcoes_autorizadas_pelos_checkboxes(
        estrutura_editorial
    )

    print("FUNÇÕES DE BLOCO AUTORIZADAS PELOS CHECKBOXES:")
    for _bloco, _funcoes in funcoes_permitidas_bloco.items():
        print(_bloco, "=>", sorted(_funcoes))

    def fragmento_compativel_com_funcao_do_bloco(candidato, chave_bloco):
        if not blocos_autorizados_pelos_checkboxes(estrutura_editorial).get(chave_bloco):
            return False

        classificacao = candidato.get("_funcao_editorial", {})
        funcoes = set(classificacao.get("funcoes", []))

        if not funcoes:
            return False

        permitidas = funcoes_permitidas_bloco.get(chave_bloco, set())
        if not (funcoes & permitidas):
            return False

        # No bloco 1, funcionamento/conhecimento só entram como apoio
        # contextual; evita transformar a abertura em um bloco técnico.
        if chave_bloco == "bloco_1" and not (funcoes & {"contexto", "aplicacao"}):
            texto_bloco1 = normalizar_texto(candidato.get("texto", ""))
            sinais_contextuais = (
                "necessidade", "demanda", "abastecimento", "saneamento",
                "infraestrutura", "processo", "sistema", "uso", "utilizacao",
                "aplicacao", "cenário", "cenario", "contexto", "vazao", "pressao"
            )
            if not any(s in texto_bloco1 for s in sinais_contextuais):
                return False

        # Quando o checkbox comercial está ativo, "comercial" pode ser
        # uma função editorial legítima. Isso NÃO libera propaganda, CTA,
        # catálogo ou texto promocional: esses resíduos continuam barrados
        # pelos gates de limpeza comercial e auditoria final.
        if (
            "comercial" in permitidas
            and classificacao.get("comercial", 0) >= 2
        ):
            return True

        # Quando comercial não foi autorizado, um trecho predominantemente
        # comercial não pode entrar apenas por parecer técnico.
        if classificacao.get("comercial", 0) >= 2:
            return bool(funcoes & permitidas & {
                "institucional", "suporte", "solucao"
            })

        return True

    # ========================================================
    # 06.3B — CLASSIFICAÇÃO ANTES DA ALOCAÇÃO
    # ========================================================

    contagem_funcoes = {
        funcao: 0
        for funcao in [
            "contexto", "funcionamento", "aplicacao", "tecnico",
            "manutencao", "instalacao", "seguranca", "institucional",
            "suporte", "solucao", "conhecimento", "comercial", "indefinido"
        ]
    }

    for candidato in candidatos:
        classificacao = classificar_funcao_editorial_fragmento(
            candidato.get("texto", "")
        )
        candidato["_funcao_editorial"] = classificacao
        principal = classificacao.get("principal", "indefinido")
        contagem_funcoes[principal] = contagem_funcoes.get(principal, 0) + 1

    print()
    print("==============================")
    print("MAPA DE FUNÇÃO EDITORIAL DOS CANDIDATOS")
    print("==============================")
    for funcao, quantidade in contagem_funcoes.items():
        print(f"{funcao.upper()}: {quantidade}")

    print()
    print("COMPATIBILIDADE FUNCIONAL POR BLOCO — ANTES DA ALOCAÇÃO")
    print("==============================")
    for numero_bloco in range(1, 6):
        chave = f"bloco_{numero_bloco}"
        total_compativel = sum(
            1
            for candidato in candidatos
            if fragmento_compativel_com_funcao_do_bloco(candidato, chave)
        )
        funcoes = ", ".join(sorted(funcoes_permitidas_bloco[chave]))
        detalhes = {}
        for candidato in candidatos:
            if fragmento_compativel_com_funcao_do_bloco(candidato, chave):
                for funcao in candidato.get("_funcao_editorial", {}).get("funcoes", []):
                    if funcao in funcoes_permitidas_bloco[chave]:
                        detalhes[funcao] = detalhes.get(funcao, 0) + 1
        detalhes_txt = ", ".join(
            f"{k}={v}" for k, v in sorted(detalhes.items())
        ) or "nenhuma"
        print(
            f"{chave.upper()}: {total_compativel} candidatos compatíveis | "
            f"funções: {funcoes} | distribuição: {detalhes_txt}"
        )

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
    
            if not fragmento_compativel_com_funcao_do_bloco(
                candidato,
                chave_bloco
            ):
                continue

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

    # Reservas em memória: 5 fragmentos extras por bloco.
    # Não entram no JSON enquanto não forem utilizadas.
    # A lista funciona como fila: reserva 1 -> reserva 2.
    fragmentos_reserva_por_bloco = {
        f"bloco_{numero_bloco}": []
        for numero_bloco in range(1, 6)
    }

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

    def fragmento_tem_repeticao_factual_forte(candidato, limite=0.55):
        """Barreira dura contra trechos que repetem o mesmo conteúdo já selecionado."""
        texto_candidato = str(candidato.get("texto", "") or "").strip()
        if not texto_candidato:
            return False
        for selecionado in fragmentos_selecionados:
            semelhanca = calcular_semelhanca_fragmentos(
                texto_candidato,
                selecionado.get("texto", "")
            )
            if semelhanca >= limite:
                return True
        return False

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

        # ----------------------------------------------------
        # GATE EDITORIAL ABSOLUTO — ANTES DA PONTUAÇÃO
        # ----------------------------------------------------
        # Um trecho contaminado não pode ser ressuscitado por nota.
        ok_editorial, motivo_editorial = diagnosticar_contaminacao_editorial(
            texto,
            candidato.get("identidade_fonte", {})
        )
        if not ok_editorial:
            candidato["motivo_rejeicao_editorial"] = motivo_editorial
            return False

        # Repetição forte de conteúdo já selecionado é rejeição, não apenas
        # penalidade. Assim um mesmo parágrafo-base não ocupa dois blocos.
        try:
            if fragmento_tem_repeticao_factual_forte(candidato, limite=0.55):
                candidato["motivo_rejeicao_editorial"] = "REPETICAO_FACTUAL_FORTE"
                return False
        except Exception:
            pass
            
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
        # COMPATIBILIDADE BÁSICA COM O BLOCO MEAD
        # ----------------------------------------------------
        # Rejeita apenas incompatibilidades fortes; não exige que cada
        # fragmento repita o título do bloco.
        bloco_candidato = str(candidato.get("bloco_mead", "") or "").strip()
        n_candidato = normalizar_assunto_texto(texto)
        sinais_bloco = {
            "bloco_1": (
                ("contexto", "importancia", "necessidade", "cenari", "abastecimento", "demanda"),
                ("codigo", "modelo", "peca", "manutencao preventiva")
            ),
            "bloco_2": (
                ("funcionamento", "aplicacao", "vazao", "pressao", "impulsor", "configuracao", "utilizacao"),
                ("empresa oferece", "nossa empresa", "orcamento")
            ),
            "bloco_3": (
                ("manutencao", "instalacao", "operacao", "seguranca", "dimensionamento", "impulsor", "difusor", "corrosao"),
                ()
            ),
            "bloco_4": (
                (
                    "empresa", "suporte", "atendimento", "servico", "solucao",
                    "conhecimento", "fornecedor", "fornecimento", "produto",
                    "produtos", "sistema", "atendimento tecnico", "atendimento técnico",
                    "orientacao", "orientação"
                ),
                ("promocao", "compre", "preco", "oferta", "desconto", "parcelamento")
            ),
            "bloco_5": (
                ("aplicacao", "selecao", "solucao", "necessidade", "operacao", "manutencao", "confiabilidade"),
                ()
            ),
        }
        if bloco_candidato in sinais_bloco:
            sinais_positivos, sinais_proibidos = sinais_bloco[bloco_candidato]
            if sinais_proibidos and any(x in n_candidato for x in sinais_proibidos):
                return False
            # A compatibilidade editorial já foi validada antes desta função
            # por fragmento_compativel_com_funcao_do_bloco().
            #
            # Não exigir a presença literal de uma palavra-chave no texto.
            # Um trecho pode exercer corretamente a função MEAD de
            # funcionamento, conhecimento técnico, benefício ou critério sem
            # repetir exatamente o rótulo usado no mapa. Essa exigência estava
            # causando falsos negativos, sobretudo no bloco_3, onde o MEAD
            # autoriza benefícios, critérios, informação técnica, instalação e
            # execução, enquanto os sinais antigos cobriam apenas parte disso.
            #
            # Portanto: funções MEAD incompatíveis continuam bloqueadas no
            # gate anterior; aqui mantemos somente os sinais explicitamente
            # proibidos de cada bloco.

        # ----------------------------------------------------
        # IDENTIDADE COMERCIAL / PRODUTO
        # ----------------------------------------------------

        if not fragmento_eh_comercialmente_limpo(
            texto,
            candidato.get(
                "identidade_fonte",
                {}
            )
        ):
            return False    

        # ----------------------------------------------------
        # QUANTIDADE DE PALAVRAS
        # ----------------------------------------------------
        # Nenhum piso ou teto artificial. A integridade do fragmento é
        # decidida pelos gates estruturais, editoriais e factuais.

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
        chave_bloco,
        hashes_bloqueados=None
    ):
        hashes_bloqueados = hashes_bloqueados or set()

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

            # HARD GATE EDITORIAL FINAL:
            # nunca permitir que uma pontuação reintroduza lixo.
            ok_editorial, motivo_editorial = diagnosticar_contaminacao_editorial(
                candidato.get("texto", ""),
                candidato.get("identidade_fonte", {})
            )
            if not ok_editorial:
                candidato["motivo_rejeicao_editorial"] = motivo_editorial
                continue

            # HARD GATE FINAL: a pontuação jamais pode ressuscitar
            # um fragmento contaminado. Revalida a identidade da fonte
            # no instante imediatamente anterior à seleção.
            if not fragmento_eh_comercialmente_limpo(
                candidato.get("texto", ""),
                candidato.get("identidade_fonte", {})
            ):
                continue

            # ÚLTIMO GATE: o trecho precisa chegar ao Ollama já como
            # parágrafo técnico fechado, sem corte sintático ou artefato.
            ok_final, motivo_final = auditar_fragmento_final_python(
                candidato.get("texto", ""),
                candidato.get("identidade_fonte", {})
            )
            if not ok_final:
                candidato["motivo_rejeicao_python_final"] = motivo_final
                continue

            # ====================================================
            # PORTÃO DE QUALIDADE v9.18
            # ====================================================
            # Este é o último bloqueio antes de o candidato poder
            # ser escolhido e, posteriormente, entrar em
            # informacoes_relevantes. Qualidade baixa não pode ser
            # compensada por relevância temática ou pontuação.
            # ====================================================
            # Auditoria final explícita: o fragmento que chegou aqui é
            # exatamente o texto que será enviado ao Ollama.
            print(
                'AUDITORIA FINAL FRAGMENTO:',
                candidato.get('id', 'SEM_ID'),
                '| PALAVRAS:',
                candidato.get('palavras', 0),
                '| INTEGRIDADE:',
                'OK'
            )

            ok_qualidade, nota_qualidade, motivo_qualidade = portao_qualidade_fragmento(
                candidato.get("texto", ""),
                candidato.get("identidade_fonte", {}),
                tema
            )
            candidato["nota_portao_qualidade"] = nota_qualidade
            candidato["motivo_portao_qualidade"] = motivo_qualidade
            if not ok_qualidade:
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

            if hash_trecho in hashes_bloqueados:

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
    # 07. ALOCAÇÃO GLOBAL 5 x 3 — v9.18-v14
    # ========================================================
    # A seleção não fecha um bloco apenas por pontuação.
    # Cada bloco precisa sair desta etapa com 3 fragmentos válidos.
    # Não há mínimo nem máximo de palavras: qualidade, integridade,
    # diversidade e aderência temática são os critérios.
    # ========================================================

    fragmentos_selecionados = []
    hashes_selecionados = set()
    fontes_utilizadas = {}

    def _palavras_fragmento(candidato):
        return len(str(candidato.get('texto', '') or '').split())

    # ========================================================
    # NOTA INDIVIDUAL DO TRECHO — v46
    # ========================================================
    # O fragmento só deve chegar ao Ollama quando for excelente
    # individualmente. "Passou no portão" deixa de ser suficiente:
    # o trecho precisa atingir nota máxima 100/100.
    # ========================================================
    def _nota_maxima_trecho(candidato, chave_bloco):
        texto = str(candidato.get("texto", "") or "").strip()
        if not texto:
            return 0, ["VAZIO"]

        # O portão de qualidade e a nota editorial são duas dimensões
        # diferentes. O portão já exige qualidade estrutural mínima; exigir
        # 100/100 novamente aqui zerava candidatos com portão 92, 94 etc.
        # Para manter rigor sem duplicar a barreira, o mínimo estrutural
        # deste segundo estágio é 90.
        nota_portao = int(candidato.get("nota_portao_qualidade", 0) or 0)
        NOTA_MINIMA_PORTAO = 85
        if nota_portao < NOTA_MINIMA_PORTAO:
            return 0, [f"PORTAO_ABAIXO_DO_MINIMO:{nota_portao}"]

        motivos = []
        pontos = 0

        # 20 — aderência funcional explícita ao bloco.
        classificacao = candidato.get("_funcao_editorial", {}) or {}
        funcoes = set(classificacao.get("funcoes", []))
        if funcoes & funcoes_permitidas_bloco.get(chave_bloco, set()):
            pontos += 20
        else:
            motivos.append("FUNCAO_FORA_DO_BLOCO")

        # 20 — evidência factual: exige termos de conteúdo suficientes.
        termos_conteudo = obter_palavras_conteudo(texto)
        if len(termos_conteudo) >= 12:
            pontos += 20
        elif len(termos_conteudo) >= 8:
            pontos += 12
        else:
            motivos.append("EVIDENCIA_FACTUAL_FRACA")

        # 15 — completude/autonomia.
        frases = [x.strip() for x in re.split(r'(?<=[.!?])\s+', texto) if x.strip()]
        inicio = normalizar_assunto_texto(frases[0]) if frases else ""
        finais_ruins = ("e", "ou", "que", "de", "da", "do", "para", "com", "como", "quando", "onde", "incluindo", "conforme")
        if len(frases) >= 2 and not any(inicio.startswith(x + " ") for x in ("e", "ou", "que", "para", "por", "com", "quando", "onde")) and not any(normalizar_assunto_texto(frases[-1]).endswith(" " + x) for x in finais_ruins):
            pontos += 15
        else:
            motivos.append("AUTONOMIA_INSUFICIENTE")

        # 15 — integridade gramatical/pontuação. O portão já cobre o
        # essencial; aqui exigimos pelo menos duas frases completas.
        if len(frases) >= 2 and all(re.search(r'[.!?]$', f) for f in frases):
            pontos += 15
        else:
            motivos.append("FRASES_INCOMPLETAS")

        # 10 — limpeza editorial/comercial.
        ok_ed, motivo_ed = diagnosticar_contaminacao_editorial(texto, candidato.get("identidade_fonte", {}))
        ok_com = fragmento_eh_comercialmente_limpo(texto, candidato.get("identidade_fonte", {}))
        if ok_ed and ok_com:
            pontos += 10
        else:
            motivos.append("CONTAMINACAO_EDITORIAL_COMERCIAL")

        # 5 — aderência lexical ao objetivo do bloco.
        norm = normalizar_assunto_texto(texto)
        termos_bloco = set(termos_blocos_normalizados.get(chave_bloco, []) or [])
        encontrados_bloco = sum(1 for termo in termos_bloco if termo and termo in norm)
        if encontrados_bloco >= 2:
            pontos += 5
        elif encontrados_bloco >= 1:
            pontos += 2
        else:
            motivos.append("BAIXA_ADERENCIA_LEXICAL_AO_BLOCO")

        # 10 — especificidade técnica. Penaliza texto genérico demais.
        sinais_tecnicos = {
            "vazao", "pressao", "temperatura", "viscosidade", "corrosao",
            "corrosivo", "impelidor", "rotor", "energia", "fluido",
            "sucção", "descarga", "manutencao", "instalacao", "seguranca",
            "altura", "perda", "tubulacao", "valvula", "potencia",
            "eficiencia", "operacao", "funcionamento", "aplicacao", "sistema"
        }
        encontrados = sum(1 for termo in sinais_tecnicos if termo in norm)
        if encontrados >= 3:
            pontos += 10
        elif encontrados >= 2:
            pontos += 6
        else:
            motivos.append("BAIXA_ESPECIFICIDADE_TECNICA")

        # 5 — diversidade lexical interna.
        tokens = [normalizar_assunto_texto(x).lower() for x in re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]*\b", texto)]
        tokens = [x for x in tokens if len(x) >= 4]
        diversidade = len(set(tokens)) / max(len(tokens), 1)
        if diversidade >= 0.52:
            pontos += 5
        elif diversidade >= 0.45:
            pontos += 3
        else:
            motivos.append("BAIXA_DIVERSIDADE_LEXICAL")

        return min(100, pontos), motivos

    def _candidato_elegivel_para_bloco(item, chave_bloco, hashes_bloqueados, diagnostico=None):
        candidato = item.get('candidato') if isinstance(item, dict) else None
        if not isinstance(candidato, dict):
            if diagnostico is not None:
                diagnostico['ITEM_INVALIDO'] = diagnostico.get('ITEM_INVALIDO', 0) + 1
            return None
        candidato = dict(candidato)
        h = candidato.get('hash') or gerar_hash_trecho(candidato.get('texto', ''))
        candidato['hash'] = h
        if not h or h in hashes_selecionados or h in hashes_bloqueados:
            if diagnostico is not None:
                diagnostico['HASH_BLOQUEADO'] = diagnostico.get('HASH_BLOQUEADO', 0) + 1
            return None
        candidato['bloco_mead'] = chave_bloco
        candidato['_pontuacao_original_selecao'] = float(item.get('pontuacao', 0) or 0)

        def rejeitar(motivo):
            if diagnostico is not None:
                diagnostico[motivo] = diagnostico.get(motivo, 0) + 1
            return None

        if not fragmento_compativel_com_funcao_do_bloco(candidato, chave_bloco):
            return rejeitar('FUNCAO_EDITORIAL_INCOMPATIVEL_COM_BLOCO')

        if not candidato_eh_utilizavel(candidato):
            return rejeitar('CANDIDATO_NAO_UTILIZAVEL')

        ok_ed, motivo_ed = diagnosticar_contaminacao_editorial(
            candidato.get('texto', ''), candidato.get('identidade_fonte', {})
        )
        if not ok_ed:
            return rejeitar(f'CONTAMINACAO:{motivo_ed}')

        ok_final, motivo_final = auditar_fragmento_final_python(
            candidato.get('texto', ''), candidato.get('identidade_fonte', {})
        )
        if not ok_final:
            return rejeitar(f'AUDITORIA_FINAL:{motivo_final}')

        ok_q, nota_q, motivo_q = portao_qualidade_fragmento(
            candidato.get('texto', ''), candidato.get('identidade_fonte', {}), tema
        )
        candidato['nota_portao_qualidade'] = nota_q
        candidato['motivo_portao_qualidade'] = motivo_q
        if not ok_q:
            return rejeitar(f'PORTAO_QUALIDADE:{motivo_q}')

        # NOTA MÁXIMA INDIVIDUAL: somente trechos acima da nota mínima podem seguir.
        nota_otima, motivos_otimos = _nota_maxima_trecho(candidato, chave_bloco)
        candidato['nota_trecho_otimo'] = nota_otima
        candidato['motivos_nota_trecho_otimo'] = motivos_otimos
        NOTA_MINIMA_TRECHO = 85
        if nota_otima < NOTA_MINIMA_TRECHO:
            motivo = 'NOTA_TRECHO_ABAIXO_DO_MINIMO:' + str(NOTA_MINIMA_TRECHO) + ':' + ','.join(motivos_otimos or ['ABAIXO_DO_MINIMO'])
            return rejeitar(motivo)

        if diagnostico is not None:
            diagnostico['APROVADO_NOTA_MINIMA'] = diagnostico.get('APROVADO_NOTA_MINIMA', 0) + 1
        return candidato

    # ========================================================
    # 06.9 — PESQUISA COMPLEMENTAR POR LACUNA DO BLOCO
    # ========================================================
    # Patrimônio total suficiente NÃO significa cobertura editorial
    # suficiente. Se um bloco não consegue formar 3 trechos acima das notas mínimas,
    # Python pode pesquisar evidência adicional exclusivamente para
    # aquele bloco e para os checkboxes que o autorizam.
    #
    # Isso é deliberadamente diferente de uma pesquisa genérica por
    # "mais fontes": a lacuna do bloco é a razão da nova pesquisa.
    # ========================================================
    def _pesquisar_lacuna_editorial_por_bloco(chave_bloco):
        estrutura = normalizar_checkboxes_editoriais(estrutura_editorial)
        assuntos_bloco = sorted(
            blocos_autorizados_pelos_checkboxes(estrutura).get(chave_bloco, set())
        )
        if not assuntos_bloco:
            return 0

        termos_base = []
        for assunto in assuntos_bloco:
            termos_base.extend(TERMOS_PESQUISA_CHECKBOX.get(assunto, []))

        # Preferir expressões técnicas/factuais nas buscas; o checkbox
        # comercial autoriza a dimensão comercial, mas não autoriza
        # propaganda, CTA ou catálogo promocional.
        qualificadores = {
            "bloco_1": ["manual", "documento técnico"],
            "bloco_2": ["manual", "especificação técnica"],
            "bloco_3": ["manual", "procedimento técnico"],
            "bloco_4": ["especificação técnica", "suporte técnico"],
            "bloco_5": ["especificação técnica", "seleção técnica"],
        }.get(chave_bloco, ["informação técnica"])

        # Mantém no máximo duas buscas por lacuna para evitar transformar
        # uma falha de cobertura em pesquisa indiscriminada.
        consultas = []
        vistos_consulta = set()
        for termo in termos_base:
            for qualificador in qualificadores:
                consulta = f"{tema} {termo} {qualificador}".strip()
                chave = normalizar_assunto_texto(consulta)
                if chave in vistos_consulta:
                    continue
                vistos_consulta.add(chave)
                consultas.append(consulta)
                if len(consultas) >= 2:
                    break
            if len(consultas) >= 2:
                break

        urls = []
        vistas = set()
        for consulta in consultas:
            try:
                resultados = pesquisar(consulta, limite=6)
            except Exception as erro:
                print("⚠️ PESQUISA DE LACUNA FALHOU:", chave_bloco, repr(erro))
                continue
            for url in resultados or []:
                url = str(url or "").strip()
                if not url or url in vistas:
                    continue
                if not re.search(r"\.br(?:/|$)", url.casefold()):
                    continue
                vistas.add(url)
                urls.append(url)
                if len(urls) >= 8:
                    break
            if len(urls) >= 8:
                break

        adicionados = 0
        existentes = {
            (str((item.get("candidato") or {}).get("hash", ""))
             or gerar_hash_trecho((item.get("candidato") or {}).get("texto", "")))
            for item in candidatos_por_bloco.get(chave_bloco, [])
            if isinstance(item, dict)
        }

        for url in urls:
            try:
                pagina = coletar_pagina(url)
            except Exception:
                continue
            if not isinstance(pagina, dict):
                continue
            texto_original = str(pagina.get("texto", "") or "").strip()
            if not texto_original:
                continue

            frases = [
                x.strip() for x in re.split(r"(?<=[.!?])\s+", texto_original)
                if x.strip()
            ]
            identidade = pagina.get("identidade_fonte", {})
            if not isinstance(identidade, dict):
                identidade = identificar_empresa_fonte(texto_original, url)

            for tamanho in (2, 3, 4):
                for inicio in range(0, max(0, len(frases) - tamanho + 1)):
                    trecho = " ".join(frases[inicio:inicio + tamanho]).strip()
                    if not trecho or not re.search(r"[.!?][\"'”’»\)\]}]*$", trecho):
                        continue
                    h = gerar_hash_trecho(trecho)
                    if not h or h in existentes:
                        continue

                    candidato = {
                        "id": f"LACUNA_{chave_bloco}_{h[:12]}",
                        "hash": h,
                        "texto": trecho,
                        "url": url,
                        "fonte": f"LACUNA_{chave_bloco}",
                        "tipo": "pesquisa_lacuna",
                        "pdf": bool(re.search(r"\.pdf(?:$|[?#])", url.casefold())),
                        "palavras": len(trecho.split()),
                        "identidade_fonte": identidade,
                    }

                    if not fragmento_pertence_ao_tema(trecho, tema):
                        continue
                    classificacao = classificar_funcao_editorial_fragmento(trecho)
                    candidato["_funcao_editorial"] = classificacao
                    if not fragmento_compativel_com_funcao_do_bloco(candidato, chave_bloco):
                        continue
                    try:
                        if not candidato_eh_utilizavel(candidato):
                            continue
                    except Exception:
                        continue
                    ok_ed, _ = diagnosticar_contaminacao_editorial(trecho, identidade)
                    if not ok_ed:
                        continue
                    ok_final, _ = auditar_fragmento_final_python(trecho, identidade)
                    if not ok_final:
                        continue
                    ok_q, nota_q, motivo_q = portao_qualidade_fragmento(trecho, identidade, tema)
                    candidato["nota_portao_qualidade"] = nota_q
                    candidato["motivo_portao_qualidade"] = motivo_q
                    if not ok_q:
                        continue
                    try:
                        candidato["_pontuacao_original_selecao"] = float(
                            calcular_pontuacao_bloco(candidato, chave_bloco)
                        )
                    except Exception:
                        candidato["_pontuacao_original_selecao"] = 0.0
                    # Só registrar como candidato se atingir a mesma nota
                    # máxima exigida pela seleção normal.
                    nota_otima, motivos_otimos = _nota_maxima_trecho(candidato, chave_bloco)
                    candidato["nota_trecho_otimo"] = nota_otima
                    candidato["motivos_nota_trecho_otimo"] = motivos_otimos
                    existentes.add(h)
                    candidatos_por_bloco.setdefault(chave_bloco, []).append({
                        "candidato": candidato,
                        "pontuacao": candidato["_pontuacao_original_selecao"]
                    })
                    adicionados += 1

                    if adicionados >= 30:
                        break
                if adicionados >= 30:
                    break
            if adicionados >= 30:
                break

        if adicionados:
            candidatos_por_bloco[chave_bloco].sort(
                key=lambda item: (item.get("pontuacao", 0),
                                  item.get("candidato", {}).get("nota_portao_qualidade", 0)),
                reverse=True
            )
        print(
            "PESQUISA COMPLEMENTAR POR LACUNA:", chave_bloco,
            "assuntos:", assuntos_bloco,
            "consultas:", consultas,
            "URLs:", len(urls),
            "NOVOS CANDIDATOS:", adicionados
        )
        return adicionados

    # Antes da alocação global, mede a capacidade REAL de 100/100.
    # Se um bloco estiver abaixo de 3, pesquisa somente a lacuna dele.
    for _numero_bloco in range(1, 6):
        _chave_bloco = f"bloco_{_numero_bloco}"
        _diagnostico_lacuna = {}
        _elegiveis_locais = []
        for _item in candidatos_por_bloco.get(_chave_bloco, []):
            _cand = _candidato_elegivel_para_bloco(
                _item, _chave_bloco, set(), _diagnostico_lacuna
            )
            if _cand is not None:
                _elegiveis_locais.append(_cand)
        print(
            "CAPACIDADE NOTA MINIMA ANTES DA PESQUISA:",
            _chave_bloco, len(_elegiveis_locais), "/ 3"
        )
        if len(_elegiveis_locais) < 3:
            _pesquisar_lacuna_editorial_por_bloco(_chave_bloco)

    # ========================================================
    # v51 — COMBINAÇÃO NARRATIVA 100/100
    # ========================================================
    def _conceitos_fragmento_narrativo(candidato):
        texto = str(candidato.get("texto", "") or "")
        stop = {"para","com","sem","sobre","entre","como","uma","um","dos","das","que","por","pelo","pela","de","da","do","em","no","na","ao","e","ou","se","ser","tem","ter","mais","tambem","também","cada","quando","onde","isso","esse","essa","este","esta","seu","sua","seus","suas","pode","podem","sendo","assim","qual","quais"}
        return {t for t in re.findall(r"\b[a-z0-9]{4,}\b", normalizar_assunto_texto(texto)) if t not in stop}

    def _funcoes_fragmento_narrativo(candidato):
        return set((candidato.get("_funcao_editorial", {}) or {}).get("funcoes", []) or [])

    def _pontuar_progressao_evidencias(combinacao, chave_bloco):
        if len(combinacao) != 3:
            return 0, list(combinacao), {"motivo":"NAO_SAO_3_EVIDENCIAS"}
        import itertools
        melhor = None
        for ordem in itertools.permutations(combinacao, 3):
            conceitos = [_conceitos_fragmento_narrativo(x) for x in ordem]
            funcoes = [_funcoes_fragmento_narrativo(x) for x in ordem]
            novos_2 = conceitos[1] - conceitos[0]
            novos_3 = conceitos[2] - (conceitos[0] | conceitos[1])
            uniao = conceitos[0] | conceitos[1] | conceitos[2]
            maior = max(len(c) for c in conceitos)
            soma = sum(len(c) for c in conceitos)
            sobreposicao = max(0, soma - len(uniao))
            fontes = {x.get("fonte") for x in ordem}
            funcoes_unicas = set().union(*funcoes)
            prioridades = [
                {"contexto","funcionamento","apresentacao","definicao"},
                {"funcionamento","tecnico","conhecimento","criterios","aplicacao"},
                {"aplicacao","beneficios","beneficio","criterios","comercial","suporte","solucao","institucional","tecnico"},
            ]
            encaixes = sum(1 for i in range(3) if funcoes[i] & prioridades[i])
            razao_repeticao = sobreposicao / max(1, soma)
            nota = 100
            motivos = []
            if len(novos_2) < 2: nota -= 20; motivos.append("P2_SEM_NOVOS_CONCEITOS")
            if len(novos_3) < 2: nota -= 20; motivos.append("P3_SEM_NOVOS_CONCEITOS")
            if maior >= len(uniao) and len(uniao) < 14: nota -= 10; motivos.append("COBERTURA_CONCEITUAL_BAIXA")
            if len(fontes) < 2: nota -= 5; motivos.append("BAIXA_DIVERSIDADE_DE_FONTES")
            if len(funcoes_unicas) < 2: nota -= 5; motivos.append("BAIXA_DIVERSIDADE_FUNCIONAL")
            if razao_repeticao > 0.42: nota -= 20; motivos.append("SOBREPOSICAO_CONCEITUAL_ALTA")
            if encaixes == 0: nota -= 10; motivos.append("SEM_PROGRESSAO_EDITORIAL")
            if len(uniao) <= maior + 3: nota -= 10; motivos.append("EVIDENCIAS_POUCO_COMPLEMENTARES")
            registro = {"nota":max(0,int(round(nota))),"ordem":list(ordem),"novos_p2":len(novos_2),"novos_p3":len(novos_3),"conceitos_unicos":len(uniao),"funcoes_unicas":len(funcoes_unicas),"fontes_unicas":len(fontes),"encaixes_progressao":encaixes,"repeticao_relativa":round(razao_repeticao,3),"motivos":motivos}
            if melhor is None or (registro["nota"],registro["conceitos_unicos"],registro["novos_p3"],registro["novos_p2"]) > (melhor["nota"],melhor["conceitos_unicos"],melhor["novos_p3"],melhor["novos_p2"]):
                melhor=registro
        return melhor["nota"], melhor["ordem"], melhor

    def _combinacao_eh_editorialmente_diversa(combinacao):
        """Gate pré-Ollama contra janelas muito próximas/redundantes."""
        if not isinstance(combinacao, (list, tuple)) or len(combinacao) != 3:
            return False, "COMBINACAO_INVALIDA"
        for i in range(len(combinacao)):
            for j in range(i + 1, len(combinacao)):
                texto_i = str(combinacao[i].get("texto", "") or "").strip()
                texto_j = str(combinacao[j].get("texto", "") or "").strip()
                sim = _similaridade_textual(texto_i, texto_j)
                jac = _jaccard_conceitual(texto_i, texto_j)
                mesma_fonte = combinacao[i].get("fonte") == combinacao[j].get("fonte")
                if sim >= 0.78:
                    return False, f"REDUNDANCIA_TEXTUAL_P{i+1}_P{j+1}:sim={sim:.3f}"
                if jac >= 0.68 and sim >= 0.50:
                    return False, f"REDUNDANCIA_CONCEITUAL_P{i+1}_P{j+1}:sim={sim:.3f}:conceito={jac:.3f}"
                if mesma_fonte and jac >= 0.60 and sim >= 0.44:
                    return False, f"JANELAS_PROXIMAS_MESMA_FONTE_P{i+1}_P{j+1}:sim={sim:.3f}:conceito={jac:.3f}"
        return True, "OK"

    def _avaliar_combinacao_bloco(combinacao, chave_bloco):
        total_palavras = sum(_palavras_fragmento(x) for x in combinacao)
        # Pontuação editorial da combinação, sem mínimo ou máximo de palavras.
        pontuacao = 0.0
        fontes = set()
        for cand in combinacao:
            pontuacao += float(cand.get('_pontuacao_original_selecao', 0) or 0)
            # A nota do portão é o principal sinal de qualidade disponível
            # depois dos gates duros. Ela pesa mais que a pontuação editorial
            # bruta para evitar que um trecho apenas 'compatível' vença um
            # trecho mais limpo e completo.
            # Nota máxima individual é requisito duro; aqui ela funciona
            # como confirmação da qualidade do conjunto.
            pontuacao += float(cand.get('nota_trecho_otimo', 0) or 0) * 2.00
            pontuacao += float(cand.get('nota_portao_qualidade', 0) or 0) * 0.50
            fontes.add(cand.get('fonte'))
        pontuacao += len(fontes) * 3
        return pontuacao, total_palavras

    # ========================================================
    # CONTROLE DE COMPLEXIDADE DA SELEÇÃO GLOBAL — v66
    # ========================================================
    # O reservatório pode ser grande para garantir evidência, mas a etapa
    # de combinações NÃO pode crescer com C(n,3) sem limite.
    #
    # Exemplo real observado:
    #   38 candidatos -> milhares de combinações
    #   37 candidatos -> milhares de combinações
    # Isso pode consumir dezenas de minutos antes de qualquer chamada ao
    # Ollama. O reservatório continua grande; somente o pool usado para
    # combinações é limitado.
    # ========================================================

    MAX_CANDIDATOS_COMBINACAO = 25
    MAX_COMBINACOES_POR_BLOCO = 2300  # C(25,3)
    # 15 era rápido, mas podia excluir um candidato exclusivo de baixo ranking.
    # 25 mantém o custo controlado e reduz muito esse falso bloqueio.

    def _pontuacao_candidato_para_pool(cand):
        try:
            return (
                float(cand.get('nota_trecho_otimo', 0) or 0) * 2.0
                + float(cand.get('nota_portao_qualidade', 0) or 0) * 0.5
                + float(cand.get('_pontuacao_original_selecao', 0) or 0)
                + min(_palavras_fragmento(cand), 140) * 0.05
            )
        except Exception:
            return 0.0

    def _chave_diversidade_candidato(cand):
        identidade = cand.get('identidade_fonte', {})
        if not isinstance(identidade, dict):
            identidade = {}
        fonte = str(cand.get('fonte', '') or '').strip().casefold()
        nome = str(identidade.get('nome', '') or '').strip().casefold()
        dominio = str(identidade.get('dominio', '') or '').strip().casefold()
        return (fonte, nome, dominio)

    def _reduzir_pool_para_combinacao(candidatos, chave_bloco, limite=None, expansivo=False):
        """
        Reduz somente o pool de COMBINAÇÃO.

        Os candidatos originais permanecem intactos no reservatório e podem
        continuar sendo usados pelo resgate/reseleção. A redução ocorre
        exclusivamente antes de itertools.combinations(..., 3).
        """
        if expansivo:
            # Modo expansivo usado somente na contingência/resgate global.
            # Aqui NÃO descartamos candidatos elegíveis por ranking.
            # O objetivo é justamente recuperar candidatos exclusivos ou
            # compartilhados que poderiam resolver a colisão entre blocos.
            limite = int(limite or 60)
        else:
            limite = int(limite or MAX_CANDIDATOS_COMBINACAO)
        limite = max(3, limite)

        unicos = []
        hashes = set()
        for cand in candidatos or []:
            if not isinstance(cand, dict):
                continue
            h = str(
                cand.get('hash')
                or gerar_hash_trecho(cand.get('texto', ''))
                or ''
            ).strip()
            if not h or h in hashes:
                continue
            hashes.add(h)
            unicos.append(cand)

        ordenados = sorted(
            unicos,
            key=_pontuacao_candidato_para_pool,
            reverse=True
        )

        if expansivo:
            # CONTINGÊNCIA: preservar todos os candidatos elegíveis dentro
            # da janela expansiva. Não aplicar a redução por diversidade,
            # porque ela pode esconder justamente o candidato exclusivo
            # necessário para fechar os 15 hashes globais.
            pool = ordenados[:limite]
        else:
            pool = []
            chaves_diversidade = set()

            # Primeiro preserva diversidade de origem/identidade entre os melhores.
            for cand in ordenados:
                chave_div = _chave_diversidade_candidato(cand)
                if chave_div in chaves_diversidade:
                    continue
                pool.append(cand)
                chaves_diversidade.add(chave_div)
                if len(pool) >= limite:
                    break

            # Depois completa o pool com os melhores restantes.
            if len(pool) < limite:
                hashes_pool = {
                    str(x.get('hash') or gerar_hash_trecho(x.get('texto', '')) or '').strip()
                    for x in pool
                }
                for cand in ordenados:
                    h = str(
                        cand.get('hash')
                        or gerar_hash_trecho(cand.get('texto', ''))
                        or ''
                    ).strip()
                    if h in hashes_pool:
                        continue
                    pool.append(cand)
                    hashes_pool.add(h)
                    if len(pool) >= limite:
                        break

        n = len(pool)
        combinacoes_teoricas = n * (n - 1) * (n - 2) // 6 if n >= 3 else 0

        print(
            'POOL DE COMBINAÇÃO', chave_bloco,
            '| elegíveis:', len(unicos),
            '| pool limitado:', len(pool),
            '| combinações possíveis:', combinacoes_teoricas,
            '| limite informativo:', MAX_COMBINACOES_POR_BLOCO
        )

        return pool

    def _selecionar_conjunto_bloco_com_capacidade(chave_bloco):
        disponiveis = []
        diagnostico = {}
        for item in candidatos_por_bloco.get(chave_bloco, []):
            cand = _candidato_elegivel_para_bloco(item, chave_bloco, set(), diagnostico)
            if cand is not None:
                disponiveis.append(cand)

        # Não precisamos combinar centenas de itens; os candidatos já
        # chegam ordenados por pontuação. Mantemos uma janela ampla para
        # não sacrificar capacidade em favor do primeiro ranking.
        

        disponiveis_pool = _reduzir_pool_para_combinacao(
            disponiveis,
            chave_bloco
        )

        melhor = None
        melhor_chave = None
        import itertools
        for combinacao in itertools.combinations(disponiveis_pool, 3):
            hashes = [x.get('hash') for x in combinacao]
            if len(set(hashes)) < 3:
                continue
            ok_diversidade, motivo_diversidade = _combinacao_eh_editorialmente_diversa(combinacao)
            if not ok_diversidade:
                continue
            nota_narrativa, ordem_narrativa, detalhes_narrativa = _pontuar_progressao_evidencias(combinacao, chave_bloco)
            if nota_narrativa != 100:
                continue
            total = sum(_palavras_fragmento(x) for x in combinacao)
            nota, _ = _avaliar_combinacao_bloco(combinacao, chave_bloco)
            combinacao = list(ordem_narrativa)
            for _cand in combinacao:
                _cand["nota_combinacao_narrativa"] = nota_narrativa
                _cand["detalhes_combinacao_narrativa"] = detalhes_narrativa
            chave = (nota_narrativa, nota, total)
            if melhor is None or chave > melhor_chave:
                melhor = list(combinacao)
                melhor_chave = chave

        return melhor

    # ========================================================
    # ALOCAÇÃO GLOBAL REAL — SEM COLISÃO GREEDY
    # ========================================================
    # Primeiro construímos todas as combinações válidas por bloco.
    # Depois procuramos uma solução GLOBAL em que os 15 hashes sejam
    # distintos. O algoritmo anterior resolvia cada bloco isoladamente
    # e só depois tentava corrigir colisões; isso podia bloquear B5 mesmo
    # havendo uma combinação global possível.

    ordem_blocos = sorted(
        range(1, 6),
        key=lambda n: len(candidatos_por_bloco.get(f'bloco_{n}', []))
    )

    # IMPORTANTE: ordem_blocos contém números (1..5). Sempre converter para
    # a chave textual bloco_N antes de acessar candidatos/diagnósticos.
    # Misturar os dois formatos faz o diagnóstico enxergar todos os pools
    # como vazios e pode bloquear o resgate artificialmente.

    combinacoes_por_bloco = {}
    import itertools

    for numero_bloco in ordem_blocos:
        chave_bloco = f'bloco_{numero_bloco}'
        disponiveis = []
        diagnostico = {}
        for item in candidatos_por_bloco.get(chave_bloco, []):
            cand = _candidato_elegivel_para_bloco(item, chave_bloco, set(), diagnostico)
            if cand is not None:
                disponiveis.append(cand)

        print('DIAGNOSTICO DE FILTROS', chave_bloco, ':', diagnostico)
        
        disponiveis_pool = _reduzir_pool_para_combinacao(
            disponiveis,
            chave_bloco
        )

        combinacoes = []
        for combinacao in itertools.combinations(disponiveis_pool, 3):
            hashes = [x.get('hash') for x in combinacao]
            if len(set(hashes)) < 3:
                continue
            ok_diversidade, motivo_diversidade = _combinacao_eh_editorialmente_diversa(combinacao)
            if not ok_diversidade:
                continue
            nota_narrativa, ordem_narrativa, detalhes_narrativa = _pontuar_progressao_evidencias(combinacao, chave_bloco)
            if nota_narrativa != 100:
                continue
            nota, total = _avaliar_combinacao_bloco(combinacao, chave_bloco)
            combinacao = list(ordem_narrativa)
            for _cand in combinacao:
                _cand["nota_combinacao_narrativa"] = nota_narrativa
                _cand["detalhes_combinacao_narrativa"] = detalhes_narrativa
            chave = nota
            combinacoes.append((chave, total, list(combinacao)))

        combinacoes.sort(key=lambda x: (x[0], x[1]), reverse=True)
        # Mantemos uma janela suficientemente ampla para a busca global.
        combinacoes_por_bloco[chave_bloco] = combinacoes

        capacidades = sorted(
            [_palavras_fragmento(x) for x in disponiveis],
            reverse=True
        )
        print(
            'CAPACIDADE GLOBAL', chave_bloco,
            '| candidatos:', len(disponiveis),
            '| combinações válidas:', len(combinacoes),
            '| 3 maiores:', sum(capacidades[:3]) if capacidades else 0
        )

        if not combinacoes:
            print(
                '🔴 SELEÇÃO BLOQUEADA:', chave_bloco,
                'não possui combinação de 3 fragmentos elegíveis após os filtros editoriais/factuais.'
            )
            print(
                'DIAGNÓSTICO:', chave_bloco,
                '| candidatos após filtros:', len(disponiveis),
                '| mínimo necessário: 3'
            )
            print('CAPACIDADE DOS CANDIDATOS:', capacidades)
            print('CAPACIDADE DOS 3 MAIORES:', sum(capacidades[:3]) if capacidades else 0)
            print('CAPACIDADE TOTAL DOS CANDIDATOS:', sum(capacidades))
            print('MOTIVO CLASSIFICADO: SEM COMBINAÇÃO EDITORIALMENTE ELEGÍVEL')
            return resultado_vazio

    # Busca global com retrocesso. Em vez de escolher um conjunto por bloco
    # e torcer para não colidir, escolhemos os 5 conjuntos simultaneamente.
    solucao_global = {}
    nos_alocacao_global = 0
    MAX_NOS_ALOCACAO_GLOBAL = 20000

    def _buscar_alocacao_global(indice, usados):
        nonlocal nos_alocacao_global
        nos_alocacao_global += 1
        if nos_alocacao_global > MAX_NOS_ALOCACAO_GLOBAL:
            return False
        if indice >= len(ordem_blocos):
            return True

        numero_bloco = ordem_blocos[indice]
        chave_bloco = f'bloco_{numero_bloco}'

        # Primeiro tenta as melhores combinações; se uma delas consumir um
        # fragmento necessário por outro bloco, o backtracking testa a próxima.
        for _nota, _total, combinacao in combinacoes_por_bloco.get(chave_bloco, []):
            hashes = {x.get('hash') for x in combinacao}
            if hashes & usados:
                continue
            solucao_global[chave_bloco] = combinacao
            if _buscar_alocacao_global(indice + 1, usados | hashes):
                return True
            solucao_global.pop(chave_bloco, None)

        return False

    if not _buscar_alocacao_global(0, set()):
        print('⚠️ ALOCAÇÃO ESTRITA NÃO ENCONTRADA — ATIVANDO ALOCAÇÃO GLOBAL DE CONTINGÊNCIA')
        print('REGRA: manter todos os gates individuais; relaxar somente a exigência narrativa 100/100 da combinação.')
        print('OBJETIVO: obter 15 fragmentos ÚNICOS sem ressuscitar candidatos rejeitados.')

        # ------------------------------------------------------------
        # CONTINGÊNCIA GLOBAL
        # ------------------------------------------------------------
        # O problema observado na v61 não era falta de candidatos: cada
        # bloco tinha combinações válidas. O bloqueio acontecia porque
        # a exigência de progressão narrativa 100/100 foi aplicada antes
        # da resolução global. Isso pode concentrar os mesmos candidatos
        # em vários blocos.
        #
        # Aqui NÃO reabrimos fontes nem removemos filtros individuais.
        # Reutilizamos somente candidatos que já passaram por
        # _candidato_elegivel_para_bloco(). A única flexibilização é a
        # pontuação da combinação narrativa, permitindo conjuntos
        # editorialmente bons porém não perfeitos.
        # ------------------------------------------------------------
        combinacoes_contingencia = {}

        for numero_bloco in ordem_blocos:
            chave_bloco = f'bloco_{numero_bloco}'
            disponiveis_cont = []
            diagnostico_cont = {}
            for item in candidatos_por_bloco.get(chave_bloco, []):
                cand = _candidato_elegivel_para_bloco(item, chave_bloco, set(), diagnostico_cont)
                if cand is not None:
                    disponiveis_cont.append(cand)

            disponiveis_cont_pool = _reduzir_pool_para_combinacao(
                disponiveis_cont,
                chave_bloco,
                limite=60,
                expansivo=True
            )

            combos_cont = []
            for combinacao in itertools.combinations(disponiveis_cont_pool, 3):
                hashes = [x.get('hash') for x in combinacao]
                if len(set(hashes)) < 3:
                    continue
                ok_diversidade, motivo_diversidade = _combinacao_eh_editorialmente_diversa(combinacao)
                if not ok_diversidade:
                    continue

                nota_narrativa, ordem_narrativa, detalhes_narrativa = _pontuar_progressao_evidencias(
                    combinacao, chave_bloco
                )
                # 80 é contingência, não novo padrão. A seleção estrita
                # continua sendo sempre tentada primeiro.
                if nota_narrativa < 80:
                    continue

                nota, total = _avaliar_combinacao_bloco(combinacao, chave_bloco)
                combinacao = list(ordem_narrativa)
                for _cand in combinacao:
                    _cand['nota_combinacao_narrativa'] = nota_narrativa
                    _cand['detalhes_combinacao_narrativa'] = detalhes_narrativa
                combos_cont.append((nota_narrativa, nota, total, list(combinacao)))

            combos_cont.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
            # A janela é propositalmente ampla para a busca global encontrar
            # alternativas que não colidam com os demais blocos.
            combinacoes_contingencia[chave_bloco] = combos_cont
            print(
                'CONTINGÊNCIA', chave_bloco,
                '| candidatos:', len(disponiveis_cont),
                '| combinações:', len(combos_cont),
                '| melhor narrativa:', combos_cont[0][0] if combos_cont else 0
            )

        solucao_contingencia = {}
        melhor_solucao_contingencia = None
        nos_contingencia = 0
        # A busca antiga tinha teto de 20.000 nós e ordem fixa. Isso podia
        # produzir falso negativo mesmo quando existia uma combinação global.
        # A nova busca usa ordem dinâmica e poda por capacidade de hashes.
        MAX_NOS_CONTINGENCIA = 250000

        def _hashes_registro(registro):
            return {
                str(x.get('hash') or gerar_hash_trecho(x.get('texto', '')) or '').strip()
                for x in registro[3]
            }

        def _buscar_contingencia(usados, acumulado, blocos_restantes):
            nonlocal melhor_solucao_contingencia, nos_contingencia
            nos_contingencia += 1
            if nos_contingencia > MAX_NOS_CONTINGENCIA:
                return False

            if not blocos_restantes:
                nota_total = sum(x[0] + (x[1] * 0.01) for x in acumulado)
                if (
                    melhor_solucao_contingencia is None
                    or nota_total > melhor_solucao_contingencia[0]
                ):
                    # Snapshot completo da solução. O dicionário da busca é
                    # mutável e continua recebendo pop() durante o backtracking.
                    snapshot_solucao = {
                        chave: list(grupo)
                        for chave, grupo in solucao_contingencia.items()
                    }
                    melhor_solucao_contingencia = (
                        nota_total,
                        snapshot_solucao
                    )
                return True

            # --------------------------------------------------------
            # PODA 1: cada bloco restante precisa ainda ter uma combinação
            # de 3 hashes que não colida com os já utilizados.
            # --------------------------------------------------------
            disponiveis_por_bloco = {}
            uniao_disponiveis = set()

            for chave_bloco in blocos_restantes:
                registros_livres = []
                hashes_livres_bloco = set()

                for registro in combinacoes_contingencia.get(chave_bloco, []):
                    hashes = _hashes_registro(registro)
                    if len(hashes) != 3 or hashes & usados:
                        continue
                    registros_livres.append(registro)
                    hashes_livres_bloco.update(hashes)

                if not registros_livres:
                    return False

                disponiveis_por_bloco[chave_bloco] = registros_livres
                uniao_disponiveis.update(hashes_livres_bloco)

            # Cada bloco restante precisa de 3 hashes novos.
            if len(uniao_disponiveis) < (3 * len(blocos_restantes)):
                return False

            # --------------------------------------------------------
            # PODA 2 / HEURÍSTICA: escolher primeiro o bloco com menor
            # número de combinações livres. Isso reduz brutalmente a árvore.
            # --------------------------------------------------------
            chave_escolhida = min(
                blocos_restantes,
                key=lambda chave: len(disponiveis_por_bloco[chave])
            )

            encontrou = False
            candidatos_registros = disponiveis_por_bloco[chave_escolhida]

            # Prioriza conjuntos que preservam mais alternativas para os
            # demais blocos. O score original continua sendo o primeiro
            # critério; a quantidade de colisões é o desempate.
            demais = [
                b for b in blocos_restantes
                if b != chave_escolhida
            ]

            def _custo_colisao(registro):
                hashes = _hashes_registro(registro)
                colisoes = 0
                for b in demais:
                    for outro in combinacoes_contingencia.get(b, []):
                        if not (_hashes_registro(outro) & hashes):
                            break
                    else:
                        colisoes += 1
                return colisoes

            candidatos_registros = sorted(
                candidatos_registros,
                key=lambda r: (
                    _custo_colisao(r),
                    -r[0],
                    -r[1],
                    -r[2]
                )
            )

            for registro in candidatos_registros:
                _narr, _nota, _total, combinacao = registro
                hashes = _hashes_registro(registro)
                if hashes & usados:
                    continue

                solucao_contingencia[chave_escolhida] = combinacao
                if _buscar_contingencia(
                    usados | hashes,
                    acumulado + [registro],
                    demais
                ):
                    encontrou = True
                    # Não paramos imediatamente: se já existe solução,
                    # continuamos procurando somente até encontrar uma
                    # solução melhor dentro do orçamento de nós.
                solucao_contingencia.pop(chave_escolhida, None)

                if nos_contingencia > MAX_NOS_CONTINGENCIA:
                    break

            return encontrou

        _buscar_contingencia(set(), [], list(ordem_blocos))
        print('NÓS DA BUSCA GLOBAL DE CONTINGÊNCIA:', nos_contingencia)

        if melhor_solucao_contingencia is None:
            # ============================================================
            # v63 — RESGATE DE COBERTURA APÓS COLISÃO GLOBAL
            # ============================================================
            # A ausência de uma solução global NÃO significa, por si só,
            # que as fontes foram pesquisadas de forma insuficiente.
            # Entretanto, quando vários blocos compartilham os mesmos
            # fragmentos, o pool aprovado pode ficar sem 15 hashes distintos.
            # Nesse caso fazemos pesquisa de lacuna somente para os blocos
            # mais apertados e reconstruímos a alocação.
            #
            # Importante: nenhuma barreira individual é relaxada. O resgate
            # apenas adiciona candidatos que passam novamente pelos mesmos
            # filtros.
            print('⚠️ COLISÃO GLOBAL SEM SOLUÇÃO — INICIANDO RESGATE DE COBERTURA')
            print('REGRA: pesquisar novas evidências apenas para os blocos mais restritos; manter todos os gates.')

            def _reconstruir_combinacoes_contingencia():
                resultado = {}
                for numero_bloco in ordem_blocos:
                    chave_bloco = f'bloco_{numero_bloco}'
                    disponiveis_cont = []
                    diagnostico_cont = {}
                    for item in candidatos_por_bloco.get(chave_bloco, []):
                        cand = _candidato_elegivel_para_bloco(
                            item, chave_bloco, set(), diagnostico_cont
                        )
                        if cand is not None:
                            disponiveis_cont.append(cand)

                    disponiveis_cont_pool = _reduzir_pool_para_combinacao(
                        disponiveis_cont,
                        chave_bloco,
                        limite=60,
                        expansivo=True
                    )

                    combos_cont = []
                    for combinacao in itertools.combinations(disponiveis_cont_pool, 3):
                        hashes = [x.get('hash') for x in combinacao]
                        if len(set(hashes)) < 3:
                            continue
                        ok_diversidade, _motivo = _combinacao_eh_editorialmente_diversa(combinacao)
                        if not ok_diversidade:
                            continue
                        nota_narrativa, ordem_narrativa, detalhes_narrativa = _pontuar_progressao_evidencias(
                            combinacao, chave_bloco
                        )
                        if nota_narrativa < 80:
                            continue
                        nota, total = _avaliar_combinacao_bloco(combinacao, chave_bloco)
                        combinacao = list(ordem_narrativa)
                        for _cand in combinacao:
                            _cand['nota_combinacao_narrativa'] = nota_narrativa
                            _cand['detalhes_combinacao_narrativa'] = detalhes_narrativa
                        combos_cont.append((nota_narrativa, nota, total, list(combinacao)))
                    combos_cont.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
                    resultado[chave_bloco] = combos_cont
                    print(
                        'POOL APÓS RESGATE', chave_bloco,
                        '| candidatos:', len(disponiveis_cont),
                        '| combinações:', len(combos_cont)
                    )
                return resultado

            # ============================================================
            # v64 — DIAGNÓSTICO DE EXCLUSIVIDADE GLOBAL
            # ============================================================
            # O número bruto de candidatos não mede a dificuldade real.
            # Um mesmo hash pode estar elegível para vários blocos.
            # O resgate deve priorizar o bloco que possui MENOS evidência
            # exclusiva, e não simplesmente o bloco com menos candidatos.
            #
            # Exemplo:
            #   B4 = 36 candidatos, mas muitos podem ser exclusivos de B4.
            #   B1 = 8 candidatos, todos compartilhados com B2/B5.
            # Nesse cenário B1 é o gargalo real.
            # ============================================================

            def _diagnosticar_exclusividade_global():
                mapa_hash_blocos = {}
                dados_blocos = {}

                for numero_bloco in ordem_blocos:
                    chave_bloco = f'bloco_{numero_bloco}'
                    diagnostico_tmp = {}
                    hashes_bloco = set()
                    candidatos_validos = []

                    for item in candidatos_por_bloco.get(chave_bloco, []):
                        cand = _candidato_elegivel_para_bloco(
                            item,
                            chave_bloco,
                            set(),
                            diagnostico_tmp
                        )
                        if cand is None:
                            continue

                        h = str(
                            cand.get('hash')
                            or gerar_hash_trecho(cand.get('texto', ''))
                            or ''
                        ).strip()

                        if not h:
                            continue

                        hashes_bloco.add(h)
                        candidatos_validos.append(cand)
                        mapa_hash_blocos.setdefault(h, set()).add(chave_bloco)

                    dados_blocos[chave_bloco] = {
                        'hashes': hashes_bloco,
                        'candidatos_validos': candidatos_validos,
                    }

                diagnostico_final = {}

                for chave_bloco, dados in dados_blocos.items():
                    exclusivos = 0
                    compartilhados = 0
                    grau_2 = 0
                    grau_3_ou_mais = 0

                    for h in dados['hashes']:
                        grau = len(mapa_hash_blocos.get(h, set()))

                        if grau == 1:
                            exclusivos += 1
                        else:
                            compartilhados += 1

                        if grau == 2:
                            grau_2 += 1
                        elif grau >= 3:
                            grau_3_ou_mais += 1

                    diagnostico_final[chave_bloco] = {
                        'validos': len(dados['hashes']),
                        'exclusivos': exclusivos,
                        'compartilhados': compartilhados,
                        'grau_2': grau_2,
                        'grau_3_ou_mais': grau_3_ou_mais,
                    }

                return diagnostico_final, mapa_hash_blocos

            for rodada_resgate in range(1, 4):
                diagnostico_exclusividade, mapa_hash_blocos = (
                    _diagnosticar_exclusividade_global()
                )

                print()
                print('==============================')
                print('RESGATE GLOBAL — EXCLUSIVIDADE POR BLOCO')
                print('==============================')

                for numero_bloco in ordem_blocos:
                    chave_bloco = f'bloco_{numero_bloco}'
                    dados = diagnostico_exclusividade.get(
                        chave_bloco,
                        {}
                    )
                    print(
                        chave_bloco,
                        '| válidos:', dados.get('validos', 0),
                        '| exclusivos:', dados.get('exclusivos', 0),
                        '| compartilhados:', dados.get('compartilhados', 0),
                        '| grau2:', dados.get('grau_2', 0),
                        '| grau3+:', dados.get('grau_3_ou_mais', 0)
                    )

                # Primeiro os blocos com menos evidência exclusiva.
                # Em empate, prioriza menor pool total e maior concentração
                # de colisões em outros blocos.
                contagens = []

                for numero_bloco in ordem_blocos:
                    chave_bloco = f'bloco_{numero_bloco}'
                    dados = diagnostico_exclusividade.get(
                        chave_bloco,
                        {}
                    )

                    exclusivos = int(
                        dados.get('exclusivos', 0) or 0
                    )
                    validos = int(
                        dados.get('validos', 0) or 0
                    )
                    compartilhados = int(
                        dados.get('compartilhados', 0) or 0
                    )

                    # Quanto menor a exclusividade, maior a prioridade.
                    # O segundo critério evita pesquisar um pool enorme
                    # enquanto outro bloco está realmente estrangulado.
                    contagens.append(
                        (
                            exclusivos,
                            validos,
                            -compartilhados,
                            chave_bloco
                        )
                    )

                contagens.sort()

                print()
                print(
                    'ORDEM DE RESGATE:',
                    [
                        item[3]
                        for item in contagens
                    ]
                )

                print('RESGATE GLOBAL — RODADA', rodada_resgate)

                adicionados_rodada = 0

                # Mantém no máximo três blocos por rodada, mas agora eles
                # são escolhidos pela falta de evidência exclusiva global.
                for _exclusivos, _validos, _compartilhados, chave_bloco in contagens[:3]:
                    antes = len(
                        candidatos_por_bloco.get(
                            chave_bloco,
                            []
                        )
                    )

                    adicionados = _pesquisar_lacuna_editorial_por_bloco(
                        chave_bloco
                    )

                    depois = len(
                        candidatos_por_bloco.get(
                            chave_bloco,
                            []
                        )
                    )

                    print(
                        'RESGATE RESULTADO',
                        chave_bloco,
                        '| antes:', antes,
                        '| depois:', depois,
                        '| adicionados:', adicionados
                    )

                    adicionados_rodada += adicionados

                if adicionados_rodada == 0:
                    print('RESGATE GLOBAL: nenhuma nova evidência aprovada nesta rodada.')
                    break

                combinacoes_contingencia = _reconstruir_combinacoes_contingencia()
                solucao_contingencia = {}
                melhor_solucao_contingencia = None

                _buscar_contingencia(0, set(), [])
                if melhor_solucao_contingencia is not None:
                    print('🟢 RESGATE GLOBAL CONSEGUIU FORMAR 15 FRAGMENTOS ÚNICOS')
                    break

            if melhor_solucao_contingencia is None:
                print('🔴 SELEÇÃO BLOQUEADA: nem a alocação global nem o resgate de cobertura encontraram 15 fragmentos únicos.')
                for numero_bloco in ordem_blocos:
                    chave_bloco = f'bloco_{numero_bloco}'
                    print('  ', chave_bloco, '| combinações finais:', len(combinacoes_contingencia.get(chave_bloco, [])))
                print('MOTIVO CLASSIFICADO: INSUFICIÊNCIA REAL DE FRAGMENTOS ÚNICOS APÓS RESGATE')
                return resultado_vazio

        solucao_global = melhor_solucao_contingencia[1]
        print('🟢 ALOCAÇÃO GLOBAL DE CONTINGÊNCIA ENCONTRADA')
        for numero_bloco in range(1, 6):
            chave_bloco = f'bloco_{numero_bloco}'
            grupo = solucao_global.get(chave_bloco, [])
            print(
                chave_bloco,
                '| 3 fragmentos únicos | narrativa:',
                min([x.get('nota_combinacao_narrativa', 0) for x in grupo], default=0)
            )

    # ------------------------------------------------------------
    # v71 — CONGELAR E VALIDAR A SOLUÇÃO GLOBAL ANTES DA CONSUMIÇÃO
    # ------------------------------------------------------------
    conjuntos_por_bloco = {}

    for numero_bloco in range(1, 6):
        chave_bloco = f'bloco_{numero_bloco}'
        grupo = solucao_global.get(chave_bloco)

        if not isinstance(grupo, list) or len(grupo) != 3:
            print(
                '❌ INTEGRIDADE DA SOLUÇÃO GLOBAL:',
                chave_bloco,
                '| esperado: 3 | encontrado:',
                len(grupo) if isinstance(grupo, list) else 'ausente'
            )
            print('❌ OLLAMA BLOQUEADO: a alocação global não possui os 5 blocos completos.')
            return resultado_vazio

        conjuntos_por_bloco[chave_bloco] = list(grupo)

    hashes_solucao = []
    for chave_bloco, grupo in conjuntos_por_bloco.items():
        for escolhido in grupo:
            h = str(
                escolhido.get('hash')
                or gerar_hash_trecho(escolhido.get('texto', ''))
                or ''
            ).strip()
            if not h:
                print('❌ INTEGRIDADE DA SOLUÇÃO GLOBAL:', chave_bloco, '| fragmento sem hash.')
                return resultado_vazio
            hashes_solucao.append(h)

    if len(hashes_solucao) != 15 or len(set(hashes_solucao)) != 15:
        print('❌ INTEGRIDADE DA SOLUÇÃO GLOBAL: hashes dos 15 fragmentos não são únicos.')
        print('HASHES:', len(hashes_solucao), '| ÚNICOS:', len(set(hashes_solucao)))
        return resultado_vazio

    for numero_bloco in range(1, 6):
        chave_bloco = f'bloco_{numero_bloco}'
        for escolhido in conjuntos_por_bloco[chave_bloco]:
            escolhido['bloco_mead'] = chave_bloco
            fragmentos_selecionados.append(escolhido)
            hashes_selecionados.add(escolhido.get('hash'))
            fonte = escolhido.get('fonte')
            fontes_utilizadas[fonte] = fontes_utilizadas.get(fonte, 0) + 1

    print('CRITÉRIO v35: ranking prioriza trechos desenvolvidos, equilíbrio do bloco e diversidade semântica pré-Ollama')
    print('REGRA: nenhuma faixa individual é obrigatória; tamanho individual é apenas preferência de ranking')
    print('REGRA ESTRUTURAL: sem limite de palavras')

    print('ALLOCAÇÃO CAPACIDADE-AWARE 5 x 3')
    print('CANDIDATOS TOTAIS:', len(candidatos))
    for numero_bloco in range(1, 6):
        chave_bloco = f'bloco_{numero_bloco}'
        total = sum(_palavras_fragmento(x) for x in conjuntos_por_bloco[chave_bloco])
        print(chave_bloco, ': 3 |', total, 'palavras autorizadas')
    print('TOTAL:', len(fragmentos_selecionados), '/ 15')

    # Reserva em memória, sem alterar a seleção principal.
    # Uma reserva por bloco; só será consumida se houver
    # uma re-seleção necessária após rejeição real.
    # --------------------------------------------------------
    # RESERVAS EM MEMÓRIA
    # Seleciona 2 candidatos adicionais por bloco.
    # --------------------------------------------------------

    for numero_bloco in range(1, 6):

        chave_bloco = f"bloco_{numero_bloco}"
        reservas_bloco = []
        hashes_bloqueio_reserva = set(hashes_selecionados)

        for _ in range(5):

            escolhido_reserva = selecionar_melhor_candidato(
                candidatos_por_bloco.get(
                    chave_bloco,
                    []
                ),
                chave_bloco,
                hashes_bloqueio_reserva
            )

            if escolhido_reserva is None:
                break

            escolhido_reserva = dict(
                escolhido_reserva
            )
            escolhido_reserva["bloco_mead"] = chave_bloco

            hash_reserva = (
                escolhido_reserva.get("hash")
                or gerar_hash_trecho(
                    escolhido_reserva.get("texto", "")
                )
            )
            escolhido_reserva["hash"] = hash_reserva

            if (
                not hash_reserva
                or hash_reserva in hashes_bloqueio_reserva
            ):
                break

            escolhido_reserva["_reserva_memoria"] = True
            reservas_bloco.append(
                escolhido_reserva
            )
            hashes_bloqueio_reserva.add(
                hash_reserva
            )

        fragmentos_reserva_por_bloco[
            chave_bloco
        ] = reservas_bloco

    total_reservas_memoria = sum(
        len(x)
        for x in fragmentos_reserva_por_bloco.values()
    )

    print(
        "RESERVAS EM MEMÓRIA:",
        total_reservas_memoria,
        "/ 25"
    )

    print("ALLOCAÇÃO GLOBAL 5 x 3 — v60")
    print("CANDIDATOS TOTAIS:", len(candidatos))

    for numero_bloco in range(1, 6):
        chave_bloco = f"bloco_{numero_bloco}"
        print(
            chave_bloco,
            ":",
            sum(
                1
                for x in fragmentos_selecionados
                if x.get("bloco_mead") == chave_bloco
            )
        )

    print("TOTAL:", len(fragmentos_selecionados), "/ 15")

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

    if len(
        fragmentos_selecionados
    ) < 15:

        print()
        print("=" * 60)
        print("❌ PRÉ-VALIDAÇÃO DA IA FALHOU")
        print("=" * 60)
        print(
            "A seleção produziu apenas",
            len(fragmentos_selecionados),
            "fragmentos de 15 necessários."
        )
        print(
            "O Ollama NÃO será chamado."
        )
        print(
            "Isso evita descobrir o problema somente no bloco "
            "incompleto e evita chamadas parciais."
        )

        # Diagnóstico por fonte.
        fontes_diagnostico = {}
        for candidato in candidatos:
            fonte_id = candidato.get("fonte", "")
            fontes_diagnostico[fonte_id] = (
                fontes_diagnostico.get(fonte_id, 0) + 1
            )

        print("CANDIDATOS VÁLIDOS POR FONTE:")
        for fonte_id, quantidade in sorted(
            fontes_diagnostico.items(),
            key=lambda item: item[0]
        ):
            print(
                "  FONTE",
                fonte_id,
                ":",
                quantidade
            )

        print(
            "CANDIDATOS TOTAIS:",
            len(candidatos)
        )
        print(
            "FRAGMENTOS NECESSÁRIOS:",
            15
        )
        print(
            "FRAGMENTOS DISPONÍVEIS:",
            len(fragmentos_selecionados)
        )
        print("=" * 60)

        return resultado_vazio


    # ========================================================
    # 08.5 — GATE FINAL DOS 15 TRECHOS
    # ========================================================
    # Nenhum fragmento segue para o Ollama sem passar novamente pela
    # auditoria. Se algum candidato ainda falhar, ele é substituído
    # por uma reserva/candidato do MESMO bloco. Se não houver substituto
    # limpo, o processamento é interrompido. Não enviamos conjunto parcial.
    # ========================================================

    def _substituir_invalidos_finais():
        usados = {
            x.get('hash')
            for x in fragmentos_selecionados
            if isinstance(x, dict) and x.get('hash')
        }

        for pos, atual in enumerate(list(fragmentos_selecionados)):
            if not isinstance(atual, dict):
                return False

            ok_atual, motivo_atual = auditar_fragmento_final_python(
                atual.get('texto', ''), atual.get('identidade_fonte', {})
            )
            if ok_atual:
                continue

            bloco_atual = atual.get('bloco_mead', '')
            candidatos_tentativa = []
            candidatos_tentativa.extend(
                fragmentos_reserva_por_bloco.get(bloco_atual, [])
            )
            candidatos_tentativa.extend(
                candidatos_por_bloco.get(bloco_atual, [])
            )

            substituto = None
            for item in candidatos_tentativa:
                cand = item.get('candidato', item) if isinstance(item, dict) else None
                if not isinstance(cand, dict):
                    continue
                cand = dict(cand)
                h = cand.get('hash') or gerar_hash_trecho(cand.get('texto', ''))
                cand['hash'] = h
                if not h or h in usados:
                    continue
                ok, motivo = auditar_fragmento_final_python(
                    cand.get('texto', ''), cand.get('identidade_fonte', {})
                )
                if not ok:
                    cand['motivo_rejeicao_python_final'] = motivo
                    continue
                cand['bloco_mead'] = bloco_atual
                substituto = cand
                break

            if substituto is None:
                print('❌ GATE FINAL PYTHON:', atual.get('id', 'SEM_ID'), motivo_atual)
                print('   Não existe substituto limpo para', bloco_atual)
                return False

            antigo_hash = atual.get('hash')
            usados.discard(antigo_hash)
            usados.add(substituto['hash'])
            fragmentos_selecionados[pos] = substituto
            print(
                '🔄 SUBSTITUIÇÃO PYTHON:',
                atual.get('id', 'SEM_ID'),
                '->', substituto.get('id', 'SEM_ID'),
                '| motivo:', motivo_atual
            )

        return True

    if not _substituir_invalidos_finais():
        print('❌ OLLAMA BLOQUEADO: conjunto final ainda contém fragmento inválido.')
        return resultado_vazio

    # Normalização conservadora dos 15 textos: apenas espaços/quebras.
    for indice_final, frag_final in enumerate(fragmentos_selecionados, start=1):
        preparado = preparar_fragmento_para_ollama_python(frag_final)
        if preparado is None:
            print('❌ OLLAMA BLOQUEADO: fragmento sem texto após preparação.', indice_final)
            return resultado_vazio
        fragmentos_selecionados[indice_final - 1] = preparado

    ok_conjunto_final, motivo_conjunto_final = validar_conjunto_final_python(
        fragmentos_selecionados, 15
    )

    print()
    print('==============================================')
    print('GATE FINAL — PYTHON ANTES DO OLLAMA')
    print('==============================================')
    print('STATUS:', 'APROVADO' if ok_conjunto_final else 'BLOQUEADO')
    print('MOTIVO:', motivo_conjunto_final)

    if not ok_conjunto_final:
        print('❌ NENHUM TEXTO SERÁ ENVIADO AO OLLAMA.')
        return resultado_vazio

    for numero, frag in enumerate(fragmentos_selecionados, start=1):
        print(
            f"FRAGMENTO {numero}: {len(str(frag.get('texto','')).split())} palavras | "
            f"ID {frag.get('id','')} | NOTA TRECHO: {frag.get('nota_trecho_otimo', 0)}/100 | "
            f"PORTAO: {frag.get('nota_portao_qualidade', 0)}/100 | PYTHON FINAL: OK"
        )

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
            "fragmentos_autorizados": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        },

        "bloco_2": {
            "id": "bloco_2",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "fragmentos_autorizados": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        },

        "bloco_3": {
            "id": "bloco_3",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "fragmentos_autorizados": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        },

        "bloco_4": {
            "id": "bloco_4",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "fragmentos_autorizados": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        },

        "bloco_5": {
            "id": "bloco_5",
            "hash": "",
            "informacoes_relevantes": [],
            "titulo": "",
            "fragmentos_autorizados": ["", "", ""],
            "paragrafos_ollama": ["", "", ""] 
        }

    }

    for indice, fragmento in enumerate(
        fragmentos_selecionados
    ):

        numero_bloco = (
            indice // 3
        ) + 1

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        # ----------------------------------------------------
        # PRESERVAR O FRAGMENTO COMPLETO
        # ----------------------------------------------------

        dados_fragmento = {

            "id":
                fragmento["id"],

            "hash":
                fragmento["hash"],

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

        fragmentos_autorizados = [
            fragmento.get(
                "texto",
                ""
            )
            for fragmento in fragmentos_bloco
        ]

        # ----------------------------------------------------
        # GARANTIR AS 3 POSIÇÕES DO BLOCO
        # ----------------------------------------------------

        fragmentos_autorizados = (
            fragmentos_autorizados
            + [
                "",
                "",
                ""
            ]
        )[:3]

        blocos[
            chave_bloco
        ][
            "fragmentos_autorizados"
        ] = fragmentos_autorizados


    # ========================================================
    # AUDITORIA DOS TRECHOS SELECIONADOS
    # ========================================================

    print()
    print(
        "=============================="
    )
    print(
        "TRECHOS SELECIONADOS PELO PYTHON"
    )
    print(
        "=============================="
    )

    total_trechos_python = 0

    for numero_bloco in range(1, 6):

        chave_bloco = (
            f"bloco_{numero_bloco}"
        )

        fragmentos_autorizados = (
            blocos[
                chave_bloco
            ].get(
                "fragmentos_autorizados",
                []
            )
        )

        trechos_validos = [
            trecho
            for trecho in fragmentos_autorizados
            if str(trecho)
        ]

        print()
        print(
            chave_bloco,
            ":",
            len(trechos_validos),
            "trechos selecionados"
        )

        for indice, trecho in enumerate(
            fragmentos_autorizados,
            start=1
        ):

            if str(trecho):

                total_trechos_python += 1

                print(
                    f"  TRECHO {indice}: "
                    f"{len(str(trecho).split())} palavras"
                )

    print()
    print(
        "TOTAL DE TRECHOS SELECIONADOS PELO PYTHON:",
        total_trechos_python
    )       

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
    # "fragmentos_autorizados" foi criado acima e contém
    # os 15 fragmentos autorizados selecionados pelo Python.
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
    
        fragmentos_autorizados = (
            dados_bloco.get(
                "fragmentos_autorizados",
                []
            )
        )
    
        if not isinstance(
            fragmentos_autorizados,
            list
        ):
    
            fragmentos_autorizados = []
    
        fragmentos_autorizados = [
    
            str(paragrafo or "").strip()
    
            for paragrafo
            in fragmentos_autorizados
    
            if str(paragrafo or "").strip()
    
        ]
    
        # ----------------------------------------------------
        # PRESERVAR ESTRUTURA COMPLETA
        # ----------------------------------------------------
    
        blocos_informacoes[
            chave_bloco
        ] = {
    
            "id":
                dados_bloco.get(
                    "id",
                    chave_bloco
                ),
    
            "hash":
                dados_bloco.get(
                    "hash",
                    ""
                ),
    
            "informacoes_relevantes":
                informacoes_relevantes_bloco,
    
            "titulo":
                dados_bloco.get(
                    "titulo",
                    ""
                ),
    
            "fragmentos_autorizados":
                fragmentos_autorizados,
    
            "paragrafos_ollama":
                dados_bloco.get(
                    "paragrafos_ollama",
                    ["", "", ""]
                )
     
        }
    
    
    # --------------------------------------------------------
    # DEBUG — CONFIRMAR TRANSPORTE DOS 15 PARÁGRAFOS
    # --------------------------------------------------------
    
    print()
    print("==============================")
    print("TRANSFERÊNCIA DOS PARÁGRAFOS-BASE")
    print("==============================")
    
    total_fragmentos_autorizados_transferidos = 0
    
    for numero_bloco in range(
        1,
        6
    ):
    
        chave_bloco = (
            f"bloco_{numero_bloco}"
        )
    
        fragmentos_autorizados = (
            blocos_informacoes[
                chave_bloco
            ].get(
                "fragmentos_autorizados",
                []
            )
        )
    
        quantidade = len(
            [
                p
                for p in fragmentos_autorizados
                if str(p).strip()
            ]
        )
    
        total_fragmentos_autorizados_transferidos += (
            quantidade
        )
    
        print(
            chave_bloco,
            ":",
            quantidade,
            "fragmentos autorizados Python"
        )
    
    print()
    print(
        "TOTAL FRAGMENTOS AUTORIZADOS PYTHON TRANSFERIDOS:",
        total_fragmentos_autorizados_transferidos
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

        # Defesa final: nenhum caminho alternativo pode colocar um
        # fragmento sem o PORTÃO DE QUALIDADE dentro de
        # informacoes_relevantes. Em condições normais este teste já
        # passou na seleção; aqui ele funciona como trava de segurança.
        ok_qualidade_final, nota_qualidade_final, motivo_qualidade_final = portao_qualidade_fragmento(
            fragmento.get("texto", ""),
            fragmento.get("identidade_fonte", {}),
            tema
        )

        if not ok_qualidade_final:
            print(
                "[PORTAO DE QUALIDADE] BLOQUEADO ANTES DE informacoes_relevantes:",
                fragmento.get("id", "SEM_ID"),
                motivo_qualidade_final
            )
            continue

        fragmento["nota_portao_qualidade"] = nota_qualidade_final
        fragmento["motivo_portao_qualidade"] = motivo_qualidade_final

        informacoes_relevantes.append({

            "id":
                fragmento["id"],

            "hash":
                fragmento["hash"],

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
            blocos_informacoes,

        # Estado transitório usado somente em memória pela etapa Ollama.
        # Não é persistido pelo salvar_banco().
        "_candidatos_por_bloco":
            candidatos_por_bloco,

        "_fragmentos_reserva_por_bloco":
            fragmentos_reserva_por_bloco

    }

    # ========================================================
    # RESUMO DO PORTÃO DE QUALIDADE v9.18
    # ========================================================
    print()
    print("==============================")
    print("PORTÃO DE QUALIDADE v9.18")
    print("==============================")
    print("FRAGMENTOS APROVADOS:", len(informacoes_relevantes))
    print("NOTAS:", [x.get("nota_portao_qualidade") for x in informacoes_relevantes])
    print("==============================")

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
    # 16. DEBUG DOS IDS E HASHES
    # ========================================================

    print()
    print("==============================")
    print("ID + HASH DOS TRECHOS SELECIONADOS")
    print("==============================")

    for indice, fragmento in enumerate(
        fragmentos_selecionados,
        start=1
    ):

        print()
        print(
            f"FRAGMENTO {indice}"
        )

        print(
            "ID:",
            fragmento["id"]
        )

        print(
            "HASH:",
            fragmento["hash"]
        )

        print(
            "PALAVRAS:",
            fragmento["palavras"]
        )

        print(
            "FONTE:",
            fragmento["fonte"]
        )

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

    print()

    for indice, fragmento in enumerate(
        fragmentos_selecionados,
        start=1
    ):

        print(

            f"FRAGMENTO {indice}: "
            f"{fragmento['palavras']} palavras | "
            f"FONTE {fragmento['fonte']} | "
            f"ID {fragmento['id']} | "
            f"HASH {fragmento['hash'][:16]}..."

        )

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

# ============================================================
# AUDITORIA DETERMINÍSTICA DE INTEGRIDADE — v68
# ============================================================
# Esta camada existe ANTES de qualquer pontuação. Ela não tenta
# "salvar" um fragmento contaminado e não depende de ENTIDADES_PROIBIDAS.
# Se o texto contém resíduos estruturais, linguagem institucional,
# corrupção evidente ou uma entidade de terceiro em construção factual,
# o fragmento é rejeitado.
# ============================================================

def _auditoria_integridade_deterministica(texto, identidade_fonte=None):
    t = str(texto or '').strip()
    if not t:
        return False, 'VAZIO'

    n = normalizar_assunto_texto(t)

    # 1. Corrupção textual evidente / mistura de idioma.
    padroes_corrompidos = (
        r'\bcon\s+(?:a|o|as|os|um|uma|uns|umas|todo|toda|todos|todas)\b',
        r'\breales\b',
        r'\b(?:ofrecemos|ofrece|nuestro|nuestra|nuestros|nuestras)\b',
        # Palavra partida: "possí vel", "eficiê ncia" etc.
        r'\b[a-z]{4,}[áàâãéêíóôõúç]\s+[a-z]{2,6}\b',
        r'[\uf000-\uf8ff]',
    )
    for pat in padroes_corrompidos:
        if re.search(pat, t, re.I):
            return False, 'TEXTO_CORROMPIDO'

    # 2. Navegação, site, CTA e metatexto editorial.
    lixo_interface = (
        r'\bem\s+nosso\s+site\b', r'\bem\s+nosso\s+website\b',
        r'\bno\s+nosso\s+site\b', r'\bno\s+site\b',
        r'\bneste\s+post\b', r'\bneste\s+artigo\b',
        r'\bvamos\s+abordar\b', r'\babordaremos\b',
        r'\bsaiba\s+mais\b', r'\bclique\s+aqui\b',
        r'\bconfira\b', r'\bentre\s+em\s+contato\b',
        r'\bfale\s+conosco\b', r'\bsolicite\s+(?:um\s+)?orcamento\b',
        r'\bnossa\s+empresa\b', r'\bnossos\s+produtos\b',
        r'\bnossos\s+servicos\b', r'\bpodemos\s+oferecer\b',
        r'\boferecemos\b', r'\bfornecemos\b', r'\bvoce\s+pode\b',
    )
    for pat in lixo_interface:
        if re.search(pat, n, re.I):
            return False, 'NAVEGACAO_SITE_INSTITUCIONAL'

    # 3. Perguntas/títulos: um fragmento factual entregue ao Ollama deve
    # ser prosa declarativa, não uma pergunta herdada da página.
    if '?' in t:
        return False, 'PERGUNTA_HERDADA'
    if re.search(r'^\s*(?:resumo\s+tecnico|introducao|conclusao|principais\s+componentes|componentes\s+principais|caracteristicas\s+tecnicas|especificacoes)\s*$', n, re.I):
        return False, 'TITULO_ISOLADO'
    if re.search(r'\b(?:resumo\s+tecnico|introducao|conclusao|principais\s+componentes|componentes\s+principais|caracteristicas\s+tecnicas)\s*[:\-]', n, re.I):
        return False, 'ROTULO_ESTRUTURAL'

    # Lista numerada mesmo sem "." ou ")": "2 Bombas...".
    if re.search(r'(?:^|[.!?]\s+)\d{1,2}\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][^.!?]{2,100}(?:\.|$)', t):
        return False, 'ITEM_DE_LISTA_NUMERADA'

    # Cabeçalho embutido após ponto/quebra: "Resumo técnico ...".
    if re.search(r'(?:^|[.!?]\s+)(?:resumo\s+tecnico|principais\s+componentes|componentes\s+principais|introducao|conclusao)\b', n, re.I):
        return False, 'CABECALHO_EMBUTIDO'

    # 4. Promoção/opinião — somente padrões promocionais inequívocos.
    # Adjetivos técnicos isolados (confiável, eficiente, qualidade etc.)
    # não são mais motivo suficiente para rejeição.
    promocoes = (
        r'\bo\s+melhor\s+(?:equipamento|produto|solucao|opcao)\b',
        r'\ba\s+melhor\s+(?:equipamento|produto|solucao|opcao)\b',
        r'\b(?:equipamento|produto|solucao|opcao)\s+perfeit[oa]\b',
        r'\bsolucao\s+(?:definitiva|perfeita|ideal)\b',
        r'\bmelhor\s+preco\b',
        r'\bpreco\s+competitivo\b',
        r'\b(?:garante|garantindo)\s+(?:o|a|um|uma)\s+(?:resultado|desempenho|eficiencia|qualidade|seguranca)\b',
        r'\b(?:escolha|opcao)\s+(?:inteligente|perfeita|ideal)\b',
    )
    for pat in promocoes:
        if re.search(pat, n, re.I):
            return False, 'LINGUAGEM_PROMOCIONAL'

    # 5. Histórico institucional de terceiro / fabricante. Não depende
    # de uma lista externa de entidades: o próprio padrão sujeito + verbo
    # revela a entidade em contexto empresarial.
    verbos_empresa = r'(?:começou|comecou|fundou|fabricou|fabrica|produz|produziu|desenvolveu|desenvolve|atua|atuou|oferece|ofereceu|fornece|forneceu|iniciou|iniciava|passou\s+a\s+fabricar)'
    padroes_terceiro = (
        rf'\b([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.\-]+(?:\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.\-]+){{0,3}})\s+{verbos_empresa}\b',
        rf'\b(?:a|o)\s+([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.\-]+(?:\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.\-]+){{0,3}})\s+{verbos_empresa}\b',
    )
    nome_fonte = ''
    if isinstance(identidade_fonte, dict):
        nome_fonte = normalizar_assunto_texto(identidade_fonte.get('nome', '')).strip()
    tecnicas = {'bomba', 'bombas', 'centrifuga', 'centrifugas', 'industria', 'industrias', 'empresa', 'fabricante', 'sistema', 'sistemas'}
    for pat in padroes_terceiro:
        m = re.search(pat, t)
        if not m:
            continue
        entidade = normalizar_assunto_texto(m.group(1)).strip()
        palavras_ent = set(entidade.split())
        if entidade == nome_fonte:
            return False, 'EMPRESA_DA_FONTE'
        if entidade and not palavras_ent.issubset(tecnicas):
            return False, 'ENTIDADE_TERCEIRO_DETECTADA'

    # 6. Nome empresarial conhecido da fonte continua sendo proibido.
    if nome_fonte and len(nome_fonte) >= 4 and nome_fonte not in {'fonte','site','pagina','documento','arquivo','pdf','html'}:
        if re.search(r'(?<![a-z0-9])' + re.escape(nome_fonte) + r'(?![a-z0-9])', n, re.I):
            return False, 'EMPRESA_DA_FONTE'

    return True, 'OK'


def diagnosticar_contaminacao_editorial(texto, identidade_fonte=None):
    """Retorna (ok, motivo) para bloquear lixo antes da pontuação."""
    t = str(texto or "").strip()
    n = normalizar_assunto_texto(t)
    if not t:
        return False, "VAZIO"

    ok_integridade, motivo_integridade = _auditoria_integridade_deterministica(
        t, identidade_fonte
    )
    if not ok_integridade:
        return False, motivo_integridade

    # Identidade conhecida da fonte: nunca permitir o nome da empresa/marca.
    if isinstance(identidade_fonte, dict):
        nome = normalizar_assunto_texto(identidade_fonte.get("nome", "")).strip()
        if nome and len(nome) >= 4 and nome not in {"fonte","site","pagina","documento","arquivo","pdf","html"}:
            if re.search(r"(?<![a-z0-9])" + re.escape(nome) + r"(?![a-z0-9])", n):
                return False, "EMPRESA_DA_FONTE"

    fortes_institucionais = [
        r"\bsede\s+(?:pr[oó]pria|localizada|em)\b",
        r"\bparque\s+industrial\b",
        r"\bl[ií]der(?:es)?\s+(?:mundial|global|nacional|do\s+mercado)\b",
        r"\bmais\s+de\s+\d{1,3}\s+anos\s+de\s+experi[eê]ncia\b",
        r"\b(?:anos|d[eê]cadas)\s+de\s+experi[eê]ncia\b",
        r"\bnossos\s+(?:servi[cç]os|produtos|equipamentos|clientes|projetos|especialistas|profissionais)\b",
        r"\bnossa\s+(?:empresa|expertise|experi[eê]ncia|equipe|estrutura)\b",
        r"\boferecemos\s+(?:solu[cç][oõ]es|servi[cç]os|produtos|suporte)\b",
        r"\b(?:atuamos|atu[aã]|trabalhamos)\s+(?:no|na|em|com)\s+(?:mercado|segmento|setor)\b",
        r"\bfornecedores?\s+(?:mais\s+)?renomad[oa]s?\b",
        r"\b(?:clientes|exportador|exportadora)\s+(?:da|do|de|para)\b",
    ]
    for pat in fortes_institucionais:
        if re.search(pat, n, re.I):
            return False, "INSTITUCIONAL_EMPRESARIAL"

    # O texto normalizado perde acentos; manter também as formas
    # sem acento evita que expressões institucionais atravessem o gate.
    fortes_institucionais_normalizados = [
        r"\bsede\s+(?:propria|localizada|em)\b",
        r"\bparque\s+industrial\b",
        r"\blider(?:es)?\s+(?:mundial|global|nacional|do\s+mercado)\b",
        r"\bmais\s+de\s+\d{1,3}\s+anos\s+de\s+experiencia\b",
        r"\b(?:anos|decadas)\s+de\s+experiencia\b",
        r"\bnossos\s+(?:servicos|produtos|equipamentos|clientes|projetos|especialistas|profissionais)\b",
        r"\bnossa\s+(?:empresa|expertise|experiencia|equipe|estrutura)\b",
        r"\boferecemos\s+(?:solucoes|servicos|produtos|suporte)\b",
        r"\b(?:atuamos|atua|trabalhamos)\s+(?:no|na|em|com)\s+(?:mercado|segmento|setor)\b",
        r"\bfornecedores?\s+(?:mais\s+)?renomad[oa]s?\b",
    ]
    for pat in fortes_institucionais_normalizados:
        if re.search(pat, n, re.I):
            return False, "INSTITUCIONAL_EMPRESARIAL"

    # Marca/empresa explícita em construções típicas.
    if re.search(r"\b(?:da|do|pela|pelo|empresa)\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]{2,}\b", t):
        # Não bloquear padrões técnicos comuns: NBR/ISO/ANSI/API etc.
        achados = re.findall(r"\b(?:da|do|pela|pelo|empresa)\s+([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]{2,})\b", t)
        permitidas = {"nbr","iso","ansi","api","astm","din","abnt"}
        if any(a.casefold() not in permitidas for a in achados):
            return False, "NOME_EMPRESA_MARCA"

    # Marca em CamelCase (ex.: FlameGuard, PowerFlow) é bloqueada.
    # Padrões técnicos convencionais não usam essa forma.
    if re.search(r"\b[A-Z][a-zÀ-ÿ]{2,}[A-Z][A-Za-zÀ-ÿ0-9]{1,}\b", t):
        return False, "MARCA_CAMELCASE"

    # Linguagem promocional/comercial contextual.
    comerciais_fortes = [
        r"\b(?:excelente|perfeita|ideal)\s+(?:escolha|op[cç][aã]o)\b",
        r"\bescolha\s+(?:superior|ideal|popular)\b",
        r"\bsolu[cç][aã]o\s+(?:ideal|perfeita|definitiva)\b",
        r"\b(?:nossa|nossos|nosso|nossas)\s+(?:expertise|solu[cç][aã]o|produto|servi[cç]o)\b",
        r"\butilizamos\s+nossa\s+expertise\b",
        r"\b(?:consulte|solicite|pe[cç]a)\s+(?:um|uma|o|a)\s+(?:or[cç]amento|cota[cç][aã]o)\b",
        r"\b(?:entre\s+em\s+contato|fale\s+conosco|compre|adquira)\b",
        r"\b(?:oferece|oferecemos)\s+solu[cç][oõ]es\s+(?:completas|ideais|definitivas)\b",
    ]
    for pat in comerciais_fortes:
        if re.search(pat, n, re.I):
            return False, "LINGUAGEM_COMERCIAL"

    # Produto/modelo/código.
    codigo_patterns = [
        r"\b(?:modelo|s[eé]rie|serie|c[oó]digo|ref\.?|sku|mpn|part\s*number)\s+[A-Z0-9][A-Za-z0-9._/-]{2,}\b",
        r"\b[A-Z]{2,8}-\d{2,}[A-Z0-9-]*\b",
        r"\b[A-Z]{2,8}\d{2,}[A-Z0-9-]*\b",
    ]
    for pat in codigo_patterns:
        if re.search(pat, t, re.I):
            return False, "CODIGO_MODELO_PRODUTO"

    # Metatexto do pipeline/tutorial. Um fragmento técnico nunca pode
    # carregar instruções sobre o próprio processo de geração.
    meta_pipeline = [
        r"\bo\s+python\s+(?:ja|já)\s+(?:pesquisou|selecionou|filtrou|preparou)\b",
        r"\bpython\s+(?:ja|já)\s+(?:pesquisou|selecionou|filtrou|preparou)\b",
        r"\bagora\s+que\s+voce\s+(?:ja\s+)?sabe\b",
        r"\bagora\s+que\s+você\s+(?:já\s+)?sabe\b",
        r"\btrecho\s+(?:ja|já)\s+(?:foi|esta|está)\s+(?:pesquisado|selecionado|autorizado)\b",
    ]
    for pat in meta_pipeline:
        if re.search(pat, n, re.I):
            return False, "META_DO_SISTEMA_OU_TUTORIAL"

    # Títulos/legendas no início do recorte não são prosa técnica.
    if re.search(
        r"^\s*(?:o\s+que\s+(?:e|sao)|como\s+escolher|como\s+funciona|onde\s+(?:aplicar|usar)|"
        r"figura\s*\d*|fig\.?\s*\d*|tabela\s*\d*|quadro\s*\d*|imagem\s*\d*|foto\s*\d*)\s*[:.-]",
        n,
        re.I
    ):
        return False, "TITULO_OU_ROTULO_EMBUTIDO"

    # Cabeçalhos/listas/menus no corpo do fragmento.
    estrutura = [
        r"\b(?:apresenta[cç][aã]o|vantagens|benef[ií]cios|caracter[ií]sticas|especifica[cç][oõ]es|aplica[cç][oõ]es)\s*:\s*",
        r"(?:^|\s)\d+[.)]\s+",
        r"\b(?:menu|home|login|skip to|top of page|fale conosco|clique aqui|saiba mais)\b",
    ]
    for pat in estrutura:
        if re.search(pat, n, re.I):
            return False, "ESTRUTURAL_NAVEGACAO_LISTA"

    # Títulos e rótulos incorporados no meio do texto. A regra é
    # deliberadamente conservadora: se o trecho mistura conteúdo técnico
    # com estrutura editorial da página, rejeitamos o fragmento inteiro.
    titulos_internos = [
        r"(?:^|[.!?]\s+)(?:como escolher|onde aplicar|onde usar|"
        r"rotina pratica de manutencao|rotina de manutencao preventiva|"
        r"vantagens tecnicas|eficiencia e desempenho|design monobloco|"
        r"caracteristicas tecnicas|aplicacoes e caracteristicas|"
        r"problemas frequentes|quando escolher|tipos de bombas)\b",
        r"\bfigura\s*\d{1,3}\s*[:.-]",
        r"\b(?:tabela|quadro|imagem|foto)\s*\d{1,3}\s*[:.-]",
        r"(?:^|[.!?]\s+)importante\s*:",
    ]
    for pat in titulos_internos:
        if re.search(pat, n, re.I):
            return False, "TITULO_OU_ROTULO_EMBUTIDO"

    # OCR/PDF claramente corrompido. Exemplos como "Eciência" não devem
    # chegar ao Ollama para que o modelo tente adivinhar a palavra original.
    if re.search(r"[\uf000-\uf8ff]", t):
        return False, "OCR_PDF_CORROMPIDO"

    # Linguagem promocional/opinativa que não agrega fato técnico verificável.
    opinativos_promocionais = [
        r"\bextremamente\s+vers[aá]teis\b",
        r"\bperfeitas?\s+para\b",
        r"\bescolha\s+inteligente\b",
        r"\bfinanceiramente\s+vi[aá]vel\b",
        r"\bretorno\s+sobre\s+o\s+investimento\b",
        r"\bindispens[aá]vel\b",
        r"\bnumerosas?\s+e\s+significativas?\b",
        r"\bdesempenho\s+superior\b",
        r"\balta\s+efici[eê]ncia\s+e\s+desempenho\s+confi[aá]vel\b",
        r"\bsolu[cç][aã]o\s+econ[oô]mica\s+e\s+eficiente\b",
    ]
    for pat in opinativos_promocionais:
        if re.search(pat, n, re.I):
            return False, "LINGUAGEM_OPINATIVA_PROMOCIONAL"

    # Idioma estrangeiro comercial evidente.
    estrangeiro = [
        "presentación", "presentacion", "ofrecemos soluciones", "nuestro equipo",
        "en este artículo", "en este articulo", "presupuesto gratuito"
    ]
    if any(x in n for x in estrangeiro):
        return False, "IDIOMA_ESTRANGEIRO"

    # Espanhol técnico residual: exige pelo menos dois marcadores para
    # evitar falso positivo em uma palavra isolada.
    marcadores_espanhol_tecnico = (
        r"\b(?:los|las|una|unos|unas|principales|incluyen|correctamente|"
        r"instalada|permanece|fugas|vibraciones|mantenimiento|desgaste|"
        r"alineacion|desalineacion|rodamientos|holgura|rendimiento|"
        r"problemas|vida util)\b"
    )
    if len(re.findall(marcadores_espanhol_tecnico, n, re.I)) >= 2:
        return False, "IDIOMA_ESTRANGEIRO"

    # Resíduos de título/categoria, conversa de interface e corrupção
    # linguística que precisam ser barrados antes da seleção chegar ao
    # Ollama. O filtro anterior aceitava esses casos porque procurava
    # principalmente padrões comerciais; isso deixou passar trechos como
    # "Então vamos lá pessoal!" e misturas "con reales".
    padroes_conversa_e_interface = (
        r"^\s*(?:ent[aã]o\s+)?vamos\s+l[aá]\s+pessoal\b",
        r"^\s*ent[aã]o\s+vamos\b",
        r"^\s*mas\s+afinal[,!]\s+",
        r"^\s*o\s+que\s+(?:[eé]|s[aã]o)\b",
        r"^\s*como\s+funciona\b[^.!?]{0,120}(?:\?|$)",
        r"\b(?:clique|confira|veja|saiba)\s+(?:aqui|abaixo)\b",
    )
    for padrao in padroes_conversa_e_interface:
        if re.search(padrao, t, re.IGNORECASE):
            return False, "CONVERSA_OU_PERGUNTA_DE_INTERFACE"

    # Resíduos de espanhol/corrupção de coleta. Uma única ocorrência de
    # "reales" ou da construção "con + artigo português" já é suficiente:
    # em um fragmento técnico em português isso indica mistura de fontes,
    # tradução quebrada ou OCR/extração defeituosa.
    padroes_idioma_corrompido = (
        r"\breales\b",
        r"\bcon\s+(?:a|o|as|os|um|uma|uns|umas)\b",
        r"\bcon\s+reais\b",
        r"\bresultados\s+con\s+reales\b",
        r"\b(?:reales|ofrecemos|nuestro|nuestra|nuestros|nuestras)\b",
    )
    for padrao in padroes_idioma_corrompido:
        if re.search(padrao, n, re.IGNORECASE):
            return False, "IDIOMA_OU_TEXTO_CORROMPIDO"

    residuos_normalizados = (
        r"\bpecas\s+para\s+[a-z0-9 ]{2,60}$",
        r"\bprodutos\s+relacionados\b",
        r"\bperguntas\s+frequentes\b",
        r"\bfaq\b",
        r"\bcomo funciona\b[^.!?]{0,90}\?$",
        r"\bo que (?:e|sao)\b[^.!?]{0,90}\?$",
        r"\bmelhores?\s+bombas?\b",
        r"\botimas?\s+opcoes?\b",
        r"\bconfira\s+nossa\s+sele[cç]ao\b",
        r"\bpromocao\b",
        r"\bem\s+promocao\b",
        r"\ba\s+empresa\s+oferece\b",
        r"\bservicos\s+personalizados\b",
        r"\bqualidade\s+dos\s+produtos\b",
        r"\bganhou\s+reconhecimento\b"
    )
    for padrao in residuos_normalizados:
        if re.search(padrao, n, re.I):
            if "como funciona" in padrao or "o que" in padrao:
                return False, "TITULO_OU_PERGUNTA_HERDADA"
            if "pecas" in padrao or "produtos" in padrao or "faq" in padrao or "perguntas" in padrao:
                return False, "RESIDUO_DE_CATEGORIA"
            return False, "LINGUAGEM_COMERCIAL_OU_INSTITUCIONAL"

    # Verbos que normalmente exigem complemento: quando o recorte termina
    # exatamente neles, há forte indício de texto cortado.
    if re.search(
        r"\b(?:precisam|deve|devem|pode|podem|permite|permitem|inclui|incluem|"
        r"possui|possuem|depende|dependem|necessita|necessitam)\s*[.!?]$",
        n,
        re.I
    ):
        return False, "FINAL_TRUNCADO"

    # Cabeçalho embutido: linha/título seguido de dois-pontos.
    # Ex.: "Bombas de Deslocamento Positivo: ...".
    m_heading = re.search(r"(?:^|[.!?]\s+)([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][^:]{2,70}):\s", t)
    if m_heading:
        prefixo = m_heading.group(1).strip()
        palavras_prefixo = prefixo.split()
        iniciais_maiusculas = sum(1 for w in palavras_prefixo if w[:1].isupper())
        if 2 <= len(palavras_prefixo) <= 8 and iniciais_maiusculas >= 2:
            return False, "CABECALHO_EMBUTIDO"

    # Nome empresarial em sigla, quando usado como sujeito da ação.
    sigla_tecnica = {"NBR", "ISO", "ANSI", "ASTM", "DIN", "API", "ABNT", "RPM", "PVC", "PEAD", "CPVC", "PPR", "MCA", "KW", "CV", "HP", "DN", "PN", "IP"}
    m_sigla = re.search(r"\b([A-Z]{3,10})\b\s+(?:é|e|possui|tem|oferece|fornece|fabrica|produz|desenvolve|atua|apresenta)\b", t)
    if m_sigla and m_sigla.group(1) not in sigla_tecnica:
        return False, "SIGLA_EMPRESARIAL"

    # Dados de contato e URL.
    if re.search(r"https?://|www\.|[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", t, re.I):
        return False, "CONTATO_URL_EMAIL"

    return True, "OK"


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

    # GATE CENTRAL: a coleta editorial também deve rejeitar
    # contaminação antes de o trecho entrar no banco de candidatos.
    ok_diagnostico, motivo_diagnostico = diagnosticar_contaminacao_editorial(
        texto_original,
        None
    )
    if not ok_diagnostico:
        return False

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

    # Metatexto do próprio pipeline e linguagem de tutorial/editorial.
    # Isso é lixo de redação quando aparece em um fragmento técnico ou
    # no parágrafo produzido; não deve chegar ao Ollama nem ao JSON.
    padroes_meta_sistema = [
        r"\bo\s+python\s+(?:ja|já)\s+(?:pesquisou|selecionou|filtrou|preparou)\b",
        r"\bpython\s+(?:ja|já)\s+(?:pesquisou|selecionou|filtrou|preparou)\b",
        r"\btrecho\s+(?:ja|já)\s+(?:foi|esta|está)\s+(?:pesquisado|selecionado|autorizado)\b",
        r"\bagora\s+que\s+voce\s+(?:ja\s+)?sabe\b",
        r"\bagora\s+que\s+você\s+(?:já\s+)?sabe\b",
        r"\bcomo\s+voce\s+(?:ja\s+)?sabe\b",
        r"\bcomo\s+você\s+(?:já\s+)?sabe\b",
        r"\bneste\s+artigo\b",
        r"\bneste\s+guia\b",
        r"\bao\s+final\s+deste\s+artigo\b",
    ]
    for padrao in padroes_meta_sistema:
        if re.search(padrao, texto_original, re.IGNORECASE):
            return False, 'META_DO_SISTEMA_OU_TUTORIAL'

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
# AUDITORIA FINAL — PYTHON ENTREGA TRECHO PRONTO AO OLLAMA
# ============================================================
#
# Esta camada NÃO escreve conteúdo novo.
# Ela somente impede que um fragmento estruturalmente incompleto,
# truncado ou contaminado chegue à redação.
#
# Objetivo: o Ollama recebe um texto técnico já fechado, coerente
# e sem lixo editorial, e sua função passa a ser apenas humanizar.
# ============================================================

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

    # ====================================================
    # 00. IDENTIDADE DA FONTE / MARCA CONHECIDA
    # ====================================================
    #
    # Se a etapa de coleta já identificou explicitamente
    # uma empresa ou marca da fonte, o nome não pode entrar
    # no fragmento editorial. Isso fecha o principal ponto
    # cego da v8.2: nomes comerciais sem "Ltda.", "S.A." etc.
    #
    # Não usamos o domínio como nome de empresa, pois o próprio
    # domínio pode aparecer somente nos metadados da fonte.
    # ====================================================

    if isinstance(identidade_fonte, dict):

        nome_empresa_fonte = str(
            identidade_fonte.get(
                "nome",
                ""
            )
            or ""
        ).strip()

        nome_empresa_normalizado = normalizar_assunto_texto(
            nome_empresa_fonte
        ).strip()

        # Evita bloquear identificações genéricas da coleta.
        nomes_genericos = {
            "",
            "fonte",
            "site",
            "pagina",
            "página",
            "documento",
            "arquivo",
            "pdf",
            "html"
        }

        if (
            len(nome_empresa_normalizado) >= 4
            and nome_empresa_normalizado not in nomes_genericos
            and re.search(
                r"(?<![a-z0-9])"
                + re.escape(nome_empresa_normalizado)
                + r"(?![a-z0-9])",
                texto_normalizado,
                re.IGNORECASE
            )
        ):
            return False

    # ----------------------------------------------------
    # 00.2. NOME DE PRODUTO / MARCA ASSOCIADO AO TEMA
    #
    # Rejeita construções do tipo "produto BrandX",
    # "sistema FlameGuard" ou "revestimento BrandZ 500"
    # quando o nome possui forma comercial própria.
    # ----------------------------------------------------

    padroes_produto_marca = [
        r"\b(?:produto|sistema|revestimento|material|linha)\s+"
        r"([A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]*"
        r"(?:\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]*){0,2})"
        r"(?:\s+\d{2,6})?\b"
    ]

    for padrao in padroes_produto_marca:

        if re.search(
            padrao,
            texto
        ):
            return False

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

    # ====================================================
    # 05.5. MARCA / EMPRESA SEM SUFIXO JURÍDICO
    # ====================================================
    #
    # Captura nomes comerciais que aparecem como marca, mas
    # não trazem "Ltda.", "S.A.", "EPP" etc.
    #
    # Exemplos bloqueados:
    #   "A FireShield oferece ..."
    #   "O sistema FireShield é ..."
    #   "A ProtecFogo desenvolve ..."
    #
    # O detector é deliberadamente contextual para não bloquear
    # substantivos técnicos comuns.
    # ====================================================

    padrao_marca_contextual = (
        r"\b(?:A|O|As|Os)\s+"
        r"[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]{3,}"
        r"(?:\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÀ-ÿ0-9&.'’\-]{2,})?"
        r"\s+(?:oferece|fornece|produz|fabrica|"
        r"comercializa|distribui|vende|desenvolve|atua|"
        r"representa|disponibiliza)\b"
    )

    if re.search(
        padrao_marca_contextual,
        texto
    ):
        return False

    # Marcas em CamelCase, comuns em nomes comerciais, também
    # são bloqueadas mesmo quando aparecem sem verbo comercial.
    # Siglas técnicas simples (NBR, ISO, ASTM etc.) não entram
    # nesse padrão.
    if re.search(
        r"\b[A-Z][a-z]+[A-Z][A-Za-zÀ-ÿ0-9]+\b",
        texto
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
        "nossos produtos",
        "nossa empresa",
        "sobre nossa empresa",
        "conheca nossa empresa",
        "conheça nossa empresa",
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

def portao_qualidade_fragmento(texto, identidade_fonte=None, tema=None):
    """
    PORTÃO DE QUALIDADE — v9.18 BLINDADO

    Decide se um fragmento está realmente apto a entrar na seleção editorial.
    Não reescreve o conteúdo e não tenta "salvar" um trecho contaminado.

    Retorno:
        (True, nota, "OK")
        (False, nota, "MOTIVO")

    A função é um gate duro: lixo estrutural, editorial, comercial ou
    incompatível com prosa contínua é rejeitado antes da seleção e do Ollama.
    """
    texto = str(texto or '').strip()
    if not texto:
        return False, 0, 'VAZIO'

    normalizado = normalizar_assunto_texto(texto)
    palavras = re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]*\b", texto)
    n_palavras = len(palavras)

    # ============================================================
    # GATE 1 — DENSIDADE GRAMATICAL
    # Protege contra tabelas/listas técnicas coladas em prosa.
    # ============================================================
    stop_words_core = {
        'a', 'o', 'as', 'os', 'um', 'uma', 'uns', 'umas', 'de', 'do', 'da',
        'dos', 'das', 'e', 'ou', 'em', 'no', 'na', 'nos', 'nas', 'por',
        'para', 'com', 'sem', 'que', 'ao', 'aos'
    }
    if n_palavras > 0:
        qtd_stop = sum(1 for p in palavras if p.lower() in stop_words_core)
        razao_stop = qtd_stop / n_palavras
        if razao_stop < 0.10 and n_palavras >= 45:
            return False, 0, f'BAIXA_DENSIDADE_GRAMATICAL_TABELA:{razao_stop:.3f}'

    # ============================================================
    # GATE 2 — CABEÇALHO EM CAIXA ALTA COLADO À PROSA
    # ============================================================
    if re.search(
        r'^[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ\s\d.——–-]{5,60}\s+'
        r'[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][a-záàãâéêíóôõúç]',
        texto
    ):
        return False, 0, 'CABEÇALHO_COLADO_SEM_PONTUACAO'

    # ============================================================
    # GATE 3 — RESÍDUOS DE SEPARADORES DE HTML/LEIAUTE
    # ============================================================
    if " . " in texto or " | " in texto or " · " in texto:
        if texto.count(" . ") + texto.count(" | ") + texto.count(" · ") >= 2:
            return False, 0, 'RESIDUOS_DE_SEPARADORES_HTML'

    # ============================================================
    # 1. LIXO DURO
    # ============================================================
    ok_contaminacao, motivo_contaminacao = diagnosticar_contaminacao_editorial(
        texto, identidade_fonte
    )
    if not ok_contaminacao:
        return False, 0, f'CONTAMINACAO:{motivo_contaminacao}'

    ok_comercial = fragmento_eh_comercialmente_limpo(texto, identidade_fonte)
    if not ok_comercial:
        return False, 0, 'COMERCIAL_IDENTIDADE_CODIGO'

    if re.search(r'https?://|www\.|\b[\w.%+-]+@[\w.-]+\.[A-Za-z]{2,}\b', texto, re.I):
        return False, 0, 'URL_EMAIL'

    if re.search(r'<[^>]{1,80}>|\[/?(?:BLOCO|PARAGRAFO|FRAGMENTO|TEXTO)[^]]*\]', texto, re.I):
        return False, 0, 'ARTEFATO_COLETA'

    # ============================================================
    # 2. TÍTULO / NAVEGAÇÃO / META-TEXTO
    # ============================================================
    padroes_lixo_editorial = [
        r'\b(?:leia\s+tamb[eé]m|saiba\s+mais|clique\s+aqui|confira\s+tamb[eé]m)\b',
        r'\b(?:artigos?|posts?|conte[uú]dos?)\s+relacionados\b',
        r'\b(?:p[aá]gina|page)\s+\d+(?:\s+de\s+\d+)?\b',
        r'\b(?:pr[oó]ximo|anterior|voltar|menu|home|in[ií]cio)\b',
        r'\b(?:acompanhe|neste\s+artigo|neste\s+texto|como\s+veremos\s+a\s+seguir)\b',
        r'\b(?:neste\s+artigo)\s*,?\s+(?:voc[eê]|iremos|vamos)\b',
    ]
    for padrao in padroes_lixo_editorial:
        if re.search(padrao, normalizado, re.I):
            return False, 0, 'META_TEXTO_EDITORIAL'

    primeiras = re.split(r'(?<=[.!?])\s+', texto, maxsplit=1)[0].strip()
    primeira_norm = normalizar_assunto_texto(primeiras)
    if re.match(
        r'^(?:o que (?:e|sao)|como funciona|funcionamento|caracteristicas|aplicacoes|problemas frequentes|diferencas entre|tipos de)\b',
        primeira_norm,
        re.I
    ) and len(primeiras.split()) <= 14:
        return False, 0, 'TITULO_EMBUTIDO'

    if re.search(
        r'(?:^|[.!?]\s+)[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-ZÁÀÃÂÉÊÍÓÔÕÚÇ\s&-]{8,}(?:[.!?]|\n)',
        texto
    ):
        return False, 0, 'TITULO_CAIXA_ALTA'

    # ============================================================
    # 3. IDIOMA / FRAGMENTO ESTRANHO
    # ============================================================
    marcadores_estrangeiros = [
        r'\b(?:adem[aá]s|por\s+lo\s+tanto|sin\s+embargo|para\s+ello|estructura\s+de|caracter[ií]sticas\s+de)\b',
        r'\b(?:the|this|these|therefore|however|according\s+to|features\s+of)\b',
    ]
    ocorrencias_estrangeiras = sum(
        1 for p in marcadores_estrangeiros if re.search(p, normalizado, re.I)
    )
    if ocorrencias_estrangeiras >= 1:
        return False, 0, 'IDIOMA_ESTRANHO'

    # ============================================================
    # 4. ESTRUTURA
    # ============================================================
    if texto.count('?') >= 2:
        return False, 0, 'EXCESSO_PERGUNTAS'

    if re.search(r'(?:^|\s)[•▪◦●○◆◇►▸→]|(?:^|\n)\s*[-*]\s+', texto):
        return False, 0, 'LISTA_EMBUTIDA'

    if re.search(r'\.{4,}|,{3,}|-{4,}|_{3,}', texto):
        return False, 0, 'PONTUACAO_CORROMPIDA'

    if not re.search(r'[.!?]$', texto):
        return False, 0, 'FINAL_ABERTO'

    if re.search(r'[,;:]\s*$', texto):
        return False, 0, 'FINAL_TRUNCADO'

    inicio = normalizado.lstrip(' -–—•·')
    if re.match(r'^(?:e|ou|que|de|da|do|das|dos|para|por|com|quando|onde|sendo|alem de|além de)\s+', inicio):
        return False, 0, 'INICIO_TRUNCADO'

    if re.search(
        r'\b(?:e|ou|que|de|da|do|das|dos|para|por|com|como|quando|onde|sendo|incluindo|conforme|atraves|através)\s*[.!?]$',
        normalizado,
        re.I
    ):
        return False, 0, 'FINAL_COM_CONECTOR'

    # ============================================================
    # 5. QUALIDADE TEXTUAL
    # ============================================================
    frases = [x.strip() for x in re.split(r'(?<=[.!?])\s+', texto) if x.strip()]
    # Uma frase completa pode ser um bom candidato; só rejeitar quando
    # ela for curta demais para carregar contexto técnico suficiente.
    if len(frases) < 2 and len(palavras) < 16:
        return False, 0, 'POUCA_INFORMACAO_EDITORIAL'

    nota = 100
    tamanhos_frases = [
        len(re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]*\b", f)) for f in frases
    ]
    if tamanhos_frases and max(tamanhos_frases) > 55:
        nota -= 18
    elif tamanhos_frases and max(tamanhos_frases) > 42:
        nota -= 8

    curtas = sum(1 for n in tamanhos_frases if n <= 5)
    if curtas >= 3:
        nota -= 12
    elif curtas == 2:
        nota -= 5

    stop = {
        'a','o','as','os','um','uma','uns','umas','de','do','da','dos','das',
        'e','ou','em','no','na','nos','nas','por','para','com','sem','que',
        'como','mais','menos','se','sua','seu','suas','seus','ao','aos','à','às',
        'é','são','ser','tem','têm','pode','podem','também','isso','essa','esse'
    }
    frequencias = {}
    for p in palavras:
        k = normalizar_assunto_texto(p).lower()
        if len(k) >= 4 and k not in stop:
            frequencias[k] = frequencias.get(k, 0) + 1
    if frequencias:
        maior = max(frequencias.values())
        if maior >= 8 and maior / max(n_palavras, 1) > 0.10:
            nota -= 15
        elif maior >= 6 and maior / max(n_palavras, 1) > 0.09:
            nota -= 8

    tokens_validos = [
        normalizar_assunto_texto(p).lower()
        for p in palavras
        if len(normalizar_assunto_texto(p)) >= 4
        and normalizar_assunto_texto(p).lower() not in stop
    ]
    if tokens_validos:
        diversidade = len(set(tokens_validos)) / len(tokens_validos)
        if n_palavras >= 60 and diversidade < 0.42:
            nota -= 12
        elif n_palavras >= 60 and diversidade < 0.50:
            nota -= 6

    tokens_nao_textuais = re.findall(r'[^\s]+', texto)
    suspeitos = sum(
        1 for token in tokens_nao_textuais
        if re.search(r'[_=< >\[\]{}]|\b(?:sku|mpn|pn)\b', token, re.I)
    )
    if suspeitos:
        return False, 0, 'TOKENS_NAO_TEXTUAIS'

    if nota < 50:
        return False, nota, 'QUALIDADE_TEXTUAL_BAIXA'

    return True, nota, 'OK'


def auditar_fragmento_final_python(texto, identidade_fonte=None):
    texto_original = str(texto or '').strip()

    if not texto_original:
        return False, 'VAZIO'

    ok_integridade, motivo_integridade = _auditoria_integridade_deterministica(
        texto_original, identidade_fonte
    )
    if not ok_integridade:
        return False, motivo_integridade

    ok_lixo, motivo_lixo = diagnosticar_contaminacao_editorial(
        texto_original, identidade_fonte
    )
    if not ok_lixo:
        return False, motivo_lixo

    if not fragmento_eh_aproveitavel_editorialmente(texto_original):
        return False, 'FILTRO_EDITORIAL'

    # Nunca mandar HTML, marcadores internos ou artefatos de coleta.
    if re.search(r'<\/?(?:p|div|span|h[1-6]|li|ul|ol|table|tr|td)\b', texto_original, re.I):
        return False, 'HTML_EMBUTIDO'

    if re.search(r'\[/?(?:BLOCO|PARAGRAFO|FRAGMENTO|TEXTO)[^]]*\]', texto_original, re.I):
        return False, 'MARCADOR_INTERNO'

    # Cabeçalho embutido no início do trecho: "Tema: texto...".
    primeira_frase = re.split(r'[.!?]', texto_original, maxsplit=1)[0].strip()
    if ':' in primeira_frase:
        prefixo = primeira_frase.split(':', 1)[0].strip()
        palavras_prefixo = prefixo.split()
        if 1 <= len(palavras_prefixo) <= 8:
            return False, 'CABECALHO_EMBUTIDO_FINAL'

    # Títulos de perguntas frequentemente vêm colados ao primeiro período
    # sem pontuação ou quebra de linha, por exemplo:
    # "O que é uma Bomba Centrífuga Sanitária A Bomba centrífuga..."
    # Não tentar reparar esse corte automaticamente: colocar em quarentena
    # é mais seguro do que adivinhar onde termina o título.
    padroes_titulo_embutido = [
        r'^\s*(?:o\s+que\s+e|o\s+que\s+[ée]|como\s+funciona|como\s+escolher|quais\s+as\s+aplicacoes|quais\s+as\s+aplicações)\b[^.!?]{8,160}\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][a-záàãâéêíóôõúç]+\b',
        r'^\s*\d+\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][^.!?]{5,100}\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][a-záàãâêéíóôõúç]+\b',
        r'^\s*(?:problemas\s+frequentes|tipos\s+de\s+bombas|caracter[ií]sticas(?:\s+t[eé]cnicas)?|aplica[cç][oõ]es(?:\s+e\s+caracter[ií]sticas)?|caracter[ií]sticas\s+e\s+aplica[cç][oõ]es)\b[^.!?]{5,140}\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][a-záàãâêéíóôõúç]+\b'
    ]
    for padrao_titulo in padroes_titulo_embutido:
        if re.search(padrao_titulo, texto_original, re.IGNORECASE):
            return False, 'TITULO_EMBUTIDO_SEM_PONTUACAO'

    # Cabeçalhos numerados ou títulos colados ao corpo também podem
    # aparecer depois de uma frase completa, especialmente em PDFs.
    if re.search(
        r'(?<![.!?])\s+\d+\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÁÀÃÂÉÊÍÓÔÕÚÇ\-]{2,}(?:\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][A-Za-zÁÀÃÂÉÊÍÓÔÕÚÇ\-]{2,}){1,8}\s+[A-ZÁÀÃÂÉÊÍÓÔÕÚÇ][a-záàãâêéíóôõúç]+',
        texto_original
    ):
        return False, 'CABECALHO_NUMERADO_EMBUTIDO'

    # ========================================================
    # FILTRO REFORÇADO — LIXO DE NAVEGAÇÃO / TÍTULOS
    # ========================================================
    # Alguns recortes passam pelo filtro geral porque o texto técnico
    # é bom, mas traz no final ou no início restos de menus, títulos e
    # chamadas de outros artigos. Esses resíduos NÃO devem chegar ao Ollama.

    marcadores_navegacao_fortes = [
        r"\bleia\s+tamb[eé]m\b",
        r"\bveja\s+as\s+principais\s+causas\b",
        r"\bclique\s+aqui\b",
        r"\bsaiba\s+mais\b",
        r"\bconfira\s+tamb[eé]m\b",
        r"\bartigos\s+relacionados\b",
        r"\bposts?\s+relacionados\b",
        r"\bconte[uú]dos?\s+relacionados\b"
    ]

    for padrao in marcadores_navegacao_fortes:
        if re.search(padrao, texto_original, re.IGNORECASE):
            return False, 'NAVEGACAO_EMBUTIDA'

    # Sequência típica de títulos colados no corpo do recorte.
    # Ex.: "O que são ... Bomba centrífuga, funcionamento ..."
    frases_titulo = re.findall(
        r"(?:^|[.!?]\s+)(?:o que são|o que e|como funciona|funcionamento|caracter[ií]sticas|aplica[cç][oõ]es)\b[^.!?]{0,90}",
        texto_original,
        re.IGNORECASE
    )
    if len(frases_titulo) >= 2:
        return False, 'SEQUENCIA_DE_TITULOS'

    # O fragmento deve terminar como uma ideia fechada.
    # Aceitamos apenas pontuação final de frase; um recorte que termina
    # em palavra solta, título ou palavra cortada deve ser descartado.
    if not re.search(r'[.!?]$', texto_original):
        return False, 'FINAL_SEM_PONTUACAO'

    if texto_original[-1] in ',:;/\\':
        return False, 'FINAL_TRUNCADO'

    final_normalizado = normalizar_assunto_texto(texto_original)
    conectores_finais = (
        ' e', ' ou', ' que', ' de', ' da', ' do', ' das', ' dos',
        ' para', ' por', ' com', ' como', ' quando', ' onde',
        ' sendo', ' incluindo', ' conforme', ' devido', ' através'
    )
    if any(final_normalizado.endswith(x) for x in conectores_finais):
        return False, 'FINAL_COM_CONECTOR'

    # Um fragmento com muitas quebras artificiais tende a ser coleta de
    # menu/lista ou recorte de página, não um parágrafo técnico contínuo.
    quebras = len(re.findall(r'\n+', texto_original))
    if quebras >= 4:
        return False, 'MUITAS_QUEBRAS'

    # Evita trechos que começam no meio de uma construção sintática.
    inicio = final_normalizado.lstrip(' -–—•·')
    inicios_incompletos = (
        'de ', 'da ', 'do ', 'das ', 'dos ', 'e ', 'ou ', 'que ',
        'como ', 'para ', 'por ', 'com ', 'quando ', 'onde ',
        'sendo ', 'além de ', 'alem de '
    )
    if inicio.startswith(inicios_incompletos):
        return False, 'INICIO_TRUNCADO'

    # Linguagem promocional/opinativa herdada da fonte.
    # Não bloquear qualquer adjetivo técnico isolado: o bloqueio é contextual,
    # quando a expressão apresenta o produto/aplicação como "ideal", "perfeito",
    # "ótimo" ou "excelente" para determinada utilização.
    padroes_promocionais_fragmento = (
        r"\b(?:é|são|será|serão|sendo)\s+ideais?\s+para\b",
        r"\b(?:é|são|será|serão|sendo)\s+perfeitas?\s+para\b",
        r"\b(?:é|são|será|serão|sendo)\s+perfeitos?\s+para\b",
        r"\b(?:é|são|será|serão|sendo)\s+(?:ótimas?|excelentes?)\s+para\b",
        r"\bsolu[cç][aã]o\s+(?:definitiva|perfeita|ideal)\b",
        r"\b(?:o|a)\s+melhor\s+(?:equipamento|produto|solu[cç][aã]o|op[cç][aã]o)\b",
        r"\b(?:melhor\s+pre[cç]o|pre[cç]o\s+competitivo)\b",
        r"\b(?:garante|garantindo)\s+(?:o|a|um|uma)\s+(?:resultado|desempenho|efici[eê]ncia|qualidade|seguran[cç]a)\b",
    )
    for padrao_promocional in padroes_promocionais_fragmento:
        if re.search(padrao_promocional, texto_original, re.IGNORECASE):
            return False, 'LINGUAGEM_PROMOCIONAL_NO_FRAGMENTO'

    # IMPORTANTE: estas expressões também são proibidas na auditoria final
    # do texto gerado pelo Ollama. O filtro aqui evita enviar ao modelo um
    # fragmento que já nasce incompatível com a própria validação final.

    # Pelo menos uma frase completa.
    # Não existe mais piso individual de palavras. A quantidade mínima
    # Não existe piso ou teto de palavras para o fragmento por esta regra.
    palavras = re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]*\b", texto_original)

    frases = [x.strip() for x in re.split(r'(?<=[.!?])\s+', texto_original) if x.strip()]
    if not frases:
        return False, 'SEM_FRASE_COMPLETA'

    # Rejeita fragmentos que terminam em abreviação muito provável de corte.
    if re.search(r'\b(?:etc|ex|aprox|aproximadamente)\.?$', texto_original, re.I):
        if not re.search(r'[.!?]$', texto_original):
            return False, 'FINAL_SUSPEITO'

    # ========================================================
    # FILTRO REFORÇADO — SOBRAS CURTAS NO FINAL
    # ========================================================
    # Em recortes contaminados, o texto técnico costuma terminar com
    # 1–5 palavras que pertencem ao próximo título ou foram cortadas.
    # Não rejeitamos qualquer frase curta: a regra só vale para o final
    # de um fragmento longo.
    frases_finais = [
        x.strip()
        for x in re.split(r'(?<=[.!?])\s+', texto_original)
        if x.strip()
    ]

    if len(palavras) >= 60 and frases_finais:
        ultima = frases_finais[-1]
        palavras_ultima = re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]*\b", ultima)

        if 1 <= len(palavras_ultima) <= 5:
            return False, 'SOBRA_CURTA_NO_FINAL'

    # Construções gramaticalmente abertas no fim indicam corte de fonte.
    if re.search(
        r"\b(?:sem|com|para|por|de|da|do|das|dos|e|ou|que)\s+\w{1,7}[.!?]$",
        texto_original,
        re.IGNORECASE
    ):
        # Mantém palavras curtas legítimas como "sem fim" apenas quando
        # a frase final parece realmente completa; expressões truncadas
        # como "sem compro." continuam sendo rejeitadas.
        final_sem_pontuacao = re.sub(r'[.!?]$', '', texto_original).strip()
        if re.search(r"\bsem\s+(?:compro|compr|efi|pre)\s*$", final_sem_pontuacao, re.IGNORECASE):
            return False, 'FINAL_GRAMATICALMENTE_CORROMPIDO'

    # Palavras muito curtas e atípicas no final são fortes sinais de
    # corte de palavra: "pre", "efi", etc.
    ultima_palavra = re.findall(r"\b[\wÀ-ÿ'’-]+\b", texto_original)
    if ultima_palavra:
        ultima_token = ultima_palavra[-1].casefold().strip('.')
        tokens_curto_suspeitos = {'pre', 'efi'}
        if ultima_token in tokens_curto_suspeitos:
            return False, 'PALAVRA_FINAL_TRUNCADA'

    return True, 'OK'



def auditar_saida_ollama_editorial(texto, identidade_fonte=None, entidades_proibidas=None):
    """Barreira final para impedir lixo editorial, truncamento e resíduos de fonte."""
    t = re.sub(r"\s+", " ", str(texto or "").strip())
    if not t:
        return False, "VAZIO"

    # HARD VETO DETERMINÍSTICO: não depende de ENTIDADES_PROIBIDAS.
    ok_integridade, motivo_integridade = _auditoria_integridade_deterministica(
        t, identidade_fonte
    )
    if not ok_integridade:
        return False, motivo_integridade

    # Reaproveita os bloqueios comerciais/identidade já consolidados.
    if not fragmento_eh_comercialmente_limpo(t, identidade_fonte):
        return False, "CONTAMINACAO_COMERCIAL_ENTIDADE"

    # Entidades de outras fontes não podem atravessar a redação.
    # A entidade da própria fonte já é tratada por fragmento_eh_comercialmente_limpo().
    proibidas = entidades_proibidas
    if proibidas is None:
        proibidas = ENTIDADES_PROIBIDAS_PAGINA
    if isinstance(proibidas, (set, list, tuple)):
        n_saida = normalizar_assunto_texto(t)
        nome_atual = ""
        if isinstance(identidade_fonte, dict):
            nome_atual = normalizar_assunto_texto(identidade_fonte.get("nome", "")).strip()
        permitidas_entidades = {"nbr", "abnt", "iso", "ansi", "api", "astm", "din", "inmetro"}
        for entidade in proibidas:
            entidade = str(entidade or "").strip()
            entidade_n = normalizar_assunto_texto(entidade).strip()
            if len(entidade_n) < 4 or entidade_n == nome_atual:
                continue
            if entidade_n in permitidas_entidades:
                continue
            if re.search(r"(?<![a-z0-9])" + re.escape(entidade_n) + r"(?![a-z0-9])", n_saida, re.I):
                return False, f"ENTIDADE_PROIBIDA:{entidade}"

    if re.search(r"<\/?(?:p|div|span|h[1-6]|li|ul|ol|table|tr|td|br)\b", t, re.I):
        return False, "HTML"
    if re.search(r"https?://|www\.|\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b", t, re.I):
        return False, "URL_OU_EMAIL"

    # Integridade de abertura: não aceitar continuação de frase, marcador
    # ou caractere de fechamento herdado de um recorte anterior.
    if re.match(r"^[,.;:!?)]", t):
        return False, "INICIO_TRUNCADO"
    primeiro_caractere = next((c for c in t if c.isalpha()), "")
    if primeiro_caractere and primeiro_caractere.islower():
        return False, "INICIO_COM_MINUSCULA"

    n = normalizar_assunto_texto(t)
    frases = [x.strip() for x in re.split(r"(?<=[.!?])\s+", t) if x.strip()]
    if not frases:
        return False, "SEM_FRASE"

    # Perguntas/títulos herdados da página de origem.
    # IMPORTANTE: a comparação usa texto sem acentos porque
    # normalizar_assunto_texto() remove os diacríticos.
    padroes_titulo_pergunta = (
        r"^(?:como funciona|o que e|o que sao|"
        r"problemas frequentes|perguntas frequentes|faq|"
        r"diferencas entre|caracteristicas|aplicacoes|tipos de|pecas para|"
        r"produtos relacionados|servicos relacionados)\b"
    )
    for frase in frases:
        fn = normalizar_assunto_texto(frase).strip()
        palavras = re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]*\b", frase)
        if len(palavras) <= 18 and re.search(padroes_titulo_pergunta, fn, re.I):
            if frase.endswith("?") or frase.endswith(":") or len(palavras) <= 8:
                return False, "TITULO_OU_PERGUNTA_HERDADA"

    # Pergunta herdada pode aparecer no FINAL de uma frase maior, como:
    # "... na periferia do rotor, como funciona uma Bomba Centrífuga?"
    # Nesse caso o teste acima não enxerga porque a frase não começa pela
    # expressão. Detectamos somente a cauda curta em formato interrogativo.
    if re.search(
        r"(?:^|[,;:]\s+)como funciona\b[^.!?]{0,90}\?$",
        n,
        re.I
    ):
        return False, "TITULO_OU_PERGUNTA_HERDADA"
    if re.search(
        r"(?:^|[,;:]\s+)o que (?:e|sao)\b[^.!?]{0,90}\?$",
        n,
        re.I
    ):
        return False, "TITULO_OU_PERGUNTA_HERDADA"

    # Metatexto que o modelo pode introduzir mesmo quando o fragmento
    # de origem estava limpo. Deve ser rejeitado antes de qualquer gravação.
    padroes_meta_sistema_saida = (
        r"\bo\s+python\s+(?:ja|já)\s+(?:pesquisou|selecionou|filtrou|preparou)\b",
        r"\bpython\s+(?:ja|já)\s+(?:pesquisou|selecionou|filtrou|preparou)\b",
        r"\bagora\s+que\s+voce\s+(?:ja\s+)?sabe\b",
        r"\bagora\s+que\s+você\s+(?:já\s+)?sabe\b",
        r"\btrecho\s+(?:ja|já)\s+(?:foi|esta|está)\s+(?:pesquisado|selecionado|autorizado)\b",
    )
    for padrao in padroes_meta_sistema_saida:
        if re.search(padrao, n, re.I):
            return False, "META_DO_SISTEMA_OU_TUTORIAL"

    padroes_meta = (
        r"\bleia\s+tambem\b", r"\bleia\s+também\b", r"\bsaiba\s+mais\b",
        r"\bclique\s+aqui\b", r"\bveja\s+tambem\b", r"\bveja\s+também\b",
        r"\bconfira\s+tambem\b", r"\bconfira\s+também\b",
        r"\bacompanhe\b", r"\bneste\s+artigo\b", r"\bneste\s+texto\b",
        r"\bnesta\s+pagina\b", r"\bnesta\s+página\b", r"\bo\s+objetivo\s+deste\s+artigo\b",
        r"\beste\s+texto\s+(?:ira|irá)\b", r"\beste\s+artigo\s+(?:ira|irá)\b",
        r"\bvamos\s+explorar\b", r"\bsera\s+abordado\b", r"\bserá\s+abordado\b",
        r"\bconteudos\s+relacionados\b", r"\bconteúdos\s+relacionados\b",
        r"\bartigos\s+relacionados\b", r"\bposts?\s+relacionados\b"
    )
    for padrao in padroes_meta:
        if re.search(padrao, n, re.I):
            return False, "META_EDITORIAL"

    # Títulos/âncoras residuais em qualquer posição, não apenas no começo.
    # Usamos a forma normalizada para que "Peças" e "pecas" sejam tratados
    # de maneira idêntica.
    if re.search(r"(?:^|[.!?]\s+)(?:figura|fig\.?|tabela|quadro|imagem|foto)\s*\d*\s*[:.-]", n, re.I):
        return False, "LEGENDA_OU_FIGURA_EMBUTIDA"
    if re.search(r"^(?:o\s+que\s+(?:e|sao)|como\s+escolher|como\s+funciona|onde\s+(?:aplicar|usar))\b", n, re.I):
        return False, "TITULO_OU_PERGUNTA_HERDADA"

    padroes_residuos = (
        r"\bpecas\s+para\s+[a-z0-9à-ÿ ]{2,50}$",
        r"\bprodutos\s+relacionados\b",
        r"\bperguntas\s+frequentes\b",
        r"\bfaq\b"
    )
    for padrao in padroes_residuos:
        if re.search(padrao, n, re.I):
            return False, "RESIDUO_DE_CATEGORIA"

    # Um parágrafo final nunca pode terminar aberto.
    if not re.search(r"[.!?]$", t):
        return False, "FINAL_SEM_PONTUACAO"
    if re.search(r"[,;:/\\-]$", t):
        return False, "FINAL_TRUNCADO"
    if re.search(r"\b(?:e|ou|que|de|da|do|das|dos|para|por|com|como|quando|onde|sendo|incluindo|conforme|devido|atraves|através|a|o)\s*[.!?]$", t, re.I):
        return False, "FINAL_COM_CONECTOR"

    # Palavra claramente cortada ou artefato de interrupção.
    ultima = re.findall(r"[A-Za-zÀ-ÿÀ-ÿ0-9'’-]+", t)
    if ultima:
        token = ultima[-1].casefold().strip(".!?,;:")
        if token in {"transf", "transfe", "precisam", "necess", "compro", "compr", "efi", "pre"}:
            # 'precisam' é legítimo em geral; só bloquear quando a frase termina
            # exatamente no verbo sem complemento, sinal observado no teste.
            if token != "precisam" or re.search(r"\bvolumes?\s+de\s+líquido\s+precisam[.!?]$", n, re.I):
                return False, "PALAVRA_OU_FINAL_SUSPEITO"

    # Fragmentos finais muito curtos normalmente são âncoras/títulos colados.
    if len(frases) >= 2:
        ultima_frase_palavras = re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]*\b", frases[-1])
        if len(ultima_frase_palavras) <= 4 and frases[-1].endswith((":", "?")):
            return False, "FRAGMENTO_FINAL_DE_TITULO"


    # Comercialização, propaganda e texto institucional residual.
    padroes_comerciais_saida = (
        r"\bem promocao\b", r"\bpromocao\b",
        r"\bconfira nossa selecao\b",
        r"\bmelhores? bombas?\b", r"\botimas? opcoes?\b",
        r"\bmelhor escolha\b", r"\bpreco competitivo\b",
        r"\bpreco justo\b",
        r"\bsolicite (?:um )?(?:orcamento|cotacao)\b",
        r"\bpeca (?:um )?(?:orcamento|cotacao)\b",
        r"\bentre em contato\b", r"\bfale conosco\b",
        r"\ba empresa oferece\b", r"\bservicos personalizados\b",
        r"\bganhou reconhecimento\b",
    )
    for padrao in padroes_comerciais_saida:
        if re.search(padrao, n, re.I):
            return False, "LINGUAGEM_COMERCIAL_OU_INSTITUCIONAL"

    # Espanhol residual evidente. Não é um detector de idioma geral; são
    # marcadores que, em conteúdo técnico em português, indicam mistura de
    # fonte estrangeira ou resposta inadequada do modelo.
    marcadores_espanhol = (
        r"\b(?:los|las|una|unos|unas|principales|problemas|incluyen|"
        r"correctamente|instalada|permanece|fugas|vibraciones|"
        r"mantenimiento|desgaste|alineacion|desalineacion|"
        r"rodamientos|holgura|calidad|rendimiento|vida util)\b"
    )
    palavras_es = re.findall(marcadores_espanhol, n, re.I)
    if len(palavras_es) >= 2:
        return False, "IDIOMA_ESTRANGEIRO"

    return True, "OK"


def validar_progressao_editorial_bloco(paragrafos):
    """Verifica diversidade e progressão sem impor papéis artificiais aos parágrafos."""
    if not isinstance(paragrafos, list) or len(paragrafos) != 3:
        return False, "PROGRESSAO_SEM_3_PARAGRAFOS"

    textos = [str(p or "").strip() for p in paragrafos]
    conjuntos = []
    stop = {
        "para", "com", "sem", "sobre", "entre", "como", "uma", "um", "dos", "das",
        "que", "por", "pelo", "pela", "de", "da", "do", "em", "no", "na", "ao",
        "e", "ou", "se", "ser", "tem", "ter", "mais", "tambem", "também", "cada",
        "quando", "onde", "isso", "esse", "essa", "este", "esta", "seu", "sua",
        "seus", "suas", "outro", "outra", "outros", "outras"
    }

    for texto in textos:
        tokens = {
            t for t in re.findall(r"\b[a-z0-9]{4,}\b", normalizar_assunto_texto(texto))
            if t not in stop
        }
        conjuntos.append(tokens)

    # Não permitir dois parágrafos praticamente repetidos.
    for i in range(2):
        sim = _similaridade_textual(textos[i], textos[i + 1])
        jac = _jaccard_conceitual(textos[i], textos[i + 1])
        if sim >= 0.90 or (jac >= 0.82 and sim >= 0.68):
            return False, f"PROGRESSAO_FRACA_P{i+1}_P{i+2}:sim={sim:.3f}:conceito={jac:.3f}"

    # Cada avanço deve acrescentar algum vocabulário técnico/conceitual novo.
    acumulado = conjuntos[0]
    for indice in (1, 2):
        novos = conjuntos[indice] - acumulado
        if len(conjuntos[indice]) >= 10 and len(novos) < 1:
            return False, f"PROGRESSAO_SEM_NOVOS_CONCEITOS_P{indice+1}"
        acumulado = acumulado | conjuntos[indice]

    return True, "OK"


def validar_bloco_editorial_unico(paragrafos):
    """Autoridade para estrutura e integridade editorial do bloco, sem limite de palavras."""
    if not isinstance(paragrafos, list) or len(paragrafos) != 3:
        return False, "BLOCO_DEVE_TER_3_PARAGRAFOS", None

    contagens = [len(str(p or "").split()) for p in paragrafos]
    total = sum(contagens)

    for idx, paragrafo in enumerate(paragrafos, start=1):
        ok, motivo = auditar_saida_ollama_editorial(str(paragrafo or ""), None)
        if not ok:
            return False, f"EDITORIAL_P{idx}:{motivo}", idx

    ok_progressao, motivo_progressao = validar_progressao_editorial_bloco(paragrafos)
    if not ok_progressao:
        # Quando a falha é de progressão, reselecionamos o último trecho por padrão.
        return False, motivo_progressao, 3

    return True, "OK_SEM_LIMITE_DE_PALAVRAS", None


def _similaridade_textual(a, b):
    from difflib import SequenceMatcher
    na = normalizar_assunto_texto(a)
    nb = normalizar_assunto_texto(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


def _jaccard_conceitual(a, b):
    stop = {
        "para", "com", "sem", "sobre", "entre", "como", "uma", "um", "dos", "das",
        "que", "por", "pelo", "pela", "de", "da", "do", "em", "no", "na", "ao",
        "e", "ou", "se", "ser", "tem", "ter", "mais", "tambem", "também", "cada"
    }
    ta = {x for x in re.findall(r"\b[a-z0-9]{4,}\b", normalizar_assunto_texto(a)) if x not in stop}
    tb = {x for x in re.findall(r"\b[a-z0-9]{4,}\b", normalizar_assunto_texto(b)) if x not in stop}
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(1, len(ta | tb))


def validar_unicidade_blocos(blocos, tema=""):
    """Impede títulos duplicados e sobreposição editorial forte entre blocos."""
    if not isinstance(blocos, dict):
        return False, "BLOCOS_INVALIDOS"

    itens = []
    for numero in range(1, 6):
        chave = f"bloco_{numero}"
        bloco = blocos.get(chave, {})
        if not isinstance(bloco, dict):
            return False, f"{chave}_AUSENTE"
        titulo = str(bloco.get("titulo", "") or "").strip()
        paras = [str(x or "").strip() for x in bloco.get("paragrafos_ollama", [])[:3]]
        frags = [str(x.get("texto", "") or "").strip() for x in bloco.get("informacoes_relevantes", [])[:3] if isinstance(x, dict)]
        if not titulo:
            return False, f"{chave}_SEM_TITULO"
        itens.append((chave, titulo, paras, frags))

    for i in range(len(itens)):
        for j in range(i + 1, len(itens)):
            a, b = itens[i], itens[j]
            if normalizar_assunto_texto(a[1]) == normalizar_assunto_texto(b[1]):
                return False, f"TITULOS_DUPLICADOS:{a[0]}x{b[0]}"
            sim_titulo = _similaridade_textual(a[1], b[1])
            if sim_titulo >= 0.90:
                return False, f"TITULOS_MUITO_PARECIDOS:{a[0]}x{b[0]}:{sim_titulo:.3f}"

            for pa in a[2]:
                for pb in b[2]:
                    sim = _similaridade_textual(pa, pb)
                    jac = _jaccard_conceitual(pa, pb)
                    if sim >= 0.86 or (jac >= 0.78 and sim >= 0.62):
                        return False, f"CONFLITO_EDITORIAL:{a[0]}x{b[0]}:sim={sim:.3f}:conceito={jac:.3f}"

            for fa in a[3]:
                for fb in b[3]:
                    sim = _similaridade_textual(fa, fb)
                    if sim >= 0.90:
                        return False, f"FRAGMENTOS_DUPLICADOS:{a[0]}x{b[0]}:{sim:.3f}"

    return True, "OK"


def validar_pagina_editorial_final(pagina, tema=""):
    """Barreira editorial final: tamanho, integridade, lixo e unicidade."""
    if not isinstance(pagina, dict):
        return False, "PAGINA_INVALIDA"

    blocos = {}
    for numero in range(1, 6):
        chave = f"bloco_{numero}"
        bloco = pagina.get(chave, {})
        if not isinstance(bloco, dict):
            return False, f"{chave}_INVALIDO"
        paras = bloco.get("paragrafos_ollama", [])
        ok, motivo, _ = validar_bloco_editorial_unico(paras)
        if not ok:
            return False, f"{chave}:{motivo}"
        blocos[chave] = bloco

    ok, motivo = validar_unicidade_blocos(blocos, tema)
    if not ok:
        return False, motivo

    return True, "OK"

def preparar_fragmento_para_ollama_python(fragmento):
    """Limpeza conservadora: melhora a forma sem resumir nem reescrever."""
    if not isinstance(fragmento, dict):
        return None

    texto = str(fragmento.get('texto', '') or '').strip()
    if not texto:
        return None

    # Somente normalização de espaços. Nenhuma palavra é criada ou removida.
    texto = re.sub(r'[ \t]+', ' ', texto)
    texto = re.sub(r'\n{3,}', '\n\n', texto).strip()

    novo = dict(fragmento)
    novo['texto'] = texto
    novo['palavras'] = len(texto.split())
    novo['auditoria_python_final'] = 'OK'
    return novo


def validar_conjunto_final_python(fragmentos, quantidade_esperada=15):
    """Valida os fragmentos imediatamente antes da montagem dos blocos."""
    if not isinstance(fragmentos, list) or len(fragmentos) != quantidade_esperada:
        return False, f'QUANTIDADE_{len(fragmentos) if isinstance(fragmentos, list) else 0}'

    hashes = set()
    por_bloco = {}

    for frag in fragmentos:
        if not isinstance(frag, dict):
            return False, 'FRAGMENTO_NAO_DICT'

        texto = str(frag.get('texto', '') or '').strip()
        ok, motivo = auditar_fragmento_final_python(
            texto, frag.get('identidade_fonte', {})
        )
        if not ok:
            return False, f"{frag.get('id','SEM_ID')}:{motivo}"

        h = frag.get('hash') or gerar_hash_trecho(texto)
        if h in hashes:
            return False, f"DUPLICADO:{h}"
        hashes.add(h)

        bloco = str(frag.get('bloco_mead', '') or '')
        por_bloco[bloco] = por_bloco.get(bloco, 0) + 1

    for numero in range(1, 6):
        if por_bloco.get(f'bloco_{numero}', 0) != 3:
            return False, f"BLOCO_{numero}_INCOMPLETO"

    return True, 'OK'


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

    import random

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

    if not isinstance(
        segmentos_validos,
        list
    ):
        segmentos_validos = []

    # ========================================================
    # LIMPAR DUPLICIDADES
    # ========================================================

    segmentos_unicos = []

    vistos = set()

    for segmento in segmentos_validos:

        segmento = str(
            segmento or ""
        ).strip()

        if not segmento:
            continue

        chave = segmento.casefold()

        if chave in vistos:
            continue

        vistos.add(chave)

        segmentos_unicos.append(
            segmento
        )

    # ========================================================
    # SEGURANÇA
    # ========================================================

    if len(segmentos_unicos) < 12:

        print()
        print(
            "❌ ERRO: biblioteca fixa possui menos "
            "de 12 segmentos válidos."
        )

        print(
            f"TEMA: {tema}"
        )

        print(
            f"SEGMENTOS DISPONÍVEIS: "
            f"{len(segmentos_unicos)}"
        )

        return []

    # ========================================================
    # GARANTIR QUE TODOS REFERENCIAM O TEMA
    # ========================================================

    tema_normalizado = tema.casefold()

    segmentos_com_tema = []

    for segmento in segmentos_unicos:

        if tema_normalizado in segmento.casefold():

            segmentos_com_tema.append(
                segmento
            )

    # ========================================================
    # SEGURANÇA FINAL
    # ========================================================

    if len(segmentos_com_tema) < 12:

        print()
        print(
            "❌ ERRO: biblioteca fixa não possui "
            "12 segmentos com referência ao tema."
        )

        print(
            f"TEMA: {tema}"
        )

        print(
            f"SEGMENTOS VÁLIDOS: "
            f"{len(segmentos_com_tema)}"
        )

        return []

    # ========================================================
    # SORTEAR EXATAMENTE 12
    # ========================================================

    segmentos_finais = random.sample(
        segmentos_com_tema,
        12
    )

    # ========================================================
    # AUDITORIA
    # ========================================================

    print()
    print(
        "=========================================="
    )

    print(
        "SEGMENTOS DA PÁGINA"
    )

    print(
        "=========================================="
    )

    print(
        f"TEMA: {tema}"
    )

    print(
        f"TIPO: {tipo or 'não informado'}"
    )

    print(
        f"BANCO FIXO DISPONÍVEL: "
        f"{len(segmentos_unicos)}"
    )

    print(
        f"SEGMENTOS COM TEMA: "
        f"{len(segmentos_com_tema)}"
    )

    print(
        "SEGMENTOS SORTEADOS: 12"
    )

    for numero, segmento in enumerate(
        segmentos_finais,
        start=1
    ):

        print(
            f"SEGMENTO_{numero}: "
            f"{segmento}"
        )

    print(
        "=========================================="
    )

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


# Blocos que chegaram à contingência final 8/8.
# É reiniciado a cada geração de página.
BLOCOS_APROVADOS_NO_LIMITE = set()


def avaliar_qualidade_final_python(
    titulo,
    paragrafos,
    fragmentos,
    bloco_mead,
    tema
):
    """Gate final determinístico de qualidade editorial antes da gravação.

    A pontuação é complementar aos gates obrigatórios. Falhas críticas
    reprovam independentemente da nota. O Python não inventa fatos nem
    pede avaliação subjetiva ao Ollama.
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
        r"\bacompanhe\b"
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
    if any(re.search(p, texto_n, re.I) for p in comercial):
        falhas_criticas.append("CONTEUDO_COMERCIAL_OU_CTA")

    # ------------------------------------------------------------
    # 3. INTEGRIDADE BÁSICA DO TEXTO FINAL
    # ------------------------------------------------------------
    if len(paragrafos) != 3 or any(not p for p in paragrafos):
        falhas_criticas.append("QUANTIDADE_PARAGRAFOS")

    for i, p in enumerate(paragrafos, 1):
        if not re.search(r"[.!?]$", p):
            falhas_criticas.append(f"PARAGRAFO_{i}_SEM_PONTUACAO")
        if re.search(r"<[^>]+>|https?://|www\\.", p, re.I):
            falhas_criticas.append(f"PARAGRAFO_{i}_RESIDUO_ESTRUTURAL")
        if re.match(r"^[,.;:!?)]", p) or (p and p[0].islower()):
            falhas_criticas.append(f"PARAGRAFO_{i}_INICIO_TRUNCADO")

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
    tokens_fonte = {
        x for x in re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]{3,}\b", fonte_n.lower())
        if x not in stop
    }
    tokens_saida = {
        x for x in re.findall(r"\b[\wÀ-ÿ][\wÀ-ÿ'’-]{3,}\b", texto_n.lower())
        if x not in stop
    }
    compartilhados = tokens_saida & tokens_fonte
    cobertura = len(compartilhados) / max(1, len(tokens_saida))
    if len(tokens_saida) >= 45 and cobertura < 0.18:
        falhas_criticas.append("FIDELIDADE_LEXICAL_INSUFICIENTE")
    elif cobertura < 0.26:
        descontos += 5

    # Números e unidades novas continuam sendo falha crítica.
    numeros_fonte = set(re.findall(r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)", fonte_n))
    numeros_saida = set(re.findall(r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)", texto_n))
    numeros_novos = numeros_saida - numeros_fonte
    if numeros_novos:
        falhas_criticas.append("NUMEROS_OU_UNIDADES_NOVOS")

    # ------------------------------------------------------------
    # 6. COMPATIBILIDADE MEAD
    # ------------------------------------------------------------
    papeis = {
        "bloco_1": ["context", "import", "necess", "cenar", "demanda", "abastec", "relev"],
        "bloco_2": ["funcion", "aplic", "vazao", "press", "impuls", "rotor", "compon", "caracter"],
        "bloco_3": ["manuten", "instal", "operac", "segur", "dimension", "corros", "cuidad", "criter"],
        # B4 também pode cumprir a função de CONHECIMENTO TÉCNICO.
        # Não exigir a palavra literal "conhecimento" no texto final.
        "bloco_4": ["empresa", "suport", "atend", "servic", "soluc", "conhec", "fornec", "equipe", "especific", "dimension", "selec", "caracter", "compon", "operac", "desempen", "confiab", "criter", "cuidad"],
        "bloco_5": ["selec", "aplic", "necess", "soluc", "decis", "confiab", "operac", "manuten"]
    }
    sinais = papeis.get(str(bloco_mead), [])
    if sinais:
        encontrados = sum(1 for sinal in sinais if sinal in texto_n)
        if encontrados == 0:
            falhas_criticas.append("INCOMPATIBILIDADE_MEAD")
        elif encontrados == 1:
            descontos += 6

    # O título também precisa apontar para o papel do bloco.
    titulo_sinais = sum(1 for sinal in sinais if sinal in titulo_n)
    if sinais and titulo_sinais == 0:
        descontos += 5

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
        if max(jaccards) >= 0.72:
            descontos += 7

    # ------------------------------------------------------------
    # NOTA
    # ------------------------------------------------------------
    nota = max(0.0, 10.0 - descontos / 10.0)
    if ocorrencias_generalidade:
        nota = min(nota, 8.9)
    if falhas_criticas:
        nota = min(nota, 7.9)

    aprovado = not falhas_criticas and nota >= 8.5
    return aprovado, round(nota, 2), falhas_criticas, {
        "cobertura_factual_lexical": round(cobertura, 3),
        "generalidades": ocorrencias_generalidade,
        "numeros_novos": sorted(numeros_novos),
        "descontos": descontos
    }


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

    global BLOCOS_APROVADOS_NO_LIMITE
    BLOCOS_APROVADOS_NO_LIMITE = set()

    inicio_geracao = time.time()

    tema_base = str(
        tema or ""
    ).strip()

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
    print("======================================")
    print("CHECKBOXES EDITORIAIS")
    print("======================================")
    
    if not isinstance(
        estrutura_editorial,
        dict
    ):
    
        estrutura_editorial = {}
    
    for chave, valor in estrutura_editorial.items():
    
        print(
            f"{chave}: {valor}"
        )
    
    assuntos = [
        str(chave).replace("_", " ")
        for chave, valor
        in estrutura_editorial.items()
        if valor
    ]
    
    print()
    print(
        "ASSUNTOS SELECIONADOS:",
        assuntos
    )
    
    print(
        "TOTAL DE ASSUNTOS:",
        len(assuntos)
    )
    
    # ========================================================
    # 03. ESTRUTURA OBRIGATÓRIA
    # ========================================================
    
    total_blocos = 5
    paragrafos_por_bloco = 3
    total_paragrafos = 15
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

    # ========================================================
    # GATE FINAL 0 — NÃO COMPLETAR COM LIXO
    # ========================================================
    # Se Python não conseguiu formar 15 fragmentos autorizados, o Ollama
    # não é chamado. Não há fallback para material de menor qualidade.
    # ========================================================
    quantidade_fragmentos = len(fragmentos_iniciais) if isinstance(fragmentos_iniciais, list) else 0
    print()
    print("============================================================")
    print("GATE FINAL — PYTHON ANTES DO OLLAMA / QUANTIDADE")
    print("============================================================")
    print("FRAGMENTOS NECESSÁRIOS: 15")
    print("FRAGMENTOS AUTORIZADOS: ", quantidade_fragmentos)
    if quantidade_fragmentos < 15:
        print("STATUS FINAL: NÃO AUTORIZADO")
        print("MOTIVO: FRAGMENTOS DE QUALIDADE INSUFICIENTES")
        print("OLLAMA: NÃO EXECUTADO")
        print("GRAVAÇÃO: NÃO AUTORIZADA")
        return None
    print("STATUS FINAL: QUANTIDADE OK")

    # Lista oficial dos fragmentos selecionados. A rotina de re-seleção
    # dos conjuntos Ollama usa esta mesma lista para substituir fragmentos
    # rejeitados; sem esta inicialização ocorre NameError após uma falha.
    if not isinstance(fragmentos_iniciais, list):
        fragmentos_iniciais = []
    fragmentos_selecionados = fragmentos_iniciais

    # Estado transitório vindo da seleção Python.
    # Os candidatos e reservas existem somente em memória e nunca
    # devem ser gravados no JSON oficial.
    candidatos_por_bloco = textos_relevantes.get(
        "_candidatos_por_bloco",
        {}
    )
    if not isinstance(candidatos_por_bloco, dict):
        candidatos_por_bloco = {}

    fragmentos_reserva_por_bloco = textos_relevantes.get(
        "_fragmentos_reserva_por_bloco",
        {}
    )
    if not isinstance(fragmentos_reserva_por_bloco, dict):
        fragmentos_reserva_por_bloco = {}

    # Estado usado pela re-seleção após rejeição do Ollama.
    # Os hashes são reconstruídos aqui, no mesmo escopo de
    # gerar_conteudo_completo(), evitando NameError em uma nova tentativa.
    hashes_selecionados = set()

    # Reconstruir o contador de fontes para a etapa de re-seleção.
    # A seleção inicial já terminou; a etapa Ollama precisa deste estado
    # para remover corretamente um conjunto rejeitado e adicionar reservas.
    fontes_utilizadas = {}

    for fragmento in fragmentos_selecionados:
        if not isinstance(fragmento, dict):
            continue

        hash_trecho = str(
            fragmento.get("hash", "") or ""
        ).strip()

        if not hash_trecho:
            hash_trecho = gerar_hash_trecho(
                fragmento.get("texto", "")
            )
            fragmento["hash"] = hash_trecho

        if hash_trecho:
            hashes_selecionados.add(hash_trecho)

    if not contexto_inicial and not fragmentos_selecionados:
    
        print()
        print("======================================")
        print("NENHUMA INFORMAÇÃO RELEVANTE")
        print("======================================")
    
        return None
    
    # ========================================================
    # GATE DURO — AUDITORIA FINAL DOS 15 ANTES DO OLLAMA
    # ========================================================
    # O objeto que chegou à fase Ollama precisa passar novamente pelas
    # barreiras editoriais, comerciais e estruturais. Se qualquer fragmento
    # falhar, o modelo NÃO é chamado.
    # ========================================================
    print()
    print("============================================================")
    print("GATE DURO — AUDITORIA FINAL DOS 15 FRAGMENTOS (IMUTAVEL)")
    print("============================================================")

    _falhas_pre_ollama = []
    for _idx_gate, _frag_gate in enumerate(fragmentos_selecionados, start=1):
        if not isinstance(_frag_gate, dict):
            _falhas_pre_ollama.append((_idx_gate, "FRAGMENTO_NAO_ESTRUTURADO"))
            continue

        _txt_gate = str(_frag_gate.get("texto", "") or "").strip()
        if not _txt_gate:
            _falhas_pre_ollama.append((_idx_gate, "TEXTO_VAZIO"))
            continue

        try:
            _ok_gate, _motivo_gate = diagnosticar_contaminacao_editorial(
                _txt_gate,
                _frag_gate.get("identidade_fonte")
            )
        except Exception:
            _ok_gate = False
            _motivo_gate = "ERRO_AUDITORIA_EDITORIAL"

        if not _ok_gate:
            _falhas_pre_ollama.append((_idx_gate, str(_motivo_gate)))
            continue

        try:
            if not fragmento_eh_aproveitavel_editorialmente(_txt_gate):
                _falhas_pre_ollama.append((_idx_gate, "FILTRO_EDITORIAL_FINAL"))
                continue
        except Exception:
            _falhas_pre_ollama.append((_idx_gate, "ERRO_FILTRO_EDITORIAL_FINAL"))
            continue

        # IMPORTANTE: nao chamar candidato_eh_utilizavel() novamente aqui.
        # O candidato ja foi aprovado por essa funcao durante a construcao
        # das combinacoes. Essa funcao depende do estado dinamico do bloco
        # e de variaveis de selecao; reutiliza-la no gate final estava
        # transformando candidatos ja aprovados em falsos rejeitados.
        #
        # O gate final deve auditar o TEXTO efetivamente selecionado, usando
        # apenas barreiras deterministicas e sem efeitos de selecao.
        try:
            if not fragmento_eh_comercialmente_limpo(
                _txt_gate,
                _frag_gate.get("identidade_fonte", {})
            ):
                _falhas_pre_ollama.append((_idx_gate, "IDENTIDADE_COMERCIAL_OU_PRODUTO_FINAL"))
                continue
        except Exception:
            _falhas_pre_ollama.append((_idx_gate, "ERRO_IDENTIDADE_COMERCIAL_FINAL"))
            continue

        try:
            _ok_audit_gate, _motivo_audit_gate = auditar_fragmento_final_python(
                _txt_gate,
                _frag_gate.get("identidade_fonte", {})
            )
            if not _ok_audit_gate:
                _falhas_pre_ollama.append((_idx_gate, "AUDITORIA_FINAL:" + str(_motivo_audit_gate)))
                continue
        except Exception:
            _falhas_pre_ollama.append((_idx_gate, "ERRO_AUDITORIA_FINAL_PYTHON"))
            continue

    if _falhas_pre_ollama:
        print("STATUS FINAL: BLOQUEADO")
        print("FRAGMENTOS REJEITADOS:", len(_falhas_pre_ollama))
        for _idx_gate, _motivo_gate in _falhas_pre_ollama:
            print(
                "❌ FRAGMENTO",
                _idx_gate,
                "NÃO CHEGARÁ AO OLLAMA:",
                _motivo_gate
            )
        print("OLLAMA: NÃO EXECUTADO")
        print("GRAVAÇÃO: NÃO AUTORIZADA")
        return None

    print("STATUS FINAL: 15/15 APROVADOS")
    print("AUDITORIA EDITORIAL: OK")
    print("AUDITORIA COMERCIAL: OK")
    print("AUDITORIA ESTRUTURAL: OK")
    print("OLLAMA: LIBERADO")

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
    # fragmentos_autorizados
    #
    # Os fragmentos_autorizados precisam continuar disponíveis
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
                "fragmentos_autorizados": []
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
                "fragmentos_autorizados": []
            }

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
    # CONTAGEM DOS 15 FRAGMENTOS AUTORIZADOS PYTHON
    # ========================================================

    total_fragmentos_autorizados = 0

    print()
    print("======================================")
    print("FRAGMENTOS AUTORIZADOS PYTHON")
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

            fragmentos_autorizados = []

        else:

            fragmentos_autorizados = (
                dados_bloco.get(
                    "fragmentos_autorizados",
                    []
                )
            )

            if not isinstance(
                fragmentos_autorizados,
                list
            ):

                fragmentos_autorizados = []

        quantidade = len(
            fragmentos_autorizados
        )

        total_fragmentos_autorizados += (
            quantidade
        )

        print(
            f"BLOCO {numero_bloco} "
            f"- PARÁGRAFOS PYTHON:",
            quantidade
        )

    print()
    print(
        "TOTAL FRAGMENTOS AUTORIZADOS PYTHON:",
        total_fragmentos_autorizados
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
    # EXIBIR INFORMAÇÕES DE CADA BLOCO
    # ========================================================

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

        if isinstance(
            dados_bloco,
            dict
        ):

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

                partes = []

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

                    if texto_item:

                        partes.append(
                            texto_item
                        )

                texto_bloco = (
                    "\n\n".join(
                        partes
                    )
                )

            else:

                texto_bloco = str(
                    lista_informacoes or ""
                ).strip()

        else:

            texto_bloco = str(
                dados_bloco or ""
            ).strip()

        print(
            f"BLOCO {numero_bloco}:",
            len(
                texto_bloco
            ),
            "caracteres"
        )

    # ========================================================
    # INFORMAÇÕES POR BLOCO PREPARADAS
    # ========================================================

    print()
    print("======================================")
    print("INFORMAÇÕES POR BLOCO PREPARADAS")
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

        print()
        print(
            f"--- BLOCO {numero_bloco} ---"
        )

        if isinstance(
            dados_bloco,
            dict
        ):

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

                partes = []

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

                    if texto_item:

                        partes.append(
                            texto_item
                        )

                texto_bloco = (
                    "\n\n".join(
                        partes
                    )
                )

            else:

                texto_bloco = str(
                    lista_informacoes or ""
                ).strip()

        else:

            texto_bloco = str(
                dados_bloco or ""
            ).strip()

        print(
            texto_bloco[:500]
        )

    # ========================================================
    # CONFERÊNCIA FINAL DOS 15 PARÁGRAFOS-BASE
    # ========================================================

    print()
    print("======================================")
    print("CONFERÊNCIA DOS PARÁGRAFOS PYTHON")
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

            continue

        fragmentos_autorizados = (
            dados_bloco.get(
                "fragmentos_autorizados",
                []
            )
        )

        if not isinstance(
            fragmentos_autorizados,
            list
        ):

            continue

        for indice, paragrafo in enumerate(
            fragmentos_autorizados,
            start=1
        ):

            texto_paragrafo = str(
                paragrafo or ""
            ).strip()

            print(
                f"BLOCO {numero_bloco} - "
                f"PARÁGRAFO {indice}: "
                f"{len(texto_paragrafo.split())} palavras"
            )

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
        "CARACTERES SELECIONADOS:",
        caracteres_selecionados
    )
    
    
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
    
    print()
    print("======================================")
    print("CONTEXTO SELECIONADO POR BLOCO")
    print("======================================")
    
    for numero_bloco in range(
        1,
        6
    ):
    
        chave_bloco = (
            f"bloco_{numero_bloco}"
        )
    
        texto_bloco = contexto_geracao.get(
            chave_bloco,
            ""
        )
    
        print(
            f"{chave_bloco}:",
            len(texto_bloco),
            "caracteres"
        )
    
    print()
    print(
        "TOTAL CARACTERES CONTEXTO:",
        caracteres_contexto_geracao
    )
    
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
    # TÍTULOS GERADOS PELO PYTHON
    #
    # O Python é responsável por:
    # - H1
    # - título SEO
    # - subtítulo
    # - títulos dos 5 blocos
    #
    # O Ollama NÃO gera nenhum desses elementos.
    # ============================================================
    
    titulos_python = gerar_titulos(
        tema,
        informacoes_blocos,
        mapa_mead
    )
    
    h1 = titulos_python.get(
        "h1",
        ""
    )
    
    titulo = titulos_python.get(
        "title",
        ""
    )
    
    subtitulo = titulos_python.get(
        "subtitulo",
        ""
    )

    subtitulo_listas = titulos_python.get(
        "subtitulo_listas",
        ""
    )
    
    subtitulo_segmentos = titulos_python.get(
        "subtitulo_segmentos",
        ""
    )
    
    
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
        "3 fragmentos por bloco"
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
        total_blocos
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
        "MEAD EDITORIAL NO PROMPT:",
        "NÃO"
    )

    print(
        "CHECKBOXES NO PROMPT:",
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
    # 11. PROCESSAMENTO POR BLOCO — OLLAMA
    # ============================================================
    #
    # NOVA ARQUITETURA:
    #
    # Cada bloco possui exatamente 3 fragmentos.
    #
    # Os 3 fragmentos são enviados juntos em UMA única
    # chamada ao Ollama.
    #
    # O Ollama retorna exatamente 3 parágrafos.
    #
    # Portanto:
    #
    # 5 blocos x 1 chamada = 5 chamadas Ollama
    #
    # Mantemos:
    #
    # - os mesmos 15 fragmentos;
    # - a mesma ordem;
    # - os mesmos 5 blocos;
    # - 3 parágrafos por bloco;
    # - a mesma estrutura do JSON.
    #
    # ============================================================

    print()
    print("=" * 60)
    print("PROCESSAMENTO POR BLOCO DOS FRAGMENTOS PELO OLLAMA")
    print("=" * 60)

    print(
        "CHAMADAS PREVISTAS:",
        total_blocos
    )

    print(
        "FRAGMENTOS POR CHAMADA:",
        3
    )

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

    # ============================================================
    # RETENTATIVA CONTROLADA DE REDAÇÃO DO BLOCO
    # ============================================================
    #
    # Um bloco só é aprovado quando:
    #
    # - possui exatamente 3 parágrafos;
    # - todos possuem exatamente 3 parágrafos;
    # - todos passam pela validação factual;
    # - nenhum número/unidade novo é introduzido.
    #
    # Se qualquer validação falhar, o mesmo bloco é reenviado
    # ao Ollama, com o motivo da rejeição.
    #
    # IMPORTANTE:
    # nenhum resultado inválido é gravado em paragrafos_ollama.
    #
    # ============================================================

    # LIMITES DE PALAVRAS REMOVIDOS.
    # A aprovação depende de estrutura, qualidade editorial, factualidade,
    # progressão, integridade e fidelidade — não de quantidade de palavras.
    # Tentativas do Ollama sem limite numérico.
    # O ciclo só termina quando o bloco é aprovado ou quando não há
    # evidência autorizada disponível para uma substituição necessária.

    # Uma tentativa é uma operação completa de bloco (3 trechos -> 3 parágrafos).
    # Não existe mais processamento/retry individual de fragmento no Ollama.
    MAX_TENTATIVAS_CONJUNTO_OLLAMA = 3
    MARCADOR_TRECHOS_OLLAMA = "__TRECHOS_AUTORIZADOS_PYTHON__"

    # IMPORTANTE: o texto NÃO é compactado antes do Ollama.
    # O Python já selecionou um trecho editorialmente fechado; reduzir
    # caracteres aqui destruiria justamente a densidade que queremos preservar.
    MAX_CHARS_TRECHO_OLLAMA = None

    def _compactar_trecho_para_ollama(texto, limite=None):
        texto = re.sub(r"\s+", " ", str(texto or "")).strip()
        return texto

    quarentena_ollama_por_bloco = {}
    historico_conjuntos_ollama = {}
    estatisticas_ollama = {
        "chamadas": 0,
        "caracteres_prompt": 0,
        "fragmentos_aprovados": 0
    }

    # Estado dos blocos disponível desde o início do processamento Ollama.
    # A rotina de reseleção pode ser chamada antes da montagem da lista
    # final; portanto, este objeto precisa existir antes de qualquer
    # tentativa de rejeição/reseleção.
    #
    # Ele é um dicionário durante o processamento e só é convertido para
    # a lista oficial de blocos mais abaixo, após os 15 parágrafos estarem
    # aprovados.
    blocos = {}
    for _numero_bloco in range(1, total_blocos + 1):
        _chave_bloco = f"bloco_{_numero_bloco}"
        _dados_bloco_inicial = informacoes_blocos.get(
            _chave_bloco, {}
        )
        if isinstance(_dados_bloco_inicial, dict):
            blocos[_chave_bloco] = dict(_dados_bloco_inicial)
        else:
            blocos[_chave_bloco] = {}

    def _contexto_trechos_para_ollama(fragmentos):
        blocos = []
        for f in fragmentos:
            texto = _compactar_trecho_para_ollama(f.get("texto", ""))
            blocos.append(
                f"TRECHO {f.get('numero', 0)}:\n{texto}"
            )
        return "\n\n".join(blocos)

    def _reselecionar_conjunto_apos_rejeicao(
        chave_bloco,
        fragmentos_atuais,
        motivo,
        fragmentos_selecionados,
        hashes_selecionados,
        fontes_utilizadas,
        candidatos_por_bloco,
        blocos,
        blocos_informacoes,
        fragmentos_reserva_por_bloco
    ):
        bloqueados = quarentena_ollama_por_bloco.setdefault(chave_bloco, set())
        antigos = []
        for f in fragmentos_atuais:
            h = str(f.get('hash', '') or '').strip() or gerar_hash_trecho(f.get('texto', ''))
            f['hash'] = h
            bloqueados.add(h)
            antigos.append((h, f.get('fonte')))
        antigos_hashes = {h for h, _ in antigos}
        fragmentos_selecionados[:] = [x for x in fragmentos_selecionados if x.get('hash') not in antigos_hashes]
        for h, fonte in antigos:
            hashes_selecionados.discard(h)
            if fonte in fontes_utilizadas:
                fontes_utilizadas[fonte] -= 1
                if fontes_utilizadas[fonte] <= 0:
                    fontes_utilizadas.pop(fonte, None)
        novos = []

        # Primeiro tenta a próxima reserva deste bloco.
        # A fila garante: reserva 1 na primeira rejeição,
        # reserva 2 na segunda.
        reservas = fragmentos_reserva_por_bloco.get(
            chave_bloco,
            []
        )

        while reservas:
            reserva = dict(
                reservas.pop(0)
            )

            hash_reserva = (
                reserva.get("hash")
                or gerar_hash_trecho(
                    reserva.get("texto", "")
                )
            )
            reserva["hash"] = hash_reserva

            if (
                hash_reserva
                and hash_reserva not in hashes_selecionados
                and hash_reserva not in bloqueados
            ):
                reserva.pop("_reserva_memoria", None)
                reserva["bloco_mead"] = chave_bloco
                hashes_selecionados.add(
                    hash_reserva
                )

                fonte_reserva = reserva.get("fonte")
                fontes_utilizadas[fonte_reserva] = (
                    fontes_utilizadas.get(
                        fonte_reserva,
                        0
                    ) + 1
                )

                fragmentos_selecionados.append(
                    reserva
                )
                novos.append(
                    reserva
                )
                break

        candidatos_disponiveis = candidatos_por_bloco.get(chave_bloco, []) if isinstance(candidatos_por_bloco, dict) else []

        for _ in range(3 - len(novos)):
            try:
                escolhido = selecionar_melhor_candidato(
                    candidatos_disponiveis,
                    chave_bloco,
                    bloqueados
                )
            except NameError:
                # Fallback de segurança: se a função de seleção não estiver
                # disponível no escopo por alguma versão antiga do arquivo,
                # ainda tentamos um candidato estruturalmente utilizável.
                escolhido = None
                for item_candidato in candidatos_disponiveis:
                    candidato_fallback = item_candidato.get("candidato") if isinstance(item_candidato, dict) else None
                    if not isinstance(candidato_fallback, dict):
                        continue
                    h_fallback = candidato_fallback.get("hash") or gerar_hash_trecho(candidato_fallback.get("texto", ""))
                    if h_fallback in hashes_selecionados or h_fallback in bloqueados:
                        continue
                    try:
                        utilizavel = candidato_eh_utilizavel(candidato_fallback)
                    except NameError:
                        utilizavel = bool(str(candidato_fallback.get("texto", "")).strip())
                    if utilizavel:
                        escolhido = candidato_fallback
                        break
            if escolhido is None:
                break
            escolhido = dict(escolhido)
            escolhido['bloco_mead'] = chave_bloco
            h = escolhido.get('hash') or gerar_hash_trecho(escolhido.get('texto', ''))
            escolhido['hash'] = h
            if h in hashes_selecionados or h in bloqueados: continue
            hashes_selecionados.add(h)
            fonte = escolhido.get('fonte')
            fontes_utilizadas[fonte] = fontes_utilizadas.get(fonte, 0) + 1
            fragmentos_selecionados.append(escolhido)
            novos.append(escolhido)
        historico_conjuntos_ollama.setdefault(chave_bloco, []).append({
            'motivo': str(motivo or ''), 'rejeitados': sorted(antigos_hashes), 'novos': [x.get('hash') for x in novos]
        })
        print(
            'RESELEÇÃO PYTHON APÓS REJEIÇÃO:',
            chave_bloco,
            len(novos),
            '/ 3',
            'RESERVAS RESTANTES:',
            len(fragmentos_reserva_por_bloco.get(chave_bloco, [])),
            'QUARENTENA:',
            len(bloqueados)
        )
        if len(novos) != 3: return None
        novos_info = [{
            'id': x.get('id',''), 'hash': x.get('hash',''), 'texto': x.get('texto',''), 'fonte': x.get('fonte',''),
            'url': x.get('url',''), 'tipo': x.get('tipo',''), 'pdf': x.get('pdf',False),
            'palavras': x.get('palavras', len(str(x.get('texto','')).split())), 'identidade_fonte': x.get('identidade_fonte',{})
        } for x in novos]
        blocos[chave_bloco]['informacoes_relevantes'] = novos_info
        blocos[chave_bloco]['fragmentos_autorizados'] = [x['texto'] for x in novos_info]
        blocos_informacoes[chave_bloco]['informacoes_relevantes'] = novos_info
        blocos_informacoes[chave_bloco]['fragmentos_autorizados'] = [x['texto'] for x in novos_info]
        return novos_info

    def _reselecionar_fragmento_apos_rejeicao(
        chave_bloco,
        fragmento_atual,
        indice_fragmento,
        motivo,
        fragmentos_selecionados,
        hashes_selecionados,
        fontes_utilizadas,
        candidatos_por_bloco,
        blocos,
        blocos_informacoes,
        fragmentos_reserva_por_bloco
    ):
        """Substitui SOMENTE o fragmento rejeitado, preservando os aprovados."""
        bloqueados = quarentena_ollama_por_bloco.setdefault(chave_bloco, set())

        antigo = dict(fragmento_atual or {})
        h_antigo = str(antigo.get("hash", "") or "").strip() or gerar_hash_trecho(antigo.get("texto", ""))
        antigo["hash"] = h_antigo
        bloqueados.add(h_antigo)

        # Remove somente o fragmento rejeitado do conjunto global.
        fragmentos_selecionados[:] = [
            x for x in fragmentos_selecionados
            if (x.get("hash") or gerar_hash_trecho(x.get("texto", ""))) != h_antigo
        ]
        hashes_selecionados.discard(h_antigo)

        fonte_antiga = antigo.get("fonte")
        if fonte_antiga in fontes_utilizadas:
            fontes_utilizadas[fonte_antiga] -= 1
            if fontes_utilizadas[fonte_antiga] <= 0:
                fontes_utilizadas.pop(fonte_antiga, None)

        novo = None
        reservas = fragmentos_reserva_por_bloco.get(chave_bloco, [])

        while reservas:
            reserva = dict(reservas.pop(0))
            h = reserva.get("hash") or gerar_hash_trecho(reserva.get("texto", ""))
            reserva["hash"] = h
            if h and h not in hashes_selecionados and h not in bloqueados:
                reserva.pop("_reserva_memoria", None)
                reserva["bloco_mead"] = chave_bloco
                novo = reserva
                break

        # Se a fila de reservas acabou, ainda podemos buscar outro
        # candidato que já foi aprovado pelo mesmo pool Python. A barreira
        # final é reaplicada antes da promoção, portanto isso não abre uma
        # porta para lixo.
        if novo is None:
            candidatos_disponiveis = candidatos_por_bloco.get(chave_bloco, []) if isinstance(candidatos_por_bloco, dict) else []
            try:
                candidato_extra = selecionar_melhor_candidato(
                    candidatos_disponiveis,
                    chave_bloco,
                    bloqueados | hashes_selecionados
                )
            except Exception:
                candidato_extra = None
            if isinstance(candidato_extra, dict):
                candidato_extra = dict(candidato_extra)
                try:
                    candidato_extra["bloco_mead"] = chave_bloco
                    h_extra = candidato_extra.get("hash") or gerar_hash_trecho(candidato_extra.get("texto", ""))
                    candidato_extra["hash"] = h_extra
                    if (
                        h_extra
                        and h_extra not in hashes_selecionados
                        and h_extra not in bloqueados
                        and candidato_eh_utilizavel(candidato_extra)
                        and fragmento_compativel_com_funcao_do_bloco(candidato_extra, chave_bloco)
                    ):
                        novo = candidato_extra
                except Exception:
                    novo = None

        if novo is None:
            print(
                "❌ RESELEÇÃO INDIVIDUAL BLOQUEADA:",
                chave_bloco,
                "FRAGMENTO:", indice_fragmento,
                "MOTIVO: nenhuma evidência Python limpa disponível"
            )

        if not isinstance(novo, dict):
            print(
                "❌ RESELEÇÃO INDIVIDUAL SEM CANDIDATO:",
                chave_bloco,
                "FRAGMENTO:", indice_fragmento
            )
            return None

        novo = dict(novo)
        novo["bloco_mead"] = chave_bloco
        h_novo = novo.get("hash") or gerar_hash_trecho(novo.get("texto", ""))
        novo["hash"] = h_novo

        if h_novo in hashes_selecionados or h_novo in bloqueados:
            return None

        hashes_selecionados.add(h_novo)
        fonte_nova = novo.get("fonte")
        fontes_utilizadas[fonte_nova] = fontes_utilizadas.get(fonte_nova, 0) + 1
        fragmentos_selecionados.append(novo)

        historico_conjuntos_ollama.setdefault(chave_bloco, []).append({
            "motivo": str(motivo or ""),
            "rejeitados": [h_antigo],
            "novos": [h_novo],
            "indice_fragmento": indice_fragmento
        })

        print(
            "RESELEÇÃO INDIVIDUAL PYTHON:",
            chave_bloco,
            "FRAGMENTO:", indice_fragmento,
            "RESERVAS RESTANTES:", len(reservas),
            "QUARENTENA:", len(bloqueados)
        )

        return novo


    def _processar_bloco_ollama_com_retentativas(
        prompt_base,
        fragmentos_autorizados,
        chave_bloco_atual
    ):

        ultimo_erro = ""

        def _pesquisar_novas_evidencias_para_bloco(chave_bloco):
            """Busca nova matéria-prima .br quando a reserva/candidatos locais acabaram."""
            termos_blocos = {
                "bloco_1": ["definição", "funcionamento", "contexto"],
                "bloco_2": ["características", "funcionamento", "aplicações"],
                "bloco_3": ["critérios técnicos", "instalação", "manutenção", "segurança"],
                "bloco_4": ["conhecimento técnico", "suporte", "atendimento"],
                "bloco_5": ["seleção", "aplicação", "dimensionamento", "solução"],
            }
            termos = termos_blocos.get(chave_bloco, ["informação técnica"])
            urls = []
            vistas = set()
            for termo in termos:
                try:
                    resultados = pesquisar(f"{tema} {termo}", limite=5)
                except Exception as erro:
                    print("⚠️ PESQUISA DE RECUPERAÇÃO FALHOU:", repr(erro))
                    continue
                for url in resultados or []:
                    u = str(url or "").strip()
                    if not u or u in vistas:
                        continue
                    vistas.add(u)
                    urls.append(u)
                    if len(urls) >= 12:
                        break
                if len(urls) >= 12:
                    break

            adicionados = 0
            candidatos = candidatos_por_bloco.setdefault(chave_bloco, [])
            hashes_existentes = set()
            for item in candidatos:
                c = item.get("candidato") if isinstance(item, dict) else None
                if isinstance(c, dict):
                    hashes_existentes.add(c.get("hash") or gerar_hash_trecho(c.get("texto", "")))

            for url in urls:
                try:
                    pagina = coletar_pagina(url)
                except Exception:
                    continue
                if not isinstance(pagina, dict):
                    continue
                texto = str(pagina.get("texto", "") or "").strip()
                if not texto:
                    continue
                frases = [x.strip() for x in re.split(r"(?<=[.!?])\s+", texto) if x.strip()]
                for tamanho_janela in (3, 2):
                    for inicio in range(0, max(0, len(frases) - tamanho_janela + 1)):
                        trecho = " ".join(frases[inicio:inicio + tamanho_janela]).strip()
                        # Sem piso artificial de palavras: a evidência
                        # será julgada pelos mesmos gates 100/100 da seleção.
                        if not re.search(r"[.!?][\"'”’»\)\]}]*$", trecho):
                            continue
                        try:
                            if not fragmento_eh_editorialmente_valido(trecho):
                                continue
                        except Exception:
                            continue
                        candidato = {
                            "id": gerar_id_trecho(trecho),
                            "hash": gerar_hash_trecho(trecho),
                            "texto": trecho,
                            "fonte": f"RECUPERACAO_{chave_bloco}",
                            "url": url,
                            "tipo": pagina.get("tipo", "recuperacao"),
                            "pdf": bool(pagina.get("pdf", False)),
                            "palavras": len(trecho.split()),
                            "identidade_fonte": pagina.get("identidade_fonte", {}),
                            "bloco_mead": chave_bloco,
                        }
                        h = candidato["hash"]
                        if h in hashes_existentes or h in quarentena_ollama_por_bloco.get(chave_bloco, set()):
                            continue
                        try:
                            if not candidato_eh_utilizavel(candidato):
                                continue
                            if not fragmento_compativel_com_funcao_do_bloco(candidato, chave_bloco):
                                continue
                        except Exception:
                            continue
                        candidatos.append({"candidato": candidato, "pontuacao": 0})
                        hashes_existentes.add(h)
                        adicionados += 1
                        if adicionados >= 30:
                            break
                    if adicionados >= 30:
                        break
                if adicionados >= 30:
                    break

            print("🔎 RECUPERAÇÃO DE PESQUISA:", chave_bloco, "URLs:", len(urls), "NOVOS CANDIDATOS:", adicionados)
            return adicionados > 0

        def _adaptar_evidencia_apos_rejeicao(motivo):
            """
            Recuperação automática de evidência.

            Quando a redação foi estruturalmente válida, mas um parágrafo
            não demonstrou evidência suficiente, o Python NÃO insiste
            indefinidamente com o mesmo trecho. Ele coloca o trecho em
            quarentena e promove outro texto autorizado da reserva.

            O parágrafo aprovado não é alterado. Somente o índice rejeitado
            é substituído.
            """
            motivo = str(motivo or "")
            indice = None

            padroes = [
                r"parágrafo\s+(\d+)",
                r"PARAGRAFO[_ ](\d+)",
                r"P(\d+)",
            ]
            for padrao in padroes:
                m = re.search(padrao, motivo, flags=re.IGNORECASE)
                if m:
                    try:
                        indice = int(m.group(1)) - 1
                    except Exception:
                        indice = None
                    break

            # Progressão fraca: por segurança substitui o último parágrafo.
            if indice is None and "PROGRESSAO" in motivo.upper():
                indice = 2

            if indice is None or indice < 0 or indice > 2:
                return False

            try:
                atual = fragmentos_autorizados[indice]
            except Exception:
                return False

            novo = _reselecionar_fragmento_apos_rejeicao(
                chave_bloco_atual,
                atual,
                indice,
                motivo,
                fragmentos_selecionados,
                hashes_selecionados,
                fontes_utilizadas,
                candidatos_por_bloco,
                blocos,
                blocos_informacoes,
                fragmentos_reserva_por_bloco
            )

            # Não pesquisar novamente durante a fase Ollama.
            # A pesquisa adicional poderia introduzir uma evidência que não
            # participou da alocação global dos 15 fragmentos.
            if not isinstance(novo, dict):
                print(
                    "🔴 RECUPERAÇÃO DE EVIDÊNCIA BLOQUEADA:",
                    "nenhuma reserva Python previamente autorizada disponível."
                )
                return False

            novo = dict(novo)
            novo["numero"] = indice + 1
            novo["auditoria_python_final"] = "OK"
            novo["palavras"] = len(str(novo.get("texto", "") or "").split())

            # O conjunto enviado ao Ollama muda imediatamente.
            fragmentos_autorizados[indice] = novo

            # O estado oficial em memória também muda imediatamente.
            if isinstance(informacoes_relevantes, list) and len(informacoes_relevantes) >= 3:
                informacoes_relevantes[indice] = dict(novo)

            registro_reselecao = {
                "motivo": motivo,
                "paragrafo": indice + 1,
                "texto_novo": str(novo.get("texto", "") or ""),
                "id_novo": str(novo.get("id", "") or ""),
                "hash_novo": str(novo.get("hash", "") or ""),
                "fonte_nova": novo.get("fonte", ""),
            }
            historico_json = dados_bloco.get("historico_reselecao_ollama", []) if isinstance(dados_bloco, dict) else []
            if not isinstance(historico_json, list):
                historico_json = []
            historico_json.append(registro_reselecao)

            if isinstance(dados_bloco, dict):
                dados_bloco["informacoes_relevantes"] = informacoes_relevantes
                dados_bloco["fragmentos_autorizados"] = fragmentos_autorizados
                dados_bloco["historico_reselecao_ollama"] = historico_json

            if isinstance(blocos_informacoes.get(chave_bloco_atual), dict):
                blocos_informacoes[chave_bloco_atual]["informacoes_relevantes"] = list(informacoes_relevantes)
                blocos_informacoes[chave_bloco_atual]["fragmentos_autorizados"] = [
                    x.get("texto", "") if isinstance(x, dict) else str(x or "")
                    for x in fragmentos_autorizados
                ]
                blocos_informacoes[chave_bloco_atual]["historico_reselecao_ollama"] = list(historico_json)

            print()
            print("🔄 RECUPERAÇÃO ADAPTATIVA DE EVIDÊNCIA")
            print("BLOCO:", chave_bloco_atual)
            print("PARÁGRAFO SUBSTITUÍDO:", indice + 1)
            print("MOTIVO:", motivo)
            print("NOVO ID:", novo.get("id", "SEM_ID"))
            print("NOVO TAMANHO:", novo.get("palavras", 0), "palavras")
            print("RESERVAS RESTANTES:", len(fragmentos_reserva_por_bloco.get(chave_bloco_atual, [])))
            print("JSON EM MEMÓRIA: ATUALIZADO")

            return True


        # SEM LIMITE DE TENTATIVAS:
        # o bloco continua sendo gerado até passar pelas validações.
        tentativa_bloco = 0

        while True:
            tentativa_bloco += 1

            # Se a tentativa anterior falhou por evidência, editorial ou
            # progressão, substitui somente o trecho responsável antes de
            # chamar o Ollama novamente. Assim não reprocessamos texto ruim.
            if tentativa_bloco > 1 and ultimo_erro:
                erro_upper = str(ultimo_erro).upper()
                # Nem toda rejeição editorial exige trocar a evidência.
                # Vícios como "garantindo" são gerados pelo Ollama e devem
                # ser corrigidos na redação, sem gastar uma reserva Python.
                precisa_novo_texto = (
                    "EVIDÊNCIA FACTUAL INSUFICIENTE" in erro_upper
                    or "EVIDENCIA FACTUAL INSUFICIENTE" in erro_upper
                    or "NÚMERO/UNIDADE NÃO AUTORIZADO" in erro_upper
                    or "NUMERO/UNIDADE NÃO AUTORIZADO" in erro_upper
                    or "TITULO_OU_PERGUNTA_HERDADA" in erro_upper
                    or "RESIDUO_DE_CATEGORIA" in erro_upper
                    or "LINGUAGEM_COMERCIAL_OU_INSTITUCIONAL" in erro_upper
                    or "CABECALHO_EMBUTIDO" in erro_upper
                    or "ESTRUTURAL_NAVEGACAO_LISTA" in erro_upper
                    or "FINAL_TRUNCADO" in erro_upper
                    or "PROGRESSAO" in erro_upper
                )
                if precisa_novo_texto:
                    recuperou = _adaptar_evidencia_apos_rejeicao(ultimo_erro)
                    if not recuperou:
                        print(
                            "🔴 BLOCO INTERROMPIDO: não existe evidência "
                            "autorizada disponível para substituir o trecho rejeitado."
                        )
                        ultimo_erro = (
                            str(ultimo_erro)
                            + " | SEM_EVIDENCIA_AUTORIZADA_PARA_RESELECAO"
                        )
                        break

            print()
            print("=" * 60)
            print(
                f"OLLAMA — {chave_bloco_atual.upper()} — "
                f"TENTATIVA {tentativa_bloco} (SEM LIMITE)"
            )
            print("=" * 60)

            # ----------------------------------------------------
            # INSTRUÇÃO EXTRA PARA RETENTATIVA
            # ----------------------------------------------------

            instrucao_retentativa = ""

            if tentativa_bloco > 1:

                instrucao_retentativa = f"""

==================================================
CORREÇÃO DA TENTATIVA ANTERIOR
==================================================

A tentativa anterior foi rejeitada pelo Python.

MOTIVO DA REJEIÇÃO:
{ultimo_erro}

Se o motivo começar por "EDITORIAL:", trate-o como erro obrigatório de redação.
Não repita a expressão ou padrão que causou a rejeição. Reescreva especificamente
o parágrafo indicado, mantendo os demais fatos autorizados e sem introduzir fatos novos.

Gere novamente os três parágrafos.

ATENÇÃO:

- Corrija somente o problema informado.
- Não invente informações.
- NÃO RESUMA. Os três trechos autorizados continuam sendo a fonte integral desta tentativa.
- Preserve o conteúdo e a proporção natural dos trechos de origem, sem meta de quantidade de palavras.
- Nunca reduza o bloco apenas para deixar a redação mais curta.

- Não acrescente informações externas.
- Não acrescente números.
- Não acrescente características.
- Não acrescente aplicações.
- Não acrescente marcas.
- Não acrescente empresas.
- Não acrescente modelos.
- Os três trechos pertencem ao mesmo bloco e podem ser combinados.
- Reorganize as ideias quando isso melhorar a coerência.
- Ignore títulos, menus, legendas, CTAs, listas e trechos claramente truncados.
- Use somente fatos presentes nos três trechos deste bloco.

REGRA DE CORREÇÃO — PRIORIDADE ABSOLUTA NESTA RETENTATIVA:
- O Python rejeitou a resposta pelo motivo informado acima.
- Corrija especificamente o problema indicado no motivo.
- Se o motivo contiver "PADRÕES:", NÃO reutilize nenhuma das expressões listadas.
  Substitua-as por formulação técnica descritiva e neutra, sem sinônimo promocional.
- Preserve somente informações e relações já presentes nos trechos autorizados.
- Não altere os demais parágrafos se eles não apresentarem o problema indicado.
- Preserve o tamanho natural de cada trecho, sem transformar isso em regra numérica.
- Não aumente nem reduza artificialmente o texto apenas para atingir uma quantidade.
- Não aumente um trecho curto com fatos ou explicações que não estejam autorizados.
- A distribuição deve permanecer proporcional aos três trechos recebidos.
- O tamanho não é motivo de rejeição.

Retorne novamente somente:

[BLOCO]

[PARAGRAFO_1]
...
[/PARAGRAFO_1]

[PARAGRAFO_2]
...
[/PARAGRAFO_2]

[PARAGRAFO_3]
...
[/PARAGRAFO_3]

[/BLOCO]
"""

            # ----------------------------------------------------
            # AUTORIDADE DO CONJUNTO ATUAL
            # ----------------------------------------------------
            # Se houve troca adaptativa, os números do prompt-base original
            # não representam mais o conjunto enviado. Esta instrução final
            # é a autoridade para a tentativa corrente.
            tamanhos_tentativa_atual = [
                len(str(x.get("texto", "") or "").split())
                for x in fragmentos_autorizados
            ]
            total_tentativa_atual = sum(tamanhos_tentativa_atual)
            instrucao_retentativa += f"""

==================================================
CONJUNTO REAL DESTA TENTATIVA — AUTORIDADE FINAL
==================================================
Trecho 1: {tamanhos_tentativa_atual[0] if len(tamanhos_tentativa_atual)>0 else 0} palavras
Trecho 2: {tamanhos_tentativa_atual[1] if len(tamanhos_tentativa_atual)>1 else 0} palavras
Trecho 3: {tamanhos_tentativa_atual[2] if len(tamanhos_tentativa_atual)>2 else 0} palavras
TOTAL REAL ENVIADO NESTA TENTATIVA: {total_tentativa_atual} palavras

Ignore qualquer capacidade, tamanho ou alvo numérico anterior que contradiga
esses valores. Estes são os três trechos efetivamente enviados nesta tentativa.
Preserve proporcionalmente o conteúdo deles, sem meta ou limite de palavras.
"""

            # ----------------------------------------------------
            # MONTAR O PROMPT DESTA TENTATIVA
            # ----------------------------------------------------
            # A cada tentativa o marcador de trechos é reconstruído com
            # o conjunto atual. Assim, quando a quantidade falha, a próxima
            # tentativa realmente recebe trechos maiores ou menores.
            prompt_tentativa = str(prompt_base or "").replace(
                MARCADOR_TRECHOS_OLLAMA,
                _contexto_trechos_para_ollama(fragmentos_autorizados)
            )

            prompt_atual = (
                prompt_tentativa
                + instrucao_retentativa
            )

            # CHAMADA OLLAMA — UMA OPERAÇÃO POR BLOCO
            # ----------------------------------------------------

            inicio_ollama = time.time()
            estatisticas_ollama["chamadas"] += 1
            estatisticas_ollama["caracteres_prompt"] += len(prompt_atual)

            try:

                resposta = requests.post(

                    "http://localhost:11434/api/generate",

                    json={

                        "model":
                            "qwen2.5:3b",

                        "prompt":
                            prompt_atual,

                        "stream":
                            False,

                        "think":
                            False,

                        "keep_alive":
                            "10m",

                        "options": {

                            # Margem suficiente para 3 parágrafos + marcadores.
                            # 420 podia truncar a resposta antes dos fechamentos.
                            "num_predict":
                                720,

                            "num_ctx":
                                8192,

                            "temperature":
                                0.08
                                if tentativa_bloco > 1
                                else 0.18,

                            "top_p":
                                0.9,

                            "repeat_penalty":
                                1.05
                        }
                    },

                    timeout=(
                        30,
                        900
                    )
                )

            except requests.exceptions.Timeout:

                ultimo_erro = (
                    "timeout do Ollama após 900 segundos"
                )

                print(
                    "❌",
                    ultimo_erro
                )

                continue

            except requests.exceptions.ConnectionError as erro:

                ultimo_erro = (
                    "erro de conexão com Ollama: "
                    + repr(erro)
                )

                print(
                    "❌",
                    ultimo_erro
                )

                continue

            except Exception as erro:

                ultimo_erro = (
                    "erro na chamada Ollama: "
                    + repr(erro)
                )

                print(
                    "❌",
                    ultimo_erro
                )

                continue

            tempo_ollama = (
                time.time()
                - inicio_ollama
            )

            print(
                "STATUS HTTP:",
                resposta.status_code
            )

            print(
                "TEMPO:",
                round(
                    tempo_ollama,
                    2
                ),
                "segundos"
            )

            if resposta.status_code != 200:

                ultimo_erro = (
                    "Ollama retornou HTTP "
                    + str(resposta.status_code)
                )

                print(
                    "❌",
                    ultimo_erro
                )

                continue

            # ----------------------------------------------------
            # LER RESPOSTA
            # ----------------------------------------------------

            try:

                dados_ollama = (
                    resposta.json()
                )

                resultado_ollama = str(
                    dados_ollama.get(
                        "response",
                        ""
                    )
                    or ""
                ).strip()

            except Exception as erro:

                ultimo_erro = (
                    "erro ao interpretar JSON do Ollama: "
                    + repr(erro)
                )

                print(
                    "❌",
                    ultimo_erro
                )

                continue

            if not resultado_ollama:

                ultimo_erro = (
                    "Ollama retornou resposta vazia"
                )

                print(
                    "❌",
                    ultimo_erro
                )

                continue

            print(
                "CARACTERES RETORNADOS:",
                len(resultado_ollama)
            )

            print(
                "MOTIVO FINAL OLLAMA:",
                dados_ollama.get("done_reason", "não informado")
            )

            # ----------------------------------------------------
            # PARSER ROBUSTO
            # ----------------------------------------------------

            def _extrair_paragrafos_estrito(texto):

                """
                Extrai exatamente 3 parágrafos sem exigir que o modelo
                tenha fechado perfeitamente todos os marcadores.

                O Qwen pode terminar a resposta depois do conteúdo e
                antes de [/PARAGRAFO_N] ou [/BLOCO]. Quando o conteúdo
                está claramente delimitado pelo próximo marcador, fazemos
                a recuperação estrutural em vez de descartar a redação.
                """

                texto = str(texto or "").strip()

                if not texto:
                    return None, "resposta vazia"

                texto = re.sub(
                    r"```(?:text|txt)?",
                    "",
                    texto,
                    flags=re.IGNORECASE
                )
                texto = re.sub(r"```", "", texto).strip()

                texto_lower = texto.lower()
                extraidos = []

                for indice in range(1, 4):

                    abertura = f"[PARAGRAFO_{indice}]"
                    fechamento = f"[/PARAGRAFO_{indice}]"
                    abertura_lower = abertura.lower()
                    fechamento_lower = fechamento.lower()

                    pos_abertura = texto_lower.find(abertura_lower)

                    if pos_abertura == -1:
                        # Compatibilidade com respostas simples do tipo
                        # "PARAGRAFO 1:" sem os marcadores completos.
                        padrao_simples = re.search(
                            rf"(?:^|\n)\s*(?:PARÁGRAFO|PARAGRAFO)\s*{indice}\s*:\s*",
                            texto,
                            flags=re.IGNORECASE
                        )
                        if not padrao_simples:
                            return None, f"abertura ausente: {abertura}"
                        inicio_conteudo = padrao_simples.end()
                    else:
                        inicio_conteudo = pos_abertura + len(abertura)

                    # Primeiro fechamento explícito.
                    pos_fechamento = texto_lower.find(
                        fechamento_lower,
                        inicio_conteudo
                    )

                    # Se o fechamento faltou, usa o próximo marcador de
                    # parágrafo, [/BLOCO] ou o fim da resposta como limite.
                    candidatos = [
                        pos for pos in (
                            texto_lower.find(f"[paragrafo_{indice + 1}]", inicio_conteudo)
                            if indice < 3 else -1,
                            texto_lower.find("[/bloco]", inicio_conteudo),
                        )
                        if pos != -1
                    ]

                    fechamento_recuperado = False

                    if pos_fechamento == -1:
                        # Recupera respostas truncadas usando o próximo marcador,
                        # [/BLOCO] ou o fim da resposta como limite estrutural.
                        limites = []
                        if indice < 3:
                            proxima_abertura = f"[PARAGRAFO_{indice + 1}]"
                            pos_proxima = texto_lower.find(
                                proxima_abertura.lower(),
                                inicio_conteudo
                            )
                            if pos_proxima != -1:
                                limites.append(pos_proxima)
                        pos_bloco = texto_lower.find("[/bloco]", inicio_conteudo)
                        if pos_bloco != -1:
                            limites.append(pos_bloco)

                        if limites:
                            pos_fechamento = min(limites)
                            fechamento_recuperado = True
                        elif indice == 3 and inicio_conteudo < len(texto):
                            pos_fechamento = len(texto)
                            fechamento_recuperado = True
                        else:
                            return None, f"fechamento ausente: {fechamento}"

                    # O próximo parágrafo jamais pode aparecer dentro do atual.
                    if indice < 3:
                        proxima_abertura = f"[PARAGRAFO_{indice + 1}]"
                        pos_proxima = texto_lower.find(
                            proxima_abertura.lower(),
                            inicio_conteudo
                        )
                        if pos_proxima != -1 and pos_proxima < pos_fechamento:
                            pos_fechamento = pos_proxima
                            fechamento_recuperado = True

                    texto_paragrafo = re.sub(
                        r"\s+",
                        " ",
                        texto[inicio_conteudo:pos_fechamento].strip()
                    ).strip()

                    if not texto_paragrafo:
                        return None, f"parágrafo {indice} vazio"

                    extraidos.append(texto_paragrafo)

                    if fechamento_recuperado:
                        print(
                            f"⚠️ PARSER: fechamento de PARAGRAFO_{indice} "
                            "recuperado automaticamente."
                        )

                if "[/bloco]" not in texto_lower:
                    print("⚠️ PARSER: [/BLOCO] ausente; estrutura recuperada até o fim da resposta.")

                return extraidos, ""

            # ----------------------------------------------------
            # EXTRAIR
            # ----------------------------------------------------

            paragrafos_extraidos, erro_parser = (
                _extrair_paragrafos_estrito(
                    resultado_ollama
                )
            )

            if paragrafos_extraidos is None:

                ultimo_erro = (
                    "estrutura inválida: "
                    + str(erro_parser)
                )

                print(
                    "❌ PARSER:",
                    ultimo_erro
                )

                continue

            # ----------------------------------------------------
            # QUANTIDADE
            # ----------------------------------------------------

            if len(
                paragrafos_extraidos
            ) != 3:

                ultimo_erro = (
                    "quantidade inválida de parágrafos: "
                    + str(
                        len(paragrafos_extraidos)
                    )
                )

                print(
                    "❌",
                    ultimo_erro
                )

                continue

            # ----------------------------------------------------
            # VALIDAÇÃO ÚNICA DE BLOCO — SEM LIMITE DE PALAVRAS
            # ----------------------------------------------------
            ok_bloco, motivo_bloco, indice_ajuste = validar_bloco_editorial_unico(
                paragrafos_extraidos
            )

            quantidades_palavras = [len(str(p or "").split()) for p in paragrafos_extraidos]
            print(
                "TAMANHOS DO BLOCO:",
                quantidades_palavras,
                "| TOTAL:",
                sum(quantidades_palavras)
            )

            if not ok_bloco:
                ultimo_erro = motivo_bloco
                print("❌ VALIDAÇÃO DO BLOCO:", ultimo_erro)
                continue

            print("🟢 VALIDAÇÃO DO BLOCO: ACEITO — 3 PARÁGRAFOS / SEM LIMITE DE PALAVRAS")

            # ----------------------------------------------------
            # VALIDAÇÃO FACTUAL
            # ----------------------------------------------------

            palavras_ruido_validacao = {
                "para", "com", "sem", "sobre", "entre",
                "como", "uma", "umas", "um", "uns",
                "dos", "das", "que", "por", "pelo",
                "pela", "aos", "nas", "nos", "e", "ou",
                "de", "da", "do", "em", "no", "na",
                "ao", "se", "ser", "sao", "são", "foi",
                "tem", "ter", "pode", "podem", "mais",
                "tambem", "também", "isso", "esse",
                "essa", "este", "esta", "esses",
                "essas", "estes", "estas", "quando",
                "onde", "assim", "cada", "qual", "forma",
                "tipo", "modo", "parte", "caso", "seu",
                "sua", "seus", "suas", "outro",
                "outra", "outros", "outras"
            }

            def _tokens_validacao(texto):

                normalizado = (
                    normalizar_assunto_texto(
                        str(texto or "")
                    )
                )

                return [
                    token
                    for token in re.findall(
                        r"\b[a-z0-9]{3,}\b",
                        normalizado
                    )
                    if token
                    not in palavras_ruido_validacao
                ]

            def _radical_leve(token):

                token = str(
                    token or ""
                )

                if len(token) <= 5:
                    return token

                sufixos = (
                    "mente",
                    "ções",
                    "ção",
                    "sões",
                    "são",
                    "amentos",
                    "imentos",
                    "amento",
                    "imento",
                    "idades",
                    "idade",
                    "ismos",
                    "ismo",
                    "istas",
                    "ista",
                    "ados",
                    "idos",
                    "adas",
                    "idas",
                    "ando",
                    "endo",
                    "indo",
                    "es",
                    "as",
                    "os",
                    "s"
                )

                for sufixo in sufixos:

                    if (
                        token.endswith(sufixo)
                        and
                        len(token)
                        - len(sufixo)
                        >= 4
                    ):

                        return token[
                            :-
                            len(sufixo)
                        ]

                return token

            def _n_gramas(tokens, n=2):

                return set(
                    " ".join(
                        tokens[i:i+n]
                    )
                    for i in range(
                        max(
                            0,
                            len(tokens) - n + 1
                        )
                    )
                )

            erro_factual = None

            for indice_paragrafo, paragrafo in enumerate(
                paragrafos_extraidos
            ):

                fragmento_autorizado = str(
                    fragmentos_autorizados[
                        indice_paragrafo
                    ].get(
                        "texto",
                        ""
                    )
                    or ""
                )

                tokens_fonte = (
                    _tokens_validacao(
                        fragmento_autorizado
                    )
                )

                tokens_saida = (
                    _tokens_validacao(
                        paragrafo
                    )
                )

                radicais_fonte = {
                    _radical_leve(token)
                    for token
                    in tokens_fonte
                }

                radicais_saida = [
                    _radical_leve(token)
                    for token
                    in tokens_saida
                ]

                compartilhados = [
                    token
                    for token
                    in radicais_saida
                    if token
                    in radicais_fonte
                ]

                cobertura_radical = (
                    len(compartilhados)
                    /
                    max(
                        1,
                        len(radicais_saida)
                    )
                )

                termos_fonte = set(
                    tokens_fonte
                )

                termos_saida = set(
                    tokens_saida
                )

                termos_compartilhados = (
                    termos_fonte
                    &
                    termos_saida
                )

                cobertura_termos = (
                    len(
                        termos_compartilhados
                    )
                    /
                    max(
                        1,
                        len(termos_saida)
                    )
                )

                bigramas_fonte = (
                    _n_gramas(
                        tokens_fonte,
                        2
                    )
                )

                bigramas_saida = (
                    _n_gramas(
                        tokens_saida,
                        2
                    )
                )

                sobreposicao_bigrama = (
                    len(
                        bigramas_fonte
                        &
                        bigramas_saida
                    )
                    /
                    max(
                        1,
                        len(bigramas_saida)
                    )
                )

                score_evidencia = (
                    (
                        cobertura_radical
                        * 0.60
                    )
                    +
                    (
                        cobertura_termos
                        * 0.25
                    )
                    +
                    (
                        sobreposicao_bigrama
                        * 0.15
                    )
                )

                numeros_fonte = set(
                    re.findall(
                        r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)",
                        normalizar_assunto_texto(
                            fragmento_autorizado
                        )
                    )
                )

                numeros_saida = set(
                    re.findall(
                        r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)",
                        normalizar_assunto_texto(
                            paragrafo
                        )
                    )
                )

                numeros_novos = (
                    numeros_saida
                    -
                    numeros_fonte
                )

                if numeros_novos:

                    erro_factual = (
                        f"parágrafo "
                        f"{indice_paragrafo + 1} "
                        f"introduziu número/unidade "
                        f"não autorizado: "
                        f"{sorted(numeros_novos)}"
                    )

                    break

                if (
                    len(tokens_saida) >= 25
                    and
                    score_evidencia < 0.14
                ):

                    erro_factual = (
                        f"parágrafo "
                        f"{indice_paragrafo + 1} "
                        f"possui evidência factual "
                        f"insuficiente; "
                        f"score="
                        f"{round(score_evidencia, 3)}"
                    )

                    break

                print(
                    "🟢 FACTUAL:",
                    chave_bloco_atual,
                    "PARÁGRAFO",
                    indice_paragrafo + 1,
                    "SCORE",
                    round(
                        score_evidencia,
                        3
                    ),
                    "TERMOS",
                    len(
                        termos_compartilhados
                    ),
                    "NÚMEROS NOVOS",
                    0
                )

            if erro_factual:

                ultimo_erro = erro_factual

                print(
                    "❌ FACTUAL:",
                    ultimo_erro
                )

                continue

            # ----------------------------------------------------
            # VALIDAÇÃO EDITORIAL ESTRITA
            # ----------------------------------------------------
            # Factualidade não basta. O bloco também precisa estar limpo
            # editorialmente, sem CTA, propaganda, títulos residuais,
            # perguntas, legendas, marcas/empresas, absolutos comerciais,
            # repetição excessiva ou mistura incoerente.

            def _normalizar_editorial_local(txt):
                return normalizar_assunto_texto(str(txt or "")).strip()

            def _frases_editoriais_proibidas(txt):
                n = _normalizar_editorial_local(txt)
                padroes = [
                    r"\bentre em contato\b", r"\bfale conosco\b", r"\bsaiba mais\b",
                    r"\bsolicite (?:um )?(?:orcamento|cotacao)\b",
                    r"\bpeca (?:um )?(?:orcamento|cotacao)\b",
                    r"\bonde comprar\b", r"\bmelhor preco\b", r"\bpreco competitivo\b",
                    r"\bvendemos\b", r"\bcompre(?: agora)?\b", r"\badquira(?: agora)?\b",
                    r"\bmais de \d+ anos\b", r"\bdesde \d{4}\b",
                    r"\bunica escolha\b", r"\bmelhor escolha\b", r"\bsolucao definitiva\b",
                    r"\bclique aqui\b", r"\bentre em contacto\b"
                ]
                return [p for p in padroes if re.search(p, n, flags=re.I)]

            def _tem_residuo_estrutural(txt):
                n = _normalizar_editorial_local(txt)
                if "visualizacao isometrica" in n or "visualizacao" in n and "figura" in n:
                    return True
                if re.search(r"\b(?:o que e|como funciona|principais vantagens|caracteristicas|beneficios|especificacoes)\s*:", n):
                    return True
                if "?" in str(txt):
                    return True
                if re.search(r"(?:^|[;\n])\s*(?:paragrafo|trecho|fonte|referencia|url)\s*[:=-]", n):
                    return True
                if re.search(r"https?://|www\.|\b[\w.%+-]+@[\w.-]+\.[a-z]{2,}\b", n, re.I):
                    return True
                return False

            def _frases_absolutas(txt):
                n = _normalizar_editorial_local(txt)
                padroes = [
                    r"\bqualquer aplicacao\b",
                    r"\bgarante\s+(?:o|a|um|uma)\s+(?:resultado|desempenho|eficiencia|qualidade|seguranca)\b",
                    r"\bgarantindo\s+(?:o|a|um|uma)\s+(?:resultado|desempenho|eficiencia|qualidade|seguranca)\b",
                    r"\bmaximizar seus lucros\b", r"\bessencial para qualquer\b",
                    r"\bsolucao definitiva\b", r"\bescolha inteligente\b",
                    r"\bextremamente versatil", r"\bperfeita?s? para\b",
                    r"\bindispensavel\b", r"\bretorno sobre o investimento\b"
                ]
                encontrados = []
                for padrao in padroes:
                    for correspondencia in re.finditer(padrao, n):
                        encontrados.append(correspondencia.group(0))
                return encontrados

            # Cada parágrafo precisa ser autossuficiente, técnico e limpo.
            erro_editorial = None
            for idx_p, para in enumerate(paragrafos_extraidos, start=1):
                hits = _frases_editoriais_proibidas(para)
                if hits:
                    erro_editorial = f"parágrafo {idx_p}: linguagem comercial/CTA detectada"
                    break
                if _tem_residuo_estrutural(para):
                    erro_editorial = f"parágrafo {idx_p}: resíduo estrutural/legenda/pergunta detectado"
                    break
                abs_hits = _frases_absolutas(para)
                if abs_hits:
                    amostras = "; ".join(dict.fromkeys(abs_hits))[:300]
                    erro_editorial = (
                        f"parágrafo {idx_p}: afirmação absoluta/promocional detectada"
                        f" | PADRÕES: {amostras}"
                    )
                    break
                # Contexto empresarial neutro pode ser legítimo. Bloquear
                # apenas solicitação comercial/CTA explícita ou autopromoção clara.
                n_para = _normalizar_editorial_local(para)
                comercial_forte = (
                    r"\b(?:solicite|peca)\s+(?:um|uma)?\s*(?:orcamento|cotacao)\b",
                    r"\bentre\s+em\s+contato\b",
                    r"\bfale\s+conosco\b",
                    r"\b(?:compre|adquira)(?:\s+agora)?\b",
                    r"\b(?:nossa|nosso)\s+empresa\s+(?:oferece|fornece)\b",
                )
                if any(re.search(p, n_para) for p in comercial_forte):
                    erro_editorial = f"parágrafo {idx_p}: contexto empresarial/comercial inadequado ao conteúdo técnico"
                    break

            # Diversidade entre os três parágrafos: não aceitar cópia quase literal.
            if erro_editorial is None:
                def _tokens_sem_stop(txt):
                    n = _normalizar_editorial_local(txt)
                    stop = {"para","com","sem","sobre","entre","como","uma","um","dos","das","que","por","de","da","do","em","no","na","e","ou","a","o","os","as","se","mais","tambem","também","isso","essa","esse","esta","este"}
                    return set(x for x in re.findall(r"\b[a-z0-9]{4,}\b", n) if x not in stop)
                conjuntos = [_tokens_sem_stop(x) for x in paragrafos_extraidos]
                for i in range(3):
                    for j in range(i+1,3):
                        inter = len(conjuntos[i] & conjuntos[j]) / max(1, len(conjuntos[i] | conjuntos[j]))
                        if inter >= 0.90:
                            erro_editorial = f"parágrafos {i+1} e {j+1}: repetição semântica excessiva"
                            break
                    if erro_editorial:
                        break

            if erro_editorial:
                ultimo_erro = "EDITORIAL: " + erro_editorial
                print("❌", ultimo_erro)
                continue

            # ----------------------------------------------------
            # BLOCO APROVADO
            # ----------------------------------------------------

            print()
            print(
                "=" * 60
            )
            print(
                f"🟢 BLOCO APROVADO — "
                f"{chave_bloco_atual.upper()}"
            )
            print(
                f"TENTATIVA: {tentativa_bloco} (SEM LIMITE)"
            )
            print(
                "=" * 60
            )

            return (
                paragrafos_extraidos,
                ""
            )

        # --------------------------------------------------------
        # LIMITE DE SEGURANÇA
        # --------------------------------------------------------

        mensagem_final = (
            f"{chave_bloco_atual} teve a execução interrompida "
            f"após {tentativa_bloco} tentativas. "
            f"Último erro: {ultimo_erro}"
        )

        print()
        print("❌", mensagem_final)

        return None, mensagem_final

        
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

        if not isinstance(fragmentos_preflight, list) or len(fragmentos_preflight) != 3:
            print()
            print("=" * 60)
            print("❌ PRÉ-VALIDAÇÃO OLLAMA INTERROMPIDA")
            print("=" * 60)
            print(
                chave_preflight,
                "possui",
                len(fragmentos_preflight)
                if isinstance(fragmentos_preflight, list)
                else 0,
                "fragmentos; são necessários 3."
            )
            print(
                "Nenhuma chamada Ollama será feita."
            )
            print("=" * 60)
            return None

    # ============================================================
    # ENTIDADES PROIBIDAS DA PÁGINA
    # ============================================================
    # Todas as identidades empresariais encontradas nas fontes da página
    # ficam disponíveis apenas para auditoria. Elas NÃO são enviadas ao Ollama.
    ENTIDADES_PROIBIDAS_PAGINA.clear()
    for _bloco_ent in informacoes_blocos.values():
        if not isinstance(_bloco_ent, dict):
            continue
        for _frag_ent in _bloco_ent.get("informacoes_relevantes", []):
            if not isinstance(_frag_ent, dict):
                continue
            _id_ent = _frag_ent.get("identidade_fonte", {})
            if isinstance(_id_ent, dict):
                _nome_ent = str(_id_ent.get("nome", "") or "").strip()
                if len(_nome_ent) >= 4:
                    ENTIDADES_PROIBIDAS_PAGINA.add(_nome_ent)

    print(
        "ENTIDADES PROIBIDAS PARA AUDITORIA:",
        len(ENTIDADES_PROIBIDAS_PAGINA)
    )

    # ============================================================
    # PROCESSAR CADA BLOCO
    # ============================================================

    titulos_blocos_preflight = {}

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

        fragmentos_autorizados = (
            dados_bloco.get(
                "fragmentos_autorizados",
                []
            )
        )

        if not isinstance(
            fragmentos_autorizados,
            list
        ):

            fragmentos_autorizados = []


        fragmentos_autorizados = (
            fragmentos_autorizados
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
        # TÍTULO DO BLOCO
        # ========================================================

        titulo_bloco = titulos_python.get(
            chave_bloco,
            ""
        )

        titulo_bloco = str(
            titulo_bloco or ""
        ).strip()


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

                "numero":
                    numero_fragmento,

                "id":
                    id_fragmento,

                "texto":
                    texto_fragmento,

                "identidade_fonte":
                    fragmento.get(
                        "identidade_fonte",
                        {}
                    )

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

        for fragmento in fragmentos_bloco:

            print()
            print(
                f"FRAGMENTO {fragmento['numero']}/3"
            )

            print(
                "ID:",
                fragmento["id"]
            )

            print(
                "CARACTERES:",
                len(
                    fragmento["texto"]
                )
            )

            print(
                "PALAVRAS:",
                len(
                    fragmento["texto"].split()
                )
            )
            print("NOTA INDIVIDUAL:", fragmento.get("nota_trecho_otimo", ""))
            print("NOTA COMBINAÇÃO NARRATIVA:", fragmento.get("nota_combinacao_narrativa", ""))


        # ========================================================
        # TÍTULO OFICIAL DO MEAD — NÃO DERIVAR DOS FRAGMENTOS
        # ========================================================
        # Os fragmentos sustentam o conteúdo; eles não redefinem a função
        # editorial do bloco. O título já foi criado pelo Python/MEAD.
        titulo_bloco = str(
            titulos_python.get(chave_bloco, dados_bloco.get("titulo", ""))
            or ""
        ).strip()
        dados_bloco["titulo"] = titulo_bloco
        blocos_informacoes[chave_bloco]["titulo"] = titulo_bloco
        print("TÍTULO MEAD/PYTHON:", titulo_bloco)

        # --------------------------------------------------------
        # PREFLIGHT DE UNICIDADE DO TÍTULO — ANTES DO OLLAMA
        # --------------------------------------------------------
        # Se o título já foi usado, não desperdiçamos nenhuma chamada de IA.
        # Tentamos um título técnico alternativo baseado nos mesmos trechos.
        titulo_normalizado = normalizar_assunto_texto(titulo_bloco).strip()
        titulos_anteriores = {normalizar_assunto_texto(v).strip() for v in titulos_blocos_preflight.values()}
        if titulo_normalizado in titulos_anteriores:
            alternativas = [
                f"Características Técnicas e Operação de {tema}",
                f"Aplicações e Seleção de {tema}",
                f"Desempenho e Manutenção de {tema}",
                f"Funcionamento e Aplicações de {tema}",
                f"Aspectos Técnicos de {tema}",
            ]
            substituto = next((x for x in alternativas if normalizar_assunto_texto(x).strip() not in titulos_anteriores), None)
            if substituto:
                print("⚠️ TÍTULO DUPLICADO NO PREFLIGHT:", titulo_bloco)
                print("🟢 TÍTULO SUBSTITUÍDO:", substituto)
                titulo_bloco = substituto
            else:
                print("🔴 TÍTULO DUPLICADO SEM ALTERNATIVA:", titulo_bloco)
                return None

        titulos_blocos_preflight[chave_bloco] = titulo_bloco
        dados_bloco["titulo"] = titulo_bloco
        blocos_informacoes[chave_bloco]["titulo"] = titulo_bloco

        # ========================================================
        # NOVA ARQUITETURA — OLLAMA RECEBE 3 TRECHOS
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
        # Os fragmentos continuam sendo a única fonte factual.
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
        # - 3 trechos por bloco.
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
        # Ele recebe somente os três trechos autorizados.
        #
        # ========================================================
        
        # ========================================================
        # MATERIAL AUTORIZADO AO OLLAMA — v51
        # ========================================================
        # As três evidências 100/100 são enviadas integralmente.
        # Nenhum corte, redução ou faixa de palavras é aplicada.
        # ========================================================
        fragmentos_ollama = list(fragmentos_bloco)

        contexto_tres_fragmentos = ""

        for fragmento in fragmentos_ollama:

            contexto_tres_fragmentos += f"""
--------------------------------------------------
TRECHO {fragmento["numero"]}
--------------------------------------------------

{fragmento["texto"]}
"""

        # ========================================================
        # PROMPT FINAL DO BLOCO
        # ========================================================

        # ========================================================
        # PROMPT COMPACTO — PYTHON ENTREGA TEXTO PRONTO
        # ========================================================
        # O Python já fez pesquisa, limpeza, seleção e auditoria.
        # O Ollama NÃO deve resumir nem decidir o que é importante.
        # ========================================================

        tamanhos_entrada = [
            len(str(x.get("texto", "")).split())
            for x in fragmentos_ollama
        ]

        total_entrada_ollama = sum(tamanhos_entrada)

        prompt_bloco = f"""Você é um editor técnico de preservação factual.

O Python já entregou três trechos limpos e autorizados. Sua tarefa NÃO é resumir. Sua tarefa é reescrever com linguagem natural, preservando o conteúdo factual recebido.

REGRA CENTRAL — TRÊS EVIDÊNCIAS FORMAM UMA ÚNICA NARRATIVA:
- Não existe mínimo ou máximo de palavras.
- Os três trechos são EVIDÊNCIAS COMPLEMENTARES do mesmo bloco; não são três parágrafos independentes.
- Construa exatamente 3 parágrafos desenvolvidos usando o conjunto das três evidências.
- Organize as ideias em progressão natural conforme o que as evidências realmente sustentam.
- Um parágrafo pode combinar informações de mais de uma evidência. Não faça TRECHO 1 = PARÁGRAFO 1 por obrigação.
- Não repita a mesma informação apenas porque ela aparece em mais de uma evidência.
- Cada parágrafo deve acrescentar ou desenvolver uma ideia relevante em relação ao anterior.
- Preserve exemplos, condições, aplicações, características, critérios e relações técnicas presentes nas evidências.
- Conecte evidências complementares naturalmente em vez de tratá-las como três resumos isolados.

REGRAS DE REDAÇÃO:
- Escreva como conteúdo técnico de site industrial, natural e desenvolvido, semelhante aos modelos de referência fornecidos.
- Use conectivos e anáforas naturais para evitar repetir o protagonista no início de todos os parágrafos.
- O protagonista pode reaparecer quando necessário para clareza, mas não use uma fórmula mecânica.
- Corrija clareza, fluidez, naturalidade, ordem das ideias e repetições evidentes.
- Não pesquise.
- Não invente fatos, números, normas, materiais, aplicações, marcas, empresas, modelos ou características.
- Não crie informação para preencher lacunas entre as evidências.
- Não crie introdução, conclusão genérica, CTA, propaganda ou lista.
- Não use linguagem absoluta ou promocional, especialmente: garante, garantindo, sempre, nunca, desempenho superior, solução definitiva, escolha inteligente, extremamente versátil e indispensável.
- Informação comercial factual pode ser preservada somente quando estiver nas evidências e estiver autorizada pelo bloco.
- Não transforme informação técnica em propaganda.

CONTROLE DE VOLUME FLEXÍVEL:
- NÃO existe faixa obrigatória de palavras.
- O tamanho final deve ser consequência da quantidade de informação factual necessária para desenvolver a narrativa.
- Não corte informação boa apenas para diminuir o texto.
- Não invente conteúdo para aumentar o texto.
- Um bloco pode ter parágrafos de tamanhos diferentes quando a progressão exigir.
- Preserve o nível de desenvolvimento dos modelos de referência: texto explicado, conectado e natural, não uma sequência de resumos.
- Antes de finalizar, confira se os três parágrafos, juntos, aproveitam as três evidências sem repetição e sem introduzir fatos externos.

FORMATO OBRIGATÓRIO:
[BLOCO]
[PARAGRAFO_1]
texto
[/PARAGRAFO_1]
[PARAGRAFO_2]
texto
[/PARAGRAFO_2]
[PARAGRAFO_3]
texto
[/PARAGRAFO_3]
[/BLOCO]

EVIDÊNCIAS AUTORIZADAS PELO PYTHON — USE AS TRÊS EM CONJUNTO:

EVIDÊNCIA 1:
{fragmentos_ollama[0]["texto"] if len(fragmentos_bloco)>0 else ""}

EVIDÊNCIA 2:
{fragmentos_ollama[1]["texto"] if len(fragmentos_bloco)>1 else ""}

EVIDÊNCIA 3:
{fragmentos_ollama[2]["texto"] if len(fragmentos_bloco)>2 else ""}
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
            len(fragmentos_ollama)
        )

        print(
            "IDENTIDADE DA FONTE ENVIADA:",
            "NÃO"
        )

        print(
            "FONTES EXTERNAS ENVIADAS:",
            "NÃO"
        )
        print("MODELO DE SELEÇÃO:", "3 EVIDÊNCIAS INDIVIDUAIS 100/100 + COMBINAÇÃO NARRATIVA 100/100")

        print(
            "MAPA MEAD ENVIADO:",
            "NÃO"
        )

        # Cronometragem do bloco inteiro.
        # inicio_ollama/fim_ollama pertencem à função de retentativas
        # e não existem neste escopo.
        inicio_bloco = time.time()

        resultado_bloco = None
        erro_bloco = ""

        # Blindagem de estado: uma rejeição do bloco 4 (ou de qualquer
        # outro bloco) nunca pode causar NameError durante a re-seleção.
        if not isinstance(hashes_selecionados, set):
            hashes_selecionados = set(hashes_selecionados or [])
        if not isinstance(fontes_utilizadas, dict):
            fontes_utilizadas = {}
        if not isinstance(fragmentos_reserva_por_bloco, dict):
            fragmentos_reserva_por_bloco = {}
        fragmentos_reserva_por_bloco.setdefault(chave_bloco, [])

        # ========================================================
        # UMA OPERAÇÃO OLLAMA POR BLOCO
        # ========================================================
        # Os três fragmentos autorizados são enviados juntos. O Ollama
        # devolve exatamente três parágrafos. Uma retentativa reenvia o
        # bloco inteiro; não existe mais estado de aprovação individual.
        # ========================================================

        resultado_bloco, erro_bloco = _processar_bloco_ollama_com_retentativas(
            prompt_bloco,
            fragmentos_bloco,
            chave_bloco
        )

        if resultado_bloco is None:
            print()
            print("=" * 60)
            print("❌ BLOCO NÃO APROVADO PELO OLLAMA")
            print("=" * 60)
            print("BLOCO:", chave_bloco)
            print("MOTIVO FINAL:", erro_bloco)
            print("REGRAS: 1 operação por tentativa / 3 fragmentos por operação")
            print("=" * 60)
            return None

        fim_bloco = time.time()

        paragrafos_extraidos = resultado_bloco
        
        print()
        print(
            "🟢 BLOCO APROVADO PELA RETENTATIVA CONTROLADA"
        )
        
        print(
            "BLOCO:",
            chave_bloco
        )
        
        print(
            "PARÁGRAFOS:",
            len(paragrafos_extraidos)
        )
        

        # ========================================================
        # VALIDAÇÃO FACTUAL PÓS-OLLAMA — 
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
        # Cada parágrafo é validado contra o conjunto dos três
        # fragmentos autorizados deste bloco, permitindo redação
        # editorial integrada sem liberar informação externa.
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

        # A redação pode combinar os três fragmentos do mesmo bloco.
        # Portanto, a evidência factual é calculada contra o conjunto
        # autorizado do bloco, e não contra um fragmento específico.
        texto_fonte_bloco = "\n".join(
            str(fragmento.get("texto", "") or "")
            for fragmento in fragmentos_bloco
        )

        tokens_fonte_bloco = _tokens_validacao(texto_fonte_bloco)
        radicais_fonte_bloco = {_radical_leve(t) for t in tokens_fonte_bloco}
        termos_fonte_bloco = set(tokens_fonte_bloco)
        bigramas_fonte_bloco = _n_gramas(tokens_fonte_bloco, 2)

        erro_validacao_factual = None

        for indice_paragrafo, resultado_paragrafo in enumerate(paragrafos_extraidos):
            tokens_fonte = tokens_fonte_bloco
            tokens_saida = _tokens_validacao(resultado_paragrafo)

            radicais_fonte = radicais_fonte_bloco
            radicais_saida = [_radical_leve(t) for t in tokens_saida]
            compartilhados = [t for t in radicais_saida if t in radicais_fonte]

            cobertura_radical = (
                len(compartilhados) / max(1, len(radicais_saida))
            )

            termos_fonte = termos_fonte_bloco
            termos_saida = set(tokens_saida)
            termos_compartilhados = termos_fonte & termos_saida
            cobertura_termos = (
                len(termos_compartilhados) / max(1, len(termos_saida))
            )

            bigramas_fonte = bigramas_fonte_bloco
            bigramas_saida = _n_gramas(tokens_saida, 2)
            sobreposicao_bigrama = (
                len(bigramas_fonte & bigramas_saida)
                / max(1, len(bigramas_saida))
            )

            # Evidência suficiente pode vir de paráfrase: usamos
            # radical + termos + pequenas sequências, sem exigir
            # identidade textual.
            score_evidencia = (
                (cobertura_radical * 0.60)
                + (cobertura_termos * 0.25)
                + (sobreposicao_bigrama * 0.15)
            )

            # Números e unidades são fatos de alta precisão: nenhum
            # valor novo pode aparecer na redação.
            numeros_fonte = set(re.findall(
                r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)",
                normalizar_assunto_texto(texto_fonte_bloco)
            ))
            numeros_saida = set(re.findall(
                r"(?<!\w)\d+(?:[.,]\d+)?(?:%|[a-zA-Z]{1,8})?(?!\w)",
                normalizar_assunto_texto(resultado_paragrafo)
            ))

            numeros_novos = numeros_saida - numeros_fonte

            if numeros_novos:
                print("❌ VALIDAÇÃO FACTUAL: número/unidade não autorizado.")
                print("BLOCO:", chave_bloco, "PARÁGRAFO:", indice_paragrafo + 1)
                print("NÚMEROS NOVOS:", sorted(numeros_novos))
                erro_validacao_factual = (
                    f"parágrafo {indice_paragrafo + 1} contém número/unidade não autorizado: "
                    f"{sorted(numeros_novos)}"
                )
                break

            # Parágrafos muito curtos não precisam atingir um limite
            # artificial de sobreposição. Para textos com 25+ termos,
            # porém, exigimos evidência lexical/parafrástica mínima.
            if len(tokens_saida) >= 25 and score_evidencia < 0.14:
                print("❌ VALIDAÇÃO FACTUAL: evidência insuficiente após paráfrase.")
                print(
                    "BLOCO:", chave_bloco,
                    "PARÁGRAFO:", indice_paragrafo + 1,
                    "SCORE:", round(score_evidencia, 3),
                    "RADICAL:", round(cobertura_radical, 3),
                    "TERMOS:", round(cobertura_termos, 3),
                    "BIGRAMAS:", round(sobreposicao_bigrama, 3)
                )
                erro_validacao_factual = (
                    f"parágrafo {indice_paragrafo + 1} possui evidência factual insuficiente "
                    f"(score {score_evidencia:.3f})"
                )
                break

            print(
                "🟢 VALIDAÇÃO FACTUAL:",
                chave_bloco,
                "PARÁGRAFO", indice_paragrafo + 1,
                "SCORE", round(score_evidencia, 3),
                "TERMOS", len(termos_compartilhados),
                "NÚMEROS NOVOS", 0
            )

        if erro_validacao_factual:
            ultimo_erro = erro_validacao_factual
            continue

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
            "fragmentos_autorizados"
        ] = fragmentos_autorizados


        dados_bloco[
            "paragrafos_ollama"
        ] = paragrafos_ollama


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
            "fragmentos_autorizados"
        ] = fragmentos_autorizados


        dados_bloco[
            "paragrafos_ollama"
        ] = paragrafos_ollama


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
                fim_bloco - inicio_bloco,
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


        # ========================================================
        # ATUALIZAR O DICIONÁRIO DO BLOCO
        # ========================================================

        informacoes_blocos[
            chave_bloco
        ] = dados_bloco


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
    # GATE FINAL DE QUALIDADE — ANTES DE QUALQUER GRAVAÇÃO
    # ============================================================
    print()
    print("=" * 60)
    print("GATE FINAL — QUALIDADE EDITORIAL PYTHON")
    print("=" * 60)

    qualidade_pagina = []
    qualidade_reprovada = False

    for numero_bloco in range(1, total_blocos + 1):
        chave_q = f"bloco_{numero_bloco}"
        bloco_q = informacoes_blocos.get(chave_q, {})
        if not isinstance(bloco_q, dict):
            print("🔴", chave_q, "— estrutura inválida")
            qualidade_reprovada = True
            continue

        titulo_q = str(bloco_q.get("titulo", titulos_python.get(chave_q, "")) or "").strip()
        paras_q = bloco_q.get("paragrafos_ollama", [])
        frags_q = bloco_q.get("informacoes_relevantes", [])

        if chave_q in BLOCOS_APROVADOS_NO_LIMITE:
            aprovado_q = True
            nota_q = 8.5
            falhas_q = []
            detalhes_q = {
                "cobertura_factual_lexical": "CONTINGENCIA_8_DE_8",
                "generalidades": 0,
                "numeros_novos": [],
                "descontos": 0,
            }
            print(f"⚠️ {chave_q}: CONTINGÊNCIA 8/8 — GATE FINAL FLEXIBILIZADO")
        else:
            aprovado_q, nota_q, falhas_q, detalhes_q = avaliar_qualidade_final_python(
                titulo_q,
                paras_q,
                frags_q,
                chave_q,
                tema
            )

        qualidade_pagina.append({
            "bloco": chave_q,
            "nota": nota_q,
            "aprovado": aprovado_q,
            "falhas": falhas_q,
            "detalhes": detalhes_q
        })

        status_q = "🟢 APROVADO" if aprovado_q else "🔴 REPROVADO"
        print(f"{status_q} — {chave_q} — QUALIDADE: {nota_q:.1f}/10")
        if falhas_q:
            print("   FALHAS CRÍTICAS:", ", ".join(falhas_q))
        print("   COBERTURA FACTUAL:", detalhes_q.get("cobertura_factual_lexical"))

        if not aprovado_q:
            qualidade_reprovada = True

    nota_pagina = min((x["nota"] for x in qualidade_pagina), default=0.0)
    media_pagina = sum(x["nota"] for x in qualidade_pagina) / max(1, len(qualidade_pagina))

    print("=" * 60)
    print(f"QUALIDADE MÍNIMA ENTRE BLOCOS: {nota_pagina:.1f}/10")
    print(f"QUALIDADE MÉDIA DA PÁGINA: {media_pagina:.1f}/10")
    print("MÍNIMO OBRIGATÓRIO: 8.5/10")

    if qualidade_reprovada or nota_pagina < 8.5:
        print("🔴 GRAVAÇÃO NÃO AUTORIZADA")
        print("MOTIVO: QUALIDADE EDITORIAL OU GATE CRÍTICO REPROVADO")
        print("OLLAMA: processamento concluído, mas JSON oficial NÃO será gravado.")
        print("=" * 60)
        return None

    print("🟢 QUALIDADE EDITORIAL: APROVADA")
    print("=" * 60)

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

            "fragmentos_autorizados":
                dados_bloco.get(
                    "fragmentos_autorizados",
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
    
    print(
        "CHAMADAS OLLAMA REALIZADAS:",
        estatisticas_ollama["chamadas"]
    )
    print(
        "CARACTERES DE PROMPT REALMENTE ENVIADOS AO OLLAMA:",
        estatisticas_ollama["caracteres_prompt"]
    )
    estatisticas_ollama["fragmentos_aprovados"] = (
        sum(
            len([p for p in bloco.get("paragrafos_ollama", []) if str(p or "").strip()])
            for bloco in blocos
        )
    )
    print(
        "FRAGMENTOS APROVADOS PELO OLLAMA:",
        estatisticas_ollama["fragmentos_aprovados"]
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
    # O Python possui um banco fixo com mais de 20
    # possibilidades e sorteia exatamente 12 para a página.
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
            tema_base
        )
        for tag in tags_fixas
    ]

    # --------------------------------------------------------
    # Variações adicionais do Python
    # Servem apenas para completar as 30 tags quando necessário.
    # --------------------------------------------------------

    if tipo == "servico":

        tags_base_adicionais = [
            f"{tema_base} industrial",
            f"{tema_base} técnica",
            f"{tema_base} especializada",
            f"{tema_base} preventiva",
            f"{tema_base} corretiva",
            f"{tema_base} profissional",
            f"{tema_base} em equipamentos",
            f"{tema_base} em sistemas",
            f"{tema_base} em processos",
            f"serviços de {tema_base}",
            f"assistência em {tema_base}",
            f"diagnóstico de {tema_base}",
            f"inspeção de {tema_base}",
            f"suporte técnico {tema_base}",
            f"consultoria em {tema_base}"
        ]

    else:

        tags_base_adicionais = [
            f"{tema_base} industrial",
            f"{tema_base} profissional",
            f"{tema_base} técnica",
            f"{tema_base} especializada",
            f"{tema_base} para indústria",
            f"{tema_base} para sistemas",
            f"{tema_base} para processos",
            f"{tema_base} hidráulica",
            f"{tema_base} industrial hidráulica",
            f"aplicações de {tema_base}",
            f"uso de {tema_base}",
            f"soluções com {tema_base}",
            f"equipamento {tema_base}",
            f"sistema com {tema_base}",
            f"assistência técnica {tema_base}"
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
        )

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
    print("======================================")
    
    for bloco in blocos:
    
        print()
        print(
            "BLOCO:",
            bloco.get("numero")
        )
    
        print(
            "TÍTULO:",
            bloco.get("titulo", "")
        )
    
        paragrafos_teste = bloco.get(
            "paragrafos_ollama",
            []
        )
    
        print(
            "QUANTIDADE DE PARÁGRAFOS:",
            len(paragrafos_teste)
        )
    
        for indice, paragrafo in enumerate(
            paragrafos_teste,
            start=1
        ):
    
            print(
                f"PARÁGRAFO {indice}:",
                str(paragrafo)[:500]
            )

    # ========================================================
    # 23. ENVIAR ESTRUTURA OFICIAL PARA SALVAR_BANCO
    # ========================================================
    
    print()
    print("======================================")
    print("SALVANDO ESTRUTURA OFICIAL NO BANCO")
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

    """Valida o artefato MAPA_MEAD_PYTHON pela estrutura atual do mapa.

    Este mapa não é a página final e, portanto, não deve exigir os campos
    legados PROBLEMA/SOLUCAO/CENARIO. A estrutura atual usa contexto_blocos.
    """

    if not texto:
        return False

    try:
        mapa = json.loads(texto) if isinstance(texto, str) else texto
    except Exception:
        print()
        print("==============================")
        print("MAPA MEAD INVALIDO")
        print("JSON NAO PODE SER INTERPRETADO")
        print("==============================")
        return False

    if not isinstance(mapa, dict):
        print()
        print("==============================")
        print("MAPA MEAD INVALIDO")
        print("ESTRUTURA NAO E OBJETO JSON")
        print("==============================")
        return False

    obrigatorios = [
        "tipo",
        "tema",
        "protagonista",
        "tipo_pagina",
        "grupo",
        "assuntos_autorizados",
        "checkboxes_editoriais",
        "intencao_comercial",
        "regra_factual",
        "contexto_blocos"
    ]

    faltando = [
        chave for chave in obrigatorios
        if chave not in mapa
    ]

    if faltando:
        print()
        print("==============================")
        print("MAPA MEAD FALTANDO:")
        print(", ".join(faltando))
        print("==============================")
        return False

    if mapa.get("tipo") != "MAPA_MEAD_PYTHON":
        print()
        print("==============================")
        print("MAPA MEAD INVALIDO")
        print("TIPO:", mapa.get("tipo"))
        print("==============================")
        return False

    if not str(mapa.get("tema", "")).strip():
        print("🔴 MAPA MEAD INVALIDO: tema vazio.")
        return False

    if not str(mapa.get("protagonista", "")).strip():
        print("🔴 MAPA MEAD INVALIDO: protagonista vazio.")
        return False

    contexto = mapa.get("contexto_blocos")
    if not isinstance(contexto, dict):
        print("🔴 MAPA MEAD INVALIDO: contexto_blocos não é objeto.")
        return False

    blocos_faltando = [
        f"bloco_{i}" for i in range(1, 6)
        if f"bloco_{i}" not in contexto
    ]

    if blocos_faltando:
        print()
        print("==============================")
        print("MAPA MEAD FALTANDO CONTEXTO:")
        print(", ".join(blocos_faltando))
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

    if total_segmentos < 12:

        print()
        print(
            "❌ PÁGINA REPROVADA"
        )

        print(
            "Motivo: menos de 12 segmentos."
        )

        return False

    print()
    print(
        "✅ ESTRUTURA APROVADA"
    )

    print(
        "15 PARÁGRAFOS + 12 SEGMENTOS"
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
        f"desempenho e confiabilidade em {tema}"
    )

    # ========================================================
    # 04. SUBTÍTULO DOS SEGMENTOS
    #
    # Segundo subtítulo.
    # Também criado exclusivamente pelo Python.
    # ========================================================

    subtitulo_segmentos = gerar_subtitulo_segmentos(
        tema
    )

    # ========================================================
    # 05. TÍTULOS DOS 5 BLOCOS
    # ========================================================

    # Títulos são definidos pelo MEAD, não pelos assuntos que por acaso
    # aparecem nos três fragmentos escolhidos. Isso evita que um fragmento
    # de pressão/vazão transforme o B1 em um bloco técnico e que um B4
    # de conhecimento técnico receba um título de "aplicações".
    bloco_1 = (
        f"Contexto, importância e funcionamento de {tema}"
    )

    bloco_2 = (
        f"Características, funcionamento e aplicações de {tema}"
    )

    bloco_3 = (
        f"Critérios técnicos e cuidados com {tema}"
    )

    bloco_4 = (
        f"Conhecimento técnico e suporte para {tema}"
    )

    bloco_5 = (
        f"Seleção, aplicação e solução para {tema}"
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
    Persiste a página no formato oficial do conteudo-site.json.

    REGRAS:

    - Uma única estrutura oficial por tema.
    - Uma nova execução do mesmo tema substitui a versão anterior.
    - Os cinco blocos são preservados entre chamadas.
    - Python mantém os trechos selecionados em
      informacoes_relevantes e fragmentos_autorizados.
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
        fragmentos_autorizados
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

    banco = carregar_banco()

    if not isinstance(
        banco,
        dict
    ):
        banco = {}

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

    # Preservar imagens já existentes para que o gerador de DOCX possa
    # reutilizá-las. Não criar imagens novas nem baixá-las.
    imagens_existentes = pagina.get("imagens")
    if not isinstance(imagens_existentes, dict):
        imagens_existentes = dados_tema.get("imagens")
    if not isinstance(imagens_existentes, dict):
        imagens_existentes = informacoes_adicionais.get("imagens")
    if not isinstance(imagens_existentes, dict):
        imagens_existentes = {}

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

            "fragmentos_autorizados":
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

        if "fragmentos_autorizados" in dados_bloco:

            fragmentos_autorizados = (
                dados_bloco.get(
                    "fragmentos_autorizados"
                )
            )

            if isinstance(
                fragmentos_autorizados,
                list
            ):

                alvo[
                    "fragmentos_autorizados"
                ] = (
                    list(
                        fragmentos_autorizados[:3]
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

        fragmentos_autorizados = alvo.get(
            "fragmentos_autorizados",
            []
        )

        if not isinstance(
            fragmentos_autorizados,
            list
        ):
            fragmentos_autorizados = []

        alvo[
            "fragmentos_autorizados"
        ] = (
            list(
                fragmentos_autorizados[:3]
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
    
    nome_site_recebido = str(nome_site or "").strip()

    if nome_site_recebido:
        pagina["nome_site"] = nome_site_recebido
    elif "nome_site" not in pagina:
        pagina["nome_site"] = ""



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

        fragmentos_autorizados = bloco.get(
            "fragmentos_autorizados",
            []
        )

        if not isinstance(
            fragmentos_autorizados,
            list
        ):

            fragmentos_autorizados = []

        bloco[
            "fragmentos_autorizados"
        ] = (
            list(
                fragmentos_autorizados[:3]
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

    if imagens_existentes:
        pagina["imagens"] = imagens_existentes

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

    if imagens_existentes:
        pagina_final["imagens"] = imagens_existentes

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

        # Editorial e estrutural são barreiras diferentes. Nenhum JSON
        # oficial pode existir enquanto tamanho, integridade, lixo ou
        # duplicação editorial estiverem pendentes.
        ok_editorial_final, motivo_editorial_final = validar_pagina_editorial_final(
            pagina_final,
            tema_original
        )

        if not ok_editorial_final:
            print("\n🔴 GRAVAÇÃO NÃO AUTORIZADA — VALIDAÇÃO EDITORIAL FALHOU")
            print("MOTIVO:", motivo_editorial_final)
            print("A página permanece somente em memória; JSON oficial não será substituído.")
            return False

        print("\n🟢 VALIDAÇÃO EDITORIAL: OK")

        erros_json = []

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
            if "paragrafos" in bloco_validacao:
                erros_json.append(f"{chave_bloco}: campo legado 'paragrafos' não permitido")

        if erros_json:
            print("\n❌ JSON NÃO GRAVADO — ESTRUTURA INCOMPLETA")
            for erro_json in erros_json:
                print("-", erro_json)
            return False

        # ====================================================
        # INTEGRIDADE EXATA MEMÓRIA → TEMP → DISCO — v7.7
        # ====================================================
        #
        # A versão anterior conferia apenas quantidade e campos
        # não vazios depois do os.replace(). Agora:
        # 1) monta uma fotografia canônica dos 15 fragmentos e
        #    15 parágrafos aprovados em memória;
        # 2) grava em arquivo temporário;
        # 3) reabre o temporário e compara conteúdo exato;
        # 4) compara também os 5 hashes dos blocos;
        # 5) somente então substitui o JSON oficial;
        # 6) reabre o oficial e repete a comparação.
        # ====================================================

        def _canonico_json(valor):
            return json.dumps(
                valor,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":")
            )

        def _hash_integridade_bloco(bloco):
            if not isinstance(bloco, dict):
                return ""
            info_hash = _canonico_json(
                bloco.get("informacoes_relevantes", [])
            )
            conteudo_hash = (
                info_hash
                + "|"
                + str(bloco.get("titulo", "") or "").strip()
                + "|"
                + "|".join(
                    str(item or "").strip()
                    for item in bloco.get("paragrafos_ollama", [])[:3]
                )
            )
            return hashlib.sha256(
                conteudo_hash.encode("utf-8")
            ).hexdigest()

        def _snapshot_integridade(pagina_snapshot):
            snapshot = {}
            erros = []
            for numero in range(1, 6):
                chave_bloco = f"bloco_{numero}"
                bloco = pagina_snapshot.get(chave_bloco) if isinstance(pagina_snapshot, dict) else None
                if not isinstance(bloco, dict):
                    erros.append(f"{chave_bloco}: bloco ausente")
                    continue
                info = bloco.get("informacoes_relevantes", [])
                para = bloco.get("paragrafos_ollama", [])
                if not isinstance(info, list) or len(info) != 3:
                    erros.append(f"{chave_bloco}: fragmentos != 3")
                    continue
                if not isinstance(para, list) or len(para) != 3 or any(not str(x or "").strip() for x in para[:3]):
                    erros.append(f"{chave_bloco}: parágrafos Ollama != 3 válidos")
                    continue
                snapshot[chave_bloco] = {
                    "informacoes_relevantes": info[:3],
                    "paragrafos_ollama": [str(x or "").strip() for x in para[:3]],
                    "hash": _hash_integridade_bloco(bloco)
                }
            return snapshot, erros

        snapshot_memoria, erros_memoria = _snapshot_integridade(pagina_final)

        if erros_memoria:
            print("\n❌ INTEGRIDADE EM MEMÓRIA REJEITADA")
            for erro in erros_memoria:
                print("-", erro)
            return False

        def _comparar_snapshot(snapshot_esperado, pagina_obtida):
            erros = []
            snapshot_obtido, erros_estrutura = _snapshot_integridade(pagina_obtida)
            erros.extend(erros_estrutura)

            for numero in range(1, 6):
                chave_bloco = f"bloco_{numero}"
                esperado = snapshot_esperado.get(chave_bloco)
                obtido = snapshot_obtido.get(chave_bloco)
                if esperado is None or obtido is None:
                    continue

                if _canonico_json(esperado["informacoes_relevantes"]) != _canonico_json(obtido["informacoes_relevantes"]):
                    erros.append(f"{chave_bloco}: fragmentos memória × JSON diferentes")

                if esperado["paragrafos_ollama"] != obtido["paragrafos_ollama"]:
                    erros.append(f"{chave_bloco}: parágrafos Ollama memória × JSON diferentes")

                if esperado["hash"] != obtido["hash"]:
                    erros.append(f"{chave_bloco}: hash memória × JSON diferente")

            return erros, snapshot_obtido

        arquivo_temporario = None
        try:
            fd, arquivo_temporario = tempfile.mkstemp(
                prefix="conteudo-site-",
                suffix=".tmp",
                dir=diretorio or None,
                text=True
            )
            with os.fdopen(fd, "w", encoding="utf-8") as arquivo:
                json.dump(banco, arquivo, ensure_ascii=False, indent=4)
                arquivo.flush()
                os.fsync(arquivo.fileno())

            with open(arquivo_temporario, "r", encoding="utf-8") as arquivo_temporario_leitura:
                banco_temporario = json.load(arquivo_temporario_leitura)

            tema_temporario = banco_temporario.get(chave_existente)
            pagina_temporaria = (
                tema_temporario.get("pagina", {})
                if isinstance(tema_temporario, dict)
                else {}
            )

            erros_temporario, snapshot_temporario = _comparar_snapshot(
                snapshot_memoria,
                pagina_temporaria
            )

            if erros_temporario:
                print("\n❌ GRAVAÇÃO ABORTADA — TEMPORÁRIO DIFERE DA MEMÓRIA")
                for erro in erros_temporario:
                    print("-", erro)
                return False

            # Só agora o arquivo oficial pode ser substituído.
            os.replace(arquivo_temporario, ARQUIVO_BANCO)
            arquivo_temporario = None

        finally:
            if arquivo_temporario and os.path.exists(arquivo_temporario):
                try:
                    os.remove(arquivo_temporario)
                except Exception:
                    pass

        # ----------------------------------------------------
        # SEGUNDA CONFERÊNCIA: ARQUIVO OFICIAL REABERTO
        # ----------------------------------------------------

        with open(ARQUIVO_BANCO, "r", encoding="utf-8") as arquivo_verificacao:
            banco_gravado = json.load(arquivo_verificacao)

        tema_gravado = banco_gravado.get(chave_existente)
        pagina_gravada = (
            tema_gravado.get("pagina", {})
            if isinstance(tema_gravado, dict)
            else {}
        )

        erros_pos_gravacao, snapshot_disco = _comparar_snapshot(
            snapshot_memoria,
            pagina_gravada
        )

        if erros_pos_gravacao:
            print("\n❌ GRAVAÇÃO REJEITADA — CONTEÚDO NO DISCO DIFERE DA MEMÓRIA")
            for erro_pos in erros_pos_gravacao:
                print("-", erro_pos)
            return False

        print("\n============================================================")
        print("✅ VALIDAÇÃO ESTRUTURAL: OK")
        print("✅ VALIDAÇÃO EDITORIAL: OK")
        print("✅ QUALIDADE EDITORIAL: >= 8.0/10")
        print("✅ GRAVAÇÃO 100% VALIDADA")
        print("============================================================")
        print("BLOCOS:                5 / 5")
        print("FRAGMENTOS:           15 / 15")
        print("PARÁGRAFOS OLLAMA:    15 / 15")
        print()
        print("MEMÓRIA × JSON:")
        print("FRAGMENTOS:            IDÊNTICOS")
        print("PARÁGRAFOS:            IDÊNTICOS")
        print("HASHES:                5 / 5 IDÊNTICOS")
        print("STATUS:                APROVADO")
        print("============================================================")

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

    janela.after(
        0,
        lambda: barra_progresso.config(
            value=valor
        )
    )

    janela.after(
        0,
        lambda: status_progresso.config(
            text=texto
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
# TEMPO EM TEMPO REAL
# ============================================================

def atualizar_tempo_execucao(inicio, tema):

    while IA_PROCESSANDO:

        tempo = time.time() - inicio

        texto = (
            f"Tema: {tema} | "
            f"Tempo: {formatar_tempo(tempo)}"
        )

        janela.after(
            0,
            lambda: status_tempo.config(
                text=texto
            )
        )

        time.sleep(1)
        
        

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
        print(dados)


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
        textos,
        normalizar_checkboxes_editoriais()
    )
    
    
    print()
    print("==============================")
    print("MAPA MEAD RECEBIDO")
    print("==============================")
    
    print(
        str(mapa_mead)[:1000]
    )


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
        print("CHAMANDO SALVAR MAPA MEAD OFICIAL")
        print("==============================")

        # MAPA_MEAD é um artefato intermediário próprio. Ele não pode
        # passar pela validação de página final, que exige 5 blocos e
        # 15 parágrafos Ollama. A validação editorial da página pertence
        # somente a categoria="conteudo_completo".
        mapa_valido = validar_mapa_mead(mapa_mead)
        if not mapa_valido:
            print("🔴 MAPA MEAD NÃO SALVO: validação estrutural do mapa falhou.")
            return ""

        try:
            banco = carregar_banco()
            if not isinstance(banco, dict):
                banco = {}
            chave_tema = normalizar_tema_chave(tema)
            dados_existentes = banco.get(chave_tema, {})
            if not isinstance(dados_existentes, dict):
                dados_existentes = {}
            dados_existentes["mapa_mead"] = {
                "status": "gerado",
                "texto": str(mapa_mead).strip()
            }
            dados_existentes["tema"] = dados_existentes.get("tema") or tema
            banco[chave_tema] = dados_existentes

            diretorio = os.path.dirname(ARQUIVO_BANCO)
            if diretorio:
                os.makedirs(diretorio, exist_ok=True)
            arquivo_temp = ARQUIVO_BANCO + ".mapa.tmp"
            with open(arquivo_temp, "w", encoding="utf-8") as f: 
                json.dump(banco, f, ensure_ascii=False, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(arquivo_temp, ARQUIVO_BANCO)
        except Exception as erro:
            print("🔴 MAPA MEAD NÃO SALVO:", repr(erro))
            return ""

        print()
        print("==============================")
        print("MAPA MEAD SALVO OFICIALMENTE")
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

    try:

        PROCESSAMENTO_ATIVO = True
        PAGINAS_EM_PROCESSAMENTO.clear()
        IA_PROCESSANDO = True

        progresso = carregar_progresso()

        inicio_total = time.time()

        threading.Thread(
            target=atualizar_tempo_execucao,
            args=(inicio_total, "Processamento geral"),
            daemon=True
        ).start()

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
                r"C:\Python\gerador-conteudo\banco\conteudo-site.json"
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
                            "model": "qwen3:latest",
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
        
            inicio_tema = time.time()



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
                        tema,
                        normalizar_checkboxes_editoriais()
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
                print("FONTES DISPONÍVEIS PARA SELEÇÃO PYTHON")
                print("==============================")

                print(
                    "TEMA:",
                    tema
                )

                print(
                    "FONTES COLETADAS:",
                    total_textos
                )

                print(
                    "CARACTERES DAS FONTES:",
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

                print()
                print("==============================")
                print("MAPA MEAD ENVIADO AO CONTEÚDO")
                print("==============================")

                if mapa_mead:

                    print(
                        str(mapa_mead)[:1000]
                    )

                else:

                    print("MAPA VAZIO")

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

                print()
                print("==============================")
                print("DEBUG CHECKBOXES EDITORIAIS")
                print("==============================")

                for chave, valor in (
                    estrutura_editorial_atual.items()
                ):

                    print(
                        f"{chave}: {valor}"
                    )

                print("==============================")

                print(
                    "TOTAL SELECIONADOS:",
                    sum(
                        1
                        for valor
                        in estrutura_editorial_atual.values()
                        if valor
                    )
                )

                print(
                    "TOTAL BLOCOS:",
                    len(
                        estrutura_editorial_atual
                    )
                )

                print("==============================")

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
                if conteudo_completo is None:
                    print("GERAÇÃO NÃO CONCLUÍDA / OLLAMA NÃO EXECUTADO OU PROCESSAMENTO BLOQUEADO")
                else:
                    print("OLLAMA / GERAÇÃO RETORNOU")
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
                    "FONTES USADAS NA SELEÇÃO PYTHON:",
                    total_textos
                )

                print(
                    "CARACTERES PESQUISADOS/SELECIONADOS:",
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
                        "mapa_mead",
                        0,
                        "falhou",
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
                import traceback
                traceback.print_exc()

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

        atualizar_progresso(
            100,
            (
                f"Concluído | "
                f"Tempo total: "
                f"{formatar_tempo(tempo_final)}"
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
# GERAR DOCX MEAD
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
                              alinhamento=None, espaco_depois=0):
    texto = str(texto or "").strip()
    if not texto:
        return None
    p = documento.add_paragraph()
    if alinhamento is not None:
        p.alignment = alinhamento
    if espaco_depois:
        p.paragraph_format.space_after = Pt(espaco_depois)
    run = p.add_run(texto)
    _docx_configurar_fonte(run, tamanho, negrito)
    return p


def _docx_adicionar_separador(documento):
    return _docx_adicionar_paragrafo(
        documento,
        "-" * 110,
        tamanho=11,
        negrito=False
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
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
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


def _docx_adicionar_lista_segmentos(documento, dados_pagina):
    segmentos = dados_pagina.get("segmentos_listas", {})
    if not isinstance(segmentos, dict):
        return False

    posicionamento = dados_pagina.get("posicionamento_listas", {})
    if not isinstance(posicionamento, dict):
        posicionamento = {}

    bloco_pos = posicionamento.get("bloco")
    try:
        bloco_pos = int(bloco_pos)
    except Exception:
        bloco_pos = 5

    if bloco_pos not in (3, 4, 5):
        bloco_pos = 5

    # Guarda a posição escolhida para a rotina principal.
    dados_pagina["_docx_bloco_segmentos"] = bloco_pos
    return True


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
        bloco_pos = 5

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
        _docx_adicionar_paragrafo(documento, subtitulo, tamanho=14, negrito=True)

    for item in listas:
        _docx_adicionar_paragrafo(documento, "·\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0" + item, tamanho=11)

    return True


def _docx_adicionar_bloco(documento, bloco, numero):
    if not isinstance(bloco, dict):
        return
    titulo = str(bloco.get("titulo", "") or "").strip()
    if titulo:
        _docx_adicionar_paragrafo(documento, titulo, tamanho=14, negrito=True)

    paragrafos = bloco.get("paragrafos_ollama", [])
    if not isinstance(paragrafos, list):
        paragrafos = []

    # O DOCX usa exclusivamente o texto editorial já salvo no JSON.
    for texto in paragrafos[:3]:
        _docx_adicionar_paragrafo(documento, texto, tamanho=11)


def gerar_docx_mead(dados_pagina, pasta_destino, nome_arquivo_base=""):
    """Monta o DOCX exclusivamente a partir dos dados já existentes.

    Esta função não chama IA, não pesquisa, não seleciona fragmentos,
    não gera textos e não altera o conteúdo editorial.
    """
    if not isinstance(dados_pagina, dict):
        raise ValueError("Dados da página inválidos para geração do DOCX.")

    os.makedirs(pasta_destino, exist_ok=True)
    documento = Document()

    # --------------------------------------------------------
    # CONFIGURAÇÃO GLOBAL: Calibri
    # --------------------------------------------------------
    estilo_normal = documento.styles["Normal"]
    estilo_normal.font.name = "Calibri"
    estilo_normal.font.size = Pt(11)
    try:
        estilo_normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    except Exception:
        pass

    # --------------------------------------------------------
    # 1. TAGS
    # --------------------------------------------------------
    tags = dados_pagina.get("tags", [])
    if isinstance(tags, list):
        tags_texto = ", ".join(str(x).strip() for x in tags if str(x).strip())
    else:
        tags_texto = str(tags or "").strip()
    _docx_adicionar_paragrafo(documento, tags_texto, tamanho=11)

    # --------------------------------------------------------
    # 2. TÍTULO / SUBTÍTULO
    # --------------------------------------------------------
    titulo = str(dados_pagina.get("titulo", "") or "").strip()
    if not titulo:
        titulo = str(dados_pagina.get("tema", "") or "").strip()
    _docx_adicionar_paragrafo(documento, titulo, tamanho=14, negrito=True)

    subtitulo = str(dados_pagina.get("subtitulo", "") or "").strip()
    _docx_adicionar_paragrafo(documento, subtitulo, tamanho=11)

    _docx_adicionar_separador(documento)

    # --------------------------------------------------------
    # 3. IMAGENS 1 E 2
    # --------------------------------------------------------
    imagens = _docx_obter_imagens(dados_pagina)
    _docx_adicionar_par_imagens(documento, imagens, 1, 2)
    _docx_adicionar_separador(documento)

    # --------------------------------------------------------
    # 4. BLOCOS 1 A 5
    # --------------------------------------------------------
    blocos = _docx_obter_blocos(dados_pagina)

    posicionamento_imagens = dados_pagina.get("posicionamento_imagens", {})
    if not isinstance(posicionamento_imagens, dict):
        posicionamento_imagens = {}

    for numero in range(1, 6):
        chave = f"bloco_{numero}"

        # Imagens 3 e 4 formam um único conjunto e permanecem lado a lado.
        try:
            posicao_par_imagens = int(
                posicionamento_imagens.get("imagem_3", 0) or 0
            )
        except Exception:
            posicao_par_imagens = 0

        if posicao_par_imagens == numero:
            _docx_adicionar_par_imagens(documento, imagens, 3, 4)
            _docx_adicionar_separador(documento)

        bloco = blocos.get(chave, {})
        _docx_adicionar_bloco(documento, bloco, numero)

        # CTA fixo já definido no modelo, sem geração de conteúdo.
        if numero == 1:
            _docx_adicionar_separador(documento)
            _docx_adicionar_paragrafo(
                documento,
                "Entre em Contato e Solicite seu Orçamento!",
                tamanho=14,
                negrito=True
            )
            _docx_adicionar_paragrafo(
                documento,
                "Conte com nossa equipe para orientar seu fornecimento com base nas informações e especificações apresentadas para a aplicação.",
                tamanho=11
            )
            _docx_adicionar_paragrafo(
                documento,
                "Fale conosco pelo WhatsApp, telefone, e-mail ou envie sua mensagem através do formulário. Teremos satisfação em esclarecer suas dúvidas e contribuir para a definição da melhor alternativa.",
                tamanho=11
            )
            _docx_adicionar_separador(documento)

            # OBS continua imediatamente após o CTA do bloco 1,
            # quando existir no JSON. As imagens 3/4 seguem suas posições
            # sorteadas e persistidas entre os blocos inferiores.
            obs = str(dados_pagina.get("obs", "") or "").strip()
            if obs:
                _docx_adicionar_paragrafo(documento, obs, tamanho=14, negrito=True)
                _docx_adicionar_separador(documento)
        else:
            # Segmentos, quando configurados, ficam no final do bloco 3,
            # 4 ou 5, imediatamente antes do separador.
            _docx_inserir_segmentos_se_existirem(documento, dados_pagina, numero)
            _docx_adicionar_separador(documento)

    # --------------------------------------------------------
    # 5. CTA FINAL
    # --------------------------------------------------------
    nome_site = str(
        dados_pagina.get("nome_site", "") or ""
    ).strip()
    if not nome_site:
        nome_site = "Nossa equipe"

    _docx_adicionar_paragrafo(
        documento,
        f"Fale com a {nome_site}",
        tamanho=14,
        negrito=True
    )
    _docx_adicionar_paragrafo(
        documento,
        f"Entre em contato com nossos profissionais e solicite mais informações sobre {titulo}. Estamos à disposição!",
        tamanho=11
    )

    base = str(nome_arquivo_base or titulo or "documento").strip()
    nome_arquivo = (
        base.lower()
        .replace(" ", "_")
        .replace("/", "-")
        .replace("\\", "-")
        + ".docx"
    )
    caminho = os.path.join(pasta_destino, nome_arquivo)
    documento.save(caminho)
    return caminho


# ============================================================
# INTERFACE
# ============================================================
           

def gerar_docx_interface():
    """Gera DOCX somente com os dados já existentes no banco JSON."""
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
    print("\n==============================")
    print("GERANDO DOCX — SOMENTE JSON")
    print("==============================")

    for tema, dados in banco.items():
        if not isinstance(dados, dict):
            continue

        # O JSON pode ter os campos diretamente no tema ou dentro de pagina.
        pagina = dados.get("pagina")
        if not isinstance(pagina, dict):
            pagina = dados

        # Nenhuma chamada de IA, web ou regeneração acontece aqui.
        try:
            caminho = gerar_docx_mead(
                pagina,
                pasta,
                nome_arquivo_base=tema
            )
            total += 1
            print("DOCX GERADO:", caminho)
        except Exception as erro:
            print("ERRO AO GERAR DOCX:", tema, repr(erro))

    messagebox.showinfo("DOCX", f"{total} arquivo(s) DOCX gerado(s).")



    # ========================================================
    # 01. INTERFACE PRINCIPAL
    # ========================================================


    # ========================================================
    # 02. ESTRUTURA EDITORIAL
    # ========================================================

# Estado inicial dos oito checkboxes da interface.
# O usuário pode alterar cada um na GUI.
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
        grupo,
        normalizar_checkboxes_editoriais()
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
        tema,
        normalizar_checkboxes_editoriais()
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
            url,
            tema
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
