
# ============================================================
# MONTAR PÁGINA OFICIAL DO JSON
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
    Monta a estrutura oficial da página que será gravada
    posteriormente no conteudo-site.json.

    Esta função:

    - cria a estrutura oficial;
    - recebe somente os dados pertencentes à nova estrutura;
    - preserva os campos existentes;
    - garante os 5 blocos;
    - garante os 12 segmentos;
    - garante as 6 imagens;
    - prepara tags;
    - prepara mapa_mead;
    - prepara informações relevantes por bloco;
    - prepara grupo_principal_projeto;
    - calcula os caracteres;
    - deixa a página pronta para gravação.

    IMPORTANTE:
    - Não pesquisa.
    - Não seleciona textos.
    - Não chama Ollama.
    - Não inventa informações.
    - Não grava o arquivo.
    - NÃO cria informacoes_adicionais.
    - NÃO cria categorias.
    """

    global pagina

    try:

        # ----------------------------------------------------
        # NORMALIZAR TEMA
        # ----------------------------------------------------

        tema = str(
            tema or ""
        ).strip()

        if not tema:

            print(
                "❌ Não foi possível montar JSON: "
                "tema vazio."
            )

            return None

        # ----------------------------------------------------
        # NORMALIZAR ENTRADAS
        # ----------------------------------------------------

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
            informacoes_relevantes,
            dict
        ):

            informacoes_relevantes = {}

        if not isinstance(
            dados_pagina,
            dict
        ):

            dados_pagina = {}

        grupo_principal_projeto = str(
            grupo_principal_projeto or ""
        ).strip()

        # ----------------------------------------------------
        # CRIAR ESTRUTURA OFICIAL
        # ----------------------------------------------------

        estrutura = criar_estrutura_json_pagina(
            tema
        )

        if not isinstance(
            estrutura,
            dict
        ):

            print(
                "❌ criar_estrutura_json_pagina() "
                "não retornou um dicionário."
            )

            return None

        if tema not in estrutura:

            print(
                "❌ Tema não encontrado na estrutura "
                "criada para o JSON."
            )

            return None

        pagina = estrutura[
            tema
        ]

        # ----------------------------------------------------
        # TEMA / GRUPO / TIPO
        # ----------------------------------------------------

        pagina[
            "tema"
        ] = tema

        pagina[
            "grupo"
        ] = str(
            grupo or ""
        ).strip()

        pagina[
            "tipo"
        ] = str(
            tipo or ""
        ).strip()

        # ----------------------------------------------------
        # GRUPO PRINCIPAL DO PROJETO
        #
        # VEM DA INTERFACE E FICA DIRETAMENTE NO TEMA.
        # NÃO FICA DENTRO DE informacoes_adicionais.
        # ----------------------------------------------------

        pagina[
            "grupo_principal_projeto"
        ] = grupo_principal_projeto

        # ----------------------------------------------------
        # TAGS
        # ----------------------------------------------------

        tags_finais = []

        for tag in tags:

            tag_limpa = str(
                tag or ""
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

        pagina[
            "tags"
        ] = tags_finais

        # ----------------------------------------------------
        # CONTROLE DE REPETIÇÕES
        # ----------------------------------------------------

        palavra_controle = str(
            controle_repeticoes.get(
                "palavra_chave",
                tema
            )
            or tema
        ).strip()

        meta_repeticoes = controle_repeticoes.get(
            "meta_repeticoes",
            60
        )

        repeticoes_realizadas = controle_repeticoes.get(
            "repeticoes_realizadas",
            0
        )

        repeticoes_faltantes = controle_repeticoes.get(
            "repeticoes_faltantes",
            60
        )

        pagina[
            "controle_repeticoes"
        ] = {

            "palavra_chave":
                palavra_controle,

            "meta_repeticoes":
                meta_repeticoes,

            "repeticoes_realizadas":
                repeticoes_realizadas,

            "repeticoes_faltantes":
                repeticoes_faltantes
        }

        # ----------------------------------------------------
        # MAPA MEAD
        # ----------------------------------------------------

        pagina[
            "mapa_mead"
        ] = {

            "status":
                str(
                    mapa_mead.get(
                        "status",
                        ""
                    )
                    or ""
                ).strip(),

            "texto":
                str(
                    mapa_mead.get(
                        "texto",
                        ""
                    )
                    or ""
                ).strip()
        }

        # ----------------------------------------------------
        # INFORMAÇÕES RELEVANTES
        #
        # OS 15 FRAGMENTOS SELECIONADOS PELO PYTHON
        # FICAM DISTRIBUÍDOS NOS 5 BLOCOS.
        #
        # NÃO EXISTE MAIS CAMPO GLOBAL.
        # ----------------------------------------------------

        fragmentos_blocos = (
            informacoes_relevantes.get(
                "blocos",
                {}
            )
        )

        if not isinstance(
            fragmentos_blocos,
            dict
        ):

            fragmentos_blocos = {}

        for numero_bloco in range(
            1,
            6
        ):

            chave_bloco = (
                f"bloco_{numero_bloco}"
            )

            informacoes_bloco = (
                fragmentos_blocos.get(
                    chave_bloco,
                    []
                )
            )

        # ========================================================
        # INFORMAÇÕES RELEVANTES
        # ========================================================
        # Preservar obrigatoriamente como LISTA DE OBJETOS.
        # Não converter os objetos para string.
        # ========================================================
        
        informacoes_bloco = bloco.get(
            "informacoes_relevantes",
            []
        )
        
        if not isinstance(
            informacoes_bloco,
            list
        ):
            informacoes_bloco = []
        
        informacoes_relevantes_finais = []
        
        for fragmento in informacoes_bloco:
        
            if not isinstance(
                fragmento,
                dict
            ):
                continue
        
            fragmento_final = dict(
                fragmento
            )
        
            if "id" in fragmento_final:
                fragmento_final["id"] = str(
                    fragmento_final.get(
                        "id"
                    )
                    or ""
                ).strip()
        
            if "hash" in fragmento_final:
                fragmento_final["hash"] = str(
                    fragmento_final.get(
                        "hash"
                    )
                    or ""
                ).strip()
        
            if "texto" in fragmento_final:
                fragmento_final["texto"] = str(
                    fragmento_final.get(
                        "texto"
                    )
                    or ""
                ).strip()
        
            if "url" in fragmento_final:
                fragmento_final["url"] = str(
                    fragmento_final.get(
                        "url"
                    )
                    or ""
                ).strip()
        
            if "tipo" in fragmento_final:
                fragmento_final["tipo"] = str(
                    fragmento_final.get(
                        "tipo"
                    )
                    or ""
                ).strip()
        
            if "pdf" in fragmento_final:
                fragmento_final["pdf"] = bool(
                    fragmento_final.get(
                        "pdf"
                    )
                )
        
            if "palavras" in fragmento_final:
                try:
                    fragmento_final["palavras"] = int(
                        fragmento_final.get(
                            "palavras"
                        )
                        or 0
                    )
                except (
                    TypeError,
                    ValueError
                ):
                    fragmento_final["palavras"] = 0
        
            if "fonte" in fragmento_final:
                try:
                    fragmento_final["fonte"] = int(
                        fragmento_final.get(
                            "fonte"
                        )
                        or 0
                    )
                except (
                    TypeError,
                    ValueError
                ):
                    fragmento_final["fonte"] = 0
        
            informacoes_relevantes_finais.append(
                fragmento_final
            )
        
        informacoes_bloco = (
            informacoes_relevantes_finais
        )

            # -----------------------------------------------
            # GARANTIR BLOCO
            # -----------------------------------------------

            if chave_bloco not in pagina:

                pagina[
                    chave_bloco
                ] = {

                    "id":
                        chave_bloco,

                    "hash":
                        "",

                    "informacoes_relevantes":
                        "",

                    "titulo":
                        "",

                    "paragrafos":
                        []
                }

            # -----------------------------------------------
            # GRAVAR INFORMAÇÕES RELEVANTES
            # DIRETAMENTE NO BLOCO
            # -----------------------------------------------

            pagina[
                chave_bloco
            ][
                "informacoes_relevantes"
            ] = informacoes_bloco

        # ========================================================
        # DADOS GERAIS DA PÁGINA
        # ========================================================
        
        dados_gerais = [
            "tema",
            "arquivo_origem",
            "h1",
            "titulo",
            "subtitulo",
            "subtitulo_listas",
            "descricao"
        ]
        
        for campo in dados_gerais:
        
            valor = pagina.get(
                campo,
                ""
            )
        
            if valor is None:
                valor = ""
        
            if isinstance(
                valor,
                str
            ):
                valor = valor.strip()
        
            pagina_json[
                campo
            ] = valor

        # ----------------------------------------------------
        # GARANTIR TEMA DA PÁGINA
        # ----------------------------------------------------

        if not pagina[
            "pagina"
        ].get(
            "tema"
        ):

            pagina[
                "pagina"
            ][
                "tema"
            ] = tema

        # ----------------------------------------------------
        # ATUALIZAR OS 5 BLOCOS
        #
        # IMPORTANTE:
        # atualizar_bloco() recebe os dados produzidos
        # pelo gerador de conteúdo.
        # ----------------------------------------------------

        for numero in range(
            1,
            6
        ):

            chave_bloco = (
                f"bloco_{numero}"
            )

            dados_bloco = dados_pagina.get(
                chave_bloco,
                {}
            )

            if not isinstance(
                dados_bloco,
                dict
            ):

                dados_bloco = {}

            atualizar_bloco(
                numero,
                dados_bloco
            )

            # ------------------------------------------------
            # GARANTIR QUE A INFORMAÇÃO RELEVANTE
            # NÃO SEJA APAGADA PELO atualizar_bloco()
            # ------------------------------------------------

            informacao_selecionada = (
                pagina[
                    chave_bloco
                ].get(
                    "informacoes_relevantes",
                    ""
                )
            )

            if not informacao_selecionada:

                informacao_selecionada = (
                    fragmentos_blocos.get(
                        chave_bloco,
                        []
                    )
                )

                if isinstance(
                    informacao_selecionada,
                    list
                ):

                    informacao_selecionada = (
                        "\n\n".join(
                            str(item).strip()
                            for item in informacao_selecionada
                            if str(item).strip()
                        )
                    )

                else:

                    informacao_selecionada = str(
                        informacao_selecionada or ""
                    ).strip()

                pagina[
                    chave_bloco
                ][
                    "informacoes_relevantes"
                ] = informacao_selecionada

        # ----------------------------------------------------
        # SEGMENTOS
        # ----------------------------------------------------

        segmentos_recebidos = dados_pagina.get(
            "segmentos_listas",
            {}
        )

        if not isinstance(
            segmentos_recebidos,
            dict
        ):

            segmentos_recebidos = {}

        segmentos_oficiais = {}

        for numero in range(
            1,
            13
        ):

            chave_segmento = (
                f"segmento_{numero}"
            )

            valor = segmentos_recebidos.get(
                chave_segmento,
                []
            )

            if not isinstance(
                valor,
                list
            ):

                valor = []

            segmentos_oficiais[
                chave_segmento
            ] = [

                str(
                    item or ""
                ).strip()

                for item in valor

                if str(
                    item or ""
                ).strip()
            ]

        pagina[
            "pagina"
        ][
            "segmentos_listas"
        ] = segmentos_oficiais

        # ----------------------------------------------------
        # POSICIONAMENTO DAS LISTAS
        # ----------------------------------------------------

        posicionamento = dados_pagina.get(
            "posicionamento_listas",
            {}
        )

        if not isinstance(
            posicionamento,
            dict
        ):

            posicionamento = {}

        pagina[
            "pagina"
        ][
            "posicionamento_listas"
        ][
            "bloco"
        ] = posicionamento.get(
            "bloco"
        )

        # ----------------------------------------------------
        # IMAGENS
        # ----------------------------------------------------

        imagens_recebidas = dados_pagina.get(
            "imagens",
            {}
        )

        if not isinstance(
            imagens_recebidas,
            dict
        ):

            imagens_recebidas = {}

        for numero in range(
            1,
            7
        ):

            chave_imagem = (
                f"imagem_{numero}"
            )

            imagem = imagens_recebidas.get(
                chave_imagem,
                {}
            )

            if not isinstance(
                imagem,
                dict
            ):

                imagem = {}

            pagina[
                "pagina"
            ][
                "imagens"
            ][
                chave_imagem
            ] = {

                "url":
                    str(
                        imagem.get(
                            "url",
                            ""
                        )
                        or ""
                    ).strip(),

                "arquivo":
                    str(
                        imagem.get(
                            "arquivo",
                            ""
                        )
                        or ""
                    ).strip(),

                "alt":
                    str(
                        imagem.get(
                            "alt",
                            ""
                        )
                        or ""
                    ).strip(),

                "descricao":
                    str(
                        imagem.get(
                            "descricao",
                            ""
                        )
                        or ""
                    ).strip()
            }

        # ----------------------------------------------------
        # GARANTIR QUE AS ESTRUTURAS ANTIGAS NÃO EXISTAM
        # ----------------------------------------------------

        pagina.pop(
            "informacoes_adicionais",
            None
        )

        pagina.pop(
            "categorias",
            None
        )

        pagina.get(
            "pagina",
            {}
        ).pop(
            "informacoes_adicionais",
            None
        )

        pagina.get(
            "pagina",
            {}
        ).pop(
            "categorias",
            None
        )

        # ----------------------------------------------------
        # CALCULAR CARACTERES DA PÁGINA
        # ----------------------------------------------------

        caracteres = 0

        pagina_json = pagina.get(
            "pagina",
            {}
        )

        for numero in range(
            1,
            6
        ):

            bloco = pagina_json.get(
                f"bloco_{numero}",
                {}
            )

            if not isinstance(
                bloco,
                dict
            ):

                continue

            caracteres += len(
                str(
                    bloco.get(
                        "informacoes_relevantes",
                        ""
                    )
                    or ""
                ).strip()
            )

            caracteres += len(
                str(
                    bloco.get(
                        "titulo",
                        ""
                    )
                    or ""
                ).strip()
            )

            for paragrafo in bloco.get(
                "paragrafos",
                []
            )[:3]:

                caracteres += len(
                    str(
                        paragrafo or ""
                    ).strip()
                )

        pagina_json[
            "caracteres"
        ] = caracteres

        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------

        pagina_json[
            "status"
        ] = "pronta_para_gravacao"

        # ----------------------------------------------------
        # RETORNO
        # ----------------------------------------------------

        print()
        print(
            "=========================================="
        )
        print(
            "✅ PÁGINA OFICIAL MONTADA"
        )
        print(
            "=========================================="
        )

        print(
            f"   Tema: {tema}"
        )

        print(
            f"   Grupo: {grupo}"
        )

        print(
            f"   Grupo principal: "
            f"{grupo_principal_projeto}"
        )

        print(
            f"   Tags: {len(tags_finais)}"
        )

        print(
            "   Blocos: 5"
        )

        blocos_com_info = 0

        total_fragmentos = 0

        for numero in range(
            1,
            6
        ):

            bloco = pagina.get(
                f"bloco_{numero}",
                {}
            )

            info = str(
                bloco.get(
                    "informacoes_relevantes",
                    ""
                )
                or ""
            ).strip()

            if info:

                blocos_com_info += 1

                total_fragmentos += len(
                    [
                        x
                        for x in info.split(
                            "\n\n"
                        )
                        if x.strip()
                    ]
                )

        print(
            f"   Blocos com informações: "
            f"{blocos_com_info}/5"
        )

        print(
            f"   Fragmentos: "
            f"{total_fragmentos}/15"
        )

        print(
            "   Segmentos: 12"
        )

        print(
            "   Imagens: 6"
        )

        print(
            f"   Caracteres: {caracteres}"
        )

        print(
            "   Status: pronta_para_gravacao"
        )

        print(
            "=========================================="
        )

        return pagina

    except Exception as erro:

        print()
        print(
            "❌ ERRO ao montar página oficial:"
        )

        print(
            erro
        )

        return None
