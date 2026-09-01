# ============================================================
# AGENTE POWER BI - VERSÃO 6
# ============================================================

import os
import re
import json
import subprocess
import requests
import sys
import time

# ============================================================
# LOG SEM BUFFER
# ============================================================
#
# Em container (stdout sem TTY) o Python usa buffer de bloco.
# Isso faz os print() aparecerem fora de ordem ou serem perdidos
# quando o processo é encerrado, impedindo o diagnóstico.
# Equivale a PYTHONUNBUFFERED=1, mas garantido no próprio código.
#
try:
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)
except Exception:
    pass

from copy import deepcopy
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from typing import Optional, Literal

from pydantic import BaseModel
from groq import Groq
from google import genai


# ============================================================
# 1. CONFIGURAÇÕES
# ============================================================

WORKSPACE_ID = "54c63939-8bd5-4731-9c15-ba8a209cd469"

DATASET_ID = "cc908d1e-61ab-4795-9091-bf2e3b3de1b4"

URL = (
    f"https://api.powerbi.com/v1.0/myorg/"
    f"groups/{WORKSPACE_ID}/datasets/"
    f"{DATASET_ID}/executeQueries"
)

# ============================================================
# AUTENTICAÇÃO POWER BI - SERVICE PRINCIPAL
# ============================================================
#
# Tenant e Client ID podem ficar com valores padrão.
# O CLIENT SECRET NÃO fica gravado no código.
# Ele deve ser informado pela variável de ambiente:
# POWERBI_CLIENT_SECRET
#
POWERBI_TENANT_ID = os.getenv(
    "POWERBI_TENANT_ID",
    "a3a95e21-7484-4ca9-9d96-7964a3c9f893"
)

POWERBI_CLIENT_ID = os.getenv(
    "POWERBI_CLIENT_ID",
    "f62c8f5d-f643-4a14-8721-2a8820f6aacf"
)

POWERBI_CLIENT_SECRET = os.getenv(
    "POWERBI_CLIENT_SECRET"
)

POWERBI_TOKEN_URL = (
    f"https://login.microsoftonline.com/"
    f"{POWERBI_TENANT_ID}/oauth2/v2.0/token"
)

POWERBI_SCOPE = (
    "https://analysis.windows.net/powerbi/api/.default"
)

GROQ_MODEL = "llama-3.3-70b-versatile"

GEMINI_MODEL = "gemini-2.5-flash"

CLAUDE_MODEL = "claude-haiku-4-5"


# ============================================================
# 2. CLIENTES DE IA
# ============================================================

if not os.getenv("GROQ_API_KEY"):

    raise RuntimeError(
        "Variável GROQ_API_KEY não encontrada."
    )


groq_client = Groq()


if os.getenv("GEMINI_API_KEY"):

    gemini_client = genai.Client()

else:

    gemini_client = None

    print(
        "Aviso: GEMINI_API_KEY não encontrada. "
        "Fallback Gemini indisponível."
    )


if not os.getenv("ANTHROPIC_API_KEY"):
    print(
        "Aviso: ANTHROPIC_API_KEY não encontrada. "
        "Fallback Claude indisponível."
    )




token_cache = {
    "access_token": None,
    "expires_on": 0
}


# ============================================================
# 3. AUTENTICAÇÃO POWER BI
# ============================================================

def obter_token_powerbi():

    agora = time.time()

    # ========================================================
    # REUTILIZAR TOKEN EM CACHE
    # ========================================================

    if (
        token_cache["access_token"]
        and token_cache["expires_on"] > agora + 300
    ):

        return token_cache["access_token"]

    # ========================================================
    # VALIDAR CREDENCIAL DO SERVICE PRINCIPAL
    # ========================================================

    if not POWERBI_CLIENT_SECRET:

        raise RuntimeError(
            "Variável POWERBI_CLIENT_SECRET não encontrada. "
            "Configure o Client Secret do Service Principal "
            "antes de iniciar a API."
        )

    # ========================================================
    # GERAR TOKEN DIRETAMENTE NO MICROSOFT ENTRA ID
    # ========================================================

    resposta = requests.post(
        POWERBI_TOKEN_URL,
        data={
            "client_id":
                POWERBI_CLIENT_ID,

            "client_secret":
                POWERBI_CLIENT_SECRET,

            "scope":
                POWERBI_SCOPE,

            "grant_type":
                "client_credentials"
        },
        timeout=30
    )

    if resposta.status_code != 200:

        raise RuntimeError(
            f"Falha ao obter token do Service Principal "
            f"({resposta.status_code}): "
            f"{resposta.text[:500]}"
        )

    dados = resposta.json()

    access_token = dados.get(
        "access_token"
    )

    if not access_token:

        raise RuntimeError(
            "Microsoft Entra ID não retornou access_token."
        )

    # Normalmente o Entra retorna expires_in ~= 3599 segundos.
    expires_in = dados.get(
        "expires_in",
        3600
    )

    try:

        expires_in = int(
            expires_in
        )

    except Exception:

        expires_in = 3600

    expiracao_timestamp = (
        agora + expires_in
    )

    # ========================================================
    # SALVAR NO CACHE
    # ========================================================

    token_cache["access_token"] = (
        access_token
    )

    token_cache["expires_on"] = (
        expiracao_timestamp
    )

    return access_token


def criar_headers():

    return {
        "Authorization":
            f"Bearer {obter_token_powerbi()}",

        "Content-Type":
            "application/json"
    }


# ============================================================
# 4. DEFAULTS DO DASHBOARD
# ============================================================

contexto_overview_comercial = {

    "ano": "Ano atual",

    "mes": "Mês atual",

    "status_bloqueio":
        "PEDIDO LIBERADO",

    "frete":
        "media",

    "data_status_pedido": [
        "MÊS VIGENTE",
        "VENCIDO"
    ],

    "tipo_pedido":
        "VENDA",

    "custo":
        "custo std ctr",

    "tipo_meta":
        "Meta"
}


# ============================================================
# 5. MAPA DE FILTROS NORMAIS
# ============================================================

mapa_campos = {
    
    "regiao": {
        "tabela": "TAB CLIENTES",
        "coluna": "desc_area_gest"
    },

    "representante": {
        "tabela": "TAB CLIENTES",
        "coluna": "desc_area_venda"
    },


    "ano": {
        "tabela": "# CALENDÁRIO",
        "coluna": "ano_atual"
    },

    "mes": {
        "tabela": "# CALENDÁRIO",
        "coluna": "mes_atual"
    },

    "status_bloqueio": {
        "tabela": "DW CARTEIRA_VENDA",
        "coluna": "status_bloqueio"
    },

    "frete": {
        "tabela": "MD FRETES",
        "coluna": "desc_frete"
    },

    "plataforma": {
        "tabela": "TAB CLIENTES",
        "coluna": "desc_cliente_nivel_2"
    },
    
    "classe": {
        "tabela": "TAB CLIENTES",
        "coluna": "classificacao"
    },

    "data_status_pedido": {
        "tabela": "DW CARTEIRA_VENDA",
        "coluna": "data_status_pedido"
    },

    "tipo_pedido": {
        "tabela": "DW CARTEIRA_VENDA",
        "coluna": "desc_cod_tpdv"
    },

    "custo": {
        "tabela": "MD CUSTO",
        "coluna": "desc_custo"
    },

    "tipo_meta": {
        "tabela": "MD META",
        "coluna": "desc_meta"
    },

    "status_entrega": {
        "tabela": "DW ENTREGAS_VENDA",
        "coluna": "status_entrega"
    },

    "cod_curva_abc": {
        "tabela": "DW D_PRODUTO",
        "coluna": "cod_curva_abc"
    },

    "analise_credito": {
        "tabela": "DW CARTEIRA_VENDA",
        "coluna": "analise_credito"
    },

    # --------------------------------------------------------
    # PRODUTO
    # --------------------------------------------------------

    "linha": {
        "tipo": "normal",
        "tabela": "DW D_PRODUTO",
        "coluna": "desc_nivel2",
        "descricao": "linha"
    },

    "familia": {
        "tipo": "normal",
        "tabela": "DW D_PRODUTO",
        "coluna": "desc_nivel3",
        "descricao": "família"
    },

    "produto": {
        "tipo": "normal",
        "tabela": "DW D_PRODUTO",
        "coluna": "desc_produto",
        "descricao": "produto"
    },

    # --------------------------------------------------------
    # CLIENTE
    # --------------------------------------------------------

    "cliente": {
        "tabela": "TAB CLIENTES",
        "coluna": "desc_cliente"
    },

    "loja": {
        "tabela": "TAB CLIENTES",
        "coluna": "desc_cliente_nivel_3"
    }
}


# ============================================================
# 6. REPRESENTANTE - TABELA DESCONECTADA
# ============================================================



# ============================================================
# 7. DIMENSÕES PARA RANKING
# ============================================================

mapa_dimensoes = {
    "regiao": {
        "tipo": "normal",
        "tabela": "TAB CLIENTES",
        "coluna": "desc_area_gest",
        "descricao": "região"
    },
    "cliente": {
        "tipo": "normal",
        "tabela": "TAB CLIENTES",
        "coluna": "desc_cliente",
        "descricao": "cliente"
    },
    "produto": {
        "tipo": "normal",
        "tabela": "DW D_PRODUTO",
        "coluna": "desc_produto",
        "descricao": "produto"
    },
    "linha": {
        "tipo": "normal",
        "tabela": "DW D_PRODUTO",
        "coluna": "desc_nivel2",
        "descricao": "linha"
    },
    "familia": {
        "tipo": "normal",
        "tabela": "DW D_PRODUTO",
        "coluna": "desc_nivel3",
        "descricao": "família"
    },
    "representante": {
        "tipo": "normal",
        "tabela": "TAB CLIENTES",
        "coluna": "desc_area_venda",
        "descricao": "representante"
    },  # <-- Vírgula adicionada aqui
    "plataforma": {
        "tipo": "normal",
        "tabela": "TAB CLIENTES",
        "coluna": "desc_cliente_nivel_2",
        "descricao": "plataforma"
    },
    "cliente_plataforma": {
        "tipo": "normal",
        "tabela": "TAB CLIENTES",
        "coluna": "desc_cliente_nivel_3",
        "descricao": "cliente plataforma"
    },

    "loja": {
        "tipo": "normal",
        "tabela": "TAB CLIENTES",
        "coluna": "desc_cliente_nivel_3",
        "descricao": "loja/CNPJ"
    },
    "classe": {
        "tipo": "normal",
        "tabela": "TAB CLIENTES",
        "coluna": "classificacao",
        "descricao": "classe"
    },
    "status_entrega": {
        "tipo": "normal",
        "tabela": "DW ENTREGAS_VENDA",
        "coluna": "status_entrega",
        "descricao": "status de entrega"
    },
    "cod_curva_abc": {
        "tipo": "normal",
        "tabela": "DW D_PRODUTO",
        "coluna": "cod_curva_abc",
        "descricao": "curva ABC"
    },
    "analise_credito": {
        "tipo": "normal",
        "tabela": "DW CARTEIRA_VENDA",
        "coluna": "analise_credito",
        "descricao": "análise de crédito"
    }
}

# ============================================================
# 8. INDICADORES VALIDADOS
# ============================================================

mapa_indicadores = {

    "faturamento": {
        "medida": "fat_si",
        "descricao": "Faturamento",
        "formato": "moeda"
    },

    "meta_faturamento": {
        "medida": "fat_si_meta",
        "descricao": "Meta de faturamento",
        "formato": "moeda"
    },

    "atingimento_meta_faturamento": {
        "medida": "atg_meta_fat_si",
        "descricao":
            "Atingimento da meta de faturamento",
        "formato": "percentual"
    },

    "desvio_meta_faturamento": {
        "medida": "desv_meta_fat_si",
        "descricao":
            "Desvio da meta de faturamento",
        "formato": "moeda"
    },

    "margem_liquida": {
        "medida": "per_mrgl",
        "descricao": "Margem líquida",
        "formato": "percentual"
    },

    "valor_margem_liquida": {
        "medida": "mrgl",
        "descricao":
            "Valor da margem líquida",
        "formato": "moeda"
    },

    "margem_bruta": {
        "medida": "per_mrgb",
        "descricao": "Margem bruta",
        "formato": "percentual"
    },

    "valor_margem_bruta": {
        "medida": "mrgb",
        "descricao":
            "Valor da margem bruta",
        "formato": "moeda"
    },

    "quantidade": {
        "medida": "qtd",
        "descricao": "Quantidade vendida",
        "formato": "inteiro"
    },

    "meta_quantidade": {
        "medida": "qtd_meta",
        "descricao": "Meta de vendas em quantidade",
        "formato": "inteiro"
    },

    "entregas": {
        "medida": "fat_si_entregas",
        "descricao": "Entregas",
        "formato": "moeda"
    },

    "faturamento_entregas": {
        "medida": "fat_e_entregas",
        "descricao": "Faturamento + entregas",
        "formato": "moeda"
    },

    "meta_margem_liquida": {
        "medida": "per_mrgl_meta",
        "descricao": "Meta de margem líquida",
        "formato": "percentual"
    },

    "carteira_faturavel_dia": {
        "medida": "fat_si_carteira_atd_dia",
        "descricao": "Carteira faturável dia",
        "formato": "moeda"
    },

    "carteira_faturavel_mes": {
        "medida": "fat_si_carteira_atd_mes",
        "descricao": "Carteira faturável mês",
        "formato": "moeda"
    },

    "faturamento_carteira": {
        "medida": "fat_si_carteira",
        "descricao": "Faturamento carteira",
        "formato": "moeda"
    },

    "previsao_faturamento": {
        "medida": "prev_faturamento",
        "descricao": "Previsão de faturamento",
        "formato": "moeda"
    }
}



# ============================================================
# INDICADORES DERIVADOS - USADOS EM PERGUNTAS COMPOSTAS
# ============================================================
#
# Estes indicadores NÃO fazem uma chamada extra ao Power BI.
# São calculados localmente a partir de medidas já retornadas
# na mesma consulta DAX.
#
mapa_indicadores_derivados = {
    "atingimento_meta_quantidade": {
        "descricao": "Atingimento da meta de quantidade",
        "formato": "percentual",
        "bases": [
            "quantidade",
            "meta_quantidade"
        ]
    },

    "desvio_meta_quantidade": {
        "descricao": "Desvio da meta de quantidade",
        "formato": "inteiro",
        "bases": [
            "quantidade",
            "meta_quantidade"
        ]
    },

    "atingimento_meta_margem_liquida": {
        "descricao": "Atingimento da meta de margem líquida",
        "formato": "percentual",
        "bases": [
            "margem_liquida",
            "meta_margem_liquida"
        ]
    },

    "desvio_meta_margem_liquida": {
        "descricao": "Desvio da meta de margem líquida",
        "formato": "percentual",
        "bases": [
            "margem_liquida",
            "meta_margem_liquida"
        ]
    }
}


# ============================================================
# FAMÍLIAS CONTEXTUAIS DE META
# ============================================================

FAMILIAS_META = {
    "faturamento": {
        "principal": "faturamento",
        "meta": "meta_faturamento",
        "atingimento": "atingimento_meta_faturamento",
        "desvio": "desvio_meta_faturamento"
    },

    "quantidade": {
        "principal": "quantidade",
        "meta": "meta_quantidade",
        "atingimento": "atingimento_meta_quantidade",
        "desvio": "desvio_meta_quantidade"
    },

    "margem_liquida": {
        "principal": "margem_liquida",
        "meta": "meta_margem_liquida",
        "atingimento": "atingimento_meta_margem_liquida",
        "desvio": "desvio_meta_margem_liquida"
    }
}


def _config_indicador_qualquer(indicador):
    """
    Retorna configuração tanto de uma medida real do Power BI
    quanto de um indicador derivado localmente.
    """
    if indicador in mapa_indicadores:
        return mapa_indicadores[indicador]

    return mapa_indicadores_derivados[indicador]


# ============================================================
# 9. REGIÕES CONHECIDAS
# ============================================================

regioes_conhecidas = [

    "MLY ASSISTENCIA TECNICA",
    "NORDESTE 1",
    "SUDESTE 1",
    "MLY E-COMMERCE 3P",
    "MLY VENDA DIRETA",
    "NORDESTE 2",
    "SUL",
    "SUDESTE 3",
    "MINAS GERAIS",
    "ESPIRITO SANTO",
    "CENTRO-OESTE",
    "NORTE 1",
    "NORTE 2",
    "MLY EXPORTACAO",
    "MLY E-COMMERCE 1P",
    "MLY ESPECIALIZADO AR",
    "SUDESTE 2",
    "MLY FUNCIONARIO",
    "NOVOS NEGOCIOS"
]


# ============================================================
# 10. TIPOS
# ============================================================

IndicadorPermitido = Literal[

    "faturamento",
    "meta_faturamento",
    "atingimento_meta_faturamento",
    "desvio_meta_faturamento",
    "margem_liquida",
    "valor_margem_liquida",
    "margem_bruta",
    "valor_margem_bruta",
    "quantidade",
    "meta_quantidade",
    "entregas",
    "faturamento_entregas",
    "meta_margem_liquida",
    "carteira_faturavel_dia",
    "carteira_faturavel_mes",
    "faturamento_carteira",
    "previsao_faturamento"
]


OperacaoPermitida = Literal[
    "valor",
    "ranking",
    "resumo_meta",
    "resumo_gerencial"
]


DimensaoPermitida = Literal[

    "regiao",
    "cliente",
    "produto",
    "linha",
    "familia",
    "representante",
    "plataforma",
    "cliente_plataforma",
    "loja",
    "classe",
    "status_entrega",
    "cod_curva_abc",
    "analise_credito"
]


OrdemPermitida = Literal[
    "desc",
    "asc"
]


# ============================================================
# 11. MODELOS PYDANTIC
# ============================================================

class FiltrosPergunta(BaseModel):

    regiao: Optional[str] = None
    ano: Optional[str] = None
    mes: Optional[str | list[str]] = None

    cliente: Optional[str] = None
    produto: Optional[str] = None
    linha: Optional[str] = None
    familia: Optional[str] = None
    representante: Optional[str] = None

    status_bloqueio: Optional[str] = None
    frete: Optional[str] = None
    data_status_pedido: Optional[list[str]] = None
    tipo_pedido: Optional[str] = None
    custo: Optional[str] = None
    tipo_meta: Optional[str] = None
    plataforma: Optional[str] = None
    classe: Optional[str] = None
    status_entrega: Optional[str] = None
    cod_curva_abc: Optional[str] = None
    analise_credito: Optional[str] = None


class PeriodoRelativoPergunta(BaseModel):
    # A IA apenas normaliza a intenção temporal.
    # O Python continua sendo a fonte de verdade para calcular as datas.
    unidade: Literal[
        "mes",
        "ano",
        "trimestre",
        "semestre",
        "bimestre"
    ]
    deslocamento: int = 0


class InterpretacaoPergunta(BaseModel):

    operacao: OperacaoPermitida

    indicador: IndicadorPermitido

    filtros: FiltrosPergunta

    agrupar_por: Optional[DimensaoPermitida] = None

    top_n: Optional[int] = None

    ordem: Optional[OrdemPermitida] = None

    # Usado somente quando a IA reconhece um período relativo em
    # linguagem natural que não foi capturado pelas regras locais.
    # Ex.: "último mês" -> {"unidade": "mes", "deslocamento": 1}.
    periodo_relativo: Optional[PeriodoRelativoPergunta] = None

    fora_escopo: bool = False

# ============================================================
# 12. CONTEXTO DA CONVERSA
# ============================================================

contexto_conversa = {
    "ultima_interpretacao": None
}


def limpar_contexto():

    contexto_conversa[
        "ultima_interpretacao"
    ] = None

    return "Contexto da conversa limpo."


# ============================================================
# 13. MESES
# ============================================================

mapa_meses = {

    "janeiro": "01",
    "fevereiro": "02",
    "março": "03",
    "abril": "04",
    "maio": "05",
    "junho": "06",
    "julho": "07",
    "agosto": "08",
    "setembro": "09",
    "outubro": "10",
    "novembro": "11",
    "dezembro": "12"
}


meses_nome = {

    "01": "janeiro",
    "02": "fevereiro",
    "03": "março",
    "04": "abril",
    "05": "maio",
    "06": "junho",
    "07": "julho",
    "08": "agosto",
    "09": "setembro",
    "10": "outubro",
    "11": "novembro",
    "12": "dezembro"
}


# ============================================================
# 14. NORMALIZAR MÊS DO POWER BI
# ============================================================

def normalizar_mes_powerbi(
    mes,
    ano=None
):

    if mes is None:

        return None


    agora = datetime.now()

    mes_atual = (
        f"{agora.month:02d}"
    )

    ano_atual = str(
        agora.year
    )


    if mes == mes_atual:

        if (
            ano is None
            or ano == "Ano atual"
            or str(ano) == ano_atual
        ):

            return "Mês atual"


    return mes


# ============================================================
# 15. NORMALIZAR INTERPRETAÇÃO
# ============================================================

def normalizar_interpretacao(
    interpretacao
):

    # --------------------------------------------------------
    # REGIÃO
    # --------------------------------------------------------

    if interpretacao.filtros.regiao:

        interpretacao.filtros.regiao = (
            interpretacao
            .filtros
            .regiao
            .strip()
            .upper()
        )


    # --------------------------------------------------------
    # ANO SEMPRE TEXTO
    # --------------------------------------------------------

    if interpretacao.filtros.ano:

        interpretacao.filtros.ano = (
            str(
                interpretacao
                .filtros
                .ano
            )
        )


    # --------------------------------------------------------
    # MÊS ATUAL
    # --------------------------------------------------------

    # --------------------------------------------------------
    # MÊS / LISTA DE MESES
    # --------------------------------------------------------

    if interpretacao.filtros.mes:
    
        # Um único mês
        if isinstance(
            interpretacao.filtros.mes,
            str
        ):
    
            interpretacao.filtros.mes = (
                normalizar_mes_powerbi(
    
                    interpretacao
                    .filtros
                    .mes,
    
                    interpretacao
                    .filtros
                    .ano
                )
            )
    
        # Vários meses (trimestre)
        elif isinstance(
            interpretacao.filtros.mes,
            list
        ):
    
            interpretacao.filtros.mes = [
                str(mes).zfill(2)
                for mes
                in interpretacao.filtros.mes
            ]


    # --------------------------------------------------------
    # RANKING
    # --------------------------------------------------------

    if interpretacao.operacao == "ranking":

        if interpretacao.top_n is None:

            interpretacao.top_n = 20


        if interpretacao.ordem is None:

            interpretacao.ordem = "desc"


    # --------------------------------------------------------
    # VALOR
    # --------------------------------------------------------

    if interpretacao.operacao == "valor":

        interpretacao.agrupar_por = None

        interpretacao.top_n = None

        interpretacao.ordem = None


    # --------------------------------------------------------
    # NORMALIZAÇÃO DE CLASSE
    # --------------------------------------------------------
    
    if interpretacao.filtros.classe:
    
        classe = (
            interpretacao
            .filtros
            .classe
            .strip()
            .upper()
        )
    
        mapa_classes = {
            "VIP+": "VIP +",
            "VIP +": "VIP +",
            "VIP": "VIP",
            "ESPECIAL": "ESPECIAL",
            "REGULAR": "REGULAR",
            "OUTRO": "OUTRO"
        }
    
        interpretacao.filtros.classe = (
            mapa_classes.get(
                classe,
                classe
            )
        )


    return interpretacao


# ============================================================
# 16. NORMALIZAR JSON DAS IAs
# ============================================================

def normalizar_json_ia(
    dados
):

    if not isinstance(
        dados,
        dict
    ):

        raise ValueError(
            "A IA não retornou JSON válido."
        )


    # --------------------------------------------------------
    # filtro -> filtros
    # --------------------------------------------------------

    if (
        "filtro" in dados
        and "filtros" not in dados
    ):

        dados["filtros"] = (
            dados.pop("filtro")
        )


    if (
        "filtros" not in dados
        or dados["filtros"] is None
    ):

        dados["filtros"] = {}


    filtros_validos = {

        "regiao",
        "ano",
        "mes",
        "cliente",
        "produto",
        "linha",
        "familia",
        "representante",
        "status_bloqueio",
        "frete",
        "data_status_pedido",
        "tipo_pedido",
        "custo",
        "tipo_meta",
        "plataforma",
        "classe",
        "status_entrega",
        "cod_curva_abc",
        "analise_credito",
    }


    # --------------------------------------------------------
    # MOVE FILTROS DA RAIZ
    # --------------------------------------------------------

    for chave in list(
        dados.keys()
    ):

        if chave in filtros_validos:

            dados["filtros"][chave] = (
                dados.pop(chave)
            )


    # --------------------------------------------------------
    # ANO PARA STRING
    # --------------------------------------------------------

    if (
        dados["filtros"].get("ano")
        is not None
    ):

        dados["filtros"]["ano"] = str(
            dados["filtros"]["ano"]
        )


    dados.setdefault(
        "agrupar_por",
        None
    )

    dados.setdefault(
        "top_n",
        None
    )

    dados.setdefault(
        "ordem",
        None
    )

    dados.setdefault(
        "fora_escopo",
        False
    )
    
    return dados


# ============================================================
# 17. DETECTAR CONTINUAÇÃO
# ============================================================

def parece_continuacao(
    pergunta
):

    texto = (
        pergunta
        .strip()
        .lower()
    )

    # Comandos explícitos abaixo são perguntas completas, mesmo quando
    # possuem poucas palavras.
    if any(
        texto.startswith(inicio)
        for inicio in [
            "resumo gerencial",
            "resuma ",
            "visão geral",
            "visao geral",
            "me dê um resumo",
            "me de um resumo"
        ]
    ):
        return False


    inicios = [

        "e em ",
        "e a ",
        "e o ",
        "e os ",
        "e as ",
        "agora ",
        "e no ",
        "e na ",
        "e nos ",
        "e nas "
    ]


    if any(
        texto.startswith(x)
        for x in inicios
    ):

        return True


    if len(
        texto.split()
    ) <= 4:

        return True


    return False


# ============================================================
# 18. INTERPRETADOR LOCAL
# ============================================================

def interpretar_localmente(
    pergunta
):

    anterior = contexto_conversa[
        "ultima_interpretacao"
    ]


    if anterior is None:

        return None


    texto = (
        pergunta
        .strip()
        .lower()
    )


    texto_upper = (
        pergunta.upper()
    )


    nova = deepcopy(
        anterior
    )


    alterou = False


    # ========================================================
    # MÊS
    # ========================================================

    for nome, numero in (
        mapa_meses.items()
    ):

        if re.search(
            r"\b" + re.escape(nome) + r"\b",
            texto
        ):

            nova["filtros"]["mes"] = (
                numero
            )

            alterou = True

            break


    # ========================================================
    # REGIÃO
    # ========================================================

    for regiao in regioes_conhecidas:

        if regiao in texto_upper:

            nova["filtros"]["regiao"] = (
                regiao
            )

            alterou = True

            break


    # ========================================================
    # INDICADORES
    # ========================================================

    if any(x in texto for x in [
        "previsão de faturamento",
        "previsao de faturamento",
        "projeção de faturamento",
        "projecao de faturamento",
        "forecast"
    ]):
        nova["indicador"] = "previsao_faturamento"
        alterou = True

    elif any(x in texto for x in [
        "mais comprou",
        "comprou mais",
        "mais vendeu",
        "vendeu mais",
        "mais produtos",
        "mais unidades"
    ]):
        nova["indicador"] = "quantidade"
        alterou = True

    elif any(x in texto for x in [
        "meta de margem líquida",
        "meta de margem liquida",
        "percentual da meta de margem líquida",
        "percentual da meta de margem liquida"
    ]):
        nova["indicador"] = "meta_margem_liquida"
        alterou = True

    elif any(x in texto for x in [
        "faturamento + entregas",
        "fat + entregas",
        "faturamento e entregas"
    ]):
        nova["indicador"] = "faturamento_entregas"
        alterou = True

    elif any(x in texto for x in [
        "carteira faturável dia", "carteira faturavel dia",
        "carteira faturável do dia", "carteira faturavel do dia"
    ]):
        nova["indicador"] = "carteira_faturavel_dia"
        alterou = True

    elif any(x in texto for x in [
        "carteira faturável mês", "carteira faturavel mes",
        "carteira faturável do mês", "carteira faturavel do mes"
    ]):
        nova["indicador"] = "carteira_faturavel_mes"
        alterou = True

    elif any(x in texto for x in [
        "faturamento carteira", "faturamento da carteira"
    ]):
        nova["indicador"] = "faturamento_carteira"
        alterou = True

    elif any(x in texto for x in [
        "meta de quantidade", "meta em quantidade", "meta de vendas"
    ]):
        nova["indicador"] = "meta_quantidade"
        alterou = True

    elif any(x in texto for x in [
        "quantidade vendida", "quantidade de vendas", "qtd vendida"
    ]) or re.search(r"\bqtd\b", texto):
        nova["indicador"] = "quantidade"
        alterou = True

    elif "entregas" in texto:
        nova["indicador"] = "entregas"
        alterou = True

    elif (
        "margem líquida" in texto
        or "margem liquida" in texto
    ):

        if (
            "valor" in texto
            or "reais" in texto
            or "r$" in texto
        ):

            nova["indicador"] = (
                "valor_margem_liquida"
            )

        else:

            nova["indicador"] = (
                "margem_liquida"
            )

        alterou = True


    elif "margem bruta" in texto:

        if (
            "valor" in texto
            or "reais" in texto
            or "r$" in texto
        ):

            nova["indicador"] = (
                "valor_margem_bruta"
            )

        else:

            nova["indicador"] = (
                "margem_bruta"
            )

        alterou = True


    elif (
        "atingimento" in texto
        or "percentual da meta" in texto
        or "bateu da meta" in texto
        or "atingiu da meta" in texto
    ):

        nova["indicador"] = (
            "atingimento_meta_faturamento"
        )

        alterou = True


    elif (
        "quanto falta" in texto
        or "desvio da meta" in texto
        or "diferença para a meta" in texto
    ):

        nova["indicador"] = (
            "desvio_meta_faturamento"
        )

        alterou = True


    elif (
        "meta" in texto
        and "margem" not in texto
    ):

        nova["indicador"] = (
            "meta_faturamento"
        )

        alterou = True


    elif (
        "faturamento" in texto
        or "faturou" in texto
        or "faturaram" in texto
        or "faturar" in texto
    ):

        nova["indicador"] = (
            "faturamento"
        )

        alterou = True


    # ========================================================
    # DIMENSÃO DE RANKING
    # ========================================================

    if (
        "representantes" in texto
        or "representante" in texto
    ):

        nova["operacao"] = "ranking"

        nova["agrupar_por"] = (
            "representante"
        )

        alterou = True


    elif (
        "clientes" in texto
        or "cliente" in texto
    ):

        nova["operacao"] = "ranking"

        nova["agrupar_por"] = "cliente"

        alterou = True


    elif (
        "produtos" in texto
        or "produto" in texto
    ):

        nova["operacao"] = "ranking"

        nova["agrupar_por"] = "produto"

        alterou = True


    elif (
        "famílias" in texto
        or "familias" in texto
        or "família" in texto
        or "familia" in texto
    ):

        nova["operacao"] = "ranking"

        nova["agrupar_por"] = "familia"

        alterou = True


    elif (
        "linhas" in texto
        or "linha" in texto
    ):

        nova["operacao"] = "ranking"

        nova["agrupar_por"] = "linha"

        alterou = True


    elif (
        "regiões" in texto
        or "regioes" in texto
        or "região" in texto
        or "regiao" in texto
        or "filial" in texto
        or "filiais" in texto
    ):

        nova["operacao"] = "ranking"

        nova["agrupar_por"] = "regiao"

        alterou = True


    # Novas dimensões comerciais
    if "status de entrega" in texto or "status_entrega" in texto:
        nova["operacao"] = "ranking"
        nova["agrupar_por"] = "status_entrega"
        alterou = True

    elif "curva abc" in texto or "cod_curva_abc" in texto:
        nova["operacao"] = "ranking"
        nova["agrupar_por"] = "cod_curva_abc"
        alterou = True

    elif "análise de crédito" in texto or "analise de credito" in texto:
        nova["operacao"] = "ranking"
        nova["agrupar_por"] = "analise_credito"
        alterou = True


    # ========================================================
    # TOP N
    # ========================================================

    match_top = re.search(
        r"\b(?:top\s*)?(\d+)\b",
        texto
    )


    if (
        match_top
        and nova.get("operacao")
            == "ranking"
    ):

        numero = int(
            match_top.group(1)
        )


        if 1 <= numero <= 100:

            nova["top_n"] = numero

            alterou = True


    # ========================================================
    # ORDEM
    # ========================================================

    if any(
        x in texto
        for x in [
            "menos",
            "menor",
            "menores",
            "piores"
        ]
    ):

        if nova.get(
            "operacao"
        ) == "ranking":

            nova["ordem"] = "asc"

            alterou = True


    elif any(
        x in texto
        for x in [
            "mais",
            "maior",
            "maiores",
            "melhores",
            "top"
        ]
    ):

        if nova.get(
            "operacao"
        ) == "ranking":

            nova["ordem"] = "desc"

            alterou = True


    if not alterou:

        return None


    try:

        interpretacao = (
            InterpretacaoPergunta(
                **nova
            )
        )

        return normalizar_interpretacao(
            interpretacao
        )

    except Exception:

        return None


# ============================================================
# 19. PROMPT DAS IAs
# ============================================================

def montar_prompt_ia(
    pergunta
):

    anterior = contexto_conversa[
        "ultima_interpretacao"
    ]


    contexto_anterior = ""


    if anterior:

        contexto_anterior = f"""
CONTEXTO ANTERIOR:

{json.dumps(
    anterior,
    ensure_ascii=False
)}

Se for continuação, mantenha somente
as informações anteriores que o usuário
não modificou.
"""


    return f"""
Você interpreta perguntas para um agente de Business Intelligence.

NÃO responda à pergunta.
NÃO calcule números.

Retorne SOMENTE um JSON válido.
Não escreva texto fora do JSON.

Estrutura:

{{
  "operacao": "valor ou ranking",
  "indicador": "indicador",
  "filtros": {{}},
  "agrupar_por": null,
  "top_n": null,
  "ordem": null,
  "periodo_relativo": null,
  "fora_escopo": false
}}


============================================================
INTERPRETAÇÃO SEMÂNTICA — REGRA PRINCIPAL
============================================================

Interprete a INTENÇÃO e o SIGNIFICADO da pergunta, e não apenas
palavras ou frases exatas listadas neste prompt.

As palavras e exemplos abaixo são referências de negócio, NÃO uma
lista fechada de sinônimos. O usuário pode usar outras formas de
falar, abreviações, linguagem informal, flexões verbais, pequenos
erros de digitação ou expressões equivalentes.

Quando uma expressão nova tiver significado claramente equivalente
a um indicador, dimensão, operação ou filtro permitido, normalize-a
para o valor canônico deste JSON.

Exemplos de equivalência semântica:
- "quanto entrou", "quanto faturou", "receita realizada" podem
  significar faturamento quando o contexto comercial deixar isso claro;
- "quem comprou mais", "maior comprador", "cliente que mais levou"
  podem representar ranking de cliente por quantidade;
- "quem vendeu mais", "maior volume vendido", "mais unidades" podem
  representar quantidade;
- "quem deu mais receita", "quem mais faturou", "maior faturamento"
  podem representar ranking por faturamento;
- "vendedor", quando usado claramente para a pessoa responsável pela
  venda, pode representar representante;
- "linha de produto" e variações semanticamente equivalentes devem ser
  normalizadas para a dimensão linha;
- erros simples como "faturameto", "margem liquda" ou abreviações
  compreensíveis não devem impedir a interpretação.

IMPORTANTE:
- Para PERÍODOS RELATIVOS, normalize a intenção no campo
  "periodo_relativo" e NÃO tente calcular mês/ano por conta própria.
  O Python calculará a faixa exata de datas.
- Estrutura de periodo_relativo:
  {{"unidade": "mes|ano|trimestre|semestre|bimestre", "deslocamento": N}}
- Exemplos semânticos:
  "este mês", "mês atual" -> {{"unidade": "mes", "deslocamento": 0}}
  "último mês", "mês passado", "mês anterior", "mês que passou"
    -> {{"unidade": "mes", "deslocamento": 1}}
  "mês retrasado" -> {{"unidade": "mes", "deslocamento": 2}}
  "último trimestre" / "trimestre passado"
    -> {{"unidade": "trimestre", "deslocamento": 1}}
  A mesma lógica vale para ano, semestre e bimestre.
- Quando periodo_relativo for usado, NÃO preencha filtros.ano ou filtros.mes
  com uma data calculada ou herdada do contexto anterior.
- NÃO invente um indicador, dimensão, filtro ou operação que não exista
  nas listas permitidas deste prompt;
- NÃO force uma interpretação quando houver ambiguidade real;
- use fora_escopo=true somente quando a pergunta realmente não pertencer
  ao domínio comercial suportado;
- se a intenção comercial estiver clara, escolha o conceito permitido
  semanticamente mais próximo, mesmo que a frase exata nunca tenha sido
  cadastrada antes;
- preserve filtros e contexto anterior somente quando forem compatíveis
  com a pergunta atual.

============================================================
OPERAÇÕES
============================================================

valor
ranking
resumo_meta

============================================================
INDICADORES
============================================================

faturamento
meta_faturamento
atingimento_meta_faturamento
desvio_meta_faturamento
margem_liquida
valor_margem_liquida
margem_bruta
valor_margem_bruta
quantidade
meta_quantidade
entregas
faturamento_entregas
meta_margem_liquida
carteira_faturavel_dia
carteira_faturavel_mes
faturamento_carteira
previsao_faturamento


============================================================
REGRAS DE INDICADORES
============================================================

faturamento / faturou / faturaram
= faturamento

meta / meta de faturamento
= meta_faturamento

atingimento / percentual da meta
= atingimento_meta_faturamento

quanto falta para a meta / desvio
= desvio_meta_faturamento

margem líquida
= margem_liquida

margem líquida em reais
= valor_margem_liquida

margem bruta
= margem_bruta

margem bruta em reais
= valor_margem_bruta

quantidade / quantidade vendida / qtd
= quantidade

meta de quantidade / meta em quantidade / meta de vendas
= meta_quantidade

entregas
= entregas

faturamento + entregas / fat + entregas / faturamento e entregas
= faturamento_entregas

meta de margem líquida / percentual da meta de margem líquida
= meta_margem_liquida

carteira faturável dia / carteira faturável do dia
= carteira_faturavel_dia

carteira faturável mês / carteira faturável do mês
= carteira_faturavel_mes

faturamento carteira / faturamento da carteira
= faturamento_carteira

previsão de faturamento / previsao de faturamento /
projeção de faturamento / forecast
= previsao_faturamento

mais comprou / comprou mais / mais vendeu / vendeu mais
= quantidade


============================================================
DIMENSÕES DE RANKING
============================================================

regiao
cliente
produto
linha
familia
representante
plataforma
cliente_plataforma
classe
status_entrega
cod_curva_abc
analise_credito


IMPORTANTE:

Neste dashboard:

"região" e "filial"
representam a mesma dimensão.

Se o usuário falar:

região
regiões
filial
filiais

use:

agrupar_por = regiao


============================================================
FILTROS DISPONÍVEIS
============================================================

regiao
ano
mes
cliente
produto
linha
familia
representante
status_bloqueio
frete
data_status_pedido
tipo_pedido
custo
tipo_meta
plataforma
classe
status_entrega
cod_curva_abc
analise_credito


============================================================
NOVOS FILTROS / AGRUPAMENTOS
============================================================

status_entrega corresponde a:
DW ENTREGAS_VENDA[status_entrega]

cod_curva_abc corresponde a:
DW D_PRODUTO[cod_curva_abc]

analise_credito corresponde a:
DW CARTEIRA_VENDA[analise_credito]

O usuário pode usar esses campos como filtro OU agrupamento
com qualquer indicador.

Exemplos:

"Faturamento por status de entrega"
→ agrupar_por = status_entrega

"Quantidade vendida da curva ABC A"
→ indicador = quantidade
→ filtros.cod_curva_abc = A

"Faturamento por curva ABC"
→ agrupar_por = cod_curva_abc

"Carteira faturável do mês por análise de crédito"
→ indicador = carteira_faturavel_mes
→ agrupar_por = analise_credito


============================================================
PRODUTOS
============================================================

linha corresponde ao agrupamento
mais amplo de produtos.

familia corresponde à família
de produtos.

Exemplos:

"produtos da família FRITADEIRAS"

agrupar_por = produto

filtros:
familia = FRITADEIRAS


"faturamento da linha COZINHA"

operacao = valor

filtros:
linha = COZINHA


============================================================
REPRESENTANTE
============================================================

"ranking por representante"

agrupar_por = representante


"faturamento do representante X"

operacao = valor

filtros:
representante = X


============================================================
PLATAFORMA E CLIENTES
============================================================

REGRA NOVA DE NEGÓCIO:
- plataforma, grupo, varejo, varejista e cliente representam
  TAB CLIENTES[desc_cliente_nivel_2] quando usados como agrupamento comercial.
  Use agrupar_por = plataforma.
- loja, CNPJ e estabelecimento representam
  TAB CLIENTES[desc_cliente_nivel_3].
  Use agrupar_por = loja.

Exemplos:
"Qual faturamento por cliente?"
→ agrupar_por = plataforma

"Qual varejista mais comprou meus produtos?"
→ indicador = quantidade
→ agrupar_por = plataforma
→ top_n = 1

"Qual previsão de faturamento por cliente?"
→ indicador = previsao_faturamento
→ agrupar_por = plataforma

"Qual faturamento por lojas?"
→ agrupar_por = loja

"Qual faturamento das lojas da AMAZON?"
→ agrupar_por = loja
→ filtros.plataforma = AMAZON


plataforma corresponde a:
TAB CLIENTES[desc_cliente_nivel_2]

cliente_plataforma corresponde a:
TAB CLIENTES[desc_cliente_nivel_3]

IMPORTANTE:

Quando o usuário pedir clientes normalmente:

"Top 5 clientes por faturamento"

use:
agrupar_por = cliente


Quando o usuário pedir clientes DE UMA PLATAFORMA:

"Quais clientes da plataforma GAZIN mais faturaram?"

use:
agrupar_por = cliente_plataforma

e:

filtros:
plataforma = GAZIN


Exemplo:

"Top 5 plataformas por faturamento"

operacao = ranking
agrupar_por = plataforma


Exemplo:

"Quanto a plataforma GAZIN faturou?"

operacao = valor

filtros:
plataforma = GAZIN


============================================================
PRIORIDADE DE INDICADOR EXPLÍCITO
============================================================

Se o usuário mencionar explicitamente um indicador,
esse indicador tem prioridade sobre expressões genéricas
como:

"como está"
"como foi"
"situação"
"desempenho"

Exemplos:

"Como está o faturamento em agosto?"
→ operacao = valor
→ indicador = faturamento
→ filtros.mes = 08

"Como está a margem líquida em agosto?"
→ operacao = valor
→ indicador = margem_liquida
→ filtros.mes = 08

"Como foi a meta em julho?"
→ operacao = valor
→ indicador = meta_faturamento
→ filtros.mes = 07

Use resumo_gerencial somente quando o usuário pedir
uma visão ampla e NÃO especificar um único indicador.

Exemplos de resumo_gerencial:

"Como está a operação em agosto?"
"Me dê uma visão geral de agosto."
"Como está o desempenho da região SUL?"






============================================================
RESUMO GERENCIAL
============================================================

Use operacao = resumo_gerencial quando o usuário pedir
um resumo, visão geral, desempenho ou situação de uma
entidade comercial.

Exemplos:

"Me dê um resumo da plataforma GAZIN em agosto de 2026."

"Como está a plataforma AMAZON?"

"Me dê uma visão geral da classe ESPECIAL em 2026."

"Como está o desempenho da região SUL?"

"Resuma o desempenho do representante X."

Nesses casos:

operacao = resumo_gerencial
indicador = faturamento
agrupar_por = null
top_n = null
ordem = null

A entidade deve ser colocada normalmente em filtros.

Exemplo:

"Me dê um resumo da plataforma GAZIN em agosto de 2026."

filtros:
plataforma = GAZIN
ano = 2026
mes = 08




============================================================
CLASSE
============================================================

O usuário pode consultar qualquer indicador por classe.

Exemplos:

"Qual o faturamento da classe ESPECIAL?"
→ filtro classe = ESPECIAL

"Ranking de faturamento por classe"
→ agrupar_por = classe

"Qual a margem líquida da classe VIP?"
→ indicador = margem_liquida
→ filtro classe = VIP

"Quais plataformas da classe ESPECIAL mais faturaram?"
→ agrupar_por = plataforma
→ filtro classe = ESPECIAL

"Quais clientes da plataforma GAZIN da classe VIP+ mais faturaram?"
→ agrupar_por = cliente_plataforma
→ filtro plataforma = GAZIN
→ filtro classe = VIP+


============================================================
RANKING
============================================================

"qual região mais faturou"

operacao = ranking
agrupar_por = regiao
top_n = 1
ordem = desc


"top 5 clientes"

operacao = ranking
agrupar_por = cliente
top_n = 5
ordem = desc


"top 10 representantes"

operacao = ranking
agrupar_por = representante
top_n = 10
ordem = desc


"5 famílias que mais faturaram"

operacao = ranking
agrupar_por = familia
top_n = 5
ordem = desc


"3 linhas que menos faturaram"

operacao = ranking
agrupar_por = linha
top_n = 3
ordem = asc


Se ranking não informar quantidade:

top_n = 20


Se ranking não indicar menor:

ordem = desc


============================================================
ESCOPO DO AGENTE
============================================================

Este agente responde SOMENTE perguntas relacionadas
aos indicadores comerciais disponíveis no dashboard.

Estão dentro do escopo assuntos como:

- faturamento
- meta
- atingimento da meta
- margem líquida
- margem bruta
- quantidade vendida
- meta de quantidade / meta de vendas
- entregas
- faturamento + entregas
- meta de margem líquida
- carteira faturável dia
- carteira faturável mês
- faturamento carteira
- status de entrega
- curva ABC
- análise de crédito
- regiões
- representantes
- clientes
- produtos
- linhas
- famílias
- plataformas
- classes
- rankings e comparações desses indicadores

Se a pergunta não estiver relacionada ao contexto
comercial disponível no agente, retorne:

"fora_escopo": true

Não tente transformar uma pergunta sem relação
com o negócio em uma consulta comercial.

Quando a pergunta estiver dentro do escopo:

"fora_escopo": false


============================================================
REGRA OBRIGATÓRIA PARA "POR ..."
============================================================

Quando a pergunta pedir um indicador "por" alguma dimensão,
a operação deve ser ranking.

Exemplos:

"Qual o faturamento por região?"
→ operacao = ranking
→ indicador = faturamento
→ agrupar_por = regiao

"Qual o faturamento de agosto de 2026 por região?"
→ operacao = ranking
→ indicador = faturamento
→ agrupar_por = regiao
→ filtros.ano = 2026
→ filtros.mes = 08

"Qual a margem líquida por plataforma?"
→ operacao = ranking
→ indicador = margem_liquida
→ agrupar_por = plataforma

"Qual a meta por representante?"
→ operacao = ranking
→ indicador = meta_faturamento
→ agrupar_por = representante

"Qual o faturamento por filial?"
→ operacao = ranking
→ agrupar_por = regiao

IMPORTANTE:
"por filial" e "por região" representam a mesma dimensão.
Use agrupar_por = regiao.

Se existir "por <dimensão>" na pergunta,
NÃO use operacao = valor.






============================================================
RESUMO DE META
============================================================

Use operacao = resumo_meta quando o usuário pedir,
na mesma pergunta, informações combinadas sobre:

- meta
- realizado/faturamento
- percentual atingido
- quanto falta para a meta

Exemplos:

"Qual a meta geral de vendas para 2026 e quanto já atingimos?"

"Quanto faturamos, qual era a meta e qual o percentual atingido?"

"Como estamos em relação à meta de 2026?"

Nesses casos:

operacao = resumo_meta
indicador = faturamento
agrupar_por = null
top_n = null
ordem = null

Os filtros devem seguir normalmente a pergunta.


============================================================
MESES
============================================================

janeiro = 01
fevereiro = 02
março = 03
abril = 04
maio = 05
junho = 06
julho = 07
agosto = 08
setembro = 09
outubro = 10
novembro = 11
dezembro = 12

============================================================
TRIMESTRES
============================================================

Quando o usuário informar um trimestre, converta o trimestre
para uma lista de meses no campo filtros.mes.

1º trimestre / primeiro trimestre / Q1:
"mes": ["01", "02", "03"]

2º trimestre / segundo trimestre / Q2:
"mes": ["04", "05", "06"]

3º trimestre / terceiro trimestre / Q3:
"mes": ["07", "08", "09"]

4º trimestre / quarto trimestre / Q4:
"mes": ["10", "11", "12"]

Exemplos:

"Me dê o resumo gerencial do primeiro trimestre de 2026"
→ operacao = resumo_gerencial
→ filtros.ano = "2026"
→ filtros.mes = ["01", "02", "03"]

"Como foi o faturamento no segundo trimestre?"
→ operacao = valor
→ indicador = faturamento
→ filtros.mes = ["04", "05", "06"]

"Qual foi a margem líquida no Q3 de 2026?"
→ operacao = valor
→ indicador = margem_liquida
→ filtros.ano = "2026"
→ filtros.mes = ["07", "08", "09"]




============================================================
REGRAS
============================================================

Use a chave "filtros" no plural.

Ano deve ser retornado como texto.
Exemplo: "2026".

Região em maiúsculas.

Não invente filtros.

Não coloque os defaults do Power BI.

Perguntas completas devem refletir
somente aquilo que o usuário pediu,
salvo quando forem claramente continuação.


{contexto_anterior}


PERGUNTA:

{pergunta}
"""


# ============================================================
# 20. GROQ
# ============================================================

def interpretar_com_groq(
    pergunta
):

    prompt = montar_prompt_ia(
        pergunta
    )


    resposta = (
        groq_client
        .chat
        .completions
        .create(

            model=GROQ_MODEL,

            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ],

            response_format={
                "type": "json_object"
            },

            temperature=0
        )
    )


    texto = (
        resposta
        .choices[0]
        .message
        .content
    )


    dados = json.loads(
        texto
    )


    dados = normalizar_json_ia(
        dados
    )


    interpretacao = (
        InterpretacaoPergunta(
            **dados
        )
    )


    return normalizar_interpretacao(
        interpretacao
    )


# ============================================================
# 21. GEMINI FALLBACK
# ============================================================

def interpretar_com_gemini(
    pergunta
):

    if gemini_client is None:

        raise RuntimeError(
            "Gemini indisponível."
        )


    prompt = montar_prompt_ia(
        pergunta
    )


    resposta = (
        gemini_client
        .models
        .generate_content(

            model=GEMINI_MODEL,

            contents=prompt,

            config={
                "response_mime_type":
                    "application/json",

                "response_schema":
                    InterpretacaoPergunta
            }
        )
    )


    return normalizar_interpretacao(
        resposta.parsed
    )


# ============================================================
# 22. ROTEADOR
# ============================================================

def corrigir_agrupamento_explicito(
    pergunta,
    interpretacao
):

    texto = (
        pergunta
        .strip()
        .lower()
    )

    regras = {

        "regiao": [
            "por região",
            "por regiao",
            "por regiões",
            "por regioes",
            "por filial",
            "por filiais"
        ],

        "representante": [
            "por representante",
            "por representantes"
        ],

        "cliente": [
            "por cliente",
            "por clientes"
        ],

        "produto": [
            "por produto",
            "por produtos"
        ],

        "linha": [
            "por linha",
            "por linhas"
        ],

        "familia": [
            "por família",
            "por familia",
            "por famílias",
            "por familias"
        ],

        "plataforma": [
            "por plataforma",
            "por plataformas"
        ],

        "classe": [
            "por classe",
            "por classes"
        ]
    }

    for dimensao, termos in regras.items():

        if any(
            termo in texto
            for termo in termos
        ):

            interpretacao.operacao = "ranking"

            interpretacao.agrupar_por = dimensao

            interpretacao.top_n = (
                interpretacao.top_n
                or 20
            )

            interpretacao.ordem = (
                interpretacao.ordem
                or "desc"
            )

            break

    return normalizar_interpretacao(
        interpretacao
    )



def _pergunta_pede_ano_inteiro(pergunta):
    """
    Detecta quando o usuário quer o acumulado/total do ano,
    sem restringir ao mês atual.
    """
    texto = _texto_normalizado(pergunta)

    padroes = [
        r"\btotal\s+(?:do|no)\s+ano\b",
        r"\bacumulad[oa]\s+(?:do|no)\s+ano\b",
        r"\bacumulado\s+anual\b",
        r"\bano\s+inteiro\b",
        r"\bano\s+atual\b",
        r"\bno\s+ano\b",
        r"\bao\s+longo\s+do\s+ano\b",
        r"\bdurante\s+o\s+ano\b",
        r"\bresultado\s+(?:do|no)\s+ano\b",
        r"\bfaturamento\s+(?:do|no)\s+ano\b",
        r"\bmargem\s+(?:liquida|bruta)\s+(?:do|no)\s+ano\b",
        r"\bmeta\s+(?:do|no)\s+ano\b",
        r"\bquantidade\s+(?:do|no)\s+ano\b",
    ]

    return any(
        re.search(padrao, texto)
        for padrao in padroes
    )


# ============================================================
# PERÍODOS EM NÍVEL DE DIA
# ============================================================

def _data_atual_negocio():
    """
    Data de referência do agente no fuso da operação no Brasil.
    Evita que "hoje"/"ontem" mudem antes da hora por causa do UTC
    do container.
    """
    try:
        return datetime.now(
            ZoneInfo("America/Fortaleza")
        ).date()
    except Exception:
        return datetime.now().date()


def _ultimo_dia_mes(ano, mes):
    if mes == 12:
        proximo = date(ano + 1, 1, 1)
    else:
        proximo = date(ano, mes + 1, 1)

    return proximo - timedelta(days=1)


def _mes_ano_explicitos_para_periodo(pergunta):
    """
    Extrai mês/ano citados. Quando um deles não é informado,
    usa o mês/ano atuais, conforme regra de negócio solicitada.
    """
    texto = _texto_normalizado(pergunta)
    hoje = _data_atual_negocio()

    meses = {
        "janeiro": 1,
        "fevereiro": 2,
        "marco": 3,
        "abril": 4,
        "maio": 5,
        "junho": 6,
        "julho": 7,
        "agosto": 8,
        "setembro": 9,
        "outubro": 10,
        "novembro": 11,
        "dezembro": 12,
    }

    mes = None
    for nome, numero in meses.items():
        if re.search(r"\b" + re.escape(nome) + r"\b", texto):
            mes = numero
            break

    match_ano = re.search(r"\b(20\d{2})\b", texto)
    ano = int(match_ano.group(1)) if match_ano else None

    # Referências relativas de mês. Só entram quando não há mês nominal explícito.
    if mes is None:
        deslocamento = None

        if re.search(r"\b(?:mes atual|este mes|neste mes|nesse mes)\b", texto):
            deslocamento = 0
        elif re.search(r"\b(?:mes passado|mes anterior)\b", texto):
            deslocamento = 1
        elif re.search(r"\bmes retrasado\b", texto):
            deslocamento = 2
        else:
            match_ha_meses = re.search(r"\bha\s+(\d+)\s+mes(?:es)?\b", texto)
            if match_ha_meses:
                deslocamento = int(match_ha_meses.group(1))

        if deslocamento is not None:
            total_meses = hoje.year * 12 + (hoje.month - 1) - deslocamento
            ano = total_meses // 12
            mes = total_meses % 12 + 1

    return {
        "mes": mes if mes is not None else hoje.month,
        "ano": ano if ano is not None else hoje.year,
        "mes_explicito": mes is not None,
        "ano_explicito": ano is not None,
    }



def _periodo_relativo_amplo(pergunta):
    """
    Resolve períodos relativos completos sem alterar as regras antigas:
    ano, trimestre, semestre e bimestre atual/passado/retrasado,
    além de "há N anos/trimestres/semestres/bimestres".
    Retorna None quando a pergunta não contém um desses períodos.
    """
    texto = _texto_normalizado(pergunta)
    hoje = _data_atual_negocio()

    def deslocamento_unidade(unidade):
        singular = unidade
        plural = {
            "ano": "anos",
            "trimestre": "trimestres",
            "semestre": "semestres",
            "bimestre": "bimestres",
        }[unidade]

        if re.search(
            rf"\b(?:{singular}\s+atual|este\s+{singular}|neste\s+{singular}|nesse\s+{singular})\b",
            texto
        ):
            return 0

        if re.search(
            rf"\b(?:{singular}\s+passado|{singular}\s+anterior)\b",
            texto
        ):
            return 1

        if re.search(rf"\b{singular}\s+retrasado\b", texto):
            return 2

        m = re.search(rf"\bha\s+(\d+)\s+{singular}(?:s)?\b", texto)
        if m:
            return int(m.group(1))

        return None

    # MÊS
    # Mantém esta regra dentro do mesmo interpretador temporal para que
    # "mês atual", "mês passado", "mês anterior", "mês retrasado"
    # e "há X meses" sejam convertidos em uma faixa real de datas.
    texto_mes = texto

    desloc_mes = None

    if re.search(
        r"\b(?:mes\s+atual|este\s+mes|neste\s+mes|nesse\s+mes)\b",
        texto_mes
    ):
        desloc_mes = 0

    elif re.search(
        r"\b(?:mes\s+passado|mes\s+anterior)\b",
        texto_mes
    ):
        desloc_mes = 1

    elif re.search(
        r"\bmes\s+retrasado\b",
        texto_mes
    ):
        desloc_mes = 2

    else:
        m_mes = re.search(
            r"\bha\s+(\d+)\s+mes(?:es)?\b",
            texto_mes
        )

        if m_mes:
            desloc_mes = int(m_mes.group(1))

    if desloc_mes is not None:
        total_meses = (
            hoje.year * 12
            + (hoje.month - 1)
            - desloc_mes
        )

        ano = total_meses // 12
        mes = total_meses % 12 + 1

        inicio = date(
            ano,
            mes,
            1
        )

        fim = _ultimo_dia_mes(
            ano,
            mes
        )

        nome_mes = meses_nome.get(
            f"{mes:02d}",
            f"{mes:02d}"
        )

        rotulo_base = (
            "neste mês"
            if desloc_mes == 0
            else "no mês passado"
            if desloc_mes == 1
            else "no mês retrasado"
            if desloc_mes == 2
            else f"há {desloc_mes} meses"
        )

        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": (
                f"{rotulo_base} "
                f"({nome_mes} de {ano})"
            ),
            "tipo": "mes",
        }

    # ANO
    desloc = deslocamento_unidade("ano")
    if desloc is not None:
        ano = hoje.year - desloc
        inicio = date(ano, 1, 1)
        fim = date(ano, 12, 31)
        rotulo_base = (
            "neste ano" if desloc == 0
            else "no ano passado" if desloc == 1
            else "no ano retrasado" if desloc == 2
            else f"há {desloc} anos"
        )
        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": f"{rotulo_base} ({ano})",
            "tipo": "ano",
        }

    # TRIMESTRE / SEMESTRE / BIMESTRE
    configuracoes = {
        "trimestre": 3,
        "semestre": 6,
        "bimestre": 2,
    }

    for unidade, tamanho_meses in configuracoes.items():
        desloc = deslocamento_unidade(unidade)
        if desloc is None:
            continue

        indice_atual = (hoje.month - 1) // tamanho_meses
        total_periodos_ano = 12 // tamanho_meses
        absoluto = hoje.year * total_periodos_ano + indice_atual - desloc

        ano = absoluto // total_periodos_ano
        indice = absoluto % total_periodos_ano
        mes_inicio = indice * tamanho_meses + 1
        mes_fim = mes_inicio + tamanho_meses - 1

        inicio = date(ano, mes_inicio, 1)
        fim = _ultimo_dia_mes(ano, mes_fim)

        numero_periodo = indice + 1
        nome = {
            "trimestre": "trimestre",
            "semestre": "semestre",
            "bimestre": "bimestre",
        }[unidade]

        rotulo_base = (
            f"neste {nome}" if desloc == 0
            else f"no {nome} passado" if desloc == 1
            else f"no {nome} retrasado" if desloc == 2
            else f"há {desloc} {nome}{'s' if desloc != 1 else ''}"
        )

        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": (
                f"{rotulo_base} "
                f"({numero_periodo}º {nome} de {ano}: "
                f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')})"
            ),
            "tipo": unidade,
        }

    return None


def _periodo_diario_pergunta(pergunta):
    """
    Converte expressões temporais em um intervalo fechado de datas.

    Retorno:
        None, quando não há período diário especial; ou
        {"inicio": date, "fim": date, "rotulo": str}
    """
    texto = _texto_normalizado(pergunta)
    hoje = _data_atual_negocio()

    # --------------------------------------------------------
    # ANO / TRIMESTRE / SEMESTRE / BIMESTRE RELATIVOS
    # --------------------------------------------------------
    periodo_amplo = _periodo_relativo_amplo(pergunta)
    if periodo_amplo is not None:
        return periodo_amplo

    # --------------------------------------------------------
    # DATA EXPLÍCITA: 13/08/2026, 13-08-2026, dia 13/08
    # --------------------------------------------------------
    m = re.search(
        r"\b(?:dia\s+)?([0-3]?\d)[/-]([01]?\d)(?:[/-](20\d{2}))?\b",
        texto
    )
    if m:
        dia = int(m.group(1))
        mes = int(m.group(2))
        ano = int(m.group(3)) if m.group(3) else hoje.year

        try:
            alvo = date(ano, mes, dia)
        except ValueError:
            return None

        return {
            "inicio": alvo,
            "fim": alvo,
            "rotulo": f"no dia {alvo.strftime('%d/%m/%Y')}"
        }

    # --------------------------------------------------------
    # DATA POR EXTENSO: dia 13 de agosto de 2026
    # --------------------------------------------------------
    meses = {
        "janeiro": 1, "fevereiro": 2, "marco": 3,
        "abril": 4, "maio": 5, "junho": 6,
        "julho": 7, "agosto": 8, "setembro": 9,
        "outubro": 10, "novembro": 11, "dezembro": 12
    }
    nomes_meses_regex = "|".join(meses.keys())

    m = re.search(
        rf"\b(?:dia\s+)?([0-3]?\d)\s+de\s+({nomes_meses_regex})"
        rf"(?:\s+de\s+(20\d{{2}}))?\b",
        texto
    )
    if m:
        dia = int(m.group(1))
        mes = meses[m.group(2)]
        ano = int(m.group(3)) if m.group(3) else hoje.year

        try:
            alvo = date(ano, mes, dia)
        except ValueError:
            return None

        return {
            "inicio": alvo,
            "fim": alvo,
            "rotulo": f"no dia {alvo.strftime('%d/%m/%Y')}"
        }

    # --------------------------------------------------------
    # SOMENTE DIA: "dia 13" -> mês e ano atuais
    # --------------------------------------------------------
    m = re.search(r"\bdia\s+([0-3]?\d)\b", texto)
    if m:
        dia = int(m.group(1))

        try:
            alvo = date(hoje.year, hoje.month, dia)
        except ValueError:
            return None

        return {
            "inicio": alvo,
            "fim": alvo,
            "rotulo": f"no dia {alvo.strftime('%d/%m/%Y')}"
        }

    # --------------------------------------------------------
    # HOJE / ONTEM
    # --------------------------------------------------------
    if re.search(r"\bhoje\b", texto):
        return {
            "inicio": hoje,
            "fim": hoje,
            "rotulo": f"hoje ({hoje.strftime('%d/%m/%Y')})"
        }

    if re.search(r"\bontem\b", texto):
        alvo = hoje - timedelta(days=1)
        return {
            "inicio": alvo,
            "fim": alvo,
            "rotulo": f"ontem ({alvo.strftime('%d/%m/%Y')})"
        }

    # --------------------------------------------------------
    # SEMANA PASSADA / ESTA SEMANA
    # Semana comercial considerada de segunda a domingo.
    # Para a semana atual, vai de segunda até hoje.
    # --------------------------------------------------------
    if re.search(r"\bsemana\s+passada\b", texto):
        inicio_semana_atual = hoje - timedelta(days=hoje.weekday())
        fim = inicio_semana_atual - timedelta(days=1)
        inicio = fim - timedelta(days=6)

        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": (
                "na semana passada "
                f"({inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')})"
            )
        }

    if re.search(
        r"\b(?:esta|essa|desta|dessa)\s+semana\b|\bsemana\s+atual\b",
        texto
    ):
        inicio = hoje - timedelta(days=hoje.weekday())
        fim = hoje

        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": (
                "nesta semana "
                f"({inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')})"
            )
        }

    referencia = _mes_ano_explicitos_para_periodo(pergunta)
    ano = referencia["ano"]
    mes = referencia["mes"]
    ultimo_mes = _ultimo_dia_mes(ano, mes)
    nome_mes = meses_nome.get(f"{mes:02d}", f"{mes:02d}")

    # --------------------------------------------------------
    # PRIMEIRA / SEGUNDA QUINZENA
    # --------------------------------------------------------
    if re.search(r"\bprimeira\s+quinzena\b|\b1[ªa]?\s+quinzena\b", texto):
        inicio = date(ano, mes, 1)
        fim = date(ano, mes, 15)

        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": f"na primeira quinzena de {nome_mes} de {ano}"
        }

    if re.search(r"\bsegunda\s+quinzena\b|\b2[ªa]?\s+quinzena\b", texto):
        inicio = date(ano, mes, 16)
        fim = ultimo_mes

        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": f"na segunda quinzena de {nome_mes} de {ano}"
        }

    # --------------------------------------------------------
    # PRIMEIRO / ÚLTIMO DIA DO MÊS
    # --------------------------------------------------------
    if re.search(r"\bprimeiro\s+dia\b|\b1[oº]?\s+dia\b", texto):
        alvo = date(ano, mes, 1)
        return {
            "inicio": alvo,
            "fim": alvo,
            "rotulo": f"no primeiro dia de {nome_mes} de {ano}"
        }

    if re.search(r"\bultimo\s+dia\b", texto):
        alvo = ultimo_mes
        return {
            "inicio": alvo,
            "fim": alvo,
            "rotulo": f"no último dia de {nome_mes} de {ano}"
        }

    # --------------------------------------------------------
    # ÚLTIMOS X DIAS
    # Sem mês/ano explícitos: termina hoje.
    # Com mês e/ou ano explícitos: termina no último dia daquele mês.
    # --------------------------------------------------------
    m = re.search(r"\bultimos\s+(\d{1,3})\s+dias\b", texto)
    if m:
        qtd = int(m.group(1))

        if not (1 <= qtd <= 366):
            return None

        if referencia["mes_explicito"] or referencia["ano_explicito"]:
            fim = ultimo_mes
        else:
            fim = hoje

        inicio = fim - timedelta(days=qtd - 1)

        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": (
                f"nos últimos {qtd} dias "
                f"({inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')})"
            )
        }

    return None


def _periodo_relativo_da_interpretacao(interpretacao):
    """
    Converte a intenção temporal normalizada pela IA em datas reais.

    A IA informa somente a unidade e o deslocamento. O cálculo de datas
    continua determinístico no Python, evitando que expressões como
    "último mês" dependam de a IA adivinhar que mês é esse.
    """
    if interpretacao is None:
        return None

    periodo_ia = getattr(interpretacao, "periodo_relativo", None)
    if periodo_ia is None:
        return None

    unidade = getattr(periodo_ia, "unidade", None)

    try:
        deslocamento = int(getattr(periodo_ia, "deslocamento", 0))
    except (TypeError, ValueError):
        return None

    if deslocamento < 0 or deslocamento > 120:
        return None

    hoje = _data_atual_negocio()

    if unidade == "mes":
        total_meses = hoje.year * 12 + (hoje.month - 1) - deslocamento
        ano = total_meses // 12
        mes = total_meses % 12 + 1
        inicio = date(ano, mes, 1)
        fim = _ultimo_dia_mes(ano, mes)
        nome_mes = meses_nome.get(f"{mes:02d}", f"{mes:02d}")
        return {
            "inicio": inicio,
            "fim": fim,
            "rotulo": f"em {nome_mes} de {ano}",
            "tipo": "mes",
        }

    if unidade == "ano":
        ano = hoje.year - deslocamento
        return {
            "inicio": date(ano, 1, 1),
            "fim": date(ano, 12, 31),
            "rotulo": f"em {ano}",
            "tipo": "ano",
        }

    tamanhos = {
        "trimestre": 3,
        "semestre": 6,
        "bimestre": 2,
    }

    tamanho_meses = tamanhos.get(unidade)
    if tamanho_meses is None:
        return None

    indice_atual = (hoje.month - 1) // tamanho_meses
    total_periodos_ano = 12 // tamanho_meses
    absoluto = hoje.year * total_periodos_ano + indice_atual - deslocamento
    ano = absoluto // total_periodos_ano
    indice = absoluto % total_periodos_ano
    mes_inicio = indice * tamanho_meses + 1
    mes_fim = mes_inicio + tamanho_meses - 1
    inicio = date(ano, mes_inicio, 1)
    fim = _ultimo_dia_mes(ano, mes_fim)

    return {
        "inicio": inicio,
        "fim": fim,
        "rotulo": (
            f"no {indice + 1}º {unidade} de {ano} "
            f"({inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')})"
        ),
        "tipo": unidade,
    }


def _aplicar_periodo_calendario(pergunta, filtros, interpretacao=None):
    """
    Injeta o intervalo diário nos filtros finais usados por qualquer
    indicador/operação. Os filtros de ano/mês são retirados para não
    limitar intervalos que atravessam mês ou ano.

    Prioridade:
    1) regras temporais locais já existentes;
    2) período relativo normalizado pela IA, apenas como fallback.
    """
    periodo = _periodo_diario_pergunta(pergunta)

    if periodo is None:
        periodo = _periodo_relativo_da_interpretacao(interpretacao)

    if periodo is None:
        return dict(filtros or {})

    novos = dict(filtros or {})

    novos.pop("ano", None)
    novos.pop("mes", None)

    novos["_data_inicio"] = periodo["inicio"].isoformat()
    novos["_data_fim"] = periodo["fim"].isoformat()
    novos["_periodo_rotulo"] = periodo["rotulo"]

    return novos


def corrigir_periodo_explicito(
    pergunta,
    interpretacao
):

    texto = (
        pergunta
        .strip()
        .lower()
    )

    # --------------------------------------------------------
    # IDENTIFICA MÊS INFORMADO NA PERGUNTA
    # --------------------------------------------------------

    mes_encontrado = None

    for nome, numero in mapa_meses.items():

        if re.search(r"\b" + re.escape(nome) + r"\b", texto):

            mes_encontrado = numero

            break

    # --------------------------------------------------------
    # IDENTIFICA ANO INFORMADO
    # --------------------------------------------------------

    match_ano = re.search(
        r"\b(20\d{2})\b",
        texto
    )

    ano_encontrado = (
        match_ano.group(1)
        if match_ano
        else None
    )

    # --------------------------------------------------------
    # MÊS RELATIVO: atual / passado / retrasado / há N meses
    # --------------------------------------------------------
    texto_normalizado = _texto_normalizado(pergunta)
    deslocamento_mes = None

    if re.search(r"\b(?:mes atual|este mes|neste mes|nesse mes)\b", texto_normalizado):
        deslocamento_mes = 0
    elif re.search(r"\b(?:mes passado|mes anterior)\b", texto_normalizado):
        deslocamento_mes = 1
    elif re.search(r"\bmes retrasado\b", texto_normalizado):
        deslocamento_mes = 2
    else:
        match_ha_meses = re.search(r"\bha\s+(\d+)\s+mes(?:es)?\b", texto_normalizado)
        if match_ha_meses:
            deslocamento_mes = int(match_ha_meses.group(1))

    if deslocamento_mes is not None:
        hoje = _data_atual_negocio()
        total_meses = hoje.year * 12 + (hoje.month - 1) - deslocamento_mes
        ano_relativo = total_meses // 12
        mes_relativo = total_meses % 12 + 1

        interpretacao.filtros.ano = str(ano_relativo)
        interpretacao.filtros.mes = f"{mes_relativo:02d}"

        return normalizar_interpretacao(interpretacao)

    pede_ano_inteiro = (
        _pergunta_pede_ano_inteiro(
            pergunta
        )
    )

    # --------------------------------------------------------
    # MÊS + ANO
    # --------------------------------------------------------

    if (
        mes_encontrado
        and ano_encontrado
    ):

        interpretacao.filtros.mes = (
            mes_encontrado
        )

        interpretacao.filtros.ano = (
            ano_encontrado
        )

    # --------------------------------------------------------
    # SOMENTE MÊS
    # --------------------------------------------------------

    elif (
        mes_encontrado
        and not ano_encontrado
    ):

        interpretacao.filtros.mes = (
            mes_encontrado
        )

        interpretacao.filtros.ano = (
            "Ano atual"
        )

    # --------------------------------------------------------
    # SOMENTE ANO / PEDIDO EXPLÍCITO DO ANO INTEIRO
    # --------------------------------------------------------

    elif (
        not mes_encontrado
        and (
            ano_encontrado
            or pede_ano_inteiro
        )
    ):

        interpretacao.filtros.ano = (
            ano_encontrado
            if ano_encontrado
            else "Ano atual"
        )

        interpretacao.filtros.mes = None

    return normalizar_interpretacao(
        interpretacao
    )


def interpretar_simples_localmente(pergunta):

    texto = (
        pergunta
        .strip()
        .lower()
    )

    # Resumo/visão geral pertence ao fluxo de resumo_gerencial.
    # Não tentar transformar em valor simples.
    if any(
        termo in texto
        for termo in [
            "resumo gerencial",
            "visão geral",
            "visao geral"
        ]
    ):
        return None

    # ========================================================
    # PEDIDO COMBINADO DE META
    # ========================================================

    pediu_percentual_meta = any(
        termo in texto
        for termo in [
            "% da meta",
            "% sobre a meta",
            "percentual da meta",
            "percentual sobre a meta",
            "atingimento da meta",
            "percentual atingido",
            "quanto atingiu",
            "quanto atingimos"
        ]
    )

    pediu_faturamento = any(
        termo in texto
        for termo in [
            "faturamento",
            "faturou",
            "faturamos",
            "realizado"
        ]
    )

    pediu_desvio_meta = any(
        termo in texto
        for termo in [
            "quanto falta",
            "desvio da meta",
            "diferença para a meta"
        ]
    )

    # A palavra "meta" dentro de "% da meta" NÃO deve contar como
    # pedido separado de meta. Só resumo quando o usuário pediu
    # realmente dois ou mais conceitos distintos.
    texto_sem_percentual = texto

    for trecho in [
        "% da meta",
        "% sobre a meta",
        "percentual da meta",
        "percentual sobre a meta",
        "atingimento da meta",
        "percentual atingido",
        "quanto atingiu",
        "quanto atingimos"
    ]:
        texto_sem_percentual = texto_sem_percentual.replace(
            trecho,
            " "
        )

    pediu_meta_base = bool(
        re.search(
            r"\bmeta\b",
            texto_sem_percentual
        )
    )

    qtd_conceitos = sum([
        1 if pediu_meta_base else 0,
        1 if pediu_percentual_meta else 0,
        1 if pediu_faturamento else 0,
        1 if pediu_desvio_meta else 0
    ])

    if qtd_conceitos >= 2:

        filtros = {}

        for nome, numero in mapa_meses.items():
            if re.search(r"\b" + re.escape(nome) + r"\b", texto):
                filtros["mes"] = numero
                break

        match_ano = re.search(r"\b(20\d{2})\b", texto)
        if match_ano:
            filtros["ano"] = match_ano.group(1)

        texto_upper = pergunta.upper()
        for regiao in regioes_conhecidas:
            if regiao in texto_upper:
                filtros["regiao"] = regiao
                break

        try:
            interpretacao = InterpretacaoPergunta(
                operacao="resumo_meta",
                indicador="faturamento",
                filtros=filtros,
                agrupar_por=None,
                top_n=None,
                ordem=None
            )

            return normalizar_interpretacao(interpretacao)

        except Exception:
            return None

    # ========================================================
    # INDICADOR
    # ========================================================

    indicador = None

    # Previsão de faturamento é um indicador próprio e precisa ser
    # reconhecida antes da palavra "faturamento" isolada.
    if any(x in texto for x in [
        "previsão de faturamento",
        "previsao de faturamento",
        "projeção de faturamento",
        "projecao de faturamento",
        "forecast"
    ]):
        indicador = "previsao_faturamento"

    elif any(x in texto for x in [
        "percentual da meta de margem líquida",
        "percentual da meta de margem liquida",
        "meta de margem líquida",
        "meta de margem liquida"
    ]):
        indicador = "meta_margem_liquida"

    elif any(x in texto for x in [
        "faturamento + entregas",
        "fat + entregas",
        "faturamento e entregas"
    ]):
        indicador = "faturamento_entregas"

    elif any(x in texto for x in [
        "carteira faturável do dia",
        "carteira faturavel do dia",
        "carteira faturável dia",
        "carteira faturavel dia"
    ]):
        indicador = "carteira_faturavel_dia"

    elif any(x in texto for x in [
        "carteira faturável do mês",
        "carteira faturavel do mes",
        "carteira faturável mês",
        "carteira faturavel mes"
    ]):
        indicador = "carteira_faturavel_mes"

    elif any(x in texto for x in [
        "faturamento da carteira",
        "faturamento carteira"
    ]):
        indicador = "faturamento_carteira"

    elif any(x in texto for x in [
        "meta de quantidade",
        "meta em quantidade",
        "meta de vendas"
    ]):
        indicador = "meta_quantidade"

    elif any(x in texto for x in [
        "quantidade vendida",
        "quantidade de vendas",
        "qtd vendida",
        " qtd "
    ]):
        indicador = "quantidade"

    elif "entregas" in texto:
        indicador = "entregas"

    elif (
        "percentual sobre a meta" in texto
        or "% sobre a meta" in texto
        or "percentual da meta" in texto
        or "% da meta" in texto
        or "atingimento da meta" in texto
    ):

        indicador = (
            "atingimento_meta_faturamento"
        )

    elif (
        "meta de faturamento" in texto
        or "meta de vendas" in texto
    ):

        indicador = "meta_faturamento"

    elif (
        "margem líquida" in texto
        or "margem liquida" in texto
    ):

        if (
            "r$" in texto
            or "reais" in texto
            or "valor" in texto
        ):
            indicador = (
                "valor_margem_liquida"
            )
        else:
            indicador = "margem_liquida"

    elif "margem bruta" in texto:

        if (
            "r$" in texto
            or "reais" in texto
            or "valor" in texto
        ):
            indicador = (
                "valor_margem_bruta"
            )
        else:
            indicador = "margem_bruta"

    elif (
        "faturamento" in texto
        or "faturou" in texto
        or "faturamos" in texto
    ):

        indicador = "faturamento"

    if indicador is None:

        return None

    # ========================================================
    # MÊS
    # ========================================================

    mes = None

    for nome, numero in mapa_meses.items():

        if re.search(
            r"\b" + re.escape(nome) + r"\b",
            texto
        ):

            mes = numero
            break

    # ========================================================
    # ANO
    # ========================================================

    ano = None

    match_ano = re.search(
        r"\b(20\d{2})\b",
        texto
    )

    if match_ano:

        ano = match_ano.group(1)

    # ========================================================
    # NÃO TRATAR LOCALMENTE RANKINGS / AGRUPAMENTOS
    # ========================================================

    termos_agrupamento = [
        " por região",
        " por regiao",
        " por filial",
        " por cliente",
        " por produto",
        " por linha",
        " por família",
        " por familia",
        " por representante",
        " por plataforma",
        " por status de entrega",
        " por status_entrega",
        " por curva abc",
        " por análise de crédito",
        " por analise de credito",
        "ranking",
        "top "
    ]

    if any(
        termo in texto
        for termo in termos_agrupamento
    ):

        return None

    # ========================================================
    # CRIAR INTERPRETAÇÃO
    # ========================================================

    filtros = {}

    # Região explícita: o parser simples antigo reconhecia o
    # indicador, mas ignorava a região (ex.: "região Sul").
    texto_upper = pergunta.upper()
    regiao_encontrada = None

    for regiao in regioes_conhecidas:
        if regiao in texto_upper:
            regiao_encontrada = regiao
            break

    if regiao_encontrada:
        filtros["regiao"] = regiao_encontrada

    # Se o usuário citou uma dimensão que o parser simples não
    # consegue identificar com segurança, deixa a IA interpretar
    # em vez de consultar o BI sem o filtro solicitado.
    termos_entidade_nao_tratados = [
        "cliente ",
        "produto ",
        "linha ",
        "família ",
        "familia ",
        "representante ",
        "plataforma ",
        "classe "
    ]

    if any(termo in texto for termo in termos_entidade_nao_tratados):
        return None

    if mes:
        filtros["mes"] = mes

    if ano:
        filtros["ano"] = ano

    try:

        interpretacao = (
            InterpretacaoPergunta(
                operacao="valor",
                indicador=indicador,
                filtros=filtros,
                agrupar_por=None,
                top_n=None,
                ordem=None
            )
        )

        return normalizar_interpretacao(
            interpretacao
        )

    except Exception:

        return None



def pergunta_analitica_complexa(pergunta):
    texto = pergunta.lower().strip()

    palavras_comparacao = [
        "compare",
        "comparar",
        "comparação",
        "versus",
        " vs ",
        "diferença entre"
    ]

    palavras_ranking = [
        "top ",
        "maiores",
        "menores",
        "maior faturamento",
        "menor faturamento",
        "ranking",
        "melhores",
        "piores"
    ]

    return (
        any(p in texto for p in palavras_comparacao)
        or
        any(p in texto for p in palavras_ranking)
    )



def interpretar_com_claude(pergunta):

    api_key = os.getenv("ANTHROPIC_API_KEY")

    if not api_key:
        raise RuntimeError("Claude indisponível: ANTHROPIC_API_KEY não encontrada.")

    prompt = montar_prompt_ia(pergunta)

    resposta = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
        },
        json={
            "model": CLAUDE_MODEL,
            "max_tokens": 2048,
            "temperature": 0,
            "messages": [
                {"role": "user", "content": prompt}
            ]
        },
        timeout=60
    )

    if resposta.status_code != 200:
        raise RuntimeError(
            f"Claude HTTP {resposta.status_code}: {resposta.text[:500]}"
        )

    dados = resposta.json()
    blocos = dados.get("content", [])
    texto = "".join(
        bloco.get("text", "")
        for bloco in blocos
        if bloco.get("type") == "text"
    ).strip()

    if not texto:
        raise RuntimeError("Claude retornou resposta vazia.")

    texto = limpar_json_texto(texto)
    interpretacao = InterpretacaoPergunta(**json.loads(texto))
    return normalizar_interpretacao(interpretacao)



def _interpretar_ajustes_hoje(pergunta):
    """
    Rota local SOMENTE para os ajustes novos.
    Se não reconhecer com segurança, retorna None e deixa
    o fluxo estável original tratar a pergunta.
    """

    t = _texto_normalizado(pergunta)

    # Resumo gerencial simples por mês/ano pode ser resolvido
    # diretamente sem depender do contexto anterior.
    if (
        "resumo gerencial" in t
        or "visao geral" in t
    ):
        # Trimestre/semestre continuam no fluxo original,
        # pois a base já possui tratamento específico para eles.
        if not any(x in t for x in [
            "trimestre",
            "semestre",
            "q1",
            "q2",
            "q3",
            "q4"
        ]):
            filtros_resumo = _extrair_periodo_pergunta(
                pergunta
            )

            return normalizar_interpretacao(
                InterpretacaoPergunta(
                    operacao="resumo_gerencial",
                    indicador="faturamento",
                    filtros=filtros_resumo,
                    agrupar_por=None,
                    top_n=None,
                    ordem=None
                )
            )

        return None

    # Não tocar em famílias estatísticas que já funcionavam na base.
    if any(x in t for x in [
        "qual mes",
        "que mes",
        "mes de maior",
        "mes de menor",
        "media mensal",
        "mediana mensal"
    ]):
        return None

    indicador = _indicador_da_pergunta(pergunta)
    filtros = _extrair_periodo_pergunta(pergunta)

    # --------------------------------------------------------
    # VALOR POR PRODUTO INFORMADO POR CÓDIGO OU NOME
    # --------------------------------------------------------
    # Exemplos:
    # "Qual faturamento do produto B94401062 em agosto?"
    # "Qual faturamento do produto VENTILADOR ECO TS PR GRAFITE 220V?"
    #
    # Não atua em "por produto", que continua sendo ranking.
    if (
        "produto" in t
        and not re.search(
            r"\bpor\s+produto(?:s)?\b",
            t
        )
    ):
        meses_regex = (
            "janeiro|fevereiro|marco|abril|maio|junho|"
            "julho|agosto|setembro|outubro|novembro|dezembro"
        )

        m_produto = re.search(
            r"\bproduto\s+(.+?)"
            r"(?="
            r"\s+(?:em|no|na|de)\s+(?:" + meses_regex + r"|20\d{2})\b"
            r"|[?!.,]"
            r"|$"
            r")",
            t,
            flags=re.IGNORECASE
        )

        if m_produto:
            produto = (
                m_produto
                .group(1)
                .strip(" -")
                .upper()
            )

            if produto:
                filtros["produto"] = produto

                return normalizar_interpretacao(
                    InterpretacaoPergunta(
                        operacao="valor",
                        indicador=indicador,
                        filtros=filtros,
                        agrupar_por=None,
                        top_n=None,
                        ordem=None
                    )
                )

    # --------------------------------------------------------
    # "POR CLIENTE/GRUPO/VAREJISTA/PLATAFORMA" = NIVEL 2
    # --------------------------------------------------------
    if re.search(
        r"\bpor\s+(cliente|clientes|grupo|grupos|plataforma|plataformas|"
        r"varejo|varejista|varejistas)\b",
        t
    ):
        return normalizar_interpretacao(
            InterpretacaoPergunta(
                operacao="ranking",
                indicador=indicador,
                filtros=filtros,
                agrupar_por="plataforma",
                top_n=20,
                ordem="desc"
            )
        )

    # --------------------------------------------------------
    # "POR LOJA/CNPJ" = NIVEL 3
    # --------------------------------------------------------
    if re.search(
        r"\bpor\s+(loja|lojas|cnpj|cnpjs|estabelecimento|estabelecimentos)\b",
        t
    ):
        return normalizar_interpretacao(
            InterpretacaoPergunta(
                operacao="ranking",
                indicador=indicador,
                filtros=filtros,
                agrupar_por="loja",
                top_n=20,
                ordem="desc"
            )
        )

    # --------------------------------------------------------
    # TOP 1 VAREJISTA/GRUPO/CLIENTE/PLATAFORMA
    # --------------------------------------------------------
    if any(
        re.search(p, t)
        for p in [
            r"\bqual\s+(?:a|o)?\s*varejista\b",
            r"\bqual\s+(?:a|o)?\s*grupo\b",
            r"\bqual\s+(?:a|o)?\s*cliente\b",
            r"\bqual\s+(?:a|o)?\s*plataforma\b"
        ]
    ) and any(x in t for x in [
        "mais comprou",
        "comprou mais",
        "mais vendeu",
        "vendeu mais",
        "mais faturou",
        "faturou mais"
    ]):

        ind = (
            "quantidade"
            if any(x in t for x in [
                "mais comprou",
                "comprou mais",
                "mais vendeu",
                "vendeu mais"
            ])
            else "faturamento"
        )

        return normalizar_interpretacao(
            InterpretacaoPergunta(
                operacao="ranking",
                indicador=ind,
                filtros=filtros,
                agrupar_por="plataforma",
                top_n=1,
                ordem="desc"
            )
        )

    # --------------------------------------------------------
    # LOJAS/CNPJ DE UMA PLATAFORMA
    # --------------------------------------------------------
    m = re.search(
        r"\b(?:loja|lojas|cnpj|cnpjs|estabelecimento|estabelecimentos)"
        r"(?:\s+por\s+loja)?\s+(?:da|do|de)\s+"
        r"(?:(?:plataforma|grupo|varejista|varejo|cliente)\s+)?"
        r"(.+?)"
        r"(?=\s+(?:em|no|na|mais|menos)\b|[?!.,]|$)",
        t
    )

    if m:
        nome_plataforma = m.group(1).strip().upper()

        filtros["plataforma"] = nome_plataforma

        return normalizar_interpretacao(
            InterpretacaoPergunta(
                operacao="ranking",
                indicador=indicador,
                filtros=filtros,
                agrupar_por="loja",
                top_n=20,
                ordem="desc"
            )
        )

    return None



def _interpretar_indicador_por_dimensao_generico(pergunta):
    """
    Permite consultar qualquer indicador disponível por qualquer
    dimensão conhecida, sem depender da formulação exata da IA.

    Exemplos:
    - faturamento por representantes em agosto
    - vendas por linha
    - margem líquida por família
    - previsão de faturamento por cliente
    - quantidade por curva ABC
    - representante que mais vendeu
    - linha que mais faturou
    """

    t = _texto_normalizado(
        pergunta
    )

    # Não intercepta fluxos especiais já existentes.
    if any(
        termo in t
        for termo in [
            "resumo gerencial",
            "visao geral",
            "resumo de ",
            "compare",
            "comparar",
            "comparacao",
            "versus",
            " vs ",
            "mes de maior",
            "mes de menor",
            "qual mes",
            "que mes",
            "media mensal",
            "mediana mensal"
        ]
    ):
        return None

    # --------------------------------------------------------
    # INDICADOR
    # --------------------------------------------------------
    if any(
        termo in t
        for termo in [
            "mais vendeu",
            "vendeu mais",
            "vendas por ",
            "venda por ",
            "mais comprou",
            "comprou mais"
        ]
    ):
        indicador = "quantidade"
    else:
        indicador = _indicador_da_pergunta(
            pergunta
        )

    filtros = _extrair_periodo_pergunta(
        pergunta
    )

    # --------------------------------------------------------
    # MAPA ÚNICO DE DIMENSÕES
    # --------------------------------------------------------
    regras_dimensoes = [
        (
            "regiao",
            [
                "regiao",
                "regioes",
                "filial",
                "filiais"
            ]
        ),
        (
            "representante",
            [
                "representante",
                "representantes"
            ]
        ),
        (
            "plataforma",
            [
                "cliente",
                "clientes",
                "grupo",
                "grupos",
                "plataforma",
                "plataformas",
                "varejo",
                "varejista",
                "varejistas"
            ]
        ),
        (
            "loja",
            [
                "loja",
                "lojas",
                "cnpj",
                "cnpjs",
                "estabelecimento",
                "estabelecimentos"
            ]
        ),
        (
            "produto",
            [
                "produto",
                "produtos"
            ]
        ),
        (
            "linha",
            [
                "linha",
                "linhas"
            ]
        ),
        (
            "familia",
            [
                "familia",
                "familias"
            ]
        ),
        (
            "classe",
            [
                "classe",
                "classes"
            ]
        ),
        (
            "status_entrega",
            [
                "status de entrega",
                "status_entrega"
            ]
        ),
        (
            "cod_curva_abc",
            [
                "curva abc",
                "cod_curva_abc"
            ]
        ),
        (
            "analise_credito",
            [
                "analise de credito",
                "analise_credito"
            ]
        )
    ]

    # --------------------------------------------------------
    # CASO 1: "POR <DIMENSÃO>"
    # --------------------------------------------------------
    dimensao = None

    for chave, termos in regras_dimensoes:
        for termo in termos:
            if re.search(
                r"\bpor\s+" + re.escape(termo) + r"\b",
                t
            ):
                dimensao = chave
                break

        if dimensao is not None:
            break

    if dimensao is not None:
        match_top = re.search(
            r"\btop\s*(\d+)\b",
            t
        )

        top_n = (
            min(max(int(match_top.group(1)), 1), 100)
            if match_top
            else 20
        )

        ordem = (
            "asc"
            if any(x in t for x in [
                "menos",
                "menor",
                "menores",
                "piores"
            ])
            else "desc"
        )

        return normalizar_interpretacao(
            InterpretacaoPergunta(
                operacao="ranking",
                indicador=indicador,
                filtros=filtros,
                agrupar_por=dimensao,
                top_n=top_n,
                ordem=ordem
            )
        )

    # --------------------------------------------------------
    # CASO 2: "<DIMENSÃO> QUE MAIS/MENOS..."
    # --------------------------------------------------------
    tem_extremo = any(
        termo in t
        for termo in [
            "mais vendeu",
            "vendeu mais",
            "mais faturou",
            "faturou mais",
            "mais comprou",
            "comprou mais",
            "maior faturamento",
            "menor faturamento",
            "menos vendeu",
            "vendeu menos"
        ]
    )

    if not tem_extremo:
        return None

    for chave, termos in regras_dimensoes:
        if any(
            re.search(
                r"\b" + re.escape(termo) + r"\b",
                t
            )
            for termo in termos
        ):
            dimensao = chave
            break

    if dimensao is None:
        return None

    if any(
        termo in t
        for termo in [
            "mais vendeu",
            "vendeu mais",
            "mais comprou",
            "comprou mais",
            "menos vendeu",
            "vendeu menos"
        ]
    ):
        indicador = "quantidade"

    elif any(
        termo in t
        for termo in [
            "mais faturou",
            "faturou mais",
            "maior faturamento",
            "menor faturamento"
        ]
    ):
        indicador = "faturamento"

    ordem = (
        "asc"
        if any(
            termo in t
            for termo in [
                "menos vendeu",
                "vendeu menos",
                "menor faturamento"
            ]
        )
        else "desc"
    )

    return normalizar_interpretacao(
        InterpretacaoPergunta(
            operacao="ranking",
            indicador=indicador,
            filtros=filtros,
            agrupar_por=dimensao,
            top_n=1,
            ordem=ordem
        )
    )


def interpretar_pergunta(pergunta):

    # ========================================================
    # QUALQUER INDICADOR POR QUALQUER DIMENSÃO
    # ========================================================
    interpretacao_generica = (
        _interpretar_indicador_por_dimensao_generico(
            pergunta
        )
    )

    if interpretacao_generica is not None:
        return (
            interpretacao_generica,
            "local_dimensao_generica"
        )

    # ========================================================
    # AJUSTES NOVOS - ROTA LOCAL SEGURA
    # ========================================================
    nova_interpretacao = _interpretar_ajustes_hoje(
        pergunta
    )

    if nova_interpretacao is not None:
        return nova_interpretacao, "local_ajustes_hoje"

    # ========================================================
    # IDENTIFICAR PERGUNTAS ANALÍTICAS MAIS COMPLEXAS
    # ========================================================

    complexa = pergunta_analitica_complexa(pergunta)

    # ========================================================
    # INTERPRETAÇÃO LOCAL SOMENTE PARA PERGUNTAS SIMPLES
    # ========================================================

    if not complexa:

        interpretacao_local = (
            interpretar_simples_localmente(
                pergunta
            )
        )

        if interpretacao_local is not None:

            print(
                "Interpretação local utilizada "
                "(Groq não foi chamada)."
            )

            return (
                interpretacao_local,
                "local"
            )

    else:

        print(
            "Pergunta analítica detectada. "
            "Interpretação local ignorada."
        )

    # --------------------------------------------------------
    # LOCAL SÓ PARA CONTINUAÇÕES
    # --------------------------------------------------------

    if (
        not complexa
        and parece_continuacao(pergunta)
    ):

        local = interpretar_localmente(
            pergunta
        )

        if local is not None:

            return local, "local"

    # --------------------------------------------------------
    # IA SEMÂNTICA (GROQ)
    # --------------------------------------------------------
    # Tudo que já foi reconhecido com segurança pelas rotas locais
    # retorna antes daqui. Quando surgir uma forma nova de perguntar,
    # a IA interpreta o significado e normaliza para os conceitos
    # canônicos já suportados pelo agente.

    try:

        interpretacao = (
            interpretar_com_groq(
                pergunta
            )
        )

        interpretacao = (
            corrigir_agrupamento_explicito(
                pergunta,
                interpretacao
            )
        )

        interpretacao = (
            corrigir_periodo_explicito(
                pergunta,
                interpretacao
            )
        )

        print("IA utilizada: GROQ")
        return (
            interpretacao,
            "groq"
        )

    except Exception as erro:

        print(
            "Groq falhou:",
            type(erro).__name__,
            "-",
            str(erro)[:300]
        )

    # --------------------------------------------------------
    # GEMINI
    # --------------------------------------------------------

    try:

        interpretacao = (
            interpretar_com_gemini(
                pergunta
            )
        )

        interpretacao = (
            corrigir_agrupamento_explicito(
                pergunta,
                interpretacao
            )
        )

        interpretacao = (
            corrigir_periodo_explicito(
                pergunta,
                interpretacao
            )
        )

        print("IA utilizada: GEMINI (fallback)")
        return (
            interpretacao,
            "gemini"
        )

    except Exception as erro:

        print(
            "Gemini falhou:",
            type(erro).__name__,
            "-",
            str(erro)[:300]
        )

    # --------------------------------------------------------
    # CLAUDE - SEGUNDO FALLBACK
    # --------------------------------------------------------

    try:
        interpretacao = interpretar_com_claude(pergunta)

        interpretacao = corrigir_agrupamento_explicito(
            pergunta,
            interpretacao
        )

        interpretacao = corrigir_periodo_explicito(
            pergunta,
            interpretacao
        )

        print("IA utilizada: CLAUDE (fallback)")
        return interpretacao, "claude"

    except Exception as erro:
        print(
            "Claude falhou:",
            type(erro).__name__,
            "-",
            str(erro)[:300]
        )

    raise RuntimeError(
        "Nenhum interpretador disponível."
    )


# ============================================================
# 23. CONTEXTO FINAL
# ============================================================

def montar_contexto_final(
    contexto_padrao,
    filtros_usuario=None
):

    contexto = deepcopy(
        contexto_padrao
    )


    if filtros_usuario:

        # ----------------------------------------------------
        # PERÍODO EM NÍVEL DE DIA
        # ----------------------------------------------------
        # Quando há uma faixa de datas explícita, ela substitui os
        # defaults de ano/mês. Isso é essencial para "últimos X dias"
        # e semanas que podem atravessar a virada do mês/ano.
        if (
            filtros_usuario.get("_data_inicio")
            or filtros_usuario.get("_data_fim")
        ):
            contexto.pop("ano", None)
            contexto.pop("mes", None)

        # ----------------------------------------------------
        # SE O USUÁRIO INFORMOU ANO, MAS NÃO INFORMOU MÊS,
        # CONSULTA O ANO INTEIRO
        # ----------------------------------------------------

        if (
            "ano" in filtros_usuario
            and filtros_usuario.get("mes") is None
        ):

            contexto.pop(
                "mes",
                None
            )


        # ----------------------------------------------------
        # APLICA OS FILTROS INFORMADOS PELO USUÁRIO
        # ----------------------------------------------------

        for chave, valor in (
            filtros_usuario.items()
        ):

            if valor is not None:

                contexto[chave] = valor


    return contexto


# ============================================================
# 24. FILTRO REPRESENTANTE VIA TREATAS
# ============================================================




# ============================================================
# 25. GERAR FILTROS DAX
# ============================================================

def gerar_filtros_dax(
    filtros
):

    filtros_dax = []


    for nome_filtro, valor in (
        filtros.items()
    ):

        if valor is None:

            continue

        # ----------------------------------------------------
        # FILTRO INTERNO DE PERÍODO DIÁRIO
        # ----------------------------------------------------
        if nome_filtro == "_periodo_rotulo":
            continue

        if nome_filtro in {"_data_inicio", "_data_fim"}:
            try:
                ano_data, mes_data, dia_data = [
                    int(parte)
                    for parte in str(valor).split("-")
                ]
            except Exception as erro:
                raise ValueError(
                    f"Data inválida no filtro {nome_filtro}: {valor}"
                ) from erro

            operador = ">=" if nome_filtro == "_data_inicio" else "<="

            filtros_dax.append(
                "'# CALENDÁRIO'[data] "
                f"{operador} DATE({ano_data}, {mes_data}, {dia_data})"
            )
            continue


        # ----------------------------------------------------
        # REPRESENTANTE
        # ----------------------------------------------------

        if nome_filtro not in mapa_campos:

            raise ValueError(
                f"Filtro não mapeado: "
                f"{nome_filtro}"
            )


        campo = mapa_campos[
            nome_filtro
        ]

        tabela = campo["tabela"]

        coluna = campo["coluna"]


        # ----------------------------------------------------
        # ANO
        # ----------------------------------------------------

        if nome_filtro == "ano":

            if valor == "Ano atual":

                filtros_dax.append(
                    "'# CALENDÁRIO'"
                    "[ano_atual] = "
                    "\"Ano atual\""
                )

            else:

                filtros_dax.append(
                    f"'# CALENDÁRIO'"
                    f"[ano] = \"{valor}\""
                )

            continue


        # ----------------------------------------------------
        # PRODUTO - CÓDIGO OU NOME PARCIAL
        # ----------------------------------------------------
        if nome_filtro == "produto":

            if isinstance(valor, list):
                partes = []

                for item in valor:
                    item_seguro = str(item).replace('"', '""')

                    partes.append(
                        "CONTAINSSTRING("
                        "'DW D_PRODUTO'[desc_produto], "
                        f'"{item_seguro}"'
                        ")"
                    )

                filtro = (
                    "FILTER("
                    "ALL('DW D_PRODUTO'[desc_produto]), "
                    + " || ".join(partes)
                    + ")"
                )

            else:
                valor_seguro = str(valor).replace('"', '""')

                filtro = (
                    "FILTER("
                    "ALL('DW D_PRODUTO'[desc_produto]), "
                    "CONTAINSSTRING("
                    "'DW D_PRODUTO'[desc_produto], "
                    f'"{valor_seguro}"'
                    ")"
                    ")"
                )

            filtros_dax.append(filtro)
            continue


        # ----------------------------------------------------
        # LISTA
        # ----------------------------------------------------

        if isinstance(
            valor,
            list
        ):

            valores = ", ".join(
                f'"{item}"'
                for item in valor
            )


            filtro = (
                f"'{tabela}'[{coluna}] "
                f"IN {{ {valores} }}"
            )


        # ----------------------------------------------------
        # VALOR
        # ----------------------------------------------------

        else:

            filtro = (
                f"'{tabela}'[{coluna}] "
                f'= "{valor}"'
            )


        filtros_dax.append(
            filtro
        )


    return filtros_dax


# ============================================================
# 26. DAX - VALOR SIMPLES
# ============================================================

def montar_dax_valor(
    indicador,
    filtros=None
):

    medida = (
        mapa_indicadores[
            indicador
        ]["medida"]
    )


    contexto = montar_contexto_final(

        contexto_overview_comercial,

        filtros
    )


    filtros_dax = gerar_filtros_dax(
        contexto
    )


    filtros_texto = ",\n        ".join(
        filtros_dax
    )


    return f"""
EVALUATE
ROW(
    "Resultado",
    CALCULATE(
        [{medida}],
        {filtros_texto}
    )
)
"""


# ============================================================
# DAX - MÚLTIPLOS INDICADORES
# ============================================================

def montar_dax_multiplos(
    indicadores,
    filtros=None
):

    contexto = montar_contexto_final(
        contexto_overview_comercial,
        filtros
    )

    filtros_dax = gerar_filtros_dax(
        contexto
    )

    filtros_texto = ",\n            ".join(
        filtros_dax
    )

    colunas = []

    for nome_saida, indicador in indicadores.items():

        medida = (
            mapa_indicadores[
                indicador
            ]["medida"]
        )

        coluna = f'''
        "{nome_saida}",
        CALCULATE(
            [{medida}],
            {filtros_texto}
        )'''

        colunas.append(
            coluna
        )

    colunas_texto = ",\n".join(
        colunas
    )

    return f"""
EVALUATE
ROW(
{colunas_texto}
)
"""


# ============================================================
# 27. DAX - RANKING NORMAL
# ============================================================

def montar_dax_ranking_normal(
    indicador,
    agrupar_por,
    top_n,
    ordem,
    filtros
):

    medida = (
        mapa_indicadores[
            indicador
        ]["medida"]
    )


    dimensao = (
        mapa_dimensoes[
            agrupar_por
        ]
    )


    contexto = montar_contexto_final(

        contexto_overview_comercial,

        filtros
    )


    # --------------------------------------------------------
    # RETIRA FILTRO DA PRÓPRIA DIMENSÃO
    # --------------------------------------------------------

    contexto.pop(
        agrupar_por,
        None
    )


    filtros_dax = gerar_filtros_dax(
        contexto
    )


    filtros_texto = ",\n            ".join(
        filtros_dax
    )


    ordem_dax = (
        "DESC"
        if ordem == "desc"
        else "ASC"
    )


    return f"""
EVALUATE

TOPN(
    {top_n},

    CALCULATETABLE(
        SUMMARIZECOLUMNS(
            '{dimensao["tabela"]}'[{dimensao["coluna"]}],

            "Resultado",
            [{medida}]
        ),

        {filtros_texto}
    ),

    [Resultado],
    {ordem_dax}
)

ORDER BY
    [Resultado] {ordem_dax}
"""


# ============================================================
# 28. DAX - RANKING REPRESENTANTE
# ============================================================




# ============================================================
# 29. DAX - RANKING GENÉRICO
# ============================================================

def montar_dax_ranking(
    indicador,
    agrupar_por,
    top_n=5,
    ordem="desc",
    filtros=None
):

    return montar_dax_ranking_normal(
        indicador,
        agrupar_por,
        top_n,
        ordem,
        filtros
    )

# ============================================================
# 30. EXECUTAR DAX
# ============================================================

def executar_dax(
    dax
):

    _t_dax_inicio = time.perf_counter()

    body = {

        "queries": [
            {
                "query": dax
            }
        ],

        "serializerSettings": {
            "includeNulls": True
        }
    }


    _t_headers_inicio = time.perf_counter()

    headers = criar_headers()

    _t_headers_fim = time.perf_counter()

    _t_http_inicio = time.perf_counter()

    resposta = requests.post(

        URL,

        headers=headers,

        json=body,

        timeout=60
    )

    _t_http_fim = time.perf_counter()

    print(
        f"[TEMPO PBI] Token/headers: "
        f"{_t_headers_fim - _t_headers_inicio:.3f}s | "
        f"HTTP Power BI: "
        f"{_t_http_fim - _t_http_inicio:.3f}s | "
        f"Total executar_dax: "
        f"{_t_http_fim - _t_dax_inicio:.3f}s"
    )


    if resposta.status_code != 200:

        raise RuntimeError(
            f"Erro Power BI "
            f"{resposta.status_code}: "
            f"{resposta.text}"
        )


    return resposta.json()


# ============================================================
# 31. EXTRAIR LINHAS
# ============================================================

def extrair_linhas(
    resultado
):

    try:

        return (
            resultado["results"][0]
            ["tables"][0]
            ["rows"]
        )

    except Exception:

        return []


# ============================================================
# 32. CONSULTAR VALOR
# ============================================================

def consultar_valor(
    indicador,
    filtros
):

    dax = montar_dax_valor(
        indicador,
        filtros
    )


    linhas = extrair_linhas(
        executar_dax(
            dax
        )
    )


    if not linhas:

        return None


    return linhas[0].get(
        "[Resultado]"
    )



# ============================================================
# CONSULTAR MÚLTIPLOS INDICADORES
# ============================================================

def consultar_multiplos(
    indicadores,
    filtros
):

    dax = montar_dax_multiplos(
        indicadores,
        filtros
    )

    linhas = extrair_linhas(
        executar_dax(
            dax
        )
    )

    if not linhas:

        return {
            nome: None
            for nome in indicadores
        }

    linha = linhas[0]

    resultado = {}

    for nome in indicadores:

        resultado[nome] = (
            linha.get(
                f"[{nome}]"
            )
        )

    return resultado

# ============================================================
# 33. CONSULTAR RANKING
# ============================================================

def consultar_ranking(
    indicador,
    agrupar_por,
    top_n,
    ordem,
    filtros
):

    dax = montar_dax_ranking(

        indicador,
        agrupar_por,
        top_n,
        ordem,
        filtros
    )


    linhas = extrair_linhas(
        executar_dax(
            dax
        )
    )


    dimensao = (
        mapa_dimensoes[
            agrupar_por
        ]
    )


    chave = (
        f"{dimensao['tabela']}"
        f"[{dimensao['coluna']}]"
    )


    ranking = []


    for linha in linhas:

        item = linha.get(
            chave
        )

        valor = linha.get(
            "[Resultado]"
        )


        if item is None:

            continue


        ranking.append(
            {
                "item": item,
                "valor": valor
            }
        )


    return ranking


# ============================================================
# 34. FORMATAR VALOR
# ============================================================

def formatar_valor(
    valor,
    formato
):

    if valor is None:

        return "Sem resultado"


    if formato == "moeda":

        return (
            f"R$ {valor:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )


    if formato == "percentual":

        return (
            f"{valor * 100:.1f}%"
            .replace(".", ",")
        )


    if formato == "inteiro":

        return (
            f"{int(round(valor)):,}"
            .replace(",", ".")
        )


    return str(
        valor
    )


def _formatar_valor_periodo(valor, formato, filtros=None):
    """
    Em consultas com período diário explícito, ausência de linhas/valor
    representa zero para apresentação ao usuário. Fora desse caminho,
    preserva exatamente o comportamento anterior de formatar_valor().
    """
    if (filtros or {}).get("_periodo_rotulo") and valor is None:
        valor = 0

    return formatar_valor(valor, formato)


# ============================================================
# 35. NOME DO MÊS PARA RESPOSTA
# ============================================================

def nome_mes_resposta(
    mes
):

    if mes == "Mês atual":

        nomes = {

            1: "janeiro",
            2: "fevereiro",
            3: "março",
            4: "abril",
            5: "maio",
            6: "junho",
            7: "julho",
            8: "agosto",
            9: "setembro",
            10: "outubro",
            11: "novembro",
            12: "dezembro"
        }

        return nomes[
            datetime.now().month
        ]


    return meses_nome.get(
        mes,
        mes
    )


# ============================================================
# 36. CONTEXTO DA RESPOSTA
# ============================================================

def construir_contexto_resposta(
    filtros
):

    partes = []


    # --------------------------------------------------------
    # REGIÃO
    # --------------------------------------------------------

    if filtros.get("regiao"):

        partes.append(
            f"da região "
            f"{filtros['regiao']}"
        )


    # --------------------------------------------------------
    # REPRESENTANTE
    # --------------------------------------------------------

    if filtros.get("representante"):

        partes.append(
            f"do representante "
            f"{filtros['representante']}"
        )


    # --------------------------------------------------------
    # PLATAFORMA
    # --------------------------------------------------------

    if filtros.get("plataforma"):

        partes.append(
            f"da plataforma "
            f"{filtros['plataforma']}"
        )


    # --------------------------------------------------------
    # CLASSE
    # --------------------------------------------------------

    if filtros.get("classe"):

        partes.append(
            f"da classe "
            f"{filtros['classe']}"
        )


    # --------------------------------------------------------
    # LINHA
    # --------------------------------------------------------

    if filtros.get("linha"):

        partes.append(
            f"da linha "
            f"{filtros['linha']}"
        )


    # --------------------------------------------------------
    # FAMÍLIA
    # --------------------------------------------------------

    if filtros.get("familia"):

        partes.append(
            f"da família "
            f"{filtros['familia']}"
        )


    # --------------------------------------------------------
    # CLIENTE
    # --------------------------------------------------------

    if filtros.get("cliente"):

        partes.append(
            f"do cliente "
            f"{filtros['cliente']}"
        )


    # --------------------------------------------------------
    # PRODUTO
    # --------------------------------------------------------

    if filtros.get("produto"):

        partes.append(
            f"do produto "
            f"{filtros['produto']}"
        )


    # --------------------------------------------------------
    # PERÍODO EM NÍVEL DE DIA
    # --------------------------------------------------------

    if filtros.get("_periodo_rotulo"):

        partes.append(
            filtros["_periodo_rotulo"]
        )


    # --------------------------------------------------------
    # MÊS
    # --------------------------------------------------------

    if filtros.get("mes"):
    
        mes_filtro = filtros["mes"]
    
        # --------------------------------------------------------
        # TRIMESTRE
        # --------------------------------------------------------
    
        if isinstance(
            mes_filtro,
            list
        ):
    
            mapa_trimestres = {
                ("01", "02", "03"): "1º trimestre",
                ("04", "05", "06"): "2º trimestre",
                ("07", "08", "09"): "3º trimestre",
                ("10", "11", "12"): "4º trimestre"
            }
    
            chave_trimestre = tuple(
                mes_filtro
            )
    
            descricao_periodo = (
                mapa_trimestres.get(
                    chave_trimestre
                )
            )
    
            if descricao_periodo:
    
                partes.append(
                    f"no {descricao_periodo}"
                )
    
            else:
    
                nomes_meses = [
                    nome_mes_resposta(mes)
                    for mes in mes_filtro
                ]
    
                partes.append(
                    "nos meses "
                    + ", ".join(nomes_meses)
                )
    
        # --------------------------------------------------------
        # MÊS ÚNICO
        # --------------------------------------------------------
    
        else:
    
            partes.append(
                f"em "
                f"{nome_mes_resposta(mes_filtro)}"
            )


    # --------------------------------------------------------
    # ANO
    # --------------------------------------------------------

    if filtros.get("ano"):

        ano = filtros["ano"]

        if ano != "Ano atual":

            partes.append(
                f"de {ano}"
            )


    return " ".join(
        partes
    )


# ============================================================
# 37. RESPOSTA - VALOR
# ============================================================

def construir_resposta_valor(
    indicador,
    filtros,
    valor
):

    config = (
        mapa_indicadores[
            indicador
        ]
    )


    valor_formatado = _formatar_valor_periodo(

        valor,

        config["formato"],

        filtros
    )


    contexto = construir_contexto_resposta(
        filtros
    )


    if contexto:

        return (
            f"{config['descricao']} "
            f"{contexto}: "
            f"{valor_formatado}."
        )


    return (
        f"{config['descricao']}: "
        f"{valor_formatado}."
    )


# ============================================================
# 38. RESPOSTA - RANKING
# ============================================================

def construir_resposta_ranking(
    indicador,
    agrupar_por,
    ranking,
    filtros,
    top_n=None,
    ordem="desc",
    pergunta=None
):

    if not ranking:

        return (
            "Não encontrei resultados "
            "para essa consulta."
        )


    config = (
        mapa_indicadores[
            indicador
        ]
    )


    descricao = (
        mapa_dimensoes[
            agrupar_por
        ]["descricao"]
    )


    contexto = construir_contexto_resposta(
        filtros
    )


    if top_n == 1:

        item = ranking[0]

        valor = formatar_valor(
            item["valor"],
            config["formato"]
        )

        t = _texto_normalizado(
            pergunta or ""
        )

        if agrupar_por == "plataforma":
            if "varejista" in t:
                sujeito = "varejista"
            elif "grupo" in t:
                sujeito = "grupo"
            elif "cliente" in t:
                sujeito = "cliente"
            else:
                sujeito = "plataforma"

        elif agrupar_por == "loja":
            sujeito = "CNPJ" if "cnpj" in t else "loja"

        else:
            sujeito = descricao

        if filtros.get("mes"):
            periodo = (
                f" em {nome_mes_resposta(filtros['mes'])}"
                if not isinstance(filtros["mes"], list)
                else ""
            )
        elif filtros.get("ano") and filtros["ano"] != "Ano atual":
            periodo = f" em {filtros['ano']}"
        else:
            periodo = " neste mês"

        artigo = "A" if sujeito in [
            "varejista",
            "plataforma",
            "loja",
            "região",
            "família",
            "linha"
        ] else "O"

        if indicador == "quantidade" and any(x in t for x in [
            "comprou",
            "compraram",
            "vendeu",
            "venderam",
            "produtos"
        ]):

            if ordem == "asc":
                return (
                    f"📉 {artigo} {sujeito} que menos comprou "
                    f"seus produtos{periodo} foi "
                    f"{item['item']}, com {valor} unidades vendidas."
                )

            return (
                f"🏆 {artigo} {sujeito} que mais comprou "
                f"seus produtos{periodo} foi "
                f"{item['item']}, com {valor} unidades vendidas."
            )

        if indicador == "faturamento":
            return (
                f"🏆 {artigo} {sujeito} com maior faturamento"
                f"{periodo} foi {item['item']}, com {valor}."
            )

        return (
            f"🏆 {artigo} {sujeito} com maior "
            f"{config['descricao'].lower()}{periodo} foi "
            f"{item['item']}, com {valor}."
        )

    titulo = (
        f"Ranking por {descricao}"
    )


    if contexto:

        titulo += (
            f" {contexto}"
        )


    linhas = []


    for posicao, item in enumerate(
        ranking,
        start=1
    ):

        valor = formatar_valor(

            item["valor"],

            config["formato"]
        )


        linhas.append(
            f"{posicao}. "
            f"{item['item']} — "
            f"{valor}"
        )


    return (
        titulo
        + ":\n\n"
        + "\n".join(
            linhas
        )
    )


# ============================================================
# 39. AGENTE PRINCIPAL
# ============================================================

# ============================================================
# 39A. MÚLTIPLOS INDICADORES + META CONTEXTUAL
# ============================================================

def _encontrar_ocorrencias(texto, termos):
    ocorrencias = []

    for termo in termos:
        inicio = 0

        while True:
            pos = texto.find(
                termo,
                inicio
            )

            if pos < 0:
                break

            ocorrencias.append(
                (
                    pos,
                    termo
                )
            )

            inicio = (
                pos + len(termo)
            )

    return sorted(
        ocorrencias,
        key=lambda x: x[0]
    )


def _familias_mencionadas_com_posicao(pergunta):
    """
    Localiza as famílias principais citadas pelo usuário e guarda
    a posição no texto. Isso permite entender a qual indicador
    uma palavra genérica como 'meta' pertence.
    """
    t = _texto_normalizado(
        pergunta
    )

    regras = {
        "faturamento": [
            "faturamento",
            "faturou",
            "faturamos",
            "realizado"
        ],

        "quantidade": [
            "quantidade vendida",
            "quantidade de vendas",
            "qtd vendida",
            "qtd"
        ],

        "margem_liquida": [
            "margem liquida",
            "mrgl"
        ]
    }

    encontrados = []

    for familia, termos in regras.items():
        ocorrencias = _encontrar_ocorrencias(
            t,
            termos
        )

        if ocorrencias:
            encontrados.append(
                {
                    "familia": familia,
                    "pos": ocorrencias[0][0]
                }
            )

    return sorted(
        encontrados,
        key=lambda x: x["pos"]
    )


def _familia_mais_proxima(
    posicao,
    familias
):
    """
    Prioriza a família citada imediatamente antes da palavra genérica.
    Se não houver, usa a primeira família citada depois.
    """
    anteriores = [
        item
        for item in familias
        if item["pos"] <= posicao
    ]

    if anteriores:
        return max(
            anteriores,
            key=lambda x: x["pos"]
        )["familia"]

    posteriores = [
        item
        for item in familias
        if item["pos"] > posicao
    ]

    if posteriores:
        return min(
            posteriores,
            key=lambda x: x["pos"]
        )["familia"]

    return None


def _resolver_meta_contextual(pergunta):
    """
    Resolve expressões genéricas de meta usando o contexto.

    Exemplos:
    - faturamento, meta e % da meta
      -> faturamento + meta_faturamento + atingimento_meta_faturamento

    - quantidade vendida, meta e % da meta
      -> quantidade + meta_quantidade + atingimento_meta_quantidade

    - margem líquida, meta e desvio
      -> margem_liquida + meta_margem_liquida
         + desvio_meta_margem_liquida
    """
    t_original = _texto_normalizado(
        pergunta
    )

    t = t_original

    resultado = []

    def adicionar(indicador):
        if indicador not in resultado:
            resultado.append(
                indicador
            )

    # --------------------------------------------------------
    # INDICADORES COMPOSTOS / EXPLÍCITOS
    # --------------------------------------------------------

    if any(x in t for x in [
        "previsao de faturamento",
        "projecao de faturamento",
        "forecast"
    ]):
        adicionar("previsao_faturamento")

        for trecho in [
            "previsao de faturamento",
            "projecao de faturamento",
            "forecast"
        ]:
            t = t.replace(
                trecho,
                " " * len(trecho)
            )

    compostos = [
        (
            [
                "faturamento + entregas",
                "fat + entregas",
                "faturamento e entregas"
            ],
            "faturamento_entregas"
        ),

        (
            [
                "meta de margem liquida",
                "percentual da meta de margem liquida"
            ],
            "meta_margem_liquida"
        ),

        (
            [
                "carteira faturavel do dia",
                "carteira faturavel dia"
            ],
            "carteira_faturavel_dia"
        ),

        (
            [
                "carteira faturavel do mes",
                "carteira faturavel mes"
            ],
            "carteira_faturavel_mes"
        ),

        (
            [
                "faturamento da carteira",
                "faturamento carteira"
            ],
            "faturamento_carteira"
        ),

        (
            [
                "meta de quantidade",
                "meta em quantidade",
                "meta de vendas"
            ],
            "meta_quantidade"
        ),

        (
            [
                "meta de faturamento"
            ],
            "meta_faturamento"
        )
    ]

    for termos, indicador in compostos:
        for termo in termos:
            if termo in t:
                adicionar(
                    indicador
                )

                t = t.replace(
                    termo,
                    " " * len(termo)
                )

    # --------------------------------------------------------
    # FAMÍLIAS PRINCIPAIS
    # --------------------------------------------------------

    # Para identificação de famílias, mascara somente as expressões
    # compostas de previsão. Assim "previsão de faturamento" permanece
    # UM indicador, sem adicionar "faturamento" novamente.
    texto_familias = t_original

    for trecho in [
        "previsao de faturamento",
        "projecao de faturamento",
        "forecast"
    ]:
        texto_familias = texto_familias.replace(
            trecho,
            " " * len(trecho)
        )

    familias = (
        _familias_mencionadas_com_posicao(
            texto_familias
        )
    )

    for item in familias:
        familia = item[
            "familia"
        ]

        adicionar(
            FAMILIAS_META[
                familia
            ]["principal"]
        )

    # --------------------------------------------------------
    # INDICADORES EXPLÍCITOS QUE NÃO DEPENDEM DE FAMÍLIA
    # --------------------------------------------------------

    if re.search(
        r"\bmrgb\b",
        t
    ):
        adicionar(
            "valor_margem_bruta"
        )

    elif "margem bruta" in t:

        if any(
            x in t
            for x in [
                "r$",
                "reais",
                "valor"
            ]
        ):
            adicionar(
                "valor_margem_bruta"
            )

        else:
            adicionar(
                "margem_bruta"
            )

    if re.search(
        r"\bentregas\b",
        t
    ):
        adicionar(
            "entregas"
        )

    # --------------------------------------------------------
    # MARCADORES DE META / ATINGIMENTO / DESVIO
    # --------------------------------------------------------

    # Primeiro localizamos expressões mais longas para que a palavra
    # "meta" dentro delas não seja tratada duas vezes.
    marcadores = []

    padroes_contextuais = [
        (
            r"(percentual de atingimento|atingimento da meta|% da meta|percentual da meta)",
            "atingimento"
        ),

        (
            r"(quanto falta para a meta|desvio da meta|diferenca para a meta)",
            "desvio"
        )
    ]

    mascara = list(
        t
    )

    for padrao, tipo in padroes_contextuais:

        for match in re.finditer(
            padrao,
            t
        ):
            marcadores.append(
                {
                    "pos": match.start(),
                    "tipo": tipo
                }
            )

            for i in range(
                match.start(),
                match.end()
            ):
                mascara[i] = " "

    texto_sem_expressoes = "".join(
        mascara
    )

    # Agora "meta" isolada.
    for match in re.finditer(
        r"\bmeta\b",
        texto_sem_expressoes
    ):
        marcadores.append(
            {
                "pos": match.start(),
                "tipo": "meta"
            }
        )

    marcadores = sorted(
        marcadores,
        key=lambda x: x["pos"]
    )

    for marcador in marcadores:

        familia = _familia_mais_proxima(
            marcador["pos"],
            familias
        )

        if familia is None:
            # Sem contexto suficiente: não inventa uma família.
            continue

        indicador_contextual = (
            FAMILIAS_META[
                familia
            ].get(
                marcador["tipo"]
            )
        )

        if indicador_contextual:
            adicionar(
                indicador_contextual
            )

    return resultado


def _indicadores_explicitos_multiplos(pergunta):
    """
    Compatibilidade com o restante do agente.
    Agora usa o resolvedor contextual.
    """
    return _resolver_meta_contextual(
        pergunta
    )


def _pergunta_multiplos_precisa_ia_para_filtros(pergunta):
    """
    Para perguntas simples com apenas período/região, evita IA.
    Para filtros de entidade mais complexos, mantém o interpretador
    atual e seus fallbacks.
    """
    t = _texto_normalizado(
        pergunta
    )

    termos_complexos = [
        "cliente",
        "produto",
        "linha",
        "familia",
        "representante",
        "plataforma",
        "classe",
        "status de entrega",
        "status_entrega",
        "curva abc",
        "cod_curva_abc",
        "analise de credito",
        "analise_credito"
    ]

    return any(
        termo in t
        for termo in termos_complexos
    )


def _filtros_para_multiplos_indicadores(pergunta):
    """
    Caminho rápido:
    - período + região são resolvidos localmente;
    - IA só entra quando há filtros de entidade mais complexos.
    """
    if not _pergunta_multiplos_precisa_ia_para_filtros(
        pergunta
    ):
        filtros = _extrair_periodo_pergunta(
            pergunta
        )

        texto_upper = pergunta.upper()

        for regiao in regioes_conhecidas:

            if regiao in texto_upper:

                filtros[
                    "regiao"
                ] = regiao

                break

        filtros = _aplicar_periodo_calendario(
            pergunta,
            filtros
        )

        return (
            filtros,
            "local"
        )

    interpretacao, origem = interpretar_pergunta(
        pergunta
    )

    filtros = (
        interpretacao
        .filtros
        .model_dump(
            exclude_none=True
        )
    )

    filtros = _aplicar_periodo_calendario(
        pergunta,
        filtros,
        interpretacao
    )

    return (
        filtros,
        origem
    )


def _bases_indicador_derivado(indicador):
    config = mapa_indicadores_derivados.get(
        indicador
    )

    if not config:
        return []

    return list(
        config.get(
            "bases",
            []
        )
    )


def _calcular_indicador_derivado(
    indicador,
    dados
):
    """
    Calcula métricas derivadas SEM consultar novamente o Power BI.
    """
    if indicador == "atingimento_meta_quantidade":

        realizado = dados.get(
            "quantidade"
        )

        meta = dados.get(
            "meta_quantidade"
        )

        if (
            realizado is None
            or meta in (
                None,
                0
            )
        ):
            return None

        return realizado / meta

    if indicador == "desvio_meta_quantidade":

        realizado = dados.get(
            "quantidade"
        )

        meta = dados.get(
            "meta_quantidade"
        )

        if (
            realizado is None
            or meta is None
        ):
            return None

        return realizado - meta

    if indicador == "atingimento_meta_margem_liquida":

        realizado = dados.get(
            "margem_liquida"
        )

        meta = dados.get(
            "meta_margem_liquida"
        )

        if (
            realizado is None
            or meta in (
                None,
                0
            )
        ):
            return None

        return realizado / meta

    if indicador == "desvio_meta_margem_liquida":

        realizado = dados.get(
            "margem_liquida"
        )

        meta = dados.get(
            "meta_margem_liquida"
        )

        if (
            realizado is None
            or meta is None
        ):
            return None

        # Diferença entre os percentuais.
        return realizado - meta

    return None


def _consultar_multiplos_indicadores_pergunta(
    pergunta
):
    indicadores = _indicadores_explicitos_multiplos(
        pergunta
    )

    if len(indicadores) < 2:
        return None

    filtros, origem = (
        _filtros_para_multiplos_indicadores(
            pergunta
        )
    )

    # --------------------------------------------------------
    # MEDIDAS REAIS NECESSÁRIAS
    # --------------------------------------------------------

    indicadores_fisicos = []

    def adicionar_fisico(indicador):

        if (
            indicador in mapa_indicadores
            and indicador
            not in indicadores_fisicos
        ):
            indicadores_fisicos.append(
                indicador
            )

    for indicador in indicadores:

        if indicador in mapa_indicadores:
            adicionar_fisico(
                indicador
            )

        else:
            for base in _bases_indicador_derivado(
                indicador
            ):
                adicionar_fisico(
                    base
                )

    # UMA única consulta DAX para todas as medidas físicas.
    mapa_saida = {
        indicador: indicador
        for indicador in indicadores_fisicos
    }

    dados = consultar_multiplos(
        mapa_saida,
        filtros
    )

    # --------------------------------------------------------
    # MÉTRICAS DERIVADAS LOCAIS
    # --------------------------------------------------------

    for indicador in indicadores:

        if indicador in mapa_indicadores_derivados:

            dados[
                indicador
            ] = _calcular_indicador_derivado(
                indicador,
                dados
            )

    return {
        "indicadores": indicadores,
        "filtros": filtros,
        "dados": dados,
        "interpretado_por": origem
    }


def _construir_resposta_multiplos_indicadores(
    resultado
):
    filtros = resultado["filtros"]
    dados = resultado["dados"]
    indicadores = resultado["indicadores"]

    contexto = construir_contexto_resposta(
        filtros
    )

    linhas = []

    for indicador in indicadores:

        config = _config_indicador_qualquer(
            indicador
        )

        valor = _formatar_valor_periodo(
            dados.get(
                indicador
            ),
            config[
                "formato"
            ],
            filtros
        )

        if contexto:

            linhas.append(
                f"{config['descricao']} "
                f"{contexto}: {valor}."
            )

        else:

            linhas.append(
                f"{config['descricao']}: "
                f"{valor}."
            )

    return "\n".join(
        linhas
    )


def _imagem_tabela_multiplos_indicadores(
    resultado
):
    filtros = resultado["filtros"]
    dados = resultado["dados"]
    indicadores = resultado["indicadores"]

    contexto = construir_contexto_resposta(
        filtros
    )

    titulo = "Indicadores comerciais"

    if contexto:
        titulo += f" {contexto}"

    linhas = []

    for indicador in indicadores:

        config = _config_indicador_qualquer(
            indicador
        )

        linhas.append([
            config[
                "descricao"
            ],
            _formatar_valor_periodo(
                dados.get(
                    indicador
                ),
                config[
                    "formato"
                ],
                filtros
            )
        ])

    return _tabela_png_base64(
        titulo,
        [
            "Indicador",
            "Resultado"
        ],
        linhas,
        fonte=9
    )


def _resposta_multiplos_indicadores(
    pergunta
):
    resultado = (
        _consultar_multiplos_indicadores_pergunta(
            pergunta
        )
    )

    if resultado is None:
        return None

    resposta = (
        _construir_resposta_multiplos_indicadores(
            resultado
        )
    )

    if _usuario_pediu_tabela(
        pergunta
    ):

        return {
            "tipo_resposta":
                "grafico",

            "tipo_grafico":
                "tabela",

            "resposta":
                resposta,

            "imagem_base64":
                _imagem_tabela_multiplos_indicadores(
                    resultado
                ),

            "nome_arquivo":
                "indicadores_comerciais.png"
        }

    return {
        "tipo_resposta":
            "texto",

        "resposta":
            resposta,

        "dados":
            resultado[
                "dados"
            ],

        "filtros":
            resultado[
                "filtros"
            ],

        "indicadores":
            resultado[
                "indicadores"
            ],

        "interpretado_por":
            resultado[
                "interpretado_por"
            ]
    }



def consultar_resumo_meta(
    filtros
):

    indicadores = {
        "meta": "meta_faturamento",
        "realizado": "faturamento",
        "atingimento": "atingimento_meta_faturamento",
        "desvio": "desvio_meta_faturamento"
    }

    return consultar_multiplos(
        indicadores,
        filtros
    )

def construir_resposta_resumo_meta(
    filtros,
    dados
):

    contexto = construir_contexto_resposta(
        filtros
    )

    meta = formatar_valor(
        dados["meta"],
        "moeda"
    )

    realizado = formatar_valor(
        dados["realizado"],
        "moeda"
    )

    atingimento = formatar_valor(
        dados["atingimento"],
        "percentual"
    )

    desvio = formatar_valor(
        dados["desvio"],
        "moeda"
    )

    titulo = "Resumo de faturamento"

    if contexto:
        titulo += f" {contexto}"

    return (
        f"{titulo}:\n\n"
        f"Meta: {meta}\n"
        f"Realizado: {realizado}\n"
        f"Atingimento: {atingimento}\n"
        f"Desvio para a meta: {desvio}"
    )

def consultar_resumo_gerencial(
    filtros
):

    # IMPORTANTE PARA PERFORMANCE:
    # todos os indicadores abaixo são buscados em UMA ÚNICA consulta DAX.
    indicadores = {
        "faturamento": "faturamento",
        "meta_faturamento": "meta_faturamento",
        "atingimento_faturamento": "atingimento_meta_faturamento",
        "margem_liquida": "margem_liquida",
        "meta_margem_liquida": "meta_margem_liquida",
        "margem_bruta": "margem_bruta",
        "quantidade": "quantidade",
        "meta_quantidade": "meta_quantidade",
        "entregas": "entregas",
        "faturamento_entregas": "faturamento_entregas",
        "carteira_faturavel_dia": "carteira_faturavel_dia",
        "carteira_faturavel_mes": "carteira_faturavel_mes",
        "faturamento_carteira": "faturamento_carteira"
    }

    try:

        return consultar_multiplos(
            indicadores,
            filtros
        )

    except Exception as erro:

        print(
            "Erro no resumo gerencial:",
            type(erro).__name__,
            "-",
            str(erro)[:300]
        )

        return {
            nome: None
            for nome in indicadores
        }
def construir_resposta_resumo_gerencial(
    filtros,
    dados
):

    contexto = construir_contexto_resposta(
        filtros
    )

    faturamento = formatar_valor(
        dados.get("faturamento"),
        "moeda"
    )

    meta = formatar_valor(
        dados.get("meta_faturamento"),
        "moeda"
    )

    atingimento = formatar_valor(
        dados.get("atingimento_faturamento"),
        "percentual"
    )

    margem_liquida = formatar_valor(
        dados.get("margem_liquida"),
        "percentual"
    )

    margem_bruta = formatar_valor(
        dados.get("margem_bruta"),
        "percentual"
    )

    titulo = "Resumo gerencial"

    if contexto:
        titulo += f" {contexto}"

    meta_margem_liquida = formatar_valor(
        dados.get("meta_margem_liquida"),
        "percentual"
    )

    quantidade = formatar_valor(
        dados.get("quantidade"),
        "inteiro"
    )

    meta_quantidade = formatar_valor(
        dados.get("meta_quantidade"),
        "inteiro"
    )

    entregas = formatar_valor(
        dados.get("entregas"),
        "moeda"
    )

    faturamento_entregas = formatar_valor(
        dados.get("faturamento_entregas"),
        "moeda"
    )

    carteira_dia = formatar_valor(
        dados.get("carteira_faturavel_dia"),
        "moeda"
    )

    carteira_mes = formatar_valor(
        dados.get("carteira_faturavel_mes"),
        "moeda"
    )

    faturamento_carteira = formatar_valor(
        dados.get("faturamento_carteira"),
        "moeda"
    )

    return (
        f"{titulo}:\n\n"
        f"Faturamento: {faturamento}\n"
        f"Meta de faturamento: {meta}\n"
        f"Atingimento da meta: {atingimento}\n"
        f"Margem líquida: {margem_liquida}\n"
        f"Meta de margem líquida: {meta_margem_liquida}\n"
        f"Margem bruta: {margem_bruta}\n"
        f"Quantidade vendida: {quantidade}\n"
        f"Meta de quantidade: {meta_quantidade}\n"
        f"Entregas: {entregas}\n"
        f"Faturamento + entregas: {faturamento_entregas}\n"
        f"Carteira faturável dia: {carteira_dia}\n"
        f"Carteira faturável mês: {carteira_mes}\n"
        f"Faturamento carteira: {faturamento_carteira}"
    )


# ============================================================
# DETECTAR SE A PERGUNTA PARECE COMPOSTA
# ============================================================

def parece_pergunta_composta(pergunta):

    texto = (
        pergunta
        .strip()
        .lower()
    )

    # --------------------------------------------------------
    # RESUMO DE META É UMA ÚNICA OPERAÇÃO
    # --------------------------------------------------------

    if (
        "meta" in texto
        and (
            "faturamento" in texto
            or "faturamos" in texto
        )
        and (
            "percentual" in texto
            or "atingimento" in texto
            or "%" in texto
        )
    ):
        return False

    # --------------------------------------------------------
    # RESUMO GERENCIAL É UMA ÚNICA OPERAÇÃO
    # --------------------------------------------------------

    if (
        "resumo gerencial" in texto
        or "visão geral" in texto
        or "visao geral" in texto
    ):
        return False

    # --------------------------------------------------------
    # DOIS OU MAIS MESES EXPLÍCITOS
    # --------------------------------------------------------

    meses_encontrados = [
        nome
        for nome in mapa_meses
        if re.search(r"\b" + re.escape(nome) + r"\b", texto)
    ]

    if len(meses_encontrados) >= 2:
        return True

    # --------------------------------------------------------
    # MÊS A MÊS
    # --------------------------------------------------------

    if (
        "mês a mês" in texto
        or "mes a mes" in texto
    ):
        return True

    # --------------------------------------------------------
    # COMPARAÇÃO
    # --------------------------------------------------------

    if (
        "compare " in texto
        or "comparar " in texto
        or "comparação " in texto
        or "comparacao " in texto
    ):
        return True

    # --------------------------------------------------------
    # DUAS SOLICITAÇÕES
    # --------------------------------------------------------

    marcadores = [
        " e também ",
        " e tambem ",
        " e como ",
        " e qual ",
        " e quanto ",
        " além disso",
        " alem disso"
    ]

    if any(
        marcador in texto
        for marcador in marcadores
    ):
        return True

    return False


# ============================================================
# IDENTIFICAR INDICADOR PARA DECOMPOSIÇÃO LOCAL
# ============================================================

def identificar_indicador_composto_local(pergunta):

    texto = (
        pergunta
        .strip()
        .lower()
    )

    if (
        "percentual sobre a meta" in texto
        or "% sobre a meta" in texto
        or "percentual da meta" in texto
        or "% da meta" in texto
        or "atingimento da meta" in texto
    ):
        return "percentual sobre a meta"

    if (
        "margem líquida" in texto
        or "margem liquida" in texto
    ):
        return "margem líquida"

    if "margem bruta" in texto:
        return "margem bruta"

    if (
        "meta de faturamento" in texto
        or "meta de vendas" in texto
    ):
        return "meta de faturamento"

    if (
        "faturamento" in texto
        or "faturou" in texto
        or "faturamos" in texto
    ):
        return "faturamento"

    return None


# ============================================================
# DECOMPOSIÇÃO LOCAL
# ============================================================

def decompor_composta_localmente(pergunta):

    texto = (
        pergunta
        .strip()
        .lower()
    )

    indicador = (
        identificar_indicador_composto_local(
            pergunta
        )
    )

    if indicador is None:
        return None

    # --------------------------------------------------------
    # ANO
    # --------------------------------------------------------

    match_ano = re.search(
        r"\b(20\d{2})\b",
        texto
    )

    ano = (
        match_ano.group(1)
        if match_ano
        else None
    )

    texto_ano = (
        f" de {ano}"
        if ano
        else ""
    )

    # --------------------------------------------------------
    # MESES EXPLÍCITOS
    # --------------------------------------------------------

    meses_encontrados = []

    for nome in mapa_meses:

        if re.search(r"\b" + re.escape(nome) + r"\b", texto):
            meses_encontrados.append(
                nome
            )

    if len(meses_encontrados) >= 2:

        return [
            (
                f"Qual o {indicador} "
                f"em {mes}{texto_ano}?"
            )
            for mes
            in meses_encontrados
        ]

    # --------------------------------------------------------
    # MÊS A MÊS - SEMESTRE / TRIMESTRE
    # --------------------------------------------------------

    if (
        "mês a mês" in texto
        or "mes a mes" in texto
    ):

        meses_periodo = None

        if (
            "primeiro semestre" in texto
            or "1º semestre" in texto
            or "1° semestre" in texto
            or "1o semestre" in texto
        ):
            meses_periodo = [
                "janeiro",
                "fevereiro",
                "março",
                "abril",
                "maio",
                "junho"
            ]

        elif (
            "segundo semestre" in texto
            or "2º semestre" in texto
            or "2° semestre" in texto
            or "2o semestre" in texto
        ):
            meses_periodo = [
                "julho",
                "agosto",
                "setembro",
                "outubro",
                "novembro",
                "dezembro"
            ]

        elif (
            "primeiro trimestre" in texto
            or "1º trimestre" in texto
            or "q1" in texto
        ):
            meses_periodo = [
                "janeiro",
                "fevereiro",
                "março"
            ]

        elif (
            "segundo trimestre" in texto
            or "2º trimestre" in texto
            or "q2" in texto
        ):
            meses_periodo = [
                "abril",
                "maio",
                "junho"
            ]

        elif (
            "terceiro trimestre" in texto
            or "3º trimestre" in texto
            or "q3" in texto
        ):
            meses_periodo = [
                "julho",
                "agosto",
                "setembro"
            ]

        elif (
            "quarto trimestre" in texto
            or "4º trimestre" in texto
            or "q4" in texto
        ):
            meses_periodo = [
                "outubro",
                "novembro",
                "dezembro"
            ]

        if meses_periodo:

            return [
                (
                    f"Qual o {indicador} "
                    f"em {mes}{texto_ano}?"
                )
                for mes
                in meses_periodo
            ]

    return None



# ============================================================
# DECOMPOR PERGUNTAS COMPOSTAS
# ============================================================


def decompor_pergunta_composta(pergunta):
    
    # ========================================================
    # NÃO É COMPOSTA → NÃO GASTA GROQ
    # ========================================================

    if not parece_pergunta_composta(
        pergunta
    ):

        return [pergunta]

    # ========================================================
    # TENTAR DECOMPOSIÇÃO LOCAL
    # ========================================================

    perguntas_locais = (
        decompor_composta_localmente(
            pergunta
        )
    )

    if perguntas_locais:

        print(
            "Decomposição local utilizada "
            "(Groq não foi chamada)."
        )

        return perguntas_locais


    
    prompt = f"""
Você recebe uma pergunta feita para um agente comercial de BI.

Sua tarefa é verificar se existem DUAS OU MAIS consultas
independentes dentro da mesma mensagem.

Se houver, divida a pergunta em perguntas simples e completas.

IMPORTANTE:
- Preserve indicador, mês, ano, região, cliente e demais filtros.
- Cada pergunta deve fazer sentido sozinha.
- Não invente informações.
- Não responda às perguntas.
- Se houver apenas uma consulta, retorne somente a pergunta original.
- Retorne SOMENTE JSON válido.

Formato:

{{
    "perguntas": [
        "pergunta 1",
        "pergunta 2"
    ]
}}

Exemplos:

Entrada:
"Qual o faturamento de julho e agosto?"

Saída:
{{
    "perguntas": [
        "Qual o faturamento de julho?",
        "Qual o faturamento de agosto?"
    ]
}}

Entrada:
"Qual o faturamento de julho e também por filial?"

Saída:
{{
    "perguntas": [
        "Qual o faturamento de julho?",
        "Qual o faturamento de julho por filial?"
    ]
}}

Entrada:
"Como foi o percentual da meta em julho e como está em agosto?"

Saída:
{{
    "perguntas": [
        "Qual foi o percentual da meta em julho?",
        "Qual é o percentual da meta em agosto?"
    ]
}}

Entrada:
"Quanto faturamos em agosto?"

Saída:
{{
    "perguntas": [
        "Quanto faturamos em agosto?"
    ]
}}

PERGUNTA:

{pergunta}
"""

    try:

        resposta = (
            groq_client
            .chat
            .completions
            .create(
                model=GROQ_MODEL,
                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                response_format={
                    "type": "json_object"
                },
                temperature=0
            )
        )

        texto = (
            resposta
            .choices[0]
            .message
            .content
        )

        dados = json.loads(texto)

        perguntas = dados.get(
            "perguntas",
            []
        )

        if not isinstance(perguntas, list):

            return [pergunta]

        perguntas = [
            p.strip()
            for p in perguntas
            if isinstance(p, str)
            and p.strip()
        ]

        if not perguntas:

            return [pergunta]

        return perguntas

    except Exception as erro:

        print(
            "Falha ao decompor pergunta:",
            type(erro).__name__,
            "-",
            str(erro)[:300]
        )

        return [pergunta]







def perguntar(
    pergunta
):

    _t_pergunta_inicio = time.perf_counter()

    print("\n" + "=" * 70)
    print(f"[PERGUNTA] {pergunta}")
    print("=" * 70)

    try:

        _t_interpretacao_inicio = time.perf_counter()

        interpretacao, origem = (
            interpretar_pergunta(
                pergunta
            )
        )

        _t_interpretacao_fim = time.perf_counter()

        print(
            f"[TEMPO] Interpretação: "
            f"{_t_interpretacao_fim - _t_interpretacao_inicio:.3f}s "
            f"| origem={origem} "
            f"| operação={interpretacao.operacao}"
        )

    except Exception as erro:

        return {
            "erro": str(erro),

            "resposta": (
                "Não consegui interpretar "
                "a pergunta neste momento."
            )
        }

    if interpretacao.fora_escopo:
    
        return {
            "fora_escopo": True,
            "interpretado_por": origem,
            "resposta": (
                "🤖 Sou o Agente Comercial Mallory e estou "
                "preparado somente para responder perguntas "
                "relacionadas aos indicadores comerciais."
            )
        }


    filtros = (
        interpretacao
        .filtros
        .model_dump(
            exclude_none=True
        )
    )

    # Períodos em nível de dia valem para qualquer indicador e
    # qualquer operação (valor, ranking, resumo de meta/gerencial).
    filtros = _aplicar_periodo_calendario(
        pergunta,
        filtros,
        interpretacao
    )


    # ========================================================
    # VALOR
    # ========================================================

    if interpretacao.operacao == "valor":

        valor = consultar_valor(
            interpretacao.indicador,
            filtros
        )

        resposta = construir_resposta_valor(
            interpretacao.indicador,
            filtros,
            valor
        )

        resultado = {
            "operacao": "valor",

            "indicador":
                interpretacao.indicador,

            "filtros":
                filtros,

            "valor":
                valor,

            "interpretado_por":
                origem,

            "resposta":
                resposta
        }


    # ========================================================
    # RESUMO DE META
    # ========================================================

    elif interpretacao.operacao == "resumo_meta":

        dados = consultar_resumo_meta(
            filtros
        )

        resposta = construir_resposta_resumo_meta(
            filtros,
            dados
        )

        resultado = {
            "operacao":
                "resumo_meta",

            "indicador":
                interpretacao.indicador,

            "filtros":
                filtros,

            "dados":
                dados,

            "interpretado_por":
                origem,

            "resposta":
                resposta
        }

    # ========================================================
    # RESUMO GERENCIAL
    # ========================================================
    
    elif interpretacao.operacao == "resumo_gerencial":
    
        dados = consultar_resumo_gerencial(
            filtros
        )
    
        resposta = construir_resposta_resumo_gerencial(
            filtros,
            dados
        )
    
        resultado = {
            "operacao":
                "resumo_gerencial",
    
            "indicador":
                interpretacao.indicador,
    
            "filtros":
                filtros,
    
            "dados":
                dados,
    
            "interpretado_por":
                origem,
    
            "resposta":
                resposta
        }
    


    # ========================================================
    # RANKING
    # ========================================================

    elif interpretacao.operacao == "ranking":

        if interpretacao.agrupar_por is None:

            return {
                "erro":
                    "dimensao_ausente",

                "resposta": (
                    "Não consegui identificar "
                    "a dimensão do ranking."
                )
            }


        ranking = consultar_ranking(
            interpretacao.indicador,
            interpretacao.agrupar_por,
            interpretacao.top_n or 20,
            interpretacao.ordem or "desc",
            filtros
        )

        resposta = construir_resposta_ranking(
            interpretacao.indicador,
            interpretacao.agrupar_por,
            ranking,
            filtros,
            top_n=interpretacao.top_n,
            ordem=interpretacao.ordem or "desc",
            pergunta=pergunta
        )

        resultado = {
            "operacao":
                "ranking",

            "indicador":
                interpretacao.indicador,

            "agrupar_por":
                interpretacao.agrupar_por,

            "top_n":
                interpretacao.top_n,

            "ordem":
                interpretacao.ordem,

            "filtros":
                filtros,

            "ranking":
                ranking,

            "interpretado_por":
                origem,

            "resposta":
                resposta
        }


    # ========================================================
    # OPERAÇÃO DESCONHECIDA
    # ========================================================

    else:

        return {
            "erro":
                "operacao_desconhecida",

            "resposta":
                "Operação não reconhecida."
        }


    # ========================================================
    # SALVAR CONTEXTO
    # ========================================================

    contexto_conversa[
        "ultima_interpretacao"
    ] = (
        interpretacao.model_dump()
    )


    _t_pergunta_fim = time.perf_counter()

    print(
        f"[TEMPO TOTAL perguntar()] "
        f"{_t_pergunta_fim - _t_pergunta_inicio:.3f}s"
    )
    print("=" * 70)

    return resultado

# ============================================================
# 40. CONVERSAR
# ============================================================
def eh_saudacao(pergunta):

    texto = (
        pergunta
        .strip()
        .lower()
    )

    # remove pontuação
    texto = re.sub(
        r"[^\w\sáàâãéèêíïóôõöúç]",
        "",
        texto
    )

    saudacoes = [
        "oi",
        "olá",
        "ola",
        "bom dia",
        "boa tarde",
        "boa noite",
        "e aí",
        "e ai",
        "opa",
        "hello",
        "hey",
        "ei",
        "oi meu amigo",
        "oi minha amiga",
        "bom dia bot",
        "boa tarde bot",
        "boa noite bot"
    ]

    return texto in saudacoes




def conversar(
    pergunta,
    mostrar_origem=False
):

    # ========================================================
    # SAUDAÇÕES
    # ========================================================

    if eh_saudacao(pergunta):

        return (
            "👋 Olá! Como posso lhe ajudar hoje?\n\n"
            "É só escrever sua pergunta. 🚀"
        )

    # ========================================================
    # VERIFICAR PERGUNTA COMPOSTA
    # ========================================================

    perguntas = decompor_pergunta_composta(
        pergunta
    )

    # ========================================================
    # PERGUNTA SIMPLES
    # ========================================================

    if len(perguntas) == 1:

        resultado = perguntar(
            perguntas[0]
        )

        if mostrar_origem:

            origem = resultado.get(
                "interpretado_por",
                "indisponível"
            )

            print(
                f"[Interpretado por: "
                f"{origem}]"
            )

        return resultado[
            "resposta"
        ]

    # ========================================================
    # PERGUNTA COMPOSTA
    # ========================================================

    respostas = []

    for pergunta_simples in perguntas:

        resultado = perguntar(
            pergunta_simples
        )

        resposta = resultado.get(
            "resposta"
        )

        if resposta:

            respostas.append(
                resposta
            )

        if mostrar_origem:

            origem = resultado.get(
                "interpretado_por",
                "indisponível"
            )

            print(
                f"[Interpretado por: "
                f"{origem}] "
                f"{pergunta_simples}"
            )

    # ========================================================
    # JUNTAR RESPOSTAS
    # ========================================================

    if not respostas:

        return (
            "Não consegui interpretar "
            "a pergunta neste momento."
        )

    return "\n\n".join(
        respostas
    )
# ============================================================
# FIM
# ============================================================

print(
    "Agente Power BI V6 "
    "carregado com sucesso!"
)







#=======================================================================================


# ============================================================
# 41. EXTENSÕES V8 - SEM ALTERAR A LÓGICA VALIDADA DA V6
# ============================================================

import io
import base64
from threading import RLock

# Guarda um contexto independente por chat sem alterar as funções internas da V6.
_contextos_por_chat = {}
_contexto_lock = RLock()


def _executar_no_contexto_usuario(chat_id, funcao):
    global contexto_conversa

    if chat_id is None:
        return funcao()

    chave = str(chat_id)

    with _contexto_lock:
        contexto_anterior_global = contexto_conversa
        contexto_usuario = _contextos_por_chat.setdefault(
            chave,
            {"ultima_interpretacao": None}
        )
        contexto_conversa = contexto_usuario

        try:
            resultado = funcao()
            _contextos_por_chat[chave] = contexto_conversa
            return resultado
        finally:
            contexto_conversa = contexto_anterior_global


def limpar_contexto_usuario(chat_id):
    chave = str(chat_id)
    with _contexto_lock:
        _contextos_por_chat.pop(chave, None)
    return {
        "ok": True,
        "chat_id": chave,
        "mensagem": "Contexto da conversa limpo."
    }


# ============================================================
# 42. COMPARAÇÕES GENÉRICAS
# ============================================================

def _texto_normalizado(texto):
    return (
        texto.lower()
        .replace("á", "a")
        .replace("à", "a")
        .replace("â", "a")
        .replace("ã", "a")
        .replace("é", "e")
        .replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c")
    )


def _tem_marcador_comparacao(pergunta):
    t = _texto_normalizado(pergunta)
    return any(x in t for x in [
        "compare", "comparar", "comparacao", "comparativo",
        "comparativa", "comparativos", "comparativas",
        " versus ", " vs ", " x ", "diferenca entre",
        "qual faturou mais", "qual vendeu mais",
        "qual foi maior", "qual e maior"
    ])


def _indicador_da_pergunta(pergunta):
    t = _texto_normalizado(pergunta)

    if any(x in t for x in [
        "previsao de faturamento",
        "projecao de faturamento",
        "forecast"
    ]):
        return "previsao_faturamento"

    if any(x in t for x in [
        "mais comprou",
        "comprou mais",
        "mais vendeu",
        "vendeu mais",
        "mais produtos",
        "mais unidades"
    ]):
        return "quantidade"

    # Novos indicadores - termos mais específicos primeiro.
    if any(x in t for x in [
        "meta de margem liquida",
        "percentual da meta de margem liquida"
    ]):
        return "meta_margem_liquida"

    if any(x in t for x in [
        "faturamento + entregas",
        "fat + entregas",
        "faturamento e entregas"
    ]):
        return "faturamento_entregas"

    if any(x in t for x in [
        "carteira faturavel do dia",
        "carteira faturavel dia"
    ]):
        return "carteira_faturavel_dia"

    if any(x in t for x in [
        "carteira faturavel do mes",
        "carteira faturavel mes"
    ]):
        return "carteira_faturavel_mes"

    if any(x in t for x in [
        "faturamento da carteira",
        "faturamento carteira"
    ]):
        return "faturamento_carteira"

    if any(x in t for x in [
        "meta de quantidade",
        "meta em quantidade",
        "meta de vendas"
    ]):
        return "meta_quantidade"

    if any(x in t for x in [
        "quantidade vendida",
        "quantidade de vendas",
        "qtd vendida"
    ]) or re.search(r"\bqtd\b", t):
        return "quantidade"

    if "entregas" in t:
        return "entregas"

    # Siglas do dashboard:
    # MRGL = valor da margem líquida em R$.
    # MRGB = valor da margem bruta em R$.
    if re.search(r"\bmrgl\b", t):
        return "valor_margem_liquida"

    if re.search(r"\bmrgb\b", t):
        return "valor_margem_bruta"

    if "margem liquida" in t:
        if any(x in t for x in ["r$", "reais", "valor"]):
            return "valor_margem_liquida"
        return "margem_liquida"

    if "margem bruta" in t:
        if any(x in t for x in ["r$", "reais", "valor"]):
            return "valor_margem_bruta"
        return "margem_bruta"

    if any(x in t for x in ["atingimento", "percentual da meta", "% da meta"]):
        return "atingimento_meta_faturamento"

    if any(x in t for x in ["quanto falta", "desvio da meta", "diferenca para a meta"]):
        return "desvio_meta_faturamento"

    if "meta" in t and "margem" not in t:
        return "meta_faturamento"

    return "faturamento"


_GRUPOS_REGIAO = {
    "SUL": ["SUL"],
    "NORDESTE": ["NORDESTE 1", "NORDESTE 2"],
    "NORDESTE 1": ["NORDESTE 1"],
    "NORDESTE 2": ["NORDESTE 2"],
    "NORTE": ["NORTE 1", "NORTE 2"],
    "NORTE 1": ["NORTE 1"],
    "NORTE 2": ["NORTE 2"],
    "SUDESTE": ["SUDESTE 1", "SUDESTE 2", "SUDESTE 3"],
    "SUDESTE 1": ["SUDESTE 1"],
    "SUDESTE 2": ["SUDESTE 2"],
    "SUDESTE 3": ["SUDESTE 3"],
    "MINAS GERAIS": ["MINAS GERAIS"],
    "ESPIRITO SANTO": ["ESPIRITO SANTO"],
    "CENTRO-OESTE": ["CENTRO-OESTE"],
    "MLY EXPORTACAO": ["MLY EXPORTACAO"],
    "MLY E-COMMERCE 1P": ["MLY E-COMMERCE 1P"],
    "MLY E-COMMERCE 3P": ["MLY E-COMMERCE 3P"],
    "MLY VENDA DIRETA": ["MLY VENDA DIRETA"],
    "MLY ASSISTENCIA TECNICA": ["MLY ASSISTENCIA TECNICA"],
    "MLY ESPECIALIZADO AR": ["MLY ESPECIALIZADO AR"],
    "MLY FUNCIONARIO": ["MLY FUNCIONARIO"],
    "NOVOS NEGOCIOS": ["NOVOS NEGOCIOS"],
}


def _extrair_periodo_pergunta(pergunta):
    texto = _texto_normalizado(pergunta)
    filtros = {}

    match_ano = re.search(r"\b(20\d{2})\b", texto)
    if match_ano:
        filtros["ano"] = match_ano.group(1)
    elif _pergunta_pede_ano_inteiro(pergunta):
        filtros["ano"] = "Ano atual"

    meses_sem_acento = {
        "janeiro": "01", "fevereiro": "02", "marco": "03",
        "abril": "04", "maio": "05", "junho": "06",
        "julho": "07", "agosto": "08", "setembro": "09",
        "outubro": "10", "novembro": "11", "dezembro": "12"
    }

    encontrados = [
        numero
        for nome, numero in meses_sem_acento.items()
        if re.search(r"\b" + re.escape(nome) + r"\b", texto)
    ]
    if len(encontrados) == 1:
        filtros["mes"] = normalizar_mes_powerbi(
            encontrados[0],
            filtros.get("ano", "Ano atual")
        )
        filtros.setdefault("ano", "Ano atual")

    return filtros


def _comparacao_regioes(pergunta):
    t = _texto_normalizado(pergunta).upper()
    achados = []

    # Mais específicos primeiro, para não capturar NORDESTE dentro de NORDESTE 1.
    for nome in sorted(_GRUPOS_REGIAO, key=len, reverse=True):
        if nome in t:
            # Não adiciona o grupo genérico se uma subdivisão dele já apareceu explicitamente.
            if nome in {"NORDESTE", "NORTE", "SUDESTE"}:
                if any(x.startswith(nome + " ") for x, _ in achados):
                    continue
            achados.append((nome, _GRUPOS_REGIAO[nome]))

    # Remove grupo genérico duplicado quando há subdivisão explícita.
    nomes = {x[0] for x in achados}
    finais = []
    for nome, valores in achados:
        if nome in {"NORDESTE", "NORTE", "SUDESTE"} and any(
            n.startswith(nome + " ") for n in nomes
        ):
            continue
        finais.append((nome, valores))

    # Mantém ordem em que aparecem na pergunta.
    finais.sort(key=lambda x: t.find(x[0]))
    return finais if len(finais) >= 2 else None


def _comparacao_meses(pergunta):
    t = _texto_normalizado(pergunta)
    meses = {
        "janeiro": "01", "fevereiro": "02", "marco": "03",
        "abril": "04", "maio": "05", "junho": "06",
        "julho": "07", "agosto": "08", "setembro": "09",
        "outubro": "10", "novembro": "11", "dezembro": "12"
    }
    encontrados = []
    for nome, num in meses.items():
        match_mes = re.search(
            r"\b" + re.escape(nome) + r"\b",
            t
        )
        if match_mes:
            encontrados.append(
                (
                    nome,
                    num,
                    match_mes.start()
                )
            )
    encontrados.sort(key=lambda x: x[2])
    return [(nome.capitalize(), num) for nome, num, _ in encontrados] if len(encontrados) >= 2 else None


def _comparacao_anos(pergunta):
    anos = re.findall(r"\b(20\d{2})\b", pergunta)
    anos = list(dict.fromkeys(anos))
    return [(ano, ano) for ano in anos] if len(anos) >= 2 else None


def _montar_resposta_comparacao(indicador, dados, descricao_dimensao):
    config = mapa_indicadores[indicador]
    linhas = [f"Comparativo de {config['descricao'].lower()} por {descricao_dimensao}:"]

    for item in dados:
        valor = item["valor"]
        texto_valor = (
            formatar_valor(valor, config["formato"])
            if valor is not None else "Sem resultado"
        )
        linhas.append(f"{item['item']}: {texto_valor}")

    validos = [x for x in dados if x["valor"] is not None]
    if len(validos) == 2:
        a, b = validos
        dif = abs(a["valor"] - b["valor"])
        maior, menor = (a, b) if a["valor"] >= b["valor"] else (b, a)
        linhas.append("")
        linhas.append("Diferença: " + formatar_valor(dif, config["formato"]))

        if menor["valor"] not in (0, None):
            pct = (maior["valor"] - menor["valor"]) / abs(menor["valor"]) * 100
            linhas.append(
                f"{maior['item']} está {pct:.1f}% acima de {menor['item']}.".replace(".", ",")
            )

    return {
        "tipo_resposta": "comparacao",
        "indicador": indicador,
        "dados": dados,
        "resposta": "\n".join(linhas)
    }


def _granularidade_temporal_comparacao(pergunta):
    """
    Retorna a granularidade temporal solicitada pelo usuário.
    Prioridade: semestre > trimestre > mês > ano.
    """
    t = _texto_normalizado(pergunta)

    if any(
        termo in t
        for termo in [
            "semestre a semestre",
            "por semestre",
            "em cada semestre",
            "semestral",
            "semestralmente",
            "1 semestre",
            "2 semestre",
            "primeiro semestre",
            "segundo semestre"
        ]
    ):
        return "semestre"

    if any(
        termo in t
        for termo in [
            "trimestre a trimestre",
            "por trimestre",
            "em cada trimestre",
            "trimestral",
            "trimestralmente",
            "1 trimestre",
            "2 trimestre",
            "3 trimestre",
            "4 trimestre",
            "primeiro trimestre",
            "segundo trimestre",
            "terceiro trimestre",
            "quarto trimestre",
            "q1", "q2", "q3", "q4"
        ]
    ):
        return "trimestre"

    if any(
        termo in t
        for termo in [
            "mes a mes",
            "por mes",
            "em cada mes",
            "evolucao mensal",
            "mensalmente"
        ]
    ):
        return "mes"

    if any(
        termo in t
        for termo in [
            "ano a ano",
            "por ano",
            "em cada ano",
            "anualmente",
            "evolucao anual"
        ]
    ):
        return "ano"

    # "ao longo de 2026" normalmente indica uma visão mensal.
    if "ao longo de" in t or "ao longo do ano" in t:
        return "mes"

    return None


def _eh_comparacao_temporal(pergunta):
    return _granularidade_temporal_comparacao(pergunta) is not None


def _extrair_dimensao_comparacao(pergunta):
    """
    Identifica a dimensão que está sendo comparada.
    Segmento ainda não é suportado porque não existe
    mapeamento de segmento no modelo atual do agente.
    """
    t = _texto_normalizado(pergunta)

    regras = [
        ("regiao", ["regioes", "regiao", "filiais", "filial"]),
        ("representante", ["representantes", "representante"]),
        ("plataforma", ["plataformas", "plataforma"]),
        ("familia", ["familias", "familia"]),
        ("produto", ["produtos", "produto"]),
        ("cliente", ["clientes", "cliente"]),
        ("linha", ["linhas", "linha"]),
        ("classe", ["classes", "classe"]),
    ]

    for dimensao, termos in regras:
        for termo in termos:
            if re.search(rf"\b{re.escape(termo)}\b", t):
                return dimensao, termo

    return None, None


def _limpar_nome_entidade_comparacao(texto):
    texto = texto.strip(" \t\n\r,;:-")

    # Remove conectivos/preposições que podem ficar nas pontas.
    texto = re.sub(
        r"^(?:a|as|o|os|da|das|do|dos|de)\s+",
        "",
        texto,
        flags=re.IGNORECASE
    )

    texto = re.sub(
        r"\s+(?:em|no|na|nos|nas)\s+(?:20\d{2}|ano atual)$",
        "",
        texto,
        flags=re.IGNORECASE
    )

    return texto.strip(" \t\n\r,;:-")


def _extrair_duas_entidades_comparacao(pergunta):
    """
    Extrai duas entidades de frases como:
    - regiões Sul e Nordeste mês a mês em 2026
    - clientes CLIENTE A e CLIENTE B mês a mês
    - representantes X com Y por mês
    - plataformas GAZIN versus AMAZON em 2026
    """
    dimensao, termo_dimensao = _extrair_dimensao_comparacao(pergunta)

    if not dimensao:
        return None

    original = pergunta.strip()
    normalizado = _texto_normalizado(original)

    # Localiza o termo da dimensão no texto normalizado.
    match_dim = re.search(
        rf"\b{re.escape(termo_dimensao)}\b",
        normalizado
    )

    if not match_dim:
        return None

    # Como normalização mantém o mesmo comprimento dos caracteres
    # usados aqui, podemos reaproveitar os índices no texto original.
    trecho = original[match_dim.end():].strip()

    # Corta marcadores temporais que vêm depois das entidades.
    marcadores_fim = [
        " mês a mês",
        " mes a mes",
        " por mês",
        " por mes",
        " em cada mês",
        " em cada mes",
        " ao longo de",
        " ao longo do ano",
        " mensalmente",
        " em 20",
        " no ano de 20",
    ]

    trecho_norm = _texto_normalizado(trecho)
    cortes = []

    for marcador in marcadores_fim:
        pos = trecho_norm.find(_texto_normalizado(marcador))
        if pos >= 0:
            cortes.append(pos)

    if cortes:
        trecho = trecho[:min(cortes)].strip()

    # Remove pontuação final.
    trecho = trecho.strip(" .?!,;:")

    # Separadores mais explícitos primeiro.
    separadores = [
        r"\s+versus\s+",
        r"\s+vs\.?\s+",
        r"\s+com\s+",
        r"\s+e\s+",
    ]

    partes = None

    for sep in separadores:
        p = re.split(
            sep,
            trecho,
            maxsplit=1,
            flags=re.IGNORECASE
        )

        if len(p) == 2:
            partes = p
            break

    if not partes:
        return None

    entidade_a = _limpar_nome_entidade_comparacao(partes[0])
    entidade_b = _limpar_nome_entidade_comparacao(partes[1])

    if not entidade_a or not entidade_b:
        return None

    return {
        "dimensao": dimensao,
        "entidade_a": entidade_a,
        "entidade_b": entidade_b
    }


def _filtro_entidade_comparacao(dimensao, entidade):
    """
    Converte a entidade informada no filtro usado por consultar_valor.
    """
    if dimensao == "regiao":
        entidade_upper = entidade.strip().upper()

        valores = _GRUPOS_REGIAO.get(
            entidade_upper,
            [entidade_upper]
        )

        return (
            valores
            if len(valores) > 1
            else valores[0]
        )

    return entidade.strip()


def _formatar_percentual_comparacao(valor):
    return f"{valor:.1f}%".replace(".", ",")


def _periodos_da_granularidade(granularidade, ano):
    """
    Retorna períodos consultáveis no Power BI.
    Cada item contém:
      - chave
      - nome
      - meses
    """
    ano_int = int(ano)
    agora = datetime.now()

    if granularidade == "mes":
        limite = agora.month if ano_int == agora.year else 12
        return [
            {
                "chave": f"{n:02d}",
                "nome": nome_mes_resposta(f"{n:02d}").capitalize(),
                "meses": [f"{n:02d}"]
            }
            for n in range(1, limite + 1)
        ]

    if granularidade == "trimestre":
        mapa = [
            ("T1", "1º trimestre", ["01", "02", "03"]),
            ("T2", "2º trimestre", ["04", "05", "06"]),
            ("T3", "3º trimestre", ["07", "08", "09"]),
            ("T4", "4º trimestre", ["10", "11", "12"]),
        ]

        if ano_int == agora.year:
            mapa = [
                item
                for item in mapa
                if int(item[2][0]) <= agora.month
            ]

        return [
            {"chave": chave, "nome": nome, "meses": meses}
            for chave, nome, meses in mapa
        ]

    if granularidade == "semestre":
        mapa = [
            ("S1", "1º semestre", ["01", "02", "03", "04", "05", "06"]),
            ("S2", "2º semestre", ["07", "08", "09", "10", "11", "12"]),
        ]

        if ano_int == agora.year:
            mapa = [
                item
                for item in mapa
                if int(item[2][0]) <= agora.month
            ]

        return [
            {"chave": chave, "nome": nome, "meses": meses}
            for chave, nome, meses in mapa
        ]

    return []


def _consultar_entidade_periodo(indicador, dimensao, entidade, ano, meses):
    """
    Consulta um indicador para uma entidade e um conjunto de meses.
    Quando há mais de um mês, usa filtros.mes como lista.
    """
    filtros = {
        "ano": ano,
        dimensao: entidade
    }

    if len(meses) == 1:
        filtros["mes"] = normalizar_mes_powerbi(
            meses[0],
            ano
        )
    else:
        filtros["mes"] = meses

    return consultar_valor(
        indicador,
        filtros
    )


def _usuario_pediu_tabela(pergunta):
    t = _texto_normalizado(pergunta)

    return any(
        termo in t
        for termo in [
            "tabela",
            "formato de tabela",
            "em tabela",
            "tabular"
        ]
    )


def _montar_tabela_comparacao_temporal(
    dados_periodos,
    entidade_a,
    entidade_b,
    config
):
    """
    Monta uma tabela textual simples para Telegram.
    """
    linhas = []

    cabecalho = (
        f"Período | {entidade_a} | {entidade_b} | Diferença | Melhor"
    )

    linhas.append(cabecalho)
    linhas.append("-" * min(len(cabecalho), 120))

    for item in dados_periodos:
        va = item["valor_a"]
        vb = item["valor_b"]

        texto_a = (
            formatar_valor(va, config["formato"])
            if va is not None
            else "Sem resultado"
        )

        texto_b = (
            formatar_valor(vb, config["formato"])
            if vb is not None
            else "Sem resultado"
        )

        if va is None or vb is None:
            diferenca = "Sem resultado"
            melhor = "-"
        else:
            diferenca = formatar_valor(
                abs(va - vb),
                config["formato"]
            )

            if va > vb:
                melhor = entidade_a
            elif vb > va:
                melhor = entidade_b
            else:
                melhor = "Empate"

        linhas.append(
            f"{item['periodo_nome']} | "
            f"{texto_a} | "
            f"{texto_b} | "
            f"{diferenca} | "
            f"{melhor}"
        )

    return "\n".join(linhas)


def _rotulo_granularidade(granularidade):
    return {
        "mes": "mês a mês",
        "trimestre": "trimestre a trimestre",
        "semestre": "semestre a semestre",
        "ano": "ano a ano"
    }.get(granularidade, granularidade)


def consultar_comparacao_temporal_generica(pergunta):
    """
    Compara duas entidades ao longo do tempo usando as dimensões
    mapeadas no agente e as granularidades:
    - mês
    - trimestre
    - semestre
    - ano

    Também suporta saída em formato de tabela.
    """
    if not _tem_marcador_comparacao(pergunta):
        return None

    granularidade = _granularidade_temporal_comparacao(
        pergunta
    )

    if granularidade is None:
        return None

    entidades = _extrair_duas_entidades_comparacao(
        pergunta
    )

    if not entidades:
        return None

    indicador = _indicador_da_pergunta(
        pergunta
    )

    config = mapa_indicadores[
        indicador
    ]

    dimensao = entidades[
        "dimensao"
    ]

    entidade_a = entidades[
        "entidade_a"
    ]

    entidade_b = entidades[
        "entidade_b"
    ]

    if dimensao not in {
        "regiao",
        "cliente",
        "representante",
        "produto",
        "familia",
        "plataforma",
        "linha",
        "classe"
    }:
        return None

    anos = re.findall(
        r"\b(20\d{2})\b",
        pergunta
    )

    anos = list(
        dict.fromkeys(
            anos
        )
    )

    # --------------------------------------------------------
    # COMPARAÇÃO ANO A ANO
    # --------------------------------------------------------
    if granularidade == "ano":
        if len(anos) < 2:
            return None

        filtro_a = _filtro_entidade_comparacao(
            dimensao,
            entidade_a
        )

        filtro_b = _filtro_entidade_comparacao(
            dimensao,
            entidade_b
        )

        dados_periodos = []

        for ano in anos:
            valor_a = consultar_valor(
                indicador,
                {
                    "ano": ano,
                    dimensao: filtro_a
                }
            )

            valor_b = consultar_valor(
                indicador,
                {
                    "ano": ano,
                    dimensao: filtro_b
                }
            )

            dados_periodos.append({
                "periodo": ano,
                "periodo_nome": ano,
                "valor_a": valor_a,
                "valor_b": valor_b
            })

    # --------------------------------------------------------
    # MÊS / TRIMESTRE / SEMESTRE
    # --------------------------------------------------------
    else:
        ano = (
            anos[0]
            if anos
            else str(
                datetime.now().year
            )
        )

        filtro_a = _filtro_entidade_comparacao(
            dimensao,
            entidade_a
        )

        filtro_b = _filtro_entidade_comparacao(
            dimensao,
            entidade_b
        )

        periodos = _periodos_da_granularidade(
            granularidade,
            ano
        )

        dados_periodos = []

        for periodo in periodos:
            valor_a = _consultar_entidade_periodo(
                indicador,
                dimensao,
                filtro_a,
                ano,
                periodo["meses"]
            )

            valor_b = _consultar_entidade_periodo(
                indicador,
                dimensao,
                filtro_b,
                ano,
                periodo["meses"]
            )

            dados_periodos.append({
                "periodo": periodo["chave"],
                "periodo_nome": periodo["nome"],
                "valor_a": valor_a,
                "valor_b": valor_b
            })

    validos_a = [
        item
        for item in dados_periodos
        if item["valor_a"] is not None
    ]

    validos_b = [
        item
        for item in dados_periodos
        if item["valor_b"] is not None
    ]

    if not validos_a and not validos_b:
        return None

    total_a = sum(
        item["valor_a"] or 0
        for item in dados_periodos
    )

    total_b = sum(
        item["valor_b"] or 0
        for item in dados_periodos
    )

    melhor_a = (
        max(
            validos_a,
            key=lambda item: item["valor_a"]
        )
        if validos_a
        else None
    )

    melhor_b = (
        max(
            validos_b,
            key=lambda item: item["valor_b"]
        )
        if validos_b
        else None
    )

    vitorias_a = 0
    vitorias_b = 0
    empates = 0

    for item in dados_periodos:
        va = item["valor_a"]
        vb = item["valor_b"]

        if va is None or vb is None:
            continue

        if va > vb:
            vitorias_a += 1

        elif vb > va:
            vitorias_b += 1

        else:
            empates += 1

    descricao_dimensao = (
        "filial"
        if (
            dimensao == "regiao"
            and any(
                termo in _texto_normalizado(pergunta)
                for termo in [
                    "filial",
                    "filiais"
                ]
            )
        )
        else mapa_dimensoes[
            dimensao
        ]["descricao"]
    )

    linhas = [
        (
            f"Comparativo "
            f"{_rotulo_granularidade(granularidade)} "
            f"de {config['descricao'].lower()} "
            f"por {descricao_dimensao}:"
        ),
        ""
    ]

    # --------------------------------------------------------
    # SAÍDA EM TABELA
    # --------------------------------------------------------
    if _usuario_pediu_tabela(
        pergunta
    ):
        linhas.append(
            _montar_tabela_comparacao_temporal(
                dados_periodos,
                entidade_a,
                entidade_b,
                config
            )
        )

    # --------------------------------------------------------
    # SAÍDA TEXTUAL NORMAL
    # --------------------------------------------------------
    else:
        for item in dados_periodos:
            texto_a = (
                formatar_valor(
                    item["valor_a"],
                    config["formato"]
                )
                if item["valor_a"] is not None
                else "Sem resultado"
            )

            texto_b = (
                formatar_valor(
                    item["valor_b"],
                    config["formato"]
                )
                if item["valor_b"] is not None
                else "Sem resultado"
            )

            linhas.append(
                f"{item['periodo_nome']}: "
                f"{entidade_a} {texto_a} | "
                f"{entidade_b} {texto_b}"
            )

    linhas.append("")
    linhas.append(
        "Resumo do período:"
    )

    linhas.append(
        f"{entidade_a}: "
        f"{formatar_valor(total_a, config['formato'])}"
    )

    linhas.append(
        f"{entidade_b}: "
        f"{formatar_valor(total_b, config['formato'])}"
    )

    if melhor_a:
        linhas.append(
            f"Melhor período de {entidade_a}: "
            f"{melhor_a['periodo_nome']} "
            f"({formatar_valor(melhor_a['valor_a'], config['formato'])})"
        )

    if melhor_b:
        linhas.append(
            f"Melhor período de {entidade_b}: "
            f"{melhor_b['periodo_nome']} "
            f"({formatar_valor(melhor_b['valor_b'], config['formato'])})"
        )

    if total_a > total_b:
        vencedor = entidade_a
        perdedor = entidade_b
        maior_total = total_a
        menor_total = total_b

    elif total_b > total_a:
        vencedor = entidade_b
        perdedor = entidade_a
        maior_total = total_b
        menor_total = total_a

    else:
        vencedor = None
        perdedor = None
        maior_total = total_a
        menor_total = total_b

    linhas.append("")

    if vencedor:
        diferenca = (
            maior_total
            - menor_total
        )

        linhas.append(
            f"Melhor desempenho no período: "
            f"{vencedor}."
        )

        linhas.append(
            "Diferença acumulada: "
            + formatar_valor(
                diferenca,
                config["formato"]
            )
            + "."
        )

        if menor_total not in (
            0,
            None
        ):
            percentual = (
                diferenca
                / abs(
                    menor_total
                )
                * 100
            )

            linhas.append(
                f"{vencedor} ficou "
                f"{_formatar_percentual_comparacao(percentual)} "
                f"acima de {perdedor} no período."
            )

    else:
        linhas.append(
            "Resultado acumulado: empate no período."
        )

    linhas.append(
        f"Períodos à frente: "
        f"{entidade_a} {vitorias_a} | "
        f"{entidade_b} {vitorias_b}"
        + (
            f" | Empates {empates}"
            if empates
            else ""
        )
        + "."
    )

    return {
        "tipo_resposta": (
            "tabela"
            if _usuario_pediu_tabela(
                pergunta
            )
            else "comparacao_temporal"
        ),
        "indicador": indicador,
        "dimensao": dimensao,
        "granularidade": granularidade,
        "entidade_a": entidade_a,
        "entidade_b": entidade_b,
        "dados": dados_periodos,
        "resposta": "\n".join(
            linhas
        )
    }


# ============================================================
# 42B. COMPARAÇÃO ENTRE PERÍODOS
# ============================================================

_MESES_COMPARACAO = {
    "janeiro": "01",
    "fevereiro": "02",
    "marco": "03",
    "abril": "04",
    "maio": "05",
    "junho": "06",
    "julho": "07",
    "agosto": "08",
    "setembro": "09",
    "outubro": "10",
    "novembro": "11",
    "dezembro": "12",
}


def _nome_periodo_mes_ano(mes, ano):
    return f"{nome_mes_resposta(str(mes).zfill(2)).capitalize()}/{ano}"


def _periodo_comparacao(chave, nome, ano, meses):
    return {
        "chave": chave,
        "nome": nome,
        "ano": str(ano),
        "meses": [str(m).zfill(2) for m in meses],
    }


def _extrair_periodos_comparacao(pergunta):
    """Extrai dois períodos explícitos da pergunta."""
    t = _texto_normalizado(pergunta)
    agora = datetime.now()
    periodos = []

    nomes_meses = "|".join(_MESES_COMPARACAO.keys())

    # Intervalos: janeiro a março de 2025
    padrao_intervalo = re.compile(
        rf"\b({nomes_meses})\s*(?:a|ate|-)\s*"
        rf"({nomes_meses})\s*(?:de\s*)?(20\d{{2}})\b"
    )
    spans_usados = []

    for match in padrao_intervalo.finditer(t):
        inicio_nome, fim_nome, ano = match.groups()
        inicio = int(_MESES_COMPARACAO[inicio_nome])
        fim = int(_MESES_COMPARACAO[fim_nome])
        if inicio > fim:
            continue
        meses = [f"{m:02d}" for m in range(inicio, fim + 1)]
        nome = f"{inicio_nome.capitalize()} a {fim_nome.capitalize()}/{ano}"
        periodos.append(
            _periodo_comparacao(
                f"{ano}-{meses[0]}-{meses[-1]}",
                nome,
                ano,
                meses
            )
        )
        spans_usados.append(match.span())

    if len(periodos) >= 2:
        return periodos[:2]

    t_sem_intervalos = t
    for ini, fim in sorted(spans_usados, reverse=True):
        t_sem_intervalos = t_sem_intervalos[:ini] + (" " * (fim - ini)) + t_sem_intervalos[fim:]

    # Trimestres
    padrao_trim = re.compile(
        r"\b("
        r"q[1-4]|"
        r"[1-4](?:º|o)?\s*trimestre|"
        r"primeiro\s+trimestre|segundo\s+trimestre|"
        r"terceiro\s+trimestre|quarto\s+trimestre"
        r")\s*(?:de\s*)?(20\d{2})\b"
    )
    mapa_trim = {
        "1": ["01", "02", "03"],
        "2": ["04", "05", "06"],
        "3": ["07", "08", "09"],
        "4": ["10", "11", "12"],
    }

    for match in padrao_trim.finditer(t):
        rotulo, ano = match.groups()
        r = rotulo.strip()
        if r.startswith("q"):
            n = r[-1]
        elif r.startswith("primeiro"):
            n = "1"
        elif r.startswith("segundo"):
            n = "2"
        elif r.startswith("terceiro"):
            n = "3"
        elif r.startswith("quarto"):
            n = "4"
        else:
            n = re.search(r"[1-4]", r).group(0)

        periodos.append(
            _periodo_comparacao(
                f"{ano}-Q{n}",
                f"{n}º trimestre/{ano}",
                ano,
                mapa_trim[n]
            )
        )

    if len(periodos) >= 2:
        return periodos[:2]

    # Semestres
    padrao_sem = re.compile(
        r"\b("
        r"[12](?:º|o)?\s*semestre|"
        r"primeiro\s+semestre|segundo\s+semestre"
        r")\s*(?:de\s*)?(20\d{2})\b"
    )

    for match in padrao_sem.finditer(t):
        rotulo, ano = match.groups()
        r = rotulo.strip()
        if r.startswith("primeiro") or r.startswith("1"):
            n = "1"
            meses = ["01", "02", "03", "04", "05", "06"]
        else:
            n = "2"
            meses = ["07", "08", "09", "10", "11", "12"]

        periodos.append(
            _periodo_comparacao(
                f"{ano}-S{n}",
                f"{n}º semestre/{ano}",
                ano,
                meses
            )
        )

    if len(periodos) >= 2:
        return periodos[:2]

    # Mês + ano por extenso
    encontrados = []
    padrao_mes_nome = re.compile(
        rf"\b({nomes_meses})\s*(?:de\s*|/|-)?\s*(20\d{{2}})\b"
    )

    for match in padrao_mes_nome.finditer(t_sem_intervalos):
        nome_mes, ano = match.groups()
        mes = _MESES_COMPARACAO[nome_mes]
        encontrados.append((
            match.start(),
            _periodo_comparacao(
                f"{ano}-{mes}",
                _nome_periodo_mes_ano(mes, ano),
                ano,
                [mes]
            )
        ))

    # Mês + ano numérico: 07/2025
    padrao_mes_num = re.compile(
        r"\b(0?[1-9]|1[0-2])\s*[/.-]\s*(20\d{2})\b"
    )

    for match in padrao_mes_num.finditer(t_sem_intervalos):
        mes, ano = match.groups()
        mes = str(mes).zfill(2)
        encontrados.append((
            match.start(),
            _periodo_comparacao(
                f"{ano}-{mes}",
                _nome_periodo_mes_ano(mes, ano),
                ano,
                [mes]
            )
        ))

    encontrados.sort(key=lambda x: x[0])
    vistos = set()
    periodos_mes_ano = []

    for _, periodo in encontrados:
        chave = (periodo["ano"], tuple(periodo["meses"]))
        if chave not in vistos:
            vistos.add(chave)
            periodos_mes_ano.append(periodo)

    if len(periodos_mes_ano) >= 2:
        return periodos_mes_ano[:2]

    # Meses citados na pergunta
    meses_pos = []
    for nome_mes, mes in _MESES_COMPARACAO.items():
        for match in re.finditer(rf"\b{re.escape(nome_mes)}\b", t):
            meses_pos.append((match.start(), nome_mes, mes))
    meses_pos.sort(key=lambda x: x[0])

    anos = re.findall(r"\b(20\d{2})\b", t)
    anos_unicos = list(dict.fromkeys(anos))

    # Dois meses e um único ano
    if len(meses_pos) >= 2 and len(anos_unicos) == 1:
        ano = anos_unicos[0]
        return [
            _periodo_comparacao(
                f"{ano}-{mes}",
                _nome_periodo_mes_ano(mes, ano),
                ano,
                [mes]
            )
            for _, _, mes in meses_pos[:2]
        ]

    # Dois meses sem ano -> ano atual
    if len(meses_pos) >= 2 and not anos_unicos:
        ano = str(agora.year)
        return [
            _periodo_comparacao(
                f"{ano}-{mes}",
                _nome_periodo_mes_ano(mes, ano),
                ano,
                [mes]
            )
            for _, _, mes in meses_pos[:2]
        ]

    # Um mês e dois anos
    if len(meses_pos) >= 1 and len(anos_unicos) >= 2:
        mes = meses_pos[0][2]
        return [
            _periodo_comparacao(
                f"{ano}-{mes}",
                _nome_periodo_mes_ano(mes, ano),
                ano,
                [mes]
            )
            for ano in anos_unicos[:2]
        ]

    # Dois anos inteiros
    if len(anos_unicos) >= 2:
        return [
            _periodo_comparacao(
                ano,
                ano,
                ano,
                []
            )
            for ano in anos_unicos[:2]
        ]

    return None



def _chave_ordenacao_periodo(periodo):
    """Chave cronológica para ordenar comparações por período."""
    try:
        ano = int(periodo.get("ano") or 0)
    except Exception:
        ano = 0

    meses = periodo.get("meses") or []

    if meses:
        try:
            mes_final = max(int(m) for m in meses)
        except Exception:
            mes_final = 12
    else:
        mes_final = 12

    return (ano, mes_final)


def _ordenar_periodos_mais_recente_primeiro(periodos):
    """
    Exibe comparações temporais sempre do período mais recente
    para o mais antigo.
    """
    return sorted(
        list(periodos or []),
        key=_chave_ordenacao_periodo,
        reverse=True
    )


def _extrair_filtros_base_comparacao(pergunta):
    """Captura filtros comerciais e remove ano/mês."""
    anterior = contexto_conversa.get("ultima_interpretacao")
    contexto_conversa["ultima_interpretacao"] = None

    try:
        for funcao in [
            interpretar_com_groq,
            interpretar_com_gemini,
            interpretar_com_claude,
        ]:
            try:
                interp = funcao(pergunta)
                filtros = interp.filtros.model_dump(exclude_none=True)
                filtros.pop("ano", None)
                filtros.pop("mes", None)
                return filtros
            except Exception:
                continue
    finally:
        contexto_conversa["ultima_interpretacao"] = anterior

    filtros = {}
    texto_upper = pergunta.upper()

    for regiao in sorted(regioes_conhecidas, key=len, reverse=True):
        if regiao in texto_upper:
            filtros["regiao"] = regiao
            break

    return filtros


def _filtros_para_periodo_comparacao(periodo, filtros_base=None):
    filtros = dict(filtros_base or {})
    filtros["ano"] = periodo["ano"]
    meses = periodo.get("meses") or []

    if len(meses) == 1:
        filtros["mes"] = normalizar_mes_powerbi(
            meses[0],
            periodo["ano"]
        )
    elif len(meses) > 1:
        filtros["mes"] = meses

    return filtros


def _diferenca_comparacao(v1, v2, formato):
    if v1 is None or v2 is None:
        return {
            "diferenca": None,
            "diferenca_formatada": "Sem resultado",
            "variacao_pct": None,
            "variacao_formatada": "-"
        }

    # A comparação segue a ordem pedida pelo usuário:
    # primeiro período menos segundo período.
    # Ex.: Julho/2026 x Julho/2025 => 2026 - 2025.
    diferenca = v1 - v2

    if formato == "percentual":
        diferenca_formatada = (
            f"{diferenca * 100:+.1f} p.p."
            .replace(".", ",")
        )
    else:
        diferenca_formatada = formatar_valor(
            diferenca,
            formato
        )
        if diferenca > 0:
            diferenca_formatada = "+" + diferenca_formatada

    variacao_pct = None
    variacao_formatada = "-"

    # A variação mede quanto o primeiro período mudou em relação
    # ao segundo período, que funciona como base da comparação.
    if v2 not in (0, None):
        variacao_pct = (diferenca / abs(v2)) * 100
        variacao_formatada = (
            f"{variacao_pct:+.1f}%"
            .replace(".", ",")
        )

    return {
        "diferenca": diferenca,
        "diferenca_formatada": diferenca_formatada,
        "variacao_pct": variacao_pct,
        "variacao_formatada": variacao_formatada
    }


def _indicadores_comparativo_gerencial():
    return [
        ("faturamento", "Faturamento"),
        ("meta_faturamento", "Meta de faturamento"),
        ("atingimento_meta_faturamento", "Atingimento da meta"),
        ("margem_liquida", "Margem líquida"),
        ("margem_bruta", "Margem bruta"),
    ]


def _pergunta_comparativo_gerencial(pergunta):
    t = _texto_normalizado(pergunta)
    return (
        "gerencial" in t
        and any(
            x in t
            for x in [
                "comparativo",
                "comparacao",
                "compare",
                "comparar"
            ]
        )
    )


def consultar_comparacao_periodos(pergunta):
    if not _tem_marcador_comparacao(pergunta):
        return None

    periodos = _extrair_periodos_comparacao(pergunta)
    if not periodos or len(periodos) < 2:
        return None

    periodos = _ordenar_periodos_mais_recente_primeiro(
        periodos[:2]
    )

    periodo_a, periodo_b = periodos
    filtros_base = _extrair_filtros_base_comparacao(pergunta)

    filtros_a = _filtros_para_periodo_comparacao(periodo_a, filtros_base)
    filtros_b = _filtros_para_periodo_comparacao(periodo_b, filtros_base)

    # Comparativo gerencial
    if _pergunta_comparativo_gerencial(pergunta):
        dados = []

        for indicador, descricao in _indicadores_comparativo_gerencial():
            config = mapa_indicadores[indicador]
            valor_a = consultar_valor(indicador, filtros_a)
            valor_b = consultar_valor(indicador, filtros_b)
            dif = _diferenca_comparacao(
                valor_a,
                valor_b,
                config["formato"]
            )

            dados.append({
                "indicador": indicador,
                "descricao": descricao,
                "formato": config["formato"],
                "valor_a": valor_a,
                "valor_b": valor_b,
                **dif
            })

        linhas = [
            f"📊 Comparativo gerencial: {periodo_a['nome']} x {periodo_b['nome']}",
            ""
        ]

        for item in dados:
            linhas.append(
                f"{item['descricao']}: "
                f"{formatar_valor(item['valor_a'], item['formato'])} → "
                f"{formatar_valor(item['valor_b'], item['formato'])} | "
                f"Dif.: {item['diferenca_formatada']}"
            )

        return {
            "tipo_resposta": "comparativo_gerencial_periodos",
            "periodo_a": periodo_a,
            "periodo_b": periodo_b,
            "filtros_base": filtros_base,
            "dados": dados,
            "resposta": "\n".join(linhas)
        }

    # Comparação de um indicador
    indicador = _indicador_da_pergunta(pergunta)
    config = mapa_indicadores[indicador]

    valor_a = consultar_valor(indicador, filtros_a)
    valor_b = consultar_valor(indicador, filtros_b)

    dif = _diferenca_comparacao(
        valor_a,
        valor_b,
        config["formato"]
    )

    resumo = ""

    if (
        valor_a is not None
        and valor_b is not None
        and dif.get("variacao_pct") is not None
    ):
        variacao_abs_txt = (
            f"{abs(dif['variacao_pct']):.1f}%"
            .replace(".", ",")
        )

        if valor_a < valor_b:
            resumo = (
                f"\n💡 Resumo: {config['descricao']} em "
                f"{periodo_a['nome']} foi {variacao_abs_txt} "
                f"menor que em {periodo_b['nome']}."
            )

        elif valor_a > valor_b:
            resumo = (
                f"\n💡 Resumo: {config['descricao']} em "
                f"{periodo_a['nome']} foi {variacao_abs_txt} "
                f"maior que em {periodo_b['nome']}."
            )

        else:
            resumo = (
                f"\n💡 Resumo: {config['descricao']} ficou "
                f"no mesmo nível nos dois períodos."
            )

    resposta = (
        f"📊 Comparativo de {config['descricao'].lower()}:\n\n"
        f"{periodo_a['nome']}: {formatar_valor(valor_a, config['formato'])}\n"
        f"{periodo_b['nome']}: {formatar_valor(valor_b, config['formato'])}\n"
        f"Diferença: {dif['diferenca_formatada']}\n"
        f"Variação: {dif['variacao_formatada']}"
        f"{resumo}"
    )

    return {
        "tipo_resposta": "comparacao_periodos",
        "indicador": indicador,
        "periodo_a": periodo_a,
        "periodo_b": periodo_b,
        "filtros_base": filtros_base,
        "dados": [
            {"item": periodo_a["nome"], "valor": valor_a},
            {"item": periodo_b["nome"], "valor": valor_b}
        ],
        "valor_a": valor_a,
        "valor_b": valor_b,
        **dif,
        "resposta": resposta
    }


def _imagem_tabela_comparacao_periodos(comparacao):
    periodo_a = comparacao["periodo_a"]["nome"]
    periodo_b = comparacao["periodo_b"]["nome"]

    if comparacao.get("tipo_resposta") == "comparativo_gerencial_periodos":
        colunas = [
            "Indicador",
            periodo_a,
            periodo_b,
            "Diferença",
            "Variação"
        ]

        linhas = []
        for item in comparacao["dados"]:
            linhas.append([
                item["descricao"],
                formatar_valor(item["valor_a"], item["formato"]),
                formatar_valor(item["valor_b"], item["formato"]),
                item["diferenca_formatada"],
                item["variacao_formatada"]
            ])

        return _tabela_png_base64(
            f"Comparativo gerencial - {periodo_a} x {periodo_b}",
            colunas,
            linhas,
            fonte=8
        )

    indicador = comparacao["indicador"]
    config = mapa_indicadores[indicador]

    colunas = [
        "Indicador",
        periodo_a,
        periodo_b,
        "Diferença",
        "Variação"
    ]

    linhas = [[
        config["descricao"],
        formatar_valor(comparacao.get("valor_a"), config["formato"]),
        formatar_valor(comparacao.get("valor_b"), config["formato"]),
        comparacao.get("diferenca_formatada", "-"),
        comparacao.get("variacao_formatada", "-")
    ]]

    return _tabela_png_base64(
        f"Comparativo de {config['descricao'].lower()} - {periodo_a} x {periodo_b}",
        colunas,
        linhas,
        fonte=8
    )


def consultar_comparacao_generica(pergunta):
    if not _tem_marcador_comparacao(pergunta):
        return None

    # Prioridade máxima para comparação entre dois períodos explícitos.
    comparacao_periodos = consultar_comparacao_periodos(
        pergunta
    )

    if comparacao_periodos is not None:
        return comparacao_periodos

    comparacao_temporal = consultar_comparacao_temporal_generica(
        pergunta
    )

    if comparacao_temporal is not None:
        return comparacao_temporal

    indicador = _indicador_da_pergunta(pergunta)

    regioes = _comparacao_regioes(pergunta)
    if regioes:
        filtros_base = _extrair_periodo_pergunta(pergunta)
        dados = []
        for nome, valores in regioes:
            filtros = dict(filtros_base)
            filtros["regiao"] = valores if len(valores) > 1 else valores[0]
            dados.append({"item": nome, "valor": consultar_valor(indicador, filtros)})
        return _montar_resposta_comparacao(indicador, dados, "região")

    meses = _comparacao_meses(pergunta)
    if meses:
        match_ano = re.search(r"\b(20\d{2})\b", pergunta)
        ano = match_ano.group(1) if match_ano else "Ano atual"
        dados = []
        for nome, numero in meses:
            filtros = {
                "ano": ano,
                "mes": normalizar_mes_powerbi(numero, ano)
            }
            dados.append({"item": nome, "valor": consultar_valor(indicador, filtros)})
        return _montar_resposta_comparacao(indicador, dados, "mês")

    anos = _comparacao_anos(pergunta)
    if anos:
        dados = []
        for nome, ano in anos:
            dados.append({"item": nome, "valor": consultar_valor(indicador, {"ano": ano})})
        return _montar_resposta_comparacao(indicador, dados, "ano")

    return None


# ============================================================
# 43. GRÁFICOS
# ============================================================

def _usuario_pediu_linha_tendencia(pergunta):
    t = _texto_normalizado(pergunta)

    return any(
        termo in t
        for termo in [
            "linha de tendencia",
            "linha tendencia",
            "tendencia linear",
            "trendline",
            "linha da tendencia"
        ]
    )


def usuario_pediu_grafico(pergunta):
    t = _texto_normalizado(pergunta)

    # "tabela" é tratada como saída visual.
    # Pedir linha de tendência também implica geração de gráfico.
    return (
        "grafico" in t
        or "chart" in t
        or _usuario_pediu_tabela(pergunta)
        or _usuario_pediu_linha_tendencia(pergunta)
    )


def _tipo_grafico_solicitado(pergunta):
    """
    Respeita o tipo de gráfico pedido explicitamente pelo usuário.

    Importante:
    "linha de tendência" NÃO transforma um gráfico de coluna
    em gráfico de linha.
    """
    t = _texto_normalizado(pergunta)

    # Remove apenas a expressão de tendência antes de detectar o tipo.
    t_tipo = t
    for termo in [
        "linha de tendencia",
        "linha tendencia",
        "tendencia linear",
        "trendline",
        "linha da tendencia"
    ]:
        t_tipo = t_tipo.replace(termo, " ")

    if "pizza" in t_tipo or "participacao" in t_tipo:
        return "pizza"

    if "coluna" in t_tipo or "colunas" in t_tipo:
        return "barras"

    if "barra" in t_tipo or "barras" in t_tipo:
        return "barras_horizontais"

    if "linha" in t_tipo or "linhas" in t_tipo:
        return "linha"

    return None


def _grafico_png_base64(
    tipo,
    titulo,
    rotulos,
    valores,
    eixo_y="Valor",
    formato="moeda",
    adicionar_tendencia=False,
    cor_serie="tab:blue",
    cor_tendencia="dimgray"
):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import math

    fig, ax = plt.subplots(figsize=(10, 6))

    def _arredondar_inteiro(valor):
        """
        Arredonda para o inteiro mais próximo.
        Ex.: 10,49 -> 10 | 10,50 -> 11 | -10,50 -> -11
        """
        valor = float(valor)

        if valor >= 0:
            return int(math.floor(valor + 0.5))

        return int(math.ceil(valor - 0.5))

    def _formatar_inteiro_ptbr(valor):
        return (
            f"{_arredondar_inteiro(valor):,}"
            .replace(",", ".")
        )

    def _formatar_rotulo_grafico(valor):
        if valor is None:
            return "Sem resultado"

        if formato == "percentual":
            percentual = float(valor) * 100
            return f"{_formatar_inteiro_ptbr(percentual)}%"

        if formato == "inteiro":
            return _formatar_inteiro_ptbr(valor)

        # Moeda: valor completo, sem casas decimais.
        return f"R$ {_formatar_inteiro_ptbr(valor)}"

    valores_plot = [
        0 if valor is None else float(valor)
        for valor in valores
    ]

    # ========================================================
    # GRÁFICO DE LINHA
    # ========================================================
    if tipo == "linha":
        ax.plot(
            rotulos,
            valores_plot,
            marker="o",
            color=cor_serie
        )

        ax.set_xlabel("Período")
        ax.set_ylabel(eixo_y)

        # Remove notação científica do eixo.
        ax.ticklabel_format(
            style="plain",
            axis="y",
            useOffset=False
        )

        # Rótulo em TODOS os pontos.
        maior = max(valores_plot) if valores_plot else 0
        deslocamento = maior * 0.025 if maior else 1

        for x, y in zip(
            rotulos,
            valores_plot
        ):
            ax.annotate(
                _formatar_rotulo_grafico(y),
                xy=(x, y),
                xytext=(0, 8),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=8
            )

        # Reserva espaço para os rótulos.
        if valores_plot:
            maior = max(valores_plot)
            menor = min(valores_plot)

            if maior > 0:
                ax.set_ylim(
                    top=maior * 1.14
                )

            if menor < 0:
                topo_atual = ax.get_ylim()[1]
                ax.set_ylim(
                    bottom=menor * 1.14,
                    top=topo_atual
                )

    # ========================================================
    # GRÁFICO DE PIZZA
    # ========================================================
    elif tipo == "pizza":
        total = sum(
            abs(v)
            for v in valores_plot
        )

        def _rotulo_pizza(pct):
            if total == 0:
                return ""

            valor = (
                pct / 100
            ) * total

            return (
                f"{_formatar_rotulo_grafico(valor)}\n"
                f"{_arredondar_inteiro(pct)}%"
            )

        ax.pie(
            valores_plot,
            labels=rotulos,
            autopct=_rotulo_pizza,
            textprops={
                "fontsize": 8
            }
        )

        ax.axis("equal")

    # ========================================================
    # BARRAS HORIZONTAIS
    # ========================================================
    elif tipo == "barras_horizontais":
        bars = ax.barh(
            rotulos,
            valores_plot,
            color=cor_serie
        )

        ax.set_xlabel(eixo_y)
        ax.invert_yaxis()

        ax.ticklabel_format(
            style="plain",
            axis="x",
            useOffset=False
        )

        labels = [
            _formatar_rotulo_grafico(v)
            for v in valores_plot
        ]

        ax.bar_label(
            bars,
            labels=labels,
            padding=3,
            fontsize=8
        )

        if valores_plot:
            maior = max(valores_plot)
            menor = min(valores_plot)

            if maior > 0:
                ax.set_xlim(
                    right=maior * 1.30
                )

            if menor < 0:
                direita_atual = ax.get_xlim()[1]

                ax.set_xlim(
                    left=menor * 1.30,
                    right=direita_atual
                )

    # ========================================================
    # BARRAS VERTICAIS
    # ========================================================
    else:
        bars = ax.bar(
            rotulos,
            valores_plot,
            color=cor_serie
        )

        ax.set_ylabel(eixo_y)

        ax.tick_params(
            axis="x",
            rotation=35
        )

        ax.ticklabel_format(
            style="plain",
            axis="y",
            useOffset=False
        )

        labels = [
            _formatar_rotulo_grafico(v)
            for v in valores_plot
        ]

        ax.bar_label(
            bars,
            labels=labels,
            padding=3,
            fontsize=8
        )

        if valores_plot:
            maior = max(valores_plot)
            menor = min(valores_plot)

            if maior > 0:
                ax.set_ylim(
                    top=maior * 1.18
                )

            if menor < 0:
                topo_atual = ax.get_ylim()[1]

                ax.set_ylim(
                    bottom=menor * 1.18,
                    top=topo_atual
                )

    # ========================================================
    # LINHA DE TENDÊNCIA LINEAR
    # ========================================================
    if (
        adicionar_tendencia
        and tipo != "pizza"
        and len(valores_plot) >= 2
    ):
        try:
            import numpy as np

            x_num = np.arange(
                len(valores_plot),
                dtype=float
            )

            y_num = np.array(
                valores_plot,
                dtype=float
            )

            coef = np.polyfit(
                x_num,
                y_num,
                1
            )

            tendencia = np.polyval(
                coef,
                x_num
            )

            if tipo == "barras_horizontais":
                # x = valor / y = posição da categoria
                ax.plot(
                    tendencia,
                    x_num,
                    linestyle="--",
                    linewidth=2,
                    label="Tendência",
                    color=cor_tendencia
                )
            else:
                ax.plot(
                    rotulos,
                    tendencia,
                    linestyle="--",
                    linewidth=2,
                    label="Tendência",
                    color=cor_tendencia
                )

            ax.legend()

        except Exception as erro:
            print(
                "Aviso: não foi possível adicionar linha de tendência:",
                type(erro).__name__,
                "-",
                str(erro)[:200]
            )

    ax.set_title(
        titulo
    )

    if tipo == "barras_horizontais":
        ax.grid(
            axis="x",
            alpha=0.25
        )

    elif tipo != "pizza":
        ax.grid(
            axis="y",
            alpha=0.25
        )

    fig.tight_layout()

    buffer = io.BytesIO()

    fig.savefig(
        buffer,
        format="png",
        dpi=140,
        bbox_inches="tight"
    )

    plt.close(fig)

    return base64.b64encode(
        buffer.getvalue()
    ).decode("ascii")

def _tabela_png_base64(
    titulo,
    colunas,
    linhas,
    larguras=None,
    fonte=9
):
    """
    Gera uma tabela como imagem PNG em base64
    com visual executivo e destaque de variações.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import io
    import base64

    qtd_linhas = max(len(linhas), 1)
    qtd_colunas = max(len(colunas), 1)

    altura = max(
        2.6,
        1.55 + (qtd_linhas * 0.52)
    )

    largura = max(
        10.0,
        2.15 * qtd_colunas
    )

    fig, ax = plt.subplots(
        figsize=(largura, altura)
    )

    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.axis("off")

    tabela = ax.table(
        cellText=linhas,
        colLabels=colunas,
        loc="center",
        cellLoc="center",
        colLoc="center"
    )

    tabela.auto_set_font_size(False)
    tabela.set_fontsize(fonte)
    tabela.scale(1, 1.55)

    cor_cabecalho = "#4F46E5"
    cor_linha_par = "#F7F7FC"
    cor_linha_impar = "#FFFFFF"
    cor_borda = "#D7D9E5"
    cor_negativo = "#D62728"
    cor_positivo = "#14833B"

    for (linha, coluna), celula in tabela.get_celld().items():

        celula.set_edgecolor(
            cor_borda
        )

        celula.set_linewidth(
            0.8
        )

        texto = celula.get_text()
        texto.set_wrap(True)

        if linha == 0:

            celula.set_facecolor(
                cor_cabecalho
            )

            texto.set_color(
                "white"
            )

            texto.set_weight(
                "bold"
            )

            texto.set_fontsize(
                fonte + 1
            )

        else:

            celula.set_facecolor(
                cor_linha_par
                if linha % 2 == 0
                else cor_linha_impar
            )

            if coluna == 0:
                texto.set_weight(
                    "bold"
                )

            valor_texto = str(
                texto.get_text()
            ).strip()

            cabecalho = (
                str(colunas[coluna]).lower()
                if coluna < len(colunas)
                else ""
            )

            if (
                "diferen" in cabecalho
                or "varia" in cabecalho
            ):

                if valor_texto.startswith("-"):

                    texto.set_color(
                        cor_negativo
                    )

                    texto.set_weight(
                        "bold"
                    )

                elif valor_texto.startswith("+"):

                    texto.set_color(
                        cor_positivo
                    )

                    texto.set_weight(
                        "bold"
                    )

    if larguras:

        for coluna, largura_coluna in enumerate(
            larguras
        ):

            for linha in range(
                0,
                qtd_linhas + 1
            ):

                chave = (
                    linha,
                    coluna
                )

                if chave in tabela.get_celld():

                    tabela.get_celld()[
                        chave
                    ].set_width(
                        largura_coluna
                    )

    ax.set_title(
        titulo,
        fontsize=14,
        fontweight="bold",
        pad=16
    )

    fig.tight_layout(
        pad=1.2
    )

    buffer = io.BytesIO()

    fig.savefig(
        buffer,
        format="png",
        dpi=150,
        bbox_inches="tight",
        facecolor="white"
    )

    plt.close(
        fig
    )

    return base64.b64encode(
        buffer.getvalue()
    ).decode(
        "ascii"
    )


def _formatar_valor_tabela(valor, formato):
    if valor is None:
        return "Sem resultado"

    return formatar_valor(
        valor,
        formato
    )


def _imagem_tabela_comparacao_temporal(
    comparacao,
    config
):
    entidade_a = comparacao["entidade_a"]
    entidade_b = comparacao["entidade_b"]

    colunas = [
        "Período",
        entidade_a,
        entidade_b,
        "Diferença",
        "Dif. %",
        "Melhor"
    ]

    linhas = []

    for item in comparacao["dados"]:
        va = item.get("valor_a")
        vb = item.get("valor_b")

        if va is None or vb is None:
            diferenca = "Sem resultado"
            diferenca_pct = "-"
            melhor = "-"
        else:
            dif_abs = abs(
                va - vb
            )

            diferenca = _formatar_valor_tabela(
                dif_abs,
                config["formato"]
            )

            menor = min(
                abs(va),
                abs(vb)
            )

            if menor > 0:
                diferenca_pct = (
                    f"{(dif_abs / menor) * 100:.1f}%"
                    .replace(".", ",")
                )
            else:
                diferenca_pct = "-"

            if va > vb:
                melhor = entidade_a
            elif vb > va:
                melhor = entidade_b
            else:
                melhor = "Empate"

        linhas.append([
            item.get(
                "periodo_nome",
                item.get("periodo", "")
            ),
            _formatar_valor_tabela(
                va,
                config["formato"]
            ),
            _formatar_valor_tabela(
                vb,
                config["formato"]
            ),
            diferenca,
            diferenca_pct,
            melhor
        ])

    titulo = (
        f"Comparativo de {config['descricao'].lower()} - "
        f"{entidade_a} x {entidade_b}"
    )

    return _tabela_png_base64(
        titulo,
        colunas,
        linhas,
        fonte=8
    )


def _imagem_tabela_comparacao_simples(
    comparacao,
    config
):
    colunas = [
        "Item",
        config["descricao"]
    ]

    linhas = [
        [
            item["item"],
            _formatar_valor_tabela(
                item.get("valor"),
                config["formato"]
            )
        ]
        for item in comparacao["dados"]
    ]

    return _tabela_png_base64(
        f"Comparativo de {config['descricao'].lower()}",
        colunas,
        linhas
    )


def _imagem_tabela_ranking(
    titulo,
    ranking,
    config
):
    colunas = [
        "Posição",
        "Item",
        config["descricao"]
    ]

    linhas = []

    for posicao, item in enumerate(
        ranking,
        start=1
    ):
        linhas.append([
            str(posicao),
            item["item"],
            _formatar_valor_tabela(
                item.get("valor"),
                config["formato"]
            )
        ])

    return _tabela_png_base64(
        titulo,
        colunas,
        linhas
    )


def _imagem_tabela_serie_temporal(
    titulo,
    rotulos,
    valores,
    config
):
    colunas = [
        "Período",
        config["descricao"]
    ]

    linhas = [
        [
            rotulo,
            _formatar_valor_tabela(
                valor,
                config["formato"]
            )
        ]
        for rotulo, valor in zip(
            rotulos,
            valores
        )
    ]

    return _tabela_png_base64(
        titulo,
        colunas,
        linhas
    )



# ============================================================
# 43B. GRÁFICO COMBINADO + LINHA DE TENDÊNCIA
# ============================================================

def _usuario_pediu_grafico_combinado(pergunta):
    t = _texto_normalizado(pergunta)

    tem_faturamento = any(
        x in t
        for x in [
            "faturamento",
            "faturou",
            "realizado"
        ]
    )

    tem_meta = (
        "meta" in t
    )

    tem_atingimento = any(
        x in t
        for x in [
            "atingimento",
            "% da meta",
            "percentual da meta",
            "percentual de atingimento"
        ]
    )

    # Pode escrever "gráfico combinado" ou simplesmente pedir
    # faturamento + meta + atingimento na mesma visualização.
    return (
        (
            tem_faturamento
            and tem_meta
            and tem_atingimento
        )
        or (
            "grafico combinado" in t
            and tem_faturamento
            and tem_meta
        )
    )


def _filtros_base_grafico(pergunta):
    """
    Extrai filtros comerciais da pergunta, mas deixa ano/mês
    sob controle da série temporal quando necessário.
    """
    anterior = contexto_conversa.get(
        "ultima_interpretacao"
    )

    contexto_conversa[
        "ultima_interpretacao"
    ] = None

    try:
        for funcao in [
            interpretar_com_groq,
            interpretar_com_gemini,
            interpretar_com_claude,
        ]:
            try:
                interp = funcao(
                    pergunta
                )

                filtros = (
                    interp
                    .filtros
                    .model_dump(
                        exclude_none=True
                    )
                )

                return filtros

            except Exception:
                continue

    finally:
        contexto_conversa[
            "ultima_interpretacao"
        ] = anterior

    return {}



# ============================================================
# CORES SOLICITADAS PELO USUÁRIO
# ============================================================

_CORES_GRAFICO_PT = {
    "azul": "tab:blue",
    "azul escuro": "navy",
    "azul claro": "lightskyblue",
    "laranja": "tab:orange",
    "verde": "tab:green",
    "verde escuro": "darkgreen",
    "verde claro": "lightgreen",
    "vermelho": "tab:red",
    "roxo": "tab:purple",
    "violeta": "violet",
    "rosa": "hotpink",
    "amarelo": "gold",
    "cinza": "tab:gray",
    "cinza escuro": "dimgray",
    "preto": "black",
    "branco": "white",
    "marrom": "tab:brown",
    "ciano": "tab:cyan",
    "turquesa": "turquoise",
    "magenta": "magenta",
    "dourado": "goldenrod",
}


def _normalizar_cor_grafico(valor):
    if not valor:
        return None

    cor = _texto_normalizado(
        str(valor)
    ).strip()

    # Hexadecimal informado pelo usuário.
    if re.fullmatch(
        r"#[0-9a-f]{6}",
        cor,
        flags=re.IGNORECASE
    ):
        return cor

    # Cores compostas primeiro.
    for nome in sorted(
        _CORES_GRAFICO_PT,
        key=len,
        reverse=True
    ):
        if nome == cor:
            return _CORES_GRAFICO_PT[
                nome
            ]

    return None


def _extrair_cor_apos_termos(
    pergunta,
    termos
):
    """
    Aceita frases como:
    - faturamento azul
    - faturamento em azul
    - barra do faturamento azul
    - meta verde
    - linha vermelha
    - atingimento roxo
    - cor #3366CC
    """
    t = _texto_normalizado(
        pergunta
    )

    nomes_cores = sorted(
        _CORES_GRAFICO_PT.keys(),
        key=len,
        reverse=True
    )

    padrao_cor = (
        r"(#[0-9a-f]{6}|"
        + "|".join(
            re.escape(cor)
            for cor in nomes_cores
        )
        + r")"
    )

    for termo in termos:
        termo_n = _texto_normalizado(
            termo
        )

        padroes = [
            rf"\b{re.escape(termo_n)}\b"
            rf"(?:\s+(?:em|na|no|de|com|cor))?"
            rf"\s+(?:cor\s+)?{padrao_cor}\b",

            rf"\bcor\s+(?:do|da|de)?\s*"
            rf"{re.escape(termo_n)}\s*[:=]?\s*"
            rf"{padrao_cor}\b",

            rf"\b{re.escape(termo_n)}\s*[:=]\s*"
            rf"{padrao_cor}\b",
        ]

        for padrao in padroes:
            match = re.search(
                padrao,
                t,
                flags=re.IGNORECASE
            )

            if match:
                cor_encontrada = (
                    match.group(
                        match.lastindex
                    )
                )

                return _normalizar_cor_grafico(
                    cor_encontrada
                )

    return None


def _cores_grafico_simples(pergunta):
    """
    Cores para gráficos simples (barras/colunas/linhas).

    Exemplos:
    - "gráfico de barras verdes"
    - "gráfico de linha vermelha"
    - "colunas azuis com linha de tendência preta"
    - "faturamento mês a mês em roxo"

    Defaults:
    - série principal: azul
    - tendência: cinza escuro
    """
    cor_serie = _extrair_cor_apos_termos(
        pergunta,
        [
            "barra", "barras", "coluna", "colunas",
            "linha", "linhas", "grafico", "faturamento",
            "meta", "mrgl", "atingimento"
        ]
    )

    # Também entende construções genéricas como "em verde" ou "cor verde".
    if cor_serie is None:
        t = _texto_normalizado(pergunta)
        nomes = sorted(_CORES_GRAFICO_PT.keys(), key=len, reverse=True)
        padrao = r"(#[0-9a-f]{6}|" + "|".join(re.escape(c) for c in nomes) + r")"
        m = re.search(rf"\b(?:em|cor)\s+{padrao}\b", t, flags=re.IGNORECASE)
        if m:
            cor_serie = _normalizar_cor_grafico(m.group(1))

    cor_tendencia = _extrair_cor_apos_termos(
        pergunta,
        ["linha de tendencia", "tendencia"]
    )

    return {
        "serie": cor_serie or "tab:blue",
        "tendencia": cor_tendencia or "dimgray"
    }


def _cores_grafico_combinado(
    pergunta
):
    """
    Define as cores do combinado.

    Exemplos entendidos:
    "faturamento azul, meta verde e atingimento vermelho"
    "barra de faturamento #3366CC, barra de meta #FF9900 e linha roxa"
    """
    cor_faturamento = _extrair_cor_apos_termos(
        pergunta,
        [
            "barra do faturamento",
            "barra de faturamento",
            "faturamento"
        ]
    )

    cor_meta = _extrair_cor_apos_termos(
        pergunta,
        [
            "barra da meta",
            "barra de meta",
            "meta"
        ]
    )

    cor_atingimento = _extrair_cor_apos_termos(
        pergunta,
        [
            "linha do atingimento",
            "linha de atingimento",
            "percentual de atingimento",
            "% de atingimento",
            "atingimento"
        ]
    )

    # "linha azul" sem especificar indicador, em gráfico combinado,
    # é interpretada como a linha do % de atingimento.
    if cor_atingimento is None:
        cor_atingimento = _extrair_cor_apos_termos(
            pergunta,
            ["linha"]
        )

    cor_tendencia = _extrair_cor_apos_termos(
        pergunta,
        [
            "linha de tendencia",
            "tendencia"
        ]
    )

    return {
        "faturamento": (
            cor_faturamento
            or "tab:blue"
        ),
        "meta": (
            cor_meta
            or "tab:orange"
        ),
        "atingimento": (
            cor_atingimento
            or "#D95F02"
        ),
        "tendencia": (
            cor_tendencia
            or "dimgray"
        )
    }


def _grafico_combinado_png_base64(
    titulo,
    rotulos,
    faturamento,
    meta,
    atingimento,
    adicionar_tendencia=False,
    tendencia_indicador="faturamento",
    cores=None
):
    """
    Barras agrupadas:
    - Faturamento
    - Meta

    Linha em eixo secundário:
    - % de atingimento da meta
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np
    import io
    import base64

    cores = cores or {
        "faturamento": "tab:blue",
        "meta": "tab:orange",
        "atingimento": "#D8640A",
        "tendencia": "dimgray"
    }

    x = np.arange(
        len(rotulos),
        dtype=float
    )

    # Um pouco mais estreitas para dar respiro ao gráfico.
    largura = 0.32

    fat = np.array(
        [
            0 if v is None else float(v)
            for v in faturamento
        ],
        dtype=float
    )

    met = np.array(
        [
            0 if v is None else float(v)
            for v in meta
        ],
        dtype=float
    )

    atg = np.array(
        [
            0 if v is None else float(v) * 100
            for v in atingimento
        ],
        dtype=float
    )

    fig, ax1 = plt.subplots(
        figsize=(12, 7)
    )

    barras_fat = ax1.bar(
        x - largura / 2,
        fat,
        largura,
        label="Faturamento",
        color=cores["faturamento"],
        zorder=2
    )

    barras_meta = ax1.bar(
        x + largura / 2,
        met,
        largura,
        label="Meta",
        color=cores["meta"],
        zorder=2
    )

    ax1.set_ylabel("Valor (R$)")
    ax1.set_xticks(x)
    ax1.set_xticklabels(
        rotulos,
        rotation=0
    )
    ax1.ticklabel_format(
        style="plain",
        axis="y",
        useOffset=False
    )
    ax1.grid(
        axis="y",
        alpha=0.18,
        zorder=0
    )

    # Folga no topo para rótulos e legenda.
    maior_valor = max(
        np.max(fat)
        if len(fat)
        else 0,
        np.max(met)
        if len(met)
        else 0,
        1
    )

    ax1.set_ylim(
        0,
        maior_valor * 1.18
    )

    ax2 = ax1.twinx()

    linha_atg, = ax2.plot(
        x,
        atg,
        marker="o",
        markersize=5.5,
        linewidth=2.4,
        label="% Atingimento",
        color=cores["atingimento"],
        zorder=5
    )

    ax2.set_ylabel(
        "% Atingimento da meta"
    )

    # Evita que pontos/rótulos da linha encostem na legenda/título.
    max_atg = (
        float(np.max(atg))
        if len(atg)
        else 100.0
    )

    limite_superior_atg = max(
        110.0,
        max_atg * 1.16
    )

    ax2.set_ylim(
        0,
        limite_superior_atg
    )

    # Os rótulos do % de atingimento são posicionados depois
    # dos rótulos das barras, para que o algoritmo possa evitar
    # colisões usando a posição real dos textos desenhados.
    # Rótulos resumidos nas barras.
    def _rotulo_moeda_curto(valor):
        valor = float(valor)

        if abs(valor) >= 1_000_000:
            return (
                f"R$ {valor / 1_000_000:.1f} mi"
                .replace(".", ",")
            )

        if abs(valor) >= 1_000:
            return (
                f"R$ {valor / 1_000:.0f} mil"
                .replace(".", ",")
            )

        return (
            f"R$ {valor:,.0f}"
            .replace(",", ".")
        )

    rotulos_barras_artistas = []

    for barras in [
        barras_fat,
        barras_meta
    ]:
        for barra in barras:
            altura = barra.get_height()

            artista_rotulo = ax1.annotate(
                _rotulo_moeda_curto(
                    altura
                ),
                xy=(
                    barra.get_x()
                    + barra.get_width() / 2,
                    altura
                ),
                xytext=(0, 4),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=7,
                rotation=90
                    if len(rotulos) > 8
                    else 0
            )

            rotulos_barras_artistas.append(
                artista_rotulo
            )

    # ========================================================
    # POSICIONAMENTO AUTOMÁTICO DOS RÓTULOS DA LINHA
    # ========================================================
    # O comportamento aqui é semelhante ao de ferramentas de BI:
    # testamos várias posições próximas ao ponto e escolhemos a
    # primeira que não colide com rótulos das barras, outros
    # percentuais ou as bordas do gráfico.
    fig.canvas.draw()

    renderer = fig.canvas.get_renderer()

    caixas_ocupadas = []

    for artista in rotulos_barras_artistas:
        try:
            caixas_ocupadas.append(
                artista.get_window_extent(
                    renderer=renderer
                ).expanded(
                    1.08,
                    1.18
                )
            )
        except Exception:
            pass

    # Ordem de preferência: perto do marcador.
    # Evitamos deslocamentos grandes como os usados na V14.
    candidatos_padrao = [
        (0, 8, "center", "bottom"),
        (0, -9, "center", "top"),
        (9, 6, "left", "bottom"),
        (-9, 6, "right", "bottom"),
        (9, -5, "left", "top"),
        (-9, -5, "right", "top"),
        (12, 0, "left", "center"),
        (-12, 0, "right", "center"),
    ]

    # Para o último ponto preferimos posições para a direita somente
    # se houver espaço; caso contrário o algoritmo automaticamente
    # encontrará outra alternativa.
    rotulos_linha_artistas = []

    for idx, (xi, valor) in enumerate(
        zip(
            x,
            atg
        )
    ):
        texto_pct = (
            f"{valor:.0f}%"
            .replace(".", ",")
        )

        melhor_artista = None
        melhor_caixa = None
        menor_penalidade = None

        for (
            desloc_x,
            desloc_y,
            ha,
            va
        ) in candidatos_padrao:

            artista_teste = ax2.annotate(
                texto_pct,
                xy=(xi, valor),
                xytext=(
                    desloc_x,
                    desloc_y
                ),
                textcoords="offset points",
                ha=ha,
                va=va,
                fontsize=8,
                fontweight="bold",
                color=cores["atingimento"],
                zorder=7
            )

            fig.canvas.draw()

            try:
                caixa = (
                    artista_teste
                    .get_window_extent(
                        renderer=renderer
                    )
                    .expanded(
                        1.06,
                        1.12
                    )
                )

                caixa_eixos = (
                    ax2
                    .get_window_extent(
                        renderer=renderer
                    )
                )

                # Penalidade por sair da área útil.
                fora = (
                    caixa.x0 < caixa_eixos.x0
                    or caixa.x1 > caixa_eixos.x1
                    or caixa.y0 < caixa_eixos.y0
                    or caixa.y1 > caixa_eixos.y1
                )

                colisoes = sum(
                    1
                    for ocupada in caixas_ocupadas
                    if caixa.overlaps(
                        ocupada
                    )
                )

                # Mantém o rótulo perto do ponto sempre que possível.
                distancia = (
                    abs(desloc_x)
                    + abs(desloc_y)
                )

                penalidade = (
                    colisoes * 1000
                    + (500 if fora else 0)
                    + distancia
                )

            except Exception:
                penalidade = 999999
                caixa = None

            if (
                menor_penalidade is None
                or penalidade < menor_penalidade
            ):
                if melhor_artista is not None:
                    melhor_artista.remove()

                melhor_artista = artista_teste
                melhor_caixa = caixa
                menor_penalidade = penalidade
            else:
                artista_teste.remove()

            # Posição sem colisão e dentro do gráfico:
            # não há motivo para procurar uma mais distante.
            if penalidade < 500:
                break

        if melhor_artista is not None:
            rotulos_linha_artistas.append(
                melhor_artista
            )

            if melhor_caixa is not None:
                caixas_ocupadas.append(
                    melhor_caixa
                )

    # --------------------------------------------------------
    # Tendência opcional
    # --------------------------------------------------------
    if (
        adicionar_tendencia
        and len(rotulos) >= 2
    ):
        alvo = tendencia_indicador

        try:
            if alvo == "meta":
                valores_tend = met
                eixo = ax1
                rotulo_tend = "Tendência da meta"

            elif alvo == "atingimento":
                valores_tend = atg
                eixo = ax2
                rotulo_tend = "Tendência do atingimento"

            else:
                valores_tend = fat
                eixo = ax1
                rotulo_tend = "Tendência do faturamento"

            coef = np.polyfit(
                x,
                valores_tend,
                1
            )

            y_tend = np.polyval(
                coef,
                x
            )

            eixo.plot(
                x,
                y_tend,
                linestyle="--",
                linewidth=2,
                color=cores["tendencia"],
                label=rotulo_tend,
                zorder=6
            )

        except Exception as erro:
            print(
                "Aviso: tendência do gráfico combinado:",
                type(erro).__name__,
                "-",
                str(erro)[:200]
            )

    # Legenda combinando os dois eixos.
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()

    ax1.legend(
        h1 + h2,
        l1 + l2,
        loc="upper center",
        bbox_to_anchor=(0.5, 1.075),
        ncol=min(
            4,
            max(
                1,
                len(l1 + l2)
            )
        ),
        frameon=False
    )

    ax1.set_title(
        titulo,
        pad=38
    )

    # Margens reservadas para título + legenda.
    fig.subplots_adjust(
        top=0.84,
        bottom=0.12,
        left=0.08,
        right=0.90
    )

    buffer = io.BytesIO()

    fig.savefig(
        buffer,
        format="png",
        dpi=150,
        bbox_inches="tight"
    )

    plt.close(fig)

    return base64.b64encode(
        buffer.getvalue()
    ).decode("utf-8")


def _alvo_tendencia_combinado(pergunta):
    t = _texto_normalizado(pergunta)

    if any(
        x in t
        for x in [
            "tendencia da meta",
            "tendencia de meta"
        ]
    ):
        return "meta"

    if any(
        x in t
        for x in [
            "tendencia do atingimento",
            "tendencia de atingimento",
            "tendencia do percentual",
            "tendencia do %"
        ]
    ):
        return "atingimento"

    # Padrão: faturamento.
    return "faturamento"



def montar_dax_series_mensais_combinadas(
    ano,
    filtros_base=None
):
    """
    Retorna faturamento, meta e atingimento de todos os meses
    do ano em UMA única chamada ao Power BI.
    """

    filtros = dict(
        filtros_base
        or {}
    )

    filtros["ano"] = str(
        ano
    )

    # A série mensal precisa enxergar todos os meses.
    filtros.pop(
        "mes",
        None
    )

    contexto = montar_contexto_final(
        contexto_overview_comercial,
        filtros
    )

    # Remove o default "Mês atual", pois o mês será a dimensão
    # agrupadora da consulta.
    contexto.pop(
        "mes",
        None
    )

    filtros_dax = gerar_filtros_dax(
        contexto
    )

    filtros_texto = ",\n        ".join(
        filtros_dax
    )

    return f"""
EVALUATE

CALCULATETABLE(
    SUMMARIZECOLUMNS(
        '# CALENDÁRIO'[mes_atual],

        "Faturamento",
        [fat_si],

        "Meta",
        [fat_si_meta],

        "Atingimento",
        [atg_meta_fat_si]
    ),

    {filtros_texto}
)
"""


def consultar_series_mensais_combinadas(
    ano,
    filtros_base=None
):
    """
    Executa uma única DAX e devolve as três métricas por mês.
    """

    _t_inicio = time.perf_counter()

    dax = montar_dax_series_mensais_combinadas(
        ano,
        filtros_base
    )

    linhas = extrair_linhas(
        executar_dax(
            dax
        )
    )

    resultado = {}

    for linha in linhas:

        mes_bruto = None

        # O Power BI REST pode devolver a coluna agrupadora
        # com variações no nome da chave. Procuramos de forma
        # tolerante.
        for chave, valor in linha.items():

            chave_norm = (
                str(chave)
                .lower()
                .replace("'", "")
                .replace("[", "")
                .replace("]", "")
                .replace(" ", "")
            )

            if (
                "calendário" in chave_norm
                or "calendario" in chave_norm
            ) and "mes_atual" in chave_norm:

                mes_bruto = valor
                break

        if mes_bruto is None:

            for chave, valor in linha.items():

                if "mes_atual" in str(chave).lower():

                    mes_bruto = valor
                    break

        if mes_bruto is None:

            continue

        mes_txt = str(
            mes_bruto
        ).strip()

        # No modelo, o mês corrente pode vir como "Mês atual".
        if mes_txt.lower() in {
            "mês atual",
            "mes atual"
        }:

            mes_num = (
                f"{datetime.now().month:02d}"
            )

        else:

            # Aceita "01", 1, "1", etc.
            match_mes = re.search(
                r"\b(0?[1-9]|1[0-2])\b",
                mes_txt
            )

            if not match_mes:

                continue

            mes_num = str(
                int(
                    match_mes.group(1)
                )
            ).zfill(2)

        resultado[mes_num] = {
            "faturamento":
                linha.get(
                    "[Faturamento]",
                    0
                )
                or 0,

            "meta":
                linha.get(
                    "[Meta]",
                    0
                )
                or 0,

            "atingimento":
                linha.get(
                    "[Atingimento]",
                    0
                )
                or 0
        }

    print(
        f"[TEMPO SÉRIE MENSAL - 1 DAX] "
        f"{time.perf_counter() - _t_inicio:.3f}s "
        f"| meses retornados={len(resultado)}"
    )

    return resultado


def _series_combinadas_mes_a_mes(
    pergunta
):
    _t_serie_inicio = time.perf_counter()

    match_ano = re.search(
        r"\b(20\d{2})\b",
        pergunta
    )

    ano = (
        match_ano.group(1)
        if match_ano
        else str(datetime.now().year)
    )

    limite = (
        datetime.now().month
        if int(ano)
        == datetime.now().year
        else 12
    )

    filtros_base = (
        _filtros_base_grafico(
            pergunta
        )
    )

    # Ano e mês são controlados pela própria série.
    filtros_base.pop(
        "ano",
        None
    )

    filtros_base.pop(
        "mes",
        None
    )

    # ========================================================
    # OTIMIZAÇÃO:
    # antes: 3 chamadas por mês (24 chamadas em jan-ago)
    # agora: 1 única chamada DAX para todos os meses/métricas
    # ========================================================

    dados_mensais = (
        consultar_series_mensais_combinadas(
            ano,
            filtros_base
        )
    )

    rotulos = []
    fat = []
    meta = []
    ating = []

    for n in range(
        1,
        limite + 1
    ):

        mes = f"{n:02d}"

        dados_mes = (
            dados_mensais.get(
                mes,
                {}
            )
        )

        rotulos.append(
            nome_mes_resposta(
                mes
            ).capitalize()
        )

        fat.append(
            dados_mes.get(
                "faturamento",
                0
            )
            or 0
        )

        meta.append(
            dados_mes.get(
                "meta",
                0
            )
            or 0
        )

        ating.append(
            dados_mes.get(
                "atingimento",
                0
            )
            or 0
        )

    print(
        f"\n[TEMPO SÉRIE MENSAL OTIMIZADA] "
        f"{time.perf_counter() - _t_serie_inicio:.3f}s"
    )

    return {
        "ano": ano,
        "rotulos": rotulos,
        "faturamento": fat,
        "meta": meta,
        "atingimento": ating,
        "filtros_base": filtros_base
    }


def _series_combinadas_dimensao(
    pergunta,
    dimensao
):
    """
    Monta séries combinadas para uma dimensão:
    região, cliente, produto, representante, etc.
    """
    t = _texto_normalizado(
        pergunta
    )

    match_top = re.search(
        r"\b(?:top\s*)?(\d+)\b",
        t
    )

    top_n = min(
        max(
            int(match_top.group(1))
            if match_top
            else 10,
            1
        ),
        20
    )

    filtros = _filtros_base_grafico(
        pergunta
    )

    # O ranking-base usa faturamento para definir a ordem.
    dados_fat = consultar_ranking(
        "faturamento",
        dimensao,
        filtros,
        top_n,
        "desc"
    )

    rotulos = []
    fat = []
    meta = []
    ating = []

    for item in dados_fat:
        nome = item[
            "item"
        ]

        rotulos.append(
            nome
        )

        fat.append(
            item.get(
                "valor"
            )
            or 0
        )

        filtros_item = dict(
            filtros
        )

        # A dimensão passa a ser filtro individual.
        filtros_item[
            dimensao
        ] = nome

        meta.append(
            consultar_valor(
                "meta_faturamento",
                filtros_item
            )
            or 0
        )

        ating.append(
            consultar_valor(
                "atingimento_meta_faturamento",
                filtros_item
            )
            or 0
        )

    return {
        "rotulos": rotulos,
        "faturamento": fat,
        "meta": meta,
        "atingimento": ating,
        "filtros": filtros
    }


def gerar_grafico_combinado(
    pergunta
):
    _t_grafico_inicio = time.perf_counter()

    if not _usuario_pediu_grafico_combinado(
        pergunta
    ):
        return None

    t = _texto_normalizado(
        pergunta
    )

    adicionar_tendencia = (
        _usuario_pediu_linha_tendencia(
            pergunta
        )
    )

    tendencia_indicador = (
        _alvo_tendencia_combinado(
            pergunta
        )
    )

    cores = _cores_grafico_combinado(
        pergunta
    )

    # --------------------------------------------------------
    # MÊS A MÊS
    # --------------------------------------------------------
    if any(
        x in t
        for x in [
            "mes a mes",
            "por mes",
            "evolucao"
        ]
    ):
        dados = (
            _series_combinadas_mes_a_mes(
                pergunta
            )
        )

        titulo = (
            "Faturamento x Meta x "
            f"% Atingimento - {dados['ano']}"
        )

        _t_render_inicio = time.perf_counter()

        imagem = (
            _grafico_combinado_png_base64(
                titulo,
                dados["rotulos"],
                dados["faturamento"],
                dados["meta"],
                dados["atingimento"],
                adicionar_tendencia,
                tendencia_indicador,
                cores
            )
        )

        print(
            f"[TEMPO GRÁFICO] Render PNG/base64: "
            f"{time.perf_counter() - _t_render_inicio:.3f}s"
        )
        print(
            f"[TEMPO GRÁFICO] Total gerar_grafico_combinado: "
            f"{time.perf_counter() - _t_grafico_inicio:.3f}s"
        )

        resposta = (
            "📊 Gráfico combinado mês a mês: "
            "faturamento e meta em barras, "
            "% de atingimento em linha."
        )

        if adicionar_tendencia:
            resposta += (
                f" Linha de tendência adicionada "
                f"ao {tendencia_indicador}."
            )

        return {
            "tipo_resposta": "grafico",
            "tipo_grafico": "combinado",
            "resposta": resposta,
            "imagem_base64": imagem,
            "nome_arquivo": "grafico_combinado_mensal.png"
        }

    # --------------------------------------------------------
    # POR DIMENSÃO
    # --------------------------------------------------------
    regras_dimensao = [
        ("regiao", ["regiao", "filial"]),
        ("plataforma", ["cliente", "clientes", "grupo", "grupos", "varejo", "varejista", "varejistas", "plataforma"]),
        ("loja", ["loja", "lojas", "cnpj", "cnpjs", "estabelecimento", "estabelecimentos"]),
        ("produto", ["produto"]),
        ("linha", ["linha"]),
        ("familia", ["familia"]),
        ("representante", ["representante"]),
        ("classe", ["classe"]),
        (
            "status_entrega",
            [
                "status de entrega",
                "status_entrega"
            ]
        ),
        (
            "cod_curva_abc",
            [
                "curva abc",
                "cod_curva_abc"
            ]
        ),
        (
            "analise_credito",
            [
                "analise de credito",
                "análise de crédito"
            ]
        ),
    ]

    dimensao = None

    for chave, termos in (
        regras_dimensao
    ):
        if any(
            termo in t
            for termo in termos
        ):
            dimensao = chave
            break

    if dimensao is None:
        # Default do gráfico combinado: mês a mês.
        # Se o usuário não informou uma dimensão, usa a série mensal.
        pergunta_mensal = (
            pergunta
            + " mês a mês"
        )

        return gerar_grafico_combinado(
            pergunta_mensal
        )

    dados = (
        _series_combinadas_dimensao(
            pergunta,
            dimensao
        )
    )

    descricao_dim = (
        mapa_dimensoes[
            dimensao
        ][
            "descricao"
        ]
    )

    imagem = (
        _grafico_combinado_png_base64(
            (
                "Faturamento x Meta x "
                f"% Atingimento por {descricao_dim}"
            ),
            dados["rotulos"],
            dados["faturamento"],
            dados["meta"],
            dados["atingimento"],
            adicionar_tendencia,
            tendencia_indicador
        )
    )

    resposta = (
        f"📊 Gráfico combinado por {descricao_dim}: "
        "faturamento e meta em barras, "
        "% de atingimento em linha."
    )

    if adicionar_tendencia:
        resposta += (
            f" Linha de tendência adicionada "
            f"ao {tendencia_indicador}."
        )

    return {
        "tipo_resposta": "grafico",
        "tipo_grafico": "combinado",
        "resposta": resposta,
        "imagem_base64": imagem,
        "nome_arquivo": "grafico_combinado_dimensao.png"
    }



def _usuario_pediu_resumo_gerencial(pergunta):
    t = _texto_normalizado(
        pergunta
    )

    return any(
        termo in t
        for termo in [
            "resumo gerencial",
            "visao gerencial",
            "visão gerencial",
            "visao geral",
            "visão geral"
        ]
    )


def _filtros_resumo_gerencial_local(pergunta):
    """
    Extrai localmente os filtros seguros para resumo gerencial.
    Evita chamar IA só para descobrir ano/mês/região.
    """
    filtros = _extrair_periodo_pergunta(
        pergunta
    )

    texto_upper = pergunta.upper()

    for regiao in regioes_conhecidas:
        if regiao in texto_upper:
            filtros["regiao"] = regiao
            break

    return filtros


def _imagem_tabela_resumo_gerencial(
    filtros,
    dados
):
    contexto = construir_contexto_resposta(
        filtros
    )

    titulo = "Resumo gerencial"

    if contexto:
        titulo += f" {contexto}"

    linhas = [
        [
            "Faturamento",
            formatar_valor(
                dados.get("faturamento"),
                "moeda"
            )
        ],
        [
            "Meta de faturamento",
            formatar_valor(
                dados.get("meta_faturamento"),
                "moeda"
            )
        ],
        [
            "Atingimento da meta",
            formatar_valor(
                dados.get("atingimento_faturamento"),
                "percentual"
            )
        ],
        [
            "Margem líquida",
            formatar_valor(
                dados.get("margem_liquida"),
                "percentual"
            )
        ],
        [
            "Meta de margem líquida",
            formatar_valor(
                dados.get("meta_margem_liquida"),
                "percentual"
            )
        ],
        [
            "Margem bruta",
            formatar_valor(
                dados.get("margem_bruta"),
                "percentual"
            )
        ],
        [
            "Quantidade vendida",
            formatar_valor(
                dados.get("quantidade"),
                "inteiro"
            )
        ],
        [
            "Meta de quantidade",
            formatar_valor(
                dados.get("meta_quantidade"),
                "inteiro"
            )
        ],
        [
            "Entregas",
            formatar_valor(
                dados.get("entregas"),
                "moeda"
            )
        ],
        [
            "Faturamento + entregas",
            formatar_valor(
                dados.get("faturamento_entregas"),
                "moeda"
            )
        ],
        [
            "Carteira faturável dia",
            formatar_valor(
                dados.get("carteira_faturavel_dia"),
                "moeda"
            )
        ],
        [
            "Carteira faturável mês",
            formatar_valor(
                dados.get("carteira_faturavel_mes"),
                "moeda"
            )
        ],
        [
            "Faturamento carteira",
            formatar_valor(
                dados.get("faturamento_carteira"),
                "moeda"
            )
        ]
    ]

    return _tabela_png_base64(
        titulo,
        [
            "Indicador",
            "Resultado"
        ],
        linhas,
        fonte=9
    )


def _resposta_resumo_gerencial_tabela(pergunta):
    """
    Trata diretamente pedidos como:
    'Mande um resumo gerencial de 2026 em tabela'

    Não deixa o fluxo genérico transformar o pedido em
    'faturamento por região'.
    """
    if not (
        _usuario_pediu_resumo_gerencial(
            pergunta
        )
        and _usuario_pediu_tabela(
            pergunta
        )
    ):
        return None

    filtros = _filtros_resumo_gerencial_local(
        pergunta
    )

    dados = consultar_resumo_gerencial(
        filtros
    )

    imagem = _imagem_tabela_resumo_gerencial(
        filtros,
        dados
    )

    resposta = construir_resposta_resumo_gerencial(
        filtros,
        dados
    )

    return {
        "tipo_resposta": "grafico",
        "tipo_grafico": "tabela",
        "resposta": resposta,
        "imagem_base64": imagem,
        "nome_arquivo": "resumo_gerencial.png"
    }



# ============================================================
# 42AA. GRÁFICOS TEMPORAIS EM NÍVEL DE DIA / SEMANA
# ============================================================
#
# Esta rotina é ADITIVA: não substitui os gráficos mensais,
# rankings, comparações ou gráficos combinados existentes.
# Ela só assume o processamento quando a pergunta pede
# explicitamente granularidade diária ou semanal.
#

def _granularidade_grafico_dia_semana(pergunta):
    """Detecta apenas pedidos explícitos de série diária/semanal."""
    t = _texto_normalizado(pergunta)

    if any(
        termo in t
        for termo in [
            "dia a dia",
            "por dia",
            "diariamente",
            "diario",
            "diaria",
            "a cada dia",
            "em cada dia",
        ]
    ):
        return "dia"

    if any(
        termo in t
        for termo in [
            "por semana",
            "semana a semana",
            "semanalmente",
            "semanal",
        ]
    ):
        return "semana"

    return None


def _periodo_mes_grafico_temporal(pergunta):
    """
    Define o mês da série diária/semanal.

    Regra de negócio:
    - mês não informado -> mês atual;
    - ano não informado -> ano atual;
    - mês atual -> somente até hoje;
    - mês passado/futuro explícito -> usa os limites naturais do mês.
    """
    periodo_amplo = _periodo_relativo_amplo(pergunta)
    hoje = _data_atual_negocio()

    if periodo_amplo is not None:
        inicio = periodo_amplo["inicio"]
        fim = periodo_amplo["fim"]

        # Em período corrente, não cria pontos futuros.
        if inicio <= hoje <= fim:
            fim = hoje

        return {
            "ano": inicio.year,
            "mes": inicio.month,
            "inicio": inicio,
            "fim": fim,
            "rotulo": periodo_amplo["rotulo"],
        }

    ref = _mes_ano_explicitos_para_periodo(pergunta)

    ano = int(ref["ano"])
    mes = int(ref["mes"])

    inicio = date(ano, mes, 1)
    fim_mes = _ultimo_dia_mes(ano, mes)

    # Para o mês corrente, não cria pontos futuros.
    if ano == hoje.year and mes == hoje.month:
        fim = min(fim_mes, hoje)
    else:
        fim = fim_mes

    nome_mes = meses_nome.get(f"{mes:02d}", f"{mes:02d}")

    return {
        "ano": ano,
        "mes": mes,
        "inicio": inicio,
        "fim": fim,
        "rotulo": f"{nome_mes} de {ano}",
    }


def _filtros_comerciais_grafico_temporal(pergunta):
    """
    Reaproveita os filtros comerciais já existentes do agente,
    mas deixa a data sob controle desta série temporal.
    """
    filtros = dict(
        _filtros_base_grafico(pergunta)
        or {}
    )

    filtros.pop("ano", None)
    filtros.pop("mes", None)
    filtros.pop("_data_inicio", None)
    filtros.pop("_data_fim", None)
    filtros.pop("_periodo_rotulo", None)

    return filtros


def _montar_dax_serie_diaria(indicador, inicio, fim, filtros_base=None):
    """Consulta todos os dias do intervalo em uma única chamada ao Power BI."""
    medida = mapa_indicadores[indicador]["medida"]

    filtros = dict(filtros_base or {})
    filtros["_data_inicio"] = inicio.isoformat()
    filtros["_data_fim"] = fim.isoformat()

    contexto = montar_contexto_final(
        contexto_overview_comercial,
        filtros
    )

    filtros_dax = gerar_filtros_dax(contexto)
    filtros_texto = ",\n        ".join(filtros_dax)

    if filtros_texto:
        corpo = f"""
CALCULATETABLE(
    SUMMARIZECOLUMNS(
        '# CALENDÁRIO'[data],
        \"Resultado\", [{medida}]
    ),
        {filtros_texto}
)"""
    else:
        corpo = f"""
SUMMARIZECOLUMNS(
    '# CALENDÁRIO'[data],
    \"Resultado\", [{medida}]
)"""

    return f"""
EVALUATE
{corpo}
ORDER BY '# CALENDÁRIO'[data]
"""


def _valor_coluna_linha(linha, sufixo):
    """Obtém uma coluna do executeQueries sem depender do prefixo usado pela API."""
    for chave, valor in linha.items():
        chave_norm = str(chave).lower().replace(" ", "")
        if chave_norm.endswith(sufixo.lower().replace(" ", "")):
            return valor
    return None


def _consultar_serie_diaria(indicador, inicio, fim, filtros_base=None):
    dax = _montar_dax_serie_diaria(
        indicador,
        inicio,
        fim,
        filtros_base
    )

    linhas = extrair_linhas(
        executar_dax(dax)
    )

    por_data = {}

    for linha in linhas:
        bruto_data = _valor_coluna_linha(linha, "[data]")
        valor = linha.get("[Resultado]")

        if bruto_data is None:
            continue

        try:
            texto_data = str(bruto_data)[:10]
            data_linha = datetime.strptime(
                texto_data,
                "%Y-%m-%d"
            ).date()
        except Exception:
            continue

        por_data[data_linha] = (
            valor if valor is not None else 0
        )

    # Inclui também dias sem movimento como zero.
    dados = []
    atual = inicio

    while atual <= fim:
        dados.append({
            "data": atual,
            "valor": por_data.get(atual, 0),
        })
        atual += timedelta(days=1)

    return dados


def _intervalos_semanais_no_mes(inicio, fim):
    """
    Gera semanas de segunda a domingo, recortadas pelos limites do mês.
    Assim nenhum dia de outro mês entra no gráfico solicitado.
    """
    intervalos = []
    atual = inicio

    while atual <= fim:
        fim_semana_calendario = atual + timedelta(
            days=(6 - atual.weekday())
        )
        fim_bloco = min(fim_semana_calendario, fim)

        intervalos.append((atual, fim_bloco))
        atual = fim_bloco + timedelta(days=1)

    return intervalos


def _montar_dax_serie_semanal(indicador, inicio, fim, filtros_base=None):
    """
    Monta UMA única consulta DAX para todas as semanas do mês.

    Cada ROW recalcula a medida dentro do intervalo daquela semana; portanto,
    percentuais, margens, metas e demais medidas continuam sendo avaliados pelo
    Power BI no contexto correto. Não há soma de percentuais diários.
    """
    medida = mapa_indicadores[indicador]["medida"]

    # Mantém os defaults comerciais do dashboard, mas remove ano/mês porque
    # o período será controlado diretamente por '# CALENDÁRIO'[data].
    contexto_base = deepcopy(contexto_overview_comercial)
    contexto_base.pop("ano", None)
    contexto_base.pop("mes", None)

    for chave, valor in dict(filtros_base or {}).items():
        if chave not in {
            "ano", "mes", "_data_inicio", "_data_fim", "_periodo_rotulo"
        } and valor is not None:
            contexto_base[chave] = valor

    filtros_dax_base = gerar_filtros_dax(contexto_base)
    blocos = []

    for numero, (ini_semana, fim_semana) in enumerate(
        _intervalos_semanais_no_mes(inicio, fim),
        start=1
    ):
        filtros_semana = list(filtros_dax_base)
        filtros_semana.extend([
            (
                "'# CALENDÁRIO'[data] >= "
                f"DATE({ini_semana.year}, {ini_semana.month}, {ini_semana.day})"
            ),
            (
                "'# CALENDÁRIO'[data] <= "
                f"DATE({fim_semana.year}, {fim_semana.month}, {fim_semana.day})"
            ),
        ])

        filtros_texto = ",\n            ".join(filtros_semana)

        bloco = (
            'ROW(\n'
            f'        "Numero", {numero},\n'
            '        "Resultado",\n'
            '        CALCULATE(\n'
            f'            [{medida}],\n'
            f'            {filtros_texto}\n'
            '        )\n'
            '    )'
        )
        blocos.append(bloco)

    return "EVALUATE\nUNION(\n    " + ",\n    ".join(blocos) + "\n)"


def _consultar_serie_semanal(indicador, inicio, fim, filtros_base=None):
    """
    Consulta todas as semanas em UMA única chamada ao Power BI.
    Mantém o mesmo resultado funcional da implementação anterior, mas elimina
    uma requisição HTTP separada para cada semana.
    """
    intervalos = _intervalos_semanais_no_mes(inicio, fim)

    dax = _montar_dax_serie_semanal(
        indicador,
        inicio,
        fim,
        filtros_base
    )

    linhas = extrair_linhas(
        executar_dax(dax)
    )

    valores_por_numero = {}

    for linha in linhas:
        numero = _valor_coluna_linha(linha, "[Numero]")
        valor = _valor_coluna_linha(linha, "[Resultado]")

        try:
            numero = int(numero)
        except Exception:
            continue

        valores_por_numero[numero] = (
            valor if valor is not None else 0
        )

    dados = []

    for numero, (ini_semana, fim_semana) in enumerate(
        intervalos,
        start=1
    ):
        dados.append({
            "numero": numero,
            "inicio": ini_semana,
            "fim": fim_semana,
            "valor": valores_por_numero.get(numero, 0),
        })

    return dados


def gerar_grafico_temporal_dia_semana(pergunta, indicador=None):
    """
    Gera gráfico diário ou semanal para qualquer indicador existente em
    mapa_indicadores. Se não houver mês/ano na pergunta, usa o mês/ano atuais.
    """
    granularidade = _granularidade_grafico_dia_semana(pergunta)

    if granularidade is None:
        return None

    indicador = indicador or _indicador_da_pergunta(pergunta)

    if indicador not in mapa_indicadores:
        return None

    config = mapa_indicadores[indicador]
    periodo = _periodo_mes_grafico_temporal(pergunta)
    filtros_base = _filtros_comerciais_grafico_temporal(pergunta)

    tipo_explicito = _tipo_grafico_solicitado(pergunta)
    tipo = tipo_explicito or "barras"
    adicionar_tendencia = _usuario_pediu_linha_tendencia(pergunta)
    cores = _cores_grafico_simples(pergunta)

    if granularidade == "dia":
        serie = _consultar_serie_diaria(
            indicador,
            periodo["inicio"],
            periodo["fim"],
            filtros_base
        )

        rotulos = [
            item["data"].strftime("%d")
            for item in serie
        ]
        valores = [
            item["valor"]
            for item in serie
        ]

        titulo = (
            f"{config['descricao']} dia a dia - "
            f"{periodo['rotulo'].capitalize()}"
        )
        resposta = (
            f"📊 {config['descricao']} dia a dia de "
            f"{periodo['rotulo']}."
        )
        nome_arquivo = "grafico_diario.png"

    else:
        serie = _consultar_serie_semanal(
            indicador,
            periodo["inicio"],
            periodo["fim"],
            filtros_base
        )

        rotulos = [
            (
                f"Sem. {item['numero']}\n"
                f"{item['inicio'].strftime('%d/%m')}–"
                f"{item['fim'].strftime('%d/%m')}"
            )
            for item in serie
        ]
        valores = [
            item["valor"]
            for item in serie
        ]

        titulo = (
            f"{config['descricao']} por semana - "
            f"{periodo['rotulo'].capitalize()}"
        )
        resposta = (
            f"📊 {config['descricao']} por semana de "
            f"{periodo['rotulo']}."
        )
        nome_arquivo = "grafico_semanal.png"

    imagem = _grafico_png_base64(
        tipo,
        titulo,
        rotulos,
        valores,
        config["descricao"],
        config["formato"],
        adicionar_tendencia,
        cores["serie"],
        cores["tendencia"]
    )

    return {
        "tipo_resposta": "grafico",
        "tipo_grafico": tipo,
        "resposta": resposta,
        "imagem_base64": imagem,
        "nome_arquivo": nome_arquivo,
    }


def gerar_resposta_temporal_texto(pergunta):
    """
    Retorna série diária/semanal em TEXTO quando o usuário pede explicitamente
    "dia a dia", "por dia", "por semana" etc. sem pedir gráfico.

    Esta rotina é aditiva e reutiliza exatamente as mesmas consultas temporais
    dos gráficos. Se mês/ano não forem informados, usa mês/ano atuais.
    """
    if usuario_pediu_grafico(pergunta):
        return None

    granularidade = _granularidade_grafico_dia_semana(pergunta)

    if granularidade is None:
        return None

    indicador = _indicador_da_pergunta(pergunta)

    if indicador not in mapa_indicadores:
        return None

    config = mapa_indicadores[indicador]
    periodo = _periodo_mes_grafico_temporal(pergunta)
    filtros_base = _filtros_comerciais_grafico_temporal(pergunta)

    if granularidade == "dia":
        serie = _consultar_serie_diaria(
            indicador,
            periodo["inicio"],
            periodo["fim"],
            filtros_base
        )

        linhas = [
            f"{config['descricao']} dia a dia de {periodo['rotulo']}:"
        ]

        for item in serie:
            valor = item["valor"]
            if valor is None:
                valor = 0

            linhas.append(
                f"{item['data'].strftime('%d')}: "
                f"{formatar_valor(valor, config['formato'])}"
            )

    else:
        serie = _consultar_serie_semanal(
            indicador,
            periodo["inicio"],
            periodo["fim"],
            filtros_base
        )

        linhas = [
            f"{config['descricao']} por semana de {periodo['rotulo']}:"
        ]

        for item in serie:
            valor = item["valor"]
            if valor is None:
                valor = 0

            linhas.append(
                f"Sem. {item['numero']} "
                f"({item['inicio'].strftime('%d/%m')}–"
                f"{item['fim'].strftime('%d/%m')}): "
                f"{formatar_valor(valor, config['formato'])}"
            )

    return {
        "tipo_resposta": "texto",
        "resposta": "\n".join(linhas)
    }


def gerar_resposta_grafico(pergunta):
    if not usuario_pediu_grafico(pergunta):
        return None

    # --------------------------------------------------------
    # PRIORIDADE MÁXIMA:
    # resumo gerencial pedido explicitamente em tabela.
    # --------------------------------------------------------
    resumo_gerencial_tabela = (
        _resposta_resumo_gerencial_tabela(
            pergunta
        )
    )

    if resumo_gerencial_tabela is not None:
        return resumo_gerencial_tabela

    # --------------------------------------------------------
    # NOVO CAMINHO ADITIVO: série diária / semanal.
    # Só entra quando a granularidade foi pedida explicitamente.
    # O restante dos gráficos antigos permanece abaixo, intacto.
    # --------------------------------------------------------
    temporal_dia_semana = gerar_grafico_temporal_dia_semana(
        pergunta
    )

    if temporal_dia_semana is not None:
        return temporal_dia_semana

    # Gráfico combinado tem prioridade sobre a interpretação
    # de indicador único.
    combinado = gerar_grafico_combinado(
        pergunta
    )

    if combinado is not None:
        return combinado

    indicador = _indicador_da_pergunta(
        pergunta
    )

    config = mapa_indicadores[
        indicador
    ]

    tipo_explicito = _tipo_grafico_solicitado(
        pergunta
    )

    pediu_tabela = _usuario_pediu_tabela(
        pergunta
    )

    adicionar_tendencia = (
        _usuario_pediu_linha_tendencia(
            pergunta
        )
    )

    # ========================================================
    # PRIORIDADE 1: COMPARAÇÕES
    # ========================================================
    comparacao = consultar_comparacao_generica(
        pergunta
    )

    if comparacao:
        tipo_resposta_comparacao = comparacao.get(
            "tipo_resposta"
        )

        # ----------------------------------------------------
        # TABELA TEM PRIORIDADE ABSOLUTA
        # ----------------------------------------------------
        if pediu_tabela:
            if tipo_resposta_comparacao in {
                "comparacao_periodos",
                "comparativo_gerencial_periodos"
            }:
                imagem = _imagem_tabela_comparacao_periodos(
                    comparacao
                )

                return {
                    "tipo_resposta": "grafico",
                    "tipo_grafico": "tabela",
                    "resposta": comparacao["resposta"],
                    "imagem_base64": imagem,
                    "nome_arquivo": "tabela_comparativa_periodos.png"
                }

            if tipo_resposta_comparacao in {
                "comparacao_temporal",
                "tabela"
            }:
                imagem = _imagem_tabela_comparacao_temporal(
                    comparacao,
                    config
                )

                return {
                    "tipo_resposta": "grafico",
                    "tipo_grafico": "tabela",
                    "resposta": comparacao["resposta"],
                    "imagem_base64": imagem,
                    "nome_arquivo": "tabela_comparativa.png"
                }

            # Comparação agregada simples.
            imagem = _imagem_tabela_comparacao_simples(
                comparacao,
                config
            )

            return {
                "tipo_resposta": "grafico",
                "tipo_grafico": "tabela",
                "resposta": comparacao["resposta"],
                "imagem_base64": imagem,
                "nome_arquivo": "tabela_comparativa.png"
            }

        # ----------------------------------------------------
        # COMPARAÇÃO DE DOIS PERÍODOS EM GRÁFICO
        # ----------------------------------------------------
        if tipo_resposta_comparacao == "comparacao_periodos":
            config_periodos = mapa_indicadores[
                comparacao["indicador"]
            ]

            rotulos = [
                comparacao["periodo_a"]["nome"],
                comparacao["periodo_b"]["nome"]
            ]

            valores = [
                comparacao.get("valor_a") or 0,
                comparacao.get("valor_b") or 0
            ]

            tipo_comparacao = tipo_explicito or "barras"

            imagem = _grafico_png_base64(
                tipo_comparacao,
                f"Comparativo de {config_periodos['descricao'].lower()}",
                rotulos,
                valores,
                config_periodos["descricao"],
                config_periodos["formato"],
                adicionar_tendencia,
                _cores_grafico_simples(pergunta)["serie"],
                _cores_grafico_simples(pergunta)["tendencia"]
            )

            return {
                "tipo_resposta": "grafico",
                "tipo_grafico": tipo_comparacao,
                "resposta": comparacao["resposta"],
                "imagem_base64": imagem,
                "nome_arquivo": "comparativo_periodos.png"
            }

        if tipo_resposta_comparacao == "comparativo_gerencial_periodos":
            imagem = _imagem_tabela_comparacao_periodos(
                comparacao
            )

            return {
                "tipo_resposta": "grafico",
                "tipo_grafico": "tabela",
                "resposta": comparacao["resposta"],
                "imagem_base64": imagem,
                "nome_arquivo": "comparativo_gerencial_periodos.png"
            }

        # ----------------------------------------------------
        # COMPARAÇÃO SIMPLES EM GRÁFICO
        # ----------------------------------------------------
        if tipo_resposta_comparacao not in {
            "comparacao_temporal",
            "tabela",
            "comparacao_periodos",
            "comparativo_gerencial_periodos"
        }:
            rotulos = [
                x["item"]
                for x in comparacao["dados"]
            ]

            valores = [
                x["valor"] or 0
                for x in comparacao["dados"]
            ]

            tipo_comparacao = (
                tipo_explicito
                or "barras"
            )

            imagem = _grafico_png_base64(
                tipo_comparacao,
                f"Comparativo de {config['descricao'].lower()}",
                rotulos,
                valores,
                config["descricao"],
                config["formato"],
                adicionar_tendencia,
                _cores_grafico_simples(pergunta)["serie"],
                _cores_grafico_simples(pergunta)["tendencia"]
            )

            return {
                "tipo_resposta": "grafico",
                "tipo_grafico": tipo_comparacao,
                "resposta": comparacao["resposta"],
                "imagem_base64": imagem,
                "nome_arquivo": "comparativo.png"
            }

        # Comparação temporal sem pedido explícito de tabela:
        # devolve texto analítico; não cai no gráfico mensal geral.
        return None

    t = _texto_normalizado(
        pergunta
    )

    # ========================================================
    # PRIORIDADE 2: SÉRIE TEMPORAL
    # ========================================================
    if any(
        x in t
        for x in [
            "mes a mes",
            "evolucao",
            "por mes"
        ]
    ):
        match_ano = re.search(
            r"\b(20\d{2})\b",
            pergunta
        )

        ano = (
            match_ano.group(1)
            if match_ano
            else str(datetime.now().year)
        )

        limite = (
            datetime.now().month
            if int(ano) == datetime.now().year
            else 12
        )

        rotulos = []
        valores = []

        for n in range(
            1,
            limite + 1
        ):
            mes = f"{n:02d}"

            filtros = {
                "ano": ano,
                "mes": normalizar_mes_powerbi(
                    mes,
                    ano
                )
            }

            rotulos.append(
                nome_mes_resposta(
                    mes
                ).capitalize()
            )

            valores.append(
                consultar_valor(
                    indicador,
                    filtros
                )
                or 0
            )

        titulo_mensal = (
            f"{config['descricao']} mês a mês - {ano}"
        )

        if pediu_tabela:
            imagem = _imagem_tabela_serie_temporal(
                titulo_mensal,
                rotulos,
                valores,
                config
            )

            return {
                "tipo_resposta": "grafico",
                "tipo_grafico": "tabela",
                "resposta": f"📋 {config['descricao']} mês a mês de {ano}.",
                "imagem_base64": imagem,
                "nome_arquivo": "tabela_mensal.png"
            }

        tipo_mensal = (
            tipo_explicito
            or "linha"
        )

        imagem = _grafico_png_base64(
            tipo_mensal,
            titulo_mensal,
            rotulos,
            valores,
            config["descricao"],
            config["formato"],
            adicionar_tendencia,
            _cores_grafico_simples(pergunta)["serie"],
            _cores_grafico_simples(pergunta)["tendencia"]
        )

        return {
            "tipo_resposta": "grafico",
            "tipo_grafico": tipo_mensal,
            "resposta": f"📈 {config['descricao']} mês a mês de {ano}.",
            "imagem_base64": imagem,
            "nome_arquivo": "grafico_mensal.png"
        }

    # ========================================================
    # PRIORIDADE 3: RANKING / DIMENSÃO
    # ========================================================
    regras_dimensao = [
        ("regiao", ["regiao", "filial"]),
        ("plataforma", ["cliente", "clientes", "grupo", "grupos", "varejo", "varejista", "varejistas", "plataforma"]),
        ("loja", ["loja", "lojas", "cnpj", "cnpjs", "estabelecimento", "estabelecimentos"]),
        ("produto", ["produto"]),
        ("linha", ["linha"]),
        ("familia", ["familia"]),
        ("representante", ["representante"]),
        ("classe", ["classe"]),
    ]

    dimensao = None

    for chave, termos in regras_dimensao:
        if any(
            termo in t
            for termo in termos
        ):
            dimensao = chave
            break

    if dimensao is None:
        # Não inventar região.
        # Se o usuário pediu tabela/gráfico sem uma dimensão explícita,
        # este caminho genérico não deve fabricar um ranking.
        return None

    match_top = re.search(
        r"\b(?:top\s*)?(\d+)\b",
        t
    )

    top_n = min(
        max(
            int(match_top.group(1))
            if match_top
            else 10,
            1
        ),
        20
    )

    filtros = _extrair_periodo_pergunta(
        pergunta
    )

    ranking = consultar_ranking(
        indicador,
        dimensao,
        top_n,
        "desc",
        filtros
    )

    descricao_dimensao = mapa_dimensoes[
        dimensao
    ]["descricao"]

    if (
        dimensao == "regiao"
        and (
            "filial" in t
            or "filiais" in t
        )
    ):
        descricao_dimensao = "filial"

    titulo = (
        f"{config['descricao']} por {descricao_dimensao}"
    )

    # --------------------------------------------------------
    # TABELA PARA RANKING
    # --------------------------------------------------------
    if pediu_tabela:
        imagem = _imagem_tabela_ranking(
            titulo,
            ranking,
            config
        )

        return {
            "tipo_resposta": "grafico",
            "tipo_grafico": "tabela",
            "resposta": f"📋 {titulo}.",
            "imagem_base64": imagem,
            "nome_arquivo": "tabela_comercial.png"
        }

    rotulos = [
        x["item"]
        for x in ranking
    ]

    valores = [
        x["valor"] or 0
        for x in ranking
    ]

    if tipo_explicito == "pizza":
        tipo = (
            "pizza"
            if len(rotulos) <= 6
            else "barras_horizontais"
        )

    elif tipo_explicito is not None:
        tipo = tipo_explicito

    elif dimensao in {
        "cliente",
        "produto",
        "representante",
        "plataforma"
    }:
        tipo = "barras_horizontais"

    else:
        tipo = "barras"

    imagem = _grafico_png_base64(
        tipo,
        titulo,
        rotulos,
        valores,
        config["descricao"],
        config["formato"],
        adicionar_tendencia,
        _cores_grafico_simples(pergunta)["serie"],
        _cores_grafico_simples(pergunta)["tendencia"]
    )

    return {
        "tipo_resposta": "grafico",
        "tipo_grafico": tipo,
        "resposta": f"📊 {titulo}.",
        "imagem_base64": imagem,
        "nome_arquivo": "grafico_comercial.png"
    }



# ============================================================
# 43A. ESTATÍSTICAS GENÉRICAS (MAIOR / MENOR / MÉDIA / MEDIANA)
# ============================================================

def _detectar_tipo_estatistica(pergunta):
    t = _texto_normalizado(pergunta)

    if any(x in t for x in ["mediana", "valor mediano"]):
        return "mediana"

    if any(x in t for x in ["media", "media mensal", "media anual", "promedio"]):
        return "media"

    if any(x in t for x in ["menor", "menores", "minimo", "minima", "pior"]):
        return "menor"

    if any(x in t for x in ["maior", "maiores", "maximo", "maxima", "melhor"]):
        return "maior"

    return None


def _detectar_dimensao_estatistica(pergunta, tipo_estatistica):
    t = _texto_normalizado(pergunta)

    if any(x in t for x in [
        "qual mes", "que mes", "mes de maior", "mes de menor",
        "por mes", "media mensal", "mediana mensal", "mensalmente"
    ]):
        return "mes"

    regras = [
        ("regiao", ["qual regiao", "qual filial", "entre regioes", "entre filiais", "por regioes", "por filiais", "media por regiao", "mediana por regiao"]),
        ("cliente", ["qual cliente", "entre clientes", "por clientes", "media por cliente", "mediana por cliente"]),
        ("produto", ["qual produto", "entre produtos", "por produtos", "media por produto", "mediana por produto"]),
        ("linha", ["qual linha", "entre linhas", "por linhas", "media por linha", "mediana por linha"]),
        ("familia", ["qual familia", "entre familias", "por familias", "media por familia", "mediana por familia"]),
        ("representante", ["qual representante", "entre representantes", "por representantes", "media por representante", "mediana por representante"]),
        ("plataforma", ["qual plataforma", "entre plataformas", "por plataformas", "media por plataforma", "mediana por plataforma"]),
        ("classe", ["qual classe", "entre classes", "por classes", "media por classe", "mediana por classe"]),
    ]

    for dimensao, termos in regras:
        if any(termo in t for termo in termos):
            return dimensao

    padroes_singulares = [
        ("regiao", r"\bpor\s+(?:regiao|filial)\s*(?:\?|$)"),
        ("cliente", r"\bpor\s+cliente\s*(?:\?|$)"),
        ("produto", r"\bpor\s+produto\s*(?:\?|$)"),
        ("linha", r"\bpor\s+linha\s*(?:\?|$)"),
        ("familia", r"\bpor\s+familia\s*(?:\?|$)"),
        ("representante", r"\bpor\s+representante\s*(?:\?|$)"),
        ("plataforma", r"\bpor\s+plataforma\s*(?:\?|$)"),
        ("classe", r"\bpor\s+classe\s*(?:\?|$)"),
    ]

    for dimensao, padrao in padroes_singulares:
        if re.search(padrao, t):
            return dimensao

    if tipo_estatistica in {"media", "mediana", "maior", "menor"}:
        return "mes"

    return None


def _pergunta_base_estatistica(pergunta, dimensao):
    texto = pergunta

    substituicoes = [
        r"\bmediana\b",
        r"\bm[eé]dia\b",
        r"\bmaior(?:es)?\b",
        r"\bmenor(?:es)?\b",
        r"\bm[aá]xim[oa]s?\b",
        r"\bm[ií]nim[oa]s?\b",
        r"\bmelhor(?:es)?\b",
        r"\bpior(?:es)?\b",
    ]

    for padrao in substituicoes:
        texto = re.sub(padrao, " ", texto, flags=re.IGNORECASE)

    if dimensao == "mes":
        texto = re.sub(r"\bqual\s+m[eê]s\b", " ", texto, flags=re.IGNORECASE)
        texto = re.sub(r"\bque\s+m[eê]s\b", " ", texto, flags=re.IGNORECASE)
        texto = re.sub(r"\bpor\s+m[eê]s\b", " ", texto, flags=re.IGNORECASE)
        texto = re.sub(r"\bmensal(?:mente)?\b", " ", texto, flags=re.IGNORECASE)

    texto = re.sub(r"\s+", " ", texto).strip(" ,.?-")

    if texto:
        return f"Qual o valor de {texto}?"

    return pergunta


def _interpretar_base_estatistica(pergunta, dimensao):
    pergunta_base = _pergunta_base_estatistica(pergunta, dimensao)

    anterior = contexto_conversa.get("ultima_interpretacao")
    contexto_conversa["ultima_interpretacao"] = None

    try:
        tentativas = [
            interpretar_com_groq,
            interpretar_com_gemini,
            interpretar_com_claude,
        ]

        ultimo_erro = None

        for funcao in tentativas:
            try:
                interp = funcao(pergunta_base)
                interp = corrigir_periodo_explicito(pergunta, interp)

                # Se estamos procurando QUAL MÊS teve maior/menor/média/mediana,
                # não aceite um mês inventado pela IA como filtro. Só preserva
                # filtros mensais quando o usuário explicitamente informou um período.
                if dimensao == "mes":
                    texto_original = _texto_normalizado(pergunta)
                    meses_explicitos = any(
                        re.search(
                            r"\b" + re.escape(nome) + r"\b",
                            texto_original
                        )
                        for nome in mapa_meses.keys()
                    )
                    trimestre_explicito = bool(re.search(
                        r"\b(?:q[1-4]|[1-4][ºo]?\s*trimestre|primeiro\s+trimestre|segundo\s+trimestre|terceiro\s+trimestre|quarto\s+trimestre)\b",
                        texto_original,
                        flags=re.IGNORECASE
                    ))
                    if not meses_explicitos and not trimestre_explicito:
                        interp.filtros.mes = None

                interp.operacao = "valor"
                interp.agrupar_por = None
                interp.top_n = None
                interp.ordem = None
                return normalizar_interpretacao(interp)
            except Exception as erro:
                ultimo_erro = erro

        if ultimo_erro:
            raise ultimo_erro

        raise RuntimeError("Não foi possível interpretar a pergunta estatística.")

    finally:
        contexto_conversa["ultima_interpretacao"] = anterior


def _meses_para_estatistica(filtros):
    filtros = dict(filtros or {})
    ano = filtros.get("ano")
    meses = filtros.get("mes")

    if isinstance(meses, list):
        return [str(m).zfill(2) for m in meses]

    if isinstance(meses, str) and meses not in {"Mês atual", "Mes atual"}:
        if re.fullmatch(r"\d{1,2}", meses):
            return [str(meses).zfill(2)]

    agora = datetime.now()

    if ano in (None, "Ano atual", str(agora.year)):
        limite = agora.month
    else:
        limite = 12

    return [f"{n:02d}" for n in range(1, limite + 1)]


def _consultar_estatistica_mensal(indicador, tipo_estatistica, filtros):
    filtros_base = dict(filtros or {})
    meses = _meses_para_estatistica(filtros_base)

    ano = filtros_base.get("ano")

    # ========================================================
    # ESTATÍSTICA MENSAL GENÉRICA PARA QUALQUER INDICADOR
    # ========================================================
    #
    # Cada mês usa exatamente o mesmo filtro mensal do dashboard.
    # Tudo é retornado em UMA única chamada DAX.
    # ========================================================

    if indicador not in mapa_indicadores:
        return None

    medida = mapa_indicadores[
        indicador
    ]["medida"]

    filtros_sem_mes = dict(
        filtros_base
    )

    filtros_sem_mes.pop(
        "mes",
        None
    )

    contexto_base = montar_contexto_final(
        contexto_overview_comercial,
        filtros_sem_mes
    )

    contexto_base.pop(
        "mes",
        None
    )

    filtros_base_dax = gerar_filtros_dax(
        contexto_base
    )

    rows = []

    for mes in meses:

        mes_num = str(
            mes
        ).zfill(2)

        mes_filtro = normalizar_mes_powerbi(
            mes_num,
            ano
        )

        mes_seguro = str(
            mes_filtro
        ).replace(
            '"',
            '""'
        )

        filtros_linha = list(
            filtros_base_dax
        )

        filtros_linha.append(
            "TREATAS("
            + '{"' + mes_seguro + '"}, '
            + "'# CALENDÁRIO'[mes_atual]"
            + ")"
        )

        filtros_texto = ",\n                    ".join(
            filtros_linha
        )

        row_dax = (
            'ROW(\n'
            f'    "Mes", "{mes_num}",\n'
            '    "Valor",\n'
            '    CALCULATE(\n'
            f'        [{medida}],\n'
            f'        {filtros_texto}\n'
            '    )\n'
            ')'
        )

        rows.append(
            row_dax
        )

    dax = (
        "EVALUATE\n"
        "UNION(\n"
        + ",\n".join(rows)
        + "\n)"
    )

    linhas = extrair_linhas(
        executar_dax(
            dax
        )
    )

    dados = []

    for linha in linhas:

        mes_retorno = (
            linha.get("[Mes]")
            or linha.get("Mes")
        )

        if "[Valor]" in linha:
            valor = linha.get("[Valor]")
        else:
            valor = linha.get("Valor")

        if (
            mes_retorno is None
            or valor is None
        ):
            continue

        mes_retorno = str(
            mes_retorno
        ).zfill(2)

        dados.append({
            "item": nome_mes_resposta(
                mes_retorno
            ).capitalize(),
            "mes": mes_retorno,
            "valor": valor
        })

    if not dados:
        return None

    valores = [
        x["valor"]
        for x in dados
    ]

    if tipo_estatistica == "maior":

        escolhido = max(
            dados,
            key=lambda x: x["valor"]
        )

        return {
            "tipo": tipo_estatistica,
            "item": escolhido["item"],
            "valor": escolhido["valor"],
            "dados": dados
        }

    if tipo_estatistica == "menor":

        escolhido = min(
            dados,
            key=lambda x: x["valor"]
        )

        return {
            "tipo": tipo_estatistica,
            "item": escolhido["item"],
            "valor": escolhido["valor"],
            "dados": dados
        }

    if tipo_estatistica == "media":

        return {
            "tipo": tipo_estatistica,
            "item": None,
            "valor": sum(valores) / len(valores),
            "dados": dados
        }

    if tipo_estatistica == "mediana":

        ordenados = sorted(
            valores
        )

        n = len(
            ordenados
        )

        meio = n // 2

        if n % 2:
            valor = ordenados[
                meio
            ]
        else:
            valor = (
                ordenados[meio - 1]
                + ordenados[meio]
            ) / 2

        return {
            "tipo": tipo_estatistica,
            "item": None,
            "valor": valor,
            "dados": dados
        }

    return None


def montar_dax_estatistica_dimensao(indicador, agrupar_por, tipo_estatistica, filtros=None):
    medida = mapa_indicadores[indicador]["medida"]
    dimensao = mapa_dimensoes[agrupar_por]

    contexto = montar_contexto_final(
        contexto_overview_comercial,
        filtros or {}
    )

    filtros_dax = gerar_filtros_dax(contexto)
    filtros_texto = ",\n        ".join(filtros_dax)

    tabela = dimensao["tabela"]
    coluna = dimensao["coluna"]

    funcao = "AVERAGEX" if tipo_estatistica == "media" else "MEDIANX"

    return f'''\nEVALUATE\nVAR Base =\n    CALCULATETABLE(\n        ADDCOLUMNS(\n            FILTER(\n                VALUES('{tabela}'[{coluna}]),\n                NOT ISBLANK('{tabela}'[{coluna}])\n            ),\n            "__Valor", CALCULATE([{medida}])\n        ),\n        {filtros_texto}\n    )\nRETURN\nROW(\n    "Resultado",\n    {funcao}(\n        FILTER(Base, NOT ISBLANK([__Valor])),\n        [__Valor]\n    )\n)\n'''


def _consultar_estatistica_dimensao(indicador, dimensao, tipo_estatistica, filtros):
    if tipo_estatistica in {"maior", "menor"}:
        ranking = consultar_ranking(
            indicador,
            dimensao,
            1,
            "desc" if tipo_estatistica == "maior" else "asc",
            filtros
        )

        if not ranking:
            return None

        return {
            "tipo": tipo_estatistica,
            "item": ranking[0]["item"],
            "valor": ranking[0]["valor"],
            "dados": ranking,
        }

    dax = montar_dax_estatistica_dimensao(
        indicador,
        dimensao,
        tipo_estatistica,
        filtros
    )

    linhas = extrair_linhas(executar_dax(dax))
    if not linhas:
        return None

    valor = linhas[0].get("[Resultado]")
    if valor is None:
        return None

    return {
        "tipo": tipo_estatistica,
        "item": None,
        "valor": valor,
        "dados": None,
    }


def _rotulo_dimensao_estatistica(dimensao):
    if dimensao == "mes":
        return "mês"
    return mapa_dimensoes.get(dimensao, {}).get("descricao", dimensao)


def _construir_resposta_estatistica(indicador, dimensao, tipo_estatistica, resultado, filtros):
    config = mapa_indicadores[indicador]
    descricao = config["descricao"].lower()
    rotulo_dim = _rotulo_dimensao_estatistica(dimensao)
    valor_fmt = formatar_valor(resultado["valor"], config["formato"])

    # Quando a própria dimensão analisada é mês, o mês encontrado é RESULTADO,
    # não deve aparecer novamente como filtro/contexto da frase.
    filtros_resposta = dict(filtros or {})
    if dimensao == "mes":
        filtros_resposta.pop("mes", None)

    contexto = construir_contexto_resposta(filtros_resposta)
    sufixo = f" {contexto}" if contexto else ""

    if tipo_estatistica == "maior":
        if dimensao == "mes":
            return (
                f"🏆 O mês com maior {descricao}{sufixo} foi "
                f"{str(resultado['item']).lower()}, com {valor_fmt}."
            )
        return (
            f"🏆 {resultado['item']} teve o maior {descricao}{sufixo}, "
            f"com {valor_fmt}."
        )

    if tipo_estatistica == "menor":
        if dimensao == "mes":
            return (
                f"📉 O mês com menor {descricao}{sufixo} foi "
                f"{str(resultado['item']).lower()}, com {valor_fmt}."
            )
        return (
            f"📉 {resultado['item']} teve o menor {descricao}{sufixo}, "
            f"com {valor_fmt}."
        )

    if tipo_estatistica == "media":
        if dimensao == "mes":
            return f"📊 A média mensal de {descricao}{sufixo} foi de {valor_fmt}."
        return f"📊 A média de {descricao} por {rotulo_dim}{sufixo} foi de {valor_fmt}."

    if dimensao == "mes":
        return f"📊 A mediana mensal de {descricao}{sufixo} foi de {valor_fmt}."

    return f"📊 A mediana de {descricao} por {rotulo_dim}{sufixo} foi de {valor_fmt}."


def consultar_estatistica_generica(pergunta):
    tipo_estatistica = _detectar_tipo_estatistica(pergunta)
    if tipo_estatistica is None:
        return None

    t = _texto_normalizado(pergunta)
    if any(x in t for x in ["top ", "ranking"]):
        return None

    dimensao = _detectar_dimensao_estatistica(
        pergunta,
        tipo_estatistica
    )

    if dimensao is None:
        return None

    # Para estatística mensal com indicador explícito, resolve tudo
    # localmente. Isso preserva o comportamento de perguntas como:
    # "Qual mês de maior faturamento?"
    if dimensao == "mes":
        try:
            indicador_local = _indicador_da_pergunta(
                pergunta
            )

            filtros = _extrair_periodo_pergunta(
                pergunta
            )

            texto_upper = pergunta.upper()

            for regiao in regioes_conhecidas:
                if regiao in texto_upper:
                    filtros["regiao"] = regiao
                    break

            resultado = _consultar_estatistica_mensal(
                indicador_local,
                tipo_estatistica,
                filtros
            )

            if resultado is None:
                return {
                    "tipo_resposta": "texto",
                    "operacao": "estatistica",
                    "resposta": "Não encontrei dados suficientes para calcular essa estatística."
                }

            resposta = _construir_resposta_estatistica(
                indicador_local,
                "mes",
                tipo_estatistica,
                resultado,
                filtros
            )

            contexto_conversa["ultima_interpretacao"] = {
                "operacao": "valor",
                "indicador": indicador_local,
                "filtros": filtros,
                "agrupar_por": None,
                "top_n": None,
                "ordem": None,
                "fora_escopo": False,
            }

            return {
                "tipo_resposta": "texto",
                "operacao": "estatistica",
                "estatistica": tipo_estatistica,
                "indicador": indicador_local,
                "agrupar_por": "mes",
                "filtros": filtros,
                "resultado": resultado,
                "resposta": resposta,
            }

        except Exception as erro:
            print(
                "Falha na estatística mensal local:",
                type(erro).__name__,
                "-",
                str(erro)[:300]
            )
            # Se algo inesperado ocorrer, cai no fluxo antigo abaixo.

    try:
        interpretacao = _interpretar_base_estatistica(
            pergunta,
            dimensao
        )
    except Exception as erro:
        print(
            "Falha ao interpretar estatística:",
            type(erro).__name__,
            "-",
            str(erro)[:300]
        )
        return None

    if interpretacao.fora_escopo:
        return None

    filtros = interpretacao.filtros.model_dump(exclude_none=True)

    if dimensao == "mes":
        resultado = _consultar_estatistica_mensal(
            interpretacao.indicador,
            tipo_estatistica,
            filtros
        )
    else:
        resultado = _consultar_estatistica_dimensao(
            interpretacao.indicador,
            dimensao,
            tipo_estatistica,
            filtros
        )

    if resultado is None:
        return {
            "tipo_resposta": "texto",
            "operacao": "estatistica",
            "resposta": "Não encontrei dados suficientes para calcular essa estatística."
        }

    resposta = _construir_resposta_estatistica(
        interpretacao.indicador,
        dimensao,
        tipo_estatistica,
        resultado,
        filtros
    )

    contexto_conversa["ultima_interpretacao"] = {
        "operacao": "valor",
        "indicador": interpretacao.indicador,
        "filtros": filtros,
        "agrupar_por": None,
        "top_n": None,
        "ordem": None,
        "fora_escopo": False,
    }

    return {
        "tipo_resposta": "texto",
        "operacao": "estatistica",
        "estatistica": tipo_estatistica,
        "indicador": interpretacao.indicador,
        "agrupar_por": dimensao,
        "filtros": filtros,
        "resultado": resultado,
        "resposta": resposta,
    }


# ============================================================
# 44. WRAPPER DE CONVERSA
# ============================================================

_conversar_v6_validado = conversar


def conversar(pergunta, mostrar_origem=False):
    estatistica = consultar_estatistica_generica(pergunta)
    if estatistica:
        return estatistica["resposta"]

    comparacao = consultar_comparacao_generica(pergunta)
    if comparacao:
        return comparacao["resposta"]

    # Toda consulta que não é estatística nem comparação continua exatamente
    # no fluxo V6 validado.
    return _conversar_v6_validado(pergunta, mostrar_origem)


# ============================================================
# 45. API DO AGENTE
# ============================================================

from fastapi import FastAPI
from pydantic import BaseModel

app = FastAPI()


# ============================================================
# WARM-UP DE STARTUP
# ============================================================

@app.on_event("startup")
def startup_leve():
    """
    Handler de startup. NÃO executa nada bloqueante.

    Motivo:
    o Uvicorn executa o lifespan de startup ANTES de abrir o socket
    da porta. Qualquer chamada externa aqui mantém a porta fechada
    e faz o healthcheck do EasyPanel reiniciar o container.

    O warm-up continua existindo, mas:
    - é OPCIONAL (ligado por WARMUP_ATIVO=1);
    - roda em thread daemon separada, depois do bind da porta;
    - nunca bloqueia a inicialização da API.
    """

    print(
        "[STARTUP] Agente Comercial iniciado com sucesso."
    )

    if os.getenv("WARMUP_ATIVO") != "1":

        print(
            "[STARTUP] Warm-up desativado "
            "(defina WARMUP_ATIVO=1 para habilitar)."
        )

        return

    import threading

    threading.Thread(
        target=_executar_warmup,
        name="warmup",
        daemon=True
    ).start()

    print(
        "[STARTUP] Warm-up agendado em background."
    )


def _executar_warmup():
    """
    Pré-aquece autenticação, Power BI e Matplotlib uma única vez.

    Importante:
    - roda FORA do startup, em thread daemon;
    - falhas de warm-up NÃO derrubam a API;
    - apenas são registradas no console;
    - nenhuma consulta de negócio pesada é executada.
    """

    _t_total = time.perf_counter()

    print("\n" + "=" * 70)
    print("[WARM-UP] Iniciando pré-aquecimento da API...")
    print("=" * 70)

    # --------------------------------------------------------
    # 1. TOKEN POWER BI
    # --------------------------------------------------------

    try:
        _t = time.perf_counter()

        obter_token_powerbi()

        print(
            f"[WARM-UP] Token Power BI: "
            f"{time.perf_counter() - _t:.3f}s"
        )

    except Exception as erro:

        print(
            f"[WARM-UP] Token Power BI FALHOU: "
            f"{type(erro).__name__}: {erro}"
        )

    # --------------------------------------------------------
    # 2. POWER BI / EXECUTE QUERIES
    # --------------------------------------------------------

    try:
        _t = time.perf_counter()

        # Consulta DAX mínima apenas para aquecer a conexão/API.
        executar_dax(
            'EVALUATE ROW("Warmup", 1)'
        )

        print(
            f"[WARM-UP] Power BI ExecuteQueries: "
            f"{time.perf_counter() - _t:.3f}s"
        )

    except Exception as erro:

        print(
            f"[WARM-UP] Power BI FALHOU: "
            f"{type(erro).__name__}: {erro}"
        )

    # --------------------------------------------------------
    # 3. MATPLOTLIB / PRIMEIRO RENDER
    # --------------------------------------------------------

    try:
        _t = time.perf_counter()

        import matplotlib

        matplotlib.use(
            "Agg"
        )

        import matplotlib.pyplot as plt

        fig, ax = plt.subplots(
            figsize=(2, 1)
        )

        ax.plot(
            [0, 1],
            [0, 1]
        )

        # Força de fato o primeiro render/font loading.
        fig.canvas.draw()

        plt.close(
            fig
        )

        print(
            f"[WARM-UP] Matplotlib/render: "
            f"{time.perf_counter() - _t:.3f}s"
        )

    except Exception as erro:

        print(
            f"[WARM-UP] Matplotlib FALHOU: "
            f"{type(erro).__name__}: {erro}"
        )

    print(
        f"[WARM-UP] TOTAL: "
        f"{time.perf_counter() - _t_total:.3f}s"
    )

    print(
        "[WARM-UP] API pronta para receber consultas."
    )

    print("=" * 70 + "\n")



def _resposta_previsao_faturamento_mes_a_mes(pergunta):
    """
    Trata somente previsão de faturamento mês a mês quando o usuário
    NÃO pediu gráfico/tabela explicitamente.

    Usa a mesma consulta mensal genérica já existente no agente.
    """

    t = _texto_normalizado(
        pergunta
    )

    if not any(
        termo in t
        for termo in [
            "previsao de faturamento",
            "projecao de faturamento",
            "forecast"
        ]
    ):
        return None

    if not any(
        termo in t
        for termo in [
            "mes a mes",
            "por mes",
            "evolucao"
        ]
    ):
        return None

    if usuario_pediu_grafico(
        pergunta
    ):
        return None

    match_ano = re.search(
        r"\b(20\d{2})\b",
        pergunta
    )

    filtros = {}

    if match_ano:
        filtros["ano"] = (
            match_ano.group(1)
        )

    serie = _consultar_estatistica_mensal(
        "previsao_faturamento",
        "maior",
        filtros
    )

    if (
        not serie
        or not serie.get("dados")
    ):
        return None

    linhas = [
        "📈 Previsão de faturamento mês a mês:"
    ]

    for item in serie["dados"]:
        valor_formatado = formatar_valor(
            item["valor"],
            mapa_indicadores[
                "previsao_faturamento"
            ]["formato"]
        )

        linhas.append(
            f"{item['item']}: "
            f"{valor_formatado}"
        )

    return {
        "tipo_resposta": "texto",
        "resposta": "\n".join(
            linhas
        )
    }


class PerguntaRequest(BaseModel):
    pergunta: str
    chat_id: Optional[str | int] = None


@app.get("/")
def inicio():
    return {"status": "Agente Comercial V15 - meta contextual e múltiplos indicadores"}


@app.post("/limpar-contexto/{chat_id}")
def limpar_contexto_endpoint(chat_id: str):
    return limpar_contexto_usuario(chat_id)


@app.post("/perguntar")
def receber_pergunta(dados: PerguntaRequest):
    _t_endpoint_inicio = time.perf_counter()

    try:
        def processar():
            _t_processar_inicio = time.perf_counter()

            # ====================================================
            # PREVISÃO DE FATURAMENTO MÊS A MÊS
            # ====================================================
            previsao_mensal = (
                _resposta_previsao_faturamento_mes_a_mes(
                    dados.pergunta
                )
            )

            if previsao_mensal is not None:
                print(
                    f"[TEMPO PROCESSAR] "
                    f"Caminho previsão mensal: "
                    f"{time.perf_counter() - _t_processar_inicio:.3f}s"
                )
                return previsao_mensal

            # ====================================================
            # SÉRIE DIÁRIA / SEMANAL EM TEXTO
            # ====================================================
            # Ex.: "faturamento dia a dia de agosto".
            # Se o usuário pedir gráfico, o fluxo continua para a rotina
            # de gráficos já existente, sem qualquer alteração nela.
            temporal_texto = gerar_resposta_temporal_texto(
                dados.pergunta
            )

            if temporal_texto is not None:
                print(
                    f"[TEMPO PROCESSAR] Caminho temporal texto: "
                    f"{time.perf_counter() - _t_processar_inicio:.3f}s"
                )
                return temporal_texto

            # ====================================================
            # MÚLTIPLOS INDICADORES
            # ====================================================
            # Exemplo:
            # "qual faturamento e quantidade vendida de agosto?"
            #
            # O gráfico combinado já existente continua tendo
            # prioridade quando o usuário pede explicitamente
            # faturamento + meta + atingimento em gráfico.
            indicadores_multiplos = (
                _indicadores_explicitos_multiplos(
                    dados.pergunta
                )
            )

            if (
                len(indicadores_multiplos) >= 2
                and not (
                    usuario_pediu_grafico(
                        dados.pergunta
                    )
                    and _usuario_pediu_grafico_combinado(
                        dados.pergunta
                    )
                )
            ):
                resposta_multipla = (
                    _resposta_multiplos_indicadores(
                        dados.pergunta
                    )
                )

                if resposta_multipla is not None:
                    print(
                        f"[TEMPO PROCESSAR] "
                        f"Caminho múltiplos indicadores: "
                        f"{time.perf_counter() - _t_processar_inicio:.3f}s"
                    )
                    return resposta_multipla

            grafico = gerar_resposta_grafico(dados.pergunta)
            if grafico is not None:
                print(
                    f"[TEMPO PROCESSAR] Caminho gráfico: "
                    f"{time.perf_counter() - _t_processar_inicio:.3f}s"
                )
                return grafico

            resposta = conversar(dados.pergunta)

            print(
                f"[TEMPO PROCESSAR] Caminho texto: "
                f"{time.perf_counter() - _t_processar_inicio:.3f}s"
            )

            return {
                "tipo_resposta": "texto",
                "resposta": resposta
            }

        resultado_endpoint = _executar_no_contexto_usuario(
            dados.chat_id,
            processar
        )

        print(
            f"[TEMPO TOTAL ENDPOINT /perguntar] "
            f"{time.perf_counter() - _t_endpoint_inicio:.3f}s"
        )
        print("=" * 70)

        return resultado_endpoint

    except Exception as erro:
        return {
            "tipo_resposta": "erro",
            "erro": type(erro).__name__,
            "detalhe": str(erro),
            "resposta": "Não consegui processar a solicitação neste momento."
        }


print("Agente Power BI V35 - previsao e margem liquida corrigidas - carregado com sucesso!")
print("Autenticação Power BI: SERVICE PRINCIPAL (sem Azure CLI).")
