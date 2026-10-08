import io
import re
import functools
import requests
import pandas as pd
import streamlit as st
from fpdf import FPDF, FontFace

INTELIPOST_URL = "https://api.intelipost.com.br/api/v1/quote_by_product"
ARQUIVO_CEPS_PADRAO = "ceps_padrao.xlsx"
DEBUG_HEADER_KEY = "80fd60a5d432a29c59a2b78990ac50fa7036e17d"
COLUNA_REF = "_ref_detalhe"

# Rótulos amigáveis para campos conhecidos do retorno da Intelipost.
# Campos não mapeados são exibidos com o nome original formatado.
ROTULOS_CAMPOS = {
    "final_shipping_cost": "Frete final (cliente)",
    "provider_shipping_cost": "Custo da transportadora",
    "shipping_cost": "Custo do frete",
    "delivery_estimate_business_days": "Prazo (dias úteis)",
    "delivery_time_business_days": "Prazo (dias úteis)",
    "delivery_method_id": "ID do método de entrega",
    "delivery_method_name": "Método de entrega",
    "delivery_method_type": "Tipo do método",
    "logistic_provider_name": "Transportadora",
    "description": "Descrição",
    "delivery_note": "Observação",
    "toll_fee": "Pedágio",
    "toll": "Pedágio",
    "pedagio": "Pedágio",
    "gris": "GRIS",
    "ad_valorem": "Ad valorem",
    "advalorem": "Ad valorem",
    "icms": "ICMS",
    "insurance": "Seguro",
    "tde": "TDE",
    "trt": "TRT",
    "tda": "TDA",
    "other_fee": "Outras taxas",
    "delivery": "Taxa de entrega",
    "emex": "EMEX",
    "suframa": "SUFRAMA",
    "base_freight_cost": "Frete base (tabela)",
    "minimum_total_freight": "Frete mínimo",
    "selected_weight": "Peso considerado",
}

# Bases de cálculo das taxas no debug da Intelipost
ROTULOS_BASES = {
    "COST_OF_GOODS": "valor da mercadoria",
    "MAX_WEIGHT": "peso considerado",
    "WEIGHT": "peso",
    "CUBIC_WEIGHT": "peso cubado",
    "FREIGHT": "frete",
    "BASE_FREIGHT": "frete base",
}

# Trechos de nomes de campo que indicam um componente monetário (taxas, pedágio etc.)
PALAVRAS_CUSTO = (
    "cost", "price", "fee", "tax", "toll", "pedagio", "gris", "valorem", "icms",
    "insurance", "seguro", "taxa", "valor", "value", "surcharge", "adicional",
    "frete", "tde", "trt", "tda", "discount", "desconto", "freight", "charge",
)
PALAVRAS_NAO_CUSTO = ("id", "days", "dias", "weight", "peso", "date", "factor", "percent", "rate", "type")

UFS_BRASIL = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA",
    "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN",
    "RS", "RO", "RR", "SC", "SP", "SE", "TO"
}
MAPA_ESTADOS = {
    "ACRE": "AC",
    "ALAGOAS": "AL",
    "AMAPA": "AP",
    "AMAPÁ": "AP",
    "AMAZONAS": "AM",
    "BAHIA": "BA",
    "CEARA": "CE",
    "CEARÁ": "CE",
    "DISTRITO FEDERAL": "DF",
    "ESPIRITO SANTO": "ES",
    "ESPÍRITO SANTO": "ES",
    "GOIAS": "GO",
    "GOIÁS": "GO",
    "MARANHAO": "MA",
    "MARANHÃO": "MA",
    "MATO GROSSO": "MT",
    "MATO GROSSO DO SUL": "MS",
    "MINAS GERAIS": "MG",
    "PARA": "PA",
    "PARÁ": "PA",
    "PARAIBA": "PB",
    "PARAÍBA": "PB",
    "PARANA": "PR",
    "PARANÁ": "PR",
    "PERNAMBUCO": "PE",
    "PIAUI": "PI",
    "PIAUÍ": "PI",
    "RIO DE JANEIRO": "RJ",
    "RIO GRANDE DO NORTE": "RN",
    "RIO GRANDE DO SUL": "RS",
    "RONDONIA": "RO",
    "RONDÔNIA": "RO",
    "RORAIMA": "RR",
    "SANTA CATARINA": "SC",
    "SAO PAULO": "SP",
    "SÃO PAULO": "SP",
    "SERGIPE": "SE",
    "TOCANTINS": "TO",
}

st.set_page_config(
    page_title="Cotação Intelipost",
    page_icon="📦",
    layout="wide",
)


def inject_css():
    st.markdown(
        """
        <style>
        .main > div { padding-top: 1.5rem; }
        .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }

        .hero {
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 45%, #0f766e 100%);
            border-radius: 22px;
            padding: 28px 32px;
            color: white;
            margin-bottom: 1.2rem;
            box-shadow: 0 18px 45px rgba(15, 23, 42, 0.22);
        }
        .hero h1 { margin: 0; font-size: 2rem; line-height: 1.15; }
        .hero p { margin: 8px 0 0 0; color: rgba(255,255,255,0.86); font-size: 0.98rem; }

        .section-card {
            background: #ffffff;
            border: 1px solid rgba(15, 23, 42, 0.08);
            border-radius: 20px;
            padding: 18px 18px 14px 18px;
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
            margin-bottom: 1rem;
        }

        .small-muted {
            color: #64748b;
            font-size: 0.92rem;
        }

        .status-ok {
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: rgba(16, 185, 129, 0.12);
            color: #047857;
            font-size: 0.85rem;
            font-weight: 600;
        }

        .status-warn {
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: rgba(245, 158, 11, 0.14);
            color: #b45309;
            font-size: 0.85rem;
            font-weight: 600;
        }

        .highlight-option {
            background: linear-gradient(180deg, rgba(15, 118, 110, 0.10), rgba(15, 118, 110, 0.04));
            border: 1px solid rgba(15, 118, 110, 0.20);
            border-radius: 16px;
            padding: 14px 14px 10px 14px;
            margin-top: 28px;
        }

        .highlight-badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 999px;
            background: rgba(15, 118, 110, 0.12);
            color: #0f766e;
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 0.02em;
            margin-bottom: 8px;
        }

        .highlight-title {
            font-size: 0.96rem;
            font-weight: 700;
            color: #0f172a;
            margin-bottom: 4px;
        }

        .highlight-text {
            font-size: 0.9rem;
            color: #475569;
            line-height: 1.4;
            margin-bottom: 8px;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 16px;
            overflow: hidden;
            border: 1px solid rgba(15, 23, 42, 0.08);
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius: 12px !important;
            font-weight: 600 !important;
            height: 42px;
        }

        .stTextInput > div > div > input,
        .stTextArea textarea,
        .stNumberInput input,
        div[data-baseweb="select"] > div {
            border-radius: 12px !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_session():
    defaults = {
        "api_key": "",
        "modo_tabela": "Tabelas vigentes",
        "todas_opcoes": [],
        "ultimo_payload": None,
        "ultimo_headers": None,
        "logs_execucao": [],
        "erros_cotacao": [],
        "detalhes_opcoes": {},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def montar_products(peso: float, largura: float, altura: float, comprimento: float, valor: float) -> list[dict]:
    return [
        {
            "weight": peso,
            "cost_of_goods": valor,
            "width": largura,
            "height": altura,
            "length": comprimento,
            "quantity": 1,
            "sku_id": "SKU123",
            "product_category": "Bebidas",
        },
    ]


def montar_payload(origin_cep: str, dest_cep: str, peso: float, largura: float, altura: float, comprimento: float, valor: float) -> dict:
    return {
        "origin_zip_code": origin_cep,
        "destination_zip_code": dest_cep,
        "quoting_mode": "DYNAMIC_BOX_ALL_ITEMS",
        "products": montar_products(peso, largura, altura, comprimento, valor),
        "additional_information": {
            "sales_channel": "meu_canal_de_vendas",
            "client_type": "gold",
            "rule_tags": ["Agendado", "Linha_Branca"],
        },
        "identification": {
            "session": "04e5bdf7ed15e571c0265c18333b6fdf1434658753",
            "ip": "000.000.000.000",
            "page_name": "carrinho",
            "url": "http://www.intelipost.com.br/checkout/cart/",
        },
    }


def montar_headers(api_key: str, usar_tabelas_rascunho: bool) -> dict:
    # O header debug é sempre enviado: sem ele a API não devolve delivery_options_debug,
    # que é onde vem a composição do frete (frete base, taxas, impostos e regras).
    headers = {
        "Content-Type": "application/json",
        "api-key": api_key,
        "debug": DEBUG_HEADER_KEY,
    }
    if usar_tabelas_rascunho:
        headers["logistic-contract-mode"] = "DRAFT"
    return headers


def cotar_frete_intelipost(
    origin_cep: str,
    dest_cep: str,
    peso: float,
    largura: float,
    altura: float,
    comprimento: float,
    valor: float,
    api_key: str,
    usar_tabelas_rascunho: bool = False,
) -> dict:
    headers = montar_headers(api_key, usar_tabelas_rascunho)
    payload = montar_payload(origin_cep, dest_cep, peso, largura, altura, comprimento, valor)
    st.session_state["ultimo_payload"] = payload
    st.session_state["ultimo_headers"] = headers
    response = requests.post(INTELIPOST_URL, json=payload, headers=headers, timeout=20)
    response.raise_for_status()
    return response.json()


def extrair_opcoes_frete(dest_cep: str, resposta: dict) -> list[dict]:
    opcoes = []
    content = resposta.get("content", {})
    delivery_options = content.get("delivery_options", [])

    chaves_fora_metadados = {"delivery_options", "delivery_options_debug", "rejected_delivery_options_debug"}
    metadados_cotacao = {k: v for k, v in content.items() if k not in chaves_fora_metadados}

    # O debug traz todos os métodos avaliados; liga cada opção ao seu pelo método + contrato
    debug_por_chave = {}
    for dbg in content.get("delivery_options_debug") or []:
        contrato = (dbg.get("freight") or {}).get("logistic_contract_id")
        debug_por_chave.setdefault((dbg.get("delivery_method_id"), contrato), dbg)
        debug_por_chave.setdefault((dbg.get("delivery_method_id"), None), dbg)

    for opt in delivery_options:
        debug_opcao = debug_por_chave.get(
            (opt.get("delivery_method_id"), opt.get("logistic_contract_id"))
        ) or debug_por_chave.get((opt.get("delivery_method_id"), None))

        ref = str(len(st.session_state.detalhes_opcoes))
        st.session_state.detalhes_opcoes[ref] = {
            "destino": dest_cep,
            "opcao": opt,
            "debug": debug_opcao,
            "cotacao": metadados_cotacao,
        }

        transportadora = opt.get("description") or opt.get("logistic_provider_name") or "N/I"
        prazo_dias = (
            opt.get("delivery_estimate_business_days")
            or opt.get("delivery_time_business_days")
            or opt.get("delivery_time")
            or 0
        )
        valor_frete = opt.get("final_shipping_cost") or opt.get("shipping_cost") or 0.0
        id_cotacao = opt.get("id") or content.get("id") or resposta.get("id") or ""

        opcoes.append(
            {
                "destino": dest_cep,
                "transportadora": transportadora,
                "prazo_dias": int(prazo_dias),
                "valor_frete": float(valor_frete),
                "id_cotacao": str(id_cotacao),
                COLUNA_REF: ref,
            }
        )
    return opcoes


def ler_ceps_de_excel_streamlit(arquivo_excel) -> list[str]:
    df_ceps = pd.read_excel(arquivo_excel, usecols="A", dtype=str)
    primeira_coluna = df_ceps.columns[0]
    ceps_series = df_ceps[primeira_coluna].dropna().astype(str).str.strip()
    ceps_unicos = pd.Series(ceps_series.unique())
    return [c for c in ceps_unicos if c]


def parse_float_br(valor: str) -> float:
    if isinstance(valor, str) and valor.count(",") == 1 and valor.count(".") > 1:
        return float(str(valor).replace(".", "").replace(",", "."))
    return float(str(valor).replace(",", "."))


def normalizar_lista_ceps(texto: str, ceps_excel: list[str]) -> list[str]:
    ceps_digitados = []
    if texto.strip():
        bruto = texto.replace("\n", ",").replace(";", ",")
        ceps_digitados = [c.strip() for c in bruto.split(",") if c.strip()]
    return list(dict.fromkeys(ceps_digitados + ceps_excel))


def gerar_excel_bytes_abas(abas: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in abas.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


MIME_EXCEL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
TITULOS_SECOES_PDF = {
    "cotacoes": "Cotações",
    "composicao_do_frete": "Composição do frete",
    "ceps_padrao": "CEPs padrão",
    "execucao": "Execução",
    "ceps_com_erro": "CEPs com erro",
}


def texto_pdf(valor) -> str:
    """As fontes padrão do PDF só cobrem latin-1; troca o que não couber (ex.: emojis)."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if isinstance(valor, float):
        valor = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(valor).encode("latin-1", "replace").decode("latin-1")


def adicionar_tabela_pdf(pdf: FPDF, df: pd.DataFrame):
    if df.empty:
        pdf.set_font("Helvetica", "I", 8)
        pdf.cell(0, 5, "Sem registros.", new_x="LMARGIN", new_y="NEXT")
        return

    linhas = [[texto_pdf(v) for v in registro] for registro in df.itertuples(index=False)]
    cabecalho = [texto_pdf(c) for c in df.columns]
    # Largura de cada coluna proporcional ao conteúdo, limitada para textos longos
    pesos = [
        min(max([len(cabecalho[i])] + [len(linha[i]) for linha in linhas]), 40) + 2
        for i in range(len(cabecalho))
    ]
    fonte = 7 if len(cabecalho) <= 10 else 6
    pdf.set_font("Helvetica", "", fonte)
    with pdf.table(
        col_widths=pesos,
        line_height=fonte * 0.6,
        text_align="LEFT",
        repeat_headings=1,
        headings_style=FontFace(emphasis="BOLD", fill_color=(230, 236, 245)),
    ) as tabela:
        tabela.row(cabecalho)
        for linha in linhas:
            tabela.row(linha)


def gerar_pdf_bytes(secoes: list[dict], titulo: str) -> bytes:
    """Gera um PDF (A4 paisagem) a partir de seções.

    Cada seção é {"titulo": str, "nivel": 1|2, "subtitulo": str opcional, "df": DataFrame opcional}.
    Nível 1 é um bloco de agrupamento (ex.: um CEP); nível 2 é o título de uma tabela.
    """
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.set_margins(10, 10, 10)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, texto_pdf(titulo), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 8)
    pdf.cell(0, 5, texto_pdf(f"Gerado em {pd.Timestamp.now():%d/%m/%Y %H:%M}"), new_x="LMARGIN", new_y="NEXT")

    for secao in secoes:
        if secao.get("nivel", 2) == 1:
            # Não deixa o cabeçalho do grupo sozinho no pé da página
            if pdf.get_y() > pdf.h - 45:
                pdf.add_page()
            pdf.ln(5)
            pdf.set_font("Helvetica", "B", 12)
            pdf.set_fill_color(31, 78, 140)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 8, texto_pdf(" " + secao["titulo"]), fill=True, new_x="LMARGIN", new_y="NEXT")
            # Volta às cores padrão para não vazar o azul nas tabelas seguintes
            pdf.set_text_color(0, 0, 0)
            pdf.set_fill_color(255, 255, 255)
            if secao.get("subtitulo"):
                pdf.set_font("Helvetica", "", 8)
                pdf.cell(0, 5, texto_pdf(secao["subtitulo"]), new_x="LMARGIN", new_y="NEXT")
        else:
            # Título de tabela não fica órfão no pé da página
            if pdf.get_y() > pdf.h - 30:
                pdf.add_page()
            pdf.ln(3)
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 6, texto_pdf(secao["titulo"]), new_x="LMARGIN", new_y="NEXT")

        if secao.get("df") is not None:
            adicionar_tabela_pdf(pdf, secao["df"])

    return bytes(pdf.output())


def gerar_pdf_bytes_abas(abas: dict[str, pd.DataFrame], titulo: str) -> bytes:
    """PDF com uma tabela por aba, no mesmo espírito do Excel."""
    return gerar_pdf_bytes(
        [
            {"titulo": TITULOS_SECOES_PDF.get(nome, nome.replace("_", " ").capitalize()), "df": df}
            for nome, df in abas.items()
        ],
        titulo,
    )


def montar_secoes_pdf_cotacoes(df_exibicao: pd.DataFrame, detalhes: dict) -> list[dict]:
    """Agrupa o PDF de cotações por CEP de destino: cotações e composição do frete de cada CEP."""
    colunas_local = [c for c in ("estado", "cidade", "tipo") if c in df_exibicao.columns]
    df_comp = montar_composicao_pdf(df_exibicao, detalhes)
    secoes = []

    for destino, df_cep in df_exibicao.groupby("destino", sort=False):
        local = df_cep.iloc[0]
        cidade_uf = " / ".join(
            str(local[c]) for c in ("cidade", "estado") if c in colunas_local and pd.notna(local[c]) and str(local[c])
        )
        titulo = f"CEP {destino}" + (f" - {cidade_uf}" if cidade_uf else "")
        if "tipo" in colunas_local and pd.notna(local["tipo"]) and str(local["tipo"]):
            titulo += f" ({local['tipo']})"

        valores = df_cep["valor_frete"].astype(float)
        prazos = df_cep["prazo_dias"].astype(float)
        subtitulo = (
            f"{len(df_cep)} opção(ões) · menor frete {formatar_moeda_br(valores.min())} "
            f"({df_cep.loc[valores.idxmin(), 'transportadora']}) · "
            f"maior frete {formatar_moeda_br(valores.max())} · menor prazo {formatar_dias(prazos.min())}"
        )

        secoes.append({"titulo": titulo, "nivel": 1, "subtitulo": subtitulo})
        secoes.append(
            {
                "titulo": "Cotações",
                "df": df_cep.drop(columns=[COLUNA_REF, "destino", *colunas_local], errors="ignore"),
            }
        )
        comp_cep = df_comp[df_comp["destino"] == destino].drop(columns=["destino"])
        if not comp_cep.empty:
            secoes.append({"titulo": "Composição do frete", "df": comp_cep})

    return secoes


def botoes_download(abas: dict[str, pd.DataFrame], nome_arquivo: str, rotulo: str, titulo_pdf: str,
                    secoes_pdf: list[dict] | None = None, chave: str = "", largura_total: bool = False):
    """Botões lado a lado de Excel e PDF para o mesmo conteúdo."""
    col_excel, col_pdf = st.columns(2) if largura_total else st.columns([1, 1, 3])[:2]
    with col_excel:
        st.download_button(
            f"💾 {rotulo} (Excel)",
            data=gerar_excel_bytes_abas(abas),
            file_name=f"{nome_arquivo}.xlsx",
            mime=MIME_EXCEL,
            use_container_width=True,
            key=f"excel_{chave or nome_arquivo}",
        )
    with col_pdf:
        try:
            if secoes_pdf is not None:
                pdf_bytes = gerar_pdf_bytes(secoes_pdf, titulo_pdf)
            else:
                pdf_bytes = gerar_pdf_bytes_abas(abas, titulo_pdf)
        except Exception as e:
            st.warning(f"Não foi possível gerar o PDF: {e}")
        else:
            st.download_button(
                f"📄 {rotulo} (PDF)",
                data=pdf_bytes,
                file_name=f"{nome_arquivo}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"pdf_{chave or nome_arquivo}",
            )


def montar_composicao_pdf(df_exibicao: pd.DataFrame, detalhes: dict) -> pd.DataFrame:
    """Composição do frete em formato longo (uma linha por componente), legível em PDF."""
    linhas = []
    for _, linha in df_exibicao.iterrows():
        detalhe = detalhes.get(linha.get(COLUNA_REF), {})
        for comp in componentes_frete(detalhe.get("debug")):
            linhas.append(
                {
                    "destino": linha.get("destino"),
                    "transportadora": linha.get("transportadora"),
                    "componente": comp["Componente"],
                    "tipo": comp["Tipo"],
                    "regra": comp["Regra"],
                    "valor": comp["Valor"],
                }
            )
    return pd.DataFrame(linhas, columns=["destino", "transportadora", "componente", "tipo", "regra", "valor"])


def montar_planilhas_cotacoes(df_exibicao: pd.DataFrame, detalhes: dict) -> dict[str, pd.DataFrame]:
    """Gera a aba de cotações com uma coluna por taxa/componente de valor e a aba com todos os campos."""
    colunas_custo = {}  # chave original -> nome da coluna no Excel
    linhas_custos = []
    linhas_detalhes = []

    for _, linha in df_exibicao.iterrows():
        detalhe = detalhes.get(linha.get(COLUNA_REF), {})
        campos = achatar_campos(detalhe.get("opcao", {}))
        campos_debug = achatar_campos(detalhe.get("debug") or {}, "debug")

        custos_linha = {}
        for componente in componentes_frete(detalhe.get("debug")):
            chave = f"componente:{componente['Componente']}"
            if chave not in colunas_custo:
                nome = componente["Componente"]
                if nome in colunas_custo.values() or nome in df_exibicao.columns:
                    nome = f"{nome} ({componente['Tipo'].lower()})"
                colunas_custo[chave] = nome
            custos_linha[colunas_custo[chave]] = custos_linha.get(colunas_custo[chave], 0.0) + componente["Valor"]

        for chave, valor in campos + campos_debug:
            if not chave.startswith("debug") and eh_campo_custo(chave, valor):
                if chave not in colunas_custo:
                    nome = rotulo_campo(chave)
                    if nome in colunas_custo.values() or nome in df_exibicao.columns:
                        nome = f"{nome} ({chave})"
                    colunas_custo[chave] = nome
                custos_linha[colunas_custo[chave]] = float(valor)

            linhas_detalhes.append(
                {
                    "destino": linha.get("destino"),
                    "transportadora": linha.get("transportadora"),
                    "id_cotacao": linha.get("id_cotacao"),
                    "campo": rotulo_campo(chave),
                    "chave": chave,
                    "valor": valor,
                }
            )
        linhas_custos.append(custos_linha)

    df_cotacoes = df_exibicao.drop(columns=[COLUNA_REF], errors="ignore").reset_index(drop=True)
    df_custos = pd.DataFrame(linhas_custos, columns=list(colunas_custo.values()))
    df_cotacoes = pd.concat([df_cotacoes, df_custos], axis=1)

    df_detalhes = pd.DataFrame(
        linhas_detalhes,
        columns=["destino", "transportadora", "id_cotacao", "campo", "chave", "valor"],
    )
    return {"cotacoes": df_cotacoes, "detalhes_cotacao": df_detalhes}


def normalizar_uf(valor: str) -> str:
    texto = str(valor).strip().upper()
    if not texto:
        return ""
    if texto in UFS_BRASIL:
        return texto
    return MAPA_ESTADOS.get(texto, "")


@functools.lru_cache(maxsize=1)
def carregar_ceps_padrao() -> pd.DataFrame:
    df = pd.read_excel(ARQUIVO_CEPS_PADRAO, dtype=str)
    colunas_normalizadas = {str(col).strip().lower(): col for col in df.columns}

    possiveis = {
        "estado": ["estado", "uf"],
        "cidade": ["cidade", "municipio", "município"],
        "cep": ["cep", "zip_code", "codigo_postal"],
        "tipo": ["tipo", "categoria", "perfil"],
    }

    renomear = {}
    for destino, aliases in possiveis.items():
        for alias in aliases:
            if alias in colunas_normalizadas:
                renomear[colunas_normalizadas[alias]] = destino
                break

    df = df.rename(columns=renomear)

    obrigatorias = ["estado", "cidade", "cep", "tipo"]
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        raise ValueError(
            "A planilha de CEPs padrão precisa ter as colunas: estado, cidade, cep e tipo. "
            f"Faltando: {', '.join(faltantes)}"
        )

    df = df[obrigatorias].copy()
    df["estado"] = df["estado"].apply(normalizar_uf)
    df["cidade"] = df["cidade"].astype(str).str.strip()
    df["cep"] = df["cep"].astype(str).str.strip()
    df["tipo"] = df["tipo"].astype(str).str.lower().str.strip()

    df = df.dropna(subset=["cep"])
    df = df[df["cep"] != ""]
    df = df[df["estado"].isin(UFS_BRASIL)]
    df = df.drop_duplicates(subset=["cep"])

    if df.empty:
        raise ValueError(
            "Nenhuma UF válida foi encontrada na planilha de CEPs padrão. "
            "Use siglas válidas como MG, SP, RJ ou nomes completos de estados brasileiros."
        )

    return df


def obter_ceps_padrao() -> list[str]:
    df = carregar_ceps_padrao()
    return df["cep"].dropna().astype(str).str.strip().unique().tolist()


def obter_info_ceps_padrao() -> pd.DataFrame:
    return carregar_ceps_padrao().copy()


def extrair_mensagem_erro_response(resp: requests.Response) -> str:
    try:
        erro_body = resp.json()
        if isinstance(erro_body, dict):
            return (
                str(erro_body.get("message"))
                if erro_body.get("message")
                else str(erro_body.get("error"))
                if erro_body.get("error")
                else str(erro_body.get("detail"))
                if erro_body.get("detail")
                else str(erro_body)
            )
        return str(erro_body)
    except Exception:
        return resp.text.strip() if resp.text else "Resposta sem conteúdo"


def formatar_moeda_br(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_dias(valor: float) -> str:
    if pd.isna(valor):
        return "0 dia(s)"
    if float(valor).is_integer():
        return f"{int(valor)} dia(s)"
    return f"{valor:.2f} dia(s)".replace(".", ",")


def rotulo_campo(caminho: str) -> str:
    partes = []
    for parte in caminho.split("."):
        nome, _, indice = parte.partition("[")
        rotulo = ROTULOS_CAMPOS.get(nome.lower(), nome.replace("_", " ").strip().capitalize())
        partes.append(f"{rotulo} [{indice}" if indice else rotulo)
    return " › ".join(partes)


def achatar_campos(obj, prefixo: str = "") -> list[tuple[str, object]]:
    """Transforma o JSON aninhado em pares (caminho, valor) apenas com valores simples."""
    itens = []
    if isinstance(obj, dict):
        for chave, valor in obj.items():
            caminho = f"{prefixo}.{chave}" if prefixo else str(chave)
            itens.extend(achatar_campos(valor, caminho))
    elif isinstance(obj, list):
        if all(not isinstance(v, (dict, list)) for v in obj):
            itens.append((prefixo, ", ".join(str(v) for v in obj)))
        else:
            for i, valor in enumerate(obj):
                itens.extend(achatar_campos(valor, f"{prefixo}[{i}]"))
    else:
        itens.append((prefixo, obj))
    return itens


def eh_campo_custo(caminho: str, valor) -> bool:
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return False
    chave = caminho.split(".")[-1].split("[")[0].lower()
    tokens = set(re.split(r"[^a-z0-9]+", chave))
    if tokens & set(PALAVRAS_NAO_CUSTO):
        return False
    return any(p in chave for p in PALAVRAS_CUSTO)


def formatar_valor_campo(valor) -> str:
    if valor is None:
        return "—"
    if isinstance(valor, bool):
        return "Sim" if valor else "Não"
    return str(valor)


def descrever_regra_taxa(regra: dict) -> str:
    """Resume a regra contratual de uma taxa (ex.: '0,2% sobre valor da mercadoria, mín. R$ 4,00')."""
    if not regra:
        return ""
    base = ROTULOS_BASES.get(str(regra.get("base")), str(regra.get("base") or "").lower())
    partes = []
    if regra.get("fixed") is not None:
        partes.append(f"fixo {formatar_moeda_br(float(regra['fixed']))}")
    if regra.get("value") is not None:
        valor = float(regra["value"])
        if regra.get("type") == "PERCENT":
            partes.append(f"{valor:g}%".replace(".", ",") + (f" sobre {base}" if base else ""))
        else:
            texto = formatar_moeda_br(valor)
            if regra.get("fraction"):
                texto += f" a cada {float(regra['fraction']):g}"
            partes.append(texto + (f" ({base})" if base else ""))
    if regra.get("min"):
        partes.append(f"mín. {formatar_moeda_br(float(regra['min']))}")
    if regra.get("max"):
        partes.append(f"máx. {formatar_moeda_br(float(regra['max']))}")
    return ", ".join(partes)


def componentes_frete(debug: dict | None) -> list[dict]:
    """Monta a composição do frete (frete base, taxas e impostos) a partir do delivery_options_debug."""
    if not debug:
        return []
    frete = debug.get("freight") or {}
    componentes = []
    # base_price é o frete base efetivo usado na soma; base_freight_cost é o valor puro da faixa
    # da tabela e pode ser menor (ex.: quando há excedente de peso)
    frete_base = frete.get("base_price", frete.get("base_freight_cost"))
    if frete_base is not None:
        regra = (
            f"região {frete.get('destination_geographic_identifier') or '—'}, "
            f"peso {frete.get('selected_weight', '—')} kg"
        )
        valor_tabela = frete.get("base_freight_cost")
        if valor_tabela is not None and abs(float(valor_tabela) - float(frete_base)) >= 0.01:
            regra += f", faixa da tabela {formatar_moeda_br(float(valor_tabela))}"
        componentes.append(
            {
                "Componente": "Frete base",
                "Tipo": "Frete",
                "Regra": regra,
                "Base de cálculo": None,
                "Valor": float(frete_base),
            }
        )
    for taxa in debug.get("fees") or []:
        componentes.append(
            {
                "Componente": rotulo_campo(str(taxa.get("name", ""))),
                "Tipo": "Taxa",
                "Regra": descrever_regra_taxa(taxa.get("logistic_fee") or {}),
                "Base de cálculo": taxa.get("base_value"),
                "Valor": float(taxa.get("final_value") or 0.0),
            }
        )
    for imposto in debug.get("taxes") or []:
        modo = {"INSIDE": "por dentro", "OUTSIDE": "por fora"}.get(imposto.get("mode"), imposto.get("mode") or "")
        regra = f"{float(imposto.get('percentage') or 0):g}%".replace(".", ",") + (f" {modo}" if modo else "")
        componentes.append(
            {
                "Componente": rotulo_campo(str(imposto.get("name", ""))),
                "Tipo": "Imposto",
                "Regra": regra,
                "Base de cálculo": imposto.get("base_value"),
                "Valor": float(imposto.get("final_value") or 0.0),
            }
        )
    return componentes


def render_composicao_debug(debug: dict):
    frete = debug.get("freight") or {}
    componentes = componentes_frete(debug)

    st.markdown("**Composição do frete**")
    if componentes:
        df_comp = pd.DataFrame(componentes)
        df_comp["Base de cálculo"] = df_comp["Base de cálculo"].map(
            lambda v: "—" if v is None or pd.isna(v) else f"{float(v):g}".replace(".", ",")
        )
        df_comp["Valor"] = df_comp["Valor"].map(formatar_moeda_br)
        st.dataframe(df_comp, use_container_width=True, hide_index=True)

        soma = sum(c["Valor"] for c in componentes)
        custo = frete.get("provider_shipping_cost")
        final = frete.get("final_shipping_cost")
        resumo = f"Soma dos componentes: **{formatar_moeda_br(soma)}**"
        if custo is not None:
            resumo += f" · Custo transportadora: **{formatar_moeda_br(float(custo))}**"
        if final is not None:
            resumo += f" · Frete final: **{formatar_moeda_br(float(final))}**"
        st.caption(resumo)
    else:
        st.info("O debug desta opção não trouxe frete base, taxas ou impostos.")

    regras = frete.get("quote_rules") or []
    if regras:
        st.markdown("**Regras de cotação aplicadas**")
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "ID": r.get("id"),
                        "Regra": r.get("name"),
                        "Ação": r.get("action_localized") or r.get("action"),
                        "Valor": r.get("value"),
                    }
                    for r in regras
                ]
            ),
            use_container_width=True,
            hide_index=True,
        )

    mensagens = (debug.get("info_messages") or []) + (debug.get("restriction_messages") or [])
    if mensagens:
        with st.expander("Mensagens do cálculo", expanded=False):
            for msg in mensagens:
                st.markdown(f"- {msg}")

    parametros = {k: v for k, v in frete.items() if not isinstance(v, (list, dict))}
    with st.expander("Parâmetros do cálculo (contrato, peso, cubagem)", expanded=False):
        st.dataframe(
            pd.DataFrame(
                [{"Campo": rotulo_campo(k), "Chave": k, "Valor": formatar_valor_campo(v)} for k, v in parametros.items()]
            ),
            use_container_width=True,
            hide_index=True,
        )
        if frete.get("volume_infos"):
            st.markdown("Volumes")
            st.dataframe(pd.json_normalize(frete["volume_infos"]), use_container_width=True, hide_index=True)


def render_detalhe_opcao(detalhe: dict):
    opcao = detalhe["opcao"]
    transportadora = opcao.get("description") or opcao.get("logistic_provider_name") or "N/I"
    prazo = (
        opcao.get("delivery_estimate_business_days")
        or opcao.get("delivery_time_business_days")
        or opcao.get("delivery_time")
    )
    frete_final = opcao.get("final_shipping_cost")
    custo_transp = opcao.get("provider_shipping_cost")

    st.markdown(f"#### 🔎 {transportadora} → CEP {detalhe['destino']}")

    metricas = st.columns(4)
    metricas[0].metric(
        "Frete final",
        formatar_moeda_br(float(frete_final)) if frete_final is not None else "—",
    )
    metricas[1].metric(
        "Custo transportadora",
        formatar_moeda_br(float(custo_transp)) if custo_transp is not None else "—",
    )
    metricas[2].metric("Prazo", formatar_dias(float(prazo)) if prazo is not None else "—")
    metricas[3].metric(
        "Método",
        opcao.get("delivery_method_name") or opcao.get("delivery_method_type") or "—",
    )

    campos = achatar_campos(opcao)
    debug = detalhe.get("debug")

    if debug:
        render_composicao_debug(debug)
    else:
        custos = [(c, v) for c, v in campos if eh_campo_custo(c, v)]
        st.markdown("**Composição de valores e taxas**")
        st.warning(
            "A API não retornou o detalhamento (delivery_options_debug) para esta opção, "
            "então só os valores totais estão disponíveis."
        )
        if custos:
            df_custos = pd.DataFrame(
                [{"Componente": rotulo_campo(c), "Valor": formatar_moeda_br(float(v))} for c, v in custos]
            )
            st.dataframe(df_custos, use_container_width=True, hide_index=True)

    with st.expander("Todos os campos retornados na opção", expanded=False):
        df_campos = pd.DataFrame(
            [{"Campo": rotulo_campo(c), "Chave": c, "Valor": formatar_valor_campo(v)} for c, v in campos]
        )
        st.dataframe(df_campos, use_container_width=True, hide_index=True)

    # Listas de objetos (ex.: volumes, taxas detalhadas) ficam melhores como tabela própria
    for chave, valor in opcao.items():
        if isinstance(valor, list) and valor and all(isinstance(v, dict) for v in valor):
            with st.expander(f"{rotulo_campo(chave)} ({len(valor)})", expanded=False):
                st.dataframe(pd.json_normalize(valor), use_container_width=True, hide_index=True)

    if detalhe.get("cotacao"):
        with st.expander("Dados gerais da cotação", expanded=False):
            st.json(detalhe["cotacao"])

    with st.expander("JSON bruto da opção", expanded=False):
        st.json(opcao)

    if debug:
        with st.expander("JSON bruto do debug (composição)", expanded=False):
            st.json(debug)


def render_header():
    st.markdown(
        """
        <div class="hero">
            <h1>📦 Cotação de Frete Intelipost</h1>
            <p>Aplicação Streamlit com interface moderna, importação de CEPs por Excel, suporte a CEPs padrão,
            análise de médias e API Key mantida apenas na sessão atual do navegador.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar():
    st.sidebar.title("⚙️ Configurações")
    st.sidebar.caption("A API Key não é salva em arquivo. Ela fica apenas na sessão atual.")

    st.sidebar.text_input(
        "API Key Intelipost",
        type="password",
        key="api_key",
        help="A chave fica somente nesta sessão do navegador.",
    )

    st.sidebar.selectbox(
        "Modo de cotação",
        options=["Tabelas vigentes", "Tabelas em rascunho"],
        key="modo_tabela",
        help="No modo rascunho, os headers de debug e logistic-contract-mode=DRAFT são enviados.",
    )

    col1, col2 = st.sidebar.columns(2)
    with col1:
        if st.button("Validar", use_container_width=True):
            if st.session_state.api_key.strip():
                st.sidebar.success("API Key disponível na sessão.")
            else:
                st.sidebar.error("Informe a API Key.")
    with col2:
        if st.button("Limpar", use_container_width=True):
            st.session_state.api_key = ""
            st.sidebar.info("API Key removida da sessão.")

    st.sidebar.markdown("---")
    if st.session_state.api_key.strip():
        st.sidebar.markdown('<span class="status-ok">API Key carregada</span>', unsafe_allow_html=True)
    else:
        st.sidebar.markdown('<span class="status-warn">API Key não informada</span>', unsafe_allow_html=True)

    if st.session_state.modo_tabela == "Tabelas em rascunho":
        st.sidebar.info("Modo rascunho ativo: headers 'debug' e 'logistic-contract-mode=DRAFT' serão enviados.")
    else:
        st.sidebar.caption("Modo vigente ativo: requisição padrão.")


def render_metricas(df: pd.DataFrame):
    total_ceps = df["destino"].nunique() if not df.empty else 0
    total_opcoes = len(df)
    menor_frete = df["valor_frete"].min() if not df.empty else 0.0
    maior_frete = df["valor_frete"].max() if not df.empty else 0.0
    valor_medio = df["valor_frete"].mean() if not df.empty else 0.0
    menor_prazo = df["prazo_dias"].min() if not df.empty else 0.0
    maior_prazo = df["prazo_dias"].max() if not df.empty else 0.0
    prazo_medio = df["prazo_dias"].mean() if not df.empty else 0.0

    linha1 = st.columns(4)
    linha1[0].metric("CEPs cotados", total_ceps)
    linha1[1].metric("Opções retornadas", total_opcoes)
    linha1[2].metric("Menor frete", formatar_moeda_br(menor_frete))
    linha1[3].metric("Maior frete", formatar_moeda_br(maior_frete))

    linha2 = st.columns(4)
    linha2[0].metric("Frete médio", formatar_moeda_br(valor_medio))
    linha2[1].metric("Menor prazo", formatar_dias(menor_prazo))
    linha2[2].metric("Maior prazo", formatar_dias(maior_prazo))
    linha2[3].metric("Prazo médio", formatar_dias(prazo_medio))


def render_debug_area():
    with st.expander("Debug técnico", expanded=False):
        if st.session_state.get("ultimo_payload"):
            st.write("Último payload enviado")
            st.json(st.session_state["ultimo_payload"])
        else:
            st.info("Nenhum payload registrado ainda.")

        if st.session_state.get("ultimo_headers"):
            st.write("Últimos headers enviados")
            st.json(st.session_state["ultimo_headers"])
        else:
            st.info("Nenhum header registrado ainda.")


def main():
    inject_css()
    init_session()
    render_sidebar()
    render_header()

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("Parâmetros da cotação")
    st.markdown(
        "<div class='small-muted'>Defina o CEP de origem e os CEPs de destino que serão utilizados "
        "na cotação (texto, planilha e/ou base padrão).</div>",
        unsafe_allow_html=True,
    )

    col_cep_origem, col_toggle_padrao = st.columns([2, 1])

    with col_cep_origem:
        origin_cep = st.text_input("CEP origem", placeholder="35620-000")

    with col_toggle_padrao:
        st.markdown(
            """
            <div class="highlight-option">
                <div class="highlight-badge">BASE COMPLETA</div>
                <div class="highlight-title">Utilize rapidamente toda a base padrão</div>
                <div class="highlight-text">
                    Ative esta opção para adicionar automaticamente todos os CEPs padrão na cotação.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        usar_ceps_padrao = st.checkbox("**Utilizar os CEPs Padrão**")

    col_destino, col_excel, col_padrao = st.columns([2.2, 1.5, 1.3])

    ceps_excel = []
    with col_destino:
        ceps_destino_str = st.text_area(
            "CEPs destino (manual)",
            placeholder="Digite separados por vírgula, ponto e vírgula ou quebra de linha",
            height=120,
        )

    with col_excel:
        uploaded_excel = st.file_uploader(
            "Importar CEPs por Excel",
            type=["xlsx", "xls"],
            help="Os CEPs devem estar na coluna A.",
        )
        if uploaded_excel is not None:
            try:
                ceps_excel = ler_ceps_de_excel_streamlit(uploaded_excel)
                if ceps_excel:
                    st.success(f"{len(ceps_excel)} CEP(s) carregado(s) da planilha.")
                    preview = ", ".join(ceps_excel[:15])
                    if len(ceps_excel) > 15:
                        preview += " ..."
                    st.caption(preview)
                else:
                    st.warning("Nenhum CEP encontrado na coluna A.")
            except Exception as e:
                st.error(f"Erro ao ler Excel: {e}")

    with col_padrao:
        st.markdown("**Base de CEPs padrão**")
        try:
            df_ceps_padrao_base = obter_info_ceps_padrao()
            st.caption(f"{len(df_ceps_padrao_base)} CEP(s) disponíveis.")
            botoes_download(
                {"ceps_padrao": df_ceps_padrao_base.sort_values(["estado", "tipo", "cidade", "cep"])},
                nome_arquivo="ceps_padrao_utilizados",
                rotulo="Base padrão",
                titulo_pdf="Base de CEPs padrão",
                largura_total=True,
            )
        except Exception as e:
            st.warning(
                "Não foi possível carregar a planilha de CEPs padrão. "
                f"Verifique o arquivo '{ARQUIVO_CEPS_PADRAO}'. Detalhe: {e}"
            )

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("Dimensões e valor do item")
    st.markdown(
        "<div class='small-muted'>Informe as dimensões e o valor total do item utilizado na cotação.</div>",
        unsafe_allow_html=True,
    )

    col_peso, col_larg, col_alt, col_comp, col_valor = st.columns(5)
    with col_peso:
        peso = st.text_input("Peso (kg)", value="1,0")
    with col_larg:
        largura = st.text_input("Largura (cm)", value="10")
    with col_alt:
        altura = st.text_input("Altura (cm)", value="10")
    with col_comp:
        comprimento = st.text_input("Comprimento (cm)", value="10")
    with col_valor:
        valor = st.text_input("Valor do item (R$)", value="100,00")

    st.markdown("</div>", unsafe_allow_html=True)

    todos_ceps = normalizar_lista_ceps(ceps_destino_str, ceps_excel)

    if usar_ceps_padrao:
        try:
            ceps_padrao = obter_ceps_padrao()
            todos_ceps = list(dict.fromkeys(todos_ceps + ceps_padrao))
        except Exception as e:
            st.error(f"Erro ao carregar CEPs padrão: {e}")

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    col_status_ceps, col_status_modo = st.columns([1.2, 3.8])
    with col_status_ceps:
        st.metric("Total de CEPs prontos para cotação", len(todos_ceps))

    usar_tabelas_rascunho = st.session_state.modo_tabela == "Tabelas em rascunho"
    with col_status_modo:
        if usar_tabelas_rascunho:
            st.caption("Modo selecionado: tabelas em rascunho (headers debug + logistic-contract-mode=DRAFT).")
        else:
            st.caption("Modo selecionado: tabelas vigentes (requisição padrão).")

    cbtn1, cbtn2, cbtn3 = st.columns([1, 1, 1])
    with cbtn1:
        st.empty()
    with cbtn2:
        cotar = st.button("🚀 Cotar fretes", use_container_width=True)
    with cbtn3:
        limpar_resultados = st.button("🧹 Limpar resultados", use_container_width=True)

    if limpar_resultados:
        st.session_state.todas_opcoes = []
        st.session_state.logs_execucao = []
        st.session_state.erros_cotacao = []
        st.session_state.detalhes_opcoes = {}
        st.info("Resultados removidos da sessão.")

    aviso_processamento = st.empty()
    status_execucao = st.empty()
    progresso_container = st.empty()

    if cotar:
        if not st.session_state.api_key.strip():
            st.error("Informe a API Key na barra lateral antes de cotar.")
        else:
            try:
                if not origin_cep.strip():
                    raise ValueError("Informe o CEP de origem.")
                if not todos_ceps:
                    raise ValueError(
                        "Informe pelo menos um CEP de destino por texto/Excel "
                        "ou marque 'Utilizar os CEPs Padrão'."
                    )

                peso_f = parse_float_br(peso)
                largura_f = parse_float_br(largura)
                altura_f = parse_float_br(altura)
                comprimento_f = parse_float_br(comprimento)
                valor_f = parse_float_br(valor)

                st.session_state.todas_opcoes = []
                st.session_state.logs_execucao = []
                st.session_state.erros_cotacao = []
                st.session_state.detalhes_opcoes = {}

                aviso_processamento.info(
                    "Processando as cotações. Aguarde até a conclusão de todos os CEPs..."
                )

                progresso = progresso_container.progress(0)

                with st.spinner("Consultando a Intelipost e processando os CEPs..."):
                    for idx, dest in enumerate(todos_ceps, start=1):
                        status_execucao.info(
                            f"Processando CEP {dest} ({idx}/{len(todos_ceps)}). Aguarde..."
                        )

                        try:
                            resposta = cotar_frete_intelipost(
                                origin_cep=origin_cep.strip(),
                                dest_cep=dest,
                                peso=peso_f,
                                largura=largura_f,
                                altura=altura_f,
                                comprimento=comprimento_f,
                                valor=valor_f,
                                api_key=st.session_state.api_key.strip(),
                                usar_tabelas_rascunho=usar_tabelas_rascunho,
                            )
                            opcoes = extrair_opcoes_frete(dest, resposta)

                            if opcoes:
                                st.session_state.todas_opcoes.extend(opcoes)
                                st.session_state.logs_execucao.append(
                                    {
                                        "destino": dest,
                                        "status": "ok",
                                        "opcoes": len(opcoes),
                                        "modo_tabela": "rascunho" if usar_tabelas_rascunho else "vigente",
                                    }
                                )
                            else:
                                mensagem_sem_opcoes = (
                                    "A API respondeu com sucesso, mas não retornou opções de entrega."
                                )
                                st.session_state.logs_execucao.append(
                                    {
                                        "destino": dest,
                                        "status": "sem_opcoes",
                                        "opcoes": 0,
                                        "modo_tabela": "rascunho" if usar_tabelas_rascunho else "vigente",
                                        "detalhe": mensagem_sem_opcoes,
                                    }
                                )
                                st.session_state.erros_cotacao.append(
                                    {
                                        "cep": dest,
                                        "tipo_erro": "SEM_OPCOES",
                                        "status_code": 200,
                                        "mensagem": mensagem_sem_opcoes,
                                    }
                                )

                        except requests.HTTPError as e:
                            resp = e.response
                            status_code = resp.status_code if resp is not None else None
                            detalhe_erro = extrair_mensagem_erro_response(resp) if resp is not None else str(e)

                            st.session_state.logs_execucao.append(
                                {
                                    "destino": dest,
                                    "status": f"http_{status_code}" if status_code else "http_erro",
                                    "opcoes": 0,
                                    "modo_tabela": "rascunho" if usar_tabelas_rascunho else "vigente",
                                    "detalhe": detalhe_erro,
                                }
                            )

                            st.session_state.erros_cotacao.append(
                                {
                                    "cep": dest,
                                    "tipo_erro": "HTTP",
                                    "status_code": status_code,
                                    "mensagem": detalhe_erro,
                                }
                            )

                        except Exception as e:
                            detalhe_erro = str(e)

                            st.session_state.logs_execucao.append(
                                {
                                    "destino": dest,
                                    "status": "erro",
                                    "opcoes": 0,
                                    "modo_tabela": "rascunho" if usar_tabelas_rascunho else "vigente",
                                    "detalhe": detalhe_erro,
                                }
                            )

                            st.session_state.erros_cotacao.append(
                                {
                                    "cep": dest,
                                    "tipo_erro": "EXCEPTION",
                                    "status_code": None,
                                    "mensagem": detalhe_erro,
                                }
                            )

                        progresso.progress(idx / len(todos_ceps))

                total_erros = len(st.session_state.erros_cotacao)
                total_sucessos = len(
                    [log for log in st.session_state.logs_execucao if log.get("status") == "ok"]
                )

                aviso_processamento.success(
                    "Processamento finalizado. Consulte as abas abaixo para ver cotações, execução e erros."
                )
                status_execucao.success(
                    f"Concluído: {len(todos_ceps)} CEP(s) processado(s), "
                    f"{total_sucessos} com retorno de cotação e {total_erros} com ocorrência."
                )

            except Exception as e:
                aviso_processamento.empty()
                status_execucao.empty()
                progresso_container.empty()
                st.error(f"Falha na validação/processamento: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.todas_opcoes or st.session_state.erros_cotacao:
        df = pd.DataFrame(st.session_state.todas_opcoes)

        if not df.empty:
            df["prazo_dias"] = pd.to_numeric(df["prazo_dias"], errors="coerce")
            df["valor_frete"] = pd.to_numeric(df["valor_frete"], errors="coerce")

            try:
                df_ceps_padrao_info = obter_info_ceps_padrao().rename(columns={"cep": "destino"})
                df = df.merge(df_ceps_padrao_info, on="destino", how="left")
            except Exception:
                pass

            render_metricas(df)

        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            ["Cotações", "Médias por transportadora", "Médias por CEP", "Execução", "CEPs com erro"]
        )

        with tab1:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            st.subheader("Opções de frete")

            if not df.empty:
                colunas_ordem = [
                    c for c in
                    [
                        "estado",
                        "cidade",
                        "tipo",
                        "destino",
                        "transportadora",
                        "prazo_dias",
                        "valor_frete",
                        "id_cotacao",
                    ]
                    if c in df.columns
                ]
                outras_colunas = [c for c in df.columns if c not in colunas_ordem]
                df_exibicao = df[colunas_ordem + outras_colunas].sort_values(
                    ["destino", "valor_frete", "prazo_dias"],
                    ascending=[True, True, True],
                )
                st.caption("Clique em uma linha para ver o detalhamento da cotação (taxas, pedágio e demais campos).")
                evento = st.dataframe(
                    df_exibicao,
                    use_container_width=True,
                    hide_index=True,
                    column_config={COLUNA_REF: None},
                    on_select="rerun",
                    selection_mode="single-row",
                    key="tabela_cotacoes",
                )
                planilhas = montar_planilhas_cotacoes(df_exibicao, st.session_state.detalhes_opcoes)
                botoes_download(
                    planilhas,
                    nome_arquivo="cotacoes_frete_intelipost",
                    rotulo="Baixar cotações",
                    titulo_pdf="Cotações de frete Intelipost",
                    secoes_pdf=montar_secoes_pdf_cotacoes(df_exibicao, st.session_state.detalhes_opcoes),
                )

                linhas_selecionadas = evento.selection.rows
                if linhas_selecionadas and COLUNA_REF in df_exibicao.columns:
                    ref = df_exibicao.iloc[linhas_selecionadas[0]][COLUNA_REF]
                    detalhe = st.session_state.detalhes_opcoes.get(ref)
                    st.divider()
                    if detalhe:
                        render_detalhe_opcao(detalhe)
                    else:
                        st.warning("Detalhes desta opção não estão mais disponíveis. Refaça a cotação.")
            else:
                st.info("Nenhuma cotação válida retornada.")

            st.markdown("</div>", unsafe_allow_html=True)

        with tab2:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            st.subheader("Médias por transportadora")

            if not df.empty:
                df_media_transp = (
                    df.groupby("transportadora", as_index=False)
                    .agg(
                        media_valor=("valor_frete", "mean"),
                        media_prazo=("prazo_dias", "mean"),
                        qtd_opcoes=("transportadora", "count"),
                    )
                    .sort_values(["media_valor", "media_prazo"], ascending=[True, True])
                )
                df_media_transp["media_valor"] = df_media_transp["media_valor"].round(2)
                df_media_transp["media_prazo"] = df_media_transp["media_prazo"].round(2)
                st.dataframe(df_media_transp, use_container_width=True, hide_index=True)
            else:
                st.info("Sem dados para calcular médias por transportadora.")

            st.markdown("</div>", unsafe_allow_html=True)

        with tab3:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            st.subheader("Médias por CEP destino")

            if not df.empty:
                agrupadores_cep = ["destino"]
                for extra in ["estado", "cidade", "tipo"]:
                    if extra in df.columns:
                        agrupadores_cep.append(extra)

                df_media_cep = (
                    df.groupby(agrupadores_cep, as_index=False)
                    .agg(
                        media_valor=("valor_frete", "mean"),
                        media_prazo=("prazo_dias", "mean"),
                        qtd_opcoes=("destino", "count"),
                    )
                    .sort_values(["media_valor", "media_prazo"], ascending=[True, True])
                )
                df_media_cep["media_valor"] = df_media_cep["media_valor"].round(2)
                df_media_cep["media_prazo"] = df_media_cep["media_prazo"].round(2)
                st.dataframe(df_media_cep, use_container_width=True, hide_index=True)
                st.bar_chart(df_media_cep.set_index("destino")["media_valor"])
            else:
                st.info("Sem dados para calcular médias por CEP.")

            st.markdown("</div>", unsafe_allow_html=True)

        with tab4:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            st.subheader("Resumo da execução")

            if st.session_state.logs_execucao:
                df_logs = pd.DataFrame(st.session_state.logs_execucao)
                st.dataframe(df_logs, use_container_width=True, hide_index=True)

                botoes_download(
                    {"execucao": df_logs},
                    nome_arquivo="execucao_cotacoes_intelipost",
                    rotulo="Baixar execução",
                    titulo_pdf="Resumo da execução - cotações Intelipost",
                )
            else:
                st.info("Nenhum log disponível.")

            st.markdown("</div>", unsafe_allow_html=True)

        with tab5:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            st.subheader("CEPs com erro")

            if st.session_state.erros_cotacao:
                df_erros = pd.DataFrame(st.session_state.erros_cotacao)
                st.dataframe(df_erros, use_container_width=True, hide_index=True)

                botoes_download(
                    {"ceps_com_erro": df_erros},
                    nome_arquivo="ceps_com_erro_intelipost",
                    rotulo="Baixar erros",
                    titulo_pdf="CEPs com erro - cotações Intelipost",
                )
            else:
                st.success("Nenhum CEP com erro nesta execução.")

            st.markdown("</div>", unsafe_allow_html=True)

    else:
        st.info("Nenhuma cotação disponível ainda. Preencha os dados e clique em 'Cotar fretes'.")

    render_debug_area()


if __name__ == "__main__":
    main()