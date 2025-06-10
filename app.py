import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
from zipfile import ZipFile
import difflib
import os
from datetime import datetime
import re


modelos_parecer = {
    "Simpress": """
PARECER

Em atendimento ao que dispõe à Cláusula 8ª, do contrato nº. 001/2024, emitimos o Parecer {{tipo}}, referente ao período de {{mes_ano}}, atestado através da Nota Fiscal Nº {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Assessor - ASSTIN
ID.: {{fiscal_id}}
""",
    "Claro": """
PARECER

Em atendimento ao que dispõe à Cláusula 3ª, do contrato nº. 002/2022, emitimos o Parecer {{tipo}}, referente ao período de {{mes_ano}}, atestado através da Nota Fiscal Nº {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Fiscal do Contrato
ID.: {{fiscal_id}}
""",
    "Ntsec": """
PARECER

Em atendimento ao que dispõe à Cláusula 7ª, do contrato nº. 011/2022, emitimos o Parecer {{tipo}}, referente ao período de {{mes_ano}}, atestado através das Notas Fiscais:

- Nota fiscal nº {{numero_nf}}

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}
""",
    "Zimbra": """
PARECER

Em atendimento ao que dispõe à Cláusula 7ª, do contrato nº. 009/2022, emitimos o Parecer {{tipo}}, referente ao período de {{mes_ano}}, atestado através da Nota Fiscal Nº {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Fiscal do Contrato
ID.: {{fiscal_id}}
""",
    "Datacorpore": """
PARECER

Em atendimento ao que dispõe à Cláusula 7ª, do contrato nº. 014/2022, emitimos o Parecer {{tipo}}, referente ao período de {{mes_ano}}, conforme atestado pela seguinte Nota Fiscal:

- Nota Fiscal nº {{numero_nf}}

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato - ASSTIN
ID.: {{gestor_id}}

{{fiscal_nome}}
Assessor - ASSTIN
ID.: {{fiscal_id}}
""",
    "Dadyilha": """
PARECER

Em atendimento ao que dispõe à Cláusula 7ª, do contrato nº. 018/2022, emitimos o Parecer {{tipo}}, referente ao período de {{mes_ano}}, conforme atestado pela seguinte Nota Fiscal:

- Nota Fiscal nº {{numero_nf}}

{{gestor_nome}}
Gestor do Contrato - ASSTIN
ID.: {{gestor_id}}

{{fiscal_nome}}
Assessor - ASSTIN
ID.: {{fiscal_id}}
""",
    "Vivo": """
PARECER

Em atendimento ao que dispõe à Cláusula 3ª, do contrato Nº. 006/2024, emitimos o Parecer {{tipo}}, referente ao período de {{mes_ano}}, atestado através da Nota Fiscal Nº {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Fiscal do Contrato
ID.: {{fiscal_id}}
""",
}

modelos_despacho = {
    "Simpress": """
À DIVISÃO FINANCEIRA - CONTAS A PAGAR
DA ASSESSORIA DE TECNOLOGIA DA INFORMAÇÃO

ASSUNTO: FATURAMENTO REFERENTE A {{mes_ano}}.
FAVORECIDO: SIMPRESS COMÉRCIO, LOCAÇÃO E SERVIÇOS LTDA.

Informamos que anexamos ao processo a seguinte documentação:

Nota Fiscal de n° {{numero_nf}} da SIMPRESS COMÉRCIO, LOCAÇÃO E SERVIÇOS LTDA no valor de R$ {{valor}} ({{valor_extenso}}), referente ao mês de {{mes_ano}} pela Prestação de serviço de locação de computadores;

Certidão negativa de débitos trabalhistas.

Certifico que o documento foi minuciosamente analisado pelos fiscais de contrato, atestando sua conformidade com as cláusulas acordadas. A certificação busca assegurar transparência e validade ao processo, fortalecendo a credibilidade das relações contratuais.

Solicitamos conferência e providências administrativas para pagamento da Nota Fiscal n° {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Assessor - ASSTIN
ID.: {{fiscal_id}}
""",
    "Claro": """
À DIVISÃO FINANCEIRA - CONTAS A PAGAR
DA ASSESSORIA TECNOLOGIA DA INFORMAÇÃO

ASSUNTO: FATURAMENTO REFERENTE A {{mes_ano}}
FAVORECIDO: CLARO S/A (FILIAL)

Prezados, informamos o faturamento referente a {{mes_ano}}, anexamos ao processo a seguinte documentação:

Nota Fiscal Nº {{numero_nf}} da CLARO S.A no valor de R$ {{valor}} ({{valor_extenso}}), referente ao mês de {{mes_ano}} pela prestação de serviços de acesso à internet;

Certidão negativa de débitos trabalhistas.

Certifico que o documento foi minuciosamente analisado pelos fiscais de contrato, atestando sua conformidade com as cláusulas acordadas. A certificação busca assegurar transparência e validade ao processo, fortalecendo a credibilidade das relações contratuais. Estamos disponíveis para esclarecimentos adicionais, reiterando nosso compromisso com a integridade do contrato.

Solicito conferência e providências administrativas para pagamento da Nota Fiscal N° {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Fiscal do Contrato
ID.: {{fiscal_id}}
""",
    "Ntsec": """
À DIVISÃO FINANCEIRA - CONTAS A PAGAR
DA ASSESSORIA TECNOLOGIA DA INFORMAÇÃO

ASSUNTO: FATURAMENTO REFERENTE A {{mes_ano}}
FAVORECIDO: NTSEC SOLUÇÕES EM TELEINFORMÁTICA LTDA

Informamos que anexamos ao processo a seguinte documentação:

Nota Fiscal de n° {{numero_nf1}} no valor de R$ {{valor1}} ({{valor_extenso1}}), referente ao mês de {{mes_ano}} pela prestação de serviços relacionados à solução antivírus;

Nota Fiscal de n° {{numero_nf2}} no valor de R$ {{valor2}} ({{valor_extenso2}}), referente ao mês de {{mes_ano}} pela prestação de serviços relacionados à solução antivírus;

Certidão negativa de débitos trabalhistas.

Certifico que os documentos foram minuciosamente analisados pelos fiscais de contrato, atestando sua conformidade com as cláusulas acordadas. A certificação busca assegurar transparência e validade ao processo, fortalecendo a credibilidade das relações contratuais.

Solicitamos conferência e providências administrativas para pagamento das notas fiscais acima.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}
""",

    "Zimbra": """
À DIVISÃO FINANCEIRA - CONTAS A PAGAR
DA ASSESSORIA TECNOLOGIA DA INFORMAÇÃO

ASSUNTO: FATURAMENTO REFERENTE A {{mes_ano}}
FAVORECIDO: PRODERJ CENTRO DE TECNOLOGIA DE INFORMAÇÃO E COMUNICAÇÃO DO ESTADO DO RIO DE JANEIRO

Anexamos ao processo a seguinte documentação:

Nota Fiscal de n° {{numero_nf}} no valor de R$ {{valor}} ({{valor_extenso}}), referente ao mês de {{mes_ano}} pela prestação de serviço de Hospedagem de Mensageria Eletrônica, incluindo armazenamento de arquivo na nuvem, infaestrutura de hardware, software, armazenamento, backup dos dados, segurança e monitoramento.

Certidão negativa de débitos trabalhistas.

Certifico que os documentos foram minuciosamente analisados pelos fiscais de contrato, atestando sua conformidade com as cláusulas acordadas. A certificação busca assegurar transparência e validade ao processo, fortalecendo a credibilidade das relações contratuais.

Solicito conferência e providências administrativas para pagamento da Nº {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Fiscal do Contrato
ID.: {{fiscal_id}}
""",
    "Datacorpore": """
À DIVISÃO FINANCEIRA - CONTAS A PAGAR
DA ASSESSORIA TECNOLOGIA DA INFORMAÇÃO

ASSUNTO: FATURAMENTO REFERENTE A {{mes_ano}}
FAVORECIDO: DATA CORPORE SERVIÇOS DE TELECOMUNICAÇÕES E INFORMATICA LTDA

Informamos que anexamos ao processo a seguinte documentação:

Nota Fiscal de n° {{numero_nf}} da DATA CORPORE SERVIÇOS DE TELECOMUNICAÇÕES E INFORMATICA LTDA no valor de R$ {{valor}} ({{valor_extenso}}), referente ao mês de {{mes_ano}} pela Prestação de serviço de locação de acesso à internet;

Certidão negativa de débitos trabalhistas.

Certifico que o documento foi minuciosamente analisado pelos fiscais de contrato, atestando sua conformidade com as cláusulas acordadas. A certificação busca assegurar transparência e validade ao processo, fortalecendo a credibilidade das relações contratuais.

Solicitamos conferência e providências administrativas para pagamento da Nota Fiscal N° {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Assessor - ASSTIN
ID.: {{fiscal_id}}
""",
    "Dadyilha": """
À DIVISÃO FINANCEIRA - CONTAS A PAGAR
DA ASSESSORIA TECNOLOGIA DA INFORMAÇÃO

ASSUNTO: FATURAMENTO REFERENTE A {{mes_ano}}
FAVORECIDO: DADY ILHA SOLUCOES INTEGRADAS LTDA

Nota Fiscal de N° {{numero_nf}} da DADY ILHA SOLUCOES INTEGRADAS LTDA no valor de R$ {{valor}} ({{valor_extenso}}), referente ao mês de {{mes_ano}} pela Prestação de serviço de solução continuada de impressão, cópia e digitalização corporativa;

Certidão Negativa de Débitos trabalhistas.

Certifico que o documento foi minuciosamente analisado pelos fiscais de contrato, atestando sua conformidade com as cláusulas acordadas. A certificação busca assegurar transparência e validade ao processo, fortalecendo a credibilidade das relações contratuais.

Solicitamos conferência e providências administrativas para pagamento da Nota Fiscal n° {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Assessor - ASSTIN
ID.: {{fiscal_id}}
""",
    "Vivo": """
À DIVISÃO FINANCEIRA - CONTAS A PAGAR
DA ASSESSORIA TECNOLOGIA DA INFORMAÇÃO

ASSUNTO: FATURAMENTO REFERENTE A {{mes_ano}}
FAVORECIDO: TELEFÔNICA BRASIL S/A (VIVO)

Informamos que anexamos ao processo a seguinte documentação:

Nota Fiscal de N° {{numero_nf}} no valor de R$ {{valor}} ({{valor_extenso}}), referente ao mês de {{mes_ano}} pela prestação de serviço de telefonia fixa;

Certidão Positiva de Débitos Trabalhistas com Efeito de Negativa.

Certifico que o documento foi minuciosamente analisado pelos fiscais de contrato, atestando sua conformidade com as cláusulas acordadas. A certificação busca assegurar transparência e validade ao processo, fortalecendo a credibilidade das relações contratuais.

Solicitamos conferência e providências administrativas para pagamento da Nota Fiscal N° {{numero_nf}}.

Atenciosamente,

{{gestor_nome}}
Gestor do Contrato
ID.: {{gestor_id}}

{{fiscal_nome}}
Fiscal do Contrato
ID.: {{fiscal_id}}
"""
}

def formatar_moeda(valor_str):
    try:
        valor_str = str(valor_str).strip()

        # Remover qualquer símbolo que não seja número, ponto ou vírgula
        valor_str = re.sub(r"[^\d.,]", "", valor_str)

        # Situações:
        # - 1.231,45 => BR → substituir "." por "" e "," por "."
        # - 1231,45  => BR sem milhar → substituir "," por "."
        # - 1231.45  => US → manter

        if "," in valor_str:
            # Sempre trata como formato brasileiro
            valor_str = valor_str.replace(".", "").replace(",", ".")
        
        valor = float(valor_str)
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao formatar moeda: {e}")
        return valor_str

def preencher_modelo(modelo, dados):
    for chave, valor in dados.items():
        modelo = modelo.replace(f"{{{{{chave}}}}}", str(valor))
    return modelo

def gerar_docx(conteudo):
    doc = Document()
    for linha in conteudo.split("\n"):
        doc.add_paragraph(linha)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def detectar_cabecalho(df):
    for i in range(5):
        if df.iloc[i].astype(str).str.contains("MÊS", case=False, na=False).any():
            return i
    return 0

def encontrar_coluna(possibilidades, colunas):
    for p in possibilidades:
        match = difflib.get_close_matches(p.lower(), [str(c).lower() for c in colunas], n=1, cutoff=0.6)
        if match:
            return next((c for c in colunas if c.lower() == match[0]), "")
    return ""

# Modelos (parecer e despacho)
# from modelos import modelos_parecer, modelos_despacho

# Interface Streamlit
st.set_page_config(page_title="Gerador de Documentos", layout="wide")
st.title("📄 Gerador de Pareceres e Despachos")

arquivo = st.file_uploader("📁 Envie a planilha .xlsm", type="xlsm")

if arquivo:
    xls = pd.ExcelFile(arquivo)
    empresas = xls.sheet_names
    empresa = st.selectbox("🏢 Selecione a empresa", empresas)

    df_raw = xls.parse(empresa, header=None)
    linha_cabecalho = detectar_cabecalho(df_raw)
    df = pd.read_excel(arquivo, sheet_name=empresa, header=linha_cabecalho)
    df = df.dropna(how="all").reset_index(drop=True)

    st.subheader("🔍 Pré-visualização da Planilha")
    st.dataframe(df.head())

    # 🔄 SELEÇÃO DE TIPO DE DOCUMENTO - FORA DO FORM
    tipo = st.radio("📌 Tipo de documento", ["Parecer", "Despacho"])

    with st.form("formulario_gerador"):
        st.markdown("### 🎯 Configuração do Documento")

        linhas = st.multiselect("Selecione uma ou mais linhas", df.index.tolist())
        gestor_nome = st.text_input("👤 Nome do Gestor", "Glauter Gaspar Valle")
        gestor_id = st.text_input("🆔 ID do Gestor", "51469944")
        fiscal_nome = st.text_input("👤 Nome do Fiscal/Assessor", "Lucas Pires Ponte")
        fiscal_id = st.text_input("🆔 ID do Fiscal/Assessor", "51567660")

        parecer_tipo = ""
        if tipo == "Parecer":
            parecer_tipo = st.radio("📌 Tipo do Parecer", ["Provisório", "Definitivo"])

        gerar = st.form_submit_button("🚀 Gerar Documento(s)")

    if gerar:
        col_mes = encontrar_coluna(["mês", "mes"], df.columns)
        col_nf = encontrar_coluna(["número da nf", "número nf", "nf"], df.columns)
        col_valor = encontrar_coluna(["valor (r$)", "valor 1", "valor"], df.columns)
        col_extenso = encontrar_coluna(["por extenso", "valor por extenso"], df.columns)

        campos_faltando = []
        if not col_mes: campos_faltando.append("Mês")
        if not col_nf: campos_faltando.append("Número da NF")
        if not col_valor: campos_faltando.append("Valor")
        if not col_extenso: campos_faltando.append("Valor por Extenso")

        if campos_faltando:
            st.error(f"⚠️ As seguintes colunas não foram encontradas na planilha: {', '.join(campos_faltando)}.")
            st.stop()

        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, "w") as zipf:
            for linha in linhas:
                linha_selecionada = df.iloc[linha]
                valor_extenso = str(linha_selecionada.get(col_extenso, "")).strip()
                if not valor_extenso:
                    st.warning(f"⚠️ Linha {linha}: O campo 'valor por extenso' está vazio.")

                empresa_padrao = empresa.strip().title()

                if empresa_padrao == "Ntsec" and tipo == "Despacho":
                    dados = {
                        "mes_ano": str(linha_selecionada.get(col_mes, "")).strip(),
                        "numero_nf1": str(linha_selecionada.get("NÚMERO DA NF 1", "")).strip(),
                        "valor1": formatar_moeda(linha_selecionada.get("VALOR 1 (R$)", "")),
                        "valor_extenso1": str(linha_selecionada.get("VALOR POR EXTENSO 1", "")).strip(),
                        "numero_nf2": str(linha_selecionada.get("NÚMERO DA NF 2", "")).strip(),
                        "valor2": formatar_moeda(linha_selecionada.get("VALOR 2 (R$)", "")),
                        "valor_extenso2": str(linha_selecionada.get("VALOR POR EXTENSO 2", "")).strip(),
                        "gestor_nome": gestor_nome.strip(),
                        "gestor_id": gestor_id.strip(),
                        "fiscal_nome": fiscal_nome.strip(), 
                        }           
                else:
                    dados = {
                        "tipo": parecer_tipo,
                        "mes_ano": str(linha_selecionada.get(col_mes, "")).strip(),
                        "numero_nf": str(linha_selecionada.get(col_nf, "")).strip(),
                        "valor": formatar_moeda(linha_selecionada.get(col_valor, "")),
                        "valor_extenso": valor_extenso,
                        "gestor_nome": gestor_nome.strip(),
                        "gestor_id": gestor_id.strip(),
                        "fiscal_nome": fiscal_nome.strip(),
                        "fiscal_id": fiscal_id.strip(),
                            }


                empresa_padrao = empresa.strip().title()
                modelo = modelos_parecer.get(empresa_padrao) if tipo == "Parecer" else modelos_despacho.get(empresa_padrao)

                if not modelo:
                    st.warning(f"Modelo não encontrado para a empresa '{empresa_padrao}' na linha {linha}.")
                    continue

                texto_final = preencher_modelo(modelo, dados)
                docx = gerar_docx(texto_final)
                zipf.writestr(f"{tipo}_{empresa_padrao}_linha{linha}.docx", docx.read())

                with st.expander(f"📄 Pré-visualização - Linha {linha}"):
                    st.text_area("", texto_final, height=300)

        zip_buffer.seek(0)
        st.download_button("📥 Baixar todos os documentos (.zip)", zip_buffer, file_name="documentos_gerados.zip")
