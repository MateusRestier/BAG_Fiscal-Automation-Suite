import os
import re
import pdfplumber
from datetime import datetime
import pandas as pd
import sys
import contextlib
import pyodbc
import pythoncom
from win32com.client import Dispatch

"""Config dotenv"""
from dotenv import load_dotenv
from pathlib import Path
def localizar_env(diretorio_raiz="PRIVATE_BAG.ENV"):
    path = Path(__file__).resolve()
    for parent in path.parents:
        possible = parent / diretorio_raiz / ".env"
        if possible.exists():
            return possible
    raise FileNotFoundError(f"Arquivo .env não encontrado dentro de '{diretorio_raiz}'.")
env_path = localizar_env()
load_dotenv(dotenv_path=env_path)

@contextlib.contextmanager
def suppress_stderr():
    with open(os.devnull, 'w') as devnull:
        old_stderr = sys.stderr
        sys.stderr = devnull
        try:
            yield
        finally:
            sys.stderr = old_stderr

# Obter data atual para montar os diretórios
today = datetime.now()
year = today.strftime("%Y")
month = today.strftime("%m")
day = today.strftime("%d")

dir_pdf = os.getenv("DIR_PDF_FICAL_BAHIA")

# Diretórios com os PDFs do dia
diretorios = [
    rf"{dir_pdf}Contabilidade\Fiscal\{year}\LUCRO REAL\SHEHRAZADE\{month}.{year}\ICMS\ICMS ANTECIPADO E ST\BAHIA\Antecipado\{year}\{month}\{day}",
    rf"{dir_pdf}Contabilidade\Fiscal\{year}\LUCRO REAL\SHEHRAZADE\{month}.{year}\ICMS\ICMS ANTECIPADO E ST\BAHIA\AntecipadoST\{year}\{month}\{day}"
]

# ---------- Funções para extrair colunas ----------

def extrair_datapag(texto: str) -> str:
    """Extrai a data de pagamento no formato YYYYMMDD."""
    match = re.search(r"pagamento até (\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if match:
        data_br = match.group(1)
        return datetime.strptime(data_br, "%d/%m/%Y").strftime("%Y%m%d")
    return None

def extrair_datavenc(texto: str) -> str:
    """Extrai a data de vencimento no formato YYYYMMDD."""
    match = re.search(r"DATA\s+DE\s+VENCIMENTO\s+(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if match:
        data_br = match.group(1)
        return datetime.strptime(data_br, "%d/%m/%Y").strftime("%Y%m%d")
    return None

def extrair_competencia(texto: str) -> str:
    """Extrai a competência com base na data de pagamento. Formato: YYYYMM"""
    match = re.search(r"pagamento até (\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if match:
        data_pag = datetime.strptime(match.group(1), "%d/%m/%Y")
        return data_pag.strftime("%Y%m")
    return None

def extrair_qtdnf(texto: str) -> str:
    """Extrai a quantidade de notas fiscais (QTDNF) a partir do trecho 'Notas Fiscais:N'."""
    match = re.search(r"Notas\s+Fiscais\s*:\s*(\d+)", texto, re.IGNORECASE)
    if match:
        return match.group(1)
    return None

def extrair_nf(texto: str) -> str:
    """
    Extrai todas as notas fiscais do bloco entre 'Notas Fiscais:' e 'Antecipado'/'AntecipadoST',
    mesmo que estejam em várias linhas ou coladas.
    """
    match = re.search(
        r"Notas\s+Fiscais\s*:\s*\d+\s*(.*?)\s*Antecipado(?:ST)?\s*-\s*[A-Z0-9]+\s*-\s*\d{4}/\d{2}/\d{2}",
        texto,
        re.IGNORECASE | re.DOTALL
    )
    if match:
        bloco_nf = match.group(1)
        nfs = re.findall(r"\d{6,}", bloco_nf)
        return ",".join(nfs)
    return None

def extrair_loja_arquivo(nome_arquivo: str) -> str:
    """
    Extrai a loja diretamente do nome do arquivo.
    Ex: 'AntecipadoST - 78 - 2025-01-02.pdf' -> retorna '78'
    """
    match = re.search(r"-(?:\s*)?([A-Z0-9]+)(?:\s*)?-\s*\d{4}-\d{2}-\d{2}", nome_arquivo)
    if match:
        return match.group(1)
    return None

def extrair_numeroguia(texto: str) -> str:
    """
    Extrai o número da guia a partir do campo 'Nº DE SÉRIE / NOSSO NÚMERO'
    que aparece no DAE. Retorna apenas os dígitos.
    """
    linhas = texto.splitlines()
    for i, linha in enumerate(linhas):
        if re.search(r"n[ºoO]?\s*de\s*s[ée]rie\s*/\s*nosso\s*n[úu]mero", linha, re.IGNORECASE):
            if i + 1 < len(linhas):
                proxima_linha = linhas[i + 1].strip()
                match = re.search(r"\d{6,}", proxima_linha)
                if match:
                    return match.group(0)
    return None

def extrair_valorprin(texto: str) -> float:
    """
    Extrai o valor principal (VALORPRIN) a partir do campo 'VALOR PRINCIPAL'.
    Converte para float no formato padrão (ex: 1062.32).
    """
    linhas = texto.splitlines()
    for i, linha in enumerate(linhas):
        if "VALOR PRINCIPAL" in linha.upper():
            if i + 1 < len(linhas):
                proxima_linha = linhas[i + 1].strip()
                match = re.search(r"R?\$?\s*([\d\.]+,\d{2})", proxima_linha)
                if match:
                    valor_str = match.group(1).replace(".", "").replace(",", ".")
                    return float(valor_str)
    return None

def extrair_valortotal(texto: str) -> float:
    """
    Extrai o valor total a recolher (VALORTOTAL) do campo 'TOTAL A RECOLHER'.
    Retorna como float sem formatação (ex: 1199.88).
    """
    linhas = texto.splitlines()
    for i, linha in enumerate(linhas):
        if "TOTAL A RECOLHER" in linha.upper():
            if i + 1 < len(linhas):
                proxima_linha = linhas[i + 1].strip()
                match = re.search(r"R?\$?\s*([\d\.]+,\d{2})", proxima_linha)
                if match:
                    valor_str = match.group(1).replace(".", "").replace(",", ".")
                    return float(valor_str)
    return None

def extrair_uf() -> str:
    """
    Retorna a UF fixa 'BA' para todos os registros.
    """
    return "BA"

# ---------- Processamento por PDF ----------

def processar_pdf(caminho_pdf: str) -> dict:
    """Abre o PDF, extrai o texto e aplica as funções de extração de colunas. Exibe resumo por coluna."""
    nome_arquivo = os.path.basename(caminho_pdf)
    with suppress_stderr():
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = "\n".join(page.extract_text() or "" for page in pdf.pages)

    resultado = {
        "Arquivo": os.path.basename(caminho_pdf),
        "DATAPAG": extrair_datapag(texto),
        "DATAVENC": extrair_datavenc(texto),
        "COMPETENCIA": extrair_competencia(texto),
        "QTDNF": extrair_qtdnf(texto),
        "NF": extrair_nf(texto),
        "LOJA": extrair_loja_arquivo(nome_arquivo),
        "NUMEROGUIA": extrair_numeroguia(texto),
        "VALORPRIN": extrair_valorprin(texto),
        "VALORTOTAL": extrair_valortotal(texto),
        "UF": extrair_uf()
    }

    print(f"📄 {resultado['Arquivo']}")
    for coluna, valor in resultado.items():
        if coluna != "Arquivo":
            status = f"✅ {valor}" if valor else "❌ Não encontrado"
            print(f"   ↳ {coluna}: {status}")
    print()

    return resultado

# ---------- Jogar no Banco ----------

def inserir_no_banco(df: pd.DataFrame):
    """
    Insere os dados do DataFrame na tabela FC_EmissaoDAEs no SQL Server
    e remove duplicatas após a inserção.
    """
    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.getenv('DB_SERVER_EXCEL')},{os.getenv('DB_PORT_EXCEL')};"
        f"DATABASE={os.getenv('DB_DATABASE_EXCEL')};"
        f"UID={os.getenv('DB_USER_EXCEL')};"
        f"PWD={os.getenv('DB_PASSWORD_EXCEL')}"
    )
    cursor = conn.cursor()

    insert_sql = """
        INSERT INTO DADOS_EXCEL.dbo.FC_EmissaoDAEs
        (Arquivo, COMPETENCIA, LOJA, NUMEROGUIA, VALORPRIN, VALORTOTAL, DATAVENC, DATAPAG, NF, QTDNF, UF)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    for _, row in df.iterrows():
        cursor.execute(insert_sql, (
            row.get("Arquivo"),
            row.get("COMPETENCIA"),
            row.get("LOJA"),
            row.get("NUMEROGUIA"),
            row.get("VALORPRIN"),
            row.get("VALORTOTAL"),
            row.get("DATAVENC"),
            row.get("DATAPAG"),
            row.get("NF"),
            row.get("QTDNF"),
            row.get("UF")
        ))

    conn.commit()
    print("✅ Dados inseridos com sucesso.")

    # Remoção de duplicatas no banco
    print("♻️ Removendo duplicatas no banco...")
    remover_duplicatas_sql = """
    WITH CTE_Duplicadas AS (
        SELECT *,
            ROW_NUMBER() OVER (
                PARTITION BY
                    Arquivo,
                    COMPETENCIA,
                    LOJA,
                    NUMEROGUIA,
                    VALORPRIN,
                    VALORTOTAL,
                    DATAVENC,
                    DATAPAG,
                    NF,
                    QTDNF,
                    UF
                ORDER BY ID
            ) AS rn
        FROM DADOS_EXCEL.dbo.FC_EmissaoDAEs
    )
    DELETE FROM CTE_Duplicadas WHERE rn > 1;
    """

    cursor.execute(remover_duplicatas_sql)
    conn.commit()
    print("✅ Duplicatas removidas com sucesso.")

    cursor.close()
    conn.close()

# ---------- Enviar Email ----------

def enviar_email_guias_emitidas(df: pd.DataFrame, diretorios: list):
    if df.empty:
        print("📭 Nenhuma guia encontrada para enviar.")
        return

    # Pega a data de pagamento da primeira guia (todas são do mesmo dia)
    data_pag = df["DATAPAG"].iloc[0]  # formato: YYYYMMDD
    venc_ddmm = f"{data_pag[6:8]}/{data_pag[4:6]}"
    venc_formatado = f"{data_pag[6:8]}/{data_pag[4:6]}/{data_pag[0:4]}"
    referencia = f"{data_pag[4:6]}/{data_pag[0:4]}"

    # Corpo com títulos
    linhas = []
    for _, row in df.iterrows():
        guia = str(row["NUMEROGUIA"]).strip()[-9:]
        valor = float(row["VALORTOTAL"])
        linhas.append(f"{guia} - R${valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    corpo = f"""Bom dia!
Seguem guias de ICMS ANTECIPADO e ST para liberação de mercadoria retida em barreira com vencimento em: {venc_formatado}.

Título(s):
{chr(10).join(linhas)}

Atenciosamente,
Mateus Restier"""

    assunto = f"ANTECIPADO BA SHEHRAZADE - VENC: {venc_ddmm} - ref. {referencia}"

    try:
        pythoncom.CoInitialize()
        outlook = Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = "pagamentos@bagaggio.com.br; beatriz.alvim@bagaggio.com.br"
        mail.CC = "bgfiscal@bagaggio.com.br; rafaella.camacho@bagaggio.com.br; jessica.rodrigues@bagaggio.com.br"
        mail.Subject = assunto
        mail.Body = corpo

        # Procurar e anexar arquivos PDF
        arquivos_anexados = 0
        for arquivo_pdf in df["Arquivo"]:
            for diretorio in diretorios:
                caminho_completo = os.path.join(diretorio, arquivo_pdf)
                if os.path.exists(caminho_completo):
                    mail.Attachments.Add(Source=caminho_completo)
                    arquivos_anexados += 1
                    break

        print(f"📎 {arquivos_anexados} PDFs anexados.")
        mail.Send()
        print("📧 E-mail enviado com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

    finally:
        pythoncom.CoUninitialize()


# ---------- Execução principal ----------

def main():
    registros = []
    for pasta in diretorios:
        if not os.path.exists(pasta):
            print(f"📁 Diretório não encontrado: {pasta}")
            continue

        for arquivo in os.listdir(pasta):
            if arquivo.lower().endswith(".pdf"):
                caminho_pdf = os.path.join(pasta, arquivo)
                print(f"📄 Processando: {arquivo}")
                registro = processar_pdf(caminho_pdf)
                registros.append(registro)

    df = pd.DataFrame(registros)
    # INSERIR NO BANCO
    inserir_no_banco(df)
    enviar_email_guias_emitidas(df, diretorios)


if __name__ == "__main__":
    main()
