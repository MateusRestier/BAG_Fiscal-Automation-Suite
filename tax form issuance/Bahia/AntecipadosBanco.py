import pyodbc
import os
from datetime import datetime, timedelta
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


def connect_databases():
    """
    Cria a conexão com os bancos DADOSADV e DADOS_EXCEL, 
    retornando os connections e cursors correspondentes.
    """

    print("🔌 Conectando aos bancos...")

    # Conexão com o DADOSADV
    conn_adv = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.getenv('DB_SERVER_ADV')},{os.getenv('DB_PORT_ADV')};"
        f"DATABASE={os.getenv('DB_DATABASE_ADV')};"
        f"UID={os.getenv('DB_USER_ADV')};"
        f"PWD={os.getenv('DB_PASSWORD_ADV')}"
    )

    cursor_adv = conn_adv.cursor()

    # Conexão com o DADOS_EXCEL
    conn_excel = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.getenv('DB_SERVER_EXCEL')},{os.getenv('DB_PORT_EXCEL')};"
        f"DATABASE={os.getenv('DB_DATABASE_EXCEL')};"
        f"UID={os.getenv('DB_USER_EXCEL')};"
        f"PWD={os.getenv('DB_PASSWORD_EXCEL')}"
    )
    cursor_excel = conn_excel.cursor()

    return conn_adv, cursor_adv, conn_excel, cursor_excel


def insert_fc_antecipado_bahia(data_inicio, cursor_adv, cursor_excel, conn_excel):
    """
    Parte 1 - Insere na tabela FC_AntecipadoBahia (POSIPI NOT IN).
    Também remove duplicatas após inserir.
    """
    print("\n=== PARTE 1: Inserindo na tabela FC_AntecipadoBahia ===")

    consulta_bahia = f"""
    SELECT 
        SUM(FT_BASEICM)   AS BASEICMS,
        FT_ALIQICM        AS ALQICMS,
        FT_ALIQIPI        AS ALQIPI,
        FT_NFISCAL        AS NF,
        FT_CHVNFE         AS CHAVE,
        FT_EMISSAO        AS EMISSÃO,
        FT_LOJA           AS LOJA,
        FT_POSIPI         AS NCM,
        FT_VALIPI
    FROM SFT010 s
    WHERE SUBSTRING(FT_FILIAL,1,4) IN ('0110')
      AND FT_ESTADO = 'BA'
      AND FT_EMISSAO >= '{data_inicio}'
      AND FT_POSIPI NOT IN (
           '42029200','42021220','42021210','42021100','83011000',
           '42029200','42021210','42021220','42021900','961700100','42029100'
      )
      AND FT_LOJA IN ('78','79','C7','F5')
    GROUP BY 
        FT_NFISCAL, FT_CHVNFE, FT_ALIQICM, FT_EMISSAO, FT_LOJA, 
        FT_ALIQIPI, FT_POSIPI, FT_VALIPI
    """

    print("📥 Executando consulta no DADOSADV (Bahia - Not IN)...")
    cursor_adv.execute(consulta_bahia)
    dados_bahia = cursor_adv.fetchall()
    qtd_inserida_bahia = len(dados_bahia)

    print(f"📝 Inserindo {qtd_inserida_bahia} linhas na tabela FC_AntecipadoBahia...")

    for row in dados_bahia:
        cursor_excel.execute("""
            INSERT INTO dbo.FC_AntecipadoBahia 
            (BASEICMS, ALQICMS, ALQIPI, NF, CHAVE, EMISSÃO, LOJA, NCM, VALIPI)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8])

    conn_excel.commit()
    print(f"✅ Inserções concluídas. ({qtd_inserida_bahia} linhas inseridas)")

    print("🔍 Verificando duplicatas em FC_AntecipadoBahia...")
    cursor_excel.execute("""
    WITH CTE_Duplicados AS (
        SELECT *,
               ROW_NUMBER() OVER (
                   PARTITION BY 
                       BASEICMS, ALQICMS, ALQIPI, NF, EMISSÃO, LOJA, NCM, VALIPI
                   ORDER BY [Data Insercao]
               ) AS rn
        FROM dbo.FC_AntecipadoBahia
    )
    SELECT COUNT(*) FROM CTE_Duplicados WHERE rn > 1
    """)
    qtd_duplicadas_bahia = cursor_excel.fetchone()[0]
    print(f"♻️ {qtd_duplicadas_bahia} linhas duplicadas encontradas.")

    if qtd_duplicadas_bahia > 0:
        cursor_excel.execute("""
        WITH CTE_Duplicados AS (
            SELECT *,
                   ROW_NUMBER() OVER (
                       PARTITION BY 
                           BASEICMS, ALQICMS, ALQIPI, NF, EMISSÃO, LOJA, NCM, VALIPI
                       ORDER BY [Data Insercao]
                   ) AS rn
            FROM dbo.FC_AntecipadoBahia
        )
        DELETE FROM CTE_Duplicados WHERE rn > 1
        """)
        conn_excel.commit()
        print("🧹 Duplicatas removidas com sucesso em FC_AntecipadoBahia.")
    else:
        print("✅ Nenhuma duplicata encontrada em FC_AntecipadoBahia. Nada foi removido.")


def insert_fc_antecipado_bahia_st(data_inicio, cursor_adv, cursor_excel, conn_excel):
    """
    Parte 2 - Insere na tabela FC_AntecipadoBahiaST (POSIPI IN).
    Também remove duplicatas após inserir.
    """
    print("\n=== PARTE 2: Inserindo na tabela FC_AntecipadoBahiaST ===")

    consulta_bahiast = f"""
    SELECT 
        SUM(FT_BASEICM) AS BASEICMS,
        FT_ALIQICM      AS ALQICMS,
        FT_ALIQIPI      AS ALQIPI,
        FT_NFISCAL      AS NF,
        FT_CHVNFE       AS CHAVE,
        FT_EMISSAO      AS EMISSÃO,
        FT_LOJA         AS LOJA,
        FT_POSIPI       AS NCM
    FROM SFT010 s
    WHERE SUBSTRING(FT_FILIAL,1,4) IN ('0110')
      AND FT_ESTADO ='BA'
      AND FT_EMISSAO >='{data_inicio}'
      AND FT_POSIPI IN (
          '42029200','42021220','42021210','42021100','83011000',
          '42029200','42021210','42021220','42021900','961700100','42029100'
      )
      AND FT_LOJA  IN ('78','79','C7','F5')
    GROUP BY 
        FT_NFISCAL, FT_CHVNFE, FT_ALIQICM, FT_EMISSAO, FT_LOJA, 
        FT_ALIQIPI, FT_POSIPI
    """

    print("📥 Executando consulta no DADOSADV (Bahia - ST - IN)...")
    cursor_adv.execute(consulta_bahiast)
    dados_bahiast = cursor_adv.fetchall()
    qtd_inserida_bahiast = len(dados_bahiast)

    print(f"📝 Inserindo {qtd_inserida_bahiast} linhas na tabela FC_AntecipadoBahiaST...")

    for row in dados_bahiast:
        cursor_excel.execute("""
            INSERT INTO dbo.FC_AntecipadoBahiaST 
            (BASEICMS, ALQICMS, ALQIPI, NF, CHAVE, EMISSÃO, LOJA, NCM)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7])

    conn_excel.commit()
    print(f"✅ Inserções concluídas. ({qtd_inserida_bahiast} linhas inseridas)")

    print("🔍 Verificando duplicatas em FC_AntecipadoBahiaST...")
    cursor_excel.execute("""
    WITH CTE_Duplicados AS (
        SELECT *,
               ROW_NUMBER() OVER (
                   PARTITION BY 
                       BASEICMS, ALQICMS, ALQIPI, NF, EMISSÃO, LOJA, NCM
                   ORDER BY [Data Insercao]
               ) AS rn
        FROM dbo.FC_AntecipadoBahiaST
    )
    SELECT COUNT(*) FROM CTE_Duplicados WHERE rn > 1
    """)
    qtd_duplicadas_bahiast = cursor_excel.fetchone()[0]
    print(f"♻️ {qtd_duplicadas_bahiast} linhas duplicadas encontradas em FC_AntecipadoBahiaST.")

    if qtd_duplicadas_bahiast > 0:
        cursor_excel.execute("""
        WITH CTE_Duplicados AS (
            SELECT *,
                   ROW_NUMBER() OVER (
                       PARTITION BY 
                           BASEICMS, ALQICMS, ALQIPI, NF, EMISSÃO, LOJA, NCM
                       ORDER BY [Data Insercao]
                   ) AS rn
            FROM dbo.FC_AntecipadoBahiaST
        )
        DELETE FROM CTE_Duplicados WHERE rn > 1
        """)
        conn_excel.commit()
        print("🧹 Duplicatas removidas com sucesso em FC_AntecipadoBahiaST.")
    else:
        print("✅ Nenhuma duplicata encontrada em FC_AntecipadoBahiaST. Nada foi removido.")


def update_mva_column(cursor_excel, conn_excel):
    """
    Parte 3 - Atualiza a coluna MVA na tabela FC_AntecipadoBahiaST 
    com base no dicionário de NCMs.
    Envia e-mail caso algum NCM não esteja no dicionário.
    """
    import pythoncom
    from win32com.client import Dispatch

    print("\n📌 Atualizando a coluna MVA com base nos NCMs...")

    mva_dict = {
        '42021210': '94,31',
        '42029200': '94,31',
        '42021220': '94,31',
        '42021900': '94,31',
        '83011000': '87,17',
        '42021100': '94,31',
    }

    cursor_excel.execute("""
        SELECT ID, NCM 
        FROM dbo.FC_AntecipadoBahiaST
        WHERE MVA IS NULL
    """)
    registros = cursor_excel.fetchall()

    qtd_atualizada = 0
    ncms_nao_mapeados = []

    for row in registros:
        id_linha = row[0]
        ncm_original = row[1] if row[1] else ""

        # Debug: ver se há espaços ou chars invisíveis
        print(f"DEBUG: ID={id_linha}, NCM lido do banco = {repr(ncm_original)}, len={len(ncm_original)}")

        # Remove espaços em branco no início/fim
        ncm_limpo = ncm_original.strip()

        # Consulta o dicionário usando o NCM "limpo"
        mva = mva_dict.get(ncm_limpo)
        if mva:
            cursor_excel.execute("""
                UPDATE dbo.FC_AntecipadoBahiaST
                SET MVA = ?
                WHERE ID = ?
            """, mva, id_linha)
            qtd_atualizada += 1
        else:
            ncms_nao_mapeados.append(ncm_original)  # Guarda o valor original para reportar

    conn_excel.commit()
    print(f"✅ MVA atualizado para {qtd_atualizada} linhas.")

    if ncms_nao_mapeados:
        ncms_unicos = sorted(set(ncms_nao_mapeados))
        print("⚠️ Os seguintes NCMs não foram encontrados no dicionário e não tiveram MVA preenchido:")
        for ncm in ncms_unicos:
            print(f"   - [{ncm}]")

        # Envia e-mail com os NCMs não mapeados
        try:
            pythoncom.CoInitialize()
            outlook = Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = "mateus.restier@bagaggio.com.br"
            mail.Subject = "AUTOMÁTICO - 🚨 NCMs sem MVA no preenchimento da FC_AntecipadoBahiaST"
            mail.Body = (
                "Olá,\n\n"
                "Durante a atualização da coluna MVA na tabela FC_AntecipadoBahiaST, "
                "os seguintes NCMs não foram encontrados no dicionário (ou estão com espaços)\n\n" +
                "\n".join(f"- {n}" for n in ncms_unicos) +
                "\n\nVerifique se precisam ser adicionados ao dicionário.\n\n"
                "Atenciosamente,\nAutomação"
            )
            mail.Send()
            print("📧 E-mail enviado com os NCMs não mapeados.")
        except Exception as e:
            print(f"❌ Falha ao enviar e-mail com os NCMs não encontrados: {e}")
        finally:
            pythoncom.CoUninitialize()
    else:
        print("🎉 Todos os NCMs com MVA pendente foram preenchidos com sucesso.")


def update_ie_column(cursor_excel, conn_excel):
    """
    Atualiza a coluna IE nas tabelas FC_AntecipadoBahia e FC_AntecipadoBahiaST
    com base no dicionário de lojas. Envia e-mail caso encontre lojas não mapeadas.
    """
    import pythoncom
    from win32com.client import Dispatch

    print("\n📌 Atualizando a coluna IE com base nas lojas...")

    loja_dict = {
        '78': '209876260',
        '79': '210949735',
        'C7': '207723108',
        'F5': '215810337',
    }

    tabelas = ['FC_AntecipadoBahia', 'FC_AntecipadoBahiaST']
    total_atualizadas = 0

    for tabela in tabelas:
        cursor_excel.execute(f"SELECT ID, LOJA FROM dbo.{tabela} WHERE IE IS NULL")
        registros = cursor_excel.fetchall()
        atualizadas = 0
        lojas_nao_mapeadas = []

        for row in registros:
            id_linha, loja = row
            loja = loja.strip() if loja else None
            ie = loja_dict.get(loja)

            if ie:
                cursor_excel.execute(
                    f"UPDATE dbo.{tabela} SET IE = ? WHERE ID = ?",
                    ie, id_linha
                )
                atualizadas += 1
            else:
                lojas_nao_mapeadas.append(loja)

        conn_excel.commit()
        total_atualizadas += atualizadas
        print(f"✅ Atualizações na tabela {tabela}: {atualizadas}")

        if lojas_nao_mapeadas:
            unicas = sorted(set(filter(None, lojas_nao_mapeadas)))
            print(f"⚠️ Lojas não mapeadas na tabela {tabela}:")
            for loja in unicas:
                print(f"   - {loja}")

            try:
                pythoncom.CoInitialize()
                outlook = Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.To = "mateus.restier@bagaggio.com.br"
                mail.Subject = f"AUTOMÁTICO - 🚨 Lojas não mapeadas na tabela {tabela}"
                mail.Body = (
                    f"Olá,\n\nDurante a execução do script `update_ie_column`, foram encontrados registros na tabela {tabela} com lojas sem IE definido:\n\n"
                    + "\n".join(f"- {loja}" for loja in unicas) +
                    "\n\nVerifique se precisam ser adicionadas ao dicionário de lojas.\n\n"
                    "Atenciosamente,\nAutomação"
                )
                mail.Send()
                print("📧 E-mail enviado com lojas não mapeadas.")
            except Exception as e:
                print(f"❌ Erro ao enviar e-mail de alerta de IE: {e}")
            finally:
                pythoncom.CoUninitialize()
        else:
            print(f"🎉 Todas as lojas na tabela {tabela} foram mapeadas com sucesso.")

    print(f"🏁 Atualização finalizada. Total de linhas atualizadas: {total_atualizadas}")



def update_guiaemitida(cursor_excel, conn_excel):
    """
    Atualiza a coluna GUIAEMITIDA para 0 em todos os registros
    das tabelas FC_AntecipadoBahia e FC_AntecipadoBahiaST onde o valor é nulo.
    """
    print("\n📌 Atualizando a coluna GUIAEMITIDA para 0 onde estiver nula...")
    
    tabelas = ['FC_AntecipadoBahia', 'FC_AntecipadoBahiaST']
    total_atualizadas = 0

    for tabela in tabelas:
        cursor_excel.execute(f"UPDATE dbo.{tabela} SET GUIAEMITIDA = 0 WHERE GUIAEMITIDA IS NULL")
        atualizadas = cursor_excel.rowcount
        total_atualizadas += atualizadas
        print(f"✅ Tabela {tabela}: {atualizadas} linhas atualizadas para 0.")

    conn_excel.commit()
    print(f"🏁 Atualização finalizada. Total de linhas atualizadas: {total_atualizadas}")


def main():
    """
    Função principal que orquestra todo o processo:
      1. Conecta aos bancos
      2. Insere registros em FC_AntecipadoBahia
      3. Insere registros em FC_AntecipadoBahiaST
      4. Atualiza coluna MVA em FC_AntecipadoBahiaST
      5. Atualiza coluna IE em ambas as tabelas
      6. Fecha as conexões
    """
    # Conectar aos bancos
    conn_adv, cursor_adv, conn_excel, cursor_excel = connect_databases()

    # Definir data de início para consultas (exemplo: 7 dias atrás)
    data_inicio = (datetime.now() - timedelta(days=7)).strftime('%Y%m%d')

    # PARTE 1
    insert_fc_antecipado_bahia(data_inicio, cursor_adv, cursor_excel, conn_excel)

    # PARTE 2
    insert_fc_antecipado_bahia_st(data_inicio, cursor_adv, cursor_excel, conn_excel)

    # PARTE 3
    update_mva_column(cursor_excel, conn_excel)

    # PARTE 4
    update_ie_column(cursor_excel, conn_excel)

    # PARTE 5
    update_guiaemitida(cursor_excel, conn_excel)

    # Finalizar
    cursor_adv.close()
    cursor_excel.close()
    conn_adv.close()
    conn_excel.close()

    print("\n🏁 Processo finalizado com sucesso!")


if __name__ == "__main__":
    main()
