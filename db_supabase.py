import os
import json
import logging
import subprocess
import sys
from datetime import datetime

try:
    import psycopg2
    from psycopg2.extras import execute_values
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "psycopg2-binary", "-q"])
    import psycopg2
    from psycopg2.extras import execute_values

try:
    from dotenv import load_dotenv
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-dotenv", "-q"])
    from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger("Senatran.DB")


def _conexao():
    return psycopg2.connect(
        host=os.getenv("SUPABASE_HOST"),
        port=int(os.getenv("SUPABASE_PORTA", 5432)),
        dbname=os.getenv("SUPABASE_BANCO", "postgres"),
        user=os.getenv("SUPABASE_USUARIO"),
        password=os.getenv("SUPABASE_SENHA"),
        sslmode="require",
        connect_timeout=10,
    )


def inserir_multas(multas: list) -> int:
    """
    Insere registros na tabela infracoes.
    Ignora silenciosamente duplicatas (ON CONFLICT DO NOTHING).
    Retorna o número de linhas efetivamente inseridas.
    """
    if not multas:
        return 0

    rows = []
    for m in multas:
        coletado_em = None
        try:
            coletado_em = datetime.strptime(m["coletado_em"], "%Y-%m-%d %H:%M:%S")
        except (ValueError, KeyError):
            pass

        rows.append((
            m.get("id"),
            coletado_em,
            m.get("data_infracao"),
            m.get("placa"),
            m.get("num_infracao"),
            m.get("descricao"),
            m.get("valor"),
            m.get("status"),
            m.get("corpo"),
            json.dumps(m.get("raw", []), ensure_ascii=False),
        ))

    sql = """
        INSERT INTO infracoes
            (id, coletado_em, data_infracao, placa, num_infracao,
             descricao, valor, status, corpo, raw)
        VALUES %s
        ON CONFLICT (id) DO NOTHING
        RETURNING id
    """
    try:
        with _conexao() as conn:
            with conn.cursor() as cur:
                result = execute_values(cur, sql, rows, fetch=True)
            conn.commit()
        inseridos = len(result)
        ignorados = len(rows) - inseridos
        msg = f"☁️  {inseridos} novo(s) registro(s) inserido(s) no Supabase"
        if ignorados:
            msg += f" ({ignorados} duplicata(s) ignorada(s))"
        logger.info(msg + ".")
        return inseridos
    except Exception as e:
        logger.error(f"❌ Supabase — erro ao inserir: {e}")
        return 0
