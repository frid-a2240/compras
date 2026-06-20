import os
from contextlib import contextmanager
from pathlib import Path

import psycopg2
from psycopg2 import pool
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv

# Ruta absoluta al .env — funciona aunque IIS arranque Python con otro CWD.
load_dotenv(Path(__file__).parent / ".env")

_pool: psycopg2.pool.SimpleConnectionPool | None = None


def init_pool() -> psycopg2.pool.SimpleConnectionPool:
    """Inicializa (una sola vez) el pool de conexiones."""
    global _pool
    if _pool is None:
        _pool = psycopg2.pool.SimpleConnectionPool(
            minconn=1,
            maxconn=10,
            host=os.getenv("DB_HOST", "localhost"),
            port=os.getenv("DB_PORT", "5432"),
            dbname=os.getenv("DB_NAME", "compras"),
            user=os.getenv("DB_USER", "postgres"),
            password=os.getenv("DB_PASSWORD", "postgres"),
        )
    return _pool


@contextmanager
def get_cursor(dict_cursor: bool = True):
    """
    Context manager que entrega un cursor listo para usar.
    - Hace commit automático si todo va bien.
    - Hace rollback si hay excepción.
    - Devuelve la conexión al pool al finalizar.

    Args:
        dict_cursor: si True (default), las filas se devuelven como dicts.
                     Si False, se devuelven como tuplas.
    """
    pool_obj = init_pool()
    conn = pool_obj.getconn()
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor) if dict_cursor else conn.cursor()
        try:
            yield cur
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            cur.close()
    finally:
        pool_obj.putconn(conn)


def close_pool() -> None:
    """Cierra todas las conexiones del pool (útil en tests o shutdown)."""
    global _pool
    if _pool is not None:
        _pool.closeall()
        _pool = None