from __future__ import annotations

import os

import psycopg
from psycopg import sql
from psycopg.conninfo import conninfo_to_dict, make_conninfo

from input_store import initialize_schema


def main() -> None:
    database_url = os.environ.get("DATABASE_URL", "").strip()
    if not database_url:
        raise SystemExit("Configura DATABASE_URL antes de preparar PostgreSQL.")
    database_name = conninfo_to_dict(database_url).get("dbname")
    if not database_name:
        raise SystemExit("DATABASE_URL debe incluir el nombre de la base de datos.")

    maintenance_url = make_conninfo(database_url, dbname="postgres")
    try:
        with psycopg.connect(maintenance_url, autocommit=True) as connection:
            exists = connection.execute(
                "SELECT 1 FROM pg_database WHERE datname = %s",
                (database_name,),
            ).fetchone()
            if not exists:
                connection.execute(
                    sql.SQL("CREATE DATABASE {}").format(
                        sql.Identifier(database_name)
                    )
                )
    except psycopg.Error as exc:
        raise SystemExit(
            "No se pudo crear o consultar la base de datos configurada."
        ) from exc

    initialize_schema()
    print(f"Base de datos lista: {database_name}")


if __name__ == "__main__":
    main()
