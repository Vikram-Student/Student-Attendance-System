import os
import sqlite3

import psycopg
from psycopg.rows import dict_row
from sqlite3 import Error as SQLiteError
from sqlite3 import IntegrityError as SQLiteIntegrityError
from psycopg import Error as PostgreSQLError
from psycopg import IntegrityError as PostgreSQLIntegrityError


DatabaseError = (SQLiteError, PostgreSQLError)
DatabaseIntegrityError = (
    SQLiteIntegrityError,
    PostgreSQLIntegrityError
)


DATABASE_PATH = os.path.join(
    os.path.dirname(__file__),
    "attendance.db"
)


class DatabaseConnection:
    """
    Compatibility wrapper for SQLite and PostgreSQL.

    The application can continue using:
        connection.execute(sql, params)

    with SQLite-style '?' placeholders.
    """

    def __init__(self, connection, is_postgresql=False):
        self.connection = connection
        self.is_postgresql = is_postgresql

    def execute(self, sql, params=None):
        if self.is_postgresql:
            sql = sql.replace("?", "%s")

        if params is None:
            return self.connection.execute(sql)

        return self.connection.execute(sql, params)

    def executemany(self, sql, params_list):
        if self.is_postgresql:
            sql = sql.replace("?", "%s")
            cursor = self.connection.cursor()
            cursor.executemany(sql, params_list)
            return cursor

        return self.connection.executemany(sql, params_list)

        return self.connection.executemany(sql, params_list)

    def execute_insert_returning_id(self, sql, params=None):
        """
        Execute an INSERT and return the generated primary key.
        """

        if self.is_postgresql:
            sql = sql.replace("?", "%s")

            if "RETURNING" not in sql.upper():
                sql = sql.rstrip().rstrip(";") + " RETURNING id"

            if params is None:
                cursor = self.connection.execute(sql)
            else:
                cursor = self.connection.execute(sql, params)

            return cursor.fetchone()["id"]

        if params is None:
            cursor = self.connection.execute(sql)
        else:
            cursor = self.connection.execute(sql, params)

        return cursor.lastrowid

    def cursor(self):
        return self.connection.cursor()

    def commit(self):
        self.connection.commit()

    def rollback(self):
        self.connection.rollback()

    def close(self):
        self.connection.close()


def get_db_connection():
    """
    Return a database connection.

    Local development:
        SQLite

    Production:
        PostgreSQL when DATABASE_URL is configured.
    """

    database_url = os.getenv("DATABASE_URL")

    if database_url:
        connection = psycopg.connect(
            database_url,
            row_factory=dict_row
        )

        return DatabaseConnection(
            connection,
            is_postgresql=True
        )

    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row

    connection.execute("PRAGMA foreign_keys = ON")

    return DatabaseConnection(
        connection,
        is_postgresql=False
    )
