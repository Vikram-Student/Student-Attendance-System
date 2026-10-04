import sqlite3
from pathlib import Path


# Get the project root directory
BASE_DIR = Path(__file__).resolve().parent.parent

# Database location
DATABASE_PATH = BASE_DIR / "database" / "attendance.db"


def get_connection():
    """Create and return a database connection."""
    connection = sqlite3.connect(DATABASE_PATH)

    # Allows accessing columns by name
    connection.row_factory = sqlite3.Row

    # Enable foreign key constraints
    connection.execute("PRAGMA foreign_keys = ON")

    return connection


def create_tables():
    """Create all required database tables."""

    connection = get_connection()
    cursor = connection.cursor()

    # Users table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK (
                role IN ('HOD', 'FACULTY', 'STUDENT')
            ),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Students table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            email TEXT UNIQUE,
            branch TEXT NOT NULL,
            year INTEGER NOT NULL,
            section TEXT NOT NULL,
            user_id INTEGER UNIQUE,
            FOREIGN KEY (user_id)
                REFERENCES users(id)
                ON DELETE SET NULL
        )
    """)

    # Faculty table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS faculty (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            faculty_id TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            email TEXT UNIQUE,
            user_id INTEGER UNIQUE,
            FOREIGN KEY (user_id)
                REFERENCES users(id)
                ON DELETE SET NULL
        )
    """)

    # Subjects table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS subjects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subject_code TEXT NOT NULL UNIQUE,
            subject_name TEXT NOT NULL,
            faculty_id INTEGER NOT NULL,
            year INTEGER NOT NULL,
            section TEXT NOT NULL,
            FOREIGN KEY (faculty_id)
                REFERENCES faculty(id)
                ON DELETE CASCADE
        )
    """)

    # Attendance sessions table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS attendance_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subject_id INTEGER NOT NULL,
            faculty_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            period INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (subject_id)
                REFERENCES subjects(id)
                ON DELETE CASCADE,

            FOREIGN KEY (faculty_id)
                REFERENCES faculty(id)
                ON DELETE CASCADE,

            UNIQUE(subject_id, date, period)
        )
    """)

    # Attendance table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            subject_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            status TEXT NOT NULL CHECK (
                status IN ('PRESENT', 'ABSENT')
            ),
            marked_by INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (student_id)
                REFERENCES students(id)
                ON DELETE CASCADE,

            FOREIGN KEY (subject_id)
                REFERENCES subjects(id)
                ON DELETE CASCADE,

            FOREIGN KEY (marked_by)
                REFERENCES faculty(id)
                ON DELETE CASCADE,

            UNIQUE(student_id, subject_id, date)
        )
    """)

    connection.commit()
    connection.close()

    print("Database and tables created successfully.")


if __name__ == "__main__":
    create_tables()