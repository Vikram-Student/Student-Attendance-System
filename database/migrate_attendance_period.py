import sqlite3

DATABASE_PATH = "database/attendance.db"

connection = sqlite3.connect(DATABASE_PATH)

try:
    connection.execute("PRAGMA foreign_keys = OFF")

    # Rename old attendance table
    connection.execute("""
        ALTER TABLE attendance
        RENAME TO attendance_old
    """)

    # Create new attendance table
    connection.execute("""
        CREATE TABLE attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,

            student_id INTEGER NOT NULL,

            subject_id INTEGER NOT NULL,

            date TEXT NOT NULL,

            period TEXT NOT NULL,

            status TEXT NOT NULL
                CHECK(status IN ('PRESENT', 'ABSENT')),

            marked_by INTEGER NOT NULL,

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (student_id)
                REFERENCES students(id),

            FOREIGN KEY (subject_id)
                REFERENCES subjects(id),

            FOREIGN KEY (marked_by)
                REFERENCES faculty(id),

            UNIQUE (
                student_id,
                subject_id,
                date,
                period
            )
        )
    """)

    # Copy existing attendance records.
    # Existing records are matched with their attendance session
    # to recover the period.
    connection.execute("""
        INSERT INTO attendance
        (
            id,
            student_id,
            subject_id,
            date,
            period,
            status,
            marked_by,
            created_at
        )
        SELECT
            attendance_old.id,
            attendance_old.student_id,
            attendance_old.subject_id,
            attendance_old.date,
            attendance_sessions.period,
            attendance_old.status,
            attendance_old.marked_by,
            attendance_old.created_at

        FROM attendance_old

        JOIN attendance_sessions
            ON attendance_sessions.subject_id =
               attendance_old.subject_id

            AND attendance_sessions.faculty_id =
                attendance_old.marked_by

            AND attendance_sessions.date =
                attendance_old.date
    """)

    # Remove old table
    connection.execute("""
        DROP TABLE attendance_old
    """)

    connection.execute("PRAGMA foreign_keys = ON")

    connection.commit()

    print("Attendance table migration completed successfully.")

except Exception as error:

    connection.rollback()

    print("Migration failed:")
    print(error)

finally:

    connection.close()