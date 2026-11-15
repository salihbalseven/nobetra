"""Bu klasördeki boş veritabanına kurgusal öğrenci ve bir örnek ceza ekler."""
from contextlib import closing
from datetime import date
from itertools import product
from pathlib import Path
import sqlite3

import app


def main():
    demo_dir = Path(__file__).resolve().parent / "veri"
    demo_db = demo_dir / "nobet.db"
    try:
        with demo_db.open("xb"):
            pass
    except FileExistsError:
        raise SystemExit("Bu klasörde zaten veri/nobet.db var. Demo için temiz bir kopya açın.")

    app.DB = demo_db
    app.setup()
    students = [
        (f"DEMO-{index:03d}", f"Kurgusal {index:03d}", "Öğrenci", cls, f"{block}{room}")
        for index, (cls, block, room) in enumerate(
            product(app.CLASSES, "AB", range(201, 209)), start=1
        )
    ]
    with closing(sqlite3.connect(demo_db)) as db, db:
        db.execute("PRAGMA foreign_keys=ON")
        db.executemany(
            "INSERT INTO students(school_no,name,surname,class,room) VALUES (?,?,?,?,?)",
            students,
        )
        db.execute(
            """INSERT INTO penalties(student_id,total,target,start_day)
            SELECT id,2,?,? FROM students WHERE school_no=?""",
            ("Devriye", date.today().isoformat(), "DEMO-001"),
        )
    with app.conn() as db:
        seed_numbers = {
            "Yatakhane Kat": "DEMO-001",
            "Devriye": "DEMO-065",
            "Yatakhane Giriş": "DEMO-129",
        }
        seed_ids = {
            kind: db.execute(
                "SELECT id FROM students WHERE school_no=?", (number,)
            ).fetchone()["id"]
            for kind, number in seed_numbers.items()
        }
        app.nobet_planlayici.configure(app, db, seed_ids)
    print(f"{len(students)} kurgusal öğrenci ve bir örnek ceza eklendi. python app.py ile açın.")


if __name__ == "__main__":
    main()
