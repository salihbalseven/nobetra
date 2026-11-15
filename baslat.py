"""Start the source edition with the existing program data folder."""
from pathlib import Path
import traceback
from tkinter import messagebox


def main():
    root = Path(__file__).resolve().parent.parent
    database = root / 'veri' / 'nobet.db'
    if not database.is_file():
        messagebox.showerror('NÖBETRA', 'Mevcut veri dosyası bulunamadı. Başlatıcıyı NÖBETRA klasöründen açın.')
        return
    try:
        import app
        app.ROOT = root
        app.DATA = database.parent
        app.DB = database
        app.run()
    except Exception:
        messagebox.showerror('NÖBETRA başlatılamadı', traceback.format_exc())


if __name__ == '__main__':
    main()
