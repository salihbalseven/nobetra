"""Review historical duty records before saving them."""
from datetime import date, datetime, timedelta
from pathlib import Path
from queue import Empty, Queue
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from PIL import Image, ImageTk

import nobet_arsiv as archive
import nobet_fotograf as photo
import pencere_yonetimi


def open_dialog(parent, app):
    kind = parent.kind.get()
    try:
        with app.conn() as c:
            roster = archive.Roster([dict(row) for row in c.execute('SELECT * FROM students')])
    except ValueError as error:
        parent.alert(error)
        return
    window = tk.Toplevel(parent)
    window.title('Geçmiş nöbet listesi yükle')
    window.transient(parent)
    frame = ttk.Frame(window, padding=12)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text=f'{kind}\nTarih, yüklediğiniz çizelgeden okunur.',
              wraplength=700).pack(anchor='w', pady=(0, 8))
    controls = ttk.Frame(frame)
    controls.pack(fill='x')
    for label, callback in [
        ('Excel seç', lambda: select_excel()),
        ('Fotoğraf seç', lambda: select_photo()),
    ]:
        ttk.Button(controls, text=label, command=callback).pack(side='left', padx=(0, 8))
    pencere_yonetimi.modal(window,parent)

    def select_excel():
        path = filedialog.askopenfilename(parent=window, filetypes=[('Excel', '*.xlsx')])
        if not path:
            return
        try:
            lists = archive.excel_lists(path, kind)
        except Exception as error:
            messagebox.showerror('Excel okunamadı', str(error), parent=window)
            return
        review({day: archive.entries(duties, values, roster)
                for day, (duties, values) in lists.items()}, Path(path).name)

    def select_photo():
        path = filedialog.askopenfilename(parent=window, filetypes=[('PNG ve JPEG', '*.png *.jpg *.jpeg')])
        if not path:
            return
        try:
            picture = photo.open_image(path)
        except Exception as error:
            messagebox.showerror('Fotoğraf açılamadı', str(error), parent=window)
            return
        for widget in controls.winfo_children():
            widget.destroy()
        hint = ttk.Label(frame, text='Tablonun dört dış köşesine tıklayın. Sol üst, sağ üst, sağ alt, sol alt.\n'
                                    'Alt köşeleri son nöbetçi/yedek satırından alın. Talimat ve imzacı adlarını dışarıda bırakın.')
        hint.pack(anchor='w', pady=6)
        choices = tk.StringVar(value='5 ekip')
        if kind == 'Devriye':
            ttk.Label(controls, text='Fotoğraftaki düzen').pack(side='left', padx=(0, 5))
            ttk.Combobox(controls, textvariable=choices, state='readonly', width=12,
                         values=('4 ekip', '5 ekip')).pack(side='left')
        canvas = tk.Canvas(frame, highlightthickness=0, background='#dbe3ed')
        canvas.pack(fill='both', expand=True)
        points = []
        width_limit = min(1050, parent.winfo_screenwidth() - 125)
        height_limit = min(650, parent.winfo_screenheight() - 255)
        size = (1, 1)
        def draw():
            nonlocal size
            preview = picture.copy()
            preview.thumbnail((width_limit, height_limit), Image.Resampling.LANCZOS)
            size = preview.size
            canvas.image = ImageTk.PhotoImage(preview, master=window)
            canvas.configure(width=size[0], height=size[1])
            canvas.delete('all')
            canvas.create_image(0, 0, anchor='nw', image=canvas.image)
            points.clear()
        def rotate():
            nonlocal picture
            picture = picture.transpose(Image.Transpose.ROTATE_90)
            draw()
        def mark(event):
            if len(points) == 4 or not (0 <= event.x < size[0] and 0 <= event.y < size[1]):
                return
            points.append((event.x * picture.width / size[0], event.y * picture.height / size[1]))
            canvas.create_oval(event.x-5, event.y-5, event.x+5, event.y+5, fill='#2563eb', outline='white')
            canvas.create_text(event.x+10, event.y+9, text=str(len(points)), fill='#2563eb')
        def read():
            if len(points) != 4:
                messagebox.showwarning('Köşeler eksik', 'Dört köşeyi sırayla işaretleyin.', parent=window)
                return
            five = choices.get() == '5 ekip'
            sections = [archive.positions(kind, five)]
            if kind == 'Yatakhane Giriş':
                sections.append(archive.positions(kind, five, second=True))
            layout = archive.layout(kind, five)
            last = {'Yatakhane Kat': 32, 'Yatakhane Giriş': 21, 'Devriye': 52 if five else 44}[kind]
            corners, image = tuple(points), picture.copy()
            addresses = [address for duties in sections for duty in duties for address in duty.addresses]
            date_groups = {'Yatakhane Kat': [[('A7', 0), ('A17', 0), ('A27', 0), ('A8', 1)]],
                           'Devriye': [[('A3', 0), ('A14' if five else 'A12', 1)]],
                           'Yatakhane Giriş': [[('A3', 0), ('A5', 0)], [('A13', 0), ('A15', 0)]]}[kind]
            status.configure(text='Fotoğraf okunuyor…')
            for widget in controls.winfo_children():
                widget.state(['disabled'])
            result = Queue()
            def worker():
                try:
                    result.put(photo.read_cells(image, corners, layout, last, addresses, date_groups))
                except Exception as error:
                    result.put(error)
            def poll():
                try:
                    output = result.get_nowait()
                except Empty:
                    window.after(120, poll)
                    return
                if isinstance(output, Exception):
                    for widget in controls.winfo_children():
                        widget.state(['!disabled'])
                    status.configure(text='')
                    messagebox.showerror('Fotoğraf okunamadı', str(output), parent=window)
                    return
                for widget in controls.winfo_children():
                    widget.state(['!disabled'])
                status.configure(text='')
                values, dates = output
                lists = {}
                for index, (duties, day) in enumerate(zip(sections, dates), 1):
                    if not day:
                        answer = simpledialog.askstring('Çizelge tarihini doğrula',
                            f'{index}. bölümün tarihi okunamadı. Fotoğraftaki tarihi gün.ay.yıl olarak girin.\n'
                            'Bölüm boşsa alanı boş bırakıp Tamam’a basın.', parent=window)
                        if answer is None:
                            return
                        if not answer.strip():
                            continue
                        try:
                            day = datetime.strptime(answer.strip(), '%d.%m.%Y').date().isoformat()
                        except ValueError:
                            messagebox.showerror('Tarih geçersiz', 'Tarihi gün.ay.yıl olarak yazın.', parent=window)
                            return
                    if day in lists:
                        messagebox.showerror('Tarihleri kontrol edin', 'İki bölüm için aynı tarih okundu.', parent=window)
                        return
                    lists[day] = archive.entries(duties, values, roster)
                if lists:
                    review(lists, Path(path).name, image)
            Thread(target=worker, daemon=True).start()
            window.after(120, poll)
        for label, command in [('Döndür', rotate), ('Köşeleri temizle', draw), ('Fotoğrafı oku', read)]:
            ttk.Button(controls, text=label, command=command).pack(side='left', padx=6)
        status = ttk.Label(frame)
        status.pack(anchor='w', pady=4)
        canvas.bind('<Button-1>', mark)
        draw()

    def review(lists, source, picture=None):
        for widget in frame.winfo_children():
            widget.destroy()
        window.geometry(f'{min(1050,parent.winfo_screenwidth()-100)}x{min(660,parent.winfo_screenheight()-150)}')
        ttk.Label(frame, text=f'{kind} · {source}').pack(anchor='w')
        dates_label = ttk.Label(frame)
        dates_label.pack(anchor='w')
        ttk.Label(frame, text='Öğrenci eşleşmelerini kontrol edin. Asıl görevleri Tuttu veya Tutmadı olarak işaretleyin.').pack(anchor='w', pady=4)
        actions = ttk.Frame(frame)
        actions.pack(fill='x', pady=4)
        rows = {f'{day}:{i}': entry for day, entries in lists.items() for i, entry in enumerate(entries)}
        table = ttk.Frame(frame)
        table.pack(fill='both', expand=True)
        headers = ['Tarih', 'Saat / Görev', 'Yer', 'Okunan bilgi', 'Eşleşen öğrenci / Öneri', 'Sonuç']
        tree = ttk.Treeview(table, columns=headers, show='headings', selectmode='extended')
        scrollbar = ttk.Scrollbar(table, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        tree.pack(side='left', fill='both', expand=True)
        for key, width in zip(headers, (95, 105, 75, 195, 245, 90)):
            tree.heading(key, text=key)
            tree.column(key, width=width, minwidth=60)
        tree.tag_configure('unmatched', foreground='#9a3412')
        status = ttk.Label(frame)
        status.pack(anchor='w', pady=6)
        def refresh():
            selected = tree.selection()
            days = dict.fromkeys(key.split(':')[0] for key in rows)
            dates_label.configure(text='Çizelge tarihi ' + ', '.join(
                date.fromisoformat(day).strftime('%d.%m.%Y') for day in days))
            for item, entry in rows.items():
                duty, student = entry.duty, entry.student
                matched = f"{student['school_no']} · {student['name']} {student['surname']}" if student else 'Eşleşmedi'
                if not student:
                    suggestion = roster.suggest(entry.raw_name, entry.raw_number)
                    if suggestion:
                        matched = 'Öneri · ' + suggestion
                actual_date = date.fromisoformat(item.split(':')[0])
                if kind != 'Yatakhane Giriş' and duty.time in archive.NIGHT[1:]:
                    actual_date += timedelta(days=1)
                values = (actual_date.strftime('%d.%m.%Y'), f'Yedek {duty.reserve}' if duty.reserve else duty.time, duty.place,
                          ' '.join(filter(None, (entry.raw_name, entry.raw_number))), matched, entry.result)
                if tree.exists(item):
                    tree.item(item, values=values, tags=() if student else ('unmatched',))
                else:
                    tree.insert('', 'end', iid=item, values=values, tags=() if student else ('unmatched',))
            tree.selection_set([item for item in selected if tree.exists(item)])
            missing = sum(entry.student is None for entry in rows.values())
            undecided = sum(entry.duty.role == 'Asıl' and entry.result == 'Planlandı' for entry in rows.values())
            status.configure(text=f'{len(rows)} satır · {missing} eşleşmeyen · {undecided} asıl sonuç bekliyor')
        def mark(result, only_primary=False):
            targets = ([entry for entry in rows.values() if not entry.duty.reserve and entry.result == 'Planlandı']
                       if only_primary else [rows[key] for key in tree.selection()])
            for entry in targets:
                entry.result = result
            refresh()
        def edit():
            chosen = tree.selection()
            if len(chosen) != 1:
                messagebox.showinfo('Öğrenciyi düzelt', 'Tablodan bir satır seçin.', parent=window)
                return
            query = simpledialog.askstring('Öğrenciyi düzelt', 'Okul numarası veya tam ad soyad girin.', parent=window)
            if not query:
                return
            query = query.strip()
            student = roster.find('', query) if query.isdigit() else roster.find(query)
            if not student:
                messagebox.showwarning('Öğrenci bulunamadı', 'Öğrencinin okul numarasıyla tekrar deneyin.', parent=window)
                return
            rows[chosen[0]].student = student
            refresh()
        def remove():
            selected = tree.selection()
            if selected and messagebox.askyesno('Satırları çıkar',
                    f'{len(selected)} satır bu aktarımdan çıkarılsın mı?', parent=window):
                for key in selected:
                    del rows[key]
                    tree.delete(key)
                refresh()
        def edit_date():
            dates = {item.split(':')[0] for item in tree.selection()}
            if len(dates) != 1:
                messagebox.showinfo('Fotoğraf tarihini düzelt', 'Tarihini düzelteceğiniz bölümden bir satır seçin.', parent=window)
                return
            old_day = dates.pop()
            answer = simpledialog.askstring('Fotoğraf tarihini düzelt',
                'Çizelgedeki tarihi gün.ay.yıl olarak girin.',
                initialvalue=date.fromisoformat(old_day).strftime('%d.%m.%Y'), parent=window)
            if not answer:
                return
            try:
                new_day = datetime.strptime(answer.strip(), '%d.%m.%Y').date().isoformat()
                if new_day == old_day:
                    return
                if any(key.startswith(new_day + ':') for key in rows):
                    raise ValueError('Bu tarih diğer bölümde kullanılıyor.')
            except ValueError as error:
                messagebox.showerror('Tarihi kontrol edin',
                    str(error) if str(error).startswith('Bu tarih') else 'Tarihi gün.ay.yıl olarak yazın.', parent=window)
                return
            updated = {new_day + key[10:] if key.startswith(old_day + ':') else key: entry
                       for key, entry in rows.items()}
            rows.clear()
            rows.update(updated)
            tree.delete(*tree.get_children())
            refresh()
        def save():
            try:
                groups = {}
                for key, entry in rows.items():
                    groups.setdefault(key.split(':')[0], []).append(entry)
                if not groups:
                    raise ValueError('Kaydedilecek görev kalmadı.')
                if picture and not messagebox.askyesno('Tarihi doğrula',
                        'Fotoğraftan alınan çizelge tarihleri ' + ', '.join(
                            date.fromisoformat(day).strftime('%d.%m.%Y') for day in groups)
                        + '. Bu tarihlere kaydedilsin mi?', parent=window):
                    return
                count = archive.save_lists(app, kind, groups, source)
            except Exception as error:
                messagebox.showerror('Geçmişe kaydedilemedi', str(error), parent=window)
                return
            window.destroy()
            parent.day.set(next(iter(groups)))
            parent.refresh_all()
            messagebox.showinfo('Geçmiş nöbet kaydedildi', f'{count} görev kaydedildi.', parent=parent)
        for label, command in [('Seçilenler Tuttu', lambda: mark('Tuttu')),
                               ('Seçilenler Tutmadı', lambda: mark('Tutmadı')),
                               ('Kalan asılları Tuttu yap', lambda: mark('Tuttu', True)),
                               ('Öğrenciyi düzelt', edit), ('Satırları çıkar', remove)]:
            ttk.Button(actions, text=label, command=command).pack(side='left', padx=(0, 4))
        tree.bind('<Double-1>', lambda event: edit())
        if picture:
            ttk.Button(frame, text='Fotoğraf tarihini düzelt', command=edit_date).pack(side='left', padx=(0, 6))
            def show_source():
                preview = tk.Toplevel(window)
                preview.title(source)
                preview.transient(window)
                image = picture.copy()
                image.thumbnail((parent.winfo_screenwidth()-80, parent.winfo_screenheight()-145),
                                Image.Resampling.LANCZOS)
                label = ttk.Label(preview)
                label.image = ImageTk.PhotoImage(image, master=preview)
                label.configure(image=label.image)
                label.pack()
                def close():
                    preview.destroy()
                ttk.Button(preview, text='Kapat', command=close).pack(pady=5)
                preview.protocol('WM_DELETE_WINDOW', close)
                pencere_yonetimi.modal(preview,window)
            ttk.Button(frame, text='Kaynak fotoğrafı göster', command=show_source).pack(side='left')
        ttk.Button(frame, text='Onayla ve geçmişe kaydet', command=save,
                   style='Success.TButton').pack(side='right')
        refresh()
