"""Prepare a duty roster from an ordered checklist or copied Excel rows."""

from collections import Counter, defaultdict
import csv
import datetime as dt
import io
import json
import re
import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk

import ceza_nobet
import excel_devriye
import nobet_planlayici
import nobet_yazdir
import pencere_yonetimi
from ogrenci_profili import norm, search


FLOORS = ('2A', '2B', '2C', '3A', '3B', '3C', '4A', '4B')
BLOCKS = ('A BLOK', 'B BLOK')
SOURCE = 'Manuel hazırlama'


def layout(kind):
    if kind == 'Yatakhane Kat':
        times, places, reserves = nobet_yazdir.NIGHT, FLOORS, FLOORS * 2
    elif kind == 'Devriye':
        times, places = nobet_yazdir.NIGHT, excel_devriye.PLACES
        reserves = tuple(f'{i + 1}. YEDEK' for i in range(excel_devriye.RESERVES))
    elif kind == 'Yatakhane Giriş':
        times, places, reserves = nobet_yazdir.GATE, BLOCKS, BLOCKS * 2
    else:
        raise ValueError('Geçersiz nöbet türü.')
    return ([(slot, place, 'Asıl') for slot in times for place in places]
            + [('Yedek', place, 'Yedek') for place in reserves])


def roster_state(c, day, kind):
    sheet = c.execute('SELECT id,status FROM sheets WHERE day=? AND kind=?', (day, kind)).fetchone()
    if not sheet:
        return None, []
    if sheet['status'] == 'Onaylandı':
        raise ValueError('Bu çizelge onaylı. Düzenlemek için önce onayını kaldırın.')
    return sheet['id'], [dict(row) for row in c.execute(
        'SELECT * FROM assignments WHERE sheet_id=? ORDER BY id', (sheet['id'],))]


def available_slots(c, day, kind):
    sheet_id, rows = roster_state(c, day, kind)
    if rows and c.execute('SELECT 1 FROM planner_runs WHERE sheet_id=?', (sheet_id,)).fetchone():
        raise ValueError('Bu tarih ve türde otomatik taslak var. Manuel hazırlamak için boş bir çizelge açın.')
    return (sheet_id, rows), unoccupied(kind, rows)


def unoccupied(kind, rows):
    if kind == 'Devriye' and nobet_yazdir.legacy_patrol(rows):
        if not any(row['source_cell'] == SOURCE for row in rows):
            return []
    occupied = occupied_positions(kind, rows)
    return [slot for index, slot in enumerate(layout(kind)) if index not in occupied]


def occupied_positions(kind, rows):
    slots = layout(kind)
    occupied = {}
    fallback = []
    for row in rows:
        role = row['role']
        time = 'Yedek' if role == 'Yedek' else nobet_yazdir.slot(row)
        place = str(row['location']).strip().upper()
        if kind == 'Yatakhane Kat':
            match = re.fullmatch(r'([234])\s*([ABC])', place) or re.search(r'([234]).*?\b([ABC])\b', place)
            if match:
                place = ''.join(match.groups())
        elif kind == 'Yatakhane Giriş' and place in ('A', 'B'):
            place += ' BLOK'
        elif kind == 'Devriye' and role == 'Asıl':
            if 'HABER' in place:
                place = 'HABER M.'
            elif re.fullmatch(r'[1-5]\s*D', place):
                place = place[0] + ' D'
        if kind == 'Devriye' and role == 'Yedek':
            place = place if re.fullmatch(r'[1-6]\. YEDEK', place) else ''
        duty = time, place, role
        position = row.get('manual_position')
        if isinstance(position, int) and 0 <= position < len(slots) and slots[position] == duty and position not in occupied:
            occupied[position] = row
        else:
            fallback.append((duty, row))
    for duty, row in fallback:
        if kind == 'Devriye' and duty[2] == 'Yedek':
            exact = next((i for i, slot in enumerate(slots) if i not in occupied and slot == duty), None)
            if exact is not None:
                occupied[exact] = row
                continue
            matches = lambda slot: slot[2] == 'Yedek'
        else:
            matches = lambda slot: slot == duty
        position = next((i for i, slot in enumerate(slots) if i not in occupied and matches(slot)), None)
        if position is None:
            raise ValueError('Mevcut taslakta şablonla uyuşmayan saat veya görev yeri var. Ana listedeki ilgili satırı kontrol edin.')
        occupied[position] = row
    return occupied


def preparation_warning(kind, rows):
    if not any(row['source_cell'] == SOURCE for row in rows):
        return ''
    try:
        remaining = len(unoccupied(kind, [dict(row) for row in rows]))
    except ValueError as error:
        return str(error)
    if remaining:
        return f'Manuel liste hazırlanıyor · {remaining} görev yeri boş. Onay için listeyi tamamlayın.'
    return ''


def require_complete(c, sheet):
    rows = c.execute('SELECT * FROM assignments WHERE sheet_id=? ORDER BY id', (sheet['id'],)).fetchall()
    warning = preparation_warning(sheet['kind'], rows)
    if warning:
        raise ValueError(f"{sheet['kind']} · {warning}")


def copied_time(text, kind):
    text = norm(text).replace('–', '-').replace('—', '-').replace('.', ':').replace(' ', '')
    if not text:
        return None
    if re.fullmatch(r'(?:\d+[:\-]?)?yedek(?:[:\-]?\d+)?', text):
        return 'Yedek'
    match = re.fullmatch(r'(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})', text)
    if match:
        a, b, c, d = (int(part) for part in match.groups())
        slot = f'{a:02}:{b:02}-{c:02}:{d:02}'
        if any(slot == entry[0] for entry in layout(kind)):
            return slot
    raise ValueError(f'Bu nöbet türü için saat/görev geçersiz: {text}')


def parse_paste(text, students, kind, used, time_parser=copied_time):
    by_number = {str(row['school_no']): row['id'] for row in students.values()}
    by_name = defaultdict(list)
    for row in students.values():
        by_name[' '.join(norm(f"{row['name']} {row['surname']}").split())].append(row['id'])
    indices = {}
    selected, errors = [], []
    seen = set(used)
    for line, cells in enumerate(csv.reader(io.StringIO(text), delimiter='\t'), 1):
        cells = [cell.strip() for cell in cells]
        if not any(cells):
            continue
        labels = [re.sub(r'[^a-z0-9]+', '', norm(cell)) for cell in cells]
        header = {field: index for index, label in enumerate(labels) for field, options in {
            'number': ('sicilokulno', 'okulno', 'sicil', 'sicilno'),
            'given': ('adi', 'ad'), 'surname': ('soyadi', 'soyad'),
            'full': ('adsoyad', 'adisoyadi'), 'time': ('saatgorev', 'saat', 'vardiya'),
        }.items() if label in options}
        if ('number' in header or 'full' in header or {'given', 'surname'} <= header.keys()
                or len(cells) == 1 and header.keys() == {'given'}):
            indices = header
            continue
        if indices:
            fields = {field: cells[index] if index < len(cells) else '' for field, index in indices.items()}
            number = fields.get('number', '')
            name = fields.get('full') or f"{fields.get('given', '')} {fields.get('surname', '')}"
            time = fields.get('time', '')
        elif len(cells) >= 5:
            name, number = ' '.join(cells[:2]), cells[4]
            time = cells[6] if len(cells) >= 7 else ''
        elif re.fullmatch(r'\d+(?:\.0+)?', cells[0]):
            number, name = cells[0], ''
            time = cells[1] if len(cells) == 2 else ''
        else:
            number, name, time = '', ' '.join(cells[:2]), ''
        number = re.sub(r'\.0+$', '', number)
        name = ' '.join(name.split())
        key = number or name
        try:
            matches = [by_number[number]] if number in by_number else (
                by_name.get(norm(name), []) if not number else [])
            if not matches:
                raise ValueError(f'Öğrenci bulunamadı: {key}')
            if len(matches) > 1:
                raise ValueError(f'Aynı ad soyad birden fazla kişide var: {key}. Okul numarasını kullanın.')
            student_id = matches[0]
            if student_id in seen:
                raise ValueError(f'Öğrenci listede zaten var: {key}')
            selected.append((student_id, time_parser(time, kind)))
            seen.add(student_id)
        except ValueError as error:
            errors.append(f'{line}. satır: {error}')
    if errors:
        raise ValueError('\n'.join(errors))
    if not selected:
        raise ValueError('Yapıştırılan alanda öğrenci satırı bulunamadı.')
    return selected


def arrange(selected, slots):
    free = dict(enumerate(slots))
    positions = {}
    for queue_index, (_, requested) in enumerate(selected):
        if isinstance(requested, int):
            if requested not in free:
                raise ValueError(f'{requested + 1}. görev yeri dolu veya artık kullanılamıyor.')
            positions[queue_index] = requested, free.pop(requested)
    for queue_index, (_, requested) in enumerate(selected):
        if requested is None or queue_index in positions:
            continue
        slot_index = next((index for index, slot in free.items() if slot[0] == requested), None)
        if slot_index is None:
            raise ValueError(f'{queue_index + 1}. öğrenci için {requested} sırasında boş yer kalmadı.')
        positions[queue_index] = slot_index, free.pop(slot_index)
    for queue_index, (_, requested) in enumerate(selected):
        if requested is not None:
            continue
        if not free:
            raise ValueError(f'Bu çizelgede en fazla {len(slots)} öğrenci daha eklenebilir.')
        slot_index = next(iter(free))
        positions[queue_index] = slot_index, free.pop(slot_index)
    return [positions[i] for i in range(len(selected))]


def exemptions(c, day):
    return {row[0] for row in c.execute("""SELECT DISTINCT student_id FROM exemptions
        WHERE start<=? AND (end IS NULL OR end='' OR end>=?)""", (day, day))}


def penalty_or_warning(c, student_id, day, kind, role):
    """Find a payable obligation or describe why an exempt duty needs consent."""
    cases = c.execute("""SELECT * FROM penalties WHERE student_id=? AND start_day<=?
        AND active=1 AND archived=0 AND voided=0 ORDER BY start_day,id""", (student_id, day))
    warnings = []
    for case in cases:
        done, reserved = ceza_nobet.counts(c, case['id'])
        if done >= case['total']:
            continue
        if case['target'] != kind:
            warnings.append(f'Kalan cezası {case["target"]} nöbetine tanımlı.')
        elif role != 'Asıl':
            warnings.append('Yedek görev cezasından düşmez.')
        elif reserved >= case['total']:
            warnings.append('Kalan ceza görevleri başka çizelgelerde planlanmış.')
        else:
            return case['id'], ''
    return None, (warnings[0] if warnings else 'Kalan aktif cezası görünmüyor.')

def replacement_rows(c,assignment_id):
    duty=c.execute("""SELECT a.*,s.day,s.kind,s.status,st.name,st.surname FROM assignments a
        JOIN sheets s ON s.id=a.sheet_id JOIN students st ON st.id=a.student_id WHERE a.id=?""",
        (assignment_id,)).fetchone()
    if not duty or duty['status']=='Onaylandı' or duty['result']!='Planlandı':
        raise ValueError('Değişim için görevin taslakta ve Planlandı durumunda olması gerekir.')
    return duty

def replacement_options(app,c,duty):
    kind,day=duty['kind'],duty['day']
    busy=app.busy_students(c,day,kind,duty['start_at'] or '',duty['end_at'] or '')
    busy.update(row[0] for row in c.execute('SELECT student_id FROM assignments WHERE sheet_id=?',(duty['sheet_id'],)))
    busy.update(nobet_planlayici.resting_reserves(c,day,kind))
    penalty=duty['duty_type']=='Ceza'
    if not penalty and kind in nobet_planlayici.NIGHT_KINDS:
        busy.update(nobet_planlayici.worked_previous_night(c,day))
    options=[]
    for student in c.execute('SELECT * FROM students WHERE active=1 ORDER BY class,id'):
        if student['id'] in busy:
            continue
        if app.eligible(c,student,day,kind,duty['slot'],allow_exempt=penalty):
            continue
        pid=penalty_or_warning(c,student['id'],day,kind,duty['role'])[0] if penalty else None
        if penalty and not pid:
            continue
        options.append((student,pid))
    return options

def replace_duty(app,original,student_id):
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        duty=replacement_rows(c,original['id'])
        if dict(duty)!=dict(original):
            raise ValueError('Görev değişti. Listeyi yenileyip tekrar deneyin.')
        options={st['id']:pid for st,pid in replacement_options(app,c,duty)}
        if student_id not in options:
            raise ValueError('Öğrenci artık bu görev için uygun değil.')
        sheet_id=duty['sheet_id']
        run=c.execute('SELECT * FROM planner_runs WHERE sheet_id=?',(sheet_id,)).fetchone()
        if run:
            nobet_planlayici.before_confirm(c,{'id':sheet_id})
        c.execute('UPDATE assignments SET student_id=?,penalty_id=? WHERE id=?',
                  (student_id,options[student_id],duty['id']))
        if run:
            rows=[dict(row) for row in c.execute('SELECT * FROM assignments WHERE sheet_id=?',(sheet_id,))]
            after=run['after_state']
            if duty['kind'] in nobet_planlayici.NIGHT_KINDS and duty['role']=='Asıl' and duty['duty_type']=='Normal':
                state=json.loads(run['before_state'])
                written={row['student_id'] for row in rows if row['role']=='Asıl' and row['duty_type']=='Normal'}
                nobet_planlayici.advance_manual(app,nobet_planlayici.ordered_pool(app,c,state),state,written)
                after=json.dumps(state)
            c.execute('UPDATE planner_runs SET after_state=?,signature=? WHERE sheet_id=?',
                      (after,nobet_planlayici.signature(rows),sheet_id))
        app.log(c,'Taslakta öğrenci değiştirildi',f'{duty["day"]}; {duty["kind"]}; {duty["student_id"]} → {student_id}')


def save(app, day, kind, original, slots, selected, accepted=None):
    if not selected:
        raise ValueError('Önce en az bir öğrenci seçin.')
    if len({student_id for student_id, _ in selected}) != len(selected):
        raise ValueError('Aynı öğrenci birden fazla sırada seçilmiş.')
    assignments = arrange(selected, slots)
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        state, current_slots = available_slots(c, day, kind)
        if state != original or current_slots != slots:
            raise ValueError('Çizelge değişti. Pencereyi yeniden açıp güncel yerleri kontrol edin.')
        sheet_id, existing = state
        occupied = occupied_positions(kind, existing)
        free_positions = [i for i in range(len(layout(kind))) if i not in occupied]
        used = {row['student_id'] for row in existing}
        exempt = exemptions(c, day)
        pending, errors = [], []
        busy_by_interval = {}
        for number, ((student_id, _), (position, (slot, place, role))) in enumerate(zip(selected, assignments), 1):
            student = c.execute('SELECT * FROM students WHERE id=? AND active=1', (student_id,)).fetchone()
            if not student:
                errors.append(f'{number}. öğrenci artık aktif değil.')
                continue
            actual_day = day
            if kind != 'Yatakhane Giriş' and role == 'Asıl' and slot != nobet_yazdir.NIGHT[0]:
                actual_day = (dt.date.fromisoformat(day) + dt.timedelta(days=1)).isoformat()
            row = nobet_planlayici.make_row(student, actual_day, slot, place, role)
            interval = row['start_at'], row['end_at']
            if interval not in busy_by_interval:
                busy_by_interval[interval] = app.busy_students(c, day, kind, *interval)
            is_exempt = student_id in exempt
            reason = app.eligible(c, student, day, kind, slot, allow_exempt=is_exempt)
            penalty_id = None
            if is_exempt:
                penalty_id, warning = penalty_or_warning(c, student_id, day, kind, role)
                if warning and (accepted or {}).get(student_id) != warning:
                    reason = reason or f'{warning} Bu göreve eklemeyi önce onaylayın.'
            if student_id in used or student_id in busy_by_interval[interval]:
                reason = reason or 'Aynı nöbet grubunda görevli veya sınıfı diğer gece listesinde.'
            if reason:
                errors.append(f'{number}. {row["name"]} {row["surname"]} · {slot} · {reason}')
            pending.append((free_positions[position], student_id, slot, place, role, *interval,
                            'Ceza' if penalty_id else 'Normal', penalty_id))
        if errors:
            raise ValueError('Hiçbir görev kaydedilmedi.\n\n' + '\n'.join(errors))
        if sheet_id is None:
            sheet_id = c.execute('INSERT INTO sheets(day,kind) VALUES(?,?)', (day, kind)).lastrowid
        elif not existing:
            c.execute('DELETE FROM planner_runs WHERE sheet_id=?', (sheet_id,))
        c.executemany('''INSERT INTO assignments
            (sheet_id,student_id,slot,location,role,start_at,end_at,duty_type,penalty_id,source_cell,manual_position)
            VALUES(?,?,?,?,?,?,?,?,?,?,?)''', ((sheet_id, *row, SOURCE, position) for position, *row in sorted(pending)))
        app.log(c, 'Manuel nöbet taslağı', f'{day} · {kind} · {len(pending)} öğrenci')
        overrides = [str(row[1]) for row in pending if row[1] in exempt and row[-1] is None]
        if overrides:
            app.log(c, 'Muaf öğrenciye manuel görev', f'{day} · {kind} · öğrenci kayıtları ' + ', '.join(overrides))
    return len(pending)


def night_history(c, day, kind):
    stats = {}
    for row in c.execute("""SELECT a.student_id,a.slot,a.start_at,a.end_at,a.duty_type,s.day
        FROM assignments a JOIN sheets s ON s.id=a.sheet_id
        WHERE s.kind=? AND s.day<? AND s.status='Onaylandı' AND a.role='Asıl' AND a.result='Tuttu'
        ORDER BY COALESCE(NULLIF(a.start_at,''),s.day),a.id""", (kind, day)):
        times, _, penalties = stats.setdefault(row['student_id'], (Counter(), '', Counter()))
        slot = nobet_yazdir.slot(dict(row))
        is_penalty = row['duty_type'] == 'Ceza'
        if slot in nobet_yazdir.NIGHT:
            times[slot] += 1
            if is_penalty:
                penalties[slot] += 1
        last = slot if slot in nobet_yazdir.NIGHT else 'Saat belirsiz'
        stats[row['student_id']] = times, last + (' · Ceza' if is_penalty else ''), penalties
    return stats


def open_dialog(parent, app, day, kind):
    dt.date.fromisoformat(day)
    if kind == 'Yatakhane Giriş' and dt.date.fromisoformat(day).isoweekday() > 5:
        raise ValueError('Hafta sonu giriş nöbeti yok.')
    with app.conn() as c:
        original, slots = available_slots(c, day, kind)
        students = {row['id']: row for row in search(c, '')}
        exempt = exemptions(c, day)
        night = kind in ('Yatakhane Kat', 'Devriye')
        history = night_history(c, day, kind) if night else {}
    if not students:
        raise ValueError('Önce öğrenci listesini aktarın.')
    if not slots:
        raise ValueError('Bu çizelgedeki tüm görev yerleri dolu.')
    existing_ids = {row['student_id'] for row in original[1]}
    full_layout = layout(kind)
    occupied = occupied_positions(kind, original[1])
    free_positions = [index for index in range(len(full_layout)) if index not in occupied]
    free_indices = {position: index for index, position in enumerate(free_positions)}
    placed = {}
    selected = []
    accepted = {}
    penalty_ids = set()

    window = tk.Toplevel(parent)
    window.title('Manuel nöbet hazırla')
    window.geometry(f'{min(1360 if night else 1160, parent.winfo_screenwidth() - 60)}x{min(710, parent.winfo_screenheight() - 100)}')
    window.minsize(830, 480)
    window.transient(parent)
    body = ttk.Frame(window, padding=12)
    body.pack(fill='both', expand=True)
    ttk.Label(body, text=f'{day}  ·  {kind}', font=('Segoe UI', 13, 'bold')).pack(anchor='w')
    ttk.Label(body, text='Sırayla eklemek için öğrencileri işaretleyin. Belirli bir görev için soldan bir öğrenci, sağdan bir sıra seçin.').pack(anchor='w', pady=(4, 8))
    footer = ttk.Frame(body, padding=(0, 6, 0, 0))
    footer.pack(side='bottom', fill='x')
    summary = ttk.Label(footer, wraplength=430)
    summary.pack(side='left')
    columns = ttk.Panedwindow(body, orient='horizontal')
    columns.pack(fill='both', expand=True)
    left = ttk.LabelFrame(columns, text='Öğrenci listesi', padding=8)
    right = ttk.LabelFrame(columns, text='Görev yerleri', padding=8)
    columns.add(left, weight=2 if night else 1)
    columns.add(right, weight=1)

    query = tk.StringVar()
    find = ttk.Entry(left, textvariable=query)
    find.pack(fill='x', pady=(0, 4))
    ttk.Label(left, text='Ad, okul numarası veya sınıf ile ara').pack(anchor='w', pady=(0, 4))
    bulk = ttk.Frame(left)
    bulk.pack(fill='x', pady=(0, 4))
    reverse_order = tk.BooleanVar(master=window, value=False)
    selection_options = ttk.Frame(left)
    selection_options.pack(fill='x')
    ttk.Label(selection_options, text='Shift ile aralık, Ctrl ile ayrı satırlar seçin.').pack(side='left')
    headers = ('Seç', 'Okul no', 'Ad Soyad', 'Sınıf')
    widths = (42, 90, 160, 55)
    if night:
        headers += (*nobet_yazdir.NIGHT, 'Son tuttuğu saat')
        widths += (76, 76, 76, 76, 136)
    catalog = parent.tree(left, headers, widths, wide_scroll=True)
    catalog.tag_configure('exempt', foreground='#a61b1b', background='#fee2e2')
    ttk.Label(left, text='Kırmızı satırlar seçili tarihte nöbetten muaf.').pack(anchor='w')
    if night:
        ttk.Label(left, text='Saat sütunları toplam/ceza biçiminde. Örneğin 3/1, 3 nöbetin 1’i ceza.').pack(anchor='w')
    controls = ttk.Frame(right)
    controls.pack(side='bottom', fill='x', pady=(5, 0))
    preview = parent.tree(right, ('Sıra', 'Ad Soyad', 'Saat', 'Yer', 'Rol', 'Görev tipi'), (42, 160, 95, 65, 55, 75),
                          wide_scroll=True)
    for column in preview['columns']:
        preview.heading(column, command='')

    def refresh_catalog(*_):
        marked = {student_id for student_id, _ in selected}
        tokens = norm(query.get()).split()
        parent.clear(catalog)
        ordered = reversed(students) if reverse_order.get() else iter(students)
        for student_id in ordered:
            row = students[student_id]
            if student_id in existing_ids:
                continue
            name = f"{row['name']} {row['surname']}"
            fields = norm(f"{name} {row['school_no']} {row['class']}")
            if all(token in fields for token in tokens):
                values = ('✓' if student_id in marked else '☐', row['school_no'], name, row['class'])
                if night:
                    counts, last, penalty_counts = history.get(student_id, (Counter(), '—', Counter()))
                    values += (*(f'{counts[slot]}/{penalty_counts[slot]}' for slot in nobet_yazdir.NIGHT), last)
                catalog.insert('', 'end', iid=str(student_id), values=values,
                               tags=('exempt',) if student_id in exempt else ())

    def refresh(focus_id=None):
        positions = arrange(selected, slots)
        placed.clear()
        placed.update({free_positions[position]: student_id for (student_id, _), (position, _) in zip(selected, positions)})
        parent.clear(preview)
        for position, (slot, place, role) in enumerate(full_layout):
            existing = occupied.get(position)
            student = students.get(existing['student_id'] if existing else placed.get(position))
            name = f"{student['name']} {student['surname']}" if student else ('Kayıtlı öğrenci' if existing else '—')
            if role == 'Yedek' and kind != 'Devriye':
                block_count = len(FLOORS if kind == 'Yatakhane Kat' else BLOCKS)
                slot = f'{1 if position < len(full_layout) - block_count else 2}. Yedek'
            duty_type = existing['duty_type'] if existing else (
                'Ceza' if placed.get(position) in penalty_ids else 'Normal' if student else '')
            preview.insert('', 'end', iid=str(position), values=(position+1, name, slot, place, role, duty_type))
        if focus_id is not None and preview.exists(focus_id):
            preview.selection_set(focus_id)
            preview.focus(focus_id)
            preview.see(focus_id)
        remaining = len(slots) - len(selected)
        suffix = ' · Liste tamamlanana kadar onay kapalı' if remaining else ''
        summary.config(text=f'{len(selected)} seçili · {remaining} boş yer{suffix}')
        save_button.state(['!disabled'] if selected else ['disabled'])
        refresh_catalog()

    def update(items, focus_id=None):
        try:
            assignments = arrange(items, slots)
        except ValueError as error:
            messagebox.showerror('Yerleştirme yapılamadı', str(error), parent=window)
            return False
        warnings, penalties = {}, set()
        with app.conn() as c:
            current_exempt = exemptions(c, day)
            for (sid, _), (_, (_, _, role)) in zip(items, assignments):
                if sid in current_exempt:
                    pid, warning = penalty_or_warning(c, sid, day, kind, role)
                    if pid:
                        penalties.add(sid)
                    elif warning:
                        warnings[sid] = warning
        missing_consent = {sid: reason for sid, reason in warnings.items() if accepted.get(sid) != reason}
        if missing_consent:
            details = '\n'.join(f"{students[sid]['name']} {students[sid]['surname']} · {reason}"
                                for sid, reason in missing_consent.items())
            if not messagebox.askyesno('Muaf öğrenciye nöbet', details +
                    '\n\nBu öğrencileri normal görev olarak eklemek istiyor musunuz?', parent=window):
                return False
        selected[:] = items
        exempt.clear()
        exempt.update(current_exempt)
        accepted.clear()
        accepted.update(warnings)
        penalty_ids.clear()
        penalty_ids.update(penalties)
        refresh(focus_id)
        return True

    def toggle(student_id):
        if any(sid == student_id for sid, _ in selected):
            update([(sid, slot) for sid, slot in selected if sid != student_id])
        else:
            update([*selected, (student_id, None)])
        if catalog.exists(str(student_id)):
            catalog.focus(str(student_id))
            catalog.selection_set(str(student_id))
            catalog.see(str(student_id))

    def click(event):
        item = catalog.identify_row(event.y)
        if item and catalog.identify_column(event.x) == '#1':
            toggle(int(item))
            catalog.focus_set()
            return 'break'

    def toggle_focus(event):
        if catalog.focus():
            toggle(int(catalog.focus()))
        return 'break'

    def add_highlighted():
        marked = {student_id for student_id, _ in selected}
        chosen = set(catalog.selection())
        additions = [(int(item), None) for item in catalog.get_children()
                     if item in chosen and int(item) not in marked]
        if additions:
            update([*selected, *additions])

    def change_order(offset):
        if not preview.selection():
            return
        item = preview.selection()[0]
        position = int(item)
        student_id = placed.get(position)
        if student_id is None:
            return
        index = next(i for i, (sid, _) in enumerate(selected) if sid == student_id)
        if not isinstance(selected[index][1], int) and 0 <= index + offset < len(selected):
            target = index + offset
            items = selected.copy()
            items[index], items[target] = items[target], items[index]
            next_slot = arrange(items, slots)[target][0]
            update(items, str(free_positions[next_slot]))
        elif isinstance(selected[index][1], int):
            next_position = position + offset
            target = free_indices.get(next_position)
            if target is None:
                return
            other = placed.get(next_position)
            items = [(sid, target if sid == student_id else free_indices[position] if sid == other else request)
                     for sid, request in selected]
            update(items, str(next_position))

    def remove():
        remove_ids = {placed.get(int(item)) for item in preview.selection()}
        update([(sid, slot) for sid, slot in selected if sid not in remove_ids])

    def place_student():
        if len(catalog.selection()) != 1 or len(preview.selection()) != 1:
            messagebox.showinfo('Sıraya öğrenci yerleştir', 'Soldan bir öğrenci ve sağdan bir görev sırası seçin.', parent=window)
            return
        student_id = int(catalog.selection()[0])
        position = int(preview.selection()[0])
        if position in occupied or (position in placed and placed[position] != student_id):
            messagebox.showwarning('Sıra dolu', 'Boş bir görev sırası seçin.', parent=window)
            return
        items = [(sid, request) for sid, request in selected if sid != student_id]
        update([*items, (student_id, free_indices[position])], str(position))

    def paste_dialog():
        paste = tk.Toplevel(window)
        paste.title('Excel satırlarını yapıştır')
        paste.geometry('720x390')
        pencere_yonetimi.modal(paste,window)
        ttk.Label(paste, text='Adı–Soyadı–Sınıf–Koğuş–Okul No–Kontrol–Saat sütunlarını veya\n'
                  'yalnızca okul numarası ya da ad soyad sütununu yapıştırın.', padding=10).pack(anchor='w')
        text = tk.Text(paste, wrap='none', height=12)
        text.pack(fill='both', expand=True, padx=10)

        def add():
            try:
                rows = parse_paste(text.get('1.0', 'end-1c'), students, kind,
                                   existing_ids | {sid for sid, _ in selected})
            except (ValueError, csv.Error) as error:
                messagebox.showerror('Excel satırlarını kontrol edin', str(error), parent=paste)
                return
            try:
                arrange([*selected, *rows], slots)
            except ValueError as error:
                messagebox.showerror('Yerleştirme yapılamadı', str(error), parent=paste)
                return
            update([*selected, *rows])
            paste.destroy()

        def close():
            paste.destroy()

        buttons = ttk.Frame(paste, padding=10)
        buttons.pack(fill='x')
        ttk.Button(buttons, text='Vazgeç', command=close).pack(side='left')
        ttk.Button(buttons, text='Seçim sırasına ekle', command=add, style='Primary.TButton').pack(side='right')
        paste.protocol('WM_DELETE_WINDOW', close)
        text.focus_set()

    def submit():
        try:
            count = save(app, day, kind, original, slots, selected, accepted)
        except (ValueError, sqlite3.Error) as error:
            messagebox.showerror('Taslak kaydedilemedi', str(error), parent=window)
            return
        window.destroy()
        parent.refresh_all()
        messagebox.showinfo('Manuel taslak kaydedildi', f'{count} görev çizelgeye eklendi.', parent=parent)

    ttk.Button(bulk, text='Seçili satırları topluca ekle', command=add_highlighted,
               style='Primary.TButton').pack(side='left', padx=(0, 6))
    ttk.Button(bulk, text='Excel’den yapıştır', command=paste_dialog).pack(side='left')
    ttk.Checkbutton(selection_options, text='Listeyi ters çevir', variable=reverse_order,
                    command=refresh_catalog).pack(side='left', padx=(8, 0))
    ttk.Button(controls, text='Seçili sıraya yerleştir', command=place_student,
               style='Primary.TButton').pack(side='left', padx=(0, 4))
    ttk.Button(controls, text='Yukarı', command=lambda: change_order(-1)).pack(side='left')
    ttk.Button(controls, text='Aşağı', command=lambda: change_order(1)).pack(side='left', padx=4)
    ttk.Button(controls, text='Seçilenleri çıkar', command=remove).pack(side='left')
    ttk.Button(footer, text='Vazgeç', command=window.destroy).pack(side='right')
    save_button = ttk.Button(footer, text='Tamam · Taslağı kaydet', command=submit, style='Success.TButton', width=22)
    save_button.pack(side='right', padx=6)
    catalog.bind('<Button-1>', click)
    catalog.bind('<space>', toggle_focus)
    preview.bind('<Double-1>', lambda event: place_student() if preview.identify_row(event.y) else None)
    query.trace_add('write', refresh_catalog)
    refresh()
    pencere_yonetimi.modal(window,parent)
    find.focus_set()
