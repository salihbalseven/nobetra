"""Read printed duty sheets and record reviewed historical attendance."""
from dataclasses import dataclass
from contextlib import nullcontext
from datetime import date, timedelta
from difflib import get_close_matches
import json
from pathlib import Path
import re

import excel_yatakhane
import ceza_nobet
import nobet_manuel
from excel_devriye import PLACES, RESERVES
from nobet_planlayici import before_confirm, make_row
from nobet_yazdir import GATE, NIGHT, slot
from ogrenci_profili import norm


@dataclass
class Duty:
    name: str
    surname: str
    klass: str
    number: str
    time: str
    place: str
    reserve: int = 0

    @property
    def role(self):
        return 'Yedek' if self.reserve else 'Asıl'

    @property
    def addresses(self):
        return [cell for cell in (self.name, self.surname, self.klass, self.number) if cell]


def positions(kind, five=False, second=False):
    if kind == 'Yatakhane Kat':
        return [
            Duty(f'{col}{r}', f'{chr(ord(col)+1)}{r}', f'{chr(ord(col)+2)}{r}', '',
                 NIGHT[i] if i < 4 else 'Yedek', f'{floor}{block}', max(0, i-3))
            for floor, top in [(2, 7), (3, 17), (4, 27)]
            for i, r in enumerate(range(top, top+6))
            for block, col in zip('ABC', 'CHM') if (floor, block) != (4, 'C')
        ]
    if kind == 'Devriye':
        places = PLACES if five else PLACES[:-2]
        duties = [(time, place, 0) for time in NIGHT for place in places]
        duties += [('Yedek', '', i+1) for i in range(RESERVES)]
        return [Duty(f'D{r}', f'E{r}', f'F{r}', f'G{r}', *duty)
                for r, duty in enumerate(duties, 3)]
    if kind == 'Yatakhane Giriş':
        top = 15 if second else 5
        return [
            Duty(f'{col}{r}', f'{chr(ord(col)+1)}{r}', f'{chr(ord(col)+2)}{r}', '',
                 GATE[i] if i < 5 else 'Yedek', f'{block} BLOK', max(0, i-4))
            for i, r in enumerate(range(top, top+7)) for block, col in zip('AB', 'CH')
        ]
    raise ValueError('Nöbet türü seçin.')


def layout(kind, five):
    return json.loads(Path(__file__).with_name('yazdirma_sablonlari.json').read_text(
        encoding='utf-8'))['Devriye 5 Ekip' if kind == 'Devriye' and five else kind]


def excel_lists(path, kind):
    title = {'Devriye': 'DEVRİYE NÖBET', 'Yatakhane Kat': 'Yatakhane Nöbet',
             'Yatakhane Giriş': 'GİRİŞ NÖBET ÇİZELGESİ'}[kind]
    book = excel_yatakhane.read_book(path)
    sheets = [cells for name, cells in book.items() if norm(name) == norm(title)]
    if len(sheets) != 1:
        raise ValueError(f'Excel’de {title} sayfasını bulamadım.')
    cells = sheets[0]
    def value(address):
        return cells.get(address, {}).get('value', '')
    five = kind == 'Devriye' and any(norm(value(f'C{i}')).replace(' ', '') == '5d'
                                     for i in range(3, 47))
    if kind == 'Yatakhane Giriş':
        anchors = ('A5' if value('A5') else 'A3', 'A15' if value('A15') else 'A13')
    else:
        anchors = ('A3',) if kind == 'Devriye' else ('A7',)
    lists = {}
    for index, anchor in enumerate(anchors):
        item = cells.get(anchor, {})
        duties = positions(kind, five, bool(index))
        values = {address: str(value(address)).strip() for duty in duties for address in duty.addresses}
        rows = [duty for duty in duties if any(values[cell] not in ('', '0') for cell in
                                               (duty.name, duty.surname, duty.number) if cell)]
        if not rows:
            continue
        if not item.get('value') or re.fullmatch(r'\d{1,2}[./]\d{1,2}[./]?', str(item['value']).strip()):
            raise ValueError(f'{title} sayfasında {anchor} için yıl dahil tarih gerekli. Excel’de tarihi kontrol edin.')
        day = excel_yatakhane.resolve_date({anchor: {**item, 'formula': None}}, anchor, date.today().year).isoformat()
        if day in lists:
            raise ValueError('Çizelgede aynı tarihe ait iki bölüm var. Excel’de tarihleri kontrol edin.')
        lists[day] = rows, values
    if not lists:
        raise ValueError('Çizelgede öğrenci okunamadı. Dosyayı Excel’de yeniden hesaplatıp kaydedin.')
    return lists


def excel_rows(path, kind, day):
    return excel_lists(path, kind)[day]


class Roster:
    def __init__(self, students):
        self.by_number = {str(student['school_no']): student for student in students if student['school_no']}
        self.by_name = {}
        for student in students:
            key = norm(f"{student['name']} {student['surname']}")
            self.by_name.setdefault(key, []).append(student)

    def find(self, name, number='', klass=''):
        key = norm(name)
        number = number.strip().removesuffix('.0')
        if number in self.by_number:
            candidate = self.by_number[number]
            if not key or key == norm(f"{candidate['name']} {candidate['surname']}"):
                return candidate
        matches = self.by_name.get(key, [])
        if len(matches) == 1 and (not number or number == str(matches[0]['school_no'])):
            return matches[0]
        same_class = [student for student in matches if klass and norm(student['class']) == norm(klass)
                      and (not number or number == str(student['school_no']))]
        return same_class[0] if len(same_class) == 1 else None

    def suggest(self, name, number=''):
        candidate = self.by_number.get(number.strip().removesuffix('.0'))
        if candidate:
            return f"{candidate['school_no']} · {candidate['name']} {candidate['surname']}"
        nearby = get_close_matches(norm(name), self.by_name, n=1, cutoff=.76)
        if nearby and len(self.by_name[nearby[0]]) == 1:
            student = self.by_name[nearby[0]][0]
            return f"{student['school_no']} · {student['name']} {student['surname']}"
        return ''


@dataclass
class Entry:
    duty: Duty
    raw_name: str
    raw_number: str
    student: object
    result: str = 'Planlandı'


def entries(duties, values, roster):
    result = []
    for duty in duties:
        name = ' '.join(values.get(address, '').strip() for address in (duty.name, duty.surname)).strip()
        number = values.get(duty.number, '').strip()
        number = '' if number == '0' else number
        student = roster.find(name, number, values.get(duty.klass, ''))
        result.append(Entry(duty, name, number, student))
    return result


def save_lists(app, kind, lists, source):
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        return sum(save(app, day, kind, rows, source, connection=c) for day, rows in lists.items())


def save(app, day, kind, rows, source, connection=None):
    if date.fromisoformat(day) > date.today():
        raise ValueError(f'{day} ileri bir tarih. Geçmişe aktarmak için nöbet gününün gelmesini bekleyin veya bu tarihin satırlarını önizlemeden çıkarın.')
    if not rows or any(row.student is None for row in rows):
        raise ValueError('Öğrencisi eşleşmeyen satırları düzeltin veya çizelgede boş olanları çıkarın.')
    ids = [row.student['id'] for row in rows]
    if len(set(ids)) != len(ids):
        raise ValueError('Aynı öğrenci birden fazla satırda bulunuyor. Listeyi kontrol edin.')
    if any(row.duty.role == 'Asıl' and row.result == 'Planlandı' for row in rows):
        raise ValueError('Tüm asıl görevleri Tuttu veya Tutmadı olarak işaretleyin.')
    with (app.conn() if connection is None else nullcontext(connection)) as c:
        if connection is None:
            c.execute('BEGIN IMMEDIATE')
        sheet = c.execute('SELECT * FROM sheets WHERE day=? AND kind=?', (day, kind)).fetchone()
        if sheet and sheet['status'] == 'Onaylandı':
            raise ValueError('Bu tarih ve nöbet türü zaten onaylı. Düzeltmek için önce onayı kaldırın.')
        if sheet:
            old = {r['student_id']: dict(r) for r in c.execute(
                'SELECT * FROM assignments WHERE sheet_id=?', (sheet['id'],))}
        else:
            old = {}
        for row in rows:
            prior = old.get(row.student['id'])
            duty = row.duty
            if old and (not prior or prior['role'] != duty.role or
                        (duty.role == 'Asıl' and slot(prior) != duty.time) or
                        ((duty.role == 'Asıl' or duty.place) and
                         norm(prior['location']).replace(' ', '') != norm(duty.place).replace(' ', ''))):
                raise ValueError('Mevcut taslak ile yüklenen çizelge farklı. Taslağı kontrol edin.')
        if old and set(old) != set(ids):
            raise ValueError('Mevcut taslaktaki öğrenci listesi yüklenen çizelgeyle tam eşleşmeli.')
        if sheet:
            sheet_id = sheet['id']
        else:
            sheet_id = c.execute('INSERT INTO sheets(day,kind) VALUES(?,?)', (day, kind)).lastrowid
        for row in rows:
            duty = row.duty
            if old:
                c.execute('UPDATE assignments SET result=? WHERE sheet_id=? AND student_id=?',
                          (row.result, sheet_id, row.student['id']))
            else:
                actual_day = date.fromisoformat(day)
                if kind != 'Yatakhane Giriş' and duty.time in NIGHT[1:]:
                    actual_day += timedelta(days=1)
                assignment = make_row(row.student, actual_day.isoformat(), duty.time, duty.place, duty.role)
                c.execute("""INSERT INTO assignments
                    (sheet_id,student_id,slot,location,role,result,start_at,end_at,source_cell)
                    VALUES (?,?,?,?,?,?,?,?,?)""", (sheet_id, assignment['student_id'], duty.time,
                    duty.place, duty.role, row.result, assignment['start_at'], assignment['end_at'], duty.name))
        if sheet and old:
            nobet_manuel.require_complete(c, sheet)
            before_confirm(c, sheet)
            ceza_nobet.validate_sheet(app, c, sheet)
        c.execute("UPDATE sheets SET status='Onaylandı' WHERE id=?", (sheet_id,))
        app.log(c, 'Geçmiş nöbet aktarıldı', f'{day} {kind} · {source} · {len(rows)} görev')
    return len(rows)
