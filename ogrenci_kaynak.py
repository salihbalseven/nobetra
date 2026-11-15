"""Fill empty profile fields from the numbered class and dormitory source lists."""

from collections import defaultdict, Counter
import re

from excel_yatakhane import read_book
from ogrenci_profili import norm


def records(path, names):
    book = read_book(path)
    for sheet in names:
        if sheet not in book:
            continue
        rows = defaultdict(dict)
        for address, cell in book[sheet].items():
            value = cell.get('value', '').strip()
            if value:
                column, number = re.fullmatch(r'([A-Z]+)(\d+)', address).groups()
                rows[int(number)][column] = value
        headers = {col: norm(label).strip() for col, label in rows.pop(1, {}).items()}
        for _, row in sorted(rows.items()):
            yield {headers[col]: value for col, value in row.items() if col in headers}


def clean(value):
    value = (value or '').strip()
    return '' if value in ('0', '-', '—') else value


def full_name(given, surname):
    return ' '.join(norm(f'{given or ""} {surname or ""}').split())

def contact_values(row):
    aliases={
        'phone':('ogrenci telefonu','tel. no','tel no','tel no1'),
        'mother_phone':('anne telefonu','anne tel no','anne cep telefonu'),
        'father_phone':('baba telefonu','baba tel no','baba cep telefonu'),
        'address':('ikamet adresi','ikametgah adresi','adres'),
    }
    values={}
    for key,headers in aliases.items():
        for header in headers:
            value=clean(row.get(header))
            if not value:
                continue
            if key!='address':
                digits=re.sub(r'\D','',re.sub(r'\.0+$','',value))
                if len(digits)==10 and digits.startswith('5'):
                    value='0'+digits
                elif len(digits)==11 and digits.startswith('0'):
                    value=digits
                elif len(digits)==12 and digits.startswith('90'):
                    value='+'+digits
                else:
                    continue
            values[key]=value
            break
    return values


def plan(c, class_path, dorm_path):
    students = {row['id']: dict(row) for row in c.execute('SELECT * FROM students')}
    by_school = {str(row['school_no']): sid for sid, row in students.items()}
    by_name = defaultdict(list)
    for sid, row in students.items():
        by_name[full_name(row['name'], row['surname'])].append(sid)
    profiles = {row['student_id']: dict(row) for row in c.execute('SELECT * FROM student_profiles')}
    changes = defaultdict(dict)
    coverage = Counter()
    for sid, row in students.items():
        if row['school_no']:
            changes[sid]['registry'] = row['school_no']
    for row in c.execute("""SELECT student_id,reason FROM exemptions
        WHERE start='1900-01-01' AND (end IS NULL OR end='') ORDER BY id"""):
        changes[row['student_id']].setdefault('official_duty', row['reason'])

    def match(row, source):
        name = full_name(row.get('adi'), row.get('soyadi'))
        if not name:
            return None
        number = re.sub(r'\.0+$', '', row.get('okul no', '').strip())
        if number and number != '0':
            sid = by_school.get(number)
            candidates = by_name[name]
            if sid not in candidates:
                coverage[f'{source}_atlandi'] += 1
                return None
            return sid
        candidates = by_name[name]
        if len(candidates) == 1:
            return candidates[0]
        coverage[f'{source}_atlandi'] += 1
        return None

    if class_path:
        rows = list(records(class_path, ('KAYNAK LİSTE', 'KAYNAK YEDEK 1', 'KAYNAK YEDEK 2')))
        if not rows:
            raise ValueError('Sınıf dosyasında KAYNAK LİSTE sayfası bulunamadı.')
        for row in rows:
            sid = match(row, 'sinif')
            if sid is None:
                continue
            values = {
                'mother': clean(row.get('anne adi')), 'father': clean(row.get('baba adi')),
                'university': clean(row.get('universitesi')), 'department': clean(row.get('bolum')),
                'hometown': ' / '.join(filter(None, (clean(row.get('nufus il')), clean(row.get('nufus ilce'))))),
                'official_duty': clean(row.get('gorevi')),
            }
            values.update(contact_values(row))
            for key, value in values.items():
                if value:
                    changes[sid].setdefault(key, value)
    if dorm_path:
        rows = list(records(dorm_path, ('YATAKHANE', 'Yatak Listesi')))
        if not rows:
            raise ValueError('Yatakhane listesi sayfası bulunamadı.')
        for row in rows:
            sid = match(row, 'yatak')
            if sid is None:
                continue
            current_room = clean(students[sid]['room'])
            room = clean(row.get('kogus no'))
            bed = clean(row.get('ranza no'))
            if room and current_room and norm(room) != norm(current_room):
                coverage['kogus_farkli'] += 1
                continue
            if room and not current_room:
                changes[sid].setdefault('room', room)
            if bed:
                changes[sid].setdefault('bed', bed)
    pending = {}
    for sid, values in changes.items():
        profile = profiles.get(sid, {})
        fields = {key: value for key, value in values.items()
                  if value and not clean(students[sid]['room'] if key == 'room' else profile.get(key))}
        if fields:
            pending[sid] = fields
    return pending, coverage


def apply(c, changes):
    for sid, values in changes.items():
        if 'room' in values:
            c.execute('UPDATE students SET room=? WHERE id=?', (values['room'], sid))
        profile = {key: value for key, value in values.items() if key != 'room'}
        if profile:
            c.execute('INSERT OR IGNORE INTO student_profiles(student_id) VALUES(?)', (sid,))
            c.execute('UPDATE student_profiles SET ' + ','.join(f'{key}=?' for key in profile) +
                      ' WHERE student_id=?', (*profile.values(), sid))
