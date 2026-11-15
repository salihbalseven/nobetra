"""Bulk student workbook import keyed by school number."""
import re
import excel_yatakhane
from ogrenci_profili import FIELDS, norm

BASE = [('school_no','Okul No'), ('name','Adı'), ('surname','Soyadı'), ('class','Sınıf'), ('room','Koğuş')]
HEADERS = BASE + FIELDS

def read(path, classes):
    book = excel_yatakhane.read_book(path)
    cells = book.get('Öğrenciler', next(iter(book.values()), {}))
    aliases = {norm(label):key for key,label in HEADERS}
    aliases.update({'ad':'name','soyad':'surname','okul numarasi':'school_no','sinif':'class'})
    mapping = {}
    for address, cell in cells.items():
        match = re.fullmatch(r'([A-Z]+)1', address)
        if match and cell.get('value','').strip():
            key = aliases.get(norm(cell['value']).strip())
            if not key:
                raise ValueError(f'{address}: tanınmayan başlık: {cell["value"]}. Şablondaki başlıkları kullanın.')
            if key in mapping.values():
                raise ValueError(f'{address}: aynı başlık birden fazla yazılmış.')
            mapping[match[1]] = key
    required = {'school_no','name','surname','class'}
    if not required <= set(mapping.values()):
        raise ValueError('İlk satırda Okul No, Adı, Soyadı ve Sınıf başlıkları olmalı.')
    rows = {}
    for address, cell in cells.items():
        match = re.fullmatch(r'([A-Z]+)(\d+)',address)
        if not match:
            continue
        column, row = match.groups()
        if int(row) <= 1 or column not in mapping:
            continue
        if cell.get('formula') is not None:
            raise ValueError(f'{address}: formül yerine değeri yapıştırın.')
        value = cell.get('value','').strip()
        if value:
            rows.setdefault(int(row),{})[mapping[column]] = value
    seen = set()
    result = []
    for row, values in sorted(rows.items()):
        if not required <= values.keys():
            raise ValueError(f'Satır {row}: Okul No, Adı, Soyadı ve Sınıf gerekli.')
        school = values['school_no']
        if school in seen:
            raise ValueError(f'Satır {row}: aynı okul numarası tekrar edilmiş.')
        seen.add(school)
        student_class = next((c for c in classes if norm(c)==norm(values['class'])), None)
        if student_class is None:
            raise ValueError(f'Satır {row}: geçersiz sınıf {values["class"]}.')
        values['class'] = student_class
        result.append(values)
    if not result:
        raise ValueError('Aktarılacak öğrenci satırı bulunamadı.')
    return result

def apply(c, rows):
    added = updated = 0
    for values in rows:
        school = values['school_no']
        existing = c.execute('SELECT id FROM students WHERE school_no=?',(school,)).fetchone()
        fields = [k for k,_ in BASE if k != 'school_no' and k in values]
        if existing:
            sid = existing[0]
            c.execute('UPDATE students SET '+','.join(k+'=?' for k in fields)+' WHERE id=?',
                      [values[k] for k in fields]+[sid])
            updated += 1
        else:
            keys = ['school_no']+fields
            sid = c.execute('INSERT INTO students('+','.join(keys)+') VALUES('+','.join('?' for k in keys)+')',
                            [values[k] for k in keys]).lastrowid
            added += 1
        details = [k for k,_ in FIELDS if k in values]
        if details:
            c.execute('INSERT OR IGNORE INTO student_profiles(student_id) VALUES(?)',(sid,))
            c.execute('UPDATE student_profiles SET '+','.join(k+'=?' for k in details)+' WHERE student_id=?',
                      [values[k] for k in details]+[sid])
    return added,updated
