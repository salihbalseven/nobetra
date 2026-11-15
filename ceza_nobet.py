"""Penalty obligations, independent of normal queues and exemptions."""
import datetime as dt
import re
import excel_yatakhane as ex
from nobet_yazdir import NIGHT as HOURS

KINDS=('Yatakhane Kat','Devriye')
CHOOSE_LATER='Planlarken seç'
def setup(c):
    c.execute('''CREATE TABLE IF NOT EXISTS historical_penalty_awards(
        source TEXT NOT NULL, source_row TEXT NOT NULL,
        student_id INTEGER NOT NULL REFERENCES students(id),
        award_day TEXT NOT NULL, total INTEGER NOT NULL CHECK(total>0),
        PRIMARY KEY(source,source_row))''')
    c.execute('''CREATE TABLE IF NOT EXISTS historical_penalty_duties(
        assignment_id INTEGER PRIMARY KEY REFERENCES assignments(id) ON DELETE CASCADE,
        source TEXT NOT NULL, source_row TEXT NOT NULL,
        FOREIGN KEY(source,source_row) REFERENCES historical_penalty_awards(source,source_row) ON DELETE CASCADE)''')
    c.execute('''CREATE TABLE IF NOT EXISTS historical_penalty_hours(
        source TEXT NOT NULL, source_row TEXT NOT NULL,
        early_count INTEGER NOT NULL CHECK(early_count>=0),
        late_count INTEGER NOT NULL CHECK(late_count>=0),
        PRIMARY KEY(source,source_row),
        FOREIGN KEY(source,source_row) REFERENCES historical_penalty_awards(source,source_row) ON DELETE CASCADE)''')
    if 'confirmed' not in {r['name'] for r in c.execute('PRAGMA table_info(historical_penalty_hours)')}:
        c.execute('ALTER TABLE historical_penalty_hours ADD COLUMN confirmed INTEGER NOT NULL DEFAULT 0 CHECK(confirmed IN (0,1))')
    c.execute('''CREATE VIEW IF NOT EXISTS historical_confirmed_penalty_hours AS
        WITH remaining AS (
            SELECT p.source,p.source_row,p.student_id,p.award_day,h.early_count,h.late_count,
                MAX(0,p.total-COUNT(a.id)) AS missing
            FROM historical_penalty_awards p JOIN historical_penalty_hours h USING(source,source_row)
            LEFT JOIN historical_penalty_duties d USING(source,source_row)
            LEFT JOIN assignments a ON a.id=d.assignment_id AND a.duty_type='Ceza'
                AND a.role='Asıl' AND a.result='Tuttu'
                AND EXISTS(SELECT 1 FROM sheets sh WHERE sh.id=a.sheet_id AND sh.status='Onaylandı')
            WHERE h.confirmed=1 GROUP BY p.source,p.source_row
        )
        SELECT source,source_row,student_id,award_day,'01.00–03.00' AS hours,MIN(early_count,missing) AS count
        FROM remaining WHERE early_count>0 AND missing>0
        UNION ALL
        SELECT source,source_row,student_id,award_day,'03.00–05.00',MIN(late_count,MAX(0,missing-early_count))
        FROM remaining WHERE late_count>0 AND missing>early_count''')
    c.execute('''CREATE TABLE IF NOT EXISTS penalties(id INTEGER PRIMARY KEY,student_id INTEGER NOT NULL REFERENCES students(id),total INTEGER NOT NULL CHECK(total>0),target TEXT NOT NULL,start_day TEXT NOT NULL,active INTEGER NOT NULL DEFAULT 1,UNIQUE(student_id,start_day))''')
    columns={r['name'] for r in c.execute('PRAGMA table_info(assignments)')}
    if 'duty_type' not in columns:c.execute("ALTER TABLE assignments ADD COLUMN duty_type TEXT NOT NULL DEFAULT 'Normal'")
    if 'penalty_id' not in columns:c.execute('ALTER TABLE assignments ADD COLUMN penalty_id INTEGER')
    c.execute('CREATE INDEX IF NOT EXISTS idx_assignments_penalty ON assignments(penalty_id)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_assignments_student ON assignments(student_id)')
    columns={r['name'] for r in c.execute('PRAGMA table_info(penalties)')}
    if 'archived' not in columns:c.execute('ALTER TABLE penalties ADD COLUMN archived INTEGER NOT NULL DEFAULT 0')
    if 'voided' not in columns:c.execute('ALTER TABLE penalties ADD COLUMN voided INTEGER NOT NULL DEFAULT 0')
    if 'preferred_slot' not in columns:
        c.execute('ALTER TABLE penalties ADD COLUMN preferred_slot TEXT')

def pasted_hour(value, kind):
    value=re.sub(r'\s+', '', value).replace('–','-').replace('—','-').replace('.',':')
    if not value:
        return None
    match=re.fullmatch(r'(\d{1,2})(?::00)?-(\d{1,2})(?::00)?',value)
    if match:
        start,end=map(int,match.groups())
        slot=f'{23 if start==11 else start:02}:00-{end:02}:00'
        if slot in HOURS:
            return slot
    raise ValueError(f'Geçersiz ceza nöbeti saati: {value}. Seçenekler: {", ".join(HOURS)}')

def managed(c,sid,day=None):return bool(c.execute('SELECT 1 FROM penalties WHERE student_id=? AND start_day<=? AND voided=0',(sid,day or '9999-12-31')).fetchone())

def rest_reason(c,sid,day,exclude_assignment=-1):
    if not managed(c,sid,day):return ''
    rows=list(c.execute("SELECT s.day,a.role FROM assignments a JOIN sheets s ON s.id=a.sheet_id WHERE a.student_id=? AND a.result!='Tutmadı' AND a.id!=?",(sid,exclude_assignment)))
    if any(r['day']==day for r in rows):return 'Ceza kaydı olan öğrenciye aynı gün yalnızca bir görev yazılabilir'
    dates={dt.date.fromisoformat(r['day']) for r in rows if r['role']=='Asıl'};candidate=dt.date.fromisoformat(day);dates.add(candidate)
    for first in dates:
        if first+dt.timedelta(days=1) in dates:
            for offset in (2,3):
                blocked=first+dt.timedelta(days=offset)
                if blocked in dates and (candidate in (first,first+dt.timedelta(days=1),blocked)):
                    return 'İki gün üst üste nöbetten sonra iki tam gün dinlenme gerekli'
    return ''

def counts(c,pid):
    done=c.execute("SELECT COUNT(*) FROM assignments a JOIN sheets s ON s.id=a.sheet_id WHERE a.penalty_id=? AND a.duty_type='Ceza' AND a.role='Asıl' AND a.result='Tuttu' AND s.status='Onaylandı'",(pid,)).fetchone()[0]
    reserved=c.execute("SELECT COUNT(*) FROM assignments WHERE penalty_id=? AND duty_type='Ceza' AND result!='Tutmadı'",(pid,)).fetchone()[0]
    return done,reserved

def record(c,sid,total,target,start,preferred_slot=None):
    dt.date.fromisoformat(start)
    if target not in KINDS:raise ValueError('Ceza yeri yatakhane kat veya devriye olmalı')
    if int(total)!=total or total<1:raise ValueError('Ceza nöbeti adedi pozitif tam sayı olmalı')
    if preferred_slot not in (None, '', *HOURS):
        raise ValueError('Geçersiz ceza nöbeti saati')
    old=c.execute('SELECT * FROM penalties WHERE student_id=? AND start_day=?',(sid,start)).fetchone()
    if old and old['archived'] and not old['voided']:raise ValueError('Aynı tarihli ceza arşivde. Önce Arşiv görünümünden geri alın veya yeni ceza için yeni kayıt tarihi kullanın.')
    if old and counts(c,old['id'])[1]>total:raise ValueError('Ceza adedi, tutulmuş ve planlanmış görevlerin toplamından az olamaz')
    if not old or total>counts(c,old['id'])[0]:
        other=c.execute('SELECT * FROM penalties WHERE student_id=? AND id!=? AND active=1 AND archived=0 AND voided=0 ORDER BY start_day',(sid,old['id'] if old else -1)).fetchall()
        if any(counts(c,p['id'])[0]<p['total'] for p in other):raise ValueError('Öğrencinin bitmemiş başka ceza kaydı var; mevcut kaydı düzenleyin')
    if preferred_slot is None and old:
        preferred_slot=old['preferred_slot']
    c.execute('''INSERT INTO penalties(student_id,total,target,start_day,preferred_slot) VALUES(?,?,?,?,?)
        ON CONFLICT(student_id,start_day) DO UPDATE SET total=excluded.total,target=excluded.target,
        preferred_slot=excluded.preferred_slot,active=1,voided=0,archived=0''',
        (sid,int(total),target,start,preferred_slot or None))

def import_excel(app,path,target,start):
    book=ex.read_book(path);title='CEZALILAR LİSTESİ'
    if title not in book:raise ValueError(f'{title} sayfası yok')
    cells=book[title];items=[]
    for row in range(2,103):
        value=lambda col:str(cells.get(f'{col}{row}',{}).get('value','')).strip()
        a,b=value('A'),value('B')
        if not a and not b:continue
        if not a or not b:raise ValueError(f'{title}!A{row}:B{row}: ad/soyad eksik')
        total=value('J')
        if not re.fullmatch(r'\d+(?:\.0+)?',total) or float(total)<1:raise ValueError(f'{title}!J{row}: ceza adedi eksik veya geçersiz')
        items.append((a,b,int(float(total)),row))
    if not items:raise ValueError('Cezalılar listesinde isim yok')
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE');names={}
        for st in c.execute('SELECT * FROM students WHERE active=1'):names.setdefault((ex.norm(st['name']),ex.norm(st['surname'])),[]).append(st)
        used=set()
        for a,b,total,row in items:
            matches=names.get((ex.norm(a),ex.norm(b)),[])
            if len(matches)!=1:raise ValueError(f'{title}!A{row}:B{row}: {a} {b} bulunamadı veya birden fazla isim eşleşmesi var')
            sid=matches[0]['id']
            if sid in used:raise ValueError(f'Satır {row}: aynı öğrenci iki kez')
            used.add(sid);record(c,sid,total,target,start)
        app.log(c,'Cezalılar Excel aktarımı',f'{len(items)} kişi; {target}; ceza kayıt tarihi {start}')
    return len(items)

def import_paste(app,text,total,target,start,fallback=None):
    from nobet_manuel import parse_paste

    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        students={row['id']:row for row in c.execute('SELECT * FROM students WHERE active=1')}
        items=parse_paste(text,students,target,(),time_parser=pasted_hour)
        for sid,hour in items:
            try:
                record(c,sid,total,target,start,hour or fallback)
            except ValueError as error:
                student=students[sid]
                raise ValueError(f"{student['name']} {student['surname']}: {error}") from error
        app.log(c,'Cezalılar Excel yapıştırma',f'{len(items)} kişi; {target}; ceza kayıt tarihi {start}')
    return len(items)

def candidates(app,c,day,kind,blocked):
    rows=[];busy=app.busy_students(c,day,kind)
    for row in c.execute('SELECT p.*,st.name,st.surname,st.school_no,st.class,st.active AS student_active FROM penalties p JOIN students st ON st.id=p.student_id WHERE p.active=1 AND p.archived=0 AND p.voided=0 AND p.target=? AND p.start_day<=? ORDER BY p.start_day,p.id',(kind,day)):
        if not row['student_active'] or counts(c,row['id'])[1]>=row['total']:continue
        if ex.norm(row['class']) in blocked or rest_reason(c,row['student_id'],day):continue
        if row['student_id'] in busy:continue
        student=dict(c.execute('SELECT * FROM students WHERE id=?',(row['student_id'],)).fetchone())
        student['preferred_penalty_slot']=row['preferred_slot']
        rows.append((student,row['id']))
    return rows

def validate_sheet(app,c,sheet):
    for row in c.execute('SELECT * FROM assignments WHERE sheet_id=? AND result!=?',(sheet['id'],'Tutmadı')).fetchall():
        if not managed(c,row['student_id'],sheet['day']):continue
        reason=rest_reason(c,row['student_id'],sheet['day'],row['id'])
        if reason:raise ValueError(f'Kayıt {row["id"]}: {reason}')
        if row['duty_type']=='Ceza':
            if c.execute('SELECT 1 FROM historical_penalty_duties WHERE assignment_id=?',(row['id'],)).fetchone():
                continue
            case=c.execute('SELECT * FROM penalties WHERE id=?',(row['penalty_id'],)).fetchone()
            if not case or case['student_id']!=row['student_id']:raise ValueError('Ceza görevinin kişi/kayıt eşleşmesi geçersiz')
            if counts(c,case['id'])[1]>case['total']:raise ValueError('Ceza adedinden fazla görev planlanmış')

def clean(app,ids,mode):
    ids=list(dict.fromkeys(int(pid) for pid in ids))
    if not ids:raise ValueError('Önce kayıt seçin')
    if mode not in ('archive','restore','void'):raise ValueError('Geçersiz işlem')
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        rows=[];restored_students=set()
        for pid in ids:
            p=c.execute('SELECT * FROM penalties WHERE id=? AND voided=0',(pid,)).fetchone()
            if not p:raise ValueError('Seçilen bir kayıt artık yok; listeyi yenileyin')
            if mode=='void' and c.execute('SELECT 1 FROM assignments WHERE penalty_id=?',(pid,)).fetchone():
                raise ValueError('Görevlendirmesi bulunan ceza kaydı silinemez. Geçmişi korumak için arşivleyin.')
            if mode=='restore':
                if counts(c,pid)[0]<p['total']:
                    if p['student_id'] in restored_students:raise ValueError('Aynı öğrencinin iki bitmemiş cezası birlikte geri alınamaz')
                    restored_students.add(p['student_id'])
                    for other in c.execute('SELECT * FROM penalties WHERE student_id=? AND id!=? AND active=1 AND archived=0 AND voided=0',(p['student_id'],pid)):
                        if counts(c,other['id'])[0]<other['total']:raise ValueError('Öğrencinin güncel bitmemiş başka cezası var; önce mevcut kaydı kontrol edin.')
            rows.append(p)
        for p in rows:
            if mode=='void':c.execute('UPDATE penalties SET voided=1,active=0 WHERE id=?',(p['id'],))
            else:c.execute('UPDATE penalties SET archived=? WHERE id=?',(int(mode=='archive'),p['id']))
        app.log(c,'Ceza listesi '+mode,','.join(map(str,ids)))
    return len(ids)

def statistics(c,start,end,query='',klass='Tümü'):
    if dt.date.fromisoformat(start)>dt.date.fromisoformat(end):raise ValueError('Başlangıç tarihi bitişten sonra olamaz')
    grouped={};needle=ex.norm(query)
    for p in c.execute('SELECT p.*,s.school_no,s.name,s.surname,s.class FROM penalties p JOIN students s ON s.id=p.student_id WHERE p.voided=0 AND p.start_day BETWEEN ? AND ? ORDER BY s.class,s.name,s.surname,p.start_day',(start,end)).fetchall():
        if needle and needle not in ex.norm(f'{p["school_no"]} {p["name"]} {p["surname"]}'):continue
        if klass!='Tümü' and ex.norm(klass)!=ex.norm(p['class']):continue
        row=grouped.setdefault(p['student_id'],[p['school_no'],p['name']+' '+p['surname'],p['class'],0,0,0,0,0,0,'',0])
        done,reserved=counts(c,p['id']);remaining=max(0,p['total']-done)
        row[3]+=1;row[4]+=p['total'];row[5]+=done;row[6]+=max(0,reserved-done)
        row[8 if p['archived'] else 7]+=remaining;row[9]=max(row[9],p['start_day'])
    for p in c.execute('''SELECT p.*,s.school_no,s.name,s.surname,s.class,
        (SELECT COUNT(*) FROM historical_penalty_duties h JOIN assignments a ON a.id=h.assignment_id
         JOIN sheets sh ON sh.id=a.sheet_id WHERE h.source=p.source AND h.source_row=p.source_row
        AND a.duty_type='Ceza' AND a.role='Asıl' AND a.result='Tuttu' AND sh.status='Onaylandı')
        + COALESCE((SELECT SUM(count) FROM historical_confirmed_penalty_hours h
                    WHERE h.source=p.source AND h.source_row=p.source_row),0) AS confirmed
        FROM historical_penalty_awards p JOIN students s ON s.id=p.student_id
        WHERE p.award_day BETWEEN ? AND ?''',(start,end)):
        if needle and needle not in ex.norm(f'{p["school_no"]} {p["name"]} {p["surname"]}'):continue
        if klass!='Tümü' and ex.norm(klass)!=ex.norm(p['class']):continue
        row=grouped.setdefault(p['student_id'],[p['school_no'],p['name']+' '+p['surname'],p['class'],0,0,0,0,0,0,'',0])
        row[3]+=1;row[4]+=p['total'];row[5]+=p['confirmed'];row[9]=max(row[9],p['award_day']);row[10]+=p['total']
    return sorted(grouped.values(),key=lambda row:(row[2] or '',ex.norm(row[1]),row[0]))
