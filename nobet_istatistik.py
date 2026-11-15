"""Independent duty-type/time filters; preserve separate and combined totals."""
import datetime as dt

NIGHT=['23:00-01:00','01:00-03:00','03:00-05:00','05:00-07:00']
SLOTS={'Yatakhane Kat':NIGHT,'Devriye':NIGHT,'Yatakhane Giriş':['08:00-10:00','10:00-12:00','12:00-14:00','14:00-16:00','16:00-17:00']}
UNKNOWN='Saat belirtilmemiş'
LABELS={'Yatakhane Kat':'Kat','Devriye':'Devriye','Yatakhane Giriş':'Giriş'}

def slot_key(row):
    text=str(row['slot']).replace('–','-').replace('—','-').replace(' ','')
    if text in SLOTS[row['kind']]:return text
    if row['start_at'] and row['end_at']:
        text=row['start_at'][11:16]+'-'+row['end_at'][11:16]
        if text in SLOTS[row['kind']]:return text
    return UNKNOWN

def calculate(c,start,end,selected,duration,only_matching=False,include_penalties=False):
    if dt.date.fromisoformat(start)>dt.date.fromisoformat(end):raise ValueError('Başlangıç tarihi bitişten sonra olamaz')
    selected=set(selected)
    for kind,slot in selected:
        if kind not in SLOTS or slot not in SLOTS[kind]+[UNKNOWN]:raise ValueError('Geçersiz nöbet/saat filtresi')
    keys=[(kind,slot) for kind,slots in SLOTS.items() for slot in slots+[UNKNOWN] if (kind,slot) in selected]
    kinds=[kind for kind in SLOTS if any(key[0]==kind for key in keys)]
    columns=['Okul No','Ad Soyad','Sınıf']+[f'{LABELS[kind]} {slot} (adet)' for kind,slot in keys]
    for kind in kinds:columns.extend([f'{LABELS[kind]} toplam (adet)',f'{LABELS[kind]} saat'])
    columns+=['Toplam nöbet','Toplam saat']
    students=c.execute('SELECT * FROM students WHERE active=1 ORDER BY class,name,surname').fetchall()
    tasks=c.execute('''SELECT a.student_id,sh.kind,a.slot,a.start_at,a.end_at,a.duty_type
        FROM assignments a JOIN sheets sh ON sh.id=a.sheet_id JOIN students st ON st.id=a.student_id AND st.active=1
        WHERE sh.status='Onaylandı' AND a.result='Tuttu' AND a.role='Asıl'
        AND COALESCE(NULLIF(substr(a.start_at,1,10),''),sh.day) BETWEEN ? AND ?''',(start,end)).fetchall()
    by={};record_count=0;total_hours=0;unknown_count=0
    for row in tasks:
        if row['duty_type']=='Ceza' and not include_penalties:continue
        if row['kind'] not in SLOTS:continue
        key=(row['kind'],slot_key(row))
        if key not in selected:continue
        hours=duration(row['kind'],row['slot'],row['start_at'],row['end_at']) if key[1]!=UNKNOWN else 0
        by.setdefault(row['student_id'],[]).append((key,hours))
        record_count+=1;total_hours+=hours;unknown_count+=key[1]==UNKNOWN
    values=[]
    for student in students:
        tasks=by.get(student['id'],[])
        if only_matching and not tasks:continue
        row=[student['school_no'],student['name']+' '+student['surname'],student['class']]
        row.extend(sum(key==target for key,hours in tasks) for target in keys)
        for kind in kinds:row.extend([sum(key[0]==kind for key,hours in tasks),sum(hours for key,hours in tasks if key[0]==kind)])
        row.extend([len(tasks),sum(hours for key,hours in tasks)])
        values.append(row)
    return columns,values,record_count,total_hours,unknown_count
