"""Import actual patrol chart positions and repair legacy generic duty labels."""
import datetime as dt
import re
import excel_yatakhane as ex

PLACES = ('HABER M.', '1 D', '1 D', '2 D', '2 D', '3 D', '3 D', '4 D', '4 D', '5 D', '5 D')
RESERVES = 6
TOTAL = 4 * len(PLACES) + RESERVES

def build_plan(path,day):
    book=ex.read_book(path)
    entry_title=next((title for title in ('50 KİŞİ GİRİŞ','42 KİŞİ GİRİŞ') if title in book),None)
    if not entry_title or 'DEVRİYE NÖBET' not in book:
        raise ValueError('DEVRİYE NÖBET ve 50 KİŞİ GİRİŞ (eski dosyada 42 KİŞİ GİRİŞ) sayfaları gerekli.')
    places=PLACES if entry_title.startswith('50') else PLACES[:-2]
    total=4*len(places)+RESERVES
    entry=book[entry_title];chart=book['DEVRİYE NÖBET'];base=dt.date.fromisoformat(day)
    def anchor(cell):return chart.get(cell,{}).get('merged_from',cell)
    def value(cells,cell):return str(cells.get(cell,{}).get('value','')).strip()
    result=[];seen=set();errors=[]
    for row in range(3,3+total):
        formula=chart.get(f'D{row}',{}).get('formula') or ''
        references=set(re.findall(rf"'{re.escape(entry_title)}'!\$?A\$?(\d+)\b",formula))
        if len(references)!=1:raise ValueError(f'DEVRİYE NÖBET!D{row}: giriş sayfasına tek bir ad bağlantısı bekleniyor')
        source=int(references.pop())
        if source not in range(2,2+total) or source in seen:raise ValueError(f'DEVRİYE NÖBET!D{row}: yinelenen/geçersiz giriş satırı')
        seen.add(source)
        first=value(entry,f'A{source}');last=value(entry,f'B{source}')
        label=f'{entry_title}!A{source}:E{source} → DEVRİYE NÖBET!D{row}'
        if not first or not last:errors.append(f'{label}: ad/soyad eksik');continue
        number=value(entry,f'E{source}')
        if re.fullmatch(r'\d+\.0+',number):number=number.split('.')[0]
        place=value(chart,anchor(f'C{row}'))
        if not place:raise ValueError(f'DEVRİYE NÖBET!C{row}: görev yeri boş')
        reserve=row>=3+4*len(places)
        if reserve:
            if 'yedek' not in ex.norm(place):raise ValueError(f'DEVRİYE NÖBET!C{row}: yedek bilgisi bekleniyor')
            slot='Yedek';start=end=''
        else:
            slot=re.sub(r'\s+','',value(chart,anchor(f'B{row}'))).replace('–','-').replace('—','-')
            if not re.fullmatch(r'\d{2}:\d{2}-\d{2}:\d{2}',slot):raise ValueError(f'DEVRİYE NÖBET!B{row}: geçersiz saat {slot!r}')
            try:duty_day=ex.resolve_date(chart,anchor(f'A{row}'),base.year)
            except ValueError as err:raise ValueError(str(err).replace('Yatakhane Nöbet','DEVRİYE NÖBET'))
            expected=base if row<3+len(places) else base+dt.timedelta(days=1)
            if duty_day!=expected:raise ValueError(f'DEVRİYE NÖBET!{anchor(f"A{row}")}: {duty_day} okunuyor; seçili gece {day}. Excel ve program tarihlerini kontrol edin.')
            a,b=slot.split('-');start=dt.datetime.combine(duty_day,dt.time.fromisoformat(a));end=dt.datetime.combine(duty_day,dt.time.fromisoformat(b))
            if end<=start:end+=dt.timedelta(days=1)
            start=start.isoformat(timespec='minutes');end=end.isoformat(timespec='minutes')
        result.append(dict(first=first,last=last,school_no=number,klass=value(entry,f'C{source}'),slot=slot,location=place,role='Yedek' if reserve else 'Asıl',start=start,end=end,source=label))
    if errors:raise ValueError('Hiçbir kayıt aktarılmadı.\n'+'\n'.join(errors))
    return result

def save_plan(app,items,day):
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        c.execute('INSERT OR IGNORE INTO sheets(day,kind) VALUES(?,?)',(day,'Devriye'))
        sheet=c.execute('SELECT * FROM sheets WHERE day=? AND kind=?',(day,'Devriye')).fetchone()
        students=list(c.execute('SELECT * FROM students WHERE active=1'))
        numbers={s['school_no']:s for s in students};names={}
        for st in students:names.setdefault((ex.norm(st['name']),ex.norm(st['surname'])),[]).append(st)
        rows=[];used=set();errors=[]
        for item in items:
            matches=[numbers[item['school_no']]] if item['school_no'] in numbers else ([] if item['school_no'] else names.get((ex.norm(item['first']),ex.norm(item['last'])),[]))
            if len(matches)!=1:errors.append(item['source']+': öğrenci bulunamadı veya isim birden çok kişiye ait');continue
            st=matches[0]
            if ex.norm(st['surname'])!=ex.norm(item['last']) or (item['klass'] and ex.norm(st['class'])!=ex.norm(item['klass'])):
                errors.append(item['source']+': okul numarası ile soyad/sınıf uyuşmuyor');continue
            if st['id'] in used:errors.append(item['source']+': aynı öğrenci iki kez yazılmış');continue
            used.add(st['id']);rows.append((st,item))
        old={r['student_id']:r for r in c.execute('SELECT * FROM assignments WHERE sheet_id=?',(sheet['id'],))}
        if old and set(old)!=used:errors.append('Mevcut liste ile Excel’in öğrencileri farklı. Bu işlem sadece aynı listenin saat/yer bilgilerini düzeltir; farklı liste için taslağı kontrol edin.')
        if sheet['status']=='Onaylandı' and not old:errors.append('Onaylı boş çizelgeye kayıt eklenemez; önce onayı kaldırın.')
        if len(items) not in (42,TOTAL):errors.append('50 kişilik veya eski 42 kişilik tam liste bekleniyor.')
        if not old:
            for st,item in rows:
                duty_day=item['start'][:10] or day
                reason=app.eligible(c,st,duty_day,'Devriye',item['slot'])
                if reason:errors.append(item['source']+': '+reason)
                if app.has_conflict(c,st['id'],day,'Devriye',item['start'],item['end']):
                    errors.append(item['source']+': aynı gece devriye/yatakhane kat grubunda başka görev var')
        if errors:raise ValueError('Hiçbir kayıt değiştirilmedi.\n'+'\n'.join(errors))
        changed=0
        for st,item in rows:
            values=tuple(item[k] for k in ('slot','location','role','start','end','source'))
            if old:
                previous=old[st['id']]
                if tuple(previous[k] for k in ('slot','location','role','start_at','end_at','source_cell'))==values:continue
                c.execute('UPDATE assignments SET slot=?,location=?,role=?,start_at=?,end_at=?,source_cell=? WHERE id=?',values+(previous['id'],))
            else:c.execute('INSERT INTO assignments(slot,location,role,start_at,end_at,source_cell,sheet_id,student_id) VALUES(?,?,?,?,?,?,?,?)',values+(sheet['id'],st['id']))
            changed+=1
        app.log(c,'Devriye Excel saat/yer aktarımı',f'{day}; {changed} kayıt; '+('mevcut sonuç ve onaylar korundu' if old else 'yeni taslak'))
        return changed,bool(old)
