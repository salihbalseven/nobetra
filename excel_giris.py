"""Read gate-duty dates from the workbook, independently of the screen date."""
import re
import datetime as dt
import excel_yatakhane as ex

def build_batches(path):
    book=ex.read_book(path);title='28 KİŞİ GİRİŞ'
    if title not in book:raise ValueError(f'{title} sayfası bulunamadı')
    cells=book[title];batches={};master=book.get('KİŞİ LİSTESİ GÜNCEL',{});lookup={};placements={}
    chart=book.get('GİRİŞ NÖBET ÇİZELGESİ')
    if chart is not None:
        for top in (5,15):
            for offset in range(7):
                for column,block in (('C','A BLOK'),('H','B BLOK')):
                    address=f'{column}{top+offset}';formula=chart.get(address,{}).get('formula') or ''
                    matches=set(re.findall(r"'28 KİŞİ GİRİŞ'!\$?E\$?(\d+)\b",formula))
                    if len(matches)!=1:raise ValueError(f'GİRİŞ NÖBET ÇİZELGESİ!{address}: giriş satırı bağlantısı okunamadı')
                    source=int(matches.pop())
                    if source in placements:raise ValueError(f'Giriş satırı {source} çizelgede iki yere bağlanmış')
                    time=str(chart.get(f'B{top+offset}',{}).get('value','')).strip()
                    placements[source]=(time,block)
    for row in range(2,428):
        def master_value(column):return str(master.get(f'{column}{row}',{}).get('value','')).strip()
        first,last=master_value('E'),master_value('F')
        if first and last:lookup.setdefault((ex.norm(first),ex.norm(last)),[]).append((master_value('C'),master_value('B')))
    for row in range(2,30):
        def value(column):return str(cells.get(f'{column}{row}',{}).get('value','')).strip()
        first,last=value('E'),value('F')
        if not first and not last:continue
        if not first or not last:raise ValueError(f'{title}!E{row}:F{row}: ad/soyad eksik')
        try:
            year=dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).year
            day=ex.resolve_date(cells,f'B{row}',year).isoformat()
        except ValueError as error:raise ValueError(str(error).replace('Yatakhane Nöbet',title))
        raw_slot,location=placements.get(row,(value('D'),value('C')))
        slot=raw_slot.replace('–','-').replace('—','-').replace(' ','')
        reserve='yedek' in ex.norm(slot)
        if reserve:slot='Yedek'
        if slot not in ('08:00-10:00','10:00-12:00','12:00-14:00','14:00-16:00','16:00-17:00','Yedek'):
            raise ValueError(f'{title}!D{row}: geçersiz saat {slot!r}')
        number=value('I')
        if re.fullmatch(r'\d+\.0+',number):number=number.split('.')[0]
        klass=value('G')
        # The template's INDEX/MATCH caches can belong to an earlier ordering.
        # Recompute only these known lookups from its raw master table, leaving typed values intact.
        matches=lookup.get((ex.norm(first),ex.norm(last)),[])
        if len(matches)==1:
            formula=cells.get(f'I{row}',{}).get('formula') or ''
            if "INDEX('KİŞİ LİSTESİ GÜNCEL'!$C$2:$C$427" in formula:number=matches[0][0]
            formula=cells.get(f'G{row}',{}).get('formula') or ''
            if "INDEX('KİŞİ LİSTESİ GÜNCEL'!$B$2:$B$427" in formula:klass=matches[0][1]
        identity={'school_no':number,'class':klass}
        batches.setdefault(day,[]).append((first,last,slot,location,'Yedek' if reserve else 'Asıl',row,identity))
    if not batches:raise ValueError('28 KİŞİ GİRİŞ sayfasında ad/soyad bulunamadı')
    for day,items in batches.items():
        if len(items)>14:raise ValueError(f'{day}: {len(items)} kişi bulundu; aynı tarih için en fazla 14 kişi olabilir')
    return batches

def save_batches(app,importer,batches,repair=False):
    """All dates commit together; identical preexisting assignments are skipped."""
    if not batches:raise ValueError('Aktarılacak tarih seçilmedi')
    summary={}
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        students=list(c.execute('SELECT * FROM students WHERE active=1'))
        numbers={str(st['school_no']).strip():st for st in students if st['school_no']}
        names={}
        for st in students:names.setdefault((ex.norm(st['name']),ex.norm(st['surname'])),[]).append(st)
        for day,items in sorted(batches.items()):
            dt.date.fromisoformat(day)
            sh=c.execute('SELECT * FROM sheets WHERE day=? AND kind=?',(day,'Yatakhane Giriş')).fetchone()
            old={r['student_id']:r for r in c.execute('SELECT * FROM assignments WHERE sheet_id=?',(sh['id'],))} if sh else {}
            pending=[];used=set();skipped=0;updates=[]
            for item in items:
                first,last,slot,location,role,row=item[:6]
                identity=item[6] if len(item)>6 else {};number=identity.get('school_no','')
                matches=([numbers[number]] if number in numbers else []) if number else names.get((ex.norm(first),ex.norm(last)),[])
                if len(matches)!=1:
                    pending.append(item);continue # The shared importer reports unknown/ambiguous identities.
                st=matches[0]
                if number and (ex.norm(st['surname'])!=ex.norm(last) or (identity.get('class') and ex.norm(st['class'])!=ex.norm(identity['class']))):
                    raise ValueError(f'{day}, satır {row}: okul numarası ile soyad/sınıf uyuşmuyor. Hiçbir tarih aktarılmadı.')
                if st['id'] in used:raise ValueError(f'{day}, satır {row}: aynı öğrenci iki kez yazılmış. Hiçbir tarih aktarılmadı.')
                used.add(st['id'])
                existing=old.get(st['id'])
                if existing:
                    if (existing['slot'],existing['location'],existing['role'])!=(slot,location,role):
                        if not repair:raise ValueError(f'{day}, satır {row}: mevcut saati/bloku/rolü farklı. Önizlemedeki Mevcut saat/blokları düzelt düğmesini kullanın. Hiçbir tarih aktarılmadı.')
                        reason=app.eligible(c,st,day,'Yatakhane Giriş',slot)
                        if reason:raise ValueError(f'{day}, satır {row}: {reason}. Hiçbir tarih değiştirilmedi.')
                        if role=='Asıl':
                            first_time,last_time=slot.split('-');start=day+'T'+first_time;end=day+'T'+last_time
                        else:start=end=''
                        updates.append((slot,location,role,start,end,f'28 KİŞİ GİRİŞ!E{row}:I{row}',existing['id']))
                    else:skipped+=1
                else:pending.append(item)
            if repair and old and set(old)!=used:raise ValueError(f'{day}: mevcut listenin öğrencileri Excel’den farklı. Bu düzeltme yalnızca aynı öğrenci listesinin saat/blok/rollerini günceller.')
            try:added=importer(pending,'Yatakhane Giriş',day,connection=c) if pending else 0
            except ValueError as error:raise ValueError(f'{day}: {error}\nSeçilen tarihlerin hiçbirinde yeni kayıt yapılmadı.') from error
            if updates:
                c.executemany('UPDATE assignments SET slot=?,location=?,role=?,start_at=?,end_at=?,source_cell=? WHERE id=?',updates)
                app.log(c,'Giriş nöbeti saat/blok düzeltmesi',f'{day}: {len(updates)} kayıt; öğrenci, sonuç ve onay durumları korundu')
            summary[day]={'added':added,'skipped':skipped}
            if repair:summary[day]['updated']=len(updates)
        app.log(c,'Çok tarihli giriş Excel aktarımı','; '.join(f'{day}: {r["added"]} eklendi, {r["skipped"]} aynı kayıt atlandı' for day,r in summary.items()))
    return summary
