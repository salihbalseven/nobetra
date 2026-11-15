"""Validated import for the supplied 48-person workbook. No Excel installation needed."""
import datetime as dt
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

def norm(value):
    return ' '.join(str(value or '').replace('İ','i').replace('I','ı').casefold().split()).replace('ı','i')

def read_book(path):
    ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with zipfile.ZipFile(path) as z:
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            strings=[''.join(t.text or '' for t in s.findall('.//m:t',ns)) for s in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si',ns)]
        rel={r.get('Id'):r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        wb=ET.fromstring(z.read('xl/workbook.xml')); book={}
        props=wb.find('m:workbookPr',ns)
        date1904=props is not None and props.get('date1904') in ('1','true')
        for s in wb.findall('m:sheets/m:sheet',ns):
            target=rel[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
            target=posixpath.normpath(target.lstrip('/') if target.startswith('/') else 'xl/'+target)
            cells={}
            sheet_root=ET.fromstring(z.read(target))
            for c in sheet_root.findall('.//m:sheetData/m:row/m:c',ns):
                v=c.find('m:v',ns); f=c.find('m:f',ns)
                value=v.text if v is not None else ''
                if c.get('t')=='s' and value:value=strings[int(value)]
                if c.get('t')=='inlineStr':value=''.join(t.text or '' for t in c.findall('.//m:t',ns))
                cells[c.get('r')]={'value':value or '', 'formula':f.text if f is not None else None,'date1904':date1904}
            for merge in sheet_root.findall('m:mergeCells/m:mergeCell',ns):
                first,last=merge.get('ref').split(':')
                def address(address):
                    letters,row=re.fullmatch(r'([A-Z]+)(\d+)',address).groups();col=0
                    for letter in letters:col=col*26+ord(letter)-64
                    return col,int(row)
                left,top=address(first);right,bottom=address(last)
                if (right-left+1)*(bottom-top+1)>10000:continue
                for row in range(top,bottom+1):
                    for col in range(left,right+1):
                        letters='';index=col
                        while index:index,remainder=divmod(index-1,26);letters=chr(65+remainder)+letters
                        cell=f'{letters}{row}'
                        if cell!=first:cells.setdefault(cell,{})['merged_from']=first
            book[s.get('name')]=cells
        return book

def resolve_date(cells,cell,year,today=None,seen=None):
    """Read saved Excel dates; evaluate only explicitly supported uncached formulas."""
    seen=set() if seen is None else set(seen)
    if cell in seen:raise ValueError(f'Yatakhane Nöbet!{cell}: döngüsel tarih formülü')
    seen.add(cell);item=cells.get(cell,{});text=str(item.get('value','')).strip()
    label=f'Yatakhane Nöbet!{cell}'
    formula=re.sub(r'\s+','',item.get('formula') or '').lstrip('=').upper()
    current=today or dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).date()
    match=re.fullmatch(r'(?:TODAY|BUGÜN)\(\)([+-]\d+)?',formula)
    if match:return current+dt.timedelta(days=int(match.group(1) or 0))
    match=re.fullmatch(r'\$?([A-Z]+)\$?(\d+)([+-]\d+)?',formula)
    if match:return resolve_date(cells,match[1]+match[2],year,current,seen)+dt.timedelta(days=int(match[3] or 0))
    match=re.fullmatch(r'IF\((\$?[A-Z]+\$?\d+)="","",(\$?[A-Z]+\$?\d+)([+-]\d+)?\)',formula)
    if match and match[1].replace('$','')==match[2].replace('$',''):
        return resolve_date(cells,match[2].replace('$',''),year,current,seen)+dt.timedelta(days=int(match[3] or 0))
    if text:
        if re.fullmatch(r'\d+(?:\.\d+)?',text) and ('.' not in text or len(text.split('.')[0])>=4):
            serial=float(text)
            if serial<1 or serial>2958465:raise ValueError(f'{label}: geçersiz Excel tarih sayısı {text}')
            if not item.get('date1904') and serial<61:raise ValueError(f'{label}: 1900 yılı tarihleri desteklenmiyor')
            epoch=dt.date(1904,1,1) if item.get('date1904') else dt.date(1899,12,30)
            return epoch+dt.timedelta(days=int(serial))
        match=re.fullmatch(r'(\d{1,2})[./](\d{1,2})(?:[./](\d{4}))?\.?',text)
        if match:
            day,month,explicit_year=match.groups()
            return dt.date(int(explicit_year or year),int(month),int(day))
        try:return dt.date.fromisoformat(text[:10])
        except ValueError:raise ValueError(f'{label}: geçersiz tarih {text!r}')
    formula=re.sub(r'\s+','',item.get('formula') or '').lstrip('=').upper()
    current=today or dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).date()
    match=re.fullmatch(r'(?:TODAY|BUGÜN)\(\)([+-]\d+)?',formula)
    if match:return current+dt.timedelta(days=int(match.group(1) or 0))
    match=re.fullmatch(r'\$?([A-Z]+)\$?(\d+)([+-]\d+)?',formula)
    if match:return resolve_date(cells,match[1]+match[2],year,current,seen)+dt.timedelta(days=int(match[3] or 0))
    raise ValueError(f'{label}: tarih okunamadı. Excel’de dosyayı yeniden hesaplatıp kaydedin. Formül: {formula or "yok"}')

def night_start_day(path,kind):
    sheet,cell={'Yatakhane Kat':('Yatakhane Nöbet','A7'),'Devriye':('DEVRİYE NÖBET','A3')}[kind]
    book=read_book(path)
    if sheet not in book:raise ValueError(f'{sheet} sayfası bulunamadı')
    year=dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).year
    try:return resolve_date(book[sheet],cell,year).isoformat()
    except ValueError as error:raise ValueError(str(error).replace('Yatakhane Nöbet',sheet)) from error

def build_plan(path, selected_day):
    book=read_book(path)
    for title in ('48 KİŞİ GİRİŞ','Yatakhane Nöbet'):
        if title not in book:raise ValueError(f'Sayfa bulunamadı: {title}. Bulunanlar: '+', '.join(book))
    entry=book['48 KİŞİ GİRİŞ']; chart=book['Yatakhane Nöbet']
    def val(cells,cell):
        item=cells.get(cell,{})
        if item.get('formula'):raise ValueError(f'{cell}: giriş alanında formül var. Ad/soyadı değer olarak yapıştırın.')
        return str(item.get('value','')).strip()
    base=dt.date.fromisoformat(selected_day)
    start_day=resolve_date(chart,'A7',base.year)
    if base!=start_day:raise ValueError(f'Yatakhane Nöbet!A7 tarihi {start_day:%d.%m.%Y}; ekranda {base:%d.%m.%Y} seçili. Tarihi düzeltin. Yıl yoksa ekrandaki yıl kullanılır.')
    next_day=base+dt.timedelta(days=1)
    # Yearless dates across New Year use the following calendar year.
    def check_date(cell,wanted):
        text=str(chart.get(cell,{}).get('value','')).strip()
        got=resolve_date(chart,cell,wanted.year)
        if got!=wanted:raise ValueError(f'Yatakhane Nöbet!{cell}: {wanted:%d.%m.%Y} bekleniyor, {text!r} bulundu.')
    for cell in ('A7','A17','A27'):check_date(cell,base)
    for cell in ('A8','A18','A28'):check_date(cell,next_day)
    blocked=set()
    for r in range(2,52):
        a=val(book.get('SORUMLULAR',{}),f'A{r}');b=val(book.get('SORUMLULAR',{}),f'B{r}')
        if a and b:blocked.add((norm(a),norm(b)))
    positions=[(2,'A',7,'C'),(2,'B',7,'H'),(2,'C',7,'M'),(3,'A',17,'C'),(3,'B',17,'H'),(3,'C',17,'M'),(4,'A',27,'C'),(4,'B',27,'H')]
    result=[]; errors=[]
    for i in range(48):
        r=i+2; first=val(entry,f'A{r}');last=val(entry,f'B{r}')
        if not first or not last:
            errors.append(f'48 KİŞİ GİRİŞ!A{r}:B{r}: ad/soyad eksik');continue
        if (norm(first),norm(last)) in blocked:errors.append(f'48 KİŞİ GİRİŞ!A{r}:B{r}: {first} {last} SORUMLULAR listesinde');continue
        group=i//8; floor,block,top,col=positions[i%8]; row=top+group
        formula=chart.get(f'{col}{row}',{}).get('formula') or ''
        if formula and f"'48 KİŞİ GİRİŞ'!$A${r}" not in formula:
            errors.append(f'Yatakhane Nöbet!{col}{row}: beklenmeyen formül bağlantısı');continue
        if group<4:
            slot=val(chart,f'B{row}')
            if not re.fullmatch(r'\d{2}:\d{2}-\d{2}:\d{2}',slot):raise ValueError(f'Yatakhane Nöbet!B{row}: geçersiz saat {slot!r}')
            expected=['23:00-01:00','01:00-03:00','03:00-05:00','05:00-07:00'][group]
            if slot!=expected:raise ValueError(f'Yatakhane Nöbet!B{row}: bu şablonda {expected} bekleniyor; {slot!r} bulundu.')
            a,b=slot.split('-'); duty_day=base if group==0 else next_day
            start=dt.datetime.combine(duty_day,dt.time.fromisoformat(a)); end=dt.datetime.combine(duty_day,dt.time.fromisoformat(b))
            if end<=start:end+=dt.timedelta(days=1)
        else:slot='Yedek';start=end=None
        result.append(dict(first=first,last=last,slot=slot,location=f'{floor}{block}',role='Asıl' if group<4 else 'Yedek',source=f'48 KİŞİ GİRİŞ!A{r}:D{r} → Yatakhane Nöbet!{col}{row}',start=start.isoformat(timespec='minutes') if start else '',end=end.isoformat(timespec='minutes') if end else ''))
    if errors:raise ValueError(f'{len(errors)} hata; hiçbir kayıt aktarılmadı.\n'+'\n'.join(errors))
    return result

def save_plan(app, items, day):
    """One transaction: either all valid rows commit, or nothing changes."""
    with app.conn() as c:
        c.execute('INSERT OR IGNORE INTO sheets(day,kind) VALUES(?,?)',(day,'Yatakhane Kat'))
        sheet=c.execute('SELECT * FROM sheets WHERE day=? AND kind=?',(day,'Yatakhane Kat')).fetchone()
        if sheet['status']=='Onaylandı':raise ValueError('Onaylı çizelge değiştirilemez')
        students={}
        for st in c.execute('SELECT * FROM students WHERE active=1'):
            students.setdefault((norm(st['name']),norm(st['surname'])),[]).append(st)
        pending=[];errors=[];used=set();skipped=0
        for item in items:
            source=item['source'];matches=students.get((norm(item['first']),norm(item['last'])),[])
            if len(matches)!=1:
                errors.append(f'{source}: '+('öğrenci bulunamadı; öğrenci listesini aktarın' if not matches else 'birden fazla isim eşleşmesi; okul numarasıyla ayırın'));continue
            st=matches[0];sid=st['id']
            if sid in used:errors.append(f'{source}: aynı öğrenci iki kez yazılmış');continue
            used.add(sid)
            old=c.execute('SELECT * FROM assignments WHERE sheet_id=? AND student_id=?',(sheet['id'],sid)).fetchone()
            if old:
                if all(old[k]==item[v] for k,v in [('slot','slot'),('location','location'),('role','role'),('start_at','start'),('end_at','end')]):skipped+=1;continue
                errors.append(f'{source}: mevcut görev farklı; taslağı kontrol edin');continue
            duty_day=item['start'][:10] or day
            reason=app.eligible(c,st,duty_day,'Yatakhane Kat',item['slot'])
            if reason:errors.append(f'{source}: {reason}');continue
            if app.has_conflict(c,sid,day,'Yatakhane Kat',item['start'],item['end']):
                errors.append(f'{source}: aynı gece devriye/yatakhane kat grubunda başka görev var');continue
            pending.append((sheet['id'],sid,item['slot'],item['location'],item['role'],item['start'],item['end'],source))
        count=c.execute('SELECT COUNT(*) FROM assignments WHERE sheet_id=?',(sheet['id'],)).fetchone()[0]
        if count+len(pending)>48:errors.append('Mevcut listeyle birlikte 48 kişilik sınır aşılıyor.')
        if errors:raise ValueError('Hiçbir kayıt aktarılmadı.\n'+'\n'.join(errors))
        c.executemany('INSERT INTO assignments(sheet_id,student_id,slot,location,role,start_at,end_at,source_cell) VALUES(?,?,?,?,?,?,?,?)',pending)
        app.log(c,'Excel yatakhane aktarımı',f'{day}: {len(pending)} eklendi, {skipped} aynı kayıt atlandı')
        return len(pending),skipped
