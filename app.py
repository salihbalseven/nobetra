"""NÖBETRA — offline Tkinter + SQLite, Python 3.10+."""
import csv, datetime as dt, html, os, shutil, sqlite3, sys, tempfile, webbrowser, zipfile, re, json, xml.etree.ElementTree as ET
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import traceback
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager, nullcontext
import excel_yatakhane
import excel_devriye
import nobet_istatistik
import excel_giris
import nobet_yazdir
from tarih_secici import DateField
import nobet_planlayici
import ceza_nobet
import arayuz
import sistem_islemleri
import istatistik_grafik
import ogrenci_profili
import ogrenci_excel
import asistan
import nobet_manuel
import ogrenci_kaynak
import nobet_arsiv_ui
import pencere_yonetimi

ASSETS=Path(__file__).resolve().parent
ROOT=Path(sys.executable).resolve().parent if getattr(sys,'frozen',False) else ASSETS
if not getattr(sys, 'frozen', False) and ASSETS.name == 'Uygulama':
    ROOT = ASSETS.parent
DATA=ROOT/'veri'
DATA.mkdir(exist_ok=True)
DB=DATA/'nobet.db'
TYPES={'Yatakhane Kat':48,'Devriye':excel_devriye.TOTAL,'Yatakhane Giriş':14}
TIMES={'Yatakhane Kat':['23:00-01:00','01:00-03:00','03:00-05:00','05:00-07:00'],
       'Devriye':['Haberleşme','Devriye','Yedek'],
       'Yatakhane Giriş':['08:00-10:00','10:00-12:00','12:00-14:00','14:00-16:00','16:00-17:00','Yedek']}
CLASSES=['A','B','C','D','E','F','G','H','İ','K','L','M','N','P','R','S']

@contextmanager
def conn():
    c=sqlite3.connect(DB)
    c.row_factory=sqlite3.Row
    c.execute('PRAGMA foreign_keys=ON')
    try:
        with c:yield c
    finally:c.close()

def setup():
    with conn() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS students(id INTEGER PRIMARY KEY,school_no TEXT UNIQUE,name TEXT NOT NULL,surname TEXT NOT NULL,class TEXT,room TEXT,active INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS exemptions(id INTEGER PRIMARY KEY,student_id INTEGER NOT NULL REFERENCES students(id),reason TEXT NOT NULL,start TEXT NOT NULL,end TEXT);
        CREATE TABLE IF NOT EXISTS restrictions(weekday INTEGER NOT NULL,class TEXT NOT NULL,PRIMARY KEY(weekday,class));
        CREATE TABLE IF NOT EXISTS sheets(id INTEGER PRIMARY KEY,day TEXT NOT NULL,kind TEXT NOT NULL,status TEXT DEFAULT 'Taslak',created TEXT DEFAULT CURRENT_TIMESTAMP,UNIQUE(day,kind));
        CREATE TABLE IF NOT EXISTS assignments(id INTEGER PRIMARY KEY,sheet_id INTEGER NOT NULL REFERENCES sheets(id),student_id INTEGER NOT NULL REFERENCES students(id),slot TEXT NOT NULL,location TEXT NOT NULL DEFAULT '',role TEXT NOT NULL DEFAULT 'Asıl',result TEXT NOT NULL DEFAULT 'Planlandı',UNIQUE(sheet_id,student_id));
        CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY,created TEXT DEFAULT CURRENT_TIMESTAMP,action TEXT,detail TEXT);
        ''')
        columns={r['name'] for r in c.execute('PRAGMA table_info(assignments)')}
        for column in ('start_at','end_at','source_cell'):
            if column not in columns:c.execute(f'ALTER TABLE assignments ADD COLUMN {column} TEXT')
        if 'manual_position' not in columns:
            c.execute('ALTER TABLE assignments ADD COLUMN manual_position INTEGER')
        nobet_planlayici.setup(c)
        ceza_nobet.setup(c)
        ogrenci_profili.setup(c)
        if not c.execute('SELECT 1 FROM restrictions LIMIT 1').fetchone():
            for day, group in enumerate([CLASSES[:4],CLASSES[4:8],CLASSES[8:12],CLASSES[12:]],1):
                c.executemany('INSERT INTO restrictions VALUES(?,?)',[(day,x) for x in group])

def log(c,action,detail):c.execute('INSERT INTO audit(action,detail) VALUES(?,?)',(action,detail))

def busy_students(c,day,kind,start='',end=''):
    if kind not in TYPES:raise ValueError('Geçersiz nöbet türü')
    # Gate duties are independent. Both night duty types share one nightly pool.
    kinds=('Yatakhane Giriş','Yatakhane Giriş') if kind=='Yatakhane Giriş' else ('Yatakhane Kat','Devriye')
    rows=list(c.execute('''SELECT a.student_id,sh.kind,st.class FROM assignments a JOIN sheets sh ON sh.id=a.sheet_id JOIN students st ON st.id=a.student_id
        WHERE a.result!='Tutmadı' AND sh.kind IN (?,?) AND
        (sh.day=? OR (?!='' AND a.start_at IS NOT NULL AND a.start_at!='' AND a.start_at<? AND a.end_at>?))''',
        (*kinds,day,start,end,start)))
    busy={r['student_id'] for r in rows}
    busy.update(r[0] for r in c.execute("SELECT DISTINCT a.student_id FROM assignments a JOIN sheets s ON s.id=a.sheet_id WHERE s.day=? AND a.result!='Tutmadı' AND EXISTS(SELECT 1 FROM penalties p WHERE p.student_id=a.student_id AND p.start_day<=? AND p.voided=0)",(day,day)))
    if kind!='Yatakhane Giriş':
        blocked={excel_yatakhane.norm(r['class']) for r in rows if r['kind']!=kind}
        busy.update(r['id'] for r in c.execute('SELECT id,class FROM students') if excel_yatakhane.norm(r['class']) in blocked)
    return busy

def has_conflict(c,student_id,day,kind,start='',end=''):
    return student_id in busy_students(c,day,kind,start,end)

def set_assignment_results(ids,result):
    ids=list(dict.fromkeys(int(aid) for aid in ids))
    if not ids:raise ValueError('Önce tablodan bir veya daha fazla görev seçin')
    if result not in ('Tuttu','Tutmadı','Planlandı'):raise ValueError('Geçersiz nöbet durumu')
    with conn() as c:
        c.execute('BEGIN IMMEDIATE')
        rows=[]
        for aid in ids:
            row=c.execute('SELECT a.sheet_id,sh.status FROM assignments a JOIN sheets sh ON sh.id=a.sheet_id WHERE a.id=?',(aid,)).fetchone()
            if not row:raise ValueError('Seçilen bir kayıt artık yok. Listeyi yenileyip tekrar seçin.')
            if row['status']=='Onaylandı':raise ValueError('Onaylı çizelge değiştirilemez; hiçbir kayıt değiştirilmedi')
            rows.append(row)
        if len({r['sheet_id'] for r in rows})!=1:raise ValueError('Görevler aynı çizelgeden seçilmeli')
        c.executemany('UPDATE assignments SET result=? WHERE id=?',[(result,aid) for aid in ids])
        log(c,'Toplu görev durumu',f'{len(ids)} kayıt → {result}; kayıtlar: '+','.join(map(str,ids)))
    return len(ids)
def eligible(c,student,day,kind,slot,*,allow_exempt=False):
    d=dt.date.fromisoformat(day); wd=d.isoweekday()
    if kind=='Yatakhane Giriş' and wd>5:return 'Hafta sonu giriş nöbeti yok'
    ex=c.execute('SELECT reason FROM exemptions WHERE student_id=? AND start<=? AND (end IS NULL OR end="" OR end>=?)',(student['id'],day,day)).fetchone()
    if ex and not allow_exempt:return 'Muaf: '+ex['reason']
    rest=ceza_nobet.rest_reason(c,student['id'],day)
    if rest:return rest
    if kind=='Yatakhane Giriş':
        if slot not in TIMES[kind]:return 'Giriş nöbeti saati 08–10, 10–12, 12–14, 14–16 veya 16–17 olmalı'
        if wd==5:
            if slot!='Yedek':
                group=CLASSES[:8] if slot=='08:00-10:00' else CLASSES[8:]
                if excel_yatakhane.norm(student['class']) not in {excel_yatakhane.norm(cls) for cls in group}:
                    return 'Cuma giriş nöbetinde 08–10 yalnızca A–H; 10–12, 12–14, 14–16 ve 16–17 yalnızca İ–S sınıflarına ayrılmıştır'
        elif c.execute('SELECT 1 FROM restrictions WHERE weekday=? AND class=?',(wd,student['class'])).fetchone():return 'Bugün sınıf nöbete yasaklı'
    return ''

def duration(kind,slot,start=None,end=None):
    if start and end:
        return (dt.datetime.fromisoformat(end)-dt.datetime.fromisoformat(start)).total_seconds()/3600
    text=str(slot).replace('–','-').replace('—','-')
    if re.fullmatch(r'\d{2}:\d{2}-\d{2}:\d{2}',text):
        a,b=text.split('-');first=dt.datetime.combine(dt.date(2000,1,1),dt.time.fromisoformat(a));last=dt.datetime.combine(first.date(),dt.time.fromisoformat(b))
        if last<=first:last+=dt.timedelta(days=1)
        return (last-first).total_seconds()/3600
    if kind=='Yatakhane Kat':return 2
    if kind=='Yatakhane Giriş':return {'08:00-10:00':2,'10:00-12:00':2,'12:00-14:00':2,'14:00-16:00':2,'16:00-17:00':1}.get(slot,0)
    return 0 # Saat aralığı olmayan eski kayıtlar Excel'den düzeltilene kadar süre bilinmez.

def read_xlsx_cells(path):
    """Read raw user-entry cells in XLSX using the standard library; formula cells are ignored."""
    n={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    with zipfile.ZipFile(path) as z:
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            root=ET.fromstring(z.read('xl/sharedStrings.xml'))
            strings=[''.join(t.text or '' for t in x.findall('.//m:t',n)) for x in root.findall('m:si',n)]
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relation={r.get('Id'):r.get('Target') for r in rels}
        result={}
        for sh in wb.findall('.//m:sheets/m:sheet',n):
            target=relation.get(sh.get('{'+n['r']+'}id'),'')
            if not target:continue
            target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            if target not in z.namelist():continue
            root=ET.fromstring(z.read(target));cells={}
            for c in root.findall('.//m:sheetData/m:row/m:c',n):
                if c.find('m:f',n) is not None:continue
                v=c.find('m:v',n);inline=c.find('m:is',n)
                value=v.text if v is not None else ''
                if c.get('t')=='s' and value:value=strings[int(value)]
                elif inline is not None:value=''.join(t.text or '' for t in inline.findall('.//m:t',n))
                if value:cells[c.get('r')]=value
            result[sh.get('name')]=cells
        return result

APP_VERSION='v2.33'
APP_TITLE=f'NÖBETRA {APP_VERSION}'


class App(tk.Tk):
    def __init__(self):
        super().__init__();self.title(APP_TITLE);self.geometry(f'{min(1420,self.winfo_screenwidth()-40)}x{min(900,self.winfo_screenheight()-80)}');self.minsize(1000,680)
        pencere_yonetimi.install(self)
        icon=ASSETS/'Nobetra.ico'
        if icon.is_file():self.iconbitmap(default=str(icon))
        arayuz.apply_theme(self);arayuz.header(self,APP_VERSION,ASSETS/'Nobetra.png')
        self.tabs=arayuz.Workspace(self);self.tabs.pack(fill='both',expand=True)
        self.pages={}
        for name in ['Öğrenciler','Nöbet Hazırla','Nöbet Geçmişi','İstatistikler','Cezalılar','Muafiyetler','Sınıf Yasakları','Yedekleme','Asistan']:
            f=ttk.Frame(self.tabs.notebook,padding=12);self.tabs.add(f,text=name);self.pages[name]=f
        self.students_ui();self.schedule_ui();self.history_ui();self.stats_ui();self.penalties_ui();self.exempt_ui();self.rules_ui();self.backup_ui()
        asistan.AssistantPanel(self.pages['Asistan'],ASSETS,ROOT/'asistan_adaylari',
                               context=lambda:f'{APP_VERSION} · seçili nöbet türü {self.kind.get()}').pack(fill='both',expand=True)
        self.refresh_all()

    def tree(self,parent,columns,widths=None,*,wide_scroll=False):
        wrap=ttk.Frame(parent);wrap.pack(fill='both',expand=True,pady=8)
        tree=arayuz.DataTree(wrap,columns=columns,show='headings',selectmode='extended')
        tree.bind('<Control-a>',lambda event:(tree.selection_set(tree.get_children()),'break')[1])
        for i,col in enumerate(columns):tree.heading(col,text=col,command=lambda column=col:tree.sort_column(column));tree.column(col,width=(widths[i] if widths else 120),minwidth=55,stretch=False,anchor='w')
        scroll_style='Wide.Vertical.TScrollbar' if wide_scroll else 'Vertical.TScrollbar'
        scroll=ttk.Scrollbar(wrap,orient='vertical',command=tree.yview,style=scroll_style);tree.configure(yscrollcommand=scroll.set)
        horizontal=ttk.Scrollbar(wrap,orient='horizontal',command=tree.xview);tree.configure(xscrollcommand=horizontal.set);horizontal.pack(side='bottom',fill='x')
        scroll.pack(side='right',fill='y');tree.pack(side='left',fill='both',expand=True);return tree
    def clear(self,tree):tree.delete(*tree.get_children())
    def dialog_owner(self):
        grabbed=self.grab_current()
        return grabbed.winfo_toplevel() if grabbed else self

    def alert(self,msg):messagebox.showerror('İşlem yapılamadı',str(msg),parent=self.dialog_owner())
    def refresh_all(self):self.refresh_students();self.refresh_schedule();self.refresh_history();self.refresh_stats();self.refresh_penalties();self.refresh_exempt();self.refresh_rules()

    def students_ui(self):
        f=self.pages['Öğrenciler'];bar=ttk.Frame(f);bar.pack(fill='x');ttk.Label(bar,text='Ad, soyad, okul no veya sınıf ara:').pack(side='left');self.search=tk.StringVar();ttk.Entry(bar,textvariable=self.search,width=30).pack(side='left',padx=8);self.search.trace_add('write',lambda *_:self.refresh_students())
        excel_actions=arayuz.ActionBar(f);excel_actions.pack(fill='x')
        excel_actions.add('Öğrenci Excel şablonunu kaydet',self.save_student_template,'Teal.TButton')
        excel_actions.add('Excel’den toplu öğrenci ekle / güncelle',self.import_student_excel,'Import.TButton')
        excel_actions.add('Sınıf ve yatakhane bilgilerini aktar',self.import_student_sources,'Import.TButton')
        ttk.Button(bar,text='CSV içe aktar',command=self.import_csv).pack(side='right',padx=4)
        ttk.Button(bar,text='Öğrenci ekle',command=self.add_student).pack(side='right')
        self.show_inactive=tk.BooleanVar(value=False)
        ttk.Checkbutton(bar,text='Pasifleri de göster',variable=self.show_inactive,command=self.refresh_students).pack(side='left',padx=6)
        filters=ttk.Frame(f);filters.pack(fill='x',pady=(4,0))
        self.student_class=tk.StringVar(value='Tümü')
        ttk.Label(filters,text='Sınıf').pack(side='left')
        group=ttk.Combobox(filters,textvariable=self.student_class,values=['Tümü',*CLASSES],state='readonly',width=8)
        group.pack(side='left',padx=6)
        group.bind('<<ComboboxSelected>>',lambda event:self.refresh_students())
        ttk.Button(filters,text='Görünen öğrencileri seç',command=lambda:self.stree.selection_set(self.stree.get_children())).pack(side='left')
        ttk.Button(filters,text='Seçimi kaldır',command=lambda:self.stree.selection_remove(self.stree.selection())).pack(side='left',padx=6)
        self.stree=self.tree(f,['ID','Okul No','Ad','Soyad','Sınıf','Koğuş'],[55,110,170,170,75,100])
        self.stree.tag_configure('exempt',foreground='#a61b1b',background='#fee2e2')
        self.stree.column('Ad',width=220,stretch=True);self.stree.column('Soyad',width=220,stretch=True)
        self.stree.bind('<Double-1>',lambda event:self.student_profile())
        self.stree.bind('<<TreeviewSelect>>',lambda event:self.update_student_actions())
        self.student_status=ttk.Label(f,text='Öğrenci seçin; çift tıklayarak profilini açabilirsiniz.');self.student_status.pack(anchor='w',pady=5)
        actions=arayuz.ActionBar(f);actions.pack(fill='x')
        self.student_edit=actions.add('Öğrenci bilgilerini düzenle',self.student_profile)
        actions.add('Öğrenci ekle',self.add_student)
        self.student_delete=actions.add('Öğrenciyi sil / pasife al',lambda:self.student_active(False),'Danger.TButton')
        self.student_restore=actions.add('Yeniden aktifleştir',lambda:self.student_active(True),'Success.TButton')
        self.update_student_actions()
        ttk.Label(f,text='CSV sütunları: okul_no, ad, soyad, sinif, kogus  •  Aynı okul numarası güncellenir.').pack(anchor='w')
    def refresh_students(self):
        if not hasattr(self,'stree'):return
        previous=[self.stree.item(i,'values')[0] for i in self.stree.selection()]
        self.clear(self.stree)
        with conn() as c:
            exempt=nobet_manuel.exemptions(c,dt.date.today().isoformat())
            for r in ogrenci_profili.search(c,self.search.get(),self.show_inactive.get()):
                if self.student_class.get() not in ('Tümü',r['class']):continue
                item=self.stree.insert('', 'end',values=(r['id'],r['school_no'],r['name'],r['surname'],r['class'],r['room']),
                                       tags=('exempt',) if r['id'] in exempt else ())
                if str(r['id']) in previous:self.stree.selection_add(item)
        if hasattr(self,'student_status'):self.update_student_actions()
    def add_student(self):
        ogrenci_profili.Profile(self,sys.modules[__name__])
    def student_profile(self):
        selected=self.stree.selection()
        if selected:ogrenci_profili.Profile(self,sys.modules[__name__],int(self.stree.item(selected[0],'values')[0]))
    def update_student_actions(self):
        selected=self.stree.selection();state='normal' if len(selected)==1 else 'disabled'
        for button in (self.student_edit,self.student_delete,self.student_restore):button.configure(state=state)
        if len(selected)==1:
            row=self.stree.item(selected[0],'values')
            with conn() as c:active=c.execute('SELECT active FROM students WHERE id=?',(row[0],)).fetchone()[0]
            self.student_status.config(text=f'{row[2]} {row[3]} • '+('Aktif' if active else 'Pasif — yeni nöbete alınmaz'))
            self.student_delete.configure(state='normal' if active else 'disabled');self.student_restore.configure(state='disabled' if active else 'normal')
        else:self.student_status.config(text=f'{len(self.stree.get_children())} öğrenci • {len(selected)} seçili')
    def student_active(self,active):
        selected=self.stree.selection()
        if len(selected)!=1:return
        sid=int(self.stree.item(selected[0],'values')[0])
        text='Öğrenci yeniden aktif olsun mu?' if active else 'Öğrenci pasife alınsın mı? Geçmişi ve mevcut görevleri korunur; yeni nöbetlere alınmaz.'
        if messagebox.askyesno('Öğrenci durumu',text,parent=self.dialog_owner()):ogrenci_profili.set_active(sys.modules[__name__],sid,active);self.refresh_all()
    def save_student(self,v):
        if not all(v[:3]):raise ValueError('Okul no, ad ve soyad zorunludur')
        with conn() as c:c.execute('INSERT INTO students(school_no,name,surname,class,room) VALUES(?,?,?,?,?)',tuple(v));log(c,'Öğrenci eklendi',v[0])
        self.refresh_all()
    def import_csv(self):
        path=filedialog.askopenfilename(filetypes=[('CSV','*.csv')],parent=self.dialog_owner());
        if not path:return
        try:
            with open(path,encoding='utf-8-sig',newline='') as f:rows=list(csv.DictReader(f))
            if not rows or not {'okul_no','ad','soyad','sinif','kogus'}.issubset(rows[0]):raise ValueError('CSV başlıkları okul_no,ad,soyad,sinif,kogus olmalı')
            with conn() as c:
                for r in rows:
                    if not r['okul_no'].strip():continue
                    c.execute('INSERT INTO students(school_no,name,surname,class,room) VALUES(?,?,?,?,?) ON CONFLICT(school_no) DO UPDATE SET name=excluded.name,surname=excluded.surname,class=excluded.class,room=excluded.room',tuple(r[k].strip() for k in ['okul_no','ad','soyad','sinif','kogus']))
                log(c,'CSV aktarıldı',f'{len(rows)} satır')
            self.refresh_all();messagebox.showinfo('Tamam','Öğrenciler aktarıldı.',parent=self.dialog_owner())
        except Exception as e:self.alert(e)

    def save_student_template(self):
        name='Ogrenci_Aktarim_Sablonu.xlsx'
        dest=filedialog.asksaveasfilename(parent=self,initialfile=name,defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')])
        if dest:
            try:shutil.copyfile(ASSETS/name,dest)
            except Exception as e:self.alert(e)

    def import_student_excel(self):
        path=filedialog.askopenfilename(parent=self,filetypes=[('Excel','*.xlsx')])
        if not path:return
        try:
            rows=ogrenci_excel.read(path,CLASSES)
            with conn() as c:
                existing={r[0] for r in c.execute('SELECT school_no FROM students')}
            added=sum(r['school_no'] not in existing for r in rows)
            updated=len(rows)-added
            if not messagebox.askyesno('Öğrenci Excel aktarımı',f'{added} yeni öğrenci eklenecek, {updated} öğrenci okul numarasına göre güncellenecek. Boş isteğe bağlı alanlar korunur. Devam edilsin mi?',parent=self):return
            with conn() as c:
                c.execute('BEGIN IMMEDIATE')
                added,updated=ogrenci_excel.apply(c,rows)
                log(c,'Excel öğrenci aktarımı',f'{added} yeni, {updated} güncellenen')
            self.refresh_all()
            messagebox.showinfo('Aktarım tamamlandı',f'{added} yeni öğrenci eklendi. {updated} öğrenci güncellendi.',parent=self)
        except Exception as e:self.alert(e)

    def import_student_sources(self):
        class_path=filedialog.askopenfilename(parent=self,title='33. Dönem Sınıf Listeleri dosyasını seçin',filetypes=[('Excel','*.xlsx')])
        dorm_path=filedialog.askopenfilename(parent=self,title='33. Dönem Yatakhane Listeleri dosyasını seçin',filetypes=[('Excel','*.xlsx')])
        if not class_path and not dorm_path:return
        try:
            with conn() as c:changes,coverage=ogrenci_kaynak.plan(c,class_path,dorm_path)
            if not changes:raise ValueError('Eşleşen öğrencilerde doldurulacak boş alan bulunamadı.')
            summary=(f'{len(changes)} öğrenciye {sum(map(len,changes.values()))} profil bilgisi eklenecek.\n'
                     f'Eşleşmeyen kaynak satırı: {coverage["sinif_atlandi"]+coverage["yatak_atlandi"]}\n'
                     f'Koğuşu farklı olduğu için ranza bilgisi atlanan: {coverage["kogus_farkli"]}')
            if not messagebox.askyesno('Kaynak bilgilerini aktar',summary,parent=self):return
            with conn() as c:
                c.execute('BEGIN IMMEDIATE')
                if ogrenci_kaynak.plan(c,class_path,dorm_path)!=(changes,coverage):
                    raise ValueError('Öğrenci kayıtları değişti. Aktarımı tekrar açın.')
                ogrenci_kaynak.apply(c,changes)
                log(c,'Öğrenci profili kaynak aktarımı',f'{len(changes)} öğrenci')
            self.refresh_all()
            messagebox.showinfo('Aktarım tamamlandı',f'{len(changes)} öğrenci güncellendi.',parent=self)
        except Exception as e:self.alert(e)

    def dialog(self,title,labels,callback,choices=None,defaults=None):
        w=tk.Toplevel(self);w.title(title);pencere_yonetimi.modal(w,self);w.resizable(False,False);vars=[]
        for i,label in enumerate(labels):
            ttk.Label(w,text=label).grid(row=i,column=0,sticky='w',padx=12,pady=7)
            v=tk.StringVar();vars.append(v)
            if choices and label in choices:wid=ttk.Combobox(w,textvariable=v,values=choices[label],state='readonly',width=34);v.set(choices[label][0])
            elif 'YYYY-AA-GG' in label:wid=DateField(w,v,width=30,allow_empty='boş' in label)
            else:wid=ttk.Entry(w,textvariable=v,width=37)
            if defaults and label in defaults:v.set(defaults[label])
            wid.grid(row=i,column=1,padx=12,pady=7)
        def submit():
            try:
                if callback([v.get().strip() for v in vars]) is not False:w.destroy()
            except Exception as e:self.alert(e)
        ttk.Button(w,text='Kaydet',command=submit).grid(row=len(labels),column=0,columnspan=2,pady=14)
        w.wait_window()

    def schedule_ui(self):
        f=self.pages['Nöbet Hazırla'];bar=ttk.Frame(f);bar.pack(fill='x')
        self.day=tk.StringVar(value=dt.date.today().isoformat());self.kind=tk.StringVar(value='Yatakhane Kat')
        ttk.Label(bar,text='Tarih').pack(side='left');DateField(bar,self.day,on_change=self.refresh_schedule).pack(side='left',padx=6)
        kindbox=ttk.Combobox(bar,textvariable=self.kind,values=list(TYPES),state='readonly',width=19);kindbox.pack(side='left',padx=6);kindbox.bind('<<ComboboxSelected>>',lambda event:self.refresh_schedule())
        for label,command in [('‹ Önceki',lambda:self.shift_day(-1)),('Bugün',self.go_today),
                              ('Sonraki ›',lambda:self.shift_day(1)),('Listeyi aç',self.refresh_schedule)]:
            ttk.Button(bar,text=label,command=command).pack(side='left',padx=3)

        status=ttk.Frame(f)
        status.pack(fill='x',pady=(8,4))
        self.pending_button=ttk.Menubutton(status,text='Bekleyen taslaklar (0)')
        self.pending_button.pack(side='right')
        self.pending_menu=tk.Menu(self.pending_button,tearoff=False)
        self.pending_button.configure(menu=self.pending_menu)
        self.sinfo=ttk.Label(status,text='',wraplength=630)
        self.sinfo.pack(side='left')

        self.conflict_panel=ttk.Frame(f)
        self.conflict_panel.pack(fill='x',pady=(0,5))
        conflict_actions=ttk.Frame(self.conflict_panel)
        conflict_actions.pack(side='right',padx=(8,0))
        for kind,label in (('Yatakhane Kat','Yatakhane'),('Devriye','Devriye')):
            ttk.Button(conflict_actions,text=f'{label} sınıf sırasını değiştir',
                       command=lambda target=kind:self.edit_class_priority(target,self.conflict_day),
                       style='Warning.TButton').pack(side='left',padx=(4,0))
        self.conflict_text=ttk.Label(self.conflict_panel,text='Gece nöbeti sınıf sırası',wraplength=390,justify='left')
        self.conflict_text.pack(side='left',fill='x',expand=True)
        self.conflict_day=None
        self.conflict_parent=f
        self.conflict_executor=ThreadPoolExecutor(max_workers=1)
        self.conflict_future=None
        self.conflict_refresh_pending=False
        f.bind('<Map>',lambda _:self.refresh_class_forecast(),add='+')
        self.after(120_000,self.refresh_class_forecast_periodically)

        prepare=ttk.LabelFrame(f,text='Listeyi hazırla ve onayla',padding=(8,2))
        prepare.pack(fill='x',pady=(2,5))
        actions=arayuz.ActionBar(prepare);actions.pack(fill='x')
        actions.add('Otomatik hazırla',self.plan_auto,'Primary.TButton')
        actions.add('Manuel hazırla',self.plan_manual,'Teal.TButton')
        actions.add('Sonraki günü hazırla ›',self.prepare_next_day,'Primary.TButton')
        self.confirm_button=actions.add('Bu listeyi onayla',self.confirm_sheet,'Success.TButton')
        actions.add('Yazdır / Excel / PDF',self.print_sheet,'Teal.TButton')
        self.preparation_warning=ttk.Label(prepare,foreground='#9a3412',wraplength=720)
        self.preparation_warning.pack(fill='x')
        self.preparation_warning.pack_forget()

        selection=ttk.LabelFrame(f,text='Nöbet sonuçlarını işle',padding=(8,2))
        selection.pack(fill='x',pady=(0,3))
        results=arayuz.ActionBar(selection);results.pack(fill='x')
        results.add('Asılları seç',self.select_primaries,'Teal.TButton')
        results.add('Yedekleri seç',lambda:self.select_role('Yedek'),'Teal.TButton')
        results.add('Seçilenler: Tuttu',lambda:self.set_result('Tuttu'),'Success.TButton')
        results.add('Seçilenler: Tutmadı',lambda:self.set_result('Tutmadı'),'Warning.TButton')
        results.add('Geçmiş nöbet listesi yükle',lambda:nobet_arsiv_ui.open_dialog(self,sys.modules[__name__]),'Primary.TButton')
        extras=ttk.LabelFrame(f,text='Liste işlemleri ve ayarlar',padding=(8,2))
        extras.pack(fill='x',pady=(2,3))
        editing=arayuz.ActionBar(extras);editing.pack(fill='x')
        for label,command in [
            ('Excel aktar',self.import_excel_schedule),('Toplu yapıştır / aktar',self.bulk_dialog),
            ('Öğrenci görevlendir',self.assign),('Seçili öğrenciyi değiştir',lambda:self.replace_selected_student()),('Az nöbet tutanları öner',self.suggest),
            ('Bu günün listelerini onayla',self.confirm_current_day),('Onayı kaldır',self.unconfirm_sheet),
            ('Durum seçenekleri',self.set_result),('Tutmadı bildirimi',self.show_missed_notice),
            ('Sıra başlangıçları',self.planner_settings),('Giriş sınıf sıraları',self.gate_class_settings),
            ('Gece nöbeti devam sırası',self.night_continuation),
            ('Sınıf önceliğini değiştir',self.edit_class_priority),
        ]:
            editing.add(label,command,'Quiet.TButton')
        editing.add('Seçilenleri sil',self.remove_assign,'Compact.Danger.TButton')
        editing.add('Tüm listeyi sil',self.clear_schedule,'Compact.Danger.TButton')
        editing.add('Planlama kuralları',self.show_planning_rules,'Quiet.TButton')

        def clear_mouse_focus(event):
            button=event.widget
            def clear():
                if button.winfo_exists() and self.focus_get()==button:
                    f.focus_set()
            self.after_idle(clear)

        for button in [*bar.winfo_children(),*actions.buttons,*results.buttons,*editing.buttons]:
            if isinstance(button,ttk.Button):
                button.bind('<ButtonRelease-1>',clear_mouse_focus,add='+')
        self.atree=self.tree(f,['Kayıt ID','Okul No','Ad Soyad','Sınıf','Saat / Görev','Kat / Blok','Rol','Durum','Görev Tipi','Değiştir'],[75,105,180,65,130,110,80,110,90,70])
        self.atree.tag_configure('exempt',foreground='#a61b1b',background='#fee2e2')
        self.atree.bind('<Button-1>',self.replacement_click,add='+')
    def show_planning_rules(self):
        messagebox.showinfo('Planlama kuralları',
            'Sıra, liste onaylanınca ilerler.\n\n'
            'Giriş nöbeti ile gece nöbeti birlikte yazılabilir. Devriye ve yatakhane kat aynı gece birlikte yazılamaz; yedekler de bu kontrole dahildir.\n\n'
            'Yedekler ve Tutmadı kayıtları tutulmuş nöbet sayılmaz.',parent=self)

    def sheet_id(self,c,create=False):
        day=self.day.get().strip();dt.date.fromisoformat(day);kind=self.kind.get()
        if create:c.execute('INSERT OR IGNORE INTO sheets(day,kind) VALUES(?,?)',(day,kind))
        r=c.execute('SELECT * FROM sheets WHERE day=? AND kind=?',(day,kind)).fetchone();return r
    def refresh_schedule(self):
        self.clear(self.atree)
        self.preparation_warning.pack_forget()
        self.confirm_button.state(['disabled'])
        try:
            with conn() as c:
                sheet=self.sheet_id(c)
                if sheet:
                    exempt=nobet_manuel.exemptions(c,sheet['day'])
                    rows=c.execute('SELECT a.*,s.school_no,s.name,s.surname,s.class FROM assignments a JOIN students s ON s.id=a.student_id WHERE sheet_id=? ORDER BY a.id',(sheet['id'],)).fetchall()
                    for r in rows:
                        slot=r['slot']
                        if r['start_at']:slot=f"{r['start_at'][8:10]}.{r['start_at'][5:7]} {slot}"
                        change='×' if sheet['status']!='Onaylandı' and r['result']=='Planlandı' else ''
                        self.atree.insert('','end',values=(r['id'],r['school_no'],r['name']+' '+r['surname'],r['class'],slot,r['location'],r['role'],r['result'],r['duty_type'],change),
                                          tags=('exempt',) if r['student_id'] in exempt else ())
                    limit=42 if sheet['kind']=='Devriye' and nobet_yazdir.legacy_patrol(rows) else TYPES[sheet['kind']]
                    if sheet['status']!='Onaylandı':
                        warning=nobet_manuel.preparation_warning(sheet['kind'],rows)
                        if warning:
                            self.preparation_warning.config(text=warning)
                            self.preparation_warning.pack(fill='x',pady=(0,4))
                            limit=TYPES[sheet['kind']]
                        elif rows:self.confirm_button.state(['!disabled'])
                    self.sinfo.config(text=f"{self.day.get()} • {self.kind.get()} • {len(rows)}/{limit} kişi • {sheet['status']}")
                else:self.sinfo.config(text='Bu tarih ve nöbet türü için henüz liste oluşturulmadı.')
                pending=sistem_islemleri.pending(c)
                self.pending_menu.delete(0,'end')
                self.pending_button.configure(text=f'Bekleyen taslaklar ({len(pending)})',
                                              state='normal' if pending else 'disabled')
                for draft in pending:
                    label=f'{draft["day"]} · {draft["kind"]} · {draft["undecided"]} asıl sonuç bekliyor'
                    self.pending_menu.add_command(label=label,command=lambda sheet=draft:self.open_pending(sheet))
                if sheet and sheet['status']=='Onaylandı':self.sinfo.config(text=self.sinfo.cget('text')+' — Tamamlandı; sonraki günü hazırlayabilirsiniz.')
        except Exception as e:self.sinfo.config(text='Tarih hatalı: YYYY-AA-GG kullanın')
        if self.conflict_parent.winfo_ismapped():
            self.refresh_class_forecast()

    def refresh_class_forecast(self):
        if self.conflict_future:
            self.conflict_refresh_pending=True
            return
        def calculate():
            with conn() as c:
                c.execute('BEGIN')
                return nobet_planlayici.forecast_class_conflicts(sys.modules[__name__],c,dt.date.today())
        self.conflict_future=self.conflict_executor.submit(calculate)
        self.after(75,self.show_class_forecast)

    def refresh_class_forecast_periodically(self):
        if self.conflict_parent.winfo_ismapped():
            self.refresh_class_forecast()
        self.after(120_000,self.refresh_class_forecast_periodically)

    def show_class_forecast(self):
        if not self.conflict_future.done():
            self.after(75,self.show_class_forecast)
            return
        result=self.conflict_future
        self.conflict_future=None
        try:
            warnings=result.result()
            if warnings:
                self.conflict_day=warnings[0][0]
                lines=[]
                for day,offset,classes in warnings:
                    when='Yarın' if offset==1 else '2 gün sonra'
                    lines.append(f'{when} ({dt.date.fromisoformat(day):%d.%m}) sınıf çakışabilir · {", ".join(classes)}')
                self.conflict_text.configure(text='\n'.join(lines),foreground='#9a3412')
            else:
                self.conflict_day=None
                self.conflict_text.configure(text='Gece nöbeti sınıf sırası',foreground='')
        except (ValueError,sqlite3.Error):
            self.conflict_day=None
            self.conflict_text.configure(text='Sınıf çakışması tahmini hesaplanamadı.',foreground='#9a3412')
        if self.conflict_refresh_pending:
            self.conflict_refresh_pending=False
            self.refresh_class_forecast()

    def edit_class_priority(self,kind=None,day=None):
        kind=kind or self.kind.get()
        if kind not in nobet_planlayici.NIGHT_KINDS:
            self.alert('Yatakhane Kat veya Devriye listesini seçin.')
            return
        try:
            with conn() as c:
                day=self.next_open_sheet_day(c,kind,day or self.day.get().strip())
                state,_,_=nobet_planlayici.queue_state(sys.modules[__name__],c,kind,day)
            if not state:
                raise ValueError('Önce nöbetin sıra başlangıcını ayarlayın.')
            order=state.get('class_order',CLASSES)
            index=state['turn']%len(order)
            labels=['Başlangıç (YYYY-AA-GG)','Öne alınacak sınıflar']
            def save(values):
                start,raw=values
                classes=[label.upper().replace('I','İ') for label in re.split(r'[-–—,\s]+',raw) if label]
                with conn() as c:
                    c.execute('BEGIN IMMEDIATE')
                    pending=nobet_planlayici.reorder_night_classes(sys.modules[__name__],c,kind,start,classes)
                text=f'{kind} sınıf sırası kaydedildi.'
                if pending:
                    text+=f' Yeni sıra {pending} sınıfı tamamlandıktan sonra başlar.'
                messagebox.showinfo('Sınıf önceliği',text,parent=self.dialog_owner())
                self.refresh_class_forecast()
            defaults={labels[0]:day,labels[1]:'-'.join(order[index:]+order[:index])}
            self.dialog(f'{kind} · Önce mevcut sınıfı tamamla',labels,save,defaults=defaults)
        except (ValueError,sqlite3.Error) as error:
            self.alert(error)
    def plan_manual(self):
        try:
            nobet_manuel.open_dialog(self,sys.modules[__name__],self.day.get().strip(),self.kind.get())
        except (ValueError,sqlite3.Error) as error:
            self.alert(error)

    def replacement_click(self,event):
        if self.atree.identify_region(event.x,event.y)!='cell' or self.atree.identify_column(event.x)!='#10':
            return
        item=self.atree.identify_row(event.y)
        values=self.atree.item(item,'values') if item else ()
        if values and values[-1]=='×':
            self.after_idle(self.replace_schedule_student,int(values[0]))
            return 'break'

    def replace_selected_student(self):
        try:
            self.replace_schedule_student(self.selected_assignment())
        except Exception as error:self.alert(error)

    def replace_schedule_student(self,assignment_id):
        try:
            with conn() as c:
                original=nobet_manuel.replacement_rows(c,assignment_id)
                choices=nobet_manuel.replacement_options(sys.modules[__name__],c,original)
            if not choices:raise ValueError('Bu görev için uygun başka öğrenci bulunamadı.')
            options={f'{st["school_no"]} · {st["name"]} {st["surname"]} · {st["class"]}':st['id'] for st,_ in choices}
            window=tk.Toplevel(self);window.title('Taslakta öğrenciyi değiştir')
            ttk.Label(window,text=f'{original["name"]} {original["surname"]} · {original["slot"]} · {original["location"]} · {original["role"]} · {original["duty_type"]}',padding=12).pack(anchor='w')
            frame=ttk.Frame(window,padding=(12,0,12,12));frame.pack(fill='both',expand=True)
            query=tk.StringVar(window)
            entry=ttk.Entry(frame,textvariable=query);entry.pack(fill='x')
            results=tk.Listbox(frame,height=11,width=66,exportselection=False);results.pack(fill='both',expand=True,pady=8)
            def search(*_):
                results.delete(0,'end')
                for label in options:
                    if excel_yatakhane.norm(query.get()) in excel_yatakhane.norm(label):
                        results.insert('end',label)
            def save(_event=None):
                selected=results.curselection()
                if not selected:return
                try:
                    nobet_manuel.replace_duty(sys.modules[__name__],original,options[results.get(selected[0])])
                except ValueError as error:
                    messagebox.showerror('Öğrenci değiştirilemedi',str(error),parent=window);return
                window.destroy();self.refresh_all()
            query.trace_add('write',search);search()
            results.bind('<Double-1>',save);results.bind('<Return>',save)
            ttk.Button(frame,text='Seçili öğrenciyi yerleştir',style='Primary.TButton',command=save).pack(side='right')
            ttk.Button(frame,text='Vazgeç',command=window.destroy).pack(side='left')
            pencere_yonetimi.modal(window,self);entry.focus_set()
        except Exception as error:self.alert(error)

    def assign(self):
        with conn() as c:
            students=c.execute('SELECT school_no,name,surname FROM students WHERE active=1 ORDER BY id').fetchall()
        options=[f"{r['school_no']} — {r['name']} {r['surname']}" for r in students]
        if not options:return self.alert('Önce öğrenci listesi aktarılmalı')
        kind=self.kind.get();slots=TIMES[kind]
        def save(v):
            school=v[0].split(' — ')[0];slot,location,role=v[1:]
            day=self.day.get()
            def penalty(c,s):
                exempt=s['id'] in nobet_manuel.exemptions(c,day)
                pid,warning=nobet_manuel.penalty_or_warning(c,s['id'],day,kind,role) if exempt else (None,'')
                return exempt,pid,warning
            with conn() as c:
                s=c.execute('SELECT * FROM students WHERE school_no=? AND active=1',(school,)).fetchone()
                if not s:raise ValueError('Öğrenci artık aktif değil.')
                expected=penalty(c,s)
            if expected[2] and not messagebox.askyesno('Muaf öğrenciye nöbet',
                    f'{s["name"]} {s["surname"]} · {expected[2]}\n\nNormal görev olarak devam etmek istiyor musunuz?',parent=self):
                return False
            with conn() as c:
                c.execute('BEGIN IMMEDIATE')
                current=c.execute('SELECT * FROM students WHERE school_no=? AND active=1',(school,)).fetchone()
                if not current or current['id']!=s['id'] or penalty(c,current)!=expected:
                    raise ValueError('Öğrencinin muafiyet veya ceza bilgisi değişti. Görevlendirmeyi tekrar açın.')
                exempt,pid,warning=expected
                err=eligible(c,current,day,kind,slot,allow_exempt=exempt)
                if err:raise ValueError(err)
                if has_conflict(c,s['id'],self.day.get(),kind):raise ValueError('Öğrenci aynı tarih için bu nöbet grubunda zaten görevli. Devriye ve yatakhane kat aynı gece birlikte yazılamaz; giriş nöbeti ayrı değerlendirilir.')
                sheet=self.sheet_id(c,True)
                if sheet['status']=='Onaylandı':raise ValueError('Onaylı çizelge değiştirilemez')
                if c.execute('SELECT COUNT(*) FROM assignments WHERE sheet_id=?',(sheet['id'],)).fetchone()[0]>=TYPES[kind]:raise ValueError('Bu listenin kişi sınırına ulaşıldı')
                c.execute('INSERT INTO assignments(sheet_id,student_id,slot,location,role,duty_type,penalty_id) VALUES(?,?,?,?,?,?,?)',
                          (sheet['id'],s['id'],slot,location,role,'Ceza' if pid else 'Normal',pid))
                if warning:log(c,'Muaf öğrenciye manuel görev',f'{day} · {kind} · öğrenci kaydı {s["id"]}')
                log(c,'Görev eklendi',f'{self.day.get()} {kind} {school} {slot}')
            self.refresh_all()
        self.dialog('Nöbet görevlendir',['Öğrenci','Saat / Görev','Kat / Blok','Rol'],save,{'Öğrenci':options,'Saat / Görev':slots,'Rol':['Asıl','Yedek']})
    def import_excel_schedule(self):
        """Import the actual input sheets from our three roster workbooks, without third-party dependencies."""
        try:
            kind=self.kind.get();day=self.day.get().strip()
            path=filedialog.askopenfilename(parent=self,title='Nöbet Excel dosyasını seçin',filetypes=[('Excel dosyaları','*.xlsx')])
            if not path:return
            if kind in ('Yatakhane Kat','Devriye'):day=excel_yatakhane.night_start_day(path,kind)
            if kind=='Yatakhane Giriş':
                batches=excel_giris.build_batches(path)
                w=tk.Toplevel(self);w.title('Excel tarihine göre giriş nöbeti');w.geometry('960x620');pencere_yonetimi.modal(w,self)
                ttk.Label(w,text='Tarihler Excel’den okundu. Tüm tarihler seçiliyken bütün günler tek işlemde kendi listelerine aktarılır.\nHer günün sınıf yasağı ayrı kontrol edilir. Aynı mevcut kayıtlar atlanır; hata olursa hiçbir güne yeni kayıt eklenmez.',wraplength=900).pack(padx=12,pady=12)
                dates=sorted(batches);chosen=tk.StringVar(value='Tüm tarihler' if len(dates)>1 else dates[0])
                choice=ttk.Combobox(w,textvariable=chosen,values=['Tüm tarihler']+dates,state='readonly',width=24);choice.pack(pady=5)
                summary=ttk.Label(w,text='');summary.pack(pady=5)
                text=tk.Text(w,wrap='word');text.pack(fill='both',expand=True,padx=12)
                def preview_gate(*args):
                    selected_dates=dates if chosen.get()=='Tüm tarihler' else [chosen.get()]
                    summary.configure(text=f'{len(selected_dates)} tarih — toplam {sum(len(batches[day]) for day in selected_dates)} öğrenci')
                    text.configure(state='normal');text.delete('1.0','end')
                    for selected_day in selected_dates:
                        items=batches[selected_day]
                        weekday=['Pazartesi','Salı','Çarşamba','Perşembe','Cuma','Cumartesi','Pazar'][dt.date.fromisoformat(selected_day).weekday()]
                        text.insert('end',f'\n{selected_day} — {weekday} — {len(items)} öğrenci\n')
                        text.insert('end','\n'.join(f'Satır {row}: {first} {last} | {slot} | {location} | {role}' for first,last,slot,location,role,row,identity in items)+'\n')
                    text.configure(state='disabled')
                def commit_gate(repair=False):
                    try:
                        selected_dates=dates if chosen.get()=='Tüm tarihler' else [chosen.get()]
                        if repair and not messagebox.askyesno('Saat ve blokları düzelt','Mevcut listedeki aynı öğrencilerin saatleri, blokları ve asıl/yedek rolleri Excel çizelgesine göre düzeltilecek. Tuttu/Tutmadı ve onay durumları korunacak; geçmiş ve istatistikler yeni saat/rollere göre hesaplanacak. Devam?',parent=w):return
                        result=excel_giris.save_batches(sys.modules[__name__],self.import_excel_rows,{day:batches[day] for day in selected_dates},repair=repair)
                        self.day.set(selected_dates[0]);w.destroy();self.refresh_all()
                        messagebox.showinfo('Giriş nöbeti aktarıldı','\n'.join(f'{day}: {r["added"]} eklendi, {r.get("updated",0)} düzeltildi, {r["skipped"]} aynı kayıt atlandı.' for day,r in result.items())+'\nYeni kayıtlar Planlandı durumundadır. Mevcut kayıtların durumu korunmuştur.',parent=self)
                    except Exception as exc:self.excel_error(exc,path)
                choice.bind('<<ComboboxSelected>>',preview_gate);preview_gate()
                ttk.Button(w,text='Seçilen tarihleri aktar',command=commit_gate).pack(pady=12)
                ttk.Button(w,text='Mevcut saat/blokları düzelt',command=lambda:commit_gate(True)).pack(pady=6)
                return
            if kind=='Devriye':
                items=excel_devriye.build_plan(path,day)
                with conn() as c:
                    sh=c.execute('SELECT * FROM sheets WHERE day=? AND kind=?',(day,kind)).fetchone()
                    existing=c.execute('SELECT COUNT(*) FROM assignments WHERE sheet_id=?',(sh['id'],)).fetchone()[0] if sh else 0
                w=tk.Toplevel(self);w.title('Devriye saat ve görev önizlemesi');w.geometry('1000x650');pencere_yonetimi.modal(w,self)
                reserve_count=sum(item['role']=='Yedek' for item in items)
                action='Mevcut aynı listenin saat ve yerleri düzeltilecek; Tuttu/Tutmadı durumları ve onayı korunacak.' if existing else f'{len(items)} kişi yeni taslağa Planlandı olarak eklenecek.'
                ttk.Label(w,text=f'Excel’den alınan başlangıç: {day} gecesi — {len(items)-reserve_count} asıl + {reserve_count} yedek\n{action}\nGece yarısından sonraki görevler ertesi gün gösterilir. Yedeklere sabit süre atanmaz.').pack(padx=12,pady=10)
                text=tk.Text(w,wrap='word');text.pack(fill='both',expand=True,padx=12)
                text.insert('1.0','\n'.join(f"{x['source']}: {x['first']} {x['last']} | {x['location']} | {x['start'] or 'Yedek'} → {x['end']}" for x in items));text.configure(state='disabled')
                def commit_patrol():
                    try:
                        count,repaired=excel_devriye.save_plan(sys.modules[__name__],items,day)
                        self.day.set(day);w.destroy();self.refresh_all()
                        messagebox.showinfo('Devriye aktarımı',f'{count} kayıt '+('düzeltildi. Mevcut durumlar ve onay korundu.' if repaired else 'eklendi. Durum: Planlandı.'),parent=self)
                    except Exception as exc:self.excel_error(exc,path)
                ttk.Button(w,text='Saat ve yerleri düzelt' if existing else f'{len(items)} kişiyi toplu aktar',command=commit_patrol).pack(pady=10)
                return
            if kind=='Yatakhane Kat':
                items=excel_yatakhane.build_plan(path,day)
                preview='\n'.join(f"{x['source']}: {x['first']} {x['last']} | {x['location']} | {x['start'] or 'Yedek'} → {x['end']}" for x in items)
                w=tk.Toplevel(self);w.title('48 kişilik aktarım önizlemesi');w.geometry('1000x650');pencere_yonetimi.modal(w,self)
                ttk.Label(w,text=f'Excel’den alınan başlangıç: {day} gecesi — 32 asıl + 16 yedek.\nGece yarısından sonraki görevler ertesi gün başlar. Excel’de yıl yazmıyorsa mevcut yıl kullanılır.').pack(padx=12,pady=10)
                text=tk.Text(w,wrap='word');text.pack(fill='both',expand=True,padx=12);text.insert('1.0',preview);text.configure(state='disabled')
                def commit():
                    try:
                        count,skipped=excel_yatakhane.save_plan(sys.modules[__name__],items,day)
                        self.day.set(day);w.destroy();self.refresh_all()
                        messagebox.showinfo('Aktarım tamamlandı',f'{count} kayıt eklendi; {skipped} aynı kayıt atlandı.\nDurum: Planlandı.',parent=self)
                    except Exception as exc:self.excel_error(exc,path)
                ttk.Button(w,text='48 kişiyi toplu aktar',command=commit).pack(pady=10)
                return
            book=read_xlsx_cells(path)
            sheet_name={'Yatakhane Kat':'48 KİŞİ GİRİŞ','Devriye':'42 KİŞİ GİRİŞ','Yatakhane Giriş':'28 KİŞİ GİRİŞ'}[kind]
            if sheet_name not in book:raise ValueError(f'Bu dosyada {sheet_name} sayfası yok. Seçili nöbet türüyle Excel aynı olmalı.')
            cells=book[sheet_name]
            identity_cells=excel_yatakhane.read_book(path)[sheet_name] if kind in ('Devriye','Yatakhane Giriş') else {}
            items=[]
            if kind in ('Yatakhane Kat','Devriye'):
                limit=TYPES[kind]
                for i in range(limit):
                    row=i+2;first=str(cells.get(f'A{row}','')).strip();last=str(cells.get(f'B{row}','')).strip()
                    if not first and not last:continue
                    if not first or not last:raise ValueError(f'{row}. satırda ad veya soyad eksik')
                    if kind=='Yatakhane Kat':
                        # 32 actual duties: four slots x eight floor/block positions; 16 reserves: two per block.
                        block=['2A','2B','2C','3A','3B','3C','4A','4B']
                        slot=TIMES[kind][i//8] if i<32 else 'Yedek'
                        loc=block[i%8] if i<32 else block[(i-32)//2]
                        role='Asıl' if i<32 else 'Yedek'
                    else:
                        slot='Haberleşme' if i<4 else ('Devriye' if i<36 else 'Yedek')
                        loc='';role='Asıl' if i<36 else 'Yedek'
                    identity={}
                    if kind=='Devriye' and 'okul' in excel_yatakhane.norm(cells.get('E1','')):
                        number=str(identity_cells.get(f'E{row}',{}).get('value','')).strip()
                        if re.fullmatch(r'\d+\.0+',number):number=number.split('.')[0]
                        identity={'school_no':number,'class':str(identity_cells.get(f'C{row}',{}).get('value','')).strip()}
                    items.append((first,last,slot,loc,role,row,identity))
            else:
                # Two-day input sheet: 14 rows per day, split into A/B blocks.
                for i in range(14):
                    row=i+2;first=str(cells.get(f'E{row}','')).strip();last=str(cells.get(f'F{row}','')).strip()
                    if not first and not last:continue
                    if not first or not last:raise ValueError(f'{row}. satırda ad veya soyad eksik')
                    slot=str(cells.get(f'D{row}','')).strip().replace('–','-').replace('—','-')
                    reserve='yedek' in excel_yatakhane.norm(slot)
                    if reserve:slot='Yedek'
                    if slot not in TIMES[kind]:raise ValueError(f'28 KİŞİ GİRİŞ!D{row}: geçersiz giriş nöbeti saati {slot!r}')
                    loc=str(cells.get(f'C{row}','')).strip()
                    number=str(identity_cells.get(f'I{row}',{}).get('value','')).strip()
                    if re.fullmatch(r'\d+\.0+',number):number=number.split('.')[0]
                    identity={'school_no':number,'class':str(identity_cells.get(f'G{row}',{}).get('value','')).strip()}
                    items.append((first,last,slot,loc,'Yedek' if reserve else 'Asıl',row,identity))
            if not items:raise ValueError('Seçilen giriş sayfasında isim bulunamadı. Önce Excel listesine ad/soyad girin.')
            if not messagebox.askyesno('Excel aktarımı',f'{Path(path).name}\nSayfa: {sheet_name}\nTarih: {day}\nTür: {kind}\n\n{len(items)} kişi okunacak.\n\nGiriş nöbeti için yalnızca seçili tarihin ilk 14 kişisi alınır.\nOnaylı listeler değiştirilemez. Devam?',parent=self.dialog_owner()):return
            count=self.import_excel_rows(items,kind,day)
            self.refresh_schedule()
            messagebox.showinfo('Excel aktarımı',f'{count} kişi aktarıldı. Saat, kat ve yedek sıralamasını kontrol edin.\nÖğrenciler henüz Planlandı durumunda.',parent=self.dialog_owner())
        except Exception as e:self.excel_error(e,locals().get('path',''))

    def excel_error(self,error,path):
        detail=f'Dosya: {path}\n{type(error).__name__}: {error}\n\n'+traceback.format_exc()
        report=DATA/'son_excel_hatasi.txt'
        try:report.write_text(detail,encoding='utf-8')
        except OSError:pass
        w=tk.Toplevel(self);w.title('Excel aktarımı yapılamadı');w.geometry('880x580');w.transient(self)
        ttk.Label(w,text='Hata ayrıntıları aşağıda. Doğrulama hatasında hiçbir kayıt eklenmez.').pack(pady=10)
        text=tk.Text(w,wrap='word');text.pack(fill='both',expand=True,padx=12);text.insert('1.0',detail);text.configure(state='disabled')
        ttk.Button(w,text='Ayrıntıları kopyala',command=lambda:(self.clipboard_clear(),self.clipboard_append(detail))).pack(pady=8)
        ttk.Label(w,text=f'Rapor: {report}').pack(pady=6)

    def import_excel_rows(self,items,kind,day,connection=None):
        dt.date.fromisoformat(day)
        with (conn() if connection is None else nullcontext(connection)) as c:
            c.execute('INSERT OR IGNORE INTO sheets(day,kind) VALUES(?,?)',(day,kind))
            sh=c.execute('SELECT * FROM sheets WHERE day=? AND kind=?',(day,kind)).fetchone()
            if sh['status']=='Onaylandı':raise ValueError('Onaylı çizelge değiştirilemez')
            current=c.execute('SELECT COUNT(*) FROM assignments WHERE sheet_id=?',(sh['id'],)).fetchone()[0]
            if current+len(items)>TYPES[kind]:raise ValueError(f'Listede {current} kişi var. Aktarım için önce mevcut listeyi temizleyin veya boş bir tarih seçin.')
            students=c.execute('SELECT * FROM students WHERE active=1').fetchall()
            norm=excel_yatakhane.norm
            names={}
            for st in students:names.setdefault((norm(st['name']),norm(st['surname'])),[]).append(st)
            numbers={str(st['school_no']).strip():st for st in students if st['school_no']}
            used=set();errors=[];records=[]
            for item in items:
                first,last,slot,loc,role,row=item[:6]
                identity=item[6] if len(item)>6 else {}
                number=identity.get('school_no','')
                if number:
                    st=numbers.get(number)
                    if not st:errors.append(f'Satır {row}: okul no {number} öğrenci listesinde bulunamadı');continue
                    if norm(st['surname'])!=norm(last) or (identity.get('class') and norm(st['class'])!=norm(identity['class'])):
                        errors.append(f'Satır {row}: okul no {number} ile soyad/sınıf uyuşmuyor. Excel’i yeniden hesaplatıp bilgileri kontrol edin.');continue
                    matches=[st]
                else:matches=names.get((norm(first),norm(last)),[])
                if len(matches)!=1:
                    errors.append(f'Satır {row}: {first} {last} - '+('bulunamadı' if not matches else 'birden fazla eşleşme'));continue
                st=matches[0]
                if st['id'] in used:errors.append(f'Satır {row}: aynı öğrenci iki kez');continue
                used.add(st['id'])
                reason=eligible(c,st,day,kind,slot)
                if reason:errors.append(f'Satır {row}: {first} {last} - {reason}');continue
                if has_conflict(c,st['id'],day,kind):
                    errors.append(f'Satır {row}: {first} {last} - aynı gün bu nöbet grubunda zaten görevli');continue
                records.append((sh['id'],st['id'],slot,loc,role))
                if number and norm(st['name'])!=norm(first):
                    log(c,'Excel okul no eşleştirmesi',f'Satır {row}; okul no {number}; Excel: {first} {last}; kayıt: {st["name"]} {st["surname"]}')
            if errors:raise ValueError('Hiçbir kişi aktarılmadı. Sorunları düzeltin:\n'+'\n'.join(errors[:15]))
            c.executemany('INSERT INTO assignments(sheet_id,student_id,slot,location,role) VALUES(?,?,?,?,?)',records)
            log(c,'Excel nöbet aktarımı',f'{day} {kind} {len(records)} kişi')
        return len(records)

    def bulk_dialog(self):
        w=tk.Toplevel(self);w.title('Toplu nöbet aktarımı');w.geometry('740x560');pencere_yonetimi.modal(w,self)
        ttk.Label(w,text='Excel’den okul numaralarını (veya Ad Soyad) satır satır yapıştırın.\nİsterseniz her satıra: Okul No [TAB] Saat [TAB] Kat/Blok [TAB] Rol yazın.',justify='left').pack(anchor='w',padx=12,pady=10)
        area=tk.Text(w,height=20,wrap='none');area.pack(fill='both',expand=True,padx=12,pady=6)
        def csv_load():
            path=filedialog.askopenfilename(parent=w,filetypes=[('CSV dosyaları','*.csv'),('Tüm dosyalar','*.*')]);
            if not path:return
            for encoding in ('utf-8-sig','cp1254'):
                try: content=Path(path).read_text(encoding=encoding);break
                except UnicodeError:continue
            area.delete('1.0','end');area.insert('1.0',content)
        def submit():
            try:
                lines=[x.strip() for x in area.get('1.0','end').splitlines() if x.strip()]
                n=self.bulk_add(lines)
                messagebox.showinfo('Aktarım tamam',f'{n} öğrenci topluca eklendi.',parent=w);w.destroy();self.refresh_schedule()
            except Exception as e:self.alert(e)
        bar=ttk.Frame(w);bar.pack(fill='x',padx=12,pady=10)
        ttk.Button(bar,text='CSV dosyası seç',command=csv_load).pack(side='left')
        ttk.Button(bar,text='Hepsini kontrol et ve ekle',command=submit).pack(side='right')

    def bulk_add(self,lines):
        if not lines:raise ValueError('Liste boş')
        kind=self.kind.get();day=self.day.get();dt.date.fromisoformat(day)
        with conn() as c:
            sh=self.sheet_id(c,True)
            if sh['status']=='Onaylandı':raise ValueError('Onaylı liste değiştirilemez')
            existing=c.execute('SELECT COUNT(*) FROM assignments WHERE sheet_id=?',(sh['id'],)).fetchone()[0]
            if existing+len(lines)>TYPES[kind]:raise ValueError(f'{kind} için en fazla {TYPES[kind]} kişi; mevcut {existing}, yeni {len(lines)}')
            students=c.execute('SELECT * FROM students WHERE active=1').fetchall()
            lookup={r['school_no']:r for r in students}
            names={}
            for r in students:names.setdefault((r['name']+' '+r['surname']).casefold().strip(),[]).append(r)
            parsed=[];used=set();errors=[]
            for i,line in enumerate(lines,1):
                cols=[x.strip() for x in (line.split('\t') if '\t' in line else line.split(';'))]
                key=cols[0];r=lookup.get(key)
                if not r:
                    matches=names.get(key.casefold(),[])
                    if len(matches)==1:r=matches[0]
                    elif len(matches)>1:errors.append(f'{i}: Aynı ad-soyad birden fazla kişide var; okul no kullanın');continue
                    else:errors.append(f'{i}: Öğrenci bulunamadı: {key}');continue
                if r['id'] in used:errors.append(f'{i}: Listede tekrar var: {key}');continue
                used.add(r['id'])
                slot=cols[1] if len(cols)>1 and cols[1] else TIMES[kind][(existing+i-1)%len(TIMES[kind])]
                loc=cols[2] if len(cols)>2 else ''
                role=cols[3] if len(cols)>3 and cols[3] else ('Yedek' if slot=='Yedek' else 'Asıl')
                if slot not in TIMES[kind]:errors.append(f'{i}: Geçersiz saat/görev: {slot}');continue
                if role not in ('Asıl','Yedek'):errors.append(f'{i}: Geçersiz rol: {role}');continue
                reason=eligible(c,r,day,kind,slot)
                if reason:errors.append(f'{i}: {key}: {reason}');continue
                if has_conflict(c,r['id'],day,kind):errors.append(f'{i}: {key}: Bu gün bu nöbet grubunda zaten görevli');continue
                parsed.append((sh['id'],r['id'],slot,loc,role))
            if errors:raise ValueError('Hiçbir kayıt eklenmedi. Hatalar:\n'+'\n'.join(errors[:18]))
            c.executemany('INSERT INTO assignments(sheet_id,student_id,slot,location,role) VALUES(?,?,?,?,?)',parsed)
            log(c,'Toplu görev aktarımı',f'{day} {kind} {len(parsed)} kişi')
        return len(parsed)

    def auto_fill(self):
        self.plan_auto()

    def planner_settings(self):
        try:
            window=tk.Toplevel(self);window.title('Otomatik nöbet sırası')
            ttk.Label(window,text='İlk boş tarih ve o tarihte geçerli sıra',padding=12).pack(anchor='w')
            with conn() as c:
                for kind in TYPES:
                    day=self.next_open_sheet_day(c,kind,self.day.get().strip())
                    state,anchor,source=nobet_planlayici.queue_state(sys.modules[__name__],c,kind,day)
                    first=nobet_planlayici.first_in_queue(sys.modules[__name__],c,state)
                    last=c.execute("""SELECT s.day,s.status,st.name,st.surname,st.class FROM sheets s
                        JOIN assignments a ON a.sheet_id=s.id JOIN students st ON st.id=a.student_id
                        WHERE s.kind=? AND s.day<? AND a.role='Asıl' AND a.duty_type='Normal'
                        ORDER BY s.day DESC,COALESCE(a.manual_position,a.id) DESC LIMIT 1""",(kind,day)).fetchone()
                    frame=ttk.LabelFrame(window,text=kind,padding=10);frame.pack(fill='x',padx=12,pady=(0,8))
                    if last:
                        ttk.Label(frame,text=f'Son listedeki son asıl · {last["day"]} · {last["name"]} {last["surname"]} ({last["class"]}) · {last["status"]}').pack(anchor='w')
                    name=f'{first["name"]} {first["surname"]} ({first["class"]})' if first else 'Başlangıç seçilmemiş'
                    ttk.Label(frame,text=f'{day} · Sıradaki aday: {name}').pack(anchor='w')
                    ttk.Label(frame,text=f'{source}'+(f' · {anchor}' if anchor else '')).pack(anchor='w')
                    if kind in nobet_planlayici.NIGHT_KINDS:
                        ttk.Label(frame,text='Sınıflar · '+'–'.join((state or {}).get('class_order',CLASSES))).pack(anchor='w')
                        command=lambda k=kind,d=day:self.night_continuation(k,d)
                    else:
                        command=self.gate_initial_student
                    ttk.Button(frame,text='Sırayı düzenle',command=command).pack(anchor='e')
            ttk.Label(window,text='Muafiyet, dinlenme ve sınıf çakışması kuralları son seçimi belirler.',padding=(12,0,12,12)).pack(anchor='w')
            pencere_yonetimi.modal(window,self)
        except Exception as e:
            if 'window' in locals():window.destroy()
            self.alert(e)

    @staticmethod
    def next_open_sheet_day(c,kind,day):
        dt.date.fromisoformat(day)
        kinds=nobet_planlayici.NIGHT_KINDS if kind=='Gece' else (kind,kind)
        occupied={row[0] for row in c.execute("""SELECT s.day FROM sheets s
            WHERE s.kind IN (?,?) AND s.day>=? AND EXISTS(SELECT 1 FROM assignments a WHERE a.sheet_id=s.id)""",(*kinds,day))}
        while day in occupied or (kind=='Yatakhane Giriş' and dt.date.fromisoformat(day).weekday()>=5):
            day=sistem_islemleri.next_preparation_day(day,kind)
        return day

    def gate_initial_student(self):
        try:
            kind='Yatakhane Giriş'
            with conn() as c:
                students=list(c.execute('SELECT * FROM students WHERE active=1 ORDER BY id'))
                current=c.execute('SELECT seed_id FROM planner_settings WHERE kind=?',(kind,)).fetchone()
            if not students:raise ValueError('Önce öğrenci listesini aktarın')
            options={f'{st["school_no"]} — {st["class"]} — {st["name"]} {st["surname"]}':st['id'] for st in students}
            selected=next((key for key,sid in options.items() if current and sid==current['seed_id']),next(iter(options)))
            label='Yatakhane Giriş başlangıcı'
            def save(values):
                with conn() as c:nobet_planlayici.configure(sys.modules[__name__],c,{kind:options[values[0]]})
            self.dialog('Yatakhane Giriş başlangıcı',[label],save,{label:list(options)},{label:selected})
        except Exception as e:self.alert(e)

    def night_continuation(self,kind=None,day=None):
        try:
            kind=kind or self.kind.get()
            if kind not in nobet_planlayici.NIGHT_KINDS:
                raise ValueError('Önce Yatakhane Kat veya Devriye türünü seçin.')
            with conn() as c:
                day=day or self.next_open_sheet_day(c,kind,self.day.get().strip())
                students=c.execute('SELECT * FROM students WHERE active=1 ORDER BY id').fetchall()
                state,_,_=nobet_planlayici.queue_state(sys.modules[__name__],c,kind,day)
                first=nobet_planlayici.first_in_queue(sys.modules[__name__],c,state)
            options={f'{st["class"]} · {st["school_no"]} · {st["name"]} {st["surname"]}':st['id'] for st in students}
            if not options:
                raise ValueError('Önce öğrenci listesini aktarın.')
            state=state or {}
            directions=['Liste sırası','Ters liste sırası']
            labels=['Başlangıç (YYYY-AA-GG)','Sınıf sırası','Sınıf içindeki sıra','Sıradaki ilk öğrenci']
            defaults=dict(zip(labels,(day,'-'.join(state.get('class_order',CLASSES)),
                directions[bool(state.get('reverse_students'))],
                next((label for label,sid in options.items() if first and sid==first['id']),next(iter(options))))))
            def save(values):
                start,order,direction,student=values
                classes=[label.strip().upper().replace('I','İ') for label in re.split(r'[-–—,\s]+',order) if label]
                with conn() as c:
                    c.execute('BEGIN IMMEDIATE')
                    nobet_planlayici.configure_continuation(sys.modules[__name__],c,kind,start,options[student],classes,direction==directions[1])
                messagebox.showinfo('Devam sırası',f'{kind} · {start} tarihinden itibaren devam sırası kaydedildi.',parent=self.dialog_owner())
            self.dialog(f'{kind} · Gece nöbeti devam sırası',labels,save,
                        {labels[2]:directions,labels[3]:list(options)},defaults)
        except Exception as error:self.alert(error)

    def choose_sick_students(self,day,kind):
        with conn() as c:
            students=ogrenci_profili.search(c,'')
        window=tk.Toplevel(self);window.title('Bugün hasta öğrenci var mı?')
        window.geometry('730x500')
        ttk.Label(window,text=f'{day} · {kind}',padding=12).pack(anchor='w')
        ttk.Label(window,text='Bu hazırlamada dışarıda bırakılacak öğrencileri işaretleyin.').pack(anchor='w',padx=12)
        query=tk.StringVar(window)
        entry=ttk.Entry(window,textvariable=query);entry.pack(fill='x',padx=12,pady=8)
        controls=ttk.Frame(window,padding=(12,0,12,6));controls.pack(side='bottom',fill='x')
        status=ttk.Label(window);status.pack(side='bottom',anchor='w',padx=12)
        tree=self.tree(window,['Seç','Okul No','Ad Soyad','Sınıf'],[55,110,400,75])
        selected=set()
        def update_selection_label():
            status.config(text=f'{len(selected)} öğrenci seçildi')
            proceed_button.config(text='Seçilenleri çıkar ve devam et' if selected else 'Hasta yok, devam et')
        labels={str(st['id']):(st['school_no'],f'{st["name"]} {st["surname"]}',st['class'] or '') for st in students}
        searchable={sid:ogrenci_profili.norm(' '.join(values)) for sid,values in labels.items()}
        def refresh(*_):
            self.clear(tree)
            tokens=ogrenci_profili.norm(query.get()).split()
            for sid,values in labels.items():
                if all(token in searchable[sid] for token in tokens):
                    tree.insert('','end',iid=sid,values=('✓' if sid in selected else '☐',*values))
            update_selection_label()
        def toggle(items):
            for sid in items:
                if sid in selected:selected.remove(sid)
                else:selected.add(sid)
                tree.set(sid,'Seç','✓' if sid in selected else '☐')
            update_selection_label()
        def click(event):
            sid=tree.identify_row(event.y)
            if sid and tree.identify_region(event.x,event.y)=='cell' and tree.identify_column(event.x)=='#1':
                toggle((sid,))
                return 'break'
        def toggle_highlighted(_event=None):
            toggle(tree.selection())
            return 'break'
        answer=None
        def proceed():
            nonlocal answer
            answer=tuple(sorted(map(int,selected)))
            window.destroy()
        ttk.Button(controls,text='Vazgeç',command=window.destroy).pack(side='left')
        ttk.Button(controls,text='Seçili satırların işaretini değiştir',command=toggle_highlighted).pack(side='left',padx=6)
        proceed_button=ttk.Button(controls,text='Hasta yok, devam et',style='Primary.TButton',command=proceed)
        proceed_button.pack(side='right')
        query.trace_add('write',refresh)
        tree.bind('<Button-1>',click)
        tree.bind('<space>',toggle_highlighted)
        refresh()
        pencere_yonetimi.modal(window,self)
        entry.focus_set()
        self.wait_window(window)
        return answer

    def plan_auto(self):
        try:
            day=self.day.get().strip();kind=self.kind.get()
            penalty_slots={} if kind in nobet_planlayici.NIGHT_KINDS else None
            with conn() as c:
                if kind in nobet_planlayici.NIGHT_KINDS:
                    devriye_day=self.next_open_sheet_day(c,'Devriye',day)
                    if devriye_day!=day:
                        target=self.next_open_sheet_day(c,'Gece',day)
                        if messagebox.askyesno('Bu gecenin devriye listesi hazır',
                            f'{day} tarihli Devriye listesi kayıtlı. İki tür için de boş gece olan {target} tarihine geçilsin mi?',
                            parent=self.dialog_owner()):
                            self.day.set(target);self.refresh_schedule()
                        return
            excluded=self.choose_sick_students(day,kind)
            if excluded is None:return
            with conn() as c:
                penalty_options=nobet_planlayici.penalty_options(sys.modules[__name__],c,day,kind,excluded)
            if penalty_options:
                penalty_slots=nobet_planlayici.suggest_penalty_slots(penalty_options,penalty_options)
            while True:
                if penalty_options:
                    penalty_slots=self.choose_penalty_slots(day,penalty_options,penalty_slots)
                    if penalty_slots is None:return
                try:
                    with conn() as c:
                        preview=nobet_planlayici.build(sys.modules[__name__],c,day,kind,penalty_slots=penalty_slots,excluded=excluded)
                    break
                except nobet_planlayici.PenaltyPlacementError as error:
                    if not penalty_options:raise
                    self.alert(error)
            candidates={}
            with conn() as c:
                for plan in preview['plans']:
                    for row in plan['rows']:
                        sid=row['student_id']
                        missed_day=plan['retries'].get(sid)
                        if not missed_day:continue
                        previous=c.execute("""SELECT a.slot,a.location,a.id FROM assignments a
                            JOIN sheets s ON s.id=a.sheet_id
                            WHERE s.kind=? AND s.day=? AND a.student_id=? AND a.role='Asıl'""",
                            (plan['kind'],missed_day,sid)).fetchone()
                        candidates[(plan['kind'],sid)]=(row,missed_day,previous)
            choices=None
            if candidates:
                choices=self.choose_retry_students(day,candidates)
                if choices is None:return
                with conn() as c:
                    preview=nobet_planlayici.build(sys.modules[__name__],c,day,kind,choices,penalty_slots,excluded)
            if preview['already_listed']:
                if messagebox.askyesno('Kayıtlı yatakhane listesinde seçili öğrenci var',
                    'Seçtiğiniz öğrencilerden bazıları kayıtlı Yatakhane Kat listesinde. Önce ilgili satırları değiştirmek için listeyi açalım mı?',
                    parent=self.dialog_owner()):
                    self.kind.set('Yatakhane Kat');self.refresh_schedule()
                    items=[item for item in self.atree.get_children()
                           if int(self.atree.item(item,'values')[0]) in preview['already_listed']]
                    self.atree.selection_set(items)
                    if items:self.atree.see(items[0])
                return
            w=tk.Toplevel(self);w.title('Otomatik nöbet önizlemesi');w.geometry('1000x700');pencere_yonetimi.modal(w,self)
            plan_note=('Ceza görevleri seçtiğiniz saatte. Normal görevler, aynı türde az tutulan vardiyalara öncelik verilerek dağıtılır.\n'
                       'Yatakhane önceliklidir. Çakışan devriye cezaları sonraki uygun geceye kalır.\n') if penalty_slots is not None else ''
            ttk.Label(w,text=f'{day}\n{plan_note}Devam sırası otomatik listeyi onayladığınızda ilerler.',wraplength=950).pack(padx=10,pady=10)
            text=tk.Text(w,wrap='word');text.pack(fill='both',expand=True,padx=10)
            if excluded:text.insert('end',f'Bu hazırlamada {len(excluded)} öğrenci dışarıda bırakıldı.\n\n')
            retry_lines=[]
            for plan in preview['plans']:
                for row in plan['rows']:
                    missed_day=plan['retries'].get(row['student_id'])
                    if missed_day:
                        retry_lines.append(f'{plan["kind"]} · {row["name"]} {row["surname"]} · {missed_day} Tutmadı → {row["slot"]}')
            if retry_lines:
                text.insert('end','TEKRAR YAZILACAK ASIL NÖBETÇİLER\n'+'\n'.join(retry_lines)+'\n\n')
            if choices is not None:
                skipped=sum(not selected for selected in choices.values())
                if skipped:text.insert('end',f'Bugün es geçilen asıl nöbetçi: {skipped}\n\n')
                scheduled={(plan['kind'],sid) for plan in preview['plans'] for sid in plan['retries']}
                unavailable=[candidates[key][0] for key,selected in choices.items() if selected and key not in scheduled]
                if unavailable:
                    names=', '.join(row['name']+' '+row['surname'] for row in unavailable)
                    text.insert('end',f'Seçildi ancak bugünkü planlama kuralları nedeniyle önceki görev türüne yazılamadı: {names}\n\n')
            if preview['preserved']:text.insert('end','Mevcut yatakhane listesi aynen korunacak.\n\n')
            for plan in preview['plans']:
                text.insert('end',f'\n{plan["kind"]} — {len(plan["rows"])} öğrenci\n')
                for note in plan['notes']:text.insert('end',note+'\n')
                if plan['kind'] in nobet_planlayici.NIGHT_KINDS:
                    text.insert('end','Geçmiş vardiya adetleri sırasıyla 23–01 / 01–03 / 03–05 / 05–07\n')
                for row in plan['rows']:
                    history=''
                    if 'slot_counts' in row:
                        counts=' / '.join(map(str,row['slot_counts']))
                        history=f' | Geçmiş: {counts} · Son: {row["last_slot"] or "—"}'
                    text.insert('end',f'{row["school_no"]} | {row["name"]} {row["surname"]} | {row["klass"]} | {row["slot"]} | {row["location"]} | {row["role"]} | {row.get("duty_type","Normal")}{history}\n')
            text.configure(state='disabled')
            def save():
                try:
                    count=nobet_planlayici.save(sys.modules[__name__],preview);w.destroy();self.refresh_all()
                    messagebox.showinfo('Taslaklar hazır',f'{count} görev Planlandı olarak eklendi. Nöbetler tamamlandığında durumlarını işaretleyip her listeyi onaylayın. Önizleme hazırlamak sırayı ilerletmez.',parent=self)
                except Exception as e:self.alert(e)
            ttk.Button(w,text='Gösterilen taslakları kaydet',command=save).pack(pady=12)
        except Exception as e:self.alert(e)

    def choose_penalty_slots(self,day,candidates,previous):
        window=tk.Toplevel(self)
        window.title('Ceza nöbeti saatlerini seç')
        window.geometry(f'900x{min(545,self.winfo_screenheight()-100)}')
        window.transient(self)
        ttk.Label(window,text=f'{day} için ceza nöbetine yazılacak öğrencileri seçin.',
                  font=('Segoe UI',12,'bold')).pack(anchor='w',padx=14,pady=(12,4))
        ttk.Label(window,text='01–03 ve 03–05 önerilir. Tikleri ve saatleri değiştirebilirsiniz.').pack(anchor='w',padx=14,pady=(0,4))
        ttk.Label(window,text=f'{len(previous)} / {len(candidates)} başlangıçta seçili. Boş saati olan öğrencileri de işaretleyebilirsiniz.').pack(anchor='w',padx=14,pady=(0,10))
        bulk=ttk.Frame(window);bulk.pack(fill='x',padx=14,pady=(0,8))
        list_frame=ttk.Frame(window)
        list_frame.pack(fill='both',expand=True,padx=14)
        canvas=tk.Canvas(list_frame,highlightthickness=0,background='white')
        scrollbar=ttk.Scrollbar(list_frame,orient='vertical',command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right',fill='y')
        canvas.pack(side='left',fill='both',expand=True)
        body=ttk.Frame(canvas)
        body.columnconfigure(0,weight=1)
        body_id=canvas.create_window((0,0),window=body,anchor='nw')
        body.bind('<Configure>',lambda _:canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',lambda event:canvas.itemconfigure(body_id,width=event.width))
        chosen={}
        selected={}
        for index,(key,(student,count)) in enumerate(candidates.items()):
            name=f"{student['name']} {student['surname']} · {student['school_no']} · {student['class']} sınıfı"
            tick=tk.BooleanVar(master=window,value=key in previous)
            selected[key]=tick
            ttk.Checkbutton(body,text=name,variable=tick).grid(row=index,column=0,sticky='w',padx=6,pady=4)
            last=student.get('last_penalty_slot') or '—'
            ttk.Label(body,text=f'{key[0]} · {count} yazılabilir · Son: {last}').grid(row=index,column=1,sticky='w',padx=8)
            value=tk.StringVar(master=window,value=previous.get(key,''))
            chosen[key]=value
            ttk.Combobox(body,state='readonly',textvariable=value,width=15,
                         values=nobet_planlayici.PENALTY_TIMES).grid(row=index,column=2,padx=6)
        def mark_all(value):
            for tick in selected.values():tick.set(value)
        def suggest():
            proposed=nobet_planlayici.suggest_penalty_slots(candidates,[key for key,tick in selected.items() if tick.get()])
            for key,tick in selected.items():
                if tick.get():chosen[key].set(proposed.get(key,''))
        ttk.Button(bulk,text='Tümünü seç',command=lambda:mark_all(True)).pack(side='left')
        ttk.Button(bulk,text='Seçimi temizle',command=lambda:mark_all(False)).pack(side='left',padx=6)
        ttk.Button(bulk,text='Seçililere otomatik saat ver',command=suggest).pack(side='left')
        answer=None
        def proceed():
            nonlocal answer
            marked=[key for key,tick in selected.items() if tick.get()]
            assigned={key:chosen[key].get() for key in marked if chosen[key].get()}
            assigned=nobet_planlayici.suggest_penalty_slots(candidates,[key for key in marked if key not in assigned],assigned)
            if len(assigned)!=len(marked):
                messagebox.showwarning('Ceza vardiyaları dolu','01–03 ve 03–05 saatlerine tüm seçimler sığmadı. Boş kalanlar için 23–01 veya 05–07 seçin ya da seçimi azaltın.',parent=window)
                for key,slot in assigned.items():chosen[key].set(slot)
                return
            outer=[key for key,slot in assigned.items() if slot not in nobet_planlayici.PENALTY_DEFAULTS]
            if outer:
                details='\n'.join(f"{candidates[key][0]['name']} {candidates[key][0]['surname']} · {key[0]} · {assigned[key]}" for key in outer)
                if not messagebox.askyesno('Erken / geç ceza vardiyası',f'Bu ceza saatlerini onaylıyor musunuz?\n\n{details}',parent=window):return
            answer=assigned
            window.destroy()
        actions=ttk.Frame(window)
        actions.pack(fill='x',padx=14,pady=12)
        ttk.Button(actions,text='Vazgeç',command=window.destroy).pack(side='left')
        ttk.Button(actions,text='Seçimle devam et',style='Primary.TButton',command=proceed).pack(side='right')
        pencere_yonetimi.modal(window,self)
        self.wait_window(window)
        return answer

    def choose_retry_students(self,day,candidates):
        window=tk.Toplevel(self)
        window.title('Bugün nöbete eklenecek asıllar')
        window.geometry(f'850x{min(560,self.winfo_screenheight()-100)}')
        window.transient(self)
        ttk.Label(window,text=f'{day} için nöbete eklenecek asılları seçin.',
                  font=('Segoe UI',12,'bold')).pack(anchor='w',padx=14,pady=(12,4))
        ttk.Label(window,text='İşaretli öğrenciler öncelikli yazılır. İşaretsiz öğrenciler bugün es geçilir.').pack(anchor='w',padx=14,pady=(0,10))
        list_frame=ttk.Frame(window)
        list_frame.pack(fill='both',expand=True,padx=14)
        canvas=tk.Canvas(list_frame,highlightthickness=0,background='white')
        scrollbar=ttk.Scrollbar(list_frame,orient='vertical',command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right',fill='y')
        canvas.pack(side='left',fill='both',expand=True)
        body=ttk.Frame(canvas)
        body_id=canvas.create_window((0,0),window=body,anchor='nw')
        body.bind('<Configure>',lambda event:canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',lambda event:canvas.itemconfigure(body_id,width=event.width))
        selected={}
        def remove(key,assignment_id,frame):
            if not messagebox.askyesno('Tekrar önerisinden çıkar',
                                       'Bu görev için tekrar önerisi kaldırılsın mı? Tutmadı kaydı geçmişte kalır.',
                                       parent=window):return
            try:
                nobet_planlayici.dismiss_missed(sys.modules[__name__],key[0],day,assignment_id)
            except ValueError as error:
                messagebox.showerror('Öneri çıkarılamadı',str(error),parent=window)
                return
            selected[key].set(False)
            frame.destroy()
        for key,(row,missed_day,previous) in candidates.items():
            value=tk.BooleanVar(master=window,value=True)
            selected[key]=value
            name=f'{row["name"]} {row["surname"]} · Okul no {row["school_no"]} · {row["klass"]} sınıfı'
            old_slot=' · '.join(part for part in (previous[:2] if previous else ()) if part)
            frame=ttk.Frame(body)
            frame.pack(fill='x',padx=6,pady=4)
            frame.columnconfigure(0,weight=1)
            ttk.Checkbutton(frame,text=name,variable=value).grid(row=0,column=0,sticky='w')
            ttk.Label(frame,text=f'{missed_day} · {key[0]} · {old_slot or "Saat belirtilmemiş"} · Tutmadı').grid(row=1,column=0,sticky='w',padx=21)
            if previous:
                ttk.Button(frame,text='Tekrar listesinden çıkar',command=lambda k=key,aid=previous['id'],f=frame:remove(k,aid,f)).grid(row=0,column=1,rowspan=2,padx=6)
        answer=None
        def proceed():
            nonlocal answer
            answer={key:value.get() for key,value in selected.items()}
            window.destroy()
        actions=ttk.Frame(window)
        actions.pack(fill='x',padx=14,pady=12)
        ttk.Button(actions,text='Vazgeç',command=window.destroy).pack(side='left')
        ttk.Button(actions,text='Seçimle listeyi hazırla',style='Primary.TButton',command=proceed).pack(side='right')
        pencere_yonetimi.modal(window,self)
        self.wait_window(window)
        return answer

    def gate_class_settings(self):
        try:
            with conn() as c:
                nobet_planlayici.ensure_settings(sys.modules[__name__],c)
                state=json.loads(c.execute("SELECT initial_state FROM planner_settings WHERE kind='Yatakhane Giriş'").fetchone()[0])
                pool=nobet_planlayici.groups(sys.modules[__name__],c)
            choices={};lookup={}
            for cls,students in pool.items():
                if not students:continue
                preferred=students[state['pointers'][cls]%len(students)]['id']
                options={f'{s["school_no"]} — {s["name"]} {s["surname"]}':s['id'] for s in students if s['active']}
                if not options:continue
                label=cls+' sınıfı';lookup[label]=options
                choices[label]=sorted(options,key=lambda key:options[key]!=preferred)
            def save(values):
                seeds={label.split()[0]:lookup[label][value] for label,value in zip(choices,values)}
                with conn() as c:nobet_planlayici.configure_gate_classes(sys.modules[__name__],c,seeds)
            self.dialog('Giriş: her sınıfın sıradaki ilk öğrencisi',list(choices),save,choices)
        except Exception as e:self.alert(e)

    def selected_assignment(self):
        selected=self.atree.selection()
        if not selected:raise ValueError('Önce tablodan bir görev seçin')
        return int(self.atree.item(selected[0],'values')[0])
    def remove_assign(self):
        try:
            selected=self.atree.selection()
            if not selected:raise ValueError('Ctrl veya Shift ile silinecek satırları seçin')
            ids=[int(self.atree.item(x,'values')[0]) for x in selected]
            with conn() as c:
                sh=self.sheet_id(c)
                if not sh or sh['status']=='Onaylandı':raise ValueError('Onaylı çizelge değiştirilemez')
                if not messagebox.askyesno('Toplu silme',f'{len(ids)} görev kaydı silinsin mi?',parent=self.dialog_owner()):return
                c.executemany('DELETE FROM assignments WHERE id=? AND sheet_id=?',[(x,sh['id']) for x in ids])
                log(c,'Toplu görev silindi',f'{sh["day"]} {sh["kind"]} {len(ids)} kayıt')
            self.refresh_schedule()
        except Exception as e:self.alert(e)
    def clear_schedule(self):
        try:
            with conn() as c:
                sh=self.sheet_id(c)
                if not sh:raise ValueError('Bu tarihte liste yok')
                if sh['status']=='Onaylandı':raise ValueError('Önce üstteki Onayı kaldır düğmesiyle listeyi taslağa alın.')
                count=c.execute('SELECT COUNT(*) FROM assignments WHERE sheet_id=?',(sh['id'],)).fetchone()[0]
                if not count:return
                if not messagebox.askyesno('Tüm listeyi sil',f'{sh["day"]} {sh["kind"]}: {count} kişinin tamamı silinecek. Emin misiniz?',parent=self.dialog_owner()):return
                c.execute('DELETE FROM assignments WHERE sheet_id=?',(sh['id'],))
                log(c,'Liste temizlendi',f'{sh["day"]} {sh["kind"]} {count} kişi')
            self.refresh_schedule()
        except Exception as e:self.alert(e)
    def set_result(self,result=None):
        try:
            selected=self.atree.selection()
            if not selected:raise ValueError('Önce tablodan bir veya daha fazla görev seçin')
            ids=[int(self.atree.item(item,'values')[0]) for item in selected]
            def apply_result(value):
                set_assignment_results(ids,value)
                self.refresh_all()
            if result is not None:
                apply_result(result)
                return
            self.dialog(f'{len(ids)} seçili görev — toplu durum değişikliği',['Durum'],lambda v:apply_result(v[0]),{'Durum':['Tuttu','Tutmadı','Planlandı']})
        except Exception as e:self.alert(e)

    def show_missed_notice(self):
        try:
            with conn() as c:
                sheet=self.sheet_id(c)
                rows=c.execute("""
                    SELECT st.school_no,st.name || ' ' || st.surname,st.class,a.slot,a.location
                    FROM assignments a JOIN students st ON st.id=a.student_id
                    WHERE a.sheet_id=? AND a.role='Asıl' AND a.result='Tutmadı' ORDER BY a.id
                """,(sheet['id'],)).fetchall() if sheet else []
            if not rows:
                messagebox.showinfo('Tutmadı bildirimi','Bu listede Tutmadı işaretli asıl nöbetçi bulunmuyor.',parent=self)
                return
            title=f'{sheet["day"]} · {sheet["kind"]} · Nöbetini tutmayan asıl öğrenciler'
            window=tk.Toplevel(self)
            window.title('Tutmadı bildirimi')
            window.geometry('790x420')
            pencere_yonetimi.modal(window,self)
            ttk.Label(window,text=title,font=('Segoe UI',12,'bold')).pack(anchor='w',padx=12,pady=10)
            columns=['Okul No','Ad Soyad','Sınıf','Saat / Görev','Yer']
            tree=self.tree(window,columns,[110,235,70,160,140])
            for row in rows:tree.insert('','end',values=tuple(row))
            actions=ttk.Frame(window)
            actions.pack(fill='x',padx=12,pady=10)
            def print_notice():
                try:
                    def cells(values,tag):
                        return ''.join(f'<{tag}>{html.escape(str(value or ""))}</{tag}>' for value in values)
                    body=''.join('<tr>'+cells(row,'td')+'</tr>' for row in rows)
                    out=DATA/f'tutmadi_{sheet["day"]}_{sheet["kind"].replace(" ","_")}.html'
                    out.write_text('<!doctype html><html lang="tr"><meta charset="utf-8"><title>Tutmadı bildirimi</title>'
                        '<style>body{font:15px Arial;margin:30px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #999;padding:8px;text-align:left}@media print{button{display:none}}</style>'
                        '<button onclick="window.print()">Yazdır / PDF</button><h2>'+html.escape(title)+'</h2><table><thead><tr>'
                        +cells(columns,'th')+'</tr></thead><tbody>'+body+'</tbody></table></html>',encoding='utf-8')
                    webbrowser.open(out.as_uri())
                except Exception as e:self.alert(e)
            ttk.Button(actions,text='Yazdır / PDF',command=print_notice,style='Teal.TButton').pack(side='left')
            ttk.Button(actions,text='Kapat',command=window.destroy).pack(side='right')
        except Exception as e:self.alert(e)
    def confirm_sheet(self):
        try:
            with conn() as c:
                sh=self.sheet_id(c)
                if not sh:raise ValueError('Henüz çizelge oluşturulmadı')
                if sh['status']=='Onaylandı':raise ValueError('Bu çizelge zaten onaylı')
                nobet_manuel.require_complete(c,sh)
                n=c.execute('SELECT COUNT(*) FROM assignments WHERE sheet_id=? AND result="Planlandı" AND role="Asıl"',(sh['id'],)).fetchone()[0]
                if n:raise ValueError(f'{n} asıl görev hâlâ Planlandı durumunda. Önce Tuttu/Tutmadı olarak işaretleyin.')
                if not messagebox.askyesno('Kesin onay','Çizelge kilitlenecek ve Tuttu işaretlenen asıl görevler geçmişe eklenecek. Devam?',parent=self.dialog_owner()):return
                c.execute('BEGIN IMMEDIATE')
                current=self.sheet_id(c)
                if not current or current['id']!=sh['id'] or current['status']=='Onaylandı':
                    raise ValueError('Çizelge değişti. Listeyi yeniden açın.')
                nobet_manuel.require_complete(c,current)
                if c.execute('SELECT 1 FROM assignments WHERE sheet_id=? AND result="Planlandı" AND role="Asıl"',(sh['id'],)).fetchone():
                    raise ValueError('Çizelge değişti. Önce asıl görevlerin sonuçlarını işaretleyin.')
                nobet_planlayici.before_confirm(c,sh)
                ceza_nobet.validate_sheet(sys.modules[__name__],c,sh)
                c.execute('UPDATE sheets SET status="Onaylandı" WHERE id=?',(sh['id'],));log(c,'Çizelge onaylandı',f'{sh["day"]} {sh["kind"]}')
            self.refresh_all()
        except Exception as e:self.alert(e)
    def unconfirm_sheet(self):
        try:
            with conn() as c:
                c.execute('BEGIN IMMEDIATE')
                sh=self.sheet_id(c)
                if not sh:raise ValueError('Bu tarihte liste yok')
                if sh['status']!='Onaylandı':raise ValueError('Liste zaten taslak; silme ve düzenleme yapabilirsiniz.')
                nobet_planlayici.before_unconfirm(c,sh)
                count=c.execute('SELECT COUNT(*) FROM assignments WHERE sheet_id=? AND role="Asıl" AND result="Tuttu"',(sh['id'],)).fetchone()[0]
                if not messagebox.askyesno('Onayı kaldır',f'{sh["day"]} — {sh["kind"]}\n\nListe taslağa alınacak. Bu listeye ait {count} tutulmuş asıl nöbet, yeniden onaylanana kadar geçmişteki nöbet sayılarından çıkarılacak.\n\nÖğrenciler ve görev durumları korunacak. Ardından listeyi düzenleyebilir veya Tüm listeyi sil ile silebilirsiniz.\n\nOnay kaldırılsın mı?',parent=self):return
                c.execute('UPDATE sheets SET status="Taslak" WHERE id=?',(sh['id'],))
                log(c,'Çizelge onayı kaldırıldı',f'{sh["day"]} {sh["kind"]}; {count} tutulmuş asıl nöbet geçmiş toplamından çıkarıldı')
            self.refresh_all()
        except Exception as e:self.alert(e)
    def suggest(self):
        try:
            with conn() as c:
                sh=self.sheet_id(c)
                used=set()
                if sh:used={r[0] for r in c.execute('SELECT student_id FROM assignments WHERE sheet_id=?',(sh['id'],))}
                rows=c.execute('''SELECT s.*,COUNT(CASE WHEN a.result='Tuttu' AND a.role='Asıl' AND sh.status='Onaylandı' THEN 1 END) as cnt FROM students s LEFT JOIN assignments a ON a.student_id=s.id LEFT JOIN sheets sh ON sh.id=a.sheet_id WHERE s.active=1 GROUP BY s.id ORDER BY cnt,s.class,s.name''').fetchall()
                slot=TIMES[self.kind.get()][0];candidates=[]
                for r in rows:
                    if r['id'] in used or eligible(c,r,self.day.get(),self.kind.get(),slot):continue
                    if has_conflict(c,r['id'],self.day.get(),self.kind.get()):continue
                    candidates.append(f"{r['name']} {r['surname']} ({r['class']}) • {r['cnt']} nöbet")
                    if len(candidates)>=12:break
            messagebox.showinfo('Az nöbet tutan uygun öğrenciler', '\n'.join(candidates) if candidates else 'Uygun öğrenci bulunamadı',parent=self.dialog_owner())
        except Exception as e:self.alert(e)
    def print_sheet(self):
        try:
            kind=self.kind.get();day=self.day.get().strip()
            dt.date.fromisoformat(day)
            window=tk.Toplevel(self);window.title('Nöbet çizelgesi çıktısı')
            ttk.Label(window,text=f'{kind} · {day}',padding=12).pack(anchor='w')
            actions=ttk.Frame(window,padding=(12,0,12,12));actions.pack(fill='x')
            def choose(output):
                window.destroy()
                self.export_sheet(kind,day,output)
            ttk.Button(actions,text='Excel olarak kaydet',command=lambda:choose('xlsx'),style='Primary.TButton').pack(side='left',padx=(0,8))
            ttk.Button(actions,text='Yazdır / PDF kaydet',command=lambda:choose('pdf'),style='Teal.TButton').pack(side='left')
            pencere_yonetimi.modal(window,self)
        except Exception as error:self.alert(error)

    def export_sheet(self,kind,day,output):
        try:
            base=dt.date.fromisoformat(day)
            with conn() as c:
                dates=[day,(base+dt.timedelta(days=1)).isoformat()] if kind=='Yatakhane Giriş' else [day]
                by_day={date:[dict(row) for row in c.execute('''SELECT s.name,s.surname,s.class,s.school_no,s.room,a.slot,a.location,a.role,a.result,a.start_at,a.end_at,a.id,a.manual_position FROM assignments a JOIN sheets sh ON sh.id=a.sheet_id JOIN students s ON s.id=a.student_id WHERE sh.day=? AND sh.kind=? ORDER BY a.id''',(date,kind))] for date in dates}
            if not any(by_day.values()):raise ValueError('Önce liste oluşturun.')
            filename=f'cizelge_{day}_{kind.replace(" ","_")}'
            if output=='xlsx':
                path=filedialog.asksaveasfilename(parent=self.dialog_owner(),title='Nöbet çizelgesini Excel olarak kaydet',
                    defaultextension='.xlsx',filetypes=[('Excel çalışma kitabı','*.xlsx')],initialfile=filename+'.xlsx')
                if path:
                    nobet_yazdir.save_excel(path,kind,day,by_day)
                    messagebox.showinfo('Excel kaydedildi',f'Çizelge kaydedildi.\n{path}',parent=self.dialog_owner())
                return
            body=nobet_yazdir.render(kind,day,by_day)
            out=DATA/(filename+'.html');out.write_text(body,encoding='utf-8');webbrowser.open(out.as_uri());messagebox.showinfo('Yazdır','Çizelge tarayıcıda açıldı. Yazdır / PDF kaydet düğmesini kullanın. Üstbilgi/altbilgiyi kapatın.',parent=self.dialog_owner())
        except Exception as e:self.alert(e)

    def history_ui(self):
        f=self.pages['Nöbet Geçmişi'];bar=ttk.Frame(f);bar.pack(fill='x');ttk.Label(bar,text='Öğrenci adı / okul no:').pack(side='left')
        self.hquery=tk.StringVar();ttk.Entry(bar,textvariable=self.hquery,width=35).pack(side='left',padx=8)
        self.hstatus=tk.StringVar(value='Tutulan asıl nöbetler')
        results=ttk.Combobox(bar,textvariable=self.hstatus,state='readonly',width=24,
            values=('Tutulan asıl nöbetler','Tüm sonuçlar','Tutmadı','Yedekler'))
        results.pack(side='left',padx=6)
        results.bind('<<ComboboxSelected>>',lambda _:self.refresh_history())
        ttk.Button(bar,text='Ara',command=self.refresh_history).pack(side='left');ttk.Button(bar,text='CSV dışa aktar',command=self.export_history).pack(side='right')
        self.htree=self.tree(f,['Tarih','Okul No','Ad Soyad','Nöbet','Saat / Görev','Yer','Süre (saat)','Görev Tipi','Rol','Sonuç'],[145,100,175,130,145,90,85,90,75,125]);self.hrows=[]
        self.htree.tag_configure('penalty',foreground='#92400e',background='#ffedd5')
    def refresh_history(self):
        self.clear(self.htree);q='%'+self.hquery.get().strip()+'%';self.hrows=[]
        condition={'Tutulan asıl nöbetler':"AND a.result='Tuttu' AND a.role='Asıl'",
                   'Tüm sonuçlar':'','Tutmadı':"AND a.result='Tutmadı'",'Yedekler':"AND a.role='Yedek'"}[self.hstatus.get()]
        with conn() as c:
            for r in c.execute(f"""SELECT sh.day,s.school_no,s.name,s.surname,sh.kind,a.* FROM assignments a
                JOIN sheets sh ON sh.id=a.sheet_id JOIN students s ON s.id=a.student_id
                WHERE sh.status='Onaylandı' {condition} AND (s.name LIKE ? OR s.surname LIKE ? OR s.school_no LIKE ? OR s.name || ' ' || s.surname LIKE ?)
                ORDER BY COALESCE(NULLIF(a.start_at,''),sh.day) DESC,a.id DESC""",(q,q,q,q)):
                hours=duration(r['kind'],r['slot'],r['start_at'],r['end_at']) if r['role']=='Asıl' and r['result']=='Tuttu' else 0
                values=((r['start_at'] or r['day'])[:10],r['school_no'],r['name']+' '+r['surname'],r['kind'],r['slot'],r['location'],hours,r['duty_type'],r['role'],r['result'])
                self.hrows.append(values)
            if self.hstatus.get() in ('Tutulan asıl nöbetler','Tüm sonuçlar'):
                tokens=ogrenci_profili.norm(self.hquery.get()).split()
                for row in c.execute("""SELECT h.award_day,h.hours,h.count,s.school_no,s.name,s.surname
                    FROM historical_confirmed_penalty_hours h JOIN students s ON s.id=h.student_id
                    ORDER BY h.award_day DESC,s.name,s.surname,h.hours"""):
                    name=f'{row["name"]} {row["surname"]}'
                    needle=ogrenci_profili.norm(f'{row["school_no"]} {name}')
                    if not all(token in needle for token in tokens):continue
                    values=(f'{row["award_day"]} (karar)',row['school_no'],name,'Geçmiş ceza',
                            row['hours'],'',2,'Ceza','Asıl','Tuttu')
                    self.hrows.extend([values]*row['count'])
                self.hrows.sort(key=lambda values:values[0][:10],reverse=True)
        for values in self.hrows:
            self.htree.insert('','end',values=values,tags=('penalty',) if values[7]=='Ceza' else ())
    def export_history(self):
        p=filedialog.asksaveasfilename(defaultextension='.csv',filetypes=[('CSV','*.csv')],parent=self.dialog_owner());
        if not p:return
        with open(p,'w',encoding='utf-8-sig',newline='') as f:
            w=csv.writer(f);w.writerow(['Tarih','Okul No','Ad Soyad','Nöbet','Saat / Görev','Yer','Süre','Görev Tipi','Rol','Sonuç']);w.writerows(self.hrows)
        messagebox.showinfo('Kaydedildi','Geçmiş CSV dosyasına aktarıldı',parent=self.dialog_owner())

    def shift_day(self,offset):
        try:self.day.set((dt.date.fromisoformat(self.day.get())+dt.timedelta(days=offset)).isoformat());self.refresh_schedule()
        except Exception as e:self.alert(e)
    def go_today(self):
        self.day.set(dt.date.today().isoformat());self.refresh_schedule()
    def select_primaries(self):
        self.select_role('Asıl')
    def select_role(self,role):
        items = [item for item in self.atree.get_children() if self.atree.item(item,'values')[6] == role]
        self.atree.selection_set(items)
        if items:
            self.atree.see(items[0])
    def open_pending(self,sheet):
        self.day.set(sheet['day'])
        self.kind.set(sheet['kind'])
        self.refresh_schedule()
    def prepare_next_day(self):
        try:
            target=sistem_islemleri.next_preparation_day(self.day.get(),self.kind.get())
            with conn() as c:waiting=sistem_islemleri.blockers(c,target,self.kind.get())
            if waiting:
                sheet=waiting[0]
                if messagebox.askyesno('Önce bekleyen listeyi tamamlayın',f"{sheet['day']} — {sheet['kind']} otomatik taslağı bekliyor. Önce sonuçlarını işaretleyip onaylayın veya taslağı temizleyin. Bu liste açılsın mı?",parent=self.dialog_owner()):
                    self.day.set(sheet['day']);self.kind.set(sheet['kind']);self.refresh_schedule()
                return
            self.day.set(target);self.refresh_schedule();self.plan_auto()
        except Exception as e:self.alert(e)
    def confirm_current_day(self):
        try:
            with conn() as c:
                lists=[r for r in sistem_islemleri.pending(c) if r['day']==self.day.get()]
                for sheet in lists:
                    try:nobet_manuel.require_complete(c,sheet)
                    except ValueError:
                        self.kind.set(sheet['kind']);self.refresh_schedule()
                        raise
            if not lists:raise ValueError('Bu tarihte onaylanacak dolu taslak yok')
            unfinished=next((s for s in lists if s['undecided']),None)
            if unfinished:
                self.kind.set(unfinished['kind']);self.refresh_schedule()
                raise ValueError(f"{unfinished['kind']}: {unfinished['undecided']} asıl görev sonuç bekliyor. Asılları seçip Tuttu / Tutmadı durumlarını işaretleyin.")
            if not messagebox.askyesno('Günün listelerini onayla',self.day.get()+' tarihinde şu listeler onaylanacak:\n'+ '\n'.join(s['kind'] for s in lists),parent=self.dialog_owner()):return
            sistem_islemleri.confirm_day(sys.modules[__name__],self.day.get());self.refresh_all()
        except Exception as e:self.alert(e)

    def stats_ui(self):
        f=self.pages['İstatistikler'];bar=ttk.Frame(f);bar.pack(fill='x');ttk.Label(bar,text='Başlangıç:').pack(side='left');self.start=tk.StringVar(value=(dt.date.today()-dt.timedelta(days=365)).isoformat());DateField(bar,self.start).pack(side='left',padx=6)
        ttk.Label(bar,text='Bitiş:').pack(side='left');self.end=tk.StringVar(value=dt.date.today().isoformat());DateField(bar,self.end).pack(side='left',padx=6);ttk.Button(bar,text='Hesapla',command=self.refresh_stats).pack(side='left')
        filters=ttk.Frame(f);filters.pack(fill='x',pady=8);self.stats_filters={}
        for index,(kind,slots) in enumerate(nobet_istatistik.SLOTS.items()):
            group=ttk.LabelFrame(filters,text=kind,padding=8);group.grid(row=0,column=index,sticky='nsew',padx=4);filters.columnconfigure(index,weight=1)
            for position,slot in enumerate(slots+[nobet_istatistik.UNKNOWN]):
                variable=tk.BooleanVar(value=True);self.stats_filters[(kind,slot)]=variable
                ttk.Checkbutton(group,text=slot,variable=variable).grid(row=position//2,column=position%2,sticky='w',padx=6,pady=2)
        options=ttk.Frame(f);options.pack(fill='x')
        def select_all(value):
            for variable in self.stats_filters.values():variable.set(value)
            self.refresh_stats()
        ttk.Button(options,text='Tüm tür ve saatleri seç',command=lambda:select_all(True)).pack(side='left',padx=4)
        ttk.Button(options,text='Seçimi temizle',command=lambda:select_all(False)).pack(side='left',padx=4)
        self.stats_only_matching=tk.BooleanVar(value=False)
        ttk.Checkbutton(options,text='Yalnızca eşleşen öğrenciler',variable=self.stats_only_matching).pack(side='left',padx=12)
        self.stats_include_penalties=tk.BooleanVar(value=False)
        ttk.Checkbutton(options,text='Ceza görevlerini de dahil et',variable=self.stats_include_penalties).pack(side='left',padx=8)
        ttk.Label(f,text='İstediğiniz tür ve saatleri işaretleyip Hesapla’ya basın. Sütunlar nöbet adedini, saat sütunları toplam süreyi gösterir.').pack(anchor='w',pady=4)
        self.statslabel=ttk.Label(f,text='',wraplength=1080);self.statslabel.pack(anchor='w',pady=6)
        views=ttk.Notebook(f);views.pack(fill='both',expand=True)
        table=ttk.Frame(views);graph=ttk.Frame(views);views.add(table,text='Tablo');views.add(graph,text='Pasta grafik')
        self.xtree=self.tree(table,['Okul No','Ad Soyad','Sınıf','Toplam nöbet','Toplam saat'],[105,180,65,110,110])
        self.stats_graph=istatistik_grafik.PieView(graph,self.refresh_stats);self.stats_graph.pack(fill='both',expand=True)
    def refresh_stats(self):
        self.clear(self.xtree)
        try:
            with conn() as c:
                selected=[key for key,variable in self.stats_filters.items() if variable.get()]
                columns,rows,count,hours,unknown=nobet_istatistik.calculate(c,self.start.get(),self.end.get(),selected,duration,self.stats_only_matching.get(),self.stats_include_penalties.get())
                if hasattr(self,'stats_graph'):
                    data=istatistik_grafik.aggregate(c,self.start.get(),self.end.get(),selected,duration,self.stats_include_penalties.get(),self.stats_graph.group.get(),self.stats_graph.measure.get())
                    self.stats_graph.set_data(data,self.stats_graph.measure.get())
            self.xtree.configure(columns=columns)
            for i,column in enumerate(columns):
                self.xtree.heading(column,text=column,command=lambda column=column:self.xtree.sort_column(column));self.xtree.column(column,width=180 if i==1 else (70 if i==2 else (110 if i==0 else 165)),minwidth=60,stretch=False,anchor='w' if i<3 else 'center')
            for row in rows:self.xtree.insert('','end',values=row)
            detail=f' • {unknown} kaydın saati belirsiz, süre toplamına eklenmedi' if unknown else ''
            if not selected:detail=' • Tür ve saat seçilmedi'
            self.statslabel.config(text=f'{len(rows)} öğrenci • filtreye uyan {count} onaylı asıl nöbet • {hours:g} saat • {self.start.get()} – {self.end.get()}{detail}')
        except Exception as e:
            self.statslabel.config(text=f'Hata: {e}')
            if hasattr(self,'stats_graph'):self.stats_graph.set_data([],self.stats_graph.measure.get())

    def penalties_ui(self):
        f=self.pages['Cezalılar']
        ttk.Label(f,text='CEZA YÖNETİMİ',font=('Segoe UI',14,'bold')).pack(anchor='w',pady=(0,4))
        ttk.Label(f,text='Ceza kayıtlarını yönetin, eski listeyi arşivleyin ve öğrenci bazında geçmişi inceleyin. Arşivleme geçmişi silmez.').pack(anchor='w',pady=(0,10))
        tabs=ttk.Notebook(f);tabs.pack(fill='both',expand=True)
        listing=ttk.Frame(tabs,padding=10);analysis=ttk.Frame(tabs,padding=10)
        tabs.add(listing,text='Ceza listesi');tabs.add(analysis,text='Ceza istatistikleri')
        filters=ttk.Frame(listing);filters.pack(fill='x',pady=(0,6))
        self.penalty_search=tk.StringVar();self.penalty_view=tk.StringVar(value='Güncel')
        ttk.Label(filters,text='Ad / okul no:').pack(side='left');ttk.Entry(filters,textvariable=self.penalty_search,width=27).pack(side='left',padx=8)
        ttk.Label(filters,text='Görünüm:').pack(side='left');box=ttk.Combobox(filters,textvariable=self.penalty_view,values=['Güncel','Arşiv','Tümü'],state='readonly',width=12);box.pack(side='left',padx=8)
        box.bind('<<ComboboxSelected>>',lambda event:self.refresh_penalties())
        self.penalty_search.trace_add('write',lambda *args:self.refresh_penalties())
        bar=arayuz.ActionBar(listing);bar.pack(fill='x')
        bar.add('Excel’den yapıştır',self.paste_penalties,'Primary.TButton')
        for text,command,style in [('Excel’den cezalıları al',self.import_penalties,'Primary.TButton'),('Ceza ekle / düzenle',self.edit_penalty,'TButton'),('Görev yerini değiştir',self.change_penalty_target,'TButton'),('Aktif / pasif',self.toggle_penalty,'TButton'),('Seçilenleri arşivle',lambda:self.clean_penalties('archive'),'TButton'),('Görünen listeyi arşivle',lambda:self.clean_penalties('archive',True),'TButton'),('Arşivden geri al',lambda:self.clean_penalties('restore'),'TButton'),('Hatalı kaydı sil',lambda:self.clean_penalties('void'),'Danger.TButton')]:bar.add(text,command,style)
        self.penalty_summary=ttk.Label(listing,text='');self.penalty_summary.pack(anchor='w',pady=6)
        self.ptree=self.tree(listing,['ID','Okul No','Ad Soyad','Sınıf','Verilen nöbet','Tuttu','Kalan','Planlı','Görev Yeri','Başlangıç','Durum','Ceza saati'],[55,100,200,65,110,70,70,70,145,115,95,120])
        controls=ttk.Frame(analysis);controls.pack(fill='x',pady=4)
        self.pstat_start=tk.StringVar(value=(dt.date.today()-dt.timedelta(days=365)).isoformat());self.pstat_end=tk.StringVar(value=dt.date.today().isoformat())
        self.pstat_search=tk.StringVar();self.pstat_class=tk.StringVar(value='Tümü')
        ttk.Label(controls,text='Başlangıç:').pack(side='left');DateField(controls,self.pstat_start).pack(side='left',padx=6)
        ttk.Label(controls,text='Bitiş:').pack(side='left');DateField(controls,self.pstat_end).pack(side='left',padx=6)
        ttk.Label(controls,text='Sınıf:').pack(side='left');ttk.Combobox(controls,textvariable=self.pstat_class,values=['Tümü']+CLASSES,state='readonly',width=7).pack(side='left',padx=6)
        searchrow=ttk.Frame(analysis);searchrow.pack(fill='x',pady=4)
        ttk.Label(searchrow,text='Ad / okul no:').pack(side='left');ttk.Entry(searchrow,textvariable=self.pstat_search,width=28).pack(side='left',padx=8)
        ttk.Button(searchrow,text='Hesapla',command=self.refresh_penalty_stats,style='Primary.TButton').pack(side='left',padx=6)
        ttk.Button(searchrow,text='CSV kaydet',command=self.export_penalty_stats).pack(side='left',padx=6)
        ttk.Label(analysis,text='Ceza verildiği tarihe göre filtrelenir. Eski Excel kararları, “Verilen nöbet” ve “Eski ceza (verilen)” sütunlarına girer. “Tuttu” için onaylı ceza görevi gerekir.',wraplength=1120).pack(anchor='w',pady=8)
        self.pstat_summary=ttk.Label(analysis,text='');self.pstat_summary.pack(anchor='w',pady=5)
        self.pstat_columns=['Okul No','Ad Soyad','Sınıf','Kaç kez ceza aldı','Verilen nöbet','Tuttu','Planlı','Güncel kalan','Arşiv kalan','Son ceza tarihi','Eski ceza (verilen)']
        self.pstat_tree=self.tree(analysis,self.pstat_columns,[100,190,65,145,115,85,85,110,110,125,145]);self.pstat_rows=[]
    def refresh_penalties(self):
        self.clear(self.ptree)
        needle=excel_yatakhane.norm(self.penalty_search.get());view=self.penalty_view.get();shown=0;remaining_total=0
        with conn() as c:
            for p in c.execute('SELECT p.*,s.name,s.surname,s.class,s.school_no FROM penalties p JOIN students s ON s.id=p.student_id WHERE p.voided=0 ORDER BY p.archived,p.active DESC,p.start_day,p.id').fetchall():
                if (view=='Güncel' and p['archived']) or (view=='Arşiv' and not p['archived']):continue
                if needle and needle not in excel_yatakhane.norm(f'{p["school_no"]} {p["name"]} {p["surname"]}'):continue
                done,reserved=ceza_nobet.counts(c,p['id'])
                self.ptree.insert('','end',values=(p['id'],p['school_no'],p['name']+' '+p['surname'],p['class'],p['total'],done,max(0,p['total']-done),max(0,reserved-done),p['target'],p['start_day'],'Arşiv' if p['archived'] else ('Bitti' if done>=p['total'] else ('Aktif' if p['active'] else 'Pasif')),p['preferred_slot'] or ceza_nobet.CHOOSE_LATER));shown+=1;remaining_total+=max(0,p['total']-done)
        self.penalty_summary.config(text=f'{shown} kayıt gösteriliyor · {remaining_total} henüz tutulmamış ceza görevi · {view}')
        self.refresh_penalty_stats()
    def clean_penalties(self,mode,visible=False):
        try:
            items=self.ptree.get_children() if visible else self.ptree.selection()
            ids=[int(self.ptree.item(item,'values')[0]) for item in items]
            if not ids:raise ValueError('İşlem yapılacak ceza kaydı seçilmedi')
            action={'archive':'arşivlenecek. Yeni ceza planlamasına alınmaz; geçmiş ve mevcut görevler korunur','restore':'güncel listeye geri alınacak','void':'hatalı kayıt olarak silinecek. Görevlendirmesi olan kayıtlar silinemez'}[mode]
            if not messagebox.askyesno('Ceza listesi işlemi',f'{len(ids)} kayıt {action}. Devam?',parent=self):return
            ceza_nobet.clean(sys.modules[__name__],ids,mode);self.refresh_all()
        except Exception as e:self.alert(e)
    def refresh_penalty_stats(self):
        self.clear(self.pstat_tree)
        self.pstat_rows=[]
        try:
            with conn() as c:self.pstat_rows=ceza_nobet.statistics(c,self.pstat_start.get(),self.pstat_end.get(),self.pstat_search.get(),self.pstat_class.get())
            for row in self.pstat_rows:self.pstat_tree.insert('','end',values=row)
            self.pstat_summary.config(text=f'{len(self.pstat_rows)} öğrenci · {sum(r[3] for r in self.pstat_rows)} ceza kaydı · {sum(r[4] for r in self.pstat_rows)} verilen nöbet · {sum(r[5] for r in self.pstat_rows)} tutulmuş ceza nöbeti')
            return True
        except Exception as e:self.pstat_summary.config(text=f'Hata: {e}');return False
    def export_penalty_stats(self):
        if not self.refresh_penalty_stats():return
        path=filedialog.asksaveasfilename(parent=self,initialfile='ceza_istatistikleri.csv',defaultextension='.csv',filetypes=[('CSV','*.csv')])
        if not path:return
        with open(path,'w',encoding='utf-8-sig',newline='') as handle:
            writer=csv.writer(handle);writer.writerow(self.pstat_columns);writer.writerows(self.pstat_rows)
        messagebox.showinfo('Kaydedildi','Ceza istatistikleri kaydedildi.',parent=self)
    def selected_penalty(self,c):
        selected=self.ptree.selection()
        if len(selected)!=1:raise ValueError('Tek bir ceza kaydı seçin')
        row=c.execute('SELECT * FROM penalties WHERE id=?',(int(self.ptree.item(selected[0],'values')[0]),)).fetchone()
        if not row:raise ValueError('Kayıt bulunamadı')
        return row
    def import_penalties(self):
        path=filedialog.askopenfilename(parent=self,title='CEZALILAR LİSTESİ sayfası olan Excel’i seçin',filetypes=[('Excel','*.xlsx')])
        if not path:return
        def save(values):
            target,start=values;count=ceza_nobet.import_excel(sys.modules[__name__],path,target,start)
            self.refresh_all();messagebox.showinfo('Cezalılar',f'{count} ceza kaydı alındı. Bu işlem henüz nöbet görevlendirmesi yapmaz.',parent=self)
        self.dialog('Ceza listesini al — başlangıç aynı cezanın kayıt tarihidir',['Görev yeri','Ceza başlangıcı (YYYY-AA-GG)'],save,{'Görev yeri':list(ceza_nobet.KINDS)},defaults={'Görev yeri':'Devriye','Ceza başlangıcı (YYYY-AA-GG)':self.day.get()})
    def paste_penalties(self):
        window=tk.Toplevel(self)
        window.title('Excel’den cezalıları yapıştır')
        window.geometry(f'720x{min(570,self.winfo_screenheight()-100)}')
        ttk.Label(window,text='Ad soyad, ayrı ad ve soyad sütunları veya okul numarası yapıştırın.',
                  wraplength=650).pack(anchor='w',padx=14,pady=(14,4))
        ttk.Label(window,text='Saat sütunu varsa o satırın saati kullanılır. Diğer satırlara aşağıdaki seçim uygulanır.',
                  wraplength=650).pack(anchor='w',padx=14,pady=(0,8))
        settings=ttk.Frame(window)
        settings.pack(fill='x',padx=14,pady=6)
        target=tk.StringVar(window,value='Devriye')
        start=tk.StringVar(window,value=self.day.get())
        total=tk.StringVar(window,value='1')
        hour=tk.StringVar(window,value=ceza_nobet.CHOOSE_LATER)
        fields=(
            ('Görev yeri',ttk.Combobox(settings,textvariable=target,values=ceza_nobet.KINDS,state='readonly',width=15)),
            ('Ceza başlangıcı',DateField(settings,start,width=12)),
            ('Kişi başı adet',ttk.Entry(settings,textvariable=total,width=10)),
            ('Ceza saati',ttk.Combobox(settings,textvariable=hour,values=(ceza_nobet.CHOOSE_LATER,*ceza_nobet.HOURS),state='readonly',width=16)),
        )
        for column,(label,widget) in enumerate(fields):
            ttk.Label(settings,text=label).grid(row=0,column=column,sticky='w',padx=5)
            widget.grid(row=1,column=column,sticky='w',padx=5,pady=4)
        text_box=tk.Text(window,wrap='none',height=12,undo=True)
        text_box.pack(fill='both',expand=True,padx=14,pady=6)
        def save():
            try:
                if not total.get().strip().isdecimal() or int(total.get())<1:
                    raise ValueError('Kişi başı ceza adedi pozitif tam sayı olmalı.')
                fallback=None if hour.get()==ceza_nobet.CHOOSE_LATER else hour.get()
                count=ceza_nobet.import_paste(sys.modules[__name__],text_box.get('1.0','end-1c'),
                                              int(total.get()),target.get(),start.get(),fallback)
            except Exception as error:
                messagebox.showerror('Ceza kayıtları alınamadı',str(error),parent=window)
                return
            window.destroy()
            self.refresh_all()
            messagebox.showinfo('Cezalılar',f'{count} ceza kaydı alındı. Görevlendirme için Nöbet Hazırla’yı kullanın.',parent=self)
        actions=ttk.Frame(window)
        actions.pack(fill='x',padx=14,pady=(4,12))
        ttk.Button(actions,text='Vazgeç',command=window.destroy).pack(side='left')
        ttk.Button(actions,text='Cezalıları kaydet',command=save,style='Primary.TButton').pack(side='right')
        pencere_yonetimi.modal(window,self)
        text_box.focus_set()

    def edit_penalty(self):
        try:
            with conn() as c:
                students=list(c.execute('SELECT * FROM students WHERE active=1 ORDER BY class,id'))
                selected=self.selected_penalty(c) if self.ptree.selection() else None
            options={f'{s["school_no"]} — {s["name"]} {s["surname"]}':s['id'] for s in students}
            if not options:raise ValueError('Öğrenci listesi boş')
            defaults={'Ceza nöbeti adedi':'1','Görev yeri':'Devriye','Ceza başlangıcı (YYYY-AA-GG)':self.day.get(),'Ceza saati':ceza_nobet.CHOOSE_LATER}
            if selected:defaults.update({'Öğrenci':next(k for k,v in options.items() if v==selected['student_id']),'Ceza nöbeti adedi':str(selected['total']),'Görev yeri':selected['target'],'Ceza başlangıcı (YYYY-AA-GG)':selected['start_day'],'Ceza saati':selected['preferred_slot'] or ceza_nobet.CHOOSE_LATER})
            def save(values):
                student,total,target,start,hour=values
                with conn() as c:
                    c.execute('BEGIN IMMEDIATE')
                    if selected and (options[student]!=selected['student_id'] or start!=selected['start_day']):raise ValueError('Düzenlemede öğrenci ve ilk kayıt tarihi değiştirilemez; yeni ceza için seçimi kaldırın.')
                    ceza_nobet.record(c,options[student],int(total),target,start,'' if hour==ceza_nobet.CHOOSE_LATER else hour)
                    log(c,'Ceza kaydı',f'{options[student]} {total} {target} {start}')
                self.refresh_all()
            self.dialog('Ceza kaydı',['Öğrenci','Ceza nöbeti adedi','Görev yeri','Ceza başlangıcı (YYYY-AA-GG)','Ceza saati'],save,{'Öğrenci':list(options),'Görev yeri':list(ceza_nobet.KINDS),'Ceza saati':(ceza_nobet.CHOOSE_LATER,*ceza_nobet.HOURS)},defaults)
        except Exception as e:self.alert(e)
    def change_penalty_target(self):
        try:
            selected=self.ptree.selection()
            if not selected:raise ValueError('Bir veya daha fazla ceza kaydı seçin')
            ids=[int(self.ptree.item(item,'values')[0]) for item in selected]
            with conn() as c:p=dict(c.execute('SELECT * FROM penalties WHERE id=?',(ids[0],)).fetchone())
            def save(values):
                with conn() as c:c.executemany('UPDATE penalties SET target=? WHERE id=?',[(values[0],pid) for pid in ids]);log(c,'Ceza görev yeri',f'{len(ids)} kayıt {values[0]}')
                self.refresh_all()
            self.dialog('Gelecek ceza görevlerinin yeri',['Görev yeri'],save,{'Görev yeri':list(ceza_nobet.KINDS)},{'Görev yeri':p['target']})
        except Exception as e:self.alert(e)
    def toggle_penalty(self):
        try:
            with conn() as c:
                p=self.selected_penalty(c)
                if not messagebox.askyesno('Ceza kaydı','Aktif/pasif durumu değiştirilsin mi? Mevcut nöbetler silinmez; yalnızca yeni ceza planlamasına alınması değişir.',parent=self):return
                c.execute('UPDATE penalties SET active=1-active WHERE id=?',(p['id'],));log(c,'Ceza aktif/pasif',str(p['id']))
            self.refresh_all()
        except Exception as e:self.alert(e)

    def exempt_ui(self):
        f=self.pages['Muafiyetler'];ttk.Label(f,text='Muafiyet başlangıç ve bitiş tarihleri dahil edilir. Bitiş boş bırakılırsa süresizdir.').pack(anchor='w')
        self.etree=self.tree(f,['ID','Okul No','Ad Soyad','Görev / Sebep','Başlangıç','Bitiş'],[55,110,180,230,110,110])
        b=ttk.Frame(f);b.pack(fill='x');ttk.Button(b,text='Muafiyet ekle',command=self.add_exempt).pack(side='left');ttk.Button(b,text='Seçileni sil',command=self.delete_exempt).pack(side='left',padx=8)
    def refresh_exempt(self):
        self.clear(self.etree)
        with conn() as c:
            for r in c.execute('SELECT e.*,s.school_no,s.name,s.surname FROM exemptions e JOIN students s ON s.id=e.student_id ORDER BY e.id DESC'):
                self.etree.insert('','end',values=(r['id'],r['school_no'],r['name']+' '+r['surname'],r['reason'],r['start'],r['end'] or 'Süresiz'))
    def add_exempt(self):
        def save(v):
            school,reason,start,end=v;dt.date.fromisoformat(start)
            if end:dt.date.fromisoformat(end)
            with conn() as c:
                s=c.execute('SELECT id FROM students WHERE school_no=?',(school,)).fetchone()
                if not s:raise ValueError('Okul numarası bulunamadı')
                c.execute('INSERT INTO exemptions(student_id,reason,start,end) VALUES(?,?,?,?)',(s['id'],reason,start,end or None));log(c,'Muafiyet eklendi',school)
            self.refresh_exempt()
        self.dialog('Muafiyet ekle',['Okul No','Görev / Sebep','Başlangıç (YYYY-AA-GG)','Bitiş (boş=süresiz)'],save)
    def delete_exempt(self):
        try:
            selection=self.etree.selection()
            if not selection:raise ValueError('Muafiyet seçin')
            eid=int(self.etree.item(selection[0],'values')[0])
            if not messagebox.askyesno('Onay','Muafiyeti kaldırmak istiyor musunuz?',parent=self.dialog_owner()):return
            with conn() as c:c.execute('DELETE FROM exemptions WHERE id=?',(eid,));log(c,'Muafiyet silindi',str(eid))
            self.refresh_exempt()
        except Exception as e:self.alert(e)

    def rules_ui(self):
        f=self.pages['Sınıf Yasakları'];ttk.Label(f,text='Pazartesi–Perşembe seçilen 4 sınıfa GİRİŞ nöbeti yazılamaz. Cuma kuralı sabittir; hafta sonu giriş nöbeti yoktur.').pack(anchor='w',pady=8)
        self.rulevars={}
        for day,name in enumerate(['Pazartesi','Salı','Çarşamba','Perşembe'],1):
            row=ttk.Frame(f);row.pack(anchor='w',pady=8);ttk.Label(row,text=name,width=15).pack(side='left')
            vars=[tk.StringVar() for _ in range(4)];self.rulevars[day]=vars
            for v in vars:ttk.Combobox(row,textvariable=v,values=CLASSES,state='readonly',width=5).pack(side='left',padx=6)
        ttk.Button(f,text='Sınıf yasaklarını kaydet',command=self.save_rules).pack(anchor='w',pady=15)
        ttk.Label(f,text='Yalnızca cuma: 08–10 → A, B, C, D, E, F, G, H | 10–12 / 12–14 / 14–16 / 16–17 → İ, K, L, M, N, P, R, S',wraplength=1000).pack(anchor='w')
    def refresh_rules(self):
        with conn() as c:
            for day,vars in self.rulevars.items():
                classes=[r[0] for r in c.execute('SELECT class FROM restrictions WHERE weekday=? ORDER BY rowid',(day,))]
                for v,x in zip(vars,classes):v.set(x)
    def save_rules(self):
        try:
            with conn() as c:
                for day,vs in self.rulevars.items():
                    classes=[v.get() for v in vs]
                    if len(set(classes))!=4 or any(x not in CLASSES for x in classes):raise ValueError('Her gün 4 farklı sınıf seçin')
                    c.execute('DELETE FROM restrictions WHERE weekday=?',(day,))
                    c.executemany('INSERT INTO restrictions(weekday,class) VALUES(?,?)',[(day,x) for x in classes])
                log(c,'Sınıf yasakları güncellendi','Pazartesi–Perşembe')
            messagebox.showinfo('Kaydedildi','Yasaklı sınıflar güncellendi',parent=self.dialog_owner())
        except Exception as e:self.alert(e)
    def backup_ui(self):
        f=self.pages['Yedekleme'];ttk.Label(f,text='Tüm öğrenci, muafiyet, nöbet ve geçmiş kayıtları veri/nobet.db dosyasında saklanır.',wraplength=700).pack(anchor='w',pady=10)
        ttk.Button(f,text='Veritabanını yedekle (.db)',command=self.backup).pack(anchor='w',pady=12)
        ttk.Label(f,text='Yedeği harici diskte veya güvenli bir klasörde saklayın. Programın veri klasörünü silmeyin.').pack(anchor='w')
        ttk.Label(f,text='Gizlilik: Veritabanına T.C. kimlik numarası ve telefon numarası alınmaz.').pack(anchor='w',pady=15)
        reset=ttk.LabelFrame(f,text='Geçmişi sıfırla',padding=16);reset.pack(fill='x',pady=20)
        ttk.Label(reset,text='Sıfırlanacak bilgileri, öğrenciyi ve nöbet geçmişi tarih aralığını kendiniz seçin. Önce otomatik yedek alınır; öğrenci adları ve okul numaraları silinmez.',wraplength=850).pack(anchor='w')
        ttk.Button(reset,text='Seçerek sıfırla',style='Danger.TButton',command=self.reset_history_dialog).pack(anchor='w',pady=12)
    def reset_history_dialog(self):
        win=tk.Toplevel(self);win.title('Seçerek sıfırla');pencere_yonetimi.modal(win,self);win.geometry(f'850x{min(760,self.winfo_screenheight()-100)}')
        container=ttk.Frame(win);container.pack(fill='both',expand=True)
        canvas=tk.Canvas(container,highlightthickness=0,bg='white');scroll=ttk.Scrollbar(container,orient='vertical',command=canvas.yview);scroll.pack(side='right',fill='y');canvas.configure(yscrollcommand=scroll.set);canvas.pack(side='left',fill='both',expand=True)
        body=ttk.Frame(canvas,padding=18);body_id=canvas.create_window((0,0),window=body,anchor='nw')
        body.bind('<Configure>',lambda event:canvas.configure(scrollregion=canvas.bbox('all')));canvas.bind('<Configure>',lambda event:canvas.itemconfigure(body_id,width=event.width))
        ttk.Label(body,text='Yalnızca işaretlediğiniz bilgiler sıfırlanır. Önce otomatik yedek alınır.',font=('Segoe UI',11,'bold')).pack(anchor='w',pady=6)
        with conn() as c:students=list(c.execute('SELECT id,school_no,name,surname FROM students ORDER BY class,name,surname'))
        students_map={'Tüm öğrenciler':None,**{f'{s["school_no"]} — {s["name"]} {s["surname"]} [#{s["id"]}]':s['id'] for s in students}}
        target=tk.StringVar(value='Tüm öğrenciler')
        ttk.Label(body,text='Öğrenci kapsamı (genel olarak belirtilen işlemler hariç):').pack(anchor='w')
        ttk.Combobox(body,textvariable=target,values=list(students_map),state='readonly',width=70).pack(anchor='w',pady=6)
        choices=ttk.LabelFrame(body,text='Sıfırlanacak bilgiler',padding=8);choices.pack(fill='x',pady=7)
        options={}
        for key,label in sistem_islemleri.RESET_LABELS.items():
            options[key]=tk.BooleanVar(value=False);ttk.Checkbutton(choices,text=label,variable=options[key]).pack(anchor='w',pady=2)
        profile_fields={key:tk.BooleanVar(value=True) for key in sistem_islemleri.PROFILE_FIELDS}
        def choose_fields():
            popup=tk.Toplevel(win);popup.title('Sıfırlanacak ek öğrenci alanları');pencere_yonetimi.modal(popup,win)
            for key,label in sistem_islemleri.PROFILE_FIELDS.items():ttk.Checkbutton(popup,text=label,variable=profile_fields[key]).pack(anchor='w',padx=20,pady=6)
            def done():popup.destroy()
            popup.protocol('WM_DELETE_WINDOW',done);ttk.Button(popup,text='Tamam',command=done).pack(pady=12)
        ttk.Button(choices,text='Ek bilgiler: sıfırlanacak alanları seç',command=choose_fields,style='Import.TButton').pack(anchor='w',pady=4)
        period=ttk.LabelFrame(body,text='Yalnızca nöbet listeleri / geçmişi için filtre',padding=8);period.pack(fill='x')
        dates=ttk.Frame(period);dates.pack(fill='x');start=tk.StringVar();end=tk.StringVar()
        ttk.Label(dates,text='Başlangıç (boş=tümü)').pack(side='left');DateField(dates,start,allow_empty=True).pack(side='left',padx=6)
        ttk.Label(dates,text='Bitiş (boş=tümü)').pack(side='left');DateField(dates,end,allow_empty=True).pack(side='left',padx=6)
        kinds={kind:tk.BooleanVar(value=True) for kind in TYPES};row=ttk.Frame(period);row.pack(fill='x',pady=5)
        for kind,v in kinds.items():ttk.Checkbutton(row,text=kind,variable=v).pack(side='left',padx=6)
        ttk.Label(body,text='Tarih ve tür filtresi yalnızca nöbet geçmişine uygulanır. Sıralar, sınıf yasakları ve işlem günlüğü geneldir. Öğrenci temel kimlik bilgileri korunur.',wraplength=790).pack(anchor='w',pady=8)
        value=tk.StringVar();entryrow=ttk.Frame(body);entryrow.pack(fill='x');ttk.Label(entryrow,text='Onay için SIFIRLA yazın:').pack(side='left');ttk.Entry(entryrow,textvariable=value,width=25).pack(side='left',padx=8)
        def plan():
            selected={k:v.get() for k,v in options.items()};selected_kinds=[k for k,v in kinds.items() if v.get()]
            if selected['history'] and not selected_kinds:raise ValueError('Nöbet geçmişi için en az bir tür seçin')
            with conn() as c:result=sistem_islemleri.reset_plan(c,selected,students_map[target.get()],start.get(),end.get(),selected_kinds,[k for k,v in profile_fields.items() if v.get()])
            return selected,selected_kinds,result
        def description(result):
            text=target.get()+'\n'+'\n'.join(sistem_islemleri.RESET_LABELS[k]+f': {v} kayıt' for k,v in result['counts'].items())
            if 'history' in result['chosen']:text+='\nGeçmiş: '+(start.get() or 'başlangıçtan')+' – '+(end.get() or 'bugüne kadar / tüm tarihler')+'; '+', '.join(k for k,v in kinds.items() if v.get())
            if 'profiles' in result['chosen']:text+='\nEk bilgi alanları: '+', '.join(sistem_islemleri.PROFILE_FIELDS[k] for k in result['profile_fields'])
            return text
        def preview():
            try:messagebox.showinfo('Sıfırlama önizlemesi',description(plan()[2])+'\n\nHenüz hiçbir kayıt değiştirilmedi.',parent=win)
            except Exception as e:messagebox.showerror('Sıfırlama',str(e),parent=win)
        def perform():
            if value.get().strip()!='SIFIRLA':messagebox.showwarning('Doğrulama','Devam etmek için SIFIRLA yazın.',parent=win);return
            try:
                selected,selected_kinds,result=plan()
                if not messagebox.askyesno('Seçili bilgileri sıfırla',description(result)+'\n\nOtomatik yedek alınıp bu bilgiler sıfırlansın mı?',parent=win):return
                backup,counts=sistem_islemleri.reset_selected(sys.modules[__name__],selected,students_map[target.get()],start.get(),end.get(),selected_kinds,result['profile_fields']);win.destroy();self.refresh_all()
                for child in self.winfo_children():
                    if isinstance(child,ogrenci_profili.Profile):child.destroy()
                messagebox.showinfo('Sıfırlandı','Seçili bilgiler sıfırlandı. Önceki kayıtların yedeği:\n'+str(backup),parent=self.dialog_owner())
            except Exception as e:messagebox.showerror('Sıfırlama',str(e),parent=win)
        buttons=arayuz.ActionBar(win);buttons.pack(fill='x',padx=18,pady=8)
        buttons.add('İptal',win.destroy);buttons.add('Önizle',preview);buttons.add('Yedeği al ve sıfırla',perform,'Danger.TButton')
    def backup(self):
        dest=filedialog.asksaveasfilename(defaultextension='.db',initialfile='nobet_yedek_'+dt.date.today().isoformat()+'.db',filetypes=[('SQLite','*.db')],parent=self.dialog_owner());
        if not dest:return
        with conn() as source,sqlite3.connect(dest) as target:source.backup(target)
        messagebox.showinfo('Tamam','Yedek kaydedildi',parent=self.dialog_owner())

def run():
    with pencere_yonetimi.single_instance(DB) as start:
        if start:
            setup()
            window=App()
            pencere_yonetimi.register(window,DB)
            window.mainloop()


if __name__=='__main__':
    run()
