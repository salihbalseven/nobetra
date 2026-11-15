"""Student profiles; stable student IDs preserve linked duty history."""
import datetime as dt
from io import BytesIO
import unicodedata
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import istatistik_grafik as charts
import arayuz
from PIL import Image, ImageOps, ImageTk
from tarih_secici import DateField

CONTACT_FIELDS=[('phone','Öğrenci telefonu'),('mother_phone','Anne telefonu'),('father_phone','Baba telefonu'),('address','İkamet adresi')]
FIELDS=[('registry','Sicil'),('bed','Yatak no'),('university','Mezun olduğu üniversite'),('department','Bölüm'),('mother','Anne adı soyadı'),('father','Baba adı soyadı'),('hometown','Memleket'),('official_duty','Görevi'),('notes','Notlar')]+CONTACT_FIELDS

def setup(c):
    c.execute('''CREATE TABLE IF NOT EXISTS student_profiles(student_id INTEGER PRIMARY KEY REFERENCES students(id),
        registry TEXT DEFAULT '',bed TEXT DEFAULT '',university TEXT DEFAULT '',department TEXT DEFAULT '',
        mother TEXT DEFAULT '',father TEXT DEFAULT '',hometown TEXT DEFAULT '',official_duty TEXT DEFAULT '',notes TEXT DEFAULT '',photo BLOB,photo_name TEXT DEFAULT '')''')
    columns={row['name'] for row in c.execute('PRAGMA table_info(student_profiles)')}
    for key,_ in CONTACT_FIELDS:
        if key not in columns:
            c.execute(f"ALTER TABLE student_profiles ADD COLUMN {key} TEXT NOT NULL DEFAULT ''")

def norm(text):
    text=str(text or '').translate(str.maketrans({'ı':'i','İ':'i'})).casefold()
    return ''.join(x for x in unicodedata.normalize('NFKD',text) if not unicodedata.combining(x))

def search(c,query,include_inactive=False):
    tokens=norm(query).split()
    rows=c.execute('''SELECT s.*,COALESCE(p.registry,'') AS registry FROM students s
        LEFT JOIN student_profiles p ON p.student_id=s.id ORDER BY s.id''')
    return [r for r in rows if (include_inactive or r['active']) and all(token in norm(' '.join(str(r[k] or '') for k in ('name','surname','school_no','class','room','registry'))) for token in tokens)]

def save(app,sid,values,photo=None,photo_name=''):
    values={k:str(v or '').strip() for k,v in values.items()}
    if not all(values.get(k) for k in ('school_no','name','surname')):raise ValueError('Okul no, ad ve soyad zorunludur')
    if values.get('class') not in app.CLASSES:raise ValueError('Geçerli bir sınıf seçin')
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        if sid:
            if not c.execute('SELECT 1 FROM students WHERE id=?',(sid,)).fetchone():raise ValueError('Öğrenci bulunamadı')
            c.execute('UPDATE students SET school_no=?,name=?,surname=?,class=?,room=? WHERE id=?',tuple(values.get(k,'') for k in ('school_no','name','surname','class','room'))+(sid,))
        else:
            sid=c.execute('INSERT INTO students(school_no,name,surname,class,room) VALUES(?,?,?,?,?)',tuple(values.get(k,'') for k in ('school_no','name','surname','class','room'))).lastrowid
        c.execute('INSERT OR IGNORE INTO student_profiles(student_id) VALUES(?)',(sid,))
        c.execute('UPDATE student_profiles SET '+','.join(k+'=?' for k,label in FIELDS)+',photo=?,photo_name=? WHERE student_id=?',tuple(values.get(k,'') for k,label in FIELDS)+(photo,photo_name,sid))
        app.log(c,'Öğrenci profili kaydedildi',str(sid))
    return sid

def set_active(app,sid,active):
    with app.conn() as c:
        c.execute('UPDATE students SET active=? WHERE id=?',(int(active),sid));app.log(c,'Öğrenci aktifliği',f'{sid}: {active}')

def save_exemption(app,sid,eid,reason,start,end):
    reason=reason.strip();dt.date.fromisoformat(start)
    if not reason:raise ValueError('Görev / muafiyet sebebi yazın')
    if end and dt.date.fromisoformat(end)<dt.date.fromisoformat(start):raise ValueError('Bitiş başlangıçtan önce olamaz')
    with app.conn() as c:
        if eid:c.execute('UPDATE exemptions SET reason=?,start=?,end=? WHERE id=? AND student_id=?',(reason,start,end or None,eid,sid))
        else:c.execute('INSERT INTO exemptions(student_id,reason,start,end) VALUES(?,?,?,?)',(sid,reason,start,end or None))
        app.log(c,'Profil muafiyet/görev kaydı',str(sid))

def summary(app,c,sid):
    rows=[dict(r) for r in c.execute('''SELECT a.*,s.kind,s.day,s.status FROM assignments a JOIN sheets s ON s.id=a.sheet_id
        WHERE a.student_id=? ORDER BY s.day DESC,a.id DESC''',(sid,))]
    totals={};slots={}
    for r in rows:
        if r['status']!='Onaylandı' or r['role']!='Asıl' or r['result']!='Tuttu':continue
        key=charts.stats.slot_key(r);hours=0 if key==charts.stats.UNKNOWN else app.duration(r['kind'],r['slot'],r['start_at'],r['end_at'])
        totals.setdefault(r['kind'],[0,0]);totals[r['kind']][0]+=1;totals[r['kind']][1]+=hours
        label=r['kind']+' '+key;slots.setdefault(label,[0,0]);slots[label][0]+=1;slots[label][1]+=hours
    return rows,totals,slots

class Profile(tk.Toplevel):
    def __init__(self,owner,app,sid=None):
        super().__init__(owner);self.owner=owner;self.app=app;self.sid=sid;self.transient(owner)
        self.title('Öğrenci bilgileri');self.geometry('940x730');self.minsize(800,600)
        with app.conn() as c:
            student=dict(c.execute('SELECT * FROM students WHERE id=?',(sid,)).fetchone()) if sid else {}
            row=c.execute('SELECT * FROM student_profiles WHERE student_id=?',(sid,)).fetchone();profile=dict(row) if row else {}
        self.photo=profile.get('photo');self.photo_name=profile.get('photo_name','');self.image_ref=None
        heading=ttk.Frame(self,padding=12);heading.pack(fill='x')
        self.heading=ttk.Label(heading,text=(student.get('name','Yeni')+' '+student.get('surname','öğrenci')),font=('Segoe UI',16,'bold'));self.heading.pack(side='left')
        ttk.Button(heading,text='Bilgileri kaydet',command=self.commit,style='Success.TButton').pack(side='right')
        self.tabs=ttk.Notebook(self);self.tabs.pack(fill='both',expand=True,padx=12,pady=(0,12))
        self.pages={}
        for label in ('Bilgiler ve fotoğraf','İletişim','Görev ve muafiyet','Cezalar','Nöbet geçmişi ve grafik'):
            page=ttk.Frame(self.tabs,padding=12);self.tabs.add(page,text=label);self.pages[label]=page
        f=self.pages['Bilgiler ve fotoğraf'];form=ttk.Frame(f);form.pack(side='left',fill='both',expand=True)
        self.values={}
        fields=[('school_no','Okul no'),('name','Adı'),('surname','Soyadı'),('class','Sınıf'),('room','Koğuş')]+FIELDS[:-len(CONTACT_FIELDS)]
        for index,(key,label) in enumerate(fields):
            ttk.Label(form,text=label).grid(row=index,column=0,sticky='w',pady=4)
            v=tk.StringVar(value=student.get(key,profile.get(key,'')));self.values[key]=v
            w=ttk.Combobox(form,textvariable=v,values=app.CLASSES,state='readonly') if key=='class' else ttk.Entry(form,textvariable=v)
            w.grid(row=index,column=1,sticky='ew',padx=12,pady=4)
        form.columnconfigure(1,weight=1)
        contact=self.pages['İletişim']
        contact.columnconfigure(1,weight=1)
        for index,(key,label) in enumerate(CONTACT_FIELDS):
            ttk.Label(contact,text=label).grid(row=index,column=0,sticky='w',pady=7)
            value=tk.StringVar(value=profile.get(key,''))
            self.values[key]=value
            ttk.Entry(contact,textvariable=value).grid(row=index,column=1,sticky='ew',padx=12,pady=7)
        pictures=ttk.Frame(f,padding=12);pictures.pack(side='right',fill='y')
        self.photo_label=ttk.Label(pictures,text='Fotoğraf yok',anchor='center',width=23);self.photo_label.pack(pady=10)
        ttk.Button(pictures,text='Fotoğraf ekle / değiştir',command=self.pick_photo,style='Import.TButton').pack(pady=6)
        ttk.Button(pictures,text='Fotoğrafı kaldır',command=self.remove_photo).pack(pady=6)
        ttk.Label(pictures,text='Fotoğraf veritabanında saklanır;\nyedeklemeye dahil edilir.\nDeğişiklikler Kaydet ile uygulanır.',wraplength=200).pack(pady=12)
        ttk.Label(pictures,text='Görevi alanı bilgi içindir.\nNöbet muafiyeti için yan sekmede\nmuafiyet tanımlayın.',wraplength=210).pack(pady=12)
        self.show_photo();self.build_exempt();self.build_penalties();self.build_history();self.refresh_related()
    def show_photo(self):
        self.photo_label.configure(image='',text='Fotoğraf yok',width=23);self.image_ref=None
        if not self.photo:return
        try:
            with Image.open(BytesIO(self.photo)) as source:
                image=ImageOps.exif_transpose(source)
                image.thumbnail((180,180))
                self.image_ref=ImageTk.PhotoImage(image.convert('RGB'),master=self)
            self.photo_label.configure(image=self.image_ref,text='',width=0)
        except (OSError, ValueError, Image.DecompressionBombError, tk.TclError):
            self.photo_label.configure(text='Fotoğraf önizlenemedi\n'+self.photo_name)
    def pick_photo(self):
        path=filedialog.askopenfilename(parent=self,filetypes=[('Fotoğraflar','*.png *.gif *.jpg *.jpeg'),('Tüm dosyalar','*.*')])
        if not path:return
        try:
            content=Path(path).read_bytes()
            if len(content)>10*1024*1024:raise ValueError('Fotoğraf en fazla 10 MB olabilir')
            with Image.open(BytesIO(content)) as image:
                if image.format not in ('PNG','GIF','JPEG'):raise ValueError('PNG, GIF veya JPEG fotoğraf seçin')
                image.verify()
            self.photo=content;self.photo_name=Path(path).name;self.show_photo()
        except Exception as e:messagebox.showerror('Fotoğraf',str(e),parent=self)
    def remove_photo(self):self.photo=None;self.photo_name='';self.show_photo()
    def commit(self):
        try:
            self.sid=save(self.app,self.sid,{k:v.get() for k,v in self.values.items()},self.photo,self.photo_name)
            self.heading.config(text=self.values['name'].get()+' '+self.values['surname'].get());self.owner.refresh_all();self.refresh_related()
            messagebox.showinfo('Kaydedildi','Öğrenci bilgileri kaydedildi.',parent=self)
        except Exception as e:messagebox.showerror('Kaydedilemedi',str(e),parent=self)
    def require_saved(self):
        if not self.sid:raise ValueError('Önce öğrenci bilgilerini kaydedin')
    def build_exempt(self):
        f=self.pages['Görev ve muafiyet'];self.duty_label=ttk.Label(f,text='',wraplength=830);self.duty_label.pack(anchor='w',pady=6)
        ttk.Label(f,text='Buradaki görev / sebep kayıtları belirtilen tarihlerde normal nöbetten muafiyet sağlar.').pack(anchor='w')
        self.etree=self.owner.tree(f,['ID','Görev / Sebep','Başlangıç','Bitiş'],[55,360,140,140])
        bar=arayuz.ActionBar(f);bar.pack(fill='x');bar.add('Muafiyet / görev ekle',lambda:self.edit_exemption(False));bar.add('Seçileni düzenle',lambda:self.edit_exemption(True));bar.add('Seçileni kaldır',self.delete_exemption,'Danger.TButton')
    def edit_exemption(self,edit):
        try:
            self.require_saved();row=None
            if edit:
                selection=self.etree.selection()
                if not selection:raise ValueError('Düzenlenecek muafiyeti seçin')
                row=self.etree.item(selection[0],'values')
            def submit(v):
                save_exemption(self.app,self.sid,int(row[0]) if row else None,*v);self.owner.refresh_all();self.refresh_related()
            labels=['Görev / Sebep','Başlangıç (YYYY-AA-GG)','Bitiş (YYYY-AA-GG, boş=süresiz)']
            self.owner.dialog('Öğrencinin muafiyeti',labels,submit,defaults=dict(zip(labels,[row[1] if row else '',row[2] if row else self.owner.day.get(),row[3] if row and row[3]!='Süresiz' else ''])))
        except Exception as e:messagebox.showerror('Muafiyet',str(e),parent=self)
    def delete_exemption(self):
        selected=self.etree.selection()
        if not selected:return
        if not messagebox.askyesno('Muafiyeti kaldır','Seçili muafiyet kaldırılsın mı?',parent=self):return
        with self.app.conn() as c:
            c.execute('DELETE FROM exemptions WHERE id=? AND student_id=?',(int(self.etree.item(selected[0],'values')[0]),self.sid));self.app.log(c,'Profil muafiyeti kaldırıldı',str(self.sid))
        self.owner.refresh_all();self.refresh_related()
    def build_penalties(self):
        f=self.pages['Cezalar'];self.penalty_label=ttk.Label(f,text='');self.penalty_label.pack(anchor='w')
        self.ptree=self.owner.tree(f,['ID','Verilen nöbet','Tuttu','Kalan','Yer','Başlangıç','Durum'],[50,105,65,65,150,110,110])
        bar=arayuz.ActionBar(f);bar.pack(fill='x');bar.add('Ceza ekle',lambda:self.edit_penalty(False));bar.add('Cezayı / görev yerini düzenle',lambda:self.edit_penalty(True));bar.add('Seçileni arşivle',self.archive_penalty)
        ttk.Label(f,text='Geçmiş ceza kararları',padding=(0,8,0,3)).pack(anchor='w')
        self.past_penalty_tree=self.owner.tree(f,['Karar tarihi','Verilen nöbet','Tuttu','Sonucu belirsiz'],[115,110,140,110])
        self.past_penalty_tree.configure(height=4)
    def edit_penalty(self,edit):
        try:
            self.require_saved();row=None
            if edit:
                items=self.ptree.selection()
                if not items:raise ValueError('Ceza kaydı seçin')
                with self.app.conn() as c:row=dict(c.execute('SELECT * FROM penalties WHERE id=? AND student_id=?',(int(self.ptree.item(items[0],'values')[0]),self.sid)).fetchone())
                if row['archived'] or row['voided']:raise ValueError('Arşiv kaydı için ana Cezalılar ekranından geri alın')
            labels=['Ceza nöbeti adedi','Görev yeri','Başlangıç (YYYY-AA-GG)']
            defaults=dict(zip(labels,[str(row['total']) if row else '1',row['target'] if row else 'Devriye',row['start_day'] if row else self.owner.day.get()]))
            def submit(v):
                if row and v[2]!=row['start_day']:raise ValueError('Mevcut cezanın ilk kayıt tarihi değiştirilemez')
                with self.app.conn() as c:
                    self.app.ceza_nobet.record(c,self.sid,int(v[0]),v[1],v[2]);self.app.log(c,'Profil ceza kaydı',str(self.sid))
                self.owner.refresh_all();self.refresh_related()
            self.owner.dialog('Öğrencinin cezası',labels,submit,{'Görev yeri':list(self.app.ceza_nobet.KINDS)},defaults)
        except Exception as e:messagebox.showerror('Ceza',str(e),parent=self)
    def archive_penalty(self):
        try:
            items=self.ptree.selection()
            if not items:return
            if not messagebox.askyesno('Arşivle','Seçili ceza arşivlensin mi? Geçmiş korunur.',parent=self):return
            self.app.ceza_nobet.clean(self.app,[int(self.ptree.item(i,'values')[0]) for i in items],'archive');self.owner.refresh_all();self.refresh_related()
        except Exception as e:messagebox.showerror('Ceza',str(e),parent=self)
    def build_history(self):
        f=self.pages['Nöbet geçmişi ve grafik'];self.history_label=ttk.Label(f,text='',wraplength=830);self.history_label.pack(anchor='w',pady=5)
        tabs=ttk.Notebook(f);tabs.pack(fill='both',expand=True);table=ttk.Frame(tabs);chart=ttk.Frame(tabs);tabs.add(table,text='Görevler');tabs.add(chart,text='Pasta grafik')
        self.htree=self.owner.tree(table,['ID','Tarih','Nöbet','Saat','Yer','Rol','Sonuç','Onay','Tip'],[45,105,145,125,90,65,85,95,65])
        ttk.Button(table,text='Seçili görevin listesini aç / görevi değiştir',command=self.open_assignment,style='Teal.TButton').pack(anchor='w',pady=5)
        ttk.Label(table,text='Liste ekranında mevcut onay, uygunluk ve çakışma kontrolleri uygulanır.').pack(anchor='w')
        self.chart=charts.PieView(chart,self.refresh_chart);self.chart.group.set('Nöbet türü');self.history_totals={};self.history_slots={}
        self.chart.pack(fill='both',expand=True)
    def refresh_chart(self):
        group=self.chart.group.get();measure=self.chart.measure.get();index=1 if measure==charts.MEASURES[1] else 0
        if group=='Nöbet türü':data={k:v[index] for k,v in self.history_totals.items()}
        elif group=='Sınıf':data={self.values['class'].get():sum(v[index] for v in self.history_totals.values())}
        elif group=='Tür + saat':data={k:v[index] for k,v in self.history_slots.items()}
        else:
            data={}
            for label,values in self.history_slots.items():
                key=label.split(' ',2)[-1] if label.startswith('Yatakhane ') else label.split(' ',1)[-1]
                data[key]=data.get(key,0)+values[index]
        self.chart.set_data(sorted(((k,v) for k,v in data.items() if v>0),key=lambda r:-r[1]),measure)
    def open_assignment(self):
        items=self.htree.selection()
        if not items:return
        row=self.htree.item(items[0],'values');self.owner.day.set(row[1]);self.owner.kind.set(row[2]);self.owner.refresh_schedule();self.owner.tabs.select(self.owner.pages['Nöbet Hazırla'])
        for item in self.owner.atree.get_children():
            if self.owner.atree.item(item,'values')[0]==row[0]:self.owner.atree.selection_set(item);self.owner.atree.see(item)
        self.destroy()
    def refresh_related(self):
        for tree in (self.etree,self.ptree,self.past_penalty_tree,self.htree):self.owner.clear(tree)
        if not self.sid:return
        today=self.owner.day.get()
        with self.app.conn() as c:
            exemptions=list(c.execute('SELECT * FROM exemptions WHERE student_id=? ORDER BY start DESC,id DESC',(self.sid,)))
            for r in exemptions:self.etree.insert('','end',values=(r['id'],r['reason'],r['start'],r['end'] or 'Süresiz'))
            current=[r['reason'] for r in exemptions if r['start']<=today and (not r['end'] or r['end']>=today)]
            profile=c.execute('SELECT official_duty FROM student_profiles WHERE student_id=?',(self.sid,)).fetchone()
            self.duty_label.config(text='Görevi: '+(profile[0] if profile and profile[0] else 'Tanımlanmadı')+'\n'+today+' için etkin muafiyet: '+(', '.join(current) or 'Yok'))
            cases=list(c.execute('SELECT * FROM penalties WHERE student_id=? AND voided=0 ORDER BY start_day DESC,id DESC',(self.sid,)))
            for r in cases:
                done,reserved=self.app.ceza_nobet.counts(c,r['id']);self.ptree.insert('','end',values=(r['id'],r['total'],done,max(0,r['total']-done),r['target'],r['start_day'],'Arşiv' if r['archived'] else ('Aktif' if r['active'] else 'Pasif')))
            past=list(c.execute("""SELECT p.award_day,p.total,COUNT(a.id)+COALESCE(
                    (SELECT SUM(count) FROM historical_confirmed_penalty_hours h
                     WHERE h.source=p.source AND h.source_row=p.source_row),0) AS matched
                FROM historical_penalty_awards p
                LEFT JOIN historical_penalty_duties h USING(source,source_row)
                LEFT JOIN assignments a ON a.id=h.assignment_id AND a.duty_type='Ceza'
                    AND a.role='Asıl' AND a.result='Tuttu'
                    AND EXISTS(SELECT 1 FROM sheets s WHERE s.id=a.sheet_id AND s.status='Onaylandı')
                WHERE p.student_id=? GROUP BY p.source,p.source_row ORDER BY p.award_day DESC""",(self.sid,)))
            for r in past:
                self.past_penalty_tree.insert('','end',values=(r['award_day'],r['total'],r['matched'],r['total']-r['matched']))
            total=sum(r['total'] for r in cases)+sum(r['total'] for r in past)
            self.penalty_label.config(text=f'{len(cases)+len(past)} ceza kararı • toplam {total} ceza nöbeti verilmiş')
            rows,self.history_totals,self.history_slots=summary(self.app,c,self.sid)
        for r in rows:self.htree.insert('','end',values=(r['id'],r['day'],r['kind'],r['slot'],r['location'],r['role'],r['result'],r['status'],r['duty_type']))
        parts=[f'{kind}: {values[0]} nöbet / {values[1]:g} saat' for kind,values in self.history_totals.items()]
        self.history_label.config(text=' • '.join(parts) or 'Henüz onaylı tutulmuş asıl nöbet yok.');self.refresh_chart()
