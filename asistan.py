"""In-app Codex chat and isolated application update candidates."""
import difflib
import hashlib
import json
import os
from pathlib import Path
import queue
import re
import shutil
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import zipfile

SOURCE_ARCHIVE = 'asistan_kaynak.zip'
ASSETS = {'Nobetra.ico', 'Nobetra.png',
          'yazdirma_sablonlari.json', 'Ogrenci_Aktarim_Sablonu.xlsx'}
HELP = """NÖBETRA uygulaması için Türkçe yardımcı ol.
Nöbet Hazırla: tarih, tür, Excel aktarımı, otomatik veya manuel hazırla, sonraki günü hazırla, Tuttu/Tutmadı ve onay. Planlandı satırındaki × ile aynı saat, yer ve rol için uygun başka öğrenci seçilir. Yazdır / Excel / PDF düğmesi düzenlenebilir XLSX veya tarayıcıdan yazdırma ve PDF kaydetme seçeneği açar.
Manuel hazırla: öğrencileri Excel sırasıyla göster, Listeyi ters çevir ile yönü değiştir. Öğrencileri sırayla seç, toplu ekle, Excel satırlarını yapıştır veya Seçili sıraya yerleştir ile belirli bir yere ata. Tüm saat ve sıra yerleri baştan görünür. Saatlerde 3/1, 3 onaylı tutulmuş asıl nöbetin 1'inin ceza olduğunu gösterir.
Yatakhane Kat 48, Devriye 50 (44 asıl, 6 yedek, 5 ekip), Yatakhane Giriş 14 kişiliktir.
Eski Devriye Excel listeleri 42 kişi olarak da aktarılabilir. Giriş nöbeti gece nöbetinden bağımsızdır.
Sonraki listede tekrar yazılacak Tutmadı işaretli asılları kullanıcı seçer. Tekrar listesinden çıkar, o görevin önerisini kaldırır ve Tutmadı kaydını geçmişte bırakır. Tuttu işaretli yedek ertesi gün atlanır.
Cezalılar: Excel’den yapıştır ile ad soyad veya okul numarasından toplu ceza kaydı eklenir. Otomatik hazırlama ekranında toplu seçim yapılır. Program 01:00–03:00 ile 03:00–05:00 vardiyalarını önerir ve mümkünse son ceza vardiyasının tersini seçer. Saatler elle değiştirilebilir. 23:00–01:00 veya 05:00–07:00 seçilirse onay istenir.
Otomatik hazırla: Sıra başlangıçları ekranı son listede yazılan son asılı ve ilk boş tarihteki devam adayını gösterir. Gece nöbeti devam sırası, seçili tür için tarih, sınıf sırası, sınıf içi yön ve sıradaki öğrenciyi kaydeder. Bu tarihten sonra onaylanan manuel ve Excel listelerindeki normal asıllar da sırayı ilerletir. Ceza saatleri sabitlenir; normal vardiyalar aynı türde geçmişte daha az tutulan saatlere öncelik verilerek dağıtılır. Önizleme geçmiş vardiya adetlerini ve son saati gösterir. Devriye ve kat aynı gece aynı sınıfta çakışamaz; Yatakhane Kat önceliklidir. Onaylı Tuttu/Asıl ceza görevi kalan cezayı azaltır.
Geçmiş nöbet listesi yükle: Excel veya tek fotoğraf seçilir. Excel tarihi dosyadan alınır ve iki günlük giriş listesi birlikte önizlenir. Mevcut fotoğraf akışında tablo köşeleri işaretlenir, OCR sonucu ve tarih kontrol edilir. 23:00'dan sonraki gece vardiyaları ertesi takvim gününe geçer. Onaylı listeyi düzeltmek için önce onay kaldırılır.
Öğrenciler: toplu Excel şablonu ve okul numarasıyla güncelleme. İstatistikler: tablo, pasta ve çubuk grafik.
Yedekleme: veritabanı yedeği ve seçerek sıfırlama. Gerçek öğrenci veya nöbet bilgilerine erişme.
Soruya kısa ve somut cevap ver. İlgili menü, düğme ve adımları sırala. Bir hata sorulursa olası nedeni gözlemden ayır ve kullanıcının uygulayabileceği kontrolü ver. Emin olmadığın özellikleri varmış gibi anlatma.
Bir değişiklik istenirse Güncelleme hazırla seçeneğini açıkla. Sohbet sırasında kodu değiştirme."""

GUIDES = {
    'Otomatik ve ceza': 'Nöbet Hazırla → tarih ve türü seç → Otomatik hazırla. Uygun ceza kayıtlarında saat seç. Tutmadı durumundaki asıl nöbetçiler için seçim yap. Önizlemeyi kontrol edip taslağı kaydet.',
    'Geçmişe aktar': 'Nöbet Hazırla → nöbet türünü seç → Geçmiş nöbet listesi yükle. Excel veya fotoğrafı seç. Tarihleri ve eşleşmeleri kontrol et, sonuçları işaretle ve geçmişe kaydet.',
    'Manuel nöbet': 'Nöbet Hazırla → Manuel hazırla. Sırayla eklemek için öğrencileri işaretle. Belirli bir yer için soldan öğrenciyi, sağdan boş sırayı seç ve Seçili sıraya yerleştir düğmesine bas.',
}


def codex_executable():
    found = shutil.which('codex.exe')
    if found:
        return Path(found)
    installed = Path(os.environ.get('LOCALAPPDATA', '')) / 'OpenAI' / 'Codex' / 'bin'
    candidates = list(installed.glob('*/codex.exe'))
    if candidates:
        return max(candidates, key=lambda p: p.stat().st_mtime)
    raise ValueError('Codex bulunamadı. Bu bilgisayarda Codex kurulu ve hesabın açık olmalı.')


def codex_command(workspace, schema=None, effort='high'):
    config = [
        'approval_policy="never"',
        'windows.sandbox="unelevated"',
        'web_search="disabled"',
        'features.shell_tool=false',
        f'model_reasoning_effort="{effort}"',
    ]
    command = [str(codex_executable()), 'exec', '--ignore-user-config', '--ephemeral',
               '--skip-git-repo-check', '--color', 'never', '--sandbox', 'read-only', '-C', str(workspace)]
    for setting in config:
        command.extend(('-c', setting))
    if schema:
        command.extend(('--output-schema',str(schema)))
    return [*command, '-']


def codex_failure(stderr):
    details = stderr.casefold()
    if any(word in details for word in ('401 unauthorized', 'not logged in', 'refresh token', 'authentication required')):
        return 'Codex oturumu doğrulanamadı. Codex’i açıp hesabına yeniden giriş yap.'
    if any(word in details for word in ('429', 'usage limit', 'rate limit')):
        return 'Codex kullanım sınırına ulaştı. Bir süre sonra tekrar dene.'
    if any(word in details for word in ('connection', 'dns', 'failed to send request', 'timed out')):
        return 'Codex’e bağlantı kurulamadı. İnternet bağlantısını kontrol edip tekrar dene.'
    return 'Codex yanıt veremedi. Codex uygulamasında hesabının açık olduğunu kontrol et.'


def run_codex(workspace, prompt, schema=None, effort='high'):
    try:
        result = subprocess.run(
            codex_command(workspace, schema, effort),
            cwd=workspace, input=prompt, capture_output=True, text=True,
            encoding='utf-8', errors='replace', timeout=300 if schema else 180,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
        )
    except subprocess.TimeoutExpired:
        raise ValueError('Codex yanıtı zamanında tamamlanmadı. Tekrar deneyebilirsin.') from None
    if result.returncode or not result.stdout.strip():
        raise ValueError(codex_failure(result.stderr))
    return result.stdout.strip()


def source_file(name):
    return (Path(name).name == name and
            (name in ASSETS or re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*\.py',name)))


def source_snapshot(directory):
    return {p.name: p.read_bytes() for p in directory.iterdir()
            if source_file(p.name) and p.is_file() and not p.is_symlink() and not p.is_junction()}


def snapshot_digest(files):
    return hashlib.sha256(json.dumps(
        {name:hashlib.sha256(data).hexdigest() for name,data in files.items()},
        sort_keys=True).encode()).hexdigest()


def unpack_source(assets, stage):
    with zipfile.ZipFile(assets / SOURCE_ARCHIVE) as archive:
        names = archive.namelist()
        if len(names)!=len(set(names)) or not all(source_file(name) for name in names):
            raise ValueError('Kaynak paketi beklenen dosyaları içermiyor.')
        original = {name:archive.read(name) for name in names}
    if not {'app.py','asistan.py',*ASSETS} <= original.keys():
        raise ValueError('Kaynak paketinde eksik uygulama dosyası var.')
    for name, contents in original.items():
        (stage / name).write_bytes(contents)
    return original


def build_settings(assets):
    config_file = assets / 'asistan_paketleme.json'
    if not config_file.is_file():
        raise ValueError('Bu kurulumda paketleme ayarları bulunamadı.')
    config = json.loads(config_file.read_text(encoding='utf-8'))
    python = Path(config['python'])
    dependencies = Path(config['dependencies'])
    if not python.is_file() or not (dependencies / 'PyInstaller').is_dir():
        raise ValueError('Python ve PyInstaller paketleme ortamı bu bilgisayarda bulunamadı.')
    return python, dependencies, config_file


def source_context(files):
    context = {name:data.decode('utf-8-sig').replace('\r\n','\n')
               for name,data in files.items() if name.endswith('.py')}
    planner=context.get('nobet_planlayici.py','')
    if planner:
        planner=re.sub(r'^DEFAULTS=.*$', 'DEFAULTS = {}  # Yerel başlangıç öğrenci adları gizlendi.',planner,flags=re.M)
        planner=re.sub(r'for target,last_name,last_surname in \[.*?\]:',
                       'for target,last_name,last_surname in []:  # Yerel öğrenci adları gizlendi.',planner)
        context['nobet_planlayici.py']=planner
    return context


def prepare_update(assets, updates, request):
    build_settings(assets)
    updates.mkdir(parents=True,exist_ok=True)
    job = Path(tempfile.mkdtemp(prefix='Guncelleme_',dir=updates))
    stage = job / 'kaynak'
    stage.mkdir()
    original = unpack_source(assets,stage)
    schema = {
        'type':'object','additionalProperties':False,'required':['summary','edits'],
        'properties':{
            'summary':{'type':'string'},
            'edits':{'type':'array','items':{
                'type':'object','additionalProperties':False,'required':['file','old','new'],
                'properties':{key:{'type':'string'} for key in ('file','old','new')}
            }}
        }
    }
    prompt = (
        'Verilen NÖBETRA Python kaynaklarına istenen güncellemeyi yaz. '
        'Araç kullanma. Her düzenlemede file dosya adı, old kaynakta tam bir kez geçen kod, new yeni kod olmalı. '
        'Yeni Python dosyası gerekiyorsa old boş olur. Değişiklikleri küçük tut. '
        'Mevcut kayıt formatını ve davranışları koru. Canlı veri, ağ, kurulum veya gizli bilgi kullanma. '
        'Gizlendi olarak işaretlenen satırları düzenleme. Test çalıştırdığını iddia etme. '
        'İstek eksikse edits boş ve summary alanı Türkçe açıklama olsun.\n\n'
        'Güncelleme isteği:\n'+request+'\n\nPython kaynakları:\n'
        +json.dumps(source_context(original),ensure_ascii=False)
    )
    with tempfile.TemporaryDirectory(prefix='nobetra_asistan_') as folder:
        workspace=Path(folder)
        schema_file=workspace/'cikti_semasi.json'
        schema_file.write_text(json.dumps(schema),encoding='utf-8')
        answer=json.loads(run_codex(workspace,prompt,schema_file))
    report=answer['summary']
    apply_edits(stage,answer['edits'])
    candidate = source_snapshot(stage)
    if not {'app.py','asistan.py',*ASSETS} <= candidate.keys():
        raise ValueError('Güncellemede gerekli uygulama dosyaları eksik.')
    if candidate==original:
        return report+'\n\nPaketlenecek kaynak değişikliği yok.', None, None
    for name, contents in candidate.items():
        if name.endswith('.py'):
            compile(contents,name,'exec')
    diff=[]
    for name in sorted(candidate.keys() | original.keys()):
        if candidate.get(name)==original.get(name):
            continue
        if name.endswith('.py'):
            diff.extend(difflib.unified_diff(
                original.get(name,b'').decode('utf-8-sig').splitlines(keepends=True),
                candidate.get(name,b'').decode('utf-8-sig').splitlines(keepends=True),
                fromfile='önceki/'+name,tofile='aday/'+name))
        else:
            diff.append(f'Güncellenen varlık: {name}\n')
    (job/'Degisiklikler.txt').write_text(''.join(diff),encoding='utf-8')
    (job/'Test_raporu.txt').write_text(report+'\n\nPython kaynakları sözdizimi kontrolünden geçti. Aday henüz çalıştırılmadı.\n',encoding='utf-8')

    return report+'\n\nAday kaynak hazır. Değişiklikleri ve test raporunu inceleyip Onayla ve kur düğmesini kullan.',job,snapshot_digest(candidate)


def apply_edits(stage, edits):
    if not isinstance(edits,list) or len(edits)>40:
        raise ValueError('Asistan beklenenden fazla değişiklik üretti. İsteği daha küçük adımlara böl.')
    changed={}
    for edit in edits:
        name,old,new=(edit[k] for k in ('file','old','new'))
        if not isinstance(name,str) or not source_file(name) or not name.endswith('.py'):
            raise ValueError('Asistan izin verilen kaynak dosyalarının dışına çıktı.')
        if not all(isinstance(value,str) for value in (old,new)):
            raise ValueError('Asistan geçersiz bir düzenleme döndürdü.')
        path=stage/name
        if old:
            if name not in changed and not path.is_file():
                raise ValueError(f'{name} kaynakta bulunamadı.')
            current=changed[name] if name in changed else path.read_text(encoding='utf-8-sig')
            if current.count(old)!=1:
                raise ValueError(f'{name} değişikliği tek bir yere uygulanamadı. İsteği tekrar dene.')
            changed[name]=current.replace(old,new,1)
        elif path.exists() or name in changed:
            raise ValueError(f'{name} için geçersiz dosya oluşturma isteği.')
        else:
            changed[name]=new
    for name,contents in changed.items():
        compile(contents,name,'exec')
    for name,contents in changed.items():
        (stage/name).write_text(contents,encoding='utf-8')


def install_update(assets, job, install_dir, digest):
    python, dependencies, config_file = build_settings(assets)
    candidate = source_snapshot(job/'kaynak')
    if snapshot_digest(candidate)!=digest:
        raise ValueError('Aday kaynak inceleme için hazırlandıktan sonra değişmiş. Yeniden güncelleme hazırla.')
    # Package only the reviewed files, without model-generated build scripts.
    clean=Path(tempfile.mkdtemp(prefix='paket_kaynagi_',dir=job))
    for name, contents in candidate.items():
        (clean/name).write_bytes(contents)
    with zipfile.ZipFile(clean/SOURCE_ARCHIVE,'w',zipfile.ZIP_DEFLATED) as archive:
        for name, contents in candidate.items():
            archive.writestr(name,contents)
    shutil.copyfile(config_file,clean/config_file.name)
    bootstrap='import runpy,sys; sys.path.insert(0,sys.argv.pop(1)); runpy.run_module("PyInstaller",run_name="__main__")'
    cmd=[str(python),'-I','-c',bootstrap,str(dependencies),'--onefile','--windowed','--exclude-module','numpy',
         '--name','NOBETRA Aday','--distpath',str(job),
         '--workpath',str(job/'derleme'),'--specpath',str(job),'--icon',str(clean/'Nobetra.ico')]
    for name in sorted(ASSETS | {SOURCE_ARCHIVE,config_file.name}):
        cmd.extend(('--add-data',str(clean/name)+';.'))
    cmd.append(str(clean/'app.py'))
    env=os.environ.copy()
    env.pop('PYTHONPATH',None)
    env.update(PYINSTALLER_CONFIG_DIR=str(job/'paketleme_onbellegi'))
    with (job/'Paketleme_gunlugu.txt').open('wb') as log:
        build=subprocess.run(cmd,cwd=job,env=env,stdout=log,stderr=subprocess.STDOUT,timeout=300,
                             creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
    exe=job/'NOBETRA Aday.exe'
    if build.returncode or not exe.is_file():
        raise ValueError('EXE paketlenemedi. Aday klasöründeki paketleme günlüğünü kontrol et.')
    name='NOBETRA '+job.name+'.exe'
    staged_path=None
    try:
        with tempfile.NamedTemporaryFile(dir=install_dir,suffix='.tmp',delete=False) as staged:
            staged_path=Path(staged.name)
            with exe.open('rb') as source:
                shutil.copyfileobj(source,staged)
        staged_path.rename(install_dir/name)
    finally:
        if staged_path:
            staged_path.unlink(missing_ok=True)
    return 'Yeni EXE programın yanına eklendi. Mevcut programı kapatıp aynı klasörde şu dosyayı aç:\n'+name


class AssistantPanel(ttk.Frame):
    def __init__(self, parent, assets, updates, context=lambda: ''):
        super().__init__(parent)
        self.assets, self.updates = assets, updates
        self.context = context
        self.history=[]
        self.results=queue.Queue()
        self.last_job=None
        self.candidate_hash=None
        ttk.Label(self,text='Sorularını sor veya bir uygulama güncellemesi hazırla.',
                  font=('Segoe UI',13,'bold')).pack(anchor='w')
        ttk.Label(self,text='Yazdığın mesajlar Codex’e gönderilir. Öğrenci bilgisi, çizelge veya şifre paylaşma.',
                  wraplength=800).pack(anchor='w',pady=(4,8))
        guides=ttk.Frame(self)
        guides.pack(fill='x')
        for label, steps in GUIDES.items():
            ttk.Button(guides,text=label,command=lambda name=label,text=steps:self.append(name,text)).pack(side='left',padx=(0,6))
        self.depth=tk.StringVar(master=self,value='Ayrıntılı')
        ttk.Combobox(guides,textvariable=self.depth,values=('Ayrıntılı','Hızlı'),state='readonly',width=10).pack(side='right')
        self.chat=ScrolledText(self,wrap='word',height=12,state='disabled',font=('Segoe UI',11),
                               bg='#f8fafc',fg='#1e293b',relief='flat',bd=0,
                               highlightthickness=1,highlightbackground='#e2e8f0',
                               padx=18,pady=16,selectbackground='#bfdbfe')
        self.chat.pack(fill='both',expand=True,pady=(10,8))
        for tag,color,left,right in [('user','#dbeafe',80,16),
                                     ('assistant','#ecfdf5',16,80),
                                     ('error','#fff1f2',16,40)]:
            self.chat.tag_configure(tag,background=color,lmargin1=left,lmargin2=left,
                                    rmargin=right,spacing1=8,spacing3=10)
            self.chat.tag_configure(tag+'_author',font=('Segoe UI',10,'bold'),
                                    foreground='#9f1239' if tag=='error' else '#0f766e',
                                    spacing1=12,spacing3=4)
        self.input=ScrolledText(self,height=3,wrap='word',font=('Segoe UI',11),
                                bg='#ffffff',fg='#1e293b',insertbackground='#2563eb',
                                relief='flat',bd=0,padx=14,pady=12,highlightthickness=1,
                                highlightbackground='#cbd5e1',highlightcolor='#2563eb')
        self.input.pack(fill='x',pady=(0,6))
        ttk.Label(self,text='Mesajını yaz • Ctrl+Enter ile gönder',
                  foreground='#64748b',font=('Segoe UI',9)).pack(anchor='w',pady=(0,4))
        self.input.bind('<Control-Return>',self.keyboard_send)
        bar=ttk.Frame(self)
        bar.pack(fill='x',pady=(2,6))
        self.ask=ttk.Button(bar,text='Asistana sor',command=lambda:self.submit(False),style='Primary.TButton')
        self.ask.pack(side='left',padx=(0,6))
        self.update=ttk.Button(bar,text='Güncelleme hazırla',command=lambda:self.submit(True),style='Teal.TButton')
        self.update.pack(side='left',padx=(0,6))
        self.open_job=ttk.Button(bar,text='Aday klasörünü aç',command=self.open_candidate,state='disabled')
        self.open_job.pack(side='left')
        self.install=ttk.Button(bar,text='Onayla ve kur',command=self.install_candidate,state='disabled')
        self.install.pack(side='left',padx=6)
        self.status=ttk.Label(self,text='Hazır')
        self.status.pack(anchor='w')
        self.append('Asistan','Nasıl yardımcı olabilirim? Örneğin “Devriye Excel’ini nasıl aktarırım?” diye sorabilirsin.')

    def append(self, author, text):
        tag='user' if author in ('Sen','Güncelleme isteğin') else (
            'error' if author=='İşlem tamamlanamadı' else 'assistant')
        self.chat.configure(state='normal')
        self.chat.insert('end',f'{author}\n',(tag,tag+'_author'))
        self.chat.insert('end',f'{text}\n',tag)
        self.chat.insert('end','\n')
        self.chat.configure(state='disabled')
        self.chat.see('end')

    def keyboard_send(self,event):
        self.submit(False)
        return 'break'

    def submit(self,update):
        if self.ask.instate(('disabled',)):
            return
        request=self.input.get('1.0','end').strip()
        if not request:return
        if len(request)>5000:
            messagebox.showinfo('Asistan','Mesajı 5.000 karakterin altında tut.',parent=self)
            return
        if update and not messagebox.askyesno('Aday güncelleme',
            'Codex bu isteği ayrı bir kaynak kopyasında hazırlayacak. Aday güncelleme oluşturulsun mu?',parent=self):
            return
        self.input.delete('1.0','end')
        self.append('Güncelleme isteğin' if update else 'Sen',request)
        self.ask.configure(state='disabled')
        self.update.configure(state='disabled')
        self.install.configure(state='disabled')
        self.status.configure(text='Aday sürüm hazırlanıyor…' if update else 'Yanıt hazırlanıyor…')
        context=self.history[-12:]
        app_context=self.context()
        effort='high' if self.depth.get()=='Ayrıntılı' else 'medium'
        def worker():
            try:
                if update:
                    answer,job,digest=prepare_update(self.assets,self.updates,request)
                else:
                    prompt=HELP+'\n\nUygulama bağlamı: '+app_context+'\n\n'+ '\n\n'.join(f'{role}: {msg}' for role,msg in context)
                    with tempfile.TemporaryDirectory(prefix='nobetra_sohbet_') as folder:
                        answer=run_codex(Path(folder),prompt+'\n\nSoru: '+request,effort=effort)
                    job=None
                    digest=None
                self.results.put((answer,job,digest,True))
            except (ValueError,OSError,subprocess.SubprocessError,zipfile.BadZipFile,SyntaxError,KeyError) as error:
                self.results.put((str(error),None,None,False))
        threading.Thread(target=worker,daemon=True).start()
        self.after(180,lambda:self.poll(request,update))

    def poll(self,request,update,installing=False):
        try:
            answer,job,digest,ok=self.results.get_nowait()
        except queue.Empty:
            self.after(180,lambda:self.poll(request,update,installing))
            return
        self.append('Asistan' if ok else 'İşlem tamamlanamadı',answer)
        if ok and not update and not installing:
            self.history.extend([('Kullanıcı',request),('Asistan',answer[:6000])])
        if job:
            self.last_job=job
            self.candidate_hash=digest
            self.open_job.configure(state='normal')
        elif update and ok:
            self.last_job=None
            self.candidate_hash=None
            self.open_job.configure(state='disabled')
        elif installing and ok:
            self.candidate_hash=None
        self.ask.configure(state='normal')
        self.update.configure(state='normal')
        self.install.configure(state='normal' if self.candidate_hash else 'disabled')
        self.status.configure(text='Hazır')

    def open_candidate(self):
        if self.last_job:
            os.startfile(self.last_job)

    def install_candidate(self):
        if self.ask.instate(('disabled',)) or not self.candidate_hash or not self.last_job:
            return
        if not messagebox.askyesno('Aday sürümü kur',
            'Aday klasöründeki değişiklikleri ve test raporunu inceledin mi? '
            'Onaylarsan aday paketlenecek ve yeni EXE mevcut programın yanına eklenecek. '
            'Yeni sürüme geçmeden önce Yedekleme ekranından veri yedeği alabilirsin.',parent=self):
            return
        job,digest=self.last_job,self.candidate_hash
        self.ask.configure(state='disabled')
        self.update.configure(state='disabled')
        self.install.configure(state='disabled')
        self.status.configure(text='Onaylanan sürüm paketleniyor…')
        def worker():
            try:
                answer=install_update(self.assets,job,self.updates.parent,digest)
                self.results.put((answer,None,None,True))
            except (OSError,ValueError,subprocess.SubprocessError,KeyError) as error:
                self.results.put((str(error),None,None,False))
        threading.Thread(target=worker,daemon=True).start()
        self.after(180,lambda:self.poll('',False,True))
