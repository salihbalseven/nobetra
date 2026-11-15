"""Common modern styling, sortable tables and wrapping action rows."""
import tkinter as tk
from tkinter import ttk

def apply_theme(root):
    root.configure(bg='#f3f0fa');root.option_add('*Font',('Segoe UI',10))
    style=ttk.Style(root);style.theme_use('clam')
    style.configure('.',font=('Segoe UI',10),background='#ffffff',foreground='#2d2540')
    style.configure('Canvas.TFrame',background='#f3f0fa')
    style.configure('Canvas.TLabel',background='#f3f0fa')
    style.configure('TFrame',background='#ffffff');style.configure('TLabel',background='#ffffff',foreground='#3d3354')
    style.configure('TButton',padding=(12,8),background='#faf8ff',foreground='#2d2540',bordercolor='#d6cde8',lightcolor='#faf8ff',darkcolor='#faf8ff',relief='flat')
    style.map('TButton',background=[('active','#e8e2f3'),('pressed','#ebe4ff')])
    style.configure('Primary.TButton',background='#5943c5',foreground='white',bordercolor='#5943c5')
    style.map('Primary.TButton',background=[('active','#4831b0'),('pressed','#412d91')],foreground=[('disabled','#d6cde8'),('!disabled','white')])
    for name,color,hover,foreground in [('Success','#15803d','#166534','white'),('Warning','#f59e0b','#d97706','#2d2540'),('Import','#7c3aed','#6d28d9','white'),('Teal','#5540b5','#443294','white'),('Danger','#dc2626','#b91c1c','white')]:
        style.configure(name+'.TButton',background=color,foreground=foreground,bordercolor=color)
        style.map(name+'.TButton',background=[('active',hover),('pressed',hover)],foreground=[('!disabled',foreground)])
    style.configure('TEntry',padding=6,fieldbackground='#ffffff',bordercolor='#d6cde8')
    style.configure('Quiet.TButton',padding=(8,5))
    style.configure('Compact.Danger.TButton',padding=(8,5))
    style.configure('TCombobox',padding=5,fieldbackground='#ffffff',bordercolor='#d6cde8')
    style.configure('TNotebook',background='#f3f0fa',borderwidth=0)
    style.configure('TNotebook.Tab',padding=(13,11),background='#e8e2f3',foreground='#544967')
    style.map('TNotebook.Tab',background=[('selected','#5540b5'),('active','#e8e2f3'),('!selected','#f7f5fc')],foreground=[('selected','#ffffff'),('!selected','#544967')],padding=[('selected',(13,11)),('!selected',(13,11))],font=[('selected',('Segoe UI',10)),('!selected',('Segoe UI',10))],expand=[('selected',(0,0,0,0)),('!selected',(0,0,0,0))])
    style.layout('Workspace.TNotebook.Tab',[])
    style.configure('Workspace.TNotebook',background='#f7f5fc',borderwidth=0,tabmargins=0)
    style.configure('Treeview',rowheight=32,background='#ffffff',fieldbackground='#ffffff',foreground='#2d2540',bordercolor='#e8e2f3')
    style.configure('Wide.Vertical.TScrollbar',arrowsize=20)
    style.configure('Treeview.Heading',font=('Segoe UI',10,'bold'),background='#e8e2f3',foreground='#3d3354',padding=(8,10),relief='flat')
    style.map('Treeview',background=[('selected','#ebe4ff')],foreground=[('selected','#33246f')])
    style.configure('TLabelframe',background='white',bordercolor='#e8e2f3');style.configure('TLabelframe.Label',font=('Segoe UI',10,'bold'),foreground='#412d91')

class DataTree(ttk.Treeview):
    def __init__(self,*args,**kwargs):
        super().__init__(*args,**kwargs);self.directions={};self.next_row=0
        self.tag_configure('stripe_even',background='#ffffff');self.tag_configure('stripe_odd',background='#f7f5fc')
        self.bind('<Control-c>',self.copy_selection)
    def insert(self,parent,index,iid=None,**kwargs):
        kwargs['tags']=tuple(kwargs.get('tags',()))+('stripe_even' if self.next_row%2==0 else 'stripe_odd',)
        self.next_row+=1
        return super().insert(parent,index,iid=iid,**kwargs)
    def delete(self,*items):
        super().delete(*items)
        if not self.get_children():self.next_row=0
    def sort_column(self,column):
        reverse=self.directions.get(column,False)
        def key(item):
            value=self.set(item,column).strip()
            try:return (0,float(value.replace(',','.')))
            except ValueError:return (1,value.replace('İ','i').replace('I','ı').casefold())
        rows=sorted(self.get_children(),key=key,reverse=reverse)
        for i,item in enumerate(rows):
            tags=tuple(tag for tag in self.item(item,'tags') if tag not in ('stripe_even','stripe_odd'))
            self.move(item,'',i);self.item(item,tags=tags+('stripe_even' if i%2==0 else 'stripe_odd',))
        self.directions[column]=not reverse
    def copy_selection(self,event=None):
        rows=['\t'.join(map(str,self.item(item,'values'))) for item in self.selection()]
        if rows:self.clipboard_clear();self.clipboard_append('\n'.join(rows))
        return 'break'

class ActionBar(ttk.Frame):
    def __init__(self,parent,**kwargs):
        super().__init__(parent,**kwargs);self.buttons=[];self.bind('<Configure>',self.reflow)
    def add(self,text,command,style='TButton'):
        if style=='TButton':style=['Primary.TButton','Import.TButton','Teal.TButton','Success.TButton'][len(self.buttons)%4]
        button=ttk.Button(self,text=text,command=command,style=style);self.buttons.append(button);self.after_idle(self.reflow);return button
    def reflow(self,event=None):
        width=max(self.winfo_width(),250);row=0;column=0;used=0
        for button in self.buttons:
            size=button.winfo_reqwidth()+8
            if used and used+size>width:row+=1;column=0;used=0
            button.grid(row=row,column=column,sticky='w',padx=(0,8),pady=4);column+=1;used+=size

def header(root, version, logo=None):
    background = '#201a39'
    bar = tk.Frame(root, bg=background, padx=22, pady=11)
    bar.pack(fill='x')
    if logo and logo.is_file():
        from PIL import Image, ImageTk
        with Image.open(logo) as original:
            display = original.convert('RGBA')
        display.thumbnail((58, 58), Image.Resampling.LANCZOS)
        bar.logo = ImageTk.PhotoImage(display, master=root)
        tk.Label(bar, image=bar.logo, bg=background, bd=0).pack(side='left', padx=(0, 14))
    brand = tk.Frame(bar, bg=background)
    brand.pack(side='left')
    tk.Label(brand, text='NÖBETRA', bg=background, fg='white',
             font=('Segoe UI', 22, 'bold')).pack(anchor='w')
    tk.Label(brand, text='Nöbet ve vardiya planlama', bg=background, fg='#c9bee5',
             font=('Segoe UI', 10)).pack(anchor='w')
    tk.Label(bar, text=version, bg='#342b54', fg='#68ecd9',
             font=('Segoe UI', 10, 'bold'), padx=12, pady=6).pack(side='right')
    tk.Label(bar, text='Geliştirici · Salih Balseven', bg=background, fg='#e5ddf5',
             font=('Segoe UI', 10)).pack(side='right', padx=(0, 16))
    tk.Frame(root, bg='#6d57e4', height=3).pack(fill='x')


PAGE_DETAILS={
    'Öğrenciler':('ÖĞ','Öğrenci rehberi','Bilgiler, öğrenci kartları ve kişisel nöbet özetleri'),
    'Nöbet Hazırla':('NH','Nöbet çalışma alanı','Tarih seçin, listeyi hazırlayın ve sonuçları birlikte yönetin'),
    'Nöbet Geçmişi':('NG','Nöbet geçmişi','Geçmiş görevleri arayın ve kayıtları dışa aktarın'),
    'İstatistikler':('İS','İstatistikler','Tür ve saat filtreleriyle tablo veya grafik görünümünü inceleyin'),
    'Cezalılar':('CZ','Ceza yönetimi','Güncel cezalar, arşiv ve öğrenci bazında ceza istatistikleri'),
    'Muafiyetler':('MF','Muafiyetler','Görev ve tarih aralığına göre muafiyetleri yönetin'),
    'Sınıf Yasakları':('SY','Sınıf kuralları','Giriş nöbetine yazılamayan sınıfları düzenleyin'),
    'Yedekleme':('YD','Yedekleme ve veri yönetimi','Yedek alın veya seçtiğiniz kayıtları kontrollü biçimde sıfırlayın'),
    'Asistan':('AS','Yapay zekâ asistanı','Program hakkında soru sor veya bir güncelleme hazırla')}

class Workspace(ttk.Frame):
    """Fixed-size side navigation; the native notebook only manages page state."""
    def __init__(self,parent):
        super().__init__(parent);self.buttons={};self.names={};self.compact=False
        self.sidebar=tk.Frame(self,bg='#201a39',width=194);self.sidebar.pack(side='left',fill='y');self.sidebar.pack_propagate(False)
        top=tk.Frame(self.sidebar,bg='#201a39',height=52);top.pack(fill='x');top.pack_propagate(False)
        self.toggle_button=tk.Button(top,text='☰  Menüyü daralt',command=self.toggle,bg='#201a39',fg='#d6cde8',activebackground='#342b54',activeforeground='white',bd=0,font=('Segoe UI',10),anchor='w',padx=15,cursor='hand2');self.toggle_button.pack(fill='both',expand=True)
        self.nav=tk.Frame(self.sidebar,bg='#201a39');self.nav.pack(fill='x',padx=8)
        self.footer=tk.Label(self.sidebar,text='Veritabanı bu bilgisayarda',bg='#201a39',fg='#bdb1d7',font=('Segoe UI',8),justify='left',anchor='w',padx=12,pady=16);self.footer.pack(side='bottom',fill='x')
        right=ttk.Frame(self,style='Canvas.TFrame');right.pack(side='left',fill='both',expand=True)
        heading=ttk.Frame(right,style='Canvas.TFrame',padding=(22,16));heading.pack(fill='x')
        self.title=ttk.Label(heading,text='',style='Canvas.TLabel',font=('Segoe UI',16,'bold'),foreground='#241a3d');self.title.pack(anchor='w')
        self.subtitle=ttk.Label(heading,text='',style='Canvas.TLabel',font=('Segoe UI',9),foreground='#615776');self.subtitle.pack(anchor='w',pady=(3,0))
        self.notebook=ttk.Notebook(right,style='Workspace.TNotebook');self.notebook.pack(fill='both',expand=True,padx=12,pady=(6,12))
        self.notebook.bind('<<NotebookTabChanged>>',self.sync)
        self.bind('<Configure>',self.resize)
        self.manual=False
    def add(self,page,text):
        self.notebook.add(page,text=text);self.names[str(page)]=text
        slot=tk.Frame(self.nav,bg='#201a39',height=47);slot.pack(fill='x',pady=3);slot.pack_propagate(False)
        b=tk.Button(slot,text=PAGE_DETAILS[text][0]+'   '+text,command=lambda:self.select(page),bg='#201a39',fg='#d6cde8',activebackground='#342b54',activeforeground='white',bd=0,relief='flat',font=('Segoe UI',10),anchor='w',padx=12,cursor='hand2',takefocus=True)
        b.pack(fill='both',expand=True);self.buttons[str(page)]=b
        def hint(event):
            if not self.compact:return
            self.subtitle.configure(text=text+' — '+PAGE_DETAILS[text][2])
        b.bind('<Enter>',hint);b.bind('<Leave>',lambda event:self.sync())
    def select(self,page=None):
        if page is None:return self.notebook.select()
        self.notebook.select(page);self.sync()
    def sync(self,event=None):
        current=self.notebook.select()
        for page,button in self.buttons.items():
            active=page==current;button.configure(bg='#5540b5' if active else '#201a39',fg='white' if active else '#d6cde8',activebackground='#443294' if active else '#342b54')
        if current in self.names:
            detail=PAGE_DETAILS[self.names[current]];self.title.configure(text=detail[1]);self.subtitle.configure(text=detail[2])
    def set_compact(self,value):
        self.compact=value;self.sidebar.configure(width=66 if value else 194)
        self.toggle_button.configure(text='☰' if value else '☰  Menüyü daralt')
        self.footer.configure(text='' if value else 'Veritabanı bu bilgisayarda')
        for page,b in self.buttons.items():
            name=self.names[page];b.configure(text=PAGE_DETAILS[name][0] if value else PAGE_DETAILS[name][0]+'   '+name,anchor='center' if value else 'w',padx=0 if value else 12)
    def toggle(self):self.manual=True;self.set_compact(not self.compact)
    def resize(self,event):
        if event.widget is self and not self.manual:
            desired=event.width<1200
            if desired!=self.compact:self.set_compact(desired)
