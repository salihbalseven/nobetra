"""Filter-consistent pie charts without external plotting dependencies."""
import tkinter as tk
from tkinter import ttk
import nobet_istatistik as stats

GROUPS=['Saat aralığı','Nöbet türü','Tür + saat','Sınıf']
MEASURES=['Nöbet sayısı','Toplam süre (saat)']
COLORS=['#2563eb','#16a34a','#f59e0b','#7c3aed','#0891b2','#e11d48','#0d9488','#9333ea','#ea580c','#0284c7','#65a30d','#db2777','#475569','#4f46e5','#b45309','#15803d']

def aggregate(c,start,end,selected,duration,include_penalties=False,group='Saat aralığı',measure='Nöbet sayısı'):
    selected=set(selected);values={}
    rows=c.execute('''SELECT s.kind,st.class,a.slot,a.start_at,a.end_at,a.duty_type
        FROM assignments a JOIN sheets s ON s.id=a.sheet_id JOIN students st ON st.id=a.student_id
        WHERE s.status='Onaylandı' AND a.role='Asıl' AND a.result='Tuttu' AND st.active=1
        AND COALESCE(NULLIF(substr(a.start_at,1,10),''),s.day) BETWEEN ? AND ?''',(start,end))
    for row in rows:
        if row['kind'] not in stats.SLOTS or (row['duty_type']=='Ceza' and not include_penalties):continue
        slot=stats.slot_key(row)
        if (row['kind'],slot) not in selected:continue
        label={'Saat aralığı':slot,'Nöbet türü':row['kind'],'Tür + saat':stats.LABELS[row['kind']]+' '+slot,'Sınıf':row['class']}[group]
        value=1 if measure=='Nöbet sayısı' else (0 if slot==stats.UNKNOWN else duration(row['kind'],row['slot'],row['start_at'],row['end_at']))
        if value>0:values[label]=values.get(label,0)+value
    return sorted(values.items(),key=lambda item:(-item[1],item[0]))

class PieView(ttk.Frame):
    def __init__(self,parent,on_change):
        super().__init__(parent);self.data=[];self.unit='nöbet';self.on_change=on_change
        controls=ttk.Frame(self);controls.pack(fill='x',pady=8)
        ttk.Label(controls,text='Dağılım:').pack(side='left');self.group=tk.StringVar(value=GROUPS[0])
        combo=ttk.Combobox(controls,textvariable=self.group,values=GROUPS,state='readonly',width=18);combo.pack(side='left',padx=8);combo.bind('<<ComboboxSelected>>',lambda event:on_change())
        ttk.Label(controls,text='Ölçü:').pack(side='left');self.measure=tk.StringVar(value=MEASURES[0])
        combo=ttk.Combobox(controls,textvariable=self.measure,values=MEASURES,state='readonly',width=22);combo.pack(side='left',padx=8);combo.bind('<<ComboboxSelected>>',lambda event:on_change())
        self.view=tk.StringVar(value='Pasta grafik')
        selector=ttk.Combobox(controls,textvariable=self.view,values=['Pasta grafik','Çubuk grafik'],state='readonly',width=16)
        selector.pack(side='left',padx=8)
        selector.bind('<<ComboboxSelected>>',lambda event:self.draw())
        self.detail=ttk.Label(self,text='');self.detail.pack(anchor='w',pady=5)
        container=ttk.Frame(self);container.pack(fill='both',expand=True)
        self.canvas=tk.Canvas(container,bg='white',highlightthickness=0,height=340)
        scroll=ttk.Scrollbar(container,orient='vertical',command=self.canvas.yview);scroll.pack(side='right',fill='y');self.canvas.configure(yscrollcommand=scroll.set);self.canvas.pack(side='left',fill='both',expand=True)
        self.canvas.bind('<Configure>',lambda event:self.draw())
    def set_data(self,data,measure):self.data=data;self.unit='saat' if measure==MEASURES[1] else 'nöbet';self.draw()
    def draw(self):
        canvas=self.canvas;canvas.delete('all');width=max(canvas.winfo_width(),650);height=max(canvas.winfo_height(),280)
        total=sum(value for label,value in self.data)
        if not total:
            self.detail.config(text='Seçilen tarih ve filtrelerde grafik için kayıt yok.')
            canvas.create_text(width/2,100,text='Grafik için eşleşen veri bulunamadı',fill='#64748b',font=('Segoe UI',13));return
        if self.view.get()=='Çubuk grafik':
            self.draw_bars(width,height,total)
            return
        summary=f'Toplam {total:g} {self.unit} · {len(self.data)} dilim · üstteki filtrelere göre'
        self.detail.config(text=summary);size=min(300,height-55,width*.43);left=25;top=25;angle=90
        canvas.create_text(left+size/2,top+size+22,text=f'{total:g} {self.unit}',fill='#0f172a',font=('Segoe UI',13,'bold'))
        legend_x=left+size+35
        for index,(label,value) in enumerate(self.data):
            color=COLORS[index%len(COLORS)];extent=360*value/total
            tag=f'slice{index}'
            if extent>359.999:canvas.create_oval(left,top,left+size,top+size,fill=color,outline='white',width=2,tags=tag)
            else:canvas.create_arc(left,top,left+size,top+size,start=angle,extent=-extent,fill=color,outline='white',width=2,tags=tag)
            angle-=extent;y=24+index*26
            canvas.create_rectangle(legend_x,y,legend_x+13,y+13,fill=color,outline='')
            text=f'{label} — {value:g} {self.unit} (%{100*value/total:.1f})'
            canvas.create_text(legend_x+23,y+7,text=text,anchor='w',fill='#334155',font=('Segoe UI',10))
            canvas.tag_bind(tag,'<Enter>',lambda event,text=text:self.detail.config(text=text));canvas.tag_bind(tag,'<Leave>',lambda event:self.detail.config(text=summary))
        canvas.configure(scrollregion=(0,0,width,max(height,len(self.data)*26+35,top+size+50)))

    def draw_bars(self,width,height,total):
        self.detail.configure(text=f'Toplam {total:g} {self.unit} · {len(self.data)} kategori')
        maximum=max(value for _,value in self.data)
        left=262
        span=width-left-105
        for index,(label,value) in enumerate(self.data):
            y=25+index*35
            end=left+span*value/maximum
            self.canvas.create_text(left-10,y+10,text=label,anchor='e',fill='#334155',font=('Segoe UI',10))
            self.canvas.create_rectangle(left,y,end,y+20,fill=COLORS[index%len(COLORS)],outline='')
            self.canvas.create_text(end+8,y+10,text=f'{value:g} {self.unit}',anchor='w',fill='#0f172a',font=('Segoe UI',10))
        self.canvas.configure(scrollregion=(0,0,width,max(height,len(self.data)*35+35)))
