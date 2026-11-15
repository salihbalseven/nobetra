"""Turkish calendar popup using only the standard Tkinter library."""
import calendar
import datetime as dt
import tkinter as tk
from tkinter import ttk

MONTHS=['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran','Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
WEEKDAYS=['Pazartesi','Salı','Çarşamba','Perşembe','Cuma','Cumartesi','Pazar']

def month_shift(year,month,offset):
    index=(year-1)*12+month-1+offset
    if not 0<=index<9999*12:return year,month
    year,month=divmod(index,12)
    return year+1,month+1

class DateField(ttk.Frame):
    def __init__(self,parent,variable,width=13,on_change=None,allow_empty=False):
        super().__init__(parent)
        self.variable=variable;self.on_change=on_change;self.allow_empty=allow_empty;self.popup=None
        self.entry=ttk.Entry(self,textvariable=variable,width=width)
        self.entry.pack(side='left',fill='x',expand=True)
        ttk.Button(self,text='▦',width=3,command=self.open_calendar).pack(side='left',padx=(2,0))
        self.entry.bind('<Button-1>',lambda event:self.open_calendar())
        self.entry.bind('<Down>',self.open_calendar);self.entry.bind('<F4>',self.open_calendar)
        self.entry.bind('<Escape>',lambda event:self.popup.close() if self.popup is not None and self.popup.winfo_exists() else None)
    def open_calendar(self,event=None):
        if self.popup is not None and self.popup.winfo_exists():self.popup.lift();return 'break' if event else None
        self.popup=CalendarPopup(self)
        return 'break' if event else None

class CalendarPopup(tk.Toplevel):
    def __init__(self,field,show=True):
        super().__init__(field);self.withdraw();self.field=field;self.root=field.winfo_toplevel();self.binding=None;self.focus_job=None
        self.overrideredirect(True);self.transient(self.root);self.configure(bg='#cbd1da')
        try:self.selected=dt.date.fromisoformat(field.variable.get().strip())
        except ValueError:self.selected=None
        self.today=dt.date.today();date=self.selected or self.today;self.year=date.year;self.month=date.month
        outer=tk.Frame(self,bg='#fafafa',padx=14,pady=12);outer.pack(padx=1,pady=1)
        self.caption=tk.Label(outer,bg='#fafafa',fg='#202124',font=('Segoe UI',11),anchor='w')
        self.caption.pack(fill='x',pady=(0,12))
        header=tk.Frame(outer,bg='#fafafa');header.pack(fill='x',pady=(0,10))
        ttk.Button(header,text='‹',width=3,command=lambda:self.shift(-1)).pack(side='left')
        self.month_var=tk.StringVar(value=MONTHS[self.month-1]);self.year_var=tk.StringVar(value=str(self.year))
        self.month_box=ttk.Combobox(header,textvariable=self.month_var,values=MONTHS,state='readonly',width=10)
        self.month_box.pack(side='left',padx=(5,3))
        years=sorted(set(range(max(1,min(self.year,self.today.year)-100),min(9999,max(self.year,self.today.year)+100)+1)))
        self.year_box=ttk.Combobox(header,textvariable=self.year_var,values=years,state='readonly',width=6)
        self.year_box.pack(side='left',padx=(0,5))
        ttk.Button(header,text='›',width=3,command=lambda:self.shift(1)).pack(side='right')
        self.month_box.bind('<<ComboboxSelected>>',self.change_month);self.year_box.bind('<<ComboboxSelected>>',self.change_month)
        self.grid_frame=tk.Frame(outer,bg='#fafafa');self.grid_frame.pack()
        footer=tk.Frame(outer,bg='#fafafa');footer.pack(fill='x',pady=(10,0))
        ttk.Button(footer,text='Bugün',command=lambda:self.choose(self.today)).pack(side='left')
        if field.allow_empty:ttk.Button(footer,text='Temizle',command=lambda:self.choose(None)).pack(side='left',padx=6)
        ttk.Button(footer,text='Kapat',command=self.close).pack(side='right')
        self.bind('<Escape>',lambda event:self.close());self.bind('<FocusOut>',self.schedule_focus_check)
        self.draw()
        if show:
            self.update_idletasks();x=field.winfo_rootx();y=field.winfo_rooty()+field.winfo_height()+3
            x=max(0,min(x,self.winfo_screenwidth()-self.winfo_reqwidth()-8))
            if y+self.winfo_reqheight()>self.winfo_screenheight():y=max(0,field.winfo_rooty()-self.winfo_reqheight()-3)
            self.geometry(f'+{x}+{y}');self.deiconify();self.lift();self.focus_set()
            self.binding=self.root.bind('<ButtonPress-1>',self.outside_click,add='+')
    def draw(self):
        for child in self.grid_frame.winfo_children():child.destroy()
        active=self.selected
        self.caption.configure(text=f'{active.day} {MONTHS[active.month-1]} {WEEKDAYS[active.weekday()]}' if active else 'Tarih seçin')
        self.month_var.set(MONTHS[self.month-1]);self.year_var.set(str(self.year));self.day_buttons={}
        for column,text in enumerate(['Pt','Sa','Ça','Pe','Cu','Ct','Pa']):
            tk.Label(self.grid_frame,text=text,bg='#fafafa',fg='#5f6368',font=('Segoe UI',9),width=4).grid(row=0,column=column,pady=(0,6))
        for row,week in enumerate(calendar.Calendar(firstweekday=0).monthdayscalendar(self.year,self.month),1):
            for column,day in enumerate(week):
                if not day:continue
                date=dt.date(self.year,self.month,day);selected=date==self.selected;today=date==self.today
                button=tk.Button(self.grid_frame,text=str(day),width=4,height=2,relief='flat',bd=0,font=('Segoe UI',10),bg='#0067c0' if selected else ('#e8edf4' if today else '#fafafa'),fg='white' if selected else '#202124',activebackground='#dcecff',cursor='hand2',command=lambda date=date:self.choose(date))
                button.grid(row=row,column=column,padx=1,pady=1);self.day_buttons[day]=button
    def shift(self,offset):self.year,self.month=month_shift(self.year,self.month,offset);self.draw()
    def change_month(self,event=None):self.year=int(self.year_var.get());self.month=MONTHS.index(self.month_var.get())+1;self.draw()
    def choose(self,date):
        self.field.variable.set(date.isoformat() if date else '')
        self.close()
        if self.field.on_change:self.field.on_change()
    def outside_click(self,event):
        name=str(event.widget)
        if not name.startswith(str(self)) and not name.startswith(str(self.field)):self.close()
    def check_focus(self):
        self.focus_job=None
        if not self.winfo_exists():return
        focus=str(self.tk.call('focus'))
        if not focus.startswith(str(self)) and not focus.startswith(str(self.field)):self.close()
    def schedule_focus_check(self,event=None):
        if self.focus_job:self.after_cancel(self.focus_job)
        self.focus_job=self.after(100,self.check_focus)
    def close(self):
        if self.focus_job:self.after_cancel(self.focus_job);self.focus_job=None
        if self.binding:
            self.root.unbind('<ButtonPress-1>',self.binding);self.binding=None
        self.field.popup=None
        self.destroy()
