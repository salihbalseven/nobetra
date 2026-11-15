"""Persistent class queues and minimum-cost night-slot balancing within each class."""
import copy
import datetime as dt
import hashlib
import json
import sqlite3
import ceza_nobet
import excel_devriye
import nobet_yazdir

NIGHT=['23:00-01:00','01:00-03:00','03:00-05:00','05:00-07:00']
NIGHT_KINDS=('Yatakhane Kat','Devriye')
PENALTY_TIMES=ceza_nobet.HOURS
PENALTY_DEFAULTS=tuple(NIGHT[1:3])

def setup(c):
    c.execute('CREATE TABLE IF NOT EXISTS planner_settings(kind TEXT PRIMARY KEY,seed_id INTEGER NOT NULL REFERENCES students(id),initial_state TEXT NOT NULL)')
    c.execute('CREATE TABLE IF NOT EXISTS planner_runs(sheet_id INTEGER PRIMARY KEY REFERENCES sheets(id),before_state TEXT NOT NULL,after_state TEXT NOT NULL,signature TEXT NOT NULL)')
    c.execute('CREATE TABLE IF NOT EXISTS planner_dismissed_missed(assignment_id INTEGER PRIMARY KEY REFERENCES assignments(id) ON DELETE CASCADE)')
    c.execute('''CREATE TABLE IF NOT EXISTS planner_continuations(
        kind TEXT NOT NULL, day TEXT NOT NULL, seed_id INTEGER NOT NULL REFERENCES students(id),
        state TEXT NOT NULL, PRIMARY KEY(kind,day))''')

def groups(app,c):
    result={cls:[] for cls in app.CLASSES}
    for row in c.execute('SELECT * FROM students ORDER BY id'):
        cls=next((cls for cls in app.CLASSES if app.excel_yatakhane.norm(cls)==app.excel_yatakhane.norm(row['class'])),None)
        if cls:result[cls].append(dict(row))
    return result

def configure_continuation(app,c,kind,day,seed_id,order,reverse_students):
    dt.date.fromisoformat(day)
    if kind not in NIGHT_KINDS:
        raise ValueError('Devam sırası Yatakhane Kat veya Devriye için ayarlanabilir.')
    if len(order)!=len(app.CLASSES) or set(order)!=set(app.CLASSES):
        raise ValueError('Her sınıfı sıraya bir kez yazın. Sınıfları tireyle ayırın.')
    if c.execute("""SELECT 1 FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id
                    WHERE s.kind=? AND s.day>=?
                    AND EXISTS (SELECT 1 FROM assignments a WHERE a.sheet_id=s.id)""",(kind,day)).fetchone():
        raise ValueError('Bu tarih veya sonrasında otomatik liste var. Devam sırası için daha ileri bir tarih seçin.')
    pool=groups(app,c)
    if reverse_students:
        for students in pool.values():
            students.reverse()
    for turn,klass in enumerate(order):
        index=next((i for i,st in enumerate(pool[klass]) if st['id']==seed_id and st['active']),None)
        if index is not None:
            break
    else:
        raise ValueError('Başlangıç öğrencisi aktif sınıf listesinde bulunamadı.')
    pointers={klass:len(pool[klass]) if i<turn else 0 for i,klass in enumerate(order)}
    pointers[order[turn]]=index
    state=dict(turn=turn,pointers=pointers,class_order=order,reverse_students=bool(reverse_students))
    c.execute('''INSERT INTO planner_continuations VALUES(?,?,?,?)
                 ON CONFLICT(kind,day) DO UPDATE SET seed_id=excluded.seed_id,state=excluded.state''',
              (kind,day,seed_id,json.dumps(state)))
    app.log(c,'Otomatik sıra devam ayarı',f'{kind}; {day}; öğrenci {seed_id}; {"-".join(order)}')

def reorder_night_classes(app,c,kind,day,priority):
    dt.date.fromisoformat(day)
    if kind not in NIGHT_KINDS or not priority or len(set(priority))!=len(priority) or not set(priority)<=set(app.CLASSES):
        raise ValueError('Öne alınacak sınıfları tekrarsız ve tireyle ayırarak yazın.')
    if c.execute("""SELECT 1 FROM sheets s JOIN assignments a ON a.sheet_id=s.id
        WHERE s.kind=? AND (s.day=? OR (s.day>? AND EXISTS(
        SELECT 1 FROM planner_runs r WHERE r.sheet_id=s.id)))""",(kind,day,day)).fetchone():
        raise ValueError('Bu tarihte veya sonrasında sıraya bağlı bir liste var. İlk boş tarihi seçin.')
    state,_,_=queue_state(app,c,kind,day)
    if not state:
        raise ValueError('Önce nöbetin sıra başlangıcını ayarlayın.')
    order=state.get('class_order',app.CLASSES)
    cycle,index=divmod(state['turn'],len(order))
    pending=order[index:]+order[:index]
    pool=ordered_pool(app,c,state)
    current,finish_cycle=state.get('finish_class',(order[index],cycle))
    members=pool[current]
    progress=state['pointers'][current]-finish_cycle*len(members)
    started=bool(state.get('finish_class')) or (
        0<progress<len(members) and any(st['active'] for st in members[progress:]))
    first_classes=[current] if started else []
    first_classes.extend(klass for klass in priority if klass not in first_classes)
    state['class_order']=first_classes+[klass for klass in pending if klass not in first_classes]
    state['turn']=cycle*len(order)
    state['class_order_first']=True
    if started:
        state['finish_class']=[current,finish_cycle]
    first=first_in_queue(app,c,state)
    if first is None:
        raise ValueError('Sırada aktif öğrenci bulunamadı.')
    c.execute('''INSERT INTO planner_continuations VALUES(?,?,?,?)
        ON CONFLICT(kind,day) DO UPDATE SET seed_id=excluded.seed_id,state=excluded.state''',
        (kind,day,first['id'],json.dumps(state)))
    app.log(c,'Sınıf önceliği değiştirildi',f'{kind} · {day} · {"-".join(state["class_order"])}')
    return current if started else None

def configure(app,c,seeds):
    setup(c);pool=groups(app,c)
    for kind,sid in seeds.items():
        previous=c.execute('SELECT seed_id FROM planner_settings WHERE kind=?',(kind,)).fetchone()
        if previous and previous[0]==sid:continue
        if c.execute('SELECT 1 FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id WHERE s.kind=? AND s.status=?',(kind,'Onaylandı')).fetchone():
            raise ValueError(f'{kind}: onaylanmış otomatik listeler var; başlangıç noktası değiştirilemez.')
        if c.execute('SELECT 1 FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id JOIN assignments a ON a.sheet_id=s.id WHERE s.kind=?',(kind,)).fetchone():
            raise ValueError(f'{kind}: önce mevcut otomatik taslağı temizleyin.')
        found=[(cls,i) for cls,students in pool.items() for i,st in enumerate(students) if st['id']==sid and st['active']]
        if len(found)!=1:raise ValueError('Başlangıç öğrencisi aktif listede bulunamadı')
        cls,index=found[0];turn=app.CLASSES.index(cls)
        pointers={cl:(len(pool[cl]) if app.CLASSES.index(cl)<turn else 0) for cl in app.CLASSES};pointers[cls]=index
        state={'turn':turn,'pointers':pointers}
        c.execute('INSERT INTO planner_settings VALUES(?,?,?) ON CONFLICT(kind) DO UPDATE SET seed_id=excluded.seed_id,initial_state=excluded.initial_state',(kind,sid,json.dumps(state)))
    app.log(c,'Otomatik sıra başlangıçları',json.dumps(seeds))

def configure_gate_classes(app,c,seeds):
    if c.execute("SELECT 1 FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id JOIN assignments a ON a.sheet_id=s.id WHERE s.kind='Yatakhane Giriş'").fetchone():
        raise ValueError('Giriş için otomatik listeler var. İlk sınıf başlangıçları yalnızca otomatik planlamaya başlamadan önce değiştirilebilir.')
    row=c.execute("SELECT * FROM planner_settings WHERE kind='Yatakhane Giriş'").fetchone();state=json.loads(row['initial_state']);pool=groups(app,c);main=state['turn']%len(app.CLASSES);seed_id=row['seed_id']
    for cls,sid in seeds.items():
        found=[i for i,st in enumerate(pool[cls]) if st['id']==sid and st['active']]
        if len(found)!=1:raise ValueError(f'{cls}: aktif öğrenci bulunamadı')
        cycle=state['turn']//len(app.CLASSES)+(app.CLASSES.index(cls)<main)
        state['pointers'][cls]=cycle*len(pool[cls])+found[0]
        if app.CLASSES.index(cls)==main:seed_id=sid
    c.execute("UPDATE planner_settings SET initial_state=?,seed_id=? WHERE kind='Yatakhane Giriş'",(json.dumps(state),seed_id))
    app.log(c,'Giriş sınıf başlangıçları',json.dumps(seeds))

def ensure_settings(app,c):
    setup(c)
    for kind in app.TYPES:
        if not c.execute('SELECT 1 FROM planner_settings WHERE kind=?',(kind,)).fetchone():
            raise ValueError(f'{kind}: Sıra başlangıçları ekranından ilk öğrenciyi seçin.')

def signature(rows):
    values=sorted((int(r['student_id']),r['slot'],r['location'],r['role'],r.get('start_at') or '',r.get('end_at') or '')+((r['penalty_id'],) if r.get('duty_type')=='Ceza' and r.get('penalty_id') is not None else ()) for r in rows)
    return hashlib.sha256(json.dumps(values,ensure_ascii=False).encode()).hexdigest()

def before_confirm(c,sheet):
    run=c.execute('SELECT * FROM planner_runs WHERE sheet_id=?',(sheet['id'],)).fetchone()
    if run:
        rows=[dict(r) for r in c.execute('SELECT * FROM assignments WHERE sheet_id=?',(sheet['id'],))]
        if signature(rows)!=run['signature']:raise ValueError('Otomatik listenin kişi/saat/yerleri elle değiştirilmiş. Sıranın yanlış ilerlememesi için taslağı temizleyip otomatik listeyi yeniden hazırlayın.')

def before_unconfirm(c,sheet):
    if c.execute('SELECT 1 FROM planner_runs WHERE sheet_id=?',(sheet['id'],)).fetchone():
        if c.execute("SELECT 1 FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id WHERE s.day>? AND s.status='Onaylandı'",(sheet['day'],)).fetchone():
            raise ValueError('Bu otomatik listeden sonra onaylanmış otomatik listeler var. Sırayı korumak için önce en yeni tarihten başlayarak onayları kaldırın.')

def state_for(app,c,kind,day):
    for run in c.execute("SELECT r.* FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id WHERE s.kind=? AND s.status='Onaylandı'",(kind,)).fetchall():
        rows=[dict(r) for r in c.execute('SELECT * FROM assignments WHERE sheet_id=?',(run['sheet_id'],))]
        if signature(rows)!=run['signature']:raise ValueError(f'{kind}: onaylı otomatik listelerden biri sonradan değiştirilmiş. Yeni sıra hazırlamadan önce ilgili listeyi kontrol edin.')
    if c.execute("SELECT 1 FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id JOIN assignments a ON a.sheet_id=s.id WHERE s.kind=? AND s.day<? AND s.status!='Onaylandı'",(kind,day)).fetchone():
        raise ValueError(f'{kind}: daha eski otomatik taslak bekliyor. Önce onaylayın veya taslağı temizleyin.')
    if c.execute("SELECT 1 FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id WHERE s.kind=? AND s.day>=? AND s.status='Onaylandı'",(kind,day)).fetchone():
        raise ValueError(f'{kind}: otomatik listeleri tarih sırasıyla hazırlayın; bu tarihte veya sonrasında onaylı otomatik liste var.')
    state,_,_=queue_state(app,c,kind,day)
    if state is None:
        raise ValueError(f'{kind}: Sıra başlangıçlarından ilk öğrenciyi seçin.')
    return state

def queue_state(app,c,kind,day):
    latest=c.execute("SELECT s.id,s.day,r.before_state,r.after_state FROM planner_runs r JOIN sheets s ON s.id=r.sheet_id WHERE s.kind=? AND s.status='Onaylandı' AND s.day<? ORDER BY s.day DESC,s.id DESC LIMIT 1",(kind,day)).fetchone()
    continuation=c.execute('SELECT day,state FROM planner_continuations WHERE kind=? AND day<=? ORDER BY day DESC LIMIT 1',(kind,day)).fetchone()
    if continuation and (not latest or continuation['day']>latest['day']):
        state=json.loads(continuation['state'])
        anchor=continuation['day']
        source='Kayıtlı devam ayarı'
    elif latest:
        state=json.loads(latest['after_state'])
        finish=json.loads(latest['before_state']).get('finish_class')
        if finish:
            missed=[r['student_id'] for r in c.execute("""SELECT a.student_id,st.class FROM assignments a
                JOIN students st ON st.id=a.student_id WHERE a.sheet_id=? AND a.role='Asıl'
                AND a.duty_type='Normal' AND a.result!='Tuttu' AND st.active=1""",(latest['id'],))
                if app.excel_yatakhane.norm(r['class'])==app.excel_yatakhane.norm(finish[0])]
            if missed:
                state['finish_class']=finish
                state['deferred']=list(dict.fromkeys([*missed,*state.get('deferred',[])]))
        anchor=latest['day']
        source='Onaylı otomatik liste'
    else:
        initial=c.execute('SELECT initial_state FROM planner_settings WHERE kind=?',(kind,)).fetchone()
        return (json.loads(initial['initial_state']) if initial else None,'','İlk sıra ayarı')
    if kind in NIGHT_KINDS:
        manual=c.execute("""SELECT s.day,a.student_id FROM sheets s JOIN assignments a ON a.sheet_id=s.id
            WHERE s.kind=? AND s.status='Onaylandı' AND s.day>=? AND s.day<?
            AND a.role='Asıl' AND a.duty_type='Normal' AND (?=0 OR a.result='Tuttu')
            AND NOT EXISTS(SELECT 1 FROM planner_runs p WHERE p.sheet_id=s.id)
            ORDER BY s.day""",(kind,anchor,day,bool(state.get('class_order_first'))))
        written={}
        for row in manual:
            written.setdefault(row['day'],set()).add(row['student_id'])
        if written:
            pool=ordered_pool(app,c,state)
            for when,students in written.items():
                advance_manual(app,pool,state,students)
            anchor=when
            source='Onaylı manuel / Excel listesi'
    return state,anchor,source

def ordered_pool(app,c,state):
    pool=groups(app,c)
    if state.get('reverse_students'):
        for students in pool.values():
            students.reverse()
    return pool

def advance_manual(app,pool,state,written):
    classes=state.get('class_order',app.CLASSES)
    cycle=state['turn']//len(classes)
    deferred=state.setdefault('deferred',[])
    deferred[:]=[sid for sid in deferred if sid not in written]
    for klass in classes:
        members=pool[klass]
        size=len(members)
        start=state['pointers'][klass]
        window=range(start,(cycle+1)*size)
        last=max((ticket for ticket in window if members[ticket%size]['id'] in written),default=start-1)
        for ticket in range(start,last):
            student=members[ticket%size]
            if student['active'] and student['id'] not in written and student['id'] not in deferred:
                deferred.append(student['id'])
        state['pointers'][klass]=last+1
    advance_turn(classes,pool,state)

def advance_turn(classes,pool,state):
    n=len(classes)
    if not any(pool.values()):
        raise ValueError('Sınıf sırası ilerletilemedi; öğrenci listesini kontrol edin.')
    turn=state['turn']
    while True:
        klass=classes[turn%n]
        size=len(pool[klass])
        if size and state['pointers'][klass]<(turn//n+1)*size:
            state['turn']=turn
            return klass
        turn+=1

def first_in_queue(app,c,state):
    if not state:
        return None
    pool=ordered_pool(app,c,state)
    students={st['id']:st for members in pool.values() for st in members if st['active']}
    if not students:
        return None
    for sid in state.get('deferred',()):
        if sid in students:
            return students[sid]
    state=copy.deepcopy(state)
    classes=state.get('class_order',app.CLASSES)
    advance_turn(classes,pool,state)
    for offset in range(len(classes)):
        klass=classes[(state['turn']+offset)%len(classes)]
        members=pool[klass]
        for i in range(len(members)):
            student=members[(state['pointers'][klass]+i)%len(members)]
            if student['active']:
                return student
    return None

def take(app,c,pool,state,day,kind,time,used,blocked,notes,busy,eligibility,resting=()):
    deferred=state.setdefault('deferred',[])
    classes=state.get('class_order',app.CLASSES);n=len(classes)
    class_order_first=state.get('class_order_first',False)
    main=advance_turn(classes,pool,state) if class_order_first else None
    finish=state.get('finish_class')
    if finish:
        locked,cycle=finish
        size=len(pool[locked])
        pending={st['id'] for st in pool[locked] if st['active'] and st['id'] in deferred}
        pending.update(st['id'] for st in pool[locked][max(0,state['pointers'][locked]-cycle*size):] if st['active'])
        if pending-used:
            main=locked
        else:
            state.pop('finish_class')
            finish=None
    for sid in list(deferred):
        found=[st for cls,students in pool.items() if app.excel_yatakhane.norm(cls) not in blocked for st in students if st['id']==sid and st['active']]
        if (found and (not class_order_first or app.excel_yatakhane.norm(found[0]['class'])==app.excel_yatakhane.norm(main))
                and sid not in used and sid not in busy and not app.eligible(c,found[0],day,kind,time)):
            deferred.remove(sid);used.add(sid);return found[0]
    if main is None:
        main=advance_turn(classes,pool,state)
    turn=state['turn']
    candidates=[]
    for offset in range(n):
        cl=classes[(turn+offset)%n]
        if finish and cl!=main:continue
        if app.excel_yatakhane.norm(cl) in blocked or not pool[cl]:continue
        pointer=state['pointers'][cl];size=len(pool[cl])
        remaining=max(0,(turn//n+1)*size-pointer) if class_order_first else size
        for delta in range(remaining):
            ticket=pointer+delta;student=pool[cl][ticket%size]
            if not student['active']:continue
            if student['id'] in used:continue
            key=(student['id'],time)
            if key not in eligibility:eligibility[key]=app.eligible(c,student,day,kind,time)
            if eligibility[key]:continue
            if student['id'] in busy:continue
            # Preserve the waiting class; fill gaps using classes earliest in their cycle.
            priority=(-1 if offset==0 else ticket//size,offset)
            candidates.append((priority,cl,ticket,student));break
    if not candidates:
        if finish:raise ValueError(f'{kind}: {main} sınıfı tamamlanmadan yeni öncelik sırasına geçilemez.')
        raise ValueError(f'{kind}: sınıfı bölmeden {time} için yeterli uygun öğrenci yok. Hiçbir liste kaydedilmedi.')
    _,cl,ticket,student=min(candidates,key=lambda item:item[0])
    for old_ticket in range(state['pointers'][cl],ticket):
        old=pool[cl][old_ticket%len(pool[cl])]
        if old['active'] and (class_order_first or old['id'] in resting or ceza_nobet.managed(c,old['id'])) and old['id'] not in deferred:
            deferred.append(old['id'])
    if cl!=main:notes.add(f'{kind}: {main} bekliyor; uygun {cl} sınıfından devam edildi.')
    state['pointers'][cl]=ticket+1;used.add(student['id'])
    return student


def missed_primaries(c,kind,day):
    """Outstanding missed normal duties in this duty type, oldest first."""
    return list(c.execute("""
        WITH latest AS (
            SELECT a.student_id,a.result,s.status,s.day,a.id,
                   ROW_NUMBER() OVER (PARTITION BY a.student_id ORDER BY s.day DESC,a.id DESC) AS position
            FROM assignments a JOIN sheets s ON s.id=a.sheet_id
            WHERE s.kind=? AND s.day<? AND a.role='Asıl' AND a.duty_type='Normal'
        )
        SELECT student_id,day,id FROM latest
        WHERE position=1 AND status='Onaylandı' AND result='Tutmadı'
          AND NOT EXISTS (SELECT 1 FROM planner_dismissed_missed d WHERE d.assignment_id=latest.id)
        ORDER BY day,id
    """,(kind,day)))

def dismiss_missed(app,kind,day,assignment_id):
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        if not any(row['id']==assignment_id for row in missed_primaries(c,kind,day)):
            raise ValueError('Tekrar önerisi değişmiş. Pencereyi kapatıp listeyi yeniden hazırlayın.')
        c.execute('INSERT INTO planner_dismissed_missed(assignment_id) VALUES(?)',(assignment_id,))
        app.log(c,'Tekrar önerisinden çıkarıldı',f'{kind}; görev kaydı {assignment_id}')

def hungarian(cost):
    """Square integer minimum-cost assignment, O(n^3), without external dependencies."""
    n=len(cost);u=[0]*(n+1);v=[0]*(n+1);p=[0]*(n+1);way=[0]*(n+1)
    for i in range(1,n+1):
        p[0]=i;j0=0;minimum=[10**18]*(n+1);used=[False]*(n+1)
        while True:
            used[j0]=True;i0=p[j0];delta=10**18;j1=0
            for j in range(1,n+1):
                if not used[j]:
                    value=cost[i0-1][j-1]-u[i0]-v[j]
                    if value<minimum[j]:minimum[j]=value;way[j]=j0
                    if minimum[j]<delta:delta=minimum[j];j1=j
            for j in range(n+1):
                if used[j]:u[p[j]]+=delta;v[j]-=delta
                else:minimum[j]-=delta
            j0=j1
            if p[j0]==0:break
        while True:
            j1=way[j0];p[j0]=p[j1];j0=j1
            if j0==0:break
    result=[0]*n
    for j in range(1,n+1):result[p[j]-1]=j-1
    return result

def night_history(c,kind,day,student_ids):
    counts={sid:[0]*len(NIGHT) for sid in student_ids};last={}
    history=c.execute("""SELECT a.student_id,a.slot,a.start_at,a.end_at,s.day FROM assignments a
        JOIN sheets s ON s.id=a.sheet_id WHERE s.kind=? AND s.status='Onaylandı'
        AND a.role='Asıl' AND a.result='Tuttu' AND a.duty_type='Normal'
        AND s.day<? ORDER BY s.day,a.start_at,a.id""",(kind,day))
    for row in history:
        if row['student_id'] not in counts:continue
        slot=nobet_yazdir.slot(dict(row))
        if slot not in NIGHT:continue
        index=NIGHT.index(slot);counts[row['student_id']][index]+=1;last[row['student_id']]=index
    return counts,last

def balance(students,capacity,counts,last):
    capacities=[capacity]*4 if isinstance(capacity,int) else capacity
    slots=[time for time, places in enumerate(capacities) for _ in range(places)]
    by_class={}
    for student in students:
        by_class.setdefault(student.get('class'),[]).append(student)

    result=[]
    offset=0
    for members in by_class.values():
        class_slots=slots[offset:offset+len(members)]
        # Keep each class's share of the timeline fixed. Only its students trade hours.
        costs=[
            [counts[student['id']][time]*100000
             +(1000 if last.get(student['id'])==time else 0)
             +((time-index%4)%4)*10 for time in class_slots]
            for index,student in enumerate(members,start=offset)
        ]
        allocation=hungarian(costs)
        result.extend((student,class_slots[column]) for student,column in zip(members,allocation))
        offset+=len(members)
    return result

def make_row(student,day,time,location,role):
    start=end=''
    if role=='Asıl':
        a,b=time.split('-');start=dt.datetime.combine(dt.date.fromisoformat(day),dt.time.fromisoformat(a))
        end=dt.datetime.combine(start.date(),dt.time.fromisoformat(b))
        if end<=start:end+=dt.timedelta(days=1)
        start=start.isoformat(timespec='minutes');end=end.isoformat(timespec='minutes')
    return dict(student_id=student['id'],name=student['name'],surname=student['surname'],school_no=student['school_no'],klass=student['class'],slot=time,location=location,role=role,start_at=start,end_at=end)

class PenaltyPlacementError(ValueError):
    pass


def resting_reserves(c, day, kind):
    yesterday = (dt.date.fromisoformat(day)-dt.timedelta(days=1)).isoformat()
    return {r[0] for r in c.execute("""
        SELECT a.student_id FROM assignments a JOIN sheets s ON s.id=a.sheet_id
        WHERE a.role='Yedek' AND a.result='Tuttu'
        AND (s.day=? OR (s.kind=? AND s.day=(
            SELECT MAX(prev.day) FROM sheets prev JOIN assignments pa ON pa.sheet_id=prev.id
            WHERE prev.kind=s.kind AND prev.day<?
        )))
    """,(yesterday,kind,day))}

def worked_previous_night(c,day):
    yesterday=(dt.date.fromisoformat(day)-dt.timedelta(days=1)).isoformat()
    return {row[0] for row in c.execute("""SELECT a.student_id FROM assignments a
        JOIN sheets s ON s.id=a.sheet_id WHERE s.day=?
        AND s.kind IN ('Yatakhane Kat','Devriye') AND a.result='Tuttu'""",(yesterday,))}


def build_one(app,c,day,kind,blocked,retry_choices=None,penalty_slots=None,excluded=(),
              *,projection=None):
    before=state_for(app,c,kind,day) if projection is None else projection[0]
    state=copy.deepcopy(before);used=set();notes=set();eligibility={}
    busy=app.busy_students(c,day,kind) if projection is None else set(projection[1])
    pool=ordered_pool(app,c,state)
    resting=resting_reserves(c,day,kind)
    busy.update(resting)
    if resting:
        notes.add(f'Önceki listede nöbet tutan {len(resting)} yedek bu listeye alınmadı.')
    resting.update(excluded)
    busy.update(excluded)
    if retry_choices is not None:
        skipped = {sid for (_,sid), include in retry_choices.items() if not include}
        resting.update(skipped)
        busy.update(skipped)
        busy.update(sid for (target,sid), include in retry_choices.items() if include and target!=kind)
    capacity=8 if kind=='Yatakhane Kat' else len(excel_devriye.PLACES)
    count=4*capacity
    if kind=='Yatakhane Giriş':count=10
    penalty_choices=[];remaining=[capacity]*4
    if kind in NIGHT_KINDS:
        candidates=({pid:student for student,pid in ceza_nobet.candidates(app,c,day,kind,blocked)
                     if student['id'] not in busy} if penalty_slots is None or penalty_slots else {})
        requests=({pid:time for (target,pid),time in penalty_slots.items() if target==kind}
                  if penalty_slots is not None else
                  {pid:student.get('preferred_penalty_slot') for pid,student in sorted(
                      candidates.items(),key=lambda entry:not entry[1].get('preferred_penalty_slot'))})
        for pid,time in requests.items():
            student=candidates.get(pid)
            if student is None:
                person=c.execute('SELECT s.name,s.surname FROM penalties p JOIN students s ON s.id=p.student_id WHERE p.id=?',(pid,)).fetchone()
                label=' '.join(person) if person else 'Seçilen öğrenci'
                raise PenaltyPlacementError(f'{kind} · {label} bu listeye ceza nöbeti olarak yerleştirilemedi. Mevcut görevleri, dinlenme süresini ve yatakhane–devriye sınıf çakışmasını kontrol edin.')
            if time is None and penalty_slots is None:
                available=[i for i in (1,2) if remaining[i]>0]
                if not available:break
                index=max(available,key=lambda i:remaining[i])
            elif time not in PENALTY_TIMES or remaining[NIGHT.index(time)]<=0:
                raise PenaltyPlacementError(f'{kind} · {time} vardiyasında ceza nöbeti için yeterli yer yok.')
            else:
                index=NIGHT.index(time)
            remaining[index]-=1
            penalty_choices.append((student,pid,index));used.add(student['id'])
        count-=len(penalty_choices)
        previous_night=worked_previous_night(c,day)
        busy.update(previous_night)
        resting.update(previous_night)
        if previous_night:
            notes.add('Önceki gece Tuttu olarak işaretlenen öğrenciler normal asıl veya yedek göreve alınmadı.')
    eligible_pool = {student['id']: student for cls, members in pool.items()
                     if app.excel_yatakhane.norm(cls) not in blocked
                     for student in members if student['active']}
    missed = [record for record in missed_primaries(c,kind,day)
              if retry_choices is None or retry_choices.get((kind,record['student_id']),False)]
    retries = {}
    students=[]
    for i in range(count):
        time=app.TIMES[kind][i//2] if kind=='Yatakhane Giriş' else NIGHT[0]
        for record in missed:
            sid=record['student_id']
            candidate=eligible_pool.get(sid)
            finish=state.get('finish_class')
            if finish and candidate and app.excel_yatakhane.norm(candidate['class'])!=app.excel_yatakhane.norm(finish[0]):
                continue
            if candidate and sid not in used and sid not in busy and not app.eligible(c,candidate,day,kind,time):
                used.add(sid)
                retries[sid]=record['day']
                students.append(candidate)
                break
        else:
            students.append(take(app,c,pool,state,day,kind,time,used,blocked,notes,busy,eligibility,resting))
    after=copy.deepcopy(state);reserves=[]
    for _ in range(16 if kind=='Yatakhane Kat' else (excel_devriye.RESERVES if kind=='Devriye' else 4)):
        reserves.append(take(app,c,pool,state,day,kind,'Yedek',used,blocked,notes,busy,eligibility,resting))
    rows=[]
    if kind in NIGHT_KINDS:
        counts,last=night_history(c,kind,day,(st['id'] for st in students))
        allocation=balance(students,remaining,counts,last)
        for index,time in enumerate(NIGHT):
            daily=(dt.date.fromisoformat(day)+dt.timedelta(days=bool(index))).isoformat()
            normal=[(st,None) for st,t in allocation if t==index]
            penalties=[(st,pid) for st,pid,t in penalty_choices if t==index]
            group=normal+penalties if penalty_slots is None else penalties+normal
            places=['2A','2B','2C','3A','3B','3C','4A','4B'] if kind=='Yatakhane Kat' else excel_devriye.PLACES
            for (st,pid),place in zip(group,places):
                row=make_row(st,daily,time,place,'Asıl')
                if pid:row.update(duty_type='Ceza',penalty_id=pid)
                else:
                    sid=st['id']
                    row.update(slot_counts=counts[sid],last_slot=NIGHT[last[sid]] if sid in last else '')
                rows.append(row)
        places=['2A','2B','2C','3A','3B','3C','4A','4B']*2 if kind=='Yatakhane Kat' else [f'{i+1}. YEDEK' for i in range(excel_devriye.RESERVES)]
    else:
        rows.extend(make_row(st,day,app.TIMES[kind][i//2],'A BLOK' if i%2==0 else 'B BLOK','Asıl') for i,st in enumerate(students))
        places=['A BLOK','B BLOK','A BLOK','B BLOK']
    rows.extend(make_row(st,day,'Yedek',place,'Yedek') for st,place in zip(reserves,places))
    return dict(kind=kind,rows=rows,before=before,after=after,notes=sorted(notes),retries=retries)

def forecast_class_conflicts(app,c,today):
    states={kind:queue_state(app,c,kind,today.isoformat())[0] for kind in NIGHT_KINDS}
    if not all(states.values()):
        return []
    previous=worked_previous_night(c,today.isoformat())
    warnings=[]
    for offset in range(3):
        day=(today+dt.timedelta(days=offset)).isoformat()
        plans={}
        for kind in NIGHT_KINDS:
            setting=c.execute('SELECT state FROM planner_continuations WHERE kind=? AND day=?',(kind,day)).fetchone()
            if setting:
                states[kind]=json.loads(setting['state'])
            saved=[dict(row) for row in c.execute("""SELECT a.*,st.class AS klass,s.status,r.after_state
                FROM sheets s JOIN assignments a ON a.sheet_id=s.id JOIN students st ON st.id=a.student_id
                LEFT JOIN planner_runs r ON r.sheet_id=s.id WHERE s.day=? AND s.kind=?""",(day,kind))]
            if saved:
                after=copy.deepcopy(states[kind])
                if saved[0]['after_state']:
                    after=json.loads(saved[0]['after_state'])
                else:
                    written={r['student_id'] for r in saved if r['role']=='Asıl'
                             and r['duty_type']=='Normal'}
                    advance_manual(app,ordered_pool(app,c,after),after,written)
                plans[kind]=dict(rows=[r for r in saved if r['result']!='Tutmadı'],after=after)
            else:
                try:
                    plans[kind]=build_one(app,c,day,kind,set(),penalty_slots={},
                                          projection=(states[kind],previous))
                except ValueError:
                    plans[kind]=dict(rows=[],after=states[kind])
        classes=[{app.excel_yatakhane.norm(r['klass']) for r in plans[kind]['rows']} for kind in NIGHT_KINDS]
        overlap=classes[0]&classes[1]
        if overlap and offset:
            warnings.append((day,offset,[k for k in app.CLASSES if app.excel_yatakhane.norm(k) in overlap]))
        states={kind:plan['after'] for kind,plan in plans.items()}
        previous={r['student_id'] for plan in plans.values() for r in plan['rows']
                  if r.get('result')=='Tuttu' or (r['role']=='Asıl' and r.get('result','Planlandı')=='Planlandı')}
    return warnings

def night_targets(c,day):
    rows=[dict(r) for r in c.execute('''SELECT a.*,st.class,s.kind FROM assignments a
        JOIN sheets s ON s.id=a.sheet_id JOIN students st ON st.id=a.student_id
        WHERE s.day=? AND s.kind IN ('Yatakhane Kat','Devriye') ORDER BY a.id''',(day,))]
    if any(row['kind']=='Devriye' for row in rows):
        raise ValueError('Bu tarihin devriye listesi dolu. Otomatik hazırlamadan önce mevcut devriye taslağını kontrol edin.')
    return rows


def penalty_options(app,c,day,kind,excluded=()):
    if kind not in NIGHT_KINDS:
        return {}
    dt.date.fromisoformat(day)
    kat=night_targets(c,day)
    blocked={app.excel_yatakhane.norm(r['class']) for r in kat}
    options={}
    for target in ('Devriye',) if kat else NIGHT_KINDS:
        resting=resting_reserves(c,day,target)
        for student,pid in ceza_nobet.candidates(app,c,day,target,blocked):
            if student['id'] not in resting and student['id'] not in excluded:
                total=c.execute('SELECT total FROM penalties WHERE id=?',(pid,)).fetchone()[0]
                _,reserved=ceza_nobet.counts(c,pid)
                last=c.execute("""SELECT a.slot FROM assignments a JOIN sheets s ON s.id=a.sheet_id
                    WHERE a.student_id=? AND a.duty_type='Ceza' AND a.role='Asıl'
                    AND a.result!='Tutmadı' AND s.day<?
                    ORDER BY s.day DESC,a.id DESC LIMIT 1""",(student['id'],day)).fetchone()
                student['last_penalty_slot']=last['slot'] if last else None
                options[target,pid]=(student,total-reserved)
    return options

def suggest_penalty_slots(candidates,keys,existing=None):
    choices=dict(existing or {})
    remaining=[]
    for key in keys:
        choices.pop(key,None)
        remaining.append(key)
    for key in remaining:
        kind,_=key
        student,_=candidates[key]
        limit=8 if kind=='Yatakhane Kat' else len(excel_devriye.PLACES)
        loads={slot:sum(target==kind and chosen==slot for (target,_),chosen in choices.items())
               for slot in PENALTY_DEFAULTS}
        available=[slot for slot in PENALTY_DEFAULTS if loads[slot]<limit]
        if available:
            last=student.get('last_penalty_slot')
            preferred=student.get('preferred_penalty_slot')
            choices[key]=min(available,key=lambda slot:(slot==last,slot!=preferred if last not in PENALTY_DEFAULTS else False,loads[slot]))
    return choices


def build(app,c,day,kind,retry_choices=None,penalty_slots=None,excluded=()):
    dt.date.fromisoformat(day);ensure_settings(app,c);plans=[];preserved=''
    excluded=tuple(sorted(set(excluded)))
    already_listed=[]
    def existing(k):return [dict(r) for r in c.execute('SELECT a.*,st.class FROM assignments a JOIN sheets s ON s.id=a.sheet_id JOIN students st ON st.id=a.student_id WHERE s.day=? AND s.kind=? ORDER BY a.id',(day,k))]
    if kind=='Yatakhane Giriş':
        if existing(kind):raise ValueError('Bu tarihin giriş listesi dolu. Mevcut liste değiştirilmedi.')
        plans.append(build_one(app,c,day,kind,set(),retry_choices,excluded=excluded))
    else:
        kat=night_targets(c,day)
        targets=('Devriye',) if kat else NIGHT_KINDS
        if penalty_slots and any(target not in targets for target,_ in penalty_slots):
            raise PenaltyPlacementError('Seçilen nöbet türünün mevcut çizelgesine ceza eklenemedi.')
        if kat:
            preserved=signature(kat);blocked={app.excel_yatakhane.norm(r['class']) for r in kat}
            already_listed=[row['id'] for row in kat if row['student_id'] in excluded]
        else:
            plan=build_one(app,c,day,'Yatakhane Kat',set(),retry_choices,penalty_slots,excluded)
            plans.append(plan)
            blocked={app.excel_yatakhane.norm(r['klass']) for r in plan['rows']}
        patrol_slots=penalty_slots
        deferred=[]
        if penalty_slots is not None:
            patrol_slots={}
            for key,time in penalty_slots.items():
                target,pid=key
                student=c.execute("""SELECT st.name,st.surname,st.class FROM penalties p
                    JOIN students st ON st.id=p.student_id WHERE p.id=?""",(pid,)).fetchone() if target=='Devriye' else None
                if student and app.excel_yatakhane.norm(student['class']) in blocked:
                    deferred.append(f'{student["name"]} {student["surname"]} ({student["class"]})')
                else:
                    patrol_slots[key]=time
        patrol=build_one(app,c,day,'Devriye',blocked,retry_choices,patrol_slots,excluded)
        if deferred:
            patrol['notes'].append('Yatakhane Kat çakışması nedeniyle devriye cezası sonraki uygun geceye kaldı: '+', '.join(deferred))
        plans.append(patrol)
    return dict(day=day,kind=kind,plans=plans,preserved=preserved,retry_choices=retry_choices,
                penalty_slots=penalty_slots,excluded=excluded,already_listed=already_listed)

def save(app,preview):
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        fresh=build(app,c,preview['day'],preview['kind'],preview.get('retry_choices'),
                    preview.get('penalty_slots'),preview['excluded'])
        if fresh!=preview:raise ValueError('Önizlemeden sonra kayıtlar değişmiş. Yeniden önizleme hazırlayın; hiçbir liste eklenmedi.')
        if fresh['already_listed']:raise ValueError('Seçilen öğrenciler önce kayıtlı Yatakhane Kat listesinden çıkarılmalı.')
        for plan in fresh['plans']:
            kind=plan['kind'];day=fresh['day']
            c.execute('INSERT OR IGNORE INTO sheets(day,kind) VALUES(?,?)',(day,kind))
            sheet=c.execute('SELECT * FROM sheets WHERE day=? AND kind=?',(day,kind)).fetchone()
            if sheet['status']=='Onaylandı':raise ValueError('Onaylı listeye otomatik görev eklenemez')
            c.execute('DELETE FROM planner_runs WHERE sheet_id=?',(sheet['id'],))
            for row in plan['rows']:
                c.execute('INSERT INTO assignments(sheet_id,student_id,slot,location,role,start_at,end_at,source_cell,duty_type,penalty_id) VALUES(?,?,?,?,?,?,?,?,?,?)',(sheet['id'],row['student_id'],row['slot'],row['location'],row['role'],row['start_at'],row['end_at'],'Sıralı otomatik plan',row.get('duty_type','Normal'),row.get('penalty_id')))
            c.execute('INSERT INTO planner_runs VALUES(?,?,?,?)',(sheet['id'],json.dumps(plan['before']),json.dumps(plan['after']),signature(plan['rows'])))
            app.log(c,'Sıralı dengeli otomatik taslak',f'{day}; {kind}; {len(plan["rows"])} öğrenci')
    return sum(len(p['rows']) for p in preview['plans'])
