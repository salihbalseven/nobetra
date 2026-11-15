"""Backed-up reset and explicit daily workflow operations."""
import datetime as dt
import sqlite3
import nobet_manuel
from ogrenci_profili import CONTACT_FIELDS

RESET_LABELS={'history':'Nöbet listeleri ve geçmişi','queues':'Tüm otomatik nöbet sıraları (genel)',
    'penalties':'Ceza kayıtları (güncel ve arşiv)','exemptions':'Muafiyet / görevden muafiyet kayıtları',
    'profiles':'Ek öğrenci bilgileri (sicil, yatak, aile, eğitim, notlar)',
    'photos':'Öğrenci fotoğrafları','duties':'Bilgi amaçlı öğrenci görevleri',
    'rules':'Sınıf yasaklarını varsayılana döndür (genel)','audit':'İşlem günlüğü (genel)'}
PROFILE_FIELDS={'registry':'Sicil','bed':'Yatak no','university':'Üniversite','department':'Bölüm','mother':'Anne adı soyadı','father':'Baba adı soyadı','hometown':'Memleket','notes':'Notlar'}
PROFILE_FIELDS.update(CONTACT_FIELDS)

def reset_plan(c,options,sid=None,start='',end='',kinds=None,profile_fields=None):
    chosen={key for key,value in options.items() if value}
    if not chosen or chosen-set(RESET_LABELS):raise ValueError('Sıfırlanacak en az bir geçerli bilgi seçin')
    profile_fields=list(PROFILE_FIELDS) if profile_fields is None else list(profile_fields)
    if 'profiles' in chosen and (not profile_fields or set(profile_fields)-set(PROFILE_FIELDS)):raise ValueError('Sıfırlanacak ek öğrenci alanlarını seçin')
    if start:dt.date.fromisoformat(start)
    if end:dt.date.fromisoformat(end)
    if start and end and start>end:raise ValueError('Başlangıç bitişten sonra olamaz')
    if sid is not None and not c.execute('SELECT 1 FROM students WHERE id=?',(sid,)).fetchone():raise ValueError('Öğrenci bulunamadı')
    kinds=set(('Yatakhane Kat','Devriye','Yatakhane Giriş') if kinds is None else kinds)
    if 'history' in chosen and (not kinds or kinds-{'Yatakhane Kat','Devriye','Yatakhane Giriş'}):raise ValueError('Nöbet geçmişi için geçerli bir nöbet türü seçin')
    all_rows=list(c.execute('SELECT a.id,a.sheet_id,a.student_id,a.penalty_id,s.day,s.kind FROM assignments a JOIN sheets s ON s.id=a.sheet_id'))
    tasks=[r for r in all_rows if 'history' in chosen and (sid is None or r['student_id']==sid) and (not start or r['day']>=start) and (not end or r['day']<=end) and r['kind'] in kinds]
    aids={r['id'] for r in tasks};sheets={r['sheet_id'] for r in tasks}
    auto={r[0] for r in c.execute('SELECT sheet_id FROM planner_runs')}
    if sheets&auto and 'queues' not in chosen:raise ValueError('Seçilen geçmiş otomatik sıra kaydı içeriyor. Sıranın tutarlı kalması için “Tüm otomatik nöbet sıraları” seçeneğini de işaretleyin.')
    where=' WHERE student_id=?' if sid is not None else '';params=(sid,) if sid is not None else ()
    cases=[r[0] for r in c.execute('SELECT id FROM penalties'+where,params)] if 'penalties' in chosen else []
    if any(r['penalty_id'] in cases and r['id'] not in aids for r in all_rows):raise ValueError('Seçilen cezalara bağlı nöbetler korunuyor. Önce bu nöbetleri de kapsayan geçmişi seçin veya Ceza kayıtları seçimini kaldırın. Arşivlemek için Cezalılar ekranını kullanabilirsiniz.')
    counts={}
    if 'history' in chosen:counts['history']=len(tasks)
    for key,table in [('penalties','penalties'),('exemptions','exemptions'),('profiles','student_profiles'),('photos','student_profiles'),('duties','student_profiles')]:
        if key in chosen:counts[key]=c.execute('SELECT COUNT(*) FROM '+table+where,params).fetchone()[0]
    if 'penalties' in chosen:counts['penalties']+=c.execute('SELECT COUNT(*) FROM historical_penalty_awards'+where,params).fetchone()[0]
    if 'photos' in chosen:counts['photos']=c.execute('SELECT COUNT(*) FROM student_profiles'+where+(' AND ' if where else ' WHERE ')+'photo IS NOT NULL',params).fetchone()[0]
    if 'queues' in chosen:counts['queues']=c.execute('SELECT (SELECT COUNT(*) FROM planner_runs)+(SELECT COUNT(*) FROM planner_settings)+(SELECT COUNT(*) FROM planner_continuations)').fetchone()[0]
    if 'rules' in chosen:counts['rules']=c.execute('SELECT COUNT(*) FROM restrictions').fetchone()[0]
    if 'audit' in chosen:counts['audit']=c.execute('SELECT COUNT(*) FROM audit').fetchone()[0]
    return {'chosen':chosen,'aids':aids,'cases':cases,'counts':counts,'sid':sid,'profile_fields':profile_fields}

def reset_selected(app,options,sid=None,start='',end='',kinds=None,profile_fields=None):
    directory=app.DATA/'yedekler';directory.mkdir(exist_ok=True)
    backup=directory/('sifirlama_oncesi_'+dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.db')
    with app.conn() as c:
        version=c.execute('PRAGMA data_version').fetchone()[0]
        reset_plan(c,options,sid,start,end,kinds,profile_fields)
        target=sqlite3.connect(backup)
        try:
            c.backup(target)
            if target.execute('PRAGMA quick_check').fetchone()[0]!='ok':raise ValueError('Yedek doğrulanamadı. Hiçbir kayıt silinmedi.')
        finally:target.close()
        c.execute('BEGIN IMMEDIATE')
        if c.execute('PRAGMA data_version').fetchone()[0]!=version:raise ValueError('Yedek sırasında kayıtlar değişti; işlem yapılmadı. Tekrar deneyin.')
        plan=reset_plan(c,options,sid,start,end,kinds,profile_fields);chosen=plan['chosen']
        if 'queues' in chosen:
            c.execute('DELETE FROM planner_runs');c.execute('DELETE FROM planner_settings')
            c.execute('DELETE FROM planner_continuations')
        c.executemany('DELETE FROM assignments WHERE id=?',[(aid,) for aid in plan['aids']])
        if 'history' in chosen:c.execute('DELETE FROM sheets WHERE NOT EXISTS(SELECT 1 FROM assignments a WHERE a.sheet_id=sheets.id) AND NOT EXISTS(SELECT 1 FROM planner_runs p WHERE p.sheet_id=sheets.id)')
        c.executemany('DELETE FROM penalties WHERE id=?',[(pid,) for pid in plan['cases']])
        where=' WHERE student_id=?' if sid is not None else '';params=(sid,) if sid is not None else ()
        if 'penalties' in chosen:
            c.execute('''UPDATE assignments SET duty_type='Normal' WHERE id IN (
                SELECT h.assignment_id FROM historical_penalty_duties h JOIN historical_penalty_awards p
                USING(source,source_row)'''+where+')',params)
            c.execute('DELETE FROM historical_penalty_awards'+where,params)
        if 'exemptions' in chosen:c.execute('DELETE FROM exemptions'+where,params)
        if 'profiles' in chosen:
            c.execute('UPDATE student_profiles SET '+','.join(field+"=''" for field in plan['profile_fields'])+where,params)
        if 'photos' in chosen:c.execute("UPDATE student_profiles SET photo=NULL,photo_name=''"+where,params)
        if 'duties' in chosen:c.execute("UPDATE student_profiles SET official_duty=''"+where,params)
        if 'rules' in chosen:
            c.execute('DELETE FROM restrictions')
            for day,group in enumerate([app.CLASSES[:4],app.CLASSES[4:8],app.CLASSES[8:12],app.CLASSES[12:]],1):c.executemany('INSERT INTO restrictions VALUES(?,?)',[(day,x) for x in group])
        if 'audit' in chosen:c.execute('DELETE FROM audit')
        app.log(c,'Seçili bilgiler sıfırlandı',str(plan['counts'])+'; öğrenci: '+str(sid)+'; yedek: '+backup.name)
    return backup,plan['counts']

def reset_history(app):
    directory=app.DATA/'yedekler';directory.mkdir(exist_ok=True)
    backup=directory/('sifirlama_oncesi_'+dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.db')
    with app.conn() as c:
        version=c.execute('PRAGMA data_version').fetchone()[0]
        target=sqlite3.connect(backup)
        try:
            c.backup(target)
            if target.execute('PRAGMA quick_check').fetchone()[0]!='ok':raise ValueError('Otomatik yedek doğrulanamadı. Sıfırlama yapılmadı.')
        finally:target.close()
        c.execute('BEGIN IMMEDIATE')
        if c.execute('PRAGMA data_version').fetchone()[0]!=version:raise ValueError('Yedek sırasında başka pencerede kayıt değişti. Sıfırlama yapılmadı; diğer pencereyi kapatıp tekrar deneyin.')
        for table in ('planner_runs','assignments','sheets','planner_settings','planner_continuations','penalties','historical_penalty_awards','audit'):c.execute('DELETE FROM '+table)
        app.log(c,'Geçmiş ve sıralar sıfırlandı','Önceki kayıtların otomatik yedeği: '+backup.name)
    return backup

def pending(c):
    return [dict(row) for row in c.execute('''SELECT s.*,COUNT(a.id) AS persons,
        SUM(CASE WHEN a.role='Asıl' AND a.result='Planlandı' THEN 1 ELSE 0 END) AS undecided
        FROM sheets s JOIN assignments a ON a.sheet_id=s.id WHERE s.status!='Onaylandı'
        GROUP BY s.id ORDER BY s.day,s.kind''')]

def next_preparation_day(day,kind):
    value=dt.date.fromisoformat(day)+dt.timedelta(days=1)
    if kind=='Yatakhane Giriş':
        while value.weekday()>4:value+=dt.timedelta(days=1)
    return value.isoformat()

def blockers(c,target,kind):
    kinds=('Yatakhane Giriş','Yatakhane Giriş') if kind=='Yatakhane Giriş' else ('Yatakhane Kat','Devriye')
    return [dict(r) for r in c.execute('''SELECT DISTINCT s.* FROM sheets s JOIN planner_runs p ON p.sheet_id=s.id
        JOIN assignments a ON a.sheet_id=s.id WHERE s.status!='Onaylandı' AND s.day<? AND s.kind IN (?,?) ORDER BY s.day,s.kind''',(target,*kinds))]

def confirm_day(app,day):
    with app.conn() as c:
        c.execute('BEGIN IMMEDIATE')
        lists=[r for r in pending(c) if r['day']==day]
        if not lists:raise ValueError('Bu tarihte onaylanacak dolu taslak yok')
        for sheet in lists:
            nobet_manuel.require_complete(c,sheet)
            if sheet['undecided']:raise ValueError(f'{sheet["kind"]}: {sheet["undecided"]} asıl görev hâlâ Planlandı. Önce Tuttu/Tutmadı durumlarını işaretleyin.')
            app.nobet_planlayici.before_confirm(c,sheet)
            app.ceza_nobet.validate_sheet(app,c,sheet)
        for sheet in lists:c.execute("UPDATE sheets SET status='Onaylandı' WHERE id=?",(sheet['id'],))
        app.log(c,'Günün listeleri topluca onaylandı',day+'; '+', '.join(s['kind'] for s in lists))
    return [s['kind'] for s in lists]
