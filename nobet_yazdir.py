"""Print layouts compiled from the user's three Excel reference sheets."""
import datetime as dt
import html
import json
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED
from excel_devriye import PLACES, RESERVES

NIGHT=['23:00-01:00','01:00-03:00','03:00-05:00','05:00-07:00']
GATE=['08:00-10:00','10:00-12:00','12:00-14:00','14:00-16:00','16:00-17:00']

def legacy_patrol(rows):
    return (len(rows)==42 and sum(r['role']=='Yedek' for r in rows)==RESERVES
            and not any(re.fullmatch(r'5\s*D',str(r['location']).strip(),re.I) for r in rows))

def slot(row):
    value=str(row['slot']).replace('–','-').replace('—','-').replace(' ','')
    if value in NIGHT+GATE:return value
    if row.get('start_at') and row.get('end_at'):return row['start_at'][11:16]+'-'+row['end_at'][11:16]
    return value

def fields(patches,row,columns,index):
    for column,key in columns:patches[f'{column}{index}']=str(row.get(key,'') or '') if row else ''

def assign(groups,key,row,limit):
    group=groups.setdefault(key,[]);group.append(row)
    if len(group)>limit:raise ValueError(f'{key}: aynı görev yerinde şablon kapasitesinden fazla öğrenci var. Nöbet listesini kontrol edin.')

def placed_row(rows, positions, index):
    fixed = {row['manual_position']: row for row in rows if row.get('manual_position') in positions}
    remaining = iter(row for row in rows if row.get('manual_position') not in fixed)
    placed = [fixed[position] if position in fixed else next(remaining, None) for position in positions]
    return placed[index]

def make_patches(kind,day,by_day):
    base=dt.date.fromisoformat(day);tomorrow=base+dt.timedelta(days=1);patches={};used=0
    rows=[dict(row) for row in by_day.get(day,[])]
    if kind=='Yatakhane Kat':
        groups={}
        for row in rows:
            location=str(row['location']).upper().strip()
            match=re.fullmatch(r'([234])\s*([ABC])',location) or re.search(r'([234]).*?\b([ABC])\b',location)
            if not match or (match[1]=='4' and match[2]=='C'):raise ValueError(f'{row["name"]} {row["surname"]}: kat/blok bilgisi anlaşılamadı. Excel’den saat ve yerleri kontrol ederek aktarın.')
            key=(int(match[1]),match[2],'Yedek' if row['role']=='Yedek' else slot(row))
            if key[2] not in NIGHT+['Yedek']:raise ValueError('Yatakhane saat aralığı eksik; Excel’den doğru saatlerle aktarın.')
            assign(groups,key,row,2 if row['role']=='Yedek' else 1)
        for floor,top in [(2,7),(3,17),(4,27)]:
            patches[f'A{top}']=base.strftime('%d.%m.%Y');patches[f'A{top+1}']=tomorrow.strftime('%d.%m.%Y')
            for block,column in [('A',3),('B',8),('C',13)]:
                if floor==4 and block=='C':continue
                for offset in range(6):
                    key=(floor,block,NIGHT[offset] if offset<4 else 'Yedek')
                    available=groups.get(key,[]);index=0 if offset<4 else offset-4
                    place_index=(floor-2)*3+column//5
                    person=(placed_row(available,(32+place_index,40+place_index),index) if offset>=4
                            else next(iter(available),None))
                    columns=[(chr(64+column+i),field) for i,field in enumerate(['name','surname','class','room'])]
                    fields(patches,person,columns,top+offset);patches[f'{chr(64+column+4)}{top+offset}']=''
                    used+=bool(person)
    elif kind=='Devriye':
        places=PLACES[:-2] if legacy_patrol(rows) else PLACES
        reserve_top=3+len(NIGHT)*len(places)
        groups={};reserves=[]
        for row in rows:
            if row['role']=='Yedek':reserves.append(row);continue
            location=str(row['location']).upper().strip()
            if 'HABER' in location:location='HABER M.'
            else:
                match=re.fullmatch(r'([1-5])\s*D',location)
                if not match:raise ValueError(f'{row["name"]} {row["surname"]}: devriye görev yeri eksik. Aynı Excel’i yeniden aktararak saat ve yerleri düzeltin.')
                location=match[1]+' D'
            if slot(row) not in NIGHT:raise ValueError('Devriye saatleri eksik. Aynı Excel’i yeniden aktararak saat ve yerleri düzeltin.')
            assign(groups,(slot(row),location),row,1 if location=='HABER M.' else 2)
        if len(reserves)>RESERVES:raise ValueError('Devriye şablonunda en fazla 6 yedek olabilir')
        patches['A3']=base.strftime('%d.%m.%Y')
        patches[f'A{3+len(places)}']=tomorrow.strftime('%d.%m.%Y')
        for group,time in enumerate(NIGHT):
            for i,place in enumerate(places):
                candidates=groups.get((time,place),[]);index=0 if i==0 else (i-1)%2
                first=i-index
                positions=tuple(group*len(places)+j for j in range(first,first+(1 if i==0 else 2)))
                person=placed_row(candidates,positions,index)
                index=3+group*len(places)+i
                fields(patches,person,list(zip('DEFGH',['name','surname','class','school_no','room'])),index);patches[f'I{index}']='';used+=bool(person)
        for i in range(RESERVES):
            person=placed_row(reserves,tuple(range(reserve_top-3,reserve_top-3+RESERVES)),i)
            index=reserve_top+i
            fields(patches,person,list(zip('DEFGH',['name','surname','class','school_no','room'])),index);patches[f'I{index}']='';used+=bool(person)
    elif kind=='Yatakhane Giriş':
        rows=[]
        for date,top in [(base,5),(tomorrow,15)]:
            date_text=date.isoformat();daily=[dict(row) for row in by_day.get(date_text,[])];rows+=daily;groups={}
            for row in daily:
                match=re.fullmatch(r'([AB])(?:\s*BLOK)?',str(row['location']).upper().strip())
                if not match:raise ValueError(f'{row["name"]} {row["surname"]}: giriş nöbetinde A/B blok bilgisi eksik')
                time='Yedek' if row['role']=='Yedek' else slot(row)
                if time not in GATE+['Yedek']:raise ValueError('Giriş nöbeti saat aralığı eksik')
                assign(groups,(match[1],time),row,2 if time=='Yedek' else 1)
            patches[f'A{top-2}']=date.strftime('%d.%m.%Y')
            for offset in range(7):
                patches[f'A{top+offset}']=date.strftime('%d.%m.%Y') if offset==0 else ''
                for block,columns,signature in [('A',list(zip('CDE',['name','surname','class'])),'F'),('B',list(zip('HIJ',['name','surname','class'])),'K')]:
                    candidates=groups.get((block,GATE[offset] if offset<5 else 'Yedek'),[]);index=0 if offset<5 else offset-5
                    block_index=0 if block=='A' else 1
                    person=(placed_row(candidates,(10+block_index,12+block_index),index) if offset>=5
                            else next(iter(candidates),None))
                    fields(patches,person,columns,top+offset);patches[f'{signature}{top+offset}']='';used+=bool(person)
    else:raise ValueError('Bilinmeyen nöbet türü')
    if used!=len(rows):raise ValueError('Bazı görevler çıktı şablonuna yerleştirilemedi. Listeyi kontrol edin.')
    return patches

def print_layout(kind,day,by_day):
    layouts=json.loads(Path(__file__).with_name('yazdirma_sablonlari.json').read_text(encoding='utf-8'))
    layout_key='Devriye 5 Ekip' if kind=='Devriye' and not legacy_patrol(by_day.get(day,[])) else kind
    return layouts[layout_key],make_patches(kind,day,by_day)


def roster_sheet(ns,kind,by_day,header_style):
    headers=('ADI','SOYADI','SINIF','KOĞUŞ NO','SİCİL / OKUL NO','NÖBET DURUMU',
             'SAAT / GÖREV','TARİH','GÖREV YERİ','ROL')
    sheet=ET.Element('worksheet',xmlns=ns)
    view=ET.SubElement(ET.SubElement(sheet,'sheetViews'),'sheetView',workbookViewId='0')
    ET.SubElement(view,'pane',ySplit='1',topLeftCell='A2',activePane='bottomLeft',state='frozen')
    cols=ET.SubElement(sheet,'cols')
    for i,width in enumerate((25,22,10,14,20,20,20,15,17,12),1):
        ET.SubElement(cols,'col',min=str(i),max=str(i),width=str(width),customWidth='1')
    data=ET.SubElement(sheet,'sheetData')
    hours=GATE if kind=='Yatakhane Giriş' else NIGHT
    order={time:i for i,time in enumerate(hours)}
    values=[headers]
    for date,duties in sorted(by_day.items()):
        reserve_numbers=defaultdict(int)
        for duty in sorted(duties,key=lambda row:(row['role']=='Yedek',order.get(slot(row),len(hours)))):
            place=duty['location']
            time=slot(duty)
            if duty['role']=='Yedek':
                reserve_numbers[place]+=1
                time=place if re.fullmatch(r'\d+\. YEDEK',place,re.I) else f'YEDEK-{reserve_numbers[place]}'
            actual_date=(duty.get('start_at') or '')[:10]
            if not actual_date:
                actual_date=date
                if kind!='Yatakhane Giriş' and time in NIGHT[1:]:
                    actual_date=(dt.date.fromisoformat(date)+dt.timedelta(days=1)).isoformat()
            values.append((duty['name'],duty['surname'],duty['class'],duty['room'],duty['school_no'],
                           duty.get('result',''),time,dt.date.fromisoformat(actual_date).strftime('%d.%m.%Y'),
                           place,duty['role']))
    for i,values_row in enumerate(values,1):
        row=ET.SubElement(data,'row',r=str(i))
        for j,value in enumerate(values_row):
            excel_text_cell(row,f'{chr(65+j)}{i}',value,header_style if i==1 else '0')
    ET.SubElement(sheet,'autoFilter',ref=f'A1:J{len(values)}')
    return sheet


def excel_text_cell(row,address,value,style):
    node=ET.SubElement(row,'c',r=address,s=style,t='inlineStr')
    string=ET.SubElement(ET.SubElement(node,'is'),'t',{'xml:space':'preserve'})
    string.text='' if value is None else str(value)


def save_excel(path,kind,day,by_day):
    layout,patches=print_layout(kind,day,by_day)
    widths,heights=layout['widths'],layout['heights']
    ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg='http://schemas.openxmlformats.org/package/2006/relationships'
    sheet=ET.Element('worksheet',xmlns=ns)
    properties=ET.SubElement(sheet,'sheetPr')
    ET.SubElement(properties,'pageSetUpPr',fitToPage='1')
    views=ET.SubElement(sheet,'sheetViews')
    ET.SubElement(views,'sheetView',showGridLines='0',workbookViewId='0')
    cols=ET.SubElement(sheet,'cols')
    for i,width in enumerate(widths,1):
        ET.SubElement(cols,'col',min=str(i),max=str(i),width=f'{max(1,(width*4/3-5)/7):.3f}',customWidth='1')
    data=ET.SubElement(sheet,'sheetData')
    by_row=defaultdict(list)
    for cell in layout['cells']:
        by_row[cell['row']].append(cell)
    styles=ET.Element('styleSheet',xmlns=ns)
    fonts,fills,borders=(ET.SubElement(styles,key) for key in ('fonts','fills','borders'))
    default_font=ET.SubElement(fonts,'font')
    ET.SubElement(default_font,'sz',val='11')
    ET.SubElement(default_font,'name',val='Calibri')
    ET.SubElement(borders,'border')
    for pattern in ('none','gray125'):
        ET.SubElement(ET.SubElement(fills,'fill'),'patternFill',patternType=pattern)
    ET.SubElement(ET.SubElement(styles,'cellStyleXfs',count='1'),'xf',numFmtId='0',fontId='0',fillId='0',borderId='0')
    xfs=ET.SubElement(styles,'cellXfs')
    ET.SubElement(xfs,'xf',numFmtId='0',fontId='0',fillId='0',borderId='0',xfId='0')
    ET.SubElement(ET.SubElement(styles,'cellStyles',count='1'),'cellStyle',name='Normal',xfId='0',builtinId='0')
    cache={}

    def style_for(cell,*,roster=False):
        sides=tuple(cell['border'].get(side) or '' for side in ('left','right','top','bottom'))
        if not roster and kind=='Yatakhane Giriş' and (4<=cell['row']<=11 or 14<=cell['row']<=21) and cell['col']!=7:
            sides=tuple(side or 'hair' for side in sides)
        key=tuple(cell[k] for k in ('font','size','bold','italic','color','fill','align','vertical','rotation'))+sides
        if key in cache:
            return cache[key]
        font_id,fill_id,border_id=str(len(fonts)),str(len(fills)),str(len(borders))
        font=ET.SubElement(fonts,'font')
        for flag,tag in (('bold','b'),('italic','i')):
            if cell[flag]:
                ET.SubElement(font,tag)
        ET.SubElement(font,'sz',val=str(cell['size']))
        ET.SubElement(font,'color',rgb='FF'+cell['color'].lstrip('#').upper())
        ET.SubElement(font,'name',val=cell['font'])
        if cell['fill'].startswith('#'):
            fill=ET.SubElement(ET.SubElement(fills,'fill'),'patternFill',patternType='solid')
            ET.SubElement(fill,'fgColor',rgb='FF'+cell['fill'][1:].upper())
        else:
            fill_id='0'
        border=ET.SubElement(borders,'border')
        for name,line in zip(('left','right','top','bottom'),sides):
            edge=ET.SubElement(border,name)
            if line:
                edge.set('style',line)
                ET.SubElement(edge,'color',rgb='FFBDBDBD' if line=='hair' else 'FF000000')
        ET.SubElement(border,'diagonal')
        index=str(len(xfs))
        xf=ET.SubElement(xfs,'xf',numFmtId='0',fontId=font_id,fillId=fill_id,borderId=border_id,
                         xfId='0',applyFont='1',applyFill='1',applyBorder='1',applyAlignment='1')
        horizontal=cell['align'] if cell['align'] in ('left','right') else 'center' if cell['align'] in ('center','centerContinuous') else 'left'
        vertical=cell['vertical'] if cell['vertical'] in ('top','bottom') else 'center'
        ET.SubElement(xf,'alignment',horizontal=horizontal,vertical=vertical,wrapText='1',
                      textRotation=str(cell['rotation'] if 0<=cell['rotation']<=180 else 0))
        cache[key]=index
        return index

    merges=[]
    for i,height in enumerate(heights,1):
        row=ET.SubElement(data,'row',r=str(i),ht=str(height),customHeight='1')
        for cell in sorted(by_row[i],key=lambda value:value['col']):
            address=cell['address']
            excel_text_cell(row,address,patches.get(address,cell['value']),style_for(cell))
            if (cell['endrow'],cell['endcol'])!=(i,cell['col']):
                merges.append(f"{address}:{chr(64+cell['endcol'])}{cell['endrow']}")
    header=style_for(dict(font='Calibri',size=11,bold=True,italic=False,color='#FFFFFF',
                          fill='#1F2937',align='left',vertical='center',rotation=0,border={}),roster=True)
    roster=roster_sheet(ns,kind,by_day,header)
    for section in (fonts,fills,borders,xfs):
        section.set('count',str(len(section)))
    if merges:
        merged=ET.SubElement(sheet,'mergeCells',count=str(len(merges)))
        for area in merges:
            ET.SubElement(merged,'mergeCell',ref=area)
    ET.SubElement(sheet,'pageMargins',left='0.315',right='0.315',top='0.315',bottom='0.315',header='0',footer='0')
    ET.SubElement(sheet,'pageSetup',paperSize='9',orientation=layout['orientation'],fitToWidth='1',fitToHeight='1')
    workbook=ET.Element('workbook',{'xmlns':ns,'xmlns:r':rel})
    ET.SubElement(ET.SubElement(workbook,'bookViews'),'workbookView')
    sheets=ET.SubElement(workbook,'sheets')
    ET.SubElement(sheets,'sheet',{'name':kind,'sheetId':'1','r:id':'rId1'})
    ET.SubElement(sheets,'sheet',{'name':'Kişi Listesi','sheetId':'2','r:id':'rId3'})
    names=ET.SubElement(workbook,'definedNames')
    ET.SubElement(names,'definedName',name='_xlnm.Print_Area',localSheetId='0').text=f"'{kind}'!$A$1:${chr(64+len(widths))}${len(heights)}"
    parts={'xl/workbook.xml':(workbook,'sheet.main'),
           'xl/worksheets/sheet1.xml':(sheet,'worksheet'),
           'xl/worksheets/sheet2.xml':(roster,'worksheet'),
           'xl/styles.xml':(styles,'styles')}
    types=ET.Element('Types',xmlns='http://schemas.openxmlformats.org/package/2006/content-types')
    ET.SubElement(types,'Default',Extension='rels',ContentType='application/vnd.openxmlformats-package.relationships+xml')
    for name,(_,content) in parts.items():
        ET.SubElement(types,'Override',PartName='/'+name,
                      ContentType=f'application/vnd.openxmlformats-officedocument.spreadsheetml.{content}+xml')
    root_rels=ET.Element('Relationships',xmlns=pkg)
    ET.SubElement(root_rels,'Relationship',Id='rId1',Type=rel+'/officeDocument',Target='xl/workbook.xml')
    sheet_rels=ET.Element('Relationships',xmlns=pkg)
    ET.SubElement(sheet_rels,'Relationship',Id='rId1',Type=rel+'/worksheet',Target='worksheets/sheet1.xml')
    ET.SubElement(sheet_rels,'Relationship',Id='rId2',Type=rel+'/styles',Target='styles.xml')
    ET.SubElement(sheet_rels,'Relationship',Id='rId3',Type=rel+'/worksheet',Target='worksheets/sheet2.xml')
    files={name:node for name,(node,_) in parts.items()}
    files.update({'[Content_Types].xml':types,'_rels/.rels':root_rels,'xl/_rels/workbook.xml.rels':sheet_rels})
    path=Path(path)
    with NamedTemporaryFile(dir=path.parent,suffix='.xlsx',delete=False) as stream:
        temp=Path(stream.name)
    try:
        with ZipFile(temp,'w',ZIP_DEFLATED) as archive:
            for name,node in files.items():
                archive.writestr(name,ET.tostring(node,encoding='utf-8',xml_declaration=True))
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def render(kind,day,by_day):
    layout,patches=print_layout(kind,day,by_day)
    width,height=(297,210) if layout['orientation']=='landscape' else (210,297)
    available_w,available_h=(width-16)*72/25.4,(height-16)*72/25.4
    widths,heights=layout['widths'],layout['heights'];scale=min(available_w/sum(widths),available_h/sum(heights))
    xs=[sum(widths[:i])*scale for i in range(len(widths)+1)];ys=[sum(heights[:i])*scale for i in range(len(heights)+1)]
    cells=[]
    for cell in layout['cells']:
        value=patches.get(cell['address'],cell['value']);r,c=cell['row']-1,cell['col']-1
        left,top=xs[c],ys[r];w,h=xs[cell['endcol']]-left,ys[cell['endrow']]-top
        css=f'left:{left:.3f}pt;top:{top:.3f}pt;width:{w:.3f}pt;height:{h:.3f}pt;font-size:{cell["size"]*scale:.3f}pt;font-weight:{700 if cell["bold"] else 400};color:{cell["color"]};background:{cell["fill"]};'
        if kind=='Yatakhane Giriş' and (4<=cell['row']<=11 or 14<=cell['row']<=21) and cell['col']!=7:css+='border:.35pt solid #bdbdbd;'
        for side,border in cell['border'].items():
            if border:css+=f'border-{side}:{max(.35,scale*(1.1 if border in ("medium","thick") else .6)):.3f}pt solid #000;'
        align='center' if cell['align'] in ('center','centerContinuous') else ('right' if cell['align']=='right' else 'left')
        vertical='flex-start' if cell['vertical']=='top' else ('flex-end' if cell['vertical']=='bottom' else 'center')
        inner=f'text-align:{align};'
        if re.fullmatch(r'\d{2}\.\d{2}\.\d{4}',str(value)) or (kind=='Devriye' and cell['row']==2):inner+='white-space:nowrap;'
        if cell['rotation'] in (90,180):inner+='writing-mode:vertical-rl;transform:rotate(180deg);'
        cells.append(f'<div class="cell" data-cell="{cell["address"]}" style="{css}justify-content:{vertical}"><span style="{inner}">{html.escape(str(value)).replace(chr(10),"<br>")}</span></div>')
    note='Giriş çıktısı seçili tarih ve ertesi günü birlikte gösterir.' if kind=='Yatakhane Giriş' else 'Çıktı kayıtlı nöbet listesinden hazırlanmıştır.'
    return f'''<!doctype html><html lang="tr"><head><meta charset="utf-8"><title>{html.escape(kind)} — {day}</title><style>
@page{{size:A4 {layout['orientation']};margin:8mm}}*{{box-sizing:border-box}}body{{margin:0;background:#e8e8e8;font-family:Arial,sans-serif}}.toolbar{{padding:12px;text-align:center;background:white;margin-bottom:12px}}button{{padding:8px 18px;cursor:pointer}}.sheet{{position:relative;width:{xs[-1]:.3f}pt;height:{ys[-1]:.3f}pt;background:white;margin:0 auto}}.cell{{position:absolute;display:flex;flex-direction:column;padding:1pt 1.5pt;overflow:hidden;line-height:1.05}}.cell span{{width:100%;white-space:normal;overflow-wrap:anywhere}}@media print{{body{{background:white}}.toolbar{{display:none}}.sheet{{margin:0 auto;break-inside:avoid}}*{{print-color-adjust:exact;-webkit-print-color-adjust:exact}}}}</style></head><body>
<div class="toolbar"><button onclick="window.print()">Yazdır / PDF kaydet</button> <span>{html.escape(note)} Yazdırma ayarlarında üstbilgi/altbilgiyi kapatın.</span></div>
<main class="sheet">{''.join(cells)}</main><script>function fit(){{for(const cell of document.querySelectorAll('.cell')){{const span=cell.firstElementChild;let size=parseFloat(getComputedStyle(cell).fontSize);while((span.scrollHeight>cell.clientHeight-2||span.scrollWidth>cell.clientWidth-2)&&size>6){{size-=0.25;cell.style.fontSize=size+'px';}}}}}}document.fonts.ready.then(fit);window.addEventListener('beforeprint',fit);</script></body></html>'''
