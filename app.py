import calendar
from collections import defaultdict
from datetime import date, datetime, timedelta, time
from io import BytesIO
import re
import unicodedata

import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title='Pontaj | Simulator retrospectiv', page_icon='📅', layout='wide')
st.title('Simulator retrospectiv de program')
st.caption('Import Excel • modelare ture • verificari orientative • export Excel')
st.warning('Rezultatele sunt SIMULARI, nu evidenta orelor efectiv lucrate si nu trebuie prezentate ca pontaj real. Fisierul sursa ramane neschimbat. Verificarea finala trebuie facuta pe baza contractelor si legislatiei aplicabile.')

MONTHS={'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'mai':5,'jun':6,'iul':7,'jul':7,'aug':8,'sep':9,'sept':9,'oct':10,'nov':11,'dec':12}

def textnorm(v):
    return ''.join(c for c in unicodedata.normalize('NFKD',str(v or '')) if not unicodedata.combining(c)).lower().strip()

def month_from_title(value):
    s=textnorm(value)
    if 'month' not in s: return None
    match=re.search(r'(20\d{2})',s)
    if not match: return None
    for key,num in MONTHS.items():
        if re.search(r'(?<![a-z])'+key+r'[a-z]*(?![a-z])',s) or key in s:
            return int(match.group(1)),num
    return None

def is_work(value):
    if value is None:return False
    s=textnorm(value)
    return s in {'x','xx','work','w','1','yes','da'}

def extract_schedule(data):
    wb=load_workbook(BytesIO(data),read_only=True,data_only=True)
    out=defaultdict(dict)
    diagnostics=[]
    for sheet in wb:
        matrix=list(sheet.iter_rows(values_only=True))
        for idx,row in enumerate(matrix):
            spec=month_from_title(row[0] if row else None)
            if not spec:continue
            year,month=spec
            if idx+1>=len(matrix):continue
            dates=matrix[idx+1]
            daycols={}
            for col,val in enumerate(dates[1:],start=1):
                try:
                    day=int(val)
                    if 1<=day<=calendar.monthrange(year,month)[1]:daycols[col]=day
                except (TypeError,ValueError):pass
            if not daycols:continue
            count=0
            for r in matrix[idx+2:]:
                a=str(r[0] or '').strip() if r else ''
                if month_from_title(a):break
                if not a or textnorm(a) in {'count','sala','bar','total','kitchen'}:continue
                # Include named staff even if they have zero marked work days in the month.
                # Each section title / summary is filtered above.
                key=(year,month,sheet.title,a)
                for col,day in daycols.items():
                    v=r[col] if col<len(r) else None
                    if v is not None and str(v).strip():
                        old=out[key].get(day)
                        if old is not None and old!=str(v).strip():diagnostics.append(f'Marcaj diferit la {a}, {day:02d}.{month:02d}.{year}; ultima valoare a fost folosita')
                        out[key][day]=str(v).strip()
                count+=1
            diagnostics.append(f'{sheet.title}, {month:02d}.{year}: {count} randuri cu marcaje de prezenta')
    wb.close()
    return out,diagnostics

def sample_shift(day, section, shift_hours, start_hour, night_start):
    # Non-factual scenario only; explicit deterministic planning assumptions.
    start=datetime.combine(day,time.min)+timedelta(hours=start_hour)
    return start,start+timedelta(hours=shift_hours)

def build_rows(schedule, year, month, hours, times, employee_overrides, meal_hours=1, leaves=None):
    leaves=leaves or {}
    rows=[]
    for (y,m,section,employee),days in sorted(schedule.items(),key=lambda k:(k[0][2],k[0][3])):
        if (y,m)!=(year,month):continue
        start_hour=employee_overrides.get((section,employee),times.get(section,10))
        for day in range(1,calendar.monthrange(year,month)[1]+1):
            original=days.get(day,'')
            d=date(year,month,day)
            leave_type=leaves.get((section,employee,d))
            work=is_work(original) and not leave_type
            if work:
                start,end=sample_shift(d,section,hours,start_hour,22)
                duration=max(0,(end-start).total_seconds()/3600-meal_hours)
            else:start=end=None;duration=0
            rows.append({'section':section,'employee':employee,'date':d,'marker':original,'is_work':work,'start':start,'end':end,'hours':duration,'meal':meal_hours if work else 0,'leave':leave_type or '', 'status':('Concediu: '+leave_type) if leave_type else ('Simulare pe baza marcajului x' if work else ('Marcaj neinterpretat' if original and textnorm(original) not in ('off','liber','l') else 'Fara prezenta simulata'))})
    return rows

def assess(rows, min_rest, weekly_limit):
    issues=[]
    grouped=defaultdict(list)
    for r in rows:grouped[(r['section'],r['employee'])].append(r)
    for (section,employee),items in grouped.items():
        worked=[r for r in sorted(items,key=lambda x:x['date']) if r['is_work']]
        for a,b in zip(worked,worked[1:]):
            gap=(b['start']-a['end']).total_seconds()/3600
            if gap<min_rest:issues.append([section,employee,str(b['date']),'Repaus intre ture',f'{gap:.1f}h (prag orientativ {min_rest}h)'])
        weeks=defaultdict(float)
        for r in worked:
            yr,wk,_=r['date'].isocalendar();weeks[(yr,wk)]+=r['hours']
        for (yr,wk),h in weeks.items():
            if h>weekly_limit:issues.append([section,employee,f'Saptamana ISO {wk}/{yr}','Ore saptamanale',f'{h:g}h (prag orientativ {weekly_limit}h)'])
        for i in range(len(items)-6):
            window=items[i:i+7]
            intervals=[]
            for r in window:
                if r['is_work']:intervals.append((r['start'],r['end']))
            anchor=datetime.combine(window[0]['date'],time.min)
            finish=datetime.combine(window[-1]['date']+timedelta(days=1),time.min)
            points=[(anchor,intervals[0][0] if intervals else finish)]
            points.extend((a[1],b[0]) for a,b in zip(intervals,intervals[1:]))
            if intervals:points.append((intervals[-1][1],finish))
            if not any((b-a).total_seconds()>=48*3600 for a,b in points):
                issues.append([section,employee,f'{window[0]["date"]} – {window[-1]["date"]}','Repaus 48h in fereastra de 7 zile','Nu exista interval continuu de 48h in aceasta fereastra; analiza legala necesita context'])
                break
    return issues

def make_export(rows,issues, source_name, year,month, assumptions,diagnostics):
    wb=Workbook(); sim=wb.active;sim.title='SIMULARE - NU PONTAJ'
    sim.append(['SECTIE','ANGAJAT','DATA','ZI','MARCAJ SURSA','INTRARE SIMULATA','IESIRE SIMULATA','ORE SIMULATE','PAUZA MASA (h)','CONCEDIU','EXPLICATIE'])
    for r in rows:
        sim.append([r['section'],r['employee'],r['date'],r['date'].strftime('%a'),r['marker'],r['start'],r['end'],r['hours'],r['meal'],r['leave'],r['status']])
    for row in sim.iter_rows(min_row=2):
        row[2].number_format='DD.MM.YYYY'; row[5].number_format='DD.MM.YYYY HH:MM'; row[6].number_format='DD.MM.YYYY HH:MM';row[7].number_format='0.00'
    a=wb.create_sheet('ABATERI SIMULARE');a.append(['SECTIE','ANGAJAT','PERIOADA','REGULA','OBSERVATIE'])
    for issue in issues:a.append(issue)
    o=wb.create_sheet('SURSA - MARCAJE');o.append(['SECTIE','ANGAJAT','DATA','MARCAJ ORIGINAL'])
    for r in rows:o.append([r['section'],r['employee'],r['date'],r['marker']]);o.cell(o.max_row,3).number_format='DD.MM.YYYY'
    meta=wb.create_sheet('IPOTEZE SI LIMITARI')
    for line in [
        ['TIP DOCUMENT','SIMULARE RETROSPECTIVA - NU ESTE EVIDENTA REALA A ORELOR LUCRATE'],
        ['Fisier sursa',source_name],['Luna',f'{month:02d}.{year}'],
        ['Metoda','Zilele cu x sunt interpretate drept prezenta; orele sunt generate prin ipoteze configurabile.'],
        ['Atentie','Fara inregistrari reale de intrare/iesire, orele efective nu pot fi deduse din marcajul x.'],
        ['Atentie','Nu se confirma conformitate legislativa; verificarea repausului, sporurilor si compensarilor presupune date contractuale si perioade adiacente.'],
        ['Atentie','Pastrati separat evidentelor reale; nu utilizati aceasta simulare ca inlocuitor al evidentei obligatorii.'],
        *[[k,str(v)] for k,v in assumptions.items()],
        ['DIAGNOSTIC IMPORT',''],*[[s,''] for s in diagnostics]
    ]:meta.append(line)
    for ws in wb:
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for cell in ws[1]:cell.fill=PatternFill('solid',fgColor='17324D');cell.font=Font(color='FFFFFF',bold=True)
        for col in ws.columns:
            letter=col[0].column_letter
            width=min(68,max(14,max(len(str(c.value or '')) for c in list(col)[:250])+2))
            ws.column_dimensions[letter].width=width
    stream=BytesIO();wb.save(stream);return stream.getvalue()

def make_cfp_export(rows, template_bytes, year, month):
    """Copy the actual CFP template unchanged in layout; fill only simulated roster and time cells.
    If roster exceeds available blocks, use additional copies of the same template as new sheets.
    """
    wb=load_workbook(BytesIO(template_bytes))
    def clear_editable(sheet, row, col):
        cell = sheet.cell(row, col)
        # MergedCell is a read-only placeholder for the non-anchor portion of a merge.
        # Keep formulas in the official template intact.
        if isinstance(cell, MergedCell) or cell.data_type == 'f':
            return
        cell.value = None

    def write_editable(sheet, row, col, value, required=False):
        cell = sheet.cell(row, col)
        if isinstance(cell, MergedCell):
            if required:
                raise ValueError(f'Campul obligatoriu {sheet.title}!{cell.coordinate} este unit cu alta celula in sablon.')
            return False
        cell.value = value
        return True
    ws=wb.worksheets[0]
    template_starts=[r for r in range(16,ws.max_row+1) if isinstance(ws.cell(r,1).value,int) and ws.cell(r,2).value]
    if not template_starts: raise ValueError('Sablonul CFP nu contine blocurile de 7 randuri asteptate.')
    names=sorted(set((r['section'],r['employee']) for r in rows))
    by_person=defaultdict(dict)
    for r in rows:by_person[(r['section'],r['employee'])][r['date'].day]=r
    sheets=[]
    for index in range(0,len(names),len(template_starts)):
        tab=ws if index==0 else wb.copy_worksheet(ws)
        tab.title=f'SIMULARE CFP {index//len(template_starts)+1}'
        tab['A6']=datetime(year,month,1)
        sheets.append(tab)
        for slot,base in enumerate(template_starts):
            pos=index+slot
            if pos>=len(names):
                clear_editable(tab,base,1)
                clear_editable(tab,base,2)
                clear_editable(tab,base,3)
                clear_editable(tab,base,4)
                for day in range(1,32):
                    c=6+2*(day-1)
                    for offset,cc in [(1,c),(1,c+1),(2,c),(2,c+1),(3,c),(3,c+1),(5,c),(5,c+1),(6,c),(6,c+1)]:
                        clear_editable(tab,base+offset,cc)
                continue
            sec,name=names[pos]
            write_editable(tab,base,1,pos+1,required=True)
            write_editable(tab,base,2,name,required=True)
            clear_editable(tab,base,3)  # do not inherit another employee's hire date
            write_editable(tab,base,4,sec)
            write_editable(tab,base,5,8)  # assumption only, not verified contracted norm
            for day in range(1,32):
                c=6+2*(day-1)
                r=by_person[(sec,name)].get(day)
                # Clear values from prior example; retain template's daily formulas and formatting.
                for offset,cc in [(1,c),(1,c+1),(2,c),(2,c+1),(3,c),(3,c+1),(5,c),(5,c+1),(6,c),(6,c+1)]:
                    clear_editable(tab,base+offset,cc)
                if r and r['is_work']:
                    start=r['start'];end=r['end']
                    write_editable(tab,base+2,c,timedelta(hours=start.hour,minutes=start.minute),required=True)
                    write_editable(tab,base+2,c+1,end-start.replace(hour=0,minute=0,second=0,microsecond=0),required=True)
                    tab.cell(base+2,c).number_format='[h]:mm'
                    tab.cell(base+2,c+1).number_format='[h]:mm'
                    write_editable(tab,base+5,c,timedelta(hours=r['meal']))
                    write_editable(tab,base+5,c+1,'masa')
                elif r and r['leave']:
                    write_editable(tab,base+1,c,r['leave'],required=True)
    # A visible disclosure on the existing sheet, outside the main time grid.
    for tab in sheets:
        tab['A136']='SIMULARE RETROSPECTIVA - NU ESTE EVIDENTA ORELOR EFECTIV LUCRATE'
        tab['A136'].font=Font(bold=True,color='C00000')
    audit=wb.create_sheet('SIMULARE - IPOTEZE')
    for line in [
        ['TIP','SIMULARE - NU PONTAJ REAL'],
        ['PERIOADA',f'{month:02d}/{year}'],
        ['PAUZA MASA','Scazuta din durata de prezenta; configurata in aplicatie'],
        ['DATE CONTRACTUALE','Norma din sablon este o ipoteza (8 ore); validati individual'],
        ['CONCEDII','Marcaje adaugate doar in scenariu, nu documente justificative'],
        ['ATENTIE','Datele din acest fisier nu pot inlocui evidenta orelor efectiv lucrate']
    ]:audit.append(line)
    stream=BytesIO();wb.save(stream);return stream.getvalue()

uploaded=st.file_uploader('1. Incarca fisierul de program (Kitchen Staff)',type=['xlsx'])
if uploaded is None:
    st.info('Incarca fisierul Kitchen Staff Sched New.xlsx. Aplicatia recunoaste sectiunile Kitchen, Sala si lunile din acelasi workbook.')
    st.stop()
try: schedule,diagnostics=extract_schedule(uploaded.getvalue())
except Exception as exc:
    st.error(f'Nu am putut interpreta fisierul: {exc}');st.stop()
periods=sorted(set((k[0],k[1]) for k in schedule),reverse=True)
if not periods:st.error('Nu au fost detectate luni cu marcaje de prezenta.');st.stop()
labels={f'{calendar.month_name[m]} {y}':(y,m) for y,m in periods}
selected=st.selectbox('2. Luna simulata',list(labels.keys()))
year,month=labels[selected]
st.subheader('3. Ipoteze pentru scenariul de ture')
st.caption('Acestea sunt ore presupuse, nu ore constatate. Poti modifica orele pe sectie si pe angajat.')
cols=st.columns(3)
shift_hours=cols[0].number_input('Durata prezentei simulate (ore, include pauza)',min_value=1.5,max_value=16.0,value=9.0,step=0.5)
min_rest=cols[1].number_input('Prag orientativ repaus intre ture (ore)',min_value=1,max_value=24,value=12)
weekly_limit=cols[2].number_input('Prag orientativ ore/saptamana',min_value=8,max_value=80,value=48)
meal_hours=st.selectbox('Pauza de masa scazuta din prezenta', [1.0,0.5,0.0], index=0, format_func=lambda v:f'{v:g} ore')
sections=sorted({k[2] for k in schedule if k[:2]==(year,month)})
times={}
for section in sections:
    times[section]=st.select_slider(f'Ora inceperii turei simulate – {section}',options=[x/2 for x in range(48)],value=10.0,format_func=lambda v:f'{int(v):02d}:{int((v%1)*60):02d}',key='start_'+section)
with st.expander('Ore de inceput personalizate pe angajat (optional)'):
    override={}
    persons=sorted({(k[2],k[3]) for k in schedule if k[:2]==(year,month)})
    for sec,name in persons:
        entry=st.select_slider(f'{sec} / {name}',options=[x/2 for x in range(48)],value=float(times[sec]),format_func=lambda v:f'{int(v):02d}:{int((v%1)*60):02d}',key='person_'+sec+'_'+name)
        override[(sec,name)]=entry
if 'leave_items' not in st.session_state:st.session_state.leave_items=[]
with st.expander('Concedii - adauga / sterge intervale',expanded=False):
    person_names=sorted({(k[2],k[3]) for k in schedule if k[:2]==(year,month)})
    person_choice=st.selectbox('Angajat',person_names,format_func=lambda x:f'{x[0]} / {x[1]}')
    leave_type=st.selectbox('Tip', ['CO','CM','CFS','Alte absente'])
    lcol1,lcol2=st.columns(2)
    first=lcol1.date_input('De la',value=date(year,month,1),min_value=date(year,month,1),max_value=date(year,month,calendar.monthrange(year,month)[1]))
    last=lcol2.date_input('Pana la',value=date(year,month,1),min_value=date(year,month,1),max_value=date(year,month,calendar.monthrange(year,month)[1]))
    if st.button('Adauga concediul'):
        if last<first:st.error('Data finala este inaintea datei initiale.')
        else:st.session_state.leave_items.append((year,month,person_choice,first,last,leave_type))
    for idx, item in enumerate(list(st.session_state.leave_items)):
        if item[:2]!=(year,month):continue
        st.write(f'{idx+1}. {item[2][1]} | {item[3]} – {item[4]} | {item[5]}')
    if st.button('Sterge toate concediile din sesiune'):st.session_state.leave_items=[];st.rerun()
leaves={}
for yy,mm,person,first,last,typ in st.session_state.leave_items:
    if (yy,mm)!=(year,month):continue
    current=first
    while current<=last:
        leaves[(*person,current)]=typ
        current+=timedelta(days=1)
rows=build_rows(schedule,year,month,shift_hours,times,override,meal_hours,leaves)
issues=assess(rows,min_rest,weekly_limit)
c1,c2,c3=st.columns(3)
c1.metric('Angajati detectati',len(set((r['section'],r['employee']) for r in rows)))
c2.metric('Zile marcate ca lucrate',sum(r['is_work'] for r in rows))
c3.metric('Semnalari orientative',len(issues))
st.subheader('Previzualizare simulare')
st.dataframe([{'Sectie':r['section'],'Angajat':r['employee'],'Data':str(r['date']),'Original':r['marker'],'Intrare simulata':r['start'].strftime('%d.%m %H:%M') if r['start'] else '', 'Iesire simulata':r['end'].strftime('%d.%m %H:%M') if r['end'] else '', 'Ore efective simulate':r['hours'],'Pauza masa':r['meal'],'Concediu':r['leave']} for r in rows if r['is_work']],use_container_width=True,hide_index=True)
if issues:
    with st.expander('Semnalari / abateri de verificat'):
        st.dataframe([{'Sectie':i[0],'Angajat':i[1],'Perioada':i[2],'Regula':i[3],'Detalii':i[4]} for i in issues],use_container_width=True,hide_index=True)
with st.expander('Diagnostic import'):
    for msg in diagnostics:st.text(msg)
assumptions={'Durata tura':f'{shift_hours} ore','Prag repaus intre ture':f'{min_rest} ore','Prag ore/saptamana':f'{weekly_limit} ore',**{f'Ora start {s}':f'{int(h):02d}:{round((h % 1) * 60):02d}' for s,h in times.items()}}
output=make_export(rows,issues,uploaded.name,year,month,assumptions,diagnostics)
st.download_button('⬇️ Descarca simularea Excel',data=output,file_name=f'SIMULARE_program_{year}_{month:02d}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
st.caption('Aplicatia nu pastreaza un istoric al fisierelor. Nu incarcatI date personale in medii care nu au controalele de acces si GDPR adecvate.')

st.subheader('Export in formatul exact al sablonului CFP')
template_upload=st.file_uploader('Incarca sablonul CFP original (.xlsx)',type=['xlsx'],key='cfp')
if template_upload is not None:
    try:
        cfp=make_cfp_export(rows,template_upload.getvalue(),year,month)
        st.download_button('Descarca simularea in format CFP',data=cfp,file_name=f'SIMULARE_CFP_{year}_{month:02d}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
        st.caption('Sablonul este pastrat, iar intervalele generate sunt marcate explicit ca simulare. Validati formulele si norma contractuala.')
    except Exception as exc:st.error(f'Exportul CFP nu a reusit: {exc}')
