import calendar
from collections import defaultdict
from datetime import date, datetime, timedelta, time
from io import BytesIO
import re
import unicodedata

import streamlit as st
from openpyxl import load_workbook, Workbook
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
    start=datetime.combine(day,time(hour=start_hour%24))
    return start,start+timedelta(hours=shift_hours)

def build_rows(schedule, year, month, hours, times, employee_overrides):
    rows=[]
    for (y,m,section,employee),days in sorted(schedule.items(),key=lambda k:(k[0][2],k[0][3])):
        if (y,m)!=(year,month):continue
        start_hour=employee_overrides.get((section,employee),times.get(section,10))
        for day in range(1,calendar.monthrange(year,month)[1]+1):
            original=days.get(day,'')
            d=date(year,month,day)
            work=is_work(original)
            if work:
                start,end=sample_shift(d,section,hours,start_hour,22)
                duration=(end-start).total_seconds()/3600
            else:start=end=None;duration=0
            rows.append({'section':section,'employee':employee,'date':d,'marker':original,'is_work':work,'start':start,'end':end,'hours':duration,'status':'Simulare pe baza marcajului x' if work else ('Marcaj neinterpretat' if original and textnorm(original) not in ('off','liber','l') else 'Fara prezenta simulata')})
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
    sim.append(['SECTIE','ANGAJAT','DATA','ZI','MARCAJ SURSA','INTRARE SIMULATA','IESIRE SIMULATA','ORE SIMULATE','EXPLICATIE'])
    for r in rows:
        sim.append([r['section'],r['employee'],r['date'],r['date'].strftime('%a'),r['marker'],r['start'],r['end'],r['hours'],r['status']])
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
shift_hours=cols[0].number_input('Durata turei simulate (ore)',min_value=1,max_value=16,value=8,step=1)
min_rest=cols[1].number_input('Prag orientativ repaus intre ture (ore)',min_value=1,max_value=24,value=12)
weekly_limit=cols[2].number_input('Prag orientativ ore/saptamana',min_value=8,max_value=80,value=48)
sections=sorted({k[2] for k in schedule if k[:2]==(year,month)})
times={}
for section in sections:
    times[section]=st.slider(f'Ora inceperii turei simulate – {section}',0,23,10,key='start_'+section)
with st.expander('Ore de inceput personalizate pe angajat (optional)'):
    override={}
    persons=sorted({(k[2],k[3]) for k in schedule if k[:2]==(year,month)})
    for sec,name in persons:
        entry=st.number_input(f'{sec} / {name}',min_value=0,max_value=23,value=int(times[sec]),key='person_'+sec+'_'+name)
        override[(sec,name)]=entry
rows=build_rows(schedule,year,month,shift_hours,times,override)
issues=assess(rows,min_rest,weekly_limit)
c1,c2,c3=st.columns(3)
c1.metric('Angajati detectati',len(set((r['section'],r['employee']) for r in rows)))
c2.metric('Zile marcate ca lucrate',sum(r['is_work'] for r in rows))
c3.metric('Semnalari orientative',len(issues))
st.subheader('Previzualizare simulare')
st.dataframe([{'Sectie':r['section'],'Angajat':r['employee'],'Data':str(r['date']),'Original':r['marker'],'Intrare simulata':r['start'].strftime('%d.%m %H:%M') if r['start'] else '', 'Iesire simulata':r['end'].strftime('%d.%m %H:%M') if r['end'] else '', 'Ore':r['hours']} for r in rows if r['is_work']],use_container_width=True,hide_index=True)
if issues:
    with st.expander('Semnalari / abateri de verificat'):
        st.dataframe([{'Sectie':i[0],'Angajat':i[1],'Perioada':i[2],'Regula':i[3],'Detalii':i[4]} for i in issues],use_container_width=True,hide_index=True)
with st.expander('Diagnostic import'):
    for msg in diagnostics:st.text(msg)
assumptions={'Durata tura':f'{shift_hours} ore','Prag repaus intre ture':f'{min_rest} ore','Prag ore/saptamana':f'{weekly_limit} ore',**{f'Ora start {s}':f'{h:02d}:00' for s,h in times.items()}}
output=make_export(rows,issues,uploaded.name,year,month,assumptions,diagnostics)
st.download_button('⬇️ Descarca simularea Excel',data=output,file_name=f'SIMULARE_program_{year}_{month:02d}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
st.caption('Aplicatia nu pastreaza un istoric al fisierelor. Nu incarcatI date personale in medii care nu au controalele de acces si GDPR adecvate.')
