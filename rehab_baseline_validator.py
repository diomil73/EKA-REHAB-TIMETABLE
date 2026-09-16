from __future__ import annotations
import re, sys, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
X14='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
XM='http://schemas.microsoft.com/office/excel/2006/main'
NS={'m':MAIN,'r':REL,'p':PKG,'x14':X14,'xm':XM}
REQUIRED={'PATIENT_PLANNER','MASTER_SCHEDULE','THERAPIST_ATTENDANCE','REPLACEMENTS','REPLACEMENT_LOG','DAILY_LOG','STATISTICS','PATIENTS','NEW_PATIENT','SETTINGS','SESSIONS','CONFLICT_LOG','THERAPIST_DAILY'}

def colnum(ref):
    letters=re.match(r'[A-Z]+',ref).group(0)
    n=0
    for ch in letters: n=n*26+ord(ch)-64
    return n

def load(path):
    z=zipfile.ZipFile(path)
    sst=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        root=ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall('m:si',NS):
            sst.append(''.join(t.text or '' for t in si.iter(f'{{{MAIN}}}t')))
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rmap={x.attrib['Id']:x.attrib['Target'] for x in rels}
    smap={}
    for s in wb.find('m:sheets',NS):
        rid=s.attrib[f'{{{REL}}}id']
        smap[s.attrib['name']]='xl/'+rmap[rid].lstrip('/')
    return z,sst,smap

def read_sheet(z,sst,path):
    root=ET.fromstring(z.read(path))
    rows=[]
    for row in root.findall('.//m:sheetData/m:row',NS):
        out={}
        for c in row.findall('m:c',NS):
            ref=c.attrib['r']; t=c.attrib.get('t'); v=c.find('m:v',NS)
            val=None if v is None else v.text
            if t=='s' and val is not None: val=sst[int(val)]
            elif t=='inlineStr': val=''.join(x.text or '' for x in c.iter(f'{{{MAIN}}}t'))
            out[colnum(ref)]=val
        rows.append((int(row.attrib['r']),out))
    return root,rows

def main(path):
    z,sst,smap=load(path)
    issues=[]; notes=[]
    missing=sorted(REQUIRED-set(smap))
    if missing: issues.append('Missing required sheets: '+', '.join(missing))
    else: notes.append('All 13 required sheets are present.')
    if 'xl/vbaProject.bin' not in z.namelist(): issues.append('Embedded VBA project is missing.')
    else: notes.append('Embedded VBA project is present.')

    proot,prows=read_sheet(z,sst,smap['PATIENTS'])
    sroot,srows=read_sheet(z,sst,smap['SETTINGS'])
    p={r:d for r,d in prows}; s={r:d for r,d in srows}
    expected=['PatientID','Θάλαμος','Ασθενής','Μολυσματικός','Κατάσταση']
    actual=[p.get(1,{}).get(i) for i in range(1,6)]
    if actual!=expected: issues.append(f'PATIENTS headers differ: {actual}')
    else: notes.append('PATIENTS headers match the agreed compact model.')

    statuses={s[r].get(6) for r in s if r>=2 and s[r].get(6)}
    yesno={s[r].get(7) for r in s if r>=2 and s[r].get(7)}
    rooms={s[r].get(8) for r in s if r>=2 and s[r].get(8)}
    ids=[]; names=[]
    active_rows=0
    for r,d in prows:
        if r<2: continue
        pid=d.get(1); room=d.get(2); name=d.get(3); inf=d.get(4); status=d.get(5)
        if not any([pid,room,name,inf,status]): continue
        # Ignore preseeded blank-ID formula rows when no business data exists
        if name is None and room is None and inf is None and status is None: continue
        active_rows+=1
        if pid is not None: ids.append(str(pid))
        if name: names.append((r,name))
        if room and room not in rooms: issues.append(f'PATIENTS row {r}: room {room!r} not in SETTINGS ROOMS.')
        if inf and inf not in yesno: issues.append(f'PATIENTS row {r}: infectious value {inf!r} not in YES_NO.')
        if status and status not in statuses: issues.append(f'PATIENTS row {r}: status {status!r} not in PATIENT_STATUS.')
    dupids=sorted({x for x in ids if ids.count(x)>1})
    if dupids: issues.append('Duplicate PatientID values: '+', '.join(dupids))
    else: notes.append(f'{active_rows} patient records checked; no duplicate PatientID found.')

    normalized={}
    for r,n in names:
        k=' '.join(n.strip().upper().split())
        normalized.setdefault(k,[]).append(r)
    dupnames={k:v for k,v in normalized.items() if len(v)>1}
    if dupnames: notes.append('Potential duplicate display names: '+str(dupnames))
    odd=[]
    for r,n in names:
        if n!=n.strip() or n.endswith('.'):
            odd.append(f'row {r}: {n!r}')
    if odd: notes.append('Name-format items worth reviewing: '+ '; '.join(odd))

    # Check CF for infectious highlighting
    cf_ok=False
    for cf in proot.findall('m:conditionalFormatting',NS):
        for rule in cf.findall('m:cfRule',NS):
            formula=rule.find('m:formula',NS)
            if formula is not None and formula.text and '$D2="Ν"' in formula.text and 'A2:E' in cf.attrib.get('sqref',''):
                cf_ok=True
    if cf_ok: notes.append('Infectious highlighting is already data-driven by column D = "Ν".')
    else: issues.append('Could not confirm infectious conditional-format rule from column D.')

    # Extended data validations
    dv_text=ET.tostring(proot,encoding='unicode')
    expected_dv=[('B','SETTINGS!$H$2:$H$200'),('D','SETTINGS!$G$2:$G$10'),('E','SETTINGS!$F$2:$F$50')]
    for col,formula in expected_dv:
        if formula not in dv_text: issues.append(f'Could not confirm validation for column {col}: {formula}')
        else: notes.append(f'Column {col} validation points to {formula}.')

    therapists=[s[r].get(1) for r in s if r>=2 and s[r].get(1)]
    students=[x for x in therapists if 'φοιτ' in x.lower() or 'student' in x.lower()]
    if students: notes.append('Existing student placeholder(s) found in therapist settings: '+', '.join(students))

    print(f'BASELINE AUDIT: {Path(path).name}')
    print('RESULT:', 'PASS WITH NOTES' if not issues else 'CHECK REQUIRED')
    print('\nCHECKS')
    for n in notes: print('  OK  -',n)
    if issues:
        print('\nISSUES')
        for i in issues: print('  !!  -',i)
    return 1 if issues else 0

if __name__=='__main__':
    sys.exit(main(sys.argv[1] if len(sys.argv)>1 else 'Rehab_Center_System_v27_1.xlsm'))
