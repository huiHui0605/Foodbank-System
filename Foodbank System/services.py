import json,os,re,uuid,hashlib,tempfile
from pathlib import Path
from pptx import Presentation
ROOT=Path(__file__).resolve().parent
OUT=ROOT/"certificates";OUT.mkdir(exist_ok=True)
HEADERS=["record_id","source","type","name","phone","email","date","external_id","certificate_pdf","certificate_signature","status"]
ALIASES={"name":["name","fullname","participantname","donorname","volunteername","姓名","名字","nama","nama penuh"],"phone":["phone","contact","contactnumber","mobile","phonenumber","电话"],"email":["email","emailaddress"],"date":["date","timestamp","submissiondate"]}
def norm(x):return re.sub(r"[^\w]+","",str(x or "").lower())
def val(x):return "" if x is None else str(x)
def replace_text(shape,values):
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                for k,v in values.items():run.text=run.text.replace('{{'+k+'}}',v)
            # Handle placeholders split across runs; keep first run formatting.
            combined=''.join(r.text for r in para.runs)
            if any('{{'+k+'}}' in combined for k in values):
                for k,v in values.items():combined=combined.replace('{{'+k+'}}',v)
                if para.runs:
                    para.runs[0].text=combined
                    for run in para.runs[1:]:run.text=''
    if shape.shape_type==6:
        for sub in shape.shapes:replace_text(sub,values)
def signature(d,template):
    stat=template.stat()
    return hashlib.sha256('|'.join([d.get(k,'') for k in ('type','name','date','external_id')]+[str(stat.st_mtime_ns),str(stat.st_size)]).encode()).hexdigest()
def generate(d):
    kind=d['type'].lower()
    if kind not in ('donor','volunteer'):return
    template=ROOT/'templates'/f'{kind}_template.pptx'
    if not template.exists():d['status']='Missing PPTX template';return
    if not d['name']:d['status']='Missing name';return
    sig=signature(d,template)
    if d.get('certificate_signature')==sig and Path(d.get('certificate_pdf') or '').is_file():return
    folder=OUT/kind;folder.mkdir(parents=True,exist_ok=True)
    filename=re.sub(r'[^\w\-]+','_',d['name'],flags=re.UNICODE)[:55] or 'participant'
    pdf=folder/f'{filename}_{d["record_id"][:8]}.pdf'
    with tempfile.TemporaryDirectory() as tmp:
        pptx=Path(tmp)/'certificate.pptx'
        prs=Presentation(str(template))
        values={'NAME':d['name'],'DATE':d['date'],'ID':d['external_id'],'TYPE':kind.title()}
        for slide in prs.slides:
            for shape in slide.shapes:replace_text(shape,values)
        prs.save(str(pptx))
        import pythoncom,win32com.client
        pythoncom.CoInitialize();app=None;presentation=None
        try:
            app=win32com.client.DispatchEx('PowerPoint.Application')
            presentation=app.Presentations.Open(str(pptx.resolve()),WithWindow=False)
            presentation.SaveAs(str(pdf.resolve()),32)
        finally:
            if presentation is not None:presentation.Close()
            if app is not None:app.Quit()
            pythoncom.CoUninitialize()
    d['certificate_pdf']=str(pdf);d['certificate_signature']=sig;d['status']='PDF ready'
def sync_google(rows):
    cfgpath=ROOT/'config.json'
    if not cfgpath.exists():raise RuntimeError('Missing config.json; copy config.example.json')
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    cfg=json.loads(cfgpath.read_text(encoding='utf-8'))
    cred=service_account.Credentials.from_service_account_file(str(ROOT/cfg['google_credentials_file']),scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    api=build('sheets','v4',credentials=cred,cache_discovery=False)
    lookup={(r['source'],r['external_id']):r for r in rows if r['source'].startswith('google:')}
    added=0;updated=0
    for item in cfg.get('sheets',[]):
        sid=item.get('spreadsheet_id','').strip();kind=item.get('type','').lower()
        if not sid:continue
        if kind not in ('donor','volunteer','beneficiary'):continue
        grid=api.spreadsheets().values().get(spreadsheetId=sid,range=item.get('range','Form Responses 1!A:Z')).execute().get('values',[])
        if len(grid)<2:continue
        headers=[norm(x) for x in grid[0]]
        def field(row,key):
            for alias in ALIASES[key]:
                a=norm(alias)
                if a in headers:
                    i=headers.index(a)
                    return val(row[i]) if i<len(row) else ''
            return ''
        for i,line in enumerate(grid[1:],start=2):
            if not any(val(x) for x in line):continue
            source='google:'+sid+':'+kind
            # row number is stable for append-only Google Form response sheets
            ext=f'row:{i}'
            key=(source,ext)
            fields={'name':field(line,'name'),'phone':field(line,'phone'),'email':field(line,'email'),'date':field(line,'date')}
            if key in lookup:
                rec=lookup[key]
                if any(rec.get(k,'')!=v for k,v in fields.items()):rec.update(fields);updated+=1
            else:
                rec={k:'' for k in HEADERS};rec.update(fields);rec.update(record_id=str(uuid.uuid4()),source=source,type=kind,external_id=ext)
                rows.append(rec);lookup[key]=rec;added+=1
    return added,updated
