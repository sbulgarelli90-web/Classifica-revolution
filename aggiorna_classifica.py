import json,re,os,sys,time,webbrowser,threading
from pathlib import Path
from openpyxl import load_workbook

BASE=Path(__file__).resolve().parent
XLSX=BASE/'Classifica Revolution - corretto.xlsx'
HTML=BASE/'classifica_revolution.html'

def num(v):
    try:return float(v)
    except:return 0

def clean(v):
    return str(v).strip() if v is not None else ''

def rank_rows(rows,key='total'):
    rows=sorted(rows,key=lambda x:(-num(x.get(key)), x.get('name','')))
    last=None; rank=0
    for i,x in enumerate(rows,1):
        if last is None or num(x.get(key))!=last: rank=i; last=num(x.get(key))
        x['rank']=rank
    return rows

def active_font(cell):
    c=cell.font.color
    if not c:return False
    if c.type=='rgb': return str(c.rgb or '').upper().endswith('FF0000') or str(c.rgb or '').upper() in ('FFFF0000','00FF0000')
    return False

def header_name(v):
    s=clean(v)
    return re.sub(r'\s*\([^)]*\)\s*$','',s).strip()

def extract_const(text,name,next_name):
    pat=r'const '+re.escape(name)+r'=(.*?);\nconst '+re.escape(next_name)+r'='
    m=re.search(pat,text,re.S)
    if not m: raise RuntimeError(f'Impossibile trovare {name}')
    return json.loads(m.group(1)),m.start(1),m.end(1)

def jsjson(o):
    return json.dumps(o,ensure_ascii=False,separators=(',',':'))

def parse_class(ws, name_col=2, total_col=4, nat_col=3, header_row=1):
    out=[]
    for r in range(header_row+1,ws.max_row+1):
        name=clean(ws.cell(r,name_col).value)
        if not name: continue
        total=ws.cell(r,total_col).value
        if total is None: continue
        out.append({'name':name,'nat':clean(ws.cell(r,nat_col).value),'total':int(num(total)),'active':active_font(ws.cell(r,name_col))})
    return rank_rows(out)

def parse_secondary(ws, start_col, limit=285):
    out=[]
    for r in range(2, min(ws.max_row,limit+1)+1):
        name=clean(ws.cell(r,start_col).value); pts=ws.cell(r,start_col+1).value
        if name and pts is not None: out.append({'name':name,'points':int(num(pts))})
    # nationality/activity from general later
    return out

def parse_general(ws):
    rows=[]
    for r in range(2,ws.max_row+1):
        name=clean(ws.cell(r,2).value)
        if not name: continue
        total=ws.cell(r,4).value
        if total is None: continue
        rows.append({'name':name,'nat':clean(ws.cell(r,3).value),'total':int(num(total)),'active':active_font(ws.cell(r,2))})
    return rank_rows(rows)

def parse_race_points(ws):
    # Actual race columns in the main sheet are E:AR (5..44), excluding legend columns.
    races=[(c,header_name(ws.cell(1,c).value)) for c in range(5,45) if ws.cell(1,c).value]
    out={}
    for r in range(2,ws.max_row+1):
        name=clean(ws.cell(r,2).value)
        if not name: continue
        arr=[]
        for c,race in races:
            v=num(ws.cell(r,c).value)
            if v: arr.append({'race':race,'points':int(v) if v.is_integer() else v})
        if arr: out[name]=arr
    return out

def parse_season(ws, general_ws):
    headers=[(c,header_name(ws.cell(1,c).value)) for c in range(3,ws.max_column+1) if ws.cell(1,c).value]
    rows=[]
    race_points={}
    for r in range(2,ws.max_row+1):
        name=clean(ws.cell(r,1).value)
        if not name: continue
        total=num(ws.cell(r,2).value)
        if total<=0: continue
        rows.append({'name':name,'total':int(total) if total.is_integer() else total})
        arr=[]
        for c,race in headers:
            v=num(ws.cell(r,c).value)
            if v: arr.append({'race':race,'points':int(v) if v.is_integer() else v})
        if arr: race_points[name]=arr
    rows=rank_rows(rows)
    top=rows[0]['total'] if rows else 0
    leaders=[x['name'] for x in rows if num(x['total'])==num(top)]
    return rows,race_points,' / '.join(leaders),top

def formula_breakdown(season_ws, general_ws):
    # Uses the workbook's own formulas to reconstruct 1st/2nd/3rd/stages/jerseys.
    out={}
    for r in range(2,season_ws.max_row+1):
        name=clean(season_ws.cell(r,1).value)
        if not name: continue
        for c in range(3,season_ws.max_column+1):
            race=header_name(season_ws.cell(1,c).value)
            if not race or race.lower().startswith('nazionali'): continue
            f=season_ws.cell(r,c).value
            if not isinstance(f,str) or not f.startswith('='): continue
            items=[]
            for coeff,col,rownum in re.findall(r'([0-9]+(?:\.[0-9]+)?)\*\'Classifica generale\'!([A-Z]+)\$(\d+)',f):
                count=num(coeff); rr=int(rownum)
                if count<=0 or rr<2: continue
                try:
                    weight=num(general_ws[f'{col}{rr}'].value)
                except: weight=0
                if weight<=0: continue
                typ={2:'1° posto',3:'2° posto',4:'3° posto',5:'tappe',6:'maglie'}.get(rr,f'componente {rr}')
                items.append({'type':typ,'count':int(count) if count.is_integer() else count,'points':int(count*weight) if float(count*weight).is_integer() else count*weight})
            if items: out.setdefault(name,{})[race]=items
    return out

def merge_breakdowns(all_years):
    career={}
    for y,d in all_years.items():
        for n,races in d.items():
            for race,items in races.items():
                dst=career.setdefault(n,{}).setdefault(race,[])
                for it in items:
                    found=next((x for x in dst if x['type']==it['type']),None)
                    if found:
                        found['count']+=it['count']; found['points']+=it['points']
                    else: dst.append(dict(it))
    return career

def parse_best(ws):
    out=[]
    for r in range(2,ws.max_row+1):
        rank=ws.cell(r,1).value; name=clean(ws.cell(r,2).value); pts=ws.cell(r,3).value; year=ws.cell(r,4).value
        if name and pts is not None and year is not None: out.append({'rank':int(num(rank)) if rank is not None else len(out)+1,'name':name,'points':int(num(pts)),'year':int(num(year)) if str(year).isdigit() else year})
    return out

def parse_no1(ws):
    out=[]
    for r in range(2,ws.max_row+1):
        name=clean(ws.cell(r,29).value)
        if name:
            out.append({'name':name,'first':int(num(ws.cell(r,30).value)),'second':int(num(ws.cell(r,31).value)),'third':int(num(ws.cell(r,32).value)),'total':int(num(ws.cell(r,33).value))})
    return out

def parse_palmares(ws):
    headers=[clean(ws.cell(1,c).value) for c in range(1,ws.max_column+1)]
    idx={h:i+1 for i,h in enumerate(headers) if h}
    out=[]
    keys=['TOUR','GIRO','VUELTA','SANREMO','FIANDRE','ROUBAIX','LIEGI','LOMBARDIA','MONDIALE','OLIMPIADE']
    for r in range(2,ws.max_row+1):
        name=clean(ws.cell(r,2).value)
        if not name: continue
        d={'rank':int(num(ws.cell(r,1).value)) if ws.cell(r,1).value is not None else r-1,'name':name,'tot':int(num(ws.cell(r,3).value))}
        for k in keys:d[k.lower()]=int(num(ws.cell(r,idx[k]).value)) if k in idx else 0
        out.append(d)
    return out

def parse_periods(ws):
    groups={'decenni':[],'lustri':[]}
    for r in list(range(2,15))+list(range(21,45)):
        label=clean(ws.cell(r,1).value)
        if not label: continue
        arr=[]
        for c in [2,4,6,8,10]:
            n=clean(ws.cell(r,c).value); p=ws.cell(r,c+1).value
            if n: arr.append({'name':n,'points':int(num(p))})
        groups['decenni' if r<20 else 'lustri'].append({'period':label,'rows':arr})
    return groups

def build():
    if not XLSX.exists(): raise FileNotFoundError(XLSX)
    wbv=load_workbook(XLSX,data_only=True,read_only=False)
    wbf=load_workbook(XLSX,data_only=False,read_only=False)
    g=wbv['Classifica generale']; gf=wbf['Classifica generale']
    general=parse_general(g)
    race_points=parse_race_points(g)
    active={x['name']:x['active'] for x in general}
    nat={x['name']:x['nat'] for x in general}
    secws=wbv['Altre classifiche']
    sec={}
    for label,col in [('GT',2),('Monumenti',6),('Mondiali + Olimpiadi',10),('Pavé',14),('Ardenne',18),('Brevi corse a tappe',22),('Nazionali',26)]:
        rows=parse_secondary(secws,col)
        for i,x in enumerate(rows,1):
            x.update(rank=i,nat=nat.get(x['name'],''),active=active.get(x['name'],False))
        sec[label]=rows
    pc=wbv['più completi']; rows=[]
    for r in range(2,pc.max_row+1):
        n=clean(pc.cell(r,2).value); p=pc.cell(r,3).value
        if n and p is not None: rows.append({'name':n,'points':int(num(p)),'rank':int(num(pc.cell(r,1).value)) if pc.cell(r,1).value else r-1,'nat':nat.get(n,''),'active':active.get(n,False)})
    sec['Più completi']=rows
    # Seasons
    html=HTML.read_text(encoding='utf-8')
    data,_,_=extract_const(html,'DATA','SEASONS')
    old_seasons,_,_=extract_const(html,'SEASONS','OTHER_CLASSIF')
    years=[]; seasons={}; season_break={}
    for s in wbv.sheetnames:
        if re.fullmatch(r'\d{4}',s):
            ws=wbv[s]; wsf=wbf[s]
            rows,rp,leader,leader_score=parse_season(ws,gf)
            seasons[s]={'rows':rows,'leader':leader,'leaderScore':leader_score,'racePoints':rp}
            season_break[s]=formula_breakdown(wsf,gf)
            years.append({'year':int(s),'first':leader.split(' / ')[0] if leader else '', 'second':rows[1]['name'] if len(rows)>1 else '', 'third':rows[2]['name'] if len(rows)>2 else ''})
    years.sort(key=lambda x:x['year'],reverse=True)
    # U25 / improved mappings
    u25={}; ws=wbv['miglior U25']
    for r in range(3,ws.max_row+1):
        y=ws.cell(r,1).value; n=clean(ws.cell(r,2).value); p=ws.cell(r,3).value
        if y is not None and n and p is not None: u25[str(int(num(y))) if num(y) else str(y)]={'name':n,'points':int(num(p))}
    improved={}; ws=wbv['most improved']
    for r in range(3,ws.max_row+1):
        y=ws.cell(r,1).value; n=clean(ws.cell(r,2).value); p=ws.cell(r,3).value
        if y is not None and n and p is not None: improved[str(int(num(y))) if num(y) else str(y)]={'name':n,'points':int(num(p))}
    for y,d in seasons.items():
        d['u25']=u25.get(y); d['improved']=improved.get(y)
        if y in old_seasons and old_seasons[y].get('nation') is not None: d['nation']=old_seasons[y].get('nation')
    # Data template / preserve profile and legend/periods, replace live sections.
    data['general']=general
    data['season26']=seasons.get('2026',{}).get('rows',[])
    for x in data['season26']:
        x.update(nat=nat.get(x['name'],''),active=active.get(x['name'],False))
    data['years']=years
    data['best']=parse_best(wbv['migliori stagioni'])
    data['no1']=parse_no1(wbv['anni al N°1'])
    data['palmares']=parse_palmares(wbv['palmares storico'])
    data['mostComplete']=sec['Più completi']
    # rebuild data periods if present
    data['periods']=parse_periods(wbv['decenni e lustri'])
    html=replace_const(html,'DATA','SEASONS',data)
    html=replace_const(html,'SEASONS','OTHER_CLASSIF',seasons)
    html=replace_const(html,'OTHER_CLASSIF','PERIOD_CLASSIFICATIONS',sec)
    html=replace_const(html,'PERIOD_CLASSIFICATIONS','SEASON_BREAKDOWN',parse_periods(wbv['decenni e lustri']))
    # The constant named PERIOD_CLASSIFICATIONS is used by period rendering; correct it below by explicit replacement.
    html=replace_const(html,'SEASON_BREAKDOWN','RACE_BREAKDOWN',season_break)
    career_break=merge_breakdowns(season_break)
    html=replace_const(html,'RACE_BREAKDOWN','RACE_POINTS',career_break)
    m=re.search(r'(const RACE_POINTS=)(.*?)(;\s*const currentLeader)',html,re.S)
    if not m: raise RuntimeError('RACE_POINTS non trovato')
    html=html[:m.start(2)]+jsjson(race_points)+html[m.end(2):]
    if 'data-excel-live-link' not in html:
        live='''<script data-excel-live-link>
(function(){let last=null;async function checkExcel(){try{const r=await fetch('/version?'+Date.now(),{cache:'no-store'});if(!r.ok)return;const v=await r.text();if(last!==null&&v!==last)location.reload();last=v}catch(e){}}checkExcel();setInterval(checkExcel,3000)})();
</script>'''; html=html.replace('</script></body></html>','</script>'+live+'</body></html>')
    HTML.write_text(html,encoding='utf-8')
    return {'general':len(general),'active':sum(active.values()),'seasons':len(seasons),'races':len(race_points),'mtime':XLSX.stat().st_mtime}

def replace_const(text,name,next_name,obj):
    pat=r"(const "+re.escape(name)+r"=)(.*?)(;\s*const "+re.escape(next_name)+r"=)"
    m=re.search(pat,text,re.S)
    if not m: raise RuntimeError(f"Costante non trovata: {name} -> {next_name}")
    return text[:m.start(2)]+jsjson(obj)+text[m.end(2):]

def replace_between(text,start,end,new):
    pat=re.escape(start)+r"(.*?)"+re.escape(end)
    m=re.search(pat,text,re.S)
    if not m: raise RuntimeError(f"Delimitatore non trovato: {start} / {end}")
    return text[:m.start(1)]+new+text[m.end(1):]

def serve():
    import http.server,socketserver
    from urllib.parse import urlparse
    class Handler(http.server.SimpleHTTPRequestHandler):
        last=0
        def log_message(self,*args): pass
        def do_GET(self):
            if self.path.startswith('/version'):
                mt=XLSX.stat().st_mtime
                if mt!=Handler.last:
                    build(); Handler.last=mt
                body=str(mt).encode()
                self.send_response(200); self.send_header('Content-Type','text/plain'); self.send_header('Cache-Control','no-store'); self.end_headers(); self.wfile.write(body); return
            if self.path=='/' : self.path='/classifica_revolution.html'
            return super().do_GET()
    os.chdir(BASE)
    Handler.last=XLSX.stat().st_mtime
    build()
    with socketserver.ThreadingTCPServer(('127.0.0.1',8765),Handler) as httpd:
        url='http://127.0.0.1:8765/classifica_revolution.html'
        print('Dashboard:',url)
        threading.Timer(0.8,lambda:webbrowser.open(url)).start()
        httpd.serve_forever()

if __name__=='__main__':
    if '--serve' in sys.argv: serve()
    else: print(build())
