"""
엑셀 데이터 처리 모듈 (Vercel Python Serverless용)
openpyxl만 사용 - 외부 프로세스 없음
"""
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import copy, io

DELETE_AC_KEYWORDS = [
    '이물질','조명','데마','콤','전도','끼임','누수','정전','청소',
    '부주의','충격','화재','멀티포스트','스피커','승차감','버튼',
    '비고장','일정수립'
]
EL_PRIORITY     = ['안전스위치']
ES_PRIORITY_TOP = ['안전스위치','동력전달장치','제동장치','제어장치']
ES_PRIORITY_BOT = ['부속장치']
KEEP_COLS    = [1,3,4,7,8,9,16,23,24,28,29]
KEEP_HEADERS = ['순번','발생일자','시간','장애유형','호선','역사',
                '신고내역','완료일자','완료시간','총복구시간','조치내역']

def build_div(sheet, col_pairs):
    d = {}
    for r in range(2, sheet.max_row+1):
        for (c1,c2,c3) in col_pairs:
            sta=sheet.cell(r,c1).value; ho=sheet.cell(r,c2).value; typ=sheet.cell(r,c3).value
            if sta and ho and typ:
                try:    d[(str(sta).strip(), str(int(ho)))] = typ
                except: d[(str(sta).strip(), str(ho).strip())] = typ
    return d

def get_fault_dict(sheet):
    headers=[sheet.cell(1,c).value for c in range(1,sheet.max_column+1)]
    fault={h:[] for h in headers if h}
    for r in range(2, sheet.max_row+1):
        for c,h in enumerate(headers,1):
            if h:
                v=sheet.cell(r,c).value
                if v: fault[h].append(str(v).strip().lower())
    return fault

def check_del_g(row):
    g=str(row[6] or '').replace(' ','')
    return '기타장애' in g or '인적장애' in g

def check_del_ac(row):
    ac=str(row[28] or '').replace(' ','')
    for kw in DELETE_AC_KEYWORDS:
        if kw in ac: return True
    return '소음' in ac and '장력조정' in ac

def classify_el(ac, fd):
    if not ac: return ''
    ac=str(ac).lower().replace(' ','')
    m=[]
    for cat,kws in fd.items():
        for kw in kws:
            if kw and kw in ac: m.append(cat); break
    if not m: return ''
    for p in EL_PRIORITY:
        if p in m: return p
    return m[0] if len(m)==1 else ', '.join(m)

def classify_es(ac, fd):
    if not ac: return ''
    ac=str(ac).lower().replace(' ','')
    m=[]
    for cat,kws in fd.items():
        for kw in kws:
            if kw and kw in ac: m.append(cat); break
    if not m: return ''
    if len(m)==1: return m[0]
    for p in ES_PRIORITY_TOP:
        if p in m: return p
    nb=[x for x in m if x not in ES_PRIORITY_BOT]
    return nb[0] if nb else m[0]

def get_io(row, div):
    sta=str(row[8] or '').strip(); j=str(row[9] or '')
    parts=j.split('-')
    try:    ho=str(int(parts[-1]))
    except: ho=parts[-1] if parts else ''
    return div.get((sta,ho),'')

def run_pipeline(file_bytes, progress_cb=None):
    def prog(msg, pct):
        if progress_cb: progress_cb(msg, pct)

    prog('파일 읽는 중...', 5)
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb['장애신고']

    prog('구분 데이터 로드 중...', 10)
    el_div  = build_div(wb['EL구분'], [(1,2,3),(5,6,7)])
    es_div  = build_div(wb['ES구분'], [(1,2,3),(5,6,7)])
    el_fd   = get_fault_dict(wb['EL장애'])
    es_fd   = get_fault_dict(wb['ES장애'])

    orig_w={}
    for i in range(1, ws.max_column+1):
        orig_w[i]=ws.column_dimensions[get_column_letter(i)].width

    prog('데이터 필터링 중...', 20)
    rows=[]
    for r in range(3, ws.max_row+1):
        row=[ws.cell(r,c).value for c in range(1,34)]
        if not check_del_g(row) and not check_del_ac(row):
            rows.append(row)

    prog('EL/ES 분리 중...', 35)
    el,es=[],[]
    for row in rows:
        j=str(row[9] or '')
        if 'EL' in j: el.append(row)
        elif 'ES' in j: es.append(row)

    def spl(lst):
        l1,l2=[],[]
        for r in lst:
            h=str(r[7] or '')
            if '1호선' in h: l1.append(r)
            elif '2호선' in h: l2.append(r)
        return l1,l2

    el1,el2=spl(el); es1,es2=spl(es)

    prog('시트 생성 중...', 50)
    hf   = Font(name=ws.cell(2,1).font.name or 'Arial',
                size=ws.cell(2,1).font.size or 9, bold=True)
    hfill= copy.copy(ws.cell(2,1).fill)
    ha   = Alignment(horizontal='center',vertical='center',wrap_text=True)
    df   = Font(name='Arial', size=9)
    odc  = {ci:ws.cell(3,ci) for ci in KEEP_COLS}

    def mk_sheet(name, data, div, fd, fn):
        ws2=wb.create_sheet(name)
        hdrs=KEEP_HEADERS+['옥내/옥외','장애종류']
        for ci,h in enumerate(hdrs,1):
            c=ws2.cell(1,ci,h); c.font=hf; c.fill=hfill; c.alignment=ha
        for ri,row in enumerate(data,2):
            ws2.cell(ri,1,ri-1).font=df
            ws2.cell(ri,1).alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            for ci,oci in enumerate(KEEP_COLS[1:],2):
                val=row[oci-1]; c=ws2.cell(ri,ci,val)
                oc=odc.get(oci)
                if oc: c.font=Font(name=oc.font.name or 'Arial',size=oc.font.size or 9); c.number_format=oc.number_format
                c.alignment=Alignment(horizontal='left' if ci in (7,11) else 'center',vertical='center',wrap_text=True)
            io_c=len(KEEP_COLS)+1
            ws2.cell(ri,io_c,get_io(row,div)).font=df
            ws2.cell(ri,io_c).alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            fc=io_c+1
            ws2.cell(ri,fc,fn(row[28],fd)).font=df
            ws2.cell(ri,fc).alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        kw={1:orig_w.get(1,5),2:orig_w.get(3,10),3:orig_w.get(4,6),
            4:orig_w.get(7,14),5:orig_w.get(8,8),6:orig_w.get(9,9),
            7:40,8:orig_w.get(23,10),9:orig_w.get(24,8),10:orig_w.get(28,9),11:40}
        for ci in range(1,len(hdrs)+1):
            col=get_column_letter(ci)
            if ci in kw: ws2.column_dimensions[col].width=kw[ci]
            elif ci==len(KEEP_COLS)+1: ws2.column_dimensions[col].width=10
            elif ci==len(KEEP_COLS)+2: ws2.column_dimensions[col].width=20
        ws2.row_dimensions[1].height=20
        for ri in range(2,len(data)+2): ws2.row_dimensions[ri].height=33.75

    mk_sheet('1호선EL',el1,el_div,el_fd,classify_el)
    mk_sheet('1호선ES',es1,es_div,es_fd,classify_es)
    mk_sheet('2호선EL',el2,el_div,el_fd,classify_el)
    mk_sheet('2호선ES',es2,es_div,es_fd,classify_es)

    prog('저장 중...', 65)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    wb2=load_workbook(buf)

    prog('미해당 행 삭제 중...', 75)
    stats={}
    for sn in ['1호선EL','1호선ES','2호선EL','2호선ES']:
        ws3=wb2[sn]; fc=ws3.max_column
        to_del=[r for r in range(2,ws3.max_row+1) if not ws3.cell(r,fc).value]
        for r in reversed(to_del): ws3.delete_rows(r)
        for ri in range(2,ws3.max_row+1): ws3.cell(ri,1).value=ri-1
        stats[sn]=ws3.max_row-1

    out=io.BytesIO(); wb2.save(out); out.seek(0)
    prog('완료', 100)
    return out.read(), stats
