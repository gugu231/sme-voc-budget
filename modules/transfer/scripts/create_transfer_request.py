"""표준 원장 → 공단 예산전용 신청서 HWPX. RHWP를 사용하지 않는다.
본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
"""
import argparse
import copy
import hashlib
import json
import math
import re
from pathlib import Path
import zipfile
from collections import defaultdict
from lxml import etree as E
import openpyxl
from hwpx import HwpxDocument

HP='http://www.hancom.co.kr/hwpml/2011/paragraph'
HH='http://www.hancom.co.kr/hwpml/2011/head'
NS={'hp':HP,'hh':HH}
def q(n):return '{'+HP+'}'+n
def money(x,sign=False):return f'{x:+,}' if sign else f'{x:,}'

STAGES = ['1차증감', '2차증감', '3차증감', '4차증감']
# 공단 양식(세부운영기준 ①~㉑)의 사업단위·비용구분 순서. 명칭은 표준 원장 기준.
CATEGORY_ORDER = [
    ('연수비', '학생연수비'), ('연수비', '교직원연수비'),
    ('프로그램운영비', '현장실습경비'), ('프로그램운영비', '실습재료비'), ('프로그램운영비', '교재개발비'),
    ('프로그램운영비', '내부수당'), ('프로그램운영비', '외부수당'), ('프로그램운영비', '간식비'),
    ('대외활동비', '업체발굴비'), ('대외활동비', '홍보 및 행사비'), ('대외활동비', '연구개발비'),
    ('실습장비구입및유지비', '실습장비 구입비'), ('실습장비구입및유지비', '유지보수비 및 시제품 설계·제작비'),
    ('실습장비구입및유지비', '실습장비 임차비'),
    ('사업관리비', '업무관리수당'), ('사업관리비', '수업대체인력운영비'), ('사업관리비', '전담인력운영비'),
    ('사업관리비', '업무추진비'), ('사업관리비', '일반수용비'),
    ('여비', '국내여비'), ('여비', '국외여비'),
]

def load_ledger(path):
    wb=openpyxl.load_workbook(path,data_only=True)
    s=wb['원장'];headers={s.cell(3,c).value:c for c in range(1,s.max_column+1)}
    required=['목','세목','사업단위','비용구분','세부항목','본예산','최종편성예산']+STAGES
    missing=[x for x in required if x not in headers]
    if missing:raise ValueError('표준 원장 머리글과 다릅니다: '+', '.join(missing))
    if '단위내전용' in headers:raise ValueError('단위내전용 열은 지원하지 않습니다. 해당 차수 증감 열에 합산한 원장을 사용하세요')
    basis='산출근거(최종예산)' if '산출근거(최종예산)' in headers else '산출근거'
    if basis not in headers:raise ValueError('산출근거 머리글 확인')
    def value(r,name):return s.cell(r,headers[name]).value
    title_rows={rng.min_row for rng in s.merged_cells.ranges if rng.min_col==1 and rng.max_col>=5}
    items=[];program=None
    for r in range(4,s.max_row+1):
        if r in title_rows:
            title=value(r,'목')
            if isinstance(title,str) and title.strip() and title.strip()!='합계':
                program=re.sub(r'^\s*[\d-]+\s*','',re.sub(r'\s*\([^()]*\)\s*$','',title.strip())) or title.strip()
            continue
        if not value(r,'세부항목'):
            if any(value(r,k) not in (None,'') for k in ['목','세목','사업단위','비용구분','본예산']+STAGES):
                raise ValueError(f'세부항목 누락: {r}행')
            continue
        if any(value(r,x) in (None,'') for x in ['목','세목','사업단위','비용구분','본예산']):raise ValueError(f'필수 입력 누락: {r}행')
        nums=[value(r,x) for x in ['본예산']+STAGES]
        nums=[0 if x is None else x for x in nums]
        if not all(type(x) in (int,float) and math.isfinite(x) and int(x)==x for x in nums):raise ValueError(f'원 단위 숫자 확인: {r}행')
        nums=list(map(int,nums))
        if any(sum(nums[:i+1])<0 for i in range(len(nums))):raise ValueError(f'차수별 예산 음수: {r}행')
        if value(r,'최종편성예산')!=sum(nums):raise ValueError(f'최종편성예산 불일치: {r}행; Excel 재계산 필요')
        items.append({'row':r,'program':program,'unit':value(r,'사업단위'),'category':value(r,'비용구분'),'name':value(r,'세부항목'),
                      'basis':value(r,basis),'base':nums[0],'deltas':nums[1:],'note':value(r,'비고')})
    if not items:raise ValueError('항목 없음')
    for i,st in enumerate(STAGES):
        if sum(x['deltas'][i] for x in items)!=0:raise ValueError(f'{st} 합계가 0이 아닙니다')
    return items

def last_stage(items):
    used=[i+1 for i in range(len(STAGES)) if any(x['deltas'][i] for x in items)]
    if not used:raise ValueError('증감이 입력된 차수가 없습니다')
    return used[-1]

def build_data(items,reasons,stage):
    stages=[];cumulative=0;total=sum(x['base'] for x in items)
    if total<=0:raise ValueError('총사업비는 0보다 커야 합니다')
    if stage not in range(1,len(STAGES)+1):raise ValueError('신청 차수 확인')
    if not any(x['deltas'][stage-1] for x in items):raise ValueError('신청 차수의 증감 내역이 없습니다')
    name_counts=defaultdict(int)
    for x in items:name_counts[x['name']]+=1
    for i in range(stage):
        cats={};units=defaultdict(int)
        for x in items:
            k=(x['unit'],x['category']);before=x['base']+sum(x['deltas'][:i]);delta=x['deltas'][i]
            c=cats.setdefault(k,{'unit':k[0],'category':k[1],'before':0,'delta':0,'after':0,'details':[]})
            c['before']+=before;c['delta']+=delta;c['after']+=before+delta
            units[x['unit']]+=delta
            if delta:
                stage_reasons=reasons.get(str(i+1),{})
                reason=stage_reasons.get(f"row:{x['row']}")
                if reason is None and name_counts[x['name']]==1:reason=stage_reasons.get(x['name'])
                if not isinstance(reason,str) or not reason.strip():raise ValueError(f"{i+1}차 {x['row']}행 사유 누락 또는 중복 항목명: row:{x['row']} 키로 지정하세요")
                c['details'].append({'name':x['name'],'program':x['program'],'delta':delta,'reason':reason,'row':x['row']})
        positive=sum(max(0,v) for v in units.values());cumulative+=positive
        stages.append({'stage':i+1,'categories':list(cats.values()),'unit_deltas':dict(units),'transfer':positive,'cumulative':cumulative,'ratio':cumulative/total})
    return {'total':total,'target':stage,'stages':stages}

def make_hwpx(template,output,data,school,year,bullets):
    if output.exists() or output.resolve()==template.resolve():raise ValueError('새 출력 파일명을 지정하세요')
    with zipfile.ZipFile(template) as z:parts={n:z.read(n) for n in z.namelist()};infos=z.infolist()
    parser=E.XMLParser(resolve_entities=False,no_network=True)
    root=E.fromstring(parts['Contents/section0.xml'],parser)
    header=E.fromstring(parts['Contents/header.xml'],parser)
    original=list(root)
    table0=root.xpath('.//hp:tbl',namespaces=NS)[1]
    tablebase=copy.deepcopy(table0)
    trbase=copy.deepcopy(tablebase.findall(q('tr'))[1])
    thbase=copy.deepcopy(tablebase.findall(q('tr'))[0])
    # 복제된 스타일 ID는 양식 내부의 기존 정의만 사용한다.
    for cp in header.xpath('.//hh:charPr',namespaces=NS):
        if cp.get('id') in ['14','15','16']:cp.set('textColor','#000000')
    cps=header.find('.//'+('{'+HH+'}charProperties'))
    def add_char(source,size):
        x=copy.deepcopy(header.xpath(f'.//hh:charPr[@id="{source}"]',namespaces=NS)[0])
        identifier=max(int(c.get('id')) for c in cps)+1;x.set('id',str(identifier));x.set('height',str(size));x.set('textColor','#000000')
        cps.append(x);cps.set('itemCnt',str(len(cps)));return str(identifier)
    bodychar=add_char('13',1200);tablechar=add_char('0',950);headchar=add_char('19',1000);titlechar=add_char('12',1600)
    pars=header.find('.//'+('{'+HH+'}paraProperties'))
    def add_para(source,align='LEFT',line=140):
        x=copy.deepcopy(header.xpath(f'.//hh:paraPr[@id="{source}"]',namespaces=NS)[0]);identifier=max(int(c.get('id')) for c in pars)+1;x.set('id',str(identifier))
        for e in x.iter():
            name=E.QName(e).localname
            if name=='align':e.set('horizontal',align)
            if name=='lineSpacing':e.set('value',str(line))
            if name in ['intent','left','right','prev','next'] and e.get('value') is not None:e.set('value','0')
        pars.append(x);pars.set('itemCnt',str(len(pars)));return str(identifier)
    bodypara=add_para('11','LEFT',150);cellpara=add_para('0','LEFT',130);centerpara=add_para('0','CENTER',130);rightpara=add_para('0','RIGHT',130)
    serial=1000
    def para(text,char=bodychar,style=bodypara,breakpage=False):
        nonlocal serial
        serial+=1;p=E.Element(q('p'),id=str(serial),paraPrIDRef=style,styleIDRef='0',pageBreak='1' if breakpage else '0',columnBreak='0',merged='0')
        run=E.SubElement(p,q('run'),charPrIDRef=char);E.SubElement(run,q('t')).text=text;return p
    def fill_cell(cell,lines,row,col,width,height,head=False):
        cell.set('header','1' if head else '0')
        sub=cell.find(q('subList'))
        for e in list(sub):sub.remove(e)
        sub.set('vertAlign','CENTER');sub.set('textWidth','0');sub.set('textHeight','0')
        for line in lines:sub.append(para(line,headchar if head else tablechar,centerpara if head or col<2 else rightpara if col==2 else cellpara))
        cell.find(q('cellAddr')).attrib.update({'colAddr':str(col),'rowAddr':str(row)})
        cell.find(q('cellSpan')).attrib.update({'colSpan':'1','rowSpan':'1'})
        cell.find(q('cellSz')).attrib.update({'width':str(width),'height':str(height)})
        mar=cell.find(q('cellMargin'))
        if mar is not None:mar.attrib.update({'left':'180','right':'180','top':'160','bottom':'160'})
    # 제목의 테두리·용지·꼬리말은 유지하고 제목 내용만 변경.
    titlep=copy.deepcopy(original[0]);titlecell=titlep.xpath('.//hp:tc',namespaces=NS)[0]
    sub=titlecell.find(q('subList'))
    for e in list(sub):sub.remove(e)
    target=data['target']
    for line in [f'{year}년 중소기업 특성화고인력양성사업', f'예산전용 신청서 (참여학교용) - {target}차']:
        sub.append(para(line,titlechar,centerpara))
    titlecell.find(q('cellSz')).set('height','4800');titlep.xpath('.//hp:tbl/hp:sz',namespaces=NS)[0].set('height','5200')
    for e in list(root):root.remove(e)
    root.append(titlep);root.append(para(f'{school}  |  {target}차 전용',headchar,rightpara));root.append(para(''))
    root.append(para('1. 전용사유','20'));root.append(para(''))
    for line in bullets:root.append(para('◦ '+line))
    root.append(para(''))
    if target>1:root.append(para(f'※ 1~{target-1}차 표는 원장 기준 재구성 내역이며, 승인 문서와 대조가 필요함.',tablechar))
    root.append(para('※ 본 신청서는 검토용이며, 결재·승인 완료를 의미하지 않음.',tablechar))
    widths=[5400,6500,7000,18500,11130]
    used=[k for k in CATEGORY_ORDER if any(c['delta'] or c['details'] for st in data['stages'] for c in st['categories'] if (c['unit'],c['category'])==k)]
    extra=[]
    for st in data['stages']:
        for c in st['categories']:
            k=(c['unit'],c['category'])
            if c['details'] and k not in CATEGORY_ORDER and k not in extra:extra.append(k)
    for stage in data['stages']:
        n=stage['stage'];root.append(para('2. 전용산출내역 및 세목별 사유' if n==1 else '2. 전용산출내역 및 세목별 사유 (계속)','20',breakpage=True))
        root.append(para(f'2-{n}. {n}차 변경'+(' (신청)' if n==target else ''),'19'))
        root.append(para('단위: 원',tablechar,rightpara))
        cats={(c['unit'],c['category']):c for c in stage['categories']}
        rows=[(['사업단위'],['비용구분'],['증감액'],['산출내역'],['전용사유(상세히 기재)'])]
        for unit,cname in used+extra:
            c=cats.get((unit,cname),{'delta':0,'details':[]})
            if c['details']:
                calc=[f'* {money(c["before"])}원 → {money(c["after"])}원 ({money(c["delta"],True)}원)']
                # 이전 차수는 프로그램 묶음 단위, 신청 차수는 세부항목 단위로 적는다.
                detailtot=defaultdict(int)
                for t in c['details']:
                    label=' / '.join(v for v in [t['program'],t['name'],f"{t['row']}행"] if v)
                    detailtot[label.replace('_',' ')]+=t['delta']
                calc += [f'- {label} {money(delta,True)}원' for label,delta in detailtot.items()]
                reasons=list(dict.fromkeys(t['reason'] for t in c['details'] if t['reason']))
                if not reasons:raise ValueError(f'{n}차 {unit}/{cname}: 작성 사유 없음')
                amount=money(c['delta'],True) if c['delta'] else '0'
            else:calc=[''];reasons=[''];amount=''
            rows.append(([unit],[cname],[amount],calc,reasons))
        rows.append((['소계'],[''],['0'],[''],['']))
        tbl=copy.deepcopy(tablebase)
        for r in tbl.findall(q('tr')):tbl.remove(r)
        tbl.set('id',str(120000+n));tbl.set('rowCnt',str(len(rows)));tbl.set('repeatHeader','1');tbl.set('pageBreak','CELL')
        tbl.find(q('pos')).set('treatAsChar','0');tbl.find(q('pos')).set('horzRelTo','COLUMN')
        heights=[]
        for ri,contents in enumerate(rows):
            tr=copy.deepcopy(thbase if ri==0 else trbase)
            cells=tr.findall(q('tc'))
            # 문자 폭을 고려해 기존 고정 높이보다 작게 만들지 않고 필요한 행 높이를 확보.
            lines=max(sum(max(1,math.ceil(sum(1 if ord(ch)<128 else 1.8 for ch in text)/(widths[ci]/480))) for text in content) for ci,content in enumerate(contents))
            height=max(1400,lines*1050+220);heights.append(height)
            for ci,(cell,content) in enumerate(zip(cells,contents)):fill_cell(cell,content,ri,ci,widths[ci],height,ri==0)
            tbl.append(tr)
        # 동일 사업단위의 첫 열을 병합해 공단 양식의 분류 묶음 보존.
        trs=tbl.findall(q('tr'));start=1
        while start<len(rows)-1:
            stop=start+1
            while stop<len(rows)-1 and rows[stop][0]==rows[start][0]:stop+=1
            if stop-start>1:
                cell=trs[start].findall(q('tc'))[0];cell.find(q('cellSpan')).set('rowSpan',str(stop-start));cell.find(q('cellSz')).set('height',str(sum(heights[start:stop])))
                for rr in range(start+1,stop):trs[rr].remove(trs[rr].findall(q('tc'))[0])
            start=stop
        tbl.find(q('sz')).set('height',str(sum(heights)));tbl.find(q('sz')).set('width',str(sum(widths)))
        p=para('',tablechar);p.find(q('run')).insert(0,tbl);root.append(p)
        root.append(para(f'사업단위 간 전용액 {money(stage["transfer"])}원 / 누적 {money(stage["cumulative"])}원 ({stage["ratio"]:.2%})',tablechar))
    # 오래된 줄배치와 메모 제어를 제거한다. 실제 편집기가 내용을 재배치한다.
    for node in root.xpath('.//hp:linesegarray',namespaces=NS):node.getparent().remove(node)
    parts['Contents/section0.xml']=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
    parts['Contents/header.xml']=E.tostring(header,encoding='UTF-8',xml_declaration=True,standalone=True)
    parts['Preview/PrvText.txt']='\r\n'.join(root.xpath('.//hp:t/text()',namespaces=NS)).encode('utf-8')
    # 기존 예시 이미지를 결과 미리보기로 오인하지 않도록 빈 투명 PNG로 치환. native 재저장 시 갱신.
    import base64
    if 'Preview/PrvImage.png' in parts:parts['Preview/PrvImage.png']=base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2l9sAAAAASUVORK5CYII=')
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'x') as z:
        for info in infos:z.writestr(info,parts[info.filename])
    report=HwpxDocument.open(output).validate()
    if report.issues:raise ValueError(str(report.issues))
    return {'schema':'passed','tables':len(data['stages']),'render':'pending native Hancom open and PDF verification'}

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--ledger',type=Path,required=True);ap.add_argument('--template',type=Path,required=True)
    ap.add_argument('--content',type=Path,required=True);ap.add_argument('--output',type=Path,required=True)
    ap.add_argument('--stage',type=int,choices=[1,2,3,4],help='신청 차수. 생략 시 증감이 입력된 마지막 차수')
    a=ap.parse_args();content=json.loads(a.content.read_text(encoding='utf-8-sig'));items=load_ledger(a.ledger)
    stage=a.stage or last_stage(items)
    data=build_data(items,content['reasons'],stage)
    status=make_hwpx(a.template,a.output,data,content['school'],content['year'],content['bullets'])
    report={'ledger_sha256':hashlib.sha256(a.ledger.read_bytes()).hexdigest(),'data':data,'validation':status,'missing_basis_rows':[x['row'] for x in items if not x['basis']]}
    a.output.with_suffix('.검산.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    cur=data['stages'][-1]
    print(json.dumps({'output':str(a.output),'total':data['total'],'stage':stage,'transfer':cur['transfer'],'cumulative':cur['cumulative'],'cumulative_ratio':round(cur['ratio'],4),**status},ensure_ascii=False))

if __name__=='__main__':main()
