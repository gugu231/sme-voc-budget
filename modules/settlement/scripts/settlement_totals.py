"""원장과 기준일별 정산 입력을 대조한다. HWPX 출력기는 아니다."""
import argparse,json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "scripts"))
from ledger_checks import require_ready_ledger
from collections import defaultdict
from datetime import date
import openpyxl

def calculate(ledger,inputs):
    require_ready_ledger(ledger)
    cutoff=date.fromisoformat(inputs['as_of'])
    sheet=openpyxl.load_workbook(ledger,data_only=True)['원장']
    headers={sheet.cell(3,c).value:c for c in range(1,sheet.max_column+1)}
    for name in ['사업단위','비용구분','세부항목','본예산','최종편성예산']:
        if name not in headers:raise ValueError('원장 머리글 누락: '+name)
    def value(r,name):return sheet.cell(r,headers[name]).value
    title_rows={rng.min_row for rng in sheet.merged_cells.ranges if rng.min_col==1 and rng.max_col>=5}
    budget=defaultdict(int)
    for r in range(4,sheet.max_row+1):
        if r in title_rows:continue
        if not value(r,'세부항목'):
            if any(value(r,k) not in (None,'') for k in ['사업단위','비용구분','본예산']):raise ValueError(f'원장 {r}행 세부항목 누락')
            continue
        unit,cat=value(r,'사업단위'),value(r,'비용구분')
        base=value(r,'본예산')
        if not unit or not cat or not isinstance(base,(float,int)):raise ValueError(f'원장 {r}행 필수 입력 확인')
        budget[(unit,cat)]+=base
    approved=dict(budget);history=[];event_sums=defaultdict(int)
    for event in sorted(inputs['approval_events'],key=lambda e:(e['date'],e['order'])):
        if date.fromisoformat(event['date'])>cutoff:continue
        if not event['document']:raise ValueError('승인/내부결재 문서번호 필요')
        key=(event['unit'],event['category']);before=approved.get(key,0)
        if before!=event['before']:raise ValueError(f'승인 이력 연결 불일치: {key}, {before} != {event["before"]}')
        approved[key]=event['after'];history.append(event)
        event_sums[(event['date'],event['document'])]+=event['after']-before
    if any(event_sums.values()):raise ValueError('결재별 증감 합계가 0이 아닙니다. 지원금 변경은 별도로 구분하세요')
    spending={}
    for item in inputs['spending_by_category']:
        key=(item['unit'],item['category'])
        if key in spending:raise ValueError('집행액 중복 키: '+str(key))
        if not isinstance(item['spent'],(int,float)):raise ValueError('집행액 미확인: '+str(key))
        spending[key]=item['spent']
    rows=[]
    for key in sorted(set(budget)|set(approved)|set(spending)):
        if key not in spending:raise ValueError('집행액 누락: '+str(key)+' (확인된 미집행만 0 입력)')
        a=approved.get(key,0);b=spending[key]
        rows.append({'unit':key[0],'category':key[1],'original':budget.get(key,0),'approved':a,'spent':b,'balance':a-b})
    a=sum(x['approved'] for x in rows);b=sum(x['spent'] for x in rows)
    if a!=inputs['funding_budget']:raise ValueError('재원별 예산액과 비용구분 합계 불일치')
    if b!=inputs['reported_spending']:raise ValueError('정산총괄표 집행액과 비용구분 합계 불일치')
    interest=inputs['interest'];disallowed=inputs['disallowed']
    if not isinstance(interest,(int,float)) or not isinstance(disallowed,(int,float)):raise ValueError('이자·불인정액 확인 필요')
    return {'as_of':str(cutoff),'summary':{'budget':a,'spent':b,'rate':b/a if a else None,'balance':a-b,'disallowed':disallowed,'interest':interest,'return_components_total':a-b+disallowed+interest},'usage':rows,'history':history,'checks':{'approval_deltas_zero':True,'spending_reconciled':True,'negative_balance_categories':[x['category'] for x in rows if x['balance']<0]},'scope':'재원 1종의 검산. 여러 재원은 각각 분리해 실행하며 대응자금 반납액은 별도 처리.'}

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--ledger',type=Path,required=True);p.add_argument('--inputs',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
    if a.output.exists():raise ValueError('새 출력 파일명 지정')
    result=calculate(a.ledger,json.loads(a.inputs.read_text(encoding='utf-8-sig')))
    a.output.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result['summary'],ensure_ascii=False))
