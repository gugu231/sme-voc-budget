"""Read-only readiness checks and classification keys for standard ledgers."""
import argparse
import json
import math
import re
from pathlib import Path

import openpyxl


def normalized_label(value):
    text = re.sub(r"\s+", "", str(value or ""))
    return re.sub(r"\([A-Za-z]+\)$", "", text)


def account_code(value, width):
    text = str(value or "").strip()
    if re.fullmatch(r"\d+", text):
        digits = text
    else:
        match = re.search(r"[（(]\s*(\d+)\s*[）)]\s*$", text)
        if not match:
            raise ValueError("목·세목 숫자 코드 확인 필요")
        digits = match[1]
    if len(digits) > width:
        raise ValueError("목·세목 코드 자릿수 확인 필요")
    return digits.zfill(width)


def mapping_key(unit, category, account, subaccount):
    """Names identify business categories; numeric codes identify accounts."""
    if not normalized_label(unit) or not normalized_label(category):
        raise ValueError("사업단위·비용구분 누락")
    return (normalized_label(unit), normalized_label(category),
            account_code(account, 3), account_code(subaccount, 2))


def check_ledger(path):
    raw = openpyxl.load_workbook(path, data_only=False)
    cached = openpyxl.load_workbook(path, data_only=True)
    errors = []
    rows = []
    totals = [0] * 5
    try:
        sheet = raw['원장']
        values = cached['원장']
        headers = {sheet.cell(3, c).value: c for c in range(1, sheet.max_column + 1)}
        required = ['목', '세목', '사업단위', '비용구분', '세부항목',
                    '산출근거(최종예산)', '본예산']
        amounts = ['본예산'] + [f'{n}차증감' for n in range(1, 5)]
        for name in required + amounts + ['최종편성예산']:
            if name not in headers:
                errors.append({'type': 'missing_header', 'field': name})
        if errors:
            return {'ready': False, 'errors': errors, 'rows': 0}
        titles = {rng.min_row for rng in sheet.merged_cells.ranges
                  if rng.min_col == 1 and rng.max_col >= 5}
        for r in range(4, sheet.max_row + 1):
            if r in titles:
                continue
            def get(name):
                return sheet.cell(r, headers[name]).value
            # A used row is determined by editable inputs, never by item name alone.
            fields = required + amounts[1:] + [k for k in ['특화형', '비고'] if k in headers]
            if all(get(k) in (None, '') for k in fields):
                continue
            rows.append(r)
            for name in required:
                value = get(name)
                if value is None or (isinstance(value, str) and not value.strip()):
                    errors.append({'type': 'missing_required', 'row': r, 'field': name})
            nums = [get(k) if get(k) is not None else 0 for k in amounts]
            if not all(type(v) in (int, float) and math.isfinite(v) and v == int(v) for v in nums):
                errors.append({'type': 'invalid_amount', 'row': r})
                continue
            for n, v in enumerate(nums):
                totals[n] += int(v)
            final = sum(nums)
            if any(sum(nums[:n + 1]) < 0 for n in range(len(nums))):
                errors.append({'type': 'negative_budget', 'row': r})
            # Calculation-basis text is required, but not parsed or amount-validated.
            try:
                mapping_key(get('사업단위'), get('비용구분'), get('목'), get('세목'))
            except ValueError:
                errors.append({'type': 'classification_key', 'row': r})
            cell = sheet.cell(r, headers['최종편성예산'])
            if cell.data_type != 'f':
                errors.append({'type': 'final_formula_missing', 'row': r})
            result = values.cell(r, headers['최종편성예산']).value
            if type(result) not in (int, float) or result != final:
                errors.append({'type': 'final_cache_missing_or_stale', 'row': r})
        if not rows:
            errors.append({'type': 'empty_ledger'})
        for n, value in enumerate(totals[1:], 1):
            if value:
                errors.append({'type': 'unbalanced_transfer', 'stage': n, 'amount': value})
        return {'ready': not errors, 'rows': len(rows), 'totals': totals,
                'errors': errors,
                'scope': '입력·코드·차수 합계·계산 캐시 검사. 산출내역은 입력 여부만 검사하며 정확성은 작업자가 확인. 원자료 대조, 한도·비율, 승인·재원 검토는 별도 필수.'}
    finally:
        raw.close()
        cached.close()


def require_ready_ledger(path):
    report = check_ledger(path)
    if not report['ready']:
        raise ValueError('원장 사전검증 실패: ' + json.dumps(report['errors'], ensure_ascii=False))
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('ledger', type=Path)
    args = parser.parse_args()
    report = check_ledger(args.ledger)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(0 if report['ready'] else 1)
