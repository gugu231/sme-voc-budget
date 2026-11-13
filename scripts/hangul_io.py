"""HWP 5 본문 추출 / HWPX 텍스트 조사 및 제한적 치환.
본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
레이아웃 렌더러 또는 범용 HWP/HWPX 편집기가 아닙니다.
"""
import argparse
import hashlib
import json
import re
import struct
import sys
import zipfile
import zlib
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

LIMIT = 64 * 1024 * 1024

def hwpx_read(path):
    with zipfile.ZipFile(path) as z:
        if sum(i.file_size for i in z.infolist()) > LIMIT:
            raise ValueError('압축 해제 크기 제한 초과')
        if len(z.namelist()) != len(set(z.namelist())):
            raise ValueError('중복 ZIP 항목')
        result = []
        for name in sorted(n for n in z.namelist() if re.fullmatch(r'Contents/section\d+\.xml', n)):
            data = z.read(name)
            root = ET.fromstring(data)
            nodes = [e for e in root.iter() if e.tag.endswith('}t')]
            result.append({'section': name, 'sha256': hashlib.sha256(data).hexdigest(),
                           'texts': [{'index': i, 'text': ''.join(e.itertext())} for i, e in enumerate(nodes)],
                           'tables': [{'table_index': ti, 'cells': [
                               {'address': next((dict(x.attrib) for x in c if x.tag.endswith('}cellAddr')), {}),
                                'text': '\n'.join(''.join(t.itertext()) for t in c.iter() if t.tag.endswith('}t'))}
                               for c in table.iter() if c.tag.endswith('}tc')]}
                               for ti, table in enumerate(e for e in root.iter() if e.tag.endswith('}tbl'))]})
        if not result:
            raise ValueError('HWPX section XML 없음')
        return result

def hwp_read(path):
    try:
        import olefile
    except ImportError as e:
        raise RuntimeError('HWP 읽기 의존성 필요: python -m pip install olefile==0.47') from e
    paragraphs = []
    with olefile.OleFileIO(path) as ole:
        header = ole.openstream('FileHeader').read()
        if not header.startswith(b'HWP Document File') or len(header) < 40:
            raise ValueError('지원하지 않는 HWP 헤더')
        if header[35] != 5:
            raise ValueError('HWP 5.x만 지원')
        flags = struct.unpack_from('<I', header, 36)[0]
        if flags & (2 | 4):
            raise ValueError('암호 또는 배포용 HWP는 지원하지 않음. 변환본 필요')
        paths = sorted((p for p in ole.listdir() if len(p) == 2 and p[0] == 'BodyText' and p[1].startswith('Section')),
                       key=lambda p: int(p[1][7:]))
        for path_parts in paths:
            data = ole.openstream(path_parts).read()
            if flags & 1:
                dec = zlib.decompressobj(-15)
                data = dec.decompress(data, LIMIT + 1)
                if len(data) > LIMIT or not dec.eof:
                    raise ValueError('압축 스트림 오류 또는 크기 제한 초과')
            pos = 0
            while pos < len(data):
                if pos + 4 > len(data):
                    raise ValueError('잘린 레코드 헤더')
                h = struct.unpack_from('<I', data, pos)[0]
                pos += 4
                tag, size = h & 1023, h >> 20
                if size == 4095:
                    if pos + 4 > len(data):
                        raise ValueError('잘린 확장 레코드')
                    size = struct.unpack_from('<I', data, pos)[0]
                    pos += 4
                if pos + size > len(data):
                    raise ValueError('잘린 레코드 본문')
                payload = data[pos:pos + size]
                pos += size
                if tag != 67:
                    continue
                if len(payload) % 2:
                    raise ValueError('잘못된 UTF16 길이')
                units = struct.unpack('<' + 'H' * (len(payload)//2), payload)
                clean, j = [], 0
                while j < len(units):
                    c = units[j]
                    if c in (1,2,3,4,5,6,7,8,9,11,12,14,15,16,17,18,19,20,21,22,23):
                        if c == 9:
                            clean.append(32)
                        j += 8
                    else:
                        if c >= 32 or c in (10,13):
                            clean.append(c)
                        j += 1
                text = struct.pack('<'+'H'*len(clean), *clean).decode('utf-16le').strip()
                if text:
                    paragraphs.append({'section': path_parts[1], 'text': text})
    if not paragraphs:
        raise ValueError('추출 가능한 본문 없음')
    return {'mode': 'hwp5-text', 'limitations': '표 셀 좌표·그림·수식·문단 서식·페이지 레이아웃 미지원', 'paragraphs': paragraphs}

def patch_hwpx(source, mapping, output):
    if source.resolve() == output.resolve() or output.exists():
        raise ValueError('원본 또는 기존 파일 덮어쓰기 금지. 새 출력 경로 사용')
    inspected = {r['section']: r for r in hwpx_read(source)}
    edits = json.loads(mapping.read_text(encoding='utf-8-sig'))
    changes = {}
    with zipfile.ZipFile(source) as zin:
        for change in edits:
            name = change['section']
            if name in changes:
                raise ValueError('한 section의 변경은 한 묶음으로 전달')
            if inspected[name]['sha256'] != change['sha256']:
                raise ValueError('양식 변경 감지: inspect를 다시 실행')
            xml = zin.read(name).decode('utf-8')
            # 태그·namespace·표 구조를 다시 직렬화하지 않고 단순 text 노드만 교체.
            pattern = re.compile(r'<(?P<p>[A-Za-z_][\w.-]*):t(?:\s[^<>]*?)?(?:\s*/>|>(?P<body>.*?)</(?P=p):t>)', re.S)
            matches = list(pattern.finditer(xml))
            if len(matches) != len(inspected[name]['texts']):
                raise ValueError('지원하지 않는 텍스트 구조')
            seen, replacements = set(), []
            for edit in change['edits']:
                i, new = edit['index'], edit['new']
                if i in seen or not isinstance(new, str) or any(ord(c)<32 for c in new):
                    raise ValueError('중복 index 또는 줄바꿈·제어문자. 문단 추가는 별도 편집 필요')
                seen.add(i)
                if inspected[name]['texts'][i]['text'] != edit['expected']:
                    raise ValueError('기존 텍스트 불일치')
                m = matches[i]
                if m.group('body') is not None and '<' in m.group('body'):
                    raise ValueError('복합 텍스트 노드 편집 불가')
                old = m.group(0)
                if old.endswith('/>'):
                    replacement = old[:-2]+'>'+escape(new)+'</'+m.group('p')+':t>'
                else:
                    replacement = old[:old.index('>')+1]+escape(new)+'</'+m.group('p')+':t>'
                replacements.append((m.start(),m.end(),replacement))
            for a,b,value in sorted(replacements,reverse=True):
                xml=xml[:a]+value+xml[b:]
            ET.fromstring(xml)
            changes[name]=xml.encode('utf-8')
        output.parent.mkdir(parents=True,exist_ok=True)
        with zipfile.ZipFile(output,'x') as zout:
            for info in zin.infolist():
                zout.writestr(info, changes.get(info.filename,zin.read(info.filename)))
    hwpx_read(output)
    return {'output':str(output),'status':'XML 검증 완료 / 시각 검증 필요',
            'preview':'기존 Preview는 갱신하지 않았으므로 결과 확인에 사용 금지. 편집기에서 열고 저장·렌더링 필요.'}

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('command',choices=['read','inspect','patch'])
    p.add_argument('input',type=Path)
    p.add_argument('--output',type=Path)
    p.add_argument('--mapping',type=Path)
    a=p.parse_args()
    if a.command=='patch':
        if not a.output or not a.mapping:
            p.error('patch에는 --mapping 및 --output 필요')
        result=patch_hwpx(a.input,a.mapping,a.output)
    else:
        result=hwpx_read(a.input) if a.input.suffix.lower()=='.hwpx' else hwp_read(a.input)
    content=json.dumps(result,ensure_ascii=False,indent=2)
    if a.output and a.command!='patch':
        with a.output.open('x',encoding='utf-8') as f:
            f.write(content)
    else:
        print(content)

if __name__=='__main__':
    if hasattr(sys.stdout,'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    main()
