"""HWP5/HWPX를 HWPX 사본으로 정규화. 원본은 보존한다."""
import argparse, json, zipfile
from pathlib import Path
import olefile
from hwpx import HwpxDocument

def normalize(source, output):
    if output.exists() or source.resolve()==output.resolve():
        raise ValueError('새 출력 파일명을 지정하세요')
    signature=source.read_bytes()[:8]
    if signature==bytes.fromhex('D0CF11E0A1B11AE1'):
        with olefile.OleFileIO(source) as ole:
            if not ole.exists('FileHeader'):raise ValueError('HWP5 파일이 아닙니다')
            header=ole.openstream('FileHeader').read()
            flags=int.from_bytes(header[36:40],'little')
            if flags & 6:raise ValueError('암호/배포용 HWP: 보호 해제된 사본이 필요합니다')
        kind='HWP5'
    elif zipfile.is_zipfile(source):
        with zipfile.ZipFile(source) as z:
            if 'Contents/header.xml' not in z.namelist():raise ValueError('HWPX 패키지가 아닙니다')
        kind='HWPX'
    else:raise ValueError('지원 형식은 HWP5/HWPX입니다. HWP97·이미지 문서는 별도 변환/OCR이 필요합니다')
    doc=HwpxDocument.open(source)
    output.parent.mkdir(parents=True,exist_ok=True)
    doc.save_to_path(output)
    report=HwpxDocument.open(output).validate()
    if report.issues:raise ValueError(str(report.issues))
    return {'input_format':kind,'output':str(output),'schema':'passed','visual':'requires rendering'}

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('source',type=Path);ap.add_argument('output',type=Path)
    args=ap.parse_args()
    print(json.dumps(normalize(args.source,args.output),ensure_ascii=False))
