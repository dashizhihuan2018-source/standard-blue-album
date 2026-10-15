"""检查正式正文小标题是否侵入页眉保留区；与版心溢出审计并行执行。"""
import argparse
import json
import re
import subprocess
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = '{http://www.w3.org/1999/xhtml}'
HEADING = re.compile(r'^[▍\s]*\d+(?:[.．]\d+)+[^\d.．]')


def audit(pdf, page_types=None, body_top_cm=4.2):
    types = json.loads(Path(page_types).read_text()) if page_types else None
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp)/'bbox.xml'
        subprocess.run(['pdftotext', '-bbox-layout', str(pdf), str(target)], check=True)
        pages = list(ET.parse(target).iter(NS+'page'))
    if types is not None and len(types) != len(pages):
        raise ValueError('page_types count differs from PDF page count')
    errors = []
    for number, page in enumerate(pages, 1):
        kind = types[number-1]['t'] if types else 'content'
        if kind in ('cover', 'blank', 'colophon', 'toc', 'chapter', 'back_cover'):
            continue
        for line in page.iter(NS+'line'):
            text = ''.join(word.text or '' for word in line.iter(NS+'word'))
            if HEADING.match(text) and float(line.get('yMin')) < body_top_cm*72/2.54-0.5:
                errors.append({'page': number, 'y_cm': round(float(line.get('yMin'))*2.54/72, 3), 'text': text[:80]})
    return errors


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf')
    parser.add_argument('page_types', nargs='?')
    parser.add_argument('--body-top-cm', type=float, default=4.2)
    args = parser.parse_args()
    failures = audit(args.pdf, args.page_types, args.body_top_cm)
    print(json.dumps({'header_intrusions': failures, 'result': 'FAIL' if failures else 'PASS'}, ensure_ascii=False, indent=2))
    raise SystemExit(bool(failures))
