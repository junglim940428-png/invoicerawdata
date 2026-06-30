import os
import json
import glob
import csv
from datetime import datetime

try:
    import openpyxl
except ImportError:
    openpyxl = None

RAW_DIR  = 'raw_data'
OUT_FILE = 'data/db.json'
SN_KEY   = 'SHIPPING_NO'


def read_xlsx(path):
    """모든 시트를 읽어서 행 합산"""
    wb = openpyxl.load_workbook(path, data_only=True)
    all_rows = []
    headers = []

    for ws in wb.worksheets:
        rows_iter = ws.iter_rows(values_only=True)
        try:
            first_row = next(rows_iter)
        except StopIteration:
            continue  # 빈 시트 건너뜀

        sheet_headers = [str(v or '').strip() for v in first_row]
        # 헤더가 모두 비어있으면 건너뜀
        if not any(sheet_headers):
            continue

        # 첫 번째 유효 시트의 헤더를 사용
        if not headers:
            headers = sheet_headers

        for row in rows_iter:
            if any(v is not None for v in row):
                obj = {sheet_headers[i]: str(row[i] if row[i] is not None else '') for i in range(len(sheet_headers))}
                
                all_rows.append(obj)

    return headers, all_rows


def read_csv(path):
    rows = []
    headers = []
    encodings = ['utf-8-sig', 'cp949', 'euc-kr', 'utf-8']
    for enc in encodings:
        try:
            with open(path, encoding=enc, newline='') as f:
                reader = csv.DictReader(f)
                headers = list(reader.fieldnames or [])
                rows = [dict(r) for r in reader]
            break
        except Exception:
            continue
    return headers, rows


def main():
    files = (
        glob.glob(os.path.join(RAW_DIR, '*.xlsx')) +
        glob.glob(os.path.join(RAW_DIR, '*.xls'))  +
        glob.glob(os.path.join(RAW_DIR, '*.csv'))
    )
    files = sorted(files)

    if not files:
        print(f"raw_data/ 폴더에 처리할 파일이 없습니다.")
        return

    all_rows   = []
    headers    = []
    seen_keys  = set()
    dup_count  = 0
    file_stats = []

    for filepath in files:
        filename = os.path.basename(filepath)
        ext      = os.path.splitext(filepath)[1].lower()
        try:
            if ext in ('.xlsx', '.xls'):
                if openpyxl is None:
                    continue
                h, rows = read_xlsx(filepath)
            elif ext == '.csv':
                h, rows = read_csv(filepath)
            else:
                continue

            if not headers and h:
                headers = h

            key_col = SN_KEY if SN_KEY in (h or headers) else (headers[0] if headers else None)

            added = 0
            for row in rows:
                dup_key = row.get(key_col, '') if key_col else json.dumps(row, ensure_ascii=False)
                if dup_key and dup_key in seen_keys:
                    dup_count += 1
                    continue
                if dup_key:
                    seen_keys.add(dup_key)
                all_rows.append(row)
                added += 1

            
            if '__sheet__' not in headers and any('__sheet__' in r for r in rows):
                headers = headers + ['__sheet__']

            file_stats.append({'file': filename, 'added': added, 'total': len(rows)})
            print(f"✅ {filename}: {added}건 추가 ({len(rows) - added}건 중복 제거)")

        except Exception as e:
            print(f"❌ {filename} 오류: {e}")

    os.makedirs('data', exist_ok=True)
    result = {
        'db':         all_rows,
        'headers':    headers,
        'lastUpdate': datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC'),
        'totalCount': len(all_rows),
        'dupRemoved': dup_count,
        'fileStats':  file_stats
    }
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, separators=(',', ':'))

    print(f"\n📦 최종 DB: {len(all_rows)}건 (중복 제거 {dup_count}건)")


if __name__ == '__main__':
    main()
