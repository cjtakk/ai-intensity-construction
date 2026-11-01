"""Audit pre-announcement availability of AI-scanned annual reports.

Inputs beside script: MA_AI_merged_corrected.xlsx and Lookaheadbiaszipped.zip
(or folder Lookaheadbias containing FY2013.xlsx ... FY2024.xlsx).
Outputs: publication_date_audit.csv and lseg_candidate_filings.csv.
No extra Python packages required. CSV files open directly in Excel.
"""
import csv
import io
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict, Counter
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path

HERE = Path(__file__).resolve().parent
MA_FILE = HERE / 'MA_AI_merged_corrected.xlsx'
LSEG_ZIP = HERE / 'Lookaheadbiaszipped.zip'
LSEG_DIR = HERE / 'Lookaheadbias'
MAIN_OUT = HERE / 'publication_date_audit.csv'
CAND_OUT = HERE / 'lseg_candidate_filings.csv'
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
EPOCH = datetime(1899, 12, 30)


def col_num(letter):
    n = 0
    for ch in letter:
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx(source):
    """Read first worksheet from filename or Excel bytes without third-party libraries."""
    if isinstance(source, bytes):
        source = io.BytesIO(source)
    with zipfile.ZipFile(source) as z:
        strs = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            strs = [''.join(t.text or '' for t in si.findall('.//m:t', NS))
                    for si in root.findall('m:si', NS)]
        xml = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        for node in xml.findall('.//m:sheetData/m:row', NS):
            row = {}
            for c in node.findall('m:c', NS):
                pos = re.match(r'([A-Z]+)', c.get('r', ''))
                if not pos:
                    continue
                v = c.find('m:v', NS)
                inline = c.find('m:is', NS)
                raw = v.text if v is not None else (
                    ''.join(t.text or '' for t in inline.findall('.//m:t', NS))
                    if inline is not None else '')
                if c.get('t') == 's' and raw != '':
                    raw = strs[int(raw)]
                row[col_num(pos.group(1))] = raw
            yield row


def as_records(rows, header_pred):
    it = iter(rows)
    for row in it:
        if header_pred(row):
            headers = {i: str(v).strip() for i, v in row.items()}
            for r in it:
                if any(r.values()):
                    yield {name: r.get(i, '') for i, name in headers.items() if name}
            return
    raise ValueError('Header row not found; check LSEG export format.')


def xldate(v):
    if v is None or str(v).strip() == '':
        return None
    try:
        return (EPOCH + timedelta(days=float(v))).date()
    except ValueError:
        pass
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%b-%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(str(v).strip(), fmt).date()
        except ValueError:
            continue
    return None


def norm_name(s):
    s = unicodedata.normalize('NFKD', str(s or '').lower())
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r'\b(publ|public|limited|ltd|plc|inc|corp|corporation|group|holding|holdings|company|co|sa|spa|ag|ab|as|ao|pao|se|nv|oyj|ad|the)\b', ' ', s)
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()


def norm_ric(s):
    return str(s or '').strip().upper().split('^')[0]


def save_csv(path, records, columns):
    with path.open('w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(records)


def main():
    if not MA_FILE.is_file():
        raise FileNotFoundError(f'Place {MA_FILE.name} beside this script.')
    ma = list(as_records(read_xlsx(MA_FILE), lambda r: 'SDC Deal No' in r.values()))
    files = []
    if LSEG_ZIP.exists():
        with zipfile.ZipFile(LSEG_ZIP) as outer:
            for name in sorted(outer.namelist()):
                if re.search(r'FY20\d\d\.xlsx$', name, re.I):
                    files.append((Path(name).name, outer.read(name)))
    elif LSEG_DIR.is_dir():
        files = [(p.name, p.read_bytes()) for p in sorted(LSEG_DIR.glob('FY*.xlsx'))]
    else:
        raise FileNotFoundError('Supply Lookaheadbiaszipped.zip or Lookaheadbias folder.')

    lseg = []
    for filename, data in files:
        fy = int(re.search(r'FY(20\d\d)', filename, re.I).group(1))
        for r in as_records(read_xlsx(data), lambda row: 'Filing Date' in row.values() and 'Company Name' in row.values()):
            if not r.get('Company Name'):
                continue
            r['FY'] = fy
            r['Export_File'] = filename
            r['_filed'] = xldate(r.get('Filing Date'))
            r['_name'] = norm_name(r.get('Company Name'))
            r['_ric'] = norm_ric(r.get('RIC'))
            lseg.append(r)

    by_fy = defaultdict(list)
    for r in lseg:
        by_fy[r['FY']].append(r)

    results, candidates = [], []
    for deal in ma:
        dealid = str(deal.get('SDC Deal No', '')).strip()
        name = deal.get('Target Full Name', '')
        fy = int(float(deal['filing_year']))
        announced = xldate(deal.get('Date Announced'))
        ric = norm_ric(deal.get('RIC'))
        target_name = norm_name(name)
        options = by_fy.get(fy, [])
        # RIC and exact normalized legal names are treated as potentially matching.
        # Fuzzy names are kept only as suggestions, never trusted automatically.
        exact = [r for r in options if ric and r['_ric'] == ric]
        method = 'RIC_EXACT' if exact else ''
        if not exact:
            exact = [r for r in options if target_name and r['_name'] == target_name]
            if exact:
                method = 'NAME_EXACT'
        if not exact:
            scored = sorted(((SequenceMatcher(None, target_name, r['_name']).ratio(), r)
                             for r in options if target_name and r['_name']),
                            key=lambda x: x[0], reverse=True)
            if scored and scored[0][0] >= .86:
                exact = [scored[0][1]]
                method = 'FUZZY_REVIEW'

        exact.sort(key=lambda r: (r['_filed'] is None, r['_filed'] or datetime.max.date()))
        before = [r for r in exact if r['_filed'] and announced and r['_filed'] < announced]
        after = [r for r in exact if r['_filed'] and announced and r['_filed'] >= announced]
        if not exact:
            status = 'NO_LSEG_MATCH_MANUAL_REVIEW'
        elif method == 'FUZZY_REVIEW':
            status = 'FUZZY_COMPANY_MATCH_REVIEW'
        elif before:
            status = 'EARLIER_FILING_CANDIDATE_VERIFY_PDF'
        elif after:
            status = 'POST_ANNOUNCEMENT_FILING_CHECK_EARLIER_RELEASE'
        else:
            status = 'MISSING_DATE_REVIEW'

        rows = {
            'SDC_Deal_No': dealid, 'Target_PermID': deal.get('Target PermID', ''),
            'Target': name, 'Target_RIC': deal.get('RIC', ''),
            'Announcement_Date': announced.isoformat() if announced else '',
            'Report_FY': fy, 'Scanned_PDF': deal.get('AI_Report', ''),
            'Match_Method': method or 'NONE', 'Candidate_Count': len(exact),
            'Earliest_LSEG_Filing': min((r['_filed'] for r in exact if r['_filed']), default=None),
            'Earliest_Pre_Announcement_Filing': min((r['_filed'] for r in before), default=None),
            'Latest_LSEG_Filing': max((r['_filed'] for r in exact if r['_filed']), default=None),
            'Timing_Status': status,
            'Manual_Review_Reason': ('Verify the filing is the exact PDF scanned, especially if multiple filings.'
                                     if exact else 'Possibly manually sourced, renamed company, or not in LSEG exports.'),
        }
        for k in ('Earliest_LSEG_Filing', 'Earliest_Pre_Announcement_Filing', 'Latest_LSEG_Filing'):
            rows[k] = rows[k].isoformat() if rows[k] else ''
        results.append(rows)
        for r in exact:
            candidates.append({
                'SDC_Deal_No': dealid, 'Target': name, 'Report_FY': fy,
                'Announcement_Date': rows['Announcement_Date'], 'Scanned_PDF': rows['Scanned_PDF'],
                'Match_Method': method, 'LSEG_Company': r.get('Company Name', ''),
                'LSEG_RIC': r.get('RIC', ''), 'Filing_Date': r['_filed'].isoformat() if r['_filed'] else '',
                'Document_Date': xldate(r.get('Document Date')).isoformat() if xldate(r.get('Document Date')) else '',
                'Report_Title': r.get('Title', ''), 'File_Size': r.get('Size', ''),
                'Pre_Announcement': bool(r['_filed'] and announced and r['_filed'] < announced),
                'Export_File': r['Export_File'],
            })
    main_cols = ['SDC_Deal_No','Target_PermID','Target','Target_RIC','Announcement_Date',
                 'Report_FY','Scanned_PDF','Match_Method','Candidate_Count','Earliest_LSEG_Filing',
                 'Earliest_Pre_Announcement_Filing','Latest_LSEG_Filing','Timing_Status','Manual_Review_Reason']
    candidate_cols = ['SDC_Deal_No','Target','Report_FY','Announcement_Date','Scanned_PDF',
                      'Match_Method','LSEG_Company','LSEG_RIC','Filing_Date','Document_Date',
                      'Report_Title','File_Size','Pre_Announcement','Export_File']
    save_csv(MAIN_OUT, results, main_cols)
    save_csv(CAND_OUT, candidates, candidate_cols)
    print(f'M&A deals: {len(ma)} | LSEG filing records: {len(lseg)} | Exports: {len(files)}')
    for status, count in Counter(r['Timing_Status'] for r in results).most_common():
        print(f'  {status}: {count}')
    print('Saved:', MAIN_OUT.name, 'and', CAND_OUT.name)
    print('CAUTION: even a pre-announcement filing must be matched to the SAME scanned PDF.')

if __name__ == '__main__':
    main()
