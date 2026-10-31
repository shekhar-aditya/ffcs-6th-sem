#!/usr/bin/env python3
"""Rebuild the course data embedded in index.html from "Course checklist.xlsx".

Usage:  python3 build_data.py            (reads the workbook, rewrites the `const DATA = ...` line)
        python3 build_data.py --check    (only prints the report, does not touch index.html)

To add a course: add its worksheet(s) to the workbook, then add one line to COURSES below.
"""
import json, re, sys
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / 'Course checklist.xlsx'
HTML = ROOT / 'index.html'

# key, code, name, category, mandatory, theory sheet, lab sheet (None = no lab), credits (None = derive from slots)
COURSES = [
    ('PHY', 'BPHY101L/P', 'Physics', 'FC', True, 'phy theory', 'phy lab', 4),
    ('SE', 'BCSE301L/P', 'Software Engineering', 'DC', True, 'software engineering', 'software engineering lab', 4),
    ('CNS', 'BCSE309L/P', 'Cryptography and Network Security', 'DC', True, 'Cryptography and Network Securi', 'Cryptography and Network lab', 4),
    ('EMB', 'BCSE305L', 'Embedded Systems', 'DC', True, 'embeded system', None, 3),
    ('ACC', 'BSTS302P', 'Advanced Competitive Coding II', 'OE', True, 'Advanced competitive coding II', None, 1.5),
    # --- Data Science & AI
    ('FDS', 'BCSE206L', 'Foundations of Data Science', 'DE', False, 'Foundations of Data Science', None, None),
    ('PDS', 'BCSE207L/P', 'Programming for Data Science', 'DE', False, 'Programming for Data Science', 'Programming for Data Sci lab', None),
    ('DM', 'BCSE208L/P', 'Data Mining', 'DE', False, 'Data Mining', 'Data Mining Lab', None),
    ('ML', 'BCSE209L/P', 'Machine Learning', 'DE', False, 'Machine Learning', 'Machine Learning Lab', None),
    ('EDA', 'BCSE331L/P', 'Exploratory Data Analysis', 'DE', False, 'Exploratory Data Analysis', 'Exploratory Data Analysis Lab', 4),
    ('DL', 'BCSE332L/P', 'Deep Learning', 'DE', False, 'Deep Learning', 'Deep Learning Lab', None),
    ('PA', 'BCSE334L', 'Predictive Analytics', 'DE', False, 'Predictive Analytics', None, None),
    ('NLP', 'BCSE409L', 'Natural Language Processing', 'DE', False, 'Natural Language Processing', None, None),
    ('XAI', 'BCSE418L', 'Explainable AI', 'DE', False, 'Explainable Artificial Intellig', None, None),
    ('RL', 'BCSE432E', 'Reinforcement Learning', 'DE', False, 'Reinforcement Learning', 'Reinforcement Learning', None),
    # --- IoT, Cloud & Computing
    ('IOT', 'BCSE312L/P', 'Programming for IoT Boards', 'DE', False, 'Programming for IoT Boards', 'Programming for IoT Boards Lab', None),
    ('FEC', 'BCSE313L', 'Fundamentals of Fog and Edge Computing', 'DE', False, 'Fundamentals of Fog and Edge Co', None, None),
    ('WC', 'BCSE315L', 'Wearable Computing', 'DE', False, 'Wearable Computing', None, None),
    ('SC', 'BCSE316L', 'Design of Smart Cities', 'DE', False, 'Design of Smart Cities', None, None),
    ('CC', 'BCSE408L', 'Cloud Computing', 'DE', False, 'Cloud Computing', None, None),
    ('PC', 'BCSE412L', 'Parallel Computing', 'DE', False, 'Parallel Computing', None, None),
    # --- Cybersecurity & Digital Forensics
    ('PIOT', 'BCSE314L', 'Privacy and Security in IoT', 'DE', False, 'Privacy and Security in IoT', None, None),
    ('IS', 'BCSE317L', 'Information Security', 'DE', False, 'Information Security', None, None),
    ('DP', 'BCSE318L', 'Data Privacy', 'DE', False, 'Data Privacy', None, None),
    ('PT', 'BCSE319L/P', 'Penetration Testing and Vulnerability Analysis', 'DE', False, 'Penetration Testing and Vulnera', 'Penetration Testing and lab', None),
    ('WAS', 'BCSE320L', 'Web Application Security', 'DE', False, 'Web Application Security', None, None),
    ('MAL', 'BCSE321L/P', 'Malware Analysis', 'DE', False, 'Malware Analysis', 'Malware Analysis Lab', None),
    ('DF', 'BCSE322L/P', 'Digital Forensics', 'DE', False, 'Digital Forensics', 'Digital Forensics Lab', None),
    ('DWM', 'BCSE323L', 'Digital Watermarking and Steganography', 'DE', False, 'Digital Watermarking and Stegan', None, None),
    # --- Blockchain & Cryptocurrency
    ('FBT', 'BCSE324L', 'Foundations of Blockchain Technology', 'DE', False, 'Foundations of Blockchain Tech', None, None),
    ('BTC', 'BCSE325L', 'Introduction to Bitcoin', 'DE', False, 'Introduction to Bitcoin', None, None),
    ('BCA', 'BCSE326L', 'Blockchain Architecture Design', 'DE', False, 'Blockchain Architecture Design', None, None),
    ('SCT', 'BCSE327L/P', 'Smart Contracts', 'DE', False, 'Smart Contracts', 'Smart Contracts Lab', None),
    ('CRY', 'BCSE328L', 'Cryptocurrency Technologies', 'DE', False, 'Cryptocurrency Technologies', None, None),
    ('DLT', 'BCSE329L/P', 'Blockchain and Distributed Ledger Technology', 'DE', False, 'Blockchain and Distributed Ledg', 'Blockchain and Distributed lab', None),
    ('PKI', 'BCSE330L', 'Public Key Infrastructure and Trust Management', 'DE', False, 'Public Key Infrastructure and T', None, None),
    # --- Computer Vision & Image Processing
    ('DIP', 'BCSE403L', 'Digital Image Processing', 'DE', False, 'Digital Image Processing', None, None),
    ('CV', 'BCSE407L', 'Computer Vision', 'DE', False, 'Computer Vision', None, None),
    ('MV', 'BCSE417L/P', 'Machine Vision', 'DE', False, 'Machine Vision', 'Machine Vision Lab', None),
    ('SLP', 'BCSE419L/P', 'Speech and Language Processing', 'DE', False, 'Speech and Language Processing', 'Speech and Language Processing ', None),
]

# VIT FFCS theory grid: (day, period). Periods 0-4 morning, 7-11 afternoon (6 = lunch).
THEORY_GRID = [
    ['A1', 'F1', 'D1', 'TB1', 'TG1', None, None, 'A2', 'F2', 'D2', 'TB2', 'TG2', None, None],
    ['B1', 'G1', 'E1', 'TC1', 'TAA1', None, None, 'B2', 'G2', 'E2', 'TC2', 'TAA2', None, None],
    ['C1', 'A1', 'F1', 'V1', 'V2', None, None, 'C2', 'A2', 'F2', 'TD2', 'TBB2', None, None],
    ['D1', 'B1', 'G1', 'TE1', 'TCC1', None, None, 'D2', 'B2', 'G2', 'TE2', 'TCC2', None, None],
    ['E1', 'C1', 'TA1', 'TF1', 'TD1', None, None, 'E2', 'C2', 'TA2', 'TF2', 'TDD2', None, None],
]
SLOT_CELLS = {}
for d, row in enumerate(THEORY_GRID):
    for p, s in enumerate(row):
        if s: SLOT_CELLS.setdefault(s, []).append([d, p])
for n in range(1, 61):  # L1-L30 mornings (periods 0-5), L31-L60 afternoons (periods 7-12), 6 per day
    d, k = ((n - 1) // 6, (n - 1) % 6) if n <= 30 else ((n - 31) // 6, 7 + (n - 31) % 6)
    SLOT_CELLS['L%d' % n] = [[d, k]]


# Real clock minutes per period (theory and lab periods start/end at different times).
T_MIN = {0: (480, 530), 1: (540, 590), 2: (600, 650), 3: (660, 710), 4: (720, 770),
         7: (840, 890), 8: (900, 950), 9: (960, 1010), 10: (1020, 1070), 11: (1080, 1130), 12: (1131, 1140), 13: (1141, 1190)}
L_MIN = {0: (480, 530), 1: (531, 580), 2: (591, 640), 3: (641, 690), 4: (700, 750), 5: (751, 800),
         7: (840, 890), 8: (891, 940), 9: (951, 1000), 10: (1001, 1050), 11: (1060, 1110), 12: (1111, 1160)}


def time_clash(ct, cl):
    """True if any theory period overlaps any lab period in real time (e.g. TE1 11:00-11:50 vs L23 11:40-12:30)."""
    for d1, p1 in ct:
        for d2, p2 in cl:
            if d1 == d2 and p1 in T_MIN and p2 in L_MIN:
                (a, b), (c, e) = T_MIN[p1], L_MIN[p2]
                if a < e and c < b: return True
    return False


def cells(slot):
    out = []
    for s in slot.split('+'):
        s = s.strip()
        if s not in SLOT_CELLS: raise ValueError('unknown slot ' + s + ' in ' + slot)
        out += SLOT_CELLS[s]
    return out


def is_lab_slot(slot, ctype):
    if ctype in ('LO', 'ELA'): return True
    if ctype in ('TH', 'ETH', 'SS'): return False
    return bool(re.match(r'^L\d', slot))


def read_sheet(wb, name):
    """Rows after the 'Slot Detail' header -> [(slot, venue, faculty, is_lab)]."""
    ws = next((w for w in wb.worksheets if w.title.strip() == name.strip() and w.title == name), None) \
        or next((w for w in wb.worksheets if w.title.strip() == name.strip()), None)
    if ws is None: raise KeyError('worksheet not found: ' + repr(name))
    rows, started = [], False
    for r in ws.iter_rows(values_only=True):
        r = list(r) + [None] * 4
        if not started:
            started = str(r[0] or '').strip().lower() == 'slot detail'; continue
        if not r[0] or not r[2]: continue
        slot = re.sub(r'\s+', '', str(r[0])).upper()
        rows.append((slot, str(r[1] or '').strip(), ' '.join(str(r[2]).split()).upper(), is_lab_slot(slot, str(r[3] or '').strip().upper())))
    return rows


def shift(cs):
    """'M' if every cell is a morning period (before lunch), 'E' if every cell is afternoon/evening."""
    halves = {'M' if p < 6 else 'E' for _, p in cs}
    if len(halves) != 1: raise ValueError('slot spans both shifts: %r' % cs)
    return halves.pop()


def theory_credits(slot):
    return 3 if any(p.startswith('T') for p in slot.split('+')) else 2


def main():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    out, report, problems = [], [], []
    for key, code, name, cat, mand, tsheet, lsheet, cr in COURSES:
        try:
            trows = [r for r in read_sheet(wb, tsheet) if not r[3]]
            lrows = [r for r in read_sheet(wb, lsheet) if r[3]] if lsheet else []
        except KeyError as e:
            problems.append('%s: %s — skipped' % (key, e)); continue
        if not trows:
            problems.append('%s %s: theory sheet "%s" has no theory rows (found %d lab rows only) — skipped; fix the sheet and rebuild'
                            % (key, code, tsheet.strip(), len(read_sheet(wb, tsheet)))); continue
        if lsheet and not lrows:
            problems.append('%s %s: lab sheet "%s" is empty — skipped' % (key, code, lsheet.strip())); continue
        units, seen, orphan, same_shift = [], {}, 0, 0
        for t in trows:
            labs = [l for l in lrows if l[2] == t[2]] if lsheet else [None]
            if not labs: orphan += 1; continue
            for l in labs:
                ct = cells(t[0]); cl = cells(l[0]) if l else []
                if time_clash(ct, cl): continue  # theory and lab of the same faculty overlap in time
                if l and shift(ct) == shift(cl): same_shift += 1; continue  # lab must be in the opposite shift
                k = (t[2], t[0], l[0] if l else '')
                if k in seen:
                    v = seen[k]['v']
                    for x in (t[1], l[1] if l else None):
                        if x and x not in v: v.append(x)
                    continue
                u = {'f': t[2], 't': t[0], 'l': l[0] if l else '', 'v': [x for x in (t[1], l[1] if l else None) if x], 'ct': ct, 'cl': cl}
                seen[k] = u; units.append(u)
        if not units:
            problems.append('%s %s: no usable section (no same-faculty theory + opposite-shift lab pair) — skipped' % (key, code)); continue
        th_cr = max(theory_credits(u['t']) for u in units)
        lab_cr = 1 if lsheet else 0
        credits = cr if cr is not None else th_cr + lab_cr
        out.append({'key': key, 'code': code, 'name': name, 'cat': cat, 'credits': credits, 'thCr': credits - lab_cr if lsheet else credits,
                    'labCr': lab_cr, 'mand': mand, 'hasLab': bool(lsheet), 'units': units})
        report.append('%-5s %-11s %-46s %s %4.1f cr  %3d sections%s' % (key, code, name[:46], cat, credits, len(units),
                      (('  (%d theory rows had no same-faculty lab)' % orphan) if orphan else '') +
                      (('  (%d same-shift pairs dropped)' % same_shift) if same_shift else '')))
    print('\n'.join(report))
    if problems: print('\nDATA ISSUES:\n  ' + '\n  '.join(problems))
    data = {'days': ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'], 'courses': out, 'issues': problems, 'source': XLSX.name}
    print('\n%d courses, %d sections' % (len(out), sum(len(c['units']) for c in out)))
    if '--check' in sys.argv: return
    html = HTML.read_text(encoding='utf-8')
    new, n = re.subn(r'^const DATA = .*$', lambda m: 'const DATA = ' + json.dumps(data, separators=(',', ':')) + ';', html, count=1, flags=re.M)
    if n != 1: sys.exit('could not find the "const DATA = ..." line in index.html')
    HTML.write_text(new, encoding='utf-8')
    print('index.html updated')


if __name__ == '__main__':
    main()
