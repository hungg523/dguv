import csv
import json
import re

file_path = 'Copy of Khung nang luc.csv'

competencies = {
    'T1': {'name': 'TẦNG 1: NĂNG LỰC CỐT LÕI', 'list': []},
    'T2': {'name': 'TẦNG 2: NĂNG LỰC CHUYÊN MÔN', 'list': []},
    'T3': {'name': 'TẦNG 3: NĂNG LỰC QUẢN LÝ', 'list': []},
    'T4': {'name': 'TẦNG 4: NĂNG LỰC NHẬN THỨC', 'list': []}
}

with open(file_path, 'r', encoding='utf-8-sig') as f:
    reader = csv.reader(f)
    rows = list(reader)

# The structure is in blocks. 
# A block starts with a row where Col 1 is "Tên năng lực" or empty, but let's look for "Định nghĩa" in col 1.
# Actually, the competencies are in cols 2, 3, 4, 5.

def clean_text(text):
    return text.strip().replace('\n', '<br>')

current_comps = {} # col_index -> dict

for i, row in enumerate(rows):
    if len(row) < 6:
        row += [''] * (6 - len(row))
        
    if "Tên năng lực" in row[1] or re.match(r'^\d+\.', row[2]) or "Tư duy" in row[2]:
        # This is a header row for a new block of competencies
        current_comps = {}
        for col_idx, t_key in [(2, 'T1'), (3, 'T2'), (4, 'T3'), (5, 'T4')]:
            if col_idx < len(row) and row[col_idx].strip():
                name = row[col_idx].strip()
                # Remove prefix like "1. ", "9. "
                name = re.sub(r'^\d+\.\s*', '', name)
                if name:
                    comp = {
                        'name': name,
                        'definition': '',
                        'levels': {}
                    }
                    current_comps[col_idx] = (t_key, comp)
                    competencies[t_key]['list'].append(comp)
    elif "Định nghĩa" in row[1]:
        for col_idx, (t_key, comp) in current_comps.items():
            if col_idx < len(row):
                comp['definition'] = clean_text(row[col_idx])
    elif row[0] in ['1', '2', '3', '4', '5']:
        level = row[0]
        level_name = row[1].strip()
        for col_idx, (t_key, comp) in current_comps.items():
            if col_idx < len(row) and row[col_idx].strip():
                comp['levels'][level] = {
                    'name': level_name,
                    'description': clean_text(row[col_idx])
                }

# remove empty competencies
for t_key in competencies:
    competencies[t_key]['list'] = [c for c in competencies[t_key]['list'] if c['name']]

with open('data.js', 'w', encoding='utf-8') as f:
    f.write('const competencyData = ' + json.dumps(competencies, ensure_ascii=False, indent=2) + ';')

print("Parsed successfully to data.js")
