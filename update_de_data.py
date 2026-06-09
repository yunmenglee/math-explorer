#!/usr/bin/env python3
"""
全面用初中DE表数据更新初中拓展条目，保留原有图片
"""
import json
import os
import re
import zipfile
import shutil
import xml.etree.ElementTree as ET
from collections import defaultdict

XL_PATH = '/Users/haosny/Desktop/7月正式课目标范围.xlsx'
PROJECT_DIR = '/Users/haosny/Desktop/math-explorer'
DATA_DIR = os.path.join(PROJECT_DIR, 'data')
IMAGES_DIR = os.path.join(PROJECT_DIR, 'images')
EXTRACT_IMAGES_DIR = os.path.join(DATA_DIR, 'extract', 'images')

NS = {
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}

def parse_shared_strings(zf):
    """Parse shared strings XML"""
    ss_xml = zf.read('xl/sharedStrings.xml').decode('utf-8')
    root = ET.fromstring(ss_xml)
    strings = []
    for si in root.findall('.//s:si', NS):
        t = si.find('s:t', NS)
        r_elements = si.findall('.//s:r/s:t', NS)
        if t is not None:
            strings.append(t.text or '')
        elif r_elements:
            text = ''.join(r.text or '' for r in r_elements)
            strings.append(text)
        else:
            strings.append('')
    return strings

def col_to_index(col_str):
    """Convert column letter to 1-based index"""
    result = 0
    for c in col_str:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result

def parse_sheet5(zf, shared_strings):
    """Parse sheet5 (初中DE) - returns list of dicts"""
    sheet_xml = zf.read('xl/worksheets/sheet5.xml').decode('utf-8')
    root = ET.fromstring(sheet_xml)
    
    rows = root.findall('.//s:row', NS)
    print(f"  Found {len(rows)} rows in sheet5")
    
    de_entries = []
    for row in rows:
        r = int(row.get('r', 0))
        if r == 1:  # Skip header
            continue
        
        cells = {}
        for cell in row.findall('s:c', NS):
            col_ref = cell.get('r', '')
            # Extract column letters
            col_match = re.match(r'([A-Z]+)', col_ref)
            if col_match:
                col = col_match.group(1)
                t = cell.get('t', '')
                v = cell.find('s:v', NS)
                val = v.text if v is not None else ''
                if t == 's' and val:
                    idx = int(val)
                    val = shared_strings[idx] if idx < len(shared_strings) else val
                cells[col] = val
        
        def get_cell(col):
            val = cells.get(col, '')
            return val.strip() if val else ''
        
        entry = {
            'row': r,
            'c1': get_cell('C'),   # 一级分类
            'c2': get_cell('D'),    # 二级分类
            't': get_cell('G'),     # 标题
            'd': get_cell('H'),     # 定义/描述
            'img': get_cell('I'),   # 图片引用
            'j': get_cell('J'),     # 一定包含的KP
            'o': get_cell('O'),     # 备注
        }
        
        if entry['t']:  # Must have a title
            de_entries.append(entry)
    
    print(f"  Extracted {len(de_entries)} DE entries from sheet5")
    return de_entries

def parse_drawing3(zf):
    """Parse drawing3.xml to map cell rows to image rIds"""
    drawing_xml = zf.read('xl/drawings/drawing3.xml').decode('utf-8')
    rels_xml = zf.read('xl/drawings/_rels/drawing3.xml.rels').decode('utf-8')
    
    root = ET.fromstring(drawing_xml)
    rels_root = ET.fromstring(rels_xml)
    
    # Map rId to media file
    rId_to_media = {}
    for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rid = rel.get('Id')
        target = rel.get('Target', '')
        if target:
            media_file = target.replace('../media/', '')
            rId_to_media[rid] = media_file
    
    # Map row to image (cells in column I = col 8)
    row_to_image = {}
    for anchor in root.findall('.//xdr:oneCellAnchor', NS):
        from_elem = anchor.find('xdr:from', NS)
        if from_elem is None:
            continue
        col = from_elem.find('xdr:col', NS)
        row = from_elem.find('xdr:row', NS)
        if col is None or row is None:
            continue
        col_val = int(col.text)
        row_val = int(row.text) + 1  # 1-indexed
        
        blip = anchor.find('.//a:blip', NS)
        if blip is not None:
            rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if rid and rid in rId_to_media:
                # Only map if in column I (col 8)
                row_to_image[row_val] = rId_to_media[rid]
    
    print(f"  Found {len(row_to_image)} images mapped in drawing3")
    return row_to_image

def copy_images(zf, row_to_image):
    """Copy images from xl/media to project images dir"""
    os.makedirs(IMAGES_DIR, exist_ok=True)
    copied = 0
    for row, media_file in row_to_image.items():
        try:
            img_data = zf.read(f'xl/media/{media_file}')
            dest = os.path.join(IMAGES_DIR, media_file)
            with open(dest, 'wb') as f:
                f.write(img_data)
            copied += 1
        except KeyError:
            print(f"  Warning: image {media_file} not found in xl/media")
    print(f"  Copied {copied} images to {IMAGES_DIR}")
    return row_to_image

def main():
    print("=" * 60)
    print("STEP 1: Parse Excel data")
    print("=" * 60)
    
    zf = zipfile.ZipFile(XL_PATH, 'r')
    
    print("Parsing shared strings...")
    shared_strings = parse_shared_strings(zf)
    print(f"  Found {len(shared_strings)} shared strings")
    
    print("Parsing sheet5 (初中DE)...")
    de_entries = parse_sheet5(zf, shared_strings)
    
    print("\nParsing drawing3 for image mappings...")
    row_to_image = parse_drawing3(zf)
    
    print("\nCopying images...")
    copy_images(zf, row_to_image)
    
    # Build row -> image file name
    row_to_img_file = row_to_image
    
    # Map row -> title from DE entries
    row_to_title = {e['row']: e['t'] for e in de_entries}
    
    # Map title -> image file from drawing
    title_to_img = {}
    for row, img_file in row_to_img_file.items():
        title = row_to_title.get(row)
        if title:
            title_to_img[title] = img_file
    
    print(f"\n  {len(title_to_img)} DE entries have images mapped")
    
    zf.close()
    
    print("\n" + "=" * 60)
    print("STEP 2: Process entries.json")
    print("=" * 60)
    
    entries_path = os.path.join(DATA_DIR, 'entries.json')
    with open(entries_path, 'r', encoding='utf-8') as f:
        entries = json.load(f)
    
    print(f"  Total entries in JSON: {len(entries)}")
    
    junior_ext = [e for e in entries if e.get('m') == '初中' and e.get('tp') == 'extend']
    print(f"  初中 extend entries: {len(junior_ext)}")
    
    # Build lookup by title for DE entries
    de_by_title = {}
    for e in de_entries:
        t = e['t']
        if t in de_by_title:
            # Merge - if duplicate, keep first (shouldn't happen often)
            pass
        de_by_title[t] = e
    
    print(f"  DE entries by title: {len(de_by_title)}")
    
    # Track stats
    matched = 0
    not_matched = 0
    updated_entries = []
    de_titles_matched = set()
    
    # Process entries - keep order
    for entry in entries:
        if entry.get('m') == '初中' and entry.get('tp') == 'extend':
            t = entry.get('t', '')
            if t in de_by_title:
                de = de_by_title[t]
                # Update with DE data
                new_x_parts = []
                if de.get('j'):
                    new_x_parts.append(de['j'])
                if de.get('o'):
                    new_x_parts.append(de['o'])
                if entry.get('x'):
                    new_x_parts.append(entry['x'])
                
                old_imt = entry.get('imt', '')
                
                new_entry = {
                    'm': '初中',
                    't': de['t'],
                    'd': de['d'],
                    'c1': de['c1'],
                    'c2': de['c2'],
                    'c3': entry.get('c3', ''),
                    'x': '；'.join(p for p in new_x_parts if p),
                    'src': 'de',
                    'tp': 'extend',
                    'img': title_to_img.get(t, ''),
                    'imt': old_imt,  # Preserve existing image data
                    'areas': []
                }
                updated_entries.append(new_entry)
                matched += 1
                de_titles_matched.add(t)
            else:
                # Not in DE, keep original
                updated_entries.append(entry)
                not_matched += 1
        else:
            updated_entries.append(entry)
    
    # Add new entries that are in DE but not in existing entries (already extend)
    # These are DE entries whose title doesn't match any existing 初中 extend
    existing_extend_titles = set(e.get('t', '') for e in entries 
                                  if e.get('m') == '初中' and e.get('tp') == 'extend')
    
    new_added = 0
    for de in de_entries:
        t = de['t']
        if t not in existing_extend_titles and t not in de_titles_matched:
            new_x_parts = []
            if de.get('j'):
                new_x_parts.append(de['j'])
            if de.get('o'):
                new_x_parts.append(de['o'])
            
            new_entry = {
                'm': '初中',
                't': t,
                'd': de['d'],
                'c1': de['c1'],
                'c2': de['c2'],
                'c3': '',
                'x': '；'.join(p for p in new_x_parts if p),
                'src': 'de',
                'tp': 'extend',
                'img': title_to_img.get(t, ''),
                'imt': '',
                'areas': []
            }
            updated_entries.append(new_entry)
            new_added += 1
    
    print(f"\n  Results:")
    print(f"    Matched and updated: {matched}")
    print(f"    Not in DE (kept as-is): {not_matched}")
    print(f"    New entries added: {new_added}")
    
    # Write updated entries
    with open(entries_path, 'w', encoding='utf-8') as f:
        json.dump(updated_entries, f, ensure_ascii=False, indent=2)
    print(f"\n  Written {len(updated_entries)} entries to entries.json")
    
    print("\n" + "=" * 60)
    print("STEP 3: Generate summary.json and categories.json")
    print("=" * 60)
    
    # Build stats
    level_stats = defaultdict(lambda: {'n': 0, 'cs': defaultdict(int)})
    categories = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    
    for entry in updated_entries:
        m = entry.get('m', '')
        c1 = entry.get('c1', '')
        c2 = entry.get('c2', '')
        
        level_stats[m]['n'] += 1
        level_stats[m]['cs'][c1] += 1
        
        categories[m][c1][c2] += 1
    
    # Convert to plain dicts
    summary = {}
    for level, data in sorted(level_stats.items()):
        summary[level] = {
            'n': data['n'],
            'cs': dict(data['cs'])
        }
    
    cats_out = {}
    for level, c1_data in sorted(categories.items()):
        cats_out[level] = {}
        for c1, c2_data in sorted(c1_data.items()):
            cats_out[level][c1] = dict(c2_data)
    
    summary_path = os.path.join(DATA_DIR, 'summary.json')
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"  Written summary.json")
    
    cats_path = os.path.join(DATA_DIR, 'categories.json')
    with open(cats_path, 'w', encoding='utf-8') as f:
        json.dump(cats_out, f, ensure_ascii=False, indent=2)
    print(f"  Written categories.json")
    
    print("\n" + "=" * 60)
    print("FINAL VERIFICATION")
    print("=" * 60)
    
    with open(entries_path, 'r', encoding='utf-8') as f:
        final_entries = json.load(f)
    
    total = len(final_entries)
    junior = [e for e in final_entries if e.get('m') == '初中']
    junior_ext = [e for e in junior if e.get('tp') == 'extend']
    junior_basic = [e for e in junior if e.get('tp') == 'basic']
    de_src = [e for e in junior_ext if e.get('src') == 'de']
    
    print(f"  Total entries: {total}")
    print(f"  初中 entries: {len(junior)}")
    print(f"  初中 basic: {len(junior_basic)}")
    print(f"  初中 extend: {len(junior_ext)}")
    print(f"  初中 extend (src=de): {len(de_src)}")
    
    # Verify summary
    if '初中' in summary:
        print(f"\n  初中 summary: {summary['初中']['n']} total entries")
        for c1, count in sorted(summary['初中']['cs'].items(), key=lambda x: -x[1]):
            print(f"    {c1}: {count}")

if __name__ == '__main__':
    main()
