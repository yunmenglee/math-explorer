#!/usr/bin/env python3
"""
修正初中DE图片映射，支持多图片显示
- 重新提取drawing3.xml中所有图片，按(row, col)分组，支持多图
- 合并col=7(H列)和col=8(I列)的图片到同一行
- 多图片用逗号分隔存入img字段
- 更新entries.json, summary.json, categories.json
- 更新index.html渲染逻辑
"""
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

XL_PATH = '/Users/haosny/Desktop/7月正式课目标范围.xlsx'
PROJECT_DIR = '/Users/haosny/Desktop/math-explorer'
DATA_DIR = os.path.join(PROJECT_DIR, 'data')
IMAGES_DIR = os.path.join(PROJECT_DIR, 'images')

NS = {
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}

def parse_shared_strings(zf):
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

def parse_sheet5(zf, shared_strings):
    """Parse sheet5 (初中DE) - returns list of dicts with row numbers"""
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
            'c1': get_cell('C'),
            'c2': get_cell('D'),
            't': get_cell('G'),
            'd': get_cell('H'),
            'img_text': get_cell('I'),  # Raw I column text (image reference text)
            'j': get_cell('J'),
            'o': get_cell('O'),   # O column = 备注
        }
        
        if entry['t']:
            de_entries.append(entry)
    
    print(f"  Extracted {len(de_entries)} DE entries from sheet5")
    return de_entries

def parse_drawing3_multi(zf):
    """
    Parse drawing3.xml to map each (col, row) to ALL image file names.
    Returns:
      cell_to_images: dict of (col, row) -> [image_filename, ...]
      row_to_all_images: dict of row -> comma-separated image filenames
    """
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
    
    # Map (col, row) to all image file names
    cell_to_images = defaultdict(list)
    for anchor in root.findall('.//xdr:oneCellAnchor', NS):
        from_elem = anchor.find('xdr:from', NS)
        if from_elem is None:
            continue
        col = from_elem.find('xdr:col', NS)
        row = from_elem.find('xdr:row', NS)
        if col is None or row is None:
            continue
        col_val = int(col.text)
        row_val = int(row.text) + 1  # Convert to 1-indexed
        
        blip = anchor.find('.//a:blip', NS)
        if blip is not None:
            rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if rid and rid in rId_to_media:
                cell_to_images[(col_val, row_val)].append(rId_to_media[rid])
    
    # Aggregate all images for each row (both col 7 and col 8)
    row_to_all_images = defaultdict(list)
    for (col, row), images in sorted(cell_to_images.items()):
        # Include images from col 7 (H=定义) and col 8 (I=图示)
        if col in (7, 8):
            row_to_all_images[row].extend(images)
    
    print(f"  Found {len(cell_to_images)} cell-image mappings in drawing3")
    print(f"  {len(row_to_all_images)} rows have images")
    
    # Show rows with multiple images
    multi_rows = {r: imgs for r, imgs in row_to_all_images.items() if len(imgs) > 1}
    print(f"  {len(multi_rows)} rows have multiple images:")
    for r in sorted(multi_rows.keys())[:10]:
        print(f"    row {r}: {', '.join(multi_rows[r])}")
    if len(multi_rows) > 10:
        print(f"    ... and {len(multi_rows)-10} more")
    
    return dict(row_to_all_images)

def copy_all_images(zf, row_to_all_images):
    """Copy ALL referenced images from xl/media to project images dir"""
    all_images = set()
    for imgs in row_to_all_images.values():
        for img in imgs:
            all_images.add(img)
    
    os.makedirs(IMAGES_DIR, exist_ok=True)
    copied = 0
    for media_file in sorted(all_images):
        try:
            img_data = zf.read(f'xl/media/{media_file}')
            dest = os.path.join(IMAGES_DIR, media_file)
            with open(dest, 'wb') as f:
                f.write(img_data)
            copied += 1
        except KeyError:
            print(f"  Warning: image {media_file} not found in xl/media")
    print(f"  Copied {copied} images to {IMAGES_DIR}")
    return all_images

def main():
    print("=" * 60)
    print("修正初中DE图片映射 - 支持多图片显示")
    print("=" * 60)
    
    zf = zipfile.ZipFile(XL_PATH, 'r')
    
    print("\nSTEP 1: Parse shared strings...")
    shared_strings = parse_shared_strings(zf)
    print(f"  Found {len(shared_strings)} shared strings")
    
    print("\nSTEP 2: Parse sheet5 (初中DE)...")
    de_entries = parse_sheet5(zf, shared_strings)
    
    print("\nSTEP 3: Parse drawing3 for MULTI-IMAGE mappings...")
    row_to_all_images = parse_drawing3_multi(zf)
    
    print("\nSTEP 4: Copy all images...")
    copy_all_images(zf, row_to_all_images)
    
    # Build row -> comma-separated image string
    row_to_img_str = {}
    for row, images in row_to_all_images.items():
        row_to_img_str[row] = ','.join(images)
    
    # Map row -> title from DE entries
    row_to_title = {e['row']: e['t'] for e in de_entries}
    
    # Map title -> image string from drawing (multi-image support)
    title_to_img = {}
    for row, img_str in row_to_img_str.items():
        title = row_to_title.get(row)
        if title:
            if title in title_to_img:
                # Merge multiple image lists for same title (shouldn't happen often)
                existing = title_to_img[title]
                merged = existing + ',' + img_str
                title_to_img[title] = ','.join(dict.fromkeys(merged.split(',')))  # deduplicate
            else:
                title_to_img[title] = img_str
    
    print(f"\n  {len(title_to_img)} DE entries have images mapped")
    
    # Show some examples
    multi_entries = {t: imgs for t, imgs in title_to_img.items() if ',' in imgs}
    print(f"  {len(multi_entries)} entries have MULTIPLE images:")
    for t, imgs in list(multi_entries.items())[:10]:
        print(f"    '{t}': {imgs}")
    
    zf.close()
    
    print("\n" + "=" * 60)
    print("STEP 5: Process entries.json")
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
        if t not in de_by_title:
            de_by_title[t] = e
        else:
            # Duplicate title - merge data
            existing = de_by_title[t]
            # Keep existing but merge non-empty fields from later
            for key in ['c1', 'c2', 'd', 'j', 'o']:
                if not existing.get(key) and e.get(key):
                    existing[key] = e[key]
    
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
                # Build x field from j + o + existing x
                new_x_parts = []
                if de.get('j'):
                    new_x_parts.append(de['j'])
                if de.get('o'):
                    new_x_parts.append(de['o'])
                if entry.get('x'):
                    new_x_parts.append(entry['x'])
                
                old_imt = entry.get('imt', '')
                
                # Get multi-image mapping
                img_value = title_to_img.get(t, '')
                
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
                    'img': img_value,
                    'imt': old_imt,
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
    
    # Add new entries that are in DE but not in existing entries
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
            
            img_value = title_to_img.get(t, '')
            
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
                'img': img_value,
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
    print("STEP 6: Generate summary.json and categories.json")
    print("=" * 60)
    
    level_stats = defaultdict(lambda: {'n': 0, 'cs': defaultdict(int)})
    categories = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    
    for entry in updated_entries:
        m = entry.get('m', '')
        c1 = entry.get('c1', '')
        c2 = entry.get('c2', '')
        
        level_stats[m]['n'] += 1
        level_stats[m]['cs'][c1] += 1
        
        categories[m][c1][c2] += 1
    
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
    print("STEP 7: Update index.html for multi-image display")
    print("=" * 60)
    
    index_path = os.path.join(PROJECT_DIR, 'index.html')
    with open(index_path, 'r', encoding='utf-8') as f:
        html = f.read()
    
    # Line 415: Replace single image rendering with multi-image support
    old_img_line = """${e.imt ? '<img class="im" src="data:image/png;base64,'+e.imt+'" alt="图示">' : (e.img && e.img.startsWith('image') ? '<img class="im" src="images/'+e.img+'" alt="图示">' : '')}"""
    
    new_img_line = """${e.imt ? '<img class="im" src="data:image/png;base64,'+e.imt+'" alt="图示">' : (e.img ? e.img.split(',').map(f => '<img class="im" src="images/'+f.trim()+'" alt="图示">').join('') : '')}"""
    
    if old_img_line in html:
        html = html.replace(old_img_line, new_img_line)
        print("  Updated image rendering to support multi-image (comma-separated)")
    else:
        # Try to find the image rendering line with a pattern
        # Maybe the existing file has slightly different whitespace
        print("  Warning: Could not find exact old image line. Searching for pattern...")
        # Find line containing imt and img
        lines = html.split('\n')
        for i, line in enumerate(lines):
            if 'imt' in line and 'img' in line and '.im' in line:
                print(f"  Found at line {i+1}: {line.strip()[:80]}...")
                # Replace this line
                if 'e.img && e.img.startsWith' in line or 'e.img' in line:
                    indent = line[:len(line) - len(line.lstrip())]
                    # Handle multi-image
                    lines[i] = indent + """${e.imt ? '<img class="im" src="data:image/png;base64,'+e.imt+'" alt="图示">' : (e.img ? e.img.split(',').map(f => '<img class="im" src="images/'+f.trim()+'" alt="图示">').join('') : '')}"""
                    html = '\n'.join(lines)
                    print("  Updated image rendering successfully")
                    break
    
    # Also add onclick handlers to make images clickable for zoom
    # Look for the image rendering and add onclick
    old_imt_click = """'<img class="im" src="data:image/png;base64,'+e.imt+'" alt="图示">'"""
    new_imt_click = """'<img class="im" src="data:image/png;base64,'+e.imt+'" alt="图示" onclick="showImg(\\''+e.imt+'\\')">'"""
    
    if old_imt_click in html:
        html = html.replace(old_imt_click, new_imt_click)
        print("  Added onclick to base64 images")
    
    # For file-based images, add onclick
    old_file_img = """'<img class="im" src="images/'+f.trim()+'" alt="图示">'"""
    new_file_img = """'<img class="im" src="images/'+f.trim()+'" alt="图示" onclick="showImgFile(\\'images/'+f.trim()+'\\')">'"""
    
    if old_file_img in html:
        html = html.replace(old_file_img, new_file_img)
        print("  Added onclick to file images")
    
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print("  Written index.html")
    
    print("\n" + "=" * 60)
    print("STEP 8: VERIFICATION")
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
    
    # Check multi-image entries
    multi_img = [e for e in de_src if e.get('img') and ',' in e['img']]
    single_img = [e for e in de_src if e.get('img') and ',' not in e['img']]
    no_img = [e for e in de_src if not e.get('img')]
    has_imt = [e for e in de_src if e.get('imt')]
    
    print(f"\n  DE entries with:")
    print(f"    Multiple images: {len(multi_img)}")
    print(f"    Single image: {len(single_img)}")
    print(f"    No image: {no_img}")
    print(f"    Has imt (base64): {len(has_imt)}")
    
    # Show multi-image examples
    if multi_img:
        print(f"\n  Multi-image examples (showing first {min(15, len(multi_img))}):")
        for e in multi_img[:15]:
            print(f"    '{e['t']}': {e['img']}")
    
    # Verify stats
    if '初中' in summary:
        total_in_summary = summary['初中']['n']
        actual_junior = len(junior)
        if total_in_summary != actual_junior:
            print(f"\n  WARNING: summary mismatch! summary={total_in_summary}, actual={actual_junior}")
        else:
            print(f"\n  Summary verified: {total_in_summary} 初中 entries")
    
    print("\n" + "=" * 60)
    print("DONE!")
    print("=" * 60)

if __name__ == '__main__':
    main()
