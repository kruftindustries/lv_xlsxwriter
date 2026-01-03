#!/usr/bin/env python3
"""
Generate LabVIEW example VIs from JSON recipes.
Adds SubVI nodes, labels, and updates string constants.
"""

import json
import os
import re
import shutil
import subprocess
import copy
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

# Required packages (installed via pipx for CLI tools)
REQUIRED_PACKAGES = ['pylabview']

def install_requirements():
    """Ensure required packages are installed via pipx."""
    for package in REQUIRED_PACKAGES:
        # Check if already installed via pipx
        result = subprocess.run(['pipx', 'list'], capture_output=True, text=True)
        if package not in result.stdout:
            print(f"Installing required package: {package}")
            subprocess.check_call(['pipx', 'install', package])
            print(f"  Installed {package}")
        else:
            print(f"  {package} already installed")

def run_pylabview(args):
    """Run readRSRC command (installed via pipx)."""
    cmd = ['readRSRC'] + args
    return subprocess.run(cmd, capture_output=True)

# Install requirements on import
install_requirements()

# Configuration
TEMPLATE_VI = 'Example_template.vi'
RECIPES_FILE = 'c_example_recipes.json'
SUBVIS_DIR = 'VIs'
OUTPUT_DIR = 'output_vis'

# Template already has these SubVIs
TEMPLATE_SUBVIS = {
    'workbook new lv.vi',
    'workbook add worksheet lv.vi',
    'workbook close.vi',
    'lxw strerror.vi'
}

# Color definitions for labels (RGB as 0xRRGGBB)
COLORS = {
    'workbook': 0x4A90D9,  # Blue
    'worksheet': 0x50C878,  # Green  
    'format': 0x9B59B6,    # Purple
    'chart': 0xE74C3C,     # Red
    'default': 0x888888    # Gray
}

def get_color_for_function(func_name):
    """Get color based on function category."""
    if func_name.startswith('workbook'):
        return COLORS['workbook']
    elif func_name.startswith('worksheet'):
        return COLORS['worksheet']
    elif func_name.startswith('format'):
        return COLORS['format']
    elif func_name.startswith('chart'):
        return COLORS['chart']
    return COLORS['default']

def c_function_to_subvi(c_function, lv_wrapper=None, use_lv_wrapper=False):
    """Convert C function name to SubVI name."""
    if lv_wrapper and use_lv_wrapper:
        return lv_wrapper.replace('_', ' ') + '.vi'
    return c_function.replace('_', ' ') + '.vi'

def encode_const_value(s):
    """Encode string for BDHb ConstValue field."""
    length = len(s)
    return f"{length:08x}" + s.encode('ascii').hex()

def load_available_subvis(subvis_dir):
    """Load set of available SubVI names."""
    available = set()
    if os.path.exists(subvis_dir):
        for f in os.listdir(subvis_dir):
            if f.endswith('.vi'):
                available.add(f)
    return available

def get_recipe_subvis(recipe, available_subvis):
    """Get list of SubVIs needed for a recipe (excluding template SubVIs)."""
    needed = []
    seen = set()
    
    for step in recipe['steps']:
        c_fn = step.get('c_function', '')
        if c_fn.startswith('UNKNOWN'):
            continue
        
        lv_wrapper = step.get('lv_wrapper', '')
        use_lv = step.get('use_lv_wrapper', False)
        
        subvi_name = c_function_to_subvi(c_fn, lv_wrapper, use_lv)
        
        if subvi_name in available_subvis and subvi_name not in TEMPLATE_SUBVIS and subvi_name not in seen:
            needed.append(subvi_name)
            seen.add(subvi_name)
    
    return needed

def get_output_filename(recipe):
    """Extract output filename from recipe's first step."""
    if recipe['steps']:
        first_step = recipe['steps'][0]
        args = first_step.get('args', [])
        if args and isinstance(args[0], str) and args[0].endswith('.xlsx'):
            return args[0]
    return None

def create_label_element(uid, text, x, y, color):
    """Create a label XML element."""
    label = ET.Element('SL__arrayElement')
    label.set('class', 'label')
    label.set('uid', str(uid))
    
    obj_flags = ET.SubElement(label, 'objFlags')
    obj_flags.text = '2098196'
    
    bounds = ET.SubElement(label, 'bounds')
    bounds.text = f'({x}, {y}, {x + 200}, {y + 15})'
    
    text_rec = ET.SubElement(label, 'textRec')
    text_rec.set('class', 'textHair')
    
    mode = ET.SubElement(text_rec, 'mode')
    mode.text = '68608'
    
    text_elem = ET.SubElement(text_rec, 'text')
    text_elem.text = text
    
    # Set color
    bg_color = ET.SubElement(text_rec, 'bgColor')
    bg_color.text = f'0x{color:06X}'
    
    return label

def add_subvi_entries(root_main, subvi_names, base_uid):
    """Add VIVI entries to LIvi and IUVI entries to LIbd."""
    livi_section = root_main.find('.//LIvi/Section/LVIN')
    libd_section = root_main.find('.//LIbd/Section/BDHP')
    
    if livi_section is None or libd_section is None:
        return []
    
    existing_iuvi = libd_section.find('IUVI')
    if existing_iuvi is None:
        return []
    
    # Count existing VIVI entries for TypeID
    existing_vivi_count = len(livi_section.findall('VIVI'))
    
    subvi_uids = []
    
    for i, subvi_name in enumerate(subvi_names):
        new_uid = base_uid + (i * 100)
        new_offset = f'0x{new_uid:04X}'
        type_id = existing_vivi_count + i + 1
        
        # Add VIVI entry
        new_vivi = ET.SubElement(livi_section, 'VIVI')
        new_vivi.set('LinkSaveFlag', '0')
        new_vivi.set('VILinkLibVersion', '1')
        new_vivi.set('VILinkFieldA', '0')
        new_vivi.set('VILinkField4', '1')
        new_vivi.set('VILinkFieldB', '')
        new_vivi.set('VILinkFieldC', '')
        new_vivi.set('VILinkFieldD', '0')
        new_vivi.set('TypedLinkFlags', '0')

        qual_name = ET.SubElement(new_vivi, 'LinkSaveQualName')
        ET.SubElement(qual_name, 'String').text = subvi_name

        path_ref = ET.SubElement(new_vivi, 'LinkSavePathRef')
        path_ref.set('Ident', 'PTH0')
        path_ref.set('TpVal', '1')
        ET.SubElement(path_ref, 'String')
        ET.SubElement(path_ref, 'String').text = 'VIs'
        ET.SubElement(path_ref, 'String').text = subvi_name

        type_desc = ET.SubElement(new_vivi, 'TypeDesc')
        type_desc.set('TypeID', str(type_id))
        
        # Add IUVI entry
        new_iuvi = copy.deepcopy(existing_iuvi)
        qual_name = new_iuvi.find('LinkSaveQualName')
        qual_name.findall('String')[0].text = subvi_name
        path_ref = new_iuvi.find('LinkSavePathRef')
        path_ref.findall('String')[2].text = subvi_name
        offset_list = new_iuvi.find('LinkOffsetList')
        offset_list.find('Offset').text = new_offset
        libd_section.append(new_iuvi)
        
        subvi_uids.append((subvi_name, new_uid))
    
    return subvi_uids

def add_iuse_nodes(root_bd, subvi_uids, start_y=280):
    """Add iUse nodes to BDHb for each SubVI."""
    diag = root_bd.find('./root[@class="diag"]')
    if diag is None:
        return
    
    zplane = diag.find('zPlaneList')
    nodelist = diag.find('nodeList')
    
    if zplane is None or nodelist is None:
        return
    
    # Find an existing iUse to clone (uid 546 = workbook_add_worksheet_lv)
    orig_iuse = None
    for child in zplane:
        if child.get('class') == 'iUse' and child.get('uid') == '546':
            orig_iuse = child
            break
    
    if orig_iuse is None:
        return
    
    y_pos = start_y
    
    for subvi_name, new_uid in subvi_uids:
        new_iuse = copy.deepcopy(orig_iuse)
        
        # Update UIDs
        def update_uids(elem, base):
            if 'uid' in elem.attrib:
                old_uid = int(elem.get('uid'))
                relative = old_uid - 546  # Offset from original base
                elem.set('uid', str(base + relative))
            for child in elem:
                update_uids(child, base)
        
        update_uids(new_iuse, new_uid)
        
        # Set position
        bounds = new_iuse.find('bounds')
        if bounds is not None:
            bounds.text = f'({y_pos}, 460, {y_pos + 32}, 492)'
        
        zplane.append(new_iuse)
        
        # Add to nodeList
        uid_ref = ET.SubElement(nodelist, 'SL__arrayElement')
        uid_ref.set('uid', str(new_uid))
        
        y_pos += 50
    
    # Update counts
    old_zplane = int(zplane.get('elements'))
    zplane.set('elements', str(old_zplane + len(subvi_uids)))
    
    old_nodelist = int(nodelist.get('elements'))
    nodelist.set('elements', str(old_nodelist + len(subvi_uids)))

def add_labels_to_bdhb(root_bd, recipe, base_uid=10000):
    """Add color-coded labels to block diagram - DISABLED FOR NOW."""
    # Labels have complex format requirements - skip for now
    return 0

def update_string_constant(content, fphb_content, bdhb_content, old_filename, new_filename):
    """Update string constant in all three locations."""
    # 1. Replace in main XML (DFDS)
    content = content.replace(old_filename, new_filename)
    
    # 2. Replace in FPHb
    fphb_content = fphb_content.replace(old_filename, new_filename)
    
    # 3. Replace in BDHb (ConstValue)
    old_const = encode_const_value(old_filename)
    new_const = encode_const_value(new_filename)
    bdhb_content = bdhb_content.replace(old_const.upper(), new_const)
    bdhb_content = bdhb_content.replace(old_const.lower(), new_const)
    
    return content, fphb_content, bdhb_content

def generate_vi(recipe, recipe_name, template_dir, output_path, available_subvis):
    """Generate a single VI from a recipe."""
    # Create working copies
    work_prefix = f'work_{recipe_name}'
    
    for f in os.listdir(template_dir):
        if f.startswith('template_') or f == 'template.xml':
            src = os.path.join(template_dir, f)
            dst = os.path.join(template_dir, f.replace('template', work_prefix))
            shutil.copy(src, dst)
    
    xml_path = os.path.join(template_dir, f'{work_prefix}.xml')
    shutil.copy(os.path.join(template_dir, 'template.xml'), xml_path)
    
    # Get SubVIs needed for this recipe
    subvi_names = get_recipe_subvis(recipe, available_subvis)
    
    # Get output filename
    output_filename = get_output_filename(recipe)
    
    # Parse main XML
    tree_main = ET.parse(xml_path)
    root_main = tree_main.getroot()
    
    # Add SubVI entries (VIVI and IUVI)
    subvi_uids = add_subvi_entries(root_main, subvi_names, base_uid=5000)
    
    # Save main XML
    content = ET.tostring(root_main, encoding='unicode')
    content = '<?xml version="1.0" encoding="utf-8"?>\n' + content
    content = content.replace('template_', f'{work_prefix}_')
    content = re.sub(r'(Name=")[^"]*(")', f'\\g<1>{recipe_name}.vi\\2', content, count=1)
    
    # Update MUID
    muid = hash(recipe_name) & 0xFFFFFFFF
    content = re.sub(r'(MUID.*?Value=")\d+(")', f'\\g<1>{muid}\\2', content, count=1, flags=re.DOTALL)
    
    with open(xml_path, 'w') as f:
        f.write(content)
    
    # Parse and modify BDHb
    bdhb_path = os.path.join(template_dir, f'{work_prefix}_BDHb.xml')
    tree_bd = ET.parse(bdhb_path)
    root_bd = tree_bd.getroot()
    
    # Add iUse nodes for SubVIs
    add_iuse_nodes(root_bd, subvi_uids)
    
    # Add labels
    labels_added = add_labels_to_bdhb(root_bd, recipe)
    
    tree_bd.write(bdhb_path, encoding='utf-8', xml_declaration=True)
    
    # Update string constant if needed
    if output_filename and output_filename != 'hello_world.xlsx':
        with open(xml_path, 'r') as f:
            content = f.read()
        
        fphb_path = os.path.join(template_dir, f'{work_prefix}_FPHb.xml')
        with open(fphb_path, 'r') as f:
            fphb_content = f.read()
        
        with open(bdhb_path, 'r') as f:
            bdhb_content = f.read()
        
        content, fphb_content, bdhb_content = update_string_constant(
            content, fphb_content, bdhb_content,
            'hello_world.xlsx', output_filename
        )
        
        with open(xml_path, 'w') as f:
            f.write(content)
        with open(fphb_path, 'w') as f:
            f.write(fphb_content)
        with open(bdhb_path, 'w') as f:
            f.write(bdhb_content)
    
    # Compile VI
    result = subprocess.run(
        ['readRSRC', '-m', xml_path, '-i', output_path, '-c'],
        capture_output=True, text=True
    )
    
    # Cleanup work files
    for f in os.listdir(template_dir):
        if f.startswith(work_prefix):
            os.remove(os.path.join(template_dir, f))
    
    if result.returncode == 0:
        return True, labels_added, len(subvi_uids), output_filename
    else:
        return False, 0, 0, result.stderr[:200]

def main():
    # Setup
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Load recipes
    with open(RECIPES_FILE) as f:
        recipes = json.load(f)
    
    # Load available SubVIs
    available_subvis = load_available_subvis(SUBVIS_DIR)
    print(f"Loaded {len(available_subvis)} available SubVIs")
    
    # Extract template if not already done
    if not os.path.exists('template.xml'):
        run_pylabview(['-i', TEMPLATE_VI, '-m', 'template.xml', '-x'])
    
    # Process each recipe
    success_count = 0
    fail_count = 0
    
    for i, recipe in enumerate(recipes):
        # Get recipe name from file field
        recipe_name = recipe['file'].replace('.c', '')
        output_path = os.path.join(OUTPUT_DIR, f'{recipe_name}.vi')
        
        success, labels, subvis, info = generate_vi(
            recipe, recipe_name, '.', output_path, available_subvis
        )
        
        if success:
            success_count += 1
            print(f"[{i+1}/{len(recipes)}] {recipe_name} -> {info}... OK ({labels} labels, {subvis} SubVIs)")
        else:
            fail_count += 1
            print(f"[{i+1}/{len(recipes)}] {recipe_name}... FAILED: {info}")
    
    print(f"\nCompleted: {success_count} success, {fail_count} failed")
    
    # Create ZIP
    if success_count > 0:
        shutil.make_archive('xlsxwriter_examples', 'zip', OUTPUT_DIR)
        print(f"\nCreated: xlsxwriter_examples.zip")

if __name__ == '__main__':
    main()
