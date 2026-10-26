import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill, Protection
from openpyxl.worksheet.dimensions import ColumnDimension
import re
import io
import copy
import json
import hashlib
import urllib.parse

# --- 1. HAFIZA (SESSION STATE) ---
if 'user_prefs' not in st.session_state:
    st.session_state['user_prefs'] = {}

# --- 2. YARDIMCI VE LOJİSTİK FONKSİYONLAR ---

def get_dim_val(pattern, text):
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        try:
            return float(match.group(1).replace(',', '.'))
        except:
            return None
    return None

def extract_overall_dims(text):
    if pd.isna(text):
        return None, None, None
    text = str(text)
    
    dia = get_dim_val(r'(?:Diameter|Çap|Dia|Ø)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)
    w = get_dim_val(r'(?:Width|Genişlik|Side to Side|\bW\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)
    h = get_dim_val(r'(?:Height|Yükseklik|Top to Bottom|\bH\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)
    d = get_dim_val(r'(?:Depth|Derinlik|Front to Back|\bD\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)
    
    if dia:
        if not w: w = dia
        if not d: d = dia
            
    if not any([w, h, d]):
        size_match = re.search(
            r'(?:Size|Ölçü|Dimension|Boyut)?[:\s]*(\d+(?:[.,]\d+)?)\s*[xX×]\s*(\d+(?:[.,]\d+)?)(?:\s*[xX×]\s*(\d+(?:[.,]\d+)?))?', 
            text
        )
        if size_match:
            h = float(size_match.group(1).replace(',', '.'))
            w = float(size_match.group(2).replace(',', '.'))
            if size_match.group(3): 
                d = float(size_match.group(3).replace(',', '.'))
                
    return h, w, d

def convert_to_inch(val):
    if val: 
        return round(val * 0.393701, 2)
    return None

def translate_features(text, do_conversion):
    if pd.isna(text): 
        return ""
    text = str(text)
    text = re.sub(r'Ø\s*[:\-]?\s*', 'Diameter: ', text)
    text = re.sub(r'(?i)\bÇap\s*[:\-]?\s*', 'Diameter: ', text)
    
    if not do_conversion: 
        return text

    def c_in(m): return f"{round(float(m.group(1).replace(',', '.')) * 0.393701, 2)}"
    def c_mm(m): return f"{round(float(m.group(1).replace(',', '.')) * 0.0393701, 2)}"
    def c_kg(m): return f"{round(float(m.group(1).replace(',', '.')) * 2.20462, 2)}"
    def c_gr(m): return f"{round(float(m.group(1).replace(',', '.')) * 0.00220462, 2)} lbs"
    def c_ml(m): return f"{round(float(m.group(1).replace(',', '.')) * 0.033814, 2)}"

    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*cm(?![a-zA-Z])', lambda m: c_in(m) + ' inches', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*mm(?![a-zA-Z])', lambda m: c_mm(m) + ' inches', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*kg(?![a-zA-Z])', lambda m: c_kg(m) + ' lbs', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*(?:gram[s]?|gr\.?|g)(?![a-zA-Z0-9])', c_gr, text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*ml(?![a-zA-Z])', lambda m: c_ml(m) + ' fl oz', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z0-9:])(\d+(?:[\.,]\d+)?)(?=\s*[xX×]\s*\d)', c_in, text)
    text = re.sub(r'(?<![a-zA-Z])\bcm\b(?![a-zA-Z])', 'inches', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z])\bmm\b(?![a-zA-Z])', 'inches', text, flags=re.IGNORECASE)
    
    return text


def calculate_freight_class_total(total_weight_lbs, total_volume_in3):
    vol_ft3 = total_volume_in3 / 1728
    if vol_ft3 == 0: 
        return "60"
    dens = total_weight_lbs / vol_ft3
    if dens < 1: return "400"
    elif dens < 2: return "300"
    elif dens < 4: return "200"
    elif dens < 6: return "150"
    elif dens < 8: return "125"
    elif dens < 10: return "100"
    else: return "60"

def extract_bedding_info(features, description, raw_h, raw_w):
    text_lower = (str(features) + " " + str(description)).lower()
    total_pieces = 0
    pillow_match = re.search(r'(?:pillowcase|pillow case|yastık kılıfı).*?(\d+)\s*(?:piece|adet|pcs)', text_lower)
    
    if pillow_match: total_pieces += int(pillow_match.group(1))
    elif 'pillowcase' in text_lower or 'yastık kılıfı' in text_lower: total_pieces += 1
        
    if any(x in text_lower for x in ['duvet cover', 'quilt cover', 'nevresim']): total_pieces += 1
    if any(x in text_lower for x in ['fitted sheet', 'çarşaf']): total_pieces += 1
    if any(x in text_lower for x in ['flat sheet', 'düz çarşaf']): total_pieces += 1
    if any(x in text_lower for x in ['bedspread', 'yatak örtüsü']): total_pieces += 1
    if any(x in text_lower for x in ['comforter', 'yorgan']): total_pieces += 1
    if any(x in text_lower for x in ['blanket', 'battaniye']): total_pieces += 1
        
    pieces_str = str(total_pieces) if total_pieces > 0 else ""
    
    set_single = ""
    if total_pieces > 1: set_single = "Set (matching pieces included)"
    elif total_pieces == 1: set_single = "Single Piece"
        
    material = ""
    if 'cotton blend' in text_lower or ('cotton' in text_lower and 'polyester' in text_lower): material = "Cotton Blend"
    elif 'cotton' in text_lower or 'pamuk' in text_lower: material = "Cotton"
    elif 'microfiber' in text_lower or 'mikrofiber' in text_lower: material = "Microfiber"
    elif 'polyester' in text_lower: material = "Polyester"
    elif 'satin' in text_lower or 'saten' in text_lower: material = "Satin"
    elif 'linen' in text_lower or 'keten' in text_lower: material = "Linen"
    elif 'flannel' in text_lower or 'pazen' in text_lower: material = "Flannel"
    elif 'silk' in text_lower or 'ipek' in text_lower: material = "Silk"
    elif 'velvet' in text_lower or 'kadife' in text_lower: material = "Velour"
    elif 'rayon' in text_lower or 'viskon' in text_lower: material = "Rayon"
    
    prod_type = ""
    is_bedding = False
    
    if 'duvet cover' in text_lower or 'nevresim' in text_lower or 'quilt cover' in text_lower: 
        prod_type = "Duvet Cover"; is_bedding = True
    elif 'bedspread' in text_lower or 'yatak örtüsü' in text_lower: 
        prod_type = "Bedspread"; is_bedding = True
    elif 'quilt' in text_lower: 
        prod_type = "Quilt"; is_bedding = True
    elif 'comforter' in text_lower or 'yorgan' in text_lower: 
        prod_type = "Comforter"; is_bedding = True
    elif 'coverlet' in text_lower: 
        prod_type = "Coverlet"; is_bedding = True
    elif 'pillowcase' in text_lower or 'yastık kılıfı' in text_lower or 'sham' in text_lower: 
        prod_type = "Sham"
        
    bed_size = ""
    max_dim = max(float(raw_h or 0), float(raw_w or 0))
    
    if max_dim > 0 and (is_bedding or prod_type == "Sham"):
        if max_dim >= 250: bed_size = "California King"
        elif max_dim >= 230: bed_size = "King"
        elif max_dim >= 200: bed_size = "Queen"
        elif max_dim >= 180: bed_size = "Full / Double"
        else: bed_size = "Twin"
        
    new_name = str(description)
    new_name = re.sub(r'\s*\([A-Z]{2,3}\)\s*', ' ', new_name)
    
    size_words = [r'\bSingle XXL\b', r'\bSingle XL\b', r'\bSingle\b', r'\bDouble\b', r'\bKing\b', r'\bSuper King\b', r'\bMega King\b', r'\bQueen\b', r'\bTwin\b', r'\bFull\b']
    for word in size_words: 
        new_name = re.sub(word, '', new_name, flags=re.IGNORECASE)
        
    new_name = re.sub(r'\s+', ' ', new_name).strip()
    if bed_size and prod_type != "Sham" and is_bedding: 
        new_name = f"{bed_size} {new_name}"
        
    return {'pieces': pieces_str, 'set_single': set_single, 'material': material, 'prod_type': prod_type, 'bed_size': bed_size, 'new_name': new_name}

def generate_bedding_note(text, h_val, w_val, bed_size, is_us):
    if not is_us or not text: return ""
    text_lower = str(text).lower()
    is_bedding = any(k in text_lower for k in ['duvet', 'quilt', 'bedspread', 'nevresim', 'yorgan', 'yatak örtüsü', 'comforter', 'coverlet'])
    is_pillow = any(k in text_lower for k in ['pillowcase', 'yastık kılıfı', 'sham'])

    if not (is_bedding or is_pillow): return ""

    h_in = round(float(h_val) * 0.393701, 2) if h_val else ""
    w_in = round(float(w_val) * 0.393701, 2) if w_val else ""
    dims_str = f"{h_in} x {w_in}" if h_in and w_in else f"{h_in or w_in}"
    
    if is_bedding and bed_size: 
        return f"This beautiful set is crafted to European standards. Measuring {dims_str} inches, it beautifully complements your {bed_size} bed, offering a cozy and elegant drape."
    elif is_pillow: 
        return f"Crafted in Türkiye, these pillowcases measure {dims_str} inches. They are a wonderful fit for standard US pillows, bringing a touch of European comfort to your bedroom."
    return ""

def get_brand_by_category(category_text):
    if pd.isna(category_text) or not str(category_text).strip(): return ""
    cat_lower = str(category_text).lower()
    if "sofa" in cat_lower or "koltuk" in cat_lower or "kanepe" in cat_lower: return "Hanah Home"
    elif "wall deco" in cat_lower or "duvar" in cat_lower or "tablo" in cat_lower: return "Wallity"
    elif "rug" in cat_lower or "carpet" in cat_lower or "halı" in cat_lower or "kilim" in cat_lower: return "Conceptum Hypnose"
    elif "kitchen" in cat_lower or "mutfak" in cat_lower: return "Hermia Concept"
    elif "lighting" in cat_lower or "aydınlatma" in cat_lower or "avize" in cat_lower or "lamba" in cat_lower: return "Opviq"
    elif "furniture" in cat_lower or "mobilya" in cat_lower: return "Skye Decor"
    elif "bathroom" in cat_lower or "banyo" in cat_lower: return "Mijölnir"
    elif "bedroom" in cat_lower or "yatak odası" in cat_lower: return "L'Essentiel Linge de Maison"
    elif "decoration" in cat_lower or "dekorasyon" in cat_lower or "aksesuar" in cat_lower: return "Evila Originals"
    return ""
# --- 3. GÖRSEL URL EŞLEŞTİRME (tek seferlik indeks ile) ---

IMG_COL_KEYWORDS = ('image', 'resim', 'url', 'link')
IMG_COL_EXCLUDE = ('number', 'sayı', 'adet')

def find_image_columns(df_data):
    cols = []
    for col in df_data.columns:
        low = str(col).lower()
        if any(k in low for k in IMG_COL_KEYWORDS) and not any(x in low for x in IMG_COL_EXCLUDE):
            cols.append(col)
    return cols

def _is_url(val):
    return bool(val) and val.lower() != 'nan' and (val.startswith('http') or val.startswith('www'))

def build_url_context(df_data, img_cols):
    """Tüm dosyadaki URL'leri ve SKU'ları bir kez toplar (eski sürüm bunu her satır için baştan yapıyordu)."""
    all_skus = set()
    if 'CODE' in df_data.columns:
        all_skus = {str(s).strip().lower() for s in df_data['CODE'].dropna().unique() if str(s).strip()}
    long_skus = [s for s in all_skus if len(s) >= 4]
    all_urls, seen = [], set()
    for _, r in df_data.iterrows():
        for col in img_cols:
            val = str(r.get(col, '')).strip()
            if _is_url(val) and val not in seen:
                seen.add(val)
                all_urls.append((val, urllib.parse.unquote(val).lower()))
    return {'long_skus': long_skus, 'all_urls': all_urls}

def extract_sku_urls(sku_key, row, img_cols, ctx):
    clean_sku = str(sku_key).strip().lower()
    row_urls = []
    for col in img_cols:
        val = str(row.get(col, '')).strip()
        if _is_url(val) and val not in row_urls:
            row_urls.append(val)

    valid_urls = []
    for url in row_urls:
        unquoted = urllib.parse.unquote(url).lower()
        belongs_to_other = any(s in unquoted for s in ctx['long_skus'] if s != clean_sku)
        if clean_sku in unquoted or not belongs_to_other:
            valid_urls.append(url)

    # Satırda SKU'ya ait URL yoksa tüm dosyada SKU adını içeren URL'leri ara
    if clean_sku and (not valid_urls or not any(clean_sku in urllib.parse.unquote(u).lower() for u in valid_urls)):
        global_urls = [u for u, ul in ctx['all_urls'] if clean_sku in ul]
        if global_urls:
            valid_urls = global_urls

    def get_img_num(url):
        m = re.search(r'[\-_/\s](\d+)\.(?:jpg|jpeg|png|webp)', urllib.parse.unquote(url), re.IGNORECASE)
        return int(m.group(1)) if m else 999

    valid_urls.sort(key=get_img_num)
    return valid_urls


# --- 4. WAYFAIR v3 TEMPLATE YAPISI ---
# Satır 1: özellik ID | 2: tekrar indeksi | 3: grup | 4: statü | 5: alan adı
# Satır 6: açıklama ([Birim] ile başlar) | 7: veri tipi | 8: Default Row | 9+: ürünler
ROW_ID, ROW_IDX, ROW_GROUP, ROW_STATUS, ROW_NAME, ROW_DESC, ROW_TYPE, ROW_DEFAULT = 1, 2, 3, 4, 5, 6, 7, 8
DATA_START_ROW = 9
MAX_DATA_ROW = 2508  # Template'teki biçimlendirme/doğrulama bu satıra kadar tanımlı

ATTR = {
    'spn': '9000', 'mpn': '9001', 'upc': '9005', 'ean': '9069', 'name': '9006', 'brand': '9007',
    'base_cost': '9017', 'msrp': '9019', 'marketing': '9020', 'bullet': '9021',
    'moq': '9026', 'fqm': '9027', 'dsq': '9028', 'product_weight': '9029',
    'ship_type': '9030', 'freight_class': '9031', 'lead_time': '9032', 'repl_lead_time': '9033',
    'carton_weight': '9037', 'carton_height': '9038', 'carton_width': '9039', 'carton_depth': '9040',
    'image': '9053',
    'overall_h': '769', 'overall_w': '770', 'overall_d': '771', 'overall_weight': '264',
    'bed_set_single': '1295', 'bed_pieces': '521', 'bed_type': '8319', 'bed_size': '1791', 'bed_material': '4200',
}
MEDIA_IDS = {'9053', '9058', '9063', '9064', '9065', '9066'}
# Ürün başına girilmemesi gereken / bot tarafından yönetilen diğer alanlar
NON_FORM_IDS = {'9008', '9009', '9010', '9011', '9012', '9004', '9070', '9018', '9034', '9035', '9036', '9020'}
CORE_REQUIRED = ['spn', 'name', 'base_cost', 'product_weight', 'carton_weight', 'image']
US_LEAD_FALLBACK = {'lead_time': 600, 'repl_lead_time': 120}
IMPERIAL_UNITS = {'pounds', 'inches', 'pound', 'inch', 'lbs'}
METRIC_UNITS = {'kilogram', 'kilograms', 'centimetres', 'centimeters', 'centimetre', 'centimeter', 'grams'}
CUSTOM_OPT = "➕ Custom Value"


class TemplateFormatError(Exception):
    pass


def _norm_id(v):
    if v is None:
        return None
    s = str(v).strip()
    if re.fullmatch(r'\d+\.0', s):
        s = s[:-2]
    return s or None

def col_key(attr, idx):
    return attr if idx is None else f"{attr}#{idx}"

def read_template_meta(wb):
    if 'WAYFAIR_USE_ONLY' in wb.sheetnames:
        try:
            return json.loads(wb['WAYFAIR_USE_ONLY']['A1'].value)
        except Exception:
            pass
    return {}

def find_class_sheets(wb, meta):
    sheets = [name for name, m in (meta.get('sheet_metadata') or {}).items()
              if m.get('sheet_type') == 'CLASS_FEED_SHEET' and name in wb.sheetnames]
    if not sheets:
        sheets = [s for s in wb.sheetnames
                  if not any(x in s for x in ["Additional", "WAYFAIR", "Instructions", "Valid Values", "Failed", "Error"])]
    return sheets

def read_allowed_values(wb):
    """allowed{ID}_fp_S_class_{CLASS} tanımlı aralıklarından geçerli değerleri okur."""
    result = {}
    for name, dn in wb.defined_names.items():
        m = re.match(r'allowed(\d+)_fp_.*?class_(\d+)', name)
        if not m:
            continue
        try:
            sheet_part, rng = dn.attr_text.rsplit('!', 1)
            sheet_name = sheet_part.strip("'")
            vals = []
            for row in wb[sheet_name][rng.replace('$', '')]:
                for c in row:
                    if c.value is not None and str(c.value).strip():
                        vals.append(str(c.value).strip())
            result[(m.group(2), m.group(1))] = vals
        except Exception:
            continue
    return result

def parse_template(wb, sheet_name=None):
    meta = read_template_meta(wb)
    sheets = find_class_sheets(wb, meta)
    if not sheets:
        raise TemplateFormatError("Template'te ürün (class) sayfası bulunamadı.")
    sheet_name = sheet_name if sheet_name in sheets else sheets[0]
    ws = wb[sheet_name]

    if _norm_id(ws.cell(ROW_ID, 1).value) != ATTR['spn'] or \
            str(ws.cell(ROW_NAME, 1).value or '').strip().lower() != 'supplier part number':
        raise TemplateFormatError(
            "Bu template beklenen Wayfair formatında değil (A1 = 9000, A5 = 'Supplier Part Number' olmalı). "
            "Partner Home'dan güncel 'Product Addition Template' indirildiğinden emin olun.")

    class_id = str((meta.get('sheet_metadata', {}).get(sheet_name) or {}).get('class_id') or '')
    if not class_id:
        m = re.match(r'\s*(\d+)\s*-', sheet_name)
        class_id = m.group(1) if m else ''
    allowed_all = read_allowed_values(wb)

    columns, by_key, by_attr = [], {}, {}
    unit_votes = {'imperial': 0, 'metric': 0}
    for c in range(1, ws.max_column + 1):
        attr = _norm_id(ws.cell(ROW_ID, c).value)
        if not attr:
            continue
        idx_raw = _norm_id(ws.cell(ROW_IDX, c).value)
        idx = int(idx_raw) if idx_raw and idx_raw.isdigit() else None
        desc = str(ws.cell(ROW_DESC, c).value or '')
        um = re.match(r'\s*\[([^\]]+)\]', desc)
        unit = um.group(1).strip() if um else ''
        if unit.lower() in IMPERIAL_UNITS: unit_votes['imperial'] += 1
        elif unit.lower() in METRIC_UNITS: unit_votes['metric'] += 1
        default_raw = ws.cell(ROW_DEFAULT, c).value
        default = None
        if default_raw is not None and str(default_raw).strip() not in ('', 'N/A', 'Default Row:'):
            default = str(default_raw).strip()
        allowed = allowed_all.get((class_id, attr))
        if allowed is None:
            allowed = next((v for (cl, a), v in allowed_all.items() if a == attr), None)
        col = {
            'c': c, 'letter': get_column_letter(c), 'attr': attr, 'idx': idx, 'key': col_key(attr, idx),
            'status': str(ws.cell(ROW_STATUS, c).value or '').strip(),
            'name': str(ws.cell(ROW_NAME, c).value or '').strip(),
            'desc': desc.strip(), 'unit': unit,
            'type': str(ws.cell(ROW_TYPE, c).value or '').strip(),
            'default': default, 'allowed': allowed,
            'can_default': str(default_raw).strip() != 'N/A' if default_raw is not None else True,
        }
        columns.append(col)
        by_key[col['key']] = col
        by_attr.setdefault(attr, []).append(col)
    for lst in by_attr.values():
        lst.sort(key=lambda x: x['idx'] or 0)

    country = (meta.get('market_context') or {}).get('country', '')
    if unit_votes['imperial'] > unit_votes['metric']: units = 'imperial'
    elif unit_votes['metric'] > unit_votes['imperial']: units = 'metric'
    else: units = None
    if units == 'imperial' or country == 'US': region = 'US'
    elif units == 'metric' or country: region = 'EU'
    else: region = None

    # Ana renk sütunu (Color / Colour / Color / Finish / Top Colour ...)
    color_col = None
    pri = ['color', 'colour', 'color / finish', 'colour / finish', 'finish']
    single = [col for col in columns if col['idx'] is None]
    for p in pri:
        color_col = next((col for col in single if col['name'].lower() == p), None)
        if color_col: break
    if not color_col:
        color_col = next((col for col in single if re.search(r'colou?r', col['name'].lower())
                          and not any(x in col['name'].lower() for x in ['leg', 'base', 'shade', 'frame'])), None)

    return {
        'sheet_name': sheet_name, 'sheets': sheets, 'class_id': class_id, 'meta': meta,
        'country': country, 'locale': (meta.get('market_context') or {}).get('locale', ''),
        'units': units, 'region': region,
        'legacy_format': meta.get('use_legacy_excel_format'),
        'columns': columns, 'by_key': by_key, 'by_attr': by_attr,
        'color_key': color_col['key'] if color_col else None,
        'max_col': ws.max_column,
    }

def auto_keys(tpl):
    """Bot tarafından otomatik doldurulan sütunlar (formda gösterilmez)."""
    keys = {col['key'] for col in tpl['columns']
            if col['attr'] in set(ATTR.values()) or col['attr'] in MEDIA_IDS or col['attr'] in NON_FORM_IDS}
    if tpl['color_key']:
        keys.add(tpl['color_key'])
    return keys


# --- 5. YAZMA YARDIMCILARI ---

def _same_value(a, b):
    def norm(v):
        s = str(v).strip()
        try:
            return '%g' % float(s)
        except ValueError:
            return ';'.join(sorted(p.strip().lower() for p in s.split(';') if p.strip()))
    return norm(a) == norm(b)

def _is_blank(v):
    if v is None: return True
    try:
        if pd.isna(v): return True
    except (TypeError, ValueError):
        pass
    return str(v).strip() == ''

def explode_column_dims(ws):
    """<col min..max> aralıklarını tek tek sütunlara böler; tek sütunu açarken komşuları bozmamak için."""
    new = {}
    for key, d in list(ws.column_dimensions.items()):
        lo = d.min or column_index_from_string(key)
        hi = d.max or lo
        for i in range(lo, hi + 1):
            L = get_column_letter(i)
            nd = ColumnDimension(ws, index=L, width=d.width, bestFit=d.bestFit, hidden=d.hidden,
                                 outlineLevel=d.outlineLevel, collapsed=d.collapsed, min=i, max=i)
            nd._style = copy.copy(d._style)
            new[L] = nd
    ws.column_dimensions.clear()
    for L in sorted(new, key=column_index_from_string):
        ws.column_dimensions[L] = new[L]


# --- 6. ANA İŞLEM ---

def process_wayfair_v20(data_bytes, template_bytes, ui_data, carton_bytes=None, sheet_name=None, progress_callback=None):
    is_us = ui_data['is_us']
    kg2lb = lambda v: round(v * 2.20462, 2)
    cm2in = lambda v: round(v * 0.393701, 2)

    df_data = pd.read_excel(io.BytesIO(data_bytes))
    df_data['__excel_row'] = df_data.index + 2
    if 'CODE' in df_data.columns:
        df_data = df_data.dropna(subset=['CODE'])
        df_data = df_data[df_data['CODE'].astype(str).str.strip() != '']
    else:
        df_data = df_data.dropna(how='all')
    df_data = df_data.reset_index(drop=True)
    cat_col_name = next((col for col in df_data.columns if 'categor' in str(col).lower() or 'kategori' in str(col).lower()), None)

    # Paket (koli) excel'i
    carton_dict = {}
    if carton_bytes:
        df_carton = pd.read_excel(io.BytesIO(carton_bytes))
        def find_col(df, keywords):
            for col in df.columns:
                if all(kw in str(col).lower() for kw in keywords): return col
            return None
        c_code_col = find_col(df_carton, ['code']) or find_col(df_carton, ['sku'])
        c_w_col = find_col(df_carton, ['weight'])
        c_x_col = find_col(df_carton, ['size', '- x'])
        c_y_col = find_col(df_carton, ['size', '- y'])
        c_z_col = find_col(df_carton, ['size', '- z'])
        if c_code_col:
            for _, r in df_carton.iterrows():
                c_sku = str(r[c_code_col]).strip()
                if c_sku and c_sku.lower() != 'nan':
                    def fv(col):
                        try: return float(r[col]) if col and pd.notna(r[col]) else 0
                        except (TypeError, ValueError): return 0
                    carton_dict.setdefault(c_sku, []).append({'kg': fv(c_w_col), 'x': fv(c_x_col), 'y': fv(c_y_col), 'z': fv(c_z_col)})

    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    tpl = parse_template(wb, sheet_name)
    ws = wb[tpl['sheet_name']]
    by_key, by_attr = tpl['by_key'], tpl['by_attr']
    has = lambda name, idx=None: col_key(ATTR[name], idx) in by_key

    missing_core = [n for n in CORE_REQUIRED
                    if not any(k == ATTR[n] or k.startswith(ATTR[n] + '#') for k in by_key)]

    img_cols = find_image_columns(df_data)
    url_ctx = build_url_context(df_data, img_cols)
    bullet_cols = by_attr.get(ATTR['bullet'], [])
    image_cols_tpl = by_attr.get(ATTR['image'], [])
    max_cartons = len(by_attr.get(ATTR['carton_weight'], []))

    processed, skipped, errors, warnings = 0, [], [], []
    written_rows = []
    written_cells = {}            # (row, col) -> sku
    max_idx_used = {}             # attr -> en yüksek kullanılan indeks
    seen_skus = set()
    total_rows = len(df_data)
    max_products = MAX_DATA_ROW - DATA_START_ROW + 1
    g_satir = DATA_START_ROW - 1

    def put(row, key, value, skip_if_default=False):
        col = by_key.get(key)
        if col is None or _is_blank(value):
            return False
        if skip_if_default and col['default'] is not None and _same_value(value, col['default']):
            return False
        ws.cell(row=row, column=col['c']).value = value
        written_cells[(row, col['c'])] = current_sku[0]
        if col['idx']:
            max_idx_used[col['attr']] = max(max_idx_used.get(col['attr'], 0), col['idx'])
        return True

    current_sku = ['']
    for index, row in df_data.iterrows():
        if progress_callback: progress_callback((index + 1) / max(total_rows, 1))
        excel_row = int(row['__excel_row'])
        sku_key = str(row.get('CODE', '')).strip()
        current_sku[0] = sku_key
        desc_short = str(row.get('DESCRIPTION', '') or '')[:60]

        if sku_key in seen_skus:
            skipped.append({'Satır': excel_row, 'Ürün Kodu': sku_key, 'Açıklama': desc_short, 'Sebep': 'Tekrarlanan SKU (Wayfair benzersiz SKU ister)'})
            continue
        if processed >= max_products:
            skipped.append({'Satır': excel_row, 'Ürün Kodu': sku_key, 'Açıklama': desc_short, 'Sebep': f'Template en fazla {max_products} ürün alıyor'})
            continue

        try:
            try: pkg_count = int(float(row.get('NUMBER OF PACKAGES', 1)))
            except (TypeError, ValueError): pkg_count = 1

            def fnum(col):
                v = row.get(col, 0)
                try: return float(v) if pd.notna(v) else 0.0
                except (TypeError, ValueError): return 0.0
            raw_cartons = [{'kg': fnum('WEIGHT (Kg)'), 'x': fnum('PACKAGING SIZE - X (cm)'),
                            'y': fnum('PACKAGING SIZE - Y (cm)'), 'z': fnum('PACKAGING SIZE - Z (cm)')}]
            leave_carton_blank = False
            if sku_key in carton_dict and carton_dict[sku_key]:
                raw_cartons = [dict(c) for c in carton_dict[sku_key]]
            elif pkg_count > 1:
                leave_carton_blank = True
                warnings.append({'Ürün Kodu': sku_key, 'Alan': 'Koli Bilgisi',
                                 'Uyarı': f'{pkg_count} paketli ürün ama Paket Excel\'inde kaydı yok — koli alanları boş bırakıldı'})

            # En büyük koli 1. koli olur
            raw_cartons.sort(key=lambda c: (c['x'] * c['y'] * c['z'], c['kg']), reverse=True)
            kg = raw_cartons[0]['kg']
            if len(raw_cartons) > 1:
                prod_weight_lbs = max(0, round((sum(c['kg'] for c in raw_cartons) * 2.20462) - 5, 2))
                prod_weight_kg = max(0, round(sum(c['kg'] for c in raw_cartons) - 2.0, 2))
            else:
                prod_weight_lbs = max(0, round((kg - 0.1) * 2.20462, 2)) if kg > 0.1 else 0
                prod_weight_kg = max(0, round(kg - 0.05, 2)) if kg > 0.05 else 0

            feat_text = row.get('FEATURES', '')
            raw_h, raw_w, raw_d = extract_overall_dims(feat_text)
            b_info = extract_bedding_info(feat_text, row.get('DESCRIPTION', ''), raw_h, raw_w)

            ean = row.get('EAN CODE', '')
            try: ean_str = "{:.0f}".format(float(ean)) if pd.notna(ean) and str(ean).strip() != '' else ""
            except (TypeError, ValueError): ean_str = str(ean).strip()

            color_val = str(row.get('COLOR', ''))
            if color_val.lower() == 'nan': color_val = ''
            else: color_val = re.sub(r'\s*;\s*', '; ', color_val.replace('\n', ';').replace(',', ';').replace('/', ';')).strip('; ')

            auto_brand = get_brand_by_category(row.get(cat_col_name, '') if cat_col_name else '')

            g_satir += 1
            r_ = g_satir
            put(r_, ATTR['spn'], sku_key)
            put(r_, ATTR['mpn'], sku_key)
            put(r_, ATTR['ean'], ean_str)
            put(r_, ATTR['upc'], ean_str)
            put(r_, ATTR['name'], b_info['new_name'] if b_info['new_name'] else row.get('DESCRIPTION'))
            put(r_, ATTR['base_cost'], row.get('PRICE'))
            put(r_, ATTR['msrp'], row.get('RETAIL PRICE'))
            if tpl['color_key']: put(r_, tpl['color_key'], color_val)
            if auto_brand: put(r_, ATTR['brand'], auto_brand, skip_if_default=True)
            for k in ('moq', 'fqm', 'dsq'): put(r_, ATTR[k], 1, skip_if_default=True)

            conv_dim = (lambda v: cm2in(v) if v else None) if is_us else (lambda v: v)
            put(r_, ATTR['overall_h'], conv_dim(raw_h))
            put(r_, ATTR['overall_w'], conv_dim(raw_w))
            put(r_, ATTR['overall_d'], conv_dim(raw_d))
            for dim_type, keys in ui_data.get('dim_mappings', {}).items():
                val = conv_dim({'h': raw_h, 'w': raw_w, 'd': raw_d}[dim_type])
                for k in keys: put(r_, k, val)

            put(r_, ATTR['bed_set_single'], b_info['set_single'])
            put(r_, ATTR['bed_type'], b_info['prod_type'])
            put(r_, ATTR['bed_size'], b_info['bed_size'])
            put(r_, ATTR['bed_material'], b_info['material'])
            put(r_, ATTR['bed_pieces'], int(b_info['pieces']) if b_info['pieces'] else None)

            # Koliler — artık ana sayfada, 40'a kadar
            if not leave_carton_blank:
                pw = prod_weight_lbs if is_us else prod_weight_kg
                put(r_, ATTR['product_weight'], pw or None)
                put(r_, ATTR['overall_weight'], pw or None)
                if len(raw_cartons) > max_cartons:
                    warnings.append({'Ürün Kodu': sku_key, 'Alan': 'Koli Bilgisi',
                                     'Uyarı': f'{len(raw_cartons)} koli var, template en fazla {max_cartons} alıyor — fazlası yazılmadı'})
                for i, c in enumerate(raw_cartons[:max_cartons], start=1):
                    vals = (kg2lb(c['kg']), cm2in(c['x']), cm2in(c['y']), cm2in(c['z'])) if is_us else (c['kg'], c['x'], c['y'], c['z'])
                    for attr_name, v in zip(('carton_weight', 'carton_height', 'carton_width', 'carton_depth'), vals):
                        put(r_, col_key(ATTR[attr_name], i), v if v else None)

                # Ship Type / Freight Class (sadece template'te varsa — US)
                if is_us and has('ship_type') and all(raw_cartons[0][k] > 0 for k in ('x', 'y', 'z')):
                    total_lbs = sum(c['kg'] for c in raw_cartons) * 2.20462
                    total_vol_in3 = sum(c['x'] * c['y'] * c['z'] for c in raw_cartons) * (0.393701 ** 3)
                    is_ltl = False
                    for c in raw_cartons:
                        dims = sorted([c['x'] * 0.393701, c['y'] * 0.393701, c['z'] * 0.393701], reverse=True)
                        if c['kg'] * 2.20462 >= 150 or (dims[0] + 2 * (dims[1] + dims[2])) >= 165 or dims[0] >= 108:
                            is_ltl = True
                            break
                    if is_ltl:
                        put(r_, ATTR['ship_type'], "LTL", skip_if_default=True)
                        put(r_, ATTR['freight_class'], calculate_freight_class_total(total_lbs, total_vol_in3))
                    else:
                        put(r_, ATTR['ship_type'], "Small Parcel", skip_if_default=True)

            if is_us:
                for k, v in US_LEAD_FALLBACK.items():
                    col = by_key.get(ATTR[k])
                    if col and col['default'] is None:
                        put(r_, ATTR[k], v)

            # Görseller — artık ana sayfada, 30'a kadar
            urls = extract_sku_urls(sku_key, row, img_cols, url_ctx)
            for i, u in enumerate(urls[:len(image_cols_tpl)], start=1):
                put(r_, col_key(ATTR['image'], i), u)
            if len(urls) > len(image_cols_tpl):
                warnings.append({'Ürün Kodu': sku_key, 'Alan': 'Görseller',
                                 'Uyarı': f'{len(urls)} görsel var, template en fazla {len(image_cols_tpl)} alıyor'})

            # Feature Bullet'lar — 8'e kadar
            satirlar = [s.strip() for s in translate_features(feat_text, is_us).split('\n') if s.strip()]
            bedding_note = generate_bedding_note(feat_text, raw_h, raw_w, b_info['bed_size'], is_us)
            all_feats = satirlar + ([bedding_note] if bedding_note else [])
            nb = len(bullet_cols)
            bullets = []
            if nb:
                if len(all_feats) < nb:
                    bullets = all_feats + ["Made In Türkiye"]
                else:
                    bullets = all_feats[:nb - 1] + [" | ".join(all_feats[nb - 1:] + ["Made In Türkiye"])]
                for i, b in enumerate(bullets, start=1):
                    put(r_, col_key(ATTR['bullet'], i), b)
                long_ones = [str(i) for i, b in enumerate(bullets, start=1) if len(b.split()) > 15]
                if long_ones:
                    warnings.append({'Ürün Kodu': sku_key, 'Alan': 'Feature Bullet',
                                     'Uyarı': f"Bullet {', '.join(long_ones)} 15 kelimeyi aşıyor (Wayfair önerisi)"})
                if len(bullets) < 3 and by_key.get(ATTR['marketing']) is not None:
                    warnings.append({'Ürün Kodu': sku_key, 'Alan': 'Feature Bullet',
                                     'Uyarı': f'{len(bullets)} bullet var — Wayfair ya Marketing Copy ya da en az 3 bullet istiyor'})

            # Formdan gelen cevaplar (default ile aynıysa yazılmaz, Wayfair default'u uygular)
            for key, val in ui_data.get('answers', {}).items():
                put(r_, key, val, skip_if_default=True)

            processed += 1
            seen_skus.add(sku_key)
            written_rows.append(r_)

        except Exception as e:
            errors.append({'Satır': excel_row, 'Ürün Kodu': sku_key, 'Açıklama': desc_short, 'Hata Detayı': str(e)})

    # --- Son işlemler ---
    # 1) Geçerli değer kontrolü (dropdown'lu alanlar)
    cols_by_c = {col['c']: col for col in tpl['columns']}
    invalid = {}
    for (r, c), sku in written_cells.items():
        col = cols_by_c[c]
        if not col['allowed']:
            continue
        allowed_low = {a.lower() for a in col['allowed']}
        val = str(ws.cell(r, c).value)
        parts = [p.strip() for p in val.split(';')] if col['type'].lower().startswith('please choose') else [val.strip()]
        bad = [p for p in parts if p and p.lower() not in allowed_low]
        if bad:
            invalid.setdefault((col['name'], '; '.join(bad)), []).append(sku)
    for (name, bad), skus in invalid.items():
        warnings.append({'Ürün Kodu': ', '.join(skus[:5]) + (f' (+{len(skus) - 5})' if len(skus) > 5 else ''),
                         'Alan': name, 'Uyarı': f"Geçerli değer listesinde yok: {bad}"})

    # 2) Boş kalan zorunlu alanlar → sarı (default'u olan alanlar boş sayılmaz)
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    empty_required = {}
    for col in tpl['columns']:
        is_req = col['status'].lower() == 'required' and col['default'] is None
        is_mkt = col['attr'] == ATTR['marketing']
        if not (is_req or is_mkt):
            continue
        for r in written_rows:
            cell = ws.cell(row=r, column=col['c'])
            if not _is_blank(cell.value):
                continue
            if is_mkt:
                n_bul = sum(1 for b in bullet_cols if not _is_blank(ws.cell(r, b['c']).value))
                if n_bul >= 3:
                    continue
            cell.fill = yellow_fill
            empty_required[col['name']] = empty_required.get(col['name'], 0) + 1

    # 3) Yazılan satırların kilidini aç — Excel'de elle düzeltme yapılabilsin
    unlocked = Protection(locked=False)
    for r in written_rows:
        for c in range(1, tpl['max_col'] + 1):
            ws.cell(row=r, column=c).protection = unlocked

    # 4) Veri yazılan gizli grupları aç — "+" butonuna gerek kalmasın
    explode_column_dims(ws)
    for attr, max_i in max_idx_used.items():
        for col in by_attr.get(attr, []):
            if col['idx'] and col['idx'] <= max_i:
                ws.column_dimensions[col['letter']].hidden = False
    # "collapsed" işareti, sağındaki grup kapalı demek; grup açıldıysa işareti kaldır
    for i in range(1, tpl['max_col']):
        d = ws.column_dimensions.get(get_column_letter(i))
        nxt = ws.column_dimensions.get(get_column_letter(i + 1))
        if d is not None and d.collapsed and nxt is not None and not nxt.hidden:
            d.collapsed = False
    # openpyxl kaydederken outlineLevelCol'u siliyor; Excel'deki gruplama çubuğu için geri koy
    max_outline = max(((d.outlineLevel or 0) for d in ws.column_dimensions.values()), default=0)
    if max_outline:
        ws.column_dimensions.max_outline = max_outline

    output = io.BytesIO()
    wb.save(output)
    return {
        'bytes': output.getvalue(), 'processed': processed, 'skipped': skipped, 'errors': errors,
        'warnings': warnings, 'empty_required': empty_required,
        'missing_core': [n for n in missing_core], 'sheet_name': tpl['sheet_name'],
    }


def process_data_excel_only(data_file, is_us):
    data_file.seek(0)
    wb = openpyxl.load_workbook(data_file)
    ws = wb.active

    if hasattr(ws, '_images'): ws._images = []

    def get_headers():
        h = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if v: h[str(v).strip()] = c
        return h

    headers = get_headers()
    code_col = headers.get('CODE')
    if code_col:
        for row in range(ws.max_row, 1, -1):
            val = ws.cell(row=row, column=code_col).value
            if val is None or str(val).strip() == '': ws.delete_rows(row, 1)
                
    headers = get_headers()
    ean_col = headers.get('EAN CODE')
    if ean_col:
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=ean_col)
            if cell.value is not None:
                try: cell.value = int(float(str(cell.value).strip())); cell.number_format = '0'
                except: cell.value = str(cell.value).strip(); cell.number_format = '@'

    color_col = headers.get('COLOR')
    if color_col:
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=color_col)
            if cell.value:
                c_val = str(cell.value)
                if c_val.lower() != 'nan':
                    c_val = c_val.replace('\n', ';').replace(',', ';').replace('/', ';')
                    cell.value = re.sub(r'\s*;\s*', '; ', c_val).strip('; ')

    ef_col = headers.get('EXTRA FEATURES')
    if ef_col:
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=ef_col)
            if cell.value: cell.value = translate_features(str(cell.value), is_us)

    feat_col = headers.get('FEATURES')
    if feat_col:
        ws.insert_cols(feat_col + 1, 5)
        ref_header = ws.cell(row=1, column=feat_col)
        ref_col_letter = get_column_letter(feat_col)
        ref_width = ws.column_dimensions[ref_col_letter].width
        
        for i, f_name in enumerate(['Feature 1', 'Feature 2', 'Feature 3', 'Feature 4', 'Feature 5']):
            target_col_idx = feat_col + 1 + i
            target_col_letter = get_column_letter(target_col_idx)
            nc = ws.cell(row=1, column=target_col_idx)
            nc.value = f_name
            if ref_width: ws.column_dimensions[target_col_letter].width = ref_width
            if ref_header.has_style:
                nc.font = copy.copy(ref_header.font); nc.border = copy.copy(ref_header.border)
                nc.fill = copy.copy(ref_header.fill); nc.alignment = copy.copy(ref_header.alignment)

        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=feat_col)
            features_to_write = ["", "", "", "", ""]
            
            if cell.value:
                translated = translate_features(str(cell.value), is_us)
                lines = [s.strip() for s in translated.split('\n') if s.strip()]
            else:
                lines = []
                
            n = len(lines)
            if n == 0:
                features_to_write[0] = "Made In Türkiye"
            elif n <= 4:
                for idx in range(n): features_to_write[idx] = lines[idx]
                features_to_write[n] = "Made In Türkiye"
            else:
                for idx in range(4): features_to_write[idx] = lines[idx]
                remaining_text = " | ".join(lines[4:])
                features_to_write[4] = f"{remaining_text} | Made In Türkiye"

            ws.cell(row=row, column=feat_col + 1).value = features_to_write[0]
            ws.cell(row=row, column=feat_col + 2).value = features_to_write[1]
            ws.cell(row=row, column=feat_col + 3).value = features_to_write[2]
            ws.cell(row=row, column=feat_col + 4).value = features_to_write[3]
            ws.cell(row=row, column=feat_col + 5).value = features_to_write[4]

    if is_us:
        headers = get_headers()
        w_col, x_col, y_col, z_col = headers.get('WEIGHT (Kg)'), headers.get('PACKAGING SIZE - X (cm)'), headers.get('PACKAGING SIZE - Y (cm)'), headers.get('PACKAGING SIZE - Z (cm)')
        feat_col_for_dims = headers.get('FEATURES')
        metric_cols = [c for c in [w_col, x_col, y_col, z_col] if c is not None]
        if metric_cols:
            insert_idx = max(metric_cols) + 1
            ws.insert_cols(insert_idx, 7)
            ws.cell(row=1, column=insert_idx).value = 'WEIGHT (Lbs)'
            ws.cell(row=1, column=insert_idx + 1).value = 'PACKAGING SIZE - X (in)'
            ws.cell(row=1, column=insert_idx + 2).value = 'PACKAGING SIZE - Y (in)'
            ws.cell(row=1, column=insert_idx + 3).value = 'PACKAGING SIZE - Z (in)'
            ws.cell(row=1, column=insert_idx + 4).value = 'Product Height(in)'
            ws.cell(row=1, column=insert_idx + 5).value = 'Product Width(in)'
            ws.cell(row=1, column=insert_idx + 6).value = 'Product Depth(in)'
            
            ref_h = ws.cell(row=1, column=metric_cols[0])
            for i in range(7):
                nh = ws.cell(row=1, column=insert_idx + i)
                if ref_h.has_style:
                    nh.font = copy.copy(ref_h.font); nh.border = copy.copy(ref_h.border)
                    nh.fill = copy.copy(ref_h.fill); nh.alignment = copy.copy(ref_h.alignment)
            
            actual_feat_col = feat_col_for_dims
            if feat_col_for_dims and feat_col_for_dims >= insert_idx:
                actual_feat_col = feat_col_for_dims + 7
                    
            for row in range(2, ws.max_row + 1):
                if w_col:
                    val = ws.cell(row=row, column=w_col).value
                    try: ws.cell(row=row, column=insert_idx).value = round(float(val) * 2.20462, 2)
                    except: pass
                if x_col:
                    val = ws.cell(row=row, column=x_col).value
                    try: ws.cell(row=row, column=insert_idx + 1).value = round(float(val) * 0.393701, 2)
                    except: pass
                if y_col:
                    val = ws.cell(row=row, column=y_col).value
                    try: ws.cell(row=row, column=insert_idx + 2).value = round(float(val) * 0.393701, 2)
                    except: pass
                if z_col:
                    val = ws.cell(row=row, column=z_col).value
                    try: ws.cell(row=row, column=insert_idx + 3).value = round(float(val) * 0.393701, 2)
                    except: pass
                if actual_feat_col:
                    feat_val = ws.cell(row=row, column=actual_feat_col).value
                    if feat_val:
                        raw_h, raw_w, raw_d = extract_overall_dims(str(feat_val))
                        ws.cell(row=row, column=insert_idx + 4).value = convert_to_inch(raw_h)
                        ws.cell(row=row, column=insert_idx + 5).value = convert_to_inch(raw_w)
                        ws.cell(row=row, column=insert_idx + 6).value = convert_to_inch(raw_d)

    headers = get_headers()
    img_cols = []
    for col_name, col_idx in headers.items():
        name_lower = col_name.lower()
        if 'image' in name_lower or 'resim' in name_lower or 'url' in name_lower or 'link' in name_lower: img_cols.append((col_idx, col_name))
            
    if img_cols:
        img_cols.sort(key=lambda x: x[0], reverse=True)
        extracted_cols = []
        for col_idx, col_name in img_cols:
            col_data = []
            col_letter = get_column_letter(col_idx)
            col_width = ws.column_dimensions[col_letter].width
            for row in range(1, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell_data = {
                    'value': cell.value, 'font': copy.copy(cell.font) if cell.has_style and cell.font else None,
                    'border': copy.copy(cell.border) if cell.has_style and cell.border else None,
                    'fill': copy.copy(cell.fill) if cell.has_style and cell.fill else None,
                    'alignment': copy.copy(cell.alignment) if cell.has_style and cell.alignment else None, 'number_format': cell.number_format
                }
                col_data.append(cell_data)
            extracted_cols.append({'name': col_name, 'width': col_width, 'data': col_data})
            ws.delete_cols(col_idx, 1)
            
        extracted_cols.reverse()
        for col_dict in extracted_cols:
            new_col_idx = ws.max_column + 1
            new_col_letter = get_column_letter(new_col_idx)
            if col_dict['width']: ws.column_dimensions[new_col_letter].width = col_dict['width']
            for row_idx, c_data in enumerate(col_dict['data'], start=1):
                new_cell = ws.cell(row=row_idx, column=new_col_idx)
                new_cell.value = c_data['value']
                if c_data['font']: new_cell.font = c_data['font']
                if c_data['border']: new_cell.border = c_data['border']
                if c_data['fill']: new_cell.fill = c_data['fill']
                if c_data['alignment']: new_cell.alignment = c_data['alignment']
                if c_data['number_format']: new_cell.number_format = c_data['number_format']

    medium_border = Border(left=Side(style='medium', color='000000'), right=Side(style='medium', color='000000'), top=Side(style='medium', color='000000'), bottom=Side(style='medium', color='000000'))
    column_colors = {
        'Feature 1': "DDEBF7", 'Feature 2': "E2EFDA", 'Feature 3': "FFF2CC", 'Feature 4': "FCE4D6", 'Feature 5': "E8D8FC",
        'WEIGHT (Lbs)': "F8CECC", 'PACKAGING SIZE - X (in)': "D1F2EB", 'PACKAGING SIZE - Y (in)': "E8F8F5", 'PACKAGING SIZE - Z (in)': "E6F2F7",
        'Product Height(in)': "D1F2EB", 'Product Width(in)': "E8F8F5", 'Product Depth(in)': "E6F2F7"
    }
    
    headers = get_headers()
    for col_name, col_idx in headers.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 12
        col_color_hex = column_colors.get(col_name)
        col_fill = PatternFill(start_color=col_color_hex, end_color=col_color_hex, fill_type="solid") if col_color_hex else None
            
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.border = medium_border
            if row == 1: cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal='justify', vertical='center', wrap_text=True)
                is_bold = cell.font.bold if (cell.font and cell.font.bold is not None) else False
                is_italic = cell.font.italic if (cell.font and cell.font.italic is not None) else False
                cell.font = Font(name='Tahoma', size=8, bold=is_bold, italic=is_italic)
                if col_fill: cell.fill = col_fill

    for row in range(1, ws.max_row + 1): ws.row_dimensions[row].height = 18

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 7. ARAYÜZ ---

# Template'te default değeri yoksa kullanılacak öneriler (alan adına göre)
FALLBACK_DEFAULTS = [
    ('warning required', ['No']),
    ('country of manufacturer', ['Turkey']), ('country of origin', ['Turkey']),
    ('uniform packaging and labeling regulations', ['Yes']),
    ('reason for restriction', ['Does Not Apply']),
    ('general certificate of conformity', ['Yes']),
    ('canada product restriction', ['No']),
    ('soffa compliant', ['Does Not Apply']), ('canfer compliant', ['Does Not Apply']),
    ('carb phase', ['Does Not Apply']), ('cal tb 117-2013', ['Does Not Apply']),
    ('california ab-1817', ['Does Not Apply']), ('sor/2016', ['Does Not Apply']),
    ('astm', ['Does Not Apply']), ('lacey act', ['Does Not Apply']), ('safety list', ['Does Not Apply']),
    ('cpsc - 16 cfr', ['Yes']), ('composite wood product (cwp)', ['Does Not Apply']),
    ('tsca title vi compliant', ['Does Not Apply']),
    ('supplier intended and approved use', ['Non Residential Use', 'Residential Use']),
    ('commercial warranty', ['Yes']), ('contains flame retardant', ['No']),
    ('wayfair compliance verified', ['No']), ('wayfair compliance vetting', ['No']),
    ('battery or batteries included', ['No']), ('additional intended use for child', ['No']),
]

@st.cache_data(show_spinner=False, max_entries=8)
def load_template_info(t_bytes, sheet_name=None):
    wb = openpyxl.load_workbook(io.BytesIO(t_bytes))
    tpl = parse_template(wb, sheet_name)
    try:
        df_v = pd.read_excel(io.BytesIO(t_bytes), sheet_name='Valid Values')
        vv_cols = {str(c).strip(): [str(o).strip() for o in df_v[c].dropna().unique() if str(o).strip()] for c in df_v.columns}
    except Exception:
        vv_cols = {}
    tpl['vv_cols'] = vv_cols
    return tpl

def field_options(col, tpl):
    opts = col['allowed'] or tpl['vv_cols'].get(col['name'])
    if not opts and col['type'].lower() == 'yes or no':
        opts = ["Yes", "No", "Does Not Apply"]
    if not opts:
        return None
    return list(dict.fromkeys(o for o in opts if o and o != 'None' and 'select all' not in o.lower()))

def initial_value(col, opts, is_multi):
    """Öncelik: bu oturumdaki son seçim > template Default Row > yedek öneriler."""
    prefs = st.session_state['user_prefs']
    if col['key'] in prefs:
        return prefs[col['key']]
    if col['default'] is not None:
        return [p.strip() for p in col['default'].split(';')] if is_multi else [col['default']]
    f_low = col['name'].lower()
    for kw, val in FALLBACK_DEFAULTS:
        if kw in f_low:
            if 'country of' in kw and opts and 'Turkey' not in opts and 'Türkiye' in opts:
                return ['Türkiye']
            return val
    return []

def render_field(col, tpl):
    """Bir template alanı için uygun widget'ı çizer, yazılacak değeri (str) döner."""
    opts = field_options(col, tpl)
    type_low = col['type'].lower()
    is_multi = type_low.startswith('please choose')
    wkey = f"ans_{tpl['class_id']}_{col['key']}"
    label = col['name'] + (f"  ·  default: {col['default']}" if col['default'] else "")
    help_txt = col['desc'][:400] if col['desc'] else None
    init = initial_value(col, opts, is_multi)

    if opts is None:
        val = st.text_input(label, value=init[0] if init else "", key=wkey, help=help_txt)
        st.session_state['user_prefs'][col['key']] = [val] if val else []
        return val.strip()

    if is_multi:
        sel = st.multiselect(label, options=opts + [CUSTOM_OPT], default=[x for x in init if x in opts],
                             key=wkey, help=help_txt)
        st.session_state['user_prefs'][col['key']] = [x for x in sel if x != CUSTOM_OPT]
        final = [x for x in sel if x != CUSTOM_OPT]
        if CUSTOM_OPT in sel:
            custom = st.text_input(f"✍️ {col['name']} için Custom Değer", key=f"custom_{wkey}")
            if custom: final.append(custom.strip())
        return "; ".join(final)

    choices = [""] + opts + [CUSTOM_OPT]
    idx = choices.index(init[0]) if init and init[0] in choices else 0
    sel = st.selectbox(label, options=choices, index=idx, key=wkey, help=help_txt)
    st.session_state['user_prefs'][col['key']] = [sel] if sel and sel != CUSTOM_OPT else []
    if sel == CUSTOM_OPT:
        return st.text_input(f"✍️ {col['name']} için Custom Değer", key=f"custom_{wkey}").strip()
    return sel


@st.dialog("🌎 Bölge Seçimi")
def region_dialog(tpl_region, tpl_country, tpl_units):
    st.write("Bu dosya hangi bölge için hazırlanıyor?")
    options = ["US (İnç / Lbs)", "EU (cm / Kg)"]
    pre = {'US': 0, 'EU': 1}.get(tpl_region)
    choice = st.radio("Bölge", options, index=pre, key="dlg_region", label_visibility="collapsed")
    if tpl_region:
        unit_txt = {'imperial': 'inç / lbs', 'metric': 'cm / kg'}.get(tpl_units, 'bilinmiyor')
        st.caption(f"Template'e göre: **{tpl_region}** (pazar: {tpl_country or '?'}, birimler: {unit_txt})")
    confirm = True
    if choice and tpl_region and not choice.startswith(tpl_region):
        st.error(f"Seçiminiz template ile uyuşmuyor. Template **{tpl_region}** birimleri bekliyor; "
                 f"devam ederseniz ölçü ve ağırlıklar yanlış birimde yazılır.")
        confirm = st.checkbox("Yine de bu bölgeyle devam et")
    if st.button("✅ Onayla ve Hazırla", type="primary", disabled=(choice is None or not confirm), width='stretch'):
        st.session_state['wf_run'] = {'is_us': choice.startswith("US")}
        st.rerun()


st.set_page_config(page_title="Wayfair & Data Akıllı Ürün Robotu V20", layout="wide")
st.title("🛡️ Wayfair & Data Akıllı Ürün Robotu V20")

with st.sidebar:
    st.header("⚙️ Genel Ayarlar")
    teklif_no = st.text_input("📄 Teklif No *", key="teklif_no", placeholder="örn. 219383").strip()
    teklif_ok = bool(re.fullmatch(r'\d+', teklif_no))
    if not teklif_no:
        st.caption("Zorunlu — çıktı dosyasının adı buna göre oluşur (örn. 219383-US-WF).")
    elif not teklif_ok:
        st.error("Teklif No sadece rakamlardan oluşmalı.")
    st.divider()
    st.subheader("🔗 Hızlı Bağlantılar")
    st.markdown("[🛠️ Asir Tools](https://excelwebpy-asirtools.streamlit.app/)")
    st.divider()

tab_wayfair, tab_data = st.tabs(["🎯 Wayfair Şablonu Hazırla", "🛠️ Sadece Data Excel'i Çevir"])

with tab_wayfair:
    st.info("💡 Bu alan, Data Excel'inizdeki ürünleri Wayfair şablonuna otomatik eşleştirip yazar.")

    u1, u2, u3 = st.columns(3)
    with u1: d_file = st.file_uploader("1. Data Excel", type="xlsx", key="wayfair_data")
    with u2: t_file = st.file_uploader("2. Template Excel", type="xlsx", key="wayfair_template")
    with u3: c_file = st.file_uploader("3. Paket Excel (Opsiyon)", type="xlsx", key="wayfair_carton")

    if t_file:
        t_bytes = t_file.getvalue()
        try:
            tpl = load_template_info(t_bytes)
        except TemplateFormatError as e:
            st.error(f"❌ {e}")
            st.stop()

        if len(tpl['sheets']) > 1:
            chosen = st.selectbox("Bu template'te birden fazla kategori sayfası var — hangisi doldurulsun?", tpl['sheets'])
            if chosen != tpl['sheet_name']:
                tpl = load_template_info(t_bytes, chosen)

        unit_txt = {'imperial': 'inç / lbs', 'metric': 'cm / kg'}.get(tpl['units'], 'tespit edilemedi')
        st.caption(f"📄 **{tpl['sheet_name']}**  ·  Pazar: **{tpl['country'] or '?'}** ({tpl['locale']})  ·  Birimler: **{unit_txt}**")

        a_keys = auto_keys(tpl)
        form_cols = [c for c in tpl['columns']
                     if c['status'].lower() == 'required' and c['key'] not in a_keys and c['attr'] not in MEDIA_IDS]

        st.markdown("---")
        st.subheader("📐 Özel Ölçü Sütun Eşleştirmeleri")
        dim_candidates = [c for c in tpl['columns'] if c['key'] not in a_keys and c['idx'] is None
                          and c['type'].lower() == 'number' and c['status'].lower() in ('required', 'conditional')]
        options_dict = {f"{c['name']} ({c['letter']})": c['key'] for c in dim_candidates}
        options_list = list(options_dict.keys())
        col_h, col_w, col_d = st.columns(3)
        with col_h: h_sel = st.multiselect("Height (Yükseklik) Yazılacaklar", options=options_list)
        with col_w: w_sel = st.multiselect("Genişlik (Width) Yazılacaklar", options=options_list)
        with col_d: d_sel = st.multiselect("Depth (Derinlik) Yazılacaklar", options=options_list)
        dim_mappings = {'h': [options_dict[x] for x in h_sel], 'w': [options_dict[x] for x in w_sel], 'd': [options_dict[x] for x in d_sel]}
        selected_dim_keys = set(sum(dim_mappings.values(), []))

        form_cols = [c for c in form_cols if c['key'] not in selected_dim_keys]
        need_cols = [c for c in form_cols if c['default'] is None]
        default_cols = [c for c in form_cols if c['default'] is not None]

        st.markdown("---")
        st.subheader(f"📋 {tpl['sheet_name']} — Doldurulması Gereken Özellikler")
        answers = {}
        if need_cols:
            st.caption(f"Template'te varsayılanı olmayan {len(need_cols)} zorunlu alan. Boş bırakılanlar çıktıda sarı işaretlenir.")
            grid = st.columns(3)
            for i, col in enumerate(need_cols):
                with grid[i % 3]:
                    answers[col['key']] = render_field(col, tpl)
        else:
            st.success("Tüm zorunlu alanların template'te varsayılanı var.")

        if default_cols:
            with st.expander(f"⚙️ Varsayılanı olan alanlar ({len(default_cols)}) — sadece farklı bir değer gerekiyorsa değiştirin"):
                st.caption("Wayfair, Default Row'daki (8. satır) değerleri tüm ürünlere otomatik uygular. "
                           "Burada değiştirdiğiniz değer sadece bu dosyadaki ürün satırlarına yazılır, kategori varsayılanı değişmez.")
                grid2 = st.columns(3)
                for i, col in enumerate(default_cols):
                    with grid2[i % 3]:
                        answers[col['key']] = render_field(col, tpl)

        st.markdown("<br>", unsafe_allow_html=True)
        ready = bool(d_file) and teklif_ok
        if not teklif_ok:
            st.warning("⬅️ Devam etmek için sol menüden **Teklif No** girin.")
        elif not d_file:
            st.warning("Data Excel'i yükleyin.")
        if st.button("🚀 Wayfair Dosyasını Hazırla", type="primary", width='stretch', disabled=not ready):
            st.session_state.pop('wf_result', None)
            region_dialog(tpl['region'], tpl['country'], tpl['units'])

        run_req = st.session_state.pop('wf_run', None)
        if run_req and ready:
            ui_data = {'is_us': run_req['is_us'], 'answers': {k: v for k, v in answers.items() if v},
                       'dim_mappings': dim_mappings}
            progress_bar = st.progress(0, text="Hazırlanıyor...")
            def update_progress(val): progress_bar.progress(min(val, 1.0), text=f"İşleniyor... %{int(val * 100)}")
            with st.spinner("Excel dosyası işleniyor..."):
                res = process_wayfair_v20(d_file.getvalue(), t_bytes, ui_data,
                                          carton_bytes=c_file.getvalue() if c_file else None,
                                          sheet_name=tpl['sheet_name'], progress_callback=update_progress)
            progress_bar.progress(1.0, text="✅ Tamamlandı!")
            res['file_name'] = f"{teklif_no}-{'US' if run_req['is_us'] else 'EU'}-WF.xlsx"
            res['sig'] = (d_file.file_id, t_file.file_id, c_file.file_id if c_file else None)
            st.session_state['wf_result'] = res

        res = st.session_state.get('wf_result')
        if res and res.get('sig') == (d_file.file_id if d_file else None, t_file.file_id, c_file.file_id if c_file else None):
            st.markdown("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("✅ İşlenen", res['processed']); m2.metric("⏭️ Atlanan", len(res['skipped']))
            m3.metric("❌ Hatalı", len(res['errors'])); m4.metric("⚠️ Uyarı", len(res['warnings']))

            if res['missing_core']:
                st.warning("Template'te beklenen temel sütunlar bulunamadı: " + ", ".join(res['missing_core']))
            if res['empty_required']:
                with st.expander(f"🟨 Boş kalan zorunlu alanlar — {len(res['empty_required'])} sütun (çıktıda sarı)", expanded=True):
                    st.dataframe(pd.DataFrame([{'Alan': k, 'Boş Ürün Sayısı': v} for k, v in res['empty_required'].items()]), width='stretch')
            if res['warnings']:
                with st.expander(f"⚠️ Uyarılar — {len(res['warnings'])}", expanded=False):
                    st.dataframe(pd.DataFrame(res['warnings']), width='stretch')
            if res['skipped']:
                with st.expander(f"⏭️ Atlanan Satırlar — {len(res['skipped'])} ürün", expanded=False):
                    st.dataframe(pd.DataFrame(res['skipped']), width='stretch')
            if res['errors']:
                with st.expander(f"❌ Hatalı Satırlar — {len(res['errors'])} ürün", expanded=True):
                    st.dataframe(pd.DataFrame(res['errors']), width='stretch')

            if res['processed'] > 0:
                st.success(f"✅ {res['processed']} ürün başarıyla işlendi → **{res['file_name']}**")
                st.download_button(label=f"📥 {res['file_name']} İndir", data=res['bytes'], file_name=res['file_name'],
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width='stretch')

with tab_data:
    st.subheader("🛠️ Data Excel Dönüştürücü (Şablonsuz)")
    st.info("💡 Bu araç sadece veri excelinizi okur, **özellikleri (Features) 5 sütuna böler**, ölçüleri (cm/kg/ml) Amerikan standartlarına (inç/lbs/fl oz) çevirir ve koli ebatları için yeni sütunlar açıp size temiz bir Excel sunar.")

    data_region = st.radio("🌎 Bölge", ["US (İnç / Lbs)", "EU (cm / Kg)"], horizontal=True, key="data_region")
    data_is_us = data_region.startswith("US")
    data_only_file = st.file_uploader("İşlenecek Data Excel'ini Yükleyin", type="xlsx", key="data_only_upload")
    if data_only_file:
        if st.button("🚀 Data Excel'i Dönüştür ve İndir", type="primary", width='stretch'):
            with st.spinner("Data Excel'iniz çevriliyor ve bölünüyor..."):
                result_excel = process_data_excel_only(io.BytesIO(data_only_file.getvalue()), data_is_us)
                st.success("✅ Dönüştürme Başarılı! Aşağıdaki butona tıklayarak yeni excelinizi indirebilirsiniz.")
                st.download_button(label="📥 Dönüştürülmüş Data Excel'i İndir", data=result_excel, file_name="Converted_Data_Excel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
