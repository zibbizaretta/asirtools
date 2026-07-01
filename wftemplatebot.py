import streamlit as st
import pandas as pd
import openpyxl
import re
import io

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
    """
    Geliştirilmiş boyut çıkarımı:
    - W:/H:/D: kısa formatları destekler
    - × karakterini de tanır
    - Çap (Diameter) tespiti
    """
    if pd.isna(text):
        return None, None, None
    text = str(text)

    # Çap (Diameter) tespiti
    dia = get_dim_val(r'(?:Diameter|Çap|Dia|Ø)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)

    # \b eklendi ki "Width" içindeki 'h' Height olarak algılanmasın
    w = get_dim_val(r'(?:Width|Genişlik|Side to Side|\bW\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)
    h = get_dim_val(r'(?:Height|Yükseklik|Top to Bottom|\bH\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)
    d = get_dim_val(r'(?:Depth|Derinlik|Front to Back|\bD\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)

    # Çap varsa ve Genişlik/Derinlik belirtilmemişse, çapı buralara ata
    if dia:
        if not w: 
            w = dia
        if not d: 
            d = dia

    # Eğer standart boyutlardan hiçbiri bulunamadıysa Size/Ölçü formatına bak
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
    """
    Geliştirilmiş çeviri:
    - Boyut sayısını inçe, kg'ı lbs'ye, ml'yi fl oz'a çevirir.
    - Wayfair'in kabul etmediği Ø veya Çap ifadelerini Diameter'a çevirir.
    """
    if pd.isna(text):
        return ""
    text = str(text)

    # Ø işaretini ve "Çap" kelimesini "Diameter: " olarak düzelt
    text = re.sub(r'Ø\s*[:\-]?\s*', 'Diameter: ', text)
    text = re.sub(r'(?i)\bÇap\s*[:\-]?\s*', 'Diameter: ', text)

    if not do_conversion:
        return text

    # Dönüşüm fonksiyonları
    def c_in(m): 
        return f"{round(float(m.group(1).replace(',', '.')) * 0.393701, 2)}"
        
    def c_mm(m): 
        return f"{round(float(m.group(1).replace(',', '.')) * 0.0393701, 2)}"
        
    def c_kg(m): 
        return f"{round(float(m.group(1).replace(',', '.')) * 2.20462, 2)}"
        
    def c_ml(m): 
        return f"{round(float(m.group(1).replace(',', '.')) * 0.033814, 2)}"

    # Sayı + birim regex değişimleri
    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*cm(?![a-zA-Z])', lambda m: c_in(m) + ' inches', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*mm(?![a-zA-Z])', lambda m: c_mm(m) + ' inches', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*kg(?![a-zA-Z])', lambda m: c_kg(m) + ' lbs', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z0-9])(\d+(?:[\.,]\d+)?)\s*ml(?![a-zA-Z])', lambda m: c_ml(m) + ' fl oz', text, flags=re.IGNORECASE)
    
    # Boyut çarpımı formatları: 50x100x30
    text = re.sub(r'(?<![a-zA-Z0-9:])(\d+(?:[\.,]\d+)?)(?=\s*[xX×]\s*\d)', c_in, text)
    
    # Kalan yalnız birim etiketlerini temizle
    text = re.sub(r'(?<![a-zA-Z])\bcm\b(?![a-zA-Z])', 'inches', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z])\bmm\b(?![a-zA-Z])', 'inches', text, flags=re.IGNORECASE)
    
    return text

def calculate_freight_class(weight_lbs, l_in, w_in, h_in):
    vol = (l_in * w_in * h_in) / 1728
    if vol == 0:
        return "60"
    dens = weight_lbs / vol
    if dens < 1: 
        return "400"
    elif dens < 2: 
        return "300"
    elif dens < 4: 
        return "200"
    elif dens < 6: 
        return "150"
    elif dens < 8: 
        return "125"
    elif dens < 10: 
        return "100"
    else: 
        return "60"

def get_ship_type(weight_lbs, l_in, w_in, h_in):
    dims = sorted([l_in, w_in, h_in], reverse=True)
    length = dims[0]
    girth = 2 * (dims[1] + dims[2])
    
    if weight_lbs < 150 and (length + girth) < 165 and length < 108:
        return "Small Parcel", None
    return "LTL", calculate_freight_class(weight_lbs, l_in, w_in, h_in)

def validate_column_mappings(col_map, mappings):
    """Template'de bulunmayan mapping sütunlarını döndürür."""
    return [k for k in mappings if k not in col_map]

# --- YENİ: KAPSAMLI BEDDING BİLGİSİ ÇIKARICI ---
def extract_bedding_info(features, description, raw_h, raw_w):
    """
    Parça sayısı, Set/Single durumu, Materyal, Yatak Tipi, Yatak Ölçüsü ve 
    Amerika standartlarına uyarlanmış yeni Ürün Başlığını hesaplar.
    """
    text_lower = (str(features) + " " + str(description)).lower()
    
    # 1. Parça Sayısı (Pieces Included) Hesaplama
    total_pieces = 0
    
    # Yastık Kılıflarını (Pillowcase) sayma (Örn: "2 pieces", "2 adet")
    pillow_match = re.search(r'(?:pillowcase|pillow case|yastık kılıfı).*?(\d+)\s*(?:piece|adet|pcs)', text_lower)
    if pillow_match:
        total_pieces += int(pillow_match.group(1))
    elif 'pillowcase' in text_lower or 'yastık kılıfı' in text_lower:
        total_pieces += 1
        
    # Diğer bileşenleri sayma
    if any(x in text_lower for x in ['duvet cover', 'quilt cover', 'nevresim']): total_pieces += 1
    if any(x in text_lower for x in ['fitted sheet', 'çarşaf']): total_pieces += 1
    if any(x in text_lower for x in ['flat sheet', 'düz çarşaf']): total_pieces += 1
    if any(x in text_lower for x in ['bedspread', 'yatak örtüsü']): total_pieces += 1
    if any(x in text_lower for x in ['comforter', 'yorgan']): total_pieces += 1
    if any(x in text_lower for x in ['blanket', 'battaniye']): total_pieces += 1
    
    pieces_str = str(total_pieces) if total_pieces > 0 else ""
    
    # 2. Set / Single Belirleme
    set_single = ""
    if total_pieces > 1:
        set_single = "Set (matching pieces included)"
    elif total_pieces == 1:
        set_single = "Single Piece"
        
    # 3. Bedding Material Belirleme
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
    
    # 4. Bedding Product Type
    prod_type = ""
    is_bedding = False
    if 'duvet cover' in text_lower or 'nevresim' in text_lower or 'quilt cover' in text_lower: 
        prod_type = "Duvet Cover"
        is_bedding = True
    elif 'bedspread' in text_lower or 'yatak örtüsü' in text_lower: 
        prod_type = "Bedspread"
        is_bedding = True
    elif 'quilt' in text_lower: 
        prod_type = "Quilt"
        is_bedding = True
    elif 'comforter' in text_lower or 'yorgan' in text_lower: 
        prod_type = "Comforter"
        is_bedding = True
    elif 'coverlet' in text_lower:
        prod_type = "Coverlet"
        is_bedding = True
    elif 'pillowcase' in text_lower or 'yastık kılıfı' in text_lower or 'sham' in text_lower: 
        prod_type = "Sham"
        
    # 5. Bedding Size
    bed_size = ""
    max_dim = max(float(raw_h or 0), float(raw_w or 0))
    if max_dim > 0 and (is_bedding or prod_type == "Sham"):
        if max_dim >= 250: bed_size = "California King"
        elif max_dim >= 230: bed_size = "King"
        elif max_dim >= 200: bed_size = "Queen"
        elif max_dim >= 180: bed_size = "Full / Double"
        else: bed_size = "Twin"
        
    # 6. Akıllı Product Name Güncellemesi
    new_name = str(description)
    # Ülke/Dil kodlarını sil (Örn: (DE), (FR), (UK) vb.)
    new_name = re.sub(r'\s*\([A-Z]{2,3}\)\s*', ' ', new_name)
    # Eski beden kelimelerini sil
    size_words = [r'\bSingle XXL\b', r'\bSingle XL\b', r'\bSingle\b', r'\bDouble\b', r'\bKing\b', r'\bSuper King\b', r'\bMega King\b', r'\bQueen\b', r'\bTwin\b', r'\bFull\b']
    for word in size_words:
        new_name = re.sub(word, '', new_name, flags=re.IGNORECASE)
    
    # Fazla boşlukları temizle
    new_name = re.sub(r'\s+', ' ', new_name).strip()
    
    # Sadece nevresim ürünleri için Amerikan ölçüsünü başa ekle (Sadece yastık kılıfı ise ekleme)
    if bed_size and prod_type != "Sham" and is_bedding:
        new_name = f"{bed_size} {new_name}"
        
    return {
        'pieces': pieces_str,
        'set_single': set_single,
        'material': material,
        'prod_type': prod_type,
        'bed_size': bed_size,
        'new_name': new_name
    }


# --- OTOMATİK NEVRESİM/YATAK BOYUTU HESAPLAYICI VE NOT OLUŞTURUCU ---
def generate_bedding_note(text, h_val, w_val, bed_size, is_us):
    """
    Özellikler metnini ve cm cinsinden ölçüleri alarak, 
    Amerikan pazarındaki yatak ölçü standartlarına (Twin, King vb.) uygun
    sıcak bir pazarlama / adaptasyon notu oluşturur.
    """
    if not is_us or not text:
        return ""

    text_lower = str(text).lower()
    is_bedding = any(k in text_lower for k in ['duvet', 'quilt', 'bedspread', 'nevresim', 'yorgan', 'yatak örtüsü', 'comforter', 'coverlet'])
    is_pillow = any(k in text_lower for k in ['pillowcase', 'yastık kılıfı', 'sham'])

    if not (is_bedding or is_pillow):
        return ""

    h_in = round(float(h_val) * 0.393701, 2) if h_val else ""
    w_in = round(float(w_val) * 0.393701, 2) if w_val else ""
    dims_str = f"{h_in} x {w_in}" if h_in and w_in else f"{h_in or w_in}"
    
    if is_bedding and bed_size:
        return f"This beautiful set is crafted to European standards. Measuring {dims_str} inches, it beautifully complements your {bed_size} bed, offering a cozy and elegant drape."
    elif is_pillow:
        return f"Crafted in Türkiye, these pillowcases measure {dims_str} inches. They are a wonderful fit for standard US pillows, bringing a touch of European comfort to your bedroom."

    return ""


# --- 3. ANA İŞLEME MOTORU ---

def process_wayfair_v13(data_file, template_file, ui_data, carton_file=None, progress_callback=None):
    data_file.seek(0)
    template_file.seek(0)

    df_data = pd.read_excel(data_file)
    
    # --- HAYALET VE ÖZET SATIR KORUMASI ---
    if 'CODE' in df_data.columns:
        df_data = df_data.dropna(subset=['CODE'])
        df_data = df_data[df_data['CODE'].astype(str).str.strip() != '']
    else:
        df_data = df_data.dropna(how='all')
        
    df_data = df_data.reset_index(drop=True)
    
    # --- PAKET (CARTON) EXCEL'İNİ İŞLEME ---
    carton_dict = {}
    if carton_file is not None:
        carton_file.seek(0)
        df_carton = pd.read_excel(carton_file)
        
        def find_col(df, keywords):
            for col in df.columns:
                if all(kw in str(col).lower() for kw in keywords): 
                    return col
            return None
            
        c_code_col = find_col(df_carton, ['code']) or find_col(df_carton, ['sku'])
        c_w_col = find_col(df_carton, ['weight'])
        c_x_col = find_col(df_carton, ['size', 'x'])
        c_y_col = find_col(df_carton, ['size', 'y'])
        c_z_col = find_col(df_carton, ['size', 'z'])
        
        if c_code_col:
            for _, r in df_carton.iterrows():
                c_sku = str(r[c_code_col]).strip()
                if c_sku and c_sku.lower() != 'nan':
                    if c_sku not in carton_dict: 
                        carton_dict[c_sku] = []
                        
                    try:
                        w_val = float(r[c_w_col]) if c_w_col and pd.notna(r[c_w_col]) else 0
                        x_val = float(r[c_x_col]) if c_x_col and pd.notna(r[c_x_col]) else 0
                        y_val = float(r[c_y_col]) if c_y_col and pd.notna(r[c_y_col]) else 0
                        z_val = float(r[c_z_col]) if c_z_col and pd.notna(r[c_z_col]) else 0
                    except:
                        w_val, x_val, y_val, z_val = 0, 0, 0, 0
                        
                    carton_dict[c_sku].append({
                        'kg': w_val, 
                        'x': x_val, 
                        'y': y_val, 
                        'z': z_val
                    })

    wb = openpyxl.load_workbook(template_file)
    
    # Ana şablonu bul
    target_sheet = None
    for s in wb.sheetnames:
        if not any(x in s for x in ["Additional", "WAYFAIR", "Instructions", "Valid Values", "Failed"]):
            target_sheet = s
            break
            
    if target_sheet is None:
        target_sheet = wb.sheetnames[0]
        
    ws_main = wb[target_sheet]

    # Akıllı Sütun Eşleştirme (Smart Column Mapping)
    col_map = {}
    for c in range(1, ws_main.max_column + 1):
        r1_val = str(ws_main.cell(row=1, column=c).value).strip() if ws_main.cell(row=1, column=c).value else ""
        r4_val = str(ws_main.cell(row=4, column=c).value).strip() if ws_main.cell(row=4, column=c).value else ""
        col_let = ws_main.cell(row=1, column=c).column_letter
        
        if r1_val: 
            col_map[r1_val] = col_let
            
        r4_lower = r4_val.lower()
        r1_lower = r1_val.lower()

        # COLOR eşleştirmesi
        if ('color' in r4_lower or 'colour' in r4_lower or r1_lower.endswith('::color')):
            if 'leg' not in r4_lower and 'base' not in r4_lower and 'shade' not in r4_lower:
                col_map['featureDescription::color'] = col_let

        # Boyut eşleştirmesi
        if 'overall height' in r4_lower or 'overallheight' in r1_lower:
            col_map['featureDescription::overallHeight'] = col_let
        elif 'overall width' in r4_lower or 'overallwidth' in r1_lower:
            col_map['featureDescription::overallWidth'] = col_let
        elif 'overall depth' in r4_lower or 'overalldepth' in r1_lower:
            col_map['featureDescription::overallDepth'] = col_let
            
        # Akıllı Bedding Sütunları Eşleştirmesi
        if 'set / single' in r4_lower: col_map['bedding::setSingle'] = col_let
        if 'bedding product type' in r4_lower: col_map['bedding::productType'] = col_let
        if 'bedding size' in r4_lower: col_map['bedding::size'] = col_let
        if 'bedding material' in r4_lower: col_map['bedding::material'] = col_let
        if 'pieces included' in r4_lower or 'total number of pieces included' in r4_lower: col_map['bedding::pieces'] = col_let

        # Resim eşleştirmesi
        for i in range(1, 6):
            if f'image file name or url {i}' in r4_lower:
                col_map[f'img_{i}'] = col_let

    feature_cols = [c.column_letter for c in ws_main[1] if str(c.value).strip() == 'featureDescription::genericFeatures']

    total_rows = len(df_data)
    processed = 0
    skipped = []
    errors = []
    missing_cols_reported = False
    written_rows = []
    additional_images_data = []
    additional_cartons_data = []
    processed_skus_for_additional = set()
    processed_skus_for_cartons = set()

    for index, row in df_data.iterrows():
        g_satir = 8 + index
        
        if progress_callback: 
            progress_callback((index + 1) / total_rows)

        sku_key = str(row.get('CODE', '')).strip()
        
        try: 
            pkg_count = int(float(row.get('NUMBER OF PACKAGES', 1)))
        except: 
            pkg_count = 1

        kg = float(row.get('WEIGHT (Kg)', 0) or 0)
        x_cm = float(row.get('PACKAGING SIZE - X (cm)', 0) or 0)
        y_cm = float(row.get('PACKAGING SIZE - Y (cm)', 0) or 0)
        z_cm = float(row.get('PACKAGING SIZE - Z (cm)', 0) or 0)
        
        prod_weight_lbs = round((kg - 0.1) * 2.20462, 2) if kg > 0.1 else 0

        # --- PAKET EXCELİ ÖNCELİĞİ ---
        if carton_file is not None and sku_key in carton_dict and len(carton_dict[sku_key]) > 0:
            cartons = carton_dict[sku_key]
            # İlk paket
            kg = cartons[0]['kg']
            x_cm = cartons[0]['x']
            y_cm = cartons[0]['y']
            z_cm = cartons[0]['z']
            
            # Tüm kolilerin toplamından 5 lbs çıkar
            total_kg = sum(c['kg'] for c in cartons)
            total_lbs = total_kg * 2.20462
            prod_weight_lbs = max(0, round(total_lbs - 5, 2))
            
            # Additional Cartons listesi
            if len(cartons) > 1 and sku_key not in processed_skus_for_cartons:
                for ext_c in cartons[1:]:
                    additional_cartons_data.append({
                        'sku': sku_key, 
                        'kg': ext_c['kg'], 
                        'x': ext_c['x'], 
                        'y': ext_c['y'], 
                        'z': ext_c['z']
                    })
                processed_skus_for_cartons.add(sku_key)
                
        elif pkg_count != 1:
            skipped.append({
                'Satır': index + 2, 
                'Ürün Kodu': sku_key, 
                'Açıklama': str(row.get('DESCRIPTION', '') or '')[:60], 
                'Neden': f"Paket Sayısı = {pkg_count} ama paket excelinde verisi yok (Atlandı)"
            })
            continue

        try:
            feat_text = row.get('FEATURES', '')
            raw_h, raw_w, raw_d = extract_overall_dims(feat_text)
            
            # --- YENİ: Bedding Verilerini Çıkar ---
            b_info = extract_bedding_info(feat_text, row.get('DESCRIPTION', ''), raw_h, raw_w)
            
            lbs = round(kg * 2.20462, 2)
            x_in = round(x_cm * 0.393701, 2)
            y_in = round(y_cm * 0.393701, 2)
            z_in = round(z_cm * 0.393701, 2)
            
            ean = row.get('EAN CODE', '')
            ean_str = "{:.0f}".format(float(ean)) if pd.notna(ean) and str(ean).strip() != '' else ""

            color_val = str(row.get('COLOR', ''))
            if color_val.lower() == 'nan':
                color_val = ''
            else:
                color_val = color_val.replace('\n', ';').replace(',', ';').replace('/', ';')
                color_val = re.sub(r'\s*;\s*', '; ', color_val).strip('; ')

            mappings = {
                'core::supplierPartNumber': sku_key,
                'core::manufacturerPartNumber': sku_key,
                'core::universalProductCode': ean_str,
                # Yeni oluşturulan ürün ismini kullan (varsa)
                'core::productName': b_info['new_name'] if b_info['new_name'] else row.get('DESCRIPTION'),
                'price::wholesalePrice': row.get('PRICE'),
                'price::manufacturerSuggestedRetailPrice': row.get('RETAIL PRICE'),
                'shippingAndFulfillment::weight': lbs,
                'shippingAndFulfillment::height': x_in,
                'shippingAndFulfillment::width': y_in,
                'shippingAndFulfillment::depth': z_in,
                'shippingAndFulfillment::productWeight': prod_weight_lbs,
                'featureDescription::overallHeight': convert_to_inch(raw_h) if ui_data['is_us'] else raw_h,
                'featureDescription::overallWidth': convert_to_inch(raw_w) if ui_data['is_us'] else raw_w,
                'featureDescription::overallDepth': convert_to_inch(raw_d) if ui_data['is_us'] else raw_d,
                'featureDescription::color': color_val,
                'core::collectionName': ui_data['coll_name'],
                'core::manufacturerId': ui_data['brand'],
                'shippingAndFulfillment::minimumOrderQuantity': 1,
                'shippingAndFulfillment::forceQuantityMultiplier': 1,
                'shippingAndFulfillment::displaySetQuantity': 1,
                
                # --- YENİ: Bedding Sütunlarına Veri Aktarımı ---
                'bedding::setSingle': b_info['set_single'],
                'bedding::productType': b_info['prod_type'],
                'bedding::size': b_info['bed_size'],
                'bedding::material': b_info['material'],
                'bedding::pieces': b_info['pieces']
            }

            # --- YENİ: RESİMLERİ ANA DATA EXCELİNDEN OKUMA ---
            urls = []
            for col in df_data.columns:
                col_str = str(col).lower()
                # Kolon adında resim/image/url kelimeleri varsa
                if 'image' in col_str or 'resim' in col_str or 'url' in col_str or 'link' in col_str:
                    # 'Number of images' gibi sayı belirten kolonları atla
                    if 'number' in col_str or 'sayı' in col_str or 'adet' in col_str: 
                        continue
                    
                    val = str(row.get(col, '')).strip()
                    # Gerçek bir URL ise (http veya www ile başlıyorsa) ve daha önce eklenmediyse listeye al
                    if val and val.lower() != 'nan' and (val.startswith('http') or val.startswith('www')) and val not in urls:
                        urls.append(val)

            # İlk 5 url'yi ana şablondaki resim sütunlarına yerleştir
            for i in range(min(5, len(urls))): 
                mappings[f'img_{i+1}'] = urls[i]
                
            # 5'ten fazla URL varsa bunları Additional Images sayfasına yönlendirmek için listeye kaydet
            if len(urls) > 5 and sku_key not in processed_skus_for_additional:
                for ext_url in urls[5:]: 
                    additional_images_data.append((sku_key, ext_url))
                processed_skus_for_additional.add(sku_key)

            # Lojistik Hesabı
            if ui_data['is_us']:
                mappings['shippingAndFulfillment::leadTime'] = 600
                mappings['shippingAndFulfillment::replacementLeadTime'] = 120
                if x_in > 0 and y_in > 0 and z_in > 0:
                    stype, fclass = get_ship_type(lbs, x_in, y_in, z_in)
                    mappings['shippingAndFulfillment::shipType'] = stype
                    if fclass: 
                        mappings['shippingAndFulfillment::freightClass'] = fclass

            # Template ile Veriyi Karşılaştırıp Hata Vermek İçin
            if not missing_cols_reported:
                missing = validate_column_mappings(col_map, mappings)
                if missing: 
                    ui_data['missing_cols'] = missing
                missing_cols_reported = True

            # Temel Sütunları Yazma
            for k, v in mappings.items():
                if k in col_map and pd.notna(v) and str(v).strip() != '':
                    ws_main[f"{col_map[k]}{g_satir}"] = v

            # Dinamik Açılır Menü Değerlerini Yazma
            for wid, val in ui_data['dyn_drops'].items():
                if wid in col_map and val:
                    if isinstance(val, list):
                        final_str = "; ".join([str(pv) for pv in val if pv and str(pv) != 'None'])
                    else:
                        final_str = str(val)
                        
                    if final_str: 
                        ws_main[f"{col_map[wid]}{g_satir}"] = final_str

            # MANUEL ÖZEL BOYUT EŞLEŞTİRMELERİNİ YAZMA (Panelden Seçilenler)
            dim_writes = {
                'h': convert_to_inch(raw_h) if ui_data['is_us'] else raw_h,
                'w': convert_to_inch(raw_w) if ui_data['is_us'] else raw_w,
                'd': convert_to_inch(raw_d) if ui_data['is_us'] else raw_d
            }
            
            for dim_type, wids in ui_data['dim_mappings'].items():
                val = dim_writes[dim_type]
                if val is not None:
                    for wid in wids:
                        if wid in col_map: 
                            ws_main[f"{col_map[wid]}{g_satir}"] = val

            # FEATURES ve YATAK BOYUTU NOTU YAZIMI
            satirlar = [s.strip() for s in translate_features(feat_text, ui_data['is_us']).split('\n') if s.strip()]
            bedding_note = generate_bedding_note(feat_text, raw_h, raw_w, b_info['bed_size'], ui_data['is_us'])

            for i, col_let in enumerate(feature_cols):
                if i < 4 and i < len(satirlar):
                    ws_main[f"{col_let}{g_satir}"] = satirlar[i]
                elif i == 4:
                    kalan_metin = " | ".join(satirlar[4:])
                    final_feats = []
                    
                    if kalan_metin: 
                        final_feats.append(kalan_metin)
                    if bedding_note: 
                        final_feats.append(bedding_note)
                        
                    final_feats.append("Made In Türkiye")
                    ws_main[f"{col_let}{g_satir}"] = " | ".join(final_feats)

            processed += 1
            written_rows.append(g_satir)

        except Exception as e:
            errors.append({
                'Satır': index + 2, 
                'Ürün Kodu': sku_key, 
                'Açıklama': str(row.get('DESCRIPTION', '') or '')[:60], 
                'Hata Detayı': str(e)
            })

    # EKSİK ZORUNLU ALANLARI SARIYA BOYA (Hex: FFFF00)
    yellow_fill = openpyxl.styles.PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for c in range(1, ws_main.max_column + 1):
        if str(ws_main.cell(row=3, column=c).value).strip().lower() == "required":
            for r in written_rows:
                cell = ws_main.cell(row=r, column=c)
                if cell.value is None or str(cell.value).strip() == "":
                    cell.fill = yellow_fill

    # EKSTRA RESİMLERİ İLGİLİ SAYFAYA YAZDIRMA
    if additional_images_data:
        add_sheet = None
        for s in wb.sheetnames:
            if 'additional' in s.lower() and 'image' in s.lower():
                add_sheet = wb[s]
                break
                
        if add_sheet:
            sku_col_let, url_col_let, start_row = 'A', 'B', 5
            
            for c in range(1, add_sheet.max_column + 1):
                r1 = str(add_sheet.cell(row=1, column=c).value).lower()
                r4 = str(add_sheet.cell(row=4, column=c).value).lower()
                
                if 'supplier part number' in r4 or 'sku' in r4 or 'part number' in r1: 
                    sku_col_let = add_sheet.cell(row=1, column=c).column_letter
                if 'image file name or url' in r4 or 'url' in r4 or 'media::' in r1: 
                    url_col_let = add_sheet.cell(row=1, column=c).column_letter
                    
            for r in range(4, add_sheet.max_row + 10):
                if not add_sheet[f"{sku_col_let}{r}"].value:
                    start_row = r
                    break
                    
            for sku, url in additional_images_data:
                add_sheet[f"{sku_col_let}{start_row}"] = sku
                add_sheet[f"{url_col_let}{start_row}"] = url
                start_row += 1

    # EKSTRA PAKETLERİ (CARTON) İLGİLİ SAYFAYA YAZDIRMA
    if additional_cartons_data:
        add_carton_sheet = None
        for s in wb.sheetnames:
            if 'additional' in s.lower() and ('carton' in s.lower() or 'package' in s.lower()):
                add_carton_sheet = wb[s]
                break
                
        if add_carton_sheet:
            col_map_c = {}
            for c in range(1, add_carton_sheet.max_column + 1):
                r1 = str(add_carton_sheet.cell(row=1, column=c).value).lower()
                r4 = str(add_carton_sheet.cell(row=4, column=c).value).lower()
                let = add_carton_sheet.cell(row=1, column=c).column_letter
                
                if 'supplier part number' in r4 or 'sku' in r4 or 'part number' in r1: 
                    col_map_c['sku'] = let
                elif 'weight' in r4 or 'weight' in r1: 
                    col_map_c['weight'] = let
                elif 'height' in r4 or 'height' in r1: 
                    col_map_c['height'] = let
                elif 'width' in r4 or 'width' in r1: 
                    col_map_c['width'] = let
                elif 'depth' in r4 or 'depth' in r1: 
                    col_map_c['depth'] = let
            
            if 'sku' in col_map_c:
                start_row = 5
                for r in range(4, add_carton_sheet.max_row + 10):
                    if not add_carton_sheet[f"{col_map_c['sku']}{r}"].value:
                        start_row = r
                        break
                        
                for c_data in additional_cartons_data:
                    c_w_final = round(c_data['kg'] * 2.20462, 2) if ui_data['is_us'] else c_data['kg']
                    c_h_final = round(c_data['x'] * 0.393701, 2) if ui_data['is_us'] else c_data['x']
                    c_w_final_2 = round(c_data['y'] * 0.393701, 2) if ui_data['is_us'] else c_data['y']
                    c_d_final = round(c_data['z'] * 0.393701, 2) if ui_data['is_us'] else c_data['z']
                    
                    add_carton_sheet[f"{col_map_c['sku']}{start_row}"] = c_data['sku']
                    if 'weight' in col_map_c: 
                        add_carton_sheet[f"{col_map_c['weight']}{start_row}"] = c_w_final
                    if 'height' in col_map_c: 
                        add_carton_sheet[f"{col_map_c['height']}{start_row}"] = c_h_final
                    if 'width' in col_map_c: 
                        add_carton_sheet[f"{col_map_c['width']}{start_row}"] = c_w_final_2
                    if 'depth' in col_map_c: 
                        add_carton_sheet[f"{col_map_c['depth']}{start_row}"] = c_d_final
                    
                    start_row += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), processed, skipped, errors

# --- 4. STREAMLIT ARAYÜZÜ (V13) ---

st.set_page_config(page_title="Wayfair Automation V13", layout="wide")
st.title("🛡️ Wayfair Akıllı Ürün Robotu V13")

# --- SOL MENÜ ---
with st.sidebar:
    st.header("⚙️ Genel Ayarlar")
    is_us = st.toggle("US Bölgesi (İnç/Lbs/Yatak Standartları)", value=True)
    
    brand_list = [
        "Lütfen Seçiniz...", "Wallity", "Hanah Home", "Conceptum Hypnose", 
        "Hermia Concept", "Opviq", "Skye Decor", "Nuit des reves", "Evila Originals"
    ]
    brand = st.selectbox("Brand (Marka)", brand_list, index=0, key="brand_sel")
    coll_name = st.text_input("Collection Name", value="", key="coll_name")

# --- DOSYA YÜKLEME ---
u1, u2, u3 = st.columns(3)
with u1: 
    d_file = st.file_uploader("1. Data Excel", type="xlsx")
with u2: 
    t_file = st.file_uploader("2. Template Excel", type="xlsx")
with u3: 
    c_file = st.file_uploader("3. Paket Excel (Opsiyon)", type="xlsx")

# --- OTOMATİK DOLDURULAN SÜTUNLAR ---
# Bu sütunlar arkaplanda sistem tarafından doldurulduğu için arayüzde sorulmaz
AUTO_MAPPED_COLS = {
    'core::supplierPartNumber', 'core::manufacturerPartNumber', 'core::universalProductCode',
    'core::productName', 'featureDescription::romanceCopy', 'featureDescription::overallHeight',
    'featureDescription::overallWidth', 'featureDescription::overallDepth', 'featureDescription::color',
    'featureDescription::genericFeatures', 'shippingAndFulfillment::weight', 'shippingAndFulfillment::height',
    'shippingAndFulfillment::width', 'shippingAndFulfillment::depth', 'price::wholesalePrice',
    'price::manufacturerSuggestedRetailPrice', 'shippingAndFulfillment::minimumOrderQuantity',
    'shippingAndFulfillment::forceQuantityMultiplier', 'shippingAndFulfillment::displaySetQuantity',
    'shippingAndFulfillment::productWeight', 'shippingAndFulfillment::leadTime',
    'shippingAndFulfillment::replacementLeadTime', 'shippingAndFulfillment::shipType',
    'shippingAndFulfillment::freightClass', 'core::collectionName', 'core::manufacturerId',
    'featureDescription::marketingCopy',
    'bedding::setSingle', 'bedding::productType', 'bedding::size', 'bedding::material', 'bedding::pieces'
}

def is_auto_mapped_by_fname(fname):
    """
    Sütun ismine göre gizleme (ÖNEMLİ DÜZELTME: Sadece birebir aynıysa gizler).
    İçinde 'color' geçtiği için 'Leg Color' gibi zorunlu alanları gizlemez.
    """
    f_low = fname.lower().strip()
    exact_matches = {
        'overall height', 'overall width', 'overall depth', 
        'overallheight', 'overallwidth', 'overalldepth', 
        'color', 'colour', 'marketing copy', 'marketingcopy',
        'set / single', 'bedding product type', 'bedding size', 
        'bedding material', 'pieces included', 'total number of pieces included'
    }
    return f_low in exact_matches

# --- DİNAMİK FORM VE ÖZEL ÖLÇÜ PANELİ ---
if d_file and t_file:
    # Streamlit Rerun Dosya Kapanma Hatasını Önlemek İçin Bellek Kopya Oluşturma
    t_bytes = t_file.getvalue()
    
    try: 
        df_v = pd.read_excel(io.BytesIO(t_bytes), sheet_name='Valid Values')
    except: 
        df_v = None

    wb_t = openpyxl.load_workbook(io.BytesIO(t_bytes))
    
    target_name = None
    for s in wb_t.sheetnames:
        if not any(x in s for x in ["Additional", "WAYFAIR", "Instructions", "Valid Values", "Failed"]):
            target_name = s
            break
            
    if not target_name:
        target_name = wb_t.sheetnames[0]
        
    ws_t = wb_t[target_name]

    # Doldurulması gereken uygun sütunları topla
    eligible_cols = []
    for c in range(1, ws_t.max_column + 1):
        wid = str(ws_t.cell(1, c).value).strip()
        status = str(ws_t.cell(3, c).value).strip()
        fname = str(ws_t.cell(4, c).value).strip()
        
        # BÜYÜK/KÜÇÜK HARF HASSASİYETİ KALDIRILDI: status.lower() == 'required'
        if status.lower() == "required" and wid not in AUTO_MAPPED_COLS and not wid.startswith('media::') and not is_auto_mapped_by_fname(fname):
            eligible_cols.append((wid, fname))

    # --- ÖZEL ÖLÇÜ PANELİ ---
    st.markdown("---")
    st.subheader("📐 Özel Ölçü Sütun Eşleştirmeleri")
    st.info("💡 Data Excelinden çekilen Yükseklik/Genişlik/Derinlik değerlerini manuel atamak istediğiniz ekstra sütunları buradan seçebilirsiniz. Seçilen sütunlar alttaki kalabalık listeden otomatik kaldırılır.")
    
    options_dict = {f"{fname} ({wid})": wid for wid, fname in eligible_cols}
    options_list = list(options_dict.keys())
    
    col_h, col_w, col_d = st.columns(3)
    with col_h: 
        h_sel = st.multiselect("Height (Yükseklik) Yazılacaklar", options=options_list)
    with col_w: 
        w_sel = st.multiselect("Width (Genişlik) Yazılacaklar", options=options_list)
    with col_d: 
        d_sel = st.multiselect("Depth (Derinlik) Yazılacaklar", options=options_list)

    # Seçilen özel ölçü sütunlarının WID'lerini ayır
    selected_dim_wids = [options_dict[x] for x in h_sel + w_sel + d_sel]
    dim_mappings = {
        'h': [options_dict[x] for x in h_sel],
        'w': [options_dict[x] for x in w_sel],
        'd': [options_dict[x] for x in d_sel]
    }

    # --- DİĞER ÖZELLİKLER FORMU ---
    st.markdown("---")
    st.subheader(f"📋 {target_name} — Doldurulması Gereken Diğer Özellikler")

    dyn_selections = {}
    cols_ui = st.columns(3)
    idx = 0

    for wid, fname in eligible_cols:
        # Eğer yukarıdaki manuel panelden seçildiyse, normal dropdownu atla
        if wid in selected_dim_wids:
            continue

        with cols_ui[idx % 3]:
            if df_v is not None and fname in df_v.columns:
                opts = list(dict.fromkeys([
                    str(o).strip() for o in df_v[fname].dropna().unique() 
                    if str(o).strip() and str(o).strip() != 'None'
                ]))
            else:
                opts = ["Yes", "No", "Does Not Apply"]

            # GIRINTI (INDENTATION) VE SESSION STATE DÜZELTMESİ
            if wid not in st.session_state['user_prefs']:
                f_low = fname.lower()
                def_val = []
                
                if 'warning required' in f_low: 
                    def_val = ['No']
                elif 'country of manufacturer' in f_low: 
                    def_val = ['Turkey'] if 'Turkey' in opts else (['Türkiye'] if 'Türkiye' in opts else [])
                elif 'uniform packaging and labeling regulations' in f_low: 
                    def_val = ['Yes']
                elif 'reason for restriction' in f_low: 
                    def_val = ['Does Not Apply']
                elif 'general certificate of conformity' in f_low: 
                    def_val = ['Yes']
                elif 'canada product restriction' in f_low: 
                    def_val = ['No']
                elif 'soffa compliant' in f_low: 
                    def_val = ['Does Not Apply']
                elif 'canfer compliant' in f_low: 
                    def_val = ['Does Not Apply']
                elif 'composite wood product (cwp)' in f_low: 
                    def_val = ['Does Not Apply']
                elif 'tsca title vi compliant' in f_low: 
                    def_val = ['Does Not Apply']
                elif 'supplier intended and approved use' in f_low:
                    def_val = [x for x in ['Non Residential Use', 'Residential Use'] if x in opts]
                    if not def_val: 
                        def_val = ['Non Residential Use', 'Residential Use']  
                elif 'commercial warranty' in f_low: 
                    def_val = ['Yes'] 
                elif 'contains flame retardant' in f_low: 
                    def_val = ['No']
                elif 'wayfair compliance verified' in f_low: 
                    def_val = ['Yes']

                st.session_state['user_prefs'][wid] = def_val

            # Buradaki değer atamaları artık doğru girintide (hata vermez)
            saved = st.session_state['user_prefs'].get(wid, [])
            sel = st.multiselect(
                fname, 
                options=opts, 
                default=[x for x in saved if x in opts], 
                key=f"sel_{wid}"
            )
            dyn_selections[wid] = sel
            st.session_state['user_prefs'][wid] = sel
        idx += 1

    # --- BUTON VE ÇALIŞTIRMA ---
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Wayfair Dosyasını Hazırla", type="primary", width='stretch'):
        
        if brand == "Lütfen Seçiniz...": 
            st.error("⚠️ Lütfen sol menüden bir Marka (Brand) seçiniz!")
            st.stop()
            
        if not coll_name.strip(): 
            st.error("⚠️ Lütfen sol menüden Collection Name alanını doldurunuz!")
            st.stop()

        ui_data = {
            'is_us': is_us,
            'brand': brand,
            'coll_name': coll_name,
            'dyn_drops': dyn_selections,
            'dim_mappings': dim_mappings,
            'missing_cols': []
        }

        progress_bar = st.progress(0, text="Hazırlanıyor...")
        def update_progress(val): 
            progress_bar.progress(min(val, 1.0), text=f"İşleniyor... %{int(val * 100)}")

        with st.spinner("Excel dosyası işleniyor..."):
            
            # Yenileme sorunlarını önlemek için işlem butonunda da taze bellek okuması kullanıyoruz
            d_io = io.BytesIO(d_file.getvalue())
            t_io = io.BytesIO(t_file.getvalue())
            c_io = io.BytesIO(c_file.getvalue()) if c_file else None
            
            # Artık resim dosyası argümanı yok
            res, processed, skipped, errors = process_wayfair_v13(
                d_io, t_io, ui_data, carton_file=c_io, progress_callback=update_progress
            )

        progress_bar.progress(1.0, text="✅ Tamamlandı!")
        
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("✅ İşlenen", processed)
        m2.metric("⏭️ Atlanan", len(skipped))
        m3.metric("❌ Hatalı", len(errors))
        m4.metric("📦 Toplam", processed + len(skipped) + len(errors))

        if ui_data.get('missing_cols'):
            with st.expander(f"⚠️ {len(ui_data['missing_cols'])} Sütun Template'de Bulunamadı", expanded=True):
                st.warning("Bu sütunlar mapping'de tanımlı ama template'de yok — ilgili veriler yazılamadı:")
                st.code("\n".join(ui_data['missing_cols']))

        if skipped:
            with st.expander(f"⏭️ Atlanan Satırlar — {len(skipped)} ürün", expanded=False): 
                st.dataframe(pd.DataFrame(skipped), width='stretch')

        if errors:
            with st.expander(f"❌ Hatalı Satırlar — {len(errors)} ürün", expanded=True):
                st.error("Hata Detayları:")
                st.dataframe(pd.DataFrame(errors), width='stretch')

        if processed > 0:
            st.success(f"✅ {processed} ürün başarıyla işlendi.")
            
            safe_coll_name = str(coll_name).strip() if str(coll_name).strip() else "Wayfair_Upload"
            st.download_button(
                label="📥 Hazır Excel'i İndir",
                data=res,
                file_name=f"{safe_coll_name}-Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
