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
    - Çap (Diameter) tespiti eklendi, \b (sınır) regex hatası düzeltildi
    """
    if pd.isna(text):
        return None, None, None
    text = str(text)

    # Çap (Diameter) tespiti
    dia = get_dim_val(r'(?:Diameter|Çap|Dia|Ø)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)

    # \b eklendi ki "Width" içindeki 'h' Height olarak, veya 'd' Depth olarak tetiklenmesin
    w = get_dim_val(r'(?:Width|Genişlik|Side to Side|\bW\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)
    h = get_dim_val(r'(?:Height|Yükseklik|Top to Bottom|\bH\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)
    d = get_dim_val(r'(?:Depth|Derinlik|Front to Back|\bD\b)\s*[:\-\s]\s*(\d+(?:[.,]\d+)?)', text)

    # Çap varsa ve Genişlik/Derinlik belirtilmemişse, çapı buralara ata
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
    """
    Geliştirilmiş çeviri:
    - Yalnızca gerçek boyut bağlamındaki sayıları çevirir
    - Alfanümerik önekli değerleri (SKU, kod, fiyat vb.) atlar
    - Wayfair'in kabul etmediği Ø veya Çap ifadelerini Diameter'a çevirir
    """
    if pd.isna(text):
        return ""
    text = str(text)

    # --- YENİ: Ø işaretini ve "Çap" kelimesini "Diameter: " olarak düzelt ---
    # Ø45, Ø:45, Ø - 45 gibi tüm kullanımları yakalar ve Diameter: 45 yapar.
    text = re.sub(r'Ø\s*[:\-]?\s*', 'Diameter: ', text)
    text = re.sub(r'(?i)\bÇap\s*[:\-]?\s*', 'Diameter: ', text)

    if not do_conversion:
        return text

    def c_in(m):
        return f"{round(float(m.group(1).replace(',', '.')) * 0.393701, 2)}"

    def c_mm(m):
        return f"{round(float(m.group(1).replace(',', '.')) * 0.0393701, 2)}"

    # Sayı + birim: "13.39cm", "30 cm", "5.5 mm" — boşluk olsa da olmasa da yakala
    # Önünde harf olmamalı (SKU/kelime içi koruma): scam, vacuum etkilenmiyor
    text = re.sub(r'(?<![a-zA-Z])(\d+(?:[\.,]\d+)?)\s*cm(?![a-zA-Z])', lambda m: c_in(m) + ' inches', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z])(\d+(?:[\.,]\d+)?)\s*mm(?![a-zA-Z])', lambda m: c_mm(m) + ' inches', text, flags=re.IGNORECASE)
    # Boyut çarpımı formatları: 50x100x30 (cm zaten üstte işlendi, sadece saf sayılar kalır)
    text = re.sub(
        r'(?<![a-zA-Z0-9:])(\d+(?:[\.,]\d+)?)(?=\s*[xX×]\s*\d)',
        c_in, text
    )
    # Kalan yalnız birim etiketlerini temizle
    text = re.sub(r'(?<![a-zA-Z])\bcm\b(?![a-zA-Z])', 'inches', text, flags=re.IGNORECASE)
    text = re.sub(r'(?<![a-zA-Z])\bmm\b(?![a-zA-Z])', 'inches', text, flags=re.IGNORECASE)
    return text

def calculate_freight_class(weight_lbs, l_in, w_in, h_in):
    vol = (l_in * w_in * h_in) / 1728
    if vol == 0:
        return "60"
    dens = weight_lbs / vol
    if dens < 1: return "400"
    elif dens < 2: return "300"
    elif dens < 4: return "200"
    elif dens < 6: return "150"
    elif dens < 8: return "125"
    elif dens < 10: return "100"
    else: return "60"

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

# --- 3. ANA İŞLEME MOTORU ---

def process_wayfair_v9(data_file, template_file, ui_data, image_file=None, carton_file=None, progress_callback=None):
    data_file.seek(0)
    template_file.seek(0)

    df_data = pd.read_excel(data_file)
    
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
                    
                    carton_dict[c_sku].append({'kg': w_val, 'x': x_val, 'y': y_val, 'z': z_val})

    # --- RESİM EXCEL'İNİ İŞLEME (SÖZLÜK OLUŞTURMA) ---
    img_dict = {}
    if image_file is not None:
        image_file.seek(0)
        df_img = pd.read_excel(image_file)
        sku_col = None
        for col in df_img.columns:
            if 'sku' in str(col).lower() or 'code' in str(col).lower():
                sku_col = col
                break
        if sku_col is not None:
            for _, r in df_img.iterrows():
                sku_val = str(r[sku_col]).strip()
                if sku_val and sku_val.lower() != 'nan':
                    urls = []
                    for col in df_img.columns:
                        col_str = str(col).lower()
                        # SKU ve sayı bildiren sütunları atla
                        if col == sku_col or 'number' in col_str or 'adet' in col_str or 'sayı' in col_str:
                            continue
                        val = str(r[col]).strip()
                        if val and val.lower() != 'nan':
                            # AYNI LİNKİN 2 KERE EKLENMESİNİ ENGELLİYORUZ
                            if val not in urls:
                                urls.append(val)
                    if urls:
                        img_dict[sku_val] = urls

    wb = openpyxl.load_workbook(template_file)

    # Ana şablonu bul
    target_sheet = None
    for s in wb.sheetnames:
        if not any(x in s for x in ["Additional", "WAYFAIR", "Instructions", "Valid Values", "Failed"]):
            target_sheet = s
            break
    ws_main = wb[target_sheet or wb.sheetnames[0]]

    # YENİ EKLENEN KISIM: Akıllı Sütun Eşleştirme (Smart Column Mapping)
    col_map = {}
    for c in range(1, ws_main.max_column + 1):
        r1_val = str(ws_main.cell(row=1, column=c).value).strip() if ws_main.cell(row=1, column=c).value else ""
        r4_val = str(ws_main.cell(row=4, column=c).value).strip() if ws_main.cell(row=4, column=c).value else ""
        col_let = ws_main.cell(row=1, column=c).column_letter
        
        if r1_val:
            col_map[r1_val] = col_let
            
        r4_lower = r4_val.lower()
        r1_lower = r1_val.lower()

        # COLOR EŞLEŞTİRMESİ
        if ('color' in r4_lower or 'colour' in r4_lower or r1_lower.endswith('::color')):
            if 'leg' not in r4_lower and 'base' not in r4_lower and 'shade' not in r4_lower:
                col_map['featureDescription::color'] = col_let

        # BOYUT (DIMENSION) EŞLEŞTİRMELERİ
        if 'overall height' in r4_lower or 'overallheight' in r1_lower:
            col_map['featureDescription::overallHeight'] = col_let
        elif 'overall width' in r4_lower or 'overallwidth' in r1_lower:
            col_map['featureDescription::overallWidth'] = col_let
        elif 'overall depth' in r4_lower or 'overalldepth' in r1_lower:
            col_map['featureDescription::overallDepth'] = col_let
            
        # RESİM SÜTUNLARI EŞLEŞTİRMESİ (İlk 5 URL)
        for i in range(1, 6):
            if f'image file name or url {i}' in r4_lower:
                col_map[f'img_{i}'] = col_let

    # Generic features sütunları
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
        pkg_count = row.get('NUMBER OF PACKAGES', 1)

        # Temel boyutları al (Varsayılan olarak data excel'den)
        kg = float(row.get('WEIGHT (Kg)', 0) or 0)
        x_cm = float(row.get('PACKAGING SIZE - X (cm)', 0) or 0)
        y_cm = float(row.get('PACKAGING SIZE - Y (cm)', 0) or 0)
        z_cm = float(row.get('PACKAGING SIZE - Z (cm)', 0) or 0)

        # Data excel'e göre varsayılan ürün ağırlığı (Paket excel yoksa)
        prod_weight_lbs = round((kg - 0.1) * 2.20462, 2) if kg > 0.1 else 0

        # --- PAKET EXCELİ ÖNCELİĞİ ---
        if carton_file is not None and sku_key in carton_dict and len(carton_dict[sku_key]) > 0:
            cartons = carton_dict[sku_key]
            # İlk paket
            kg = cartons[0]['kg']
            x_cm = cartons[0]['x']
            y_cm = cartons[0]['y']
            z_cm = cartons[0]['z']
            
            # --- PRODUCT WEIGHT: TÜM KOLİLERİN TOPLAMINDAN 5 LBS ÇIKAR ---
            total_kg = sum(c['kg'] for c in cartons)
            total_lbs = total_kg * 2.20462
            prod_weight_lbs = round(total_lbs - 5, 2)
            if prod_weight_lbs < 0:
                prod_weight_lbs = 0
            
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

            if not any([raw_h, raw_w, raw_d]) and feat_text and str(feat_text).strip():
                skipped.append({
                    'Satır': index + 2,
                    'Ürün Kodu': sku_key,
                    'Açıklama': str(row.get('DESCRIPTION', '') or '')[:60],
                    'Neden': f"⚠️ Boyut (H/W/D) features metninden çıkarılamadı"
                })
            
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
                'core::productName': row.get('DESCRIPTION'),
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
                'shippingAndFulfillment::displaySetQuantity': 1
            }

            # RESİMLERİ EŞLEŞTİRME
            if sku_key in img_dict:
                urls = img_dict[sku_key]
                for i in range(min(5, len(urls))):
                    mappings[f'img_{i+1}'] = urls[i]
                
                if len(urls) > 5 and sku_key not in processed_skus_for_additional:
                    for ext_url in urls[5:]:
                        additional_images_data.append((sku_key, ext_url))
                    processed_skus_for_additional.add(sku_key)

            if ui_data['is_us']:
                mappings['shippingAndFulfillment::leadTime'] = 600
                mappings['shippingAndFulfillment::replacementLeadTime'] = 120
                if x_in > 0 and y_in > 0 and z_in > 0:
                    stype, fclass = get_ship_type(lbs, x_in, y_in, z_in)
                    mappings['shippingAndFulfillment::shipType'] = stype
                    if fclass:
                        mappings['shippingAndFulfillment::freightClass'] = fclass

            if not missing_cols_reported:
                missing = validate_column_mappings(col_map, mappings)
                if missing:
                    ui_data['missing_cols'] = missing
                missing_cols_reported = True

            for k, v in mappings.items():
                if k in col_map and pd.notna(v) and str(v).strip() != '':
                    ws_main[f"{col_map[k]}{g_satir}"] = v

            # DİNAMİK AÇILIR MENÜ DEĞERLERİNİ YAZMA (MAGIC OPTIONS DAHİL)
            for wid, val in ui_data['dyn_drops'].items():
                if wid in col_map and val:
                    if isinstance(val, list):
                        processed_vals = []
                        for v_item in val:
                            if v_item == "📏 DataHeight":
                                processed_vals.append(str(mappings.get('featureDescription::overallHeight', '')))
                            elif v_item == "📏 DataWidth":
                                processed_vals.append(str(mappings.get('featureDescription::overallWidth', '')))
                            elif v_item == "📏 DataDepth":
                                processed_vals.append(str(mappings.get('featureDescription::overallDepth', '')))
                            else:
                                processed_vals.append(str(v_item))
                        
                        final_str = "; ".join([pv for pv in processed_vals if pv and pv != 'None'])
                        if final_str:
                            ws_main[f"{col_map[wid]}{g_satir}"] = final_str
                    else:
                        ws_main[f"{col_map[wid]}{g_satir}"] = str(val)

            satirlar = [s.strip() for s in translate_features(feat_text, ui_data['is_us']).split('\n') if s.strip()]
            for i, col_let in enumerate(feature_cols):
                if i < 4 and i < len(satirlar):
                    ws_main[f"{col_let}{g_satir}"] = satirlar[i]
                elif i == 4:
                    kalan_metin = " | ".join(satirlar[4:])
                    ws_main[f"{col_let}{g_satir}"] = (kalan_metin + " | Made In Türkiye").strip(" | ")

            processed += 1
            written_rows.append(g_satir)

        except Exception as e:
            errors.append({
                'Satır': index + 2,
                'Ürün Kodu': sku_key,
                'Açıklama': str(row.get('DESCRIPTION', '') or '')[:60],
                'Hata Detayı': str(e)
            })

    # EKSİK ZORUNLU ALANLARI PEMBEYE BOYA
    pink_fill = openpyxl.styles.PatternFill(start_color="FFFFC0CB", end_color="FFFFC0CB", fill_type="solid")
    for c in range(1, ws_main.max_column + 1):
        status = str(ws_main.cell(row=3, column=c).value).strip()
        if status == "Required":
            for r in written_rows:
                cell = ws_main.cell(row=r, column=c)
                if cell.value is None or str(cell.value).strip() == "":
                    cell.fill = pink_fill

    # EKSTRA RESİMLERİ YAZDIRMA
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

    # EKSTRA PAKETLERİ YAZDIRMA
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
                    c_kg, c_x, c_y, c_z = c_data['kg'], c_data['x'], c_data['y'], c_data['z']
                    c_w_final = round(c_kg * 2.20462, 2) if ui_data['is_us'] else c_kg
                    c_h_final = round(c_x * 0.393701, 2) if ui_data['is_us'] else c_x
                    c_w_final_2 = round(c_y * 0.393701, 2) if ui_data['is_us'] else c_y
                    c_d_final = round(c_z * 0.393701, 2) if ui_data['is_us'] else c_z
                    
                    add_carton_sheet[f"{col_map_c['sku']}{start_row}"] = c_data['sku']
                    if 'weight' in col_map_c: add_carton_sheet[f"{col_map_c['weight']}{start_row}"] = c_w_final
                    if 'height' in col_map_c: add_carton_sheet[f"{col_map_c['height']}{start_row}"] = c_h_final
                    if 'width' in col_map_c: add_carton_sheet[f"{col_map_c['width']}{start_row}"] = c_w_final_2
                    if 'depth' in col_map_c: add_carton_sheet[f"{col_map_c['depth']}{start_row}"] = c_d_final
                    start_row += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), processed, skipped, errors


# --- 4. STREAMLIT ARAYÜZÜ (V9) ---

st.set_page_config(page_title="Wayfair Automation V9", layout="wide")
st.title("🛡️ Wayfair Akıllı Ürün Robotu V9")

# --- SOL MENÜ ---
with st.sidebar:
    st.header("⚙️ Genel Ayarlar")
    is_us = st.toggle("US Bölgesi (İnç / Lojistik Seçimi)", value=True)

    brand_list = [
        "Lütfen Seçiniz...", "Wallity", "Hanah Home", "Conceptum Hypnose", "Hermia Concept",
        "Opviq", "Skye Decor", "Nuit des reves", "Evila Originals"
    ]
    brand = st.selectbox("Brand (Marka)", brand_list, index=0, key="brand_sel")
    coll_name = st.text_input("Collection Name", value="", key="coll_name")


# --- DOSYA YÜKLEME ---
u1, u2, u3, u4 = st.columns(4)
with u1:
    d_file = st.file_uploader("1. Data Excel", type="xlsx")
with u2:
    t_file = st.file_uploader("2. Template Excel", type="xlsx")
with u3:
    i_file = st.file_uploader("3. Image Excel (Opsiyon)", type="xlsx")
with u4:
    c_file = st.file_uploader("4. Paket Excel (Opsiyon)", type="xlsx")

# --- VERİ ÖNİZLEME ---
if d_file:
    with st.expander("📊 Veri Önizleme (İlk 5 Satır)", expanded=False):
        d_file.seek(0)
        df_preview = pd.read_excel(d_file)
        st.dataframe(df_preview.head(5), width='stretch')
        multi_pkg = df_preview[df_preview.get('NUMBER OF PACKAGES', pd.Series(dtype=int)) != 1] if 'NUMBER OF PACKAGES' in df_preview.columns else pd.DataFrame()
        col1, col2, col3 = st.columns(3)
        col1.metric("Toplam Ürün", len(df_preview))
        col2.metric("Toplam Sütun", len(df_preview.columns))
        col3.metric("Çoklu Paket (Ek Veri İster)", len(multi_pkg))
        d_file.seek(0)
        
    if len(multi_pkg) > 0 and c_file is None:
        st.warning(f"⚠️ DİKKAT: Ana veride {len(multi_pkg)} adet çoklu paketli ürün tespit edildi! İşlenebilmeleri için '4. Paket Excel'ini yüklemelisiniz, aksi halde bu ürünler atlanacaktır.")

# --- OTOMATİK DOLDURULAN SÜTUNLAR ---
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
    'featureDescription::marketingCopy'
}

AUTO_MAPPED_FNAME_KEYWORDS = [
    'overall height', 'overall width', 'overall depth',
    'overallheight', 'overallwidth', 'overalldepth',
    'color', 'colour', 'marketing copy', 'marketingcopy'
]

def is_auto_mapped_by_fname(fname):
    fl = fname.lower()
    return any(kw in fl for kw in AUTO_MAPPED_FNAME_KEYWORDS)

# --- DİNAMİK FORM ---
if d_file and t_file:
    try:
        df_v = pd.read_excel(t_file, sheet_name='Valid Values')
    except:
        df_v = None

    t_file.seek(0)
    wb_t = openpyxl.load_workbook(t_file)

    target_name = None
    for s in wb_t.sheetnames:
        if not any(x in s for x in ["Additional", "WAYFAIR", "Instructions", "Valid Values", "Failed"]):
            target_name = s
            break
    if not target_name:
        target_name = wb_t.sheetnames[0]
    ws_t = wb_t[target_name]
    t_file.seek(0)

    st.subheader(f"📋 {target_name} — Doldurulması Gereken Özellikler")

    dyn_selections = {}
    cols_ui = st.columns(3)
    idx = 0

    for c in range(1, ws_t.max_column + 1):
        wid = str(ws_t.cell(1, c).value).strip()
        status = str(ws_t.cell(3, c).value).strip()
        fname = str(ws_t.cell(4, c).value).strip()

        if status == "Required" and wid not in AUTO_MAPPED_COLS and not wid.startswith('media::') and not is_auto_mapped_by_fname(fname):
            with cols_ui[idx % 3]:
                if 'total' in fname.lower() and 'piece' in fname.lower():
                    val = st.text_input(
                        fname,
                        value=st.session_state['user_prefs'].get(wid, "1"),
                        key=f"input_{wid}"
                    )
                    dyn_selections[wid] = val
                    st.session_state['user_prefs'][wid] = val
                else:
                    if df_v is not None and fname in df_v.columns:
                        opts = list(dict.fromkeys([
                            str(o).strip() for o in df_v[fname].dropna().unique()
                            if str(o).strip() and str(o).strip() != 'None'
                        ]))
                    else:
                        opts = ["Yes", "No", "Does Not Apply"]

                    # MAGIC OPTIONS EKLENİYOR
                    magic_opts = ["📏 DataHeight", "📏 DataWidth", "📏 DataDepth"]
                    opts = magic_opts + opts

                    # DEFAULT ATAMALARI
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
                        elif 'tsca title vi compliant (formaldehyde emissions)' in f_low:
                            def_val = ['Does Not Apply']
                        elif 'supplier intended and approved use' in f_low:
                            def_val = [x for x in ['Non Residential Use', 'Residential Use'] if x in opts]
                            if not def_val:
                                def_val = ['Non Residential Use', 'Residential Use']  
                        elif 'commercial warranty' in f_low:
                            def_val = ['Yes'] 
                        elif 'contains flame retardant materials' in f_low:
                            def_val = ['No']
                        elif 'wayfair compliance verified program' in f_low:
                            def_val = ['Yes']

                        st.session_state['user_prefs'][wid] = def_val

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

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Wayfair Dosyasını Hazırla", type="primary", width='stretch'):

        # Başlamadan önce Marka ve Koleksiyon Adı kontrolü
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
            'missing_cols': []
        }

        progress_bar = st.progress(0, text="Hazırlanıyor...")

        def update_progress(val):
            progress_bar.progress(min(val, 1.0), text=f"İşleniyor... %{int(val * 100)}")

        with st.spinner("Excel dosyası işleniyor..."):
            res, processed, skipped, errors = process_wayfair_v9(
                d_file, t_file, ui_data, image_file=i_file, carton_file=c_file, progress_callback=update_progress
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
                st.error(f"{len(errors)} satırda beklenmedik hata oluştu. Aşağıda detaylar:")
                st.dataframe(pd.DataFrame(errors), width='stretch')

        if processed > 0:
            st.success(f"✅ {processed} ürün başarıyla işlendi. Dosyayı indirebilirsiniz.")
            
            safe_coll_name = str(coll_name).strip() if str(coll_name).strip() else "Wayfair_Upload"
            
            st.download_button(
                label="📥 Hazır Excel'i İndir",
                data=res,
                file_name=f"{safe_coll_name}-Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Hiç ürün işlenemedi. Lütfen hata raporunu inceleyin.")
