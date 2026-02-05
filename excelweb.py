import pandas as pd
import math
import re
import io
import streamlit as st

# --- Sabitler ---
KG_TO_LBS = 2.20462
CM_TO_INCH = 0.393701
MADE_IN_TURKEY = "Made In Türkiye"

# --- Yardımcı Fonksiyonlar (Orijinal Kodundan Alındı) ---
def extract_dimensions_from_string(text_to_search):
    def find_dimension_value(pattern, text):
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            try:
                value_str = match.group(1).replace(',', '.')
                return float(value_str)
            except (ValueError, TypeError, AttributeError):
                return None
        return None

    width_pattern = r'(?:Width|Genişlik):\s*(\d+(?:[.,]\d+)?)(?:\s*cm)?'
    height_pattern = r'(?:Height|Yükseklik):\s*(\d+(?:[.,]\d+)?)(?:\s*cm)?'
    depth_pattern = r'(?:Depth|Derinlik):\s*(\d+(?:[.,]\d+)?)(?:\s*cm)?'
    length_pattern = r'(?:Length|Uzunluk):\s*(\d+(?:[.,]\d+)?)(?:\s*cm)?'
    diameter_pattern = r'(?:Diameter|Çap):\s*(\d+(?:[.,]\d+)?)(?:\s*cm)?'

    w = find_dimension_value(width_pattern, text_to_search)
    h = find_dimension_value(height_pattern, text_to_search)
    d = find_dimension_value(depth_pattern, text_to_search)
    l = find_dimension_value(length_pattern, text_to_search)
    diam = find_dimension_value(diameter_pattern, text_to_search)

    y_dim_source_value = d if d is not None else l

    if w is not None and h is not None and y_dim_source_value is not None:
        return (w, y_dim_source_value, h)

    if diam is not None and h is not None:
        return (diam, diam, h)

    dimension_pattern_xyz = r'(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)(?:\s*x\s*(\d+(?:[.,]\d+)?))?(?:\s*cm)?'
    match_xyz = re.search(dimension_pattern_xyz, text_to_search, re.IGNORECASE)
    if match_xyz:
        try:
            x = float(match_xyz.group(1).replace(',', '.'))
            y = float(match_xyz.group(2).replace(',', '.'))
            z = float(match_xyz.group(3).replace(',', '.')) if match_xyz.group(3) else None
            return (x, y, z)
        except:
            return None
    return None

def clean_feature_list(features_str):
    if pd.isna(features_str) or features_str == "":
        return []
    features = re.split(r'\s*(?:\\n|\n)\s*', str(features_str).strip())
    return [f.strip() for f in features if f and f.strip()]

def convert_cm_to_inch(cm_str):
    cm_input_str = str(cm_str) if cm_str is not None else ''
    try:
        cleaned_cm_str = cm_input_str.replace(',', '.').strip()
        if cleaned_cm_str:
            cm_numeric = float(cleaned_cm_str)
            if not pd.isna(cm_numeric):
                return round(cm_numeric * CM_TO_INCH, 2)
        return ''
    except:
        return cm_input_str

# --- Streamlit Arayüzü ---
st.set_page_config(page_title="Excel Dönüştürücü", page_icon="📊")
st.title("📊 Excel Veri İşleme Paneli")
st.markdown("Arkadaşlarınızla paylaşabileceğiniz web tabanlı Excel işleme aracı.")

uploaded_file = st.file_uploader("İşlemek istediğiniz Excel dosyasını seçin", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, keep_default_na=False, dtype=str)
        st.info(f"Dosya okundu: {len(df)} satır bulundu. İşlem başlıyor...")
        
        processed_data = []
        output_headers = [
            'CODE', 'EAN CODE', 'COLOR', 'DESCRIPTION',
            'Feature 1', 'Feature 2', 'Feature 3', 'Feature 4', 'Feature 5',
            'IMAGE', 'PRICE', 'RETAIL PRICE', 'NUMBER OF PACKAGES',
            'PRODUCT SIZE - X (inch)', 'PRODUCT SIZE - Y (inch)', 'PRODUCT SIZE - Z (inch)',
            'CARTON WEIGHT (LBS)', 'WEIGHT (LBS)',
            'PACKAGING SIZE - X (inch)', 'PACKAGING SIZE - Y (inch)', 'PACKAGING SIZE - Z (inch)'
        ]

        # İşleme döngüsü (Orijinal mantığın aynısı)
        for index, row in df.iterrows():
            combined_features_text = str(row.get('FEATURES', '')) + "\n" + str(row.get('EXTRA FEATURES', ''))
            dims = extract_dimensions_from_string(combined_features_text)
            
            p_x, p_y, p_z = dims if dims else ('', '', '')
            
            # Renk ve Özellik İşleme
            processed_color = str(row.get('COLOR', '')).replace('\\n', ';').replace('\n', ';')
            processed_color = re.sub(r';+', ';', processed_color).strip(';')
            
            feat_list = clean_feature_list(row.get('FEATURES', ''))
            extra_feat = str(row.get('EXTRA FEATURES', ''))
            if extra_feat and "number of packages" not in extra_feat.lower():
                feat_list.extend(clean_feature_list(extra_feat))
            
            feature_cols = [""] * 5
            for i in range(min(len(feat_list), 4)): feature_cols[i] = feat_list[i]
            if len(feat_list) >= 5: feature_cols[4] = "\n".join(feat_list[4:])
            
            # Made in Turkey eklemesi
            if not any(MADE_IN_TURKEY in str(f) for f in feature_cols):
                idx = min(len(feat_list), 4)
                feature_cols[idx] = (str(feature_cols[idx]) + f"\n{MADE_IN_TURKEY}").strip()

            # Ağırlık Dönüşümü
            try:
                w_kg = float(str(row.get('WEIGHT (Kg)', '')).replace(',', '.'))
                c_lbs = round(w_kg * KG_TO_LBS, 2)
                p_lbs = max(0, round(c_lbs - 0.01, 2))
            except:
                c_lbs = p_lbs = row.get('WEIGHT (Kg)', '')

            processed_row = {
                'CODE': row.get('CODE', ''),
                'EAN CODE': row.get('EAN CODE', ''),
                'COLOR': processed_color,
                'DESCRIPTION': row.get('DESCRIPTION', ''),
                'Feature 1': feature_cols[0], 'Feature 2': feature_cols[1],
                'Feature 3': feature_cols[2], 'Feature 4': feature_cols[3],
                'Feature 5': feature_cols[4],
                'IMAGE': row.get('IMAGE', ''), 'PRICE': row.get('PRICE', ''),
                'RETAIL PRICE': row.get('RETAIL PRICE', ''),
                'NUMBER OF PACKAGES': row.get('NUMBER OF PACKAGES', ''),
                'PRODUCT SIZE - X (inch)': convert_cm_to_inch(p_x),
                'PRODUCT SIZE - Y (inch)': convert_cm_to_inch(p_y),
                'PRODUCT SIZE - Z (inch)': convert_cm_to_inch(p_z),
                'CARTON WEIGHT (LBS)': c_lbs, 'WEIGHT (LBS)': p_lbs,
                'PACKAGING SIZE - X (inch)': convert_cm_to_inch(row.get('PACKAGING SIZE - X (cm)', '')),
                'PACKAGING SIZE - Y (inch)': convert_cm_to_inch(row.get('PACKAGING SIZE - Y (cm)', '')),
                'PACKAGING SIZE - Z (inch)': convert_cm_to_inch(row.get('PACKAGING SIZE - Z (cm)', ''))
            }
            processed_data.append(processed_row)

        # Sonuçları hazırla
        output_df = pd.DataFrame(processed_data, columns=output_headers)
        st.success("İşlem tamamlandı!")
        st.dataframe(output_df.head()) # Önizleme

        # İndirme Butonu
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 İşlenmiş Excel'i İndir",
            data=output.getvalue(),
            file_name=f"islenmis_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Hata oluştu: {e}")
