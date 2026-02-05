import pandas as pd
import math
import re
import io
import streamlit as st

# --- Sabitler ---
KG_TO_LBS = 2.20462
CM_TO_INCH = 0.393701
MADE_IN_TURKEY = "Made In Türkiye"

# --- Yardımcı Fonksiyonlar ---
def extract_dimensions_from_string(text_to_search):
    def find_dimension_value(pattern, text):
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            try:
                value_str = match.group(1).replace(',', '.')
                return float(value_str)
            except: return None
        return None

    w = find_dimension_value(r'(?:Width|Genişlik):\s*(\d+(?:[.,]\d+)?)', text_to_search)
    h = find_dimension_value(r'(?:Height|Yükseklik):\s*(\d+(?:[.,]\d+)?)', text_to_search)
    d = find_dimension_value(r'(?:Depth|Derinlik):\s*(\d+(?:[.,]\d+)?)', text_to_search)
    l = find_dimension_value(r'(?:Length|Uzunluk):\s*(\d+(?:[.,]\d+)?)', text_to_search)
    diam = find_dimension_value(r'(?:Diameter|Çap):\s*(\d+(?:[.,]\d+)?)', text_to_search)

    y_val = d if d is not None else l
    if w is not None and h is not None and y_val is not None: return (w, y_val, h)
    if diam is not None and h is not None: return (diam, diam, h)

    xyz_pattern = r'(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)(?:\s*x\s*(\d+(?:[.,]\d+)?))?'
    match = re.search(xyz_pattern, text_to_search)
    if match:
        try:
            x = float(match.group(1).replace(',', '.'))
            y = float(match.group(2).replace(',', '.'))
            z = float(match.group(3).replace(',', '.')) if match.group(3) else None
            return (x, y, z)
        except: return None
    return None

def clean_feature_list(features_str):
    if pd.isna(features_str) or features_str == "": return []
    features = re.split(r'\s*(?:\\n|\n)\s*', str(features_str).strip())
    return [f.strip() for f in features if f and f.strip()]

def convert_value(val, unit_choice):
    """Kullanıcı inch seçerse çevirir, cm seçerse olduğu gibi bırakır."""
    if not val or pd.isna(val) or val == '': return ''
    try:
        num_val = float(str(val).replace(',', '.'))
        if unit_choice == "inch":
            return round(num_val * CM_TO_INCH, 2)
        return round(num_val, 2)
    except:
        return val

# --- Streamlit Arayüzü ---
st.set_page_config(page_title="Asir Tools", layout="wide")
st.title("📊 Excel Veri Dönüştürücü")

# Seçenekler Paneli
with st.sidebar:
    st.header("Ayarlar")
    unit_choice = st.radio("Ölçü Birimi Seçin: (Sadece kutu ölçüleri çevrilir)", ("cm", "inch"), index=1) # Default inch
    st.write(f"Şu an seçili: **{unit_choice}**")

uploaded_file = st.file_uploader("Excel dosyasını seçin", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str).fillna('')
        
        processed_data = []
        # Kolon başlıklarını isteğine göre düzenledim
        unit_label = f"({unit_choice})"
        output_headers = [
            'CODE', 'EAN CODE', 'COLOR', 'DESCRIPTION',
            'Feature 1', 'Feature 2', 'Feature 3', 'Feature 4', 'Feature 5',
            'IMAGE', 'PRICE', ' ', 'RETAIL PRICE', 'NUMBER OF PACKAGES', 
            'WEIGHT (LBS)', 
            f'PRODUCT SIZE - X {unit_label}', f'PRODUCT SIZE - Y {unit_label}', f'PRODUCT SIZE - Z {unit_label}',
            'CARTON WEIGHT (LBS)',
            f'PACKAGING SIZE - X {unit_label}', f'PACKAGING SIZE - Y {unit_label}', f'PACKAGING SIZE - Z {unit_label}'
        ]

        for index, row in df.iterrows():
            # 🔥 Boş satır kontrolü 
            if not str(row.get('CODE', '')).strip():
                continue

            # Özelliklerden boyut çekme
            combined_text = str(row.get('FEATURES', '')) + "\n" + str(row.get('EXTRA FEATURES', ''))
            dims = extract_dimensions_from_string(combined_text)
            p_x, p_y, p_z = dims if dims else ('', '', '')

            # Özellik kolonlarını hazırlama (Made in Türkiye dahil)
            feat_list = clean_feature_list(row.get('FEATURES', ''))
            if "number of packages" not in str(row.get('EXTRA FEATURES', '')).lower():
                feat_list.extend(clean_feature_list(row.get('EXTRA FEATURES', '')))
            
            feature_cols = [""] * 5
            for i in range(min(len(feat_list), 4)): feature_cols[i] = feat_list[i]
            if len(feat_list) >= 5: feature_cols[4] = "\n".join(feat_list[4:])
            
            if not any(MADE_IN_TURKEY in str(f) for f in feature_cols):
                for i in range(5):
                    if feature_cols[i] == "":
                        feature_cols[i] = MADE_IN_TURKEY
                        break
                else: feature_cols[4] += f"\n{MADE_IN_TURKEY}"

            # Ağırlık Hesaplama
            try:
                w_kg = float(str(row.get('WEIGHT (Kg)', '')).replace(',', '.'))
                c_lbs = round(w_kg * KG_TO_LBS, 2)
                p_lbs = max(0.0, round(c_lbs - 0.01, 2))
            except:
                c_lbs = p_lbs = row.get('WEIGHT (Kg)', '')

            # Yeni Satırı Oluştur (Sıralama isteğine göre)
            processed_row = [
                row.get('CODE', ''), row.get('EAN CODE', ''), 
                str(row.get('COLOR', '')).replace('\n', ';'), row.get('DESCRIPTION', ''),
                feature_cols[0], feature_cols[1], feature_cols[2], feature_cols[3], feature_cols[4],
                row.get('IMAGE', ''), row.get('PRICE', ''), 
                '', # 👈 Boş kolon
                row.get('RETAIL PRICE', ''), row.get('NUMBER OF PACKAGES', ''),
                p_lbs, # 👈 WEIGHT (LBS) buraya taşındı
                convert_value(p_x, unit_choice), convert_value(p_y, unit_choice), convert_value(p_z, unit_choice),
                c_lbs,
                convert_value(row.get('PACKAGING SIZE - X (cm)', ''), unit_choice),
                convert_value(row.get('PACKAGING SIZE - Y (cm)', ''), unit_choice),
                convert_value(row.get('PACKAGING SIZE - Z (cm)', ''), unit_choice)
            ]
            processed_data.append(processed_row)

        output_df = pd.DataFrame(processed_data, columns=output_headers)
        st.success(f"İşlem tamamlandı! ({unit_choice} birimi kullanıldı)")
        st.dataframe(output_df)

        # Excel olarak indirme
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 İşlenmiş Excel'i İndir",
            data=output.getvalue(),
            file_name=f"asir_islenmis_{unit_choice}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Bir hata oluştu: {e}")

