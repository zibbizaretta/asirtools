import pandas as pd
import math
import re
import io
import os
import streamlit as st
from openpyxl.styles import Alignment

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

    xyz_pattern = r'(\d+(?:[.,]\d+)?)\s*[xX]\s*(\d+(?:[.,]\d+)?)(?:\s*[xX]\s*(\d+(?:[.,]\d+)?))?'
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
    # Satır sonlarına veya \n ifadesine göre böler
    features = re.split(r'\s*(?:\\n|\n)\s*', str(features_str).strip())
    return [f.strip() for f in features if f and f.strip()]

def convert_size_value(val, unit_choice):
    if val is None or val == '' or (isinstance(val, float) and math.isnan(val)): return ''
    try:
        num_val = float(str(val).replace(',', '.'))
        if unit_choice == "inch":
            return round(num_val * CM_TO_INCH, 2)
        return round(num_val, 2)
    except:
        return val

def convert_weight_value(val_kg, weight_unit_choice):
    if val_kg is None or val_kg == '' or (isinstance(val_kg, float) and math.isnan(val_kg)): return ''
    try:
        num_val = float(str(val_kg).replace(',', '.'))
        if weight_unit_choice == "LBS":
            c_lbs = round(num_val * KG_TO_LBS, 2)
            p_lbs = max(0.0, round(c_lbs - 0.01, 2))
            return c_lbs, p_lbs
        return round(num_val, 2), round(num_val, 2)
    except:
        return val_kg, val_kg

# --- Streamlit Arayüzü ---
st.set_page_config(page_title="Asir Tools", layout="wide")
st.title("📊 Excel Veri Dönüştürücü")

# Sidebar Ayarları
with st.sidebar:
    st.header("⚙️ Ayarlar")
    size_unit = st.radio("Ölçü Birimi (Boyut):", ("cm", "inch"), index=1)
    weight_unit = st.radio("Ağırlık Birimi:", ("KG", "LBS"), index=1)
    
    st.divider()
    st.subheader("Özellik (Feature) Ayarları")
    add_made_in_tr = st.checkbox("Made in Türkiye Eklensin mi?", value=True)
    feature_count = st.slider("Özellik Kolon Sayısı:", min_value=1, max_value=10, value=5)
    
    st.divider()
    if st.button("🏠 Ana Sayfa", use_container_width=True):
        st.write('<meta http-equiv="refresh" content="0;url=https://excelwebpy-asirtools.streamlit.app/">', unsafe_allow_html=True)
        st.stop()

uploaded_file = st.file_uploader("İşlemek istediğiniz Excel dosyasını seçin", type=["xlsx", "xls"])

if uploaded_file:
    try:
        input_filename = uploaded_file.name
        file_base, file_ext = os.path.splitext(input_filename)
        output_filename = f"{file_base}_islenmis{file_ext}"

        df = pd.read_excel(uploaded_file, dtype=str).fillna('')
        processed_data = []
        size_label = f"({size_unit})"
        weight_label = f"({weight_unit})"
        
        # Dinamik Özellik Başlıkları Oluşturma
        feature_headers = [f'Feature {i+1}' for i in range(feature_count)]
        
        output_headers = [
            'CODE', 'EAN CODE', 'COLOR', 'DESCRIPTION'
        ] + feature_headers + [
            'IMAGE', 'PRICE', ' ', 'RETAIL PRICE', 'NUMBER OF PACKAGES', 
            f'WEIGHT {weight_label}',
            f'PRODUCT SIZE - X {size_label}', f'PRODUCT SIZE - Y {size_label}', f'PRODUCT SIZE - Z {size_label}',
            f'CARTON WEIGHT {weight_label}',
            f'PACKAGING SIZE - X {size_label}', f'PACKAGING SIZE - Y {size_label}', f'PACKAGING SIZE - Z {size_label}'
        ]

        for index, row in df.iterrows():
            code_val = str(row.get('CODE', '')).strip()
            if not code_val or code_val.lower() == 'nan':
                continue

            # 1. Özellikleri topla ve temizle
            features_text = str(row.get('FEATURES', ''))
            extra_text = str(row.get('EXTRA FEATURES', ''))
            feat_list = clean_feature_list(features_text)
            if "number of packages" not in extra_text.lower():
                feat_list.extend(clean_feature_list(extra_text))
            
            # 2. Boyutları çıkar (ham metin üzerinden)
            dims = extract_dimensions_from_string(features_text + "\n" + extra_text)
            p_x, p_y, p_z = dims if dims else ('', '', '')

            # 3. Made in Türkiye Ekleme (Seçenek aktifse)
            if add_made_in_tr:
                if not any(MADE_IN_TURKEY in str(f) for f in feat_list):
                    feat_list.append(MADE_IN_TURKEY)

            # 4. Özellikleri Kolonlara Dağıt (Dinamik Mantık)
            feature_cols = [""] * feature_count
            if feature_count > 1:
                # Son kolon hariç her birine bir madde
                for i in range(min(len(feat_list), feature_count - 1)):
                    feature_cols[i] = feat_list[i]
                # Geri kalan her şeyi son kolona ekle
                if len(feat_list) >= feature_count:
                    feature_cols[feature_count-1] = "\n".join(feat_list[feature_count-1:])
            elif feature_count == 1 and feat_list:
                feature_cols[0] = "\n".join(feat_list)

            # 5. Ağırlık Dönüşümü
            c_weight, p_weight = convert_weight_value(row.get('WEIGHT (Kg)', ''), weight_unit)

            # 6. Satırı Oluştur
            processed_row = [
                code_val, row.get('EAN CODE', ''), 
                str(row.get('COLOR', '')).replace('\n', ';'), row.get('DESCRIPTION', '')
            ] + feature_cols + [
                row.get('IMAGE', ''), row.get('PRICE', ''), '', 
                row.get('RETAIL PRICE', ''), row.get('NUMBER OF PACKAGES', ''),
                p_weight,
                convert_size_value(p_x, size_unit), convert_size_value(p_y, size_unit), convert_size_value(p_z, size_unit),
                c_weight,
                convert_size_value(row.get('PACKAGING SIZE - X (cm)', ''), size_unit),
                convert_size_value(row.get('PACKAGING SIZE - Y (cm)', ''), size_unit),
                convert_size_value(row.get('PACKAGING SIZE - Z (cm)', ''), size_unit)
            ]
            processed_data.append(processed_row)

        output_df = pd.DataFrame(processed_data, columns=output_headers)
        st.success(f"✅ İşlem tamamlandı! {len(output_df)} ürün hazır.")
        st.dataframe(output_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            
            wrap_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            wrap_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            for col_idx, column_name in enumerate(output_headers, 1):
                column_letter = worksheet.cell(row=1, column=col_idx).column_letter
                
                # Genişlik Ayarları
                if "Feature" in str(column_name):
                    worksheet.column_dimensions[column_letter].width = 15
                elif any(word in str(column_name) for word in ["PRICE", "SIZE", "WEIGHT", "PACKAGES"]):
                    max_data_len = 0
                    for row_idx in range(2, len(output_df) + 2):
                        val = worksheet.cell(row=row_idx, column=col_idx).value
                        max_data_len = max(max_data_len, len(str(val)) if val else 0)
                    worksheet.column_dimensions[column_letter].width = max_data_len + 5 
                else:
                    max_len = 0
                    for row_idx in range(1, len(output_df) + 2):
                        val = worksheet.cell(row=row_idx, column=col_idx).value
                        max_len = max(max_len, len(str(val)) if val else 0)
                    worksheet.column_dimensions[column_letter].width = min(max_len + 2, 40)

                for row_idx in range(1, len(output_df) + 2):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    if row_idx == 1:
                        cell.alignment = wrap_center
                    else:
                        cell.alignment = wrap_left if "Feature" in str(column_name) else wrap_center
                    
                    if row_idx > 1:
                        worksheet.row_dimensions[row_idx].height = 15

            worksheet.row_dimensions[1].height = 45

        st.download_button(
            label=f"📥 İşlenmiş Excel'i İndir",
            data=output.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Beklenmedik bir hata oluştu: {e}")
else:
    st.info("👋 Hoş geldiniz! Lütfen işlem yapmak istediğiniz Excel dosyasını yukarıdan seçin.")
