import pandas as pd
import math
import re
import io
import streamlit as st

# --- Sabitler ---
KG_TO_LBS = 2.20462
CM_TO_INCH = 0.393701
MADE_IN_TURKEY = "Made In Türkiye"

# --- Mevcut Fonksiyonların (extract_dimensions_from_string, clean_feature_list, convert_cm_to_inch) 
# dokunulmadan aynı kaldığını varsayıyoruz. Yukardaki koddan aynen kopyalayabilirsin ---

# ... (Buraya daha önce verdiğim yardımcı fonksiyonları ekle) ...

# --- Streamlit Arayüzü ---
st.set_page_config(page_title="Excel İşleyici", layout="centered")
st.title("📊 Excel Veri Dönüştürücü")

uploaded_file = st.file_uploader("Excel dosyanızı yükleyin", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Dosyayı oku
        df = pd.read_excel(uploaded_file, dtype=str).fillna('')
        
        processed_data = []
        output_headers = [
            'CODE', 'EAN CODE', 'COLOR', 'DESCRIPTION',
            'Feature 1', 'Feature 2', 'Feature 3', 'Feature 4', 'Feature 5',
            'IMAGE', 'PRICE', 'RETAIL PRICE', 'NUMBER OF PACKAGES',
            'PRODUCT SIZE - X (inch)', 'PRODUCT SIZE - Y (inch)', 'PRODUCT SIZE - Z (inch)',
            'CARTON WEIGHT (LBS)', 'WEIGHT (LBS)',
            'PACKAGING SIZE - X (inch)', 'PACKAGING SIZE - Y (inch)', 'PACKAGING SIZE - Z (inch)'
        ]

        for index, row in df.iterrows():
            # 🔥 KRİTİK DÜZELTME: Eğer CODE sütunu boşsa bu satırı atla (21. satır hatasını önler)
            if str(row.get('CODE', '')).strip() == '':
                continue
            
            # --- İşleme Mantığı Başlangıcı ---
            features_str = row.get('FEATURES', '')
            extra_features_str = row.get('EXTRA FEATURES', '')
            combined_text = str(features_str) + "\n" + str(extra_features_str)
            
            # Boyut çıkarma
            dims = extract_dimensions_from_string(combined_text)
            p_x, p_y, p_z = dims if dims else ('', '', '')

            # Renk temizleme
            processed_color = str(row.get('COLOR', '')).replace('\\n', ';').replace('\n', ';')
            processed_color = re.sub(r';+', ';', processed_color).strip(';')

            # Özellik listesi oluşturma
            feat_list = clean_feature_list(features_str)
            if str(extra_features_str).strip() and "number of packages" not in str(extra_features_str).lower():
                feat_list.extend(clean_feature_list(extra_features_str))

            feature_cols = [""] * 5
            for i in range(min(len(feat_list), 4)):
                feature_cols[i] = feat_list[i]
            if len(feat_list) >= 5:
                feature_cols[4] = "\n".join(feat_list[4:])

            # Made in Türkiye ekleme (Sadece liste boşsa veya içinde yoksa)
            if not any(MADE_IN_TURKEY in str(f) for f in feature_cols):
                empty_idx = next((i for i, v in enumerate(feature_cols) if v == ""), None)
                if empty_idx is not None:
                    feature_cols[empty_idx] = MADE_IN_TURKEY
                else:
                    feature_cols[4] += f"\n{MADE_IN_TURKEY}"

            # Ağırlık dönüşümü
            try:
                w_kg = float(str(row.get('WEIGHT (Kg)', '')).replace(',', '.'))
                c_lbs = round(w_kg * KG_TO_LBS, 2)
                p_lbs = max(0.0, round(c_lbs - 0.01, 2))
            except:
                c_lbs = p_lbs = row.get('WEIGHT (Kg)', '')

            # Satırı ekle
            processed_data.append([
                row.get('CODE', ''), row.get('EAN CODE', ''), processed_color, row.get('DESCRIPTION', ''),
                feature_cols[0], feature_cols[1], feature_cols[2], feature_cols[3], feature_cols[4],
                row.get('IMAGE', ''), row.get('PRICE', ''), row.get('RETAIL PRICE', ''), row.get('NUMBER OF PACKAGES', ''),
                convert_cm_to_inch(p_x), convert_cm_to_inch(p_y), convert_cm_to_inch(p_z),
                c_lbs, p_lbs,
                convert_cm_to_inch(row.get('PACKAGING SIZE - X (cm)', '')),
                convert_cm_to_inch(row.get('PACKAGING SIZE - Y (cm)', '')),
                convert_cm_to_inch(row.get('PACKAGING SIZE - Z (cm)', ''))
            ])

        # Yeni DataFrame ve İndirme Butonu
        output_df = pd.DataFrame(processed_data, columns=output_headers)
        st.dataframe(output_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False)
        
        st.download_button(label="📥 İşlenmiş Excel'i İndir", data=output.getvalue(), file_name="islenmis_liste.xlsx")

    except Exception as e:
        st.error(f"Hata: {e}")
