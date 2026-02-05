import pandas as pd
import math
import re
import io
import os # 👈 Dosya uzantısını ayırmak için eklendi
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
    features = re.split(r'\s*(?:\\n|\n)\s*', str(features_str).strip())
    return [f.strip() for f in features if f and f.strip()]

def convert_value(val, unit_choice):
    if val is None or val == '' or (isinstance(val, float) and math.isnan(val)): return ''
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

with st.sidebar:
    st.header("⚙️ Ayarlar")
    unit_choice = st.radio("Ölçü Birimi Seçin:", ("cm", "inch"), index=1)
    st.info(f"Seçili Birim: **{unit_choice.upper()}**")
    st.divider()
    if st.button("🔄 Uygulamayı Sıfırla"):
        st.rerun()

uploaded_file = st.file_uploader("İşlemek istediğiniz Excel dosyasını seçin", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Dinamik Dosya Adı Oluşturma
        input_filename = uploaded_file.name
        file_base, file_ext = os.path.splitext(input_filename)
        output_filename = f"{file_base}_islenmis{file_ext}"

        df = pd.read_excel(uploaded_file, dtype=str).fillna('')
        processed_data = []
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
            code_val = str(row.get('CODE', '')).strip()
            if not code_val or code_val.lower() == 'nan':
                continue

            features_text = str(row.get('FEATURES', ''))
            extra_text = str(row.get('EXTRA FEATURES', ''))
            combined_text = features_text + "\n" + extra_text
            dims = extract_dimensions_from_string(combined_text)
            p_x, p_y, p_z = dims if dims else ('', '', '')

            feat_list = clean_feature_list(features_text)
            if "number of packages" not in extra_text.lower():
                feat_list.extend(clean_feature_list(extra_text))
            
            feature_cols = [""] * 5
            for i in range(min(len(feat_list), 4)):
                feature_cols[i] = feat_list[i]
            if len(feat_list) >= 5:
                feature_cols[4] = "\n".join(feat_list[4:])
            
            if not any(MADE_IN_TURKEY in str(f) for f in feature_cols):
                for i in range(5):
                    if feature_cols[i] == "":
                        feature_cols[i] = MADE_IN_TURKEY
                        break
                else:
                    feature_cols[4] += f"\n{MADE_IN_TURKEY}"

            try:
                weight_input = str(row.get('WEIGHT (Kg)', '')).replace(',', '.')
                w_kg = float(weight_input)
                c_lbs = round(w_kg * KG_TO_LBS, 2)
                p_lbs = max(0.0, round(c_lbs - 0.01, 2))
            except:
                c_lbs = p_lbs = row.get('WEIGHT (Kg)', '')

            processed_row = [
                code_val, row.get('EAN CODE', ''), 
                str(row.get('COLOR', '')).replace('\n', ';'), row.get('DESCRIPTION', ''),
                feature_cols[0], feature_cols[1], feature_cols[2], feature_cols[3], feature_cols[4],
                row.get('IMAGE', ''), row.get('PRICE', ''), '', 
                row.get('RETAIL PRICE', ''), row.get('NUMBER OF PACKAGES', ''),
                p_lbs,
                convert_value(p_x, unit_choice), convert_value(p_y, unit_choice), convert_value(p_z, unit_choice),
                c_lbs,
                convert_value(row.get('PACKAGING SIZE - X (cm)', ''), unit_choice),
                convert_value(row.get('PACKAGING SIZE - Y (cm)', ''), unit_choice),
                convert_value(row.get('PACKAGING SIZE - Z (cm)', ''), unit_choice)
            ]
            processed_data.append(processed_row)

        output_df = pd.DataFrame(processed_data, columns=output_headers)
        st.success(f"İşlem başarıyla tamamlandı! Dosya: {output_filename}")
        st.dataframe(output_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            
            wrap_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            wrap_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            for col_idx, column_name in enumerate(output_headers, 1):
                column_letter = worksheet.cell(row=1, column=col_idx).column_letter
                
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
            label=f"📥 İşlenmiş Excel'i İndir ({unit_choice.upper()})",
            data=output.getvalue(),
            file_name=output_filename, # 👈 Artık orijinal isme göre iniyor
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Beklenmedik bir hata oluştu: {e}")
else:
    st.info("Lütfen başlamak için bir Excel dosyası yükleyin.")
