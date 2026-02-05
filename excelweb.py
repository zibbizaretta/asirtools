import pandas as pd
import math
import re
import io
import streamlit as st

# --- Sabitler ---
KG_TO_LBS = 2.20462
CM_TO_INCH = 0.393701
MADE_IN_TURKEY = "Made In Türkiye"

# --- Fonksiyonlar (Mevcut mantığını koruyoruz) ---
def extract_dimensions_from_string(text_to_search):
    if not isinstance(text_to_search, str):
        return None
    
    def find_dimension_value(pattern, text):
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            try:
                val_str = match.group(1).replace(',', '.')
                return float(val_str)
            except (ValueError, IndexError):
                return None
        return None

    w = find_dimension_value(r"Width[:\s-]*([\d,.]+)", text_to_search)
    h = find_dimension_value(r"Height[:\s-]*([\d,.]+)", text_to_search)
    d = find_dimension_value(r"Depth[:\s-]*([\d,.]+)", text_to_search)
    l = find_dimension_value(r"Length[:\s-]*([\d,.]+)", text_to_search)
    dia = find_dimension_value(r"Diameter[:\s-]*([\d,.]+)", text_to_search)

    depth_val = d if d is not None else l

    if w is not None and h is not None and depth_val is not None:
        return (w, depth_val, h)
    if dia is not None and h is not None:
        return (dia, dia, h)

    xyz_pattern = r"(\d+(?:[.,]\d+)?)\s*[xX*]\s*(\d+(?:[.,]\d+)?)(?:\s*[xX*]\s*(\d+(?:[.,]\d+)?))?"
    match = re.search(xyz_pattern, text_to_search)
    if match:
        try:
            x = float(match.group(1).replace(',', '.'))
            y = float(match.group(2).replace(',', '.'))
            z = float(match.group(3).replace(',', '.')) if match.group(3) else None
            return (x, y, z)
        except (ValueError, TypeError):
            return None
    return None

# --- Streamlit Arayüzü ---
st.set_page_config(page_title="Excel İşleyici", layout="centered")

st.title("📊 Excel Veri Dönüştürücü")
st.write("Excel dosyanızı yükleyin, hesaplamaları yapalım ve işlenmiş halini indirin.")

uploaded_file = st.file_uploader("Bir Excel dosyası seçin", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Excel'i oku
        df = pd.read_excel(uploaded_file)
        st.success("Dosya başarıyla yüklendi! İşleniyor...")

        # --- İşleme Mantığı (Senin kodun) ---
        # Örnek: Eğer kodunda özel sütun işlemleri varsa buraya ekleyebilirsin.
        # Mevcut excel.py içeriğindeki dönüşüm mantığını buraya uyguluyoruz:
        
        # (Burada df üzerinde yaptığın tüm transformasyonları yapabilirsin)
        # Örnek sütun oluşturma:
        if 'Dimensions' in df.columns:
            df['Parsed_Dims'] = df['Dimensions'].apply(extract_dimensions_from_string)
        
        # İşlenmiş veriyi göster (ilk 5 satır)
        st.write("Önizleme (İlk 5 Satır):")
        st.dataframe(df.head())

        # Excel'i belleğe (memory) yazdır (dosya olarak indirmek için)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sonuç')
        
        processed_data = output.getvalue()

        # İndirme Butonu
        st.download_button(
            label="📥 İşlenmiş Dosyayı İndir",
            data=processed_data,
            file_name=f"islenmis_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Bir hata oluştu: {e}")

else:
    st.info("Lütfen işlem yapmak için bir Excel dosyası yükleyin.")