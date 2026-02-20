import pandas as pd
import math
import re
import io
import os
import streamlit as st
import pypdf
from openpyxl.styles import Alignment

# --- SABİTLER ---
KG_TO_LBS = 2.20462
CM_TO_INCH = 0.393701
MADE_IN_TURKEY = "Made In Türkiye"

# --- YARDIMCI FONKSİYONLAR (EXCEL DÖNÜŞTÜRÜCÜ) ---
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

def convert_size_value(val, unit_choice):
    if val is None or val == '' or (isinstance(val, float) and math.isnan(val)): return ''
    try:
        num_val = float(str(val).replace(',', '.'))
        return round(num_val * CM_TO_INCH, 2) if unit_choice == "inch" else round(num_val, 2)
    except: return val

def convert_weight_value(val_kg, weight_unit_choice):
    if val_kg is None or val_kg == '' or (isinstance(val_kg, float) and math.isnan(val_kg)): return ''
    try:
        num_val = float(str(val_kg).replace(',', '.'))
        if weight_unit_choice == "LBS":
            c_lbs = round(num_val * KG_TO_LBS, 2)
            return c_lbs, max(0.0, round(c_lbs - 0.01, 2))
        return round(num_val, 2), round(num_val, 2)
    except: return val_kg, val_kg

# --- YARDIMCI FONKSİYONLAR (PO TRACKING - GELİŞTİRİLMİŞ MANTIK) ---
def process_pdfs_advanced(pdf_files):
    all_data = {}
    # PO: CS veya CA ile başlayan 9+ hane
    po_pattern = re.compile(r"((?:CS|CA)\d{9,})")
    # Tracking: Sadece 12 haneli olanlar. 
    # Ama barkod verilerini elemek için "sayı dizisinin önünde harf olmasın" kuralı ekledik (J26... gibi verileri eler)
    trk_pattern = re.compile(r"(?<![a-zA-Z0-9])(\d{12})(?![a-zA-Z0-9])")
    
    for pdf_file in pdf_files:
        try:
            reader = pypdf.PdfReader(pdf_file)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    # PDF'i PO'lara göre bölüyoruz
                    chunks = po_pattern.split(text)
                    for j in range(1, len(chunks), 2):
                        po_number = chunks[j].strip()
                        content_after_po = chunks[j+1]
                        
                        # Bu PO bloğundaki kargo numaralarını bul
                        found_trks = trk_pattern.findall(content_after_po)
                        if found_trks:
                            if po_number not in all_data:
                                all_data[po_number] = set()
                            for t in found_trks:
                                # PDF'in en sağındaki dikey barkod verileri genelde metin içinde 
                                # izole durmaz veya yanında başka kodlar olur. 
                                # Gerçek kargo numarası ise genelde TRK# etiketinden sonra gelir.
                                all_data[po_number].add(t)
        except Exception as e:
            st.error(f"Hata: {pdf_file.name} - {e}")
    
    final_rows = []
    for po in sorted(all_data.keys()):
        # Senin yerel sonucun gibi TRK'ları virgülle birleştiriyoruz
        trks = sorted(list(all_data[po]))
        if trks:
            final_rows.append({"PO": po, "TRK": ", ".join(trks)})
    return pd.DataFrame(final_rows)

# --- STREAMLIT ARAYÜZÜ ---
st.set_page_config(page_title="Asir Tools Pro", layout="wide")

with st.sidebar:
    st.title("🛠️ Asir Tools Pro")
    page = st.radio("Bir araç seçin:", ["Excel Dönüştürücü", "PO Tracking Çıkarıcı"])
    st.divider()
    if st.button("🏠 Ana Sayfa", use_container_width=True):
        st.write('<meta http-equiv="refresh" content="0;url=https://excelwebpy-asirtools.streamlit.app/">', unsafe_allow_html=True)
        st.stop()

# --- SAYFA 1: EXCEL DÖNÜŞTÜRÜCÜ ---
if page == "Excel Dönüştürücü":
    st.header("📊 Excel Veri Dönüştürücü")
    with st.sidebar:
        st.subheader("⚙️ Excel Ayarları")
        size_unit = st.radio("Ölçü Birimi:", ("cm", "inch"), index=1)
        weight_unit = st.radio("Ağırlık Birimi:", ("KG", "LBS"), index=1)
        add_made_in_tr = st.checkbox("Made in Türkiye Eklensin mi?", value=True)
        feature_count = st.slider("Özellik Kolon Sayısı:", 1, 10, 5)

    uploaded_file = st.file_uploader("Excel dosyasını yükleyin", type=["xlsx", "xls"])
    if uploaded_file:
        try:
            file_base, file_ext = os.path.splitext(uploaded_file.name)
            output_filename = f"{file_base}_islenmis{file_ext}"
            df = pd.read_excel(uploaded_file, dtype=str).fillna('')
            processed_data = []
            
            f_headers = [f'Feature {i+1}' for i in range(feature_count)]
            u_s, u_w = f"({size_unit})", f"({weight_unit})"
            output_headers = ['CODE', 'EAN CODE', 'COLOR', 'DESCRIPTION'] + f_headers + \
                             ['IMAGE', 'PRICE', ' ', 'RETAIL PRICE', 'NUMBER OF PACKAGES', f'WEIGHT {u_w}',
                              f'PRODUCT SIZE - X {u_s}', f'PRODUCT SIZE - Y {u_s}', f'PRODUCT SIZE - Z {u_s}',
                              f'CARTON WEIGHT {u_w}', f'PACKAGING SIZE - X {u_s}', f'PACKAGING SIZE - Y {u_s}', f'PACKAGING SIZE - Z {u_s}']

            for index, row in df.iterrows():
                code = str(row.get('CODE', '')).strip()
                if not code or code.lower() == 'nan': continue
                
                feat_list = clean_feature_list(row.get('FEATURES', ''))
                extra = str(row.get('EXTRA FEATURES', ''))
                if "number of packages" not in extra.lower(): feat_list.extend(clean_feature_list(extra))
                if add_made_in_tr and not any(MADE_IN_TURKEY in str(f) for f in feat_list): feat_list.append(MADE_IN_TURKEY)
                
                f_cols = [""] * feature_count
                if feature_count > 1:
                    for i in range(min(len(feat_list), feature_count - 1)): f_cols[i] = feat_list[i]
                    if len(feat_list) >= feature_count: f_cols[feature_count-1] = "\n".join(feat_list[feature_count-1:])
                elif feat_list: f_cols[0] = "\n".join(feat_list)

                dims = extract_dimensions_from_string(str(row.get('FEATURES', '')) + "\n" + extra)
                px, py, pz = dims if dims else ('', '', '')
                c_w, p_w = convert_weight_value(row.get('WEIGHT (Kg)', ''), weight_unit)

                processed_data.append([code, row.get('EAN CODE', ''), str(row.get('COLOR', '')).replace('\n', ';'), row.get('DESCRIPTION', '')] + f_cols + \
                                      [row.get('IMAGE', ''), row.get('PRICE', ''), '', row.get('RETAIL PRICE', ''), row.get('NUMBER OF PACKAGES', ''),
                                       p_w, convert_size_value(px, size_unit), convert_size_value(py, size_unit), convert_size_value(pz, size_unit),
                                       c_w, convert_size_value(row.get('PACKAGING SIZE - X (cm)', ''), size_unit),
                                       convert_size_value(row.get('PACKAGING SIZE - Y (cm)', ''), size_unit), convert_size_value(row.get('PACKAGING SIZE - Z (cm)', ''), size_unit)])

            out_df = pd.DataFrame(processed_data, columns=output_headers)
            st.success(f"✅ Hazır: {output_filename}")
            st.dataframe(out_df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                out_df.to_excel(writer, index=False, sheet_name='Sheet1')
                ws = writer.sheets['Sheet1']
                al_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
                al_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
                for c_idx, c_name in enumerate(output_headers, 1):
                    letter = ws.cell(row=1, column=c_idx).column_letter
                    if "Feature" in str(c_name): ws.column_dimensions[letter].width = 15
                    elif any(w in str(c_name) for w in ["PRICE", "SIZE", "WEIGHT", "PACKAGES"]):
                        m_data = max([len(str(ws.cell(row=r, column=c_idx).value)) for r in range(2, len(out_df)+2)] + [0])
                        ws.column_dimensions[letter].width = m_data + 5
                    else:
                        m_all = max([len(str(ws.cell(row=r, column=c_idx).value)) for r in range(1, len(out_df)+2)] + [0])
                        ws.column_dimensions[letter].width = min(m_all + 2, 40)
                    for r_idx in range(1, len(out_df) + 2):
                        cell = ws.cell(row=r_idx, column=c_idx)
                        cell.alignment = al_left if (r_idx > 1 and "Feature" in str(c_name)) else al_center
                        if r_idx > 1: ws.row_dimensions[r_idx].height = 15
                ws.row_dimensions[1].height = 45
            st.download_button("📥 Excel'i İndir", output.getvalue(), output_filename, use_container_width=True)
        except Exception as e: st.error(f"Hata: {e}")

# --- SAYFA 2: PO TRACKING (GELİŞTİRİLMİŞ) ---
elif page == "PO Tracking Çıkarıcı":
    st.header("📄 PDF'den PO ve Tracking Numarası Çıkarıcı")
    st.write("PDF'leri yükleyin; sistem dikey kodları ve barkod verilerini akıllıca eler.")
    pdf_files = st.file_uploader("PDF dosyalarını seçin", type="pdf", accept_multiple_files=True)
    if pdf_files:
        if st.button("🚀 Verileri Ayrıştır", use_container_width=True):
            with st.spinner("Dosyalar taranıyor..."):
                results_df = process_pdfs_advanced(pdf_files)
                if not results_df.empty:
                    st.success(f"✅ İşlem Tamam! {len(results_df)} PO bulundu.")
                    st.dataframe(results_df)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        results_df.to_excel(writer, index=False)
                    st.download_button("📥 Sonuçları Excel Olarak İndir", output.getvalue(), "PO_Tracking_Final.xlsx", use_container_width=True)
                else:
                    st.warning("Eşleşen geçerli PO veya Tracking numarası bulunamadı.")
