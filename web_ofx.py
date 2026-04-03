import streamlit as st
import tempfile
from pathlib import Path
from engine import convertpdf

st.set_page_config(page_title="OFX Bridge", layout="wide")

st.title("💳 OFX Bridge")
st.markdown("Convertisseur PDF bancaire → OFX")

# Sidebar
st.sidebar.title("📤 Upload")
uploaded_file = st.sidebar.file_uploader("PDF bancaire", type="pdf")
target = st.sidebar.selectbox("Logiciel", ["quadra", "sage", "ebp"])

if uploaded_file is not None:
    st.success(f"Fichier: **{uploaded_file.name}**")
    
    if st.button("🔄 Convertir", type="primary"):
        with st.spinner('Conversion...'):
            # Fichier temporaire
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(uploaded_file.getvalue())
                pdf_path = tmp.name
            
            try:
                output_path, nb_txns, info, bank = convertpdf(pdf_path, target=target)
                
                # Download
                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"📥 {bank}.ofx ({nb_txns} txns)",
                        data=f.read(),
                        file_name=f"releve_{bank}.ofx",
                        mime="text/plain"
                    )
                st.balloons()
                
            except Exception as e:
                st.error(f"❌ {e}")
            finally:
                Path(pdf_path).unlink(missing_ok=True)
else:
    st.info("👈 **Upload PDF dans la sidebar**")