import streamlit as st
import tempfile
from pathlib import Path
from engine import convertpdf

st.set_page_config(
    page_title="OFX Bridge Web",
    page_icon="💳",
    layout="wide"
)

st.title("💳 OFX Bridge - Convertisseur PDF → OFX")
st.markdown("**Transforme tes relevés bancaires PDF en OFX pour ta compta**")

# Sidebar
st.sidebar.header("📁 Fichiers PDF")
uploaded_file = st.sidebar.file_uploader("Choisis un PDF bancaire", type="pdf")

st.sidebar.header("⚙️ Logiciel comptable")
target = st.sidebar.selectbox(
    "Sélectionne ton logiciel", 
    ["quadra", "sage", "ebp", "myunisoft"]
)

# Main content
col1, col2 = st.columns([1, 3])

if uploaded_file is not None:
    col1.success(f"📄 **{uploaded_file.name}**")
    
    if st.button("🚀 Convertir en OFX", type="primary"):
        with st.spinner("Conversion en cours..."):
            # Sauvegarde temporaire du fichier
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                pdf_path = tmp_file.name
            
            try:
                # Conversion avec ton moteur
                output_path, nb_txns, info, bank = convertpdf(
                    pdf_path, target=target
                )
                
                # Téléchargement
                with open(output_path, "rb") as ofx_file:
                    st.download_button(
                        label=f"✅ Télécharger {Path(output_path).name}",
                        data=ofx_file.read(),
                        file_name=f"releve_{bank}.ofx",
                        mime="application/x-ofx"
                    )
                
                st.success(f"**{nb_txns} transactions converties**")
                st.info(f"**Banque** : {bank}")
                st.info(f"**IBAN** : {info.get('iban', 'Non détecté')}")
                
            except Exception as e:
                st.error(f"❌ Erreur : {str(e)}")
            
            finally:
                # Nettoyage
                Path(pdf_path).unlink(missing_ok=True)
                try:
                    Path(output_path).unlink()
                except:
                    pass
else:
    col1.info("👆 **Upload un PDF bancaire** dans la sidebar")
    col2.markdown("""
    # 🎯 Comment ça marche ?
    
    1. **Upload** ton PDF Qonto/SG/LCL/etc.
    2. **Choisis** Quadra/Sage/etc.
    3. **Clique Convertir**
    4. **Télécharge** le fichier OFX !
    
    **Formats supportés** : Qonto, LCL, Société Générale, Crédit Agricole...
    """)

st.markdown("---")
st.markdown("*Développé avec ❤️")