import streamlit as st
from pathlib import Path
from engine import convertpdf, parse_cic
import tempfile
import pandas as pd

st.set_page_config(page_title="OFX Bridge", layout="wide")
st.title("💳 OFX Bridge")
st.markdown("**PDF bancaire → OFX pour comptabilité**")

uploaded_file = st.sidebar.file_uploader("📤 PDF bancaire", type="pdf")
target = st.sidebar.selectbox("💻 Logiciel", ["quadra", "myunisoft", "sage", "ebp"])

if uploaded_file:
    st.success(f"✅ **{uploaded_file.name}** ({uploaded_file.size/1024:.0f} KB)")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.getvalue())
        pdf_path = tmp.name

    try:
        info, transactions = parse_cic(pdf_path)

        col1, col2, col3 = st.columns(3)
        col1.metric("🏦 Banque", "CIC")
        col2.metric("📊 Transactions", len(transactions))
        col3.metric("💳 IBAN", f"{info.get('iban', '')[:10]}...")

        st.subheader(f"📋 Aperçu des {len(transactions)} transactions")

        df_data = []
        total_debit = 0.0
        total_credit = 0.0

        for txn in transactions:
            debit = abs(txn["amount"]) if txn["amount"] < 0 else 0.0
            credit = txn["amount"] if txn["amount"] > 0 else 0.0

            df_data.append({
                "Date": txn["date"],
                "Libellé": txn["name"],
                "Mémo": txn["memo"],
                "Débit": f"{debit:,.2f}€".replace(",", " ").replace(".", ",") if debit else "",
                "Crédit": f"{credit:,.2f}€".replace(",", " ").replace(".", ",") if credit else ""
            })

            total_debit += debit
            total_credit += credit

        if df_data:
            st.dataframe(pd.DataFrame(df_data), use_container_width=True, hide_index=True)
        else:
            st.warning("Aucune transaction détectée dans ce PDF.")

        c1, c2, c3 = st.columns(3)
        c1.metric("📉 Débits", f"{total_debit:,.2f}€".replace(",", " ").replace(".", ","))
        c2.metric("📈 Crédits", f"{total_credit:,.2f}€".replace(",", " ").replace(".", ","))
        c3.metric("💰 Solde", f"{(total_credit - total_debit):,.2f}€".replace(",", " ").replace(".", ","))

        if st.button("🚀 Exporter OFX", type="primary", use_container_width=True):
            ofx_path, count, info_out, bank = convertpdf(pdf_path, None, target)
            with open(ofx_path, "r", encoding="latin-1", errors="replace") as f:
                ofx_data = f.read()

            st.download_button(
                label=f"📥 Télécharger releve_{bank}_{target}.ofx",
                data=ofx_data,
                file_name=f"releve_{bank}_{target}_{Path(uploaded_file.name).stem}.ofx",
                mime="application/x-ofx",
                use_container_width=True
            )
            st.success("🎉 OFX exporté avec succès !")

    finally:
        Path(pdf_path).unlink(missing_ok=True)

else:
    st.info("👈 Upload PDF → Aperçu automatique → Exporter")