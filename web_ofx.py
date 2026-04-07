import streamlit as st
from datetime import datetime
import fitz
import tempfile
import re
from pathlib import Path
import pandas as pd

st.set_page_config(page_title="OFX Bridge", layout="wide")

st.title("💳 OFX Bridge")
st.markdown("**PDF bancaire → OFX pour comptabilité**")

uploaded_file = st.sidebar.file_uploader("📤 PDF bancaire", type="pdf")
target = st.sidebar.selectbox("💻 Logiciel", ["quadra", "myunisoft", "sage", "ebp"])

def parse_amount_fr(value):
    if not value:
        return None
    s = value.replace("€", "").replace("\u202f", "").replace(" ", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None

def clean_text(value):
    return re.sub(r"\s+", " ", value or "").strip()

if uploaded_file:
    st.success(f"✅ **{uploaded_file.name}** ({uploaded_file.size/1024:.0f} KB)")

    with st.spinner("🔍 Analyse PDF..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.getvalue())
            pdf_path = tmp.name

        doc = fitz.open(pdf_path)
        full_text = ""
        for page in doc:
            full_text += page.get_text() + "\n"
        doc.close()

        text_upper = full_text.upper()

        if "QONTO" in text_upper:
            bank = "QONTO"
        elif "SOCIETE GENERALE" in text_upper or "SOCIÉTÉ GÉNÉRALE" in text_upper:
            bank = "SG"
        elif "LCL" in text_upper:
            bank = "LCL"
        elif "CREDIT AGRICOLE" in text_upper or "CRÉDIT AGRICOLE" in text_upper:
            bank = "CA"
        elif "CREDIT INDUSTRIEL ET COMMERCIAL" in text_upper or "CMCIFRPP" in text_upper or "CIC" in text_upper:
            bank = "CIC"
        elif "BANQUE POSTALE" in text_upper:
            bank = "LBP"
        else:
            bank = Path(uploaded_file.name).stem[:12].upper()

        iban_match = re.search(r'(FR[A-Z0-9]{2}[A-Z0-9]{4}[0-9]{5}[A-Z0-9]{11,})', text_upper)
        iban = iban_match.group(1) if iban_match else f"FR76{Path(uploaded_file.name).stem[:20].upper()}"

        lines = [line.strip() for line in full_text.split("\n") if line.strip()]
        transactions = []

        skip_keywords = [
            "TOTAL DES MOUVEMENTS",
            "SOLDE CREDITEUR",
            "SOLDE DÉBITEUR",
            "SOLDE DEBITEUR",
            "IBAN",
            "RELEVE ET INFORMATIONS BANCAIRES",
            "RELEVÉ ET INFORMATIONS BANCAIRES",
            "CREDIT INDUSTRIEL ET COMMERCIAL",
            "CIC EZANVILLE",
            "PAGE",
            "WWW.CIC.FR",
            "RCS PARIS",
            "TVA INTRACOMMUNAUTAIRE",
            "MÉDIATEUR",
            "MEDIATEUR",
            "GARANTIE",
            "ORIAS",
            "DATE DATE VALEUR OPERATION DEBIT EUROS CREDIT EUROS",
            "DATE DATE VALEUR OPÉRATION DÉBIT EUROS CRÉDIT EUROS"
        ]

        for i, line in enumerate(lines):
            up = line.upper()

            if any(k in up for k in skip_keywords):
                continue

            m = re.match(
                r"^(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(.*?)(\d[\d\s.,]*)$",
                line
            )
            if not m:
                continue

            date_op = m.group(1)
            date_val = m.group(2)
            rest = clean_text(m.group(3))
            amount_raw = m.group(4).strip()
            amount = parse_amount_fr(amount_raw)

            if amount is None:
                continue

            if abs(amount) < 0.01 or abs(amount) > 1000000:
                continue

            label = rest[:80]

            memo = ""
            if i + 1 < len(lines):
                next_line = clean_text(lines[i + 1])
                if next_line and not re.match(r"^\d{2}/\d{2}/\d{4}", next_line):
                    if not any(k in next_line.upper() for k in skip_keywords):
                        memo = next_line[:120]

            label_upper = label.upper()
            if any(x in label_upper for x in ["PAIEMENT", "PRLV", "PRELEVEMENT", "FACT"]):
                signed_amount = -abs(amount)
                txn_type = "DEBIT"
            else:
                signed_amount = amount if amount >= 0 else abs(amount)
                txn_type = "CREDIT" if signed_amount >= 0 else "DEBIT"

            transactions.append({
                "date": date_op.replace("/", ""),
                "date_value": date_val.replace("/", ""),
                "amount": signed_amount,
                "label": label,
                "memo": memo,
                "type": txn_type
            })

        unique = []
        seen = set()
        for t in transactions:
            key = (t["date"], t["label"], round(t["amount"], 2))
            if key not in seen:
                seen.add(key)
                unique.append(t)
        transactions = unique

    col1, col2, col3 = st.columns(3)
    col1.metric("🏦 Banque", bank)
    col2.metric("📊 Transactions", len(transactions))
    col3.metric("💳 IBAN", f"{iban[:10]}...")

    st.subheader(f"📋 **Aperçu des {len(transactions)} transactions**")

    df_data = []
    total_debit = 0.0
    total_credit = 0.0

    for txn in transactions:
        debit = abs(txn["amount"]) if txn["amount"] < 0 else 0
        credit = txn["amount"] if txn["amount"] > 0 else 0

        df_data.append({
            "Date": txn["date"][:8],
            "Libellé": txn["label"],
            "Mémo": txn["memo"],
            "Débit": f"{debit:,.2f}€" if debit else "",
            "Crédit": f"{credit:,.2f}€" if credit else ""
        })

        total_debit += debit
        total_credit += credit

    df = pd.DataFrame(df_data)
    st.dataframe(df, use_container_width=True, hide_index=True)

    col_total1, col_total2, col_total3 = st.columns(3)
    col_total1.metric("📉 Débits", f"{total_debit:,.2f}€")
    col_total2.metric("📈 Crédits", f"{total_credit:,.2f}€")
    col_total3.metric("💰 Solde", f"{total_credit - total_debit:,.2f}€")

    if st.button("🚀 **Exporter OFX**", type="primary", use_container_width=True):
        with st.spinner("📤 Génération OFX..."):
            dn = datetime.now().strftime("%Y%m%d%H%M%S")

            ofx = f"""OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
<SIGNONMSGSRSV1><SONRS>
<STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>
<DTSERVER>{dn}</DTSERVER>
<LANGUAGE>FRA</LANGUAGE>
</SONRS></SIGNONMSGSRSV1>
<BANKMSGSRSV1><STMTTRNRS>
<TRNUID>1001</TRNUID>
<STMTRS>
<CURDEF>EUR</CURDEF>
<BANKACCTFROM>
<BANKID>30000</BANKID>
<ACCTID>{iban}</ACCTID>
<ACCTTYPE>CHECKING</ACCTTYPE>
</BANKACCTFROM>
<BANKTRANLIST>
<DTSTART>20250701</DTSTART>
<DTEND>20250731</DTEND>"""

            for txn in transactions:
                ofx += f"""
<STMTTRN>
<TRNTYPE>{txn["type"]}</TRNTYPE>
<DTPOSTED>{txn["date"]}</DTPOSTED>
<TRNAMT>{txn["amount"]:.2f}</TRNAMT>
<FITID>{bank}{txn["date"]}{abs(txn["amount"]):.2f}</FITID>
<NAME>{txn["label"][:64]}</NAME>
<MEMO>{txn["memo"][:128] if txn["memo"] else Path(uploaded_file.name).stem}</MEMO>
</STMTTRN>"""

            ofx += f"""
</BANKTRANLIST>
<LEDGERBAL>
<BALAMT>{total_credit - total_debit:.2f}</BALAMT>
<DTASOF>20250731</DTASOF>
</LEDGERBAL>
</STMTRS>
</STMTTRNRS>
</BANKMSGSRSV1>
</OFX>"""

            st.download_button(
                label=f"📥 Télécharger **releve_{bank}_{target}.ofx**",
                data=ofx,
                file_name=f"releve_{bank}_{target}_{Path(uploaded_file.name).stem}.ofx",
                mime="application/x-ofx",
                use_container_width=True
            )
            st.success("🎉 **OFX exporté avec succès !**")
            st.balloons()

    Path(pdf_path).unlink(missing_ok=True)

else:
    st.info("👈 **Upload PDF** → **Aperçu automatique** → **Exporter**")