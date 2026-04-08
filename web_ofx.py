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
    s = str(value).replace("€", "").replace("\u202f", "").replace(" ", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None

def clean_text(value):
    return re.sub(r"\s+", " ", value or "").strip()

def is_skip_line(line):
    u = line.upper()
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
    return any(k in u for k in skip_keywords)

def detect_bank(text_upper, filename):
    if "CREDIT INDUSTRIEL ET COMMERCIAL" in text_upper or "CMCIFRPP" in text_upper or "CIC" in text_upper:
        return "CIC"
    if "QONTO" in text_upper:
        return "QONTO"
    if "SOCIETE GENERALE" in text_upper or "SOCIÉTÉ GÉNÉRALE" in text_upper:
        return "SG"
    if "LCL" in text_upper:
        return "LCL"
    if "CREDIT AGRICOLE" in text_upper or "CRÉDIT AGRICOLE" in text_upper:
        return "CA"
    if "BANQUE POSTALE" in text_upper:
        return "LBP"
    return Path(filename).stem[:12].upper()

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
        bank = detect_bank(text_upper, uploaded_file.name)

        iban_match = re.search(r'(FR[A-Z0-9]{2}[A-Z0-9]{4}[A-Z0-9]{11,})', text_upper)
        iban = iban_match.group(1) if iban_match else f"FR76{Path(uploaded_file.name).stem[:20].upper()}"

        lines = [clean_text(line) for line in full_text.split("\n") if clean_text(line)]
        transactions = []

        # Extraction robuste CIC : une ligne = une écriture quand elle contient une date
        # On accepte plusieurs variantes :
        # - 01/07/2025 01/07/2025 LIBELLE 132,73
        # - 01/07/2025 LIBELLE 132,73
        # - date + libellé + montant sur ligne suivante si besoin
        date_regex = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(.*)$')

        i = 0
        while i < len(lines):
            line = lines[i]
            if is_skip_line(line):
                i += 1
                continue

            m = date_regex.match(line)
            if not m:
                i += 1
                continue

            date_op = m.group(1)
            rest = m.group(2).strip()

            # Cas CIC fréquent : date valeur répétée au début
            rest = re.sub(r'^\d{2}/\d{2}/\d{4}\s+', '', rest).strip()

            # On cherche un montant à la fin de la ligne
            amount = None
            label = rest
            memo = ""

            amt_match = re.search(r'(-?\d[\d\s.,]*\d)$', rest)
            if amt_match:
                amount = parse_amount_fr(amt_match.group(1))
                label = rest[:amt_match.start()].strip()

            # Si pas trouvé sur la même ligne, regarder la ligne suivante
            if amount is None and i + 1 < len(lines):
                next_line = lines[i + 1]
                if not is_skip_line(next_line):
                    next_amt = re.search(r'(-?\d[\d\s.,]*\d)$', next_line)
                    if next_amt:
                        amount = parse_amount_fr(next_amt.group(1))
                        memo = next_line[:next_amt.start()].strip()

            # Si toujours rien, essayer sur 2 lignes suivantes
            if amount is None and i + 2 < len(lines):
                next_line = lines[i + 1]
                next_next_line = lines[i + 2]
                if not is_skip_line(next_line):
                    next_amt = re.search(r'(-?\d[\d\s.,]*\d)$', next_next_line)
                    if next_amt:
                        amount = parse_amount_fr(next_amt.group(1))
                        memo = next_line[:120]
                        if not label:
                            label = next_line[:80]

            if amount is not None:
                label = clean_text(label)
                memo = clean_text(memo)

                if not label:
                    label = f"Transaction {len(transactions) + 1}"

                # Heuristique simple : certains relevés CIC utilisent des montants positifs
                # On garde le signe tel quel s'il existe, sinon on classe selon le libellé.
                if amount > 0 and any(x in label.upper() for x in ["PAIEMENT", "PRLV", "PRELEVEMENT", "FACT"]):
                    amount = -abs(amount)

                txn_type = "DEBIT" if amount < 0 else "CREDIT"

                transactions.append({
                    "date": date_op.replace("/", ""),
                    "date_display": date_op,
                    "amount": amount,
                    "label": label[:80],
                    "memo": memo[:120],
                    "type": txn_type
                })

            i += 1

        # Suppression des doublons
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
        debit = abs(txn["amount"]) if txn["amount"] < 0 else 0.0
        credit = txn["amount"] if txn["amount"] > 0 else 0.0

        df_data.append({
            "Date": txn["date_display"],
            "Libellé": txn["label"],
            "Mémo": txn["memo"],
            "Débit": f"{debit:,.2f}€".replace(",", " ").replace(".", ",") if debit else "",
            "Crédit": f"{credit:,.2f}€".replace(",", " ").replace(".", ",") if credit else ""
        })

        total_debit += debit
        total_credit += credit

    st.dataframe(pd.DataFrame(df_data), use_container_width=True, hide_index=True)

    col_total1, col_total2, col_total3 = st.columns(3)
    col_total1.metric("📉 Débits", f"{total_debit:,.2f}€".replace(",", " ").replace(".", ","))
    col_total2.metric("📈 Crédits", f"{total_credit:,.2f}€".replace(",", " ").replace(".", ","))
    col_total3.metric("💰 Solde", f"{(total_credit - total_debit):,.2f}€".replace(",", " ").replace(".", ","))

    if st.button("🚀 **Exporter OFX**", type="primary", use_container_width=True):
        with st.spinner("📤 Génération OFX..."):
            dn = datetime.now().strftime("%Y%m%d%H%M%S")
            dtstart = "20250701"
            dtend = "20250731"

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
<DTSTART>{dtstart}</DTSTART>
<DTEND>{dtend}</DTEND>"""

            for txn in transactions:
                ofx += f"""
<STMTTRN>
<TRNTYPE>{txn["type"]}</TRNTYPE>
<DTPOSTED>{txn["date"]}</DTPOSTED>
<TRNAMT>{txn["amount"]:.2f}</TRNAMT>
<FITID>{bank}{txn["date"]}{abs(txn["amount"]):.2f}</FITID>
<NAME>{txn["label"][:64]}</NAME>
<MEMO>{(txn["memo"] if txn["memo"] else Path(uploaded_file.name).stem)[:128]}</MEMO>
</STMTTRN>"""

            ofx += f"""
</BANKTRANLIST>
<LEDGERBAL>
<BALAMT>{(total_credit - total_debit):.2f}</BALAMT>
<DTASOF>{dtend}</DTASOF>
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