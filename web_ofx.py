import streamlit as st
from datetime import datetime
import fitz
import tempfile
import re
from pathlib import Path

st.set_page_config(page_title="OFX Bridge", layout="wide")

st.title("💳 OFX Bridge Pro")
st.markdown("**PDF bancaire → Aperçu → OFX**")

uploaded_file = st.sidebar.file_uploader("📤 PDF bancaire", type="pdf")
target = st.sidebar.selectbox("💻 Logiciel", ["quadra", "myunisoft", "sage", "ebp"])

if uploaded_file:
    st.success(f"✅ **{uploaded_file.name}** ({uploaded_file.size/1024:.0f} KB)")
    
    with st.spinner("🔍 Analyse intelligente..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.getvalue())
            pdf_path = tmp.name
        
        doc = fitz.open(pdf_path)
        full_text = ""
        for page in doc:
            full_text += page.get_text()
        doc.close()
        
        # DÉTECTION BANQUE
        text_upper = full_text.upper()
        if "QONTO" in text_upper: bank = "QONTO"
        elif "SOCIETE GENERALE" in text_upper: bank = "SG"  
        elif "LCL" in text_upper: bank = "LCL"
        elif "CREDIT AGRICOLE" in text_upper: bank = "CA"
        else: bank = Path(uploaded_file.name).stem[:12].upper()
        
        # EXTRACTION IBAN
        iban_match = re.search(r'(FR[A-Z0-9]{2}[A-Z0-9]{4}[0-9]{5}([A-Z0-9]?){11})', full_text)
        iban = iban_match.group(1)[:34] if iban_match else "FR7612345678901234567890"
        
        # EXTRACTION TRANSACTIONS INTELLIGENTE
        lines = full_text.split('\n')
        transactions = []
        
        for i, line in enumerate(lines):
            date_match = re.search(r'(\d{2}[./-]\d{2})', line)
            amount_match = re.search(r'[-+]?\s*(\d{1,3}[.,]\d{2})', line)
            
            if date_match and amount_match:
                date = date_match.group(1).replace('/', '').replace('-', '').replace('.', '')
                amount_str = amount_match.group(1).replace(' ', '').replace(',', '.')
                
                try:
                    amount = float(amount_str)
                    # FILTRAGE INTELLIGENT : montants raisonnables (0.01-5000€)
                    if 0.01 < abs(amount) < 5000:
                        label = (lines[i+1][:50] if i+1 < len(lines) else 
                                lines[i][:50] if len(lines[i]) > 10 else f"TXN {len(transactions)+1}").strip()
                        
                        transactions.append({
                            'date': f'202607{date}',
                            'amount': amount,
                            'label': label[:50],
                            'type': 'CREDIT' if amount >= 0 else 'DEBIT'
                        })
                except:
                    continue
        
        transactions = transactions[:30]  # Max 30 transactions
        
    # PRÉVISUALISATION
    col1, col2, col3 = st.columns(3)
    col1.metric("🏦 Banque", bank)
    col2.metric("📊 Transactions", len(transactions))
    col3.metric("💳 IBAN", f"{iban[:10]}...")
    
    st.subheader(f"📋 **Aperçu des {len(transactions)} transactions**")
    
    df_data = []
    total_debit = total_credit = 0
    for txn in transactions:
        df_data.append({
            'Date': txn['date'][6:],
            'Montant': f"{txn['amount']:,.2f}€",
            'Libellé': txn['label'],
            'Type': txn['type']
        })
        if txn['amount'] < 0:
            total_debit += abs(txn['amount'])
        else:
            total_credit += txn['amount']
    
    st.dataframe(df_data, use_container_width=True, hide_index=True)
    
    col_total1, col_total2, col_total3 = st.columns(3)
    col_total1.metric("📉 Débits", f"-{total_debit:,.2f}€")
    col_total2.metric("📈 Crédits", f"+{total_credit:,.2f}€")
    col_total3.metric("💰 Solde", f"{total_credit-total_debit:,.2f}€")
    
    # EXPORT
    if st.button("🚀 **Exporter OFX**", type="primary", use_container_width=True):
        dn = datetime.now().strftime('%Y%m%d%H%M%S')
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
<SIGNONMSGSRSV1><SONRS><STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS><DTSERVER>{dn}</DTSERVER><LANGUAGE>FRA</LANGUAGE></SONRS></SIGNONMSGSRSV1>
<BANKMSGSRSV1><STMTTRNRS><TRNUID>1001</TRNUID><STMTRS><CURDEF>EUR</CURDEF><BANKACCTFROM><BANKID>30000</BANKID><ACCTID>{iban}</ACCTID><ACCTTYPE>CHECKING</ACCTTYPE></BANKACCTFROM><BANKTRANLIST><DTSTART>20260701</DTSTART><DTEND>20260731</DTEND>"""
        
        for txn in transactions:
            ofx += f"<STMTTRN><TRNTYPE>{txn['type']}</TRNTYPE><DTPOSTED>{txn['date']}</DTPOSTED><TRNAMT>{txn['amount']:.2f}</TRNAMT><FITID>{bank}{txn['date']}</FITID><NAME>{txn['label'][:64]}</NAME><MEMO>{Path(uploaded_file.name).stem}</MEMO></STMTTRN>"
        
        ofx += f"</BANKTRANLIST><LEDGERBAL><BALAMT>{total_credit-total_debit:.2f}</BALAMT><DTASOF>20260731</DTASOF></LEDGERBAL></STMTRS></STMTTRNRS></BANKMSGSRSV1></OFX>"
        
        st.download_button(
            label=f"📥 **releve_{bank}_{target}.ofx** ({len(transactions)} txn)",
            data=ofx,
            file_name=f"releve_{bank}_{target}_{Path(uploaded_file.name).stem}.ofx",
            mime="application/x-ofx"
        )
        st.balloons()
    
    Path(pdf_path).unlink(missing_ok=True)

else:
    st.info("👈 **Upload PDF** → **Aperçu transactions** → **Export**")