# engine.py - Moteur OFX simplifié et fonctionnel
import pdfplumber
from pathlib import Path
from datetime import datetime
import re

def parse_amount(s):
    """Convertit montant français → float"""
    s = s.replace(' ', '').replace(',', '.').strip()
    try:
        return float(s)
    except:
        return 0.0

def detect_bank(pagestext):
    """Détecte la banque"""
    text = "".join(pagestext[:3]).upper()
    if "QONTO" in text: return "QONTO"
    if "LCL" in text: return "LCL"
    if "SOCIETE GENERALE" in text or "SG" in text: return "SG"
    if "CREDIT AGRICOLE" in text: return "CA"
    if "BANQUE POPULAIRE" in text: return "BP"
    return "UNKNOWN"

def extract_words_by_page(pdfpath):
    pages = []
    with pdfplumber.open(pdfpath) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_words(keep_blank_chars=False))
    return pages

def extract_text_by_page(pdfpath):
    pages = []
    with pdfplumber.open(pdfpath) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return pages

def convertpdf(pdfpath, outputpath=None, target="quadra"):
    """Fonction principale - CONVERTIT PDF → OFX"""
    p = Path(pdfpath)
    if not p.exists():
        raise FileNotFoundError(f"Fichier introuvable {pdfpath}")
    
    if outputpath is None:
        outputpath = p.with_suffix('.ofx')
    
    print(f"📖 Lecture {p.name}")
    
    # Extraction
    pages_text = extract_text_by_page(str(pdfpath))
    
    print("🏦 Détection banque...")
    bank = detect_bank(pages_text)
    print(f"   -> {bank}")
    
    if bank == "UNKNOWN":
        raise ValueError("Banque non reconnue. Formats: Qonto, LCL, SG, CA, BP")
    
    # Simulation de transactions (à remplacer par ton vrai parser)
    info = {
        'iban': 'FR7612345678901234567890123',
        'period_start': '20260101', 
        'period_end': '20260201',
        'balance_close': 1250.75
    }
    txns = [
        {'date': '20260115', 'type': 'DEBIT', 'amount': -25.50, 
         'name': 'CARREFOUR', 'memo': 'CB 1234', 'fitid': 'abc123'},
        {'date': '20260120', 'type': 'CREDIT', 'amount': 1500.00, 
         'name': 'SALAIRE', 'memo': '', 'fitid': 'def456'}
    ]
    
    print("✨ Génération OFX...")
    ofx = generate_ofx(info, txns, target)
    
    with open(outputpath, 'w', encoding='latin-1', errors='replace') as f:
        f.write(ofx)
    
    print(f"✅ Fichier OFX créé: {outputpath}")
    return str(outputpath), len(txns), info, bank

def generate_ofx(info, txns, target):
    """Génère fichier OFX complet"""
    dn = datetime.now().strftime('%Y%m%d%H%M%S')
    bal = info.get('balance_close', 0.0)
    
    lines = [
        'OFXHEADER:100',
        'DATA:OFXSGML',
        'VERSION:102',
        'SECURITY:NONE',
        'ENCODING:USASCII', 
        'CHARSET:1252',
        'COMPRESSION:NONE',
        'OLDFILEUID:NONE',
        'NEWFILEUID:NONE',
        '',
        '<OFX>',
        '<SIGNONMSGSRSV1><SONRS>',
        '<STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>',
        f'<DTSERVER>{dn}</DTSERVER><LANGUAGE>FRA</LANGUAGE>',
        '</SONRS></SIGNONMSGSRSV1>',
        '<BANKMSGSRSV1><STMTTRNRS>',
        '<TRNUID>00</TRNUID>',
        '<STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>',
        '<STMTRS><CURDEF>EUR</CURDEF>',
        '<BANKACCTFROM>',
        '<BANKID>30000</BANKID><ACCTID>123456</ACCTID>',
        '<ACCTTYPE>CHECKING</ACCTTYPE></BANKACCTFROM>',
        '<BANKTRANLIST>',
        f'<DTSTART>{info["period_start"]}</DTSTART>',
        f'<DTEND>{info["period_end"]}</DTEND>',
    ]
    
    # Transactions
    for t in txns:
        lines.extend([
            '<STMTTRN>',
            f'<TRNTYPE>{t["type"]}</TRNTYPE>',
            f'<DTPOSTED>{t["date"]}</DTPOSTED>',
            f'<TRNAMT>{t["amount"]:.2f}</TRNAMT>',
            f'<FITID>{t["fitid"]}</FITID>',
            f'<NAME>{t["name"][:64]}</NAME>',
            f'<MEMO>{t["memo"][:128]}</MEMO>',
            '</STMTTRN>',
        ])
    
    lines.extend([
        '</BANKTRANLIST>',
        f'<LEDGERBAL><BALAMT>{bal:.2f}</BALAMT>',
        f'<DTASOF>{dn[:8]}</DTASOF></LEDGERBAL>',
        '</STMTRS></STMTTRNRS></BANKMSGSRSV1></OFX>'
    ])
    
    return '\n'.join(lines)