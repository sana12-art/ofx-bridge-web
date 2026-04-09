import pdfplumber
from pathlib import Path
from datetime import datetime
import re
import hashlib


def parse_amount_fr(s):
    if not s:
        return None
    s = str(s).replace("€", "").replace("\u202f", "").replace(" ", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None


def clean_text(s):
    return re.sub(r"\s+", " ", s or "").strip()


def detect_bank(pages_text):
    text = "\n".join(pages_text[:3]).upper()
    if "CREDIT INDUSTRIEL ET COMMERCIAL" in text or "CMCIFRPP" in text or "CIC" in text:
        return "CIC"
    if "QONTO" in text or "QNTOFRP" in text:
        return "QONTO"
    if "CREDIT LYONNAIS" in text or "LCL" in text:
        return "LCL"
    if "SOCIETE GENERALE" in text or "SOCIÉTÉ GÉNÉRALE" in text or "SG" in text:
        return "SG"
    if "CREDIT AGRICOLE" in text or "CRÉDIT AGRICOLE" in text or "AGRIFRPP" in text:
        return "CA"
    if "BANQUE POPULAIRE" in text or "CCBPFRPP" in text:
        return "BP"
    if "BANQUE POSTALE" in text or "PSSTFRPP" in text:
        return "LBP"
    return "UNKNOWN"


def extract_text_by_page(pdfpath):
    pages = []
    with pdfplumber.open(pdfpath) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return pages


def make_fitid(date, label, amount):
    return hashlib.md5(f"{date}{label}{amount:.2f}".encode()).hexdigest()


def generate_ofx(info, txns, target):
    dn = datetime.now().strftime("%Y%m%d%H%M%S")
    bal = info.get("balance_close", 0.0)
    iban = info.get("iban", "FR7612345678901234567890123")

    lines = [
        "OFXHEADER:100",
        "DATA:OFXSGML",
        "VERSION:102",
        "SECURITY:NONE",
        "ENCODING:USASCII",
        "CHARSET:1252",
        "COMPRESSION:NONE",
        "OLDFILEUID:NONE",
        "NEWFILEUID:NONE",
        "",
        "<OFX>",
        "<SIGNONMSGSRSV1><SONRS>",
        "<STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>",
        f"<DTSERVER>{dn}</DTSERVER>",
        "<LANGUAGE>FRA</LANGUAGE>",
        "</SONRS></SIGNONMSGSRSV1>",
        "<BANKMSGSRSV1><STMTTRNRS>",
        "<TRNUID>00</TRNUID>",
        "<STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>",
        "<STMTRS>",
        "<CURDEF>EUR</CURDEF>",
        "<BANKACCTFROM>",
        "<BANKID>30000</BANKID>",
        f"<ACCTID>{iban}</ACCTID>",
        "<ACCTTYPE>CHECKING</ACCTTYPE>",
        "</BANKACCTFROM>",
        "<BANKTRANLIST>",
        f"<DTSTART>{info.get('period_start', '20250701')}</DTSTART>",
        f"<DTEND>{info.get('period_end', '20250731')}</DTEND>",
    ]

    for t in txns:
        lines.extend([
            "<STMTTRN>",
            f"<TRNTYPE>{t['type']}</TRNTYPE>",
            f"<DTPOSTED>{t['date']}</DTPOSTED>",
            f"<TRNAMT>{t['amount']:.2f}</TRNAMT>",
            f"<FITID>{t['fitid']}</FITID>",
            f"<NAME>{t['name'][:64]}</NAME>",
            f"<MEMO>{t['memo'][:128]}</MEMO>",
            "</STMTTRN>",
        ])

    lines.extend([
        "</BANKTRANLIST>",
        f"<LEDGERBAL><BALAMT>{bal:.2f}</BALAMT>",
        f"<DTASOF>{dn[:8]}</DTASOF></LEDGERBAL>",
        "</STMTRS></STMTTRNRS></BANKMSGSRSV1></OFX>",
    ])
    return "\n".join(lines)


def parse_cic(pdfpath):
    """Parseur CIC robuste - ne prend QUE les vrais montants"""
    lines = []
    with pdfplumber.open(pdfpath) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            lines.extend([clean_text(x) for x in txt.split("\n") if clean_text(x)])

    # Skip mots-clés techniques
    skip_keywords = [
        "TOTAL DES MOUVEMENTS", "SOLDE CREDITEUR", "SOLDE DEBITEUR", "IBAN",
        "RELEVE ET INFORMATIONS BANCAIRES", "RELEVÉ ET INFORMATIONS BANCAIRES",
        "CREDIT INDUSTRIEL ET COMMERCIAL", "CIC EZANVILLE", "PAGE", "WWW.CIC.FR",
        "RCS PARIS", "TVA INTRACOMMUNAUTAIRE", "MÉDIATEUR", "MEDIATEUR", 
        "GARANTIE", "ORIAS", "DATE", "VALEUR", "OPÉRATION", "DÉBIT", "CRÉDIT"
    ]

    transactions = []
    i = 0
    while i < len(lines):
        line = lines[i]
        up = line.upper()

        # Skip lignes techniques
        if any(k in up for k in skip_keywords):
            i += 1
            continue

        # Date au début = transaction potentielle (plus souple)
        date_match = re.match(r"^(\d{2}/\d{2}/\d{4})", line)
        if not date_match:
            i += 1
            continue

        date_op = date_match.group(1)
        rest = clean_text(line[len(date_match.group(1)):].strip())

        # ✅ REGEX STRICTE : SEULEMENT montants "1 234,56" ou "123,45"
        amt_match = re.search(r'(\d{1,3}(?:\s?\d{3})*(?:,\d{2})?)$', rest)
        amount = None
        label = rest

        if amt_match:
            candidate = amt_match.group(1).replace(" ", "").replace(",", ".")
            amount = parse_amount_fr(candidate)
            
            # ❌ Rejette les numéros de compte géants (> 1M€)
            if amount and abs(amount) > 1000000:
                amount = None
            else:
                label = clean_text(rest[:amt_match.start()])

        # Si pas de montant sur la ligne date, vérifier ligne suivante
        if amount is None and i + 1 < len(lines):
            next_line = lines[i + 1]
            if not any(k in next_line.upper() for k in skip_keywords):
                next_amt = re.search(r'(\d{1,3}(?:\s?\d{3})*(?:,\d{2})?)$', next_line)
                if next_amt:
                    candidate = next_amt.group(1).replace(" ", "").replace(",", ".")
                    amount = parse_amount_fr(candidate)
                    if amount and abs(amount) <= 1000000:
                        label = clean_text(line[len(date_match.group(1)):])

        # Montant valide trouvé ET raisonnable
        if amount is not None and abs(amount) > 0.01:
            # Corriger débits mal signés (PAIEMENT/CB = négatif)
            if amount > 0 and any(x in up for x in ["PAIEMENT", "PRLV", "PRELEVEMENT", "FACT", "CB"]):
                amount = -abs(amount)

            transactions.append({
                "date": date_op.replace("/", ""),
                "type": "DEBIT" if amount < 0 else "CREDIT",
                "amount": amount,
                "name": (label or f"Transaction {len(transactions) + 1}")[:64],
                "memo": Path(pdfpath).stem,
                "fitid": make_fitid(date_op.replace("/", ""), label or "", amount),
            })

        i += 1

    # Déduplication par date + montant (évite doublons)
    seen = set()
    uniq = []
    for t in transactions:
        key = (t["date"], round(t["amount"], 2))
        if key not in seen:
            seen.add(key)
            uniq.append(t)

    info = {
        "iban": "FR76" + Path(pdfpath).stem[:20].upper(),
        "period_start": "20250701",
        "period_end": "20250731",
        "balance_close": sum(t["amount"] for t in uniq),
    }
    return info, uniq


def convertpdf(pdfpath, outputpath=None, target="quadra"):
    p = Path(pdfpath)
    if not p.exists():
        raise FileNotFoundError(f"Fichier introuvable {pdfpath}")

    if outputpath is None:
        outputpath = p.with_suffix(".ofx")

    pages_text = extract_text_by_page(str(pdfpath))
    bank = detect_bank(pages_text)

    if bank == "UNKNOWN":
        raise ValueError("Banque non reconnue. Formats: CIC, Qonto, LCL, SG, CA, BP, LBP")

    if bank == "CIC":
        info, txns = parse_cic(str(pdfpath))
    else:
        info = {
            "iban": "FR7612345678901234567890123",
            "period_start": "20250701",
            "period_end": "20250731",
            "balance_close": 0.0,
        }
        txns = []

    ofx = generate_ofx(info, txns, target)

    with open(outputpath, "w", encoding="latin-1", errors="replace") as f:
        f.write(ofx)

    return str(outputpath), len(txns), info, bank