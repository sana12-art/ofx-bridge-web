# -*- coding: utf-8 -*-
"""
OFX Bridge - Convertisseur PDF Qonto vers OFX
Interface graphique moderne (Tkinter)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import re
import hashlib
from datetime import datetime
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "pdfplumber"])
    import pdfplumber


# ══════════════════════════════════════════════════════════════
# UTILITAIRES COMMUNS
# ══════════════════════════════════════════════════════════════

def extract_words_by_page(pdf_path):
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_words(keep_blank_chars=False))
    return pages

def extract_text_by_page(pdf_path):
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return pages

def parse_amount(s):
    """Convertit un montant français en float: '2.870,45' ou '1 234,56' ou '870,45'"""
    s = s.replace('\xa0','').replace(' ','').replace('*','').strip()
    # Format X.XXX,XX (point milliers, virgule décimale)
    if re.match(r'^\d{1,3}(\.\d{3})*,\d{2}$', s):
        return float(s.replace('.','').replace(',','.'))
    # Format X XXX,XX (espace milliers)
    if re.match(r'^\d+,\d{2}$', s):
        return float(s.replace(',','.'))
    # Format décimal point
    if re.match(r'^\d+\.\d{2}$', s):
        return float(s)
    # Fallback: supprimer tout sauf chiffres et dernière virgule/point
    cleaned = re.sub(r'[^\d,.]', '', s)
    cleaned = cleaned.replace(',', '.')
    try:
        return float(cleaned)
    except ValueError:
        return None

def group_words_by_row(words, tol=3.0):
    if not words:
        return []
    rows, cur, top = [], [words[0]], words[0]['top']
    for w in words[1:]:
        if abs(w['top'] - top) <= tol:
            cur.append(w)
        else:
            rows.append(sorted(cur, key=lambda x: x['x0']))
            cur, top = [w], w['top']
    if cur:
        rows.append(sorted(cur, key=lambda x: x['x0']))
    return sorted(rows, key=lambda r: r[0]['top'])

def clean_label(s):
    return re.sub(r'\s+', ' ', s).strip()

def _is_technical_label(label):
    """
    Retourne True si le label ressemble à un code technique peu lisible :
    - Pattern CB carte banque: DDMMYY CB****XXXX CODE  (ex: "011224 CB****6410 FII73LG")
    - Ligne entièrement composée de codes sans mots lisibles (3+ lettres consécutives)
    """
    if not label:
        return True
    # Pattern carte bancaire banque populaire/CA: "DDMMYY CB****XXXX CODE"
    if re.match(r'^\d{6}\s+CB\*+\d+\s+\w+\s*$', label):
        return True
    # Aucun mot avec 3+ lettres alphabétiques consécutives
    if not re.search(r'[A-Za-zÀ-ÿ]{3,}', label):
        return True
    return False

def _is_human_readable(label):
    """
    Retourne True si le label contient un nom de fournisseur lisible :
    - Au moins 2 mots avec des lettres
    - Ne commence pas par un code technique long (15+ chars alphanumériques)
    - Ne contient pas de référence ultra-longue
    """
    if not label:
        return False
    # Rejeter les lignes de références techniques longues
    if re.search(r'[A-Z0-9]{15,}', label):
        return False
    # Rejeter les lignes qui ne sont que des chiffres/codes
    if re.match(r'^[\d\s\-\/.,]+$', label):
        return False
    # Doit contenir au moins 2 tokens avec des lettres lisibles
    readable_words = [w for w in label.split() if re.search(r'[A-Za-zÀ-ÿ]{2,}', w)
                      and not re.match(r'^\d', w)]
    return len(readable_words) >= 2

def smart_label(main_label, memo_lines):
    """
    Choisit le meilleur nom pour une transaction.
    Si le label principal est un code technique et qu'une ligne de mémo
    contient un nom lisible, on préfère ce nom.
    Le label principal original est conservé dans le mémo.
    Retourne (name, memo_str).
    """
    label = clean_label(main_label)
    memos = [clean_label(m) for m in memo_lines if clean_label(m)]

    if _is_technical_label(label) and memos:
        # Chercher la première ligne de mémo lisible
        for candidate in memos:
            if _is_human_readable(candidate):
                # Utiliser cette ligne comme nom, garder l'original en mémo
                remaining = ' | '.join(
                    m for m in memos if m != candidate and m
                )
                return candidate, (label + (' | ' + remaining if remaining else ''))
        # Aucune ligne lisible trouvée: garder l'original
        return label, ' | '.join(memos)

    # Label principal déjà lisible: le garder, mémo = autres lignes
    return label, ' | '.join(memos)

def make_fitid(date, label, amount):
    return hashlib.md5(f"{date}{label}{amount:.2f}".encode()).hexdigest()

def date_jjmm_to_ofx(jjmm, year):
    p = jjmm.replace('.', '/').split('/')
    if len(p) == 2:
        return f"{year}{p[1].zfill(2)}{p[0].zfill(2)}"
    return f"{year}0101"

def date_full_to_ofx(date_str):
    date_str = date_str.replace('.', '/')
    p = date_str.split('/')
    if len(p) == 3:
        return f"{p[2]}{p[1].zfill(2)}{p[0].zfill(2)}"
    return datetime.now().strftime('%Y%m%d')

def extract_iban(text):
    m = re.search(r'IBAN\s*:?\s*((?:[A-Z]{2}\d{2}[\s\d]+))', text)
    if m:
        return re.sub(r'\s+', '', m.group(1)).strip()
    return ''

def iban_to_rib(iban):
    c = iban.replace(' ', '').upper()
    if c.startswith('FR') and len(c) == 27:
        r = c[4:]
        return r[0:5], r[5:10], r[10:21]
    return '99999', '00001', c[-11:] if len(c) >= 11 else c

def _year_from_text(text):
    m = re.search(r'\b(20\d{2})\b', text)
    return int(m.group(1)) if m else datetime.now().year

def _parse_col_amount(words):
    """Parse un montant depuis une liste de mots.
    Supporte: 2.870,45 (milliers) | 870,45 (simple) | 870.45 (point décimal)
    """
    if not words:
        return None
    full = ' '.join(w['text'] for w in words).replace('\xa0', ' ').strip()
    if full in ('.', ',', ''):
        return None
    # Chercher d'abord les montants avec séparateur de milliers: 1.234,56 ou 1 234,56
    m = re.search(r'(\d{1,3}(?:[.\s]\d{3})+,\d{2})', full)
    if m:
        val = parse_amount(m.group(1).replace(' ', '.'))
        if val is not None and val > 0:
            return val
    # Montants simples: 123,45
    m2 = re.search(r'(\d+,\d{2})', full)
    if m2:
        val = parse_amount(m2.group(1))
        if val is not None and val > 0:
            return val
    return None

def _parse_signed_amount(words):
    if not words:
        return None
    full = ' '.join(w['text'] for w in words).replace('\xa0', ' ').strip()
    m = re.search(r'([+\-])\s*([\d\s]+[,.][\d]{2})', full)
    if m:
        sign = 1.0 if m.group(1) == '+' else -1.0
        val = parse_amount(m.group(2))
        if val is not None:
            return sign * val
    m2 = re.search(r'([\d\s]+[,.][\d]{2})', full)
    if m2:
        val = parse_amount(m2.group(1))
        if val is not None:
            return val
    return None

def _make_txn(date_ofx, amount, label, memo=''):
    return {
        'date':   date_ofx,
        'type':   'CREDIT' if amount >= 0 else 'DEBIT',
        'amount': amount,
        'name':   clean_label(label)[:64],
        'memo':   clean_label(memo)[:128],
        'fitid':  make_fitid(date_ofx, label, amount)
    }


# ══════════════════════════════════════════════════════════════
# DETECTION DE LA BANQUE
# ══════════════════════════════════════════════════════════════

def detect_bank(pages_text):
    text = pages_text[0][:3000].upper()
    if 'QONTO' in text or 'QNTOFRP' in text:
        return 'QONTO'
    if 'CREDIT LYONNAIS' in text or ('LCL' in text and 'RELEVE DE COMPTE COURANT' in text):
        return 'LCL'
    # Société Générale — avant Crédit Agricole pour éviter faux positif
    text_nospace = text.replace(' ', '')
    if ('SOCIETE GENERALE' in text or 'SOCIÉTÉ GÉNÉRALE' in text
            or '552 120 222' in text or 'SOCIETEGENERALE' in text_nospace
            or 'SG.FR' in text or 'PROFESSIONNELS.SG.FR' in text):
        return 'SG'
    if 'CREDIT AGRICOLE' in text or 'AGRIFRPP' in text:
        return 'CA'
    # CGD (Caixa Geral de Depositos) — avant Caisse d'Épargne
    if 'CAIXA GERAL' in text or 'CGDIFRPP' in text or 'CGD' in text[:500]:
        return 'CGD'
    if "CAISSE D'EPARGNE" in text or "CAISSE D.EPARGNE" in text or 'CEPAFRPP' in text:
        return 'CE'
    if 'BANQUE POPULAIRE' in text or 'CCBPFRPP' in text:
        return 'BP'
    # La Banque Postale — après Banque Populaire
    if 'BANQUE POSTALE' in text or 'PSSTFRPP' in text or 'LABANQUEPOSTALE' in text:
        return 'LBP'
    if 'CREDIT INDUSTRIEL' in text or 'CMCIFRPP' in text or ('CIC' in text and 'RELEVE' in text):
        return 'CIC'
    if ('BNP PARIBAS' in text or 'BNPAFRPP' in text or 'BNP' in text[:500]
            or 'BANQUE NATIONALE DE PARIS' in text):
        return 'BNP'
    if 'MYPOS' in text or 'MYPOS LTD' in text or 'MPOS99' in text or 'MY POS' in text:
        return 'MYPOS'
    return 'UNKNOWN'


# ══════════════════════════════════════════════════════════════
# PARSEUR QONTO
# ══════════════════════════════════════════════════════════════

def parse_qonto(pages_words, pages_text):
    info = _extract_qonto_header(pages_text[0])
    year = int(info['period_start'].split('/')[2]) if info.get('period_start') else _year_from_text(pages_text[0])
    txns = []
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _qonto_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 130 <= w['x0'] < 410).strip()
            amount = _qonto_amount(row)
            memo = ''
            j = i + 1
            while j < len(rows) and not _qonto_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 130 <= w['x0'] < 410).strip()
                na = _qonto_amount(rows[j])
                if na is not None and amount is None:
                    amount = na; memo = nl; j += 1; break
                elif na is None and nl:
                    memo = nl; j += 1; break
                else:
                    break
            i = j
            if amount is None or not label or label in ('Transactions', 'Date de valeur'):
                continue
            memo_clean = memo if memo.strip() not in ('', '-', '+') else ''
            name, memo_out = smart_label(label, [memo_clean] if memo_clean else [])
            txns.append(_make_txn(date_jjmm_to_ofx(date_str, year), amount, name, memo_out))
    return info, txns

def _qonto_date(row):
    for w in row:
        if w['x0'] < 120 and re.match(r'^\d{2}/\d{2}$', w['text']):
            return w['text']
    return ''

def _qonto_amount(row):
    aw = [w for w in row if w['x0'] >= 400]
    if not aw: return None
    full = ' '.join(w['text'] for w in aw).replace('EUR','').replace('\xa0',' ').strip()
    m = re.search(r'([+\-])\s*([\d\s]+[.,]\d{2})', full)
    if m:
        sign = 1.0 if m.group(1)=='+' else -1.0
        try: return sign * float(m.group(2).replace(' ','').replace(',','.'))
        except: pass
    m2 = re.search(r'([\d\s]+[.,]\d{2})', full)
    if m2:
        sign = 1.0
        for w in aw:
            if w['text'] in ('+','-'): sign = 1.0 if w['text']=='+' else -1.0; break
            sm = re.match(r'^([+\-])([\d,.]+)$', w['text'])
            if sm: sign = 1.0 if sm.group(1)=='+' else -1.0; break
        try: return sign * float(m2.group(1).replace(' ','').replace(',','.'))
        except: pass
    return None

def _extract_qonto_header(text):
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    m = re.search(r'Du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text)
    if m: info['period_start'], info['period_end'] = m.group(1), m.group(2)
    bals = re.findall(r'Solde au \d{2}/\d{2}\s*[+\-]\s*([\d]+\.[\d]{2})\s*EUR', text)
    if len(bals) >= 1: info['balance_open']  = float(bals[0])
    if len(bals) >= 2: info['balance_close'] = float(bals[-1])
    return info


# ══════════════════════════════════════════════════════════════
# PARSEUR LCL
# Colonnes: DATE x~42 | LIBELLE x~197 | VALEUR x~365 | DEBIT x~433 | CREDIT x~504
# ══════════════════════════════════════════════════════════════

def parse_lcl(pages_words, pages_text):
    info = _extract_lcl_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _lcl_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 70 <= w['x0'] < 360).strip()
            # Exclure les dates valeur (format JJ.MM.AA) de la zone débit
            debit_words  = [w for w in row if 360 <= w['x0'] < 490
                            and not re.match(r'^\d{2}\.\d{2}(\.\d{2,4})?$', w['text'])]
            debit_amt  = _parse_col_amount(debit_words)
            credit_amt = _parse_col_amount([w for w in row if w['x0'] >= 490])
            memo = ''
            j = i + 1
            while j < len(rows) and not _lcl_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 70 <= w['x0'] < 360).strip()
                if nl and nl not in ('DEBIT','CREDIT','VALEUR','DATE','LIBELLE','ANCIEN SOLDE'):
                    memo = (memo + ' ' + nl).strip()
                j += 1
            i = j
            if not label or label in ('DEBIT','CREDIT','VALEUR','DATE','LIBELLE','ANCIEN SOLDE'):
                continue
            date_ofx = date_jjmm_to_ofx(date_str, year)
            memo_parts = [memo] if memo else []
            name, memo_out = smart_label(label, memo_parts)
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo_out))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo_out))
    return info, txns

def _lcl_date(row):
    for w in row:
        if w['x0'] < 100 and re.match(r'^\d{2}\.\d{2}$', w['text']):
            return w['text']
    return ''

def _extract_lcl_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    m = re.search(r'du\s+(\d{2}\.\d{2}\.\d{4})\s+au\s+(\d{2}\.\d{2}\.\d{4})', text, re.IGNORECASE)
    if m:
        info['period_start'] = m.group(1).replace('.','/')
        info['period_end']   = m.group(2).replace('.','/') 
    m2 = re.search(r'ANCIEN SOLDE\s+([\d\s]+[,.][\d]{2})', text)
    if m2: info['balance_open'] = parse_amount(m2.group(1)) or 0.0
    m3 = re.search(r'SOLDE EN EUROS\s+([\d\s]+[,.][\d]{2})', text)
    if m3: info['balance_close'] = parse_amount(m3.group(1)) or 0.0
    return info


# ══════════════════════════════════════════════════════════════
# PARSEUR CRÉDIT AGRICOLE
# Colonnes exactes (mesurées sur PDF réel) :
#   Date opé  x ~19  | Date val x ~50  | Libellé x ~79..419
#   Débit     dernier token x ~430-448  (préfixe entier à ~420-432)
#   Crédit    dernier token x ~490-530  (préfixe entier à ~490-515)
# Les gros montants sont splittés en 2 tokens: '1' + '231,49' = 1 231,49
# ══════════════════════════════════════════════════════════════

def _ca_parse_zone(row, x_min, x_max):
    """
    Extrait un montant depuis la zone [x_min, x_max[ d'une ligne.
    Gère les montants splittes en milliers: token entier + token 'XXX,XX'.
    Retourne float ou None.
    """
    # Filtrer uniquement les tokens numériques (ignorer '¨', 'þ', etc.)
    col = [w for w in row if x_min <= w['x0'] < x_max
           and re.match(r'^\d', w['text'])]
    if not col:
        return None
    last = col[-1]['text']
    # Le dernier token doit être un montant décimal (centimes) : "231,49" ou "19,90"
    if not re.match(r'^\d+,\d{2}$', last):
        return None
    if len(col) == 1:
        return parse_amount(last)
    # Tokens précédents = préfixe milliers (entiers purs)
    prefix_tokens = [w['text'] for w in col[:-1]]
    if all(re.match(r'^\d+$', p) for p in prefix_tokens):
        prefix = ''.join(prefix_tokens)
        decimal_str = last.replace(',', '.')
        try:
            return float(prefix + decimal_str)
        except ValueError:
            pass
    # Fallback: concaténer tout
    combined = ''.join(w['text'] for w in col).replace(',', '.')
    try:
        return float(combined)
    except ValueError:
        return None

def parse_ca(pages_words, pages_text):
    info = _extract_ca_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    SKIP = {'Débit','Crédit','Date','Libellé','Total des opérations',
            'Nouveau solde','opé.','valeur','Libellé des opérations',
            'Ancien solde débiteur','Nouveau solde débiteur'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _ca_date(row)
            if not date_str:
                i += 1; continue
            # Libellé: entre les deux dates (x~79) et avant la zone montants (x<420)
            label = ' '.join(w['text'] for w in row if 70 <= w['x0'] < 420).strip()
            # Débit: dernier token autour de x=430-448, préfixe possible à x=420-432
            debit_amt  = _ca_parse_zone(row, 415, 490)
            # Crédit: dernier token autour de x=490-530, préfixe possible à x=490-515
            credit_amt = _ca_parse_zone(row, 490, 560)
            # Mémo: lignes suivantes sans date
            memo_parts = []
            j = i + 1
            while j < len(rows) and not _ca_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 70 <= w['x0'] < 420).strip()
                if not nl or nl in SKIP or len(nl) <= 1:
                    pass
                elif re.match(r'^Page\s+\d+\s*/\s*\d+$', nl):
                    pass  # "Page 1 / 5"
                elif any(k in nl for k in ('Crédit Agricole Brie Picardie', 'RCS AMIENS',
                                           'Vos réserves', 'garantiedesdepots')):
                    pass  # pied de page
                else:
                    memo_parts.append(nl)
                j += 1
            i = j
            if not label or any(s in label for s in ('Total des','Nouveau solde','Ancien solde','Vos réserves')):
                continue
            if label in SKIP:
                continue
            date_ofx = date_jjmm_to_ofx(date_str, year)
            name, memo = smart_label(label, memo_parts)
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, txns

def _ca_date(row):
    for w in row:
        if w['x0'] < 50 and re.match(r'^\d{2}\.\d{2}$', w['text']):
            return w['text']
    return ''

def _extract_ca_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    mois_map = {'janvier':'01','février':'02','mars':'03','avril':'04','mai':'05','juin':'06',
                'juillet':'07','août':'08','septembre':'09','octobre':'10','novembre':'11','décembre':'12'}
    m = re.search(r'Date d.arrêté\s*:\s*(\d+)\s+(\w+)\s+(\d{4})', text)
    if m:
        mn = mois_map.get(m.group(2).lower(), '01')
        info['period_end']   = f"{m.group(1).zfill(2)}/{mn}/{m.group(3)}"
        info['period_start'] = f"01/{mn}/{m.group(3)}"
    m2 = re.search(r'Ancien solde\s+\w+\s+au[^\d]+([\d\s]+[,.][\d]{2})', text, re.IGNORECASE)
    if m2: info['balance_open'] = parse_amount(m2.group(1)) or 0.0
    m3 = re.search(r'Nouveau solde\s+\w+\s+au[^\d]+([\d\s]+[,.][\d]{2})', text, re.IGNORECASE)
    if m3: info['balance_close'] = parse_amount(m3.group(1)) or 0.0
    return info


# ══════════════════════════════════════════════════════════════
# PARSEUR CAISSE D'ÉPARGNE
# Colonnes: DATE OPÉRATION x~56 | DATE VALEUR x~120 | LIBELLÉ x~183 | MONTANT x~511 (signé)
# ══════════════════════════════════════════════════════════════

def parse_ce(pages_words, pages_text):
    info = _extract_ce_header(pages_text)
    txns = []
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _ce_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 155 <= w['x0'] < 500).strip()
            amount = _parse_signed_amount([w for w in row if w['x0'] >= 500])
            memo = ''
            j = i + 1
            while j < len(rows) and not _ce_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 155 <= w['x0'] < 500).strip()
                if nl and len(nl) > 2:
                    memo = (memo + ' ' + nl).strip()
                j += 1
            i = j
            if not label or amount is None:
                continue
            skip_kw = {'DATE','VALEUR','MONTANT','OPERATIONS','SOLDE','TOTAL','DETAIL'}
            if any(s in label.upper() for s in skip_kw):
                continue
            date_ofx = date_full_to_ofx(date_str)
            memo_parts = [memo] if memo else []
            name, memo_out = smart_label(label, memo_parts)
            txns.append(_make_txn(date_ofx, amount, name, memo_out))
    return info, txns

def _ce_date(row):
    for w in row:
        if w['x0'] < 100 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']):
            return w['text']
    return ''

def _extract_ce_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    m = re.search(r'au\s+(\d{2}/\d{2}/\d{4})', text)
    if m:
        info['period_end'] = m.group(1)
        p = m.group(1).split('/')
        info['period_start'] = f"01/{p[1]}/{p[2]}"
    all_soldes = re.findall(r'SOLDE CREDITEUR AU[^\d]+([\d\s]+[,.][\d]{2})', text)
    if all_soldes:
        info['balance_open']  = parse_amount(all_soldes[0])  or 0.0
        info['balance_close'] = parse_amount(all_soldes[-1]) or 0.0
    return info


# ══════════════════════════════════════════════════════════════
# PARSEUR BANQUE POPULAIRE
# Colonnes: DATE COMPTA x~51 | LIBELLE x~94..350 | MONTANT signé x~490+
# Structure montant: '-' token + 'XX,XX' token + '€' token
# ATTENTION: la page "DETAIL DE VOS MOUVEMENTS SEPA" duplique certaines
# transactions SANS le signe '-' → il faut ignorer cette section.
# Règle: on n'accepte QUE les transactions avec un montant SIGNÉ ('-')
# car toutes les opérations BP sont des débits sur ce type de relevé
# (les crédits éventuels ont un '+' ou montant sans signe en colonne séparée).
# ══════════════════════════════════════════════════════════════

def parse_bp(pages_words, pages_text):
    info = _extract_bp_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    # Mots-clés indiquant une section à ignorer (SEPA detail, publicité…)
    SECTION_SKIP = {'DETAIL DE VOS MOUVEMENTS SEPA', 'DETAIL DE VOS PRELEVEMENTS SEPA',
                    'VOTRE COMPTE COURANT'}
    SKIP_KW = {'DATE','LIBELLE','REFERENCE','COMPTA','VALEUR','MONTANT',
               'SOLDE','TOTAL','DETAIL','OPERATION'}

    for pw in pages_words:
        rows = group_words_by_row(pw)
        # Détecter les zones à ignorer (sections secondaires)
        skip_from = None
        for idx, row in enumerate(rows):
            row_text = ' '.join(w['text'] for w in row).upper()
            if any(s in row_text for s in ('DETAIL DE VOS MOUVEMENTS SEPA',
                                           'DETAIL DE VOS PRELEVEMENTS SEPA RECUS')):
                skip_from = idx
                break

        i = 0
        while i < len(rows):
            # Arrêter si on entre dans la section SEPA secondaire
            if skip_from is not None and i >= skip_from:
                break
            row = rows[i]
            date_str = _bp_date(row)
            if not date_str:
                i += 1; continue

            label = ' '.join(w['text'] for w in row if 80 <= w['x0'] < 355).strip()
            amount = _bp_amount([w for w in row if w['x0'] >= 490])

            # Mémo: lignes suivantes sans date
            memo_parts = []
            j = i + 1
            while j < len(rows) and not _bp_date(rows[j]):
                if skip_from is not None and j >= skip_from:
                    break
                nl = ' '.join(w['text'] for w in rows[j] if 80 <= w['x0'] < 355).strip()
                # Exclure lignes de taux de change, totaux et mentions légales
                if not nl or len(nl) <= 2:
                    pass
                elif re.match(r'^[\d\s.,€%=\-EUR]+$', nl):
                    pass
                elif re.search(r'\d+EUR\s+1\s+EURO\s*=', nl):
                    pass  # taux de change "88,79EUR 1 EURO = 1,000000"
                elif any(k in nl.upper() for k in ('TOTAL DES MOUVEMENTS','SOLDE CREDITEUR',
                                                    'SOUS RESERVE','NE JUSTIFIE PAS',
                                                    'BANQUE POPULAIRE','SOCIETE ANONYME',
                                                    "D'ENREGISTREMENT", 'PROVISION SUFFISANTE',
                                                    'DEDUCTION DE LA TVA')):
                    pass  # lignes de totaux et mentions légales
                else:
                    memo_parts.append(nl)
                j += 1
            i = j
            if not label or amount is None:
                continue
            if any(s in label.upper() for s in SKIP_KW):
                continue
            date_ofx = date_jjmm_to_ofx(date_str, year)
            name, memo = smart_label(label, memo_parts)
            txns.append(_make_txn(date_ofx, amount, name, memo))
    return info, txns

def _bp_date(row):
    for w in row:
        if w['x0'] < 80 and re.match(r'^\d{2}/\d{2}$', w['text']):
            return w['text']
    return ''

def _bp_amount(words):
    """
    Parse le montant BP: token '-' suivi de 'XX,XX' (débit)
    ou '+' suivi de 'XX,XX' (crédit).
    Les gros montants sont splittés: '-' '1' '000,00' → -1000.00
    Dans le format BP Rives de Paris, les débits ont un '-' explicite et
    les crédits (remises chèques, virements entrants) n'ont PAS de signe.
    Un montant sans signe est donc traité comme un crédit (positif).
    """
    if not words: return None
    full = ' '.join(w['text'] for w in words).replace('€','').replace('\xa0',' ').strip()

    # Débit: '-' explicite suivi du montant (éventuellement en milliers splittés)
    m = re.search(r'-\s*([\d\s]+[,.][\d]{2})', full)
    if m:
        val_str = m.group(1).replace(' ', '').replace(',', '.')
        try: return -abs(float(val_str))
        except: pass

    # Crédit: '+' explicite suivi du montant
    m2 = re.search(r'\+\s*([\d\s]+[,.][\d]{2})', full)
    if m2:
        val_str = m2.group(1).replace(' ', '').replace(',', '.')
        try: return abs(float(val_str))
        except: pass

    # Pas de signe → crédit (montant positif) dans le format BP
    # Les doublons SEPA sont déjà évités par le skip_from sur la section SEPA
    m3 = re.search(r'([\d\s]+[,.][\d]{2})', full)
    if m3:
        val_str = m3.group(1).replace(' ', '').replace(',', '.')
        try:
            val = float(val_str)
            if val > 0:
                return val
        except: pass

    return None

def _extract_bp_header(pages_text):
    # Utiliser uniquement la page 1 pour éviter de mélanger avec la section SEPA
    text = pages_text[0] if pages_text else ''
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    # Date de fin du relevé: "au 31/12/2024"
    m = re.search(r'au\s+(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
    if m:
        info['period_end'] = m.group(1)
        p = m.group(1).split('/')
        info['period_start'] = f"01/{p[1]}/{p[2]}"
    # Solde ouverture et clôture: "SOLDE CREDITEUR AU DD/MM/YYYY NNN,NN €"
    all_soldes = re.findall(
        r'SOLDE CREDITEUR AU[^\n]+?(\b\d{1,3}(?:\s\d{3})*,\d{2})\s*€', text)
    if all_soldes:
        info['balance_open']  = parse_amount(all_soldes[0])  or 0.0
        info['balance_close'] = parse_amount(all_soldes[-1]) or 0.0
    return info


# ══════════════════════════════════════════════════════════════
# PARSEUR CIC
# Colonnes: Date x~52 | Date valeur x~100 | Opération x~148 | Débit x~435-500 | Crédit x~500+
# ══════════════════════════════════════════════════════════════

def parse_cic(pages_words, pages_text):
    info = _extract_cic_header(pages_text)
    txns = []
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _cic_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 140 <= w['x0'] < 430).strip()
            debit_amt  = _parse_col_amount([w for w in row if 420 <= w['x0'] < 500])
            credit_amt = _parse_col_amount([w for w in row if w['x0'] >= 500])
            memo = ''
            j = i + 1
            while j < len(rows) and not _cic_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 140 <= w['x0'] < 430).strip()
                if nl and len(nl) > 2 and not re.match(r'^[\d.,]+$', nl):
                    memo = (memo + ' ' + nl).strip()
                j += 1
            i = j
            if not label:
                continue
            skip_kw = {'DATE','DÉBIT','CRÉDIT','EUROS','SOLDE CREDITEUR','CREDIT INDUSTRIEL','TOTAL DES MOUVEMENTS'}
            if any(s in label.upper() for s in skip_kw):
                continue
            date_ofx = date_full_to_ofx(date_str)
            memo_parts = [memo] if memo else []
            name, memo_out = smart_label(label, memo_parts)
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo_out))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo_out))
    return info, txns

def _cic_date(row):
    for w in row:
        if w['x0'] < 100 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']):
            return w['text']
    return ''

def _extract_cic_header(pages_text):
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    # IBAN souvent en dernière page
    for pt in reversed(pages_text):
        iban = extract_iban(pt)
        if iban:
            info['iban'] = iban; break
    text = pages_text[0]
    mois_map = {'janvier':'01','février':'02','mars':'03','avril':'04','mai':'05','juin':'06',
                'juillet':'07','août':'08','septembre':'09','octobre':'10','novembre':'11','décembre':'12'}
    m = re.search(r'(\d{2})\s+(janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)\s+(\d{4})', text, re.IGNORECASE)
    if m:
        mn = mois_map.get(m.group(2).lower(), '01')
        info['period_end']   = f"{m.group(1).zfill(2)}/{mn}/{m.group(3)}"
        info['period_start'] = f"01/{mn}/{m.group(3)}"
    all_text = ' '.join(pages_text)
    all_soldes = re.findall(r'SOLDE CREDITEUR AU[^\d]+([\d.]+[,.][\d]{2})', all_text)
    if all_soldes:
        info['balance_open']  = parse_amount(all_soldes[0])  or 0.0
        info['balance_close'] = parse_amount(all_soldes[-1]) or 0.0
    return info



# ══════════════════════════════════════════════════════════════
# PARSEUR CGD (Caixa Geral de Depositos)
# Colonnes: Date DD MM x~24,42 | Libelle x~77..310 | Debit x~395-500 | Credit x~500+
# ══════════════════════════════════════════════════════════════

def parse_cgd(pages_words, pages_text):
    info = _extract_cgd_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    SKIP = {'A REPORTER', 'REPORT', 'TOTAL', 'NOUVEAU', 'ANCIEN', 'SARL', 'CPT ORD'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            if not (len(row) >= 2
                    and re.match(r'^\d{2}$', row[0]['text']) and row[0]['x0'] < 50
                    and re.match(r'^\d{2}$', row[1]['text']) and row[1]['x0'] < 55):
                i += 1; continue
            dd, mm = row[0]['text'], row[1]['text']
            label = ' '.join(w['text'] for w in row if 70 <= w['x0'] < 310).strip()
            if not label or any(s in label.upper() for s in SKIP):
                i += 1; continue
            debit_amt  = _cgd_amount_in_zone(row, 395, 500)
            credit_amt = _cgd_amount_in_zone(row, 500, 570)
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if (len(r2) >= 2 and re.match(r'^\d{2}$', r2[0]['text']) and r2[0]['x0'] < 50):
                    break
                nl = ' '.join(w['text'] for w in r2 if 70 <= w['x0'] < 310).strip()
                if nl and not any(s in nl.upper() for s in
                                  ('A REPORTER', 'REPORTER', 'PAGE', 'TOTAL', 'NOUVEAU',
                                   'CAIXA GERAL', 'SARL MAC', 'GARANTIE', 'MEDIATEUR')):
                    memo_parts.append(nl)
                j += 1
            i = j
            date_ofx = f"{year}{mm.zfill(2)}{dd.zfill(2)}"
            name, memo = smart_label(label, memo_parts)
            if debit_amt:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, txns

def _cgd_amount_in_zone(row, x_min, x_max):
    col = [w for w in row if x_min <= w['x0'] < x_max and re.match(r'^\d', w['text'])]
    if not col: return None
    return parse_amount(col[-1]['text'])

def _extract_cgd_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban': '', 'period_start': '', 'period_end': '', 'balance_open': 0.0, 'balance_close': 0.0}
    info['iban'] = extract_iban(text)
    m = re.search(r'(\d{2}/\d{2}/\d{4})\s+AU\s+(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
    if m:
        info['period_start'] = m.group(1)
        info['period_end']   = m.group(2)
    m2 = re.search(r'ANCIEN SOLDE[^\d]*(\d[\d.]*,\d{2})', text)
    if m2: info['balance_open'] = parse_amount(m2.group(1)) or 0.0
    m3 = re.search(r'NOUVEAU SOLDE EN EUR\s+([+\-]?\d[\d.,]*)', text)
    if m3:
        try: info['balance_close'] = float(m3.group(1).replace('.','').replace(',','.').replace('+',''))
        except: pass
    return info


# ══════════════════════════════════════════════════════════════
# PARSEUR LA BANQUE POSTALE (LBP)
# Colonnes: Date DD/MM x~53 | Operation x~90..430 | Debit x~430-500 | Credit x~500+
# ══════════════════════════════════════════════════════════════

def parse_lbp(pages_words, pages_text):
    info = _extract_lbp_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    SKIP = {'TOTAL DES', 'NOUVEAU SOLDE', 'ANCIEN SOLDE', 'VOS OPERATIONS',
            'DATE OPERATION', 'SITUATION DU', 'PAGE'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            if not (row[0]['x0'] < 60 and re.match(r'^\d{2}/\d{2}$', row[0]['text'])):
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 85 <= w['x0'] < 430).strip()
            label = re.sub(r'\(cid:\d+\)', '', label).strip()
            if not label or any(s in label.upper() for s in SKIP):
                i += 1; continue
            debit_amt  = _lbp_amount_in_zone(row, 430, 500)
            credit_amt = _lbp_amount_in_zone(row, 500, 560)
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2[0]['x0'] < 60 and re.match(r'^\d{2}/\d{2}$', r2[0]['text']): break
                nl = ' '.join(w['text'] for w in r2 if 85 <= w['x0'] < 430).strip()
                nl = re.sub(r'\(cid:\d+\)', '', nl).strip()
                if nl and len(nl) > 2 and not any(s in nl.upper() for s in
                                                   ('TOTAL DES', 'NOUVEAU SOLDE', 'PAGE',
                                                    'LA BANQUE POSTALE', 'FRAIS ET COTIS',
                                                    'IL VOUS EST')):
                    memo_parts.append(nl)
                j += 1
            i = j
            date_ofx = f"{year}{row[0]['text'][3:5]}{row[0]['text'][:2]}"
            name, memo = smart_label(label, memo_parts)
            if debit_amt:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, txns

def _lbp_amount_in_zone(row, x_min, x_max):
    col = [w for w in row if x_min <= w['x0'] < x_max and re.match(r'^\d', w['text'])]
    if not col: return None
    last = col[-1]['text']
    if not re.match(r'^\d+,\d{2}$', last): return None
    if len(col) == 1: return parse_amount(last)
    prefix_tokens = [w['text'] for w in col[:-1]]
    if all(re.match(r'^\d+$', p) for p in prefix_tokens):
        try: return float(''.join(prefix_tokens) + last.replace(',', '.'))
        except: pass
    return parse_amount(last)

def _extract_lbp_header(pages_text):
    text = ' '.join(pages_text[:2])
    text = re.sub(r'\(cid:\d+\)', ' ', text)
    info = {'iban': '', 'period_start': '', 'period_end': '', 'balance_open': 0.0, 'balance_close': 0.0}
    info['iban'] = extract_iban(text)
    m = re.search(r'du\s+(\d{1,2})\s+au\s+(\d{1,2})\s+(\w+)\s+(\d{4})', text, re.IGNORECASE)
    if m:
        mois_map = {'janvier':'01','fevrier':'02','fevrier':'02','mars':'03','avril':'04',
                    'mai':'05','juin':'06','juillet':'07','aout':'08','septembre':'09',
                    'octobre':'10','novembre':'11','decembre':'12'}
        mn = mois_map.get(m.group(3).lower(), '01')
        yr = m.group(4)
        info['period_start'] = f"01/{mn}/{yr}"
        info['period_end']   = f"{m.group(2).zfill(2)}/{mn}/{yr}"
    text_up = text.upper()
    m2 = re.search(r'ANCIEN SOLDE AU[^\d]+(\d[\d\s]*[,.]\d{2})', text_up)
    if m2: info['balance_open'] = parse_amount(m2.group(1)) or 0.0
    m3 = re.search(r'NOUVEAU SOLDE AU[^\d]+(\d[\d\s]*[,.]\d{2})', text_up)
    if m3: info['balance_close'] = parse_amount(m3.group(1)) or 0.0
    return info


# ══════════════════════════════════════════════════════════════
# PARSEUR SOCIETE GENERALE (SG)
# Date DD/MM/YYYY x~31 | Valeur x~77 | Libelle x~124 | Debit x~430-510 | Credit x~510+
# Montants avec '*' (exonere de commission) comptent quand meme.
# ══════════════════════════════════════════════════════════════

def parse_sg(pages_words, pages_text):
    info = _extract_sg_header(pages_text)
    txns = []
    SKIP = {'TOTAUX DES', 'NOUVEAU SOLDE', 'SOLDE PRECEDENT', 'PROGRAMME DE',
            'RAPPEL DES', 'MONTANT CUMULE'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            if not (row[0]['x0'] < 45 and re.match(r'^\d{2}/\d{2}/\d{4}$', row[0]['text'])):
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 120 <= w['x0'] < 430).strip()
            if not label or any(s in label.upper() for s in SKIP):
                i += 1; continue
            debit_amt  = _sg_amount_in_zone(row, 430, 510)
            credit_amt = _sg_amount_in_zone(row, 510, 570)
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2[0]['x0'] < 45 and re.match(r'^\d{2}/\d{2}/\d{4}$', r2[0]['text']): break
                nl = ' '.join(w['text'] for w in r2 if 120 <= w['x0'] < 430).strip()
                if nl and not any(s in nl.upper() for s in
                                  ('TOTAUX', 'NOUVEAU', 'PROGRAMME', 'RAPPEL',
                                   'SOCIETE GENERALE', 'SUITE >>>', 'PAGE')):
                    memo_parts.append(nl)
                j += 1
            i = j
            date_ofx = date_full_to_ofx(row[0]['text'])
            name, memo = smart_label(label, memo_parts)
            if debit_amt:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, txns

def _sg_amount_in_zone(row, x_min, x_max):
    col = [w for w in row if x_min <= w['x0'] < x_max]
    if not col: return None
    return parse_amount(col[-1]['text'])

def _extract_sg_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban': '', 'period_start': '', 'period_end': '', 'balance_open': 0.0, 'balance_close': 0.0}
    info['iban'] = extract_iban(text)
    m = re.search(r'du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
    if m:
        info['period_start'] = m.group(1)
        info['period_end']   = m.group(2)
    m2 = re.search(r'SOLDE\s+PR[EE]C[EE]DENT\s+AU[^\d]+(\d[\d.,]*)', text, re.IGNORECASE)
    if m2: info['balance_open'] = parse_amount(m2.group(1)) or 0.0
    m3 = re.search(r'NOUVEAU SOLDE AU[^\d]+[+]?\s*(\d[\d.,]*)', text, re.IGNORECASE)
    if m3: info['balance_close'] = parse_amount(m3.group(1)) or 0.0
    return info



# ══════════════════════════════════════════════════════════════
# PARSEUR BNP PARIBAS
# Colonnes: Date DD/MM/YY x~30 | Libellé x~90..430 | Valeur x~430-500 | Débit x~500-560 | Crédit x~560+
# Le relevé BNP utilise des dates JJ/MM/AA (2 chiffres pour l'année)
# ══════════════════════════════════════════════════════════════

def parse_bnp(pages_words, pages_text):
    info = _extract_bnp_header(pages_text)
    year = _year_from_text(' '.join(pages_text[:2]))
    txns = []
    SKIP = {'DATE', 'LIBELLE', 'VALEUR', 'DEBIT', 'CREDIT', 'EUROS',
            'SOLDE', 'TOTAL', 'OPERATIONS', 'ANCIEN SOLDE', 'NOUVEAU SOLDE',
            'VIREMENT RECU', 'RELEVE DE COMPTE'}

    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _bnp_date(row)
            if not date_str:
                i += 1; continue

            label = ' '.join(w['text'] for w in row if 85 <= w['x0'] < 430).strip()
            if not label or any(s in label.upper() for s in SKIP):
                i += 1; continue

            # BNP: débit en zone 490-560, crédit en zone 560+
            debit_amt  = _parse_col_amount([w for w in row if 480 <= w['x0'] < 560])
            credit_amt = _parse_col_amount([w for w in row if w['x0'] >= 560])

            # Mémo: lignes suivantes sans date
            memo_parts = []
            j = i + 1
            while j < len(rows) and not _bnp_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 85 <= w['x0'] < 430).strip()
                if nl and len(nl) > 2 and not any(s in nl.upper() for s in
                                                   ('TOTAL', 'SOLDE', 'BNP PARIBAS', 'PAGE')):
                    memo_parts.append(nl)
                j += 1
            i = j

            date_ofx = _bnp_date_to_ofx(date_str, year)
            name, memo = smart_label(label, memo_parts)
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, txns

def _bnp_date(row):
    """Détecte une date BNP: JJ/MM/AA ou JJ/MM/AAAA en début de ligne."""
    for w in row:
        if w['x0'] < 80:
            if re.match(r'^\d{2}/\d{2}/\d{2}$', w['text']):
                return w['text']
            if re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']):
                return w['text']
    return ''

def _bnp_date_to_ofx(date_str, year_hint):
    """Convertit JJ/MM/AA ou JJ/MM/AAAA en AAAAMMJJ."""
    parts = date_str.split('/')
    if len(parts) == 3:
        dd, mm = parts[0].zfill(2), parts[1].zfill(2)
        yy = parts[2]
        if len(yy) == 2:
            # Siècle: 00-30 → 20xx, 31-99 → 19xx
            full_year = (2000 + int(yy)) if int(yy) <= 30 else (1900 + int(yy))
        else:
            full_year = int(yy)
        return f"{full_year}{mm}{dd}"
    return str(year_hint) + '0101'

def _extract_bnp_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban': '', 'period_start': '', 'period_end': '', 'balance_open': 0.0, 'balance_close': 0.0}
    info['iban'] = extract_iban(text)
    # Période: "du JJ/MM/AAAA au JJ/MM/AAAA" ou "du JJ/MM/AA au JJ/MM/AA"
    m = re.search(r'du\s+(\d{2}/\d{2}/\d{2,4})\s+au\s+(\d{2}/\d{2}/\d{2,4})', text, re.IGNORECASE)
    if m:
        info['period_start'] = m.group(1).replace('/', '/')
        info['period_end']   = m.group(2).replace('/', '/')
    # Soldes
    m2 = re.search(r'(?:ANCIEN SOLDE|SOLDE PRECEDENT)[^\d]+([\d\s]+[,.][\d]{2})', text, re.IGNORECASE)
    if m2: info['balance_open'] = parse_amount(m2.group(1)) or 0.0
    m3 = re.search(r'(?:NOUVEAU SOLDE|SOLDE FINAL)[^\d]+([\d\s]+[,.][\d]{2})', text, re.IGNORECASE)
    if m3: info['balance_close'] = parse_amount(m3.group(1)) or 0.0
    return info


# ══════════════════════════════════════════════════════════════
# GÉNÉRATION OFX
# ══════════════════════════════════════════════════════════════

def period_to_ofx(date_str):
    try:
        p = date_str.split('/')
        return f"{p[2]}{p[1].zfill(2)}{p[0].zfill(2)}"
    except:
        return datetime.now().strftime('%Y%m%d')

def generate_ofx(info, txns, target='quadra'):
    # target='quadra'    -> NAME porte le libelle (Quadra/Cegid)
    # target='myunisoft' -> MEMO porte le libelle (MyUnisoft/Sage/EBP)
    bid, brid, aid = iban_to_rib(info.get('iban',''))
    ds  = period_to_ofx(info.get('period_start',''))
    de  = period_to_ofx(info.get('period_end',''))
    dn  = datetime.now().strftime('%Y%m%d%H')
    bal = info.get('balance_close', 0.0)
    lines = [
        'OFXHEADER:100','DATA:OFXSGML','VERSION:102','SECURITY:NONE',
        'ENCODING:USASCII','CHARSET:1252','COMPRESSION:NONE',
        'OLDFILEUID:NONE','NEWFILEUID:NONE',
        '<OFX>','<SIGNONMSGSRSV1>','<SONRS>','<STATUS>',
        '<CODE>0','<SEVERITY>INFO','</STATUS>',
        f'<DTSERVER>{dn}','<LANGUAGE>FRA',
        '</SONRS>','</SIGNONMSGSRSV1>',
        '<BANKMSGSRSV1>','<STMTTRNRS>','<TRNUID>00',
        '<STATUS>','<CODE>0','<SEVERITY>INFO','</STATUS>',
        '<STMTRS>','<CURDEF>EUR','<BANKACCTFROM>',
        f'<BANKID>{bid}',f'<BRANCHID>{brid}',
        f'<ACCTID>{aid}','<ACCTTYPE>CHECKING','</BANKACCTFROM>',
        '<BANKTRANLIST>',f'<DTSTART>{ds}',f'<DTEND>{de}',
    ]
    memo_carries_label = target in ('myunisoft', 'sage', 'ebp')
    for t in txns:
        name = t['name']
        memo = t.get('memo', '') or ''
        if memo_carries_label:
            name_tag = name
            memo_tag = (name + ' | ' + memo) if memo else name
        else:
            name_tag = name
            memo_tag = memo
        lines += [
            '<STMTTRN>',
            f"<TRNTYPE>{t['type']}",
            f"<DTPOSTED>{t['date']}",
            f"<TRNAMT>{t['amount']:.2f}",
            f"<FITID>{t['fitid']}",
            '<NAME>' + name_tag,
            '<MEMO>' + memo_tag,
            '</STMTTRN>',
        ]
    lines += [
        '</BANKTRANLIST>',
        '<LEDGERBAL>',f'<BALAMT>{bal:.2f}',f'<DTASOF>{dn}','</LEDGERBAL>',
        '<AVAILBAL>',f'<BALAMT>{bal:.2f}',f'<DTASOF>{dn}','</AVAILBAL>',
        '</STMTRS>','</STMTTRNRS>','</BANKMSGSRSV1>','</OFX>',
    ]
    return '\n'.join(lines) + '\n'



# ══════════════════════════════════════════════════════════════
# PARSEUR myPOS
# Format: relevé mensuel myPOS Ltd (IBAN IE32MPOS...)
# Colonnes: Date | Via | Type | Description | Taux | Débit | Crédit
# Chaque transaction s'étale sur plusieurs lignes dans le PDF.
# On parse le texte brut ligne par ligne.
# ══════════════════════════════════════════════════════════════

def parse_mypos(pages_words, pages_text):
    """
    Format réel myPOS extrait par pdfplumber :
    La description est sur la ligne AVANT la date+type+montants.
    Ex:
        SAINT MAURDO - 000409 / Transaction fee, -0.89
        31.12.2024 23:03 System Fee 1.0000 0.89 0.00
        EUR
        SAINT MAURDO - 000409 / Payment on TID
        31.12.2024 23:03 myPOS Payment 1.0000 0.00 52.90
        80163012, 52.90 EUR
    On parcourt toutes les lignes et on cherche le pattern:
       DD.MM.YYYY HH:MM  <TYPE>  1.0000  <debit>  <credit>
    La ligne précédente (ou 2 lignes avant) est la description.
    """
    info = _extract_mypos_header(pages_text)
    txns = []

    full_text = '\n'.join(pages_text)
    lines = [l.strip() for l in full_text.splitlines()]

    # Regex: date heure type taux débit crédit
    txn_re = re.compile(
        r'^(\d{2}\.\d{2}\.\d{4})\s+\d{2}:\d{2}\s+'
        r'(System Fee|myPOS Payment|Glass Payment|Outgoing bank transfer|POS Payment|Mobile)\s*'
        r'.*?1\.0000\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s*$'
    )
    # Regex sans le taux (fallback)
    txn_re2 = re.compile(
        r'^(\d{2}\.\d{2}\.\d{4})\s+\d{2}:\d{2}\s+'
        r'(System Fee|myPOS Payment|Glass Payment|Outgoing bank transfer|POS Payment|Mobile)'
        r'\s+.*?([\d]+\.[\d]{2})\s+([\d]+\.[\d]{2})\s*$'
    )

    for idx, line in enumerate(lines):
        m = txn_re.match(line)
        if not m:
            m = txn_re2.match(line)
        if not m:
            continue

        date_raw  = m.group(1)           # DD.MM.YYYY
        txn_type  = m.group(2).strip()
        try:
            debit_val  = float(m.group(3).replace(',', ''))
            credit_val = float(m.group(4).replace(',', ''))
        except ValueError:
            continue

        date_ofx = f"{date_raw[6:10]}{date_raw[3:5]}{date_raw[0:2]}"

        # Description = ligne précédente (parfois 2 lignes avant)
        description = ''
        for back in (1, 2):
            if idx >= back:
                prev = lines[idx - back].strip()
                # Exclure les lignes qui sont elles-mêmes des transactions ou des en-têtes
                if (prev and
                        not re.match(r'^\d{2}\.\d{2}\.\d{4}', prev) and
                        not prev.startswith('Monthly statement') and
                        not prev.startswith('Date de valeur') and
                        not prev.startswith('Commande') and
                        not re.match(r'^(Débit|Crédit|EUR)$', prev)):
                    description = prev
                    break

        # Construire nom et mémo
        if txn_type == 'System Fee':
            name = 'myPOS Fee'
            memo = description
        elif 'Outgoing bank transfer' in txn_type or txn_type == 'Mobile':
            name = 'Virement sortant'
            memo = description
        elif txn_type == 'POS Payment':
            name = description or 'Paiement carte'
            memo = 'POS Payment'
        else:
            # myPOS Payment, Glass Payment…
            # Extraire un libellé court: "SAINT MAURDO - 000409 / Payment on TID 80163012"
            # → on garde juste la partie après "/"
            slash_m = re.search(r'/\s*(.+)', description)
            if slash_m:
                name = slash_m.group(1).strip()
                # Retirer le montant en fin : ", 52.90 EUR"
                name = re.sub(r',\s*[\d.]+\s*EUR\s*$', '', name).strip()
            else:
                name = description or txn_type
            memo = description

        # Signe
        if debit_val > 0:
            amount = -debit_val
        elif credit_val > 0:
            amount = credit_val
        else:
            continue

        txns.append(_make_txn(date_ofx, amount, name[:64], memo[:128]))

    return info, txns


def _extract_mypos_header(pages_text):
    info = {'iban': '', 'period_start': '', 'period_end': '',
            'balance_open': 0.0, 'balance_close': 0.0}
    text = pages_text[0] if pages_text else ''

    # IBAN
    m = re.search(r'IBAN\s*:?\s*(IE\d{2}[A-Z0-9]+)', text)
    if m:
        info['iban'] = m.group(1).replace(' ', '')

    # Période depuis le titre "Monthly statement - MM.YYYY"
    m2 = re.search(r'Monthly statement\s*[-–]\s*(\d{2})\.(\d{4})', text, re.IGNORECASE)
    if m2:
        month, year = m2.group(1), m2.group(2)
        info['period_start'] = f"01/{month}/{year}"
        # Dernier jour du mois
        import calendar
        last_day = calendar.monthrange(int(year), int(month))[1]
        info['period_end'] = f"{last_day:02d}/{month}/{year}"

    # Soldes
    m3 = re.search(r'Solde.*?ouverture\s*:?\s*([\d,.]+)', text, re.IGNORECASE)
    if m3:
        info['balance_open'] = parse_amount(m3.group(1)) or 0.0
    m4 = re.search(r'Solde de cl[ôo]ture\s*:?\s*([\d,.]+)', text, re.IGNORECASE)
    if m4:
        info['balance_close'] = parse_amount(m4.group(1)) or 0.0

    return info


# ══════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════

BANK_LABELS = {
    'QONTO': 'Qonto',
    'LCL':   'LCL (Crédit Lyonnais)',
    'CA':    'Crédit Agricole',
    'CE':    "Caisse d'Épargne",
    'BP':    'Banque Populaire',
    'CIC':   'CIC',
    'CGD':   'Caixa Geral de Depositos',
    'LBP':   'La Banque Postale',
    'SG':    'Société Générale',
    'BNP':   'BNP Paribas',
    'MYPOS': 'myPOS',
}

def convert(pdf_path, output_path=None, target='quadra'):
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"Fichier introuvable: {pdf_path}")
    if output_path is None:
        output_path = pdf_path.with_suffix('.ofx')

    print(f"[1/4] Lecture : {pdf_path.name}")
    pages_words = extract_words_by_page(str(pdf_path))
    pages_text  = extract_text_by_page(str(pdf_path))
    print(f"      {len(pages_words)} page(s)")

    print("[2/4] Détection banque...")
    bank = detect_bank(pages_text)
    print(f"      → {BANK_LABELS.get(bank, 'Non reconnue')}")

    if bank == 'UNKNOWN':
        raise ValueError("Banque non reconnue. Formats supportés : Qonto, LCL, CA, CGD, CE, BP, LBP, SG, CIC, BNP Paribas, myPOS")

    print("[3/4] Parsing des transactions...")
    parsers = {'QONTO':parse_qonto,'LCL':parse_lcl,'CA':parse_ca,'CE':parse_ce,'BP':parse_bp,'CIC':parse_cic,
               'CGD':parse_cgd,'LBP':parse_lbp,'SG':parse_sg,'BNP':parse_bnp,'MYPOS':parse_mypos}
    info, txns = parsers[bank](pages_words, pages_text)
    print(f"      {len(txns)} transaction(s)")
    print(f"      Période  : {info.get('period_start','')} → {info.get('period_end','')}")
    print(f"      IBAN     : {info.get('iban','N/A')}")

    print("[4/4] Génération OFX...")
    ofx = generate_ofx(info, txns, target=target)
    with open(output_path, 'w', encoding='latin-1', errors='replace') as f:
        f.write(ofx)
    print(f"\n✅  Fichier OFX créé : {output_path}")
    return str(output_path), len(txns), info, bank


def convert_pdf(pdf_path, output_dir=None, target='quadra'):
    """Wrapper pour l'interface graphique — appelle le moteur multi-banques."""
    p  = Path(pdf_path)
    od = Path(output_dir) if output_dir else p.parent
    op = od / p.with_suffix(".ofx").name
    out_path, nb_txns, info, bank = convert(str(p), str(op), target=target)
    return str(op), nb_txns, info


# ─────────────────────────────────────────────────────────
# Interface graphique — Design moderne dark/slate
# ─────────────────────────────────────────────────────────

C = {
    "bg":           "#0f1117",   # fond principal très sombre
    "sidebar":      "#16181f",   # sidebar légèrement plus clair
    "card":         "#1e2130",   # cartes
    "card2":        "#252839",   # cartes secondaires / hover
    "accent":       "#6366f1",   # indigo vif
    "accent_dark":  "#4f46e5",
    "accent_glow":  "#23254a",
    "green":        "#10b981",
    "green_dim":    "#0d2e22",
    "red":          "#f43f5e",
    "red_dim":      "#2e0d16",
    "amber":        "#f59e0b",
    "text":         "#e2e8f0",
    "text2":        "#94a3b8",
    "text3":        "#475569",
    "border":       "#2d3148",
    "border2":      "#3d4168",
    "row_even":     "#1e2130",
    "row_odd":      "#1a1d2b",
    "row_hover":    "#252839",
    "debit":        "#f43f5e",
    "credit":       "#10b981",
    "header_bg":    "#12141e",
}

FONT_TITLE  = ("Segoe UI", 20, "bold")
FONT_HEAD   = ("Segoe UI", 11, "bold")
FONT_SUBHEAD= ("Segoe UI", 10, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_TINY   = ("Segoe UI", 8)
FONT_MONO   = ("Consolas", 9)
FONT_BTN    = ("Segoe UI", 10, "bold")


def _styled_btn(parent, text, command, bg, fg="white", pad_x=20, pad_y=9, font=None):
    f = font or FONT_BTN
    b = tk.Button(parent, text=text, command=command, font=f,
                  bg=bg, fg=fg, relief="flat", bd=0,
                  activebackground=bg, activeforeground=fg,
                  padx=pad_x, pady=pad_y, cursor="hand2")
    b.bind("<Enter>", lambda e: b.config(bg=_darken(bg)))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b

def _darken(hex_color):
    """Assombrit légèrement une couleur hex."""
    try:
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
        r = max(0, r - 20); g = max(0, g - 20); b = max(0, b - 20)
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return hex_color


class OFXBridgeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OFX Bridge")
        self.geometry("1100x750")
        self.minsize(900, 620)
        self.configure(bg=C["bg"])
        self.resizable(True, True)

        self.pdf_files       = []
        self.preview_txns    = []
        self.preview_info    = {}
        self.output_dir      = tk.StringVar()
        self.target_software = tk.StringVar(value='quadra')
        self.is_running      = False

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── Layout principal ─────────────────────────────────
    def _build_ui(self):
        # Barre supérieure
        self._build_topbar()
        # Corps: sidebar + contenu
        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True)
        self._build_sidebar(body)
        self._build_main(body)

    # ── Topbar ───────────────────────────────────────────
    def _build_topbar(self):
        bar = tk.Frame(self, bg=C["header_bg"], height=54)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        # Séparateur bas
        sep = tk.Frame(self, bg=C["border"], height=1)
        sep.pack(fill="x")

        logo = tk.Frame(bar, bg=C["header_bg"])
        logo.pack(side="left", padx=24)
        tk.Label(logo, text="*", font=("Segoe UI", 16, "bold"),
                 fg=C["accent"], bg=C["header_bg"]).pack(side="left")
        tk.Label(logo, text="  OFX Bridge", font=("Segoe UI", 13, "bold"),
                 fg=C["text"], bg=C["header_bg"]).pack(side="left")

        # Badge version
        badge = tk.Frame(logo, bg=C["accent_glow"], padx=8, pady=2)
        badge.pack(side="left", padx=(10, 0))
        tk.Label(badge, text="v2.0", font=FONT_TINY,
                 fg=C["accent"], bg=C["accent_glow"]).pack()

        # Info droite
        tk.Label(bar, text="Releves bancaires PDF > OFX", font=FONT_SMALL,
                 fg=C["text3"], bg=C["header_bg"]).pack(side="right", padx=24)

    # ── Sidebar ──────────────────────────────────────────
    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=C["sidebar"], width=260)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)

        # Séparateur droite
        sep = tk.Frame(parent, bg=C["border"], width=1)
        sep.pack(side="left", fill="y")

        # Sidebar scrollable via canvas interne
        sb_canvas = tk.Canvas(sb, bg=C["sidebar"], highlightthickness=0)
        sb_vsb = ttk.Scrollbar(sb, orient="vertical", command=sb_canvas.yview)
        sb_canvas.configure(yscrollcommand=sb_vsb.set)
        sb_vsb.pack(side="right", fill="y")
        sb_canvas.pack(side="left", fill="both", expand=True)

        pad = tk.Frame(sb_canvas, bg=C["sidebar"])
        _sb_win = sb_canvas.create_window((0, 0), window=pad, anchor="nw")
        pad.bind("<Configure>", lambda e: sb_canvas.configure(
            scrollregion=sb_canvas.bbox("all")))
        sb_canvas.bind("<Configure>", lambda e: sb_canvas.itemconfig(
            _sb_win, width=e.width))
        # Molette sur la sidebar
        sb_canvas.bind("<MouseWheel>",
            lambda e: sb_canvas.yview_scroll(-1*(e.delta//120), "units"))

        def _pad(content_fn):
            f = tk.Frame(pad, bg=C["sidebar"])
            f.pack(fill="x", padx=20)
            content_fn(f)

        # ── 1. FICHIERS ──────────────────────────────────
        top = tk.Frame(pad, bg=C["sidebar"])
        top.pack(fill="x", padx=20, pady=(24, 0))

        tk.Label(top, text="FICHIERS", font=("Segoe UI", 8, "bold"),
                 fg=C["text3"], bg=C["sidebar"]).pack(anchor="w", pady=(0, 8))

        drop = tk.Frame(top, bg=C["card"],
                        highlightthickness=1, highlightbackground=C["border2"],
                        cursor="hand2")
        drop.pack(fill="x")
        drop.bind("<Button-1>", lambda e: self._add_files())
        inner = tk.Frame(drop, bg=C["card"])
        inner.pack(pady=16, padx=12)
        inner.bind("<Button-1>", lambda e: self._add_files())
        tk.Label(inner, text="^", font=("Segoe UI", 18),
                 fg=C["accent"], bg=C["card"]).pack()
        tk.Label(inner, text="Déposer des PDF ici", font=FONT_SMALL,
                 fg=C["text2"], bg=C["card"]).pack(pady=(4, 1))
        tk.Label(inner, text="ou cliquer pour parcourir",
                 font=FONT_TINY, fg=C["text3"], bg=C["card"]).pack()

        btn_frame = tk.Frame(top, bg=C["sidebar"])
        btn_frame.pack(fill="x", pady=(8, 0))
        btn_add = _styled_btn(btn_frame, "+ Ajouter PDF", self._add_files,
                              bg=C["accent"], pad_x=12, pad_y=7, font=FONT_SMALL)
        btn_add.pack(side="left", fill="x", expand=True)
        btn_clr = _styled_btn(btn_frame, "X", self._clear_files,
                              bg=C["card2"], fg=C["text2"], pad_x=10, pad_y=7, font=FONT_SMALL)
        btn_clr.pack(side="left", padx=(6, 0))

        self.files_info = tk.Label(top, text="Aucun fichier sélectionné",
                                    font=FONT_TINY, fg=C["text3"],
                                    bg=C["sidebar"], wraplength=210, justify="left")
        self.files_info.pack(anchor="w", pady=(6, 0))

        # ── 2. LOGICIEL COMPTABLE ───────────────
        soft = tk.Frame(pad, bg=C["sidebar"])
        soft.pack(fill="x", padx=20, pady=(16, 0))
        tk.Frame(soft, bg=C["border"], height=1).pack(fill="x", pady=(0, 10))
        tk.Label(soft, text="LOGICIEL COMPTABLE", font=("Segoe UI", 8, "bold"),
                 fg=C["text3"], bg=C["sidebar"]).pack(anchor="w", pady=(0, 6))

        self._soft_btns = {}
        SOFTWARE_OPTIONS = [
            ('Quadra / Cegid', 'quadra'),
            ('MyUnisoft',      'myunisoft'),
            ('Sage',           'sage'),
            ('EBP',            'ebp'),
        ]

        def _make_soft_btn(parent, label, value):
            is_selected = (self.target_software.get() == value)
            btn_bg = C["accent_glow"] if is_selected else C["card"]
            btn_fg = C["accent"]      if is_selected else C["text2"]
            border = C["accent"]      if is_selected else C["border2"]

            frame = tk.Frame(parent, bg=border, padx=1, pady=1)
            frame.pack(fill="x", pady=2)

            inner = tk.Frame(frame, bg=btn_bg, cursor="hand2")
            inner.pack(fill="x")

            dot = tk.Label(inner, text="●" if is_selected else "○",
                           font=("Segoe UI", 8), fg=btn_fg, bg=btn_bg,
                           padx=8, pady=5)
            dot.pack(side="left")

            lbl_widget = tk.Label(inner, text=label, font=FONT_TINY,
                                  fg=btn_fg, bg=btn_bg, anchor="w", pady=5)
            lbl_widget.pack(side="left", fill="x", expand=True, padx=(0, 8))

            def _select(v=value):
                self.target_software.set(v)
                self._refresh_soft_btns()

            for w in (frame, inner, dot, lbl_widget):
                w.bind("<Button-1>", lambda e, v=value: _select(v))

            self._soft_btns[value] = (frame, inner, dot, lbl_widget)

        for _lbl, _val in SOFTWARE_OPTIONS:
            _make_soft_btn(soft, _lbl, _val)

        def _refresh_soft_btns_method():
            selected = self.target_software.get()
            for val, (frame, inner, dot, lbl_widget) in self._soft_btns.items():
                is_sel = (val == selected)
                btn_bg = C["accent_glow"] if is_sel else C["card"]
                btn_fg = C["accent"]      if is_sel else C["text2"]
                border = C["accent"]      if is_sel else C["border2"]
                frame.config(bg=border)
                inner.config(bg=btn_bg)
                dot.config(text="●" if is_sel else "○", fg=btn_fg, bg=btn_bg)
                lbl_widget.config(fg=btn_fg, bg=btn_bg)

        self._refresh_soft_btns = _refresh_soft_btns_method

        # ── 3. TOTAUX ───────────────────────────────
        tot = tk.Frame(pad, bg=C["sidebar"])
        tot.pack(fill="x", padx=20, pady=(16, 0))
        tk.Frame(tot, bg=C["border"], height=1).pack(fill="x", pady=(0, 12))

        tk.Label(tot, text="TOTAUX", font=("Segoe UI", 8, "bold"),
                 fg=C["text3"], bg=C["sidebar"]).pack(anchor="w", pady=(0, 6))

        self.info_labels = {}
        totaux_row = tk.Frame(tot, bg=C["sidebar"])
        totaux_row.pack(fill="x")
        for key, label, color, bg_dim in [
            ("Total Débit",  "Débit",  C["red"],   C["red_dim"]),
            ("Total Crédit", "Crédit", C["green"], C["green_dim"]),
        ]:
            pill = tk.Frame(totaux_row, bg=bg_dim, padx=10, pady=7)
            pill.pack(side="left", fill="x", expand=True, padx=(0, 4))
            tk.Label(pill, text=label, font=FONT_TINY,
                     fg=color, bg=bg_dim, anchor="w").pack(anchor="w")
            lbl = tk.Label(pill, text="—", font=("Consolas", 10, "bold"),
                           fg=color, bg=bg_dim, anchor="w")
            lbl.pack(anchor="w")
            self.info_labels[key] = lbl

        # ── 3. DOSSIER DE SORTIE ─────────────────────────
        doss = tk.Frame(pad, bg=C["sidebar"])
        doss.pack(fill="x", padx=20, pady=(16, 0))
        tk.Frame(doss, bg=C["border"], height=1).pack(fill="x", pady=(0, 12))

        tk.Label(doss, text="DOSSIER DE SORTIE", font=("Segoe UI", 8, "bold"),
                 fg=C["text3"], bg=C["sidebar"]).pack(anchor="w", pady=(0, 4))
        tk.Label(doss, text="Par défaut : même dossier que le PDF",
                 font=FONT_TINY, fg=C["text3"], bg=C["sidebar"],
                 wraplength=210, justify="left").pack(anchor="w", pady=(0, 6))

        dir_row = tk.Frame(doss, bg=C["sidebar"])
        dir_row.pack(fill="x")
        entry = tk.Entry(dir_row, textvariable=self.output_dir,
                         font=FONT_TINY, bg=C["card"], fg=C["text"],
                         insertbackground=C["accent"],
                         relief="flat", bd=0,
                         highlightthickness=1, highlightbackground=C["border2"],
                         highlightcolor=C["accent"])
        entry.pack(side="left", fill="x", expand=True, ipady=6)
        btn_dir = _styled_btn(dir_row, "...", self._choose_output,
                              bg=C["card2"], fg=C["text2"], pad_x=8, pad_y=6,
                              font=FONT_TINY)
        btn_dir.pack(side="left", padx=(4, 0))

        # ── 4. COMPTE DÉTECTÉ (moins critique, en bas) ───
        acct = tk.Frame(pad, bg=C["sidebar"])
        acct.pack(fill="x", padx=20, pady=(16, 24))
        tk.Frame(acct, bg=C["border"], height=1).pack(fill="x", pady=(0, 12))

        tk.Label(acct, text="COMPTE DETECTE", font=("Segoe UI", 8, "bold"),
                 fg=C["text3"], bg=C["sidebar"]).pack(anchor="w", pady=(0, 6))

        for key, icon in [("Compte", "B"), ("Transactions", "T"),
                           ("Format", "F"), ("Devise", "E")]:
            row = tk.Frame(acct, bg=C["sidebar"])
            row.pack(fill="x", pady=2)
            tk.Label(row, text=f"{key}", font=FONT_TINY,
                     fg=C["text3"], bg=C["sidebar"], width=12, anchor="w").pack(side="left")
            lbl = tk.Label(row, text="—", font=FONT_TINY,
                           fg=C["text2"], bg=C["sidebar"], anchor="w")
            lbl.pack(side="left", fill="x", expand=True)
            self.info_labels[key] = lbl

        self.info_labels["Format"].config(text="OFX Standard", fg=C["text2"])
        self.info_labels["Devise"].config(text="EUR",           fg=C["text2"])

    # ── Zone principale ──────────────────────────────────
    def _build_main(self, parent):
        main = tk.Frame(parent, bg=C["bg"])
        main.pack(side="left", fill="both", expand=True)

        # En-tête de la zone principale
        hdr = tk.Frame(main, bg=C["bg"])
        hdr.pack(fill="x", padx=28, pady=(22, 0))

        title_col = tk.Frame(hdr, bg=C["bg"])
        title_col.pack(side="left", fill="both", expand=True)

        tk.Label(title_col, text="Apercu des ecritures",
                 font=FONT_TITLE, fg=C["text"], bg=C["bg"]).pack(anchor="w")

        self.tbl_count = tk.Label(title_col, text="Sélectionnez un PDF pour voir les transactions",
                                   font=FONT_SMALL, fg=C["text3"], bg=C["bg"])
        self.tbl_count.pack(anchor="w", pady=(3, 0))

        # Bouton convertir en haut à droite
        self.convert_btn = _styled_btn(
            hdr, "Convertir en OFX", self._start_conversion,
            bg=C["green"], pad_x=22, pad_y=10, font=FONT_BTN
        )
        self.convert_btn.pack(side="right", padx=(16, 0))

        # Barre de progression (cachée par défaut)
        self.progress_bar = ttk.Progressbar(main, mode="indeterminate")

        # Message résultat
        self.result_label = tk.Label(main, text="", font=FONT_SMALL,
                                      fg=C["green"], bg=C["bg"])
        self.result_label.pack(anchor="e", padx=28)

        # Séparateur
        tk.Frame(main, bg=C["border"], height=1).pack(fill="x", padx=28, pady=(12, 0))

        # ── Tableau scrollable ──
        self._build_table(main)

    def _build_table(self, parent):
        """Construit le tableau de transactions avec scroll vertical."""
        wrapper = tk.Frame(parent, bg=C["bg"])
        wrapper.pack(fill="both", expand=True, padx=28, pady=16)

        # En-têtes fixes
        hdr_frame = tk.Frame(wrapper, bg=C["header_bg"],
                             highlightthickness=1, highlightbackground=C["border"])
        hdr_frame.pack(fill="x")

        cols = [
            ("#",           3,  "center"),
            ("Date",        9,  "w"),
            ("Libellé",     35, "w"),
            ("Mémo",        22, "w"),
            ("Débit",       11, "e"),
            ("Crédit",      11, "e"),
        ]
        for col_name, col_w, anchor in cols:
            fg = C["text3"]
            tk.Label(hdr_frame, text=col_name, font=("Segoe UI", 8, "bold"),
                     fg=fg, bg=C["header_bg"],
                     width=col_w, anchor=anchor, padx=8, pady=8).pack(side="left")

        # Zone scrollable
        scroll_wrap = tk.Frame(wrapper, bg=C["bg"],
                               highlightthickness=1, highlightbackground=C["border"])
        scroll_wrap.pack(fill="both", expand=True)

        self.tbl_canvas = tk.Canvas(scroll_wrap, bg=C["row_even"],
                                     highlightthickness=0)
        vsb = ttk.Scrollbar(scroll_wrap, orient="vertical",
                             command=self.tbl_canvas.yview)
        self.tbl_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tbl_canvas.pack(side="left", fill="both", expand=True)

        self.tbl_body = tk.Frame(self.tbl_canvas, bg=C["row_even"])
        self._tbl_win_id = self.tbl_canvas.create_window(
            (0, 0), window=self.tbl_body, anchor="nw")

        self.tbl_body.bind("<Configure>", lambda e: self.tbl_canvas.configure(
            scrollregion=self.tbl_canvas.bbox("all")))
        self.tbl_canvas.bind("<Configure>", lambda e: self.tbl_canvas.itemconfig(
            self._tbl_win_id, width=e.width))

        # Molette souris
        self.tbl_canvas.bind("<MouseWheel>",
            lambda e: self.tbl_canvas.yview_scroll(-1*(e.delta//120), "units"))
        self.tbl_canvas.bind("<Button-4>",
            lambda e: self.tbl_canvas.yview_scroll(-1, "units"))
        self.tbl_canvas.bind("<Button-5>",
            lambda e: self.tbl_canvas.yview_scroll(1, "units"))

        # Placeholder
        self._show_placeholder()

    def _show_placeholder(self):
        for w in self.tbl_body.winfo_children():
            w.destroy()
        holder = tk.Frame(self.tbl_body, bg=C["row_even"])
        holder.pack(expand=True, pady=60)
        tk.Label(holder, text="[  ]", font=("Segoe UI", 36),
                 fg=C["text3"], bg=C["row_even"]).pack()
        tk.Label(holder, text="Aucun fichier chargé",
                 font=FONT_HEAD, fg=C["text3"], bg=C["row_even"]).pack(pady=(8, 2))
        tk.Label(holder, text="Sélectionnez un PDF dans le panneau de gauche",
                 font=FONT_SMALL, fg=C["text3"], bg=C["row_even"]).pack()

    # ── Actions ──────────────────────────────────────────
    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Sélectionner des relevés PDF",
            filetypes=[("Fichiers PDF", "*.pdf"), ("Tous", "*.*")]
        )
        added = 0
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
                added += 1

        if not self.pdf_files:
            return

        n = len(self.pdf_files)
        self.files_info.config(
            text=(f"{'1 fichier' if n == 1 else f'{n} fichiers'} : "
                  + ", ".join(Path(f).name for f in self.pdf_files[:2])
                  + ("..." if n > 2 else "")),
            fg=C["text2"]
        )
        if added > 0:
            self._load_preview(self.pdf_files[-1])

    def _clear_files(self):
        self.pdf_files.clear()
        self.files_info.config(text="Aucun fichier sélectionné", fg=C["text3"])
        self._reset_info_labels()
        self._show_placeholder()
        self.tbl_count.config(text="Sélectionnez un PDF pour voir les transactions")

    def _reset_info_labels(self):
        for k in ["Compte", "Transactions", "Total Débit", "Total Crédit"]:
            self.info_labels[k].config(text="-", fg=C["text2"])

    def _choose_output(self):
        folder = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if folder:
            self.output_dir.set(folder)

    def _load_preview(self, pdf_path):
        """Charge et affiche toutes les transactions du PDF."""
        # Spinner dans le tableau
        for w in self.tbl_body.winfo_children():
            w.destroy()
        tk.Label(self.tbl_body, text="Analyse en cours...",
                 font=FONT_BODY, fg=C["text3"], bg=C["row_even"],
                 pady=40).pack()
        self.tbl_count.config(text="Chargement...")

        def _do():
            try:
                pages_words = extract_words_by_page(pdf_path)
                pages_text  = extract_text_by_page(pdf_path)
                bank        = detect_bank(pages_text)
                parsers = {'QONTO': parse_qonto, 'LCL': parse_lcl, 'CA': parse_ca,
                           'CE': parse_ce, 'BP': parse_bp, 'CIC': parse_cic,
                           'CGD': parse_cgd, 'LBP': parse_lbp, 'SG': parse_sg,
                           'BNP': parse_bnp, 'MYPOS': parse_mypos}
                if bank not in parsers:
                    raise ValueError(
                        "Banque non reconnue. Formats supportés : "
                        "Qonto, LCL, Crédit Agricole, CGD, Caisse d'Épargne, "
                        "Banque Populaire, La Banque Postale, Société Générale, CIC, BNP Paribas, myPOS"
                    )
                info, txns = parsers[bank](pages_words, pages_text)
                self.after(0, lambda: self._update_preview(txns, info, bank))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda msg=err_msg: self._show_preview_error(msg))

        threading.Thread(target=_do, daemon=True).start()

    def _update_preview(self, txns, info, bank=""):
        self.preview_txns = txns
        self.preview_info = info

        # Vider
        for w in self.tbl_body.winfo_children():
            w.destroy()

        if not txns:
            tk.Label(self.tbl_body,
                     text="Aucune transaction détectée dans ce fichier.",
                     font=FONT_SMALL, fg=C["text3"], bg=C["row_even"], pady=30).pack()
        else:
            # Afficher TOUTES les transactions
            col_widths = [3, 9, 35, 22, 11, 11]
            for i, t in enumerate(txns):
                bg = C["row_even"] if i % 2 == 0 else C["row_odd"]
                row = tk.Frame(self.tbl_body, bg=bg, cursor="arrow")
                row.pack(fill="x")

                # Hover
                row.bind("<Enter>", lambda e, r=row: r.config(bg=C["row_hover"]))
                row.bind("<Leave>", lambda e, r=row, b=bg: r.config(bg=b))

                d = t["date"]
                date_fmt = f"{d[6:8]}/{d[4:6]}/{d[0:4]}"
                is_debit  = t["type"] == "DEBIT"
                amt_color = C["debit"] if is_debit else C["credit"]

                # Numéro de ligne
                tk.Label(row, text=str(i + 1), font=FONT_TINY,
                         fg=C["text3"], bg=bg,
                         width=col_widths[0], anchor="center", padx=4, pady=7
                         ).pack(side="left")

                # Date
                tk.Label(row, text=date_fmt, font=FONT_MONO,
                         fg=C["text2"], bg=bg,
                         width=col_widths[1], anchor="w", padx=6).pack(side="left")

                # Libellé — badge coloré selon type
                name_fr = tk.Frame(row, bg=bg)
                name_fr.pack(side="left", fill="x", expand=True)
                tag_bg  = C["red_dim"]   if is_debit else C["green_dim"]
                tag_fg  = C["red"]       if is_debit else C["green"]
                tag_txt = "-"            if is_debit else "+"
                tk.Label(name_fr, text=tag_txt, font=("Segoe UI", 8, "bold"),
                         fg=tag_fg, bg=tag_bg, padx=3, pady=1
                         ).pack(side="left", padx=(6, 4), pady=5)
                tk.Label(name_fr, text=t["name"][:45], font=FONT_SMALL,
                         fg=C["text"], bg=bg, anchor="w").pack(side="left")

                # Mémo
                tk.Label(row, text=(t.get("memo", "") or "")[:30],
                         font=FONT_TINY, fg=C["text3"], bg=bg,
                         width=col_widths[3], anchor="w", padx=4).pack(side="left")

                # Débit
                debit_txt = f"{abs(t['amount']):,.2f} €".replace(",", " ") if is_debit else ""
                tk.Label(row, text=debit_txt, font=FONT_MONO,
                         fg=C["debit"], bg=bg,
                         width=col_widths[4], anchor="e", padx=8).pack(side="left")

                # Crédit
                credit_txt = f"{t['amount']:,.2f} €".replace(",", " ") if not is_debit else ""
                tk.Label(row, text=credit_txt, font=FONT_MONO,
                         fg=C["credit"], bg=bg,
                         width=col_widths[5], anchor="e", padx=8).pack(side="left")

                # Séparateur fin
                tk.Frame(self.tbl_body, bg=C["border"], height=1).pack(fill="x")

        # Remettre le scroll en haut
        self.tbl_canvas.yview_moveto(0)

        # Infos sidebar
        acctid     = iban_to_rib(info.get("iban", ""))[2]
        bank_label = BANK_LABELS.get(bank, bank) if bank else "-"

        self.info_labels["Compte"].config(text=acctid or "-", fg=C["text2"])
        self.info_labels["Transactions"].config(
            text=f"{len(txns)} ops - {bank_label}", fg=C["text2"])

        total_debit  = sum(abs(t["amount"]) for t in txns if t["type"] == "DEBIT")
        total_credit = sum(t["amount"]      for t in txns if t["type"] == "CREDIT")
        self.info_labels["Total Débit"].config(
            text=f"{total_debit:,.2f} €".replace(",", " "), fg=C["debit"])
        self.info_labels["Total Crédit"].config(
            text=f"{total_credit:,.2f} €".replace(",", " "), fg=C["green"])

        # Compteur
        p_start = info.get("period_start", "")
        p_end   = info.get("period_end",   "")
        period  = f"  {p_start} -> {p_end}" if p_start else ""
        self.tbl_count.config(
            text=f"{len(txns)} transaction(s) détectée(s){period}",
            fg=C["text3"])

    def _show_preview_error(self, msg):
        for w in self.tbl_body.winfo_children():
            w.destroy()
        err_frame = tk.Frame(self.tbl_body, bg=C["row_even"])
        err_frame.pack(expand=True, pady=40)
        tk.Label(err_frame, text="(!)", font=("Segoe UI", 28),
                 fg=C["red"], bg=C["row_even"]).pack()
        tk.Label(err_frame, text="Erreur lors de la lecture",
                 font=FONT_HEAD, fg=C["red"], bg=C["row_even"]).pack(pady=(6, 2))
        tk.Label(err_frame, text=msg, font=FONT_SMALL,
                 fg=C["text3"], bg=C["row_even"],
                 wraplength=500, justify="center").pack()
        self.tbl_count.config(text="Erreur", fg=C["red"])

    def _start_conversion(self):
        if self.is_running:
            return
        if not self.pdf_files:
            messagebox.showwarning("Aucun fichier", "Ajoutez au moins un fichier PDF.")
            return
        self.is_running = True
        self.convert_btn.config(text="Conversion en cours...", state="disabled")
        self.result_label.config(text="")
        self.progress_bar.pack(fill="x", padx=28, pady=(0, 4))
        self.progress_bar.start(10)
        threading.Thread(target=self._run_conversion, daemon=True).start()

    def _run_conversion(self):
        od = self.output_dir.get() or None
        ok, err, paths = 0, 0, []
        for pdf in self.pdf_files:
            try:
                path, n, info = convert_pdf(pdf, od, target=self.target_software.get())
                paths.append(path)
                ok += 1
            except Exception:
                err += 1
        self.after(0, lambda: self._conversion_done(ok, err, paths))

    def _conversion_done(self, ok, err, paths):
        self.is_running = False
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.convert_btn.config(text="Convertir en OFX", state="normal")

        if err == 0:
            self.result_label.config(
                text=f"OK {ok} fichier(s) converti(s) avec succes !",
                fg=C["green"])
            if ok == 1:
                folder = str(Path(paths[0]).parent)
                messagebox.showinfo(
                    "Conversion réussie",
                    f"OK Fichier OFX cree !\n\n{Path(paths[0]).name}\n\nDans : {folder}"
                )
            else:
                messagebox.showinfo(
                    "Conversion réussie",
                    f"OK {ok} fichiers OFX crees !\n\n"
                    + "\n".join(f"- {Path(p).name}" for p in paths)
                )
        else:
            self.result_label.config(
                text=f"[!] {ok} reussi(s), {err} erreur(s)",
                fg=C["red"])
            messagebox.showwarning(
                "Conversion terminée",
                f"{ok} réussi(s), {err} erreur(s).\n"
                "Vérifiez que les fichiers sont bien des relevés bancaires supportés."
            )


if __name__ == "__main__":
    app = OFXBridgeApp()
    app.mainloop()
