#!/usr/bin/env python3
"""
Automatisation des factures Deztination
========================================
1. Cree le folder de la semaine (dimanche)
2. Copie le template Excel dedans
3. Lit tous les PDFs du folder _zappier
4. Remplit le Excel automatiquement

Usage: python3 factures_auto.py
"""

import os
import re
import shutil
from datetime import datetime, timedelta
from pathlib import Path

ONEDRIVE = Path.home() / "OneDrive"
BASE_FOLDER = ONEDRIVE / "2_Client_Corpo" / "ITI" / "Factures"
ZAPIER_FOLDER = BASE_FOLDER / "_zappier"
TEMPLATE_FILE = BASE_FOLDER / "Visa_report_2026Blank.xlsx"

def check_and_install():
    try:
        import pypdf
        import openpyxl
    except ImportError:
        print("Installing required modules...")
        os.system("pip3 install pypdf openpyxl --quiet")
check_and_install()

from pypdf import PdfReader
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

BLUE_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
BLUE_FONT = Font(color="0070C0", bold=True)

def get_sunday_of_week():
    today = datetime.today()
    sunday = today + timedelta(days=(6 - today.weekday()))
    return sunday.strftime("%Y%m%d")

def get_or_create_weekly_folder():
    week_name = get_sunday_of_week()
    week_folder = BASE_FOLDER / week_name
    if not week_folder.exists():
        print(f"📁 Creating folder: {week_name}")
        week_folder.mkdir(parents=True)
    else:
        print(f"📁 Folder exists: {week_name}")
    return week_folder

def get_or_create_weekly_excel(week_folder):
    week_name = week_folder.name
    excel_name = f"Visa_report_{week_name}.xlsx"
    excel_path = week_folder / excel_name
    if not excel_path.exists():
        if not TEMPLATE_FILE.exists():
            print(f"❌ Template not found: {TEMPLATE_FILE}")
            exit(1)
        shutil.copy2(TEMPLATE_FILE, excel_path)
        print(f"📄 Template copied: {excel_name}")
    else:
        print(f"📄 Excel exists: {excel_name}")
    return excel_path

def detect_airline(conf_number):
    """
    Returns (airline_code, needs_review)
    AC:      confirmation starts with 014 and is 10 digits
    Porter:  confirmation is 6 alphanumeric chars
    Unknown: needs review
    """
    if not conf_number:
        return "VERIFY", True
    conf = conf_number.strip()
    if re.match(r'^014\d+$', conf):
        return "AC", False
    if re.match(r'^[A-Z0-9]{6}$', conf):
        return "Porter", False
    return "VERIFY", True

def extract_pdf_data(pdf_path):
    """
    Returns a LIST of charge dicts, one per payment line.
    Each dict: date, dossier, payer, payer_needs_review, description, facture, currency, total, passengers, project
    """
    try:
        reader = PdfReader(str(pdf_path))
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"

        lines = text.split('\n')

        # --- Shared fields ---
        dossier = None
        date = None
        currency = None
        passengers = []
        project = None
        seen_names = set()

        # Dossier
        m = re.search(r'Dossier:\s*(\d+)', text)
        if m:
            dossier = m.group(1)

        # Date
        for i, line in enumerate(lines):
            if 'Conseiller:' in line:
                for j in range(i+1, min(i+5, len(lines))):
                    m2 = re.match(r'^\s*(\d{1,2}/\d{1,2}/\d{4})\s*$', lines[j])
                    if m2:
                        try:
                            dt = datetime.strptime(m2.group(1), "%m/%d/%Y")
                            date = dt.strftime("%d-%m-%Y")
                        except:
                            date = m2.group(1)
                        break
                break

        # Project
        for line in lines:
            m = re.search(r'Projet/Project:\s*(\S+)', line, re.IGNORECASE)
            if m:
                val = m.group(1).strip()
                if val:
                    project = val
                break

        # Passengers
        def clean_name(raw):
            raw = re.sub(r'\b(Mrs?|Ms|Mme|M)\.\s*', '', raw, flags=re.IGNORECASE).strip()
            raw = re.split(r'\s+(?:REFUSE|CDN|PASSPORT|ASSURANCE|CANCELLATION)', raw, flags=re.IGNORECASE)[0].strip()
            raw = raw.strip(' -')
            if raw and re.match(r'[A-Za-zÀ-ÿ\-\ ]+$', raw):
                return re.sub(r'\s+', '', raw.title())
            return None

        for i, line in enumerate(lines):
            if 'Passagers:' in line:
                m = re.match(r'(.+?)Passagers:', line)
                if m:
                    name = clean_name(m.group(1))
                    if name and name not in seen_names:
                        passengers.append(name)
                        seen_names.add(name)
                for j in range(i+1, min(i+5, len(lines))):
                    next_line = lines[j].strip()
                    if re.match(r'(RÉFÉRENCE|TOTAL|PAIEMENTS|SERVICE FEES|Vols|FOND)', next_line, re.IGNORECASE):
                        break
                    if re.search(r'REFUSE', next_line, re.IGNORECASE):
                        name = clean_name(next_line)
                        if name and name not in seen_names:
                            passengers.append(name)
                            seen_names.add(name)
                break

        # Currency
        if re.search(r'\bUSD\b', text):
            currency = "USD"
        elif re.search(r'\bEUR\b', text):
            currency = "EUR"

        # --- Extract payment lines ---
        payments = []
        in_payments = False
        for line in lines:
            if 'PAIEMENTS' in line:
                in_payments = True
                continue
            if in_payments:
                if 'Total des paiements' in line:
                    break
                m = re.match(r'(\d{8})\s+CAD.*?\s+([-\d,]+\.\d{2})\s*$', line)
                if m:
                    payments.append({
                        'facture': str(int(m.group(1))),
                        'amount': float(m.group(2).replace(',', ''))
                    })

        # --- Extract descriptions (skip FOND D'INDEMNISATION) ---
        descriptions = []
        in_ref = False
        current_desc = []

        for line in lines:
            if 'RÉFÉRENCE' in line or 'REFERENCE' in line:
                in_ref = True
                continue
            if in_ref:
                if 'PAIEMENTS' in line or re.match(r'\s*Total:', line):
                    if current_desc:
                        desc = ' '.join(current_desc).strip()
                        if not re.search(r'FOND D.INDEMNISATION', desc, re.IGNORECASE):
                            desc = re.sub(r'\s+[\d,]+\.\d{2}\s+CAD.*$', '', desc).strip()
                            desc = re.sub(r'\s{2,}', ' ', desc).strip()[:60]
                            descriptions.append(desc)
                    break
                if re.match(r'^\s*[-\d]', line):
                    if current_desc:
                        desc = ' '.join(current_desc).strip()
                        if not re.search(r'FOND D.INDEMNISATION', desc, re.IGNORECASE):
                            desc = re.sub(r'\s+[\d,]+\.\d{2}\s+CAD.*$', '', desc).strip()
                            desc = re.sub(r'\s{2,}', ' ', desc).strip()[:60]
                            descriptions.append(desc)
                    current_desc = []
                    continue
                if line.strip():
                    current_desc.append(line.strip())

        # --- Build charge list, one per payment ---
        # Match descriptions to payments by amount when possible
        # First try direct order match, fallback to amount matching
        charges = []
        used_desc_indices = set()

        def find_desc_for_amount(amount, descriptions, used):
            """Find best matching description for a payment amount"""
            # For negative amounts, look for REFUND descriptions
            if amount < 0:
                for i, d in enumerate(descriptions):
                    if i not in used and re.search(r'REFUND|REMBOURS', d, re.IGNORECASE):
                        return i, d
            # For HERTZ amounts, look for HERTZ descriptions
            for i, d in enumerate(descriptions):
                if i not in used and re.search(r'HERTZ', d, re.IGNORECASE):
                    # Check if amount roughly matches
                    return i, d
            # Default: next unused description
            for i, d in enumerate(descriptions):
                if i not in used:
                    return i, d
            return None, ""

        for idx, payment in enumerate(payments):
            if idx < len(descriptions):
                # Try smart matching for special cases
                if payment['amount'] < 0 or (idx < len(descriptions) and
                    re.search(r'REFUND|HERTZ', descriptions[idx] if idx < len(descriptions) else '', re.IGNORECASE)):
                    di, desc = find_desc_for_amount(payment['amount'], descriptions, used_desc_indices)
                    if di is not None:
                        used_desc_indices.add(di)
                    else:
                        desc = ""
                else:
                    # Use next available description in order
                    di = idx
                    while di in used_desc_indices and di < len(descriptions):
                        di += 1
                    if di < len(descriptions):
                        desc = descriptions[di]
                        used_desc_indices.add(di)
                    else:
                        desc = ""
            else:
                desc = ""

            # Determine payer from description
            is_flight = re.search(r'RESERVATION VOL', desc, re.IGNORECASE) and re.search(r'Vols.*Conf', desc, re.IGNORECASE)
            is_train = re.search(r'VIA RAIL|RESERVATION TRAIN', desc, re.IGNORECASE)

            is_hertz = re.search(r'HERTZ', desc, re.IGNORECASE)

            if is_flight or re.search(r'Vols.*Conf', desc, re.IGNORECASE):
                conf_match = re.search(r'Vols[^#]*#\s*([A-Z0-9]+)', desc, re.IGNORECASE)
                conf = conf_match.group(1).strip() if conf_match else None
                if conf and re.match(r'^014\d+$', conf):
                    payer, needs_review = "AC", False
                elif conf and re.match(r'^[A-Z0-9]{6}$', conf):
                    payer, needs_review = "Porter", False
                else:
                    payer, needs_review = "VERIFY", True
            elif is_train:
                payer, needs_review = "Via", False
            elif is_hertz:
                payer, needs_review = "Hertz", False
            else:
                payer, needs_review = "Deztination", False
                desc = "Frais de service/FICAV"

            charges.append({
                "date": date,
                "dossier": dossier,
                "payer": payer,
                "payer_needs_review": needs_review,
                "description": desc,
                "facture": payment["facture"],
                "currency": currency,
                "total": payment["amount"],
                "passengers": passengers,
                "project": project,
            })

        return charges if charges else None

    except Exception as e:
        print(f"   ⚠️  Error reading PDF: {e}")
        return None

def get_already_processed(ws):
    processed = set()
    for row in ws.iter_rows(min_row=12, max_row=41, values_only=True):
        if row[8] and row[11]:  # facture + total
            processed.add(f"{row[8]}_{row[11]}")
    return processed

def fill_excel(excel_path, all_data):
    wb = load_workbook(str(excel_path))
    ws = wb.active
    already_done = get_already_processed(ws)

    added = 0
    skipped = 0

    # all_data is a list of lists (each PDF returns a list of charges)
    for charge_list in all_data:
        if not charge_list:
            continue
        for charge in charge_list:
            key = f"{charge['facture']}_{charge['total']}"
            if key in already_done:
                skipped += 1
                continue

            # Find next empty row (12-41)
            target_row = None
            for row_num in range(12, 42):
                if ws.cell(row=row_num, column=2).value is None and ws.cell(row=row_num, column=3).value is None:
                    target_row = row_num
                    break

            if target_row is None:
                print("   ⚠️  Excel full — no more rows available (max 30)")
                break

            ws.cell(row=target_row, column=2).value = charge["date"]
            ws.cell(row=target_row, column=5).value = charge["description"]
            ws.cell(row=target_row, column=8).value = charge["dossier"]
            ws.cell(row=target_row, column=9).value = charge["facture"]
            ws.cell(row=target_row, column=10).value = charge["currency"]
            ws.cell(row=target_row, column=12).value = charge["total"]

            payer_cell = ws.cell(row=target_row, column=3)
            if charge["payer_needs_review"]:
                payer_cell.value = "VERIFY"
                payer_cell.fill = BLUE_FILL
                payer_cell.font = BLUE_FONT
            else:
                payer_cell.value = charge["payer"]

            already_done.add(key)
            added += 1

    wb.save(str(excel_path))
    return added, skipped


def rename_pdf(pdf_path, data):
    """Renames PDF to include facture number, passengers and project"""
    dossier = data.get("dossier", "")
    facture = data.get("facture", "")
    passengers = data.get("passengers", [])
    project = data.get("project", "")

    if not facture:
        return

    parts = [f"FactureCL_Deztination-{dossier}", facture]
    parts += passengers
    if project:
        parts.append(project)

    new_name = "_".join(parts) + ".pdf"
    new_path = pdf_path.parent / new_name

    if pdf_path.name != new_name:
        try:
            pdf_path.rename(new_path)
            print(f"   📝 Renamed: {new_name}")
        except Exception as e:
            print(f"   ⚠️  Could not rename: {e}")
    else:
        print(f"   📝 Name already correct")

def main():
    print("=" * 50)
    print("  INVOICE AUTOMATION — DEZTINATION")
    print("=" * 50)
    print()

    if not BASE_FOLDER.exists():
        print(f"❌ Base folder not found: {BASE_FOLDER}")
        return

    if not ZAPIER_FOLDER.exists():
        print(f"❌ Zappier folder not found: {ZAPIER_FOLDER}")
        return

    week_folder = get_or_create_weekly_folder()
    excel_path = get_or_create_weekly_excel(week_folder)

    pdfs = list(ZAPIER_FOLDER.glob("*.pdf"))

    if not pdfs:
        print(f"\n📭 No PDFs found in: {ZAPIER_FOLDER}")
        print("   Add PDFs to the _zappier folder and run again.")
        return

    print(f"\n🔍 {len(pdfs)} PDF(s) found\n")

    all_data = []
    for pdf in pdfs:
        print(f"   Reading: {pdf.name}")
        charges = extract_pdf_data(pdf)
        if charges:
            for charge in charges:
                status = "⚠️  VERIFY payer" if charge["payer_needs_review"] else "✅"
                print(f"   {status} Invoice {charge['facture']} — {charge['payer']} — {charge['total']} {'CAD' if not charge['currency'] else charge['currency']}")
            all_data.append(charges)
            rename_pdf(pdf, charges[0])
        else:
            print(f"   ❌ Could not read this PDF")

    print(f"\n📊 Updating Excel...")
    added, skipped = fill_excel(excel_path, all_data)

    print()
    print("=" * 50)
    print(f"  ✅ DONE")
    print(f"  {added} invoice(s) added")
    if skipped > 0:
        print(f"  {skipped} invoice(s) already present — skipped")
    print(f"  File: {excel_path.name}")
    if any(c["payer_needs_review"] for charges in all_data for c in charges):
        print(f"  ⚠️  Some payers marked VERIFY in blue — please fill in manually")
    print("=" * 50)

if __name__ == "__main__":
    main()
