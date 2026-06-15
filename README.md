# Invoice Automation

Drops PDF invoices in, gets a filled-out weekly Excel report 
back — no manual copying.

## Business Impact

This replaced a weekly routine that used to mean opening each 
PDF one by one and typing vendor names, amounts, dates, and 
invoice numbers into Excel by hand. On one production use 
case, that was 5-6 hours a week — about 280 hours a year.

It also checks invoice number plus amount before adding a row, 
so re-running it on the same folder never creates duplicates. 
And when it can't confidently match a vendor, it doesn't 
guess - it flags the row "VERIFY" in blue so a human checks 
just that one thing.

## How It Works

Drop PDFs into a watch folder and run the script. It pulls out 
vendor, date, amount, currency, invoice number, and (when 
relevant) traveler names and project codes, then fills the 
next open row in a weekly Excel report built from a template. 
PDFs get renamed automatically with a consistent naming 
pattern. Safe to run again on the same folder - anything 
already processed gets skipped.

## Tech Stack

Python 3, pypdf for reading the PDFs, openpyxl for the Excel 
side. Field extraction is regex-based and configurable - 
there's a CONFIG block at the top of the file for folder 
paths, the Excel template, and vendor keywords.

## Setup

Clone the repo, edit the CONFIG section in `factures_auto.py` 
(base folder, inbox folder, Excel template path, vendor 
keywords), then run:
