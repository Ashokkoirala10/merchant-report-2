#!/usr/bin/env python3
"""
CBS Excel → SQLite Dumper
=========================
Usage:
    python dump_cbs_to_sqlite.py

Run from inside your merchant_report/ Django project directory.
This script reads cbs_fonepay.xlsx and cbs_nepalpay.xlsx and
upserts all rows into the Django SQLite database.

Make sure both CBS Excel files are in the same directory as this script,
or update the paths below.
"""

import os, sys, django

# ── Point at your Django project ────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
django.setup()

import pandas as pd
from core.models import FonepayMerchantCBS, NepalpayMerchantCBS

# ── File paths — edit if your files are elsewhere ────────────────────────────
FONEPAY_CBS_FILE  = os.path.join(BASE_DIR, 'cbs_fonepay.xlsx')
NEPALPAY_CBS_FILE = os.path.join(BASE_DIR, 'cbs_nepalpay.xlsx')


def dump_fonepay_cbs(filepath):
    if not os.path.exists(filepath):
        print(f"[SKIP] File not found: {filepath}")
        return
    df = pd.read_excel(filepath, dtype={'merchant_id': str})
    df.columns = [c.strip().lower() for c in df.columns]
    created, updated = 0, 0
    for _, row in df.iterrows():
        mid = str(row.get('merchant_id', '')).strip()
        if not mid:
            continue
        obj, was_created = FonepayMerchantCBS.objects.update_or_create(
            merchant_id=mid,
            defaults=dict(
                merchant_name = str(row.get('merchant_name', '')),
                province      = str(row.get('province',      '')),
                district      = str(row.get('district',      '')),
                municipality  = str(row.get('municipality',  '')),
                address1      = str(row.get('address1',      '')),
                address3      = str(row.get('address3',      '')),
                gender        = str(row.get('gender',        '')),
            )
        )
        if was_created: created += 1
        else:           updated += 1
    print(f"[Fonepay CBS]  Created: {created:>3}  |  Updated: {updated:>3}  |  Total: {created+updated}")


def dump_nepalpay_cbs(filepath):
    if not os.path.exists(filepath):
        print(f"[SKIP] File not found: {filepath}")
        return

    df = pd.read_excel(filepath, dtype={'merchant_code': str, 'merchant_account': str})
    df.columns = [c.strip().lower() for c in df.columns]

    created, updated = 0, 0
    for _, row in df.iterrows():
        mc = str(row.get('merchant_code', '')).strip()
        ma = str(row.get('merchant_account', '')).strip()

        if not mc:
            continue

        # Use composite key: (merchant_code + merchant_account)
        obj, was_created = NepalpayMerchantCBS.objects.update_or_create(
            merchant_code=mc,
            merchant_account=ma if ma and ma not in ('nan', 'None', '') else None,
            defaults={
                'merchant_name': str(row.get('merchant_name', '')),
                'province':      str(row.get('province', '')),
                'district':      str(row.get('district', '')),
                'municipality':  str(row.get('municipality', '')),
                'address1':      str(row.get('address1', '')),
                'address3':      str(row.get('address3', '')),
                'gender':        str(row.get('gender', '')),
            }
        )
        if was_created:
            created += 1
        else:
            updated += 1

    print(f"[Nepalpay CBS] Created: {created:>3} | Updated: {updated:>3} | Total: {created+updated}")
if __name__ == '__main__':
    print("=" * 50)
    print("  CBS Excel → SQLite Dumper")
    print("=" * 50)
    dump_fonepay_cbs(FONEPAY_CBS_FILE)
    dump_nepalpay_cbs(NEPALPAY_CBS_FILE)
    print("=" * 50)
    print("Done. CBS tables updated successfully.")
