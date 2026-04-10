"""
views.py – Merchant Report Django Views
========================================

Flow:
    upload → process_session → review → [download processed files]
                                      → [reupload corrected files]  → generate_report → complete

Key changes vs original:
  • process_session no longer seeds mock CBS data (seeding is done by dump_cbs_to_sqlite.py).
  • The review page shows how many rows still have missing geo-fields so the
    operator knows whether to download-fix-reupload before generating the report.
  • generate_report reads the (possibly re-uploaded) processed files and
    calls generate_final_report directly – no re-processing needed.
"""

import os
import traceback

import pandas as pd
from django.conf import settings
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_POST

from .models import UploadSession, FonepayMerchantCBS, NepalpayMerchantCBS
from .processors import (
    generate_final_report,
    process_fonepay,
    process_nepalpay,
    save_processed_excel,
)

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

GEO_COLS = ['PROVINCE', 'DISTRICT', 'MUNICIPALITY']

def _count_missing_geo(df: pd.DataFrame) -> int:
    """Return number of rows where at least one geo-field is blank."""
    cols = [c for c in GEO_COLS if c in df.columns]
    if not cols:
        return 0
    mask = df[cols].isnull().any(axis=1) | (df[cols] == '').any(axis=1)
    return int(mask.sum())


def _media(path_field) -> str:
    """Full filesystem path from a FileField."""
    return os.path.join(settings.MEDIA_ROOT, path_field.name)


# ─────────────────────────────────────────────────────────────────────────────
#  INDEX
# ─────────────────────────────────────────────────────────────────────────────

def index(request):
    sessions = UploadSession.objects.order_by('-created_at')[:10]
    return render(request, 'core/index.html', {'sessions': sessions})


# ─────────────────────────────────────────────────────────────────────────────
#  UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

def upload(request):
    if request.method == 'POST':
        month   = request.POST.get('month_name', 'Ashwin')
        fp_file = request.FILES.get('fonepay_file')
        np_file = request.FILES.get('nepalpay_file')

        if not fp_file or not np_file:
            return render(request, 'core/upload.html',
                          {'error': 'Both Fonepay and NepaPay files are required.'})

        session = UploadSession.objects.create(
            month_name=month,
            fonepay_file=fp_file,
            nepalpay_file=np_file,
            status='uploaded',
        )
        return redirect('process', session_id=session.id)

    return render(request, 'core/upload.html')


# ─────────────────────────────────────────────────────────────────────────────
#  PROCESS  (uploaded files → enriched/processed files)
# ─────────────────────────────────────────────────────────────────────────────

def process_session(request, session_id):
    session = get_object_or_404(UploadSession, id=session_id)
    try:
        fp_path = _media(session.fonepay_file)
        np_path = _media(session.nepalpay_file)

        # CBS enrichment (no mock seeding – CBS is pre-populated via dump_cbs_to_sqlite.py)
        fp_df, fp_count = process_fonepay(fp_path)
        np_df, np_count = process_nepalpay(np_path)

        proc_dir = os.path.join(settings.MEDIA_ROOT, 'processed')
        os.makedirs(proc_dir, exist_ok=True)

        fp_out = os.path.join(proc_dir, f'fonepay_processed_{session.id}.xlsx')
        np_out = os.path.join(proc_dir, f'nepalpay_processed_{session.id}.xlsx')

        # Save with red-cell highlighting for missing geo-fields
        save_processed_excel(fp_df, fp_out)
        save_processed_excel(np_df, np_out)

        session.fonepay_processed  = f'processed/fonepay_processed_{session.id}.xlsx'
        session.nepalpay_processed = f'processed/nepalpay_processed_{session.id}.xlsx'
        session.fonepay_row_count  = fp_count
        session.nepalpay_row_count = np_count
        session.status             = 'processed'
        session.save()

        fp_missing = _count_missing_geo(fp_df)
        np_missing = _count_missing_geo(np_df)

        return render(request, 'core/review.html', {
            'session':        session,
            'fp_null_count':  fp_missing,
            'np_null_count':  np_missing,
            'total_missing':  fp_missing + np_missing,
        })

    except Exception as e:
        return render(request, 'core/error.html',
                      {'error': str(e), 'trace': traceback.format_exc()})


# ─────────────────────────────────────────────────────────────────────────────
#  REVIEW  (re-visit the review page without re-processing)
# ─────────────────────────────────────────────────────────────────────────────

def review(request, session_id):
    session = get_object_or_404(UploadSession, id=session_id)

    fp_path = _media(session.fonepay_processed)
    np_path = _media(session.nepalpay_processed)

    try:
        fp_df = pd.read_excel(fp_path)
        np_df = pd.read_excel(np_path)
    except Exception:
        fp_df = np_df = pd.DataFrame()

    fp_missing = _count_missing_geo(fp_df)
    np_missing = _count_missing_geo(np_df)

    return render(request, 'core/review.html', {
        'session':        session,
        'fp_null_count':  fp_missing,
        'np_null_count':  np_missing,
        'total_missing':  fp_missing + np_missing,
    })


# ─────────────────────────────────────────────────────────────────────────────
#  DOWNLOAD  (processed files or final report)
# ─────────────────────────────────────────────────────────────────────────────

def download_file(request, session_id, file_type):
    session = get_object_or_404(UploadSession, id=session_id)
    file_map = {
        'fonepay_processed':  session.fonepay_processed,
        'nepalpay_processed': session.nepalpay_processed,
        'final_report':       session.final_report,
    }
    field = file_map.get(file_type)
    if not field:
        raise Http404("Unknown file type.")

    path = _media(field)
    if not os.path.exists(path):
        raise Http404("File not found on disk.")

    fname = os.path.basename(path)
    response = FileResponse(
        open(path, 'rb'),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = f'attachment; filename="{fname}"'
    return response


# ─────────────────────────────────────────────────────────────────────────────
#  RE-UPLOAD  (operator has fixed missing fields and re-uploads corrected files)
#
#  After re-upload we go straight to generate_report.
#  The re-uploaded files are treated as the final processed files –
#  they are NOT re-enriched from CBS.  The operator has already done the work.
# ─────────────────────────────────────────────────────────────────────────────

def reupload(request, session_id):
    session = get_object_or_404(UploadSession, id=session_id)

    if request.method == 'POST':
        fp_file = request.FILES.get('fonepay_file')
        np_file = request.FILES.get('nepalpay_file')

        proc_dir = os.path.join(settings.MEDIA_ROOT, 'processed')
        os.makedirs(proc_dir, exist_ok=True)

        if fp_file:
            fp_path = os.path.join(proc_dir, f'fonepay_processed_{session.id}.xlsx')
            with open(fp_path, 'wb') as fh:
                for chunk in fp_file.chunks():
                    fh.write(chunk)
            session.fonepay_processed = f'processed/fonepay_processed_{session.id}.xlsx'

        if np_file:
            np_path = os.path.join(proc_dir, f'nepalpay_processed_{session.id}.xlsx')
            with open(np_path, 'wb') as fh:
                for chunk in np_file.chunks():
                    fh.write(chunk)
            session.nepalpay_processed = f'processed/nepalpay_processed_{session.id}.xlsx'

        session.status = 'reviewed'
        session.save()
        return redirect('generate_report', session_id=session.id)

    return render(request, 'core/reupload.html', {'session': session})


# ─────────────────────────────────────────────────────────────────────────────
#  GENERATE FINAL REPORT
#  Reads the (possibly re-uploaded / corrected) processed files and builds
#  the 4-sheet Excel report directly.  No re-processing from CBS.
# ─────────────────────────────────────────────────────────────────────────────

def generate_report(request, session_id):
    session = get_object_or_404(UploadSession, id=session_id)
    try:
        fp_path = _media(session.fonepay_processed)
        np_path = _media(session.nepalpay_processed)

        fp_df = pd.read_excel(fp_path)
        np_df = pd.read_excel(np_path)

        report_dir = os.path.join(settings.MEDIA_ROOT, 'reports')
        os.makedirs(report_dir, exist_ok=True)
        report_path = os.path.join(report_dir, f'final_report_{session.id}.xlsx')

        generate_final_report(fp_df, np_df, session.month_name, report_path)

        session.final_report = f'reports/final_report_{session.id}.xlsx'
        session.status       = 'final'
        session.save()

        return render(request, 'core/complete.html', {'session': session})

    except Exception as e:
        return render(request, 'core/error.html',
                      {'error': str(e), 'trace': traceback.format_exc()})