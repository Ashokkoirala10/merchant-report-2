# NRB Merchant Report Generator

A Django web application to generate Nepal Rastra Bank (NRB) province-wise, district-wise, local-level and gender-wise merchant transaction reports from Fonepay and NepalpPay month-end data.

## Setup & Run

### Requirements
- Python 3.8+
- pip

### Install dependencies

```bash
pip install -r requirements.txt
```

### Initialize the database

```bash
python manage.py migrate
```

### Run the development server

```bash
python manage.py runserver
```

Then open: **http://127.0.0.1:8000**

---

## How It Works

### Step 1 — Upload Files
Upload the Fonepay month-end `.xlsx` file and the NepalpPay month-end `.xlsx` file. Enter the report month name (e.g. "Ashwin 2081").

**Fonepay columns used:** `MERCHANT_ID`, `MERCHANT_NAME`, `ISSUER_NAME`, `TERMINAL_DETAILS_ID`, `PROVINCE`, `DISTRICT`, `MUNICIPALITY`, `ORIGINAL_AMOUNT`, `PAYMENT_MODULE`

**NepalpPay columns used:** `Merchant Code`, `Merchant Name`, `Amount`, `QR Type`, `Transaction Date`, `Issuer Id`

### Step 2 — CBS Mapping (Automatic)
The system:
- Extracts unique merchant IDs (Fonepay) and merchant codes (NepalpPay)
- Looks them up in the mock CBS SQLite database (`FonepayMerchantCBS`, `NepalpayMerchantCBS` tables)
- Fills in missing `PROVINCE`, `DISTRICT`, `MUNICIPALITY`, `GENDER` from CBS
- Appends `ADDRESS1` and `ADDRESS3` for manual review of any remaining gaps
- **Row count is always preserved** — no rows are dropped during mapping

### Step 3 — Review & Download
- Download the processed Excel files
- Rows with missing geographic data are highlighted in **red**
- Use `ADDRESS1`/`ADDRESS3` columns to manually fill in gaps
- Either proceed directly or re-upload the corrected files

### Step 4 — Final Report
Generated Excel report contains 4 sheets:
1. **Province Wise** — QR-enabled (Fonepay), POS-enabled (empty), Online-enabled (NepalpPay) by 7 provinces
2. **District Wise** — All 77 districts breakdown
3. **Local Level Wise** — Metro / Sub-Metro / Municipality / Rural Municipality
4. **Gender Wise** — Male, Female, Others, Company

---

## CBS Database

The app uses a mock SQLite CBS database. In production:
- Replace `seed_mock_fonepay_cbs()` and `seed_mock_nepalpay_cbs()` in `core/processors.py` with real CBS API/DB queries
- The `FonepayMerchantCBS` and `NepalpayMerchantCBS` models can be replaced with direct CBS connections

You can also populate the CBS tables via Django Admin at: `http://127.0.0.1:8000/admin/`

---

## Project Structure

```
merchant_report/
├── config/             # Django project settings & URLs
├── core/
│   ├── models.py       # FonepayMerchantCBS, NepalpayMerchantCBS, UploadSession
│   ├── views.py        # Upload, process, review, generate views
│   ├── processors.py   # All data processing & Excel generation logic
│   └── templates/core/ # HTML templates
├── media/              # Uploaded & generated files (auto-created)
├── manage.py
└── requirements.txt
```
