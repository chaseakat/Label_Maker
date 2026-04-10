# Label Generator Dashboard

Simple Flask website + dashboard to run your label-generation logic from uploaded work orders.

## What It Does
- Accepts a work order file (`.pdf` or image)
- Accepts logo input (`.jpg/.png/...` or `.zip` containing an image)
- Runs your parsing and PDF label generation logic
- Returns generated labels as a direct PDF download (single file) or zip (two files)
- Lets you choose the output directory where PDFs are saved

## Setup
```bash
cd /root/label_dashboard
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run
```bash
cd /root/label_dashboard
source .venv/bin/activate
python app.py
```

Open `http://localhost:5000`.

## One-Step Install And Run
```bash
cd /root/label_dashboard
chmod +x install_and_run.sh
./install_and_run.sh
```

## Notes
- OCR for image inputs requires Tesseract installed on the machine.
- Generated output defaults to `/root/label_dashboard/generated_labels` unless overridden in the dashboard.
