#!/bin/bash
set -a
source "$(dirname "$0")/.env"
set +a
source "$(dirname "$0")/venv/bin/activate"
python3 "$(dirname "$0")/CreateInvoice.py"
