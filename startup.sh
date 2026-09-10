#!/bin/bash
# Self-locating, path-agnostic startup script for Azure App Service.
#
# IMPORTANT: Azure/Oryx does NOT always extract your app to
# /home/site/wwwroot - with compressed package deployments it extracts
# to a dynamically-named temp folder (e.g. /tmp/8df0e6d1251db15) instead,
# and changes into that directory before running this script. So this
# script never hardcodes a base path - it always searches relative to
# its own current working directory ($(pwd)) at the moment it runs.
set -e

ROOT_DIR="$(pwd)"
echo "startup.sh running from: $ROOT_DIR"

echo "Searching for app.py under $ROOT_DIR ..."
APP_FILE=$(find "$ROOT_DIR" -maxdepth 4 -name "app.py" -not -path "*/antenv/*" -not -path "*/__pycache__/*" 2>/dev/null | head -n 1)

if [ -z "$APP_FILE" ]; then
    echo "ERROR: Could not find app.py anywhere under $ROOT_DIR."
    echo "Directory listing for debugging:"
    find "$ROOT_DIR" -maxdepth 3
    exit 1
fi

APP_DIR=$(dirname "$APP_FILE")
echo "Found app.py in: $APP_DIR"

echo "Searching for antenv virtual environment under $ROOT_DIR ..."
ANTENV_ACTIVATE=$(find "$ROOT_DIR" -maxdepth 3 -path "*/antenv/bin/activate" 2>/dev/null | head -n 1)
if [ -n "$ANTENV_ACTIVATE" ]; then
    echo "Activating virtual environment: $ANTENV_ACTIVATE"
    # shellcheck disable=SC1090
    source "$ANTENV_ACTIVATE"
else
    echo "WARNING: No antenv virtual environment found under $ROOT_DIR - gunicorn/uvicorn may not be importable."
fi

cd "$APP_DIR"
echo "Launching gunicorn from: $(pwd)"
exec gunicorn --bind=0.0.0.0:8000 --timeout 600 --workers 2 -k uvicorn.workers.UvicornWorker app:app