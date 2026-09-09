#!/bin/bash
# Self-locating startup script for Azure App Service.
# Instead of hardcoding a --chdir path that has to exactly match wherever
# Azure/Oryx unpacks the repo, this finds app.py wherever it actually is
# under /home/site/wwwroot and starts the app from there. This makes the
# startup command immune to future changes in folder layout.
set -e

echo "Searching for app.py under /home/site/wwwroot ..."
APP_FILE=$(find /home/site/wwwroot -maxdepth 4 -name "app.py" -not -path "*/antenv/*" -not -path "*/__pycache__/*" | head -n 1)

if [ -z "$APP_FILE" ]; then
    echo "ERROR: Could not find app.py anywhere under /home/site/wwwroot."
    echo "Directory listing for debugging:"
    find /home/site/wwwroot -maxdepth 3
    exit 1
fi

APP_DIR=$(dirname "$APP_FILE")
echo "Found app.py in: $APP_DIR"
cd "$APP_DIR"

exec gunicorn --bind=0.0.0.0:8000 --timeout 600 --workers 2 -k uvicorn.workers.UvicornWorker app:app