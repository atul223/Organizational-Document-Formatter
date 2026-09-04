# Deploying the Formatting Tool on your Ubuntu VM (free-of-cost)

This assumes the Ubuntu VM your IT person mentioned is a normal Ubuntu
Server (20.04/22.04/24.04) with SSH access. No Azure, no paid services —
everything below runs on the VM itself.

## 0. What you're deploying

```
formatting-tool/
├── backend/
│   ├── app.py                 (FastAPI wrapper — untouched pipeline underneath)
│   ├── requirements.txt
│   └── pipeline/               (your original phase-1/2/3 scripts, unmodified)
│       ├── format_doc.py
│       ├── policy_extractor.py
│       ├── structure_classifier.py
│       ├── formatting_engine.py
│       ├── integrity_validator.py
│       ├── toc_card_builder.py
│       ├── docx_fast.py
│       └── ooxml_fonts.py
└── frontend/
    └── index.html              (single-page upload UI)
```

## 1. Copy the project to the VM

From your machine:
```bash
scp -r formatting-tool youruser@<vm-ip>:/tmp/formatting-tool
```

On the VM:
```bash
sudo mkdir -p /opt/formatting-tool
sudo mv /tmp/formatting-tool/* /opt/formatting-tool/
```

## 2. Create a dedicated (non-login) service user

Running the app as its own low-privilege user, rather than root or your
personal account, is standard practice and costs nothing extra:
```bash
sudo useradd --system --no-create-home --shell /usr/sbin/nologin formattingtool
sudo chown -R formattingtool:formattingtool /opt/formatting-tool
```

## 3. Install Python and create a virtual environment

Ubuntu ships Python 3 by default; just confirm venv support is installed:
```bash
sudo apt update
sudo apt install -y python3-venv python3-pip
```

Create the venv and install pinned dependencies:
```bash
cd /opt/formatting-tool/backend
sudo -u formattingtool python3 -m venv venv
sudo -u formattingtool venv/bin/pip install --upgrade pip
sudo -u formattingtool venv/bin/pip install -r requirements.txt
```

## 4. Quick manual test (before wiring up systemd)

```bash
sudo -u formattingtool venv/bin/uvicorn app:app --host 127.0.0.1 --port 8000
```
In another terminal on the VM:
```bash
curl http://127.0.0.1:8000/api/health
# should print: {"status":"ok"}
```
Stop it with Ctrl+C once confirmed.

## 5. Set it up as a systemd service (so it survives reboots/crashes)

```bash
sudo cp /opt/formatting-tool/deploy/formatting-tool.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable formatting-tool
sudo systemctl start formatting-tool
sudo systemctl status formatting-tool     # confirm it's "active (running)"
```

Logs any time:
```bash
sudo journalctl -u formatting-tool -f
```

## 6. Put Nginx in front of it (so employees hit a normal URL, not :8000)

```bash
sudo apt install -y nginx
sudo cp /opt/formatting-tool/deploy/nginx_formatting-tool.conf /etc/nginx/sites-available/formatting-tool
```
Edit `server_name` in that file to the VM's internal hostname or IP, then:
```bash
sudo ln -s /etc/nginx/sites-available/formatting-tool /etc/nginx/sites-enabled/
sudo nginx -t          # test config syntax
sudo systemctl reload nginx
```

Employees can now open `http://<vm-hostname-or-ip>/` from any office
machine on the same network and use the tool directly in their browser —
no installs, no VS Code, no local Python needed on their end.

## 7. Firewall (if ufw is active)

```bash
sudo ufw allow 'Nginx Full'
```
Do **not** open port 8000 to the network — only Nginx (port 80) should be
externally reachable; the app itself stays bound to 127.0.0.1.

## 8. Restricting access to your organization only

Since this handles internal company documents, two free options:
- **Simplest**: the VM is already only reachable on your office LAN/VPN —
  if true, you're done; nothing external can reach it.
- **Add a login prompt**: ask IT to put an HTTP Basic Auth layer in Nginx
  (a few lines with `htpasswd`), or better, front it with your existing
  Azure AD/Entra ID via an Nginx `auth_request` + OAuth2 Proxy — free,
  open-source, and reuses your M365 logins. Worth a follow-up once the
  tool is validated internally.

## 9. Updating the tool later

```bash
# From your machine:
scp -r backend/pipeline youruser@<vm-ip>:/tmp/pipeline-update
# On the VM:
sudo systemctl stop formatting-tool
sudo cp -r /tmp/pipeline-update/* /opt/formatting-tool/backend/pipeline/
sudo systemctl start formatting-tool
```

## Notes on the temp files

Each request creates a private temp folder under `/tmp/formatting_tool_jobs/`
holding the two uploaded files and intermediate JSON/artifacts; it's deleted
automatically right after the formatted file is sent back to the browser.
Nothing from an employee's documents is kept on the server between requests.
