# Deploy LJIET Attendance Portal to PythonAnywhere

This guide walks you through deploying the attendance portal on [PythonAnywhere](https://www.pythonanywhere.com).

---

## 1. Create a PythonAnywhere Account

1. Go to [pythonanywhere.com](https://www.pythonanywhere.com) and sign up (free tier available).
2. Note your username (e.g. `dhruvpd77`). Your site will be at `https://dhruvpd77.pythonanywhere.com`.

---

## 2. Upload the Project

### Option A: Clone from GitHub (recommended)

1. Open a **Bash console** on PythonAnywhere (Dashboard → Consoles → Bash).
2. Run:

```bash
cd ~
git clone https://github.com/dhruvpd77/attendance_portal.git
cd attendance_portal
```

### Option B: Upload manually

1. Zip your project folder locally.
2. In PythonAnywhere: **Files** tab → navigate to your home directory → **Upload a file**.
3. Upload the zip, then in Bash: `unzip attendance_portal.zip` and `cd attendance_portal`.

---

## 3. Create a Virtual Environment

In the Bash console:

```bash
cd ~/attendance_portal
python3.10 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

> Use the Python version shown in your PythonAnywhere account (e.g. 3.10 or 3.11).

---

## 4. Set Up Environment Variables

Create a `.env` file in the project root:

```bash
cd ~/attendance_portal
nano .env
```

Add (replace `YOUR_USERNAME` with your PythonAnywhere username):

```
SECRET_KEY=your-long-random-secret-key-here
DEBUG=False
ALLOWED_HOSTS=YOUR_USERNAME.pythonanywhere.com
```

Generate a secret key:

```bash
python -c "import secrets; print(secrets.token_urlsafe(50))"
```

Save and exit (Ctrl+O, Enter, Ctrl+X in nano).

---

## 5. Run Migrations and Collect Static Files

```bash
cd ~/attendance_portal
source venv/bin/activate
python manage.py migrate
python manage.py collectstatic --noinput
```

---

## 6. Create a Web App on PythonAnywhere

1. Go to **Web** tab in the PythonAnywhere dashboard.
2. Click **Add a new web app** → choose **Manual configuration** (not Django) → select your Python version.
3. After creation, in **Code** section:
   - **Source code**: `/home/YOUR_USERNAME/attendance_portal`
   - **Working directory**: `/home/YOUR_USERNAME/attendance_portal`

4. **WSGI configuration file**: Click the link to edit it. Replace the entire content with:

```python
import os
import sys

# Add project to path
path = '/home/YOUR_USERNAME/attendance_portal'
if path not in sys.path:
    sys.path.insert(0, path)

os.environ['DJANGO_SETTINGS_MODULE'] = 'config.settings'

from django.core.wsgi import get_wsgi_application
application = get_wsgi_application()
```

Replace `YOUR_USERNAME` with your actual PythonAnywhere username in all three places.

5. **Virtualenv**: Set to `/home/YOUR_USERNAME/attendance_portal/venv`

6. **Static files** (in the Web tab, scroll to Static files):
   - URL: `/static/`
   - Directory: `/home/YOUR_USERNAME/attendance_portal/staticfiles`

7. **Media files** (optional, if you use file uploads):
   - URL: `/media/`
   - Directory: `/home/YOUR_USERNAME/attendance_portal/media`

---

## 7. Reload the Web App

Click the green **Reload** button in the Web tab.

---

## 8. Create a Superuser (First-Time Setup)

In the Bash console:

```bash
cd ~/attendance_portal
source venv/bin/activate
python manage.py createsuperuser
```

Follow the prompts to create an admin account.

---

## 9. Access Your Site

- **Main site**: `https://YOUR_USERNAME.pythonanywhere.com/`
- **Admin**: `https://YOUR_USERNAME.pythonanywhere.com/admin/`
- **Portal (login)**: `https://YOUR_USERNAME.pythonanywhere.com/portal/`

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| 500 error | Check **Web** → **Error log** for details |
| Static files not loading | Ensure `collectstatic` ran and static mapping is correct |
| Import errors | Verify virtualenv path and that `pip install -r requirements.txt` completed |
| Database errors | Run `python manage.py migrate` again |

---

## Updating the App (After Code Changes)

```bash
cd ~/attendance_portal
git pull   # if using Git
source venv/bin/activate
pip install -r requirements.txt
python manage.py migrate
python manage.py collectstatic --noinput
```

Then **Reload** the web app in the Web tab.
