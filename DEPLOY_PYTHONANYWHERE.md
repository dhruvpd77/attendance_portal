# Deploy LJIET_Attendance on PythonAnywhere

## 0. Update Existing Deployment (fetch latest code from GitHub)

If the app is already deployed and you want to pull the latest changes:

```bash
cd ~/attendance_portal
git pull origin main
pip install -r requirements.txt
python manage.py migrate
python manage.py collectstatic --noinput
```

Then go to **Web** tab → click **Reload** to restart the app.

---

## 1. General Setup

### 1.1 Create PythonAnywhere account
- Sign up at [pythonanywhere.com](https://www.pythonanywhere.com)
- Free accounts: limited emails (see below). Paid accounts ($5/mo): full SMTP access.

### 1.2 Clone your repo
```bash
cd ~
git clone https://github.com/dhruvpd77/attendance_portal.git
cd attendance_portal
```

### 1.3 Create virtualenv
```bash
mkvirtualenv attendance_portal --python=python3.10
pip install -r requirements.txt
```

### 1.4 Create `.env` file
```bash
cp .env.example .env
nano .env   # or use the editor in PythonAnywhere
```

---

## 2. Required Settings for PythonAnywhere

Add these to your `.env` (replace `YOUR_USERNAME` with your PythonAnywhere username):

```env
# Required
SECRET_KEY=your-secret-key-here
DEBUG=False
ALLOWED_HOSTS=YOUR_USERNAME.pythonanywhere.com

# Optional: keep default for local SQLite
# DATABASE_URL=sqlite:///db.sqlite3
```

---

## 3. Email Configuration (for sending mentor reports)

### Option A: Gmail SMTP (works on paid PythonAnywhere accounts)

```env
EMAIL_BACKEND=django.core.mail.backends.smtp.EmailBackend
DEFAULT_FROM_EMAIL=LJIET Attendance <your-email@gmail.com>
EMAIL_HOST=smtp.gmail.com
EMAIL_PORT=587
EMAIL_USE_TLS=True
EMAIL_HOST_USER=your-email@gmail.com
EMAIL_HOST_PASSWORD=your-app-password
```

**Gmail setup:**
1. Enable 2-factor authentication on your Google account
2. Go to [Google App Passwords](https://myaccount.google.com/apppasswords)
3. Create an app password for "Mail"
4. Use that 16-character password as `EMAIL_HOST_PASSWORD`

### Option B: Outlook / Office 365

```env
EMAIL_BACKEND=django.core.mail.backends.smtp.EmailBackend
EMAIL_HOST=smtp.office365.com
EMAIL_PORT=587
EMAIL_USE_TLS=True
EMAIL_HOST_USER=your-email@outlook.com
EMAIL_HOST_PASSWORD=your-password
DEFAULT_FROM_EMAIL=LJIET Attendance <your-email@outlook.com>
```

### Option C: Free PythonAnywhere account

- **Free accounts can only send emails TO your own PythonAnywhere email** (e.g. `YOUR_USERNAME@pythonanywhere.com`).
- Outbound SMTP to Gmail/Outlook is blocked on free tier.
- To send real emails: upgrade to a paid account ($5/mo) or use a transactional email service (SendGrid, Mailgun, etc.).

### Option D: SendGrid (free tier: 100 emails/day)

```env
EMAIL_BACKEND=django.core.mail.backends.smtp.EmailBackend
EMAIL_HOST=smtp.sendgrid.net
EMAIL_PORT=587
EMAIL_USE_TLS=True
EMAIL_HOST_USER=apikey
EMAIL_HOST_PASSWORD=your-sendgrid-api-key
DEFAULT_FROM_EMAIL=LJIET Attendance <verified-sender@yourdomain.com>
```

---

## 4. Web App Configuration

1. Go to **Web** tab on PythonAnywhere
2. **Add a new web app** → Manual configuration
3. **Source code:** `/home/YOUR_USERNAME/attendance_portal`
4. **Working directory:** `/home/YOUR_USERNAME/attendance_portal`
5. **WSGI file:** Edit `/var/www/YOUR_USERNAME_pythonanywhere_com_wsgi.py`:

```python
import os
import sys

path = '/home/YOUR_USERNAME/attendance_portal'
if path not in sys.path:
    sys.path.insert(0, path)

os.environ['DJANGO_SETTINGS_MODULE'] = 'config.settings'

from django.core.wsgi import get_wsgi_application
application = get_wsgi_application()
```

6. **Static files:**
   - Run: `python manage.py collectstatic --noinput`
   - Add mapping: URL `/static/` → Directory `/home/YOUR_USERNAME/attendance_portal/staticfiles`

7. **Reload** the web app.

---

## 5. Database

```bash
cd ~/attendance_portal
python manage.py migrate
python manage.py createsuperuser
```

---

## 6. Checklist

| Item | Status |
|------|--------|
| `.env` with SECRET_KEY, DEBUG=False, ALLOWED_HOSTS | ☐ |
| Email vars (EMAIL_BACKEND, EMAIL_HOST, etc.) for SMTP | ☐ |
| Gmail App Password or SendGrid API key | ☐ |
| `collectstatic` run | ☐ |
| Static files mapping in Web tab | ☐ |
| WSGI file configured | ☐ |
| `migrate` and `createsuperuser` | ☐ |

---

## 7. Emails Not Sending? Troubleshooting

### Option A: Use Web App Environment Variables (most reliable on PythonAnywhere)

Instead of relying on `.env`, set variables in the **Web** tab:

1. Open your Web app → **Configuration for...**
2. Scroll to **Environment variables**
3. Add each variable (one per line):
   ```
   EMAIL_BACKEND=django.core.mail.backends.smtp.EmailBackend
   EMAIL_HOST=smtp.gmail.com
   EMAIL_PORT=587
   EMAIL_USE_TLS=True
   EMAIL_HOST_USER=your-email@gmail.com
   EMAIL_HOST_PASSWORD=your-16-char-app-password-no-spaces
   DEFAULT_FROM_EMAIL=LJIET Attendance <your-email@gmail.com>
   ```
4. **Important:** For `EMAIL_HOST_PASSWORD`, use the Gmail App Password **without spaces** (e.g. `abcdabcdabcdabcd` not `abcd abcd abcd abcd`)
5. Click **Reload**

### Option B: Fix .env

- Ensure `.env` is in `/home/LJIET/attendance_portal/.env` (same folder as `manage.py`)
- Remove spaces from the Gmail App Password in `.env`
- Restart the web app after editing `.env`

### Option C: Test in PythonAnywhere Console

1. Open **Consoles** → **Python 3.10**
2. Run:
   ```python
   import os
   os.chdir('/home/LJIET/attendance_portal')
   from dotenv import load_dotenv
   load_dotenv('.env')
   print('EMAIL_BACKEND:', os.environ.get('EMAIL_BACKEND'))
   print('EMAIL_HOST:', os.environ.get('EMAIL_HOST'))
   print('EMAIL_HOST_USER:', os.environ.get('EMAIL_HOST_USER'))
   ```
3. If these print empty, `.env` is not loading. Use Web tab env vars instead.

### Option D: Check error log

- Web tab → **Error log** – look for SMTP/auth errors
- Common: "Username and Password not accepted" → wrong App Password or use without spaces

---

## 8. Other Troubleshooting

**Emails not sending:**
- Gmail App Password: use **no spaces** (16 chars: `vvmmiirpdihocrlf`)
- On free PythonAnywhere: Gmail SMTP should work (PA has firewall exception)

**Static files not loading:**
- Run `python manage.py collectstatic`
- Ensure static URL `/static/` → `staticfiles/` in Web tab

**502 Bad Gateway:**
- Check WSGI path and `DJANGO_SETTINGS_MODULE`
- Check error log in Web tab
