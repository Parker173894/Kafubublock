KAFUBU BLOCK SECONDARY SCHOOL PORTAL
SECURITY AND BACKUP UPGRADE
====================================

This package upgrades the existing Flask school portal without including a
database, uploaded school files, or a secret key. Preserve the live school.db
and uploads folder when deploying it.

MAIN SECURITY IMPROVEMENTS
--------------------------
1. Known demonstration passwords are never accepted in production.
2. Demonstration accounts and sample pupils are not created in production.
3. New pupils receive a random temporary password instead of their pupil number.
4. New and reset accounts must create a private password at first login.
5. All POST forms are protected with a CSRF security token.
6. Production requires a strong FLASK_SECRET_KEY stored in environment variables.
7. Password-reset tokens are stored as hashes, not usable links.
8. Recovery links are not displayed or stored when production email is unavailable.
9. Session cookies require HTTPS in production, with HSTS and no-store protection.
10. Untrusted X-Forwarded-For headers are ignored unless TRUST_PROXY is enabled.
11. The default request limit is reduced from 750 MB to 50 MB.
12. Images are limited to 8 MB, documents to 25 MB, and videos to 50 MB by default.

SUBJECT TEACHER PORTAL UPGRADE
------------------------------
- Subject Teachers have a separate Teacher Login page at /teacher-login.
- Each teacher uses an individual username and password.
- Successful teacher login opens that teacher's own editable profile.
- Teachers can update their name, position, department, phone, email,
  qualification, address, biography and profile picture.
- Teachers can change a known password from the Change Password page.
- Forgotten passwords are reset by the Headteacher or HR.
- Under New Subject Teacher Registration, an HOD creates the teacher's
  individual username and temporary password.
- The HOD can manually enter or securely generate a strong temporary password.
- After registration, a printable one-time handoff page shows the login details.
- The new teacher must replace the temporary password at first login.
- A new Subject Teacher is automatically assigned to the registering HOD's
  department; the browser cannot override this assignment.
- Teacher profile edits automatically update the live department staff page
  and the matching staff-return record.
- Only active Subject Teachers are displayed on public department pages.

MODERN DESIGN UPGRADE
---------------------
- New high-standard visual system across the public website and portals.
- Modern green-and-gold school identity, typography, cards, forms and tables.
- Responsive navigation with a mobile menu for phones and tablets.
- Redesigned homepage, department directory and department staff pages.
- Clear live-record indicators and automatic department-assignment messages.
- Improved keyboard focus, reduced-motion support and readable mobile layouts.

PUPIL RESULT LOOKUP UPGRADE
---------------------------
- A public Check Pupil Results page is available at /pupil-results.
- Pupils enter both their student number and full registered name.
- Results are shown only when both details match the same pupil record.
- The lookup follows the existing management-controlled results portal status.
- Repeated unsuccessful searches are temporarily limited for privacy.
- Public result pages are marked no-cache and are hidden from search engines.

VISUAL AND LOGIN REFINEMENTS
----------------------------
- Guidance and Counselling posts use dark readable text on a soft green-and-gold
  background, including the homepage post preview.
- Uploaded active website backgrounds are now clearly visible through a lighter
  overlay while all page text remains readable.
- "Other Portal Login" has been renamed to "Admin Login".
- The duplicate Subject Teacher login button was removed from the Admin Login
  page; teachers continue using the dedicated Teacher Login.

HR TEACHER AND PUPIL RECORDS
----------------------------
- HR has a combined records page showing active Subject Teachers, active HODs,
  total teaching staff and total registered pupils.
- Teaching-staff totals are summarised by department.
- Pupil totals are summarised by grade/form.
- HR can search pupils and change their current grade, class and class teacher.
- Every grade change records the old and new details, HR officer and date.
- Historical academic results are preserved under the grade in which they were
  earned.
- Selecting a grade displays every pupil currently registered in that grade.
- HR can change all pupils in a selected grade at once, with optional shared
  class and class-teacher assignments.

FULL BACKGROUND PICTURE DISPLAY
-------------------------------
- Website background pictures use full-image display instead of cropping.
- The slideshow zoom effect was removed so picture edges remain visible.
- The light overlay was reduced so uploaded pictures appear more clearly.
- Background previews use full-image display and can be opened at full size.

BACKUP AND RESTORE
------------------
- Headteacher and HR accounts can open Portal > Backup and Restore Centre.
- A complete backup contains a consistent database snapshot and uploaded files.
- Backups can be downloaded and stored away from the hosting server.
- Automatic backups run every 24 hours in production by default.
- Up to 30 automatic backups are retained by default.
- Restoring requires the current manager password and typing RESTORE.
- A safety backup is automatically created before every restore.

IMPORTANT: A backup on the same server is not enough. Download regular copies
to a protected computer, encrypted drive, or approved cloud storage.

UPGRADING THE ALREADY-HOSTED APP
--------------------------------
1. Download a copy of the live school.db and uploads folder before doing anything.
2. Stop the hosted web service.
3. Replace the old application code with the files in this package.
4. KEEP the existing live school.db and uploads folder. Do not overwrite them.
5. Set the production environment variables listed below.
6. Install the updated requirements.
7. Start the service using: gunicorn wsgi:app
8. Log in as Headteacher.
9. If the previous Headteacher password was still the public demonstration
   password, use INITIAL_ADMIN_PASSWORD for the first login.
10. Create a new private password when prompted.
11. Open Manage / Reset Portal Passwords and issue temporary passwords to any
    older accounts that were disabled because they still used demonstration passwords.
12. Create and download a backup from the Backup and Restore Centre.

REQUIRED PRODUCTION ENVIRONMENT VARIABLES
-----------------------------------------
APP_ENV=production
FLASK_SECRET_KEY=<random value of at least 32 characters>
INITIAL_ADMIN_PASSWORD=<at least 12 characters with upper/lowercase and a number>
RECOVERY_ALLOWED_EMAIL=<official school recovery email>

INITIAL_ADMIN_PASSWORD is used only when creating the first Headteacher account
or replacing the old public Headteacher password. After the first successful
login, change the password in the portal and remove INITIAL_ADMIN_PASSWORD from
the hosting environment if your provider allows it.

EMAIL RECOVERY VARIABLES
------------------------
SMTP_HOST=<mail server>
SMTP_PORT=587
SMTP_USER=<mail account>
SMTP_PASSWORD=<mail password or app password>
SMTP_FROM=<sender address>

OPTIONAL SETTINGS
-----------------
TRUST_PROXY=1                 Only if the host uses a trusted reverse proxy.
MAX_UPLOAD_MB=50              Whole-request limit; allowed range is 5-100 MB.
MAX_GUIDANCE_VIDEO_MB=50      Must not exceed MAX_UPLOAD_MB.
MAX_DOCUMENT_MB=25
MAX_IMAGE_MB=8
BACKUP_DIR=<persistent private directory>
AUTO_BACKUP_ENABLED=1
AUTO_BACKUP_HOURS=24
BACKUP_RETENTION=30

HOSTING REQUIREMENTS
--------------------
- HTTPS/SSL must be enabled.
- school.db, uploads, and backups must be stored on persistent storage.
- On Render, attach a persistent disk and set DATA_DIR=/var/data.
- Do not place backups inside a public/static directory.
- Configure the host request limit to match MAX_UPLOAD_MB.
- Use one Gunicorn worker with SQLite unless the database is migrated to PostgreSQL.
- Follow RENDER_DEPLOYMENT_GUIDE.txt when upgrading an existing Render service.

INSTALL AND START
-----------------
pip install -r requirements.txt
gunicorn wsgi:app

LOCAL DEMONSTRATION ONLY
------------------------
For an offline demonstration with sample accounts, run:

python app.py

Directly running app.py automatically uses local development/demo mode. Render
uses Gunicorn to import wsgi:app, which remains locked to production mode.

Never enable APP_DEMO_MODE on the hosted school website.

DATA PROTECTION NOTES
---------------------
- Staff returns can contain NRC numbers, addresses, dates of birth, next-of-kin
  details and phone numbers. Give HR access only to authorised staff.
- Do not send database backups through public messaging groups.
- Remove accounts promptly when a staff member leaves the school.
- Review Security Audit Records and Password Change Records regularly.
