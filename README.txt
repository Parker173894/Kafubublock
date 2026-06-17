KAFUBU BLOCK SECONDARY SCHOOL WEBSITE

TECHNOLOGIES USED
- Python
- Flask
- SQLite
- HTML
- CSS
- JavaScript-ready structure
- Werkzeug password hashing
- File upload support

NEW FEATURES ADDED
1. Student portal
   - Students login and view their Test 1, Test 2 and End Term results.
   - Example student login: student1 / student123

2. Teacher results entry
   - Teachers, HODs, Headteacher and Deputy can add student results.
   - Results include Test 1, Test 2 and End Term marks.
   - The system calculates total and average automatically.

3. Student register
   - Staff can register students before adding results.

4. HR account under staff portal
   - HR can login using: hr / hr123
   - HR can access staff portal functions and upload official/HR documents.

5. Staff profile update
   - Staff can update full name, phone, email, qualification, address and bio.
   - Staff can upload profile pictures.
   - Profile pictures show on the profiles page and department pages.

DEFAULT LOGIN DETAILS
Headteacher: head / head123
Deputy: deputy / deputy123
HR: hr / hr123
Subject Teacher: teacher / teacher123
HOD Mathematics: hod_math / math123
HOD Natural Science: hod_natural / natural123
HOD Social Science: hod_social / social123
HOD Computer Science: hod_computer / computer123
HOD Business: hod_business / business123
HOD Home Economics: hod_home / home123
HOD Language: hod_language / language123
Student 1: student1 / student123
Student 2: student2 / student123

HOW TO RUN
1. Extract the ZIP file.
2. Open Command Prompt inside the project folder.
3. Run: pip install -r requirements.txt
4. Run: python app.py
5. Open: http://127.0.0.1:5000


LATEST UPDATES ADDED
--------------------
1. Subject teacher profiles now have a Department field under My Profile.
   When a teacher chooses Mathematics, Natural Science, Social Science, Computer Science, Business, Home Economics or Language, the teacher profile is automatically shown on that department page.

2. Leadership, HR and HOD profile editing is supported through Portal > Update My Profile.
3. Users can change their portal password under My Profile from the dashboard.
   They can update full name, position, phone, email, qualification, address, bio and profile picture.

3. Pupil grades include Form 1, Form 2, Form 3, Form 4, Form 5, Grade 10, Grade 11 and Grade 12.
   These forms appear in the student register and results entry pages.


UPDATED PORTAL ACCESS
---------------------
The School Portal Login now has provisions for Headteacher, Deputy Headteacher, HR, all HODs, Subject Teacher and Student accounts.

Full access accounts:
- Headteacher: head / head123
- HR: hr / hr123

Full access means the user can access department materials, official/HR documents, student register, student results, donor needs and leadership/department profile sections.

Deputy Headteacher portal:
- deputy / deputy123

HOD portals:
- Mathematics HOD: hod_math / math123
- Natural Science HOD: hod_natural / natural123
- Social Science HOD: hod_social / social123
- Computer Science HOD: hod_computer / computer123
- Business HOD: hod_business / business123
- Home Economics HOD: hod_home / home123
- Language HOD: hod_language / language123

LATEST UPDATE - STUDENT RESULT ANALYSIS
- Added Student Result Analysis page under the student/staff portal.
- Students can view their own analysis of Test 1, Test 2 and End Term marks.
- Staff can filter result analysis by student, grade/form, term and academic year.
- HODs, Headteacher, Deputy Headteacher and HR can download result analysis as a CSV file.

LATEST UPDATE: DEPUTY HEADTEACHER STAFF ACCESS
- The Deputy Headteacher can now manage HOD profiles and Subject Teacher profiles.
- Deputy Headteacher can edit names, positions, departments where applicable, contact details, qualifications, bios and profile pictures.
- Deputy Headteacher can delete HOD and Subject Teacher profiles, but cannot delete the currently logged-in account.
- HOD departments remain controlled by their HOD account role to avoid assigning an HOD to the wrong department.

NEW UPDATE: GUIDANCE AND COUNSELLING PORTAL
- Added Guidance and Counselling Teacher account.
- Guidance teachers can upload educational posts and videos for learners.
- Students can view guidance messages and videos from their student portal.
- Headteacher, Deputy Headteacher and HR can also upload/delete guidance posts.

Guidance and Counselling login:
Username: guidance
Password: guidance123


PUBLIC GUIDANCE AND COUNSELLING UPDATE
- Guidance and Counselling Posts / Videos are now public.
- Visitors, parents, guardians and learners can open the Guidance & Counselling menu without logging in.
- Uploaded guidance videos are also viewable publicly.
- Only authorised users can upload or delete guidance posts: Headteacher, Deputy Headteacher, HR and Guidance and Counselling Teacher.

LATEST UPDATE: STUDENT PDF RESULTS AND PORTAL CONTROL
-----------------------------------------------------
1. Students can now download their results as a PDF report form from the Student Portal.
   - Login as a student.
   - Open Dashboard.
   - Click Download Results PDF or View My Results then Download Results in PDF.

2. The PDF report form includes:
   - School name
   - Student name
   - Student number
   - Grade/Form and class
   - Class teacher name
   - Test 1, Test 2 and End Term marks
   - Total, average and teacher comments
   - Result analysis summary
   - Spaces for Class Teacher and Headteacher signatures

3. Student Register now includes a Class Teacher field.
   - When registering a student, enter the name of the class teacher.
   - This name appears on the result report form.

4. Headteacher, Deputy Headteacher and HR can activate or deactivate the Student Results Portal.
   - Login as Headteacher, Deputy or HR.
   - Open Dashboard.
   - Click Activate / Deactivate Student Results Portal.
   - Select Active when results are ready to be viewed.
   - Select Inactive when results should be hidden from students.

5. When the Student Results Portal is inactive:
   - Students cannot view results.
   - Students cannot download PDF result reports.
   - Students cannot view result analysis.
   - Staff can still enter and manage results.

ADDITIONAL REQUIREMENT
----------------------
This version uses ReportLab to generate PDF files.
It is included in requirements.txt, so run:

pip install -r requirements.txt

LATEST UPDATE: STUDENT DELETION AND RESULT DOWNLOAD RECORDS
- Headteacher, Deputy Headteacher, HR, HODs and Subject Teachers can delete students from the Student Register.
- Deleting a student also removes that student's portal login account and related result records.
- The system now records every PDF result download in a result_download_logs table.
- Staff can view this from Dashboard > Result Download Records.
- The download record shows student number, student name, grade/form, class, downloader name, role and date/time.


BACKGROUND SLIDESHOW UPDATE
- Headteacher, Deputy Headteacher and HODs can upload/manage website background pictures.
- Active background pictures rotate automatically every 15 seconds.
- The slideshow uses a smooth fade and slight zoom effect, similar to modern school websites.
- To change the timing later, open templates/base.html and edit slideDuration = 15000.


HR Staff Return Details Setup:
- Login as HR: hr / hr123
- Go to Dashboard -> Set Staff Return Details Required from Teachers
- Add any extra details HR wants teachers to submit, for example payroll number, bank name, number of periods taught, union membership, or deployment status.
- Active details automatically appear on the Staff Return Registration form.
- HR can edit teachers' staff return records and download all information in CSV or Excel.

SECURITY UPDATE ADDED
---------------------
This secured version includes:
1. Login attempt protection: after 5 wrong attempts from the same computer/IP, login is blocked for 15 minutes.
2. Password strength rules: new/reset passwords must have at least 8 characters, one capital letter, one small letter and one number.
3. Security audit records: Headteacher and HR can view successful logins, failed logins, blocked login attempts and logouts.
4. Password audit records: Headteacher and HR can see password changes and resets.
5. Safer file uploads: the system blocks dangerous file types and allows only approved documents, images and videos.
6. Upload size limit: files above 25MB are blocked.
7. Safer session cookies: cookies are HttpOnly and SameSite=Lax.
8. Security headers: the app sends browser security headers such as X-Frame-Options, X-Content-Type-Options and Content-Security-Policy.
9. Secret key protection: the app now generates a strong secret key automatically instead of using a weak fixed key.

IMPORTANT HOSTING ADVICE
------------------------
For real online hosting:
- Use HTTPS/SSL.
- Set FLASK_SECRET_KEY as an environment variable.
- Set FLASK_HTTPS_ONLY=1 only after HTTPS is active.
- Change all default passwords before showing the system to the public.
- Keep a backup copy of school.db and the uploads folder.


UPDATE: Guidance and Counselling video uploads now support videos up to 750MB. If hosting online, configure the hosting server/proxy upload limit to at least 750MB as well.
