import os, io, secrets, sqlite3, smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, abort
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook

load_dotenv()
app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'change-this-secret-before-deployment')
DB_PATH = os.getenv('DATABASE_PATH', 'candidates.db')
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
ALLOWED_EXTENSIONS = {'.xlsx'}


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


def init_db():
    with db() as con:
        con.executescript('''
        CREATE TABLE IF NOT EXISTS users (
          id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
          password_hash TEXT NOT NULL, role TEXT NOT NULL CHECK(role IN ('admin','candidate')),
          candidate_id INTEGER
        );
        CREATE TABLE IF NOT EXISTS candidates (
          id INTEGER PRIMARY KEY AUTOINCREMENT, candidate_code TEXT UNIQUE NOT NULL,
          full_name TEXT NOT NULL, email TEXT NOT NULL UNIQUE, phone TEXT DEFAULT '',
          assigned_task TEXT DEFAULT '', task_status TEXT DEFAULT 'Pending',
          attendance_status TEXT DEFAULT 'Not scheduled', last_review TEXT DEFAULT '',
          next_review TEXT DEFAULT '', task_deadline TEXT DEFAULT '', meet_link TEXT DEFAULT '',
          feedback TEXT DEFAULT '', review_status TEXT DEFAULT 'Waiting',
          created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS reviews (
          id INTEGER PRIMARY KEY AUTOINCREMENT, candidate_id INTEGER NOT NULL,
          scheduled_at TEXT NOT NULL, meet_link TEXT DEFAULT '', status TEXT DEFAULT 'Scheduled',
          feedback TEXT DEFAULT '', task_snapshot TEXT DEFAULT '', created_at TEXT NOT NULL,
          FOREIGN KEY(candidate_id) REFERENCES candidates(id) ON DELETE CASCADE
        );
        ''')
        # Seed admin account once; override with environment variables in .env.
        username = os.getenv('ADMIN_USERNAME', 'admin')
        password = os.getenv('ADMIN_PASSWORD', 'ChangeMe123!')
        row = con.execute('SELECT id FROM users WHERE username=?', (username,)).fetchone()
        if not row:
            con.execute('INSERT INTO users(username,password_hash,role) VALUES(?,?,?)',
                        (username, generate_password_hash(password), 'admin'))


def send_email(to, subject, body):
    host = os.getenv('SMTP_HOST', '')
    user = os.getenv('SMTP_USERNAME', '')
    password = os.getenv('SMTP_PASSWORD', '')
    sender = os.getenv('EMAIL_FROM', user)
    if not all([host, user, password, sender, to]):
        app.logger.warning('Email not sent: SMTP settings are incomplete (recipient hidden).')
        return False
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to
    msg.set_content(body)
    try:
        with smtplib.SMTP(host, int(os.getenv('SMTP_PORT', '587')), timeout=15) as smtp:
            smtp.starttls()
            smtp.login(user, password)
            smtp.send_message(msg)
        return True
    except Exception:
        app.logger.exception('Email delivery failed')
        return False


def login_required(role=None):
    def decorator(fn):
        @wraps(fn)
        def wrapped(*args, **kwargs):
            if not session.get('user_id'):
                return redirect(url_for('login'))
            if role and session.get('role') != role:
                abort(403)
            return fn(*args, **kwargs)
        return wrapped
    return decorator


def candidate_for_user():
    if session.get('role') != 'candidate' or not session.get('candidate_id'):
        abort(403)
    with db() as con:
        candidate = con.execute('SELECT * FROM candidates WHERE id=?', (session['candidate_id'],)).fetchone()
    if not candidate:
        session.clear()
        abort(403)
    return candidate


def parse_dt(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def schedule_review(candidate_id, scheduled_at, meet_link=''):
    when = parse_dt(scheduled_at)
    if not when:
        raise ValueError('Choose a valid review date and time.')
    with db() as con:
        overlap = con.execute('''SELECT id FROM reviews WHERE scheduled_at=? AND status='Scheduled' ''',
                              (when.isoformat(timespec='minutes'),)).fetchone()
        if overlap:
            raise ValueError('That exact time slot is already booked. Choose another time.')
        candidate = con.execute('SELECT * FROM candidates WHERE id=?', (candidate_id,)).fetchone()
        if not candidate:
            raise ValueError('Candidate not found.')
        link = meet_link.strip() or os.getenv('DEFAULT_MEET_LINK', '')
        con.execute('''INSERT INTO reviews(candidate_id,scheduled_at,meet_link,status,task_snapshot,created_at)
                       VALUES(?,?,?,'Scheduled',?,?)''',
                    (candidate_id, when.isoformat(timespec='minutes'), link, candidate['assigned_task'], datetime.now().isoformat(timespec='seconds')))
        con.execute('''UPDATE candidates SET next_review=?, meet_link=?, attendance_status='Scheduled', review_status='Scheduled'
                       WHERE id=?''', (when.isoformat(timespec='minutes'), link, candidate_id))
        return candidate['email'], candidate['full_name'], candidate['candidate_code'], candidate['assigned_task'], link, when


@app.route('/')
def index():
    if session.get('role') == 'admin':
        return redirect(url_for('dashboard'))
    if session.get('role') == 'candidate':
        return redirect(url_for('candidate_home'))
    return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        with db() as con:
            user = con.execute('SELECT * FROM users WHERE username=?', (username,)).fetchone()
        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['role'] = user['role']
            session['candidate_id'] = user['candidate_id']
            return redirect(url_for('dashboard' if user['role'] == 'admin' else 'candidate_home'))
        flash('Invalid username or password.', 'error')
    return render_template('login.html')


@app.route('/logout', methods=['POST'])
@login_required()
def logout():
    session.clear()
    flash('You have logged out successfully.', 'success')
    return redirect(url_for('login'))


@app.route('/admin')
@login_required('admin')
def dashboard():
    with db() as con:
        candidates = con.execute('SELECT * FROM candidates ORDER BY full_name').fetchall()
        reviews = con.execute('''SELECT r.*, c.full_name, c.candidate_code, c.email FROM reviews r
                                 JOIN candidates c ON c.id=r.candidate_id ORDER BY r.scheduled_at LIMIT 30''').fetchall()
    stats = {
        'total': len(candidates),
        'tasks_done': sum(c['task_status'] == 'Completed' for c in candidates),
        'waiting': sum(c['review_status'] in ('Waiting', 'Scheduled') for c in candidates),
        'reviewed': sum(c['review_status'] == 'Completed' for c in candidates),
        'absent': sum(c['attendance_status'] == 'Absent' for c in candidates),
        'overdue': sum(bool(c['task_deadline'] and c['task_deadline'] < datetime.now().date().isoformat() and c['task_status'] != 'Completed') for c in candidates)
    }
    return render_template('dashboard.html', candidates=candidates, reviews=reviews, stats=stats)


@app.route('/candidate/add', methods=['GET', 'POST'])
@login_required('admin')
def add_candidate():
    if request.method == 'POST':
        code = request.form.get('candidate_code', '').strip()
        name = request.form.get('full_name', '').strip()
        email = request.form.get('email', '').strip().lower()
        phone = request.form.get('phone', '').strip()
        if not code or not name or not email or '@' not in email:
            flash('Candidate ID, name, and a valid email are required.', 'error')
            return render_template('candidate_form.html', candidate=None)
        try:
            with db() as con:
                cur = con.execute('''INSERT INTO candidates(candidate_code,full_name,email,phone,created_at)
                                     VALUES(?,?,?,?,?)''', (code, name, email, phone, datetime.now().isoformat(timespec='seconds')))
                cid = cur.lastrowid
                # Create a candidate login with generated password; admin can reset it by re-importing only after deleting the user.
                temp_password = secrets.token_urlsafe(9)
                username = email
                con.execute('INSERT INTO users(username,password_hash,role,candidate_id) VALUES(?,?,?,?)',
                            (username, generate_password_hash(temp_password), 'candidate', cid))
            sent = send_email(email, 'Candidate portal login',
                              f'Hello {name},\n\nYour candidate portal account has been created.\nUsername: {email}\nTemporary password: {temp_password}\nPlease change this password after your first login (password change feature must be enabled before production use).')
            flash('Candidate added. Login details emailed.' if sent else 'Candidate added. Email was not sent; configure SMTP and securely share login details.', 'success')
            return redirect(url_for('dashboard'))
        except sqlite3.IntegrityError:
            flash('Candidate ID or email already exists.', 'error')
    return render_template('candidate_form.html', candidate=None)


@app.route('/candidate/<int:cid>/edit', methods=['GET', 'POST'])
@login_required('admin')
def edit_candidate(cid):
    with db() as con:
        candidate = con.execute('SELECT * FROM candidates WHERE id=?', (cid,)).fetchone()
    if not candidate:
        abort(404)
    if request.method == 'POST':
        try:
            with db() as con:
                con.execute('''UPDATE candidates SET candidate_code=?,full_name=?,email=?,phone=?,assigned_task=?,task_status=?,
                               task_deadline=?,feedback=? WHERE id=?''',
                            (request.form.get('candidate_code','').strip(), request.form.get('full_name','').strip(),
                             request.form.get('email','').strip().lower(), request.form.get('phone','').strip(),
                             request.form.get('assigned_task','').strip(), request.form.get('task_status','Pending'),
                             request.form.get('task_deadline',''), request.form.get('feedback','').strip(), cid))
                con.execute('UPDATE users SET username=? WHERE candidate_id=?', (request.form.get('email','').strip().lower(), cid))
            flash('Candidate updated.', 'success')
            return redirect(url_for('dashboard'))
        except sqlite3.IntegrityError:
            flash('Candidate ID or email already belongs to another candidate.', 'error')
    return render_template('candidate_form.html', candidate=candidate)


@app.route('/candidate/<int:cid>/delete', methods=['POST'])
@login_required('admin')
def delete_candidate(cid):
    with db() as con:
        con.execute('DELETE FROM users WHERE candidate_id=?', (cid,))
        con.execute('DELETE FROM candidates WHERE id=?', (cid,))
    flash('Candidate deleted.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/candidate/<int:cid>/task', methods=['POST'])
@login_required('admin')
def update_task(cid):
    with db() as con:
        c = con.execute('SELECT * FROM candidates WHERE id=?', (cid,)).fetchone()
        if not c: abort(404)
        task = request.form.get('assigned_task', '').strip()
        status = request.form.get('task_status', 'Pending')
        deadline = request.form.get('task_deadline', '')
        con.execute('UPDATE candidates SET assigned_task=?,task_status=?,task_deadline=? WHERE id=?', (task, status, deadline, cid))
    flash('Task details saved.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/candidate/<int:cid>/schedule', methods=['GET', 'POST'])
@login_required('admin')
def schedule(cid):
    with db() as con:
        candidate = con.execute('SELECT * FROM candidates WHERE id=?', (cid,)).fetchone()
        reviews = con.execute('SELECT * FROM reviews WHERE candidate_id=? ORDER BY scheduled_at DESC', (cid,)).fetchall()
    if not candidate: abort(404)
    if request.method == 'POST':
        try:
            email, name, code, task, link, when = schedule_review(cid, request.form.get('scheduled_at', ''), request.form.get('meet_link', ''))
            body = (f'Hello {name},\n\nYour review has been scheduled.\nCandidate ID: {code}\nTask: {task or "Not assigned"}\n'
                    f'Date and time: {when.strftime("%d %B %Y, %I:%M %p")}\nGoogle Meet: {link or "The meeting link will be shared by the administrator."}\n\nPlease be ready to explain your completed task.')
            sent = send_email(email, 'Your candidate review is scheduled', body)
            flash('Review scheduled and email sent.' if sent else 'Review scheduled. Email not sent; check SMTP configuration.', 'success')
            return redirect(url_for('schedule', cid=cid))
        except ValueError as e:
            flash(str(e), 'error')
    return render_template('schedule.html', candidate=candidate, reviews=reviews)


@app.route('/review/<int:rid>/complete', methods=['POST'])
@login_required('admin')
def complete_review(rid):
    feedback = request.form.get('feedback', '').strip()
    result = request.form.get('result', 'Completed')
    with db() as con:
        review = con.execute('''SELECT r.*, c.full_name,c.email,c.candidate_code,c.assigned_task FROM reviews r
                                JOIN candidates c ON c.id=r.candidate_id WHERE r.id=?''', (rid,)).fetchone()
        if not review: abort(404)
        status = 'Absent' if result == 'Absent' else 'Completed'
        con.execute('UPDATE reviews SET status=?,feedback=? WHERE id=?', (status, feedback, rid))
        if status == 'Absent':
            con.execute('UPDATE candidates SET attendance_status="Absent",review_status="Waiting",feedback=? WHERE id=?', (feedback, review['candidate_id']))
        else:
            con.execute('''UPDATE candidates SET attendance_status='Attended',review_status='Completed',last_review=?,feedback=?
                           WHERE id=?''', (review['scheduled_at'], feedback, review['candidate_id']))
    body = (f'Hello {review["full_name"]},\n\nReview status: {status}\nTask: {review["task_snapshot"] or review["assigned_task"] or "Not assigned"}\n'
            f'Reviewer feedback: {feedback or "No feedback entered."}\n\nPlease sign in to the portal to check your next review schedule.')
    sent = send_email(review['email'], 'Candidate review update', body)
    flash(f'Review marked {status.lower()}.' + ('' if sent else ' Email not sent; check SMTP configuration.'), 'success')
    return redirect(url_for('dashboard'))


@app.route('/candidate/home')
@login_required('candidate')
def candidate_home():
    candidate = candidate_for_user()
    with db() as con:
        reviews = con.execute('SELECT * FROM reviews WHERE candidate_id=? ORDER BY scheduled_at DESC', (candidate['id'],)).fetchall()
    return render_template('candidate_home.html', candidate=candidate, reviews=reviews)


@app.route('/candidate/review-summary')
@login_required('candidate')
def review_summary():
    candidate = candidate_for_user()
    with db() as con:
        next_review = con.execute("SELECT * FROM reviews WHERE candidate_id=? AND status='Scheduled' ORDER BY scheduled_at LIMIT 1", (candidate['id'],)).fetchone()
    return render_template('review_summary.html', candidate=candidate, next_review=next_review)


@app.route('/admin/export.xlsx')
@login_required('admin')
def export_excel():
    with db() as con:
        rows = con.execute('''SELECT candidate_code,full_name,email,phone,assigned_task,task_status,attendance_status,
                              last_review,next_review,task_deadline,meet_link,feedback,review_status FROM candidates ORDER BY full_name''').fetchall()
    wb = Workbook(); ws = wb.active; ws.title = 'Candidates'
    headers = ['Candidate ID','Full Name','Email Address','Phone Number','Assigned Task','Task Status','Review Attendance Status','Last Review Date','Next Review Date and Time','Task Deadline','Google Meet Link','Reviewer Feedback','Review Status']
    ws.append(headers)
    for row in rows: ws.append(list(row))
    for col in ws.columns:
        max_len = min(max(len(str(cell.value or '')) for cell in col) + 2, 45)
        ws.column_dimensions[col[0].column_letter].width = max_len
    output = io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output, as_attachment=True, download_name='candidate_review_export.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/import.xlsx', methods=['POST'])
@login_required('admin')
def import_excel():
    file = request.files.get('file')
    if not file or not file.filename.lower().endswith('.xlsx'):
        flash('Choose an .xlsx Excel file.', 'error'); return redirect(url_for('dashboard'))
    try:
        wb = load_workbook(file, read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        headers = [str(x).strip().lower() if x is not None else '' for x in next(rows, [])]
        aliases = {'candidate id':'candidate_code','candidate_id':'candidate_code','full name':'full_name','name':'full_name','email address':'email','email':'email','phone number':'phone','phone':'phone','assigned task':'assigned_task','task status':'task_status','review attendance status':'attendance_status','last review date':'last_review','next review date and time':'next_review','task deadline':'task_deadline','google meet link':'meet_link','reviewer feedback':'feedback','review status':'review_status'}
        mapped = {aliases[h]: i for i,h in enumerate(headers) if h in aliases}
        if not {'candidate_code','full_name','email'}.issubset(mapped):
            flash('Excel must contain Candidate ID, Full Name, and Email Address columns.', 'error'); return redirect(url_for('dashboard'))
        added = updated = skipped = 0
        for row in rows:
            def val(key, default=''):
                i = mapped.get(key)
                return str(row[i]).strip() if i is not None and i < len(row) and row[i] is not None else default
            code, name, email = val('candidate_code'), val('full_name'), val('email').lower()
            if not code or not name or '@' not in email:
                skipped += 1; continue
            with db() as con:
                existing = con.execute('SELECT id FROM candidates WHERE candidate_code=? OR email=?', (code,email)).fetchone()
                fields = [val('phone'),val('assigned_task'),val('task_status','Pending'),val('attendance_status','Not scheduled'),val('last_review'),val('next_review'),val('task_deadline'),val('meet_link'),val('feedback'),val('review_status','Waiting')]
                if existing:
                    con.execute('''UPDATE candidates SET candidate_code=?,full_name=?,email=?,phone=?,assigned_task=?,task_status=?,attendance_status=?,last_review=?,next_review=?,task_deadline=?,meet_link=?,feedback=?,review_status=? WHERE id=?''',
                                (code,name,email,*fields,existing['id']))
                    updated += 1
                else:
                    cur=con.execute('''INSERT INTO candidates(candidate_code,full_name,email,phone,assigned_task,task_status,attendance_status,last_review,next_review,task_deadline,meet_link,feedback,review_status,created_at)
                                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(code,name,email,*fields,datetime.now().isoformat(timespec='seconds')))
                    cid=cur.lastrowid
                    # Imported candidate login is created with a random password; credentials are not emailed in bulk.
                    con.execute('INSERT OR IGNORE INTO users(username,password_hash,role,candidate_id) VALUES(?,?,?,?)',(email,generate_password_hash(secrets.token_urlsafe(16)),'candidate',cid))
                    added += 1
        flash(f'Import finished: {added} added, {updated} updated, {skipped} skipped.', 'success')
    except Exception:
        app.logger.exception('Excel import failed')
        flash('Could not read the Excel file. Check its format and column headers.', 'error')
    return redirect(url_for('dashboard'))


@app.errorhandler(403)
def forbidden(_):
    return render_template('error.html', code=403, message='You do not have permission to view this page.'), 403

@app.errorhandler(404)
def not_found(_):
    return render_template('error.html', code=404, message='The requested page was not found.'), 404

init_db()
if __name__ == '__main__':
    app.run(debug=os.getenv('FLASK_DEBUG', 'false').lower() == 'true')
