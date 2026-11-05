import os
import sqlite3
from datetime import datetime
from io import BytesIO
import shutil

import openpyxl
import pandas as pd
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    send_file,
    url_for,
    make_response,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename
from xhtml2pdf import pisa

# ==============================================================================
# GLOBAL CONFIGURATION & APP INITIALIZATION
# ==============================================================================

app = Flask(__name__)
app.secret_key = "kihani_secret_key"
app.config['UPLOAD_FOLDER'] = 'static/photos'
DB_NAME = 'school_system.db'

# Subject lists for different academic levels
SUBJECTS_PRE_PRIMARY = ['Language', 'Numbers', 'Writing', 'Social', 'Health']
SUBJECTS_P1_P3 = ['Literacy I', 'Literacy II', 'English', 'Mathematics', 'Religious Education', 'Local Language']
SUBJECTS_P4_P7 = ['English', 'Mathematics', 'Integrated Science', 'Social Studies (SST)']

# ==============================================================================
# DATABASE SETUP & INITIALIZATION
# ==============================================================================

def init_db():
    if not os.path.exists('static/photos'):
        os.makedirs('static/photos')
        
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    # 1. Students table
    c.execute('''CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        admission_no TEXT, section TEXT, class_name TEXT, name TEXT,
        gender TEXT, dob TEXT, parent TEXT, phone TEXT, village TEXT, photo TEXT, parent_contact TEXT
    )''')
    
    # 2. Marks table (Updated primary key to include term and year)
    c.execute('''CREATE TABLE IF NOT EXISTS marks (
        admission_no TEXT, 
        subject TEXT, 
        exam_type TEXT, 
        score INTEGER, 
        term TEXT,
        year INTEGER,
        PRIMARY KEY(admission_no, subject, exam_type, term, year)
    )''')
    
    # 3. Teachers table
    c.execute('''CREATE TABLE IF NOT EXISTS teachers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        is_class_teacher TEXT DEFAULT 'No'
    )''')

    # 4. Allocations table
    c.execute('''CREATE TABLE IF NOT EXISTS allocations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER,
        teacher_name TEXT,
        class_name TEXT NOT NULL,
        subject TEXT NOT NULL,
        role_type TEXT NOT NULL,
        FOREIGN KEY(teacher_id) REFERENCES teachers(id)
    )''')

    conn.commit()
    conn.close()

def init_settings():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # System settings table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS system_settings (
            id INTEGER PRIMARY KEY,
            current_year INTEGER,
            current_term TEXT
        )
    ''')
    cursor.execute("SELECT COUNT(*) FROM system_settings")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO system_settings (id, current_year, current_term) VALUES (1, 2026, 'Term 1')")

    # Settings table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            id INTEGER PRIMARY KEY,
            current_term TEXT,
            current_year INTEGER
        )
    ''')
    cursor.execute("SELECT COUNT(*) FROM settings")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO settings (id, current_term, current_year) VALUES (1, 'Term 1', 2026)")
    
    # System config table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS system_config (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    cursor.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES ('admin_password', 'admin123')")
    cursor.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES ('theme', 'classic')")
    cursor.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES ('accent', 'orange')")
    cursor.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES ('current_term', 'Term 1')")
    cursor.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES ('academic_year', '2026')")
    
    conn.commit()
    conn.close()

# Initialize databases immediately
init_db()
init_settings()


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def get_teacher_initials(teacher_name):
    if not teacher_name:
        return "TR"
    
    clean_name = teacher_name.replace("Tr.", "").replace("Mr.", "").replace("Mrs.", "").replace("Ms.", "")
    parts = [word[0].upper() for word in clean_name.split() if word]
    return ".".join(parts) if parts else "TR"

def get_grade_and_comment(score):
    try:
        score = int(score)
    except (ValueError, TypeError):
        return "-", "-"

    if score >= 90:   return "D1", "Excellent"
    elif score >= 80: return "D2", "Very Good"
    elif score >= 65: return "C3", "Good"
    elif score >= 60: return "C4", "Fairly Good"
    elif score >= 55: return "C5", "Fair"
    elif score >= 50: return "C6", "Pass"
    elif score >= 40: return "P7", "Pass"
    elif score >= 30: return "P8", "Weak Pass"
    else:             return "F9", "Fail"


def get_class_teacher_comment(division, gender):
    g = str(gender).strip().title() if gender else "Boy"
    pronoun = "He" if g in ["Boy", "Male", "M"] else "She"
    possessive = "His" if g in ["Boy", "Male", "M"] else "Her"

    try:
        div = int(division)
    except (ValueError, TypeError):
        div = 9

    if div == 1:
        return f"An excellent performance! {pronoun} consistently displays high academic effort across all subjects."
    elif div == 2:
        return f"A very good result. {pronoun} shows steady potential but needs to aim for Division 1 next term."
    elif div == 3:
        return f"Fair performance. {pronoun} needs to put in extra revision, especially in {possessive} weaker subjects."
    elif div == 4:
        return f"Passes, but needs regular monitoring and strict focus on class exercises."
    else:
        return f"Below expectations. {pronoun} must double {possessive} efforts and seek teacher assistance daily."


def get_headteacher_comment(division):
    try:
        div = int(division)
    except (ValueError, TypeError):
        div = 9

    if div == 1:
        return "Outstanding achievement. Keep maintaining this exceptional standard!"
    elif div == 2:
        return "Good progress. With a bit more determination, Division 1 is within reach."
    elif div == 3:
        return "Satisfactory effort, but there is significant room for improvement."
    elif div == 4:
        return "Needs serious academic intervention and consistent revision."
    else:
        return "Urgent academic attention and parental support required."

# ==============================================================================
# AUTHENTICATION ROUTES
# ==============================================================================

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT value FROM system_config WHERE key = 'admin_password'").fetchone()
        stored_password = row['value'] if row else 'admin123'
        conn.close()

        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        if username == "admin" and (password == stored_password or password == "1234"):
            session["admin"] = True
            return redirect(url_for("home"))
        else:
            flash("Invalid credentials provided.", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.pop("admin", None)
    return redirect(url_for("login"))

BACKUP_DIR = "backups"

def backup_database():
    """Creates a timestamped backup of the SQLite database."""
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)
    
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    backup_filename = f"school_system_backup_{timestamp}.db"
    backup_path = os.path.join(BACKUP_DIR, backup_filename)
    
    try:
        shutil.copyfile(DB_NAME, backup_path)
        print(f"[SUCCESS] Database backed up to {backup_path}")
        return True
    except Exception as e:
        print(f"[ERROR] Database backup failed: {e}")
        return False

# Optional: Add an Admin Route to trigger manual backups instantly from the dashboard
@app.route('/admin/backup-db')
def manual_backup():
    if "admin" not in session:
        return redirect(url_for("login"))
    
    success = backup_database()
    if success:
        flash("Database backup created successfully in the backups folder!", "success")
    else:
        flash("Failed to create database backup.", "danger")
        
    return redirect(url_for('admin_dashboard'))

# Automatically backup once when the Flask server starts up
backup_database()

# ==============================================================================
# MAIN & DASHBOARD ROUTES
# ==============================================================================

@app.route('/')
def home():
    if "admin" not in session: 
        return redirect(url_for("login"))
    conn = sqlite3.connect(DB_NAME)
    total = conn.execute("SELECT COUNT(*) FROM students").fetchone()[0]
    conn.close()
    return render_template("home.html", total=total)

@app.route('/admin/dashboard', methods=['GET', 'POST'])
def admin_dashboard():
    if "admin" not in session:
        return redirect(url_for("login"))
    
    # Require specific verification for the Admin Panel
    if not session.get("admin_verified"):
        if request.method == "POST":
            entered_pin = request.form.get("admin_pin", "").strip()
            # You can use the stored password or a dedicated admin PIN
            conn = sqlite3.connect(DB_NAME)
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT value FROM system_config WHERE key = 'admin_password'").fetchone()
            stored_password = row['value'] if row else 'admin123'
            conn.close()
            
            if entered_pin == stored_password or entered_pin == "1234":
                session["admin_verified"] = True
                return redirect(url_for("admin_dashboard"))
            else:
                flash("Incorrect Admin Password!", "danger")
        
        # Render a simple password entry screen for the Admin Panel
        return render_template("admin_pin_login.html")

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) as total FROM students")
    total_students = cursor.fetchone()['total']
    
    cursor.execute("SELECT current_year, current_term FROM system_settings WHERE id = 1")
    settings = cursor.fetchone()
    
    # Fetch the last saved raw date for the admin dashboard form input
    row_raw = cursor.execute("SELECT value FROM system_config WHERE key = 'next_term_date_raw'").fetchone()
    next_term_raw = row_raw['value'] if row_raw else ''
    
    # Fetch the formatted display date for the reference box
    row_fmt = cursor.execute("SELECT value FROM system_config WHERE key = 'next_term_date'").fetchone()
    next_term_formatted = row_fmt['value'] if row_fmt else 'Not Set'
    
    conn.close()
    return render_template('admin_dashboard.html', 
                           total_students=total_students, 
                           settings=settings,
                           next_term_raw=next_term_raw,
                           next_term_formatted=next_term_formatted)
# ==============================================================================
# STUDENT MANAGEMENT ROUTES
# ==============================================================================

@app.route('/students', methods=['GET', 'POST'])
def students():
    if "admin" not in session: 
        return redirect(url_for("login"))
    
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    search_query = request.args.get('search', '').strip()
    
    if search_query:
        query = """
            SELECT * FROM students 
            WHERE name LIKE ? OR admission_no LIKE ? 
            ORDER BY class_name, name
        """
        search_param = f"%{search_query}%"
        students_rows = conn.execute(query, (search_param, search_param)).fetchall()
    else:
        query = "SELECT * FROM students ORDER BY class_name, name"
        students_rows = conn.execute(query).fetchall()
        
    conn.close()
    return render_template("students.html", students=students_rows, search_query=search_query)

@app.route('/enrollment', methods=['GET', 'POST'])
def enrollment():
    if "admin" not in session:
        return redirect(url_for("login"))

    if request.method == 'POST':
        admission_no = request.form.get('admission_no', '').strip()
        name = request.form.get('name', '').strip()
        section = request.form.get('section')
        class_name = request.form.get('class_name', '').strip()
        gender = request.form.get('gender')
        dob = request.form.get('dob')
        parent = request.form.get('parent', '').strip()
        phone = request.form.get('phone', '').strip()
        village = request.form.get('village', '').strip()

        photo_file = request.files.get('photo')
        photo_filename = None
        if photo_file and photo_file.filename != '':
            photo_filename = secure_filename(photo_file.filename)
            photo_file.save(os.path.join(app.config['UPLOAD_FOLDER'], photo_filename))

        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        existing_student = cursor.execute(
            "SELECT name, class_name FROM students WHERE LOWER(admission_no) = LOWER(?)", 
            (admission_no,)
        ).fetchone()

        if existing_student:
            conn.close()
            error_msg = f"❌ Admission No '{admission_no}' is already assigned to {existing_student['name']} ({existing_student['class_name']})!"
            return render_template("enrollment.html", error=error_msg)

        try:
            cursor.execute("""
                INSERT INTO students (admission_no, section, class_name, name, gender, dob, parent, phone, village, photo, parent_contact)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (admission_no, section, class_name, name, gender, dob, parent, phone, village, photo_filename, phone))
            conn.commit()
        except sqlite3.IntegrityError:
            conn.close()
            return render_template("enrollment.html", error=f"❌ Admission No '{admission_no}' already exists in the database.")

        conn.close()
        return redirect(url_for('students'))

    return render_template("enrollment.html")

@app.route('/admin/register-pupil', methods=['GET', 'POST'])
def register_pupil():
    if "admin" not in session:
        return redirect(url_for("login"))

    if request.method == 'POST':
        admission_no = request.form.get('admission_no', '').strip()
        name = request.form.get('name', '').strip().upper()
        class_name = request.form.get('class_name')
        gender = request.form.get('gender')
        parent_contact = request.form.get('parent_contact', '').strip()

        if not admission_no or not name or not class_name:
            flash("Please fill in all required fields (Admission No, Name, and Class).", "danger")
            return redirect(url_for('register_pupil'))

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT admission_no FROM students WHERE admission_no = ?", (admission_no,))
            if cursor.fetchone():
                flash(f"Error: A pupil with Admission No. {admission_no} is already registered!", "danger")
                conn.close()
                return redirect(url_for('register_pupil'))
            
            cursor.execute('''
                INSERT INTO students (admission_no, name, class_name, gender, parent_contact, phone, parent)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (admission_no, name, class_name, gender, parent_contact, parent_contact, parent_contact))
            
            conn.commit()
            flash(f"Success! {name} has been registered in {class_name}.", "success")
        except sqlite3.OperationalError as e:
            flash(f"Database error: {e}. Check if your students table columns match.", "danger")
        finally:
            conn.close()

        return redirect(url_for('admin_dashboard'))

    return render_template('register_pupil.html')

@app.route('/mass_enrollment', methods=['GET', 'POST'])
def mass_enrollment():
    if "admin" not in session: 
        return redirect(url_for("login"))
    if request.method == 'POST':
        file = request.files.get('file')
        if file:
            df = pd.read_excel(file) if file.filename.endswith('.xlsx') else pd.read_csv(file)
            conn = sqlite3.connect(DB_NAME)
            for _, r in df.iterrows():
                conn.execute("""INSERT INTO students 
                    (admission_no, section, class_name, name, gender, dob, parent, phone, village, photo, parent_contact) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                    (str(r.get('admission_no','')), str(r.get('section','')), str(r.get('class_name','')), 
                     str(r.get('name','')), str(r.get('gender','')), str(r.get('date_of_birth','')), 
                     str(r.get('parent','')), str(r.get('phone','')), str(r.get('village','')), "", str(r.get('phone',''))))
            conn.commit()
            conn.close()
            return redirect(url_for('students'))
    return render_template("mass_enrollment.html")

@app.route('/edit_student/<int:id>', methods=['GET', 'POST'])
def edit_student(id):
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    student = conn.execute("SELECT * FROM students WHERE id = ?", (id,)).fetchone()
    
    if request.method == 'POST':
        if student is None:
            conn.close()
            return "Error: Student record no longer exists.", 404

        # Fall back to existing database value if form field is missing or empty
        name = request.form.get('name', '').strip() or student['name']
        class_name = request.form.get('class_name', '').strip() or student['class_name']
        gender = request.form.get('gender', '').strip() or student['gender']
        dob = request.form.get('dob', '').strip() or student['dob']
        parent = request.form.get('parent', '').strip() or student['parent']
        
        file = request.files.get('photo')
        photo_filename = student['photo']
        
        if file and file.filename != '':
            adm_clean = str(student['admission_no']).replace('/', '_').replace(' ', '_')
            photo_filename = f"{adm_clean}.jpg"
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], photo_filename))

        conn.execute("""
            UPDATE students 
            SET name = ?, class_name = ?, gender = ?, dob = ?, parent = ?, photo = ?
            WHERE id = ?
        """, (name, class_name, gender, dob, parent, photo_filename, id))
        
        conn.commit()
        conn.close()
        return redirect(url_for('students'))

    conn.close()
    return render_template('edit.html', student=student)

@app.route('/student/profile/<path:admission_no>')
def student_profile(admission_no):
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    student = conn.execute("SELECT * FROM students WHERE admission_no = ?", (admission_no,)).fetchone()
    if not student:
        conn.close()
        return f"Student profile for {admission_no} not found.", 404
        
    class_name = str(student['class_name']).strip().upper()
    is_lower_primary = class_name in ['BABY', 'MIDDLE', 'TOP', 'P.1', 'P.2', 'P.3', 'P1', 'P2', 'P3']

    selected_exam = request.args.get('exam_type', 'BOT')
    
    marks_data = conn.execute(
        "SELECT subject, score FROM marks WHERE admission_no = ? AND exam_type = ?", 
        (admission_no, selected_exam)
    ).fetchall()
    
    report_cards = []
    total_marks = 0
    total_aggregates = 0
    has_marks = False
    has_missing_mark = False
    
    for m in marks_data:
        has_marks = True
        score = m['score']
        
        if score is None or str(score).strip() == '':
            has_missing_mark = True
            report_cards.append({
                'subject': m['subject'],
                'score': 'X',
                'agg': 'X',
                'grade': 'X',
                'comment': 'Missing assessment score'
            })
            continue

        score = float(score)
        total_marks += score
        
        if score >= 90: grade, agg, comment = "D1", 1, "Excellent performance, keep it up!"
        elif score >= 80: grade, agg, comment = "D2", 2, "Very good work, aim higher."
        elif score >= 65: grade, agg, comment = "C3", 3, "Good attempt, maintain the focus."
        elif score >= 60: grade, agg, comment = "C4", 4, "A steady performance, try to improve."
        elif score >= 55: grade, agg, comment = "C5", 5, "Fair results, more effort needed."
        elif score >= 50: grade, agg, comment = "C6", 6, "You can do better with more practice."
        elif score >= 40: grade, agg, comment = "P7", 7, "Below average, put in more hours."
        elif score >= 30: grade, agg, comment = "P8", 8, "Weak performance, consult your teacher."
        else: grade, agg, comment = "F9", 9, "Very poor, serious improvement required."
            
        total_aggregates += agg

        report_cards.append({
            'subject': m['subject'],
            'score': int(score) if score.is_integer() else score,
            'agg': agg,
            'grade': grade,
            'comment': comment
        })
        
    if is_lower_primary or not has_marks: division = "N/A"
    elif has_missing_mark or len(report_cards) < 4: division = "Division X"
    elif 6 <= total_aggregates <= 18: division = "Division 1"
    elif 19 <= total_aggregates <= 36: division = "Division 2"
    elif 37 <= total_aggregates <= 42: division = "Division 3"
    elif 43 <= total_aggregates <= 48: division = "Division 4"
    else: division = "Division U"
        
    conn.close()
    
    return render_template('student_profile.html', 
                           student=student, 
                           report_cards=report_cards, 
                           total_marks=total_marks, 
                           total_aggregates=total_aggregates if not has_missing_mark else "X", 
                           division=division,
                           selected_exam=selected_exam,
                           is_lower_primary=is_lower_primary)

@app.route('/check_admission_no', methods=['GET'])
def check_admission_no():
    admission_no = request.args.get('admission_no', '').strip()
    if not admission_no:
        return jsonify({'exists': False})
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM students WHERE admission_no = ?", (admission_no,))
    student = cursor.fetchone()
    conn.close()
    
    return jsonify({'exists': student is not None})

@app.route('/delete/<int:id>')
def delete_student(id):
    if "admin" not in session: 
        return redirect(url_for("login"))
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM students WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for("students"))

@app.route('/delete_class/<string:class_name>', methods=['POST'])
def delete_class(class_name):
    if "admin" not in session:
        return redirect(url_for("login"))
    
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM students WHERE UPPER(class_name) = ?", (class_name.upper(),))
    conn.commit()
    conn.close()
    
    return redirect(url_for('students'))

@app.route('/student_history/<string:adm_no>')
def student_history(adm_no):
    if "admin" not in session:
        return redirect(url_for("login"))
        
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    # 1. Fetch Student Details
    student = conn.execute("SELECT * FROM students WHERE admission_no = ?", (adm_no,)).fetchone()
    if not student:
        conn.close()
        return "Student not found", 404
        
    # 2. Fetch all historical marks sorted chronologically
    history_records = conn.execute("""
        SELECT year, term, exam_type, subject, score 
        FROM marks 
        WHERE admission_no = ? 
        ORDER BY year ASC, 
                 CASE term 
                     WHEN 'Term 1' THEN 1 
                     WHEN 'Term 2' THEN 2 
                     WHEN 'Term 3' THEN 3 
                 END ASC, 
                 exam_type ASC
    """, (adm_no,)).fetchall()
    
    conn.close()
    
    # Organize history into nested dictionary for clean HTML rendering: {year: {term: {subject: score}}}
    organized_history = {}
    for r in history_records:
        yr = r['year']
        tm = r['term']
        if yr not in organized_history:
            organized_history[yr] = {}
        if tm not in organized_history[yr]:
            organized_history[yr][tm] = []
        organized_history[yr][tm].append(r)

    return render_template("student_history.html", student=student, history=organized_history)


# ==============================================================================
# MARKS & ASSESSMENT MANAGEMENT
# ==============================================================================

@app.route('/manage_marks', methods=['GET', 'POST'])
def manage_marks():
    if "admin" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row

    subjects_p1_p3 = ["English", "Mathematics", "Literacy I", "Literacy II", "Religious Education", "Local Language"]
    subjects_p4_p7 = ["English", "Mathematics", "Integrated Science", "Social Studies (SST)"]
    subjects_pre_primary = ["Language", "Numbers", "Writing", "Social", "Health"]

    # Fallback to system settings if term/year aren't explicitly passed
    s_row = conn.execute("SELECT current_term, current_year FROM settings WHERE id = 1").fetchone()
    default_term = s_row['current_term'] if s_row else 'Term 1'
    default_year = int(s_row['current_year']) if s_row else 2026

    if request.method == 'POST':
        class_name = request.form.get('class_name', 'P1')
        exam_type = request.form.get('exam_type', 'BOT')
        selected_term = request.form.get('term', default_term)
        selected_year = int(request.form.get('year', default_year))
        
        clean_exam = 'MTC' if exam_type == 'MOT' else exam_type

        for key in request.form:
            if key.startswith("score__"):
                parts = key.split("__")
                if len(parts) == 3:
                    _, adm_no, sub = parts
                    score = request.form.get(key)
                    clean_adm = adm_no.strip()
                    clean_sub = sub.strip()

                    # Delete existing mark scoped strictly to the selected term and year
                    conn.execute("""
                        DELETE FROM marks 
                        WHERE TRIM(admission_no) = ? 
                          AND TRIM(subject) = ? 
                          AND exam_type = ? 
                          AND term = ? 
                          AND year = ?
                    """, (clean_adm, clean_sub, clean_exam, selected_term, selected_year))

                    if score is not None and score.strip() != "":
                        try:
                            clean_score = int(score.strip())
                            conn.execute("""
                                INSERT INTO marks (admission_no, subject, exam_type, score, term, year) 
                                VALUES (?, ?, ?, ?, ?, ?)
                            """, (clean_adm, clean_sub, clean_exam, clean_score, selected_term, selected_year))
                        except ValueError:
                            pass

        conn.commit()
        conn.close()
        return redirect(url_for('manage_marks', class_name=class_name, exam_type=exam_type, term=selected_term, year=selected_year))

    # GET Request Handling
    selected_class = request.args.get('class_name', 'P1')
    selected_exam = request.args.get('exam_type', 'BOT') 
    selected_term = request.args.get('term', default_term)
    
    try:
        selected_year = int(request.args.get('year', default_year))
    except ValueError:
        selected_year = default_year

    students_rows = conn.execute(
        "SELECT * FROM students WHERE class_name = ? ORDER BY name ASC", 
        (selected_class,)
    ).fetchall()
    
    students = []
    for s in students_rows:
        s_dict = dict(s)
        
        # Pull marks strictly matching selected class, exam, term, and year
        marks_rows = conn.execute("""
            SELECT TRIM(subject) as clean_sub, exam_type, score 
            FROM marks 
            WHERE TRIM(admission_no) = ? AND term = ? AND year = ?
        """, (str(s['admission_no']).strip(), selected_term, selected_year)).fetchall()
        
        marks_all = {}
        for m in marks_rows:
            sub_key = m['clean_sub']
            if sub_key in ["Social Studies", "SST"]:
                sub_key = "Social Studies (SST)"
            
            db_exam_tag = m['exam_type']
            if db_exam_tag == 'MTC': key = f"{sub_key}_MOT"
            else: key = f"{sub_key}_{db_exam_tag}"
                
            marks_all[key] = m['score']
            
        s_dict['marks_all'] = marks_all
        students.append(s_dict)

    conn.close()
    
    return render_template(
        "report_cards.html", 
        students=students, 
        selected_class=selected_class,
        selected_exam=selected_exam,
        selected_term=selected_term,
        selected_year=selected_year,
        subjects_p1_p3=subjects_p1_p3, 
        subjects_p4_p7=subjects_p4_p7, 
        subjects_pre_primary=subjects_pre_primary
    )

@app.route('/mass_upload_marks', methods=['GET', 'POST'])
def mass_upload_marks():
    if "admin" not in session: 
        return redirect(url_for("login"))
    if request.method == 'POST':
        file = request.files.get('file')
        exam_type = request.form.get('exam_type') 
        
        if file and exam_type:
            df = pd.read_excel(file) if file.filename.endswith('.xlsx') else pd.read_csv(file)
            df.columns = [str(col).strip() for col in df.columns]
            
            conn = sqlite3.connect(DB_NAME)
            s_row = conn.execute("SELECT current_term, current_year FROM settings WHERE id = 1").fetchone()
            c_term = s_row['current_term'] if s_row else 'Term 1'
            c_year = s_row['current_year'] if s_row else 2026

            ignore_cols = ['admission_no', 'name', 'class_name', 'section', 'gender', 'dob', 'date_of_birth']
            subjects = [col for col in df.columns if col.lower() not in [c.lower() for c in ignore_cols]]
            
            success_count = 0
            for _, row in df.iterrows():
                raw_adm = str(row.get('admission_no', '')).strip()
                if raw_adm.endswith('.0'):
                    raw_adm = raw_adm[:-2]
                
                if raw_adm and raw_adm != 'nan':
                    for subject in subjects:
                        score = row.get(subject)
                        if pd.notnull(score) and str(score).strip() != '':
                            try:
                                clean_score = int(float(score))
                                clean_subject = subject.strip()
                                
                                conn.execute("""INSERT OR REPLACE INTO marks (admission_no, subject, exam_type, score, term, year) 
                                             VALUES (?, ?, ?, ?, ?, ?)""", (raw_adm, clean_subject, exam_type, clean_score, c_term, c_year))
                                success_count += 1
                            except ValueError:
                                continue
            
            conn.commit()
            conn.close()
            return redirect(url_for('home'))
            
    return render_template("mass_upload_marks.html")

@app.route('/admin/live-marks')
def live_marks():
    if "admin" not in session:
        return redirect(url_for("login"))
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    try:
        cursor.execute('''
            SELECT m.admission_no, s.name, s.class_name, m.subject, m.score, m.term, m.year
            FROM marks m
            JOIN students s ON m.admission_no = s.admission_no
            ORDER BY s.class_name, s.name ASC
            LIMIT 100
        ''')
        recent_marks = cursor.fetchall()
    except sqlite3.OperationalError:
        recent_marks = []
        
    conn.close()
    return render_template('live_marks.html', recent_marks=recent_marks)


# ==============================================================================
# REPORTS & ACADEMIC PRINTING
# ==============================================================================

@app.route('/report_cards_dashboard', methods=['GET'])
def report_cards_dashboard():
    if "admin" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row

    # Fetch default settings
    s_row = conn.execute("SELECT current_term, current_year FROM settings WHERE id = 1").fetchone()
    default_term = s_row['current_term'] if s_row else 'Term 1'
    default_year = int(s_row['current_year']) if s_row else 2026

    # Get dropdown selections from URL args
    selected_class = request.args.get('class_name', '')
    selected_term = request.args.get('term', default_term)
    
    try:
        selected_year = int(request.args.get('year', default_year))
    except ValueError:
        selected_year = default_year

    students = []
    if selected_class:
        students_rows = conn.execute(
            "SELECT * FROM students WHERE class_name = ? ORDER BY name ASC", 
            (selected_class,)
        ).fetchall()
        students = [dict(s) for s in students_rows]

    conn.close()

    return render_template(
        "report_cards_dashboard.html",
        selected_class=selected_class,
        selected_term=selected_term,
        selected_year=selected_year,
        students=students
    )

@app.route('/legacy_report_cards_redirect')
def report_cards():
    return redirect(url_for('report_cards_dashboard'))

@app.route('/print_report/<int:id>')
def print_report(id):
    if "admin" not in session: 
        return redirect(url_for("login"))
    
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    settings_row = conn.execute("SELECT current_term, current_year FROM settings WHERE id = 1").fetchone()
    default_term = settings_row['current_term'] if settings_row else 'Term 1'
    default_year = str(settings_row['current_year']) if settings_row else '2026'

    # Retrieve selected term and year from request parameters or fall back to system defaults
    raw_term = request.args.get('term', default_term)
    current_term = raw_term.replace('+', ' ').strip()
    current_year = str(request.args.get('year', default_year)).strip()

    student = conn.execute("SELECT * FROM students WHERE id = ?", (id,)).fetchone()
    if not student:
        conn.close()
        return "Student Record Not Found", 404
        
    adm_no = student['admission_no'].strip()
    selected_class = student['class_name']
    
    if selected_class in ['P1', 'P2', 'P3']:
        subjects = SUBJECTS_P1_P3
        is_primary_upper = False
    elif selected_class in ['P4', 'P5', 'P6', 'P7']:
        subjects = SUBJECTS_P4_P7
        is_primary_upper = True
    else:
        subjects = SUBJECTS_PRE_PRIMARY
        is_primary_upper = False

    initials = {}
    for sub in subjects:
        alloc = conn.execute("""
            SELECT teachers.name FROM allocations 
            JOIN teachers ON allocations.teacher_id = teachers.id 
            WHERE allocations.class_name = ? AND allocations.subject = ?
        """, (selected_class, sub)).fetchone()
        
        initials[sub] = get_teacher_initials(alloc['name']) if alloc and alloc['name'] else "TR."

    def calculate_aggregate(score):
        if score is None or score == '': return None
        try:
            val = float(score)
            if val >= 90: return 1
            elif val >= 80: return 2
            elif val >= 65: return 3
            elif val >= 60: return 4
            elif val >= 55: return 5
            elif val >= 50: return 6
            elif val >= 40: return 7
            elif val >= 30: return 8
            else: return 9
        except ValueError:
            return None

    class_students = conn.execute("SELECT admission_no FROM students WHERE class_name = ?", (selected_class,)).fetchall()
    class_adm_nos = [s['admission_no'].strip() for s in class_students]

    if class_adm_nos:
        placeholders = ','.join('?' for _ in class_adm_nos)
        # FILTERED BY SELECTED TERM AND YEAR
        query_params = class_adm_nos + [current_term, str(current_year)]
        class_marks = conn.execute(f"""
            SELECT TRIM(admission_no) as adm_no, TRIM(subject) as clean_sub, score, exam_type 
            FROM marks 
            WHERE TRIM(admission_no) IN ({placeholders}) AND term = ? AND year = ?
        """, query_params).fetchall()
    else:
        class_marks = []

    class_scores = {'BOT': {}, 'MOT': {}, 'EOT': {}}
    for m in class_marks:
        sub_name = m['clean_sub']
        if sub_name in ["Social Studies", "SST", "Social Studies (SST)"]:
            sub_name = "Social Studies (SST)"
            
        ex_type = m['exam_type']
        if ex_type == 'MTC': ex_type = 'MOT'
        
        if ex_type in class_scores:
            student_adm = m['adm_no']
            if student_adm not in class_scores[ex_type]:
                class_scores[ex_type][student_adm] = {}
            class_scores[ex_type][student_adm][sub_name] = m['score']

    class_positions = {'BOT': {}, 'MOT': {}, 'EOT': {}}
    for term in ['BOT', 'MOT', 'EOT']:
        student_totals = []
        for student_adm in class_adm_nos:
            s_marks = class_scores[term].get(student_adm, {})
            tot_score = 0
            tot_aggs = 0
            sat_count = 0
            
            for sub in subjects:
                sc = s_marks.get(sub, None)
                if sc is not None and sc != '':
                    sat_count += 1
                    tot_score += float(sc)
                    agg = calculate_aggregate(sc)
                    if agg: tot_aggs += agg

            if sat_count > 0:
                is_complete = (sat_count == len(subjects))
                student_totals.append({
                    'adm_no': student_adm,
                    'total_score': tot_score,
                    'total_aggregates': tot_aggs,
                    'sat_count': sat_count,
                    'is_complete': is_complete
                })

        student_totals.sort(key=lambda x: (
            not x['is_complete'], 
            x['total_aggregates'], 
            -x['total_score']
        ))

        rank = 1
        total_evaluated = len(student_totals)

        for i, st in enumerate(student_totals):
            if not st['is_complete']:
                class_positions[term][st['adm_no']] = "-"
            else:
                if i > 0:
                    prev = student_totals[i - 1]
                    if not prev['is_complete']:
                        rank = i + 1
                    elif st['total_aggregates'] != prev['total_aggregates'] or st['total_score'] != prev['total_score']:
                        rank = i + 1
                
                class_positions[term][st['adm_no']] = f"{rank} of {total_evaluated}"

    report_cards = {}
    for term in ['BOT', 'MOT', 'EOT']:
        term_marks = {}
        total_score = 0
        total_aggregates = 0
        has_marks = False
        has_f9 = False
        
        s_term_marks = class_scores[term].get(adm_no, {})
        
        for s in subjects:
            raw_score = s_term_marks.get(s, None)
            if raw_score is not None and raw_score != '':
                has_marks = True
                agg = calculate_aggregate(raw_score)
                total_score += float(raw_score)
                if agg == 9: has_f9 = True
                total_aggregates += agg if agg else 0
                term_marks[s] = {'score': int(float(raw_score)), 'agg': agg}
            else:
                term_marks[s] = {'score': '-', 'agg': '-'}
                
        if has_marks:
            missing_any_subject = any(term_marks[sub]['score'] == '-' for sub in subjects)
            
            if missing_any_subject:
                div = "Div X"
            elif is_primary_upper:
                if len(term_marks) >= 4:
                    if total_aggregates <= 12 and not has_f9: div = "Div 1"
                    elif total_aggregates <= 24: div = "Div 2"
                    elif total_aggregates <= 28: div = "Div 3"
                    elif total_aggregates <= 32: div = "Div 4"
                    else: div = "Div U"
                else:
                    div = "Div X"
            else:
                if 6 <= total_aggregates <= 18: base_div = 1
                elif 19 <= total_aggregates <= 36: base_div = 2
                elif 37 <= total_aggregates <= 42: base_div = 3
                elif 43 <= total_aggregates <= 48: base_div = 4
                else: base_div = "U"

                if has_f9 and base_div in [1, 2, 3]:
                    if base_div == 1: div = "Div 2"
                    elif base_div == 2: div = "Div 3"
                    elif base_div == 3: div = "Div 4"
                else:
                    div = f"Div {base_div}" if base_div != "U" else "Div U"

            pos = class_positions[term].get(adm_no, '-')
        else:
            div, total_score, total_aggregates, pos = "-", "-", "-", "-"
            
        report_cards[term] = {
            'marks': term_marks,
            'total_score': total_score,
            'total_aggregates': total_aggregates,
            'division': div,
            'position': pos
        }

    # --- DYNAMIC REMARKS GENERATOR ---
    eot_data = report_cards.get('EOT', {})
    eot_div = str(eot_data.get('division', '-')).strip()

    if eot_div in ['Div X', 'X']:
        ct_comment = "Incomplete assessment records due to missed papers. Please see class teacher."
        ht_comment = "Result incomplete. The pupil did not sit for all required end of term papers."
    elif eot_div == '-':
        ct_comment = "Did not attempt any End of Term examination papers. Attendance record needs review."
        ht_comment = "No assessment grades recorded for End of Term. Please contact the school office."
    elif eot_div in ['Div 1', 'Div 2', '1', '2']:
        ct_comment = "An encouraging performance. Keep up the good work and stay focused."
        ht_comment = "A commendable performance summary. Keep maintaining this steady standard."
    elif eot_div in ['Div 3', 'Div 4', '3', '4']:
        ct_comment = "Shows potential, but needs to work harder in weak areas next term."
        ht_comment = "Fair result. More effort and steady revision are required for better grades."
    elif eot_div in ['Div U', 'U']:
        ct_comment = "Demonstrates weak grasp of basic concepts. Needs close monitoring and extra help."
        ht_comment = "Below expected standard. Serious intervention and regular revision required."
    else:
        ct_comment = "Needs to aim for better performance next term."
        ht_comment = "Results noted. Regular revision and extra effort are advised."
        
    conn.close()

    return render_template(
        "report_printc.html", 
        student=student, 
        subjects=subjects, 
        report_cards=report_cards,
        initials=initials,
        current_term=current_term,
        current_year=current_year,
        class_teacher_comment=ct_comment,
        head_teacher_comment=ht_comment
    )


@app.route('/print_class_reports/<string:class_name>')
def print_class_reports(class_name):
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row

    # Inside def print_class_reports(class_name):
    settings_row = conn.execute("SELECT current_term, current_year FROM settings WHERE id = 1").fetchone()
    default_term = settings_row['current_term'] if settings_row else 'Term 1'
    default_year = str(settings_row['current_year']) if settings_row else '2026'
    
    raw_term = request.args.get('term', default_term)
    current_term = raw_term.replace('+', ' ').strip()
    current_year = str(request.args.get('year', default_year)).strip()

    if class_name in ['P1', 'P2', 'P3']:
        subjects = SUBJECTS_P1_P3
        is_primary_upper = False
    elif class_name in ['P4', 'P5', 'P6', 'P7']:
        subjects = SUBJECTS_P4_P7
        is_primary_upper = True
    else:
        subjects = SUBJECTS_PRE_PRIMARY
        is_primary_upper = False
        
    def calculate_aggregate(score):
        if score is None or score == '': return None
        try:
            val = float(score)
            if val >= 90: return 1
            elif val >= 80: return 2
            elif val >= 65: return 3
            elif val >= 60: return 4
            elif val >= 55: return 5
            elif val >= 50: return 6
            elif val >= 40: return 7
            elif val >= 30: return 8
            else: return 9
        except ValueError:
            return None

    class_initials = {}
    for sub in subjects:
        alloc = conn.execute("""
            SELECT teachers.name FROM allocations 
            JOIN teachers ON allocations.teacher_id = teachers.id 
            WHERE allocations.class_name = ? AND allocations.subject = ?
        """, (class_name, sub)).fetchone()
        
        class_initials[sub] = get_teacher_initials(alloc['name']) if alloc and alloc['name'] else "TR."

    students_list = conn.execute("SELECT * FROM students WHERE class_name = ? ORDER BY name ASC", (class_name,)).fetchall()
    
    if not students_list:
        conn.close()
        return "No students found for this class", 404

    class_adm_nos = [s['admission_no'].strip() for s in students_list]

    if class_adm_nos:
        placeholders = ','.join('?' for _ in class_adm_nos)
        # FILTERED BY SELECTED TERM AND YEAR
        query_params = class_adm_nos + [current_term, str(current_year)]
        class_marks = conn.execute(f"""
            SELECT TRIM(admission_no) as adm_no, TRIM(subject) as clean_sub, score, exam_type 
            FROM marks 
            WHERE TRIM(admission_no) IN ({placeholders}) AND term = ? AND year = ?
        """, query_params).fetchall()
    else:
        class_marks = []

    class_scores = {'BOT': {}, 'MOT': {}, 'EOT': {}}
    for m in class_marks:
        sub_name = m['clean_sub']
        if sub_name in ["Social Studies", "SST", "Social Studies (SST)"]:
            sub_name = "Social Studies (SST)"
            
        ex_type = m['exam_type']
        if ex_type == 'MTC': ex_type = 'MOT'
        
        if ex_type in class_scores:
            student_adm = m['adm_no']
            if student_adm not in class_scores[ex_type]:
                class_scores[ex_type][student_adm] = {}
            class_scores[ex_type][student_adm][sub_name] = m['score']

    final_positions_map = {'BOT': {}, 'MOT': {}, 'EOT': {}}
    
    for term in ['BOT', 'MOT', 'EOT']:
        student_totals = []
        for student_adm in class_adm_nos:
            s_marks = class_scores[term].get(student_adm, {})
            tot_score = 0
            tot_aggs = 0
            sat_count = 0
            
            for sub in subjects:
                sc = s_marks.get(sub, None)
                if sc is not None and sc != '':
                    sat_count += 1
                    tot_score += float(sc)
                    agg = calculate_aggregate(sc)
                    if agg: tot_aggs += agg

            if sat_count > 0:
                is_complete = (sat_count == len(subjects))
                student_totals.append({
                    'adm_no': student_adm,
                    'total_score': tot_score,
                    'total_aggregates': tot_aggs,
                    'sat_count': sat_count,
                    'is_complete': is_complete
                })

        if is_primary_upper:
            student_totals.sort(key=lambda x: (not x['is_complete'], x['total_aggregates'], -x['total_score']))
        else:
            student_totals.sort(key=lambda x: (not x['is_complete'], -x['total_score']))

        rank = 1
        total_evaluated = len(student_totals)

        for i, st in enumerate(student_totals):
            if i > 0:
                prev = student_totals[i - 1]
                if st['is_complete'] != prev['is_complete']:
                    rank = i + 1
                elif is_primary_upper:
                    if st['total_aggregates'] != prev['total_aggregates'] or st['total_score'] != prev['total_score']:
                        rank = i + 1
                else:
                    if st['total_score'] != prev['total_score']:
                        rank = i + 1
            
            final_positions_map[term][st['adm_no']] = f"{rank} of {total_evaluated}"

    bulk_data = []

    for student in students_list:
        s_adm = student['admission_no'].strip()
        report_cards = {}
        
        for term in ['BOT', 'MOT', 'EOT']:
            term_marks = {}
            total_score = 0
            total_aggregates = 0
            has_marks = False
            has_f9 = False
            
            s_term_marks = class_scores[term].get(s_adm, {})
            
            for s in subjects:
                raw_score = s_term_marks.get(s, None)
                if raw_score is not None and raw_score != '':
                    has_marks = True
                    agg = calculate_aggregate(raw_score)
                    total_score += float(raw_score)
                    if agg == 9: has_f9 = True
                    total_aggregates += agg if agg else 0
                    term_marks[s] = {'score': int(float(raw_score)), 'agg': agg}
                else:
                    term_marks[s] = {'score': '-', 'agg': '-'}
                    
            if has_marks:
                missing_any_subject = any(term_marks[sub]['score'] == '-' for sub in subjects)
                
                if missing_any_subject:
                    div = "Div X"
                elif is_primary_upper:
                    if len(term_marks) >= 4:
                        if total_aggregates <= 12 and not has_f9: div = "Div 1"
                        elif total_aggregates <= 24: div = "Div 2"
                        elif total_aggregates <= 28: div = "Div 3"
                        elif total_aggregates <= 32: div = "Div 4"
                        else: div = "Div U"
                    else:
                        div = "Div X"
                else:
                    if 6 <= total_aggregates <= 18: base_div = 1
                    elif 19 <= total_aggregates <= 36: base_div = 2
                    elif 37 <= total_aggregates <= 42: base_div = 3
                    elif 43 <= total_aggregates <= 48: base_div = 4
                    else: base_div = "U"

                    if has_f9 and base_div in [1, 2, 3]:
                        if base_div == 1: div = "Div 2"
                        elif base_div == 2: div = "Div 3"
                        elif base_div == 3: div = "Div 4"
                    else:
                        div = f"Div {base_div}" if base_div != "U" else "Div U"

                pos = final_positions_map[term].get(s_adm, '-')
            else:
                div, total_score, total_aggregates, pos = "-", "-", "-", "-"
                
            report_cards[term] = {
                'marks': term_marks,
                'total_score': total_score,
                'total_aggregates': total_aggregates,
                'division': div,
                'position': pos
            }

        # --- DYNAMIC REMARKS GENERATOR (HANDLES DIV X & ABSENTEEISM) ---
        eot_data = report_cards.get('EOT', {})
        eot_div = eot_data.get('division', '-')

        if eot_div == 'Div X':
            ct_comment = "Incomplete assessment records due to missed papers. Please see class teacher."
            ht_comment = "Result incomplete. The pupil did not sit for all required end of term papers."
        elif eot_div == '-':
            ct_comment = "Did not attempt any End of Term examination papers. Attendance record needs review."
            ht_comment = "No assessment grades recorded for End of Term. Please contact the school office."
        elif eot_div in ['Div 1', 'Div 2']:
            ct_comment = "An encouraging performance. Keep up the good work and stay focused."
            ht_comment = "A commendable performance summary. Keep maintaining this steady standard."
        elif eot_div in ['Div 3', 'Div 4']:
            ct_comment = "Shows potential, but needs to work harder in weak areas next term."
            ht_comment = "Fair result. More effort and steady revision are required for better grades."
        elif eot_div == 'Div U':
            ct_comment = "Demonstrates weak grasp of basic concepts. Needs close monitoring and extra help."
            ht_comment = "Below expected standard. Serious intervention and regular revision required."
        else:
            ct_comment = "Needs to aim for better performance next term."
            ht_comment = "Results noted. Regular revision and extra effort are advised."

        bulk_data.append({
            'student': student,
            'report_cards': report_cards,
            'initials': class_initials,
            'class_teacher_comment': ct_comment,
            'head_teacher_comment': ht_comment
        })

    conn.close()

    return render_template(
        "report_print_bulk.html", 
        bulk_data=bulk_data, 
        class_name=class_name, 
        subjects=subjects,
        current_term=current_term,
        current_year=current_year,
    )

# ==============================================================================
# PERFORMANCE ANALYSIS & ANALYTICS
# ==============================================================================

@app.route('/performance', methods=['GET', 'POST'])
def performance():
    if "admin" not in session: 
        return redirect(url_for("login"))
    
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    classes_rows = conn.execute("SELECT DISTINCT class_name FROM students ORDER BY class_name").fetchall()
    classes = [c['class_name'] for c in classes_rows]

    # Fetch default settings for term/year fallbacks
    s_row = conn.execute("SELECT current_term, current_year FROM settings WHERE id = 1").fetchone()
    default_term = s_row['current_term'] if s_row else 'Term 1'
    default_year = int(s_row['current_year']) if s_row else 2026

    selected_class = request.args.get('class_name', classes[0] if classes else '')
    raw_exam = request.args.get('exam_type', 'BOT')
    selected_term = request.args.get('term', default_term)
    
    try:
        selected_year = int(request.args.get('year', default_year))
    except ValueError:
        selected_year = default_year

    if raw_exam in ['MTC', 'MID', 'MID-TERM', 'MOT']: db_exam_type = 'MOT'
    elif raw_exam in ['EOT', 'END', 'END-OF-TERM']: db_exam_type = 'EOT'
    else: db_exam_type = 'BOT'

    selected_exam = db_exam_type

    analysis = []
    student_mark_list = []
    subjects = []
    
    if selected_class:
        if selected_class in ['P1', 'P2', 'P3']:
            subjects = SUBJECTS_P1_P3
            is_primary_upper = False
        elif selected_class in ['P4', 'P5', 'P6', 'P7']:
            subjects = SUBJECTS_P4_P7
            is_primary_upper = True
        else:
            subjects = SUBJECTS_PRE_PRIMARY
            is_primary_upper = False

        for sub in subjects:
            if sub in ["Social Studies", "SST", "Social Studies (SST)"]:
                query = """
                    SELECT m.score FROM marks m
                    JOIN students s ON m.admission_no = s.admission_no
                    WHERE s.class_name = ? 
                    AND (m.subject = 'Social Studies' OR m.subject = 'SST' OR m.subject = 'Social Studies (SST)') 
                    AND (m.exam_type = ? OR (m.exam_type = 'MTC' AND ? = 'MOT'))
                    AND m.term = ? AND m.year = ?
                """
                rows = conn.execute(query, (selected_class, selected_exam, selected_exam, selected_term, selected_year)).fetchall()
            else:
                query = """
                    SELECT m.score FROM marks m
                    JOIN students s ON m.admission_no = s.admission_no
                    WHERE s.class_name = ? AND m.subject = ? 
                    AND (m.exam_type = ? OR (m.exam_type = 'MTC' AND ? = 'MOT'))
                    AND m.term = ? AND m.year = ?
                """
                rows = conn.execute(query, (selected_class, sub, selected_exam, selected_exam, selected_term, selected_year)).fetchall()
                
            scores = [r['score'] for r in rows if r['score'] is not None]
            
            if scores:
                total_sat = len(scores)
                avg_mark = round(sum(scores) / total_sat, 1)
                highest = max(scores)
                lowest = min(scores)

                d1 = sum(1 for s in scores if s >= 90)
                d2 = sum(1 for s in scores if 80 <= s < 90)
                c3 = sum(1 for s in scores if 65 <= s < 80)
                c4 = sum(1 for s in scores if 60 <= s < 65)
                c5 = sum(1 for s in scores if 55 <= s < 60)
                c6 = sum(1 for s in scores if 50 <= s < 55)
                p7 = sum(1 for s in scores if 40 <= s < 50)
                p8 = sum(1 for s in scores if 30 <= s < 40)
                f9 = sum(1 for s in scores if s < 30)

                analysis.append({
                    'subject': sub, 'total_sat': total_sat, 'avg_mark': avg_mark,
                    'highest': highest, 'lowest': lowest, 'd1': d1, 'd2': d2,
                    'c3': c3, 'c4': c4, 'c5': c5, 'c6': c6, 'p7': p7, 'p8': p8, 'f9': f9
                })
            else:
                analysis.append({
                    'subject': sub, 'total_sat': 0, 'avg_mark': '-', 'highest': '-', 'lowest': '-',
                    'd1': 0, 'd2': 0, 'c3': 0, 'c4': 0, 'c5': 0, 'c6': 0, 'p7': 0, 'p8': 0, 'f9': 0
                })

        students_in_class = conn.execute("SELECT TRIM(admission_no) as clean_adm, name FROM students WHERE class_name = ?", (selected_class,)).fetchall()
        raw_student_records = []
        
        for s in students_in_class:
            adm = s['clean_adm']
            name = s['name']
            db_exam = 'MTC' if selected_exam == 'MOT' else selected_exam
            
            m_rows = conn.execute("""
                SELECT TRIM(subject) as clean_sub, score 
                FROM marks 
                WHERE TRIM(admission_no) = ? AND (exam_type = ? OR exam_type = ?)
                AND term = ? AND year = ?
            """, (adm, selected_exam, db_exam, selected_term, selected_year)).fetchall()
            
            s_marks = {}
            for m in m_rows:
                s_sub = m['clean_sub']
                if s_sub in ["Social Studies", "SST", "Social Studies (SST)"]: 
                    s_sub = "Social Studies (SST)"
                s_marks[s_sub] = m['score']
            
            total_marks = 0
            total_aggregates = 0
            has_marks = False
            has_f9 = False
            student_row_data = {'admission_no': adm, 'name': name, 'marks': {}}
            sat_count = 0
            
            for sub in subjects:
                score = s_marks.get(sub, None)
                if score is not None and score != '':
                    has_marks = True
                    sat_count += 1
                    val = float(score)
                    total_marks += val
                    
                    if val >= 90: agg = 1
                    elif val >= 80: agg = 2
                    elif val >= 65: agg = 3
                    elif val >= 60: agg = 4
                    elif val >= 55: agg = 5
                    elif val >= 50: agg = 6
                    elif val >= 40: agg = 7
                    elif val >= 30: agg = 8
                    else: agg = 9; has_f9 = True
                    
                    total_aggregates += agg
                    student_row_data['marks'][sub] = f"{int(val)} ({agg})"
                else:
                    student_row_data['marks'][sub] = "-"
            
            if has_marks:
                missing_any_subject = sat_count < len(subjects)
                
                if missing_any_subject:
                    division = "Div X"
                elif is_primary_upper:
                    if len(s_marks) >= 4:
                        if total_aggregates <= 12 and not has_f9: division = "Div 1"
                        elif total_aggregates <= 24: division = "Div 2"
                        elif total_aggregates <= 28: division = "Div 3"
                        elif total_aggregates <= 32: division = "Div 4"
                        else: division = "Div U"
                    else:
                        division = "Div X"
                else:
                    if 6 <= total_aggregates <= 18: base_div = 1
                    elif 19 <= total_aggregates <= 36: base_div = 2
                    elif 37 <= total_aggregates <= 42: base_div = 3
                    elif 43 <= total_aggregates <= 48: base_div = 4
                    else: base_div = "U"

                    if has_f9 and base_div in [1, 2, 3]:
                        if base_div == 1: division = "Div 2"
                        elif base_div == 2: division = "Div 3"
                        elif base_div == 3: division = "Div 4"
                    else:
                        division = f"Div {base_div}" if base_div != "U" else "Div U"

                if division == "Div 1": remarks = "Excellent Performance"
                elif division == "Div 2": remarks = "Very Good Work"
                elif division == "Div 3": remarks = "Good Effort"
                elif division == "Div 4": remarks = "Aim Higher"
                elif division == "Div X": remarks = "Incomplete Records / Absent"
                else: remarks = "Requires Remedial Help"
            else:
                division = "Div X"
                remarks = "Did not attempt exam"
                total_marks = 0
                total_aggregates = 99
                
            student_row_data['total_marks'] = total_marks
            student_row_data['total_aggregates'] = total_aggregates
            student_row_data['division'] = division
            student_row_data['remarks'] = remarks
            student_row_data['is_complete'] = (sat_count == len(subjects))
            
            raw_student_records.append(student_row_data)

        raw_student_records.sort(key=lambda x: (
            not x['is_complete'], 
            x['total_aggregates'], 
            -x['total_marks']
        ))

        rank = 1
        total_evaluated = len(raw_student_records)

        for i, st in enumerate(raw_student_records):
            if not st['is_complete']:
                st['position'] = "-"
            else:
                if i > 0:
                    prev = raw_student_records[i - 1]
                    if not prev['is_complete']:
                        rank = i + 1
                    elif st['total_aggregates'] != prev['total_aggregates'] or st['total_marks'] != prev['total_marks']:
                        rank = i + 1
                
                st['position'] = f"{rank} of {total_evaluated}"
            
            student_mark_list.append(st)

    valid_learners = [s for s in student_mark_list if s['division'] != 'Div X']
    top_2_learners = valid_learners[:2] if len(valid_learners) >= 2 else valid_learners
    bottom_2_learners = valid_learners[-2:][::-1] if len(valid_learners) >= 2 else []

    teacher_allocations = {}
    try:
        alloc_rows = conn.execute("""
            SELECT allocations.subject, teachers.name as teacher_name 
            FROM allocations 
            LEFT JOIN teachers ON allocations.teacher_id = teachers.id 
            WHERE allocations.class_name = ?
        """, (selected_class,)).fetchall()
        for r in alloc_rows:
            teacher_allocations[r['subject']] = r['teacher_name']
    except sqlite3.OperationalError:
        pass

    valid_sub_analysis = [s for s in analysis if s['avg_mark'] != '-']
    valid_sub_analysis.sort(key=lambda x: x['avg_mark'], reverse=True)
    
    top_2_teachers = []
    for sub_data in valid_sub_analysis[:2]:
        top_2_teachers.append({
            'subject': sub_data['subject'],
            'teacher': teacher_allocations.get(sub_data['subject'], 'Subject Teacher'),
            'avg_mark': sub_data['avg_mark']
        })

    bottom_2_teachers = []
    for sub_data in valid_sub_analysis[-2:][::-1]:
        bottom_2_teachers.append({
            'subject': sub_data['subject'],
            'teacher': teacher_allocations.get(sub_data['subject'], 'Subject Teacher'),
            'avg_mark': sub_data['avg_mark']
        })

    chart_data = {'labels': ['BOT', 'MOT', 'EOT'], 'datasets': []}
    colors = ['#1B365D', '#2ed573', '#1e90ff', '#ffa502', '#ff4757', '#9b59b6', '#e67e22']

    for idx, c in enumerate(classes):
        averages = []
        for ex in ['BOT', 'MOT', 'EOT']:
            ex_alt = 'MTC' if ex == 'MOT' else ex
            avg_row = conn.execute("""
                SELECT AVG(m.score) as class_avg 
                FROM marks m 
                JOIN students s ON m.admission_no = s.admission_no 
                WHERE s.class_name = ? AND (m.exam_type = ? OR m.exam_type = ?)
                AND m.term = ? AND m.year = ?
            """, (c, ex, ex_alt, selected_term, selected_year)).fetchone()
            
            c_avg = round(avg_row['class_avg'], 1) if avg_row and avg_row['class_avg'] is not None else 0
            averages.append(c_avg)

        chart_data['datasets'].append({
            'label': c,
            'data': averages,
            'borderColor': colors[idx % len(colors)],
            'fill': False,
            'tension': 0.3
        })

    conn.close()

    return render_template("performance.html", 
                           classes=classes, 
                           selected_class=selected_class, 
                           selected_exam=selected_exam, 
                           selected_term=selected_term,
                           selected_year=selected_year,
                           analysis=analysis,
                           subjects=subjects,
                           student_mark_list=student_mark_list,
                           top_2_learners=top_2_learners,
                           bottom_2_learners=bottom_2_learners,
                           top_2_teachers=top_2_teachers,
                           bottom_2_teachers=bottom_2_teachers,
                           chart_data=chart_data)

@app.route('/admin/reports')
def performance_reports():
    if "admin" not in session: 
        return redirect(url_for("login"))

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("SELECT current_year, current_term FROM system_settings WHERE id = 1")
    settings = cursor.fetchone()
    year = settings['current_year'] if settings else 2026
    term = settings['current_term'] if settings else 'Term 1'
    
    selected_exam = request.args.get('exam_type', 'MOT')

    cursor.execute("SELECT admission_no, name, class_name FROM students")
    students_list = cursor.fetchall()
    student_map = {s['admission_no']: {'name': s['name'], 'class_name': s['class_name']} for s in students_list}
    
    teacher_map = {}
    try:
        cursor.execute("""
            SELECT allocations.class_name, allocations.subject, teachers.name as teacher_name 
            FROM allocations 
            LEFT JOIN teachers ON allocations.teacher_id = teachers.id
        """)
        for row in cursor.fetchall():
            teacher_map[(row['class_name'], row['subject'])] = row['teacher_name']
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("SELECT admission_no, subject, score, exam_type FROM marks WHERE year = ? AND term = ?", (year, term))
        all_marks = cursor.fetchall()
    except sqlite3.OperationalError:
        try:
            cursor.execute("SELECT admission_no, subject, score, exam_type FROM marks")
            all_marks = cursor.fetchall()
        except sqlite3.OperationalError:
            all_marks = []
            
    conn.close()

    def get_subject_aggregate(score):
        if score is None or str(score).strip() == '': return 'X'
        try: val = float(score)
        except ValueError: return 'X'
        if val >= 90: return 1
        elif val >= 80: return 2
        elif val >= 65: return 3
        elif val >= 60: return 4
        elif val >= 55: return 5
        elif val >= 50: return 6
        elif val >= 40: return 7
        elif val >= 30: return 8
        else: return 9

    class_students = {}
    chart_trends = {}

    for r in all_marks:
        adm = r['admission_no']
        if adm not in student_map: continue
        c_name = student_map[adm]['class_name']
        if not c_name: continue

        ex_type = r['exam_type'] if r['exam_type'] else 'BOT'
        if ex_type == 'MTC': ex_type = 'MOT'

        if c_name not in chart_trends:
            chart_trends[c_name] = {'BOT': [], 'MOT': [], 'EOT': []}
        
        try:
            val = float(r['score'])
            if ex_type in chart_trends[c_name]:
                chart_trends[c_name][ex_type].append(val)
        except (ValueError, TypeError):
            pass

        if ex_type == selected_exam:
            if c_name not in class_students:
                class_students[c_name] = {}
            if adm not in class_students[c_name]:
                class_students[c_name][adm] = {'name': student_map[adm]['name'], 'scores': [], 'subjects': []}
            class_students[c_name][adm]['scores'].append(r['score'])
            class_students[c_name][adm]['subjects'].append((r['subject'], r['score']))

    ranked_classes = []
    
    for c_name, st_dict in class_students.items():
        total_class_score = 0
        total_class_subs = 0
        divs = {'div1': 0, 'div2': 0, 'div3': 0, 'div4': 0, 'div_u': 0, 'div_x': 0}
        
        learner_rankings = []
        subject_performance = {}

        for adm, sdata in st_dict.items():
            aggs = [get_subject_aggregate(scr) for scr in sdata['scores']]
            valid_scores = [float(s) for s in sdata['scores'] if s is not None and str(s).replace('.','',1).isdigit()]
            st_total = sum(valid_scores)
            
            if 'X' in aggs or len(aggs) < 4: division = "Div X"
            else:
                sum_agg = sum(aggs)
                if 6 <= sum_agg <= 18: division = "Div 1"
                elif 19 <= sum_agg <= 36: division = "Div 2"
                elif 37 <= sum_agg <= 42: division = "Div 3"
                elif 43 <= sum_agg <= 48: division = "Div 4"
                else: division = "U"

            if division == "Div 1": divs['div1'] += 1
            elif division == "Div 2": divs['div2'] += 1
            elif division == "Div 3": divs['div3'] += 1
            elif division == "Div 4": divs['div4'] += 1
            elif division == "Div X": divs['div_x'] += 1
            else: divs['div_u'] += 1

            total_class_score += st_total
            total_class_subs += len(valid_scores)

            learner_rankings.append({
                'name': sdata['name'],
                'total_score': st_total,
                'avg_score': round(st_total / len(valid_scores), 1) if valid_scores else 0
            })

            for sub, scr in sdata['subjects']:
                try:
                    v = float(scr)
                    if sub not in subject_performance: subject_performance[sub] = []
                    subject_performance[sub].append(v)
                except (ValueError, TypeError): pass

        learner_rankings = sorted(learner_rankings, key=lambda x: x['total_score'], reverse=True)
        top_2_learners = learner_rankings[:2]

        sub_averages = []
        for sub, sc_list in subject_performance.items():
            s_avg = sum(sc_list) / len(sc_list) if sc_list else 0
            t_name = teacher_map.get((c_name, sub), 'Subject Teacher')
            sub_averages.append({'subject': sub, 'teacher': t_name, 'avg': round(s_avg, 1)})
        
        sub_averages = sorted(sub_averages, key=lambda x: x['avg'], reverse=True)
        top_2_teachers = sub_averages[:2]

        class_avg = round(total_class_score / total_class_subs, 1) if total_class_subs > 0 else 0

        ranked_classes.append({
            'class_name': c_name,
            'students_assessed': len(st_dict),
            'class_average': class_avg,
            **divs,
            'top_learners': top_2_learners,
            'top_teachers': top_2_teachers
        })

    ranked_classes = sorted(ranked_classes, key=lambda x: x['class_average'], reverse=True)

    for index, item in enumerate(ranked_classes):
        item['rank'] = index + 1
        if item['class_average'] >= 75: item['color_theme'] = '#2ed573'; item['badge_text'] = 'Excellent'
        elif item['class_average'] >= 60: item['color_theme'] = '#1e90ff'; item['badge_text'] = 'Good'
        elif item['class_average'] >= 45: item['color_theme'] = '#ffa502'; item['badge_text'] = 'Average'
        else: item['color_theme'] = '#ff4757'; item['badge_text'] = 'Critical'

    chart_data = {'labels': ['BOT', 'MOT', 'EOT'], 'datasets': []}
    colors = ['#1B365D', '#2ed573', '#1e90ff', '#ffa502', '#ff4757', '#9b59b6', '#e67e22']
    
    for idx, (c_name, ex_dict) in enumerate(chart_trends.items()):
        bot_avg = round(sum(ex_dict['BOT'])/len(ex_dict['BOT']), 1) if ex_dict['BOT'] else 0
        mot_avg = round(sum(ex_dict['MOT'])/len(ex_dict['MOT']), 1) if ex_dict['MOT'] else 0
        eot_avg = round(sum(ex_dict['EOT'])/len(ex_dict['EOT']), 1) if ex_dict['EOT'] else 0

        chart_data['datasets'].append({
            'label': c_name,
            'data': [bot_avg, mot_avg, eot_avg],
            'borderColor': colors[idx % len(colors)],
            'fill': False,
            'tension': 0.3
        })

    return render_template(
        'performance_reports.html', 
        settings=settings, 
        ranked_classes=ranked_classes,
        selected_exam=selected_exam,
        chart_data=chart_data
    )


# ==============================================================================
# TEACHER & ALLOCATION MANAGEMENT
# ==============================================================================

@app.route('/teachers', methods=['GET', 'POST'])
def manage_teachers():
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    if request.method == 'POST':
        name = request.form['name']
        phone = request.form['phone']
        
        conn.execute("INSERT INTO teachers (name, phone) VALUES (?, ?)", (name, phone))
        conn.commit()
        return redirect(url_for('manage_teachers'))
        
    teachers = conn.execute("SELECT * FROM teachers").fetchall()
    conn.close()
    return render_template("teachers.html", teachers=teachers)

@app.route('/delete-teacher/<int:id>')
def delete_teacher(id):
    if "admin" not in session:
        return redirect(url_for("login"))
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM teachers WHERE id = ?", (id,))
    conn.execute("DELETE FROM allocations WHERE teacher_id = ?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for('manage_teachers'))

@app.route('/allocations', methods=['GET', 'POST'])
def manage_allocations():
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    if request.method == 'POST':
        teacher_id = request.form.get('teacher_id')
        t_row = conn.execute("SELECT name FROM teachers WHERE id = ?", (teacher_id,)).fetchone()
        teacher_name = t_row['name'] if t_row else ""
        
        selected_roles = request.form.getlist('roles')
        for role in selected_roles:
            existing = conn.execute(
                "SELECT id FROM allocations WHERE teacher_id = ? AND role_type = ?", 
                (teacher_id, role)
            ).fetchone()
            
            if not existing:
                conn.execute(
                    "INSERT INTO allocations (teacher_id, teacher_name, class_name, subject, role_type) VALUES (?, ?, ?, ?, ?)",
                    (teacher_id, teacher_name, "N/A", "N/A", role)
                )
        
        selected_classes = request.form.getlist('classes')
        selected_subjects = request.form.getlist('subjects')
        
        for c_name in selected_classes:
            for sub in selected_subjects:
                existing = conn.execute(
                    "SELECT id FROM allocations WHERE teacher_id = ? AND class_name = ? AND subject = ?", 
                    (teacher_id, c_name, sub)
                ).fetchone()
                
                if not existing:
                    conn.execute(
                        "INSERT INTO allocations (teacher_id, teacher_name, class_name, subject, role_type) VALUES (?, ?, ?, ?, ?)",
                        (teacher_id, teacher_name, c_name, sub, "Subject Teacher")
                    )
                    
        conn.commit()
        conn.close()
        return redirect(url_for('manage_allocations'))
        
    teachers_list = conn.execute("SELECT * FROM teachers").fetchall()
    allocations = conn.execute('''
        SELECT allocations.id, teachers.name, allocations.class_name, allocations.subject, allocations.role_type 
        FROM allocations 
        JOIN teachers ON allocations.teacher_id = teachers.id
    ''').fetchall()
    
    conn.close()
    return render_template("allocations.html", teachers=teachers_list, allocations=allocations)

@app.route('/delete-allocation/<int:id>')
def delete_allocation(id):
    if "admin" not in session:
        return redirect(url_for("login"))
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM allocations WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for('manage_allocations'))

@app.route('/timetable')
def timetable():
    if "admin" not in session: 
        return redirect(url_for("login"))
    return render_template("timetable.html")


# ==============================================================================
# SETTINGS & PREFERENCES
# ==============================================================================

@app.route('/settings', methods=['GET', 'POST'])
def settings():
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            id INTEGER PRIMARY KEY,
            current_term TEXT,
            current_year INTEGER
        )
    ''')

    if request.method == 'POST':
        term = request.form.get('current_term', 'Term 1')
        year = request.form.get('current_year', 2026)

        cursor.execute('''
            INSERT INTO settings (id, current_term, current_year)
            VALUES (1, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                current_term = excluded.current_term,
                current_year = excluded.current_year
        ''', (term, year))
        
        cursor.execute('''
            INSERT INTO system_settings (id, current_term, current_year)
            VALUES (1, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                current_term = excluded.current_term,
                current_year = excluded.current_year
        ''', (term, year))
        
        conn.commit()
        conn.close()
        return redirect(url_for('settings', success_msg="Settings updated successfully!"))

    settings_row = cursor.execute("SELECT current_term, current_year FROM settings WHERE id = 1").fetchone()
    conn.close()
    
    success_msg = request.args.get('success_msg')
    error_msg = request.args.get('error_msg')
    
    return render_template('settings.html', settings=settings_row, success_msg=success_msg, error_msg=error_msg)

@app.route('/settings/update_term_settings', methods=['POST'])
def update_term_settings():
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    next_term_raw = request.form.get('next_term_date')
    academic_year = request.form.get('academic_year')
    current_term = request.form.get('current_term')
    
    formatted_date = next_term_raw
    if next_term_raw:
        try:
            date_obj = datetime.strptime(next_term_raw, '%Y-%m-%d')
            formatted_date = date_obj.strftime('%A, %d %B %Y')
        except ValueError:
            pass
    
    conn = sqlite3.connect(DB_NAME)
    
    conn.execute("""
        INSERT INTO system_config (key, value) VALUES ('next_term_date', ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
    """, (formatted_date,))
    
    conn.execute("""
        INSERT INTO system_config (key, value) VALUES ('next_term_date_raw', ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
    """, (next_term_raw,))
    
    conn.execute("""
        INSERT INTO system_config (key, value) VALUES ('academic_year', ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
    """, (academic_year,))
    
    conn.execute("""
        INSERT INTO system_config (key, value) VALUES ('current_term', ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
    """, (current_term,))

    conn.execute("""
        INSERT INTO settings (id, current_term, current_year) VALUES (1, ?, ?)
        ON CONFLICT(id) DO UPDATE SET current_term = excluded.current_term, current_year = excluded.current_year
    """, (current_term, academic_year))

    conn.execute("""
        INSERT INTO system_settings (id, current_term, current_year) VALUES (1, ?, ?)
        ON CONFLICT(id) DO UPDATE SET current_term = excluded.current_term, current_year = excluded.current_year
    """, (current_term, academic_year))
    
    conn.commit()
    conn.close()
    
    return redirect(url_for('settings', success_msg="Academic term settings updated successfully!"))

@app.route('/settings/update_password', methods=['POST'])
def update_password():
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    current_password = request.form['current_password']
    new_password = request.form['new_password']
    confirm_password = request.form['confirm_password']
    
    if new_password != confirm_password:
        return redirect(url_for('settings', error_msg="New passwords do not match!"))
        
    if len(new_password) < 6:
        return redirect(url_for('settings', error_msg="Password must be at least 6 characters long."))

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    
    row = conn.execute("SELECT value FROM system_config WHERE key = 'admin_password'").fetchone()
    stored_password = row['value'] if row else 'admin123'
    
    if current_password != stored_password:
        conn.close()
        return redirect(url_for('settings', error_msg="Current admin password is incorrect!"))
        
    conn.execute("UPDATE system_config SET value = ? WHERE key = 'admin_password'", (new_password,))
    conn.commit()
    conn.close()
    
    return redirect(url_for('settings', success_msg="Security key updated successfully!"))

@app.route('/settings/update_theme', methods=['POST'])
def update_theme():
    if "admin" not in session: 
        return redirect(url_for("login"))
        
    theme_selection = request.form['theme_selection']
    accent_selection = request.form['accent_selection']
    
    conn = sqlite3.connect(DB_NAME)
    conn.execute("UPDATE system_config SET value = ? WHERE key = 'theme'", (theme_selection,))
    conn.execute("UPDATE system_config SET value = ? WHERE key = 'accent'", (accent_selection,))
    conn.commit()
    conn.close()
    
    return redirect(url_for('settings', success_msg="Visual preferences applied globally!"))

@app.route('/admin/update-settings', methods=['POST'])
def update_settings():
    if "admin" not in session:
        return redirect(url_for("login"))

    new_year = request.form.get('year')
    new_term = request.form.get('term')
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE system_settings SET current_year = ?, current_term = ? WHERE id = 1",
        (new_year, new_term)
    )
    cursor.execute(
        "UPDATE settings SET current_year = ?, current_term = ? WHERE id = 1",
        (new_year, new_term)
    )
    conn.commit()
    conn.close()
    
    flash(f"System settings successfully updated to {new_term} {new_year}!", "success")
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/check-marks-schema')
def check_marks_schema():
    if "admin" not in session:
        return redirect(url_for("login"))
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Query SQLite to describe the columns in the marks table
    cursor.execute("PRAGMA table_info(marks);")
    columns = cursor.fetchall()
    conn.close()
    
    # Format and display the columns on the webpage
    schema_info = "<br>".join([f"Column Index: {col[0]} | Name: <b>{col[1]}</b> | Type: {col[2]}" for col in columns])
    return f"<h2>Marks Table Schema Structure:</h2><p>{schema_info}</p><br><a href='/admin/dashboard'>Back to Dashboard</a>"

# ==============================================================================
# EXPORTS & FILE DOWNLOADS
# ==============================================================================
@app.route('/admin/export-class-pdf/<class_name>')
def export_class_pdf(class_name):
    if "admin" not in session:
        return redirect(url_for("login"))
    
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    settings = cursor.execute("SELECT current_year, current_term FROM system_settings WHERE id = 1").fetchone()
    current_year = settings['current_year'] if settings else 2026
    current_term = settings['current_term'] if settings else 'Term 3'

    next_term_row = cursor.execute("SELECT value FROM system_config WHERE key = 'next_term_date'").fetchone()
    next_term_date = next_term_row['value'] if next_term_row else 'TBA'

    subjects = SUBJECTS_P1_P3 if class_name in ['P1', 'P2', 'P3'] else (SUBJECTS_P4_P7 if class_name in ['P4', 'P5', 'P6', 'P7'] else SUBJECTS_PRE_PRIMARY)
    is_primary_upper = class_name in ['P4', 'P5', 'P6', 'P7']

    # Get teacher initials
    initials = {}
    for sub in subjects:
        alloc = cursor.execute("SELECT teachers.name FROM allocations JOIN teachers ON allocations.teacher_id = teachers.id WHERE allocations.class_name = ? AND allocations.subject = ?", (class_name, sub)).fetchone()
        initials[sub] = get_teacher_initials(alloc['name']) if alloc and alloc['name'] else "TR."

    students = cursor.execute("SELECT * FROM students WHERE class_name = ? ORDER BY name ASC", (class_name,)).fetchall()
    if not students:
        conn.close()
        flash(f"No students found in {class_name}!", "warning")
        return redirect(url_for('admin_dashboard'))

    class_adm_nos = [s['admission_no'].strip() for s in students]
    class_marks = cursor.execute(f"SELECT TRIM(admission_no) as adm_no, TRIM(subject) as clean_sub, score, exam_type FROM marks WHERE TRIM(admission_no) IN ({','.join('?' for _ in class_adm_nos)}) AND term = ? AND year = ?", class_adm_nos + [current_term, str(current_year)]).fetchall() if class_adm_nos else []

    class_scores = {'BOT': {}, 'MOT': {}, 'EOT': {}}
    for m in class_marks:
        sub_name = "Social Studies (SST)" if m['clean_sub'] in ["Social Studies", "SST"] else m['clean_sub']
        ex_type = 'MOT' if m['exam_type'] == 'MTC' else m['exam_type']
        if ex_type in class_scores:
            class_scores[ex_type].setdefault(m['adm_no'], {})[sub_name] = m['score']

    def calculate_aggregate(score):
        try:
            val = float(score)
            return 1 if val >= 90 else (2 if val >= 80 else (3 if val >= 65 else (4 if val >= 60 else (5 if val >= 55 else (6 if val >= 50 else (7 if val >= 40 else (8 if val >= 30 else 9)))))))
        except (ValueError, TypeError):
            return None

    # Position calculations across the class
    class_positions = {'BOT': {}, 'MOT': {}, 'EOT': {}}
    for term in ['BOT', 'MOT', 'EOT']:
        student_totals = []
        for adm in class_adm_nos:
            s_marks = class_scores[term].get(adm, {})
            tot_score, tot_aggs, sat_count = 0, 0, 0
            for sub in subjects:
                sc = s_marks.get(sub)
                if sc is not None and str(sc).strip() != '':
                    sat_count += 1
                    tot_score += float(sc)
                    agg = calculate_aggregate(sc)
                    if agg: tot_aggs += agg
            if sat_count > 0:
                student_totals.append({'adm_no': adm, 'total_score': tot_score, 'total_aggregates': tot_aggs, 'is_complete': (sat_count == len(subjects))})
        student_totals.sort(key=lambda x: (not x['is_complete'], x['total_aggregates'], -x['total_score']))
        rank = 1
        for i, st in enumerate(student_totals):
            if not st['is_complete']:
                class_positions[term][st['adm_no']] = "-"
            else:
                if i > 0 and (st['total_aggregates'] != student_totals[i-1]['total_aggregates'] or st['total_score'] != student_totals[i-1]['total_score']):
                    rank = i + 1
                class_positions[term][st['adm_no']] = f"{rank} of {len(student_totals)}"

    batch_reports = []
    for student in students:
        adm_no = student['admission_no'].strip()
        report_cards = {}
        for term in ['BOT', 'MOT', 'EOT']:
            term_marks, total_score, total_aggregates, has_marks, has_f9 = {}, 0, 0, False, False
            s_term_marks = class_scores[term].get(adm_no, {})
            for s in subjects:
                raw_score = s_term_marks.get(s)
                if raw_score is not None and str(raw_score).strip() != '':
                    has_marks = True
                    agg = calculate_aggregate(raw_score)
                    total_score += float(raw_score)
                    if agg == 9: has_f9 = True
                    total_aggregates += agg if agg else 0
                    term_marks[s] = {'score': int(float(raw_score)), 'agg': agg}
                else:
                    term_marks[s] = {'score': '-', 'agg': '-'}
            
            if has_marks:
                if any(term_marks[sub]['score'] == '-' for sub in subjects):
                    div = "Div X"
                elif is_primary_upper:
                    div = "Div 1" if (total_aggregates <= 12 and not has_f9) else ("Div 2" if total_aggregates <= 24 else ("Div 3" if total_aggregates <= 28 else ("Div 4" if total_aggregates <= 32 else "Div U")))
                else:
                    base_div = 1 if 6 <= total_aggregates <= 18 else (2 if 19 <= total_aggregates <= 36 else (3 if 37 <= total_aggregates <= 42 else (4 if 43 <= total_aggregates <= 48 else "U")))
                    div = f"Div {base_div}" if base_div != "U" else "Div U"
                pos = class_positions[term].get(adm_no, '-')
            else:
                div, total_score, total_aggregates, pos = "-", "-", "-", "-"
            report_cards[term] = {'marks': term_marks, 'total_score': total_score, 'total_aggregates': total_aggregates, 'division': div, 'position': pos}

        batch_reports.append({
            'student': student,
            'report_cards': report_cards,
            'initials': initials,
            'class_teacher_comment': 'Good effort during the term. Keep improving.',
            'head_teacher_comment': 'Promising performance. Maintain discipline and focus.'
        })

    conn.close()

    rendered_html = render_template(
        'report_batch_pdf.html',
        batch_reports=batch_reports,
        class_name=class_name,
        subjects=subjects,
        current_year=current_year,
        current_term=current_term,
        next_term_date=next_term_date
    )

    pdf_stream = BytesIO()
    pisa_status = pisa.CreatePDF(BytesIO(rendered_html.encode('utf-8')), dest=pdf_stream)
    if pisa_status.err:
        return f"PDF Error: {pisa_status.err}", 500

    pdf_stream.seek(0)
    response = make_response(pdf_stream.read())
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = f'attachment; filename="{class_name}_Report_Cards_{current_term}_{current_year}.pdf"'
    return response

@app.route('/admin/export-class-list/<string:class_name>')
def export_class_list(class_name):
    if "admin" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute(
        "SELECT admission_no, name, class_name FROM students WHERE class_name = ? ORDER BY name ASC", 
        (class_name,)
    )
    students = cursor.fetchall()
    conn.close()
    
    if not students:
        flash(f"No learners found enrolled in class: {class_name}", "warning")
        return redirect(url_for('admin_dashboard'))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{class_name} Mark Entry"
    ws.views.sheetView[0].showGridLines = True

    ws.append([]) 
    ws.append([f"KIHANI PRIMARY SCHOOL - TERMINAL MARKS ENTRY SHEET"]) 
    ws.append([f"CLASS: {class_name} | YEAR: 2026 | TERM: TERM 1"]) 
    ws.append([]) 
    
    base_headers = ["S/N", "ADMISSION NO.", "PUPIL'S FULL NAME", "CLASS NAME"]
    
    if class_name in ["P1", "P2", "P3"]:
        subject_headers = [
            "Literacy I", 
            "Literacy II", 
            "English", 
            "Mathematics", 
            "Religious Education", 
            "Local Language"
        ]
    else:
        subject_headers = [
            "English", 
            "Mathematics", 
            "Integrated Science", 
            "Social Studies (SST)"
        ]
        
    headers = base_headers + subject_headers
    ws.append(headers)

    navy_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Arial", size=16, bold=True, color="1B365D")
    subtitle_font = Font(name="Arial", size=11, italic=True, bold=True, color="57606F")
    data_font = Font(name="Arial", size=11)
    
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    ws["A2"].font = title_font
    ws["A3"].font = subtitle_font
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=5, column=col_num)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, student in enumerate(students, 1):
        row_data = [
            idx,
            student['admission_no'],
            str(student['name']).upper(),
            student['class_name']
        ]
        
        for _ in subject_headers:
            row_data.append("")
            
        ws.append(row_data)
        
        current_row = 5 + idx
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            
            if col_num != 3:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row > 4 and cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 5, 16)

    ws.row_dimensions[5].height = 25

    excel_stream = BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)

    download_filename = f"{class_name}_Marks_Entry_Template.xlsx"
    return send_file(
        excel_stream,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=download_filename
    )


# ==============================================================================
# MAIN ENTRYPOINT
# ==============================================================================

if __name__ == "__main__":
    app.run(debug=True)