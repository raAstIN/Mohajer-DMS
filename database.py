import sqlite3
import os
import shutil
import datetime
import jdatetime


ROOT = os.path.dirname(__file__)
FILES_DIR = os.path.join(ROOT, 'files')
UPLOADS_DIR = os.path.join(FILES_DIR, 'uploads')
BACKUP_DIR = os.path.join(FILES_DIR, 'backup')
DB_PATH = os.path.join(FILES_DIR, 'cases.db')


def ensure_dirs():
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    os.makedirs(BACKUP_DIR, exist_ok=True)


def get_connection():
    ensure_dirs()
    conn = sqlite3.connect(DB_PATH)
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''
    CREATE TABLE IF NOT EXISTS cases (
        id TEXT PRIMARY KEY,
        title TEXT,
        date TEXT,
        duration TEXT,
        duration_from TEXT,
        duration_to TEXT,
        mojer TEXT,
        mostajjer TEXT,
        karfarma TEXT,
        piman TEXT,
        subject TEXT,
        contract_amount TEXT,
        bank_owner_name TEXT,
        bank_account_number TEXT,
        bank_shaba_number TEXT,
        bank_card_number TEXT,
        description TEXT,
        folder_path TEXT,
        case_type TEXT
    )
    ''')

    # Simple migration: check if columns exist and add them if not.
    cur.execute("PRAGMA table_info(cases)")
    columns = [info[1] for info in cur.fetchall()]
    if 'contract_amount' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN contract_amount TEXT")
    if 'case_type' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN case_type TEXT")
    if 'duration_from' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN duration_from TEXT")
    if 'duration_to' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN duration_to TEXT")
    if 'mojer' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN mojer TEXT")
    if 'mostajjer' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN mostajjer TEXT")
    if 'karfarma' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN karfarma TEXT")
    if 'piman' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN piman TEXT")
    if 'bank_owner_name' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN bank_owner_name TEXT")
    if 'bank_account_number' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN bank_account_number TEXT")
    if 'bank_shaba_number' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN bank_shaba_number TEXT")
    if 'bank_card_number' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN bank_card_number TEXT")
    # Keep 'parties' column for backward compatibility if it exists
    if 'parties' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN parties TEXT")
    if 'status' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN status TEXT DEFAULT 'در جریان'")
    if 'bank_name' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN bank_name TEXT")
    if 'bank_branch' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN bank_branch TEXT")
    if 'payment_id' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN payment_id TEXT")
    if 'guarantee_amount' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN guarantee_amount TEXT")
    if 'guarantee_type' not in columns:
        cur.execute("ALTER TABLE cases ADD COLUMN guarantee_type TEXT")

    conn.commit()
    conn.close()


def add_case(data: dict):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''INSERT INTO cases (id, title, date, duration, duration_from, duration_to, mojer, mostajjer, karfarma, piman, subject, contract_amount, bank_owner_name, bank_account_number, bank_shaba_number, bank_card_number, bank_name, bank_branch, payment_id, guarantee_amount, guarantee_type, description, folder_path, case_type, status)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', (
        data['id'], data.get('title'), data.get('date'), data.get('duration'), data.get('duration_from'), data.get('duration_to'),
        data.get('mojer'), data.get('mostajjer'), data.get('karfarma'), data.get('piman'),
        data.get('subject'), data.get('contract_amount'), 
        data.get('bank_owner_name'), data.get('bank_account_number'), data.get('bank_shaba_number'), data.get('bank_card_number'),
        data.get('bank_name'), data.get('bank_branch'), data.get('payment_id'),
        data.get('guarantee_amount'), data.get('guarantee_type'),
        data.get('description'), data.get('folder_path'), data.get('case_type'), data.get('status', 'در جریان')
    ))
    conn.commit()
    conn.close()
    backup_db()


def get_case_by_id(case_id: str):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM cases WHERE id = ?', (case_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return None
    # Use cursor.description to get the actual column order from the DB; this
    # avoids mismatches when the schema has been migrated (columns added).
    col_names = [d[0] for d in cur.description]
    result = dict(zip(col_names, row))
    conn.close()
    return result


def update_case(case_id: str, data: dict):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''UPDATE cases SET title=?, date=?, duration=?, duration_from=?, duration_to=?, mojer=?, mostajjer=?, karfarma=?, piman=?, subject=?, contract_amount=?, bank_owner_name=?, bank_account_number=?, bank_shaba_number=?, bank_card_number=?, bank_name=?, bank_branch=?, payment_id=?, guarantee_amount=?, guarantee_type=?, description=?, folder_path=?, case_type=?, status=? WHERE id=?''', (
        data.get('title'), data.get('date'), data.get('duration'), data.get('duration_from'), data.get('duration_to'),
        data.get('mojer'), data.get('mostajjer'), data.get('karfarma'), data.get('piman'),
        data.get('subject'), data.get('contract_amount'),
        data.get('bank_owner_name'), data.get('bank_account_number'), data.get('bank_shaba_number'), data.get('bank_card_number'),
        data.get('bank_name'), data.get('bank_branch'), data.get('payment_id'),
        data.get('guarantee_amount'), data.get('guarantee_type'),
        data.get('description'), data.get('folder_path'), data.get('case_type'), data.get('status', 'در جریان'), case_id
    ))
    conn.commit()
    conn.close()
    backup_db()


def delete_case(case_id: str):
    # remove db record
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT folder_path FROM cases WHERE id = ?', (case_id,))
    row = cur.fetchone()
    folder = None
    if row:
        folder = row[0]
    cur.execute('DELETE FROM cases WHERE id = ?', (case_id,))
    conn.commit()
    conn.close()
    # remove folder if exists
    if folder and os.path.exists(folder):
        try:
            shutil.rmtree(folder)
        except Exception:
            pass
    backup_db()


def search_cases(filter_type: str, query: str):
    conn = get_connection()
    cur = conn.cursor()
    q = f"%{query}%"
    # Always include 'subject' in the returned columns so the UI can display it
    if filter_type == 'title':
           cur.execute('SELECT id, title, subject, date, case_type, duration, status, contract_amount FROM cases WHERE title LIKE ? ORDER BY date DESC', (q,))
    elif filter_type == 'subject':
           cur.execute('SELECT id, title, subject, date, case_type, duration, status, contract_amount FROM cases WHERE subject LIKE ? ORDER BY date DESC', (q,))
    elif filter_type == 'date':
           cur.execute('SELECT id, title, subject, date, case_type, duration, status, contract_amount FROM cases WHERE date LIKE ? ORDER BY date DESC', (q,))
    elif filter_type == 'case_type':
           cur.execute('SELECT id, title, subject, date, case_type, duration, status, contract_amount FROM cases WHERE case_type LIKE ? ORDER BY date DESC', (q,))
    elif filter_type == 'id':
           cur.execute('SELECT id, title, subject, date, case_type, duration, status, contract_amount FROM cases WHERE id LIKE ? ORDER BY date DESC', (q,))
    else:  # full search
           cur.execute('''SELECT id, title, subject, date, case_type, duration, status, contract_amount FROM cases
                       WHERE id LIKE ? OR title LIKE ? OR subject LIKE ? OR description LIKE ? OR contract_amount LIKE ? OR case_type LIKE ? ORDER BY date DESC''', (q, q, q, q, q, q))
    rows = cur.fetchall()
    conn.close()
    return rows


def backup_db():
    ensure_dirs()
    if not os.path.exists(DB_PATH):
        return
    ts = jdatetime.datetime.now().strftime('%Y%m%d%H%M%S')
    dst = os.path.join(BACKUP_DIR, f'cases_{ts}.db')
    try:
        shutil.copy2(DB_PATH, dst)
    except Exception:
        pass


# Initialize DB on import
init_db()
