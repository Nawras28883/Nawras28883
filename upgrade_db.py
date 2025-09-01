import sqlite3
import sys

DB_PATH = 'shipping.db'

def upgrade_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    try:
        c.execute('PRAGMA foreign_keys=off;')
        c.execute('BEGIN TRANSACTION;')
        # إنشاء جدول جديد مع receipt_number فريد
        c.execute('''
            CREATE TABLE IF NOT EXISTS shipment_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shopiny_number TEXT UNIQUE NOT NULL,
                receipt_number TEXT UNIQUE,
                order_number TEXT,
                delivery_date DATETIME,
                from_governorate TEXT NOT NULL,
                to_governorate TEXT NOT NULL,
                carrier_company TEXT,
                notes TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        # جلب كل الشحنات مع أول receipt_number فريد فقط
        rows = c.execute('SELECT * FROM shipment').fetchall()
        seen_receipts = set()
        for row in rows:
            receipt = row[2]
            if receipt and receipt in seen_receipts:
                continue
            seen_receipts.add(receipt)
            c.execute('''INSERT INTO shipment_new (id, shopiny_number, receipt_number, order_number, delivery_date, from_governorate, to_governorate, carrier_company, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', row)
        c.execute('DROP TABLE shipment;')
        c.execute('ALTER TABLE shipment_new RENAME TO shipment;')
        c.execute('COMMIT;')
        c.execute('PRAGMA foreign_keys=on;')
        print('Database upgraded successfully.')
    except Exception as e:
        print('Error during upgrade:', e)
        c.execute('ROLLBACK;')
    finally:
        conn.close()

if __name__ == '__main__':
    upgrade_db()
