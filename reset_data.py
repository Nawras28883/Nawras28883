import sqlite3

DB_PATH = 'shipping.db'

def reset_all_data():
    tables = [
        'shipment_item', 'shipment', 'shipment_type',
        'department', 'carrier_company', 'governorate'
    ]
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute('PRAGMA foreign_keys = OFF')
        for table in tables:
            conn.execute(f'DELETE FROM {table}')
        conn.execute('PRAGMA foreign_keys = ON')
        conn.commit()
        print('تم تصفير جميع البيانات بنجاح.')
    except Exception as e:
        print('حدث خطأ أثناء التصفير:', e)
        conn.rollback()
    finally:
        conn.close()

if __name__ == '__main__':
    reset_all_data()
