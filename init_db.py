import sqlite3
from app import DATABASE

def init_sample_data():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    # Add shipping types if they don't exist
    shipping_types = ['طرد', 'صندوق', 'كرتون', 'كيس']
    for type_name in shipping_types:
        cursor.execute('INSERT OR IGNORE INTO shipment_type (name) VALUES (?)', (type_name,))

    # Add departments if they don't exist
    departments = ['المبيعات', 'المخزن', 'التوصيل', 'خدمة العملاء']
    for dept_name in departments:
        cursor.execute('INSERT OR IGNORE INTO department (name) VALUES (?)', (dept_name,))

    conn.commit()
    conn.close()
    print('تم إضافة البيانات الأولية بنجاح!')

if __name__ == '__main__':
    init_sample_data()