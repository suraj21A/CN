# db_helper.py

import mysql.connector
from mysql.connector import Error
from mysql_config import MYSQL_CONFIG

def get_db_connection():
    """Database connection banao"""
    try:
        conn = mysql.connector.connect(**MYSQL_CONFIG)
        return conn
    except Error as e:
        print(f"❌ MySQL Connection Error: {e}")
        return None

def save_cn_to_mysql(row_data):
    """
    CN row ko MySQL mein save karo
    
    row_data = {
        'CN No': '5280029184',
        'CN Type': 'Billed',
        ... sab 26 columns
    }
    """
    conn = get_db_connection()
    if not conn:
        print("⚠️  Database connection failed - skipping MySQL save")
        return False
    
    try:
        cursor = conn.cursor()
        
        # Build INSERT query
        columns = ", ".join([f"`{col}`" for col in row_data.keys()])
        placeholders = ", ".join(["%s"] * len(row_data))
        
        query = f"INSERT INTO cn_master ({columns}) VALUES ({placeholders})"
        values = tuple(row_data.values())
        
        # Execute
        cursor.execute(query, values)
        conn.commit()
        
        print(f"✓ MySQL: CN {row_data.get('CN No')} saved")
        cursor.close()
        conn.close()
        return True
        
    except Error as e:
        print(f"❌ MySQL Insert Error for CN {row_data.get('CN No')}: {e}")
        if conn and conn.is_connected():
            conn.close()
        return False