"""
Error Reporter logic for querying processing anomalies and exporting them
to error reports (CSV).
"""
import os
import csv
import sqlite3

def get_error_count(conn: sqlite3.Connection) -> int:
    """
    Get the total number of ingestion or validation errors logged.
    
    Args:
        conn (sqlite3.Connection): DB connection.
        
    Returns:
        int: Number of errors.
    """
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM etl_errors")
    return cursor.fetchone()[0]

def export_errors_to_csv(conn: sqlite3.Connection, output_path: str):
    """
    Query all rows from etl_errors and save them to a styled CSV report.
    
    Args:
        conn (sqlite3.Connection): DB connection.
        output_path (str): File path where CSV should be exported.
    """
    cursor = conn.cursor()
    cursor.execute("SELECT id, source_file, error_reason FROM etl_errors ORDER BY id ASC")
    rows = cursor.fetchall()
    
    # Ensure directory exists
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    
    # Write CSV
    with open(output_path, mode='w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        # Header Row
        writer.writerow(["Error ID", "Source File Name", "Error/Validation Reason"])
        for row in rows:
            writer.writerow(row)

def clear_errors(conn: sqlite3.Connection):
    """
    Truncate all logged records from the etl_errors table.
    
    Args:
        conn (sqlite3.Connection): DB connection.
    """
    cursor = conn.cursor()
    cursor.execute("DELETE FROM etl_errors")
    conn.commit()
