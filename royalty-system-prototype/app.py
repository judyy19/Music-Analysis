"""
Royalty System Prototype main entry point.
Provides a CLI to initialize the SQLite database, seed user accounts,
trigger ETL file ingestion, and launch the Streamlit portal.
"""
import os
import sys
import argparse
import sqlite3
import subprocess

import config
from catalog.catalog_logic import init_catalog_tables, import_catalog_from_excel
from processor.etl_logic import init_etl_tables, run_etl_pipeline

def init_database():
    """
    Initialize SQLite database tables and import catalog split fee rates.
    """
    print("Initializing SQLite database...")
    
    # Establish connection
    conn = sqlite3.connect(config.DB_PATH)
    
    try:
        # Create all schemas
        print("  - Creating table schemas...")
        init_catalog_tables(conn)
        init_etl_tables(conn)
        
        # Import Catalog splits
        print(f"  - Seeding catalog splits from: {config.ROYALTY_RATE_PATH} ...")
        if os.path.exists(config.ROYALTY_RATE_PATH):
            import_catalog_from_excel(conn, config.ROYALTY_RATE_PATH)
            print("    [SUCCESS] Catalog splits successfully seeded.")
        else:
            print(f"    [WARNING] Royalty rate file not found at {config.ROYALTY_RATE_PATH}. Splits table remains empty.")
            
        print("\nDatabase initialization completed successfully!")
        
    except Exception as e:
        print(f"[ERROR] Error during database initialization: {e}")
        conn.rollback()
    finally:
        conn.close()

def trigger_etl():
    """Trigger the ETL ingestion pipeline from command line."""
    print(f"Starting ETL Ingestion Pipeline scanning: {config.RAW_DATA_DIR} ...")
    
    if not os.path.exists(config.RAW_DATA_DIR):
        print(f"[ERROR] Raw data directory not found at: {config.RAW_DATA_DIR}")
        sys.exit(1)
        
    conn = sqlite3.connect(config.DB_PATH)
    try:
        summary = run_etl_pipeline(conn, config.RAW_DATA_DIR)
        print("\nIngestion Pipeline Summary:")
        print(f"  - Files Processed: {summary['files_processed']}")
        print(f"  - Records Added:   {summary['records_added']}")
        print(f"  - Duplicate/Skipped Files: {summary['records_skipped_duplicate']}")
        print(f"  - Errors Logged:   {summary['errors_logged']}")
        print("\nETL ingestion completed.")
    except Exception as e:
        print(f"[ERROR] Ingestion Pipeline error: {e}")
    finally:
        conn.close()

def launch_portal():
    """Launch the Streamlit web interface using subprocess."""
    print("Launching Streamlit Royalty System Portal...")
    dashboard_path = os.path.join(config.PROTOTYPE_DIR, "portal", "dashboard_logic.py")
    
    try:
        subprocess.run([
            "streamlit", "run", dashboard_path,
            "--server.port=8501",
            "--server.address=0.0.0.0",
            "--server.enableCORS=false",
            "--server.enableXsrfProtection=false",
            "--server.maxUploadSize=200"
        ], check=True)
    except KeyboardInterrupt:
        print("\nPortal stopped.")
    except Exception as e:
        print(f"[ERROR] Error launching Streamlit portal: {e}")

def main():
    parser = argparse.ArgumentParser(
        description="Royalty System Prototype Command Line Tool",
        epilog="Examples:\n  python app.py --init\n  python app.py --etl\n  python app.py --run-portal",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('--init', action='store_true', help='Initialize database tables, load catalog rates, and seed credentials.')
    parser.add_argument('--etl', action='store_true', help='Run the ETL pipeline to ingest raw spreadsheets into SQLite records.')
    parser.add_argument('--run-portal', action='store_true', help='Launch the interactive Streamlit user dashboard portal.')
    
    args = parser.parse_args()
    
    if not any(vars(args).values()):
        parser.print_help()
        sys.exit(0)
        
    if args.init:
        init_database()
    if args.etl:
        trigger_etl()
    if args.run_portal:
        launch_portal()

if __name__ == "__main__":
    main()
