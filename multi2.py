import psycopg2
from psycopg2 import sql
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT

def connect_db(database_name=None):
    """Connect to the PostgreSQL database."""
    try:
        conn = psycopg2.connect(
            dbname=database_name or "postgres",  
            user="postgres",
            password="Gajanan@1535",
            host="localhost",
            port="5432"
        )
        return conn
    except Exception as e:
        print(f"Error connecting to database: {e}")
        return None

def list_databases():
    """List all available databases."""
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT datname FROM pg_database WHERE datistemplate = false;")
    databases = cursor.fetchall()
    cursor.close()
    conn.close()
    return [db[0] for db in databases]

def list_tables(conn):
    """List all tables in the current database."""
    cursor = conn.cursor()
    cursor.execute("""SELECT table_name FROM information_schema.tables WHERE table_schema = 'public';""")
    tables = cursor.fetchall()
    cursor.close()
    return [table[0] for table in tables]

def get_table_columns(conn, table_name):
    """Get columns of a table in the current database."""
    cursor = conn.cursor()
    cursor.execute(sql.SQL("""
        SELECT column_name, data_type FROM information_schema.columns
        WHERE table_name = %s;
    """), (table_name,))
    columns = cursor.fetchall()
    cursor.close()
    return columns

def create_database():
    """Create a new database."""
    db_name = input("Enter the name for the new database: ")
    
    try:
        
        conn = psycopg2.connect(
            dbname="postgres",
            user="postgres",
            password="Gajanan@1535",
            host="localhost",
            port="5432"
        )
        
       
        conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
        
        cursor = conn.cursor()
        
       
        cursor.execute(sql.SQL("CREATE DATABASE {}").format(sql.Identifier(db_name)))
        print(f"Database '{db_name}' created successfully.")
        
    except Exception as e:
        print(f"Error creating database: {e}")
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'conn' in locals():
            conn.close()

def create_table(conn):
    """Create a new table in the selected database."""
    table_name = input("Enter table name: ")
    columns = []
    
    while True:
        col_name = input("Enter column name (or 'done' to finish): ")
        if col_name.lower() == "done":
            break
        col_type = input(f"Enter data type for {col_name} (e.g., VARCHAR(50), INTEGER, TEXT): ")
        columns.append(f"{col_name} {col_type}")
    
    if not columns:
        print("No columns provided. Table not created.")
        return
    
    cursor = conn.cursor()
    try:
        create_table_query = sql.SQL("CREATE TABLE {} ({})").format(
            sql.Identifier(table_name),
            sql.SQL(", ").join(sql.SQL(col) for col in columns)
        )
        cursor.execute(create_table_query)
        conn.commit()
        print(f"Table '{table_name}' created successfully.")
    except Exception as e:
        print(f"Error creating table: {e}")
    finally:
        cursor.close()

def insert_data(conn, table_name):
    """Insert data into a specified table."""
    columns = get_table_columns(conn, table_name)
    values = []
    for col_name, col_type in columns:
        value = input(f"Enter value for {col_name} ({col_type}): ")
        values.append(value)
    
    cursor = conn.cursor()
    try:
        insert_query = sql.SQL("""
            INSERT INTO {} ({}) VALUES ({})
        """).format(
            sql.Identifier(table_name),
            sql.SQL(", ").join(map(sql.Identifier, [col[0] for col in columns])),
            sql.SQL(", ").join(sql.Placeholder() for _ in values)
        )
        cursor.execute(insert_query, values)
        conn.commit()
        print("Data inserted successfully.")
    except Exception as e:
        print(f"Error inserting data: {e}")
    finally:
        cursor.close()

def update_data(conn, table_name):
    """Update data in a specified table."""
    columns = get_table_columns(conn, table_name)
    primary_key = columns[0][0]
    record_id = input(f"Enter {primary_key} of the record to update: ")
    updates = {}
    for col_name, col_type in columns:
        value = input(f"Enter new value for {col_name} ({col_type}) (leave empty to skip): ")
        if value:
            updates[col_name] = value
    
    if not updates:
        print("No updates provided.")
        return
    
    cursor = conn.cursor()
    try:
        update_query = sql.SQL("""
            UPDATE {} SET {} WHERE {} = %s;
        """).format(
            sql.Identifier(table_name),
            sql.SQL(", ").join(
                sql.SQL("{} = {}").format(sql.Identifier(k), sql.Placeholder()) for k in updates.keys()
            ),
            sql.Identifier(primary_key)
        )
        cursor.execute(update_query, list(updates.values()) + [record_id])
        conn.commit()
        print("Record updated successfully.")
    except Exception as e:
        print(f"Error updating data: {e}")
    finally:
        cursor.close()

def delete_data(conn, table_name):
    """Delete data from a specified table."""
    columns = get_table_columns(conn, table_name)
    primary_key = columns[0][0]
    record_id = input(f"Enter {primary_key} of the record to delete: ")
    
    cursor = conn.cursor()
    try:
        delete_query = sql.SQL("""
            DELETE FROM {} WHERE {} = %s;
        """).format(sql.Identifier(table_name), sql.Identifier(primary_key))
        cursor.execute(delete_query, (record_id,))
        conn.commit()
        print("Record deleted successfully.")
    except Exception as e:
        print(f"Error deleting data: {e}")
    finally:
        cursor.close()

def main():
    while True:
        print("\nAvailable options:")
        print("1. Create Database")
        print("2. Select Database")
        print("3. Exit")
        choice = input("Choose an option: ")

        if choice == "1":
            create_database()
        elif choice == "2":
            databases = list_databases()
            print("\nAvailable databases:")
            for idx, db in enumerate(databases, 1):
                print(f"{idx}. {db}")
            try:
                db_choice = int(input("\nSelect a database by number: "))
                selected_db = databases[db_choice - 1]
                print(f"\nConnecting to database: {selected_db}")
                conn = connect_db(selected_db)
                if conn:
                    while True:
                        tables = list_tables(conn)
                        print("\nAvailable tables:")
                        for idx, table in enumerate(tables, 1):
                            print(f"{idx}. {table}")
                        
                        print("\nOptions:")
                        print("1. Create Table")
                        print("2. Insert Data")
                        print("3. Update Data")
                        print("4. Delete Data")
                        print("5. Return to database selection")

                        action = input("\nSelect action (1-5): ").strip()

                        if action == "1":
                            create_table(conn)
                        elif action == "2":
                            table_name = input("Enter table name: ")
                            insert_data(conn, table_name)
                        elif action == "3":
                            table_name = input("Enter table name: ")
                            update_data(conn, table_name)
                        elif action == "4":
                            table_name = input("Enter table name: ")
                            delete_data(conn, table_name)
                        elif action == "5":
                            print("Returning to database selection.")
                            break
                        else:
                            print("Invalid action selected.")
                    
                    conn.close()
            except (ValueError, IndexError):
                print("Invalid database selection. Please try again.")
        elif choice == "3":
            print("Exiting program.")
            break
        else:
            print("Invalid option selected.")

if __name__ == "__main__":
    main()
