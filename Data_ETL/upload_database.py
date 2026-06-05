import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError
import time

def upload_with_error_handling(df, table_name, db_config, if_exists='append'):
    
    try:
        connection_string = f"postgresql://{db_config['user']}:{db_config['password']}@{db_config['host']}:{db_config['port']}/{db_config['database']}"
        engine = create_engine(connection_string)
        
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print("Database connection successful")
        
        start_time = time.time()
        
        df.to_sql(
            name=table_name,
            con=engine,
            if_exists=if_exists,
            index=False,
            method='multi',
            chunksize=5000
        )
        
        elapsed_time = time.time() - start_time
        
        print(f"Successfully uploaded {len(df)} rows to '{table_name}'")
        print(f"Time taken: {elapsed_time:.2f} seconds")
        print(f"Rows per second: {len(df)/elapsed_time:.0f}")
        
        engine.dispose()
        return True
        
    except SQLAlchemyError as e:
        print(f"Database error: {e}")
        return False
    except Exception as e:
        print(f"Unexpected error: {e}")
        return False

db_config = {
    'host': 'localhost',
    'port': 5432,
    'database': 'template_data',
    'user': 'ddi20241312005',
    'password': "Welcome2005"
}
