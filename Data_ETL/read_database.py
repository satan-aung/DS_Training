from sqlalchemy import create_engine, text
import pandas as pd
from Data_ETL.upload_database import db_config

def read_data():
    """ conn = psycopg2.connect(
        host=db_config["host"],
        port=db_config["port"],
        database=db_config["database"],
        user=db_config["user"],
        password=db_config["password"]
    )
    cursor = conn.cursor()
    df = pd.read_sql("SELECT * FROM template_data LIMIT 10", conn)
    print("\nYour first data 10 rows:")
    print(df) """

    connection_string = f"postgresql://{db_config['user']}:{db_config['password']}@{db_config['host']}:{db_config['port']}/{db_config['database']}"
    engine = create_engine(connection_string)
        
    with engine.connect() as conn:
        df = pd.read_sql(text("SELECT * FROM template_data LIMIT 10"), conn)
    print("\nYour first data 10 rows:")
    print(df)

    engine.dispose()