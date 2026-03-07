from sqlalchemy import create_engine, text
from sqlalchemy.orm import scoped_session, sessionmaker
from dotenv import load_dotenv
import os
# def resource_path(relative_path):
#     """ Get absolute path to resource, works for dev and for PyInstaller """
#     try:
#         # PyInstaller creates a temp folder and stores path in _MEIPASS
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.abspath(".")

#     return os.path.join(base_path, relative_path)


class Database_process():
    def __init__(self):
        load_dotenv() 
        db_url = os.getenv("DB_URL")
        for attempt in range(3):
            try:
                self.engine = create_engine(
                    db_url,
                    connect_args={"charset": "utf8mb4"},
                    future=True)
                with self.engine.connect() as conn:
                    pass
                break
            except Exception as retry_e:
                if attempt < 2:
                    continue 
                else:
                    raise ConnectionError(f"Database connection failed after 3 attempts: {retry_e}")

        self.Session = scoped_session(sessionmaker(bind=self.engine))

    def query(self, sql=None, params=None):
        if sql is None:
            raise ValueError("SQL query must be provided.")

        with self.Session() as session:
            try:
                if isinstance(params, list): 
                    session.execute(text(sql), params)
                    session.commit()
                    return len(params)
                else:
                    result = session.execute(text(sql), params or {})
                    if sql.strip().lower().startswith("select"):
                        return result.fetchall()
                    else:
                        session.commit()
                        return result.rowcount
            except Exception as e:
                session.rollback()
                raise e
    def executemany(self, sql, params_list):
        if not sql:
            raise ValueError("SQL query must be provided.")
        if not params_list or not isinstance(params_list, list):
            raise ValueError("params_list must be a non-empty list.")

        with self.Session() as session:
            try:
                result = session.execute(
                    text(sql),
                    params_list,
                    execution_options={"executemany": True}
                )
                session.commit()
                return result.rowcount
            except Exception as e:
                session.rollback()
                raise e

    def close(self):
        if self.engine:
            self.engine.dispose()
        if self.Session:
            self.Session.remove()
        self.engine = None
        self.Session = None



