import pandas as pd
import sys
import os
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

db = Database_process()

frame = db.query(sql = '''SELECT machine_code, status, line_name
                 FROM maintenance_with_status AS ms
                 WHERE status = "Near due" COLLATE utf8mb4_unicode_ci;''')

df = pd.DataFrame(frame, columns=['machine_code', 'status', 'line_name'])
output_path = os.path.join(PROJECT_ROOT, 'exported_files', 'Mar.xlsx')
df.to_excel(output_path, index=False)