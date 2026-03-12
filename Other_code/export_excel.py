import pandas as pd
import sys
import os
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

db = Database_process()

frame = db.query(sql = '''SELECT machine_name,machine_code,line_name,maintenance_date FROM view_record_pending
WHERE line_name IN ("MA16")
ORDER BY line_name ASC, machine_code ASC;''')

df = pd.DataFrame(frame, columns=['machine_name', 'machine_code', 'line_name', 'maintenance_date'])
output_path = os.path.join(PROJECT_ROOT, 'exported_files', 'MTPI_MA16.xlsx')
df.to_excel(output_path, index=False)