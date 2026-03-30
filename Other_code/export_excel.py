import pandas as pd
import sys
import os
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

db = Database_process()

frame = db.query(sql = '''SELECT machine_name,machine_code, line_name,status FROM maintenance_with_status
WHERE department_name = "PE3" COLLATE utf8mb4_unicode_ci AND status COLLATE utf8mb4_unicode_ci IN ("Near due", "Overdue") 
ORDER BY line_name ASC, machine_code ASC;''')

df = pd.DataFrame(frame, columns=['machine_name', 'machine_code', 'line_name', 'status'])
output_path = os.path.join(PROJECT_ROOT, 'exported_files', 'PE3_MAR.xlsx')
df.to_excel(output_path, index=False)