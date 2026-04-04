import pandas as pd
import sys
import os
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

db = Database_process()

# frame = db.query(sql = '''SELECT m.machine_code,m.machine_name,pl.line_name, d.department_name
#                             FROM machines m
#                             JOIN maintenance_plan mp 
#                                 ON m.machine_id = mp.machine_id
#                             JOIN production_lines pl 
#                                 ON mp.line_id = pl.line_id
#                             JOIN departments d 
#                                 ON pl.department_id = d.department_id
#                             WHERE mp.maintenance_date IS NULL
#                             AND ( mp.status IS NULL OR mp.status = '') AND d.department_name = 'PE3' AND mp.month_year_id = 15;''')

# frame = db.query(sql = '''SELECT machine_code,machine_name,line_name,department_name
#                             FROM maintenance_with_status
#                             WHERE department_name = 'PE3' COLLATE utf8mb4_unicode_ci  AND status = "Near due" COLLATE utf8mb4_unicode_ci;''')


frame = db.query(sql = '''SELECT machine_code,machine_name,line_name,department_name
                            FROM view_record_pending
                            WHERE department_name = 'PE3';''')

df = pd.DataFrame(frame, columns=['machine_code', 'machine_name', 'line_name', 'department_name'])
output_path = os.path.join(PROJECT_ROOT, 'exported_files', 'PE3_PENDING_MAR.xlsx')
df.to_excel(output_path, index=False)