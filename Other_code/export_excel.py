import pandas as pd
import sys
import os
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

db = Database_process()

frame = db.query(sql = '''SELECT lot.operation_date , pl.line_name,lot.model_running, lot.operation_hours, lot.setup_time,lot.break_time
 FROM `line_operation_times` as lot
 JOIN production_lines as pl ON lot.line_id = pl.line_id
 WHERE lot.operation_date BETWEEN '2026-06-01' AND '2026-06-30';''')

# frame = db.query(sql = '''SELECT machine_code,machine_name,line_name,department_name
#                             FROM maintenance_with_status
#                             WHERE department_name = 'PE3' COLLATE utf8mb4_unicode_ci  AND status = "Near due" COLLATE utf8mb4_unicode_ci;''')


# frame = db.query(sql = '''SELECT pr.report_title, report_date, pr.line_id, pl.line_name, d.department_name, pr.report_type_id, rt.report_type_name, pr.issue_description, pr.corrective_action, pr.reported_by, pr.status, pr.notes, pr.report_file_path, pr.path_type
# FROM problem_reports as pr
# JOIN production_lines as pl ON pr.line_id = pl.line_id
# JOIN departments AS d ON pr.department_id = d.department_id
# JOIN report_types AS rt ON pr.report_type_id = rt.report_type_id;''')

df = pd.DataFrame(frame, columns=['operation_date', 'line_name', 'model_running', 'operation_hours', 'setup_time', 'break_time'])
output_path = os.path.join(PROJECT_ROOT, 'exported_files', 'WT_Jun_26.xlsx')
df.to_excel(output_path, index=False)