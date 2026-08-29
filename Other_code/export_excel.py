import pandas as pd
import sys
import os
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

db = Database_process()

# frame = db.query(sql = '''SELECT lot.operation_date , pl.line_name,lot.model_running, lot.operation_hours, lot.setup_time,lot.break_time
#  FROM `line_operation_times` as lot
#  JOIN production_lines as pl ON lot.line_id = pl.line_id
#  WHERE lot.operation_date BETWEEN '2026-06-01' AND '2026-06-30';''')

# frame = db.query(sql = '''SELECT m.machine_code, d.department_name, pl.line_name, d.department_name, mp.status FROM maintenance_plan as mp
# JOIN machines as m ON mp.machine_id = m.machine_id
# JOIN production_lines as pl ON mp.line_id = pl.line_id
# JOIN departments as d ON pl.department_id = d.department_id
# WHERE mp.`status` = "Overdue" AND mp.maintenance_date IS NULL AND mp.month_year_id = 19;
# ''')


# frame = db.query(sql = '''SELECT m.machine_code, pl.line_name, mp.week
# FROM maintenance_plan AS mp
# JOIN machines AS m ON mp.machine_id = m.machine_id
# JOIN production_lines AS pl ON mp.line_id = pl.line_id
# WHERE mp.machine_id NOT IN (
#     SELECT rp.machine_id
#     FROM record_pending as rp
#     WHERE rp.technical = "VINH"
# ) AND mp.month_year_id = 20 AND mp.maintenance_date IS NULL AND mp.status = "Near due" ;''')

frame = db.query(sql = '''SELECT * FROM downtime_report
WHERE `Date` BETWEEN '2026-07-01' AND '2026-07-31';''')

df = pd.DataFrame(frame)
output_path = os.path.join(PROJECT_ROOT, 'exported_files', 'dt_SCA_Kha.xlsx')
df.to_excel(output_path, index=False)