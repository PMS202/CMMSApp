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


frame = db.query(sql = '''SELECT dr.Date, dr.Start_Time, dr.Start_Repair_Time , dr.End_Time, dr.Total_Loss, dr.Repair_Time, 
                                dr.Staff_Name, dr.Error_Code, dr.Machine_Code, dr.Line_Name FROM downtime_report as dr;''')

df = pd.DataFrame(frame, columns=['Date', 'Start_Time', 'Start_Repair_Time', 'End_Time', 'Total_Loss', 'Repair_Time', 'Staff_Name', 'Error_Code', 'Machine_Code', 'Line_Name'])
output_path = os.path.join(PROJECT_ROOT, 'exported_files', 'downtime_SCA.xlsx')
df.to_excel(output_path, index=False)