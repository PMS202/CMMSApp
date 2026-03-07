import pymsteams
from sqlalchemy import create_engine, text
from sqlalchemy.orm import scoped_session, sessionmaker
from dotenv import load_dotenv
import requests
import datetime
import os
from Database.MariaDB import Database_process


# Webhook URL
# myTeamsMessage = pymsteams.connectorcard(
#     "https://yageo.webhook.office.com/webhookb2/43c60a89-1113-44b3-bbf9-563c8da25dbd@66d44188-4e30-446d-9ab1-a4ae1962a1b1/IncomingWebhook/b25dba6b7a8e48b2839d4f66636c519f/743e9a0f-79fa-41cc-acb7-5b896ae80061/V2LDZ-GY-si73pIecY4DKPjZgyRgnAPisvnS7XYUBbOc41"
# )
workflow_url = "https://default66d441884e30446d9ab1a4ae1962a1.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a9cee6f5f061445791bafe12b56136a5/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=6Wz-3UI5ibZb--T869Igbhnug9VEXWn6uZCNMSPy6mM"

class Notification():
    def __init__(self,DB):
        self.today = datetime.datetime.now().date()
        self.week = self.company_week_number(self.today)
        self.DB = DB

    def company_week_number(self,date: datetime.date) -> int:
        year_start = datetime.date(date.year, 1, 1)
        days_to_sunday = (6 - year_start.weekday()) % 7 
        first_sunday = year_start + datetime.timedelta(days=days_to_sunday)
        if date < first_sunday:
                return 1
        delta_days = (date - first_sunday).days
        return delta_days // 7 + 1

    def stock_notification(self):
        script_dict = {
            "total_part":''' SELECT COUNT(*) as total_part FROM `spare_part_view`;''',

            "items_need_order": '''SELECT COUNT(*) as items_need_order FROM `spare_part_view` WHERE stockup > 0;''',

            "part_need_order": '''SELECT sum(stockup) as part_need_order FROM `spare_part_view` WHERE stockup > 0;''',
            "total_cost": '''SELECT currency,sum(total_cost)  AS total_cost_sum
                            FROM `spare_part_view` 
                            WHERE stockup > 0 
                            GROUP BY currency;'''
        }
        result_dict = {}
        url = "https://open.er-api.com/v6/latest/USD"
        response = requests.get(url)
        exchange_rates = response.json()["rates"]
        try:
            for key, sql in script_dict.items():
                result_dict[key]  = self.DB.query(sql)
        except:
            pass
        total_cost = 0
        for item in result_dict["total_cost"]:
            total_cost += round(float(item[1]) / float(exchange_rates[item[0]]),2)

        # rows = [
        #     ("Total part:", str(result_dict["total_part"][0][0])),
        #     ("Items need order:",str(result_dict["items_need_order"][0][0])),
        #     ("Part need order:", str(result_dict["part_need_order"][0][0])),
        #     ("Order cost:", f"{total_cost} USD")
        # ]
        # myTeamsMessage = pymsteams.connectorcard( "https://yageo.webhook.office.com/webhookb2/43c60a89-1113-44b3-bbf9-563c8da25dbd@66d44188-4e30-446d-9ab1-a4ae1962a1b1/IncomingWebhook/4bc21be0b9d54eb9b32c01123e3f0616/743e9a0f-79fa-41cc-acb7-5b896ae80061/V2wvV6BoCOpaZUUhwGq9JlB7ckYlMOEwduoMVdmTUtgIg1")
        # myTeamsMessage.title("Thông báo từ CMMS system")
        # myTeamsMessage.text(f"Báo cáo Stock control tuần {self.week} ngày {self.today}")

        # for line, machines in rows:
        #     section = pymsteams.cardsection()
        #     section.addFact(line, machines)
        #     myTeamsMessage.addSection(section)

        # myTeamsMessage.send()

    def maintenance_notification(self):
        dep = "PE1"
        script_dict = {
            "total_line":f'''SELECT COUNT(DISTINCT mp.line_id) AS total
                            FROM `Maintenance_plan` mp
                            JOIN `Production_Lines` p ON p.line_id = mp.line_id
                            JOIN `Departments` d ON p.department_id = d.department_id
                            JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id
                            WHERE week = {self.week} AND d.department_id < 6 AND my.year = {self.today.year};''',

            "plan": f'''SELECT p.line_name, COUNT(p.line_name) AS plan_count
                    FROM Production_Lines p
                    JOIN Maintenance_plan mp ON p.line_id = mp.line_id
                    JOIN Departments d ON p.department_id = d.department_id
                    JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id
                    WHERE d.department_name = "{dep}" AND mp.week = {self.week} AND my.year = {self.today.year}
                    GROUP BY p.line_name;''',

            "result": f'''SELECT 
                            COALESCE(
                                COUNT(DISTINCT mp.line_id) 
                                - COUNT(DISTINCT CASE WHEN mp.status IS NULL OR mp.status = 'Overdue' THEN mp.line_id END), 0
                            ) AS complete,
                            COALESCE(
                                COUNT(mp.machine_id) 
                                - COUNT(CASE WHEN mp.status IS NULL OR mp.status = 'Overdue' THEN mp.machine_id END), 0
                            ) AS complete_mc
                        FROM (
                            SELECT 'PE1' AS department_name
                            UNION ALL SELECT 'PE2'
                            UNION ALL SELECT 'PE3'
                            UNION ALL SELECT 'PE4'
                            UNION ALL SELECT 'PE5'
                        ) AS d
                        LEFT JOIN `Departments` as dep ON dep.department_name = d.department_name
                        LEFT JOIN `Production_Lines` as p ON p.department_id = dep.department_id
                        LEFT JOIN `Maintenance_plan` as mp ON mp.line_id = p.line_id AND mp.week = {self.week}
                        LEFT JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id AND my.year = {self.today.year}
                        GROUP BY d.department_name
                        ORDER BY d.department_name;''',
        }
        total_line = []
        plan = {}

        try:
            total_line = self.DB.query(script_dict["total_line"])
            for item in ["PE1","PE2","PE3","PE4","PE5"]:
                dep = item
                plan[item] = self.DB.query(f'''SELECT p.line_name, COUNT(p.line_name) AS plan_count
                    FROM Production_Lines p
                    JOIN Maintenance_plan mp ON p.line_id = mp.line_id
                    JOIN Departments d ON p.department_id = d.department_id
                    JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id
                    WHERE d.department_name = "{dep}" AND mp.week = {self.week} AND my.year = {self.today.year}
                    GROUP BY p.line_name;''')

            result = self.DB.query(script_dict["result"])
        except:
            pass
        plan_maintenance = {}
        for group, items in plan.items():
            if not items:   
                plan_maintenance[group] = {
                    "total": 0,
                    "count": 0,
                    "lines": []
                }
                continue
            
            total = sum(i[1] for i in items)
            count = len(items)
            lines = [i[0] for i in items]

            plan_maintenance[group] = {
                "total": total,
                "count": count,
                "lines": lines}
        result_maintenance = {}
        for index in range(len(result)):
            result_maintenance[f"PE{index+1}"] = result[index]
        message_text = self.build_teams_message(total_line, plan_maintenance, result_maintenance, self.week, self.today)
        myTeamsMessage = pymsteams.connectorcard( "https://yageo.webhook.office.com/webhookb2/43c60a89-1113-44b3-bbf9-563c8da25dbd@66d44188-4e30-446d-9ab1-a4ae1962a1b1/IncomingWebhook/d320c80956a14a059df2325267bbb46a/743e9a0f-79fa-41cc-acb7-5b896ae80061/V2SWQ0CJWbCKwGcLeJ_A6I0p9FUMJVNkrYb_I-_Sm0lOw1")
        myTeamsMessage.title("Thông báo từ CMMS system")
        myTeamsMessage.text(message_text)
        myTeamsMessage.send()
    
    def build_teams_message(self,total_line, plan, result, week, today):
        if isinstance(plan, tuple):
            plan = plan[0]
        total_mc = sum(plan[g]["total"] for g in plan)
        total_lines = total_line[0][0]

        msg = []
        msg.append(f"**Báo cáo bảo trì tuần {week} ngày {today}**\n")
        msg.append(f"**Total**")
        msg.append(f"- Line: **{total_lines}**")
        msg.append(f"- Qty machines: **{total_mc}**\n")

        for g in plan:
            lines = ", ".join(plan[g]["lines"]) if plan[g]["lines"] else "-"
            msg.append(f"**{g}**")
            msg.append(f"- Line: {lines}")
            msg.append(f"- Machine plan: {plan[g]['total']}")
            msg.append(f"- Complete line: {result[g][0]}")
            msg.append(f"- Complete machine: {result[g][1]}\n")
        
        return "\n".join(msg)
    
    def call_notification(self):
        self.stock_notification()
        # self.maintenance_notification()

if __name__ == "__main__":
    db = Database_process()
    notification = Notification( DB= db)
    notification.call_notification()
    db.close()