import fitz  
import cv2
import numpy as np
import os
import zxingcpp
import json
class Scan_record_process():
    def __init__(self):
        self.roi_size = {"x":100 , "y":100, "width":1500 , "height" : 1500}
        self.y = self.roi_size["y"]
        self.height = self.y + self.roi_size["height"]
        self.x = self.roi_size["x"]
        self.width = self.x + self.roi_size["width"]
        self.oneFile = None

    def scanning_oneFile(self,path):
        pdf = fitz.open(path)
        readed_code = []
        for page_num in range(pdf.page_count):
            page = pdf[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(8, 8)) 
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if pix.n == 4:
                img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
            roi = img[self.y:self.height, self.x:self.width]
            gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            result = zxingcpp.read_barcode(gray) # ISO/IEC 18004
            if result is None:  
                readed_code.append("NA")
            else:
                readed_code.append(result.text)
            # else:  
            #     try:
            #         data = json.loads(result.text)
            #         if isinstance(data, dict):
            #             data["page_num"] = page_num 
            #             result_text_with_page = json.dumps(data, ensure_ascii=False)
            #             readed_code.append(result_text_with_page)
            #         else:
            #             readed_code.append(result.text)
                # except json.JSONDecodeError:
                #     readed_code.append(result.text)
        return readed_code

    def scanning_dir(self,path):
        pdf = fitz.open(path)
        page = pdf[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(8, 8)) 
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
        img2 = img.copy()
        roi = img2[self.y:self.height, self.x:self.width]
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        result = zxingcpp.read_barcode(gray)
        if result:
            return result.text
        else:
            return "text"
    
                
    def split_pdf(self,input_file, start, end, output_file):
        doc = fitz.open(input_file)
        new_doc = fitz.open()
        for i in range(start, end + 1):
            page = doc[i]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5),alpha=False)  
            new_page = new_doc.new_page(width=pix.width, height=pix.height)
            new_page.insert_image(new_page.rect, pixmap=pix)
        new_doc.save(output_file,        
                     deflate=True,
                        garbage=4)
        new_doc.close()
        doc.close()
    
    def return_form_page(self,path):
        pdf = fitz.open(path)
        return pdf.page_count

    def paths(self,link):
        self.link = link
        self.path = None
        if self.link != "" and not self.link.lower().endswith(".pdf"):
            dirs = os.listdir(self.link)
            self.path = [os.path.join(self.link,dir) for dir in dirs]
            self.oneFile = False
        else:
            self.path = [self.link,]
            self.oneFile = True
        return self.path
    
if __name__ == "__main__":
    scanner = Scan_record_process()
    import os
    import sys
    from pathlib import Path
#     PROJECT_ROOT = Path(__file__).resolve().parents[1]
#     if str(PROJECT_ROOT) not in sys.path:
#         sys.path.insert(0, str(PROJECT_ROOT))
#     from Database.MariaDB import Database_process
    # paths = scanner.paths(r"//172.30.73.156/share/14032026154102.pdf")
#     database = Database_process()
#     plan_query = ''' SELECT m.machine_code,m.machine_name,d.department_name,pl.line_name,mp.maintenance_date,(SELECT "VINH") as technical_name, mf.page_num FROM maintenance_plan as mp
#                     JOIN machines AS m ON mp.machine_id = m.machine_id
#                     JOIN production_lines AS pl ON m.line_id = pl.line_id
#                     JOIN departments AS d ON pl.department_id = d.department_id
#                     JOIN maintenance_form_register AS mfr ON mfr.machine_id = m.machine_id
#                     JOIN maintenance_form AS mf ON mfr.form_id = mf.form_id
#                     WHERE pl.line_name = "Z06" AND mp.maintenance_date = "2026-03-05";'''
#     output_file = database.query(sql = plan_query)
#     start_page_dict = {
#         "688714": 0,
#         "MCG151765": 2,
#         "ACS-048": 4,
#         "MCG152284": 5,
#         "MCG150267": 6,
#         "ACS-049": 8,
#         "MCG152736": 9,
#         "687936": 10,
#         "ACS-051": 12,
#         "MCG151862": 13,
#         "688250": 14,
#         "ACS-053": 16,
#         "MCG152750": 17,
#         "MCG152117": 18,
#         "ACS-052": 20,
#         "MCG160005": 21,
#         "MCG150262": 22,
#         "ACS-050": 24,
#         "MCG151567": 25,
#         "687793": 26,
#         "ACS-072": 28,
#         "MCG151573": 29,
#         "ZAJ-037": 30,
#         "ZAJ-036": 31,
#         "ZAJ-034": 32,
#         "ZAJ-033": 33,
#         "ZAJ-032": 34,
#         "ZAJ-031": 35,
#         "TAJ-006": 36,
#         "ZAJ-041": 37,
#         "MCG152876": 38,
#         "MCG152191": 39,
#         "MCG170292": 40,
#         "MCG170121": 41,
#         "MCG152165": 42,
#         "ZJ-093": 43,
#         "ZJ-101": 44,
#         "ZJ-195": 45,
#         "ZJ-088": 46,
#         "ZJ-086": 47,
#         "ZJ-092": 48,
#         "ZJ-090": 49,
#         "ZJ-141": 50,
#         "ZJ-221": 51,
#         "ZJ-139": 52,
#         "ZJ-220": 53}
#     for output in output_file:
#         file_name = rf"X:\Scan\{output[0]}_{output[1]}_{output[2]}_{output[3]}_{output[5]}_{output[4]}.pdf"
#         scanner.split_pdf(r"\\172.30.73.156\share\14032026154102.pdf",start_page_dict[output[0]],start_page_dict[output[0]]+output[6]-1,file_name)  

    split_list = [
        {"code":"MCG240007", "file_name":"MCG240007_Arrange_Products_Into_Trays_MC_PE3_FA27_LONG_2026-09-29.pdf", "page_start":0, "page_end":1},
        {"code":"MCG250152", "file_name":"MCG250152_Industrial_Electron_Microscop_GP660V_PE3_FA27_LONG_2026-09-29.pdf", "page_start":2, "page_end":2},
        {"code":"1741213-5/7", "file_name":"1741213-5-7_Wire_cutting_machine_PE3_FA27_LONG_2026-09-29.pdf", "page_start":3, "page_end":4},
        {"code":"MCG220016", "file_name":"MCG220016_Vacuum_Cleaner_BF502_PE3_FA27_LONG_2026-09-29.pdf", "page_start":3, "page_end":4},
        {"code":"1741213-6/7", "file_name":"1741213-6-7_Piece_cutting_machine_PE3_FA27_LONG_2026-09-29.pdf", "page_start":5, "page_end":6},
        {"code":"MCG230021", "file_name":"MCG230021_Belt_Over_PE3_FA27_LONG_2026-09-29.pdf", "page_start":7, "page_end":8},
        {"code":"1741213-3/7", "file_name":"1741213-3-7_LASER_PEELING_MACHINE_PE3_FA27_LONG_2026-09-29.pdf", "page_start":9, "page_end":10},
        {"code":"1741213-4/7", "file_name":"1741213-4-7_Laser_welding_machine_PE3_FA27_LONG_2026-09-29.pdf", "page_start":11, "page_end":12},
        {"code":"1741213-4/7+1", "file_name":"1741213-4-7+1_Power_suplly_for_welding_PE3_FA27_LONG_2026-09-29.pdf", "page_start":11, "page_end":12},
        {"code":"1741213-7/7", "file_name":"1741213-7-7_Industrial_chiller_machine_PE3_FA27_LONG_2026-09-29.pdf", "page_start":11, "page_end":12},
        {"code":"1741211-1/2", "file_name":"1741211-1-2_Terminal_Forming_machine(6)_PE3_FA27_LONG_2026-09-29.pdf", "page_start":13, "page_end":14},
        {"code":"1741211-2/2", "file_name":"1741211-2-2_Terminal_Forming_machine(6)_PE3_FA27_LONG_2026-09-29.pdf", "page_start":15, "page_end":16},
        {"code":"1741213-2/7", "file_name":"1741213-2-7_Winding_Machine_PE3_FA27_LONG_2026-09-29.pdf", "page_start":17, "page_end":18},
        {"code":"1741213-1/7", "file_name":"1741213-1-7_Winding_Machine_PE3_FA27_LONG_2026-09-29.pdf", "page_start":19, "page_end":20},
        {"code":"1741210", "file_name":"1741210_Multi-function_Molding_machine_PE3_FA27_LONG_2026-09-29.pdf", "page_start":21, "page_end":22},
        {"code":"MCG250117", "file_name":"MCG250117_Vacuum_Cleaner_VF-5N_PE3_FA27_LONG_2026-09-29.pdf", "page_start":21, "page_end":22},
        {"code":"1818214", "file_name":"1818214_DRYING_MACHINE_(OVEN)-YAMATO-411I_PE3_FA27_LONG_2026-09-29.pdf", "page_start":23, "page_end":24},
        {"code":"1818218", "file_name":"1818218_DRYING_MACHINE_(OVEN)-YAMATO-411I_PE3_FA27_LONG_2026-09-29.pdf", "page_start":25, "page_end":26},
        {"code":"1741209-2/15", "file_name":"1741209-2-15_Eletrical_Inspection_Tapping_Machine_PE3_FA27_LONG_2026-09-29.pdf", "page_start":27, "page_end":28},
    ]

    for item in split_list:
        scanner.split_pdf(
            r"C:\Users\2173452100291\Downloads\02102026133716.pdf",
            item["page_start"],
            item["page_end"],
            rf"X:\Scan_Result_M\{item['file_name']}"
        )