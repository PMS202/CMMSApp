import win32print
import fitz  
import segno
from io import BytesIO
import os,sys
import subprocess
import json

def resource_path(relative_path):
    if getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.argv[0]) 
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))  
    return os.path.normpath(os.path.join(base_path, relative_path))

class Printer_process():
    def __init__(self):
        self.printers_list = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        exe_dir = os.path.dirname(sys.argv[0]) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
        app_name = os.path.splitext(os.path.basename(sys.argv[0]))[0] if getattr(sys, "frozen", False) else "CMMSApp"
        dist_dir = os.path.join(exe_dir, f"{app_name}.dist")
        candidates = [
            resource_path(r"SumatraPDF-3.5.2-64\SumatraPDF-3.5.2-64.exe"),               
            os.path.join(dist_dir, r"SumatraPDF-3.5.2-64\SumatraPDF-3.5.2-64.exe"),      
            os.path.join(os.path.dirname(exe_dir), r"SumatraPDF-3.5.2-64\SumatraPDF-3.5.2-64.exe"),  
        ]
        self.sumatra_path = next((p for p in candidates if os.path.exists(p)), None)
        if not self.sumatra_path:
            raise FileNotFoundError("Không tìm thấy SumatraPDF-3.5.2-64.exe. Hãy đặt file cạnh exe hoặc cấu hình đường dẫn.")
    
    def choice_printer(self,printer_name):
        self.printer_name = printer_name
        
    def send_to_printer(self,input_pdf, data, attached_machine = None, file_index = 0 ,page_number=0, x=10, y=25, size=50):
        try:
            qr_json = json.dumps({
                    "machine_code": data[0],
                    "machine_name":data[1],
                    "group":data[2],
                    "line": data[3],
                    "technical": data[4],
                    "maintenance_date": data[5],
                    "attached_machine": "" if attached_machine is None else attached_machine
                })
            qr = segno.make(qr_json, error='h')  #  ISO/IEC 18004
            buffer = BytesIO()
            qr.save(buffer, kind='png', scale=10)  
            qr_bytes = buffer.getvalue()

            self.doc = fitz.open(input_pdf)
            page = self.doc[page_number]
            if page.rotation != 0:
                page.set_rotation(0)
            rect = fitz.Rect(x, y, x + size, y + size)

            page.insert_image(rect, stream=qr_bytes)
            page.insert_text((x+50, y+10), f"Code: {data[0]}",         
                            fontsize=8,
                            fontname="helv",
                            color=(0, 0, 0)  
                        )
            page.insert_text((x+50, y+20), f"Machine name: {data[1]}",      
                            fontsize=8,
                            fontname="helv",
                            color=(0, 0, 0)  
                        )
            page.insert_text((x+50, y+30), f"Line: {data[3]}",      
                            fontsize=8,
                            fontname="helv",
                            color=(0, 0, 0)  
                        )
            page.insert_text((x+50, y+40), f"Techinical: {data[4]}",      
                            fontsize=8,
                            fontname="helv",
                            color=(0, 0, 0)  
                        )
            page.insert_text((x+50, y+50), f"Date: {data[5]}",      
                            fontsize=8,
                            fontname="helv",
                            color=(0, 0, 0)  
                        )
            if attached_machine is not None:
                page.insert_text((x+50, y+60), f"Attached equipment: {', '.join(attached_machine)}",      
                                fontsize=8,
                                fontname="helv",
                                color=(0, 0, 0)  
                            )
            temp_dir = os.path.join(os.getcwd(), "Temp")
            os.makedirs(temp_dir, exist_ok=True)

            temp_pdf_path = os.path.join(temp_dir, f"print_job.pdf")
            self.doc.save(temp_pdf_path)
            self.doc.close()
            printer = getattr(self, "printer_name", None)
            cmd = [self.sumatra_path, "-silent"]
            cmd += ["-print-settings", "fit,portrait"]
            if printer:
                cmd += ["-print-to", printer, temp_pdf_path]
            else:
                cmd += ["-print-to-default", temp_pdf_path]
            subprocess.run(cmd, check=True)
        except Exception as e:
            raise e


