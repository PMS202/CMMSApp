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
            if not result:  
                readed_code.append("text")
            else:  
                try:
                    data = json.loads(result.text)
                    if isinstance(data, dict):
                        data["page_num"] = page_num 
                        result_text_with_page = json.dumps(data, ensure_ascii=False)
                        readed_code.append(result_text_with_page)
                    else:
                        readed_code.append(result.text)
                except json.JSONDecodeError:
                    readed_code.append(result.text)
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
    
# if __name__ == "__main__":
    # scanner = Scan_record_process()
    # paths = scanner.paths(r"\\172.30.73.156\share\13012026190443.pdf")
    # scanner.split_pdf(r"C:\Users\2173452100291\Documents\program\30012026134600.pdf",60,61,r"X:\Scan\MCG150061_Forming_PE1_WSPE1_2026-01-19.pdf")  
    # scanner.scanning_oneFile(r"\\172.30.73.156\nd_ie2\Noise Device - IE Data\DANH MUC THIET BI BAO TRI\2026\F07.pdf")    