import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.scrolledtext as scrolledtext
import zipfile
import xml.etree.ElementTree as ET
import re

# Namespace ของ Microsoft Word XML
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def w_tag(tag):
    return f"{{{W_NS}}}{tag}"

def to_roman(num):
    if num <= 0: return str(num)
    val =[1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb =["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
    roman_num = ""
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman_num += syb[i]
            num -= val[i]
        i += 1
    return roman_num

def format_number(count, fmt):
    if not fmt: return str(count)
    f = fmt.lower()
    
    if "thainumber" in f or "thainum" in f or "thaicounting" in f:
        thai_digits =['๐','๑','๒','๓','๔','๕','๖','๗','๘','๙']
        return ''.join(thai_digits[int(d)] for d in str(count))
    if "thailetter" in f:
        th_alphabets = "กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ"
        return th_alphabets[(count - 1) % 42] if count > 0 else str(count)
        
    if fmt == "upperLetter": return chr(64 + count) if 1 <= count <= 26 else str(count)
    if fmt == "lowerLetter": return chr(96 + count) if 1 <= count <= 26 else str(count)
    if fmt == "upperRoman": return to_roman(count).upper()
    if fmt == "lowerRoman": return to_roman(count).lower()
    if fmt == "decimalZero": return f"0{count}" if count < 10 else str(count)
    if fmt == "bullet": return "•"
    
    return str(count)

class DocxExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรมสกัดข้อความดิบจากไฟล์ Word (.docx)")
        self.root.geometry("900x700")
        self.root.configure(bg="#f0f2f5")
        
        # Header
        tk.Label(root, text="📄 เครื่องมือดึงข้อความดิบจาก Word (Python Version)", 
                 font=("Segoe UI", 18, "bold"), bg="#f0f2f5", fg="#1a73e8").pack(pady=(20, 5))
        tk.Label(root, text="รองรับ .docx ดึงทุกตัวอักษร รักษาเลขไทย 100%\nจัดกลุ่ม [ส่วนหัว] บนสุด, [เนื้อหาหลัก] ตรงกลาง, [ส่วนท้าย] ล่างสุด และกรองข้อความเบิ้ลซ้ำ", 
                 font=("Segoe UI", 12), bg="#f0f2f5", fg="#555", justify="center").pack(pady=(0, 20))
        
        # Buttons Frame
        btn_frame = tk.Frame(root, bg="#f0f2f5")
        btn_frame.pack(fill="x", padx=30)
        
        tk.Button(btn_frame, text="📂 1. อัปโหลดไฟล์ Word (.docx)", font=("Segoe UI", 12, "bold"), 
                  bg="#1a73e8", fg="white", padx=10, pady=5, command=self.upload_file).pack(side="left")
        
        self.lbl_filename = tk.Label(btn_frame, text="ยังไม่ได้เลือกไฟล์...", font=("Segoe UI", 11, "italic"), bg="#f0f2f5", fg="#666")
        self.lbl_filename.pack(side="left", padx=15)
        
        # Text Area
        self.text_area = scrolledtext.ScrolledText(root, font=("Courier New", 12), height=20, width=80)
        self.text_area.pack(expand=True, fill="both", padx=30, pady=20)
        self.text_area.insert(tk.END, "ข้อความดิบทั้งหมดจะแสดงที่นี่...")
        
        # Action Buttons Frame
        action_frame = tk.Frame(root, bg="#f0f2f5")
        action_frame.pack(fill="x", padx=30, pady=(0, 20))
        
        tk.Button(action_frame, text="📋 2. คัดลอกข้อความทั้งหมด", font=("Segoe UI", 12, "bold"), 
                  bg="#34a853", fg="white", padx=10, pady=5, command=self.copy_text).pack(side="left", padx=(0, 10))
        tk.Button(action_frame, text="💾 3. เซฟเป็นไฟล์ .txt", font=("Segoe UI", 12, "bold"), 
                  bg="#fbbc05", fg="black", padx=10, pady=5, command=self.download_txt).pack(side="left")

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if not file_path: return
        
        self.lbl_filename.config(text=file_path.split("/")[-1])
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, "กำลังวิเคราะห์และจัดเรียงข้อความ (พร้อมกรองข้อความซ้ำ)...\n")
        self.root.update()
        
        try:
            self.process_docx(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการอ่านไฟล์:\n{str(e)}")
            self.text_area.delete(1.0, tk.END)

    def process_docx(self, filepath):
        with zipfile.ZipFile(filepath, 'r') as docx_zip:
            numbering_data = self.parse_numbering(docx_zip)
            styles_data = self.parse_styles(docx_zip)
            list_counters = {}
            
            def extract_raw_from_xml(filename):
                if filename not in docx_zip.namelist(): return ""
                xml_content = docx_zip.read(filename)
                root = ET.fromstring(xml_content)
                
                text_result = ""
                for p in root.findall(f".//{w_tag('p')}"):
                    p_text = ""
                    
                    # ตรวจจับ Numbering / Bullets
                    numPr = p.find(f".//{w_tag('numPr')}")
                    list_num_id = None
                    list_ilvl = 0
                    
                    if numPr is not None:
                        numIdNode = numPr.find(f"{w_tag('numId')}")
                        ilvlNode = numPr.find(f"{w_tag('ilvl')}")
                        if numIdNode is not None: list_num_id = numIdNode.attrib.get(w_tag('val'))
                        if ilvlNode is not None: list_ilvl = int(ilvlNode.attrib.get(w_tag('val'), 0))
                    else:
                        pStyle = p.find(f".//{w_tag('pStyle')}")
                        if pStyle is not None and styles_data:
                            style_id = pStyle.attrib.get(w_tag('val'))
                            if style_id in styles_data:
                                list_num_id = styles_data[style_id]['numId']
                                list_ilvl = styles_data[style_id]['ilvl']
                                
                    if list_num_id and list_num_id != "0" and numbering_data:
                        p_text += self.get_list_string(list_num_id, list_ilvl, numbering_data, list_counters)
                        
                    # ดึงข้อความและแท็บ
                    for elem in p.iter():
                        if elem.tag == w_tag('t') and elem.text:
                            p_text += elem.text
                        elif elem.tag == w_tag('tab'):
                            p_text += "\t"
                        elif elem.tag in[w_tag('br'), w_tag('cr')]:
                            p_text += "\n"
                        elif elem.tag == w_tag('sym'):
                            p_text += "•"
                            
                    text_result += p_text + "\n"
                return text_result

            # 1. จัดการส่วนหัว (Headers) + Deduplication
            header_text = ""
            seen_headers = set()
            header_files = sorted([f for f in docx_zip.namelist() if re.match(r"^word/header\d+\.xml$", f)])
            for h in header_files:
                extracted = extract_raw_from_xml(h)
                trimmed = extracted.strip()
                if trimmed and trimmed not in seen_headers:
                    seen_headers.add(trimmed)
                    header_text += extracted + "\n"
                    
            # 2. จัดการเนื้อหาหลัก (Main)
            main_text = ""
            main_text += extract_raw_from_xml("word/document.xml")
            main_text += extract_raw_from_xml("word/footnotes.xml")
            main_text += extract_raw_from_xml("word/endnotes.xml")
            
            # 3. จัดการส่วนท้าย (Footers) + Deduplication
            footer_text = ""
            seen_footers = set()
            footer_files = sorted([f for f in docx_zip.namelist() if re.match(r"^word/footer\d+\.xml$", f)])
            for f in footer_files:
                extracted = extract_raw_from_xml(f)
                trimmed = extracted.strip()
                if trimmed and trimmed not in seen_headers and trimmed not in seen_footers: # เลี่ยงส่วนท้ายที่อาจซ้ำกับส่วนหัว
                    seen_footers.add(trimmed)
                    footer_text += extracted + "\n"
                    
            # ประกอบร่าง
            final_output = ""
            if header_text.strip():
                final_output += "---[ส่วนหัว] ---\n" + header_text.strip() + "\n\n\n"
            if main_text.strip():
                final_output += "---[เนื้อหาหลัก] ---\n" + main_text.strip() + "\n\n\n"
            if footer_text.strip():
                final_output += "---[ส่วนท้าย] ---\n" + footer_text.strip() + "\n"
                
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, final_output.strip() if final_output.strip() else "ไม่พบข้อความใดๆ ในไฟล์นี้")

    def parse_numbering(self, docx_zip):
        if "word/numbering.xml" not in docx_zip.namelist(): return None
        root = ET.fromstring(docx_zip.read("word/numbering.xml"))
        
        abstract_num_map = {}
        for abs_num in root.findall(f".//{w_tag('abstractNum')}"):
            abs_id = abs_num.attrib.get(w_tag('abstractNumId'))
            lvl_map = {}
            for lvl in abs_num.findall(f".//{w_tag('lvl')}"):
                ilvl = int(lvl.attrib.get(w_tag('ilvl')))
                start = lvl.find(f"{w_tag('start')}")
                numFmt = lvl.find(f"{w_tag('numFmt')}")
                lvlText = lvl.find(f"{w_tag('lvlText')}")
                
                lvl_map[ilvl] = {
                    'start': int(start.attrib.get(w_tag('val'))) if start is not None else 1,
                    'numFmt': numFmt.attrib.get(w_tag('val')) if numFmt is not None else 'decimal',
                    'lvlText': lvlText.attrib.get(w_tag('val')) if lvlText is not None else ''
                }
            abstract_num_map[abs_id] = lvl_map
            
        num_map = {}
        for num in root.findall(f".//{w_tag('num')}"):
            num_id = num.attrib.get(w_tag('numId'))
            abs_node = num.find(f"{w_tag('abstractNumId')}")
            if abs_node is not None:
                num_map[num_id] = abs_node.attrib.get(w_tag('val'))
                
        return {'abstractNumMap': abstract_num_map, 'numMap': num_map}

    def parse_styles(self, docx_zip):
        if "word/styles.xml" not in docx_zip.namelist(): return {}
        root = ET.fromstring(docx_zip.read("word/styles.xml"))
        
        styles_map = {}
        for style in root.findall(f".//{w_tag('style')}"):
            style_id = style.attrib.get(w_tag('styleId'))
            numPr = style.find(f".//{w_tag('numPr')}")
            if numPr is not None:
                numIdNode = numPr.find(f"{w_tag('numId')}")
                ilvlNode = numPr.find(f"{w_tag('ilvl')}")
                if numIdNode is not None:
                    styles_map[style_id] = {
                        'numId': numIdNode.attrib.get(w_tag('val')),
                        'ilvl': int(ilvlNode.attrib.get(w_tag('val'))) if ilvlNode is not None else 0
                    }
        return styles_map

    def get_list_string(self, num_id, ilvl, numbering_data, list_counters):
        if not numbering_data: return ""
        abs_id = numbering_data['numMap'].get(num_id)
        if abs_id is None: return ""
        
        abstract_map = numbering_data['abstractNumMap'].get(abs_id)
        if not abstract_map: return ""
        
        lvl_info = abstract_map.get(ilvl)
        if not lvl_info: return ""
        
        if num_id not in list_counters: list_counters[num_id] = {}
        
        # Reset counters for deeper levels if we moved up a level
        last_level = list_counters[num_id].get('lastLevel', -1)
        if last_level != -1 and ilvl < last_level:
            for l in range(ilvl + 1, 10):
                if l in list_counters[num_id]:
                    del list_counters[num_id][l]
                    
        list_counters[num_id]['lastLevel'] = ilvl
        
        if ilvl not in list_counters[num_id]:
            list_counters[num_id][ilvl] = lvl_info['start']
        else:
            list_counters[num_id][ilvl] += 1
            
        if lvl_info['numFmt'] == 'bullet': return "• "
        
        result = lvl_info['lvlText']
        for i in range(ilvl + 1):
            count = list_counters[num_id].get(i)
            parent_lvl_info = abstract_map.get(i)
            
            if count is None:
                count = parent_lvl_info['start'] if parent_lvl_info else 1
                list_counters[num_id][i] = count
                
            fmt = parent_lvl_info['numFmt'] if parent_lvl_info else 'decimal'
            formatted_val = format_number(count, fmt)
            result = result.replace(f"%{i+1}", str(formatted_val))
            
        return result + " "

    def copy_text(self):
        txt = self.text_area.get(1.0, tk.END).strip()
        if not txt or txt == "ข้อความดิบทั้งหมดจะแสดงที่นี่...":
            messagebox.showwarning("เตือน", "ไม่มีข้อความให้คัดลอก กรุณาอัปโหลดไฟล์ก่อนครับ")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)
        self.root.update()
        messagebox.showinfo("สำเร็จ", "✅ คัดลอกข้อความสำเร็จ ไม่มีข้อความ Header/Footer เบิ้ลซ้ำ!")

    def download_txt(self):
        txt = self.text_area.get(1.0, tk.END).strip()
        if not txt or txt == "ข้อความดิบทั้งหมดจะแสดงที่นี่...":
            messagebox.showwarning("เตือน", "ไม่มีข้อความให้บันทึก กรุณาอัปโหลดไฟล์ก่อนครับ")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt", 
            initialfile="extracted_formatted_text.txt",
            filetypes=[("Text Files", "*.txt")]
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(txt)
            messagebox.showinfo("สำเร็จ", "✅ บันทึกไฟล์สำเร็จ!")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxExtractorApp(root)
    root.mainloop()