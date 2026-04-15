import streamlit as st
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

def parse_numbering(docx_zip):
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

def parse_styles(docx_zip):
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

def get_list_string(num_id, ilvl, numbering_data, list_counters):
    if not numbering_data: return ""
    abs_id = numbering_data['numMap'].get(num_id)
    if abs_id is None: return ""
    
    abstract_map = numbering_data['abstractNumMap'].get(abs_id)
    if not abstract_map: return ""
    
    lvl_info = abstract_map.get(ilvl)
    if not lvl_info: return ""
    
    if num_id not in list_counters: list_counters[num_id] = {}
    
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

def process_docx(file_buffer):
    with zipfile.ZipFile(file_buffer, 'r') as docx_zip:
        numbering_data = parse_numbering(docx_zip)
        styles_data = parse_styles(docx_zip)
        list_counters = {}
        
        # ฟังก์ชันสกัดข้อความปกติ
        def extract_raw_from_xml(filename):
            if filename not in docx_zip.namelist(): return ""
            xml_content = docx_zip.read(filename)
            root = ET.fromstring(xml_content)
            
            text_result = ""
            for p in root.findall(f".//{w_tag('p')}"):
                p_text = ""
                for elem in p.iter():
                    if elem.tag == w_tag('t') and elem.text:
                        p_text += elem.text
                    elif elem.tag == w_tag('tab'):
                        p_text += "\t"
                    elif elem.tag in[w_tag('br'), w_tag('cr')]:
                        if elem.attrib.get(w_tag('type')) != 'page':
                            p_text += "\n"
                    elif elem.tag == w_tag('sym'):
                        p_text += "•"
                text_result += p_text + "\n"
            return text_result

        # ฟังก์ชันสกัดข้อความแบบ "แบ่งหน้ากระดาษ (Page by Page)"
        def extract_pages_from_xml(filename):
            if filename not in docx_zip.namelist(): return[]
            xml_content = docx_zip.read(filename)
            root = ET.fromstring(xml_content)
            
            pages =[]
            current_page_text = ""
            
            for p in root.findall(f".//{w_tag('p')}"):
                p_text = ""
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
                    p_text += get_list_string(list_num_id, list_ilvl, numbering_data, list_counters)
                    
                for elem in p.iter():
                    # ตรวจจับการขึ้นหน้าใหม่ (Page Break)
                    if elem.tag == w_tag('lastRenderedPageBreak') or (elem.tag == w_tag('br') and elem.attrib.get(w_tag('type')) == 'page'):
                        if current_page_text.strip() or p_text.strip():
                            pages.append(current_page_text + p_text)
                        current_page_text = ""
                        p_text = ""
                    elif elem.tag == w_tag('t') and elem.text:
                        p_text += elem.text
                    elif elem.tag == w_tag('tab'):
                        p_text += "\t"
                    elif elem.tag in[w_tag('br'), w_tag('cr')]:
                        if elem.attrib.get(w_tag('type')) != 'page':
                            p_text += "\n"
                    elif elem.tag == w_tag('sym'):
                        p_text += "•"
                        
                current_page_text += p_text + "\n"
                
            if current_page_text.strip():
                pages.append(current_page_text)
            return pages

        # 1. ดึงข้อความ Header ชุดตั้งต้น (Master Header)
        master_header = ""
        seen_headers = set()
        header_files = sorted([f for f in docx_zip.namelist() if re.match(r"^word/header\d+\.xml$", f)])
        for h in header_files:
            extracted = extract_raw_from_xml(h)
            trimmed = extracted.strip()
            if trimmed and trimmed not in seen_headers:
                seen_headers.add(trimmed)
                master_header += extracted + "\n"
                
        # 2. ดึงข้อความ Footer ชุดตั้งต้น (Master Footer)
        master_footer = ""
        seen_footers = set()
        footer_files = sorted([f for f in docx_zip.namelist() if re.match(r"^word/footer\d+\.xml$", f)])
        for f in footer_files:
            extracted = extract_raw_from_xml(f)
            trimmed = extracted.strip()
            if trimmed and trimmed not in seen_headers and trimmed not in seen_footers:
                seen_footers.add(trimmed)
                master_footer += extracted + "\n"

        # 3. สกัดเนื้อหาหลักออกเป็นรายหน้า (Pages)
        pages = extract_pages_from_xml("word/document.xml")
        
        # เพิ่มเชิงอรรถ/อ้างอิง ไว้ที่ท้ายสุดของเอกสาร
        notes_text = extract_raw_from_xml("word/footnotes.xml") + extract_raw_from_xml("word/endnotes.xml")
        if notes_text.strip() and pages:
            pages[-1] += "\n\n--- [เชิงอรรถ/อ้างอิง] ---\n" + notes_text.strip()

        # ===============================================
        # ฟังก์ชันปรับเปลี่ยนตัวเลขหน้าใน Header อัตโนมัติ (AI Logic)
        # ===============================================
        def update_header_page_num(header_text, page_num):
            def repl(match):
                prefix = match.group(1)
                num_str = match.group(2)
                # เช็คว่าเป็นเลขไทยหรือไม่ ถ้าใช่ให้รันเลขไทย ถ้าไม่ให้รันเลขอาราบิก
                if any(c in '๐๑๒๓๔๕๖๗๘๙' for c in num_str):
                    new_num = ''.join(['๐','๑','๒','๓','๔','๕','๖','๗','๘','๙'][int(d)] for d in str(page_num))
                else:
                    new_num = str(page_num)
                return f"{prefix} {new_num}"
            
            # วิ่งหาคำว่า "หน้า 1", "หน้าที่ ๑", "Page 2" ในข้อความ
            return re.sub(r"(page|หน้า|หน้าที่)\s*([๐-๙0-9]+)", repl, header_text, flags=re.IGNORECASE)

        # 4. ประกอบร่างข้อความ (Assemble)
        final_output = ""
        if not pages:
            pages = [""] # กันเหนียวกรณีไฟล์เปล่า
            
        for i, page_text in enumerate(pages, start=1):
            # เสก Header ใหม่ตามหมายเลขหน้า
            page_header = update_header_page_num(master_header, i)
            
            if page_header.strip():
                final_output += f"--- [ส่วนหัว หน้าที่ {i}] ---\n{page_header.strip()}\n\n"
                
            final_output += f"---[เนื้อหาหลัก หน้าที่ {i}] ---\n{page_text.strip()}\n\n\n"
            
        # แปะ Footer ไว้ท้ายสุดของไฟล์เพียงชุดเดียว
        if master_footer.strip():
            final_output += f"--- [ส่วนท้าย] ---\n{master_footer.strip()}\n"
            
        return final_output.strip()

# ================= UI STREAMLIT =================
st.set_page_config(page_title="Word Text Extractor", page_icon="📄", layout="wide")

st.title("📄 เครื่องมือดึงข้อความดิบจาก Word (แยกหน้ากระดาษ)")
st.markdown("""
รองรับ **.docx** รักษารูปแบบตัวเลขและอักขระไทย **100%**
✅ จำลองการตัดหน้ากระดาษ (Page Break)
✅ แยกแสดง `[ส่วนหัว หน้าที่ 1]`, `[ส่วนหัว หน้าที่ 2]` ... พร้อมรันเลขหน้าให้อัตโนมัติ
✅ รวบ `[ส่วนท้าย]` ทั้งหมดไปแสดงไว้ล่างสุดเพียงชุดเดียวเพื่อไม่ให้รก
""")

st.divider()

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Word (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner("กำลังวิเคราะห์โครงสร้าง ตัดหน้ากระดาษ และคำนวณเลขหน้า..."):
        try:
            extracted_text = process_docx(uploaded_file)
            
            if extracted_text:
                st.success("✅ สกัดข้อความสำเร็จและแบ่งหน้าเรียบร้อยแล้ว!")
                
                st.text_area("ข้อความที่สกัดได้ (คัดลอกไปเทียบ Text Diff ได้เลย)", value=extracted_text, height=600)
                
                st.download_button(
                    label="💾 ดาวน์โหลดเป็นไฟล์ .txt",
                    data=extracted_text,
                    file_name="extracted_formatted_pages.txt",
                    mime="text/plain"
                )
            else:
                st.warning("ไม่พบข้อความใดๆ ในไฟล์นี้")
                
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")
