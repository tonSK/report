import streamlit as st
import pdfplumber
import openpyxl
import re
import json
from collections import defaultdict
import gspread
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------------------------
# ตั้งค่าหน้าเว็บ
# ---------------------------------------------------------------------------
st.set_page_config(page_title="แปลงรายงานเบี้ยประกัน PDF -> Google Sheet", page_icon="📄")
st.title("📄 แปลงรายงานเบี้ยประกัน PDF ➜ Google Sheet")
st.write("อัปโหลดไฟล์ PDF รายงานเบี้ยประกัน ระบบจะอ่านและเขียนข้อมูลเข้า Google Sheet ให้อัตโนมัติ")

# ---------------------------------------------------------------------------
# ตรรกะแปลง PDF (เหมือนที่ทดสอบไว้แล้ว)
# ---------------------------------------------------------------------------
COLS = [
    ("ค่าบำเหน็จ", 225, 300),
    ("เบี้ยเครดิต", 300, 370),
    ("ปีต่อไป", 370, 432),
    ("รวม", 432, 496),
    ("ฐานเบี้ย", 496, 540),
    ("%เบี้ยปีแรก/ฐาน", 540, 999),
]
NUM_RE = re.compile(r"^-?[\d,]+(\.\d+)?$")
SKIP_SUBSTRINGS = [
    "วันที่ออกรายงาน", "รายงานเบี้ยประกัน", "ข้อมูลเบี้ยตัด",
    "หมายเหตุ", "ตัวเลขด้านล่าง", "ที่มา", "โครงการนักขายดิจิทัล",
    "เบี้ยประกันรับ", "ฐานเบี้ย", "ค่าบำเหน็จ",
]
HEADER_WORDS = {"ปีแรก", "ปีต่อไป", "รวม", "แรก/ฐาน", "%", "เบี้ยปี", "ตำแหน่ง", "ชื่อ", "-", "สกุล"}


def is_number(tok):
    return bool(NUM_RE.match(tok))


def col_for_x(x0):
    for name, lo, hi in COLS:
        if lo <= x0 < hi:
            return name
    return None


def extract_lines(file_obj):
    all_lines = []
    with pdfplumber.open(file_obj) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            words = sorted(page.extract_words(), key=lambda w: w["top"])
            groups = []
            current_group = []
            current_top = None
            # รวมคำเป็นบรรทัดเดียวกันถ้า top ใกล้กันในระยะ 3pt แทนการปัดเศษเป๊ะๆ
            # เพราะบางไฟล์คำนำหน้าตำแหน่ง (ศูนย์/หน่วย) จะสูงกว่าตัวเลขในแถวเดียวกันเล็กน้อย
            for w in words:
                if current_top is not None and abs(w["top"] - current_top) <= 3:
                    current_group.append(w)
                else:
                    if current_group:
                        groups.append(current_group)
                    current_group = [w]
                    current_top = w["top"]
            if current_group:
                groups.append(current_group)
            for g in groups:
                ws = sorted(g, key=lambda w: w["x0"])
                all_lines.append((pno, g[0]["top"], ws))
    return all_lines


def should_skip(ws):
    texts = [w["text"] for w in ws]
    joined = " ".join(texts)
    if any(s in joined for s in SKIP_SUBSTRINGS):
        return True
    if set(texts) <= HEADER_WORDS:
        return True
    return False


def is_continuation_number_line(ws):
    texts = [w["text"] for w in ws]
    if not texts:
        return False
    return all(is_number(t) for t in texts)


def parse_data_line(ws, current_label):
    texts_with_x = [(w["text"], w["x0"]) for w in ws]
    idx = 0
    label = current_label
    if texts_with_x and texts_with_x[0][1] < 40:
        label = texts_with_x[0][0]
        idx = 1
    code = ""
    if idx < len(texts_with_x) and texts_with_x[idx][1] < 95:
        code = texts_with_x[idx][0]
        idx += 1
    name_tokens = []
    while idx < len(texts_with_x) and texts_with_x[idx][1] < 225:
        name_tokens.append(texts_with_x[idx][0])
        idx += 1
    name = " ".join(name_tokens).replace(" - ", "-").strip()

    # เผื่อไฟล์บางแบบ (เช่นรายงานรายวัน) ตำแหน่ง x ของ "รหัส" ไม่ตรงช่วงที่คาดไว้
    # จึงหลุดไปรวมอยู่ในชื่อแทน เช่น "98 สุภัทรชัย โกศาคาร"
    # ถ้ายังไม่มีรหัส และคำแรกของชื่อเป็นตัวเลขล้วน ให้แยกออกมาเป็นรหัส
    if not code and name:
        first_word, _, rest = name.partition(" ")
        if re.match(r"^[\d.\-/]+$", first_word):
            code = first_word
            name = rest.strip()

    # บางครั้งบรรทัด "ตัวแทน" แถวแรกของกลุ่มไม่มีคำว่า "ตัวแทน" กำกับ (หลุดจากการพิมพ์)
    # แต่รหัสตัวแทนจะเป็นตัวเลขล้วน 5 หลักเสมอ (ต่างจากรหัสศูนย์/หน่วยที่มีขีด และรหัสภาคที่มีจุด)
    # จึงใช้รูปแบบรหัสยืนยันตำแหน่งให้แม่นกว่าการไล่ตามบรรทัดก่อนหน้าอย่างเดียว
    if re.match(r"^\d{5}$", code):
        label = "ตัวแทน"

    row = {
        "ตำแหน่ง": label, "รหัส": code, "ชื่อ-สกุล": name,
        "ค่าบำเหน็จ": "", "เบี้ยเครดิต": "", "เบี้ยไม่คิดผลงาน": "",
        "ปีต่อไป": "", "รวม": "", "ฐานเบี้ย": "", "%เบี้ยปีแรก/ฐาน": "",
    }
    for text, x0 in texts_with_x[idx:]:
        col = col_for_x(x0)
        if col:
            row[col] = text
    return row


def parse_pdf(file_obj):
    lines = extract_lines(file_obj)
    rows = []
    current_label = ""
    # ใช้ไล่ตามลำดับสังกัด: ฝ่าย > ภาค/ภาค(GL) > ศูนย์ > หน่วย > ตัวแทน
    # เวลาเจอตำแหน่งที่สูงกว่าใหม่ ต้องล้างตำแหน่งที่ต่ำกว่าทั้งหมด เพราะยังไม่มีข้อมูลของสังกัดใหม่
    current_fah = ""
    current_pak = ""
    current_soon = ""
    current_nuay = ""

    for pno, top, ws in lines:
        if should_skip(ws):
            continue
        if is_continuation_number_line(ws):
            if rows:
                rows[-1]["เบี้ยไม่คิดผลงาน"] = ws[0]["text"]
            continue

        row = parse_data_line(ws, current_label)
        current_label = row["ตำแหน่ง"]
        name = row["ชื่อ-สกุล"]

        if current_label == "ฝ่าย":
            current_fah = name
            current_pak = ""
            current_soon = ""
            current_nuay = ""
        elif current_label in ("ภาค", "ภาค(GL)"):
            current_pak = name
            current_soon = ""
            current_nuay = ""
        elif current_label == "ศูนย์":
            current_soon = name
            current_nuay = ""
        elif current_label == "หน่วย":
            current_nuay = name
        # current_label == "ตัวแทน" -> ไม่ต้องเปลี่ยนอะไร ใช้สังกัดปัจจุบันตามที่ไล่มา

        row["ฝ่าย"] = current_fah
        row["ภาค"] = current_pak
        row["ศูนย์"] = current_soon
        row["หน่วย"] = current_nuay

        rows.append(row)
    return rows


FIELDNAMES = ["ตำแหน่ง", "รหัส", "ชื่อ-สกุล", "ฝ่าย", "ภาค", "ศูนย์", "หน่วย",
              "ค่าบำเหน็จ", "เบี้ยเครดิต", "เบี้ยไม่คิดผลงาน", "ปีต่อไป", "รวม",
              "ฐานเบี้ย", "%เบี้ยปีแรก/ฐาน"]


# ---------------------------------------------------------------------------
# ตรรกะแปลงไฟล์ xlsx ประกอบ (ต้องมาคู่กับ PDF)
# ---------------------------------------------------------------------------
def parse_xlsx_col_a(text):
    """แปลงข้อความคอลัมน์ A เช่น 'ตัวแทน A07106274Z : นายพงศ์ปณต ชลชีพ'
    ให้ได้ (ตำแหน่ง, รหัส(ตัดตัวอักษรหัวท้ายออก), ชื่อ-สกุล)"""
    left, _, name = text.partition(":")
    left = left.strip()
    name = name.strip()
    parts = left.split(None, 1)
    label = parts[0] if parts else ""
    code_raw = parts[1].strip() if len(parts) > 1 else ""
    middle = code_raw
    if middle and middle[0].isalpha():
        middle = middle[1:]
    if middle and middle[-1].isalpha():
        middle = middle[:-1]
    return label, middle, name


def match_code_for_label(label, code):
    """รหัสในไฟล์ xlsx ของ 'ตัวแทน' จะมีรหัสสาขานำหน้าติดมาด้วย (เช่น 07106274)
    ให้ตัดเหลือ 5 หลักท้ายเพื่อเทียบกับรหัสตัวแทนใน Import (เช่น 06274)
    ระดับอื่น (ฝ่าย/ภาค/ศูนย์/หน่วย) ใช้รหัสทั้งเส้นเทียบตรงๆ"""
    if label == "ตัวแทน":
        return code[-5:]
    return code


def parse_xlsx(file_obj):
    """อ่านไฟล์ xlsx คืนค่า list of (label, match_code, colC_value)"""
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.worksheets[0]
    result = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        col_a, col_c = row[0], row[2]
        if not col_a:
            continue
        label, code, _name = parse_xlsx_col_a(str(col_a))
        match_code = match_code_for_label(label, code)
        result.append((label, match_code, col_c))
    return result


def extract_date_from_filename(filename):
    """Find a dd-mm-yy date inside the filename, e.g.
    'Lallmonthpremium31-08-69.pdf' -> '31/08/2569'."""
    match = re.search(r"(\d{2})-(\d{2})-(\d{2})", filename)
    if not match:
        return None
    dd, mm, yy = match.groups()
    return f"{dd}/{mm}/25{yy}"

# ---------------------------------------------------------------------------
# เชื่อมต่อ Google Sheets โดยใช้ Service Account (เก็บไว้ใน Streamlit Secrets)
# ---------------------------------------------------------------------------
def get_gsheet_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds_dict = json.loads(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


# ---------------------------------------------------------------------------
# หน้าเว็บ
# ---------------------------------------------------------------------------
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1fYTibLa8riOyUPzPiu_579f5Ntu9_4B5yGLQJ8-LnA4/edit?usp=sharing"

# ไม่แสดงลิงก์ Google Sheet และชื่อชีตในหน้าแอป (ป้องกันคนอื่นเห็น) — ใช้ค่าคงที่แทน
sheet_url = DEFAULT_SHEET_URL
worksheet_name = "Import"

# ใช้ตัวนับต่อท้าย key ของ file_uploader เพื่อให้ปุ่ม Clear
# สามารถ "รีเซ็ต" ไฟล์ที่แนบไว้ได้ (เปลี่ยน key ทำให้ widget เริ่มใหม่)
if "uploader_generation" not in st.session_state:
    st.session_state["uploader_generation"] = 0
gen = st.session_state["uploader_generation"]

uploaded_file = st.file_uploader(
    "เลือกไฟล์ PDF รายงานเบี้ยประกัน",
    type=["pdf"],
    accept_multiple_files=False,
    key=f"pdf_uploader_{gen}",
)
uploaded_xlsx = st.file_uploader(
    "แนบไฟล์ xlsx ประกอบ (ไม่บังคับ — ต้องมีข้อมูลใน Import จากไฟล์ PDF อยู่ก่อนแล้ว)",
    type=["xlsx"],
    accept_multiple_files=False,
    key=f"xlsx_uploader_{gen}",
)


def process_and_write(file, target_sheet_url, target_worksheet_name):
    with st.spinner("กำลังอ่านไฟล์ PDF..."):
        rows = parse_pdf(file)
    st.success(f"อ่านสำเร็จ พบข้อมูล {len(rows)} แถว")

    report_date = extract_date_from_filename(file.name)
    if report_date:
        st.info(f"วันที่ของรายงาน (จากชื่อไฟล์): {report_date}")
    else:
        st.warning("หาวันที่จากชื่อไฟล์ไม่เจอ (คาดรูปแบบ dd-mm-yy ในชื่อไฟล์) — จะไม่เขียนวันที่ลง O1")

    st.dataframe(rows)

    if target_sheet_url:
        with st.spinner("กำลังเขียนข้อมูลเข้า Google Sheet..."):
            try:
                gc = get_gsheet_client()
                sh = gc.open_by_url(target_sheet_url)
                try:
                    ws = sh.worksheet(target_worksheet_name)
                except gspread.WorksheetNotFound:
                    ws = sh.add_worksheet(title=target_worksheet_name, rows=1000, cols=20)

                # ล้างข้อมูลเดิมทั้งหมดในชีตนี้ก่อนเสมอ
                ws.clear()

                # เขียนหัวตารางและข้อมูลใหม่
                ws.append_row(FIELDNAMES)
                data_rows = [[r[f] for f in FIELDNAMES] for r in rows]
                ws.append_rows(data_rows)

                # เขียนวันที่ของรายงาน (จากชื่อไฟล์) ที่ O1
                if report_date:
                    ws.update(range_name="O1", values=[[report_date]])

                st.success("เขียนเข้า Google Sheet เรียบร้อยแล้ว ✅")
                st.markdown(f"[เปิด Google Sheet]({target_sheet_url})")
            except Exception as e:
                st.error(f"เขียนเข้า Google Sheet ไม่สำเร็จ: {e}")
    else:
        st.info("ยังไม่ได้ใส่ลิงก์ Google Sheet — ดูตารางด้านบนได้เลย หรือใส่ลิงก์แล้วลองใหม่อีกครั้ง")


def process_xlsx_merge(xlsx_file, target_sheet_url, target_worksheet_name):
    with st.spinner("กำลังอ่านไฟล์ xlsx..."):
        xlsx_rows = parse_xlsx(xlsx_file)
    st.success(f"อ่านไฟล์ xlsx สำเร็จ พบข้อมูล {len(xlsx_rows)} แถว")

    try:
        gc = get_gsheet_client()
        sh = gc.open_by_url(target_sheet_url)
        ws = sh.worksheet(target_worksheet_name)
    except Exception as e:
        st.error(f"เปิด Google Sheet ไม่สำเร็จ: {e}")
        return

    # อ่านข้อมูลปัจจุบันในชีต Import เพื่อหาว่าแต่ละแถวอยู่บรรทัดไหน
    existing = ws.get_all_values()
    if len(existing) < 2:
        st.warning("ยังไม่มีข้อมูลใน Import — กรุณาแนบไฟล์ PDF คู่กัน หรืออัปโหลด PDF ก่อน")
        return

    header = existing[0]
    try:
        idx_pos = header.index("ตำแหน่ง")
        idx_code = header.index("รหัส")
    except ValueError:
        st.error("ไม่พบคอลัมน์ ตำแหน่ง/รหัส ใน Import — ตรวจสอบว่าอัปโหลด PDF ไปแล้ว")
        return

    row_lookup = {}
    for i, row in enumerate(existing[1:], start=2):  # แถวที่ 2 เป็นต้นไป (แถวจริงในชีต)
        if len(row) > max(idx_pos, idx_code):
            key = (row[idx_pos], row[idx_code].replace("-", ""))
            row_lookup[key] = i

    updates = []
    matched = 0
    for label, match_code, col_c in xlsx_rows:
        key = (label, match_code)
        if key in row_lookup:
            row_num = row_lookup[key]
            try:
                value = f"{round(float(col_c) * 100, 2):.2f}%"
            except (TypeError, ValueError):
                value = col_c
            updates.append({"range": f"P{row_num}", "values": [[value]]})
            matched += 1

    if updates:
        # เพิ่มหัวข้อคอลัมน์ P1
        updates.append({"range": "P1", "values": [["%ผลบังคับ"]]})
        with st.spinner("กำลังเขียนคอลัมน์ P เข้า Google Sheet..."):
            ws.batch_update(updates)
        st.success(f"จับคู่และเขียนคอลัมน์ P สำเร็จ {matched} แถว จากทั้งหมด {len(xlsx_rows)} แถวในไฟล์ xlsx")
    else:
        st.warning("ไม่พบแถวที่จับคู่ได้เลย — ตรวจสอบว่าอัปโหลด PDF เดือนเดียวกันไปแล้วหรือยัง")


ALLOWED_FILENAME_KEYWORDS = ("monthpremium", "lalldailypremium")


def is_allowed_filename(filename):
    lower = filename.lower()
    return any(kw in lower for kw in ALLOWED_FILENAME_KEYWORDS)


col_import, col_clear = st.columns(2)
with col_import:
    import_clicked = st.button("🚀 Import", use_container_width=True)
with col_clear:
    clear_clicked = st.button("🧹 Clear", use_container_width=True)

if clear_clicked:
    st.session_state["uploader_generation"] += 1
    st.rerun()

if import_clicked:
    if not uploaded_file and not uploaded_xlsx:
        st.warning("กรุณาแนบไฟล์อย่างน้อย 1 ไฟล์ก่อนกด Import")
    else:
        if uploaded_file:
            if not is_allowed_filename(uploaded_file.name):
                st.error(
                    "ชื่อไฟล์ PDF นี้ไม่ตรงรูปแบบที่รองรับ — ต้องมีคำว่า \"monthpremium\" หรือ "
                    "\"Lalldailypremium\" อยู่ในชื่อไฟล์ ระบบจะไม่ประมวลผลไฟล์นี้"
                )
            else:
                process_and_write(uploaded_file, sheet_url, worksheet_name)

        if uploaded_xlsx:
            process_xlsx_merge(uploaded_xlsx, sheet_url, worksheet_name)
