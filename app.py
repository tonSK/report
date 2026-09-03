import streamlit as st
import pdfplumber
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
            words = page.extract_words()
            lines = defaultdict(list)
            for w in words:
                lines[round(w["top"], 1)].append(w)
            for top in sorted(lines):
                ws = sorted(lines[top], key=lambda w: w["x0"])
                all_lines.append((pno, top, ws))
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
    for pno, top, ws in lines:
        if should_skip(ws):
            continue
        if is_continuation_number_line(ws):
            if rows:
                rows[-1]["เบี้ยไม่คิดผลงาน"] = ws[0]["text"]
            continue
        row = parse_data_line(ws, current_label)
        current_label = row["ตำแหน่ง"]
        rows.append(row)
    return rows


FIELDNAMES = ["ตำแหน่ง", "รหัส", "ชื่อ-สกุล", "ค่าบำเหน็จ", "เบี้ยเครดิต",
              "เบี้ยไม่คิดผลงาน", "ปีต่อไป", "รวม", "ฐานเบี้ย", "%เบี้ยปีแรก/ฐาน"]


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

sheet_url = st.text_input("ลิงก์ Google Sheet ปลายทาง", value=DEFAULT_SHEET_URL)
worksheet_name = st.text_input("ชื่อชีต (tab) ที่จะเขียนข้อมูลลง", value="Import")
uploaded_file = st.file_uploader("เลือกไฟล์ PDF รายงานเบี้ยประกัน", type=["pdf"])

if uploaded_file and st.button("🚀 แปลงและส่งเข้า Google Sheet"):
    with st.spinner("กำลังอ่านไฟล์ PDF..."):
        rows = parse_pdf(uploaded_file)
    st.success(f"อ่านสำเร็จ พบข้อมูล {len(rows)} แถว")

    report_date = extract_date_from_filename(uploaded_file.name)
    if report_date:
        st.info(f"วันที่ของรายงาน (จากชื่อไฟล์): {report_date}")
    else:
        st.warning("หาวันที่จากชื่อไฟล์ไม่เจอ (คาดรูปแบบ dd-mm-yy ในชื่อไฟล์) — จะไม่เขียนวันที่ลง K1/K2")

    st.dataframe(rows)

    if sheet_url:
        with st.spinner("กำลังเขียนข้อมูลเข้า Google Sheet..."):
            try:
                gc = get_gsheet_client()
                sh = gc.open_by_url(sheet_url)
                try:
                    ws = sh.worksheet(worksheet_name)
                except gspread.WorksheetNotFound:
                    ws = sh.add_worksheet(title=worksheet_name, rows=1000, cols=20)

                # ล้างข้อมูลเดิมทั้งหมดในชีตนี้ก่อนเสมอ
                ws.clear()

                # เขียนหัวตารางและข้อมูลใหม่
                ws.append_row(FIELDNAMES)
                data_rows = [[r[f] for f in FIELDNAMES] for r in rows]
                ws.append_rows(data_rows)

                # เขียนหัวข้อ "วันที่" ที่ K1 และวันที่ของรายงาน (จากชื่อไฟล์) ที่ K2
                if report_date:
                    ws.update(range_name="K1", values=[["วันที่"]])
                    ws.update(range_name="K2", values=[[report_date]])

                st.success("เขียนเข้า Google Sheet เรียบร้อยแล้ว ✅")
                st.markdown(f"[เปิด Google Sheet]({sheet_url})")
            except Exception as e:
                st.error(f"เขียนเข้า Google Sheet ไม่สำเร็จ: {e}")
    else:
        st.info("ยังไม่ได้ใส่ลิงก์ Google Sheet — ดูตารางด้านบนได้เลย หรือใส่ลิงก์แล้วกดปุ่มใหม่อีกครั้ง")
