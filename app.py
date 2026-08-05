import streamlit as st
import pandas as pd
from docx import Document
import pdfplumber
import io
import re
from openpyxl.styles import PatternFill

# 1. 파일명 규칙 정제 함수

# 발음 파일명에서 사용하지 않을 줄임말 -> 원래 표현 매핑
# (긴 패턴을 먼저 처리해야 짧은 패턴에 의해 잘못 치환되지 않음)
def expand_abbreviations(text):
    t = text
    t = re.sub(r'\bsb\s*/\s*sth\b', 'something', t)   # sb/sth -> something
    t = re.sub(r'\bV-ing\b', 'ing', t)                # V-ing -> ing
    t = re.sub(r'\bto\s+V\b', 'to', t)                # to V -> to
    t = re.sub(r'\bsb\b', 'somebody', t)              # sb -> somebody
    t = re.sub(r'\bsth\b', 'something', t)            # sth -> something
    return t

def clean_filename(text):
    text = expand_abbreviations(text)
    text = text.replace(" ", "_")
    text = re.sub(r'[^a-zA-Z0-9_]', '', text)
    # 대소문자는 vocabulary 원문 그대로 유지 (예: protect A from B -> protect_A_from_B.mp3)
    return text + ".mp3"

# 챕터 표기 감지: "Chapter 1", "CHAPTER 01", "CH01", "CH02" 등 여러 형식 지원
CHAPTER_LINE_PATTERN = re.compile(r'(?i)^(?:chapter|ch)\.?\s*0*(\d+)\s*$')

def detect_chapter_number(line):
    m = CHAPTER_LINE_PATTERN.match(line.strip())
    if m:
        return int(m.group(1))
    return None

# 2. 워드 파싱 함수
def parse_word_content(docx_file):
    doc = Document(docx_file)
    data = []
    raw_texts = []
    
    current_chapter_num = 1
    current_week_title = "Week01/"
    current_week_num = 1
    
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    # 번호는 "1. word" (마침표)와 "01  word" (공백 2칸) 형식을 모두 지원
    entry_pattern = re.compile(r'^(?:\d+[\.\s]+)?([^:]+):\s*(.*)$')
    keyword_only_pattern = re.compile(r'(?i)^(Syn|Ant|Atn|유의어|반의어)$')
    tail_keyword_pattern = re.compile(r'(?i)^(Syn|Ant|Atn|유의어|반의어)\s*:\s*(.*)$')

    for line in paragraphs:
        raw_texts.append(line)
        
        # 챕터 감지 (Week 번호 연동)
        chapter_num = detect_chapter_number(line)
        if chapter_num is not None:
            current_week_num = chapter_num
            current_week_title = f"Week{current_week_num:02d}/"
            current_chapter_num = 1
            continue

        # 단어: 뜻 [탭/2칸 이상 공백 + Syn/Ant: 유의어] 추출
        match = entry_pattern.match(line)

        if match:
            raw_word = match.group(1).strip()

            # "Syn: xxx" 처럼 단어 없이 유의어만 있는 줄은 아래 유의어 처리 블록으로 넘김
            if not keyword_only_pattern.match(raw_word):
                rest = match.group(2)

                # 뜻과 유의어/반의어를 탭 또는 2칸 이상 공백으로 분리
                parts = re.split(r'\t|\s{2,}', rest, maxsplit=1)
                meaning = parts[0].strip()

                extra_info = ""
                if len(parts) > 1:
                    tail = parts[1].strip()
                    tail_match = tail_keyword_pattern.match(tail)
                    if tail_match:
                        extra_info = tail_match.group(2).strip()

                if re.search(r'[가-힣]', meaning):
                    sound_file = clean_filename(raw_word)

                    data.append({
                        "lesson_title": current_week_title,
                        "lesson_order_seq": current_week_num,
                        "page_order_seq": current_chapter_num,
                        "vocabulary": raw_word,
                        "vocabulary_kor": meaning,
                        "vocabulary_sound": sound_file,
                        "vocabulary_excep": extra_info,
                        "prompt": ""
                    })
                    current_chapter_num += 1
                continue
        
        # 유의어 처리 (Syn/Ant가 별도의 줄로 존재하는 예전 형식 지원)
        if any(keyword in line for keyword in ["Syn", "Atn", "Ant", "유의어", "반의어"]):
            extra_info = line.split(":", 1)[-1].strip()
            if data:
                first_extra = re.split(r'\s{2,}', extra_info)[0].strip()
                data[-1]["vocabulary_excep"] = first_extra

    return data, raw_texts

# 2-1. PDF 파싱 함수
def parse_pdf_content(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = [ln.strip() for ln in full_text.split("\n") if ln.strip()]

    data = []
    raw_texts = []

    current_chapter_num = 1
    current_week_title = "Week01/"
    current_week_num = 1

    # PDF 텍스트 추출 시 "01 word: 뜻   Syn: xxx" 처럼 번호가 마침표 없이
    # 붙는 형태와, 뜻과 Syn/Ant 사이가 한 칸 공백으로만 구분되는 경우가 많아
    # 워드용 정규식과는 별도의 패턴을 사용합니다.
    entry_pattern = re.compile(r'^\d{1,2}\s+([^:]+):\s*(.*)$')
    syn_pattern = re.compile(r'\b(Syn|Ant)\s*:\s*(.+)$')

    for line in lines:
        raw_texts.append(line)

        # 챕터 감지 (Week 번호 연동)
        chapter_num = detect_chapter_number(line)
        if chapter_num is not None:
            current_week_num = chapter_num
            current_week_title = f"Week{current_week_num:02d}/"
            current_chapter_num = 1
            continue

        # 번호 + 단어: 뜻 [Syn/Ant: 유의어] 추출
        match = entry_pattern.match(line)
        if match:
            raw_word = match.group(1).strip()
            rest = match.group(2).strip()

            syn_match = syn_pattern.search(rest)
            if syn_match:
                meaning = rest[:syn_match.start()].strip()
                extra_info = syn_match.group(2).strip()
            else:
                meaning = rest
                extra_info = ""

            if re.search(r'[가-힣]', meaning):
                sound_file = clean_filename(raw_word)

                data.append({
                    "lesson_title": current_week_title,
                    "lesson_order_seq": current_week_num,
                    "page_order_seq": current_chapter_num,
                    "vocabulary": raw_word,
                    "vocabulary_kor": meaning,
                    "vocabulary_sound": sound_file,
                    "vocabulary_excep": extra_info,
                    "prompt": ""
                })
                current_chapter_num += 1

    return data, raw_texts

# --- 3. Streamlit UI (여기서 변수가 정의됩니다) ---
st.set_page_config(page_title="최종 단어 변환기", layout="wide")
st.title("📑 주차 연동 및 색상 지정 시스템")

# 사이드바 설정
with st.sidebar:
    st.header("⚙️ 고정값 설정")
    service_code = st.text_input("service_code", "SVC170")
    track_code = st.text_input("track_code", "RSV_TRK01")
    top_cors_id = st.text_input("top_cors_id", "1879")
    level_code = st.text_input("level_code", "TO_R_E_SP")
    component_code = st.text_input("component_code", "COM170")
    book_code = st.text_input("book_code", "SVC170")
    act_name = st.text_input("act_name", "Vocablist")
    show_debug = st.checkbox("디버그 모드", value=False)

# 중요: 여기서 uploaded_file 변수가 정의됩니다!
uploaded_file = st.file_uploader("워드(.docx) 또는 PDF(.pdf) 파일을 업로드하세요", type=["docx", "pdf"])

# 변수가 정의된 이후에 사용합니다.
if uploaded_file is not None:
    file_ext = uploaded_file.name.lower().rsplit(".", 1)[-1]

    if file_ext == "pdf":
        parsed_data, raw_texts = parse_pdf_content(uploaded_file)
    else:
        parsed_data, raw_texts = parse_word_content(uploaded_file)
    
    if parsed_data:
        df = pd.DataFrame(parsed_data)
        
        # 고정값 채우기
        df['service_code'] = service_code
        df['track_code'] = track_code
        df['top_cors_id'] = top_cors_id
        df['level_code'] = level_code
        df['component_code'] = component_code
        df['book_code'] = book_code
        df['act_code'] = "RSV_ACT002"
        df['act_name'] = act_name
        
        # 컬럼 순서 재배치
        final_cols = [
            "service_code", "track_code", "top_cors_id", "level_code", 
            "component_code", "book_code", "lesson_order_seq", "lesson_title", 
            "act_code", "act_name", "page_order_seq", "vocabulary", 
            "vocabulary_kor", "vocabulary_sound", "vocabulary_excep", "prompt"
        ]
        df_final = df.reindex(columns=final_cols).fillna("")

        st.success("✅ 분석 완료!")
        st.dataframe(df_final)

        # 엑셀 생성 및 색칠
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='원고작성')
            workbook = writer.book
            worksheet = writer.sheets['원고작성']
            
            # RGB(191, 191, 191) 색상 (HEX: BFBFBF)
            gray_fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')
            
            # 색상 적용 대상 컬럼
            target_cols = ["service_code", "track_code", "top_cors_id", "component_code", "book_code", "act_code"]
            
            for col_name in target_cols:
                if col_name in df_final.columns:
                    col_idx = df_final.columns.get_loc(col_name) + 1
                    for row in range(1, len(df_final) + 2):
                        cell = worksheet.cell(row=row, column=col_idx)
                        cell.fill = gray_fill

        st.download_button(
            label="📊 최종 결과물 다운로드",
            data=output.getvalue(),
            file_name=f"Standard_Vocab_{level_code}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    if show_debug:
        with st.expander("🔍 디버그 데이터"):
            for i, txt in enumerate(raw_texts):
                st.text(f"L{i+1}: {txt}")
