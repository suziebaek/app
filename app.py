import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
import pdfplumber
import io
import re
import os
import zipfile
import datetime
from copy import copy
from openpyxl import load_workbook

# 1. 파일명 규칙 정제 함수

# vocabulary 및 발음 파일명 모두에 적용: sb/sth 계열 줄임말 풀어쓰기
# (sb/sth처럼 함께 붙어있는 경우를 먼저 처리해야 단독 sb, sth 규칙에 의해 잘못 치환되지 않음)
def expand_sb_sth(text):
    t = text
    t = re.sub(r'\bsb\s*/\s*sth\b', 'something', t)   # sb/sth -> something
    t = re.sub(r'\bsb\b', 'somebody', t)              # sb -> somebody
    t = re.sub(r'\bsth\b', 'something', t)            # sth -> something
    return t

# 발음 파일명에서만 추가로 사용하지 않을 줄임말 -> 원래 표현 매핑
def expand_abbreviations(text):
    t = expand_sb_sth(text)
    t = re.sub(r'\bV-ing\b', 'ing', t)                # V-ing -> ing
    t = re.sub(r'\bto\s+V\b', 'to', t)                # to V -> to
    return t

def clean_filename(text):
    text = expand_abbreviations(text)
    # 공백/하이픈은 언더바로 치환 (예: voice-assisted -> voice_assisted)
    text = re.sub(r'[\s\-]+', '_', text)
    text = re.sub(r'[^a-zA-Z0-9_]', '', text)
    # 대소문자는 vocabulary 원문 그대로 유지 (예: protect A from B -> protect_A_from_B.mp3)
    return text + ".mp3"

# 챕터/회차 표기 감지: "Chapter 1", "CHAPTER 01", "CH01", "CH02", "1회", "04회" 등 지원
CHAPTER_LINE_PATTERN = re.compile(r'(?i)^(?:chapter|ch)\.?\s*0*(\d+)\b')
ROUND_LINE_PATTERN = re.compile(r'^제?\s*0*(\d+)\s*회\s*$')

def detect_chapter_number(line):
    stripped = line.strip()
    m = CHAPTER_LINE_PATTERN.match(stripped)
    if m:
        return int(m.group(1))
    m = ROUND_LINE_PATTERN.match(stripped)
    if m:
        return int(m.group(1))
    return None

def is_round_marker(line):
    return ROUND_LINE_PATTERN.match(line.strip()) is not None

# 번호 없이 이름만 있는 단원 제목 감지 (예: "Mastery", "Review", "Final Test")
# - 영문/숫자/공백 위주의 짧은 한 줄 (5단어 이하), 콜론·마침표·한글 없음
# - 제목만으로는 새 회차로 보지 않고, 바로 다음 단어 번호가 1로 다시 시작될 때만 새 회차로 인정
#   (예문 문장은 마침표로 끝나므로 제목으로 오인되지 않음)
SECTION_HEADING_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9&'\- ]{1,40}$")

def is_section_heading(line):
    s = line.strip()
    return bool(SECTION_HEADING_PATTERN.match(s)) and len(s.split()) <= 5

# 주차 배정 방식: 원본 회차 번호 n을 그대로 / 홀수 주(2n-1) / 짝수 주(2n)로 바꿉니다.
# 예) 홀수 주: CHAPTER 1 -> W01, CHAPTER 2 -> W03, ... / 짝수 주: CHAPTER 1 -> W02, CHAPTER 2 -> W04, ...
WEEK_MODES = {
    "자동 (파일명의 교재명으로 판단)": "auto",
    "그대로 (1, 2, 3, …)": "seq",
    "홀수 주 (1, 3, 5, …)": "odd",
    "짝수 주 (2, 4, 6, …)": "even",
}

# 교재명(파일명 맨 앞 등)으로 주차 배정 방식을 자동 결정합니다.
# 새 교재가 생기면 아래 목록에 교재명만 추가하면 됩니다.
ODD_WEEK_BOOKS = ["SR", "CSR", "CR", "SCR"]     # 홀수 주차 (1, 3, 5, …)
EVEN_WEEK_BOOKS = ["RP", "ARP", "URP"]          # 짝수 주차 (2, 4, 6, …)
AUTO_WEEK_MODE = "auto"
# 레슨 12개가 모두 있는 레벨은 교재와 상관없이 회차 번호 그대로(1~12주차) 배정합니다.
FULL_LESSON_LEVELS = ["S", "T"]

def detect_level(file_name, level_code=""):
    """레벨을 찾습니다. 1순위: 사이드바 level_code의 세 번째 값 (예: TO_R_T_AU -> T)
    2순위: 파일명 속 레벨 표기 (예: CSR_S1_Vocab.docx -> S, CR_H_Vocab.docx -> H)
    돌려주는 값: (레벨 또는 None, 어디서 찾았는지)"""
    parts = [p for p in str(level_code).strip().upper().split("_") if p]
    if len(parts) >= 3 and re.fullmatch(r'[A-Z]', parts[2]):
        return parts[2], "level_code"
    stem = os.path.splitext(os.path.basename(file_name))[0]
    for t in re.split(r'[_\s\-\.]+', stem):
        m = re.fullmatch(r'([A-Za-z])\d*', t)
        if m:
            return m.group(1).upper(), "파일명"
    return None, None

def detect_book_week_mode(file_name):
    """파일명을 _ / 공백 / - 기준으로 나눠 교재명을 찾습니다.
    예) CSR_E1_Vocab_List_초안.docx -> CSR -> 홀수 주
    돌려주는 값: (mode 또는 None, 찾은 교재명 또는 None)"""
    stem = os.path.splitext(os.path.basename(file_name))[0]
    tokens = [t.upper() for t in re.split(r'[_\s\-\.]+', stem) if t]
    odd = [t for t in tokens if t in ODD_WEEK_BOOKS]
    even = [t for t in tokens if t in EVEN_WEEK_BOOKS]
    if odd and not even:
        return "odd", odd[0]
    if even and not odd:
        return "even", even[0]
    return None, None

def resolve_week_mode(file_name, selected_mode, level_code=""):
    """'자동'이면 레벨(S/T는 그대로) -> 파일명의 교재명 순서로 결정하고, 못 찾으면 '그대로'를 씁니다.
    돌려주는 값: (실제 적용할 mode, 화면에 보여줄 설명)"""
    if selected_mode != AUTO_WEEK_MODE:
        return selected_mode, None
    level, level_src = detect_level(file_name, level_code)
    if level in FULL_LESSON_LEVELS:
        return "seq", (f"{level} 레벨({level_src} 기준)은 레슨 12개가 모두 있어 "
                       f"교재와 상관없이 회차 번호 그대로(1, 2, 3, …) 배정했습니다.")
    mode, book = detect_book_week_mode(file_name)
    if mode == "odd":
        return "odd", f"교재명 '{book}' → 홀수 주차(1, 3, 5, …)로 배정했습니다."
    if mode == "even":
        return "even", f"교재명 '{book}' → 짝수 주차(2, 4, 6, …)로 배정했습니다."
    return "seq", ("파일명에서 홀수/짝수 교재명을 찾지 못해 회차 번호 그대로(1, 2, 3, …) 배정했습니다. "
                   "필요하면 사이드바에서 직접 선택하세요.")

def remap_weeks(data, mode):
    if mode == "seq":
        return data
    new_data = []
    for d in data:
        n = int(d["lesson_order_seq"])
        week = 2 * n - 1 if mode == "odd" else 2 * n
        nd = dict(d)
        nd["lesson_order_seq"] = week
        nd["lesson_title"] = f"Week {week:02d}/"
        new_data.append(nd)
    return new_data

# 2. 워드 파싱 함수
def parse_word_content(docx_file):
    doc = Document(docx_file)
    data = []
    raw_texts = []
    
    current_chapter_num = 1
    current_week_title = "Week 01/"
    current_week_num = 1
    
    # 문서 내부 텍스트(마커/번호)뿐 아니라, 워드의 '구역 나누기(section break)'도
    # 회차 경계로 활용합니다. PDF -> 워드 변환 과정에서 "1회/2회..." 같은 텍스트 라벨은
    # 사라져도, 원래 페이지 구분에 쓰였던 구역 나누기는 그대로 남아있는 경우가 많습니다.
    # 각 항목은 ("text", 줄내용) 또는 ("section_break", None) 형태의 토큰입니다.
    tokens = []
    for p in doc.paragraphs:
        if p.text.strip():
            # Shift+Enter(줄바꿈)로 한 문단 안에 여러 줄(예: "word: 뜻" 다음 줄에 "Syn: ...")이
            # 들어있는 경우를 대비해 문단 내부의 줄바꿈 기준으로도 분리
            for sub in p.text.split("\n"):
                sub = sub.strip()
                if sub:
                    tokens.append(("text", sub))

        pPr = p._p.find(qn('w:pPr'))
        if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
            tokens.append(("section_break", None))

    paragraphs = [t[1] for t in tokens if t[0] == "text"]

    # 번호는 "1. word" (마침표)와 "01  word" (공백 2칸) 형식을 모두 지원
    entry_pattern = re.compile(r'^(?:(\d+)[\.\s]+)?([^:]+):\s*(.*)$')
    keyword_only_pattern = re.compile(r'(?i)^(Syn|Ant|Atn|유의어|반의어)$')
    # vocabulary_excep에는 유의어(Syn)만 채웁니다 (반의어/Ant는 제외)
    tail_keyword_pattern = re.compile(r'(?i)^(Syn|유의어)\s*:\s*(.*)$')

    # "1회", "04회"처럼 "~회" 표기가 문서 어딘가에 하나라도 있으면, 그 표기가
    # 누락된 곳도 있을 수 있다고 보고 번호가 정확히 1로 리셋되는 지점(예: ...20. -> 1.)을
    # 새 회차 시작으로 자동 보완합니다. "CH01" 같은 챕터 표기만 쓰는 문서(예: 회차 내부에
    # 번호가 우연히 반복/오탈자로 흐트러지는 경우)에는 적용하지 않아 기존 동작을 유지합니다.
    use_round_reset_fallback = any(is_round_marker(l) for l in paragraphs)
    previous_entry_num = None
    just_saw_explicit_marker = False
    pending_heading = False   # 직전 줄이 "Mastery" 같은 단원 제목이었는지

    for token_type, line in tokens:
        if token_type == "section_break":
            # 구역 나누기 = 새 회차(챕터) 시작 신호
            current_week_num += 1
            current_week_title = f"Week {current_week_num:02d}/"
            current_chapter_num = 1
            previous_entry_num = None
            just_saw_explicit_marker = True
            pending_heading = False
            continue

        raw_texts.append(line)
        
        # 챕터 감지 (Week 번호 연동)
        chapter_num = detect_chapter_number(line)
        if chapter_num is not None:
            current_week_num = chapter_num
            current_week_title = f"Week {current_week_num:02d}/"
            current_chapter_num = 1
            previous_entry_num = None
            just_saw_explicit_marker = True
            pending_heading = False
            continue

        # "Mastery" 같은 이름만 있는 단원 제목 (다음 단어 번호가 1이면 새 회차)
        if is_section_heading(line):
            pending_heading = True
            continue

        # 단어: 뜻 [탭/2칸 이상 공백 + Syn/Ant: 유의어] 추출
        match = entry_pattern.match(line)

        if match:
            entry_num_str = match.group(1)
            raw_word = expand_sb_sth(match.group(2).strip())

            # "Syn: xxx" 처럼 단어 없이 유의어만 있는 줄은 아래 유의어 처리 블록으로 넘김
            if not keyword_only_pattern.match(raw_word):
                # 번호가 이전보다 작아지며 리셋되면 새 회차(챕터) 시작으로 간주
                # (단, 직전에 명시적 챕터/회차 마커를 이미 처리했다면 중복 증가하지 않음)
                if entry_num_str is not None:
                    entry_num = int(entry_num_str)
                    # 번호가 정확히 1로 리셋되는 경우만 새 회차 시작으로 인정
                    # (오탈자로 인한 불규칙한 감소는 무시하여 오탐을 방지)
                    restarted = (previous_entry_num is not None and entry_num == 1
                                 and previous_entry_num != 1 and not just_saw_explicit_marker)
                    if restarted and (use_round_reset_fallback or pending_heading):
                        current_week_num += 1
                        current_week_title = f"Week {current_week_num:02d}/"
                        current_chapter_num = 1
                    previous_entry_num = entry_num

                just_saw_explicit_marker = False
                pending_heading = False

                rest = match.group(3)

                # 뜻과 유의어/반의어를 탭 또는 2칸 이상 공백으로 분리
                parts = re.split(r'\t|\s{2,}', rest, maxsplit=1)
                meaning = parts[0].strip()

                extra_info = ""
                if len(parts) > 1:
                    tail = parts[1].strip()
                    tail_match = tail_keyword_pattern.match(tail)
                    if tail_match:
                        extra_info = expand_sb_sth(tail_match.group(2).strip())

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
        
        # 유의어 처리 (Syn이 별도의 줄로 존재하는 예전 형식 지원, Ant/반의어는 제외)
        if any(keyword in line for keyword in ["Syn", "유의어"]):
            extra_info = line.split(":", 1)[-1].strip()
            if data:
                first_extra = expand_sb_sth(re.split(r'\s{2,}', extra_info)[0].strip())
                data[-1]["vocabulary_excep"] = first_extra

    return data, raw_texts

# 2-0. 워드 파싱 함수 (용어집/Glossary 형식)
# "word / 뜻 / (부가 설명) / syn: ... / ant: ... / 예문 / 활용 표현" 처럼
# 각 항목이 한 줄에 한 요소씩, 콜론으로 단어-뜻이 묶이지 않고 여러 줄에 나뉘어
# 있는 문서 형식을 지원합니다. 단어 줄은 "그 줄 자체엔 한글이 없고, 바로 다음
# 줄에는 한글이 있다"는 특징으로 판별합니다.
def _has_korean(text):
    return bool(re.search(r'[가-힣]', text))

def _is_glossary_word_start(lines, idx):
    if idx >= len(lines):
        return False
    line = lines[idx]
    if re.match(r'(?i)^(syn|ant|atn|유의어|반의어)\s*:', line):
        return False
    if _has_korean(line):
        return False
    if detect_chapter_number(line) is not None:
        return False
    if idx + 1 >= len(lines):
        return False
    next_line = lines[idx + 1]
    # 활용/연어(collocation) 목록 줄은 "/"로 구분되며, 그 안에 "+ 기간"처럼
    # 한글 플레이스홀더가 섞여 있어도 뜻(meaning) 줄로 오인하지 않도록 제외합니다.
    if "/" in next_line:
        return False
    return _has_korean(next_line)

def parse_word_content_glossary_style(docx_file):
    doc = Document(docx_file)

    # 워드 안에서 줄바꿈(Shift+Enter)으로 나뉜 줄까지 모두 개별 라인으로 펼침
    raw_lines = []
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        for sub in p.text.split("\n"):
            sub = sub.strip()
            if sub:
                raw_lines.append(sub)

    data = []
    raw_texts = list(raw_lines)

    current_chapter_num = 1
    current_week_title = "Week 01/"
    current_week_num = 1

    # vocabulary_excep에는 유의어(Syn)만 채웁니다 (반의어/Ant는 제외)
    tail_keyword_pattern = re.compile(r'(?i)^(syn|유의어)\s*:\s*(.*)$')

    i = 0
    n = len(raw_lines)
    while i < n:
        line = raw_lines[i]

        # 챕터 감지 (Week 번호 연동)
        chapter_num = detect_chapter_number(line)
        if chapter_num is not None:
            current_week_num = chapter_num
            current_week_title = f"Week {current_week_num:02d}/"
            current_chapter_num = 1
            i += 1
            continue

        if _is_glossary_word_start(raw_lines, i):
            raw_word = expand_sb_sth(line)
            meaning = raw_lines[i + 1].strip()

            # 단어 다음, 다음 단어(혹은 챕터)가 나오기 전까지의 구간에서
            # 첫 번째로 등장하는 syn/ant 값을 유의어/반의어로 채택
            extra_info = ""
            j = i + 2
            while j < n and not _is_glossary_word_start(raw_lines, j) and detect_chapter_number(raw_lines[j]) is None:
                tail_match = tail_keyword_pattern.match(raw_lines[j])
                if tail_match and not extra_info:
                    extra_info = expand_sb_sth(tail_match.group(2).strip())
                j += 1

            if _has_korean(meaning):
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

            i = j
            continue

        i += 1

    return data, raw_texts

# 2-1. PDF 파싱 함수
def parse_pdf_content(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = [ln.strip() for ln in full_text.split("\n") if ln.strip()]

    data = []
    raw_texts = []

    current_chapter_num = 1
    current_week_title = "Week 01/"
    current_week_num = 1
    previous_entry_num = None
    pending_heading = False

    # PDF 텍스트 추출 시 "01 word: 뜻   Syn: xxx" 처럼 번호가 마침표 없이
    # 붙는 형태와, 뜻과 Syn/Ant 사이가 한 칸 공백으로만 구분되는 경우가 많아
    # 워드용 정규식과는 별도의 패턴을 사용합니다.
    entry_pattern = re.compile(r'^\d{1,2}\s+([^:]+):\s*(.*)$')
    # Syn/Ant 둘 다 감지해서 뜻(meaning)에서는 잘라내되, vocabulary_excep에는 Syn만 채웁니다
    tail_start_pattern = re.compile(r'\b(Syn|Ant)\s*:')
    syn_value_pattern = re.compile(r'\bSyn\s*:\s*(.+)$')

    for line in lines:
        raw_texts.append(line)

        # 챕터 감지 (Week 번호 연동, 페이지 하단에 챕터 번호가 반복 출력되는
        # 문서도 있어 실제로 챕터가 바뀔 때만 카운터를 리셋합니다)
        chapter_num = detect_chapter_number(line)
        if chapter_num is not None:
            if chapter_num != current_week_num:
                current_week_num = chapter_num
                current_week_title = f"Week {current_week_num:02d}/"
                current_chapter_num = 1
                previous_entry_num = None
            pending_heading = False
            continue

        # "Mastery" 같은 이름만 있는 단원 제목 (다음 단어 번호가 1이면 새 회차)
        if is_section_heading(line):
            pending_heading = True
            continue

        # 번호 + 단어: 뜻 [Syn/Ant: 유의어] 추출
        match = entry_pattern.match(line)
        if match:
            entry_num = int(re.match(r'^(\d{1,2})', line).group(1))
            if pending_heading and previous_entry_num is not None and entry_num == 1 and previous_entry_num != 1:
                current_week_num += 1
                current_week_title = f"Week {current_week_num:02d}/"
                current_chapter_num = 1
            previous_entry_num = entry_num
            pending_heading = False

            raw_word = expand_sb_sth(match.group(1).strip())
            rest = match.group(2).strip()

            tail_match = tail_start_pattern.search(rest)
            if tail_match:
                meaning = rest[:tail_match.start()].strip()
                tail = rest[tail_match.start():]
                syn_match = syn_value_pattern.search(tail)
                extra_info = expand_sb_sth(syn_match.group(1).strip()) if syn_match else ""
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

# 2-2. PDF 파싱 함수 (품사 표기형: "01 word 품사. 뜻" 형식, 콜론 없음, 단원 코드 접미)
# 예: "01 spicy adj. 매운 LZ" / "02 taste n. 맛 v. ~한 맛이 나다 LZ"
def parse_pdf_content_pos_style(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = [ln.strip() for ln in full_text.split("\n") if ln.strip()]

    data = []
    raw_texts = []

    current_chapter_num = 1
    current_week_title = "Week 01/"
    current_week_num = 1
    previous_entry_num = None
    pending_heading = False

    # 번호, 단어(구), 품사(n./v./adj./adv./prep./conj., 필요시 phr. 포함), 뜻,
    # 그리고 마지막에 붙는 단원 코드(LZ/U1/U2/TS)까지 한 줄에 모두 들어있는 형식
    entry_pattern = re.compile(
        r"^(\d{1,2})\s+([A-Za-z][A-Za-z'\-\s]*?)\s+"
        r"((?:n|v|adj|adv|prep|conj)\.\s*(?:phr\.\s*)?)(.*?)(?:\s+(?:LZ|U1|U2|TS))?$"
    )
    # 뜻 안에 품사가 여러 번(예: "맛 v. ~한 맛이 나다") 섞여 있으면 쉼표로 구분
    pos_inline_pattern = re.compile(r"\s*\b(?:n|v|adj|adv|prep|conj)\.\s*(?:phr\.\s*)?")

    for line in lines:
        raw_texts.append(line)

        chapter_num = detect_chapter_number(line)
        if chapter_num is not None:
            if chapter_num != current_week_num:
                current_week_num = chapter_num
                current_week_title = f"Week {current_week_num:02d}/"
                current_chapter_num = 1
                previous_entry_num = None
            pending_heading = False
            continue

        # "Mastery" 같은 이름만 있는 단원 제목 (다음 단어 번호가 1이면 새 회차)
        if is_section_heading(line):
            pending_heading = True
            continue

        match = entry_pattern.match(line)
        if match:
            entry_num = int(match.group(1))
            if pending_heading and previous_entry_num is not None and entry_num == 1 and previous_entry_num != 1:
                current_week_num += 1
                current_week_title = f"Week {current_week_num:02d}/"
                current_chapter_num = 1
            previous_entry_num = entry_num
            pending_heading = False

            raw_word = expand_sb_sth(match.group(2).strip())
            rest = match.group(4).strip()
            meaning = pos_inline_pattern.sub(", ", rest).strip(", ").strip()

            if re.search(r'[가-힣]', meaning):
                sound_file = clean_filename(raw_word)

                data.append({
                    "lesson_title": current_week_title,
                    "lesson_order_seq": current_week_num,
                    "page_order_seq": current_chapter_num,
                    "vocabulary": raw_word,
                    "vocabulary_kor": meaning,
                    "vocabulary_sound": sound_file,
                    "vocabulary_excep": "",
                    "prompt": ""
                })
                current_chapter_num += 1

    return data, raw_texts

# 4. 엑셀 생성 헬퍼: 제공된 서식 템플릿(TO_Recap_S_원고템플릿.xlsx)을 그대로 사용해서
#    "원고작성" 시트에 데이터를 채우고, "문서개정" 시트는 그대로 포함시킵니다.
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "template.xlsx")

# 템플릿의 회차(Prompt) 행 규칙: act_code=RSV_ACT001 / act_name=Prompt / page_order_seq=1 / prompt=prompt.FOLDER
PROMPT_ACT_CODE = "RSV_ACT001"
PROMPT_ACT_NAME = "Prompt"
PROMPT_VALUE = "prompt.FOLDER"

FINAL_COLS = [
    "service_code", "track_code", "top_cors_id", "level_code",
    "component_code", "book_code", "lesson_order_seq", "lesson_title",
    "act_code", "act_name", "page_order_seq", "vocabulary",
    "vocabulary_kor", "vocabulary_sound", "vocabulary_excep", "prompt"
]

def build_excel_bytes(df_subset):
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["원고작성"]

    # 템플릿 예시 행(2행=Prompt 행, 3행=단어 행)에서 컬럼별 스타일을 미리 복사해둔 뒤
    # 예시 데이터를 지우고, 실제 데이터를 같은 스타일로 다시 채웁니다.
    prompt_row_styles = {c: copy(ws.cell(row=2, column=c)._style) for c in range(1, 17)}
    vocab_row_styles = {c: copy(ws.cell(row=3, column=c)._style) for c in range(1, 17)}

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    def write_row(row_idx, values, styles):
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value if value != "" else None)
            cell._style = styles[col_idx]

    row_idx = 2
    # lesson_order_seq(주차) 등장 순서를 유지하면서 그룹핑
    week_order = list(dict.fromkeys(df_subset["lesson_order_seq"].tolist()))

    for week_num in week_order:
        week_df = df_subset[df_subset["lesson_order_seq"] == week_num]
        first = week_df.iloc[0]

        # 회차(Prompt) 행 - 고정값들은 해당 주차 단어 행과 동일한 값을 그대로 사용
        write_row(row_idx, [
            first["service_code"], first["track_code"], first["top_cors_id"],
            first["level_code"], first["component_code"], first["book_code"],
            week_num, first["lesson_title"], PROMPT_ACT_CODE, PROMPT_ACT_NAME, 1,
            "", "", "", "", PROMPT_VALUE
        ], prompt_row_styles)
        row_idx += 1

        # 단어 행들
        for _, r in week_df.iterrows():
            write_row(row_idx, [
                r["service_code"], r["track_code"], r["top_cors_id"], r["level_code"],
                r["component_code"], r["book_code"], r["lesson_order_seq"], r["lesson_title"],
                r["act_code"], r["act_name"], r["page_order_seq"], r["vocabulary"],
                r["vocabulary_kor"], r["vocabulary_sound"], r["vocabulary_excep"], r["prompt"]
            ], vocab_row_styles)
            row_idx += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================================================
# 5. 원본 파일 파싱 (탭1, 탭2 공용)
# ==========================================================================
import asyncio
import time
import shutil
import tempfile
import subprocess
import collections

def parse_source_file(uploaded, week_mode="seq"):
    """워드/PDF를 기존 파서로 분석하고, 주차 배정 방식(그대로/홀수/짝수)을 적용합니다. (탭1, 탭2 공용)"""
    file_ext = uploaded.name.lower().rsplit(".", 1)[-1]
    uploaded.seek(0)
    if file_ext == "pdf":
        parsed_data, raw_texts = parse_pdf_content(uploaded)
        if not parsed_data:
            uploaded.seek(0)
            parsed_data, raw_texts = parse_pdf_content_pos_style(uploaded)
    else:
        parsed_data, raw_texts = parse_word_content(uploaded)
        if not parsed_data:
            uploaded.seek(0)
            parsed_data, raw_texts = parse_word_content_glossary_style(uploaded)
    return remap_weeks(parsed_data, week_mode), raw_texts


# ==========================================================================
# 6. 자동 검수 (탭2)
# ==========================================================================
HANGUL = re.compile(r'[가-힣]')
ABBR_PATTERN = re.compile(r'\b(sb|sth|V-ing|to V)\b')
EXCEL_NAME_PATTERN = re.compile(r'^(.+)_W(\d{2})_(\d{6})\.xlsx$')
FIXED_COLS = ["service_code", "track_code", "top_cors_id", "level_code", "component_code", "book_code"]
VOCAB_ACT_CODE = "RSV_ACT002"

def _s(v):
    """셀 값을 문자열로 (None -> '', 1.0 -> '1')"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)

def normalize_spaces(text):
    """앞뒤 공백 제거 + 연속 공백을 한 칸으로"""
    return re.sub(r'\s{2,}', ' ', text.strip())

def load_revision_styles():
    """템플릿 '문서개정' 시트의 수정 차수별 예시 셀(F20~F24) 서식을 그대로 가져옵니다."""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["문서개정"]
    styles = {}
    for i, row in enumerate(range(20, 25), start=1):
        cell = ws.cell(row=row, column=6)
        styles[f"{i}차 수정"] = (copy(cell.fill), copy(cell.font.color))
    return styles

def check_workbook(file_name, file_bytes, source_rows=None, expected_fixed=None):
    """엑셀 한 개를 검수해서 문제 목록을 돌려줍니다.
    각 문제: 파일, 행, 열, 수준(오류/경고), 내용, 현재값, 제안값, 자동수정(True/False)"""
    issues = []

    def add(row, col, level, msg, current="", suggest="", fixable=False):
        issues.append({
            "파일": file_name, "행": row, "열": col, "수준": level, "내용": msg,
            "현재값": current, "제안값": suggest, "자동수정": fixable,
        })

    try:
        wb = load_workbook(io.BytesIO(file_bytes))
    except Exception as e:
        add("-", "-", "오류", f"엑셀 파일을 열 수 없습니다: {e}")
        return issues

    # (1) 시트 구성
    for sheet in ["원고작성", "문서개정"]:
        if sheet not in wb.sheetnames:
            add("-", "-", "오류", f"'{sheet}' 시트가 없습니다. (템플릿 정보 변경 금지)")
    if "원고작성" not in wb.sheetnames:
        return issues
    ws = wb["원고작성"]

    # (2) 헤더
    header = [_s(c.value).strip() for c in ws[1]][:len(FINAL_COLS)]
    if header != FINAL_COLS:
        add(1, "-", "오류", "헤더(1행)가 템플릿과 다릅니다.", " | ".join(header), " | ".join(FINAL_COLS))
        return issues
    col_idx = {name: i + 1 for i, name in enumerate(FINAL_COLS)}

    rows = []
    for r in range(2, ws.max_row + 1):
        vals = {name: ws.cell(row=r, column=col_idx[name]).value for name in FINAL_COLS}
        if all(_s(v).strip() == "" for v in vals.values()):
            continue
        rows.append((r, vals))

    if not rows:
        add("-", "-", "오류", "데이터 행이 없습니다.")
        return issues

    # (3) 모든 문자열 셀: 불필요한 띄어쓰기
    for r, vals in rows:
        for name in FINAL_COLS:
            v = vals[name]
            if isinstance(v, str) and v != normalize_spaces(v):
                add(r, name, "오류", "불필요한 띄어쓰기(앞뒤 공백 또는 연속 공백)",
                    repr(v), normalize_spaces(v), True)

    # (4) 행 단위 검사
    sound_seen = {}
    for r, vals in rows:
        act = _s(vals["act_code"]).strip()
        vocab = normalize_spaces(_s(vals["vocabulary"]))
        kor = normalize_spaces(_s(vals["vocabulary_kor"]))
        sound = normalize_spaces(_s(vals["vocabulary_sound"]))
        excep = normalize_spaces(_s(vals["vocabulary_excep"]))
        prompt = normalize_spaces(_s(vals["prompt"]))

        if act == PROMPT_ACT_CODE:
            if prompt != PROMPT_VALUE:
                add(r, "prompt", "오류", f"Prompt 행의 prompt 값은 '{PROMPT_VALUE}'이어야 합니다.", prompt, PROMPT_VALUE)
            if _s(vals["page_order_seq"]).strip() != "1":
                add(r, "page_order_seq", "오류", "Prompt 행의 page_order_seq는 1이어야 합니다.", _s(vals["page_order_seq"]), "1")
            if vocab or kor or sound or excep:
                add(r, "vocabulary", "오류", "Prompt 행에는 단어 정보가 비어 있어야 합니다.")
            continue

        if act != VOCAB_ACT_CODE:
            add(r, "act_code", "오류", "알 수 없는 act_code입니다.", act, f"{PROMPT_ACT_CODE} 또는 {VOCAB_ACT_CODE}")
            continue

        if not vocab:
            add(r, "vocabulary", "오류", "단어가 비어 있습니다.")
        elif HANGUL.search(vocab):
            add(r, "vocabulary", "오류", "단어 칸에 한글이 들어 있습니다.", vocab)
        elif ABBR_PATTERN.search(vocab):
            add(r, "vocabulary", "경고", "줄임말(sb/sth/V-ing/to V)이 남아 있습니다. 표기 규칙을 확인하세요.", vocab)

        if not kor:
            add(r, "vocabulary_kor", "오류", "뜻이 비어 있습니다.")
        elif not HANGUL.search(kor):
            add(r, "vocabulary_kor", "경고", "뜻에 한글이 없습니다.", kor)

        if vocab and not HANGUL.search(vocab):
            expected = clean_filename(vocab)
            if sound != expected:
                add(r, "vocabulary_sound", "오류", "mp3 파일명이 단어에서 만든 규칙과 다릅니다.",
                    sound, expected, True)
        if sound:
            if sound in sound_seen:
                add(r, "vocabulary_sound", "오류", f"파일명이 {sound_seen[sound]}행과 중복됩니다.", sound)
            else:
                sound_seen[sound] = r

        if excep:
            if HANGUL.search(excep):
                add(r, "vocabulary_excep", "경고", "유의어 칸에 한글이 들어 있습니다.", excep)
            if re.match(r'(?i)^(syn|ant)\s*:', excep):
                add(r, "vocabulary_excep", "오류", "유의어 칸에 'Syn:/Ant:' 표기가 남아 있습니다.", excep)

        if prompt:
            add(r, "prompt", "경고", "단어 행의 prompt 칸은 비어 있어야 합니다.", prompt)

    # (5) 회차(lesson)별 구성 검사
    lessons = collections.OrderedDict()
    for r, vals in rows:
        lessons.setdefault(_s(vals["lesson_order_seq"]).strip(), []).append((r, vals))

    for seq, lrows in lessons.items():
        first_row = lrows[0][0]
        try:
            seq_num = int(seq)
            expected_title = f"Week {seq_num:02d}/"
        except ValueError:
            add(first_row, "lesson_order_seq", "오류", "lesson_order_seq가 숫자가 아닙니다.", seq)
            continue

        for r, vals in lrows:
            title = normalize_spaces(_s(vals["lesson_title"]))
            if title != expected_title:
                add(r, "lesson_title", "오류", "lesson_title이 회차 번호와 맞지 않습니다.", title, expected_title)

        prompt_rows = [r for r, v in lrows if _s(v["act_code"]).strip() == PROMPT_ACT_CODE]
        if len(prompt_rows) != 1:
            add(first_row, "act_code", "오류", f"Week {seq_num:02d}의 Prompt 행이 {len(prompt_rows)}개입니다. (1개여야 함)")
        elif prompt_rows[0] != first_row:
            add(prompt_rows[0], "act_code", "경고", "Prompt 행이 회차의 첫 행이 아닙니다.")

        vocab_rows = [(r, v) for r, v in lrows if _s(v["act_code"]).strip() == VOCAB_ACT_CODE]
        for expected_n, (r, v) in enumerate(vocab_rows, start=1):
            actual = _s(v["page_order_seq"]).strip()
            if actual != str(expected_n):
                add(r, "page_order_seq", "오류", "번호가 1부터 순서대로 이어지지 않습니다.", actual, str(expected_n))

    # (6) 고정값 일관성
    for name in FIXED_COLS + ["act_name"]:
        if name == "act_name":
            values = {_s(v[name]).strip() for _, v in rows if _s(v["act_code"]).strip() == VOCAB_ACT_CODE}
        else:
            values = {_s(v[name]).strip() for _, v in rows}
        if len(values) > 1:
            add("-", name, "오류", "같은 파일 안에서 값이 서로 다릅니다.", ", ".join(sorted(values)))
        if expected_fixed and name in expected_fixed:
            exp = str(expected_fixed[name]).strip()
            wrong = sorted(v for v in values if v != exp)
            if wrong:
                add("-", name, "오류", "사이드바 고정값과 다릅니다.", ", ".join(wrong), exp)

    # (7) 엑셀 파일명 규칙: 접두어_W##_yymmdd.xlsx
    m = EXCEL_NAME_PATTERN.match(file_name)
    if not m:
        add("-", "파일명", "경고", "파일명이 '접두어_W##_yymmdd.xlsx' 형식이 아닙니다.", file_name)
    else:
        try:
            datetime.datetime.strptime(m.group(3), "%y%m%d")
        except ValueError:
            add("-", "파일명", "오류", "파일명의 날짜(yymmdd)가 올바르지 않습니다.", m.group(3))
        lesson_keys = [k for k in lessons.keys()]
        if len(lesson_keys) == 1 and lesson_keys[0].isdigit() and int(lesson_keys[0]) != int(m.group(2)):
            add("-", "파일명", "오류", "파일명의 주차(W##)와 lesson_order_seq가 다릅니다.",
                f"W{m.group(2)}", f"W{int(lesson_keys[0]):02d}")

    # (8) 원본(워드/PDF)과 대조
    if source_rows is not None:
        src_by_key = {}
        for d in source_rows:
            src_by_key[(str(d["lesson_order_seq"]), str(d["page_order_seq"]))] = d
        for seq, lrows in lessons.items():
            vocab_rows = [(r, v) for r, v in lrows if _s(v["act_code"]).strip() == VOCAB_ACT_CODE]
            src_count = sum(1 for k in src_by_key if k[0] == seq)
            if src_count == 0:
                add("-", "-", "경고", f"원본에서 Week {seq} 회차를 찾지 못해 대조를 건너뜁니다.")
                continue
            if src_count != len(vocab_rows):
                add("-", "-", "오류", f"Week {seq} 단어 수가 원본({src_count}개)과 다릅니다.",
                    f"{len(vocab_rows)}개", f"{src_count}개")
            for r, v in vocab_rows:
                src = src_by_key.get((seq, _s(v["page_order_seq"]).strip()))
                if not src:
                    continue
                for col in ["vocabulary", "vocabulary_kor", "vocabulary_excep"]:
                    excel_val = normalize_spaces(_s(v[col]))
                    src_val = normalize_spaces(str(src.get(col, "")))
                    if excel_val != src_val:
                        add(r, col, "경고", "원본과 내용이 다릅니다. (의도한 수정인지 확인)", excel_val, src_val)

    return issues

def apply_auto_fixes(file_bytes, issues, revision_style):
    """자동수정 가능한 문제를 고치고, 고친 셀에 수정 차수 서식(바탕색+붉은 글씨)을 입힙니다."""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb["원고작성"]
    col_idx = {name: i + 1 for i, name in enumerate(FINAL_COLS)}
    fill, font_color = revision_style

    # 같은 셀에 여러 수정(공백 + 파일명)이 겹치면 마지막(파일명 규칙) 값이 적용됩니다.
    fixes = {}
    for it in issues:
        if it["자동수정"] and isinstance(it["행"], int):
            fixes[(it["행"], it["열"])] = it["제안값"]

    for (r, col), value in fixes.items():
        cell = ws.cell(row=r, column=col_idx[col])
        cell.value = value if value != "" else None
        cell.fill = copy(fill)
        new_font = copy(cell.font)
        new_font.color = copy(font_color)
        cell.font = new_font

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), len(fixes)


# ==========================================================================
# 7. MP3 만들기 (탭3) - 기존 Colab 코드의 규칙을 그대로 옮김
# ==========================================================================
DEFAULT_TTS = {
    "VOICE": "en-US-AriaNeural",
    "RATE": "-10%",
    "PITCH": "+0Hz",
    "REPEAT_COUNT": 1,
    "CONCURRENCY": 2,
    "SLEEP_MS": 500,
    "MAX_WORDS": 120,
    "BITRATE": "32k",
}
LOG_FILE = "build_log.csv"
CLEAN_EXCEL = "정리된_단어목록.xlsx"

def make_speech_text(raw_word):
    """기존 Colab 코드의 발음 문장 규칙 + 하이픈을 띄어 읽도록 한 줄 추가"""
    base_text = re.sub(r'\.mp3$', '', raw_word, flags=re.IGNORECASE).strip()
    base_text = base_text.replace('-', ' ')   # 추가: mass-produce -> mass produce
    clean_text_only = re.sub(r'[^a-zA-Z0-9\s_]', '', base_text)
    speech_text = clean_text_only.replace('_', ' ')
    speech_text = re.sub(r'\s+', ' ', speech_text).strip()
    speech_text = re.sub(r'\bA\b', 'A.', speech_text)
    return speech_text

def read_vocab_items(file_bytes):
    """엑셀의 단어 행을 읽어옵니다. 파일명은 vocabulary_sound 열을 그대로 사용합니다."""
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="원고작성", dtype=str).fillna("")
    df = df[df["act_code"].str.strip() == VOCAB_ACT_CODE]
    items = []
    for idx, row in df.iterrows():
        raw_word = normalize_spaces(row["vocabulary"])
        filename = normalize_spaces(row["vocabulary_sound"])
        if not raw_word or not filename:
            continue
        items.append({
            "idx": len(items),
            "orig_word": raw_word,
            "speech_text": make_speech_text(raw_word),
            "final_filename": filename,
        })
    return items

def get_ffmpeg():
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        return shutil.which("ffmpeg")

def reencode(ffmpeg, src, dst, bitrate):
    subprocess.run([ffmpeg, "-y", "-loglevel", "error", "-i", src, "-b:a", bitrate, dst], check=True)

async def convert_word(item, semaphore, log_data, cfg, out_dir, ffmpeg, on_done):
    import edge_tts
    async with semaphore:
        speech_text = item['speech_text']
        filename = item['final_filename']
        filepath = os.path.join(out_dir, filename)
        tmp_path = filepath + ".tmp.mp3"

        single_text = speech_text if speech_text.endswith('.') else f"{speech_text}."
        text_to_speak = " ".join([single_text] * cfg["REPEAT_COUNT"])

        start_time = time.time()
        last_error = ""
        success = False

        for attempt in range(3):
            try:
                communicate = edge_tts.Communicate(text=text_to_speak, voice=cfg["VOICE"],
                                                   rate=cfg["RATE"], pitch=cfg["PITCH"])
                await communicate.save(tmp_path)
                if os.path.getsize(tmp_path) == 0:
                    raise RuntimeError("빈 오디오 수신 (0 bytes)")
                await asyncio.to_thread(reencode, ffmpeg, tmp_path, filepath, cfg["BITRATE"])
                success = True
                break
            except Exception as e:
                last_error = f"{type(e).__name__}: {e}"
                await asyncio.sleep(2 * (attempt + 1))
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)

        if not success and os.path.exists(filepath):
            os.remove(filepath)

        status = "Success" if success else f"Error: {last_error}"
        log_data.append([item['idx'] + 1, item['orig_word'], speech_text, filename, status,
                         f"{time.time() - start_time:.2f}s"])
        on_done(item, success)
        await asyncio.sleep(cfg["SLEEP_MS"] / 1000.0)

async def run_tts(items, cfg, out_dir, ffmpeg, on_done):
    semaphore = asyncio.Semaphore(cfg["CONCURRENCY"])
    log_data = []
    await asyncio.gather(*[convert_word(it, semaphore, log_data, cfg, out_dir, ffmpeg, on_done) for it in items])
    return log_data

def check_mp3_files(items, out_dir, bitrate, min_sec, max_sec):
    """만들어진 mp3를 엑셀 기준으로 검수합니다."""
    from mutagen.mp3 import MP3
    expected_kbps = int(re.sub(r'\D', '', bitrate) or 0)
    rows = []
    for it in items:
        fn = it["final_filename"]
        path = os.path.join(out_dir, fn)
        problems = []
        size = os.path.getsize(path) if os.path.exists(path) else None
        kbps = length = sr = None
        if size is None:
            problems.append("파일 없음")
        elif size == 0:
            problems.append("0KB")
        else:
            try:
                info = MP3(path).info
                kbps = round(info.bitrate / 1000)
                length = round(info.length, 2)
                sr = info.sample_rate
                if expected_kbps and kbps != expected_kbps:
                    problems.append(f"비트레이트 {kbps}k (기준 {expected_kbps}k)")
                if length < min_sec or length > max_sec:
                    problems.append(f"길이 {length}초 (기준 {min_sec}~{max_sec}초)")
            except Exception as e:
                problems.append(f"mp3 읽기 실패: {e}")
        if not re.fullmatch(r'[A-Za-z0-9_]+\.mp3', fn):
            problems.append("파일명에 허용되지 않는 문자")
        rows.append({
            "파일명": fn, "단어": it["orig_word"], "크기(KB)": round(size / 1024, 1) if size else 0,
            "비트레이트(kbps)": kbps, "길이(초)": length, "샘플레이트": sr,
            "결과": "통과" if not problems else " / ".join(problems),
        })
    created = set(os.listdir(out_dir))
    extras = sorted(created - {it["final_filename"] for it in items})
    return pd.DataFrame(rows), extras

def build_mp3_zip(items, log_data, out_dir):
    df_result = pd.DataFrame(log_data, columns=["번호", "원본_단어", "정제된_발음문장", "생성된_파일명", "상태", "소요시간"])
    df_result = df_result.sort_values(by="번호").reset_index(drop=True)
    log_bytes = df_result.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    clean_buf = io.BytesIO()
    df_result[["생성된_파일명"]].to_excel(clean_buf, index=False, engine='openpyxl')

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w') as zf:
        for it in items:
            path = os.path.join(out_dir, it["final_filename"])
            if os.path.exists(path) and os.path.getsize(path) > 0:
                zf.write(path, arcname=os.path.join("ContentsFiles", it["final_filename"]))
        zf.writestr(LOG_FILE, log_bytes)
        zf.writestr(CLEAN_EXCEL, clean_buf.getvalue())
    return zip_buf.getvalue(), df_result


# ==========================================================================
# 8. Streamlit UI
# ==========================================================================
st.set_page_config(page_title="단어 원고 · 검수 · MP3 통합 도구", layout="wide")
st.title("📑 단어 원고 · 검수 · MP3 통합 도구")

# 사이드바 설정 (기존과 동일)
with st.sidebar:
    st.header("⚙️ 고정값 설정")
    service_code = st.text_input("service_code", "SVC170")
    track_code = st.text_input("track_code", "RSV_TRK01")
    top_cors_id = st.text_input("top_cors_id", "1879")
    level_code = st.text_input("level_code", "TO_R_E_SP")
    component_code = st.text_input("component_code", "COM170")
    book_code = st.text_input("book_code", "SVC170")
    act_name = st.text_input("act_name", "Vocab list")
    show_debug = st.checkbox("디버그 모드", value=False)

    st.header("📅 주차 배정")
    week_mode_label = st.radio(
        "회차를 몇 주차로 넣을까요?", list(WEEK_MODES.keys()),
        help="자동: " + "/".join(FULL_LESSON_LEVELS) + " 레벨은 그대로, 그 외 홀수 주 교재(" + ", ".join(ODD_WEEK_BOOKS)
             + "), 짝수 주 교재(" + ", ".join(EVEN_WEEK_BOOKS) + ")\n\n"
             "홀수 주: CHAPTER 1 → W01, CHAPTER 2 → W03 …  /  짝수 주: CHAPTER 1 → W02, CHAPTER 2 → W04 …",
    )
    week_mode = WEEK_MODES[week_mode_label]

    st.header("📁 파일명 설정")
    st.caption("예: TOT_AU_Recap_S_W04_260731.xlsx")
    file_name_prefix = st.text_input("파일명 접두어 (주차/날짜 앞부분)", "TOT_AU_Recap_S")
    file_date = st.date_input("파일명 날짜", value=datetime.date.today())
    date_str = file_date.strftime("%y%m%d")

tab1, tab2, tab3 = st.tabs(["① 엑셀 만들기", "② 엑셀 검수", "③ MP3 만들기"])

# --------------------------------------------------------------------------
# 탭 1: 엑셀 만들기 (기존 기능 그대로)
# --------------------------------------------------------------------------
with tab1:
    uploaded_file = st.file_uploader("워드(.docx) 또는 PDF(.pdf) 파일을 업로드하세요", type=["docx", "pdf"], key="src_upload")

    if uploaded_file is not None:
        applied_mode, mode_note = resolve_week_mode(uploaded_file.name, week_mode, level_code)
        if mode_note:
            st.info("📅 " + mode_note)
        parsed_data, raw_texts = parse_source_file(uploaded_file, applied_mode)

        if parsed_data:
            df = pd.DataFrame(parsed_data)

            df['service_code'] = service_code
            df['track_code'] = track_code
            try:
                df['top_cors_id'] = int(top_cors_id)
            except (ValueError, TypeError):
                df['top_cors_id'] = top_cors_id
            df['level_code'] = level_code
            df['component_code'] = component_code
            df['book_code'] = book_code
            df['act_code'] = "RSV_ACT002"
            df['act_name'] = act_name

            df_final = df.reindex(columns=FINAL_COLS).fillna("")

            st.success("✅ 분석 완료!")
            st.dataframe(df_final)

            excel_bytes_all = build_excel_bytes(df_final)
            all_file_name = f"{file_name_prefix}_{date_str}.xlsx"

            st.download_button(
                label="📊 전체 결과물 다운로드 (모든 주차 통합)",
                data=excel_bytes_all,
                file_name=all_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.subheader("📅 주차별 다운로드")

            week_numbers = sorted(df_final["lesson_order_seq"].unique())
            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for week_num in week_numbers:
                    week_df = df_final[df_final["lesson_order_seq"] == week_num]
                    week_bytes = build_excel_bytes(week_df)
                    week_file_name = f"{file_name_prefix}_W{int(week_num):02d}_{date_str}.xlsx"
                    zf.writestr(week_file_name, week_bytes)

                    st.download_button(
                        label=f"⬇️ Week {int(week_num):02d} 다운로드 ({len(week_df)}개 단어)",
                        data=week_bytes,
                        file_name=week_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"week_download_{week_num}"
                    )

            st.download_button(
                label="🗂️ 주차별 파일 전체 ZIP으로 한꺼번에 다운로드",
                data=zip_buffer.getvalue(),
                file_name=f"{file_name_prefix}_ALL_{date_str}.zip",
                mime="application/zip"
            )
        else:
            st.warning("단어를 찾지 못했습니다. 파일 형식을 확인하거나 디버그 모드를 켜보세요.")

        if show_debug:
            with st.expander("🔍 디버그 데이터"):
                for i, txt in enumerate(raw_texts):
                    st.text(f"L{i+1}: {txt}")

# --------------------------------------------------------------------------
# 탭 2: 엑셀 검수
# --------------------------------------------------------------------------
with tab2:
    st.markdown("엑셀을 올리면 문서 작성 규정과 파일명 규칙을 자동으로 검사합니다. "
                "자동으로 고칠 수 있는 항목(불필요한 띄어쓰기, mp3 파일명)은 수정본을 내려받을 수 있고, "
                "고친 셀에는 **수정 차수 서식(바탕색 + 붉은색 글씨)**이 입혀집니다.")

    review_files = st.file_uploader("검수할 엑셀(.xlsx) 파일 (여러 개 가능)", type=["xlsx"],
                                    accept_multiple_files=True, key="review_upload")

    c1, c2 = st.columns(2)
    with c1:
        compare_source = st.file_uploader("(선택) 원본 워드/PDF와 대조", type=["docx", "pdf"], key="review_source")
    with c2:
        revision_styles = load_revision_styles()
        revision_label = st.selectbox("자동 수정 시 표시할 수정 차수", list(revision_styles.keys()))
        compare_fixed = st.checkbox("사이드바 고정값과도 비교", value=False,
                                    help="사이드바에 이번 교재의 고정값을 입력해둔 경우에만 켜세요.")

    if review_files:
        source_rows = None
        if compare_source is not None:
            src_mode, src_note = resolve_week_mode(compare_source.name, week_mode, level_code)
            if src_note:
                st.info("📅 원본 대조: " + src_note)
            source_rows, _ = parse_source_file(compare_source, src_mode)
            if not source_rows:
                st.warning("원본 파일에서 단어를 찾지 못해 원본 대조는 건너뜁니다.")
                source_rows = None

        expected_fixed = None
        if compare_fixed:
            expected_fixed = {
                "service_code": service_code, "track_code": track_code, "top_cors_id": top_cors_id,
                "level_code": level_code, "component_code": component_code, "book_code": book_code,
                "act_name": act_name,
            }

        all_issues = []
        fixed_outputs = []
        for f in review_files:
            file_bytes = f.getvalue()
            issues = check_workbook(f.name, file_bytes, source_rows, expected_fixed)
            all_issues.extend(issues)
            if any(it["자동수정"] for it in issues):
                fixed_bytes, n = apply_auto_fixes(file_bytes, issues, revision_styles[revision_label])
                fixed_outputs.append((f.name, fixed_bytes, n))

        st.subheader("📋 검수 결과")
        summary = []
        for f in review_files:
            f_issues = [i for i in all_issues if i["파일"] == f.name]
            summary.append({
                "파일": f.name,
                "오류": sum(1 for i in f_issues if i["수준"] == "오류"),
                "경고": sum(1 for i in f_issues if i["수준"] == "경고"),
                "자동수정 가능": sum(1 for i in f_issues if i["자동수정"]),
                "판정": "✅ 통과" if not any(i["수준"] == "오류" for i in f_issues) else "❌ 수정 필요",
            })
        st.dataframe(pd.DataFrame(summary), hide_index=True, use_container_width=True)

        if all_issues:
            df_issues = pd.DataFrame(all_issues)
            level_filter = st.multiselect("표시할 수준", ["오류", "경고"], default=["오류", "경고"])
            st.dataframe(df_issues[df_issues["수준"].isin(level_filter)], hide_index=True, use_container_width=True)

            report_buf = io.BytesIO()
            df_issues.to_excel(report_buf, index=False, engine="openpyxl")
            st.download_button("📝 검수 보고서 다운로드 (.xlsx)", report_buf.getvalue(),
                               file_name=f"검수보고서_{datetime.date.today():%y%m%d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("발견된 문제가 없습니다! 🎉")

        if fixed_outputs:
            st.subheader("🛠️ 자동 수정본")
            if len(fixed_outputs) == 1:
                name, data, n = fixed_outputs[0]
                st.download_button(f"⬇️ {name} 수정본 ({n}개 셀 수정)", data, file_name=name,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                zbuf = io.BytesIO()
                with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name, data, n in fixed_outputs:
                        zf.writestr(name, data)
                st.download_button(f"⬇️ 수정본 {len(fixed_outputs)}개 ZIP 다운로드", zbuf.getvalue(),
                                   file_name=f"수정본_{datetime.date.today():%y%m%d}.zip", mime="application/zip")
            st.caption("수정본은 원래 파일명 그대로 저장됩니다. 문서개정 시트의 개정 이력은 직접 기입해주세요.")

# --------------------------------------------------------------------------
# 탭 3: MP3 만들기 + 파일명/규격 검수
# --------------------------------------------------------------------------
with tab3:
    st.markdown("검수를 마친 엑셀을 올리면 **vocabulary_sound 열의 파일명 그대로** mp3를 만들고, "
                "만든 뒤 개수·파일명·용량·비트레이트·길이를 자동으로 검사합니다.")

    with st.expander("⚙️ 음성 설정 (기본 목소리: en-US-AriaNeural)"):
        cfg = dict(DEFAULT_TTS)
        cfg["VOICE"] = st.text_input("VOICE", DEFAULT_TTS["VOICE"])
        cfg["RATE"] = st.text_input("RATE", DEFAULT_TTS["RATE"])
        cfg["PITCH"] = st.text_input("PITCH", DEFAULT_TTS["PITCH"])
        cfg["REPEAT_COUNT"] = st.number_input("REPEAT_COUNT", 1, 5, DEFAULT_TTS["REPEAT_COUNT"])
        cfg["CONCURRENCY"] = st.number_input("CONCURRENCY (동시 처리 수)", 1, 5, DEFAULT_TTS["CONCURRENCY"])
        cfg["SLEEP_MS"] = st.number_input("SLEEP_MS (요청 사이 대기, ms)", 0, 5000, DEFAULT_TTS["SLEEP_MS"])
        cfg["MAX_WORDS"] = st.number_input("MAX_WORDS", 1, 1000, DEFAULT_TTS["MAX_WORDS"])
        cfg["BITRATE"] = st.text_input("BITRATE", DEFAULT_TTS["BITRATE"])
        st.markdown("**mp3 검수 기준**")
        min_sec = st.number_input("최소 길이(초)", 0.0, 60.0, 0.3, step=0.1)
        max_sec = st.number_input("최대 길이(초)", 0.5, 120.0, 15.0, step=0.5)

    mp3_excel = st.file_uploader("mp3를 만들 엑셀(.xlsx) 파일", type=["xlsx"], key="mp3_upload")

    if mp3_excel is not None:
        file_bytes = mp3_excel.getvalue()
        file_stem = os.path.splitext(mp3_excel.name)[0]

        # 시작 전 사전 점검: 파일명 관련 오류가 있으면 먼저 알려줌
        pre_issues = [i for i in check_workbook(mp3_excel.name, file_bytes)
                      if i["수준"] == "오류" and i["열"] in ("vocabulary", "vocabulary_sound", "-")]
        try:
            items = read_vocab_items(file_bytes)
        except Exception as e:
            items = []
            st.error(f"엑셀을 읽을 수 없습니다: {e}")

        if items:
            if len(items) > cfg["MAX_WORDS"]:
                st.warning(f"단어가 {len(items)}개라 MAX_WORDS({cfg['MAX_WORDS']})까지만 만듭니다.")
                items = items[:int(cfg["MAX_WORDS"])]

            st.write(f"단어 **{len(items)}개**를 찾았습니다.")
            with st.expander("읽을 문장과 파일명 미리보기"):
                st.dataframe(pd.DataFrame([{"단어": i["orig_word"], "읽을 문장": i["speech_text"],
                                            "파일명": i["final_filename"]} for i in items]),
                             hide_index=True, use_container_width=True)

            proceed = True
            if pre_issues:
                st.error(f"파일명/단어 관련 오류가 {len(pre_issues)}건 있습니다. ② 엑셀 검수 탭에서 먼저 확인하세요.")
                st.dataframe(pd.DataFrame(pre_issues)[["행", "열", "내용", "현재값", "제안값"]], hide_index=True)
                proceed = st.checkbox("오류를 확인했고, 그래도 진행합니다.")

            ffmpeg = get_ffmpeg()
            if not ffmpeg:
                st.error("ffmpeg를 찾을 수 없습니다. `pip install imageio-ffmpeg`를 실행해주세요.")
                proceed = False

            if st.button("🎙️ MP3 만들기 시작", type="primary", disabled=not proceed):
                out_dir = tempfile.mkdtemp(prefix="audio_")
                progress = st.progress(0.0, text="변환 준비 중...")
                status_box = st.empty()
                counter = {"done": 0, "fail": 0}

                def on_done(item, success):
                    counter["done"] += 1
                    if not success:
                        counter["fail"] += 1
                    progress.progress(counter["done"] / len(items),
                                      text=f"{counter['done']}/{len(items)} 완료 (실패 {counter['fail']}개)")
                    status_box.caption(("✅ " if success else "❌ ") + item["final_filename"])

                log_data = asyncio.run(run_tts(items, cfg, out_dir, ffmpeg, on_done))
                df_check, extras = check_mp3_files(items, out_dir, cfg["BITRATE"], min_sec, max_sec)
                zip_bytes, df_log = build_mp3_zip(items, log_data, out_dir)
                shutil.rmtree(out_dir, ignore_errors=True)

                st.session_state["mp3_result"] = {
                    "zip": zip_bytes, "zip_name": f"{file_stem}.zip",
                    "check": df_check, "extras": extras, "log": df_log, "total": len(items),
                }

    result = st.session_state.get("mp3_result")
    if result:
        st.subheader("🔎 MP3 검수 결과")
        df_check = result["check"]
        n_pass = int((df_check["결과"] == "통과").sum())
        n_fail = result["total"] - n_pass

        if n_fail == 0:
            st.success(f"전체 {result['total']}개 모두 통과했습니다! 🎉")
        else:
            st.error(f"{result['total']}개 중 {n_fail}개에 문제가 있습니다. 문제 있는 파일은 ZIP에서 빠지거나 확인이 필요합니다.")
            errors = result["log"][result["log"]["상태"] != "Success"]
            if len(errors) and errors["상태"].str.contains("403").all():
                st.warning("모든 실패가 403 에러입니다. edge-tts 서버가 막힌 상태일 수 있습니다. "
                           "`python -m pip install --upgrade edge-tts`로 업데이트한 뒤 다시 시도해보세요.")
        if result["extras"]:
            st.warning(f"엑셀에 없는 파일이 생성되었습니다: {', '.join(result['extras'])}")

        st.dataframe(df_check, hide_index=True, use_container_width=True)
        with st.expander("변환 로그 (build_log.csv 내용)"):
            st.dataframe(result["log"], hide_index=True, use_container_width=True)

        st.download_button(f"📦 {result['zip_name']} 다운로드", result["zip"],
                           file_name=result["zip_name"], mime="application/zip", type="primary")
