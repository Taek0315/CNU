import html
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

# -------------------------------------------------------------------
# 설정
# -------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
SURVEY_FILE = BASE_DIR / "측정문항_2025-1118_충남대학교 창의융합대학.xlsx"
OUTPUT_CSV = "responses.csv"

st.set_page_config(
    page_title="충남대학교 창의융합대학 종단연구 설문",
    layout="centered",
)

CUSTOM_CSS = """
<style>
section.main > div,
.main > div {
    max-width: 900px;
    margin: 0 auto;
    padding: 0 1.5rem 2rem;
}
@media (max-width: 640px) {
    section.main > div,
    .main > div {
        padding: 0 1rem 2rem;
    }
}
[data-testid="stSidebar"] {
    width: 0 !important;
    min-width: 0 !important;
}
.set-divider {
    border: none;
    border-top: 1px solid #e2e8f0;
    margin: 1.5rem 0 1.25rem;
}
[data-testid="stCheckbox"] > div {
    align-items: center;
}
[data-testid="stCheckbox"] > label {
    border: 1px solid #dbeafe;
    border-radius: 999px;
    padding: 0.35rem 0.75rem;
    width: 100%;
    display: flex;
    gap: 0.4rem;
}
[data-testid="stCheckbox"] > label span {
    flex: 1;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

QUESTION_CSS = """
<style>
.question-text {
    font-size: 1.05rem;
    line-height: 1.55;
    margin-bottom: 0.25rem;
}
.scale-desc {
    font-size: 0.95rem;
    color: #4b5563;
    margin: 0 0 0.35rem 0;
}
[data-testid="stRadio"] > div {
    gap: 0.45rem;
    flex-wrap: wrap;
}
[data-testid="stRadio"] label {
    border: 1px solid #dbeafe;
    border-radius: 999px;
    padding: 0.35rem 0.85rem;
    min-width: 2.5rem;
    justify-content: center;
}
[data-testid="stRadio"] label:hover {
    border-color: #93c5fd;
}
</style>
"""
st.markdown(QUESTION_CSS, unsafe_allow_html=True)

LIKERT_VALUES = [1, 2, 3, 4, 5]
LIKERT_LABELS = [
    "매우 그렇지 않다",
    "그렇지 않다",
    "보통이다",
    "그렇다",
    "매우 그렇다",
]

VALIDITY_ITEM_SPECS = [
    {
        "position": 23,
        "key": "validity_1",
        "text": "다음 보기 중 매우 그렇지 않다를 선택해주세요",
    },
    {
        "position": 44,
        "key": "validity_2",
        "text": "다음 보기 중 2번을 선택해주세요",
    },
    {
        "position": 65,
        "key": "validity_3",
        "text": "다음 보기 중 1+3의 값에 해당되는 숫자를 선택해주세요.",
    },
]

BASIC_FIELD_DEFAULTS = {
    "name": "",
    "gender": "",
    "student_id": "",
    "consent_research": False,
}


def get_response_store() -> dict[str, object]:
    """Widget 값이 사라져도 보존되는 응답 저장소."""
    if "_response_store" not in st.session_state:
        st.session_state["_response_store"] = {}
    return st.session_state["_response_store"]


def get_registered_keys() -> list[str]:
    """CSV/구글 시트 헤더 생성을 위한 키 목록."""
    if "question_keys" not in st.session_state:
        st.session_state["question_keys"] = []
    return st.session_state["question_keys"]


def register_key(key: str | None):
    """위젯/컬럼 키를 등록하고, 저장소에 값이 있으면 위젯 상태를 복원."""
    if not key:
        return
    keys = get_registered_keys()
    if key not in keys:
        keys.append(key)
    store = get_response_store()
    if key in store and key not in st.session_state:
        st.session_state[key] = store[key]


def record_answer(key: str | None, value):
    """현재 선택/입력 값을 별도 저장소에 기록."""
    if not key:
        return value
    store = get_response_store()
    store[key] = value
    return value


def get_saved_value(key: str, default=None):
    """위젯이 언마운트되어도 마지막 응답 값을 반환."""
    store = get_response_store()
    if key in store:
        return store[key]
    return st.session_state.get(key, default)


def reset_answer(key: str | None):
    """특정 키의 응답 값을 완전히 제거."""
    if not key:
        return
    store = get_response_store()
    if key in store:
        del store[key]
    if key in st.session_state:
        del st.session_state[key]


def render_question_text(text: str):
    """모바일 가독성을 높인 질문 텍스트 렌더러."""
    if not text:
        return
    safe = html.escape(str(text)).replace("\n", "<br/>")
    st.markdown(f"<p class='question-text'>{safe}</p>", unsafe_allow_html=True)


def render_pills(
    options: list,
    *,
    key: str,
    selection_mode: str = "single",
    format_func=None,
):
    """Streamlit pills helper with consistent styling."""
    register_key(key)
    formatter = format_func or (lambda x: x)
    value = st.pills(
        "",
        options=options,
        key=key,
        selection_mode=selection_mode,
        format_func=formatter,
        width="stretch",
        label_visibility="collapsed",
    )
    return record_answer(key, value)


def render_multi_checkbox_grid(
    *,
    key: str,
    options: list[str],
    option_keys: list[str],
    columns: int = 3,
    grid_threshold: int = 3,
):
    """복수 선택 항목을 칩 형태의 체크박스로 그리드 배치."""
    register_key(key)
    stored = st.session_state.get(key, [])
    if not isinstance(stored, list):
        stored = []

    total = len(options)
    if total == 0:
        st.session_state[key] = []
        return []

    col_count = 1 if total < grid_threshold else max(1, min(columns, total))

    for idx, option in enumerate(options):
        if idx % col_count == 0:
            cols = st.columns(col_count, gap="small")
        col = cols[idx % col_count]
        with col:
            checkbox_key = option_keys[idx]
            default_value = option in stored
            st.checkbox(
                option,
                key=checkbox_key,
                value=default_value,
            )

    selected = [
        options[i]
        for i in range(total)
        if st.session_state.get(option_keys[i], False)
    ]
    st.session_state[key] = selected
    record_answer(key, selected)
    return selected


# -------------------------------------------------------------------
# 유틸 함수
# -------------------------------------------------------------------
def make_key(*parts: str) -> str:
    """질문/영역 이름을 기반으로 위젯 key 및 저장용 컬럼명을 만들기."""
    text = "_".join(str(p) for p in parts if p not in [None, ""])
    text = re.sub(r"\s+", "_", text)  # 공백 -> _
    # 한글, 영문, 숫자, _ 만 남기기
    text = re.sub(r"[^\w가-힣_]+", "", text)
    return text


def clean_option(opt: str) -> str:
    """'❑', '□' 같은 불릿과 앞뒤 공백 제거."""
    if not isinstance(opt, str):
        return opt
    return re.sub(r"^[■□❑◻︎\s]+", "", opt).strip()


def normalize_value(v):
    """CSV 저장 전에 리스트/튜플을 문자열로 변환."""
    if isinstance(v, (list, tuple)):
        return " / ".join(str(x) for x in v)
    return v


def render_scale_description(
    low: str | None = None, mid: str | None = None, high: str | None = None
):
    """Likert 척도 안내 문구를 출력."""
    low = low or LIKERT_LABELS[0]
    mid = mid or LIKERT_LABELS[2]
    high = high or LIKERT_LABELS[4]
    desc = f"응답 척도: 1점 = {low} — 3점 = {mid} — 5점 = {high}"
    st.markdown(f"<p class='scale-desc'>{html.escape(desc)}</p>", unsafe_allow_html=True)


def clear_state(keys: list[str]):
    """조건 불충족 시 하위 문항 값을 초기화."""
    for key in keys:
        reset_answer(key)


def question_state_keys(question: dict) -> list[str]:
    """배경 문항 등에서 본문 key + 파생 key 목록을 반환."""
    keys = [question["key"]]
    keys.extend(question.get("option_keys", []))
    return keys


def clear_question_state(question: dict):
    """주어진 문항 관련 상태를 모두 초기화."""
    clear_state(question_state_keys(question))


def save_record_to_csv(record: dict, filename: str = OUTPUT_CSV):
    """한 응답자를 한 행으로 CSV에 누적 저장."""
    df_new = pd.DataFrame([record])

    path = Path(filename)
    if path.exists():
        try:
            df_old = pd.read_csv(path)
            df_all = pd.concat([df_old, df_new], ignore_index=True)
        except Exception:
            # 형식이 안 맞을 경우 그냥 새로 생성
            df_all = df_new
    else:
        df_all = df_new

    df_all.to_csv(path, index=False, encoding="utf-8-sig")


GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
GOOGLE_SPREADSHEET_ID = "1-25xZC2Z0u9Ac-pvfg7yG2doDKWrI5xaEYExLlCmVLE"


def get_worksheet():
    """st.secrets 기반으로 Google Sheets worksheet 객체를 반환."""
    if gspread is None or Credentials is None:
        return None

    try:
        service_account_info = st.secrets["gcp_service_account"]
    except Exception:
        return None

    try:
        credentials = Credentials.from_service_account_info(
            service_account_info, scopes=GOOGLE_SCOPES
        )
        client = gspread.authorize(credentials)
        spreadsheet = client.open_by_key(GOOGLE_SPREADSHEET_ID)
        worksheet = spreadsheet.get_worksheet(0)
    except Exception:
        return None

    return worksheet


def save_record_to_google_sheet(record: dict):
    """한 응답자를 Google Sheet에 한 행으로 저장."""
    worksheet = get_worksheet()
    if worksheet is None:
        raise RuntimeError("Google 스프레드시트 워크시트를 열 수 없습니다.")

    existing_header = worksheet.row_values(1)
    if not existing_header:
        headers = list(record.keys())
        worksheet.append_row(headers, value_input_option="USER_ENTERED")
        existing_header = headers

    new_keys = [k for k in record.keys() if k not in existing_header]
    if new_keys:
        updated_header = existing_header + new_keys
        worksheet.update("A1", [updated_header])
        existing_header = updated_header

    values = [record.get(k, "") for k in existing_header]
    worksheet.append_row(values, value_input_option="USER_ENTERED")


# -------------------------------------------------------------------
# 엑셀 로딩 / 파싱 함수
# -------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_instructions(path: Path = SURVEY_FILE) -> str:
    """안내문 시트에서 제목+본문 텍스트 가져오기."""
    path = Path(path)
    df = pd.read_excel(path, sheet_name="안내문", engine="openpyxl")
    title = str(df.loc[0, "Unnamed: 1"]).strip()
    body = str(df.loc[1, "Unnamed: 1"]).strip()
    return f"### {title}\n\n{body}"


@st.cache_data(show_spinner=False)
def load_main_items(path: Path = SURVEY_FILE) -> pd.DataFrame:
    """
    [전공 및 진로확신, 학습역량 측정 도구] 시트에서
    5개 영역 × 하위요인 × 80문항을 구조화하여 DataFrame으로 반환.
    """
    path = Path(path)
    df = pd.read_excel(path, sheet_name="전공 및 진로확신, 학습역량 측정 도구", engine="openpyxl")
    # '측정 영역'이 적혀 있는 행 찾기
    header_idx = df.index[df.iloc[:, 0] == "측정 영역"][0]

    items = df.iloc[header_idx + 1 :].copy()
    items.columns = [
        "area",
        "subscale",
        "item_no",
        "item_text",
        "opt1",
        "opt2",
        "opt3",
        "opt4",
        "opt5",
    ]
    items["area"] = items["area"].ffill()
    items["subscale"] = items["subscale"].ffill()
    items = items[items["item_no"].notna()].copy()

    return items


def inject_validity_items(items: list[dict]) -> list[dict]:
    """메인 Likert 문항 사이에 주어진 위치로 타당도 문항을 삽입."""
    if not items:
        return items

    sequence = list(items)
    specs = sorted(VALIDITY_ITEM_SPECS, key=lambda spec: spec["position"])
    for spec in specs:
        insert_idx = max(0, min(spec["position"] - 1, len(sequence)))
        reference_idx = min(max(insert_idx - 1, 0), len(sequence) - 1)
        reference = sequence[reference_idx]
        sequence.insert(
            insert_idx,
            {
                "area": reference.get("area"),
                "subscale": reference.get("subscale"),
                "item_no": None,
                "text": spec["text"],
                "key": spec["key"],
                "is_validity": True,
            },
        )
    return sequence


@st.cache_data(show_spinner=False)
def build_main_scale_sequence() -> list[dict]:
    """엑셀 문항 + 타당도 문항을 포함한 순차 리스트 생성."""
    df = load_main_items()
    questions = []
    for _, row in df.iterrows():
        num_str = str(row["item_no"]).strip()
        questions.append(
            {
                "area": row["area"],
                "subscale": row["subscale"],
                "item_no": num_str,
                "text": str(row["item_text"]).strip(),
                "key": make_key("quant", row["area"], row["subscale"], f"Q{num_str}"),
                "is_validity": False,
            }
        )
    return inject_validity_items(questions)


@st.cache_data(show_spinner=False)
def load_belong_blocks(path: Path = SURVEY_FILE):
    """
    [대학 소속감 및 비공식관계망] 시트에서
    3단 반구조화 블록(예/아니오 → 빈도 → 서술)을 모두 추출.

    NOTE: 문항 내용은 자유롭게 바꿔도 되지만, 각 블록이 정확히 3개의 연속된 행(예/아니오,
    빈도, 서술)과 Unnamed:3~7 열에 빈도 옵션이 배치된다는 구조는 유지되어야 합니다.
    """
    path = Path(path)
    df = pd.read_excel(path, sheet_name="대학 소속감 및 비공식관계망", engine="openpyxl")
    df = df.copy()
    df["area"] = df["대학 소속감 및 비공식관계망 질문지"]
    df["area"] = df["area"].replace("영역", np.nan)
    df["area"] = df["area"].ffill()

    starts = [
        idx
        for idx in df.index
        if isinstance(df.loc[idx, "Unnamed: 1"], str)
        and df.loc[idx, "Unnamed: 1"] not in ["하위 영역", "지시문"]
        and idx >= 3
    ]

    blocks = []
    for idx in starts:
        yes_row = df.loc[idx]
        freq_row = df.loc[idx + 1]
        open_row = df.loc[idx + 2]

        freq_opts = [
            freq_row[c]
            for c in ["Unnamed: 3", "Unnamed: 4", "Unnamed: 5", "Unnamed: 6", "Unnamed: 7"]
            if isinstance(freq_row[c], str)
        ]

        block = {
            "area": yes_row["area"],
            "dimension": yes_row["Unnamed: 1"],
            "yesno_question": yes_row["Unnamed: 2"],
            "yes_label": yes_row["Unnamed: 3"],
            "no_label": yes_row["Unnamed: 4"],
            "freq_question": freq_row["Unnamed: 2"],
            "freq_options": freq_opts,
            "open_question": open_row["Unnamed: 2"],
            "row_index": idx,
        }
        blocks.append(block)

    return blocks


@st.cache_data(show_spinner=False)
def load_dropout_blocks(path: Path = SURVEY_FILE):
    """
    [이탈 방지 질문지] 시트에서
    각 축별 3단 반구조화 블록(예/아니오 → 빈도 → 서술)을 추출.

    NOTE: 구조는 [예/아니오 행 → 빈도 행 → 서술 행] 순서를 가정합니다. 문항 텍스트나
    옵션 내용은 자유롭게 바꿀 수 있지만, 이 3행 패턴이 깨지면 파서 수정이 필요합니다.
    """
    path = Path(path)
    df = pd.read_excel(path, sheet_name="이탈 방지 질문지", engine="openpyxl")
    df = df.copy()

    starts = [
        idx
        for idx in df.index
        if isinstance(df.loc[idx, "대학생 이탈 방지 질문지"], str)
        and df.loc[idx, "대학생 이탈 방지 질문지"] not in ["지시문", "측정 영역"]
    ]

    blocks = []
    for idx in starts:
        yes_row = df.loc[idx]
        freq_row = df.loc[idx + 1]
        open_row = df.loc[idx + 2]

        freq_opts = [
            freq_row[c]
            for c in ["Unnamed: 2", "Unnamed: 3", "Unnamed: 4", "Unnamed: 5", "Unnamed: 6"]
            if isinstance(freq_row[c], str)
        ]

        block = {
            "axis": yes_row["대학생 이탈 방지 질문지"],
            "yesno_question": yes_row["Unnamed: 1"],
            "yes_label": yes_row["Unnamed: 2"],
            "no_label": yes_row["Unnamed: 3"],
            "freq_question": freq_row["Unnamed: 1"],
            "freq_options": freq_opts,
            "open_question": open_row["Unnamed: 1"],
            "row_index": idx,
        }
        blocks.append(block)

    return blocks


@st.cache_data(show_spinner=False)
def load_background_questions(path: Path = SURVEY_FILE):
    """
    [질적정보] 시트에서 학생 배경 및 교과/비교과 참여 문항 전부를 추출.
    - section: 큰 영역(창의 융합 대학 분류, 학력/입학, 가정/SES, 생활/근로, 교과 참여, 비교과 참여 등)
    - question: 실제 질문 텍스트
    - options: 보기 목록(없으면 자유 응답)
    - type: 'single' | 'multi' | 'text'
    - optional: 선택 문항 여부(괄호에 '선택'이 들어있는 경우 등)
    """
    path = Path(path)
    df = pd.read_excel(path, sheet_name="질적정보", engine="openpyxl")
    df = df.copy()

    df["section"] = df.iloc[:, 0]
    df.loc[df["section"].isin(["지시문", "영역"]), "section"] = np.nan
    df["section"] = df["section"].ffill()

    option_cols = [c for c in df.columns if c.startswith("Unnamed:") and c != "Unnamed: 1"]

    questions = []
    for idx, row in df.iterrows():
        if idx < 3:
            continue
        qtext = row.get("Unnamed: 1", None)
        if not isinstance(qtext, str) or qtext.strip() == "":
            continue

        options = [row[c] for c in option_cols if isinstance(row[c], str)]
        normalized_text = re.sub(r"\s+", "", qtext)
        multi = "복수선택" in normalized_text
        optional = "선택" in qtext

        qtype = "text"
        if len(options) > 0:
            qtype = "multi" if multi else "single"

        questions.append(
            {
                "idx": idx,
                "section": row["section"],
                "question": qtext,
                "options": options,
                "type": qtype,
                "optional": optional,
            }
        )

    return questions


BACKGROUND_BRANCH_RULES = [
    {
        "parent_idx": 36,
        "child_indices": [37, 38],
        "affirmative_values": None,
    },
]

BACKGROUND_ENTRY_IDX = 3
PROGRAM_ONLY_SECTIONS = {"교과 참여", "비교과 참여", "전공탐색 지원"}
FIRST_CHOICE_TOKEN = "1지망"
FIRST_CHOICE_FOLLOWUP_LABEL = "어떤 1지망 계열/학과였는지 작성해주세요."


@st.cache_data(show_spinner=False)
def get_main_scale_items() -> list[dict]:
    """캐시된 메인 척도 문항 리스트."""
    return build_main_scale_sequence()


@st.cache_data(show_spinner=False)
def get_belong_blocks_with_keys() -> list[dict]:
    """소속감/비공식 관계망 블록에 위젯 key 및 옵션을 부여."""
    annotated = []
    for block in load_belong_blocks():
        key_prefix = make_key("belong", block["area"], block["dimension"], block["row_index"])
        annotated.append(
            {
                **block,
                "yes_key": f"{key_prefix}_yn",
                "freq_key": f"{key_prefix}_freq",
                "open_key": f"{key_prefix}_open",
                "yes_value": clean_option(block["yes_label"]),
                "no_value": clean_option(block["no_label"]),
                "freq_options": [clean_option(opt) for opt in block["freq_options"]],
            }
        )
    return annotated


@st.cache_data(show_spinner=False)
def get_dropout_blocks_with_keys() -> list[dict]:
    """중도탈락 스캐닝 블록에 위젯 key 및 옵션을 부여."""
    annotated = []
    for block in load_dropout_blocks():
        key_prefix = make_key("drop", block["axis"], block["row_index"])
        annotated.append(
            {
                **block,
                "yes_key": f"{key_prefix}_yn",
                "freq_key": f"{key_prefix}_freq",
                "open_key": f"{key_prefix}_open",
                "yes_value": clean_option(block["yes_label"]),
                "no_value": clean_option(block["no_label"]),
                "freq_options": [clean_option(opt) for opt in block["freq_options"]],
            }
        )
    return annotated


@st.cache_data(show_spinner=False)
def get_background_question_bank():
    """배경 문항 + 분기 메타데이터."""
    questions = []
    entry_question = None
    for q in load_background_questions():
        options = [clean_option(opt) for opt in q["options"]]
        key = make_key("bg", q["section"], q["idx"])
        option_keys = [f"{key}__opt_{idx}" for idx in range(len(options))]
        question_entry = {**q, "options": options, "key": key, "option_keys": option_keys}
        questions.append(question_entry)
        if q["idx"] == BACKGROUND_ENTRY_IDX:
            entry_question = question_entry

    parent_children: dict[str, list[str]] = {}
    child_parent: dict[str, str] = {}
    parent_affirmative: dict[str, list[str]] = {}

    for rule in BACKGROUND_BRANCH_RULES:
        parent = next((q for q in questions if q["idx"] == rule["parent_idx"]), None)
        if parent is None:
            continue
        parent_key = parent["key"]
        children_keys = []
        for child_idx in rule.get("child_indices", []):
            child = next((q for q in questions if q["idx"] == child_idx), None)
            if child is None:
                continue
            children_keys.append(child["key"])
            child_parent[child["key"]] = parent_key
        if children_keys:
            parent_children[parent_key] = children_keys

        if rule.get("affirmative_values"):
            show_values = [clean_option(v) for v in rule["affirmative_values"]]
        elif parent["options"]:
            show_values = [parent["options"][0]]
        else:
            show_values = ["예"]
        cleaned_values = [val for val in show_values if val]
        if not cleaned_values:
            cleaned_values = ["예"]
        parent_affirmative[parent_key] = cleaned_values

    entry_question_key = entry_question["key"] if entry_question else None
    entry_affirm_values = []
    if entry_question:
        entry_affirm_values = [
            opt for opt in entry_question["options"] if "창의" in opt
        ]
        if not entry_affirm_values and entry_question["options"]:
            entry_affirm_values = [entry_question["options"][0]]

    question_lookup = {q["key"]: q for q in questions}

    return {
        "questions": questions,
        "parent_children": parent_children,
        "child_parent": child_parent,
        "parent_affirmative": parent_affirmative,
        "entry_question_key": entry_question_key,
        "entry_affirm_values": entry_affirm_values,
        "question_lookup": question_lookup,
    }


def get_first_choice_major_meta(questions: list[dict] | None = None) -> tuple[str | None, str | None]:
    """
    1지망 계열/학과 여부 문항과 텍스트 후속 질문 key를 찾는다.

    NOTE: 질적정보 시트에서 해당 문항 텍스트에 '1지망' 문자열이 포함되어 있다는 가정에
    의존한다. Excel에서 해당 문구를 완전히 바꾸면 FIRST_CHOICE_TOKEN을 업데이트해야 한다.
    """
    source_questions = questions
    if source_questions is None:
        source_questions = get_background_question_bank()["questions"]
    for question in source_questions:
        question_text = question.get("question")
        if isinstance(question_text, str) and FIRST_CHOICE_TOKEN in question_text:
            base_key = question.get("key")
            major_key = make_key(base_key, "1지망", "학과", "텍스트")
            return base_key, major_key
    return None, None


def is_answer_missing(value) -> bool:
    """필수 응답 누락 여부 판별."""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, (list, tuple, set)):
        return len(value) == 0
    return False


def validate_main_scale_required() -> bool:
    """메인 척도 전체 문항 응답 확인."""
    for item in get_main_scale_items():
        value = get_saved_value(item["key"])
        if is_answer_missing(value):
            st.warning("모든 전공 및 진로 확신 · 학습 역량 문항에 응답해주세요.")
            return False
    return True


def validate_yesno_screeners(blocks: list[dict], warning_message: str) -> bool:
    """예/아니오 스크리닝 문항 응답 확인."""
    for block in blocks:
        value = get_saved_value(block["yes_key"])
        if is_answer_missing(value):
            st.warning(warning_message)
            return False
    return True


def can_advance_from_step(step: int) -> bool:
    """다음 단계로 이동 가능한지 검증."""
    if step == 1:
        return validate_main_scale_required()
    if step == 2:
        return validate_yesno_screeners(
            get_belong_blocks_with_keys(),
            "대학 소속감 및 비공식 관계망의 모든 예/아니오 문항을 응답해주세요.",
        )
    if step == 3:
        return validate_yesno_screeners(
            get_dropout_blocks_with_keys(),
            "중도 탈락 문항의 모든 예/아니오 문항을 응답해주세요.",
        )
    return True


# -------------------------------------------------------------------
# 각 단계 화면 렌더링
# -------------------------------------------------------------------
def show_step_0_consent_and_basic():
    st.markdown(load_instructions())

    st.markdown("#### 기본 정보 입력")

    name_key = "name"
    gender_key = "gender"
    sid_key = "student_id"

    render_question_text("이름")
    register_key(name_key)
    name_value = st.text_input(
        "",
        key=name_key,
        label_visibility="collapsed",
        placeholder="이름을 입력하세요",
    )
    record_answer(name_key, name_value)

    render_question_text("성별")
    render_pills(["남자", "여자"], key=gender_key)

    render_question_text("학번")
    register_key(sid_key)
    sid_value = st.text_input(
        "",
        key=sid_key,
        label_visibility="collapsed",
        placeholder="학번을 입력하세요",
    )
    record_answer(sid_key, sid_value)

    consent_key = "consent_research"
    render_question_text("위 안내문을 읽었으며, 자발적으로 연구 참여에 동의합니다.")
    register_key(consent_key)
    consent_value = st.checkbox("동의합니다.", key=consent_key)
    record_answer(consent_key, consent_value)

    st.info("※ 동의 여부와 상관없이 언제든지 설문 참여를 중단하실 수 있습니다.")


def show_step_1_main_scale():
    st.markdown("### 1. 전공 및 진로 확신, 학습 역량 측정 문항")

    items = get_main_scale_items()

    current_area = None
    for item in items:
        area = item.get("area")
        if area != current_area:
            if current_area is not None:
                st.divider()
            current_area = area

        key = item["key"]
        render_question_text(item["text"])
        render_scale_description()
        render_pills(
            LIKERT_VALUES,
            key=key,
            format_func=lambda v: str(v),
        )


def show_step_2_belonging_and_informal():
    st.markdown("### 2. 대학 소속감 및 비공식 관계망 문항")

    blocks = get_belong_blocks_with_keys()

    current_area = None
    total_blocks = len(blocks)
    for idx, block in enumerate(blocks):
        if block["area"] != current_area:
            if current_area is not None:
                st.divider()
            current_area = block["area"]

        # (1) 예/아니오 스크리닝
        yn_key = block["yes_key"]
        render_question_text(block["yesno_question"])
        yn_answer = render_pills(
            [block["yes_value"], block["no_value"]],
            key=yn_key,
        )

        show_followups = yn_answer == block["yes_value"]
        followup_keys = [block["freq_key"], block["open_key"]]
        for fk in followup_keys:
            register_key(fk)

        if show_followups:
            # (2) 빈도
            freq_opts = block["freq_options"]
            render_question_text(block["freq_question"])
            freq_value = st.radio(
                "",
                freq_opts,
                index=None,
                key=block["freq_key"],
                horizontal=False,
                label_visibility="collapsed",
            )
            record_answer(block["freq_key"], freq_value)

            # (3) 서술형
            render_question_text(block["open_question"])
            open_value = st.text_area(
                "",
                key=block["open_key"],
                label_visibility="collapsed",
            )
            record_answer(block["open_key"], open_value)
        else:
            clear_state(followup_keys)

        if idx < total_blocks - 1:
            next_area = blocks[idx + 1]["area"]
            if next_area == block["area"]:
                st.markdown("<hr class='set-divider' />", unsafe_allow_html=True)


def show_step_3_dropout_scanning():
    st.markdown("### 3. 중도탈락 위험 스캐닝 문항")

    blocks = get_dropout_blocks_with_keys()
    current_axis = None

    for block in blocks:
        if block["axis"] != current_axis:
            if current_axis is not None:
                st.divider()
            current_axis = block["axis"]

        # (1) 예/아니오
        yn_key = block["yes_key"]
        render_question_text(block["yesno_question"])
        yn_answer = render_pills(
            [block["yes_value"], block["no_value"]],
            key=yn_key,
        )

        show_followups = yn_answer == block["yes_value"]
        followup_keys = [block["freq_key"], block["open_key"]]
        for fk in followup_keys:
            register_key(fk)

        if show_followups:
            # (2) 빈도
            freq_opts = block["freq_options"]
            render_question_text(block["freq_question"])
            freq_value = st.radio(
                "",
                freq_opts,
                index=None,
                key=block["freq_key"],
                horizontal=False,
                label_visibility="collapsed",
            )
            record_answer(block["freq_key"], freq_value)

            # (3) 서술형
            render_question_text(block["open_question"])
            open_value = st.text_area(
                "",
                key=block["open_key"],
                label_visibility="collapsed",
            )
            record_answer(block["open_key"], open_value)
        else:
            clear_state(followup_keys)


def show_step_4_background_and_programs():
    st.markdown("### 4. 학생 배경 및 교과·비교과 프로그램 참여 문항")

    bg_data = get_background_question_bank()
    questions = bg_data["questions"]
    parent_children = bg_data["parent_children"]
    child_parent = bg_data["child_parent"]
    parent_affirm = bg_data["parent_affirmative"]
    entry_key = bg_data.get("entry_question_key")
    entry_affirm_values = bg_data.get("entry_affirm_values", [])
    question_lookup = bg_data.get("question_lookup", {})
    current_section = None
    entry_answer = get_saved_value(entry_key, "") if entry_key else ""
    show_program_sections = (
        not entry_key or (entry_answer and entry_answer in entry_affirm_values)
    )
    first_choice_key, first_choice_major_key = get_first_choice_major_meta(questions)

    for q in questions:
        section = q["section"]
        if section != current_section:
            if current_section is not None:
                st.divider()
            current_section = section

        qtext = q["question"]
        options = q["options"]
        qtype = q["type"]
        optional = q["optional"]

        key = q["key"]
        register_key(key)

        if section in PROGRAM_ONLY_SECTIONS and not show_program_sections:
            clear_question_state(q)
            if first_choice_major_key and key == first_choice_key:
                reset_answer(first_choice_major_key)
            continue

        parent_key = child_parent.get(key)
        if parent_key:
            allowed = parent_affirm.get(parent_key, [])
            parent_answer = get_saved_value(parent_key, "")
            if not parent_answer or parent_answer not in allowed:
                clear_question_state(q)
                if first_choice_major_key and key == first_choice_key:
                    reset_answer(first_choice_major_key)
                continue

        question_label = qtext + (" (선택)" if optional else "")
        render_question_text(question_label)

        if qtype == "single" and options:
            if len(options) >= 3:
                render_pills(
                    options,
                    key=key,
                )
            else:
                value = st.radio(
                    "",
                    options,
                    index=None,
                    key=key,
                    horizontal=False,
                    label_visibility="collapsed",
                )
                record_answer(key, value)
        elif qtype == "multi" and options:
            render_multi_checkbox_grid(
                key=key,
                options=options,
                option_keys=q.get("option_keys", []),
            )
        else:
            value = st.text_area(
                "",
                key=key,
                label_visibility="collapsed",
            )
            record_answer(key, value)

        if key == first_choice_key and first_choice_major_key:
            register_key(first_choice_major_key)
            choice_answer = get_saved_value(key, "")
            if choice_answer and "있다" in str(choice_answer):
                render_question_text(FIRST_CHOICE_FOLLOWUP_LABEL)
                default_major = get_saved_value(first_choice_major_key, "")
                major_value = st.text_input(
                    "",
                    value=default_major,
                    key=first_choice_major_key,
                    label_visibility="collapsed",
                )
                record_answer(first_choice_major_key, major_value)
            else:
                reset_answer(first_choice_major_key)

        if key in parent_children:
            allowed = parent_affirm.get(key, [])
            answer = get_saved_value(key, "")
            if not answer or answer not in allowed:
                keys_to_clear: list[str] = []
                for child_key in parent_children[key]:
                    child_question = question_lookup.get(child_key)
                    if child_question:
                        keys_to_clear.extend(question_state_keys(child_question))
                    else:
                        keys_to_clear.append(child_key)
                clear_state(keys_to_clear)


def build_record():
    """모든 단계 응답을 하나의 flat dict로 정리."""
    record: dict[str, object] = {}
    record["submitted_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    for key, default in BASIC_FIELD_DEFAULTS.items():
        value = get_saved_value(key, default)
        if key == "consent_research":
            record[key] = bool(value)
        else:
            record[key] = normalize_value(value or "")

    # 1. 메인 Likert 척도 (타당도 문항 포함)
    for item in get_main_scale_items():
        record[item["key"]] = normalize_value(get_saved_value(item["key"], ""))

    # 2. 소속감/비공식 관계망
    for block in get_belong_blocks_with_keys():
        yn_value = get_saved_value(block["yes_key"], "")
        record[block["yes_key"]] = normalize_value(yn_value)
        followup_allowed = yn_value == block["yes_value"]
        if followup_allowed:
            record[block["freq_key"]] = normalize_value(
                get_saved_value(block["freq_key"], "")
            )
            record[block["open_key"]] = normalize_value(
                get_saved_value(block["open_key"], "")
            )
        else:
            record[block["freq_key"]] = ""
            record[block["open_key"]] = ""

    # 3. 중도탈락 스캐닝
    for block in get_dropout_blocks_with_keys():
        yn_value = get_saved_value(block["yes_key"], "")
        record[block["yes_key"]] = normalize_value(yn_value)
        followup_allowed = yn_value == block["yes_value"]
        if followup_allowed:
            record[block["freq_key"]] = normalize_value(
                get_saved_value(block["freq_key"], "")
            )
            record[block["open_key"]] = normalize_value(
                get_saved_value(block["open_key"], "")
            )
        else:
            record[block["freq_key"]] = ""
            record[block["open_key"]] = ""

    # 4. 배경/교과·비교과
    bg_data = get_background_question_bank()
    parent_affirm = bg_data["parent_affirmative"]
    child_parent = bg_data["child_parent"]
    first_choice_key, first_choice_major_key = get_first_choice_major_meta(bg_data["questions"])

    for q in bg_data["questions"]:
        key = q["key"]
        parent_key = child_parent.get(key)
        if parent_key:
            allowed = parent_affirm.get(parent_key, [])
            parent_answer = get_saved_value(parent_key, "")
            if not parent_answer or parent_answer not in allowed:
                record[key] = ""
                continue

        record[key] = normalize_value(get_saved_value(key, ""))
        if first_choice_major_key and key == first_choice_key:
            record[first_choice_major_key] = normalize_value(
                get_saved_value(first_choice_major_key, "")
            )

    return record


def render_navigation(step_labels):
    total = len(step_labels)
    step = st.session_state.get("step", 0)

    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])

    with col1:
        if step > 0:
            if st.button("◀ 이전 단계"):
                st.session_state["step"] = step - 1
                st.rerun()

    with col3:
        consent_ok = bool(get_saved_value("consent_research", False))
        next_disabled = step == 0 and not consent_ok

        if step < total - 1:
            if st.button("다음 단계 ▶", disabled=next_disabled):
                if can_advance_from_step(step):
                    st.session_state["step"] = step + 1
                    st.rerun()
            if next_disabled:
                st.caption("연구 참여에 동의해야 다음 단계로 이동할 수 있습니다.")
        else:
            if st.button("응답 제출"):
                record = build_record()
                save_record_to_csv(record)
                try:
                    save_record_to_google_sheet(record)
                except Exception:
                    st.warning(
                        "응답은 CSV에 저장되었으나 Google 스프레드시트 저장 중 오류가 발생했습니다."
                    )
                st.success("응답이 저장되었습니다. 참여해 주셔서 감사합니다.")
                st.balloons()


# -------------------------------------------------------------------
# 메인
# -------------------------------------------------------------------
def main():
    if "step" not in st.session_state:
        st.session_state["step"] = 0

    step_labels = [
        "참여 동의 및 기본 정보",
        "전공 및 진로 확신 · 학습 역량",
        "대학 소속감 및 비공식 관계망",
        "중도탈락 위험 스캐닝",
        "학생 배경 및 교과·비교과 참여",
    ]

    step = st.session_state["step"]
    step = max(0, min(step, len(step_labels) - 1))
    st.session_state["step"] = step

    if step == 0:
        st.title("충남대학교 창의융합대학 종단연구 설문")

    if step == 0:
        show_step_0_consent_and_basic()
    elif step == 1:
        show_step_1_main_scale()
    elif step == 2:
        show_step_2_belonging_and_informal()
    elif step == 3:
        show_step_3_dropout_scanning()
    elif step == 4:
        show_step_4_background_and_programs()

    render_navigation(step_labels)


if __name__ == "__main__":
    main()
