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
.main > div {
    max-width: 900px;
    margin: 0 auto;
}
[data-testid="stSidebar"] {
    width: 0 !important;
    min-width: 0 !important;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

QUESTION_CSS = """
<style>
.question-text {
    font-size: 1.05rem;
    line-height: 1.6;
}
</style>
"""
st.markdown(QUESTION_CSS, unsafe_allow_html=True)

LIKERT_CSS = """
<style>
[data-testid="stRadio"] label p {
    white-space: pre-line;
    text-align: center;
}
</style>
"""
st.markdown(LIKERT_CSS, unsafe_allow_html=True)

LIKERT_VALUES = [1, 2, 3, 4, 5]
LIKERT_LABELS = [
    "매우 그렇지 않다",
    "그렇지 않다",
    "보통이다",
    "그렇다",
    "매우 그렇다",
]

BASIC_FIELD_DEFAULTS = {
    "name": "",
    "gender": "",
    "student_id": "",
    "consent_research": False,
}


def format_likert_option(value: int) -> str:
    """Likert 옵션을 '숫자\\n라벨' 형태로 렌더링."""
    try:
        label = LIKERT_LABELS[value - 1]
    except Exception:
        label = ""
    return f"{value}\n{label}".strip()


def render_question_text(text: str):
    """모바일 가독성을 높인 질문 텍스트 렌더러."""
    if not text:
        return
    safe = html.escape(str(text)).replace("\n", "<br/>")
    st.markdown(f"<p class='question-text'>{safe}</p>", unsafe_allow_html=True)


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


def register_key(key: str):
    """저장할 질문 key를 전역 리스트에 모으기."""
    if "question_keys" not in st.session_state:
        st.session_state["question_keys"] = []
    if key not in st.session_state["question_keys"]:
        st.session_state["question_keys"].append(key)


def normalize_value(v):
    """CSV 저장 전에 리스트/튜플을 문자열로 변환."""
    if isinstance(v, (list, tuple)):
        return " / ".join(str(x) for x in v)
    return v


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
def load_instructions(path: Path = SURVEY_FILE) -> str:
    """안내문 시트에서 제목+본문 텍스트 가져오기."""
    path = Path(path)
    df = pd.read_excel(path, sheet_name="안내문", engine="openpyxl")
    title = str(df.loc[0, "Unnamed: 1"]).strip()
    body = str(df.loc[1, "Unnamed: 1"]).strip()
    return f"### {title}\n\n{body}"


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


def load_belong_blocks(path: Path = SURVEY_FILE):
    """
    [대학 소속감 및 비공식관계망] 시트에서
    3단 반구조화 블록(예/아니오 → 빈도 → 서술)을 모두 추출.
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


def load_dropout_blocks(path: Path = SURVEY_FILE):
    """
    [이탈 방지 질문지] 시트에서
    각 축별 3단 반구조화 블록(예/아니오 → 빈도 → 서술)을 추출.
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
        multi = "복수 선택" in qtext.replace(" ", "")
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


# -------------------------------------------------------------------
# 각 단계 화면 렌더링
# -------------------------------------------------------------------
def show_step_0_consent_and_basic():
    st.markdown(load_instructions())

    st.markdown("#### 기본 정보 입력")

    name_key = "name"
    gender_key = "gender"
    sid_key = "student_id"

    st.text_input("이름", key=name_key)
    st.radio(
        "성별",
        ["남", "여", "기타/응답 거절"],
        index=None,
        key=gender_key,
        horizontal=True,
    )
    st.text_input("학번", key=sid_key)

    register_key(name_key)
    register_key(gender_key)
    register_key(sid_key)

    consent_key = "consent_research"
    consent = st.checkbox("위 안내문을 읽었으며, 자발적으로 연구 참여에 동의합니다.", key=consent_key)
    register_key(consent_key)

    st.info("※ 동의 여부와 상관없이 언제든지 설문 참여를 중단하실 수 있습니다.")


def show_step_1_main_scale():
    st.markdown("### 1. 전공 및 진로 확신, 학습 역량 측정 문항")

    items = load_main_items()

    first_area = True
    for area, df_area in items.groupby("area"):
        if not first_area:
            st.divider()
        first_area = False
        for subscale, df_sub in df_area.groupby("subscale"):
            for _, row in df_sub.iterrows():
                num_str = str(row["item_no"]).strip()
                label = f"{num_str}. {row['item_text']}"
                key = make_key("quant", area, subscale, f"Q{num_str}")

                register_key(key)
                render_question_text(label)
                st.radio(
                    "",
                    LIKERT_VALUES,
                    index=None,
                    key=key,
                    horizontal=True,
                    format_func=format_likert_option,
                    label_visibility="collapsed",
                )


def show_step_2_belonging_and_informal():
    st.markdown("### 2. 대학 소속감 및 비공식 관계망 문항")

    blocks = load_belong_blocks()

    current_area = None
    for block in blocks:
        if block["area"] != current_area:
            if current_area is not None:
                st.divider()
            current_area = block["area"]

        dim = block["dimension"]

        # (1) 예/아니오 스크리닝
        yn_key = make_key("belong", current_area, dim, block["row_index"], "yn")
        register_key(yn_key)
        render_question_text(block["yesno_question"])
        st.radio(
            "",
            [clean_option(block["yes_label"]), clean_option(block["no_label"])],
            index=None,
            key=yn_key,
            horizontal=True,
            label_visibility="collapsed",
        )

        # (2) 빈도
        freq_key = make_key("belong", current_area, dim, block["row_index"], "freq")
        register_key(freq_key)
        freq_opts = [clean_option(o) for o in block["freq_options"]]
        render_question_text(block["freq_question"])
        st.radio(
            "",
            freq_opts,
            index=None,
            key=freq_key,
            horizontal=True,
            label_visibility="collapsed",
        )

        # (3) 서술형
        open_key = make_key("belong", current_area, dim, block["row_index"], "open")
        register_key(open_key)
        render_question_text(block["open_question"])
        st.text_area(
            "",
            key=open_key,
            label_visibility="collapsed",
        )


def show_step_3_dropout_scanning():
    st.markdown("### 3. 중도탈락 위험 스캐닝 문항")

    blocks = load_dropout_blocks()
    current_axis = None

    for block in blocks:
        if block["axis"] != current_axis:
            if current_axis is not None:
                st.divider()
            current_axis = block["axis"]

        # (1) 예/아니오
        yn_key = make_key("drop", current_axis, block["row_index"], "yn")
        register_key(yn_key)
        render_question_text(block["yesno_question"])
        st.radio(
            "",
            [clean_option(block["yes_label"]), clean_option(block["no_label"])],
            index=None,
            key=yn_key,
            horizontal=True,
            label_visibility="collapsed",
        )

        # (2) 빈도
        freq_key = make_key("drop", current_axis, block["row_index"], "freq")
        register_key(freq_key)
        freq_opts = [clean_option(o) for o in block["freq_options"]]
        render_question_text(block["freq_question"])
        st.radio(
            "",
            freq_opts,
            index=None,
            key=freq_key,
            horizontal=True,
            label_visibility="collapsed",
        )

        # (3) 서술형
        open_key = make_key("drop", current_axis, block["row_index"], "open")
        register_key(open_key)
        render_question_text(block["open_question"])
        st.text_area(
            "",
            key=open_key,
            label_visibility="collapsed",
        )


def show_step_4_background_and_programs():
    st.markdown("### 4. 학생 배경 및 교과·비교과 프로그램 참여 문항")

    questions = load_background_questions()
    current_section = None

    for q in questions:
        section = q["section"]
        if section != current_section:
            if current_section is not None:
                st.divider()
            current_section = section

        qtext = q["question"]
        options = [clean_option(o) for o in q["options"]]
        qtype = q["type"]
        optional = q["optional"]

        key = make_key("bg", current_section, q["idx"])
        register_key(key)

        question_label = qtext + (" (선택)" if optional else "")
        render_question_text(question_label)

        if qtype == "single" and options:
            # 단일 선택
            st.radio(
                "",
                options,
                index=None,
                key=key,
                horizontal=True,
                label_visibility="collapsed",
            )
        elif qtype == "multi" and options:
            # 복수 선택
            st.multiselect(
                "",
                options,
                key=key,
                label_visibility="collapsed",
            )
        else:
            # 자유 응답
            # 길이가 좀 있는 서술형일 가능성이 높으니 text_area 사용
            st.text_area(
                "",
                key=key,
                label_visibility="collapsed",
            )


def build_record():
    """session_state에 모인 응답을 한 행짜리 dict로 정리."""
    record = {}
    record["submitted_at"] = datetime.now().isoformat(timespec="seconds")

    for key, default in BASIC_FIELD_DEFAULTS.items():
        record[key] = normalize_value(st.session_state.get(key, default))

    keys = st.session_state.get("question_keys", [])
    for k in keys:
        if k in record:
            continue
        record[k] = normalize_value(st.session_state.get(k))
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
        if step < total - 1:
            if st.button("다음 단계 ▶"):
                st.session_state["step"] = step + 1
                st.rerun()
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
    if "question_keys" not in st.session_state:
        st.session_state["question_keys"] = []

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

    st.write(
        f"**현재 단계:** {step + 1} / {len(step_labels)} — {step_labels[step]}"
    )

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
