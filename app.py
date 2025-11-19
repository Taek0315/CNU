import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# -------------------------------------------------------------------
# 설정
# -------------------------------------------------------------------
SURVEY_FILE = "측정문항_2025-1118_충남대학교 창의융합대학.xlsx"
OUTPUT_CSV = "responses.csv"

st.set_page_config(
    page_title="충남대학교 창의융합대학 종단연구 설문",
    layout="wide",
)


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


# -------------------------------------------------------------------
# 엑셀 로딩 / 파싱 함수
# -------------------------------------------------------------------
def load_instructions(path: str = SURVEY_FILE) -> str:
    """안내문 시트에서 제목+본문 텍스트 가져오기."""
    df = pd.read_excel(path, sheet_name="안내문")
    title = str(df.loc[0, "Unnamed: 1"]).strip()
    body = str(df.loc[1, "Unnamed: 1"]).strip()
    return f"### {title}\n\n{body}"


def load_main_items(path: str = SURVEY_FILE) -> pd.DataFrame:
    """
    [전공 및 진로확신, 학습역량 측정 도구] 시트에서
    5개 영역 × 하위요인 × 80문항을 구조화하여 DataFrame으로 반환.
    """
    df = pd.read_excel(path, sheet_name="전공 및 진로확신, 학습역량 측정 도구")
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


def load_belong_blocks(path: str = SURVEY_FILE):
    """
    [대학 소속감 및 비공식관계망] 시트에서
    3단 반구조화 블록(예/아니오 → 빈도 → 서술)을 모두 추출.
    """
    df = pd.read_excel(path, sheet_name="대학 소속감 및 비공식관계망")
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


def load_dropout_blocks(path: str = SURVEY_FILE):
    """
    [이탈 방지 질문지] 시트에서
    각 축별 3단 반구조화 블록(예/아니오 → 빈도 → 서술)을 추출.
    """
    df = pd.read_excel(path, sheet_name="이탈 방지 질문지")
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


def load_background_questions(path: str = SURVEY_FILE):
    """
    [질적정보] 시트에서 학생 배경 및 교과/비교과 참여 문항 전부를 추출.
    - section: 큰 영역(창의 융합 대학 분류, 학력/입학, 가정/SES, 생활/근로, 교과 참여, 비교과 참여 등)
    - question: 실제 질문 텍스트
    - options: 보기 목록(없으면 자유 응답)
    - type: 'single' | 'multi' | 'text'
    - optional: 선택 문항 여부(괄호에 '선택'이 들어있는 경우 등)
    """
    df = pd.read_excel(path, sheet_name="질적정보")
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

    col1, col2 = st.columns(2)
    with col1:
        name_key = "basic_name"
        sid_key = "basic_student_id"
        name = st.text_input("이름", key=name_key)
        sid = st.text_input("학번", key=sid_key)
        register_key(name_key)
        register_key(sid_key)

    with col2:
        gender_key = "basic_gender"
        grade_key = "basic_grade"
        major_type_key = "basic_major_type"

        gender = st.radio("성별", ["남", "여", "기타/응답 거절"], index=None, key=gender_key)
        grade = st.selectbox("학년", ["1학년", "2학년", "3학년", "4학년 이상"], index=None, key=grade_key)
        major_type = st.radio(
            "입학 유형",
            ["창의융합대학 전공자율선택제", "일반 학과 입학"],
            index=None,
            key=major_type_key,
        )

        register_key(gender_key)
        register_key(grade_key)
        register_key(major_type_key)

    consent_key = "consent_research"
    consent = st.checkbox("위 안내문을 읽었으며, 자발적으로 연구 참여에 동의합니다.", key=consent_key)
    register_key(consent_key)

    st.info("※ 동의 여부와 상관없이 언제든지 설문 참여를 중단하실 수 있습니다.")


def show_step_1_main_scale():
    st.markdown("### 1. 전공 및 진로 확신, 학습 역량 측정 문항")

    items = load_main_items()

    likert_5 = [
        "1 매우 그렇지 않다",
        "2 그렇지 않다",
        "3 보통이다",
        "4 그렇다",
        "5 매우 그렇다",
    ]

    for area, df_area in items.groupby("area"):
        st.markdown(f"#### 영역: {area}")
        for subscale, df_sub in df_area.groupby("subscale"):
            st.markdown(f"**하위요인: {subscale}**")
            for _, row in df_sub.iterrows():
                num_str = str(row["item_no"]).strip()
                label = f"{num_str}. {row['item_text']}"
                key = make_key("quant", area, subscale, f"Q{num_str}")

                register_key(key)
                st.radio(
                    label,
                    likert_5,
                    index=None,
                    key=key,
                    horizontal=True,
                )


def show_step_2_belonging_and_informal():
    st.markdown("### 2. 대학 소속감 및 비공식 관계망 문항")

    blocks = load_belong_blocks()

    current_area = None
    for block in blocks:
        if block["area"] != current_area:
            current_area = block["area"]
            st.markdown(f"#### 영역: {current_area}")

        dim = block["dimension"]
        st.markdown(f"**하위영역: {dim}**")

        # (1) 예/아니오 스크리닝
        yn_key = make_key("belong", current_area, dim, block["row_index"], "yn")
        register_key(yn_key)
        st.radio(
            block["yesno_question"],
            [clean_option(block["yes_label"]), clean_option(block["no_label"])],
            index=None,
            key=yn_key,
            horizontal=True,
        )

        # (2) 빈도
        freq_key = make_key("belong", current_area, dim, block["row_index"], "freq")
        register_key(freq_key)
        freq_opts = [clean_option(o) for o in block["freq_options"]]
        st.radio(
            block["freq_question"],
            freq_opts,
            index=None,
            key=freq_key,
            horizontal=True,
        )

        # (3) 서술형
        open_key = make_key("belong", current_area, dim, block["row_index"], "open")
        register_key(open_key)
        st.text_area(block["open_question"], key=open_key)


def show_step_3_dropout_scanning():
    st.markdown("### 3. 중도탈락 위험 스캐닝 문항")

    blocks = load_dropout_blocks()
    current_axis = None

    for block in blocks:
        if block["axis"] != current_axis:
            current_axis = block["axis"]
            st.markdown(f"#### 축: {current_axis}")

        # (1) 예/아니오
        yn_key = make_key("drop", current_axis, block["row_index"], "yn")
        register_key(yn_key)
        st.radio(
            block["yesno_question"],
            [clean_option(block["yes_label"]), clean_option(block["no_label"])],
            index=None,
            key=yn_key,
            horizontal=True,
        )

        # (2) 빈도
        freq_key = make_key("drop", current_axis, block["row_index"], "freq")
        register_key(freq_key)
        freq_opts = [clean_option(o) for o in block["freq_options"]]
        st.radio(
            block["freq_question"],
            freq_opts,
            index=None,
            key=freq_key,
            horizontal=True,
        )

        # (3) 서술형
        open_key = make_key("drop", current_axis, block["row_index"], "open")
        register_key(open_key)
        st.text_area(block["open_question"], key=open_key)


def show_step_4_background_and_programs():
    st.markdown("### 4. 학생 배경 및 교과·비교과 프로그램 참여 문항")

    questions = load_background_questions()
    current_section = None

    for q in questions:
        section = q["section"]
        if section != current_section:
            current_section = section
            st.markdown(f"#### 영역: {current_section}")

        qtext = q["question"]
        options = [clean_option(o) for o in q["options"]]
        qtype = q["type"]
        optional = q["optional"]

        key = make_key("bg", current_section, q["idx"])
        register_key(key)

        if qtype == "single" and options:
            # 단일 선택
            st.radio(
                qtext + (" (선택)" if optional else ""),
                options,
                index=None,
                key=key,
                horizontal=True,
            )
        elif qtype == "multi" and options:
            # 복수 선택
            st.multiselect(
                qtext + (" (선택)" if optional else ""),
                options,
                key=key,
            )
        else:
            # 자유 응답
            # 길이가 좀 있는 서술형일 가능성이 높으니 text_area 사용
            st.text_area(
                qtext + (" (선택)" if optional else ""),
                key=key,
            )


def build_record():
    """session_state에 모인 응답을 한 행짜리 dict로 정리."""
    record = {
        "submitted_at": datetime.now().isoformat(timespec="seconds"),
    }
    keys = st.session_state.get("question_keys", [])
    for k in keys:
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

    st.title("충남대학교 창의융합대학 종단연구 설문")

    st.write(
        f"**현재 단계:** {step + 1} / {len(step_labels)} — {step_labels[step]}"
    )

    st.progress((step + 1) / len(step_labels))

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
