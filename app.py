import streamlit as st
import pandas as pd

st.set_page_config(page_title="한국건축규정 체크리스트", layout="wide")
st.title("한국건축규정 체크리스트")

FACILITY_SHEET = "개별용도 시설기준 (26개 시설, 146개)"

# ---------------- 표준 분류 목록 ----------------
ZONE_TYPES = [
    # 대분류(4개)
    "도시지역", "관리지역", "농림지역", "자연환경보전지역",
    # 중분류(9개) 중 대분류에 없는 것들
    "주거지역", "상업지역", "공업지역", "녹지지역",
    "보전관리지역", "생산관리지역", "계획관리지역",
    # 소분류(21개)
    "제1종전용주거지역", "제2종전용주거지역", "제1종일반주거지역", "제2종일반주거지역", "제3종일반주거지역", "준주거지역",
    "중심상업지역", "일반상업지역", "근린상업지역", "유통상업지역",
    "전용공업지역", "일반공업지역", "준공업지역",
    "보전녹지지역", "생산녹지지역", "자연녹지지역",
    # 개별법 전용 추가 항목
    "역사문화환경 보존지역",
]

ZONE_DISTRICTS = [
    "자연경관지구", "시가지경관지구", "특화경관지구",
    "고도지구", "방화지구",
    "시가지방재지구", "자연방재지구",
    "역사문화환경보호지구", "중요시설물보호지구", "생태계보호지구",
    "자연취락지구", "집단취락지구",
    "주거개발진흥지구", "산업유통개발진흥지구", "관광휴양개발진흥지구", "복합개발진흥지구", "특정개발진흥지구",
    "특정용도제한지구", "복합용도지구",
]

ZONE_AREAS = [
    "개발제한구역", "도시자연공원구역", "시가화조정구역", "수산자원보호구역",
    "도시혁신구역", "복합용도구역", "도시·군계획시설입체복합구역",
    # 개별법 전용 추가 항목
    "상대보호구역", "절대보호구역", "지구단위계획구역", "접도구역", "항만구역", "도로구역",
    "산림보호구역", "상수원보호구역",
    "국가지정문화유산구역",
    "과밀억제권역", "성장관리권역", "자연보전권역",
    "습지보호지역", "습지주변관리지역", "역사문화권정비구역", "연구개발특구", "시·도 생태·경관보전지역",
    "무선방위측정장치보호구역", "철도보호지구", "하천구역", "택지개발지구", "혁신도시개발예정지구", "특별대책지역",
]

# 용도지역·지구·구역을 하나의 선택목록으로 통합
ZONE_ALL = ZONE_TYPES + ZONE_DISTRICTS + ZONE_AREAS

# 건물 용도 화면 표시 순서 (요청하신 순서 그대로)
FACILITY_DISPLAY_ORDER = [
    "단독주택", "공동주택", "제1종 근린생활시설", "제2종 근린생활시설",
    "문화 및 집회시설", "종교시설", "판매시설", "운수시설", "의료시설",
    "교육연구시설", "노유자시설", "수련시설", "운동시설", "업무시설",
    "숙박시설", "위락시설", "공장", "창고시설", "위험물 저장 및 처리 시설",
    "자동차 관련 시설", "동물 및 식물 관련 시설", "자원순환 관련 시설",
    "교정시설", "국방·군사시설", "방송통신시설", "발전시설",
    "묘지 관련 시설", "관광 휴게시설", "장례시설", "야영장 시설",
]

# 화면 표시명 -> 실제 데이터("시설 구분") 값 매핑
FACILITY_DISPLAY_TO_DATA = {
    "단독주택": "주택",
    "공동주택": "주택",
    "제1종 근린생활시설": "근린생활시설",
    "제2종 근린생활시설": "근린생활시설",
    "자동차 관련 시설": "자동차관련시설",
    "교정시설": "교정 및 군사 시설",
    "국방·군사시설": "교정 및 군사 시설",
    "장례시설": "장례식장",
}

JIMOK_LIST = [
    "전", "답", "과수원", "목장용지", "임야", "광천지", "염전", "대", "공장용지", "학교용지",
    "주차장", "주유소용지", "창고용지", "도로", "철도용지", "제방", "하천", "구거", "유지",
    "양어장", "수도용지", "공원", "체육용지", "유원지", "종교용지", "사적지", "묘지", "잡종지",
]


@st.cache_data
def load_data():
    sheets = pd.read_excel("regulations_full.xlsx", sheet_name=None)
    # 시트마다 열 구성이 조금씩 달라도(예: 조항번호/내용 없는 시트) 화면 정렬이
    # 어긋나지 않도록, 표시에 쓰는 표준 열이 없으면 빈 값으로 채워둔다.
    for name, df in sheets.items():
        for col in ["조항번호", "내용"]:
            if col not in df.columns:
                df[col] = ""
        sheets[name] = df
    return sheets


sheets = load_data()
facility_df = sheets[FACILITY_SHEET]
common_categories = [name for name in sheets.keys() if name != FACILITY_SHEET]


# ---------------- 사이드바: 설계 개요 입력 ----------------
st.sidebar.header("설계 개요 입력")

selected_uses_display = st.sidebar.multiselect("건물 용도(시설 구분)", FACILITY_DISPLAY_ORDER)
selected_uses = list({FACILITY_DISPLAY_TO_DATA.get(u, u) for u in selected_uses_display})

address_input = st.sidebar.text_input("대지 주소", placeholder="예: 서울특별시 종로구 삼일대로 451")


def classify_eup_myeon(address: str) -> str:
    """주소 문자열에서 읍/면 여부를 판별. 읍 또는 면으로 끝나는 행정동 단위가 있으면 '대상', 없으면(동 등) '비대상'."""
    if not address:
        return "비대상"
    tokens = address.split()
    for t in tokens:
        if t.endswith("읍"):
            return "대상"
    for t in tokens:
        if t.endswith("면"):
            return "대상"
    return "비대상"


is_eup_myeon = classify_eup_myeon(address_input)
if address_input:
    st.sidebar.caption(f"읍 · 면 판별 결과: **{is_eup_myeon}**")

selected_jimok = st.sidebar.multiselect("지목", JIMOK_LIST)

selected_zones = st.sidebar.multiselect("용도지역 · 지구 · 구역", ZONE_ALL)

is_urban_planning_facility = st.sidebar.radio("도시 · 군계획시설", ["비대상", "대상"], horizontal=True)

site_area = st.sidebar.number_input("대지면적(㎡)", min_value=0.0, value=0.0, step=10.0)
building_area = st.sidebar.number_input("건축면적(㎡)", min_value=0.0, value=0.0, step=10.0)
total_floor_area = st.sidebar.number_input("연면적(㎡)", min_value=0.0, value=0.0, step=10.0)
height = st.sidebar.number_input("높이(m)", min_value=0.0, value=0.0, step=1.0)
num_floors_above = st.sidebar.number_input("층수(지상)", min_value=0, value=0, step=1)
num_floors_below = st.sidebar.number_input("층수(지하)", min_value=0, value=0, step=1)

# 숫자형 입력값
numeric_values = {
    "대지면적": site_area,
    "건축면적": building_area,
    "연면적": total_floor_area,
    "높이": height,
    "지상층수": num_floors_above,
    "지하층수": num_floors_below,
}

# 다중선택(리스트)형 입력값
list_values = {
    "용도지역지구구역": selected_zones,
    "지목": selected_jimok,
    "건물용도": selected_uses,
    "건물용도표시": selected_uses_display,
}

# Y/N형 입력값
yn_values = {
    "도시군계획시설": is_urban_planning_facility,
    "읍면": is_eup_myeon,
}

show_common = st.sidebar.checkbox("공통 체크리스트(용도 무관) 표시", value=True)

st.sidebar.markdown("---")
st.sidebar.caption(
    "기준항목/기준값이 입력된 조항만 자동판별됩니다. "
    "비어있는 조항은 확인필요로 표시되며 원문 확인이 필요합니다."
)


# ---------------- 판별 로직 ----------------
def compare_numeric(value, op, threshold):
    try:
        threshold = float(threshold)
    except (TypeError, ValueError):
        return None
    if op == ">=":
        return value >= threshold
    elif op == ">":
        return value > threshold
    elif op == "<=":
        return value <= threshold
    elif op == "<":
        return value < threshold
    elif op == "eq":
        return value == threshold
    return None


def eval_condition(item, op, threshold):
    """조건 하나(기준항목/연산자/기준값)를 평가. 해당 없으면 None 반환"""
    if not isinstance(item, str) or item.strip() == "":
        return None
    item = item.strip()

    if item in numeric_values:
        return compare_numeric(numeric_values[item], op, threshold)

    if item in yn_values:
        if op == "eq":
            return yn_values[item] == threshold
        return None

    if item in list_values:
        selected = list_values[item]
        if op == "is_empty":
            return len(selected) == 0
        if op == "count_gte":
            try:
                return len(selected) >= float(threshold)
            except (TypeError, ValueError):
                return None
        if op == "count_lte":
            try:
                return len(selected) <= float(threshold)
            except (TypeError, ValueError):
                return None
        if not isinstance(threshold, str):
            return None
        candidates = [t.strip() for t in threshold.split(",")]
        if op == "contains_any":
            return any(c in selected for c in candidates)
        elif op == "contains_all":
            return all(c in selected for c in candidates)
    return None


def judge_row(row):
    # 기본판정이 채워져 있으면 조건 계산 없이 그대로 사용
    base = row.get("기본판정")
    if isinstance(base, str) and base.strip() != "":
        return base.strip()

    cond1 = eval_condition(row.get("기준항목1"), row.get("연산자1"), row.get("기준값1"))
    cond2 = eval_condition(row.get("기준항목2"), row.get("연산자2"), row.get("기준값2"))

    has_any_condition_defined = isinstance(row.get("기준항목1"), str) and row.get("기준항목1").strip() != ""
    if not has_any_condition_defined:
        return "검토필요"
    if cond1 is None:
        return "검토필요"

    conds = [c for c in [cond1, cond2] if c is not None]
    return "검토필요" if all(conds) else "대상아님"


def render_table(df):
    df = df.copy()
    df["판정"] = df.apply(judge_row, axis=1)

    color_map = {"검토필요": "🟢 검토필요", "대상아님": "🔴 대상아님"}
    df["판정_표시"] = df["판정"].map(color_map).fillna(df["판정"])

    show_cols = [c for c in ["연번", "법명", "조항번호", "내용", "판정_표시"] if c in df.columns]
    for _, row in df.iterrows():
        cols = st.columns([1, 3, 2, 4, 1.3, 2])
        for i, c in enumerate(show_cols):
            cols[i].write(row[c])
        if "링크" in df.columns and pd.notna(row["링크"]):
            cols[len(show_cols)].markdown(f"[원문보기]({row['링크']})")
    return df


CIRCLED_NUMS = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]

# ---------------- ① 개별용도 시설기준 ----------------
st.header(f"{CIRCLED_NUMS[0]} {FACILITY_SHEET}")

filtered = pd.DataFrame()
if selected_uses:
    filtered = facility_df[facility_df["시설 구분"].isin(selected_uses)]
    st.write(f"선택한 용도({', '.join(selected_uses)})에 해당하는 법령 **{len(filtered)}건**")
    filtered = render_table(filtered)
else:
    st.info("왼쪽에서 건물 용도를 선택하면 해당 용도의 개별 시설기준 법령이 표시됩니다.")

# ---------------- ②~ 공통 체크리스트 (자동판별 포함) ----------------
result_frames = []
if show_common:
    for idx, cat in enumerate(common_categories, start=1):
        st.header(f"{CIRCLED_NUMS[idx]} {cat}")
        with st.expander("목록 펼치기 / 접기", expanded=True):
            result_df = render_table(sheets[cat])
            result_frames.append(result_df)

# ---------------- 결과 다운로드 ----------------
st.markdown("---")
if st.button("현재 체크리스트 엑셀로 저장"):
    combined = []
    if not filtered.empty:
        tmp = filtered.copy()
        tmp["구분"] = "개별용도 시설기준"
        combined.append(tmp)
    if show_common:
        for cat, rdf in zip(common_categories, result_frames):
            tmp = rdf.copy()
            tmp["구분"] = cat
            combined.append(tmp)
    result_df = pd.concat(combined, ignore_index=True)
    output_path = "검토결과.xlsx"
    result_df.to_excel(output_path, index=False)
    with open(output_path, "rb") as f:
        st.download_button("다운로드", f, file_name="검토결과.xlsx")
