import streamlit as st
import pandas as pd

st.set_page_config(page_title="건축 인허가 법규검토 체크리스트", layout="wide")
st.title("건축 인허가 법규검토 체크리스트")

FACILITY_SHEET = "개별용도 시설기준 (26개 시설, 146개)"

# ---------------- 표준 분류 목록 ----------------
ZONE_TYPES = [
    "제1종전용주거지역", "제2종전용주거지역", "제1종일반주거지역", "제2종일반주거지역", "제3종일반주거지역", "준주거지역",
    "중심상업지역", "일반상업지역", "근린상업지역", "유통상업지역",
    "전용공업지역", "일반공업지역", "준공업지역",
    "보전녹지지역", "생산녹지지역", "자연녹지지역",
    "보전관리지역", "생산관리지역", "계획관리지역",
    "농림지역", "자연환경보전지역",
]

ZONE_DISTRICTS = [
    "자연경관지구", "시가지경관지구", "특화경관지구",
    "고도지구", "방화지구",
    "시가지방재지구", "자연방재지구",
    "역사문화환경보호지구", "중요시설물보호지구", "생태계보호지구",
    "자연취락지구", "집단취락지구",
    "주거개발진흥지구", "산업유통개발진흥지구", "관광휴양개발진흥지구", "복합개발진흥지구", "특정개발진흥지구",
    "특정용도제한지구", "복합용도지구",
    # 개별법 전용 추가 항목
    "역사문화환경보존지역",
]

ZONE_AREAS = [
    "개발제한구역", "도시자연공원구역", "시가화조정구역", "수산자원보호구역",
    "도시혁신구역", "복합용도구역", "도시·군계획시설입체복합구역",
    # 개별법 전용 추가 항목
    "상대보호구역", "절대보호구역", "지구단위계획구역",
]

# 용도지역·지구·구역을 하나의 선택목록으로 통합
ZONE_ALL = ZONE_TYPES + ZONE_DISTRICTS + ZONE_AREAS

JIMOK_LIST = [
    "전", "답", "과수원", "목장용지", "임야", "광천지", "염전", "대", "공장용지", "학교용지",
    "주차장", "주유소용지", "창고용지", "도로", "철도용지", "제방", "하천", "구거", "유지",
    "양어장", "수도용지", "공원", "체육용지", "유원지", "종교용지", "사적지", "묘지", "잡종지",
]


@st.cache_data
def load_data():
    return pd.read_excel("regulations_full.xlsx", sheet_name=None)


sheets = load_data()
facility_df = sheets[FACILITY_SHEET]
facility_list = sorted(facility_df["시설 구분"].unique().tolist())
common_categories = [name for name in sheets.keys() if name != FACILITY_SHEET]


# ---------------- 사이드바: 설계 개요 입력 ----------------
st.sidebar.header("설계 개요 입력")

selected_uses = st.sidebar.multiselect("건물 용도(시설 구분)", facility_list)

selected_jimok = st.sidebar.multiselect("지목", JIMOK_LIST)

selected_zones = st.sidebar.multiselect("용도지역 · 지구 · 구역", ZONE_ALL)

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
    elif op == "==":
        return value == threshold
    return None


def eval_condition(item, op, threshold):
    """조건 하나(기준항목/연산자/기준값)를 평가. 해당 없으면 None 반환"""
    if not isinstance(item, str) or item.strip() == "":
        return None
    item = item.strip()

    if item in numeric_values:
        return compare_numeric(numeric_values[item], op, threshold)

    if item in list_values:
        selected = list_values[item]
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

    conds = [c for c in [cond1, cond2] if c is not None]
    has_any_condition_defined = isinstance(row.get("기준항목1"), str) and row.get("기준항목1").strip() != ""

    if not has_any_condition_defined:
        return "확인필요"
    if any(c is None for c in [cond1] if row.get("기준항목1")):
        return "확인필요"

    if not conds:
        return "확인필요"

    return "대상" if all(conds) else "제외"


def render_table(df):
    df = df.copy()
    df["판정"] = df.apply(judge_row, axis=1)

    color_map = {"대상": "🔴 대상", "제외": "⚪ 제외", "확인필요": "🟡 확인필요"}
    df["판정_표시"] = df["판정"].map(color_map).fillna(df["판정"])

    show_cols = [c for c in ["연번", "법명", "조항번호", "내용", "판정_표시"] if c in df.columns]
    for _, row in df.iterrows():
        cols = st.columns([1, 3, 2, 4, 1.3, 2])
        for i, c in enumerate(show_cols):
            cols[i].write(row[c])
        if "링크" in df.columns and pd.notna(row["링크"]):
            cols[len(show_cols)].markdown(f"[원문보기]({row['링크']})")
    return df


# ---------------- ① 용도별 개별 시설기준 ----------------
st.header("① 용도별 개별 시설기준")

filtered = pd.DataFrame()
if selected_uses:
    filtered = facility_df[facility_df["시설 구분"].isin(selected_uses)]
    st.write(f"선택한 용도({', '.join(selected_uses)})에 해당하는 법령 **{len(filtered)}건**")
    filtered = render_table(filtered)
else:
    st.info("왼쪽에서 건물 용도를 선택하면 해당 용도의 개별 시설기준 법령이 표시됩니다.")

# ---------------- ② 공통 체크리스트 (자동판별 포함) ----------------
result_frames = []
if show_common:
    st.header("② 공통 체크리스트 (대상 / 제외 / 확인필요 자동판별)")
    for cat in common_categories:
        with st.expander(cat):
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
