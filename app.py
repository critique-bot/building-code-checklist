import streamlit as st
import pandas as pd

st.set_page_config(page_title="건축 인허가 법규검토 체크리스트", layout="wide")
st.title("건축 인허가 법규검토 체크리스트")

FACILITY_SHEET = "개별용도 시설기준 (26개 시설, 146개)"


@st.cache_data
def load_data():
    return pd.read_excel("regulations_full.xlsx", sheet_name=None)


sheets = load_data()
facility_df = sheets[FACILITY_SHEET]
facility_list = sorted(facility_df["시설 구분"].unique().tolist())
common_categories = [name for name in sheets.keys() if name != FACILITY_SHEET]


def render_table(df):
    show_cols = [c for c in ["연번", "법명", "조항번호", "내용"] if c in df.columns]
    for _, row in df.iterrows():
        cols = st.columns([1, 3, 2, 4, 2])
        for i, c in enumerate(show_cols):
            cols[i].write(row[c])
        if "링크" in df.columns and pd.notna(row["링크"]):
            cols[4].markdown(f"[원문보기]({row['링크']})")


# ---------------- 사이드바: 설계 개요 입력 ----------------
st.sidebar.header("설계 개요 입력")
selected_uses = st.sidebar.multiselect("건물 용도(시설 구분) 선택", facility_list)
show_common = st.sidebar.checkbox("공통 체크리스트(용도 무관) 표시", value=True)

# ---------------- ① 용도별 개별 시설기준 ----------------
st.header("① 용도별 개별 시설기준")

filtered = pd.DataFrame()
if selected_uses:
    filtered = facility_df[facility_df["시설 구분"].isin(selected_uses)]
    st.write(f"선택한 용도({', '.join(selected_uses)})에 해당하는 법령 **{len(filtered)}건**")
    render_table(filtered)
else:
    st.info("왼쪽에서 건물 용도를 선택하면 해당 용도의 개별 시설기준 법령이 표시됩니다.")

# ---------------- ② 공통 체크리스트 ----------------
if show_common:
    st.header("② 공통 체크리스트 (모든 프로젝트 공통 확인 필요)")
    for cat in common_categories:
        with st.expander(cat):
            render_table(sheets[cat])

# ---------------- 결과 다운로드 ----------------
st.markdown("---")
if st.button("현재 체크리스트 엑셀로 저장"):
    combined = []
    if not filtered.empty:
        tmp = filtered.copy()
        tmp["구분"] = "개별용도 시설기준"
        combined.append(tmp)
    for cat in common_categories:
        tmp = sheets[cat].copy()
        tmp["구분"] = cat
        combined.append(tmp)
    result_df = pd.concat(combined, ignore_index=True)
    output_path = "검토결과.xlsx"
    result_df.to_excel(output_path, index=False)
    with open(output_path, "rb") as f:
        st.download_button("다운로드", f, file_name="검토결과.xlsx")
