import streamlit as st
import pandas as pd

st.title("건축법규 자동검토 프로그램")

reg_df = pd.read_excel("regulations.xlsx")

# 입력 UI
building_height = st.number_input("건물 높이(m)", min_value=0, value=22)
total_floor_area = st.number_input("연면적(㎡)", min_value=0, value=1200)
use_options = reg_df["용도"].unique().tolist()
building_use = st.multiselect("건물 용도 선택", use_options)

if st.button("검토 실행"):
    results = []
    for index, row in reg_df.iterrows():
        use_match = (row["용도"] in building_use) or (row["용도"] == "전체")
        
        if row["기준항목"] == "연면적":
            value_match = total_floor_area >= row["기준값"]
        elif row["기준항목"] == "높이":
            value_match = building_height >= row["기준값"]
        else:
            value_match = False
        
        status = "대상" if (use_match and value_match) else "제외"
        results.append({"조항": row["조항"], "내용": row["내용"], "결과": status})
    
    result_df = pd.DataFrame(results)
    st.write("=== 검토 결과 ===")
    st.dataframe(result_df)