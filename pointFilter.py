import streamlit as st
import pandas as pd
from datetime import datetime
import pandas.tseries.offsets as offsets

st.set_page_config(
    page_title="복지포인트 지급 대상 자동화 프로그램",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 사용자 정의 CSS
st.markdown("""
<style>
/* 전체 문서에 대해 폰트 크기, 줄 간격, 폰트 패밀리 지정 */
html, body, [class*="css"]  {
    font-family: "Malgun Gothic", "맑은 고딕", sans-serif; /* Windows 환경에서 한글 폰트 */
    font-size: 15px;           /* 전체 글자 크기 조정 */
    line-height: 1.4;          /* 줄 간격 */
    word-wrap: break-word;     /* 길이가 긴 단어도 자동 줄바꿈 */
}

/* h1 태그: st.title()에 해당 */
h1 {
    font-size: 1.8rem !important;  /* 타이틀 폰트 크기 */
    line-height: 1.2 !important;   /* 타이틀 줄 간격 */
    margin-bottom: 0.5rem;         /* 타이틀 아래쪽 여백 */
}
</style>
""", unsafe_allow_html=True)

st.title("복지포인트 지급 대상 자동화 프로그램")
st.write("""
이 프로그램은 업로드된 Excel 파일을 기반으로 **복지포인트 지급 대상 여부**를 자동으로 판별합니다.  
좌측 사이드바에서 **지급 기준일**을 선택하고, **Excel 파일**을 업로드하세요.  
조건에 따라 **'대상'**, **'확인필요'** 행만 남기고 CSV 파일로 다운로드할 수 있습니다.
""")

# 4) 사이드바
st.sidebar.header("설정")
pay_date_input = st.sidebar.date_input("지급 기준일을 선택하세요", datetime(2025, 4, 3))
pay_date = datetime.combine(pay_date_input, datetime.min.time())
uploaded_file = st.sidebar.file_uploader("Excel 파일 업로드 (.xlsx)", type=["xlsx"])

def process_data(df, pay_date):
    df["퇴직일"] = pd.to_datetime(df["퇴직일"], errors="coerce")
    df["입사일"] = pd.to_datetime(df["입사일"], errors="coerce")

    name_counts = df["영문이름"].value_counts()

    def check_row(row):
        # 1. 재직상태 "퇴직"이면 제외
        if row["재직상태"] == "퇴직":
            return "제외"
        # 2. 지급 기준일보다 퇴직일이 이전이면 "확인필요"
        if pd.notnull(row["퇴직일"]) and row["퇴직일"] < pay_date:
            return "확인필요"
        
        # 기타 제외 조건
        cond_name = row["회사 내 이름"] in ["Jake.Kim", "Jae.Kim", "TEST HR"]
        cond_intern = row["직위"] == "Intern"
        cond_duplicate = name_counts.get(row["회사 내 이름"], 0) > 1
        
        join_threshold = pay_date - offsets.DateOffset(months=3)
        cond_join_exclude = pd.isnull(row["입사일"]) or (row["입사일"] >= join_threshold)

        if cond_name or cond_intern or cond_duplicate or cond_join_exclude:
            return "제외"
        return "대상"

    df["지급대상여부"] = df.apply(check_row, axis=1)

    # 제외 대상 행은 제거, 대상/확인필요만 남김
    df_filtered = df[df["지급대상여부"] != "제외"].copy()

    cols_to_keep = ["사번", "이름", "회사 내 이름", "입사일", "지급대상여부"]
    df_final = df_filtered[cols_to_keep]
    return df_final

# 5) 메인 컨텐츠
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        
        st.subheader("업로드 된 원본 데이터 (미리보기)")
        st.dataframe(df.head(10))
        st.write("컬럼 헤더:", df.columns.tolist())
        
        df_result = process_data(df, pay_date=pay_date)
        
        st.markdown("---")
        st.subheader("필터링 후 최종 데이터")
        st.dataframe(df_result)
        
        # CSV 다운로드
        csv_data = df_result.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="CSV 파일 다운로드",
            data=csv_data,
            file_name="filtered_output.csv",
            mime="text/csv"
        )
    except Exception as e:
        st.error(f"오류가 발생하였습니다: {e}")
else:
    st.info("좌측 사이드바에서 지급 기준일을 설정하고, Excel 파일을 업로드하세요.")


#
# import streamlit as st
# import pandas as pd
# from datetime import datetime
# import pandas.tseries.offsets as offsets

# # 1) 페이지 설정 (와이드 레이아웃 사용)
# st.set_page_config(
#     page_title="복지포인트 지급 대상 자동화 프로그램",
#     layout="wide",  # wide 모드
#     initial_sidebar_state="expanded"
# )

# # 2) 사용자 정의 CSS
# st.markdown("""
#     <style>
#     /* 메인 영역의 최대 너비를 제한하여 가운데 정렬 효과 */
#     .main > div {
#         max-width: 900px; /* 원하는 최대 너비 */
#         margin: 0 auto;
#         padding: 1rem;
#         background-color: #ffffff; /* 배경색을 흰색으로 */
#         border-radius: 8px;
#         box-shadow: 0 2px 5px rgba(0,0,0,0.1);
#     }
    
#     /* 제목 스타일 */
#     .centered-title {
#         text-align: center;
#         color: #2c3e50;
#         font-size: 2.5rem;
#         font-weight: 700;
#         margin-top: 1rem;
#         margin-bottom: 0.5rem;
#     }
    
#     /* 설명(Description) 스타일 */
#     .centered-description {
#         text-align: center;
#         color: #34495e;
#         font-size: 1.1rem;
#         line-height: 1.6;
#         margin: 0 auto 1.5rem auto;
#         max-width: 700px; /* 설명 문단의 최대 너비 */
#     }
    
#     /* 사이드바 스타일 */
#     .sidebar .sidebar-content {
#         background-color: #f1f4f8;
#         padding: 1rem;
#     }
#     </style>
# """, unsafe_allow_html=True)

# # 3) 메인 레이아웃: 중앙 정렬된 제목과 설명
# st.markdown("<h1 class='centered-title'>복지포인트 지급 대상 자동화 프로그램</h1>", unsafe_allow_html=True)
# st.markdown("""
# <p class='centered-description'>
# 이 프로그램은 업로드된 Excel 파일을 기반으로 복지포인트 지급 대상 여부를 자동으로 판별합니다. 
# 사용자는 사이드바에서 <strong>지급 기준일</strong>을 선택하고, <strong>Excel 파일</strong>을 업로드하세요.
# 이후 조건에 따라 '대상' 또는 '확인필요'인 행만 남기고 CSV 파일로 다운로드할 수 있습니다.
# </p>
# """, unsafe_allow_html=True)

# # 4) 사이드바 구성
# st.sidebar.title("설정")
# pay_date_input = st.sidebar.date_input("지급 기준일을 선택하세요", datetime(2025, 4, 3))
# pay_date = datetime.combine(pay_date_input, datetime.min.time())
# uploaded_file = st.sidebar.file_uploader("Excel 파일 업로드 (.xlsx)", type=["xlsx"])

# def process_data(df, pay_date):
#     # "퇴직일"과 "입사일" 열을 날짜형으로 변환
#     df["퇴직일"] = pd.to_datetime(df["퇴직일"], errors="coerce")
#     df["입사일"] = pd.to_datetime(df["입사일"], errors="coerce")
    
#     # 지급 기준일 이전 퇴직자 조건
#     condition_date = df["퇴직일"].notnull() & (df["퇴직일"] < pay_date)
    
#     # "영문이름" 열의 빈도 계산
#     name_counts = df["영문이름"].value_counts()
    
#     def check_row(row):
#         # 1. 재직상태가 "퇴직"이면 바로 제외
#         if row["재직상태"] == "퇴직":
#             return "제외"
        
#         # 2. 퇴직일이 지급 기준일보다 이전이면 "확인필요"
#         if pd.notnull(row["퇴직일"]) and row["퇴직일"] < pay_date:
#             return "확인필요"
        
#         # 3. 기타 제외 조건
#         cond_name = row["회사 내 이름"] in ["Jake.Kim", "Jae.Kim", "TEST HR"]
#         cond_intern = row["직위"] == "Intern"
#         cond_duplicate = name_counts.get(row["회사 내 이름"], 0) > 1
        
#         # 지급 기준일로부터 3개월 전
#         join_threshold = pay_date - offsets.DateOffset(months=3)
#         cond_join_exclude = pd.isnull(row["입사일"]) or (row["입사일"] >= join_threshold)
        
#         if cond_name or cond_intern or cond_duplicate or cond_join_exclude:
#             return "제외"
        
#         return "대상"
    
#     df["지급대상여부"] = df.apply(check_row, axis=1)
    
#     # "제외"는 제거, "대상"과 "확인필요"만 남김
#     df_filtered = df[df["지급대상여부"] != "제외"].copy()
    
#     # 최종 결과
#     cols_to_keep = ["사번", "이름", "회사 내 이름", "입사일", "지급대상여부"]
#     df_final = df_filtered[cols_to_keep]
#     return df_final

# # 5) 메인 컨텐츠
# if uploaded_file:
#     try:
#         df = pd.read_excel(uploaded_file)
        
#         st.subheader("업로드 된 원본 데이터 (미리보기)")
#         st.dataframe(df.head(10))  # 처음 10행 미리보기
#         st.write("컬럼 헤더:", df.columns.tolist())
        
#         df_result = process_data(df, pay_date=pay_date)
        
#         st.markdown("---")
#         st.subheader("필터링 후 최종 데이터")
#         st.dataframe(df_result)
        
#         # CSV 다운로드
#         csv_data = df_result.to_csv(index=False).encode("utf-8")
#         st.download_button(
#             label="CSV 파일 다운로드",
#             data=csv_data,
#             file_name="filtered_output.csv",
#             mime="text/csv"
#         )
#     except Exception as e:
#         st.error(f"오류가 발생하였습니다: {e}")
# else:
#     st.info("좌측 사이드바에서 지급 기준일을 설정하고, Excel 파일을 업로드하세요.")



# Ver1.0
# import streamlit as st
# import pandas as pd
# from datetime import datetime
# import pandas.tseries.offsets as offsets

# def process_data(df, pay_date):
#     # 1. "퇴직일"과 "입사일" 열을 날짜형으로 변환 (오류 발생 시 NaT 처리)
#     df["퇴직일"] = pd.to_datetime(df["퇴직일"], errors="coerce")
#     df["입사일"] = pd.to_datetime(df["입사일"], errors="coerce")
    
#     # 2. COUNTIF($F:$F, C2) 조건을 위해 "영문이름" 열의 값 빈도를 계산합니다.
#     name_counts = df["영문이름"].value_counts()
    
#     # 3. 각 행을 검사하는 함수 (조건 순서를 재배치)
#     def check_row(row):
#         # (1) 재직상태가 "퇴직"이면 바로 "제외" (최우선)
#         if row["재직상태"] == "퇴직":
#             return "제외"
#         # (2) 퇴직일이 존재하고, 지급 기준일(pay_date)보다 이전이면 "확인필요"
#         if pd.notnull(row["퇴직일"]) and row["퇴직일"] < pay_date:
#             return "확인필요"
#         # (3) 그 외의 조건들을 검사
#         # 회사 내 이름이 "Jake.Kim" 또는 "Jae.Kim"
#         cond_name = row["회사 내 이름"] in ["Jake.Kim", "Jae.Kim", "TEST HR"]
#         # 직위가 "Intern"
#         cond_intern = row["직위"] == "Intern"
#         # 회사 내 이름의 중복 여부: "영문이름" 열에서 해당 값의 등장 횟수가 1보다 크면
#         cond_duplicate = name_counts.get(row["회사 내 이름"], 0) > 1
#         # 지급 기준일로부터 3개월 내에 입사한 경우는 제외 (즉, 입사일이 지급 기준일 - 3개월 이후이거나 누락)
#         join_threshold = pay_date - offsets.DateOffset(months=3)
#         cond_join_exclude = pd.isnull(row["입사일"]) or (row["입사일"] >= join_threshold)
        
#         if cond_name or cond_intern or cond_duplicate or cond_join_exclude:
#             return "제외"
#         else:
#             return "대상"
    
#     # 4. 각 행에 대해 지급대상여부를 결정
#     df["지급대상여부"] = df.apply(check_row, axis=1)
    
#     # 5. 지급대상여부가 "제외"인 행은 제거 (즉, "대상"과 "확인필요"인 행만 남김)
#     df_filtered = df[df["지급대상여부"] != "제외"].copy()
    
#     # 6. 최종 결과: "사번", "이름", "회사 내 이름", "입사일"과 지급대상여부(마지막 열)만 남김
#     cols_to_keep = ["사번", "이름", "회사 내 이름", "입사일", "지급대상여부"]
#     df_final = df_filtered[cols_to_keep]
    
#     return df_final

# def main():
#     st.title("복지포인트 지급 대상 자동화")
    
#     # 지급 기준일을 사용자가 직접 입력 (기본값: 2025-04-03)
#     pay_date_input = st.date_input("지급 기준일을 선택하세요", datetime(2025, 4, 3))
#     pay_date = datetime.combine(pay_date_input, datetime.min.time())
#     st.write("선택된 지급 기준일:", pay_date.strftime("%Y-%m-%d"))
    
#     st.write("Excel 파일을 업로드하면, 지급 기준 조건에 따라 지급 대상('대상' 또는 '확인필요')인 행만 남긴 CSV 파일로 처리됩니다. " \
#              "출력 결과의 마지막 열에 '지급대상여부'가 표기됩니다.")
    
#     uploaded_file = st.file_uploader("Excel 파일 업로드 (.xlsx)", type=["xlsx"])
    
#     if uploaded_file is not None:
#         try:
#             df = pd.read_excel(uploaded_file)
#             st.subheader("업로드 된 원본 데이터 (미리보기)")
#             st.dataframe(df.head())
#             st.write("컬럼 헤더:", df.columns.tolist())
            
#             df_result = process_data(df, pay_date=pay_date)
            
#             st.subheader("필터링 후 최종 데이터 (사번, 이름, 회사 내 이름, 입사일, 지급대상여부)")
#             st.dataframe(df_result)
            
#             csv_data = df_result.to_csv(index=False).encode("utf-8")
#             st.download_button(
#                 label="CSV 파일 다운로드",
#                 data=csv_data,
#                 file_name="filtered_output.csv",
#                 mime="text/csv"
#             )
#         except Exception as e:
#             st.error(f"오류가 발생하였습니다: {e}")

# if __name__ == '__main__':
#     main()
