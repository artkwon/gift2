import pandas as pd
import streamlit as st
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="광고 수익 계산기", layout="wide")
st.title("📦 광고 마진 계산기")

uploaded_file = st.file_uploader("👉 분석할 엑셀 파일을 업로드하세요", type=["xlsx"])
st.caption("※ .xlsx 형식의 광고 데이터 파일만 업로드 가능합니다.")

def convert_df_to_excel(result_df, summary_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name='광고유형별 계산 결과')
        summary_df.to_excel(writer, index=False, sheet_name='전체 수익 요약')
    output.seek(0)
    return output

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        required_cols = ['광고집행 옵션ID', '광고집행 상품명', '광고유형', '광고비', '총 판매수량(14일)']
        for col in required_cols:
            if col not in df.columns:
                st.error(f"엑셀에 '{col}' 열이 없습니다. 현재 열 목록: {df.columns.tolist()}")
                st.stop()

        df['상품키'] = df['광고집행 옵션ID'].astype(str) + '_' + df['광고집행 상품명']
        grouped_inputs = df[['상품키', '광고집행 옵션ID', '광고집행 상품명']].drop_duplicates()

        st.subheader("🧾 기본 정보 입력")
        user_inputs = {}

        for idx, row in grouped_inputs.iterrows():
            상품키 = row['상품키']
            옵션ID = row['광고집행 옵션ID']
            상품명 = row['광고집행 상품명']

            with st.container():
                st.markdown(f"**옵션ID:** `{옵션ID}` &nbsp;&nbsp;&nbsp; **상품명:** {상품명}")
                cols = st.columns([1, 1, 1, 1, 1])
                총판매량 = cols[0].number_input("총판매량", key=f"총판매량_{상품키}", min_value=0, step=1, format="%d")
                판매가 = cols[1].number_input("판매가", key=f"판매가_{상품키}", min_value=0, step=100, format="%d")
                원가 = cols[2].number_input("원가", key=f"원가_{상품키}", min_value=0, step=100, format="%d")
                수수료율 = cols[3].number_input("수수료율(%)", key=f"수수료율_{상품키}", min_value=0, step=1, format="%d")
                배송비 = cols[4].number_input("배송비", key=f"배송비_{상품키}", min_value=0, step=500, format="%d")
                user_inputs[상품키] = (총판매량, 판매가, 원가, 수수료율, 배송비)

        st.divider()

        grouped = df.groupby(['상품키', '광고유형', '광고집행 옵션ID', '광고집행 상품명'], as_index=False).agg({
            '총 판매수량(14일)': 'sum',
            '광고비': 'sum'
        })

        광고수익_map = defaultdict(float)
        광고비합_map = defaultdict(float)
        기타판매_map = defaultdict(int)
        단가마진_map = {}
        전체판매량_map = {}

        result_rows = []

        for _, row in grouped.iterrows():
            상품키 = row['상품키']
            광고유형 = row['광고유형']
            옵션ID = row['광고집행 옵션ID']
            상품명 = row['광고집행 상품명']
            광고비 = row['광고비']
            판매14 = row['총 판매수량(14일)']

            if 상품키 not in user_inputs or any(v == 0 for v in user_inputs[상품키]):
                광고수익 = '-'
                평균광고비 = '-'
            else:
                총판매량, 판매가, 원가, 수수료율, 배송비 = user_inputs[상품키]
                try:
                    단가마진 = 판매가 - 원가 - (판매가 * 수수료율 / 100) - 배송비
                    광고수익 = (단가마진 * 판매14) - 광고비
                    기타판매 = max(0, 총판매량 - 판매14)
                    평균광고비 = 광고비 / 판매14 if 판매14 else 0
                except:
                    광고수익 = 0
                    평균광고비 = 0

                광고수익_map[상품키] += 광고수익
                광고비합_map[상품키] += 광고비
                기타판매_map[상품키] = 기타판매
                단가마진_map[상품키] = 단가마진
                전체판매량_map[상품키] = 총판매량

            result_rows.append({
                "광고유형": 광고유형,
                "옵션ID": 옵션ID,
                "상품명": 상품명,
                "총 판매수량": 판매14,
                "광고비": 광고비,
                "평균 광고비": round(평균광고비) if isinstance(평균광고비, (int, float)) else 평균광고비,
                "광고수익": 광고수익
            })

        merged_result = {}
        merged_summary = []

        for row in result_rows:
            상품키 = f"{row['옵션ID']}_{row['상품명']}"
            if 상품키 in merged_result:
                continue
            if 상품키 not in 단가마진_map or any(v == 0 for v in user_inputs.get(상품키, (0,0,0,0,0))):
                전체수익 = '-'
                개당마진 = '-'
            else:
                전체수익 = 광고수익_map[상품키] + (단가마진_map[상품키] * 기타판매_map[상품키])
                개당마진 = 전체수익 / 전체판매량_map[상품키] if 전체판매량_map[상품키] else 0

            merged_result[상품키] = {
                "옵션ID": row['옵션ID'],
                "상품명": row['상품명'],
                "전체수익": 전체수익,
                "개당마진": round(개당마진) if isinstance(개당마진, (int, float)) else 개당마진
            }
            merged_summary.append({
                "옵션ID": row['옵션ID'],
                "상품명": row['상품명'],
                "총 판매량": 전체판매량_map.get(상품키, '-'),
                "총 광고비": 광고비합_map.get(상품키, '-'),
                "개당수익": round(개당마진) if isinstance(개당마진, (int, float)) else 개당마진,
                "총 수익": 전체수익
            })

        result_df = pd.DataFrame(result_rows)
        result_df['총 판매수량'] = result_df['총 판매수량'].apply(lambda x: f"{x:,}" if isinstance(x, (int, float)) else x)
        result_df['광고비'] = result_df['광고비'].apply(lambda x: f"{int(x):,}" if isinstance(x, (int, float)) else x)
        result_df['평균 광고비'] = result_df['평균 광고비'].apply(lambda x: f"{x:,}" if isinstance(x, (int, float)) else x)
        result_df['광고수익'] = result_df['광고수익'].apply(lambda x: f"<span style='color:red;font-weight:bold;'>{int(x):,}</span>" if isinstance(x, (int, float)) else '-')

        summary_df = pd.DataFrame(merged_summary)
        summary_df['총 판매량'] = summary_df['총 판매량'].apply(lambda x: f"{x:,}" if isinstance(x, (int, float)) else x)
        summary_df['총 광고비'] = summary_df['총 광고비'].apply(lambda x: f"{int(x):,}" if isinstance(x, (int, float)) else x)
        summary_df['개당수익'] = summary_df['개당수익'].apply(lambda x: f"{x:,}" if isinstance(x, (int, float)) else x)
        summary_df['총 수익'] = summary_df['총 수익'].apply(lambda x: f"<span style='color:red;font-weight:bold;'>{int(x):,}</span>" if isinstance(x, (int, float)) else '-')

        st.markdown("### 광고유형별 계산 결과", unsafe_allow_html=True)
        st.markdown("<style>thead th, tbody td {text-align: center !important;} td span {font-family: inherit;}</style>", unsafe_allow_html=True)
        st.write(result_df.to_html(escape=False, index=False), unsafe_allow_html=True)

        st.subheader("📈 전체 수익 요약")
        st.write(summary_df.to_html(escape=False, index=False), unsafe_allow_html=True)

        excel_data = convert_df_to_excel(pd.DataFrame(result_rows), pd.DataFrame(merged_summary))
        st.download_button(
            label="📥 결과 엑셀로 다운로드",
            data=excel_data,
            file_name="광고수익_계산결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"파일 처리 중 오류 발생: {e}")
