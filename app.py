import io
import pandas as pd
import streamlit as st

# Streamlit 페이지 기본 설정
st.set_page_config(
    page_title="주문 데이터 자동 가공 대시보드",
    page_icon="📦",
    layout="wide",
)

st.title("📦 CAFE24 주문 & 상품 데이터 자동 가공 시스템")
st.markdown("왼쪽 사이드바에서 4개의 원본 파일을 모두 업로드한 후 **[데이터 자동 가공 시작]** 버튼을 누르세요.")

st.divider()

# 사이드바 파일 업로드
st.sidebar.header("📁 원본 파일 업로드")

file_latest = st.sidebar.file_uploader("1. 카페24 상품 전체(최신)", type=["xlsx", "xls"])
file_past = st.sidebar.file_uploader("2. 카페24 상품 전체(과거)", type=["xlsx", "xls"])
file_order = st.sidebar.file_uploader("3. CAFE24 주문서 원본", type=["xlsx", "xls"])
file_master = st.sidebar.file_uploader("4. 마스터정보 & 출고제한코드", type=["xlsx", "xls"])

def to_excel_bytes(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# 실행 버튼 클릭 시 동작
if st.sidebar.button("🚀 데이터 자동 가공 시작", type="primary"):
    if not (file_latest and file_past and file_order and file_master):
        st.error("❌ 4개의 원본 엑셀 파일을 모두 업로드해 주세요.")
    else:
        try:
            with st.spinner("데이터를 분석하고 가공하는 중입니다..."):
                df_latest = pd.read_excel(file_latest)
                df_past = pd.read_excel(file_past)
                df_order = pd.read_excel(file_order)

                df_master = pd.read_excel(file_master, sheet_name="마스터정보")
                df_restricted = pd.read_excel(file_master, sheet_name="출고제한 코드")

                # 1. 신규 SKU 추출
                past_codes = set(df_past["상품코드"].dropna().astype(str))
                df_new_sku = df_latest[~df_latest["상품코드"].astype(str).isin(past_codes)].copy()

                # 2. 주문서 + 마스터정보 매칭
                master_cols = ["상품코드", "MD 구분", "현장수령 여부", "이벤트 구분", "발주처 구분"]
                valid_cols = [c for c in master_cols if c in df_master.columns]
                df_order_merged = pd.merge(df_order, df_master[valid_cols], on="상품코드", how="left")

                # 3. 출고제한 코드 묶음 처리
                restricted_codes = set(df_restricted.iloc[:, 0].dropna().astype(str))
                order_restricted_mask = (
                    df_order_merged.groupby("주문번호")["상품코드"]
                    .transform(lambda s: s.fillna("").astype(str).isin(restricted_codes).any())
                    .fillna(False)
                )

                df_unreleased = df_order_merged[order_restricted_mask].copy()
                df_released = df_order_merged[~order_restricted_mask].copy()

                # 4. 배송국가 기준 분리
                country_col = "배송국가" if "배송국가" in df_released.columns else df_released.columns[85]

                df_easyadmin = df_released[df_released[country_col].isna() | (df_released[country_col] == "")].copy()
                df_cj = df_released[df_released[country_col] == "KR"].copy()
                df_fastbox = df_released[
                    (~df_released[country_col].isna())
                    & (df_released[country_col] != "")
                    & (df_released[country_col] != "KR")
                ].copy()

                result_dfs = {
                    "revised_이지어드민": df_easyadmin,
                    "revised_CJ대한통운": df_cj,
                    "revised_패스트박스": df_fastbox,
                    "revised_미출고건": df_unreleased,
                    "revised_마스터 신규SKU": df_new_sku,
                }

                st.success("✅ 가공이 완벽하게 완료되었습니다!")

                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric("이지어드민", f"{len(df_easyadmin):,} 건")
                col2.metric("CJ대한통운", f"{len(df_cj):,} 건")
                col3.metric("패스트박스", f"{len(df_fastbox):,} 건")
                col4.metric("미출고건", f"{len(df_unreleased):,} 건", delta="-제외됨")
                col5.metric("신규 SKU", f"{len(df_new_sku):,} 건")

                st.divider()

                merged_excel = to_excel_bytes(result_dfs)
                st.download_button(
                    label="📥 전체 가공 결과 (통합 엑셀) 다운로드",
                    data=merged_excel,
                    file_name="CAFE24_주문_통합가공결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

                st.subheader("📋 시트별 미리보기 및 개별 다운로드")

                tabs = st.tabs(list(result_dfs.keys()))
                for tab, (sheet_name, df) in zip(tabs, result_dfs.items()):
                    with tab:
                        st.write(f"총 **{len(df):,}** 개의 행이 있습니다.")
                        single_excel = to_excel_bytes({sheet_name: df})
                        st.download_button(
                            label=f"💾 {sheet_name}.xlsx 다운로드",
                            data=single_excel,
                            file_name=f"{sheet_name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{sheet_name}",
                        )
                        st.dataframe(df, use_container_width=True, height=400)
        except Exception as e:
            st.error(f"❌ 데이터 가공 중 오류가 발생했습니다: {e}")
