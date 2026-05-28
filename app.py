# streamlit run app.py
# git add .
# git commit -m "장비 리스트 연동 + 로고 + UI 수정"
# git push

import streamlit as st
from datetime import datetime
import pandas as pd

# =========================
# 설정
# =========================
st.set_page_config(page_title="장비 불출 관리", layout="wide")

LOGO = "aaron_logo.jpg"
EQUIP_FILE = "AARON_Equipments_List.xlsx"
EQUIP_SHEET = "장비통합"   # 장비 리스트 시트명
KEY_COL = "관리 NO."      # 선택 키 컬럼명

# =========================
# 로고 + 타이틀 (요구사항 1)
# =========================
head_l, head_r = st.columns([1, 6])
with head_l:
    try:
        st.image(LOGO, width=120)  # 👈 여기서 로고 크기 조정 가능
    except Exception:
        st.empty()

with head_r:
    st.title("장비 불출 시스템")

# =========================
# 장비 마스터 로딩 (요구사항 2)
# =========================
@st.cache_data(show_spinner=False)
def load_equipment_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=EQUIP_SHEET, engine="openpyxl")

    # 엑셀 병합/빈열 때문에 생기는 Unnamed 컬럼 제거
    df.columns = df.columns.map(str)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

    # 결측치 정리
    df = df.fillna("")

    # 키 컬럼 정리
    if KEY_COL in df.columns:
        df[KEY_COL] = df[KEY_COL].astype(str).str.strip()
        df = df[df[KEY_COL] != ""]

    return df

try:
    equip_df = load_equipment_master(EQUIP_FILE)
except Exception as e:
    st.error(
        f"장비 리스트 파일을 읽지 못했습니다: {EQUIP_FILE}\n\n"
        f"- 레포/폴더에 파일이 있는지 확인\n"
        f"- requirements.txt에 openpyxl 포함 여부 확인\n\n"
        f"에러: {e}"
    )
    st.stop()

if equip_df.empty or KEY_COL not in equip_df.columns:
    st.error(f"장비 리스트에서 '{KEY_COL}' 컬럼을 찾지 못했거나 데이터가 비어 있습니다.")
    st.stop()

# =========================
# 세션 초기화
# =========================
if "records" not in st.session_state:
    st.session_state.records = []

# =========================
# 장비 선택 UI (요구사항 4)
# =========================
st.subheader("🧰 장비 선택 (관리 NO. 기준)")

# 옵션 리스트 (중복 없게)
manage_no_list = equip_df[KEY_COL].astype(str).tolist()

sel_col1, sel_col2 = st.columns([2, 3])

with sel_col1:
    selected_manage_no = st.selectbox("관리 NO. 선택", manage_no_list, key="selected_manage_no")

# 선택된 장비 row
selected_row = equip_df.loc[equip_df[KEY_COL].astype(str) == str(selected_manage_no)].iloc[0]

# 자주 쓰는 컬럼들(있으면 표시)
def safe_get(col):
    return str(selected_row[col]) if col in equip_df.columns else ""

equip_info = {
    "관리 NO.": safe_get("관리 NO."),
    "serial NO.": safe_get("serial NO."),
    "장비명": safe_get("장비명"),
    "모델명": safe_get("모델명"),
    "제조회사": safe_get("제조회사"),
    "장비 구분": safe_get("장비 구분"),
    "최근 위치": safe_get("최근 위치(위치명, 확인날짜기록)"),
    "위치 확인 날짜": safe_get("위치 확인 날짜"),
    "비고": safe_get("비고"),
}

with sel_col2:
    st.markdown("**선택 장비 정보**")
    st.dataframe(
    pd.DataFrame([equip_info]).T.rename(columns={0: "값"}),
    use_container_width=True
)


with st.expander("📌 장비 리스트 전체 보기 (마스터 데이터)", expanded=False):
    st.dataframe(equip_df, use_container_width=True)

# =========================
# 불출 입력 폼
# =========================
st.subheader("🔐 장비 불출 등록")

with st.form("issue_form"):
    col1, col2, col3 = st.columns(3)

    # ---- 사용자 정보 (요구사항 3: 사번/ID 삭제) ----
    with col1:
        user_name = st.text_input("불출자 이름")
        department = st.text_input("소속 부서")

    # ---- 장비 정보: 선택된 관리NO 기반 자동 표시 ----
    with col2:
        st.text_input("관리 NO.", value=equip_info["관리 NO."], disabled=True)
        st.text_input("시리얼(serial NO.)", value=equip_info["serial NO."], disabled=True)
        st.text_input("장비명", value=equip_info["장비명"], disabled=True)
        st.text_input("모델명", value=equip_info["모델명"], disabled=True)

    # ---- 불출 조건 ----
    with col3:
        warehouse = st.selectbox("창고 위치", ["3층", "4층"])
        quantity = st.number_input("수량", min_value=1, step=1)
        return_date = st.date_input("반납 예정일")

    purpose = st.selectbox(
        "사용 목적",
        ["프로젝트", "유지보수", "테스트", "임시대여", "기타"]
    )

    submitted = st.form_submit_button("✅ 불출 등록")

# =========================
# 제출 처리
# =========================
if submitted:
    # 필수 체크: 불출자 이름 + (장비는 선택된 상태라 항상 있음)
    if not user_name:
        st.warning("⚠️ 불출자 이름은 필수입니다.")
    else:
        record = {
            # 사용자
            "불출자": user_name,
            "부서": department,

            # 장비(마스터 연동)
            "관리 NO.": equip_info["관리 NO."],
            "serial NO.": equip_info["serial NO."],
            "장비명": equip_info["장비명"],
            "모델명": equip_info["모델명"],
            "제조회사": equip_info["제조회사"],
            "장비 구분": equip_info["장비 구분"],
            "최근 위치": equip_info["최근 위치"],
            "비고": equip_info["비고"],

            # 불출 조건
            "수량": int(quantity),
            "창고": warehouse,
            "사용 목적": purpose,
            "불출 일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "반납 예정일": return_date.strftime("%Y-%m-%d"),
            "상태": "불출"
        }

        st.session_state.records.append(record)
        st.success("✅ 불출 등록 완료!")

# =========================
# 불출 현황 테이블
# =========================
st.divider()
st.subheader("📋 현재 불출 현황")

if st.session_state.records:
    df = pd.DataFrame(st.session_state.records)
    st.dataframe(df, use_container_width=True)
else:
    st.info("아직 등록된 불출 내역이 없습니다.")