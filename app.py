# streamlit run app.py
# git add .
# git commit -m "UI/마스터 업데이트/사용목적 입력 방식 수정"
# git push

import base64
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

# 마스터 업데이트 대상 컬럼명(엑셀에 존재)
COL_RECENT_LOC = "최근 위치(위치명, 확인날짜기록)"
COL_LOC_DATE   = "위치 확인 날짜"
COL_OUT_DATE   = "반출 일자"
COL_OUT_LOC    = "반출 위치"

# =========================
# 로고를 HTML로 렌더링(제목 옆 + 아래로 약간 내리기)
# =========================
def img_to_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

def render_header():
    try:
        b64 = img_to_base64(LOGO)
        st.markdown(
            f"""
            <div style="display:flex; align-items:flex-start; gap:14px; margin: 6px 0 12px 0;">
                <img src="data:image/jpg;base64,{b64}"
                     style="
                        width:120px;
                        margin-top:18px;   /* 👈 로고를 아래로 내리는 핵심 */
                     " />
                <div>
                    <div style="font-size:44px; font-weight:800; line-height:1.05; margin:0;">
                        장비 불출 시스템
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    except Exception:
        st.title("장비 불출 시스템")

render_header()

# =========================
# 장비 마스터 로딩
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

def save_equipment_master(path: str, df: pd.DataFrame):
    # 원본 시트명 유지해서 덮어쓰기
    # (환경에 따라 쓰기 실패 가능 → 호출부에서 try/except)
    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=EQUIP_SHEET, index=False)

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
# 불출 기록을 기준으로 "표시용 마스터"를 갱신
# =========================
def build_updated_master(df_master: pd.DataFrame, records: list[dict]) -> pd.DataFrame:
    """
    화면 표시용 마스터 데이터:
    - 불출 기록이 있는 관리 NO.에 대해
      최근 위치 = '불출'
      위치 확인 날짜 = 불출일자(오늘)
      반출 일자 = 불출일자(오늘)
      반출 위치 = 사용 목적(입력 문자열)
    """
    df = df_master.copy()

    # 컬럼 없으면 생성(안전)
    for c in [COL_RECENT_LOC, COL_LOC_DATE, COL_OUT_DATE, COL_OUT_LOC]:
        if c not in df.columns:
            df[c] = ""

    for r in records:
        mno = str(r.get("관리 NO.", "")).strip()
        if not mno:
            continue
        mask = df[KEY_COL].astype(str).str.strip() == mno
        if not mask.any():
            continue

        issue_date = r.get("불출 일자(표시)", "")  # yyyy.mm.dd
        purpose = r.get("사용 목적", "")

        df.loc[mask, COL_RECENT_LOC] = "불출"
        df.loc[mask, COL_LOC_DATE] = issue_date
        df.loc[mask, COL_OUT_DATE] = issue_date
        df.loc[mask, COL_OUT_LOC] = purpose

    return df

# =========================
# 장비 선택 UI
# =========================
st.subheader("장비 선택 (관리 NO. 기준)")

manage_no_list = equip_df[KEY_COL].astype(str).tolist()
sel_col1, sel_col2 = st.columns([2, 3])

with sel_col1:
    selected_manage_no = st.selectbox("관리 NO. 선택", manage_no_list, key="selected_manage_no")

selected_row = equip_df.loc[equip_df[KEY_COL].astype(str) == str(selected_manage_no)].iloc[0]

def safe_get(col):
    return str(selected_row[col]) if col in equip_df.columns else ""

equip_info = {
    "관리 NO.": safe_get("관리 NO."),
    "serial NO.": safe_get("serial NO."),
    "장비명": safe_get("장비명"),
    "모델명": safe_get("모델명"),
    "제조회사": safe_get("제조회사"),
    "장비 구분": safe_get("장비 구분"),
    "최근 위치": safe_get(COL_RECENT_LOC),
    "위치 확인 날짜": safe_get(COL_LOC_DATE),
    "비고": safe_get("비고"),
}

with sel_col2:
    st.markdown("**선택 장비 정보**")
    st.dataframe(
        pd.DataFrame([equip_info]).T.rename(columns={0: "값"}),
        use_container_width=True
    )

# =========================
# 불출 입력 폼
# =========================
st.subheader("🔐 장비 불출 등록")

with st.form("issue_form"):
    col1, col2, col3 = st.columns(3)

    with col1:
        user_name = st.text_input("불출자 이름")
        department = st.text_input("소속 부서")

    with col2:
        st.text_input("관리 NO.", value=equip_info["관리 NO."], disabled=True)
        st.text_input("시리얼(serial NO.)", value=equip_info["serial NO."], disabled=True)
        st.text_input("장비명", value=equip_info["장비명"], disabled=True)
        st.text_input("모델명", value=equip_info["모델명"], disabled=True)

    with col3:
        warehouse = st.selectbox("창고 위치", ["3층", "4층"])
        return_date = st.date_input("반납 예정일")

    # ✅ 요구사항 2: 사용 목적 직접 입력
    purpose = st.text_input("사용 목적", placeholder="예: 현장 불출, 유지보수, 프로젝트명 등")

    submitted = st.form_submit_button("✅ 불출 등록")

# =========================
# 제출 처리 (수량 제거 + 마스터 업데이트)
# =========================
if submitted:
    if not user_name:
        st.warning("⚠️ 불출자 이름은 필수입니다.")
    elif not purpose.strip():
        st.warning("⚠️ 사용 목적은 필수입니다. (직접 입력)")
    else:
        now = datetime.now()
        issue_datetime = now.strftime("%Y-%m-%d %H:%M:%S")
        issue_date_dot = now.strftime("%Y.%m.%d")  # 엑셀 기존 표기와 유사하게 점(.) 포맷

        record = {
            "불출자": user_name,
            "부서": department,

            "관리 NO.": equip_info["관리 NO."],
            "serial NO.": equip_info["serial NO."],
            "장비명": equip_info["장비명"],
            "모델명": equip_info["모델명"],
            "제조회사": equip_info["제조회사"],
            "장비 구분": equip_info["장비 구분"],

            "창고": warehouse,
            "사용 목적": purpose.strip(),

            "불출 일시": issue_datetime,
            "불출 일자(표시)": issue_date_dot,
            "반납 예정일": return_date.strftime("%Y-%m-%d"),
            "상태": "불출",
        }

        st.session_state.records.append(record)

        # ✅ 마스터(화면표시용) 갱신 + 엑셀 저장 시도
        updated_master = build_updated_master(equip_df, st.session_state.records)

        try:
            save_equipment_master(EQUIP_FILE, updated_master)
            # 저장 성공 시, 캐시 갱신을 위해 clear
            st.cache_data.clear()
            equip_df = load_equipment_master(EQUIP_FILE)
        except Exception:
            # 저장이 안 되는 환경도 있을 수 있으니, 화면에서만 갱신되도록 유지
            pass

        st.success("✅ 불출 등록 완료!")

# =========================
# 현재 불출 현황
# =========================
st.divider()
st.subheader("📋 현재 불출 현황")

if st.session_state.records:
    df_records = pd.DataFrame(st.session_state.records).drop(columns=["불출 일자(표시)"], errors="ignore")
    st.dataframe(df_records, use_container_width=True)
else:
    st.info("아직 등록된 불출 내역이 없습니다.")

# =========================
# ✅ 요구사항 3: 마스터 데이터는 '현재 불출 현황' 아래에 항상 표시 + 자동 업데이트
# =========================
st.subheader("📌 장비 리스트 전체 보기 (마스터 데이터)")

# 화면 표시는 항상 records 기반으로 즉시 갱신된 버전 사용
updated_master_view = build_updated_master(equip_df, st.session_state.records)
st.dataframe(updated_master_view, use_container_width=True)