# streamlit run app.py
# git add .
# git commit -m "UI/마스터 업데이트/사용목적 입력 방식 수정"
# git push

import os
import streamlit as st
from datetime import datetime
import pandas as pd

# =========================
# 설정
# =========================
st.set_page_config(page_title="장비 불출 관리", layout="wide")

LOGO = "aaron_logo.jpg"

EQUIP_FILE = "AARON_Equipments_List.xlsx"
EQUIP_SHEET = "장비통합"
KEY_COL = "관리 NO."

# ✅ 불출 이력 파일(새로 추가)
HISTORY_FILE = "AARON_Issue_History.xlsx"
HISTORY_SHEET = "불출이력"

# 마스터 업데이트 대상 컬럼(엑셀에 존재) 
COL_RECENT_LOC = "최근 위치(위치명, 확인날짜기록)"
COL_LOC_DATE   = "위치 확인 날짜"
COL_OUT_DATE   = "반출 일자"
COL_OUT_LOC    = "반출 위치"

# =========================
# 헤더(안전하게 로고를 제목 위)
# =========================
try:
    st.image(LOGO, width=170)  # 필요하면 여기만 조절
except Exception:
    pass

st.title("장비 불출 시스템")

# =========================
# 장비 마스터 로딩/저장
# =========================
@st.cache_data(show_spinner=False)
def load_equipment_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=EQUIP_SHEET, engine="openpyxl")
    df.columns = df.columns.map(str)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    df = df.fillna("")

    if KEY_COL in df.columns:
        df[KEY_COL] = df[KEY_COL].astype(str).str.strip()
        df = df[df[KEY_COL] != ""]

    return df

def save_equipment_master(path: str, df: pd.DataFrame):
    # "장비통합" 시트를 덮어쓰기
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
# ✅ 불출 이력: 생성/누적 저장 유틸
# =========================
HISTORY_COLS = [
    "불출 일시", "불출 일자",
    "불출자", "부서",
    "관리 NO.", "serial NO.", "장비명", "모델명", "제조회사", "장비 구분",
    "창고", "사용 목적",
    "반납 예정일",
    "상태",
]

def append_issue_history(new_row: dict):
    """
    - HISTORY_FILE 없으면 생성
    - 있으면 읽어서 append 후 저장
    """
    row_df = pd.DataFrame([{c: new_row.get(c, "") for c in HISTORY_COLS}])

    if not os.path.exists(HISTORY_FILE):
        # 파일이 없으면 새로 생성
        with pd.ExcelWriter(HISTORY_FILE, engine="openpyxl") as writer:
            row_df.to_excel(writer, sheet_name=HISTORY_SHEET, index=False)
        return

    # 파일이 있으면 기존 로드 후 append
    try:
        old = pd.read_excel(HISTORY_FILE, sheet_name=HISTORY_SHEET, engine="openpyxl")
        old.columns = old.columns.map(str)
    except Exception:
        # 시트가 없거나 깨졌으면 새로 생성(안전)
        old = pd.DataFrame(columns=HISTORY_COLS)

    # 컬럼 보정
    for c in HISTORY_COLS:
        if c not in old.columns:
            old[c] = ""

    combined = pd.concat([old[HISTORY_COLS], row_df], ignore_index=True)

    # 덮어쓰기 저장
    with pd.ExcelWriter(HISTORY_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        combined.to_excel(writer, sheet_name=HISTORY_SHEET, index=False)

@st.cache_data(show_spinner=False)
def load_issue_history() -> pd.DataFrame:
    if not os.path.exists(HISTORY_FILE):
        return pd.DataFrame(columns=HISTORY_COLS)
    try:
        df = pd.read_excel(HISTORY_FILE, sheet_name=HISTORY_SHEET, engine="openpyxl")
        df.columns = df.columns.map(str)
        return df
    except Exception:
        return pd.DataFrame(columns=HISTORY_COLS)

# =========================
# 세션 초기화(화면 표시용)
# =========================
if "records" not in st.session_state:
    st.session_state.records = []

# ✅ 앱 시작 시 불출 이력에서 세션 복원
if not st.session_state.get("initialized", False):
    history_df = load_issue_history()
    if not history_df.empty:
        st.session_state.records = history_df.to_dict("records")
    st.session_state.initialized = True


# =========================
# 화면 표시용 마스터 업데이트
# =========================
def build_updated_master(df_master: pd.DataFrame, records: list[dict]) -> pd.DataFrame:
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

        issue_date_dot = r.get("불출 일자", "")
        purpose = r.get("사용 목적", "")

        df.loc[mask, COL_RECENT_LOC] = "불출"
        df.loc[mask, COL_LOC_DATE] = issue_date_dot
        df.loc[mask, COL_OUT_DATE] = issue_date_dot
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
    st.dataframe(pd.DataFrame([equip_info]).T.rename(columns={0: "값"}), use_container_width=True)

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

    # 사용 목적 직접 입력
    purpose = st.text_input("사용 목적", placeholder="예: 현장 불출, 유지보수, 프로젝트명 등")

    submitted = st.form_submit_button("✅ 불출 등록")

# =========================
# 제출 처리: 화면 이력 + 엑셀 이력 append + 마스터 업데이트
# =========================
if submitted:
    if not user_name:
        st.warning("⚠️ 불출자 이름은 필수입니다.")
    elif not purpose.strip():
        st.warning("⚠️ 사용 목적은 필수입니다.")
    else:
        now = datetime.now()
        issue_datetime = now.strftime("%Y-%m-%d %H:%M:%S")
        issue_date_dot = now.strftime("%Y.%m.%d")

        # 화면 표시용 record
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
            "불출 일자": issue_date_dot,
            "반납 예정일": return_date.strftime("%Y-%m-%d"),
            "상태": "불출",
        }
        st.session_state.records.append(record)

        # ✅ (1) 불출 이력 엑셀에 누적 저장
        try:
            append_issue_history(record)
            st.cache_data.clear()  # load_issue_history 캐시 갱신
        except Exception as e:
            st.warning(f"이력 저장 중 오류가 발생했습니다(환경 권한/경로 확인 필요): {e}")

        # ✅ (2) 마스터 업데이트(화면 + 가능하면 파일에도 반영)
        updated_master = build_updated_master(equip_df, st.session_state.records)

        try:
            save_equipment_master(EQUIP_FILE, updated_master)
            st.cache_data.clear()
            equip_df = load_equipment_master(EQUIP_FILE)
        except Exception:
            # 저장이 막힌 환경이면 화면만 갱신됨
            pass

        st.success("✅ 불출 등록 완료!")

# =========================
# 현재 불출 현황
# =========================
st.divider()
st.subheader("📋 현재 불출 현황")

if st.session_state.records:
    st.dataframe(pd.DataFrame(st.session_state.records), use_container_width=True)
else:
    st.info("아직 등록된 불출 내역이 없습니다.")

# =========================
# ✅ 불출 이력(영구 로그) 표시
# =========================
st.subheader("📒 불출 이력 ")

history_df = load_issue_history()
st.dataframe(history_df, use_container_width=True)

# =========================
# 마스터 데이터(항상 표시 + 자동 업데이트)
# =========================
st.subheader("📌 장비 리스트 전체 보기 (마스터 데이터)")

updated_master_view = build_updated_master(equip_df, st.session_state.records)
st.dataframe(updated_master_view, use_container_width=True)