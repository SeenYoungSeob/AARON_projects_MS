# streamlit run app.py
# git add .
# git commit -m "UI/마스터 업데이트/사용목적 입력 방식 수정"
# git push

# streamlit run app.py

import os
from datetime import datetime

import pandas as pd
import streamlit as st
from supabase import create_client, Client

# =========================
# 설정
# =========================
st.set_page_config(page_title="장비 불출 관리", layout="wide")

LOGO = "aaron_logo.jpg"

EQUIP_FILE = "AARON_Equipments_List.xlsx"
EQUIP_SHEET = "장비통합"
KEY_COL = "관리 NO."

COL_RECENT_LOC = "최근 위치(위치명, 확인날짜기록)"
COL_LOC_DATE   = "위치 확인 날짜"
COL_OUT_DATE   = "반출 일자"
COL_OUT_LOC    = "반출 위치"

SUPABASE_TABLE = "issue_history"

# =========================
# 헤더
# =========================
try:
    st.image(LOGO, width=170)
except Exception:
    pass

st.title("장비 불출 시스템")

# =========================
# Supabase 연결
# =========================
@st.cache_resource
def get_supabase_client() -> Client | None:
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception as e:
        st.error(f"Supabase 연결 정보를 읽지 못했습니다: {e}")
        return None

supabase = get_supabase_client()

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
    """
    기존과 동일하게 장비통합 시트를 덮어쓰기 시도.
    배포 환경에서 파일 쓰기가 막혀 있으면 except에서 무시하고
    화면 갱신만 유지.
    """
    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=EQUIP_SHEET, index=False)

try:
    equip_df = load_equipment_master(EQUIP_FILE)
except Exception as e:
    st.error(
        f"장비 리스트 파일을 읽지 못했습니다: {EQUIP_FILE}\n\n"
        f"- 파일 존재 여부 확인\n"
        f"- requirements.txt에 openpyxl 포함 여부 확인\n\n"
        f"에러: {e}"
    )
    st.stop()

if equip_df.empty or KEY_COL not in equip_df.columns:
    st.error(f"장비 리스트에서 '{KEY_COL}' 컬럼을 찾지 못했거나 데이터가 비어 있습니다.")
    st.stop()

# =========================
# DB 컬럼 정의
# =========================
DB_TO_KR = {
    "issue_datetime": "불출 일시",
    "issue_date": "불출 일자",
    "user_name": "불출자",
    "department": "부서",
    "manage_no": "관리 NO.",
    "serial_no": "serial NO.",
    "item_name": "장비명",
    "model_name": "모델명",
    "manufacturer": "제조회사",
    "item_group": "장비 구분",
    "warehouse": "창고",
    "purpose": "사용 목적",
    "return_date": "반납 예정일",
    "status": "상태",
}

KR_TO_DB = {v: k for k, v in DB_TO_KR.items()}

HISTORY_COLS = [
    "불출 일시", "불출 일자",
    "불출자", "부서",
    "관리 NO.", "serial NO.", "장비명", "모델명", "제조회사", "장비 구분",
    "창고", "사용 목적",
    "반납 예정일",
    "상태",
]

# =========================
# DB 저장 / 조회
# =========================
def normalize_history_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=HISTORY_COLS)

    # 영문 컬럼 -> 한글 컬럼 변환
    rename_map = {c: DB_TO_KR[c] for c in df.columns if c in DB_TO_KR}
    df = df.rename(columns=rename_map)

    # 필요한 컬럼 보정
    for c in HISTORY_COLS:
        if c not in df.columns:
            df[c] = ""

    # 보기용 순서 정렬
    df = df[HISTORY_COLS].fillna("")
    return df

def save_db(record: dict) -> tuple[bool, str]:
    if supabase is None:
        return False, "Supabase 클라이언트가 준비되지 않았습니다."

    payload = {
        "issue_datetime": record.get("불출 일시", ""),
        "issue_date": record.get("불출 일자", ""),
        "user_name": record.get("불출자", ""),
        "department": record.get("부서", ""),
        "manage_no": record.get("관리 NO.", ""),
        "serial_no": record.get("serial NO.", ""),
        "item_name": record.get("장비명", ""),
        "model_name": record.get("모델명", ""),
        "manufacturer": record.get("제조회사", ""),
        "item_group": record.get("장비 구분", ""),
        "warehouse": record.get("창고", ""),
        "purpose": record.get("사용 목적", ""),
        "return_date": record.get("반납 예정일", ""),
        "status": record.get("상태", ""),
    }

    try:
        supabase.table(SUPABASE_TABLE).insert(payload).execute()
        return True, "저장 완료"
    except Exception as e:
        return False, str(e)

@st.cache_data(show_spinner=False)
def load_db_history() -> pd.DataFrame:
    if supabase is None:
        return pd.DataFrame(columns=HISTORY_COLS)

    try:
        # 정렬이 안 되는 환경도 있어서 우선 전체 select
        res = supabase.table(SUPABASE_TABLE).select("*").execute()
        data = res.data if hasattr(res, "data") and res.data else []
        df = pd.DataFrame(data)
        df = normalize_history_df(df)

        # 불출 일시 기준 정렬 시도
        if not df.empty and "불출 일시" in df.columns:
            try:
                df["_sort_dt"] = pd.to_datetime(df["불출 일시"], errors="coerce")
                df = df.sort_values("_sort_dt", ascending=False).drop(columns=["_sort_dt"])
            except Exception:
                pass

        return df
    except Exception as e:
        st.warning(f"DB 이력 조회 중 오류가 발생했습니다: {e}")
        return pd.DataFrame(columns=HISTORY_COLS)

# =========================
# 세션 초기화
# =========================
if "records" not in st.session_state:
    st.session_state.records = []

if not st.session_state.get("initialized", False):
    history_df = load_db_history()
    if not history_df.empty:
        st.session_state.records = history_df.to_dict("records")
    st.session_state.initialized = True

# =========================
# 화면 표시용 마스터 업데이트
# =========================
def build_updated_master(df_master: pd.DataFrame, records: list[dict]) -> pd.DataFrame:
    df = df_master.copy()

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

    purpose = st.text_input("사용 목적", placeholder="예: 현장 불출, 유지보수, 프로젝트명 등")

    submitted = st.form_submit_button("✅ 불출 등록")

# =========================
# 제출 처리
# =========================
if submitted:
    if not user_name.strip():
        st.warning("⚠️ 불출자 이름은 필수입니다.")
    elif not purpose.strip():
        st.warning("⚠️ 사용 목적은 필수입니다.")
    else:
        now = datetime.now()
        issue_datetime = now.strftime("%Y-%m-%d %H:%M:%S")
        issue_date_dot = now.strftime("%Y.%m.%d")

        record = {
            "불출자": user_name.strip(),
            "부서": department.strip(),
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
            "반납 예정일": return_date.strftime("%Y-%m-%d") if return_date else "",
            "상태": "불출",
        }

        # 1) DB 저장
        ok, msg = save_db(record)
        if not ok:
            st.error(f"DB 저장 실패: {msg}")
        else:
            # 2) 화면 반영
            st.session_state.records.append(record)

            # 3) 마스터 업데이트(화면 + 가능하면 파일 반영)
            updated_master = build_updated_master(equip_df, st.session_state.records)

            try:
                save_equipment_master(EQUIP_FILE, updated_master)
                st.cache_data.clear()
                equip_df = load_equipment_master(EQUIP_FILE)
            except Exception as e:
                # 배포환경에서는 파일 쓰기 막힐 수 있으므로 경고만
                st.warning(f"마스터 파일 저장은 생략되었습니다(권한/환경 이슈 가능): {e}")

            st.cache_data.clear()
            st.success("✅ 불출 등록 완료!")
            st.rerun()

# =========================
# 현재 불출 현황
# =========================
st.divider()
st.subheader("📋 현재 불출 현황")

if st.session_state.records:
    current_df = pd.DataFrame(st.session_state.records)
    st.dataframe(current_df, use_container_width=True)
else:
    st.info("아직 등록된 불출 내역이 없습니다.")

# =========================
# DB 불출 이력
# =========================
st.subheader("📒 불출 이력 (Supabase)")

history_df = load_db_history()
st.dataframe(history_df, use_container_width=True)

# =========================
# 마스터 데이터 보기
# =========================
st.subheader("📌 장비 리스트 전체 보기 (마스터 데이터)")

updated_master_view = build_updated_master(equip_df, st.session_state.records)
st.dataframe(updated_master_view, use_container_width=True)