# streamlit run app.py
# git add .
# git commit -m "불출/반납 통합, KST 적용, 권한 분리, 원위치 복구 추가"
# git push

import json
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from supabase import create_client, Client

# =========================
# 설정
# =========================
st.set_page_config(page_title="장비 불출/반납 관리", layout="wide")

LOGO = "aaron_logo.jpg"

EQUIP_FILE = "AARON_Equipments_List.xlsx"
EQUIP_SHEET = "장비통합"
KEY_COL = "관리 NO."

COL_RECENT_LOC = "최근 위치(위치명, 확인날짜기록)"
COL_LOC_DATE   = "위치 확인 날짜"
COL_OUT_DATE   = "반출 일자"
COL_OUT_LOC    = "반출 위치"

SUPABASE_TABLE = "issue_history"
ORIGIN_CACHE_FILE = "issue_origin_cache.json"

KST = ZoneInfo("Asia/Seoul")

# =========================
# 헤더
# =========================
try:
    st.image(LOGO, width=170)
except Exception:
    pass

st.title("장비 불출/반납 시스템")
st.caption(f"현재 한국 표준시(KST): {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")

# =========================
# 공통 함수
# =========================
def now_kst() -> datetime:
    return datetime.now(KST)

def format_dt(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y-%m-%d %H:%M:%S")

def format_date_dot(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y.%m.%d")

def format_date_dash(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y-%m-%d")

def parse_dt_safe(value):
    try:
        return pd.to_datetime(value, errors="coerce")
    except Exception:
        return pd.NaT

def safe_str(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()

# =========================
# 권한 설정
# =========================
# secrets.toml 예시
# ADMIN_USERS = ["admin", "manager1"]
# USER_USERS = ["user1", "user2"]

def get_user_role(user_id: str) -> str:
    user_id = safe_str(user_id)
    admin_users = st.secrets.get("ADMIN_USERS", [])
    user_users = st.secrets.get("USER_USERS", [])

    if user_id in admin_users:
        return "관리자"
    if user_id in user_users:
        return "사용자"
    return "사용자"

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
# 원위치 캐시 (반납 복구용)
# =========================
def load_origin_cache() -> dict:
    if not os.path.exists(ORIGIN_CACHE_FILE):
        return {}
    try:
        with open(ORIGIN_CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_origin_cache(data: dict):
    try:
        with open(ORIGIN_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.warning(f"원위치 캐시 저장 실패(배포환경/권한 이슈 가능): {e}")

def set_origin_cache(manage_no: str, payload: dict):
    cache = load_origin_cache()
    cache[manage_no] = payload
    save_origin_cache(cache)

def get_origin_cache(manage_no: str) -> dict:
    cache = load_origin_cache()
    return cache.get(manage_no, {})

def pop_origin_cache(manage_no: str):
    cache = load_origin_cache()
    if manage_no in cache:
        cache.pop(manage_no, None)
        save_origin_cache(cache)

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

    rename_map = {c: DB_TO_KR[c] for c in df.columns if c in DB_TO_KR}
    df = df.rename(columns=rename_map)

    for c in HISTORY_COLS:
        if c not in df.columns:
            df[c] = ""

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
        res = supabase.table(SUPABASE_TABLE).select("*").execute()
        data = res.data if hasattr(res, "data") and res.data else []
        df = pd.DataFrame(data)
        df = normalize_history_df(df)

        if not df.empty and "불출 일시" in df.columns:
            df["_sort_dt"] = pd.to_datetime(df["불출 일시"], errors="coerce")
            df = df.sort_values("_sort_dt", ascending=False).drop(columns=["_sort_dt"])

        return df
    except Exception as e:
        st.warning(f"DB 이력 조회 중 오류가 발생했습니다: {e}")
        return pd.DataFrame(columns=HISTORY_COLS)

# =========================
# 상태 계산 유틸
# =========================
def get_latest_records(records: list[dict]) -> dict[str, dict]:
    latest_map: dict[str, dict] = {}

    for r in records:
        mno = safe_str(r.get("관리 NO.", ""))
        if not mno:
            continue

        cur_dt = parse_dt_safe(r.get("불출 일시", ""))
        prev = latest_map.get(mno)

        if prev is None:
            latest_map[mno] = r
            continue

        prev_dt = parse_dt_safe(prev.get("불출 일시", ""))
        if pd.isna(prev_dt) or (not pd.isna(cur_dt) and cur_dt >= prev_dt):
            latest_map[mno] = r

    return latest_map

def get_current_issued_records(records: list[dict]) -> list[dict]:
    latest_map = get_latest_records(records)
    current = []

    for _, r in latest_map.items():
        if safe_str(r.get("상태", "")) == "불출":
            current.append(r)

    current.sort(key=lambda x: parse_dt_safe(x.get("불출 일시", "")), reverse=True)
    return current

# =========================
# 세션 초기화
# =========================
if "records" not in st.session_state:
    st.session_state.records = []

if "initialized" not in st.session_state:
    history_df = load_db_history()
    if not history_df.empty:
        st.session_state.records = history_df.to_dict("records")
    st.session_state.initialized = True

if "selected_return_target" not in st.session_state:
    st.session_state.selected_return_target = ""

# =========================
# 화면 표시용 마스터 업데이트
# =========================
def build_updated_master(df_master: pd.DataFrame, records: list[dict]) -> pd.DataFrame:
    df = df_master.copy()

    for c in [COL_RECENT_LOC, COL_LOC_DATE, COL_OUT_DATE, COL_OUT_LOC]:
        if c not in df.columns:
            df[c] = ""

    latest_map = get_latest_records(records)

    for mno, r in latest_map.items():
        mask = df[KEY_COL].astype(str).str.strip() == str(mno).strip()
        if not mask.any():
            continue

        event_date = safe_str(r.get("불출 일자", ""))
        purpose = safe_str(r.get("사용 목적", ""))
        warehouse = safe_str(r.get("창고", ""))
        status = safe_str(r.get("상태", ""))

        if status == "불출":
            df.loc[mask, COL_RECENT_LOC] = "불출"
            df.loc[mask, COL_LOC_DATE] = event_date
            df.loc[mask, COL_OUT_DATE] = event_date
            df.loc[mask, COL_OUT_LOC] = purpose if purpose else warehouse

        elif status == "반납":
            origin_info = get_origin_cache(mno)
            restored_loc = safe_str(origin_info.get("recent_loc", ""))
            restored_date = safe_str(origin_info.get("loc_date", ""))

            if not restored_loc:
                # fallback: 창고 기준 복구
                restored_loc = f"본사, {warehouse} 창고" if warehouse else "반납"
            if not restored_date:
                restored_date = event_date

            df.loc[mask, COL_RECENT_LOC] = restored_loc
            df.loc[mask, COL_LOC_DATE] = restored_date
            df.loc[mask, COL_OUT_DATE] = ""
            df.loc[mask, COL_OUT_LOC] = ""

    return df

# =========================
# 사이드바: 사용자/권한/모드
# =========================
with st.sidebar:
    st.subheader("사용자 설정")
    input_user_id = st.text_input("사용자 ID", value=st.session_state.get("user_id", ""))
    input_user_name = st.text_input("이름", value=st.session_state.get("user_name", ""))
    input_department = st.text_input("부서", value=st.session_state.get("department", ""))

    st.session_state.user_id = safe_str(input_user_id)
    st.session_state.user_name = safe_str(input_user_name)
    st.session_state.department = safe_str(input_department)
    st.session_state.role = get_user_role(st.session_state.user_id)

    st.success(f"권한: {st.session_state.role}")

    mode = st.radio(
        "작업 모드",
        ["🔐 불출", "📥 반납"],
        horizontal=False
    )

# =========================
# 장비 선택 UI (불출용)
# =========================
manage_no_list = equip_df[KEY_COL].astype(str).tolist()

def get_selected_row(df: pd.DataFrame, manage_no: str) -> pd.Series:
    row = df.loc[df[KEY_COL].astype(str) == str(manage_no)].iloc[0]
    return row

def build_equip_info(selected_row: pd.Series, df: pd.DataFrame) -> dict:
    def safe_get(col):
        return str(selected_row[col]) if col in df.columns else ""

    return {
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

# =========================
# 불출 모드
# =========================
if mode == "🔐 불출":
    st.subheader("장비 선택 (관리 NO. 기준)")

    sel_col1, sel_col2 = st.columns([2, 3])

    with sel_col1:
        selected_manage_no = st.selectbox("관리 NO. 선택", manage_no_list, key="selected_manage_no")

    selected_row = get_selected_row(equip_df, selected_manage_no)
    equip_info = build_equip_info(selected_row, equip_df)

    latest_map = get_latest_records(st.session_state.records)
    latest_status = safe_str(latest_map.get(selected_manage_no, {}).get("상태", ""))

    with sel_col2:
        view_info = equip_info.copy()
        view_info["현재 상태(시스템)"] = latest_status if latest_status else "미등록"
        st.markdown("**선택 장비 정보**")
        st.dataframe(
            pd.DataFrame([view_info]).T.rename(columns={0: "값"}),
            use_container_width=True
        )

    st.subheader("🔐 장비 불출 등록")

    with st.form("issue_form"):
        col1, col2, col3 = st.columns(3)

        with col1:
            user_name = st.text_input("불출자 이름", value=st.session_state.get("user_name", ""))
            department = st.text_input("소속 부서", value=st.session_state.get("department", ""))

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

    if submitted:
        current_latest = latest_map.get(equip_info["관리 NO."], {})
        current_status = safe_str(current_latest.get("상태", ""))

        if current_status == "불출":
            st.warning("⚠️ 이미 불출 중인 장비입니다. 먼저 반납 처리해주세요.")
        elif not user_name.strip():
            st.warning("⚠️ 불출자 이름은 필수입니다.")
        elif not purpose.strip():
            st.warning("⚠️ 사용 목적은 필수입니다.")
        else:
            now = now_kst()
            issue_datetime = format_dt(now)
            issue_date_dot = format_date_dot(now)

            # 불출 전 원래 위치 저장 (반납 시 복구용)
            origin_payload = {
                "recent_loc": safe_str(equip_info.get("최근 위치", "")),
                "loc_date": safe_str(equip_info.get("위치 확인 날짜", "")),
                "saved_at": issue_datetime,
            }
            set_origin_cache(equip_info["관리 NO."], origin_payload)

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

            ok, msg = save_db(record)
            if not ok:
                st.error(f"DB 저장 실패: {msg}")
            else:
                st.session_state.records.append(record)

                updated_master = build_updated_master(equip_df, st.session_state.records)

                try:
                    save_equipment_master(EQUIP_FILE, updated_master)
                    st.cache_data.clear()
                    equip_df = load_equipment_master(EQUIP_FILE)
                except Exception as e:
                    st.warning(f"마스터 파일 저장은 생략되었습니다(권한/환경 이슈 가능): {e}")

                st.cache_data.clear()
                st.success("✅ 불출 등록 완료! (한국 표준시 기준)")
                st.rerun()

# =========================
# 반납 모드
# =========================
elif mode == "📥 반납":
    st.subheader("📥 불출 중 장비 목록")

    current_issued = get_current_issued_records(st.session_state.records)

    if not current_issued:
        st.info("현재 불출 중인 장비가 없습니다.")
    else:
        # 자동 필터 목록 + 반납 버튼
        for rec in current_issued:
            mno = safe_str(rec.get("관리 NO.", ""))
            item_name = safe_str(rec.get("장비명", ""))
            purpose = safe_str(rec.get("사용 목적", ""))
            issue_dt = safe_str(rec.get("불출 일시", ""))
            issue_user = safe_str(rec.get("불출자", ""))

            c1, c2, c3, c4, c5 = st.columns([2, 3, 3, 3, 1])
            with c1:
                st.write(f"**{mno}**")
            with c2:
                st.write(item_name)
            with c3:
                st.write(purpose if purpose else "-")
            with c4:
                st.write(f"{issue_user} / {issue_dt}")
            with c5:
                if st.button("반납", key=f"return_btn_{mno}"):
                    st.session_state.selected_return_target = mno

        st.divider()

        selected_target = st.session_state.get("selected_return_target", "")
        if not selected_target:
            st.info("반납할 장비의 **반납 버튼**을 눌러주세요.")
        else:
            target_record = None
            for r in current_issued:
                if safe_str(r.get("관리 NO.", "")) == selected_target:
                    target_record = r
                    break

            if target_record is None:
                st.warning("선택한 반납 대상 정보를 찾을 수 없습니다.")
            else:
                st.subheader(f"↩️ 반납 등록 - {selected_target}")

                target_manage_no = safe_str(target_record.get("관리 NO.", ""))
                target_row = equip_df.loc[equip_df[KEY_COL].astype(str) == target_manage_no]
                target_master = target_row.iloc[0] if not target_row.empty else None

                with st.form("return_form"):
                    col1, col2, col3 = st.columns(3)

                    with col1:
                        return_user = st.text_input("반납자 이름", value=st.session_state.get("user_name", ""))
                        return_department = st.text_input("소속 부서", value=st.session_state.get("department", ""))

                    with col2:
                        st.text_input("관리 NO.", value=target_manage_no, disabled=True)
                        st.text_input("장비명", value=safe_str(target_record.get("장비명", "")), disabled=True)
                        st.text_input("모델명", value=safe_str(target_record.get("모델명", "")), disabled=True)

                    with col3:
                        return_warehouse = st.selectbox("반납 창고 위치", ["3층", "4층"], key="return_warehouse")
                        st.text_input("기존 불출자", value=safe_str(target_record.get("불출자", "")), disabled=True)

                    return_note = st.text_input(
                        "반납 메모",
                        placeholder="예: 정상 반납, 점검 필요, 부속품 포함 등"
                    )

                    return_submitted = st.form_submit_button("✅ 반납 등록")

                if return_submitted:
                    if not return_user.strip():
                        st.warning("⚠️ 반납자 이름은 필수입니다.")
                    else:
                        now = now_kst()
                        return_datetime = format_dt(now)
                        return_date_dot = format_date_dot(now)

                        return_record = {
                            "불출자": return_user.strip(),  # 이벤트 처리자
                            "부서": return_department.strip(),
                            "관리 NO.": safe_str(target_record.get("관리 NO.", "")),
                            "serial NO.": safe_str(target_record.get("serial NO.", "")),
                            "장비명": safe_str(target_record.get("장비명", "")),
                            "모델명": safe_str(target_record.get("모델명", "")),
                            "제조회사": safe_str(target_record.get("제조회사", "")),
                            "장비 구분": safe_str(target_record.get("장비 구분", "")),
                            "창고": return_warehouse,
                            "사용 목적": return_note.strip() if return_note.strip() else "반납",
                            "불출 일시": return_datetime,
                            "불출 일자": return_date_dot,
                            "반납 예정일": safe_str(target_record.get("반납 예정일", "")),
                            "상태": "반납",
                        }

                        ok, msg = save_db(return_record)
                        if not ok:
                            st.error(f"DB 저장 실패: {msg}")
                        else:
                            st.session_state.records.append(return_record)

                            updated_master = build_updated_master(equip_df, st.session_state.records)

                            try:
                                save_equipment_master(EQUIP_FILE, updated_master)
                                st.cache_data.clear()
                                equip_df = load_equipment_master(EQUIP_FILE)
                            except Exception as e:
                                st.warning(f"마스터 파일 저장은 생략되었습니다(권한/환경 이슈 가능): {e}")

                            # 반납 완료 후 origin cache 제거
                            pop_origin_cache(target_manage_no)
                            st.session_state.selected_return_target = ""

                            st.cache_data.clear()
                            st.success("✅ 반납 등록 완료! (한국 표준시 기준)")
                            st.rerun()

# =========================
# 현재 불출 현황
# =========================
st.divider()
st.subheader("📋 현재 불출 현황")

current_df = pd.DataFrame(get_current_issued_records(st.session_state.records))
if not current_df.empty:
    st.dataframe(current_df, use_container_width=True)
else:
    st.info("현재 불출 중인 장비가 없습니다.")

# =========================
# 관리자 전용: DB 이력 / 마스터 보기
# =========================
if st.session_state.get("role", "사용자") == "관리자":
    st.divider()
    st.subheader("📒 불출/반납 이력 (Supabase)")

    history_df = load_db_history()
    st.dataframe(history_df, use_container_width=True)

    st.subheader("📌 장비 리스트 전체 보기 (마스터 데이터)")
    updated_master_view = build_updated_master(equip_df, st.session_state.records)
    st.dataframe(updated_master_view, use_container_width=True)
else:
    st.divider()
    st.info("사용자 권한에서는 전체 이력/마스터 데이터 전체 보기는 숨김 처리됩니다.")