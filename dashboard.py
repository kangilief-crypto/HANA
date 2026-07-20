import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path


# =========================
# 기본 설정
# =========================
st.set_page_config(
    page_title="HANA 유통점 휴일 프로모션 대시보드",
    page_icon="📊",
    layout="wide",
)


# =========================
# 스타일
# =========================
st.markdown(
    """
    <style>
    .main {
        background-color: #fbf7ff;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #7c3aed 0%, #14b8a6 100%);
        padding: 18px;
        border-radius: 18px;
        color: white;
        text-align: center;
        margin-bottom: 12px;
    }
    .metric-card h3 {
        font-size: 15px;
        margin-bottom: 6px;
        font-weight: 600;
    }
    .metric-card p {
        font-size: 28px;
        margin: 0;
        font-weight: 800;
    }
    .section-title {
        font-size: 22px;
        font-weight: 800;
        color: #4c1d95;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    .sub-box {
        background-color: #ffffff;
        border: 1px solid #e9d5ff;
        border-radius: 16px;
        padding: 18px;
        margin-bottom: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# 비밀번호 설정
# =========================
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.title("HANA 유통점 휴일 프로모션 대시보드")
    password = st.text_input("비밀번호를 입력하세요", type="password")

    if st.button("접속"):
        try:
            correct_password = st.secrets["password"]
        except Exception:
            correct_password = ""

        if password == correct_password:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")

    return False


if not check_password():
    st.stop()


# =========================
# 데이터 로드
# =========================
DATA_FILE = Path("weekly_data.xlsx")


@st.cache_data
def load_excel_data(file_path):
    xls = pd.ExcelFile(file_path)
    sheets = {}

    for sheet_name in xls.sheet_names:
        sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)

    return sheets


def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {
        "담당팀": "담당팀명",
        "대리점명": "대리점한글명",
        "예약인원": "순예약인원",
        "예약": "순예약인원",
    }

    for old, new in rename_map.items():
        if old in df.columns and new not in df.columns:
            df = df.rename(columns={old: new})

    return df


def to_number(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


def find_sheet(sheets, candidates):
    for name in candidates:
        if name in sheets:
            return sheets[name]
    return None


try:
    all_sheets = load_excel_data(DATA_FILE)
except FileNotFoundError:
    st.error("weekly_data.xlsx 파일을 찾을 수 없습니다. GitHub 저장소 최상단에 업로드했는지 확인해 주세요.")
    st.stop()
except Exception as e:
    st.error(f"엑셀 파일을 불러오는 중 오류가 발생했습니다: {e}")
    st.stop()


raw_final = find_sheet(all_sheets, ["최종실적", "최종 실적", "실적", "Sheet1"])
raw_dept = find_sheet(all_sheets, ["부서별 모객현황", "부서별모객현황", "부서별 예약인원"])
raw_participant = find_sheet(all_sheets, ["잠재대리점", "잠재 대리점", "참여대리점", "참여 대리점"])

if raw_final is None:
    st.error("엑셀에서 '최종실적' 시트를 찾지 못했습니다.")
    st.stop()

final_df = normalize_columns(raw_final)
dept_df = normalize_columns(raw_dept) if raw_dept is not None else pd.DataFrame()
participant_df = normalize_columns(raw_participant) if raw_participant is not None else pd.DataFrame()


# =========================
# 필수 컬럼 보정
# =========================
if "대리점한글명" not in final_df.columns:
    possible_agency_cols = [c for c in final_df.columns if "대리점" in c or "거래처" in c]
    if possible_agency_cols:
        final_df = final_df.rename(columns={possible_agency_cols[0]: "대리점한글명"})

if "담당부서" not in final_df.columns:
    possible_dept_cols = [c for c in final_df.columns if "부서" in c]
    if possible_dept_cols:
        final_df = final_df.rename(columns={possible_dept_cols[0]: "담당부서"})

if "담당팀명" not in final_df.columns:
    possible_team_cols = [c for c in final_df.columns if "팀" in c]
    if possible_team_cols:
        final_df = final_df.rename(columns={possible_team_cols[0]: "담당팀명"})

if "순예약인원" not in final_df.columns:
    possible_res_cols = [c for c in final_df.columns if "예약" in c and "인원" in c]
    if possible_res_cols:
        final_df = final_df.rename(columns={possible_res_cols[0]: "순예약인원"})

if "순예약인원" in final_df.columns:
    final_df["순예약인원"] = to_number(final_df["순예약인원"])
else:
    final_df["순예약인원"] = 0


for col in ["담당부서", "담당팀명", "대리점한글명"]:
    if col not in final_df.columns:
        final_df[col] = "미분류"

final_df["담당부서"] = final_df["담당부서"].astype(str).str.strip()
final_df["담당팀명"] = final_df["담당팀명"].astype(str).str.strip()
final_df["대리점한글명"] = final_df["대리점한글명"].astype(str).str.strip()


# =========================
# 사이드바 필터
# =========================
st.sidebar.title("필터")

dept_options = sorted(final_df["담당부서"].dropna().unique().tolist())
team_options = sorted(final_df["담당팀명"].dropna().unique().tolist())

selected_depts = st.sidebar.multiselect(
    "담당부서",
    options=dept_options,
    default=dept_options,
)

selected_teams = st.sidebar.multiselect(
    "담당팀명",
    options=team_options,
    default=team_options,
)

filtered_df = final_df.copy()

if selected_depts:
    filtered_df = filtered_df[filtered_df["담당부서"].isin(selected_depts)]

if selected_teams:
    filtered_df = filtered_df[filtered_df["담당팀명"].isin(selected_teams)]


# =========================
# 순위 및 리그 로직
# =========================
LEAGUE_THRESHOLD = 30


def clean_agency_name(name):
    return str(name).replace(" ", "").strip()


def format_final_rank(row):
    agency = clean_agency_name(row.get("대리점한글명", ""))
    if agency == clean_agency_name("대한여행사[AK분당점]"):
        return "성장률순위 1등"

    rank = row.get("모객순위", None)
    if pd.isna(rank):
        return ""

    try:
        return f"모객순위 {int(rank)}등"
    except Exception:
        return f"모객순위 {rank}등"


ranking_df = filtered_df.copy()
ranking_df = ranking_df.sort_values("순예약인원", ascending=False).reset_index(drop=True)
ranking_df["모객순위"] = ranking_df.index + 1
ranking_df["최종순위"] = ranking_df.apply(format_final_rank, axis=1)
ranking_df["리그"] = ranking_df["순예약인원"].apply(lambda x: "A리그" if x >= LEAGUE_THRESHOLD else "B리그")


# =========================
# 메인 화면
# =========================
st.title("HANA 유통점 휴일 프로모션 대시보드")
# 주간 업데이트 텍스트 제거 완료

st.markdown(
    """
    <div class="sub-box">
        <b>유통점 휴일 프로모션 개요</b><br>
        기간: 26년 4월 11일 ~ 6월 28일<br>
        기준: 유통점별 순예약인원 및 부서별 예약 현황
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================
# 핵심 지표
# =========================
total_reservations = int(filtered_df["순예약인원"].sum())
total_agencies = int(filtered_df["대리점한글명"].nunique())
a_league_count = int((ranking_df["리그"] == "A리그").sum())
b_league_count = int((ranking_df["리그"] == "B리그").sum())

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(
        f"""
        <div class="metric-card">
            <h3>총 순예약인원</h3>
            <p>{total_reservations:,}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col2:
    st.markdown(
        f"""
        <div class="metric-card">
            <h3>대리점 수</h3>
            <p>{total_agencies:,}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col3:
    st.markdown(
        f"""
        <div class="metric-card">
            <h3>A리그</h3>
            <p>{a_league_count:,}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col4:
    st.markdown(
        f"""
        <div class="metric-card">
            <h3>B리그</h3>
            <p>{b_league_count:,}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================
# 탭 구성
# =========================
tab_rank, tab_dept, tab_league, tab_raw = st.tabs(
    ["최종순위", "부서별 예약인원", "리그 구성비", "원본데이터"]
)


# =========================
# 최종순위 탭
# =========================
with tab_rank:
    st.markdown('<div class="section-title">최종순위</div>', unsafe_allow_html=True)

    display_cols = [
        "모객순위",
        "최종순위",
        "리그",
        "담당부서",
        "담당팀명",
        "대리점한글명",
        "순예약인원",
    ]

    display_cols = [c for c in display_cols if c in ranking_df.columns]
    display_df = ranking_df[display_cols].copy()

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
    )


# =========================
# 부서별 예약인원 탭
# =========================
with tab_dept:
    st.markdown('<div class="section-title">부서별 예약인원</div>', unsafe_allow_html=True)

    dept_summary = (
        filtered_df.groupby("담당부서", as_index=False)
        .agg(
            순예약인원=("순예약인원", "sum"),
            대리점수=("대리점한글명", "nunique"),
        )
        .sort_values("순예약인원", ascending=False)
    )

    fig = px.bar(
        dept_summary,
        x="담당부서",
        y="순예약인원",
        text="순예약인원",
        color="순예약인원",
        color_continuous_scale=["#14b8a6", "#7c3aed"],
    )

    fig.update_traces(texttemplate="%{text:,}", textposition="outside")

    fig.update_layout(
        height=430,
        margin=dict(l=20, r=20, t=30, b=20),
        xaxis_title="담당부서",
        yaxis_title="순예약인원",
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    try:
        if not participant_df.empty:
            participant = participant_df.copy()

            for col in ["담당부서", "담당팀명", "대리점한글명"]:
                if col in participant.columns:
                    participant[col] = participant[col].astype(str).str.strip()

            if "담당부서" in participant.columns and selected_depts:
                participant = participant[participant["담당부서"].isin(selected_depts)]

            if "담당팀명" in participant.columns and selected_teams:
                participant = participant[participant["담당팀명"].isin(selected_teams)]

            if "담당부서" in participant.columns:
                total_by_dept = (
                    participant.groupby("담당부서")
                    .agg(참여대리점수=("대리점한글명", "nunique"))
                    .reset_index()
                )
            else:
                total_by_dept = dept_summary[["담당부서", "대리점수"]].rename(
                    columns={"대리점수": "참여대리점수"}
                )
        else:
            total_by_dept = dept_summary[["담당부서", "대리점수"]].rename(
                columns={"대리점수": "참여대리점수"}
            )

        reserve_by_dept = (
            filtered_df[filtered_df["순예약인원"] > 0]
            .groupby("담당부서")
            .agg(예약발생대리점수=("대리점한글명", "nunique"))
            .reset_index()
        )

        dept_holiday_rate = total_by_dept.merge(
            reserve_by_dept,
            on="담당부서",
            how="left",
        )

        dept_holiday_rate["예약발생대리점수"] = (
            dept_holiday_rate["예약발생대리점수"].fillna(0).astype(int)
        )

        dept_holiday_rate["참여대리점수"] = (
            dept_holiday_rate["참여대리점수"].replace(0, pd.NA)
        )

        dept_holiday_rate["휴일예약발생률"] = (
            dept_holiday_rate["예약발생대리점수"]
            / dept_holiday_rate["참여대리점수"]
            * 100
        ).fillna(0)

        dept_holiday_rate["휴일예약발생률"] = dept_holiday_rate["휴일예약발생률"].round(1)

        dept_holiday_rate = dept_holiday_rate.sort_values(
            "휴일예약발생률",
            ascending=False,
        )

        with st.expander("부서별 휴일 예약 발생률 보기", expanded=True):
            st.dataframe(
                dept_holiday_rate,
                use_container_width=True,
                hide_index=True,
            )

    except Exception as e:
        st.warning(f"부서별 휴일 예약 발생률 계산 중 오류가 발생했습니다: {e}")


# =========================
# 리그 구성비 탭
# =========================
with tab_league:
    st.markdown('<div class="section-title">리그 구성비</div>', unsafe_allow_html=True)

    league_summary = (
        ranking_df.groupby("리그", as_index=False)
        .agg(
            대리점수=("대리점한글명", "nunique"),
            순예약인원=("순예약인원", "sum"),
        )
        .sort_values("리그")
    )

    fig_league = px.pie(
        league_summary,
        names="리그",
        values="대리점수",
        hole=0.45,
        color="리그",
        color_discrete_map={
            "A리그": "#7c3aed",
            "B리그": "#14b8a6",
        },
    )

    fig_league.update_layout(
        height=430,
        margin=dict(l=20, r=20, t=30, b=20),
        paper_bgcolor="white",
        plot_bgcolor="white",
    )

    st.plotly_chart(fig_league, use_container_width=True)

    st.dataframe(
        league_summary,
        use_container_width=True,
        hide_index=True,
    )


# =========================
# 원본데이터 탭
# =========================
with tab_raw:
    st.markdown('<div class="section-title">원본데이터</div>', unsafe_allow_html=True)

    st.dataframe(
        filtered_df,
        use_container_width=True,
        hide_index=True,
    )
