import streamlit as st
import pandas as pd
import plotly.express as px

# ════════════════════════════════════════════
# 페이지 설정
# ════════════════════════════════════════════
st.set_page_config(
    page_title="유통점 휴일 프로모션",
    layout="wide",
)

# ════════════════════════════════════════════
# 비밀번호 보호
# ════════════════════════════════════════════
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "비밀번호를 입력하세요",
            type="password",
            on_change=password_entered,
            key="password",
        )
        return False

    if not st.session_state["password_correct"]:
        st.text_input(
            "비밀번호를 입력하세요",
            type="password",
            on_change=password_entered,
            key="password",
        )
        st.error("비밀번호가 틀렸습니다.")
        return False

    return True


if not check_password():
    st.stop()

# ════════════════════════════════════════════
# 브랜드 디자인 시스템
# ════════════════════════════════════════════
BRAND_PURPLE = "#5E2BB8"
BRAND_MINT = "#08D1D9"
PURPLE_LIGHT = "#F5F0FF"
MINT_LIGHT = "#E6FBFC"
TEXT_DARK = "#1F1B2E"
TEXT_GRAY = "#6B7280"
BORDER_LIGHT = "#E5E7EB"

st.markdown(
    f"""
<style>
    html, body, [class*="css"] {{
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, sans-serif;
    }}

    .stApp {{
        background-color: #FAFAFC;
    }}

    h1 {{
        color: {BRAND_PURPLE} !important;
        font-weight: 800 !important;
        letter-spacing: -0.02em !important;
        border-bottom: 3px solid {BRAND_MINT};
        padding-bottom: 12px !important;
        margin-bottom: 8px !important;
    }}

    h2, h3 {{
        color: {TEXT_DARK} !important;
        font-weight: 700 !important;
        letter-spacing: -0.01em !important;
    }}

    .stCaption, [data-testid="stCaptionContainer"] {{
        color: {TEXT_GRAY} !important;
        font-size: 14px !important;
    }}

    [data-testid="stMetric"] {{
        background: linear-gradient(135deg, {PURPLE_LIGHT} 0%, {MINT_LIGHT} 100%);
        border: 1px solid {BORDER_LIGHT};
        border-left: 4px solid {BRAND_PURPLE};
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 1px 3px rgba(94, 43, 184, 0.05);
    }}

    [data-testid="stMetricLabel"] {{
        color: {TEXT_GRAY} !important;
        font-weight: 600 !important;
        font-size: 13px !important;
    }}

    [data-testid="stMetricValue"] {{
        color: {BRAND_PURPLE} !important;
        font-weight: 800 !important;
        font-size: 28px !important;
    }}

    [data-testid="stSidebar"] {{
        background-color: #FFFFFF;
        border-right: 1px solid {BORDER_LIGHT};
    }}

    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px;
        background-color: transparent;
        border-bottom: 2px solid {BORDER_LIGHT};
    }}

    .stTabs [data-baseweb="tab"] {{
        background-color: transparent;
        color: {TEXT_GRAY};
        font-weight: 600;
        padding: 10px 20px;
        border-radius: 8px 8px 0 0;
        border: none;
    }}

    .stTabs [aria-selected="true"] {{
        background-color: {BRAND_PURPLE} !important;
        color: white !important;
    }}

    .stDataFrame {{
        border: 1px solid {BORDER_LIGHT};
        border-radius: 10px;
        overflow: hidden;
    }}

    [data-baseweb="tag"] {{
        background-color: {BRAND_MINT} !important;
        color: white !important;
        font-weight: 600 !important;
    }}

    hr {{
        border-color: {BRAND_MINT}33 !important;
        margin: 24px 0 !important;
    }}

    .stButton button {{
        background-color: {BRAND_PURPLE};
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        padding: 8px 20px;
    }}

    .stButton button:hover {{
        background-color: {BRAND_MINT};
        color: white;
    }}

    .overview-box {{
        background-color: #FFFFFF;
        border: 1px solid {BORDER_LIGHT};
        border-left: 6px solid {BRAND_PURPLE};
        border-radius: 12px;
        padding: 20px 24px;
        margin-top: 12px;
        margin-bottom: 22px;
    }}

    .overview-title {{
        color: {BRAND_PURPLE};
        font-size: 26px;
        font-weight: 800;
        margin-bottom: 14px;
    }}

    .overview-table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 15px;
    }}

    .overview-table th {{
        width: 110px;
        text-align: left;
        color: {BRAND_PURPLE};
        background-color: {PURPLE_LIGHT};
        padding: 10px 12px;
        border: 1px solid {BORDER_LIGHT};
        vertical-align: top;
    }}

    .overview-table td {{
        color: {TEXT_DARK};
        padding: 10px 12px;
        border: 1px solid {BORDER_LIGHT};
        line-height: 1.6;
        vertical-align: top;
        background-color: #FFFFFF;
    }}
</style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "weekly_data.xlsx"

# ════════════════════════════════════════════
# 공통 함수
# ════════════════════════════════════════════
def normalize_columns(df):
    df = df.copy()
    df.columns = df.columns.astype(str)
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    rename_map = {
        "담당팀": "담당팀명",
        "대리점명": "대리점한글명",
        "지급대상": "모객순위",
        "최종 순위": "최종순위",
        "최종순위": "최종순위",
    }

    df = df.rename(columns=rename_map)
    return df


def clean_sales_df(df):
    df = normalize_columns(df)

    for col in ["담당부서", "담당팀명", "리그"]:
        if col in df.columns:
            df[col] = df[col].ffill()

    if "번호" in df.columns:
        df = df[df["번호"] != "번호"].copy()
        df["번호"] = pd.to_numeric(df["번호"], errors="coerce")
        df = df.dropna(subset=["번호"])

    for col in ["순예약인원", "전분기 예약인원", "모객성장률", "지원금액", "최종순위", "모객순위"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "순예약인원" in df.columns:
        df = df.dropna(subset=["순예약인원"])

    return df


@st.cache_data(ttl=3600)
def load_excel_book():
    return pd.ExcelFile(EXCEL_FILE)


@st.cache_data(ttl=3600)
def load_final_data():
    df = pd.read_excel(EXCEL_FILE, sheet_name="최종실적", header=1)
    df = clean_sales_df(df)

    required = {
        "담당부서",
        "담당팀명",
        "담당세일즈",
        "대리점코드",
        "대리점한글명",
        "리그",
        "계약관계",
        "순예약인원",
    }

    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"최종실적 시트에 필수 컬럼이 없습니다: {', '.join(missing)}")

    if "최종순위" not in df.columns:
        df["최종순위"] = None

    if "지원금액" not in df.columns:
        df["지원금액"] = None

    return df


@st.cache_data(ttl=3600)
def load_dept_data():
    df = pd.read_excel(EXCEL_FILE, sheet_name="부서별 모객현황", header=1)
    df = clean_sales_df(df)

    required = {
        "담당부서",
        "담당팀명",
        "담당세일즈",
        "대리점코드",
        "대리점한글명",
        "리그",
        "계약관계",
        "순예약인원",
    }

    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"부서별 모객현황 시트에 필수 컬럼이 없습니다: {', '.join(missing)}")

    return df


@st.cache_data(ttl=3600)
def load_participation_list():
    df_list = pd.read_excel(EXCEL_FILE, sheet_name="참여대리점 리스트", header=1)
    df_list = normalize_columns(df_list)

    for col in ["담당부서", "담당팀명"]:
        if col in df_list.columns:
            df_list[col] = df_list[col].ffill()

    if "번호" in df_list.columns:
        df_list = df_list[df_list["번호"] != "번호"].copy()
        df_list["번호"] = pd.to_numeric(df_list["번호"], errors="coerce")
        df_list = df_list.dropna(subset=["번호"])

    if "비고" in df_list.columns:
        exclude_keywords = ["퇴점", "거래중지"]
        df_list = df_list[
            ~df_list["비고"].astype(str).str.contains("|".join(exclude_keywords), na=False)
        ].copy()

    return df_list


try:
    final_df = load_final_data()
    dept_df = load_dept_data()
    participant_df = load_participation_list()
except FileNotFoundError:
    st.error("weekly_data.xlsx 파일을 찾을 수 없습니다. GitHub 저장소 최상단에 업로드해주세요.")
    st.stop()
except Exception as e:
    st.error(f"엑셀 데이터를 불러오는 중 오류가 발생했습니다: {e}")
    st.stop()

# 기본 분석 데이터
# - 리그 순위표: 최종실적
# - 부서별 예약/발생률: 부서별 모객현황
df = final_df.copy()

# ════════════════════════════════════════════
# 제목 및 개요
# ════════════════════════════════════════════
st.title("유통점 휴일 프로모션")
st.caption("주간 업데이트")

st.markdown(
    """
<div class="overview-box">
    <div class="overview-title">[유통점 휴일 프로모션]</div>
    <table class="overview-table">
        <tr>
            <th>기간</th>
            <td>26년 4월 11일 ~ 6월 28일 기간 내 휴일(주말+공휴일) 최초 예약일 기준</td>
        </tr>
        <tr>
            <th>조건</th>
            <td>P,B,U 속성(주말 유통점 현장 예약에 한함)</td>
        </tr>
        <tr>
            <th>제외</th>
            <td>코브랜드, 닷컴, 랜드온리, 유아/아동 제외</td>
        </tr>
        <tr>
            <th>지급</th>
            <td>
                1) 리그별 모객 순위(1위 ~ 10위), 모객성장률(1위~10위) 대리점 대상 VI 지급<br>
                2) 최소 모객인원 30명 허들 미달성시 지급대상에서 제외<br>
                3) 모객 동일 수 예약 시 상위매출 순서대로 순위 인정
            </td>
        </tr>
        <tr>
            <th>대상</th>
            <td>공식인증예약센터 212개 (유통점 172개_현대 제외 + 직계약 유통점 15개 + 부서별 선정대리점 5개)</td>
        </tr>
    </table>
</div>
""",
    unsafe_allow_html=True,
)

# ════════════════════════════════════════════
# 사이드바 필터
# ════════════════════════════════════════════
st.sidebar.header("필터")
st.sidebar.caption("리그 순위표 기준: 최종실적")
st.sidebar.caption("부서별 예약/발생률 기준: 부서별 모객현황")

df_f = df.copy()
dept_f = dept_df.copy()

if "담당부서" in df.columns:
    부서목록 = sorted(df["담당부서"].dropna().astype(str).unique())
    선택부서 = st.sidebar.multiselect("담당부서", 부서목록, default=부서목록)
    df_f = df_f[df_f["담당부서"].astype(str).isin(선택부서)]
    dept_f = dept_f[dept_f["담당부서"].astype(str).isin(선택부서)]

if "담당팀명" in df.columns:
    팀목록 = sorted(df_f["담당팀명"].dropna().astype(str).unique())
    선택팀 = st.sidebar.multiselect("담당팀명", 팀목록, default=팀목록)
    df_f = df_f[df_f["담당팀명"].astype(str).isin(선택팀)]
    dept_f = dept_f[dept_f["담당팀명"].astype(str).isin(선택팀)]

if "담당세일즈" in df.columns:
    담당자목록 = sorted(df_f["담당세일즈"].dropna().astype(str).unique())
    선택담당자 = st.sidebar.multiselect("담당세일즈", 담당자목록, default=담당자목록)
    df_f = df_f[df_f["담당세일즈"].astype(str).isin(선택담당자)]
    dept_f = dept_f[dept_f["담당세일즈"].astype(str).isin(선택담당자)]

# ════════════════════════════════════════════
# 핵심 지표
# ════════════════════════════════════════════
st.subheader("핵심 지표")

c1, c2, c3, c4 = st.columns(4)

c1.metric("총 대리점 수", f"{df_f['대리점코드'].nunique():,}")
c2.metric("총 담당자 수", f"{df_f['담당세일즈'].nunique():,}")
c3.metric("총 예약 인원", f"{df_f['순예약인원'].sum():,.0f}")
c4.metric("평균 예약 인원", f"{df_f['순예약인원'].mean():,.1f}")

st.divider()

# ════════════════════════════════════════════
# 메인 탭
# ════════════════════════════════════════════
tab_a, tab_b, tab_dept, tab_league, tab_contract, tab_raw = st.tabs(
    ["A리그 (30명 이상)", "B리그 (30명 이상)", "부서별 예약인원", "리그 구성비", "계약관계별", "원본데이터"]
)

# 리그 순위표 데이터: 최종실적 기준
대리점별 = (
    df_f.groupby(["리그", "대리점한글명"], as_index=False)
    .agg(
        순예약인원=("순예약인원", "sum"),
        담당부서=("담당부서", "first"),
        담당팀명=("담당팀명", "first"),
        담당세일즈=("담당세일즈", "first"),
        대리점코드=("대리점코드", "first"),
        계약관계=("계약관계", "first"),
        최종순위=("최종순위", "first"),
        지원금액=("지원금액", "first"),
    )
)

대리점별_30 = 대리점별[대리점별["순예약인원"] >= 30].copy()


def render_league(league_name, color):
    league_df = 대리점별_30[대리점별_30["리그"] == league_name].copy()
    league_df = league_df.sort_values(
        ["최종순위", "순예약인원"],
        ascending=[True, False],
        na_position="last",
    ).reset_index(drop=True)

    league_all = 대리점별[대리점별["리그"] == league_name].copy()

    if len(league_all) == 0:
        st.info(f"{league_name}에 데이터가 없습니다.")
        return

    st.caption(f"{league_name} — 순예약인원 30명 이상 허들 달성 대리점 기준")

    if len(league_df) == 0:
        st.info(f"{league_name}에 30명 이상인 대리점이 없습니다.")
    else:
        league_df["표시명"] = (
            league_df["대리점한글명"].astype(str)
            + " ("
            + league_df["담당팀명"].astype(str)
            + " / "
            + league_df["담당세일즈"].astype(str)
            + ")"
        )

        chart_height = max(400, len(league_df) * 28)

        fig = px.bar(
            league_df.sort_values("순예약인원", ascending=True),
            x="순예약인원",
            y="표시명",
            orientation="h",
            text_auto=True,
            title=f"{league_name} — 순예약인원 30명 이상",
            color_discrete_sequence=[color],
        )

        fig.update_layout(
            height=chart_height,
            yaxis={"title": ""},
            xaxis={"title": "순예약인원"},
            plot_bgcolor="white",
            paper_bgcolor="white",
        )

        st.plotly_chart(fig, use_container_width=True)

        st.markdown(f"**{league_name} 순위표**")

                # 최종순위 표시 문구 변환

        # 기본: 최종순위 1~10 → 모객순위 1등 ~ 모객순위 10등

        # 예외: 대한여행사[AK분당점] → 성장률순위 1등

        table_source = league_df.copy()



        def format_final_rank(row):

            agency_name = str(row.get("대리점한글명", ""))

            final_rank = row.get("최종순위", None)



            if agency_name == "대한여행사[AK 분당점]":

                return "성장률순위 1등"



            if pd.notna(final_rank):

                try:

                    rank_num = int(final_rank)

                    if 1 <= rank_num <= 10:

                        return f"모객순위 {rank_num}등"

                except Exception:

                    pass



            return ""



        table_source["최종순위표시"] = table_source.apply(format_final_rank, axis=1)



        table = table_source[

            [

                "담당부서",

                "담당팀명",

                "대리점코드",

                "대리점한글명",

                "순예약인원",

                "담당세일즈",

                "계약관계",

                "최종순위표시",

                "지원금액",

            ]

        ].copy()



        table = table.rename(columns={"최종순위표시": "최종순위"})



        table.index = table.index + 1

        table.index.name = "순번"



                # 순위표 가운데 정렬 HTML 테이블 출력
        display_table = table.copy().reset_index()

        if "지원금액" in display_table.columns:
            display_table["지원금액"] = display_table["지원금액"].apply(
                lambda x: f"{int(x):,}" if pd.notna(x) else ""
            )

        if "순예약인원" in display_table.columns:
            display_table["순예약인원"] = display_table["순예약인원"].apply(
                lambda x: f"{int(x):,}" if pd.notna(x) else ""
            )

        table_html = display_table.to_html(
            index=False,
            escape=False,
            classes="rank-table",
        )

        st.markdown(
            """
            <style>
                .rank-table {
                    width: 100%;
                    border-collapse: collapse;
                    background-color: #FFFFFF;
                    font-size: 14px;
                    border: 1px solid #E5E7EB;
                }

                .rank-table thead th {
                    background-color: #F5F0FF;
                    color: #5E2BB8;
                    font-weight: 700;
                    text-align: center !important;
                    padding: 10px 8px;
                    border: 1px solid #E5E7EB;
                    white-space: nowrap;
                }

                .rank-table tbody td {
                    text-align: center !important;
                    padding: 9px 8px;
                    border: 1px solid #E5E7EB;
                    color: #1F1B2E;
                    vertical-align: middle;
                    white-space: nowrap;
                }

                .rank-table tbody tr:nth-child(even) {
                    background-color: #FAFAFC;
                }

                .rank-table tbody tr:hover {
                    background-color: #E6FBFC;
                }
            </style>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(table_html, unsafe_allow_html=True)




    st.markdown("---")
    st.markdown(f"##### {league_name} 전체 요약")

    cc1, cc2, cc3 = st.columns(3)
    cc1.metric(f"{league_name} 대리점 수", f"{len(league_all):,}")
    cc2.metric("총 예약 인원", f"{league_all['순예약인원'].sum():,.0f}")
    cc3.metric("평균 예약 인원", f"{league_all['순예약인원'].mean():,.1f}")


with tab_a:
    render_league("A리그", BRAND_PURPLE)

with tab_b:
    render_league("B리그", BRAND_MINT)

with tab_dept:
    dept = (
        dept_f.groupby("담당부서", as_index=False)["순예약인원"]
        .sum()
        .sort_values("순예약인원", ascending=False)
    )

    fig = px.bar(
        dept,
        x="담당부서",
        y="순예약인원",
        title="부서별 총 예약 인원",
        text_auto=True,
        color_discrete_sequence=[BRAND_PURPLE],
    )

    fig.update_layout(
        height=420,
        xaxis={"title": ""},
        yaxis={"title": "순예약인원"},
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    with st.expander("부서별 휴일 예약 발생률 보기", expanded=False):
        try:
            # 분모: 참여대리점 리스트
            participant = participant_df.copy()

            # 사이드바 필터와 동일하게 반영
            if "담당부서" in participant.columns:
                participant = participant[participant["담당부서"].astype(str).isin(선택부서)]

            if "담당팀명" in participant.columns:
                participant = participant[participant["담당팀명"].astype(str).isin(선택팀)]

            total_by_dept = (
                participant.groupby("담당부서")
                .agg(참여_대리점수=("대리점코드", "nunique"))
                .reset_index()
            )

            # 분자: 부서별 모객현황 내 예약 발생 대리점
            active_base = dept_f[dept_f["순예약인원"] > 0].copy()

            active_by_dept = (
                active_base.groupby("담당부서")
                .agg(예약발생_대리점수=("대리점코드", "nunique"))
                .reset_index()
            )

            rate = total_by_dept.merge(active_by_dept, on="담당부서", how="left")
            rate["예약발생_대리점수"] = rate["예약발생_대리점수"].fillna(0).astype(int)
            rate["미예약_대리점수"] = rate["참여_대리점수"] - rate["예약발생_대리점수"]
            rate["발생률(%)"] = (rate["예약발생_대리점수"] / rate["참여_대리점수"] * 100).round(1)

            rate = rate[rate["담당부서"] != "기타"].copy()
            rate = rate.sort_values("발생률(%)", ascending=False).reset_index(drop=True)

            total_cnt = int(rate["참여_대리점수"].sum())
            active_cnt = int(rate["예약발생_대리점수"].sum())
            zero_cnt = int(rate["미예약_대리점수"].sum())
            total_rate = round(active_cnt / total_cnt * 100, 1) if total_cnt > 0 else 0

            st.caption(
                "기준: 참여대리점 리스트의 전체 대리점 수 대비, 부서별 모객현황에서 순예약인원이 1명 이상 발생한 대리점 비율 "
                "(비고가 퇴점 또는 거래중지인 대리점 제외)"
            )

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("총 참여 대리점", f"{total_cnt:,}")
            k2.metric("예약 발생", f"{active_cnt:,}")
            k3.metric("미예약 대리점", f"{zero_cnt:,}")
            k4.metric("전체 발생률", f"{total_rate}%")

            fig = px.bar(
                rate,
                x="담당부서",
                y="발생률(%)",
                title="부서별 휴일 예약 발생률",
                text=rate["발생률(%)"].astype(str) + "%",
                color_discrete_sequence=[BRAND_PURPLE],
            )

            fig.update_traces(textposition="outside")
            fig.update_layout(
                height=420,
                plot_bgcolor="white",
                paper_bgcolor="white",
                yaxis={"range": [0, 100], "title": "발생률 (%)"},
                xaxis={"title": ""},
            )

            st.plotly_chart(fig, use_container_width=True)

            st.markdown("**부서별 상세**")

            display = rate[
                ["담당부서", "참여_대리점수", "예약발생_대리점수", "미예약_대리점수", "발생률(%)"]
            ].copy()

            display.index = display.index + 1
            display.index.name = "순위"

            st.dataframe(
                display,
                use_container_width=False,
                width=850,
                column_config={
                    "순위": st.column_config.NumberColumn("순위", width="small"),
                    "담당부서": st.column_config.TextColumn("담당부서", width="medium"),
                    "참여_대리점수": st.column_config.NumberColumn("참여 대리점", width="small", format="%d"),
                    "예약발생_대리점수": st.column_config.NumberColumn("예약 발생", width="small", format="%d"),
                    "미예약_대리점수": st.column_config.NumberColumn("미예약", width="small", format="%d"),
                    "발생률(%)": st.column_config.ProgressColumn(
                        "발생률 (%)",
                        format="%.1f%%",
                        min_value=0,
                        max_value=100,
                        width="medium",
                    ),
                },
            )

        except Exception as e:
            st.error(f"부서별 발생률 데이터를 불러오는 중 오류가 발생했습니다: {e}")

with tab_league:
    league = dept_f.groupby("리그", as_index=False)["순예약인원"].sum()

    fig = px.pie(
        league,
        names="리그",
        values="순예약인원",
        title="리그별 예약 인원 구성비",
        hole=0.4,
        color_discrete_sequence=[BRAND_PURPLE, BRAND_MINT],
    )

    fig.update_layout(height=420, paper_bgcolor="white")
    st.plotly_chart(fig, use_container_width=True)

with tab_contract:
    contract = dept_f.groupby("계약관계", as_index=False)["순예약인원"].sum()

    fig = px.pie(
        contract,
        names="계약관계",
        values="순예약인원",
        title="계약관계별 예약 인원 구성비",
        hole=0.4,
        color_discrete_sequence=[
            BRAND_PURPLE,
            BRAND_MINT,
            "#9B6FE0",
            "#5DD9DD",
            "#7C4FCE",
        ],
    )

    fig.update_layout(height=420, paper_bgcolor="white")
    st.plotly_chart(fig, use_container_width=True)

with tab_raw:
    st.subheader("원본데이터")
    st.caption("기준 시트: 최종실적")

    raw_tab1, raw_tab2, raw_tab3 = st.tabs(["최종실적", "부서별 모객현황", "참여대리점 리스트"])

    with raw_tab1:
        st.dataframe(final_df, use_container_width=True, hide_index=True)
        csv = final_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("최종실적 CSV 다운로드", csv, "final_data.csv", "text/csv")

    with raw_tab2:
        st.dataframe(dept_df, use_container_width=True, hide_index=True)
        csv = dept_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("부서별 모객현황 CSV 다운로드", csv, "dept_data.csv", "text/csv")

    with raw_tab3:
        st.dataframe(participant_df, use_container_width=True, hide_index=True)
        csv = participant_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("참여대리점 리스트 CSV 다운로드", csv, "participant_data.csv", "text/csv")
