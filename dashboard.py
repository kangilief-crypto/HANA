import streamlit as st
import pandas as pd
import plotly.express as px

# ════════════════════════════════════════════
# 페이지 설정
# ════════════════════════════════════════════
st.set_page_config(
    page_title="휴일 프로모션",
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

    .stCaption, [data-testid="stCaptionContainer"] {{
        color: {TEXT_GRAY} !important;
        font-size: 14px !important;
    }}

    h2, h3 {{
        color: {TEXT_DARK} !important;
        font-weight: 700 !important;
        letter-spacing: -0.01em !important;
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

    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {{
        color: {BRAND_PURPLE} !important;
        border-bottom: 2px solid {BRAND_MINT};
        padding-bottom: 8px;
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

    [data-baseweb="select"] {{
        border-radius: 8px !important;
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

    .streamlit-expanderHeader {{
        background-color: {PURPLE_LIGHT} !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        color: {BRAND_PURPLE} !important;
    }}
</style>
""",
    unsafe_allow_html=True,
)

st.title("휴일 프로모션")
st.caption("주간 업데이트")

# ════════════════════════════════════════════
# 데이터 로드 + 전처리
# ════════════════════════════════════════════
@st.cache_data(ttl=3600)
def load_data():
    df = pd.read_excel("weekly_data.xlsx", header=1)

    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    if "담당부서" in df.columns:
        df["담당부서"] = df["담당부서"].ffill()

    if "담당팀명" in df.columns:
        df["담당팀명"] = df["담당팀명"].ffill()

    for col in ["순예약인원", "모객순위"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


@st.cache_data(ttl=3600)
def load_participation_data():
    df_list = pd.read_excel(
        "weekly_data.xlsx",
        sheet_name="참여대리점 리스트",
        header=1,
    )

    df_list = df_list.loc[:, ~df_list.columns.str.contains("^Unnamed")]

    if "담당부서" in df_list.columns:
        df_list["담당부서"] = df_list["담당부서"].ffill()

    if "담당팀명" in df_list.columns:
        df_list["담당팀명"] = df_list["담당팀명"].ffill()

    if "번호" in df_list.columns:
        df_list = df_list[df_list["번호"] != "번호"].copy()
        df_list["번호"] = pd.to_numeric(df_list["번호"], errors="coerce")
        df_list = df_list.dropna(subset=["번호"])

    if "비고" in df_list.columns:
        exclude_keywords = ["퇴점", "거래중지"]
        df_list = df_list[
            ~df_list["비고"].astype(str).str.contains(
                "|".join(exclude_keywords),
                na=False,
            )
        ].copy()

    df_zero = pd.read_excel(
        "weekly_data.xlsx",
        sheet_name="휴일예약 0명 대리점",
        header=1,
    )

    df_zero = df_zero.loc[:, ~df_zero.columns.str.contains("^Unnamed")]

    if "담당부서" in df_zero.columns:
        df_zero["담당부서"] = df_zero["담당부서"].ffill()

    return df_list, df_zero


try:
    df = load_data()
except FileNotFoundError:
    st.error("weekly_data.xlsx 파일을 찾을 수 없습니다. GitHub 저장소 최상단에 업로드해주세요.")
    st.stop()
except Exception as e:
    st.error(f"엑셀 파일을 불러오는 중 오류가 발생했습니다: {e}")
    st.stop()

# ════════════════════════════════════════════
# 사이드바 필터
# ════════════════════════════════════════════
st.sidebar.header("필터")

df_f = df.copy()

if "담당부서" in df.columns:
    부서목록 = sorted(df["담당부서"].dropna().unique())
    선택부서 = st.sidebar.multiselect(
        "담당부서",
        부서목록,
        default=부서목록,
    )
    df_f = df_f[df_f["담당부서"].isin(선택부서)]

if "담당팀명" in df.columns:
    팀목록 = sorted(df_f["담당팀명"].dropna().unique())
    선택팀 = st.sidebar.multiselect(
        "담당팀명",
        팀목록,
        default=팀목록,
    )
    df_f = df_f[df_f["담당팀명"].isin(선택팀)]

if "담당세일즈" in df.columns:
    담당자목록 = sorted(df_f["담당세일즈"].dropna().unique())
    선택담당자 = st.sidebar.multiselect(
        "담당세일즈",
        담당자목록,
        default=담당자목록,
    )
    df_f = df_f[df_f["담당세일즈"].isin(선택담당자)]

# ════════════════════════════════════════════
# KPI 카드
# ════════════════════════════════════════════
st.subheader("핵심 지표")

c1, c2, c3, c4 = st.columns(4)

if "대리점코드" in df_f.columns:
    c1.metric("총 대리점 수", f"{df_f['대리점코드'].nunique():,}")
else:
    c1.metric("총 대리점 수", f"{len(df_f):,}")

if "담당세일즈" in df_f.columns:
    c2.metric("총 담당자 수", f"{df_f['담당세일즈'].nunique():,}")
else:
    c2.metric("총 담당자 수", "0")

if "순예약인원" in df_f.columns:
    c3.metric("총 예약 인원", f"{df_f['순예약인원'].sum():,.0f}")
    c4.metric("평균 예약 인원", f"{df_f['순예약인원'].mean():,.1f}")
else:
    c3.metric("총 예약 인원", "0")
    c4.metric("평균 예약 인원", "0")

st.divider()

# ════════════════════════════════════════════
# 메인 탭
# ════════════════════════════════════════════
tab_a, tab_b, tab_dept, tab_league, tab_contract = st.tabs(
    [
        "A리그 (30명 이상)",
        "B리그 (30명 이상)",
        "부서별",
        "리그 구성비",
        "계약관계별",
    ]
)

필수컬럼 = {
    "대리점한글명",
    "순예약인원",
    "담당세일즈",
    "리그",
    "담당부서",
    "담당팀명",
}

if 필수컬럼.issubset(df_f.columns):
    대리점별 = (
        df_f.groupby(["리그", "대리점한글명"], as_index=False)
        .agg(
            순예약인원=("순예약인원", "sum"),
            담당부서=("담당부서", "first"),
            담당팀명=("담당팀명", "first"),
            담당세일즈=("담당세일즈", "first"),
        )
    )
    대리점별_30 = 대리점별[대리점별["순예약인원"] >= 30].copy()
else:
    대리점별 = pd.DataFrame()
    대리점별_30 = pd.DataFrame()


def render_league(league_name, color):
    if 대리점별.empty:
        st.warning("대리점별 분석에 필요한 컬럼이 부족합니다.")
        return

    league_df = 대리점별_30[대리점별_30["리그"] == league_name].copy()
    league_df = league_df.sort_values("순예약인원", ascending=False).reset_index(drop=True)

    league_all = df_f[df_f["리그"] == league_name].copy()
    league_all_grouped = (
        league_all.groupby("대리점한글명", as_index=False)
        .agg(순예약인원=("순예약인원", "sum"))
    )

    if len(league_all_grouped) == 0:
        st.info(f"{league_name}에 데이터가 없습니다.")
        return

    st.caption(f"{league_name} — 순예약인원 30명 이상 대리점 기준 차트 및 순위표")

    if len(league_df) == 0:
        st.info(f"{league_name}에 30명 이상인 대리점이 없습니다.")
    else:
        league_df["표시명"] = (
            league_df["대리점한글명"]
            + " ("
            + league_df["담당팀명"].astype(str)
            + " / "
            + league_df["담당세일즈"].astype(str)
            + ")"
        )

        높이 = max(400, len(league_df) * 28)

        fig = px.bar(
            league_df,
            x="순예약인원",
            y="표시명",
            orientation="h",
            text_auto=True,
            title=f"{league_name} — 순예약인원 30명 이상",
            color_discrete_sequence=[color],
        )

        fig.update_layout(
            height=높이,
            yaxis={"categoryorder": "total ascending", "title": ""},
            xaxis={"title": "순예약인원"},
            plot_bgcolor="white",
            paper_bgcolor="white",
        )

        st.plotly_chart(fig, use_container_width=True)

        st.markdown(f"**{league_name} 순위표**")

        표 = league_df[
            [
                "담당부서",
                "담당팀명",
                "대리점한글명",
                "순예약인원",
                "담당세일즈",
            ]
        ].copy()

        표.index = 표.index + 1
        표.index.name = "순위"

        st.dataframe(
            표,
            use_container_width=False,
            width=900,
            height=min(600, 40 + len(표) * 35),
            column_config={
                "순위": st.column_config.NumberColumn("순위", width="small"),
                "담당부서": st.column_config.TextColumn("담당부서", width="small"),
                "담당팀명": st.column_config.TextColumn("담당팀명", width="small"),
                "대리점한글명": st.column_config.TextColumn("대리점한글명", width="medium"),
                "순예약인원": st.column_config.NumberColumn(
                    "순예약인원",
                    width="small",
                    format="%d",
                ),
                "담당세일즈": st.column_config.TextColumn("담당세일즈", width="small"),
            },
        )

    st.markdown("---")
    st.markdown(f"##### {league_name} 전체 요약")

    cc1, cc2, cc3 = st.columns(3)
    cc1.metric(f"{league_name} 대리점 수", f"{len(league_all_grouped):,}")
    cc2.metric("총 예약 인원", f"{league_all_grouped['순예약인원'].sum():,.0f}")
    cc3.metric("평균 예약 인원", f"{league_all_grouped['순예약인원'].mean():,.1f}")


with tab_a:
    render_league("A리그", BRAND_PURPLE)

with tab_b:
    render_league("B리그", BRAND_MINT)

with tab_dept:
    if {"담당부서", "순예약인원"}.issubset(df_f.columns):
        부서별 = (
            df_f.groupby("담당부서", as_index=False)["순예약인원"]
            .sum()
            .sort_values("순예약인원", ascending=False)
        )

        fig = px.bar(
            부서별,
            x="담당부서",
            y="순예약인원",
            title="부서별 총 예약 인원",
            text_auto=True,
            color_discrete_sequence=[BRAND_PURPLE],
        )

        fig.update_layout(
            height=400,
            xaxis={"title": ""},
            yaxis={"title": "순예약인원"},
            plot_bgcolor="white",
            paper_bgcolor="white",
        )

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("부서별 차트에 필요한 컬럼이 부족합니다.")

with tab_league:
    if {"리그", "순예약인원"}.issubset(df_f.columns):
        리그별 = df_f.groupby("리그", as_index=False)["순예약인원"].sum()

        fig = px.pie(
            리그별,
            names="리그",
            values="순예약인원",
            title="리그별 예약 인원 구성비",
            hole=0.4,
            color_discrete_sequence=[BRAND_PURPLE, BRAND_MINT],
        )

        fig.update_layout(height=400, paper_bgcolor="white")

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("리그 구성비에 필요한 컬럼이 부족합니다.")

with tab_contract:
    if {"계약관계", "순예약인원"}.issubset(df_f.columns):
        계약별 = df_f.groupby("계약관계", as_index=False)["순예약인원"].sum()

        fig = px.pie(
            계약별,
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

        fig.update_layout(height=400, paper_bgcolor="white")

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("계약관계별 차트에 필요한 컬럼이 부족합니다.")

st.divider()

# ════════════════════════════════════════════
# 부서별 휴일 예약 발생률
# ════════════════════════════════════════════
with st.expander("부서별 휴일 예약 발생률 보기", expanded=False):
    try:
        df_list, df_zero = load_participation_data()

        if "담당부서" not in df_list.columns or "담당부서" not in df_zero.columns:
            st.warning("참여대리점 리스트 또는 휴일예약 0명 대리점 시트에 담당부서 컬럼이 없습니다.")
        else:
            부서_참여 = df_list.groupby("담당부서").size().reset_index(name="참여_대리점수")
            부서_미예약 = df_zero.groupby("담당부서").size().reset_index(name="미예약_대리점수")

            발생률표 = 부서_참여.merge(부서_미예약, on="담당부서", how="left")
            발생률표["미예약_대리점수"] = 발생률표["미예약_대리점수"].fillna(0).astype(int)
            발생률표["예약발생_대리점수"] = (
                발생률표["참여_대리점수"] - 발생률표["미예약_대리점수"]
            )
            발생률표["발생률(%)"] = (
                발생률표["예약발생_대리점수"] / 발생률표["참여_대리점수"] * 100
            ).round(1)

            발생률표 = 발생률표[발생률표["담당부서"] != "기타"].copy()
            발생률표 = 발생률표.sort_values("발생률(%)", ascending=False).reset_index(drop=True)

            총_참여 = int(발생률표["참여_대리점수"].sum())
            총_발생 = int(발생률표["예약발생_대리점수"].sum())
            총_미예약 = int(발생률표["미예약_대리점수"].sum())
            전체_발생률 = round(총_발생 / 총_참여 * 100, 1) if 총_참여 > 0 else 0

            st.caption(
                "기준: 참여대리점 리스트에서 비고가 퇴점 또는 거래중지인 대리점 제외 "
                "후, 휴일예약 0명 대리점 시트를 기준으로 산출"
            )

            fig = px.bar(
                발생률표,
                x="담당부서",
                y="발생률(%)",
                title="부서별 휴일 예약 발생률",
                text=발생률표["발생률(%)"].astype(str) + "%",
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

            표시표 = 발생률표[
                [
                    "담당부서",
                    "참여_대리점수",
                    "예약발생_대리점수",
                    "미예약_대리점수",
                    "발생률(%)",
                ]
            ].copy()

            표시표.index = 표시표.index + 1
            표시표.index.name = "순위"

            st.dataframe(
                표시표,
                use_container_width=False,
                width=800,
                column_config={
                    "순위": st.column_config.NumberColumn("순위", width="small"),
                    "담당부서": st.column_config.TextColumn("담당부서", width="medium"),
                    "참여_대리점수": st.column_config.NumberColumn(
                        "참여 대리점",
                        width="small",
                        format="%d",
                    ),
                    "예약발생_대리점수": st.column_config.NumberColumn(
                        "예약 발생",
                        width="small",
                        format="%d",
                    ),
                    "미예약_대리점수": st.column_config.NumberColumn(
                        "미예약",
                        width="small",
                        format="%d",
                    ),
                    "발생률(%)": st.column_config.ProgressColumn(
                        "발생률 (%)",
                        format="%.1f%%",
                        min_value=0,
                        max_value=100,
                        width="medium",
                    ),
                },
            )

            st.markdown("---")
            st.markdown("##### 전체 요약")

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("총 참여 대리점", f"{총_참여:,}")
            k2.metric("예약 발생", f"{총_발생:,}")
            k3.metric("미예약 대리점", f"{총_미예약:,}")
            k4.metric("전체 발생률", f"{전체_발생률}%")

    except Exception as e:
        st.error(f"부서별 발생률 데이터를 불러오는 중 오류가 발생했습니다: {e}")
        st.caption("엑셀 파일에 '참여대리점 리스트'와 '휴일예약 0명 대리점' 시트가 모두 있어야 합니다.")

st.divider()

# ════════════════════════════════════════════
# 원본 데이터 + 다운로드
# ════════════════════════════════════════════
st.subheader("원본 데이터")
st.dataframe(df_f, use_container_width=True, hide_index=True)

csv = df_f.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "CSV 다운로드",
    csv,
    "filtered_data.csv",
    "text/csv",
)
