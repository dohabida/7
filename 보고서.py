# -*- coding: utf-8 -*-
import io
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(
    page_title="사이버범죄 추이 대시보드 (2014-2020)",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Helper: Plot Export
# -----------------------------
def fig_to_png_download_button(fig, filename_prefix="figure"):
    """Plotly Figure를 PNG로 변환하여 다운로드 버튼 제공"""
    try:
        import kaleido  # noqa: F401  # ensure installed
        png_bytes = fig.to_image(format="png", scale=2)
        st.download_button(
            label="⬇️ PNG로 저장",
            data=png_bytes,
            file_name=f"{filename_prefix}.png",
            mime="image/png",
            use_container_width=True
        )
    except Exception as e:
        st.info("PNG 저장을 위해 'kaleido' 설치가 필요합니다:  \n`pip install -U kaleido`")
        st.caption(f"오류 메시지: {e}")

# -----------------------------
# Sidebar: File Upload & Year Selection
# -----------------------------
st.sidebar.title("📥 데이터 업로드")
uploaded = st.sidebar.file_uploader(
    "엑셀 파일(.xlsx)을 업로드하세요\n(시트명 예: '경찰청_월별 사이버범죄 발생건수와 검거건수 현황_2020', '경찰청_연도별 사이버 범죄 통계 현황_20200831')",
    type=["xlsx"]
)

# -----------------------------
# Read & Prepare Data
# -----------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def load_data(file) -> dict:
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names

    # 예상 시트명 패턴
    monthly_sheet = next((s for s in sheets if "월별" in s), None)
    yearly_sheet  = next((s for s in sheets if "연도별" in s or "통계 현황" in s), None)

    if monthly_sheet is None:
        raise ValueError("월별 시트를 찾을 수 없습니다. 시트명에 '월별'이 포함되도록 해주세요.")
    monthly_df = pd.read_excel(file, sheet_name=monthly_sheet)

    yearly_df = None
    if yearly_sheet is not None:
        yearly_df = pd.read_excel(file, sheet_name=yearly_sheet)

    return {
        "monthly_raw": monthly_df,
        "yearly_raw": yearly_df,
        "monthly_sheet": monthly_sheet,
        "yearly_sheet": yearly_sheet,
    }

def tidy_monthly(df: pd.DataFrame) -> pd.DataFrame:
    """
    입력 예시:
      연도 | 구분 | 1월 | 2월 | ... | 12월
      구분: 발생건수 / 검거건수
    출력을 long화:
      연도, 월(1~12), 발생건수, 검거건수, 검거율
    """
    df = df.copy()
    # 월 컬럼 후보
    month_cols = [c for c in df.columns if any(str(m) in str(c) for m in range(1,13)) or ("월" in str(c))]
    id_cols = [c for c in df.columns if c not in month_cols]
    # melt
    m = df.melt(id_vars=id_cols, value_vars=month_cols, var_name="월", value_name="값")
    # 월 문자 정리 (예: "1월" -> 1)
    m["월"] = m["월"].astype(str).str.replace("월", "", regex=False).str.strip()
    m["월"] = pd.to_numeric(m["월"], errors="coerce")
    # 발생/검거 pivot
    if "구분" not in m.columns:
        raise ValueError("월별 시트에 '구분' 컬럼(발생건수/검거건수 구분)이 필요합니다.")

    pv = m.pivot_table(
        index=["연도", "월"],
        columns="구분",
        values="값",
        aggfunc="sum"
    ).reset_index()

    # 컬럼 표준화
    col_map = {}
    for c in pv.columns:
        if isinstance(c, tuple):
            col_map[c] = c[-1]
    pv.rename(columns=col_map, inplace=True)

    # 예상 명칭 보정
    # (발생건수/검거건수 외 명칭일 경우에도 최대한 매칭)
    def pick_col(cols, keys):
        for k in keys:
            for c in cols:
                if k in str(c):
                    return c
        return None

    occur_col = pick_col(pv.columns, ["발생", "발생건수"])
    arrest_col = pick_col(pv.columns, ["검거", "검거건수"])

    if occur_col is None or arrest_col is None:
        raise ValueError("발생/검거 컬럼을 식별할 수 없습니다. '발생건수'와 '검거건수' 구조를 확인해주세요.")

    pv = pv.rename(columns={occur_col: "발생건수", arrest_col: "검거건수"})
    pv["검거율"] = np.where(pv["발생건수"] > 0, pv["검거건수"] / pv["발생건수"], np.nan)

    # 정렬
    pv = pv.sort_values(["연도", "월"]).reset_index(drop=True)
    return pv

def yearly_totals_from_monthly(monthly_long: pd.DataFrame) -> pd.DataFrame:
    g = monthly_long.groupby("연도", as_index=False).agg(
        발생건수=("발생건수", "sum"),
        검거건수=("검거건수", "sum"),
    )
    g["검거율"] = np.where(g["발생건수"] > 0, g["검거건수"] / g["발생건수"], np.nan)
    return g

def tidy_yearly_category(df: pd.DataFrame) -> pd.DataFrame:
    """
    연도별 시트(범죄유형 x 발생/검거) → 긴 형태:
    ['연도','구분','유형','값'] + 유형별 검거율 계산
    """
    if df is None:
        return None
    t = df.copy()
    # 기본 식별 컬럼
    id_cols = [c for c in ["연도","구분"] if c in t.columns]
    val_cols = [c for c in t.columns if c not in id_cols]
    long = t.melt(id_vars=id_cols, value_vars=val_cols, var_name="유형", value_name="값")
    # 발생/검거 pivot하여 유형별 검거율
    pv = long.pivot_table(index=["연도","유형"], columns="구분", values="값", aggfunc="sum").reset_index()
    # 이름 보정
    def pick_col(cols, keys):
        for k in keys:
            for c in cols:
                if k in str(c):
                    return c
        return None
    occur_col = pick_col(pv.columns, ["발생", "발생건수"])
    arrest_col = pick_col(pv.columns, ["검거", "검거건수"])

    if occur_col is None or arrest_col is None:
        # 연도별 시트 구조가 상이한 경우 None 반환
        return None

    pv = pv.rename(columns={occur_col:"발생건수", arrest_col:"검거건수"})
    pv["검거율"] = np.where(pv["발생건수"]>0, pv["검거건수"]/pv["발생건수"], np.nan)
    return pv

# -----------------------------
# UI: Load data or show example hint
# -----------------------------
st.title("🛡️ 2014~2020 사이버범죄 연·월별 추이 대시보드")
st.caption("경영진 의사결정 지원용: Executive Summary → 현황분석 → 문제도출 → 해결방안 흐름에 맞춘 핵심 지표 탐색")

if not uploaded:
    st.info("왼쪽 사이드바에서 엑셀(.xlsx)을 업로드해 주세요.")
    st.stop()

# Load
try:
    data = load_data(uploaded)
    monthly_raw = data["monthly_raw"]
    yearly_raw  = data["yearly_raw"]
except Exception as e:
    st.error(f"데이터 읽기 중 오류: {e}")
    st.stop()

# Tidy
try:
    monthly_long = tidy_monthly(monthly_raw)  # [연도, 월, 발생건수, 검거건수, 검거율]
    yearly_totals = yearly_totals_from_monthly(monthly_long)  # [연도, 발생건수, 검거건수, 검거율]
    yearly_categories = tidy_yearly_category(yearly_raw)      # [연도, 유형, 발생건수, 검거건수, 검거율] or None
except Exception as e:
    st.error(f"전처리 과정 오류: {e}")
    st.stop()

# Sidebar: Year selector
years = sorted(monthly_long["연도"].unique())
default_year = max(yers) if (yers:=years) else None  # safe trick for default
sel_year = st.sidebar.slider("연도 선택", int(min(years)), int(max(years)), int(default_year))

# -----------------------------
# Executive Summary (3p 대응)
# -----------------------------
st.header("1) Executive Summary")
c1, c2, c3 = st.columns(3)

# KPI Cards (선택 연도)
y_df = monthly_long[monthly_long["연도"] == sel_year]
yr_total_occ = int(y_df["발생건수"].sum())
yr_total_arr = int(y_df["검거건수"].sum())
yr_rate = (yr_total_arr / yr_total_occ) if yr_total_occ > 0 else np.nan

c1.metric("총 발생 (연)", f"{yr_total_occ:,}")
c2.metric("총 검거 (연)", f"{yr_total_arr:,}")
c3.metric("검거율 (연)", f"{yr_rate*100:,.1f}%")

# 연도별 총 발생/검거 추이 (라인)
fig_line = go.Figure()
fig_line.add_trace(go.Scatter(
    x=yearly_totals["연도"], y=yearly_totals["발생건수"],
    mode="lines+markers", name="발생건수"
))
fig_line.add_trace(go.Scatter(
    x=yearly_totals["연도"], y=yearly_totals["검거건수"],
    mode="lines+markers", name="검거건수"
))
fig_line.update_layout(
    title="연도별 총 발생/검거 추이",
    xaxis_title="연도", yaxis_title="건수",
    hovermode="x unified", margin=dict(t=60, l=10, r=10, b=10)
)
st.plotly_chart(fig_line, use_container_width=True)
fig_to_png_download_button(fig_line, f"연도별_발생_검거_추이")

# 연도별 검거율 바차트
fig_rate = px.bar(
    yearly_totals,
    x="연도", y=(yearly_totals["검거율"]*100),
    labels={"y":"검거율(%)"},
    title="연도별 검거율"
)
fig_rate.update_yaxes(title="검거율(%)", rangemode="tozero")
st.plotly_chart(fig_rate, use_container_width=True)
fig_to_png_download_button(fig_rate, f"연도별_검거율")

st.divider()

# -----------------------------
# 현황 분석 (6p 대응)
# -----------------------------
st.header("2) 현황 분석")

tab1, tab2, tab3, tab4 = st.tabs([
    "연도별 추세", "월별 패턴", "범주(유형) 현황", "저조 유형"
])

with tab1:
    # 발생-검거 Gap (라인, 18p 중 2.1 확장)
    yt = yearly_totals.copy()
    yt["Gap(발생-검거)"] = yt["발생건수"] - yt["검거건수"]
    fig_gap = px.line(
        yt, x="연도", y=["발생건수","검거건수","Gap(발생-검거)"],
        title="연도별 발생·검거·Gap 추이"
    )
    st.plotly_chart(fig_gap, use_container_width=True)
    fig_to_png_download_button(fig_gap, f"연도별_발생_검거_Gap")

with tab2:
    # 월별 평균 발생 추이 (전체연도 평균)
    m_avg = monthly_long.groupby("월", as_index=False)["발생건수"].mean()
    fig_mavg = px.line(m_avg, x="월", y="발생건수", markers=True,
                       title="월별 평균 발생 추이 (전체연도)")
    st.plotly_chart(fig_mavg, use_container_width=True)
    fig_to_png_download_button(fig_mavg, f"월별_평균_발생_추이")

    # 연도 x 월 검거율 히트맵
    pivot = monthly_long.pivot_table(index="연도", columns="월", values="검거율", aggfunc="mean")
    fig_hm = px.imshow(
        (pivot*100).round(1),
        aspect="auto", color_continuous_scale="Blues",
        labels=dict(color="검거율(%)"),
        title="연도 x 월 검거율 히트맵"
    )
    st.plotly_chart(fig_hm, use_container_width=True)
    fig_to_png_download_button(fig_hm, f"연도x월_검거율_히트맵")

with tab3:
    if yearly_categories is not None and sel_year in yearly_categories["연도"].unique():
        yc_year = yearly_categories[yearly_categories["연도"] == sel_year].copy()
        # 파이차트(발생 비중)
        fig_pie = px.pie(
            yc_year, values="발생건수", names="유형",
            title=f"{sel_year}년 범주별 발생 비중"
        )
        st.plotly_chart(fig_pie, use_container_width=True)
        fig_to_png_download_button(fig_pie, f"{sel_year}년_범주별_발생_비중")

        # 유형별 검거율 바차트
        fig_cat_rate = px.bar(
            yc_year.sort_values("검거율"),
            x="유형", y=yc_year["검거율"]*100,
            title=f"{sel_year}년 유형별 검거율 (오름차순)",
            labels={"y":"검거율(%)"}
        )
        fig_cat_rate.update_yaxes(title="검거율(%)", rangemode="tozero")
        fig_cat_rate.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(fig_cat_rate, use_container_width=True)
        fig_to_png_download_button(fig_cat_rate, f"{sel_year}년_유형별_검거율")
    else:
        st.info("연도별 범주(유형) 데이터가 없거나 선택 연도에 해당 데이터가 없습니다. (연도별 시트 구조를 확인하세요)")

with tab4:
    if yearly_categories is not None:
        # 최근 연도(데이터에 존재하는 최대 연도) 기준 저조 유형 TOP 5
        last_year = int(yearly_categories["연도"].max())
        yc_last = yearly_categories[yearly_categories["연도"] == last_year].copy()
        yc_last = yc_last[yc_last["발생건수"] > 0]
        worst5 = yc_last.sort_values("검거율").head(5)

        c1, c2 = st.columns([2,1])
        with c1:
            fig_worst = px.bar(
                worst5, x="유형", y=worst5["검거율"]*100,
                title=f"{last_year}년 검거율 저조 유형 TOP 5",
                labels={"y":"검거율(%)"}
            )
            fig_worst.update_yaxes(range=[0, max(10, (worst5['검거율']*100).max()+5)])
            fig_worst.update_layout(xaxis_tickangle=-30)
            st.plotly_chart(fig_worst, use_container_width=True)
            fig_to_png_download_button(fig_worst, f"{last_year}년_검거율_저조_유형_TOP5")

        with c2:
            st.subheader("🔎 시사점")
            st.markdown(
                "- **저검거율 유형**은 집중대응 필요 (전담팀·OSINT·플랫폼 협력 강화)\n"
                "- **증거 휘발성 높은 유형** 우선 확보 프로토콜 정비\n"
                "- **광고/유통 채널 차단** 등 비수사·행정 지원 병행"
            )
    else:
        st.info("유형별 데이터가 없어 저조 유형 분석을 건너뜁니다.")

st.divider()

# -----------------------------
# 문제 도출 (4p 대응)
# -----------------------------
st.header("3) 문제 도출")
c1, c2 = st.columns(2)

# 3.1 발생은 급증·검거는 미흡(연도별 추세의 메시지 카드)
with c1:
    # 최근 3개 연도 비교
    yt = yearly_totals.sort_values("연도")
    if len(yt) >= 3:
        tail = yt.tail(3)
        msg = (
            f"최근 3개 연도 발생 증가율: "
            f"{(tail.iloc[-1]['발생건수'] - tail.iloc[0]['발생건수'])/max(1,tail.iloc[0]['발생건수']):.1%}\n\n"
            f"최근 3개 연도 검거 증가율: "
            f"{(tail.iloc[-1]['검거건수'] - tail.iloc[0]['검거건수'])/max(1,tail.iloc[0]['검거건수']):.1%}"
        )
    else:
        msg = "데이터 연도 수가 적어 추세 해석이 제한됩니다."
    st.markdown(f"**발생 급증 vs 검거 정체**  \n{msg}")

# 3.2/3.3 월별 자원 불균형 (히트맵에서 낮은 월 식별)
with c2:
    low_months = (
        monthly_long
        .groupby("월")["검거율"].mean()
        .sort_values()
        .head(3)
        .index
        .tolist()
    )
    st.markdown(
        f"**월별 집중 발생/낮은 검거율 시기**  \n"
        f"평균 검거율 하위 월(Top3): **{', '.join(map(str, low_months))}월**  \n"
        f"→ 이 구간에 자원 재배치 필요"
    )

st.divider()

# -----------------------------
# 해결 방안 (5~6p 대응)
# -----------------------------
st.header("4) 해결 방안")
st.markdown("""
**4.1 AI 기반 탐지 고도화**: 패턴 학습 및 이상감지 → 초기 경보 민감도 조정  
**4.2 유형 맞춤 전략**: 도박/음란물/해킹 등 특성별 전담 플레이북  
**4.3 자원 최적화**: 월별 패턴과 저검거율 시간대에 인력/예산 배분  
**4.4 협력 강화**: 플랫폼·금융사·국제 공조 채널 표준화  
**4.5 KPI 로드맵**: 연도별 검거율 **90%** 달성 단계(단기: 데이터 정합성, 중기: 자동화, 장기: 예측형 배치)
""")

# KPI 로드맵 간단 시뮬레이션 (선형 가정)
roadmap_c1, roadmap_c2 = st.columns([2,1])
with roadmap_c1:
    base = yearly_totals.sort_values("연도")
    if not base.empty:
        last_year = int(base["연도"].max())
        last_rate = float(base[base["연도"]==last_year]["검거율"].iloc[0])
        target = 0.90
        future_years = [last_year+i for i in range(1,6)]
        step = (target - last_rate)/len(future_years)
        sim_years = list(base["연도"]) + future_years
        sim_rates = list(base["검거율"]) + [min(1.0, last_rate + step*i) for i in range(1, len(future_years)+1)]

        fig_kpi = go.Figure()
        fig_kpi.add_trace(go.Scatter(x=base["연도"], y=base["검거율"]*100,
                                     mode="lines+markers", name="실적"))
        fig_kpi.add_trace(go.Scatter(x=future_years, y=[r*100 for r in sim_rates[-len(future_years):]],
                                     mode="lines+markers", name="로드맵(가정)"))
        fig_kpi.add_hline(y=90, line_dash="dot", annotation_text="목표 90%")
        fig_kpi.update_layout(title="검거율 90% 달성 로드맵(단순 시뮬레이션)",
                              xaxis_title="연도", yaxis_title="검거율(%)")
        st.plotly_chart(fig_kpi, use_container_width=True)
        fig_to_png_download_button(fig_kpi, f"검거율_90_로드맵")
    else:
        st.info("연도 데이터가 없어 로드맵 시뮬레이션을 생략합니다.")

with roadmap_c2:
    st.subheader("🎯 투자 우선순위")
    st.markdown(
        "- **데이터 파이프라인 정합성 개선** (단기)\n"
        "- **패턴탐지/링크분석 자동화** (중기)\n"
        "- **예측형 배치·실시간 공조** (장기)"
    )

st.divider()

# -----------------------------
# Download: 가공 데이터셋
# -----------------------------
st.subheader("📦 가공 데이터 다운로드")
proc1 = monthly_long.copy()
proc2 = yearly_totals.copy()
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    proc1.to_excel(writer, index=False, sheet_name="월별_long")
    proc2.to_excel(writer, index=False, sheet_name="연도별_totals")
    if yearly_categories is not None:
        yearly_categories.to_excel(writer, index=False, sheet_name="연도별_유형별")
st.download_button(
    "엑셀(가공데이터) 다운로드",
    data=buf.getvalue(),
    file_name="사이버범죄_가공데이터.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.caption("※ 차트는 상단 PNG 저장 버튼으로 슬라이드에 바로 삽입할 수 있습니다.")
