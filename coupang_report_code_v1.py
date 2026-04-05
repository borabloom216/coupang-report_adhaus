import streamlit as st
import pandas as pd
import re
from datetime import datetime
from collections import defaultdict

# ─────────────────────────────────────────
# 상수
# ─────────────────────────────────────────
ROAS_THRESHOLD = 300

SURFACE_MAP = {
    "비검색": "비검색영역",
    "검색":   "검색영역",
    "어드미니스": "어드미니스플러스",
}

SURFACE_LABEL = {
    "비검색영역":    "캠페인의 비검색 영역",
    "검색영역":      "캠페인의 검색 영역",
    "어드미니스플러스": "캠페인의 어드미니스 영역",
    "전체영역":      "캠페인의 전체 영역",
}


# ─────────────────────────────────────────
# 숫자 유틸
# ─────────────────────────────────────────
def fmt(n):
    try:
        return f"{int(n):,}"
    except Exception:
        return "0"

def roas(sales, cost):
    return round(sales / cost * 100) if cost else 0

def cpc(cost, clicks):
    return round(cost / clicks) if clicks else 0


# ─────────────────────────────────────────
# 지면 분류
# ─────────────────────────────────────────
def classify_surface(x):
    x = str(x).strip()
    if "비검색" in x:
        return "비검색영역"
    if "검색" in x:
        return "검색영역"
    if "어드미니스" in x or "Product Ad" in x:
        return "어드미니스플러스"
    return "기타"


# ─────────────────────────────────────────
# 파일명 파싱
# ─────────────────────────────────────────
def parse_filename(filename):
    """
    형식: {업체코드}_{보고서종류}_{시작일}_{종료일}.xlsx
    예)  A00536370_pa_total_campaign_20260401_20260401.xlsx
    반환: dict 또는 None
    """
    name = re.sub(r"\.xlsx?$", "", filename, flags=re.IGNORECASE)
    m = re.match(r"^([A-Z]\d+)_(pa_[a-z_]+)_(\d{8})_(\d{8})$", name)
    if not m:
        return None
    code, report_type, start_str, end_str = m.groups()
    start = datetime.strptime(start_str, "%Y%m%d")
    end   = datetime.strptime(end_str,   "%Y%m%d")
    delta = (end - start).days

    if delta == 0:
        period = "daily"
    elif delta <= 7:
        period = "weekly"
    else:
        period = "monthly"

    return {
        "code":        code,
        "report_type": report_type,   # pa_total_campaign | pa_daily_keyword
        "start":       start,
        "end":         end,
        "period":      period,
    }


# ─────────────────────────────────────────
# DataFrame 전처리
# ─────────────────────────────────────────

# 전각 괄호·공백 등 표기 차이를 흡수하는 컬럼명 정규화 맵
_COL_NORMALIZE = {
    "총 상품매출액(1일)":  ["총 상품매출액(1일)", "총상품매출액(1일)",
                          "총 상품매출액（1일）", "총상품매출액（1일）"],
    "광고비":             ["광고비"],
    "클릭수":             ["클릭수"],
    "광고 노출 지면":     ["광고 노출 지면", "광고노출지면"],
    "광고지표 상품명":    ["광고지표 상품명", "광고지표상품명"],
    "광고지표 옵션ID":    ["광고지표 옵션ID", "광고지표옵션ID", "광고지표 옵션id"],
    "캠페인명":           ["캠페인명"],
    "키워드":             ["키워드"],
}

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """컬럼명 앞뒤 공백 제거 후, 표기 변형을 표준 이름으로 통일한다."""
    df = df.copy()
    df.columns = df.columns.str.strip()          # 앞뒤 공백 제거
    rename = {}
    for standard, variants in _COL_NORMALIZE.items():
        for v in variants:
            if v in df.columns and v != standard:
                rename[v] = standard
    if rename:
        df = df.rename(columns=rename)
    return df

def prep(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    for col in ["광고비", "총 상품매출액(1일)", "클릭수"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "광고 노출 지면" in df.columns:
        df["광고 노출 지면"] = df["광고 노출 지면"].astype(str).str.strip()
        df["_지면"] = df["광고 노출 지면"].apply(classify_surface)
    if "키워드" in df.columns:
        df["키워드"] = df["키워드"].astype(str).str.replace(" ", "").str.lower()
    return df


# ─────────────────────────────────────────
# 상품명 분석
# ─────────────────────────────────────────
def analyze_products(df: pd.DataFrame):
    key_cols = ["광고지표 상품명", "광고지표 옵션ID"]
    missing = [c for c in key_cols if c not in df.columns]
    if missing:
        return pd.DataFrame(), pd.DataFrame()

    g = (
        df.groupby(key_cols)
        .agg({"광고비": "sum", "총 상품매출액(1일)": "sum", "클릭수": "sum"})
        .reset_index()
    )
    g["CPC"]  = g.apply(lambda r: cpc(r["광고비"], r["클릭수"]), axis=1)
    g["ROAS"] = g.apply(lambda r: roas(r["총 상품매출액(1일)"], r["광고비"]), axis=1)

    top_sales = (
        g[g["총 상품매출액(1일)"] > 0]
        .sort_values("총 상품매출액(1일)", ascending=False)
        .head(3)
    )
    top_loss = (
        g[g["ROAS"] < ROAS_THRESHOLD]
        .sort_values(["광고비", "ROAS"], ascending=[False, True])
        .head(3)
    )
    return top_sales, top_loss


# ─────────────────────────────────────────
# 키워드 분석
# ─────────────────────────────────────────
def analyze_keywords(df: pd.DataFrame):
    if "키워드" not in df.columns:
        return pd.DataFrame(), pd.DataFrame()

    filtered = df[~df["키워드"].isin(["-", "—", "–", ""])]
    g = (
        filtered.groupby("키워드")
        .agg({"광고비": "sum", "총 상품매출액(1일)": "sum", "클릭수": "sum"})
        .reset_index()
    )
    g["CPC"]  = g.apply(lambda r: cpc(r["광고비"], r["클릭수"]), axis=1)
    g["ROAS"] = g.apply(lambda r: roas(r["총 상품매출액(1일)"], r["광고비"]), axis=1)

    top_sales = (
        g[g["총 상품매출액(1일)"] > 0]
        .sort_values("총 상품매출액(1일)", ascending=False)
        .head(3)
    )
    top_loss = (
        g[g["ROAS"] < ROAS_THRESHOLD]
        .sort_values(["광고비", "ROAS"], ascending=[False, True])
        .head(3)
    )
    return top_sales, top_loss


# ─────────────────────────────────────────
# 캠페인 분석
# ─────────────────────────────────────────
def analyze_campaigns(df: pd.DataFrame):
    """
    매출 원인
      이익 지면(ROAS≥300 & 매출>0 & 광고비>0) 개수에 따라:
        2개 이상 → 캠페인 전체 합산
        정확히 1개 → 해당 지면
        0개 → 제외

    손실 원인
      이익 지면 존재 여부에 따라:
        이익 지면 있음 → 손실 지면 중 광고비 최대 1개 (동률 시 ROAS 낮은 순)
        이익 지면 없음 → 캠페인 전체 합산
        손실 지면 없음 → 제외
    """
    if "캠페인명" not in df.columns or "_지면" not in df.columns:
        return pd.DataFrame(), pd.DataFrame()

    results_sales = []
    results_loss  = []

    for camp_name, camp_df in df.groupby("캠페인명"):
        # 지면별 집계
        sg = (
            camp_df.groupby("_지면")
            .agg({"광고비": "sum", "총 상품매출액(1일)": "sum", "클릭수": "sum"})
            .reset_index()
        )
        sg["ROAS"] = sg.apply(lambda r: roas(r["총 상품매출액(1일)"], r["광고비"]), axis=1)

        profit = sg[
            (sg["ROAS"] >= ROAS_THRESHOLD) &
            (sg["총 상품매출액(1일)"] > 0) &
            (sg["광고비"] > 0)
        ]

        # ── 매출 원인 ──
        n_profit = len(profit)
        if n_profit >= 2:
            tc = camp_df["광고비"].sum()
            ts = camp_df["총 상품매출액(1일)"].sum()
            tk = camp_df["클릭수"].sum()
            results_sales.append({
                "캠페인명": camp_name,
                "지면라벨": SURFACE_LABEL["전체영역"],
                "광고비": tc, "총 상품매출액(1일)": ts, "클릭수": tk,
                "CPC": cpc(tc, tk), "ROAS": roas(ts, tc),
            })
        elif n_profit == 1:
            r_ = profit.iloc[0]
            results_sales.append({
                "캠페인명": camp_name,
                "지면라벨": SURFACE_LABEL.get(r_["_지면"], r_["_지면"]),
                "광고비": r_["광고비"], "총 상품매출액(1일)": r_["총 상품매출액(1일)"],
                "클릭수": r_["클릭수"],
                "CPC": cpc(r_["광고비"], r_["클릭수"]), "ROAS": r_["ROAS"],
            })
        # n_profit == 0 → 제외

        # ── 손실 원인 ──
        active = sg[sg["광고비"] > 0]
        loss   = active[active["ROAS"] < ROAS_THRESHOLD]

        if loss.empty:
            continue  # 손실 지면 없음 → 제외

        if n_profit > 0:
            # 이익 지면 있음 → 손실 지면 중 광고비 최대 1개 선택
            picked = loss.sort_values(["광고비", "ROAS"], ascending=[False, True]).iloc[0]
            results_loss.append({
                "캠페인명": camp_name,
                "지면라벨": SURFACE_LABEL.get(picked["_지면"], picked["_지면"]),
                "광고비": picked["광고비"], "총 상품매출액(1일)": picked["총 상품매출액(1일)"],
                "클릭수": picked["클릭수"],
                "CPC": cpc(picked["광고비"], picked["클릭수"]), "ROAS": picked["ROAS"],
            })
        else:
            # 모든 지면 ROAS < 300 → 캠페인 전체 합산
            tc = camp_df["광고비"].sum()
            ts = camp_df["총 상품매출액(1일)"].sum()
            tk = camp_df["클릭수"].sum()
            results_loss.append({
                "캠페인명": camp_name,
                "지면라벨": SURFACE_LABEL["전체영역"],
                "광고비": tc, "총 상품매출액(1일)": ts, "클릭수": tk,
                "CPC": cpc(tc, tk), "ROAS": roas(ts, tc),
            })

    sales_df = pd.DataFrame(results_sales) if results_sales else pd.DataFrame()
    loss_df  = pd.DataFrame(results_loss)  if results_loss  else pd.DataFrame()

    if not sales_df.empty:
        sales_df = sales_df.sort_values("총 상품매출액(1일)", ascending=False).head(3)
    if not loss_df.empty:
        loss_df = loss_df.sort_values(["광고비", "ROAS"], ascending=[False, True]).head(3)

    return sales_df, loss_df


# ─────────────────────────────────────────
# 보고서 섹션 빌더
# ─────────────────────────────────────────
def build_summary(df: pd.DataFrame, title: str = "광고 전체 요약") -> str:
    tc = df["광고비"].sum()
    ts = df["총 상품매출액(1일)"].sum()
    tr = roas(ts, tc)

    lines = [
        f"📊 {title}",
        f"광고비    : {fmt(tc)} 원",
        f"광고매출 : {fmt(ts)} 원",
        f"ROAS     : {tr} %",
    ]

    if "_지면" in df.columns:
        for surf, label in [
            ("검색영역",        "🔹 검색지면"),
            ("비검색영역",      "🔹 비검색지면"),
            ("어드미니스플러스", "🔹 어드미니스 플러스"),
        ]:
            sub = df[df["_지면"] == surf]
            if sub.empty:
                continue
            c = sub["광고비"].sum()
            s = sub["총 상품매출액(1일)"].sum()
            lines += [
                "",
                label,
                f"광고비    : {fmt(c)} 원",
                f"광고매출 : {fmt(s)} 원",
                f"ROAS     : {roas(s, c)} %",
            ]

    return "\n".join(lines)


def build_product_section(df: pd.DataFrame) -> str:
    top_sales, top_loss = analyze_products(df)
    lines = ["✅ 주요 매출 원인 (상품명 기준)"]
    if top_sales.empty:
        lines.append("  해당 없음")
    else:
        for _, row in top_sales.iterrows():
            lines += [
                f"  {row['광고지표 상품명']}",
                f"  (옵션ID : {int(row['광고지표 옵션ID'])})",
                f"  CPC {fmt(row['CPC'])} / 광고비 {fmt(row['광고비'])} / ROAS {int(row['ROAS'])} %",
                "",
            ]
    lines.append("")
    lines.append("⚠️ 주요 손실 원인 (상품명 기준)")
    if top_loss.empty:
        lines.append("  해당 없음")
    else:
        for _, row in top_loss.iterrows():
            lines += [
                f"  {row['광고지표 상품명']}",
                f"  (옵션ID : {int(row['광고지표 옵션ID'])})",
                f"  CPC {fmt(row['CPC'])} / 광고비 {fmt(row['광고비'])} / ROAS {int(row['ROAS'])} %",
                "",
            ]
    return "\n".join(lines)


def build_keyword_section(df: pd.DataFrame) -> str:
    top_sales, top_loss = analyze_keywords(df)
    lines = ["✅ 주요 매출 원인 (키워드 기준)"]
    if top_sales.empty:
        lines.append("  해당 없음")
    else:
        for _, row in top_sales.iterrows():
            lines += [
                f"  (키워드) {row['키워드']}",
                f"  CPC {fmt(row['CPC'])} / 광고비 {fmt(row['광고비'])} / ROAS {int(row['ROAS'])} %",
                "",
            ]
    lines.append("")
    lines.append("⚠️ 주요 손실 원인 (키워드 기준)")
    if top_loss.empty:
        lines.append("  해당 없음")
    else:
        for _, row in top_loss.iterrows():
            lines += [
                f"  (키워드) {row['키워드']}",
                f"  CPC {fmt(row['CPC'])} / 광고비 {fmt(row['광고비'])} / ROAS {int(row['ROAS'])} %",
                "",
            ]
    return "\n".join(lines)


def build_campaign_section(df: pd.DataFrame) -> str:
    top_sales, top_loss = analyze_campaigns(df)
    lines = ["✅ 주요 매출 원인 (캠페인 기준)"]
    if top_sales.empty:
        lines.append("  해당 없음")
    else:
        for _, row in top_sales.iterrows():
            lines += [
                f"  {row['캠페인명']}",
                f"  {row['지면라벨']}",
                f"  CPC {fmt(row['CPC'])} / 광고비 {fmt(row['광고비'])} / ROAS {int(row['ROAS'])} %",
                "",
            ]
    lines.append("")
    lines.append("⚠️ 주요 손실 원인 (캠페인 기준)")
    if top_loss.empty:
        lines.append("  해당 없음")
    else:
        for _, row in top_loss.iterrows():
            lines += [
                f"  {row['캠페인명']}",
                f"  {row['지면라벨']}",
                f"  CPC {fmt(row['CPC'])} / 광고비 {fmt(row['광고비'])} / ROAS {int(row['ROAS'])} %",
                "",
            ]
    return "\n".join(lines)


def build_comparison_delta(df_prev, df_curr) -> str:
    lines = ["📈 주요 지표 변화 (전기 → 당기)"]
    for col, name in [("광고비", "광고비"), ("총 상품매출액(1일)", "광고매출")]:
        pv = df_prev[col].sum()
        cv = df_curr[col].sum()
        diff = cv - pv
        pct  = round(diff / pv * 100) if pv else 0
        arrow = "▲" if diff > 0 else ("▼" if diff < 0 else "─")
        lines.append(f"  {name}: {fmt(pv)} → {fmt(cv)} ({arrow}{abs(pct)}%)")

    pr = roas(df_prev["총 상품매출액(1일)"].sum(), df_prev["광고비"].sum())
    cr = roas(df_curr["총 상품매출액(1일)"].sum(), df_curr["광고비"].sum())
    dr = cr - pr
    arrow = "▲" if dr > 0 else ("▼" if dr < 0 else "─")
    lines.append(f"  ROAS: {pr}% → {cr}% ({arrow}{abs(dr)}%p)")
    return "\n".join(lines)


# ─────────────────────────────────────────
# 완성 보고서 생성
# ─────────────────────────────────────────
def report_daily(df: pd.DataFrame, meta: dict, company: str) -> str:
    date_str = meta["end"].strftime("%m월 %d일")
    sep = "=" * 40
    parts = [
        sep,
        f"일시 : {date_str}   |   업체명 : {company}",
        sep,
        "안녕하세요.",
        "전일 광고 동향 간략 보고 드립니다.",
        f"일시 : {date_str}",
        "",
        build_summary(df),
        "",
        build_campaign_section(df),
        "",
        build_product_section(df),
    ]
    return "\n".join(parts)


def report_weekly_comparison(
    df_prev, df_curr, meta_prev, meta_curr, company: str
) -> str:
    lp = f"{meta_prev['start'].strftime('%m/%d')}~{meta_prev['end'].strftime('%m/%d')}"
    lc = f"{meta_curr['start'].strftime('%m/%d')}~{meta_curr['end'].strftime('%m/%d')}"
    sep = "=" * 40
    parts = [
        sep,
        f"업체명 : {company}",
        f"전전주({lp}) vs 전주({lc}) 광고 비교 보고",
        sep,
        "",
        f"── 전전주 ({lp}) ──",
        build_summary(df_prev),
        "",
        build_campaign_section(df_prev),
        "",
        build_product_section(df_prev),
        "",
        f"── 전주 ({lc}) ──",
        build_summary(df_curr),
        "",
        build_campaign_section(df_curr),
        "",
        build_product_section(df_curr),
        "",
        build_comparison_delta(df_prev, df_curr),
    ]
    return "\n".join(parts)


def report_monthly_comparison(
    df_prev, df_curr, meta_prev, meta_curr, company: str
) -> str:
    lp = meta_prev["start"].strftime("%Y년 %m월")
    lc = meta_curr["start"].strftime("%Y년 %m월")
    sep = "=" * 40
    parts = [
        sep,
        f"업체명 : {company}",
        f"{lp} vs {lc} 광고 비교 보고",
        sep,
        "",
        f"── {lp} ──",
        build_summary(df_prev),
        "",
        build_keyword_section(df_prev),
        "",
        build_product_section(df_prev),
        "",
        f"── {lc} ──",
        build_summary(df_curr),
        "",
        build_keyword_section(df_curr),
        "",
        build_product_section(df_curr),
        "",
        build_comparison_delta(df_prev, df_curr),
    ]
    return "\n".join(parts)


# ─────────────────────────────────────────
# Streamlit 앱
# ─────────────────────────────────────────
def load_df(uploaded_file) -> pd.DataFrame:
    return prep(pd.read_excel(uploaded_file, engine="openpyxl"))


def main():
    st.set_page_config(page_title="쿠팡 광고 분석 보고", layout="wide")
    st.title("📦 쿠팡 광고 간략 보고 생성기")

    # ── 업체코드 매핑 (선택) ──
    with st.sidebar:
        st.header("⚙️ 설정")
        map_file = st.file_uploader("업체코드 리스트 (선택)", type=["xlsx"])
        company_map: dict = {}
        if map_file:
            try:
                cdf = pd.read_excel(map_file)
                company_map = dict(zip(cdf.iloc[:, 0].astype(str), cdf.iloc[:, 1]))
            except Exception as e:
                st.warning(f"업체코드 파일 오류: {e}")

        st.markdown("---")
        st.markdown("""
**지원 보고서 유형**

| 파일 종류 | 기간 | 결과 |
|---|---|---|
| `pa_total_campaign` | 1일 | 일간 보고 |
| `pa_total_campaign` × 2 | 주간 | 전전주/전주 비교 |
| `pa_daily_keyword` × 2 | 월간 | 전전달/전달 비교 |
        """)

    # ── 파일 업로드 ──
    st.header("① 파일 업로드")
    uploaded = st.file_uploader(
        "분석할 엑셀 파일을 업로드하세요 (여러 파일 동시 가능)",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info("파일을 업로드하면 기간을 자동 감지해 보고서를 생성합니다.")
        return

    # ── 파일명 파싱 ──
    parsed = []
    for f in uploaded:
        meta = parse_filename(f.name)
        if meta is None:
            st.warning(f"⚠️ 파일명 형식 불일치 (건너뜀): `{f.name}`")
            continue
        meta["file"]    = f
        meta["company"] = company_map.get(meta["code"], meta["code"])
        parsed.append(meta)

    if not parsed:
        st.error("유효한 파일이 없습니다. 파일명 형식을 확인하세요.")
        return

    # ── 업체별 그룹핑 ──
    by_company: dict = defaultdict(list)
    for p in parsed:
        by_company[p["code"]].append(p)

    st.header("② 보고서 결과")

    for code, files in by_company.items():
        company = files[0]["company"]
        st.subheader(f"🏢 {company}  ({code})")

        campaign_files = sorted(
            [f for f in files if f["report_type"] == "pa_total_campaign"],
            key=lambda x: x["start"],
        )
        keyword_files = sorted(
            [f for f in files if f["report_type"] == "pa_daily_keyword"],
            key=lambda x: x["start"],
        )

        # ── 일간 보고 ──
        daily_files = [f for f in campaign_files if f["period"] == "daily"]
        for meta in daily_files:
            label = meta["end"].strftime("%m/%d") + " 일간 보고"
            with st.expander(f"📄 {label}", expanded=True):
                df = load_df(meta["file"])

                # 필수 컬럼 누락 확인
                required = ["광고비", "총 상품매출액(1일)", "클릭수"]
                missing = [c for c in required if c not in df.columns]
                if missing:
                    st.error(f"필수 컬럼을 찾을 수 없습니다: {missing}")
                    st.info(f"실제 파일의 컬럼명: {list(df.columns)}")
                    continue

                report = report_daily(df, meta, company)
                st.text(report)
                fname = f"일간보고_{company}_{meta['end'].strftime('%Y%m%d')}.txt"
                st.download_button(
                    "💾 다운로드", report.encode("utf-8"),
                    file_name=fname, mime="text/plain", key=fname,
                )

        # ── 전전주/전주 비교 ──
        weekly_files = [f for f in campaign_files if f["period"] == "weekly"]
        if len(weekly_files) >= 2:
            prev_meta, curr_meta = weekly_files[-2], weekly_files[-1]
            lp = f"{prev_meta['start'].strftime('%m/%d')}~{prev_meta['end'].strftime('%m/%d')}"
            lc = f"{curr_meta['start'].strftime('%m/%d')}~{curr_meta['end'].strftime('%m/%d')}"
            label = f"전전주({lp}) / 전주({lc}) 비교"
            with st.expander(f"📄 {label}", expanded=True):
                df_prev = load_df(prev_meta["file"])
                df_curr = load_df(curr_meta["file"])
                report = report_weekly_comparison(df_prev, df_curr, prev_meta, curr_meta, company)
                st.text(report)
                fname = (
                    f"주간비교_{company}_"
                    f"{prev_meta['start'].strftime('%m%d')}_"
                    f"{curr_meta['end'].strftime('%m%d')}.txt"
                )
                st.download_button(
                    "💾 다운로드", report.encode("utf-8"),
                    file_name=fname, mime="text/plain", key=fname,
                )

        # ── 전전달/전달 비교 ──
        monthly_files = [f for f in keyword_files if f["period"] == "monthly"]
        if len(monthly_files) >= 2:
            prev_meta, curr_meta = monthly_files[-2], monthly_files[-1]
            lp = prev_meta["start"].strftime("%Y년 %m월")
            lc = curr_meta["start"].strftime("%Y년 %m월")
            label = f"{lp} / {lc} 비교"
            with st.expander(f"📄 {label}", expanded=True):
                df_prev = load_df(prev_meta["file"])
                df_curr = load_df(curr_meta["file"])
                report = report_monthly_comparison(df_prev, df_curr, prev_meta, curr_meta, company)
                st.text(report)
                fname = (
                    f"월간비교_{company}_"
                    f"{prev_meta['start'].strftime('%Y%m')}_"
                    f"{curr_meta['start'].strftime('%Y%m')}.txt"
                )
                st.download_button(
                    "💾 다운로드", report.encode("utf-8"),
                    file_name=fname, mime="text/plain", key=fname,
                )


if __name__ == "__main__":
    main()
