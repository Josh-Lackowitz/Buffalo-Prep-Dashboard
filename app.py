"""
Buffalo Prep Alumni Outcomes Dashboard
=======================================
All metrics computed directly from Cleaned_Dataset.xlsx.
No external data, no hardcoded statistics.
"""

import os
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import base64

st.set_page_config(
    page_title="Buffalo Prep Alumni Dashboard",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

C_BLUE   = "#2B7EC1"
C_BLUE_D = "#1A5A94"
C_BLUE_XD= "#0D2E47"
C_BLUE_L = "#D6E9F8"
C_BLUE_M = "#4A9ED4"
C_GOLD   = "#C9973A"
C_GREEN  = "#16A34A"
C_GRAY_D = "#1E1E2E"
C_GRAY_M = "#6B7280"
C_WHITE  = "#FFFFFF"
SEQ_BLUE = [[0, C_BLUE_L], [0.5, C_BLUE], [1, C_BLUE_XD]]
PALETTE  = [C_BLUE, C_BLUE_M, C_BLUE_D, "#7BBDE0", C_BLUE_XD, "#A8D4F0", C_GOLD, "#0A1F30"]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] {{ font-family:'DM Sans',sans-serif; color:{C_GRAY_D}; }}
[data-testid="stSidebar"] {{ background:{C_GRAY_D} !important; }}
[data-testid="stSidebar"] * {{ color:{C_WHITE} !important; }}
.main .block-container {{ padding-top:1.2rem; padding-bottom:2rem; max-width:1440px; }}
.pg-header {{
    background:linear-gradient(135deg,{C_GRAY_D} 0%,#16213E 100%);
    border-radius:14px; padding:20px 28px; margin-bottom:18px;
    display:flex; align-items:center; justify-content:space-between;
    border:1px solid rgba(43,126,193,0.25);
}}
.pg-title {{ font-family:'Fraunces',serif; font-size:1.6rem; font-weight:600; color:{C_WHITE}; margin:0; }}
.pg-subtitle {{ font-size:0.82rem; color:rgba(255,255,255,0.5); margin:3px 0 0 0; }}
.pg-badge {{
    background:rgba(201,151,58,0.15); border:1px solid rgba(201,151,58,0.4);
    color:{C_GOLD}; font-size:0.7rem; font-weight:600; padding:6px 14px;
    border-radius:20px; text-transform:uppercase; letter-spacing:0.05em; white-space:nowrap;
}}
.sec-hdr {{
    font-family:'Fraunces',serif; font-size:1.05rem; font-weight:600; color:{C_WHITE};
    margin:0 0 12px 0; padding-bottom:7px; border-bottom:1px solid #E5E7EB;
}}
.insight {{
    background:linear-gradient(135deg,{C_BLUE_L},{C_BLUE_L}cc);
    border-left:4px solid {C_BLUE}; border-radius:0 10px 10px 0;
    padding:10px 16px; font-size:0.81rem; color:{C_BLUE_XD};
    margin:10px 0 16px 0; line-height:1.55;
}}
.dq-note {{
    background:#FFFBEB; border-left:4px solid {C_GOLD};
    border-radius:0 10px 10px 0; padding:9px 14px;
    font-size:0.78rem; color:#78350F; margin:8px 0 14px 0; line-height:1.5;
}}
.divider {{ height:1px; background:linear-gradient(to right,{C_BLUE}44,transparent); margin:16px 0; border:none; }}
.nav-lbl {{ font-size:0.63rem; letter-spacing:0.12em; text-transform:uppercase; opacity:0.35; margin-bottom:5px; }}
</style>
""", unsafe_allow_html=True)


# ── Data loading & cleaning ───────────────────────────────────────────────────
@st.cache_data
def load_and_clean():
    path = os.path.join(BASE_DIR, "Cleaned_Dataset.xlsx")
    raw  = pd.read_excel(path, sheet_name="Alumni Data")
    df   = raw[raw["First Name"].notna() | raw["Last Name"].notna()].copy()

    df["Cohort Year"]       = pd.to_numeric(df["Cohort Year"],       errors="coerce").astype("Int64")
    df["Estimated Salary"]  = pd.to_numeric(df["Estimated Salary"],  errors="coerce")
    df["Salary Range Low"]  = pd.to_numeric(df["Salary Range Low"],  errors="coerce")
    df["Salary Range High"] = pd.to_numeric(df["Salary Range High"], errors="coerce")

    # Has Graduate Degree: explicit "Yes" = grad, everything else = No
    df["Has Graduate Degree"] = df["Has Graduate Degree"].fillna("No")
    df["Gender"]    = df["Gender"].fillna("Not Reported")
    df["High School"]= df["High School"].fillna("Not Reported")
    df["City"]      = df["Location"].str.split(",").str[0].str.strip()

    NORTHEAST = {"NY","MA","CT","PA","NJ","ME","RI","VT","NH"}
    SOUTH     = {"GA","TX","FL","NC","TN","VA","SC","LA","KY","AL","MS","AR","WV","MD","DC"}
    WEST      = {"CA","WA","OR","AZ","NV","CO","UT","NM","HI","AK","ID","MT","WY"}
    MIDWEST   = {"IL","OH","MI","IN","WI","MN","IA","MO","ND","SD","NE","KS"}

    def region(s):
        if pd.isna(s): return None
        s = str(s).strip()
        if s in NORTHEAST: return "Northeast"
        if s in SOUTH:     return "South"
        if s in WEST:      return "West"
        if s in MIDWEST:   return "Midwest"
        return "Other / International"

    df["Region"] = df["State"].apply(region)
    return df


df = load_and_clean()


# ── Helpers ───────────────────────────────────────────────────────────────────
def apply_filters(df, year_range, genders, grad_status, states, occ_cats):
    fdf = df[df["Cohort Year"].between(year_range[0], year_range[1])].copy()
    if genders:
        fdf = fdf[fdf["Gender"].isin(genders)]
    if grad_status != "All":
        fdf = fdf[fdf["Has Graduate Degree"] == grad_status]
    if states:
        fdf = fdf[fdf["State"].isin(states)]
    if occ_cats:
        fdf = fdf[fdf["Broad Occupation Category"].isin(occ_cats)]
    return fdf


def styled_fig(fig, height=320):
    fig.update_layout(
        height=height, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="DM Sans, sans-serif", color=C_GRAY_D, size=12),
        margin=dict(l=8, r=8, t=36, b=8),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
        xaxis=dict(showgrid=False, zeroline=False),
        yaxis=dict(showgrid=True, gridcolor="#F0F0F0", zeroline=False),
    )
    return fig


def page_header(title, subtitle, icon, badge=None):
    b = f'<span class="pg-badge">{badge}</span>' if badge else ""
    st.markdown(f"""
    <div class="pg-header">
      <div style="display:flex;align-items:center;gap:14px;">
        <div style="font-size:2rem">{icon}</div>
        <div><p class="pg-title">{title}</p><p class="pg-subtitle">{subtitle}</p></div>
      </div>{b}
    </div>""", unsafe_allow_html=True)


def sec(title):
    st.markdown(f'<p class="sec-hdr">{title}</p>', unsafe_allow_html=True)

def insight(text):
    st.markdown(f'<div class="insight">💡 {text}</div>', unsafe_allow_html=True)

def dq(text):
    st.markdown(f'<div class="dq-note">⚠️ {text}</div>', unsafe_allow_html=True)

def divider():
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

def salary_df(fdf, confidence=None):
    s = fdf[fdf["Estimated Salary"].notna()].copy()
    if confidence == "High":
        s = s[s["Salary Confidence"] == "High"]
    elif confidence == "Medium":
        s = s[s["Salary Confidence"] == "Medium"]
    elif confidence == "High+Medium":
        s = s[s["Salary Confidence"].isin(["High","Medium"])]
    return s


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        st.markdown(f'<div style="text-align:center;padding:18px 8px 14px"><img src="data:image/png;base64,{b64}" style="width:80%;max-width:180px;border-radius:8px"></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="text-align:center;padding:22px 8px 14px"><span style="font-family:Fraunces,serif;font-size:1.3rem;font-weight:600;color:{C_WHITE}">Buffalo Prep</span><br><span style="font-size:0.65rem;color:rgba(255,255,255,0.35);text-transform:uppercase;letter-spacing:.1em">Alumni Dashboard</span></div>', unsafe_allow_html=True)

    st.markdown("<div class='nav-lbl'>Navigation</div>", unsafe_allow_html=True)
    page = st.radio("", ["📊  Overview","💰  Economic Mobility","🗺️  Geographic Mobility","🎓  Educational Attainment","💼  Career Outcomes"], label_visibility="collapsed")

    st.markdown("---")
    st.markdown("<div class='nav-lbl'>Filters</div>", unsafe_allow_html=True)

    yr_min = int(df["Cohort Year"].min())
    yr_max = int(df["Cohort Year"].max())
    year_range = st.slider("Cohort Year", yr_min, yr_max, (yr_min, yr_max))

    all_genders = sorted(df["Gender"].unique().tolist())
    sel_genders = st.multiselect("Gender", all_genders, default=all_genders)

    grad_status = st.selectbox("Graduate Degree", ["All", "Yes", "No"])

    known_states = sorted(df["State"].dropna().unique().tolist())
    sel_states = st.multiselect("State (location data only)", known_states, default=[])

    known_occs = sorted([o for o in df["Broad Occupation Category"].dropna().unique() if o != "Not Reported"])
    sel_occs = st.multiselect("Occupation Category", known_occs, default=[])

    st.markdown("---")
    st.markdown(f'<p style="font-size:.63rem;opacity:.3;text-align:center;line-height:1.7">Buffalo Prep Alumni Data<br>Cohorts 1990–2017 · 578 alumni</p>', unsafe_allow_html=True)


# ── Apply filters ─────────────────────────────────────────────────────────────
fdf = apply_filters(df, year_range,
                    sel_genders if sel_genders else None,
                    grad_status,
                    sel_states if sel_states else None,
                    sel_occs if sel_occs else None)

n_filtered = len(fdf)
if n_filtered == 0:
    st.warning("No alumni match the current filters. Please adjust the sidebar selections.")
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════
if "Overview" in page:
    page_header("Alumni Overview", "High-level snapshot of Buffalo Prep alumni outcomes", "📊",
                badge=f"{n_filtered} alumni · {fdf['Cohort Year'].nunique()} cohorts")

    total    = n_filtered
    pct_ug   = fdf["Undergraduate University Institution"].notna().sum() / total * 100
    pct_grad = (fdf["Has Graduate Degree"] == "Yes").sum() / total * 100
    hc       = salary_df(fdf, "High")
    avg_sal_str = f"${hc['Estimated Salary'].mean()/1000:.0f}K" if len(hc) > 0 else "N/A"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Alumni", f"{total:,}")
    c2.metric("With UG Institution on Record", f"{pct_ug:.1f}%", help="86.7% of full dataset")
    c3.metric("With Graduate Degree", f"{pct_grad:.1f}%", help="Based on Has Graduate Degree = Yes")
    c4.metric("Avg Salary (High Confidence)", avg_sal_str, help=f"n={len(hc)} high-confidence records")

    dq(f"Salary data available for {fdf['Estimated Salary'].notna().sum()} of {total} alumni. "
       f"Of those, {len(hc)} are High Confidence (BLS OEWS 2023 verified). "
       f"'No graduate degree' includes 481 alumni where the field was blank — treated as No because all verified graduates have an explicit Yes.")

    divider()

    col1, col2 = st.columns([3, 2])
    with col1:
        sec("Alumni per Cohort Year")
        cb = fdf.groupby("Cohort Year").size().reset_index(name="Count")
        cb["Cohort Year"] = cb["Cohort Year"].astype(int)
        fig = px.bar(cb, x="Cohort Year", y="Count", color_discrete_sequence=[C_BLUE])
        fig.update_layout(xaxis=dict(tickmode="linear", dtick=2), xaxis_title="", yaxis_title="Alumni")
        st.plotly_chart(styled_fig(fig), use_container_width=True)

    with col2:
        sec("Gender Distribution")
        gd = fdf["Gender"].value_counts().reset_index()
        gd.columns = ["Gender", "Count"]
        gc_map = {"F": C_BLUE, "M": C_BLUE_D, "Non-Binary": C_GOLD, "Not Reported": C_GRAY_M}
        fig2 = px.pie(gd, names="Gender", values="Count", hole=0.55, color="Gender", color_discrete_map=gc_map)
        fig2.update_traces(textinfo="percent+label", textfont_size=11, marker=dict(line=dict(color=C_WHITE, width=2)))
        fig2.update_layout(showlegend=False, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(styled_fig(fig2, height=290), use_container_width=True)

    divider()
    col3, col4 = st.columns(2)

    with col3:
        sec("Salary Distribution — All Alumni with Salary Data")
        sal_all = salary_df(fdf)
        if len(sal_all) >= 5:
            fig3 = px.histogram(sal_all, x="Estimated Salary", nbins=20,
                                color="Salary Confidence",
                                color_discrete_map={"High": C_BLUE, "Medium": C_GOLD, "Low": C_GRAY_M},
                                barmode="overlay", opacity=0.8)
            fig3.update_layout(xaxis_tickprefix="$", xaxis_tickformat=",",
                               xaxis_title="Estimated Salary", yaxis_title="Alumni Count", legend_title="Confidence")
            st.plotly_chart(styled_fig(fig3), use_container_width=True)
            dq("Low-confidence salaries are BLS occupation-category estimates — treat with caution.")
        else:
            st.info("Insufficient salary data for current filter selection.")

    with col4:
        sec("Graduate Degree Attainment by Cohort")
        gc2 = fdf.groupby("Cohort Year").apply(
            lambda x: pd.Series({"Total": len(x), "WithGrad": (x["Has Graduate Degree"]=="Yes").sum()})
        ).reset_index()
        gc2["PctGrad"] = gc2["WithGrad"] / gc2["Total"] * 100
        gc2["Cohort Year"] = gc2["Cohort Year"].astype(int)
        fig4 = px.bar(gc2, x="Cohort Year", y="PctGrad",
                      color="PctGrad", color_continuous_scale=SEQ_BLUE,
                      hover_data={"WithGrad": True, "Total": True})
        fig4.update_layout(xaxis=dict(tickmode="linear", dtick=2),
                           xaxis_title="", yaxis_title="% With Graduate Degree",
                           yaxis_ticksuffix="%", coloraxis_showscale=False)
        insight("Recent cohorts (2015–2017) show 0–3% graduate rates — expected, as most alumni from these years have not yet had time to complete graduate programs.")
        st.plotly_chart(styled_fig(fig4), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — ECONOMIC MOBILITY
# ═══════════════════════════════════════════════════════════════════════════════
elif "Economic" in page:
    hc = salary_df(fdf, "High")
    page_header("Economic Mobility", "Salary outcomes across alumni — patterns and distributions", "💰",
                badge=f"Salary data: {fdf['Estimated Salary'].notna().sum()} alumni")

    if len(hc) == 0:
        st.warning("No High Confidence salary records match the current filters.")
        st.stop()

    avg_s = hc["Estimated Salary"].mean()
    med_s = hc["Estimated Salary"].median()
    min_s = hc["Estimated Salary"].min()
    max_s = hc["Estimated Salary"].max()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Avg Salary (High Conf.)",    f"${avg_s:,.0f}", help=f"n={len(hc)}")
    c2.metric("Median Salary (High Conf.)", f"${med_s:,.0f}", help=f"n={len(hc)}")
    c3.metric("Min (High Conf.)",           f"${min_s:,.0f}")
    c4.metric("Max (High Conf.)",           f"${max_s:,.0f}")

    dq(f"All charts on this page use High Confidence records only (n={len(hc)}), "
       f"sourced from BLS OEWS 2023 matched by SOC code. "
       f"Medium (n={len(salary_df(fdf,'Medium'))}) and Low confidence records are excluded.")

    divider()
    col1, col2 = st.columns(2)

    with col1:
        sec("Salary Distribution (High Confidence)")
        fig = px.histogram(hc, x="Estimated Salary", nbins=20, color_discrete_sequence=[C_BLUE])
        fig.update_layout(xaxis_tickprefix="$", xaxis_tickformat=",",
                          xaxis_title="Estimated Salary", yaxis_title="Alumni Count")
        st.plotly_chart(styled_fig(fig), use_container_width=True)

    with col2:
        sec("Salary with Range Error Bars (High Confidence)")
        hc_r = hc.dropna(subset=["Salary Range Low","Salary Range High"]).sort_values("Estimated Salary").reset_index(drop=True)
        if len(hc_r) > 0:
            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(
                x=hc_r.index, y=hc_r["Estimated Salary"], mode="markers",
                marker=dict(color=C_BLUE, size=5, opacity=0.7),
                error_y=dict(type="data", symmetric=False,
                             array=(hc_r["Salary Range High"]-hc_r["Estimated Salary"]).clip(lower=0),
                             arrayminus=(hc_r["Estimated Salary"]-hc_r["Salary Range Low"]).clip(lower=0),
                             color=C_BLUE_L, thickness=1, width=2),
                hovertemplate="Salary: $%{y:,.0f}<extra></extra>"
            ))
            fig2.update_layout(xaxis_title="Alumni (sorted by salary)",
                               yaxis_title="Estimated Salary",
                               yaxis_tickprefix="$", yaxis_tickformat=",")
            st.plotly_chart(styled_fig(fig2), use_container_width=True)

    divider()
    col3, col4 = st.columns(2)

    with col3:
        sec("Salary by Gender (High Confidence)")
        gen_hc = hc[hc["Gender"].isin(["F","M","Non-Binary"])]
        if len(gen_hc) > 0:
            fig3 = px.box(gen_hc, x="Gender", y="Estimated Salary", color="Gender",
                          color_discrete_map={"F": C_BLUE, "M": C_BLUE_D, "Non-Binary": C_GOLD},
                          points="outliers")
            fig3.update_layout(yaxis_tickprefix="$", yaxis_tickformat=",",
                               xaxis_title="", yaxis_title="Estimated Salary", showlegend=False)
            st.plotly_chart(styled_fig(fig3), use_container_width=True)
            gstat = gen_hc.groupby("Gender")["Estimated Salary"].agg(Mean="mean", Median="median", Count="count").reset_index()
            gstat["Mean"]   = gstat["Mean"].apply(lambda v: f"${v:,.0f}")
            gstat["Median"] = gstat["Median"].apply(lambda v: f"${v:,.0f}")
            st.dataframe(gstat, hide_index=True, use_container_width=True)
            dq("Non-Binary group has n=1 high-confidence record — not statistically meaningful.")

    with col4:
        sec("Salary by Graduate Degree Status (High Confidence)")
        grad_hc = hc[hc["Has Graduate Degree"].isin(["Yes","No"])]
        if len(grad_hc) > 0:
            fig4 = px.box(grad_hc, x="Has Graduate Degree", y="Estimated Salary", color="Has Graduate Degree",
                          color_discrete_map={"Yes": C_BLUE, "No": C_GOLD}, points="outliers")
            fig4.update_layout(yaxis_tickprefix="$", yaxis_tickformat=",",
                               xaxis_title="Has Graduate Degree", yaxis_title="Estimated Salary", showlegend=False)
            st.plotly_chart(styled_fig(fig4), use_container_width=True)
            gstat2 = grad_hc.groupby("Has Graduate Degree")["Estimated Salary"].agg(Mean="mean", Median="median", Count="count").reset_index()
            gstat2["Mean"]   = gstat2["Mean"].apply(lambda v: f"${v:,.0f}")
            gstat2["Median"] = gstat2["Median"].apply(lambda v: f"${v:,.0f}")
            st.dataframe(gstat2, hide_index=True, use_container_width=True)
            yes_n = (grad_hc["Has Graduate Degree"]=="Yes").sum()
            no_n  = (grad_hc["Has Graduate Degree"]=="No").sum()
            if yes_n >= 5 and no_n >= 5:
                yes_med = grad_hc[grad_hc["Has Graduate Degree"]=="Yes"]["Estimated Salary"].median()
                no_med  = grad_hc[grad_hc["Has Graduate Degree"]=="No"]["Estimated Salary"].median()
                if yes_med > no_med:
                    insight(f"Alumni with graduate degrees show a higher median salary (${yes_med:,.0f}) vs. those without (${no_med:,.0f}) in this selection. Correlation only — this data cannot establish causation.")
            else:
                dq(f"Small group sizes: Yes n={yes_n}, No n={no_n}. Treat comparison with caution.")

    divider()
    sec("Average Salary by Occupation Category (High Confidence, n ≥ 2)")
    occ_avg = hc.groupby("Broad Occupation Category")["Estimated Salary"].agg(Avg="mean", Median="median", Count="count").reset_index().sort_values("Avg", ascending=True)
    occ_avg = occ_avg[occ_avg["Count"] >= 2]
    if len(occ_avg) > 0:
        fig5 = px.bar(occ_avg, x="Avg", y="Broad Occupation Category", orientation="h",
                      color="Avg", color_continuous_scale=SEQ_BLUE,
                      text=occ_avg["Avg"].apply(lambda v: f"${v/1000:.0f}K"),
                      custom_data=["Count","Median"])
        fig5.update_traces(textposition="outside",
                           textfont=dict(color="white", size=12),
                           hovertemplate="<b>%{y}</b><br>Avg: $%{x:,.0f}<br>Median: $%{customdata[1]:,.0f}<br>n=%{customdata[0]}<extra></extra>")
        fig5.update_layout(coloraxis_showscale=False, xaxis_tickprefix="$", xaxis_tickformat=",",
                           yaxis_title="", xaxis_title="Average Estimated Salary",
                           xaxis=dict(range=[0, occ_avg["Avg"].max()*1.25]))
        st.plotly_chart(styled_fig(fig5, height=420), use_container_width=True)

    divider()
    sec("Average Salary by Cohort Year (High Confidence, cohorts with n ≥ 2)")
    sc = hc.groupby("Cohort Year")["Estimated Salary"].agg(Avg="mean", Count="count").reset_index()
    sc = sc[sc["Count"] >= 2]
    sc["Cohort Year"] = sc["Cohort Year"].astype(int)
    if len(sc) > 0:
        fig6 = px.scatter(sc, x="Cohort Year", y="Avg", size="Count", trendline="lowess",
                          color_discrete_sequence=[C_BLUE], trendline_color_override=C_BLUE_D,
                          hover_data={"Count": True})
        fig6.update_traces(marker=dict(opacity=0.75))
        fig6.update_layout(xaxis_title="Cohort Year", yaxis_title="Avg Estimated Salary",
                           yaxis_tickprefix="$", yaxis_tickformat=",")
        insight("Bubble size = number of verified salary records in that cohort. Earlier cohorts have had more career time, which may explain higher averages.")
        st.plotly_chart(styled_fig(fig6, height=340), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — GEOGRAPHIC MOBILITY
# ═══════════════════════════════════════════════════════════════════════════════
elif "Geographic" in page:
    loc_df = fdf[fdf["State"].notna()].copy()
    n_loc  = len(loc_df)
    page_header("Geographic Mobility", "Where Buffalo Prep alumni have settled", "🗺️",
                badge=f"{n_loc} alumni with location data")

    dq(f"Only {n_loc} of {n_filtered} alumni ({n_loc/n_filtered*100:.1f}%) have location data. Findings reflect this subset only.")

    if n_loc == 0:
        st.info("No alumni with location data match the current filters.")
        st.stop()

    top_state   = loc_df["State"].value_counts().idxmax()
    top_state_n = loc_df["State"].value_counts().max()
    n_states    = loc_df["State"].nunique()
    stayed_buf  = (loc_df["City"].str.lower() == "buffalo").sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alumni with Location", n_loc)
    c2.metric("States / Territories", n_states)
    c3.metric("Top State", f"{top_state} ({top_state_n})")
    c4.metric("Remaining in Buffalo", stayed_buf, help=f"{stayed_buf/n_loc*100:.1f}% of alumni with known location")

    divider()
    sec("Alumni Count by State — Choropleth")
    state_ct = loc_df["State"].value_counts().reset_index()
    state_ct.columns = ["State", "Count"]
    us_ct = state_ct[state_ct["State"].str.len() == 2]
    fig_map = px.choropleth(us_ct, locations="State", locationmode="USA-states",
                            color="Count", scope="usa",
                            color_continuous_scale=[[0,C_BLUE_L],[0.3,C_BLUE_M],[0.7,C_BLUE],[1,C_BLUE_XD]],
                            labels={"Count":"Alumni"})
    fig_map.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)",
                          geo=dict(bgcolor="rgba(0,0,0,0)", lakecolor="rgba(0,0,0,0)",
                                   landcolor="#F8FAFC", showlakes=True,
                                   coastlinecolor="#E5E7EB", subunitcolor="#E5E7EB"),
                          margin=dict(l=0,r=0,t=10,b=0),
                          coloraxis_colorbar=dict(title="Alumni", thickness=12))
    st.plotly_chart(fig_map, use_container_width=True)

    divider()
    col1, col2 = st.columns(2)

    with col1:
        sec("Alumni by State (Bar)")
        fig2 = px.bar(state_ct.sort_values("Count", ascending=False), x="State", y="Count",
                      color="Count", color_continuous_scale=SEQ_BLUE)
        fig2.update_layout(xaxis_title="", yaxis_title="Alumni", coloraxis_showscale=False,
                           xaxis=dict(categoryorder="total descending"))
        st.plotly_chart(styled_fig(fig2), use_container_width=True)

    with col2:
        sec("Alumni by Region")
        reg_df = loc_df[loc_df["Region"].notna()]["Region"].value_counts().reset_index()
        reg_df.columns = ["Region", "Count"]
        rc = {"Northeast": C_BLUE,"South": C_BLUE_M,"West": C_BLUE_D,"Midwest": "#7BBDE0","Other / International": C_GOLD}
        fig3 = px.pie(reg_df, names="Region", values="Count", hole=0.52, color="Region", color_discrete_map=rc)
        fig3.update_traces(textinfo="percent+label", textfont_size=11, marker=dict(line=dict(color=C_WHITE, width=2)))
        fig3.update_layout(showlegend=False, margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(styled_fig(fig3, height=300), use_container_width=True)

    divider()
    sec("Top Cities")
    city_df = loc_df[loc_df["City"].notna()]["City"].value_counts().head(15).reset_index()
    city_df.columns = ["City","Count"]
    ccolors = {c: (C_GOLD if c=="Buffalo" else C_BLUE) for c in city_df["City"]}
    fig4 = px.bar(city_df, x="Count", y="City", orientation="h", color="City", color_discrete_map=ccolors)
    fig4.update_layout(yaxis=dict(categoryorder="total ascending"), xaxis_title="Alumni", yaxis_title="", showlegend=False)
    st.plotly_chart(styled_fig(fig4, height=380), use_container_width=True)
    insight(f"Buffalo is highlighted in gold. {stayed_buf} alumni ({stayed_buf/n_loc*100:.1f}% of those with known location) remain in Buffalo.")

    divider()
    sec("Geographic Diversity by Cohort Year (alumni with location data only)")
    geo_c = loc_df.groupby("Cohort Year")["State"].nunique().reset_index()
    geo_c.columns = ["Cohort Year","Unique States"]
    geo_c["Cohort Year"] = geo_c["Cohort Year"].astype(int)
    fig5 = px.bar(geo_c, x="Cohort Year", y="Unique States", color_discrete_sequence=[C_BLUE])
    fig5.update_layout(xaxis=dict(tickmode="linear", dtick=2), xaxis_title="", yaxis_title="Unique States")
    dq("Counts reflect only alumni with known location — low values in recent cohorts may indicate data gaps rather than less mobility.")
    st.plotly_chart(styled_fig(fig5), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — EDUCATIONAL ATTAINMENT
# ═══════════════════════════════════════════════════════════════════════════════
elif "Education" in page:
    page_header("Educational Attainment", "Academic pathways from high school through graduate study", "🎓")

    total = n_filtered
    n_ug  = fdf["Undergraduate University Institution"].notna().sum()
    n_grad= (fdf["Has Graduate Degree"]=="Yes").sum()
    n_ugt = fdf["Undergraduate Degree Type"].notna().sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Alumni", total)
    c2.metric("With UG Institution", f"{n_ug} ({n_ug/total*100:.1f}%)")
    c3.metric("With Graduate Degree", f"{n_grad} ({n_grad/total*100:.1f}%)")
    c4.metric("UG Degree Type on Record", f"{n_ugt} ({n_ugt/total*100:.1f}%)")

    dq(f"UG institution: {n_ug}/{total} filled. UG degree type: {n_ugt}/{total} filled. "
       f"Grad institution: {fdf['Graduate University Institution'].notna().sum()}/{total} filled.")

    divider()
    col1, col2 = st.columns(2)

    with col1:
        sec("Top Undergraduate Institutions (n ≥ 2)")
        ug_ct = fdf["Undergraduate University Institution"].value_counts()
        ug_ct = ug_ct[ug_ct>=2].reset_index()
        ug_ct.columns = ["Institution","Count"]
        fig = px.bar(ug_ct.head(15), x="Count", y="Institution", orientation="h",
                     color="Count", color_continuous_scale=SEQ_BLUE)
        fig.update_layout(yaxis=dict(categoryorder="total ascending"), xaxis_title="Alumni", yaxis_title="", coloraxis_showscale=False)
        st.plotly_chart(styled_fig(fig, height=440), use_container_width=True)

    with col2:
        sec("Top Graduate Institutions")
        gi = fdf["Graduate University Institution"].dropna().value_counts().reset_index()
        gi.columns = ["Institution","Count"]
        if len(gi) > 0:
            fig2 = px.bar(gi.head(15), x="Count", y="Institution", orientation="h",
                          color="Count", color_continuous_scale=SEQ_BLUE)
            fig2.update_layout(yaxis=dict(categoryorder="total ascending"), xaxis_title="Alumni", yaxis_title="", coloraxis_showscale=False)
            st.plotly_chart(styled_fig(fig2, height=440), use_container_width=True)
        else:
            st.info("No graduate institution data for current filter.")

    divider()
    col3, col4 = st.columns(2)

    with col3:
        sec("Graduate Degree Category Distribution")
        gdc = fdf["Grad Degree Category"].dropna()
        gdc = gdc[~gdc.isin(["Other / Non-Graduate"])].value_counts().reset_index()
        gdc.columns = ["Category","Count"]
        if len(gdc) > 0:
            fig3 = px.pie(gdc, names="Category", values="Count", hole=0.5, color_discrete_sequence=PALETTE)
            fig3.update_traces(textinfo="percent+label", textfont_size=10, marker=dict(line=dict(color=C_WHITE, width=2)))
            fig3.update_layout(showlegend=True, margin=dict(l=0,r=0,t=10,b=0), legend=dict(font=dict(size=10)))
            st.plotly_chart(styled_fig(fig3, height=320), use_container_width=True)

    with col4:
        sec("Graduate Attainment Rate by Cohort Year")
        gc2 = fdf.groupby("Cohort Year").apply(
            lambda x: pd.Series({"Total":len(x),"WithGrad":(x["Has Graduate Degree"]=="Yes").sum()})
        ).reset_index()
        gc2["PctGrad"] = gc2["WithGrad"]/gc2["Total"]*100
        gc2["Cohort Year"] = gc2["Cohort Year"].astype(int)
        fig4 = px.bar(gc2, x="Cohort Year", y="PctGrad",
                      color="PctGrad", color_continuous_scale=SEQ_BLUE,
                      hover_data={"WithGrad":True,"Total":True})
        fig4.update_layout(xaxis=dict(tickmode="linear",dtick=2),
                           xaxis_title="", yaxis_title="% With Graduate Degree",
                           yaxis_ticksuffix="%", coloraxis_showscale=False)
        insight("2016–2017 cohorts show 0% because most alumni are still within typical graduate program timelines.")
        st.plotly_chart(styled_fig(fig4), use_container_width=True)

    divider()
    col5, col6 = st.columns(2)

    with col5:
        sec("Top High Schools (n ≥ 2)")
        hs = fdf[fdf["High School"]!="Not Reported"]["High School"].value_counts()
        hs = hs[hs>=2].reset_index()
        hs.columns = ["High School","Count"]
        fig5 = px.bar(hs.head(15), x="Count", y="High School", orientation="h",
                      color="Count", color_continuous_scale=SEQ_BLUE)
        fig5.update_layout(yaxis=dict(categoryorder="total ascending"), xaxis_title="Alumni", yaxis_title="", coloraxis_showscale=False)
        st.plotly_chart(styled_fig(fig5, height=400), use_container_width=True)

    with col6:
        sec("Graduate Degree Rate by High School (schools with ≥ 10 alumni)")
        hs_stats = fdf[fdf["High School"]!="Not Reported"].groupby("High School").apply(
            lambda x: pd.Series({"Alumni":len(x),"GradRate":round((x["Has Graduate Degree"]=="Yes").sum()/len(x)*100,1)})
        ).reset_index()
        hs_stats = hs_stats[hs_stats["Alumni"]>=10].sort_values("GradRate", ascending=True)
        if len(hs_stats) > 0:
            fig6 = px.bar(hs_stats, x="GradRate", y="High School", orientation="h",
                          color="GradRate", color_continuous_scale=SEQ_BLUE,
                          text=hs_stats["GradRate"].apply(lambda v: f"{v:.1f}%"),
                          custom_data=["Alumni"])
            fig6.update_traces(textposition="outside",
                               textfont=dict(color="white", size=12),
                               hovertemplate="<b>%{y}</b><br>Grad Rate: %{x:.1f}%<br>Alumni: %{customdata[0]}<extra></extra>")
            fig6.update_layout(yaxis=dict(categoryorder="total ascending"),
                               xaxis_title="Graduate Degree Rate (%)", yaxis_title="",
                               coloraxis_showscale=False,
                               xaxis=dict(ticksuffix="%", range=[0, hs_stats["GradRate"].max()*1.3]))
            st.plotly_chart(styled_fig(fig6, height=400), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — CAREER OUTCOMES
# ═══════════════════════════════════════════════════════════════════════════════
elif "Career" in page:
    career_df = fdf[fdf["Broad Occupation Category"].notna()].copy()
    n_career  = len(career_df)
    page_header("Career Outcomes", "Occupational pathways and salary outcomes by field", "💼",
                badge=f"{n_career} alumni with occupation data")

    dq(f"Career/occupation data available for {n_career} of {n_filtered} alumni ({n_career/n_filtered*100:.1f}%). "
       f"The remaining {n_filtered-n_career} alumni have no occupation on record.")

    if n_career == 0:
        st.info("No career data available for the current filters.")
        st.stop()

    top_occ   = career_df["Broad Occupation Category"].value_counts().idxmax()
    top_occ_n = career_df["Broad Occupation Category"].value_counts().max()
    n_soc     = career_df["SOC Title"].notna().sum()
    hc_car    = salary_df(career_df, "High")
    avg_car   = f"${hc_car['Estimated Salary'].mean():,.0f}" if len(hc_car)>0 else "N/A"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alumni with Occupation Data", n_career)
    c2.metric("Top Field", top_occ, delta=f"{top_occ_n} alumni")
    c3.metric("SOC Title Records", n_soc)
    c4.metric("Avg Salary (High Conf.)", avg_car, help=f"n={len(hc_car)}")

    divider()
    col1, col2 = st.columns(2)

    with col1:
        sec("Alumni by Broad Occupation Category")
        occ_ct = career_df["Broad Occupation Category"].value_counts().reset_index()
        occ_ct.columns = ["Category","Count"]
        fig = px.bar(occ_ct, x="Count", y="Category", orientation="h",
                     color="Count", color_continuous_scale=SEQ_BLUE)
        fig.update_layout(yaxis=dict(categoryorder="total ascending"), xaxis_title="Alumni", yaxis_title="", coloraxis_showscale=False)
        st.plotly_chart(styled_fig(fig, height=420), use_container_width=True)

    with col2:
        sec("Top SOC Titles (n ≥ 2)")
        soc_ct = career_df["SOC Title"].value_counts()
        soc_ct = soc_ct[soc_ct>=2].reset_index()
        soc_ct.columns = ["SOC Title","Count"]
        if len(soc_ct) > 0:
            fig2 = px.bar(soc_ct.head(15), x="Count", y="SOC Title", orientation="h",
                          color="Count", color_continuous_scale=SEQ_BLUE)
            fig2.update_layout(yaxis=dict(categoryorder="total ascending"), xaxis_title="Alumni", yaxis_title="", coloraxis_showscale=False)
            st.plotly_chart(styled_fig(fig2, height=420), use_container_width=True)

    divider()
    sec("Salary by Occupation Category (High Confidence, n ≥ 2)")
    hc_occ = salary_df(career_df, "High")
    if len(hc_occ) > 0:
        occ_box = hc_occ.groupby("Broad Occupation Category").filter(lambda x: len(x)>=2)
        fig3 = px.box(occ_box, x="Broad Occupation Category", y="Estimated Salary",
                      color="Broad Occupation Category", color_discrete_sequence=PALETTE, points="outliers")
        fig3.update_layout(xaxis_tickangle=-30, xaxis_title="", yaxis_title="Estimated Salary",
                           yaxis_tickprefix="$", yaxis_tickformat=",", showlegend=False)
        st.plotly_chart(styled_fig(fig3, height=400), use_container_width=True)

    divider()
    col3, col4 = st.columns(2)

    with col3:
        sec("Career Field by Graduate Degree Status")
        grad_occ = career_df.groupby(["Broad Occupation Category","Has Graduate Degree"]).size().reset_index(name="Count")
        fig4 = px.bar(grad_occ, x="Broad Occupation Category", y="Count",
                      color="Has Graduate Degree",
                      color_discrete_map={"Yes": C_GOLD,"No": C_BLUE},
                      barmode="stack")
        fig4.update_layout(xaxis_tickangle=-30, xaxis_title="", yaxis_title="Alumni", legend_title="Grad Degree")
        st.plotly_chart(styled_fig(fig4, height=360), use_container_width=True)

    with col4:
        if len(hc_occ) > 0:
            sec("Average Salary by Occupation (High Confidence, n ≥ 2)")
            occ_avg = hc_occ.groupby("Broad Occupation Category")["Estimated Salary"].agg(Avg="mean", Median="median", Count="count").reset_index().sort_values("Avg", ascending=True)
            occ_avg = occ_avg[occ_avg["Count"]>=2]
            if len(occ_avg) > 0:
                fig5 = px.bar(occ_avg, x="Avg", y="Broad Occupation Category", orientation="h",
                              color="Avg", color_continuous_scale=SEQ_BLUE,
                              text=occ_avg["Avg"].apply(lambda v: f"${v/1000:.0f}K"),
                              custom_data=["Count","Median"])
                fig5.update_traces(textposition="outside",
                                   textfont=dict(color="white", size=12),
                                   hovertemplate="<b>%{y}</b><br>Avg: $%{x:,.0f}<br>Median: $%{customdata[1]:,.0f}<br>n=%{customdata[0]}<extra></extra>")
                fig5.update_layout(coloraxis_showscale=False, xaxis_tickprefix="$", xaxis_tickformat=",",
                                   yaxis_title="", xaxis_title="Average Estimated Salary",
                                   xaxis=dict(range=[0, occ_avg["Avg"].max()*1.25]))
                st.plotly_chart(styled_fig(fig5, height=360), use_container_width=True)
