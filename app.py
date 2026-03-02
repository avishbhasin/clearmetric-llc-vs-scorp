"""
LLC vs S-Corp Tax Calculator — ClearMetric
https://clearmetric.gumroad.com

Helps business owners compare LLC (sole prop/partnership) vs S-Corp taxation
to see which structure saves more.
"""

import streamlit as st
import plotly.graph_objects as go
import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="LLC vs S-Corp Tax Calculator — ClearMetric",
    page_icon="🏢",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Custom CSS (navy/indigo theme)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; max-width: 1200px; }
    .stMetric { background: #f8f9fa; border-radius: 8px; padding: 12px; border-left: 4px solid #2C3E8F; }
    h1 { color: #2C3E8F; }
    h2, h3 { color: #1A2766; }
    .cta-box {
        background: linear-gradient(135deg, #1A2766 0%, #2C3E8F 100%);
        color: white; padding: 24px; border-radius: 12px; text-align: center;
        margin: 20px 0;
    }
    .cta-box a { color: #D6DBEF; text-decoration: none; font-weight: bold; font-size: 1.1rem; }
    div[data-testid="stSidebar"] { background: #f8f9fa; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Tax Constants (2026)
# ---------------------------------------------------------------------------
SS_WAGE_BASE_2026 = 184_500
STANDARD_DEDUCTION = {
    "Single": 16_100,
    "Married Filing Jointly": 32_200,
    "Head of Household": 24_150,
}

BRACKETS = {
    "Single": [
        (0.10, 0), (0.12, 12_400), (0.22, 50_400), (0.24, 105_700),
        (0.32, 201_775), (0.35, 256_225), (0.37, 640_600),
    ],
    "Married Filing Jointly": [
        (0.10, 0), (0.12, 24_800), (0.22, 100_800), (0.24, 211_400),
        (0.32, 403_550), (0.35, 512_450), (0.37, 768_700),
    ],
    "Head of Household": [
        (0.10, 0), (0.12, 17_700), (0.22, 67_450), (0.24, 105_700),
        (0.32, 201_775), (0.35, 256_200), (0.37, 640_600),
    ],
}

STATE_TAX_RATES = {
    "Alabama": 0.05, "Alaska": 0, "Arizona": 0.025, "Arkansas": 0.045,
    "California": 0.093, "Colorado": 0.0455, "Connecticut": 0.065,
    "Delaware": 0.066, "District of Columbia": 0.0975, "Florida": 0,
    "Georgia": 0.055, "Hawaii": 0.09, "Idaho": 0.058, "Illinois": 0.0495,
    "Indiana": 0.0315, "Iowa": 0.044, "Kansas": 0.057, "Kentucky": 0.045,
    "Louisiana": 0.0425, "Maine": 0.075, "Maryland": 0.0575,
    "Massachusetts": 0.05, "Michigan": 0.0425, "Minnesota": 0.0985,
    "Mississippi": 0.05, "Missouri": 0.045, "Montana": 0.069,
    "Nebraska": 0.0684, "Nevada": 0, "New Hampshire": 0,
    "New Jersey": 0.1075, "New Mexico": 0.059, "New York": 0.109,
    "North Carolina": 0.0475, "North Dakota": 0.029, "Ohio": 0.0395,
    "Oklahoma": 0.0475, "Oregon": 0.099, "Pennsylvania": 0.0307,
    "Rhode Island": 0.0599, "South Carolina": 0.065, "South Dakota": 0,
    "Tennessee": 0, "Texas": 0, "Utah": 0.0485, "Vermont": 0.0875,
    "Virginia": 0.0575, "Washington": 0, "West Virginia": 0.065,
    "Wisconsin": 0.0765, "Wyoming": 0,
}

S_CORP_COSTS = {
    "Payroll service": 1_200,
    "Tax preparation": 1_500,
    "State filing": 800,
}
S_CORP_TOTAL_COST = sum(S_CORP_COSTS.values())


def federal_income_tax(taxable_income: float, filing_status: str) -> float:
    """Compute federal income tax using 2026 brackets."""
    if taxable_income <= 0:
        return 0.0
    brackets = BRACKETS[filing_status]
    tax = 0.0
    prev = 0
    for rate, thresh in brackets:
        if taxable_income <= thresh:
            tax += (taxable_income - prev) * rate
            break
        tax += (thresh - prev) * rate
        prev = thresh
    else:
        tax += (taxable_income - prev) * brackets[-1][0]
    return max(0, tax)


def se_tax_llc(net_income: float, w2_income: float) -> float:
    """SE tax for LLC: 15.3% on 92.35% of net income, SS capped at wage base."""
    se_taxable = net_income * 0.9235
    remaining_ss_cap = max(0, SS_WAGE_BASE_2026 - w2_income)
    ss_taxable = min(se_taxable, remaining_ss_cap)
    ss_tax = ss_taxable * 0.124
    medicare_tax = se_taxable * 0.029
    return ss_tax + medicare_tax


def fica_scorp(salary: float, w2_other: float) -> tuple[float, float]:
    """FICA for S-Corp salary: returns (employee+employer total, employer_portion)."""
    remaining_ss_cap = max(0, SS_WAGE_BASE_2026 - w2_other)
    ss_taxable = min(salary, remaining_ss_cap)
    ss_tax = ss_taxable * 0.124  # 6.2% employee + 6.2% employer
    medicare_tax = salary * 0.029  # 1.45% + 1.45%
    total_fica = ss_tax + medicare_tax
    employer_portion = ss_taxable * 0.062 + salary * 0.0145
    return total_fica, employer_portion


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown("# 🏢 LLC vs S-Corp Tax Calculator")
st.markdown("**Compare taxation** — see which structure saves you more.")
st.markdown("---")

# ---------------------------------------------------------------------------
# Sidebar — User inputs
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## Your Numbers")
    st.button("🔄 Update Results", use_container_width=True)

    st.markdown("### Business Income")
    business_net = st.number_input(
        "Business net income ($)",
        value=120_000,
        min_value=0,
        step=5_000,
        help="Profit after expenses",
    )
    biz_expenses = st.number_input(
        "Business expenses already deducted ($)",
        value=20_000,
        min_value=0,
        step=1_000,
        help="Expenses used to arrive at net income",
    )

    st.markdown("### Filing & Other Income")
    filing_status = st.selectbox(
        "Filing status",
        ["Single", "Married Filing Jointly", "Head of Household"],
    )
    w2_other = st.number_input(
        "Other W-2 income ($)",
        value=0,
        min_value=0,
        step=5_000,
    )

    st.markdown("### S-Corp Settings")
    scorp_salary = st.number_input(
        "S-Corp reasonable salary ($)",
        value=60_000,
        min_value=0,
        step=5_000,
        help="Must be 'reasonable' per IRS — typically 30–50% of profit",
    )

    st.markdown("### Deductions & State")
    state = st.selectbox(
        "State",
        list(STATE_TAX_RATES.keys()),
        index=list(STATE_TAX_RATES.keys()).index("California"),
    )
    health_insurance = st.number_input(
        "Health insurance premiums ($/year)",
        value=6_000,
        min_value=0,
        step=500,
    )
    retirement = st.number_input(
        "Retirement contribution ($)",
        value=20_000,
        min_value=0,
        step=1_000,
        help="SEP-IRA (max 25%) or Solo 401k",
    )
    qbi_eligible = st.checkbox(
        "QBI deduction eligible?",
        value=True,
        help="20% deduction on qualified business income",
    )

# ---------------------------------------------------------------------------
# Calculations — LLC path
# ---------------------------------------------------------------------------
deduction = STANDARD_DEDUCTION[filing_status]
state_rate = STATE_TAX_RATES[state]

# LLC: all net income subject to SE tax
se_tax_llc_val = se_tax_llc(business_net, w2_other)
se_deduction_llc = se_tax_llc_val * 0.5

qbi_llc = 0.20 * business_net if qbi_eligible else 0
agi_llc = w2_other + business_net - health_insurance - retirement - se_deduction_llc
taxable_llc = max(0, agi_llc - deduction)
qbi_lim = min(qbi_llc, 0.20 * taxable_llc) if qbi_eligible else 0
taxable_llc = max(0, taxable_llc - qbi_lim)

fed_llc = federal_income_tax(taxable_llc, filing_status)
state_llc = taxable_llc * state_rate if state_rate > 0 else 0
total_llc = fed_llc + state_llc + se_tax_llc_val

# ---------------------------------------------------------------------------
# Calculations — S-Corp path
# ---------------------------------------------------------------------------
# Cap salary so distributions stay non-negative (employer FICA is business expense)
# Solve: salary + employer_fica <= business_net. Approx: salary <= net / 1.0765
max_salary = business_net / 1.0765 if business_net > 0 else 0
salary_actual = min(scorp_salary, max_salary, business_net)
fica_total, fica_employer = fica_scorp(salary_actual, w2_other)

# Distributions = what's left after salary and employer FICA (business expense)
distributions = max(0, business_net - salary_actual - fica_employer)

qbi_scorp = 0.20 * distributions if qbi_eligible else 0
agi_scorp = w2_other + salary_actual + distributions - health_insurance - retirement
taxable_scorp = max(0, agi_scorp - deduction)
qbi_lim_scorp = min(qbi_scorp, 0.20 * taxable_scorp) if qbi_eligible else 0
taxable_scorp = max(0, taxable_scorp - qbi_lim_scorp)

fed_scorp = federal_income_tax(taxable_scorp, filing_status)
state_scorp = taxable_scorp * state_rate if state_rate > 0 else 0
total_scorp_tax = fed_scorp + state_scorp + fica_total
total_scorp = total_scorp_tax + S_CORP_TOTAL_COST

# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------
savings = total_llc - total_scorp
verdict = "S-Corp saves" if savings > 0 else ("LLC saves" if savings < 0 else "Roughly equal")

# ---------------------------------------------------------------------------
# Break-even analysis
# ---------------------------------------------------------------------------
incomes = np.arange(50_000, 310_000, 10_000)
llc_taxes = []
scorp_taxes = []
for inc in incomes:
    se = se_tax_llc(inc, w2_other)
    se_ded = se * 0.5
    qbi = 0.20 * inc if qbi_eligible else 0
    agi = w2_other + inc - health_insurance - retirement - se_ded
    tax_inc = max(0, agi - deduction)
    qbi_lim = min(qbi, 0.20 * tax_inc) if qbi_eligible else 0
    tax_inc = max(0, tax_inc - qbi_lim)
    fed = federal_income_tax(tax_inc, filing_status)
    st_tax = tax_inc * state_rate if state_rate > 0 else 0
    llc_taxes.append(fed + st_tax + se)

    sal_ratio = scorp_salary / business_net if business_net > 0 else 0.5
    sal = min(inc * sal_ratio, inc / 1.0765)  # reasonable salary scales with income
    fica_t, fica_emp = fica_scorp(sal, w2_other)
    dist = max(0, inc - sal - fica_emp)
    qbi_s = 0.20 * dist if qbi_eligible else 0
    agi_s = w2_other + sal + dist - health_insurance - retirement
    tax_inc_s = max(0, agi_s - deduction)
    qbi_lim_s = min(qbi_s, 0.20 * tax_inc_s) if qbi_eligible else 0
    tax_inc_s = max(0, tax_inc_s - qbi_lim_s)
    fed_s = federal_income_tax(tax_inc_s, filing_status)
    st_s = tax_inc_s * state_rate if state_rate > 0 else 0
    scorp_taxes.append(fed_s + st_s + fica_t + S_CORP_TOTAL_COST)

break_even_df = pd.DataFrame({
    "Business Income": incomes,
    "LLC Total Tax": llc_taxes,
    "S-Corp Total Cost": scorp_taxes,
    "Savings (S-Corp)": np.array(llc_taxes) - np.array(scorp_taxes),
})
break_even_income = break_even_df[break_even_df["Savings (S-Corp)"] > 0]["Business Income"].min()
break_even_point = int(break_even_income) if not pd.isna(break_even_income) else None

# ---------------------------------------------------------------------------
# Display — Key metrics
# ---------------------------------------------------------------------------
st.markdown("## Key Results")

m1, m2, m3, m4 = st.columns(4)
m1.metric("LLC Total Tax", f"${total_llc:,.0f}", None)
m2.metric("S-Corp Total Cost", f"${total_scorp:,.0f}", f"Incl. ${S_CORP_TOTAL_COST:,.0f} S-Corp fees")
m3.metric("Annual Savings", f"${abs(savings):,.0f}", verdict)
m4.metric("Verdict", verdict, f"Break-even ~${break_even_point:,.0f}" if break_even_point else "See chart")

st.markdown("---")

# ---------------------------------------------------------------------------
# Side-by-side bar chart
# ---------------------------------------------------------------------------
st.markdown("## LLC vs S-Corp Comparison")

fig = go.Figure(data=[
    go.Bar(name="LLC", x=["Total Tax"], y=[total_llc], marker_color="#2C3E8F"),
    go.Bar(name="S-Corp", x=["Total Cost"], y=[total_scorp], marker_color="#1A2766"),
])
fig.update_layout(
    barmode="group",
    height=350,
    showlegend=True,
    legend=dict(orientation="h", y=1.02),
    margin=dict(t=40, b=40),
    template="plotly_white",
    yaxis_title="Amount ($)",
)
st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ---------------------------------------------------------------------------
# Tax breakdown table
# ---------------------------------------------------------------------------
st.markdown("## Tax Breakdown")

breakdown_data = {
    "Component": [
        "Federal Income Tax",
        "State Income Tax",
        "SE Tax / FICA",
        "S-Corp Costs (payroll, tax prep, filing)",
        "**Total**",
    ],
    "LLC": [
        f"${fed_llc:,.0f}",
        f"${state_llc:,.0f}",
        f"${se_tax_llc_val:,.0f}",
        "$0",
        f"**${total_llc:,.0f}**",
    ],
    "S-Corp": [
        f"${fed_scorp:,.0f}",
        f"${state_scorp:,.0f}",
        f"${fica_total:,.0f}",
        f"${S_CORP_TOTAL_COST:,.0f}",
        f"**${total_scorp:,.0f}**",
    ],
}
st.dataframe(pd.DataFrame(breakdown_data), use_container_width=True, hide_index=True)

st.markdown("---")

# ---------------------------------------------------------------------------
# When does S-Corp make sense?
# ---------------------------------------------------------------------------
st.markdown("## When Does S-Corp Make Sense?")

if break_even_point:
    st.info(
        f"**Break-even point:** S-Corp typically becomes worthwhile when business net income "
        f"exceeds **~${break_even_point:,.0f}** (with your inputs). Below that, LLC costs less "
        "because S-Corp administrative fees outweigh the SE tax savings."
    )
else:
    st.info(
        "With your inputs, S-Corp saves at all income levels shown. "
        "The break-even is below $50K business income."
    )

fig_breakeven = go.Figure()
fig_breakeven.add_trace(go.Scatter(
    x=break_even_df["Business Income"],
    y=break_even_df["LLC Total Tax"],
    name="LLC Total Tax",
    line=dict(color="#2C3E8F", width=2),
))
fig_breakeven.add_trace(go.Scatter(
    x=break_even_df["Business Income"],
    y=break_even_df["S-Corp Total Cost"],
    name="S-Corp Total Cost",
    line=dict(color="#1A2766", width=2),
))
fig_breakeven.add_trace(go.Scatter(
    x=break_even_df["Business Income"],
    y=break_even_df["Savings (S-Corp)"],
    name="Savings (S-Corp)",
    line=dict(color="#27ae60", width=2, dash="dash"),
    yaxis="y2",
))
fig_breakeven.update_layout(
    height=400,
    xaxis_title="Business Net Income ($)",
    yaxis_title="Tax / Cost ($)",
    yaxis2=dict(title="Savings ($)", overlaying="y", side="right", showgrid=False),
    legend=dict(orientation="h", y=1.08),
    template="plotly_white",
    margin=dict(t=60),
)
st.plotly_chart(fig_breakeven, use_container_width=True)

st.markdown("---")

# ---------------------------------------------------------------------------
# CTA — Excel
# ---------------------------------------------------------------------------
st.markdown("""
<div class="cta-box">
    <h3 style="color: white; margin: 0 0 8px 0;">Get the Full Excel Calculator</h3>
    <p style="margin: 0 0 16px 0;">
        <strong>ClearMetric LLC vs S-Corp Calculator</strong> — $12.99<br>
        ✓ Side-by-side LLC vs S-Corp comparison<br>
        ✓ Break-even analysis ($50K–$300K income)<br>
        ✓ All inputs editable, formulas included<br>
        ✓ How To Use guide
    </p>
    <a href="https://clearmetric.gumroad.com/l/llc-vs-scorp" target="_blank">
        Get It on Gumroad — $12.99 →
    </a>
</div>
""", unsafe_allow_html=True)

# Cross-sell
st.markdown("### More from ClearMetric")
cx1, cx2, cx3 = st.columns(3)
with cx1:
    st.markdown("""
    **📋 Side Hustle Tax Estimator** — $12.99
    Estimate tax liability for freelance, Etsy, Uber income.
    [Get it →](https://clearmetric.gumroad.com/l/side-hustle-tax)
    """)
with cx2:
    st.markdown("""
    **📊 Freelancer Tax Planner** — $14.99
    Quarterly estimates, deductions, SE tax, full-year projection.
    [Get it →](https://clearmetric.gumroad.com/l/freelancer-tax-planner)
    """)
with cx3:
    st.markdown("""
    **💰 Budget Planner** — $13.99
    Track income, expenses, savings with 50/30/20 framework.
    [Get it →](https://clearmetric.gumroad.com/l/budget-planner)
    """)

# Footer
st.markdown("---")
st.caption(
    "© 2026 ClearMetric | [clearmetric.gumroad.com](https://clearmetric.gumroad.com) | "
    "This tool is for educational purposes only. Not financial or tax advice. Consult a CPA."
)
