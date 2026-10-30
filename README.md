# The OI Ratio Edge

### Expiry Day Index Options Intelligence & Performance Analytics

> **Nifty, BankNifty & Sensex Derivative Analytics using Open Interest Ratio Imbalance, 1-Hour Price Action Retest & Advanced MS Excel Modeling**

[![Excel](https://img.shields.io/badge/MS%20Excel-Advanced%20Formulas%20%26%20KPIs-107C41?style=for-the-badge&logo=microsoftexcel&logoColor=white)](#)
[![Market](https://img.shields.io/badge/Derivatives-NIFTY%20%7C%20BANKNIFTY%20%7C%20SENSEX-FF6B6B?style=for-the-badge)](#)
[![Sensibull](https://img.shields.io/badge/Sensibull-OI%20Change%20Analytics-6366F1?style=for-the-badge)](#)
[![Zero Indicators](https://img.shields.io/badge/Indicators-ZERO-E74C3C?style=for-the-badge)](#)

---

## 📌 Executive Summary

Over 90% of retail option buyers in the Indian market lose money primarily due to two reasons:
1. **Indicator Noise:** Relying on lagging technical indicators on 1-min/5-min charts that produce frequent false breakouts.
2. **Trading Blind Against Sellers:** Ignoring institutional Option Sellers (Call & Put Writers) who dominate expiry levels.

This project presents a **systematic, zero-indicator derivative analytics framework** that combines:
- **Macro Trend & Key Support/Resistance** identified on the **1-Hour Timeframe** (TradingView).
- **Institutional Derivative Data** extracted via **Sensibull Open Interest (OI) Change**.
- **Performance & Risk Analytics** modeled dynamically in **MS Excel** with automated formulas.

---

## 🎯 How the System Works (Core Methodology)

### Step 1: 1-Hour Chart Analysis (TradingView)
- Plot the **1-Hour timeframe** on Nifty, BankNifty, or Sensex.
- Identify consolidation zones formed by a **minimum of 2 candles**.
- Draw horizontal support and resistance lines along these zones.
- **Entry Rule:** Wait patiently for the price to **retest** the marked 1-Hour zone. No trade is taken without a retest.

### Step 2: Sensibull OI Change Ratio Confirmation
- Navigate to **Sensibull → Analyse → Open Interest → OI Change** (full-day tracking).
- Compare the rate of Call Sellers vs. Put Sellers to compute the **OI Ratio**:

| OI Ratio (Call : Put Change) | Market Meaning | System Action |
|---|---|---|
| **1:2 to 1:5** (or + / − Divergence) | **Strong Imbalance** — One side has significantly fewer sellers. Price moves towards the path of least resistance. | ✅ **TRADE** (90.0% Win Rate) |
| **~1:1 or below 1:2** | **Weak / Range** — Sellers are evenly distributed; market chops sideways. | ⚠️ **Small Stop Losses Hit** |
| **Equal both sides** | **No Edge** — Zero directional edge. | ❌ **SKIP / NO TRADE** (Capital Protection) |

### Step 3: Morning Gap Dynamics
- If the index opens with a significant **Gap Up**, Put premiums open depressed (cheap).
- Approximately 90% of the time intraday, the deeply discounted premium recovers towards fair value, creating an early-session edge.
- **Strict Execution Discipline:** Same-day entry and exit. Zero overnight risk.

### Step 4: Quantitative Risk Management & Sizing
- **Defined Stop Loss:**
  - Nifty: 15 points (Option Premium)
  - BankNifty: 30 points (Option Premium)
  - Sensex: 30 points (Option Premium)
- **Position Sizing:** Full allocation deployed at base/support levels; scaled-down sizing at extended/resistance levels to optimize risk.
- **Dynamic Targets:** High OI imbalance ratios (1:3 to 1:5) target 3x–5x of risk; standard ratios target 2x–3x.

---

## 📊 Performance Analytics & Findings

### Overall Execution Summary

| Metric | Value | Significance |
|---|---|---|
| **Total Tracking Sessions** | **65** | Documented audit trail across expiry cycles |
| **Trades Executed** | **57** | Executed when setup criteria were met |
| **Disciplined Skips** | **8** | Protected capital when no directional edge existed |
| **Wins / Losses** | **47 / 10** | High-probability execution distribution |
| **Overall Win Rate** | **82.5%** | Aggregate across all executed sessions |
| **Total Realized Net PnL** | **₹3,72,635** | Net profit after all STT, exchange turnover fees & taxes |
| **Total Transaction Friction** | **₹4,922** | Brokerage and regulatory charges tracked per contract |
| **Profit Factor** | **> 2.3** | Gross profits generated per unit of rupee lost |

---

### The OI Ratio Edge (Key Analytical Proof)

This table demonstrates the direct correlation between institutional OI imbalance and trade profitability:

| Category | Trades | Win Rate | Net PnL Contribution | Key Insight |
|---|---|---|---|---|
| **🔥 Strong Edge** (1:2 to 1:5 Ratio or Divergence) | **50** | **90.0%** | **₹3,66,480** | When institutional sellers are absent on one side, execution accuracy reaches 90.0%. |
| **⚠️ Weak / Range** (~1:1 Ratio) | **7** | **28.6%** | **₹6,155** | Even distribution of sellers produces chop and repeated small stop-losses. |
| **❌ No Edge** (Equal Both Sides) | **8** | **SKIPPED** | **₹0.00** | Strict rule compliance prevented drawdowns on indecisive days. |

---

### Index Performance Breakdown

| Index | Trades | Win Rate | Realized Net PnL | Preferred Instrument Rationale |
|---|---|---|---|---|
| **SENSEX** | 17 | **94.1%** | **₹1,27,010** | High gamma expansion & clean trends on Friday expiry |
| **BANKNIFTY** | 21 | **85.7%** | **₹1,52,059** | Superior liquidity & sharp institutional moves on Wednesday expiry |
| **NIFTY** | 19 | **68.4%** | **₹93,567** | Reliable index trends across Thursday expiry |

---

## 📗 Excel Workbook Architecture (`The_OI_Ratio_Edge_Analytics.xlsx`)

The workbook is structured into 3 distinct operational sheets:

### Sheet 1: `Trading_System_Blueprint`
- Formalizes the 5-step operational framework, entry/exit criteria, and risk allocation logic.

### Sheet 2: `Expiry_Trade_Log`
- Complete chronological audit log of all 65 sessions.
- Clean numeric data types across all computational fields.
- Dynamic formulas tracking:
  - Points Captured: `=Exit_Premium - Entry_Premium`
  - Gross PnL: `=Points_Captured * Quantity`
  - STT & Taxes: `=ROUND(50 + (Entry_Premium + Exit_Premium) * Quantity * 0.0007, 2)`
  - Net PnL: `=Gross_PnL - Charges_STT`
  - Result: `=IF(Net_PnL > 0, "WIN", "LOSS")`

### Sheet 3: `Performance_Dashboard`
- Formula-driven KPI scorecard utilizing advanced Excel functions:
  - `=COUNTIFS()` for granular segment tracking
  - `=SUMIFS()` for index and category PnL attribution
  - `=AVERAGEIF()` for expected payoff modeling
  - Dynamic Profit Factor computation

---

## 🛠️ Tools & Competencies

- **Analytical Tool:** MS Excel (Advanced Formulas, Data Modeling, Pivot Logic, KPI Scorecards)
- **Derivative Analysis:** Open Interest (OI) Analysis, Sensibull Option Chain, Institutional Seller Dynamics, Risk-Reward Modeling
- **Execution Strategy:** Multi-Timeframe Confluence, 1-Hour Chart Price Action, Systematic Drawdown Control

---

## 📬 Contact & Portfolio

- **Ajay Saini**  
- **Email:** [ajayysaiini@gmail.com](mailto:ajayysaiini@gmail.com)  
- **GitHub:** [github.com/ajayysainii](https://github.com/ajayysainii)
