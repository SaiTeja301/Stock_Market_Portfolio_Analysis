---
name: trm-algorithm
id: FIN-03
description: >
  Production-grade quantitative skill for the Teja Recovery Model (TRM) — the user's original algorithmic invention.
  Covers the complete TRM family (Original, V1, V2, V3, V4, V5, V6, V7, V8), including recursive recovery math,
  percentage-triggered allocation, actual-loss adaptation, dual-path profit/loss handling, current-price share laws,
  adaptive actual-P/L calculations, market-regime-aware exit/re-entry orchestration, principal-preserving continuous
  profit harvesting, portfolio wealth-floor protection, cost-aware execution, anti-churn controls, strict dimensional
  integrity, data provenance, walk-forward validation requirements, deterministic state machines, standardized output
  contracts, and Python/Java 17 implementation guidance.
version: 4.0.0
category: Domain - Finance
applicable_stacks: [Universal, Python, Java]
tags:
  - trm-algorithm
  - teja-recovery-model
  - trend-regression-momentum
  - ols-regression-channel
  - mathematical-expectancy
  - quantitative-risk
  - cash-equity
  - capital-allocation
  - position-sizing
  - ledger-accounting
  - zero-cost-runner
  - trm-v4
  - trm-v5
  - trm-v6
  - trm-v7
  - trm-v8
  - adaptive-meta-controller
  - principal-preservation
  - profit-harvesting
  - wealth-floor
  - anti-churn
  - walk-forward-validation
author: "Teja (original algorithm inventor) — documented by AI Agent Skills & Architecture Registry"
inputs:
  - Chart screenshots / OHLCV time series data
  - Ticker symbol & exchange (NSE/BSE, NYSE/NASDAQ)
  - Existing holding data (average price, capital allocated, cycle step, unrealized P/L, harvested cash, protected principal, peak wealth)
  - Account bankroll & risk allocation parameters (base tranche S, max rounds N, position caps, risk budget)
  - Technical and market-state inputs (OHLCV, trend, regression channel, VWAP, EMA/SMA, RSI, MACD, ATR, volume/RVOL)
  - Fundamental/corporate context (debt, liquidity, governance flags, events, sector/index context)
  - Optional external analysis-agent output, treated as untrusted input requiring validation
outputs:
  - Multi-timeframe trend & regression channel analysis
  - TRM model selection/routing across Original, V1, V2, V3, V4, V5, V6, V7, and V8
  - Actual P/L, cycle-wealth, peak-wealth, drawdown and wealth-floor state
  - Exact entry/recovery/harvest/exit triggers, position sizing, and thesis invalidation stops
  - Principal-preserving continuous profit harvesting and dynamic runner logic
  - Exit/re-entry zones, confirmation gates, anti-churn controls, and execution-cost checks
  - Copy-paste-ready V8 master report, audit ledger, and risk scorecard
prerequisites:
  - Cash equity trading account (strictly NO unhedged F&O, margin leverage, or shorting)
  - Verified liquid instruments (large-cap equities, index ETFs, clean balance sheets)
---

# Master TRM (Teja Recovery Model / Trend-Regression-Momentum) Algorithm Skill
## Directory: `.skills/.Stock_Market_Skills/TRM_ Strategies/`
### Version: 4.0 Final Enterprise Adaptive Edition (TRM V8 Master Controller)
### Platform: Multi-Platform (Indian NSE/BSE & U.S. Equities)

---

## Table of Contents
1. [Executive Summary & Architectural Clarification](#1-executive-summary--architectural-clarification)
2. [Agent Persona & Operational Boundaries](#2-agent-persona--operational-boundaries)
3. [Data Integrity & Provenance Tagging Protocol](#3-data-integrity--provenance-tagging-protocol)
4. [Mathematical Engines: Complete TRM Family (Original + V1–V8)](#4-mathematical-engines-complete-trm-family-original--v1v8)
   - 4.1 [TRM Original](#41-trm-original-recursive-cumulative-capital-allocation)
   - 4.2 [TRM V1](#42-trm-v1-percentage-triggered-stock-equity-model)
   - 4.3 [TRM V2](#43-trm-v2-actual-portfolio-loss-adaptive-model)
   - 4.4 [TRM V3](#44-trm-v3-asymmetric-dual-path-model-profit-retention--loss-recovery)
   - 4.5 [TRM V4](#45-trm-v4-50-current-price-pl-share-mechanism)
   - 4.6 [TRM V5](#46-trm-v5-adaptive-actual-pl-dual-path-engine)
   - 4.7 [TRM V6](#47-trm-v6-adaptive-multi-agent-exit--re-entry-rotation-model)
   - 4.8 [TRM V7](#48-trm-v7-principal-preservation--continuous-profit-harvesting)
   - 4.9 [TRM V8](#49-trm-v8-master-adaptive-meta-controller)
   - 4.10 [Universal Invariants & Model Routing Matrix](#410-universal-invariants--model-routing-matrix)
5. [Technical Market Regime & Confluence Engine](#5-technical-market-regime--confluence-engine-trend-regression-momentum)
6. [TRM V8 Deterministic Execution Pipeline](#6-trm-v8-deterministic-execution-pipeline)
7. [V8 Decision Scorecard & Calibration](#7-v8-decision-scorecard--calibration)
   - 7.1 [Evidence Dimensions](#71-evidence-dimensions)
   - 7.2 [Gating vs Scoring](#72-gating-vs-scoring)
   - 7.3 [Calibration Policy](#73-calibration-policy)
   - 7.4 [Legacy 100-Point Rubric](#74-legacy-quantitative-setup-scoring-system-100-point-rubric)
8. [Pre-Flight Execution Checklist](#8-pre-flight-execution-checklist)
9. [Comprehensive Edge Cases & Anomaly Protocols](#9-comprehensive-edge-cases--anomaly-protocols)
10. [Standardized Output Contracts](#10-standardized-output-contracts-copy-paste-ascii-templates)
   - 10.1 [Legacy V2 Report](#101-mode-a-trm-v2-dip-averaging-execution-report)
   - 10.2 [Legacy V3 Report](#102-mode-b-trm-v3-dual-path-execution-report)
   - 10.3 [Multi-Round Audit](#103-mode-c-trm-state-machine--multi-round-portfolio-audit)
   - 10.4 [TRM V7/V8 Master Report](#104-trm-v7v8-master-execution-report)
11. [Production Reference Implementations](#11-production-reference-implementations)
   - 11.1 [Existing Python Reference](#111-python-production-engine-decimal-precise)
   - 11.2 [Existing Java 17 Reference](#112-java-17-enterprise-implementation)
   - 11.3 [V4–V8 Implementation Contracts](#113-v4v8-implementation-contracts)
12. [End-to-End Numerical Case Studies](#12-end-to-end-concrete-numerical-case-studies)
13. [Validation, Backtesting & Deployment Readiness](#13-validation-backtesting--deployment-readiness)
14. [Mandatory Regulatory Compliance & Risk Disclaimer](#14-mandatory-regulatory-compliance--risk-disclaimer)

---

# 1. Executive Summary & Architectural Clarification

### 1.1 Resolution of TRM Nomenclature: The Teja Recovery Model
In quantitative finance and algorithmic portfolio management, ambiguity regarding acronyms can lead to fatal execution errors. In this repository and throughout the foundational research prompts (`TRM_Chart_Discussion.txt` and `TRM-Master-Research-Prompt.txt`), **TRM** stands authoritatively for the **Teja Recovery Model** `[OBSERVED]`.

The Teja Recovery Model is a recursive capital-allocation, dynamic position-sizing, and loss-recovery mathematical framework designed specifically for **cash equity stock investing**. It governs how capital is staged across adverse market outcomes, how drawdowns are neutralized through dynamic averaging or calibrated recovery additions, and how trading cycles unconditionally reset upon reaching defined profit targets.

### 1.2 TRM V8 Final Architecture: Master Adaptive Controller
TRM V8 is the **master controller** for the complete TRM family. It is not a claim of guaranteed profitability and it must not be treated as a single fixed buy/sell formula. V8 routes the current market/portfolio state into the most appropriate specialized TRM engine, then applies portfolio-level risk, data-quality, execution-cost, and wealth-floor constraints before any order can be emitted.

**Core design objective:** maximize decision quality and reduce avoidable capital destruction by using the least aggressive TRM engine sufficient for the validated market state.

```text
MARKET + PORTFOLIO STATE
        |
        v
DATA / UNIT / QUALITY GATES
        |
        v
FUNDAMENTAL + LIQUIDITY GATE
        |
        v
ACTUAL P/L + CYCLE WEALTH ENGINE
        |
        v
MARKET REGIME + STRUCTURAL STATE
        |
        +-------------------+-------------------+
        |                   |                   |
   STRONG BULL          NEUTRAL            BEARISH / RISK
        |                   |                   |
       V7/V5             V4/V5/V6          V2-V8/V6
        |                   |                   |
        +-------------------+-------------------+
                            |
                            v
                 PORTFOLIO WEALTH FLOOR
                            |
                            v
                   RISK / CAPITAL GATE
                            |
                            v
                    COST / LIQUIDITY GATE
                            |
                            v
                     EXECUTION ENGINE
                            |
                            v
                  POST-TRADE LEDGER / MONITOR
                            |
                            v
                     NEW V8 STATE
```

**V8 governing principle:**

> **Detect → Validate → Classify → Select → Calculate → Risk-Gate → Execute → Reconcile → Monitor → Adapt.**

A recommendation is valid only when all mandatory gates pass. `WAIT`, `HOLD`, `REDUCE`, `EXIT`, and `NO_TRADE` are first-class outcomes.

### 1.3 Evolutionary Synthesis of the TRM Family
| Version | Core contribution | V8 keeps | V8 constrains/replaces |
|---|---|---|---|
| Original | Recursive cumulative recovery | Mathematical benchmark | Never default live engine; geometric risk is capped |
| V1 | Percentage-triggered stock adaptation | Trigger concept | No blind geometric capital escalation |
| V2 | Actual monetary-loss adaptation | Actual-loss accounting | Recovery requires quality/structure/capital gates |
| V3 | Dual profit/loss path | Dual-path architecture + recovery surcharge concept | Profit trigger cannot override actual observed P/L |
| V4 | 50% current-price share law | Current-price share conversion | Fixed 50% is a mode, not a universal rule |
| V5 | Actual-P/L + adaptive runner | Actual-P/L authority + current-price law | Runner fraction must be regime-aware |
| V6 | Multi-agent exit/re-entry lifecycle | Validated external analysis + exit/re-entry orchestration | AI output is untrusted until validated |
| V7 | Principal-preserving continuous harvest | Protected principal + excess-only harvesting | Harvesting is disabled under structural/risk deterioration |
| V8 | Adaptive meta-controller | **All compatible components** | Prevents one formula from dominating every regime |

### 1.4 Claims, Evidence & Validation Boundary
The V8 skill is a **decision-support and research architecture**. “More adaptive,” “more risk-aware,” or “more complete” are design properties. They are **not evidence that V8 produces higher returns**. Any claim of superiority must be established by controlled backtesting and out-of-sample/walk-forward evaluation using identical assumptions across versions, including slippage, brokerage, taxes/charges, liquidity constraints, corporate actions, gaps, and rejected orders.


### 1.5 The Dual-Engine Institutional Architecture
To achieve deterministic market outperformance, the system integrates two interlocking subsystems operating in a continuous feedback loop:

```mermaid
flowchart TD
    subgraph MarketFeeds ["1. Market Data Ingestion & Indicator Feeds"]
        direction LR
        OHLCV["OHLCV Price Action (NSE/BSE & US Equities)"]
        Indicators["Technical Indicators (SuperTrend 10:3, 20/50/200 EMAs, RSI, MACD, ATR)"]
        RegChannel["100-Period Linear Regression Channel (±2.0σ Bands)"]
    end

    subgraph SubsystemA ["SUBSYSTEM A: Signal & Timing Engine (Trend-Regression-Momentum)"]
        direction TB
        TrendAudit["1. Multi-Timeframe Trend Audit<br/>(Weekly > 200 SMA & Weekly SuperTrend Bullish)"]
        ValueMap["2. Structural Value Bounds<br/>(Demand at -2.0σ / Rejection at +2.0σ)"]
        TriggerGen["3. Precision Entry & Dip Triggers<br/>(RSI Oversold + Volume Candlestick Rejection)"]
        TrendAudit --> ValueMap --> TriggerGen
    end

    subgraph SubsystemB ["SUBSYSTEM B: Capital Allocation & Recovery Engine (Teja Recovery Model)"]
        direction TB
        ModelSelect{"Model Selection Gate"}
        TRM_V2["TRM V2: Actual-Loss Adaptive<br/>(B_next = 2 × ActualLoss)"]
        TRM_V3["TRM V3: Asymmetric Dual-Path<br/>(Runner Retention + Scaled Loss Recovery)"]
        TrancheMath["Dynamic Position & Tranche Sizing<br/>(Shares = ⌊Allocation / LimitPrice⌋)"]
        ResetGate["Universal Reset Invariant<br/>(Unconditional Reset to S upon Target Hit)"]
        
        ModelSelect -- "Core Blue-Chip Dip Averaging" --> TRM_V2 --> TrancheMath
        ModelSelect -- "Active Swing & Runner Accumulation" --> TRM_V3 --> TrancheMath
        TrancheMath --> ResetGate
    end

    subgraph PortfolioTreasury ["Portfolio Capital & Execution Infrastructure"]
        direction LR
        CashPool[("Dedicated Cash Bankroll (W)<br/>(Strict 15% Cap per Ticker)")]
        OrderRouter["Cash-Equity Broker Order Router<br/>(Strictly NO Leverage / NO F&O)"]
        RunnerVault[("Compounding Equity Vault<br/>(Zero-Cost 'House Money' Shares)")]
    end

    MarketFeeds --> SubsystemA
    SubsystemA -- "Qualified Timing Signal (WHEN)" --> SubsystemB
    CashPool <--> SubsystemB
    SubsystemB -- "Exact Order Parameters (HOW MUCH)" --> OrderRouter
    ResetGate -- "Transfer Retained Runner Shares" --> RunnerVault

    classDef feed fill:#1E293B,stroke:#94A3B8,stroke-width:1px,color:#F8FAFC;
    classDef subA fill:#0F172A,stroke:#38BDF8,stroke-width:2px,color:#F8FAFC;
    classDef subB fill:#1E1B4B,stroke:#818CF8,stroke-width:2px,color:#FFFFFF;
    classDef treasury fill:#064E3B,stroke:#34D399,stroke-width:2px,color:#ECFDF5;
    classDef gate fill:#312E81,stroke:#C7D2FE,stroke-width:2px,color:#FFFFFF;

    class OHLCV,Indicators,RegChannel feed;
    class TrendAudit,ValueMap,TriggerGen subA;
    class TRM_V2,TRM_V3,TrancheMath,ResetGate subB;
    class ModelSelect gate;
    class CashPool,OrderRouter,RunnerVault treasury;
```

> [!NOTE]
> **Functional Separation of Concerns**:
> - **Subsystem A (Trend-Regression-Momentum)**: Evaluates macro trend health, calculates mean-reversion regression bounds, and filters for momentum exhaustion to determine **WHEN** to trade.
> - **Subsystem B (Teja Recovery Model)**: Computes the mathematical capital tranches, share allocations, break-even shift, and risk-adjusted recovery targets to determine **HOW MUCH** capital to deploy.

---

# 2. Agent Persona & Operational Boundaries

### 2.1 Multi-Role Persona Declaration
You are an elite **Senior Quantitative Portfolio Manager, Algorithmic Systems Architect, Financial Mathematician, and Risk Management Specialist**. Your primary mandate is to evaluate cash equity assets, analyze multi-timeframe price action, map regression channels, compute staged recovery capital allocations, and enforce deterministic risk controls under the **Teja Recovery Model (TRM)**.

### 2.2 Institutional Trading Priority Stack
Every analytical assessment and trade proposal MUST obey this strict, non-negotiable priority sequence:

```mermaid
flowchart TD
    subgraph PriorityStack ["INSTITUTIONAL TRADING PRIORITY STACK (NON-NEGOTIABLE EXECUTION HIERARCHY)"]
        direction TB
        L1["TIER 1: CAPITAL PRESERVATION (ZERO RUIN)<br/><b>Absolute Sovereign Mandate</b>: Eliminating catastrophic drawdown overrides all return objectives"]
        L2["TIER 2: STRICT CASH-EQUITY BOUNDARY<br/><b>Vehicle Constraint</b>: Unleveraged spot shares only — Zero margin, zero F&O, zero liquidation risk"]
        L3["TIER 3: STRUCTURAL INVALIDATION STOP<br/><b>Safety Floor</b>: Hard price floor where technical thesis is voided — Absolute trade termination"]
        L4["TIER 4: STAGED CAPITAL TRANCHE SIZING<br/><b>Position Sizing</b>: TRM algorithmic allocation (V2 loss-adaptive or V3 surcharge) — No arbitrary sizing"]
        L5["TIER 5: MATHEMATICAL EXPECTANCY (R:R)<br/><b>Edge Validation</b>: Positive probability edge with asymmetric payoff ratio (R:R ≥ 1:1.5)"]
        L6["TIER 6: DISCIPLINED CYCLE MANAGEMENT<br/><b>State Discipline</b>: Completed recovery/exit cycles reset their escalation state to base S; V7/V8 may remain in an active harvesting cycle until a valid exit/reset condition occurs."]

        L1 ==>|Overrules Lower Tiers| L2
        L2 ==>|Enforces Safe Scope| L3
        L3 ==>|Protects Bankroll| L4
        L4 ==>|Calibrates Risk| L5
        L5 ==>|Yields Cycle Victory| L6
    end

    classDef t1 fill:#7F1D1D,stroke:#F87171,stroke-width:2px,color:#FEF2F2;
    classDef t2 fill:#78350F,stroke:#FBBF24,stroke-width:2px,color:#FFFBEB;
    classDef t3 fill:#831843,stroke:#F472B6,stroke-width:2px,color:#FDF2F8;
    classDef t4 fill:#1E1B4B,stroke:#818CF8,stroke-width:2px,color:#EEF2FF;
    classDef t5 fill:#0C4A6E,stroke:#38BDF8,stroke-width:2px,color:#F0F9FF;
    classDef t6 fill:#064E3B,stroke:#34D399,stroke-width:2px,color:#ECFDF5;

    class L1 t1;
    class L2 t2;
    class L3 t3;
    class L4 t4;
    class L5 t5;
    class L6 t6;
```

1. **Capital Preservation (Zero Ruin)**: Preventing catastrophic drawdown takes absolute precedence over any profit target or recovery attempt.
2. **Strict Cash-Equity Boundary**: Trading is restricted solely to unleveraged common equity shares.
3. **Structural Invalidation Stop**: Every trade must have a pre-calculated, non-negotiable price floor where the thesis is voided.
4. **Staged Capital Tranche Sizing**: Position sizing is governed strictly by the TRM mathematical engines.
5. **Mathematical Expectancy ($R:R$)**: Every cycle must possess a positive mathematical expectation with an asymmetry of at least $1:1.5$.
6. **Disciplined Cycle Management**: Completed recovery/exit cycles reset their escalation state to base $S$; V7/V8 may remain in an active harvesting cycle until a valid exit/reset condition occurs.

### 2.3 Strict Cash-Equity Operational Scope
- **Permitted Asset Class**: **Cash Equity / Spot Market Only** (holding unleveraged shares in delivery accounts).
- **Strictly Prohibited Instruments**:
  - ❌ Futures & Options (F&O)
  - ❌ Leveraged ETFs / Inverse Leveraged Funds
  - ❌ Intraday Margin / Margin Trading Facility (MTF) Borrowing
  - ❌ Naked Short Selling
  - ❌ Speculative Penny Stocks (Micro-caps < ₹500 Cr / $100M Market Cap)
  - ❌ Cryptocurrencies, Forex Pairs, and Binary Options

### 2.4 Negative Constraints & Behavioral Guardrails
- **No Profit Guarantees**: Never assert or imply that TRM eliminates risk or guarantees profit. Markets are non-stationary and stochastic.
- **Anti-Anchoring Mandate**: Never anchor forward analysis to the user's historical purchase price. Assess the asset strictly on current market structure and liquidity.
- **No Uncapped Averaging**: Averaging down without a maximum round limit is prohibited. Hard stop is enforced at **maximum 3 to 5 rounds**.
- **No Averaging into Impaired Assets**: If an asset faces fraud allegations, promoter pledging > 50%, debt default, or regulatory bans, averaging is permanently banned and immediate liquidation is mandated.
- **Never Move Stops**: A structural stop-loss can never be widened or deleted during drawdown.

### 2.5 Asset Suitability & Fundamental Quality Gate
Before any TRM capital allocation is calculated, the target instrument must pass the **Cash Equity Quality Gate**:
1. **Market Capitalization**: Large-Cap or High-Quality Mid-Cap (Nifty 50, Nifty Next 50, S&P 500, Nasdaq 100 components preferred).
2. **Solvency & Debt**: Debt-to-Equity Ratio $< 0.80$ (exempting banking/NBFC institutions which are screened for Gross NPA $< 2.5\%$).
3. **Promoter Integrity (Indian Equities)**: Promoter Pledging $= 0.0\%$ (or $< 5\%$ with stable institutional holding).
4. **Liquidity / ADV**: 20-day Average Daily Volume $> 500,000$ shares and Average Daily Turnover $> ₹10 \text{ Cr}$ ($\$5\text{M}$ for US equities).
5. **Bankruptcy Immunity**: The company must demonstrate positive operating cash flows (OCF) in at least 3 of the past 4 years.

---

# 3. Data Integrity & Provenance Tagging Protocol

### 3.1 Quad-Tagging System Specification
Every price, volume, indicator reading, calculation, and assumption in the analysis must be tagged with explicit data provenance:

- **`[OBSERVED]`**: Facts directly visible on the price chart, verified quote, or raw OHLCV dataset.
  - *Examples*: Last Traded Price ₹1,420.50 `[OBSERVED]`, Session Volume 4.2M `[OBSERVED]`, Daily 20 EMA ₹1,390.00 `[OBSERVED]`.
- **`[CALCULATED]`**: Values derived mathematically via formal quantitative equations.
  - *Examples*: TRM Tranche 3 Size = ₹600.00 `[CALCULATED]`, Volume-Weighted Average Cost = ₹14.28 `[CALCULATED]`, Retained Quantity = 1.9565 shares `[CALCULATED]`.
- **`[USER-PROVIDED]`**: Values supplied directly by the user's prompt or account context.
  - *Examples*: Account Bankroll $W = ₹5,00,000$ `[USER-PROVIDED]`, Base Tranche $S = ₹10,000$ `[USER-PROVIDED]`, Original Entry Price ₹18.00 `[USER-PROVIDED]`.
- **`[ASSUMED]`**: Values inferred due to unobservable context, missing data, or baseline models.
  - *Examples*: Normal execution liquidity assumed `[ASSUMED]`, Zero slippage assumed on limit order `[ASSUMED]`.

### 3.2 Data Insufficiency & Degraded Mode Handling
If key market data (e.g., volume profile, index beta, ATR) is missing or unreadable from an image:
1. Do **NOT** invent or hallucinate missing data.
2. Emit the standardized banner:
   > **⚠️ Data Limitation Warning**: Missing [Volume / VWAP / ATR]. Analysis proceeds under baseline technical assumptions. Quality score confidence capped at `MEDIUM`.
3. Fall back to visible structural swing levels only.

### 3.3 The Strict Dimensional Integrity Law
In financial algorithms, mixing units leads to catastrophic calculation errors. The TRM algorithm strictly enforces the **Dimensional Separation Law** `[OBSERVED]`:

$$\boxed{\text{Invested Capital } (\$) \xrightarrow{\times \text{P/L Rate } (\%)} \text{Monetary P/L } (\$) \xrightarrow{\times \text{Factor } (\%)} \text{Retention/Recovery } (\$) \xrightarrow{\div P_{\text{current}} (\$/\text{share})} \text{Share Quantity}}$$

**Prohibited Arithmetic Operations**:
- ❌ NEVER add a Currency Amount to a Share Quantity ($₹100 + 10 \text{ shares} = \text{ERROR}$).
- ❌ NEVER add a Percentage to a Currency Amount ($₹1,000 + 15\% = \text{ERROR}$; must multiply: $₹1,000 \times 0.15 = ₹150$).
- ❌ NEVER divide Retained Profit Amount by Entry Price $P_0$ (MUST divide by Current Price $P_1$).

---

# 4. Mathematical Engines: All 4 TRM Editions

```mermaid
flowchart LR
    subgraph TRM_Orig ["TRM Original: Recursive Cumulative"]
        direction TB
        OrigDesc["<b>Recursive Bankroll Progression</b><br/>Rule: $B_{n+1} = 2 T_n$<br/>Scaling: $T_n = S \times 3^{n-1}$<br/>Complexity: $\mathcal{O}(3^n)$ Exponential<br/><i>Abstract Math — Ruin risk under finite bankroll</i>"]
    end

    subgraph TRM_1 ["TRM V1: Percentage-Triggered"]
        direction TB
        V1Desc["<b>Stock Market Dip Adaptation</b><br/>Trigger: $P_n = P_0(1-L)^n$<br/>Scaling: Geometric Tranches $\mathcal{O}(3^n)$<br/>Intensity: ₹5.90M for 10 drops of 20%<br/><i>Impractical capital requirements</i>"]
    end

    subgraph TRM_2 ["TRM V2: Loss-Adaptive Averaging"]
        direction TB
        V2Desc["<b>Actual-Loss Dynamic Scaling</b><br/>Rule: $B_{n+1} = m \times (T_n - Q_n P_n)$<br/>Deployment: ₹18,712.11 for 10 drops<br/><b>99.68% Capital Reduction</b><br/><i>Institutional Cash Equity Viable</i>"]
    end

    subgraph TRM_3 ["TRM V3: Asymmetric Dual-Path"]
        direction TB
        V3Desc["<b>Position Retention & Reversal Engine</b><br/>Path A: Free Runner Retention ($Q_{\text{retain}}$)<br/>Path B: Calibrated Stop-Loss Surcharge<br/>Invariant: Unconditional Reset to Base $S$<br/><i>Perpetual Equity Compounding</i>"]
    end

    TRM_Orig ==>|"1. Stock Market Dip Adaptation"| TRM_1
    TRM_1 ==>|"2. 99.68% Mathematical Optimization"| TRM_2
    TRM_2 ==>|"3. Bidirectional Asymmetric Expansion"| TRM_3

    classDef origBox fill:#1E293B,stroke:#94A3B8,stroke-width:2px,color:#F8FAFC;
    classDef v1Box fill:#312E81,stroke:#818CF8,stroke-width:2px,color:#EEF2FF;
    classDef v2Box fill:#064E3B,stroke:#34D399,stroke-width:2px,color:#ECFDF5;
    classDef v3Box fill:#1E1B4B,stroke:#F59E0B,stroke-width:2px,color:#FFFBEB;

    class OrigDesc origBox;
    class V1Desc v1Box;
    class V2Desc v2Box;
    class V3Desc v3Box;
```

---

### 4.1 TRM Original: Recursive Cumulative Capital Allocation

#### 4.1.1 Foundational Premise & Recurrence Relation
The Original Teja Recovery Model is a pure mathematical progression designed to recover all prior accumulated losses with a single recovery win, yielding a cycle profit equal to the base investment scaled by the round index.

**Input Variables**:
- $S$: Initial base investment (reset amount), $S > 0$.
- $N$: Target maximum consecutive losing rounds planned to withstand.
- $B_n$: Capital allocated at the $n$-th step.
- $T_n$: Total cumulative capital invested after $n$ losing steps ($T_n = \sum_{i=1}^n B_i$).

**Governing Recurrence Equations**:
$$B_1 = S$$
$$B_{n+1} = 2 \times T_n = 2 \times \sum_{i=1}^n B_i \quad (n \ge 1)$$
$$T_n = T_{n-1} + B_n$$

#### 4.1.2 First-Principles Derivation & Closed-Form Proof

To understand why doubling cumulative losses ($B_{n+1} = 2 T_n$) forces an exponential base-3 progression, we construct the rigorous first-principles derivation from Step 1 through Step 5.

Let $S$ denote the initial base capital allocation ($S > 0$).

##### A. Step-by-Step Derivation Matrix

| Step ($n$) | Stage Description | Allocation Rule ($B_n$) | Tranche Sizing Math | Tranche Value ($B_n$) | Cumulative Capital Math ($T_n = T_{n-1} + B_n$) | Total Invested ($T_n$) | Closed-Form ($S \cdot 3^k$) | Step Ratio ($T_n / T_{n-1}$) |
| :---: | :--- | :--- | :--- | :---: | :--- | :---: | :---: | :---: |
| **Step 1** | Baseline Position Entry | $B_1 = S$ | Baseline initialization | $\mathbf{1S}$ | $T_1 = B_1 = S$ | $\mathbf{1S}$ | $\mathbf{S \cdot 3^0}$ | Baseline ($1\times$) |
| **Step 2** | Round 1 Loss Recovery | $B_2 = 2 \times T_1$ | $2 \times (S)$ | $\mathbf{2S}$ | $T_2 = T_1 + B_2 = S + 2S$ | $\mathbf{3S}$ | $\mathbf{S \cdot 3^1}$ | $\mathbf{3\times}$ |
| **Step 3** | Round 2 Loss Recovery | $B_3 = 2 \times T_2$ | $2 \times (3S)$ | $\mathbf{6S}$ | $T_3 = T_2 + B_3 = 3S + 6S$ | $\mathbf{9S}$ | $\mathbf{S \cdot 3^2}$ | $\mathbf{3\times}$ |
| **Step 4** | Round 3 Loss Recovery | $B_4 = 2 \times T_3$ | $2 \times (9S)$ | $\mathbf{18S}$ | $T_4 = T_3 + B_4 = 9S + 18S$ | $\mathbf{27S}$ | $\mathbf{S \cdot 3^3}$ | $\mathbf{3\times}$ |
| **Step 5** | Round 4 Loss Recovery | $B_5 = 2 \times T_4$ | $2 \times (27S)$ | $\mathbf{54S}$ | $T_5 = T_4 + B_5 = 27S + 54S$ | $\mathbf{81S}$ | $\mathbf{S \cdot 3^4}$ | $\mathbf{3\times}$ |

---

##### B. Step-by-Step Mathematical Walkthrough

Each step computes the new tranche allocation and tracks the cumulative committed bankroll:

1. **Step 1: Baseline Position Deployment ($n = 1$)**
   - **Tranche Allocation**:
     $$B_1 = S$$
   - **Cumulative Capital Committed**:
     $$T_1 = B_1 = S = \mathbf{S \cdot 3^0}$$
   - *State*: Cycle opens with 1 unit of base capital.

2. **Step 2: Recovery Deployment after 1 Loss ($n = 2$)**
   - **Tranche Sizing Rule**: Double prior cumulative invested capital ($2 \times T_1$):
     $$B_2 = 2 \times T_1 = 2 \times (S) = \mathbf{2S}$$
   - **Updated Cumulative Capital**:
     $$T_2 = T_1 + B_2 = S + 2S = 3S = \mathbf{S \cdot 3^1}$$
   - *Step Multiplier*: $\frac{T_2}{T_1} = \frac{3S}{S} = \mathbf{3}$ *(Total cumulative capital triples)*.

3. **Step 3: Recovery Deployment after 2 Losses ($n = 3$)**
   - **Tranche Sizing Rule**: Double prior cumulative invested capital ($2 \times T_2$):
     $$B_3 = 2 \times T_2 = 2 \times (3S) = \mathbf{6S}$$
   - **Updated Cumulative Capital**:
     $$T_3 = T_2 + B_3 = 3S + 6S = 9S = \mathbf{S \cdot 3^2}$$
   - *Step Multipliers*:
     $$\frac{B_3}{B_2} = \frac{6S}{2S} = \mathbf{3} \quad \text{and} \quad \frac{T_3}{T_2} = \frac{9S}{3S} = \mathbf{3}$$

4. **Step 4: Recovery Deployment after 3 Losses ($n = 4$)**
   - **Tranche Sizing Rule**: Double prior cumulative invested capital ($2 \times T_3$):
     $$B_4 = 2 \times T_3 = 2 \times (9S) = \mathbf{18S}$$
   - **Updated Cumulative Capital**:
     $$T_4 = T_3 + B_4 = 9S + 18S = 27S = \mathbf{S \cdot 3^3}$$
   - *Step Multipliers*:
     $$\frac{B_4}{B_3} = \frac{18S}{6S} = \mathbf{3} \quad \text{and} \quad \frac{T_4}{T_3} = \frac{27S}{9S} = \mathbf{3}$$

5. **Step 5: Recovery Deployment after 4 Losses ($n = 5$)**
   - **Tranche Sizing Rule**: Double prior cumulative invested capital ($2 \times T_4$):
     $$B_5 = 2 \times T_4 = 2 \times (27S) = \mathbf{54S}$$
   - **Updated Cumulative Capital**:
     $$T_5 = T_4 + B_5 = 27S + 54S = 81S = \mathbf{S \cdot 3^4}$$
   - *Step Multipliers*:
     $$\frac{B_5}{B_4} = \frac{54S}{18S} = \mathbf{3} \quad \text{and} \quad \frac{T_5}{T_4} = \frac{81S}{27S} = \mathbf{3}$$

---

##### C. Mathematical Induction Proof of the Tripling Invariant

> [!IMPORTANT]
> **The Fundamental Tripling Invariant of TRM Original**:
> Because the allocation rule mandates investing **twice the entire cumulative sum** of all previous rounds ($B_{n+1} = 2 T_n$), adding this new tranche to the existing portfolio yields:
> $$T_{n+1} = T_n + B_{n+1} = T_n + 2T_n = (1 + 2) T_n = \mathbf{3 \times T_n}$$
> Therefore, every consecutive round unconditionally **triples** the total cumulative capital invested.

**Formal Proof by Induction**:
1. **Base Case ($n = 1$)**:
   $$T_1 = S = S \cdot 3^{1 - 1} = S \cdot 3^0 = S \quad \checkmark \text{ (Base case holds)}$$

2. **Inductive Hypothesis**:
   Assume the closed-form equation holds for step $k$, where $k \ge 1$:
   $$T_k = S \cdot 3^{k - 1}$$

3. **Inductive Step ($k \to k + 1$)**:
   By the TRM Original recurrence relations:
   $$B_{k+1} = 2 \times T_k = 2 \cdot (S \cdot 3^{k - 1})$$
   $$T_{k+1} = T_k + B_{k+1} = (S \cdot 3^{k - 1}) + 2 \cdot (S \cdot 3^{k - 1})$$
   Factoring out $(S \cdot 3^{k - 1})$:
   $$T_{k+1} = (1 + 2) \cdot (S \cdot 3^{k - 1}) = 3 \cdot (S \cdot 3^{k - 1}) = \mathbf{S \cdot 3^k} = \mathbf{S \cdot 3^{(k+1) - 1}}$$
   Thus, the formula holds for $k + 1$. By mathematical induction:
   $$\mathbf{T_n = S \times 3^{n-1} \quad \forall \; n \ge 1} \quad \blacksquare$$

4. **Constant Tripling Growth Ratios**:
   - **Cumulative Capital Growth Ratio**:
     $$\frac{T_n}{T_{n-1}} = \frac{S \cdot 3^{n-1}}{S \cdot 3^{n-2}} = \mathbf{3} \quad (n \ge 2) \quad [CALCULATED]$$
   - **Tranche Allocation Growth Ratio**:
     $$\frac{B_n}{B_{n-1}} = \frac{2S \cdot 3^{n-2}}{2S \cdot 3^{n-3}} = \mathbf{3} \quad (n \ge 3) \quad [CALCULATED]$$

---

##### D. Governing Closed-Form Equations Reference

| Metric / Parameter | Variable | Governing Formula | Provenance | Concrete Example ($S = \$5, n = 4$) |
| :--- | :---: | :--- | :---: | :---: |
| **Cumulative Invested Capital** | $T_n$ | $\mathbf{T_n = S \times 3^{n-1}}$ | `[CALCULATED]` | $5 \times 3^3 = \mathbf{\$135.00}$ |
| **Individual Tranche Allocation** | $B_n$ | $\mathbf{B_n = \begin{cases} S & n = 1 \\ 2S \times 3^{n-2} & n \ge 2 \end{cases}}$ | `[CALCULATED]` | $2(5) \times 3^2 = \mathbf{\$90.00}$ |
| **Next Recovery Tranche Required** | $B_{N+1}$ | $\mathbf{B_{N+1} = 2S \times 3^{N-1}}$ | `[CALCULATED]` | $2(5) \times 3^3 = \mathbf{\$270.00}$ ($N=4$) |
| **Total Bankroll Capacity Required** | $M$ | $\mathbf{M = T_N + B_{N+1} = S \times 3^N}$ | `[CALCULATED]` | $5 \times 3^4 = \mathbf{\$405.00}$ ($N=4$) |
| **Maximum Surviving Loss Streak** | $n_{\max}$ | $\mathbf{n_{\max} = \left\lfloor \log_3\left(\frac{W}{S}\right) \right\rfloor}$ | `[CALCULATED]` | For $W = \$405 \implies \lfloor \log_3(81) \rfloor = \mathbf{4}$ |

#### 4.1.3 Sequence Verification Table ($S = \$5, N = 10$)
| Step ($n$) | Allocation ($B_n$) | Formula ($2S \cdot 3^{n-2}$) | Cumulative Invested ($T_n$) | Closed-Form ($S \cdot 3^{n-1}$) |
| :---: | :---: | :---: | :---: | :---: |
| 1 | \$5.00 | $S$ | \$5.00 | $5 \times 3^0 = \$5.00$ |
| 2 | \$10.00 | $2 \times 5 \times 3^0$ | \$15.00 | $5 \times 3^1 = \$15.00$ |
| 3 | \$30.00 | $2 \times 5 \times 3^1$ | \$45.00 | $5 \times 3^2 = \$45.00$ |
| 4 | \$90.00 | $2 \times 5 \times 3^2$ | \$135.00 | $5 \times 3^3 = \$135.00$ |
| 5 | \$270.00 | $2 \times 5 \times 3^3$ | \$405.00 | $5 \times 3^4 = \$405.00$ |
| 6 | \$810.00 | $2 \times 5 \times 3^4$ | \$1,215.00 | $5 \times 3^5 = \$1,215.00$ |
| 7 | \$2,430.00 | $2 \times 5 \times 3^5$ | \$3,645.00 | $5 \times 3^6 = \$3,645.00$ |
| 8 | \$7,290.00 | $2 \times 5 \times 3^6$ | \$10,935.00 | $5 \times 3^7 = \$10,935.00$ |
| 9 | \$21,870.00 | $2 \times 5 \times 3^7$ | \$32,805.00 | $5 \times 3^8 = \$32,805.00$ |
| 10 | \$65,610.00 | $2 \times 5 \times 3^8$ | \$98,415.00 | $5 \times 3^9 = \$98,415.00$ |
| **Next ($n=11$)** | **\$196,830.00** | $2 \times T_{10}$ | — | — |
| **Total Bankroll ($M$)** | — | — | **\$295,245.00** | $\mathbf{5 \times 3^{10} = \$295,245.00}$ |

#### 4.1.4 Comparison with Classic Martingale
- **Classic Martingale**: Doubles the *previous* bet ($B_{n+1} = 2 B_n$). Capital scaling is geometric order $\mathbf{O(2^n)}$. Total bankroll after 10 losses for $S=5$ is $5 \times (2^{11}-1) = \$10,235$.
- **TRM Original**: Doubles the *cumulative* invested capital ($B_{n+1} = 2 T_n$), causing tranches to triple ($B_{n+1} = 3 B_n$). Capital scaling is geometric order $\mathbf{O(3^n)}$. Total bankroll after 10 losses for $S=5$ is $\mathbf{\$295,245}$ (nearly $29\times$ more capital-intensive than Martingale).
- **Ruin Paradox**: In finite human or institutional portfolios, unconstrained TRM Original leads to mathematical certainty of ruin ($P_{\text{ruin}} \to 1.0$) as streak length approaches bankroll limits. This mathematically necessitates TRM V2 and V3.

---

### 4.2 TRM V1: Percentage-Triggered Stock Equity Model

#### 4.2.1 Conceptual Bridge
TRM V1 transitions the abstract recovery model to stock investing by defining a **predefined percentage stock price decline** ($L$) as the trigger for successive allocation tranches `[OBSERVED]`:
> *"Instead of waiting for an asset to become worthless, trigger the next TRM tranche whenever the stock price declines by a predefined threshold (e.g., $L = 20\%$)."*

#### 4.2.2 Mathematical State Engine

##### A. State Variables & Input Vector Reference

| Variable | Domain | Financial / Mathematical Meaning | Provenance |
| :---: | :---: | :--- | :---: |
| $P_0$ | $P_0 > 0$ | Initial base purchase price at Round 1 | `[USER-PROVIDED]` |
| $L$ | $0 < L < 1.0$ | Predefined percentage price drop trigger (default $0.20 = 20\%$) | `[USER-PROVIDED]` |
| $S$ | $S > 0$ | Initial base cash investment amount ($B_1 = S$) | `[USER-PROVIDED]` |
| $n$ | $n \in \mathbb{N}_{\ge 1}$ | Execution round index | `[CALCULATED]` |
| $P_n$ | $P_n > 0$ | Trigger stock price at round $n$ | `[CALCULATED]` |
| $B_n$ | $B_n > 0$ | Capital tranche allocated at round $n$ | `[CALCULATED]` |
| $T_n$ | $T_n = \sum_{i=1}^n B_i$ | Total cumulative invested capital after $n$ rounds | `[CALCULATED]` |
| $Q_n$ | $Q_n \ge 0$ | Delivery shares purchased in tranche $n$ | `[CALCULATED]` |
| $Q_{\text{total}, n}$ | $Q_{\text{total}, n} = \sum_{i=1}^n Q_i$ | Total cumulative delivery shares accumulated | `[CALCULATED]` |
| $AC_n$ | $AC_n > 0$ | Volume-weighted average cost per share | `[CALCULATED]` |
| $MV_n$ | $MV_n \ge 0$ | Current mark-to-market portfolio liquidation value | `[CALCULATED]` |
| $PnL_n$ | $PnL_n \in \mathbb{R}$ | Unrealized monetary profit / loss | `[CALCULATED]` |

---

##### B. Governing Mathematical State Engine Reference

| Metric / Parameter | Symbol | Exact Governing Equation | Unit | Provenance | Operational Role |
| :--- | :---: | :--- | :---: | :---: | :--- |
| **Drop Trigger Price** | $P_n$ | $$\mathbf{P_n = P_0 \times (1 - L)^n}$$ | $\$/\text{share}$ | `[CALCULATED]` | Determines exact price level where tranche $n+1$ fires |
| **Tranche Allocation Sizing** | $B_{n+1}$ | $$\mathbf{B_{n+1} = 2 \times T_n = 2 \times \sum_{i=1}^n B_i}$$ | $\$$ | `[CALCULATED]` | Doubles cumulative capital; causes tranches to triple ($B_{n+1} = 3 B_n$) |
| **New Shares Acquired** | $Q_n$ | $$\mathbf{Q_n = \frac{B_n}{P_{n-1}} \quad \text{or} \quad \frac{B_n}{P_n}}$$ | shares | `[CALCULATED]` | Number of delivery shares acquired at execution limit |
| **Cumulative Share Holdings** | $Q_{\text{total}, n}$ | $$\mathbf{Q_{\text{total}, n} = \sum_{i=1}^n Q_i}$$ | shares | `[CALCULATED]` | Running inventory of unleveraged common equity |
| **Volume-Weighted Avg Cost** | $AC_n$ | $$\mathbf{AC_n = \frac{T_n}{Q_{\text{total}, n}} = \frac{\sum_{i=1}^n B_i}{\sum_{i=1}^n Q_i}}$$ | $\$/\text{share}$ | `[CALCULATED]` | Exact breakeven exit price floor for full recovery |
| **Mark-to-Market Valuation** | $MV_n$ | $$\mathbf{MV_n = Q_{\text{total}, n} \times P_n}$$ | $\$$ | `[CALCULATED]` | Instantaneous liquidation value of stock holdings |
| **Unrealized Monetary Loss** | $PnL_n$ | $$\mathbf{PnL_n = MV_n - T_n = Q_{\text{total}, n} P_n - T_n}$$ | $\$$ | `[CALCULATED]` | Actual portfolio drawdown in currency units |

---

##### C. Step-by-Step State Transition Calculus

At each successive price drop round $k$ ($k \ge 2$), the mathematical state vector transitions deterministically:

1. **Price Trajectory**:
   $$P_k = P_{k-1} \times (1 - L) = P_0 \times (1 - L)^k$$
2. **Tranche Deployment**:
   $$B_k = 2 \times T_{k-1} = 2S \times 3^{k-2}$$
3. **Cumulative Capital Committed**:
   $$T_k = T_{k-1} + B_k = S \times 3^{k-1}$$
4. **Volume-Weighted Cost Neutralization**:
   $$AC_k = \frac{S \times 3^{k-1}}{\sum_{i=1}^k \frac{B_i}{P_{i-1}}}$$
   Because stock price $P_i$ decreases while tranche allocation $B_i$ scales geometrically as $\mathcal{O}(3^i)$, the latest tranches acquire proportionally massive share volumes, rapidly shifting the average cost $AC_k$ toward $P_k$.

#### 4.2.3 Comprehensive Walkthrough ($P_0 = ₹18.00, S = ₹100, L = 20\%$)
| Round | Price ($P$) | Investment ($B_n$) | Shares Bought | Total Invested ($T_n$) | Total Shares ($Q$) | Avg Cost ($AC$) | Market Value ($MV$) | Unrealized P/L |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| 1 | ₹18.00 | ₹100.00 | 5.555556 | ₹100.00 | 5.555556 | ₹18.00 | ₹100.00 | ₹0.00 |
| 2 | ₹14.40 | ₹200.00 | 13.888889 | ₹300.00 | 19.444444 | ₹15.43 | ₹280.00 | -₹20.00 |
| 3 | ₹11.52 | ₹600.00 | 52.083333 | ₹900.00 | 71.527778 | ₹12.58 | ₹824.00 | -₹76.00 |
| 4 | ₹9.22 | ₹1,800.00 | 195.312500 | ₹2,700.00 | 266.840278 | ₹10.12 | ₹2,459.20 | -₹240.80 |
| 5 | ₹7.37 | ₹5,400.00 | 732.421875 | ₹8,100.00 | 999.262153 | ₹8.11 | ₹7,367.36 | -₹732.64 |
| 6 | ₹5.90 | ₹16,200.00 | 2,746.582031 | ₹24,300.00 | 3,745.844184 | ₹6.49 | ₹22,093.89 | -₹2,206.11 |
| 7 | ₹4.72 | ₹48,600.00 | 10,299.682617 | ₹72,900.00 | 14,045.526802 | ₹5.19 | ₹66,275.11 | -₹6,624.89 |
| 8 | ₹3.77 | ₹145,800.00 | 38,623.809814 | ₹218,700.00 | 52,669.336616 | ₹4.15 | ₹198,820.09 | -₹19,879.91 |
| 9 | ₹3.02 | ₹437,400.00 | 144,839.286804 | ₹656,100.00 | 197,508.623421 | ₹3.32 | ₹596,456.07 | -₹59,643.93 |
| 10 | ₹2.42 | ₹1,312,200.00 | 543,147.325516 | ₹1,968,300.00 | 740,655.948937 | ₹2.66 | ₹1,789,364.86 | -₹178,935.14 |
| 11 | ₹1.93 | ₹3,936,600.00 | 2,036,802.470684 | ₹5,904,900.00 | 2,777,458.419620 | ₹2.13 | ₹5,368,091.89 | -₹536,808.11 |

#### 4.2.4 Proof of the Fatal Flaw in TRM V1
Although TRM V1 successfully drives the average cost down aggressively (from ₹18.00 to ₹2.13), its capital allocation formula ($B_{n+1} = 2 T_n$) operates under the false assumption that all prior capital has been completely lost. 

In reality, after 10 consecutive 20% drops, the stock has declined by $89.28\%$ (from ₹18.00 to ₹1.93), yet the position retains **₹5,368,091.89** in liquidation market value against **₹5,904,900.00** invested. The actual portfolio unrealized loss is only **₹536,808.11**. Demanding a next allocation of **₹11,809,800.00** is an irrational, catastrophic capital escalation that renders TRM V1 unsuitable for live equity trading.

---

### 4.3 TRM V2: Actual-Portfolio-Loss Adaptive Model

#### 4.3.1 Core Breakthrough & Philosophy
TRM V2 resolves the fatal flaw of V1 by introducing the **Actual-Loss Adaptive Principle** `[OBSERVED]`:
> *"Calculate how much money your existing stock position is currently losing in actual currency, then invest a configurable multiplier ($M$, default $2.0$) of that ACTUAL loss. It does NOT assume that the entire previous investment is lost."*

#### 4.3.2 Mathematical State Engine

##### A. State Variables & Multipliers Vector Reference

| Variable | Mathematical Domain | Operational Meaning | Provenance |
| :---: | :---: | :--- | :---: |
| $P_{\text{current}}$ | $P_{\text{current}} > 0$ | Current market price at dip evaluation | `[OBSERVED]` |
| $T_{\text{invested}}$ | $T_{\text{invested}} > 0$ | Total prior cumulative capital deployed | `[CALCULATED]` |
| $Q_{\text{total}}$ | $Q_{\text{total}} > 0$ | Total delivery shares held prior to rebalancing | `[CALCULATED]` |
| $m$ | $m \ge 1.0$ | Actual-loss multiplier factor (default parameter $m = 2.0$) | `[USER-PROVIDED]` |
| $MV$ | $MV \ge 0$ | Instantaneous portfolio mark-to-market valuation | `[CALCULATED]` |
| $L_{\text{actual}}$ | $L_{\text{actual}} \ge 0$ | True monetary unrealized drawdown in currency | `[CALCULATED]` |
| $B_{\text{next}}$ | $B_{\text{next}} \ge 0$ | Dynamic capital tranche allocated to current dip | `[CALCULATED]` |
| $\Delta Q$ | $\Delta Q \ge 0$ | Additional shares acquired in this recovery tranche | `[CALCULATED]` |
| $T_{\text{new}}$ | $T_{\text{new}} = T_{\text{invested}} + B_{\text{next}}$ | Updated total cumulative capital deployed | `[CALCULATED]` |
| $Q_{\text{new}}$ | $Q_{\text{new}} = Q_{\text{total}} + \Delta Q$ | Updated total delivery share count | `[CALCULATED]` |
| $AC_{\text{new}}$ | $AC_{\text{new}} > 0$ | Updated volume-weighted breakeven cost basis | `[CALCULATED]` |

---

##### B. Governing Adaptive Mathematical State Engine Reference

| Metric / Parameter | Symbol | Exact Governing Equation | Unit | Provenance | Operational Role |
| :--- | :---: | :--- | :---: | :---: | :--- |
| **Mark-to-Market Value** | $MV$ | $$\mathbf{MV = Q_{\text{total}} \times P_{\text{current}}}$$ | $\$$ | `[CALCULATED]` | Computes real-time gross liquidation value |
| **Actual Monetary Loss** | $L_{\text{actual}}$ | $$\mathbf{L_{\text{actual}} = \max(0, T_{\text{invested}} - MV)}$$ | $\$$ | `[CALCULATED]` | Extracts true currency deficit without assuming 100% loss |
| **Adaptive Next Tranche** | $B_{\text{next}}$ | $$\mathbf{B_{\text{next}} = m \times L_{\text{actual}}}$$ | $\$$ | `[CALCULATED]` | Scales investment strictly to actual deficit ($m=2.0$) |
| **Incremental Shares** | $\Delta Q$ | $$\mathbf{\Delta Q = \left\lfloor \frac{B_{\text{next}}}{P_{\text{current}}} \right\rfloor \quad \text{or} \quad \frac{B_{\text{next}}}{P_{\text{current}}}}$$ | shares | `[CALCULATED]` | Floor-truncated integer share quantity for cash order |
| **Updated Capital Deployed** | $T_{\text{new}}$ | $$\mathbf{T_{\text{new}} = T_{\text{invested}} + B_{\text{next}}}$$ | $\$$ | `[CALCULATED]` | Increments total portfolio cash committed |
| **Updated Total Shares** | $Q_{\text{new}}$ | $$\mathbf{Q_{\text{new}} = Q_{\text{total}} + \Delta Q}$$ | shares | `[CALCULATED]` | Increments aggregate share inventory |
| **Adjusted Breakeven Cost** | $AC_{\text{new}}$ | $$\mathbf{AC_{\text{new}} = \frac{T_{\text{new}}}{Q_{\text{new}}} = \frac{T_{\text{invested}} + B_{\text{next}}}{Q_{\text{total}} + \Delta Q}}$$ | $\$/\text{share}$ | `[CALCULATED]` | Shifts breakeven threshold just above current price |

---

##### C. The 5-Phase Dynamic Calculus Walkthrough

At each adverse price dip trigger, the agent executes the 5-phase rebalancing pipeline:

1. **Phase 1: Mark-to-Market Valuation**:
   Determine the current market liquidation value of existing delivery shares:
   $$MV = Q_{\text{total}} \times P_{\text{current}}$$
2. **Phase 2: Unrealized Loss Extraction**:
   Isolate the net monetary deficit. If $MV \ge T_{\text{invested}}$, then $L_{\text{actual}} = 0$ (no averaging authorized):
   $$L_{\text{actual}} = \max(0, T_{\text{invested}} - MV)$$
3. **Phase 3: Adaptive Allocation Sizing**:
   Multiply the actual deficit by $m = 2.0$ to fund both the loss and the recovery margin:
   $$B_{\text{next}} = 2.0 \times L_{\text{actual}}$$
4. **Phase 4: Integer Share Order Generation**:
   Compute exact whole delivery shares for cash broker routing:
   $$\Delta Q = \left\lfloor \frac{B_{\text{next}}}{P_{\text{current}}} \right\rfloor$$
5. **Phase 5: Cost Basis Recalibration**:
   Verify that $AC_{\text{new}}$ compresses toward $P_{\text{current}}$, establishing a tight recovery ceiling:
   $$AC_{\text{new}} = \frac{T_{\text{invested}} + (\Delta Q \cdot P_{\text{current}})}{Q_{\text{total}} + \Delta Q} < AC_{\text{prior}}$$

#### 4.3.3 Comprehensive Walkthrough ($P_0 = ₹18.00, S = ₹100, L = 20\%, m = 2.0$)
| Round | Price | Tranche ($B_n$) | Total Invested ($T$) | Total Shares ($Q$) | Avg Cost ($AC$) | Market Value ($MV$) | Unrealized P/L | Next Allocation ($2 \times \text{Loss}$) |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| 1 | ₹18.00 | ₹100.00 | ₹100.00 | 5.555556 | ₹18.00 | ₹100.00 | ₹0.00 | — |
| 2 | ₹14.40 | ₹40.00 | ₹140.00 | 8.333333 | ₹16.80 | ₹120.00 | -₹20.00 | ₹40.00 |
| 3 | ₹11.52 | ₹88.00 | ₹228.00 | 15.972222 | ₹14.27 | ₹184.00 | -₹44.00 | ₹88.00 |
| 4 | ₹9.22 | ₹161.60 | ₹389.60 | 33.506944 | ₹11.63 | ₹308.80 | -₹80.80 | ₹161.60 |
| 5 | ₹7.37 | ₹285.12 | ₹674.72 | 72.179400 | ₹9.35 | ₹532.16 | -₹142.56 | ₹285.12 |
| 6 | ₹5.90 | ₹497.98 | ₹1,172.70 | 156.602263 | ₹7.49 | ₹923.71 | -₹248.99 | ₹497.98 |
| 7 | ₹4.72 | ₹867.47 | ₹2,040.17 | 340.443743 | ₹5.99 | ₹1,606.44 | -₹433.73 | ₹867.47 |
| 8 | ₹3.77 | ₹1,510.04 | ₹3,550.22 | 740.505704 | ₹4.79 | ₹2,795.20 | -₹755.02 | ₹1,510.04 |
| 9 | ₹3.02 | ₹2,628.12 | ₹6,178.34 | 1,610.849925 | ₹3.84 | ₹4,864.28 | -₹1,314.06 | ₹2,628.12 |
| 10 | ₹2.42 | ₹4,573.83 | ₹10,752.17 | 3,503.734898 | ₹3.07 | ₹8,465.25 | -₹2,286.92 | ₹4,573.83 |
| 11 | ₹1.93 | ₹7,959.94 | ₹18,712.11 | 7,622.430409 | ₹2.45 | ₹14,732.14 | -₹3,979.97 | ₹7,959.94 |

#### 4.3.4 Empirical Proof of 99.68% Capital Reduction
Under an identical market crash scenario (10 consecutive 20% drops, cumulative decline of $89.28\%$):
- **TRM V1 Total Capital Deployed**: **₹5,904,900.00**
- **TRM V2 Total Capital Deployed**: **₹18,712.11**
- **Capital Savings**:
  $$\text{Savings} = \frac{5,904,900.00 - 18,712.11}{5,904,900.00} = \frac{5,886,187.89}{5,904,900.00} = \mathbf{99.6829\%} \approx \mathbf{99.68\%} \quad [CALCULATED]$$

*Conclusion*: TRM V2 achieves virtually the same average-cost neutralization ($₹2.45$ vs $₹2.13$) while consuming less than $0.32\%$ of the capital required by TRM V1. It is fully viable for institutional cash equity accounts.

---

### 4.4 TRM V3: Asymmetric Dual-Path Model (Profit Retention + Loss Recovery)

#### 4.4.1 Architectural Design
TRM V3 expands the Teja Recovery Model from a pure downward averaging model into a complete, bidirectional **Position-Retention and Cycle-Reversal Framework** `[OBSERVED]`. It operates symmetrically across both profitable and losing exits:

```mermaid
flowchart TD
    TradeExitGate{"TRM V3 Trade Resolution Gate<br/>Price Action Reaches Defined Exit Bound"}

    subgraph PathA ["PATH A: PROFIT TARGET HIT (+p%) — RUNNER RETENTION"]
        direction TB
        A1["1. Liquidate Original Capital (I)<br/>Principal cash returned to liquid portfolio bankroll"]
        A2["2. Compute Profit Allocation & Runner Shares<br/>$R = \text{Net Profit} \times r$<br/>$Q_{\text{retain}} = \lfloor R / P_{\text{current}} \rfloor$"]
        A3[("3. Transfer Shares to Wealth Vault<br/>Zero-cost 'House Money' perpetual runner equity")]
        A4["4. Universal Reset Invariant<br/>Active trading tranche resets to Base Capital $S$<br/>Cycle profit counter: $N=0$"]
        A1 --> A2 --> A3 --> A4
    end

    subgraph PathB ["PATH B: STOP-LOSS EXIT (-l%) — LOSS RECOVERY ENGINE"]
        direction TB
        B1["1. Hard Liquidation to Cash<br/>100% position liquidated immediately at Stop-Loss price<br/>Capital preservation takes absolute priority"]
        B2["2. Compute Realized Loss in Currency<br/>$\text{Loss}_{\$} = I \times l$ (Exact Rupee/Dollar Loss)"]
        B3["3. Calculate Loss-Recovery Surcharge<br/>$R_{\text{loss}} = \text{Loss}_{\$} \times f$<br/>(Fraction $f$ calibrated to next trade target R:R)"]
        B4["4. Calibrated Next Setup Allocation<br/>$A_{\text{next}} = A_{\text{base}} + R_{\text{loss}}$<br/>(Subject to strict 15% maximum bankroll cap)"]
        B5["5. Reversal Trade Execution<br/>Deploy $A_{\text{next}}$ upon next qualified confluence entry<br/>On recovery win: Complete reset to Base $S$"]
        B1 --> B2 --> B3 --> B4 --> B5
    end

    subgraph PortfolioVault ["Portfolio Capital Management"]
        direction LR
        CashBankroll[("Liquid Cash Bankroll (W)")]
        RunnerCompounding[("Compounding Wealth Vault<br/>(Accumulated Free Shares)")]
    end

    TradeExitGate ==>|"Profit Target Achieved (+p%)"| PathA
    TradeExitGate ==>|"Structural Stop Triggered (-l%)"| PathB

    A1 -.->|"Liquid Cash Principal"| CashBankroll
    A3 ==>|"Deposit Retained Shares"| RunnerCompounding
    B1 -.->|"Preserve Remaining Cash"| CashBankroll
    A4 -.->|"Next Trade at Base S"| CashBankroll
    B5 -.->|"On Win: Return Surcharge to Cash"| CashBankroll

    classDef exitGate fill:#312E81,stroke:#C7D2FE,stroke-width:2px,color:#FFFFFF;
    classDef profitPath fill:#064E3B,stroke:#34D399,stroke-width:2px,color:#ECFDF5;
    classDef lossPath fill:#7F1D1D,stroke:#F87171,stroke-width:2px,color:#FEF2F2;
    classDef vault fill:#0F172A,stroke:#38BDF8,stroke-width:2px,color:#F8FAFC;

    class TradeExitGate exitGate;
    class A1,A2,A4 profitPath;
    class A3 vault;
    class B1,B2,B3,B4,B5 lossPath;
    class CashBankroll,RunnerCompounding vault;
```

#### 4.4.2 Mathematical Specification: Profit Path (Profit Retention)

##### A. Configurable Inputs & Valuation Vector Reference

| Variable | Domain | Financial / Mathematical Meaning | Provenance |
| :---: | :---: | :--- | :---: |
| $P_{\text{entry}}$ | $P_{\text{entry}} > 0$ | Original purchase execution price ($P_0$) | `[USER-PROVIDED]` |
| $Q_{\text{original}}$ | $Q_{\text{original}} \in \mathbb{N}_{\ge 1}$ | Initial delivery shares purchased ($Q_0$) | `[USER-PROVIDED]` |
| $I$ | $I = P_{\text{entry}} \times Q_{\text{original}}$ | Total principal cash capital deployed | `[CALCULATED]` |
| $p$ | $0 < p < 1.0$ | Target profit percentage (e.g., $0.15 = 15\%$) | `[USER-PROVIDED]` |
| $r$ | $0 < r < 1.0$ | Profit retention factor (e.g., $0.15 = 15\%$) | `[USER-PROVIDED]` |
| $P_{\text{current}}$ | $P_{\text{current}} > P_{\text{entry}}$ | Target exit price reached by market | `[OBSERVED]` |
| $P_{\text{amount}}$ | $P_{\text{amount}} > 0$ | Gross realized monetary profit | `[CALCULATED]` |
| $R$ | $R > 0$ | Monetary profit value dedicated to share retention | `[CALCULATED]` |
| $Q_{\text{retain}}$ | $Q_{\text{retain}} \ge 0$ | Number of permanent runner shares vaulted | `[CALCULATED]` |
| $Q_{\text{sell}}$ | $Q_{\text{sell}} \ge 0$ | Number of delivery shares liquidated to cash | `[CALCULATED]` |
| $\text{Cash Realized}$ | $\text{Cash Realized} > 0$ | Total liquid cash returned to trading bankroll | `[CALCULATED]` |

---

##### B. Governing Profit Path Mathematical Engine Reference

| Metric / Parameter | Symbol | Exact Governing Equation | Unit | Provenance | Operational Role |
| :--- | :---: | :--- | :---: | :---: | :--- |
| **Target Exit Price** | $P_{\text{current}}$ | $$\mathbf{P_{\text{current}} = P_{\text{entry}} \times (1 + p)}$$ | $\$/\text{share}$ | `[CALCULATED]` | Target price floor triggering profit harvest |
| **Gross Profit Amount** | $P_{\text{amount}}$ | $$\mathbf{P_{\text{amount}} = I \times p = (P_{\text{current}} - P_{\text{entry}}) \times Q_{\text{original}}}$$ | $\$$ | `[CALCULATED]` | Total monetary gains generated by trade |
| **Retained Profit Value** | $R$ | $$\mathbf{R = P_{\text{amount}} \times r = (I \times p) \times r}$$ | $\$$ | `[CALCULATED]` | Cash value converted into compounding runner shares |
| **Retained Share Quantity** | $Q_{\text{retain}}$ | $$\mathbf{Q_{\text{retain}} = \frac{R}{P_{\text{current}}} = \frac{I \times p \times r}{P_{\text{entry}} \times (1 + p)}}$$ | shares | `[CALCULATED]` | **The Current Price Law**: Evaluated at current market price |
| **Liquidated Shares** | $Q_{\text{sell}}$ | $$\mathbf{Q_{\text{sell}} = Q_{\text{original}} - Q_{\text{retain}}}$$ | shares | `[CALCULATED]` | Shares sold to recover principal + net cash profit |
| **Liquid Cash Proceeds** | $\text{Cash Realized}$ | $$\mathbf{Q_{\text{sell}} \times P_{\text{current}} = I + P_{\text{amount}} \times (1 - r)}$$ | $\$$ | `[CALCULATED]` | Cash credited immediately back to trading bankroll |
| **Perpetual Runner Basis** | $\text{Cost Basis}$ | $$\mathbf{\text{Effective Basis} = \$0.00}$$ | $\$/\text{share}$ | `[CALCULATED]` | 100% of initial principal ($I$) is extracted |

---

##### C. The Current Price Law & Zero-Cost Basis Proof

> [!IMPORTANT]
> **The Current Price Law of Share Retention**:
> Converting monetary profit $R$ into shares MUST divide by the **current market price ($P_{\text{current}}$)**:
> $$Q_{\text{retain}} = \frac{R}{P_{\text{current}}} \quad \text{(CORRECT)}$$
> Dividing by entry price $P_{\text{entry}}$ is a **fatal dimensional violation**:
> $$\frac{R}{P_{\text{entry}}} = \frac{R}{P_{\text{current}} / (1 + p)} = Q_{\text{retain}} \times (1 + p) > Q_{\text{retain}} \quad \text{(WRONG: Distorts Cash Ledger)}$$

**Mathematical Proof of Zero-Cost Runner Basis**:
1. Initial Capital Deployed: $I$
2. Total Cash Proceeds Returned to Bankroll:
   $$\text{Cash Realized} = Q_{\text{sell}} \times P_{\text{current}} = (Q_{\text{original}} - Q_{\text{retain}}) \times P_{\text{current}}$$
   $$= Q_{\text{original}} P_{\text{current}} - \left(\frac{R}{P_{\text{current}}}\right) P_{\text{current}} = [I + P_{\text{amount}}] - R = I + [P_{\text{amount}} - R]$$
3. Since $r \in (0, 1)$, $P_{\text{amount}} - R = P_{\text{amount}}(1 - r) > 0$.
4. Therefore, $\text{Cash Realized} > I$. The initial investment $I$ is **100% returned in cash**, plus a net cash profit of $P_{\text{amount}}(1 - r)$.
5. The remaining $Q_{\text{retain}}$ shares carry an **effective net acquisition cost of exactly $\$0.00$** ("House Money"). $\blacksquare$

---

##### D. Profit Path Ledger Accounting & Reconciliation ($P_{\text{entry}} = ₹10.00, Q = 100, p = 15\%, r = 15\%$)

| Accounting Ledger Item | Mathematical Expression | Computed Value | Financial Meaning |
| :--- | :--- | :---: | :--- |
| **Principal Capital Deployed ($I$)** | $100 \times ₹10.00$ | **₹1,000.00** | Initial cash drawn from liquid bankroll |
| **Target Exit Price ($P_{\text{current}}$)** | $₹10.00 \times (1 + 0.15)$ | **₹11.50** | Market quote at execution limit |
| **Gross Realized Profit ($P_{\text{amount}}$)** | $₹1,000.00 \times 15\%$ | **₹150.00** | Total monetary profit generated |
| **Retained Profit Value ($R$)** | $₹150.00 \times 15\%$ | **₹22.50** | Value converted to perpetual runner equity |
| **Retained Runner Shares ($Q_{\text{retain}}$)**| $\frac{₹22.50}{₹11.50}$ | **1.956522 shares** | Free shares transferred to Compounding Vault |
| **Shares Liquidated to Cash ($Q_{\text{sell}}$)** | $100 - 1.956522$ | **98.043478 shares** | Sold on spot exchange |
| **Initial Capital Returned** | $I$ | **₹1,000.00** | 100% principal preservation |
| **Net Cash Profit Realized** | $₹150.00 \times (1 - 0.15)$ | **₹127.50** | Liquid profit credited to cash |
| **Total Liquid Cash Returned** | $₹1,000.00 + ₹127.50$ | **₹1,127.50** | $98.043478 \times ₹11.50 = ₹1,127.50$ |
| **Perpetual Runner Equity Value** | $1.956522 \times ₹11.50$ | **₹22.50** | Zero-cost basis equity asset |
| **Total Cycle Reconciliation** | $₹1,127.50 \text{ (Cash)} + ₹22.50 \text{ (Vault)}$ | **₹1,150.00** | **100.00% Reconciled ($100 \times ₹11.50$)** |

---

#### 4.4.3 Mathematical Specification: Loss Path (Loss Recovery Reversal)

##### A. Configurable Inputs & Loss State Reference

| Variable | Domain | Financial / Mathematical Meaning | Provenance |
| :---: | :---: | :--- | :---: |
| $P_{\text{entry}}$ | $P_{\text{entry}} > 0$ | Original purchase execution price ($P_0$) | `[USER-PROVIDED]` |
| $Q_{\text{original}}$ | $Q_{\text{original}} \in \mathbb{N}_{\ge 1}$ | Initial delivery shares purchased ($Q_0$) | `[USER-PROVIDED]` |
| $I$ | $I = P_{\text{entry}} \times Q_{\text{original}}$ | Total principal cash capital deployed | `[CALCULATED]` |
| $l$ | $0 < l < 1.0$ | Hard structural stop-loss percentage (e.g., $0.15 = 15\%$) | `[USER-PROVIDED]` |
| $f$ | $0 < f \le 1.0$ | Loss-recovery surcharge factor (default $0.15 = 15\%$) | `[USER-PROVIDED]` |
| $A_{\text{base}}$ | $A_{\text{base}} > 0$ | Standard base capital allocation for next trade ($S$) | `[USER-PROVIDED]` |
| $P_{\text{current}}$ | $P_{\text{current}} < P_{\text{entry}}$ | Stop-loss execution price | `[OBSERVED]` |
| $L_{\text{amount}}$ | $L_{\text{amount}} > 0$ | Gross realized monetary loss in currency | `[CALCULATED]` |
| $R_{\text{loss}}$ | $R_{\text{loss}} > 0$ | Calibrated recovery surcharge added to next trade | `[CALCULATED]` |
| $A_{\text{next}}$ | $A_{\text{next}} > A_{\text{base}}$ | Total capital deployed into next qualified reversal setup | `[CALCULATED]` |

---

##### B. Governing Loss Path Mathematical Engine Reference

| Metric / Parameter | Symbol | Exact Governing Equation | Unit | Provenance | Operational Role |
| :--- | :---: | :--- | :---: | :---: | :--- |
| **Stop-Loss Exit Price** | $P_{\text{current}}$ | $$\mathbf{P_{\text{current}} = P_{\text{entry}} \times (1 - l)}$$ | $\$/\text{share}$ | `[CALCULATED]` | Structural exit floor triggering liquidation |
| **Gross Realized Loss** | $L_{\text{amount}}$ | $$\mathbf{L_{\text{amount}} = I \times l = (P_{\text{entry}} - P_{\text{current}}) \times Q_{\text{original}}}$$ | $\$$ | `[CALCULATED]` | Exact monetary loss realized on 100% exit |
| **Preserved Liquid Cash** | $\text{Cash Recovered}$ | $$\mathbf{\text{Cash} = I \times (1 - l) = Q_{\text{original}} \times P_{\text{current}}}$$ | $\$$ | `[CALCULATED]` | Cash preserved and credited to liquid bankroll |
| **Loss-Recovery Surcharge** | $R_{\text{loss}}$ | $$\mathbf{R_{\text{loss}} = L_{\text{amount}} \times f}$$ | $\$$ | `[CALCULATED]` | Incremental capital added to recover prior deficit |
| **Next Trade Tranche** | $A_{\text{next}}$ | $$\mathbf{A_{\text{next}} = A_{\text{base}} + R_{\text{loss}}}$$ | $\$$ | `[CALCULATED]` | Capital deployed in next trade (capped at 15% bankroll) |
| **Bankroll Cap Guardrail** | $\text{Cap Check}$ | $$\mathbf{A_{\text{next}} \le 0.15 \times W}$$ | $\$$ | `[CALCULATED]` | Strict maximum position risk constraint |

---

##### C. Mathematical Calibration of Recovery Factor ($f$)

> [!NOTE]
> **Linear Surcharge vs. Exponential Martingale**:
> Unlike TRM Original or Classic Martingale which double or triple total capital ($O(3^n)$ or $O(2^n)$), TRM V3 adds a calibrated **linear surcharge** $R_{\text{loss}} = L_{\text{amount}} \times f$.
> If the next qualified trade achieves profit target $p_{\text{target}}$ with risk-to-reward ratio $R:R \ge 1.5$, setting:
> $$f = \frac{1}{(R:R) \times p_{\text{target}}}$$
> ensures that the recovery trade recaptures the prior loss without creating exponential bankroll hazard.

---

##### D. Loss Path Ledger Accounting & Reconciliation ($P_{\text{entry}} = ₹10.00, Q = 100, l = 15\%, f = 15\%, A_{\text{base}} = ₹100.00$)

| Accounting Ledger Item | Mathematical Expression | Computed Value | Financial Meaning |
| :--- | :--- | :---: | :--- |
| **Principal Capital Deployed ($I$)** | $100 \times ₹10.00$ | **₹1,000.00** | Initial trade allocation |
| **Stop-Loss Exit Price ($P_{\text{current}}$)** | $₹10.00 \times (1 - 0.15)$ | **₹8.50** | Hard stop hit; immediate liquidation |
| **Preserved Cash Returned** | $100 \times ₹8.50$ | **₹850.00** | 85% of capital preserved in bankroll |
| **Realized Monetary Loss ($L_{\text{amount}}$)** | $₹1,000.00 - ₹850.00$ | **₹150.00** | Deficit booked to ledger |
| **Recovery Surcharge ($R_{\text{loss}}$)** | $₹150.00 \times 15\%$ | **₹22.50** | Linear loss surcharge |
| **Standard Base Tranche ($A_{\text{base}}$)** | Baseline $S$ | **₹100.00** | Normal setup allocation |
| **Next Reversal Tranche ($A_{\text{next}}$)** | $₹100.00 + ₹22.50$ | **₹122.50** | Deployed only upon next qualified confluence |
| **Bankroll Cap Audit ($15\% \times W$)** | For $W = ₹5,00,000$ (Cap = ₹75,000) | **PASS** | ₹122.50 is well below the ₹75,000 ceiling |
| **Strategic Recovery Cycle** | Target gain on ₹122.50 | **Recaptures Loss** | Cycle unconditionally resets to Base $S$ upon win |

---

### 4.5 TRM V4: 50% Current-Price P/L Share Mechanism

#### 4.5.1 Purpose
TRM V4 converts **actual monetary P/L** into current-price-equivalent shares and applies a 50% factor. It is intentionally simpler than V5/V6/V7.

#### 4.5.2 Profit Path
Let:
- $I=P_{entry}\times Q_0$
- $MV=P_{current}\times Q$
- $Profit=\max(0,MV-I)$

Then:
$$\boxed{Q_{profit-equivalent}=\frac{Profit}{P_{current}}}$$
$$\boxed{Q_{retain}=\frac{Profit}{2P_{current}}}$$
$$\boxed{Q_{sell}=Q-Q_{retain}}$$

For executable cash-equity orders, round only at the order layer and reconcile residual cash. Do not mix shares and currency in a single state variable.

#### 4.5.3 Loss Path
$$\boxed{Q_{recovery,raw}=\frac{ActualLoss}{2P_{current}}}$$
This is a **calculation**, not an automatic order. Recovery requires risk/capital/structure gates.

#### 4.5.4 Strengths / Limitations
- Strengths: simple, transparent, current-price correct, easy to automate.
- Limitations: fixed 50% factor is not regime-adaptive and can still over-liquidate a strong trend winner.

---

### 4.6 TRM V5: Adaptive Actual P/L Dual-Path Engine

#### 4.6.1 Core Invariants
1. Trigger percentages are gates; actual observed P/L is the calculation basis.
2. Monetary values are converted to shares only at the current execution price.
3. Loss recovery is risk-gated; no blind averaging.
4. No geometric escalation from Original/V1 is permitted in V5 live mode.

#### 4.6.2 Profit
$$I=P_{entry}\times Q$$
$$MV=P_{current}\times Q$$
$$PnL=MV-I$$
$$ActualProfit=\max(0,PnL)$$
$$Q_{retain}=\frac{ActualProfit}{2P_{current}}$$
$$Q_{sell}=\max(0,Q-Q_{retain})$$

V5 is the reference engine when a deterministic actual-P/L calculation is required but full V6/V8 orchestration is unnecessary.

#### 4.6.3 Loss
$$ActualLoss=\max(0,-PnL)$$
$$Q_{recovery,raw}=\frac{ActualLoss}{2P_{current}}$$

The raw quantity is advisory and must be constrained by available cash, position cap, risk budget, stop distance, liquidity and maximum recovery rounds.

---

### 4.7 TRM V6: Adaptive Multi-Agent Exit & Re-Entry Rotation Model

#### 4.7.1 Role
V6 is the **market-lifecycle intelligence engine**. It consumes validated market/technical/fundamental analysis and calculates exit zones, re-entry zones, confirmation states, and risk-gated position sizing.

#### 4.7.2 External Analysis Contract
External-agent fields may include:
- trend, market regime, support/resistance,
- regression mean / upper / lower bands,
- VWAP, EMA20, EMA50, SMA200,
- RSI, MACD, ATR, volume, RVOL,
- sector/index trend,
- fundamental status, debt status, promoter risk,
- corporate events, news risk, structural invalidation.

**Validation is mandatory.** Missing, stale, contradictory, or impossible values produce `WAIT`/`DEGRADED_MODE`; V6 must never fabricate levels.

#### 4.7.3 Exit-Zone Engine
Generate candidates from validated:
- resistance,
- regression upper band,
- user target,
- prior swing high,
- ATR extension.

Cluster candidates using:
$$Tolerance=\max(0.5\times ATR,0.01\times P_{current})$$

Output:
`exit_low`, `exit_mid`, `exit_high`, `primary_exit`, `exit_trigger`.

#### 4.7.4 Re-Entry-Zone Engine
Candidates:
- major support,
- regression mean/lower band,
- VWAP,
- 20 EMA / 50 SMA,
- breakout-retest,
- validated external-agent levels.

Output:
`aggressive_entry`, `primary_entry`, `confirmation_entry`, `invalidation`.

If candidate disagreement is excessive, output `NO_CLEAN_REENTRY_ZONE`.

#### 4.7.5 Re-Entry Confirmation
Require, as applicable:
- price enters the zone,
- structure is intact,
- bullish rejection/reclaim,
- acceptable volume confirmation,
- sector/index conditions acceptable,
- capital/risk limits pass.

#### 4.7.6 Position Sizing
$$A_{candidate}=\min(BaseAllocation,AvailableCash,RemainingCapitalCap)$$
$$Q_{cash}=\left\lfloor A_{candidate}/P_{entry}\right\rfloor$$
$$Q_{risk}=\left\lfloor RiskBudget/(P_{entry}-StopPrice)\right\rfloor$$
$$Q_{reentry}=\min(Q_{cash},Q_{risk},Q_{max})$$

Never invent `RiskBudget`, `StopPrice`, or `Q_max`.

---

### 4.8 TRM V7: Principal Preservation & Continuous Profit Harvesting

#### 4.8.1 Purpose
V7 addresses the specific failure mode where an algorithm sells too much of a strong winner and leaves insufficient exposure to a continuing uptrend.

Instead of retaining a configured fraction of profit, V7 preserves a **protected monetary principal** and harvests only the excess value above that principal while trend structure remains valid.

Let:
$$S=ProtectedPrincipal$$
$$MV=P_{current}\times Q$$

Then:
$$\boxed{HarvestAmount=\max(0,MV-S)}$$
$$\boxed{Q_{sell}=\frac{HarvestAmount}{P_{current}}}$$
$$\boxed{Q_{keep}=\frac{S}{P_{current}}}$$

After the harvest, residual position value should reconcile to approximately $S$, subject to whole-share execution and transaction costs.

#### 4.8.2 Harvest Trigger vs Calculation Basis
A configured move such as `+10%` determines **when** V7 evaluates harvesting. It does not replace the actual observed profit calculation.

Example:
- Protected principal = ₹40,000.
- Current position value = ₹48,000.
- Harvest = ₹8,000.
- Keep approximately ₹40,000 of stock exposure.

If the remaining position later rises to ₹44,000, harvest ₹4,000 again and return the stock value toward ₹40,000.

#### 4.8.3 Cycle Wealth Ledger
V7 MUST maintain both position value and harvested cash:
$$HarvestedCash_t=\sum_i Harvest_i$$
$$CycleWealth_t=CurrentPositionValue_t+HarvestedCash_t$$
$$CyclePnL_t=CycleWealth_t-I_0$$

A drop of the stock position below protected principal does **not** automatically mean the overall cycle is at a loss if previously harvested cash offsets it.

#### 4.8.4 Downside Routing
When trend quality deteriorates:
`STOP HARVESTING → CHECK STRUCTURE → CHECK CYCLE WEALTH → ROUTE TO V3/V4/V5/V6`.

Structural invalidation overrides V7 and requires exit according to the risk policy.

#### 4.8.5 Whole-Share Execution
If $Q_{keep}$ is fractional, the order layer must round to executable whole shares and reconcile the residual cash explicitly. Mathematical quantities may remain high precision internally.

---

### 4.9 TRM V8: Master Adaptive Meta-Controller

#### 4.9.1 Definition
V8 is the **final master orchestration layer** for TRM. It does not assume that one TRM equation is optimal in every market condition. V8 selects among the specialized engines using validated evidence, portfolio state, and risk constraints.

#### 4.9.2 Required State Vector
```text
Instrument:
  symbol, exchange, assetClass, currency, timestamp

Market:
  price, OHLCV, volume, RVOL, ATR, VWAP
  EMA20, EMA50, SMA200, RSI, MACD
  regressionMean, regressionUpper, regressionLower
  trendState, marketRegime, sectorTrend, indexTrend

Fundamental:
  marketCap, debtToEquity, liquidity, governanceFlags
  promoterRisk, corporateEvents, structuralRisk

Position:
  quantity, averageEntryPrice, investedCapital
  protectedPrincipal, currentMarketValue
  harvestedCash, cumulativeRealizedPnL
  cyclePnL, peakCycleWealth, currentDrawdown
  recoveryRound, maxRecoveryRounds

Risk:
  riskBudget, positionCap, availableCash
  wealthFloor, maxDrawdown, stopPrice
  transactionCostEstimate, slippageEstimate

Analysis:
  externalAgentPayload, validationStatus, confidenceState

Execution:
  orderType, limitPrice, executableQuantity
  action, rationale, blockers
```

#### 4.9.3 Wealth-Floor Engine
Define:
$$W_t=PositionValue_t+HarvestedCash_t$$
$$W^{peak}_t=\max(W^{peak}_{t-1},W_t)$$
$$DD_t=\frac{W^{peak}_t-W_t}{W^{peak}_t}$$

For configurable maximum drawdown $D_{max}$:
$$\boxed{WealthFloor_t=\max(I_0,W^{peak}_t(1-D_{max}))}$$

The wealth floor is a **risk control**, not a guarantee that execution can occur at that exact price.

#### 4.9.4 Regime Routing
```text
IF dataInvalid or contradictory:
    WAIT
ELSE IF structuralBreakdown:
    FULL_EXIT / NO_NEW_CAPITAL
ELSE IF wealthFloorBreached:
    DE-RISK / EXIT / V6
ELSE IF regime == BULLISH_STRONG:
    V7
ELSE IF regime == BULLISH_EXTENDED:
    V7 + V5/V6
ELSE IF regime == NEUTRAL:
    V4/V5/V6
ELSE IF regime == BEARISH_CORRECTION:
    V2-V8 only when recovery gates pass
ELSE IF regime == BEARISH:
    V6 / EXIT
ELSE:
    WAIT
```

#### 4.9.5 Adaptive Profit Engine
V8 may choose among:
- **V7** when a validated trend is strong and continued exposure is desirable.
- **V5** when actual P/L management is preferred without full lifecycle complexity.
- **V4** for simple deterministic 50% P/L-share handling.
- **V6** when resistance, exhaustion, event risk, or re-entry structure dominates.

#### 4.9.6 Adaptive Loss Engine
Use the least aggressive valid recovery model:
- V3 for controlled linear recovery surcharge.
- V4/V5 for calculated loss-equivalent share sizing under strict gates.
- V2 for qualified correction accumulation in high-quality assets.
- V6 for structural exit/re-entry orchestration.
- Original/V1 remain research/reference modes unless explicitly enabled for simulation.

#### 4.9.7 Recovery Budget Cap
The V8 recovery allocation MUST satisfy:
$$\boxed{A_{recovery}\le\min(RiskBudget,AvailableCash,PositionCap,RecoveryBudget)}$$

No recovery rule may create an unbounded capital sequence.

#### 4.9.8 Cost-Aware Harvest Gate
Do not harvest when the expected benefit is not materially greater than estimated execution friction:
$$\boxed{HarvestAmount>MinimumEconomicActionValue}$$
where `MinimumEconomicActionValue` includes configurable brokerage/fees, taxes where relevant, slippage buffer, and a minimum absolute value threshold.

#### 4.9.9 Anti-Churn Hysteresis
V8 must use separate activation/deactivation thresholds and optional cooldown periods so that minor oscillations around one level do not repeatedly trigger `BUY/SELL/BUY/SELL`.

#### 4.9.10 First-Class Outcomes
V8 can emit only one primary action per evaluation cycle:
`NO_TRADE`, `WAIT`, `HOLD`, `HARVEST`, `REDUCE`, `ENTER`, `REENTER`, `RECOVER`, `EXIT`, `EMERGENCY_EXIT`.

#### 4.9.11 Priority Order
```text
1. Data integrity
2. Structural validity
3. Wealth-floor / portfolio risk
4. Liquidity / execution feasibility
5. Capital preservation
6. Market-regime fit
7. Profit harvesting
8. Recovery allocation
9. Additional exposure
```

#### 4.9.12 V8 Pseudocode
```text
validateInputs()
validateUnits()
validateExternalAnalysis()

if dataInvalid:
    return WAIT

calculateActualPnL()
calculateCycleWealth()
calculatePeakWealth()
calculateDrawdown()
calculateWealthFloor()
classifyMarketRegime()

if structuralBreakdown:
    return EMERGENCY_EXIT

if wealthFloorBreached:
    return DE_RISK_OR_EXIT

switch regime:
    BULLISH_STRONG:
        evaluateV7Harvest()
    BULLISH_EXTENDED:
        evaluateV7ThenV6Resistance()
    NEUTRAL:
        evaluateV4OrV5OrV6()
    BEARISH_CORRECTION:
        evaluateV2OrV3RecoveryGate()
    BEARISH:
        evaluateV6Exit()
    UNKNOWN:
        return WAIT

applyRiskCaps()
applyLiquidityGate()
applyEconomicCostGate()
applyAntiChurnGate()
applyWholeShareRounding()
reconcileLedger()
return one_primary_action()
```

---

### 4.10 Universal Invariants & Model Routing Matrix

#### 4.10.1 Non-Negotiable Invariants
| Invariant | Rule |
|---|---|
| Actual-P/L Authority | Actual observed P/L is authoritative once the relevant trigger is reached |
| Current-Price Law | Monetary profit/loss converted to shares MUST divide by current execution price |
| Dimensional Separation | Currency, percentage, price/share, and quantity remain separate units |
| No Blind Averaging | A price decline alone never authorizes additional capital |
| No Unlimited Recovery | Recovery rounds and dollar exposure are hard capped |
| Structural Override | Thesis invalidation overrides all recovery/harvest formulas |
| Cost Awareness | Small uneconomic trades are blocked |
| Anti-Churn | Hysteresis/cooldown blocks rapid flip-flopping |
| Principal Protection | V7 keeps protected principal-sized exposure while valid |
| Wealth Protection | V8 maintains a configurable cycle wealth floor |
| External-Agent Skepticism | AI analysis is evidence, not authority |
| No Fabrication | Missing data produces WAIT/DEGRADED_MODE, never invented numbers |
| Cycle Isolation | Completed cycles must not inherit prior recovery escalation |
| Whole-Share Execution | Fractional internal math is rounded only at order execution |

#### 4.10.2 Routing Matrix
| State | Preferred engine | Secondary engine | Automatic capital increase? |
|---|---|---|---|
| Strong bullish + valid | V7 | V5/V6 | No, unless sizing gate passes |
| Bullish extended | V7 | V6 | No |
| Neutral profitable | V5/V4 | V6 | No |
| Pullback, thesis valid | V5/V6 | V2-V8 | Only if recovery gates pass |
| Bearish correction | V2-V8 | V6 | Strictly capped |
| Structural breakdown | V6 | None | **Never** |
| Post-exit re-entry | V6 | V5 | Only after confirmation |
| Invalid/stale data | None | None | **Never** |

#### 4.10.3 Model-selection rationale
V8 is successful as an architecture only when it **refuses to use an unsuitable engine**. The controller must prefer `WAIT` to an aggressive formula when required evidence is missing or conflicting.

### 4.11 The Universal Reset & Cycle Accounting Invariant
For **completed recovery/exit cycles**, the allocation escalation state MUST reset to base:
$$\boxed{\text{Completed Recovery/Exit} \implies \text{Reset Escalation State to Base } S}$$

- Prior tranche sizes, recovery multipliers, and escalation counters are **NEVER** carried into a new independent cycle.
- V7/V8 profit harvesting can remain within the same active cycle because harvesting is not a completed recovery/exit event.
- On a true cycle close, reset:
  $$\text{RoundNumber} \leftarrow 1$$
  $$\text{CumulativeInvested} \leftarrow 0$$
  $$\text{RecoveryState} \leftarrow 0$$
  $$\text{NextAllocation} \leftarrow S$$
- Harvested cash and long-term vault assets must remain in the ledger according to the configured accounting policy; they must not be silently deleted during reset.

```mermaid
stateDiagram-v2
    [*] --> Idle_State: Dedicated Bankroll Established (W)

    Idle_State --> Base_Active_T1: Signal Confirmed (Deploy Tranche 1 = S)
    
    state Base_Active_T1 {
        [*] --> Position_Live
        Position_Live --> Target_Hit: Price >= Target (T1 / T2)
        Position_Live --> Adverse_Drop: Price <= Dip Trigger (-10%)
        Position_Live --> Invalidation_Breach: Price <= Structural Invalidation
    }

    state Dip_Recovery_Loop {
        [*] --> Tranche_Calculated: Compute B_next = 2 × ActualLoss
        Tranche_Calculated --> Capital_Deployed: Buy Additional Shares
        Capital_Deployed --> Price_Monitoring: Cost Basis Shifted Down
        Price_Monitoring --> Recovery_Exit: Price >= Break-Even + Margin
        Price_Monitoring --> Further_Drop: Price drops next 10% (n < 5)
        Further_Drop --> Tranche_Calculated: Step n = n + 1
    }

    Base_Active_T1 --> Target_Hit: Standard Win
    Base_Active_T1 --> Dip_Recovery_Loop: Adverse Pullback (V2 Mode)
    Base_Active_T1 --> Invalidation_Breach: Stop Loss Triggered (V3 Mode)

    state Profit_Harvest_V3 {
        [*] --> Retained_Value: R = Profit × r
        Retained_Value --> Convert_Shares: Q_retain = R / P_current
        Convert_Shares --> Liquidate_Balance: Cash Credited to Bankroll
    }

    Target_Hit --> Profit_Harvest_V3: V3 Dual-Path Mode
    Profit_Harvest_V3 --> Runner_Vault: Transfer Free Runner Shares
    
    Recovery_Exit --> Universal_Reset: 100% Cash Liquidated at Profit
    Runner_Vault --> Universal_Reset: Retained Equity Locked in Vault
    
    Invalidation_Breach --> Loss_Recovery_V3: Realized Loss L = I × l
    Loss_Recovery_V3 --> Universal_Reset: Next Tranche A_next = A_base + (L × f)
    
    state Universal_Reset {
        [*] --> Purge_Cycle_State: Clear Tranche Counter
        Purge_Cycle_State --> Reset_Variables: RoundNumber = 1, Allocation = S
        Reset_Variables --> [*]
    }

    Universal_Reset --> Idle_State: Ready for Next Independent Cycle
```

> [!IMPORTANT]
> **Anti-Compounding Principle**:
> Never carry the allocation size of a completed recovery cycle into the next trade. If Round 4 deployed ₹45,000 to recover a dip, the very next trade must deploy **₹5,000 (Base $S$)**, NOT ₹45,000. Compounding tranche sizes post-recovery is the single most common cause of algorithmic capital destruction.

---

### 4.12 Historical Cross-Model Comparison Matrix (Original–V3)
> This historical matrix is preserved for auditability. The complete Original–V8 comparison is defined in Section 4.13.
| Architectural Dimension | TRM Original | TRM V1 (Drop Model) | TRM V2 (Actual Loss) | TRM V3 (Dual-Path) |
| :--- | :--- | :--- | :--- | :--- |
| **Primary Domain** | Theoretical Math | Equity Dip Averaging | Equity Dip Averaging | Reversal / Swing Trading |
| **Trigger Mechanism** | Prior Losing Step | Predefined Price Drop ($L$) | Unrealized Portfolio Loss | Target Gain ($p$) or Stop ($l$) |
| **Capital Progression** | Geometric ($O(3^n)$) | Geometric ($O(3^n)$) | Sub-Geometric Adaptive | Linear Surcharge ($A_{\text{base}} + \Delta$) |
| **Mark-to-Market Accounting**| Assumes 100% loss | Ignores Residual Equity | Real-Time Mark-to-Market | Real-Time Mark-to-Market |
| **Profit Handling** | 100% Exit & Reset | 100% Exit & Reset | 100% Exit & Reset | Partial Retention ($Q_{\text{retain}}$) + Reset |
| **Capital Required (10 Rounds)**| \$295,245 ($S=5$) | ₹5,904,900 ($S=100$) | **₹18,712** ($S=100$) | Controlled by $A_{\text{base}}$ |
| **Capital Efficiency** | Zero (Exponential) | Extremely Fragile | **99.68% Savings vs V1** | **Maximum Capital Efficiency** |
| **Production Suitability**| Educational / Math Only | Not Recommended | **Approved (Max 3-5 Rounds)** | **Approved (Standard Production)**|

---


### 4.13 Complete Original–V8 Architecture Matrix

| Dimension | Original | V1 | V2 | V3 | V4 | V5 | V6 | V7 | **V8** |
|---|---|---|---|---|---|---|---|---|---|
| Primary role | Recursive recovery | Price-drop allocation | Actual-loss allocation | Dual-path | 50% P/L-share | Adaptive actual P/L | Exit/re-entry lifecycle | Principal-preserving harvest | **Adaptive meta-controller** |
| Actual P/L authority | No | No | Yes | Yes | Yes | Yes | Yes | Yes | **Yes** |
| Current-price share law | No | No | Partial | Yes | **Core** | **Core** | **Core** | **Core** | **Mandatory** |
| Profit handling | Full exit | Full exit | Full exit | Retention factor | 50% rule | 50% actual-P/L rule | Dynamic | Excess-only | **Route by regime** |
| Loss handling | Geometric | Geometric | Actual loss | Surcharge | Raw loss shares | Actual loss shares | Risk-gated | Delegated | **Adaptive + capped** |
| Trend/regime awareness | No | No | No | Limited | Limited | Limited | High | High | **High** |
| Exit/re-entry | No | No | No | Partial | No | Partial | **Core** | Partial | **Core** |
| Principal preservation | No | No | No | Partial | Partial | Partial | Partial | **Core** | **Core** |
| Wealth floor | No | No | No | No | No | No | No | No | **Core** |
| Cost-aware execution | No | No | No | Limited | Limited | Limited | Yes | Needed | **Mandatory** |
| Anti-churn | No | No | No | No | No | No | Yes | Needed | **Mandatory** |
| Model routing | No | No | No | No | No | No | Partial | No | **Core** |
| No-trade / wait | No | No | No | No | No | Partial | Yes | Yes | **Core** |
| Capital escalation | Unbounded geometric | Geometric | Loss-adaptive | Capped surcharge | None | Capped | Risk-gated | None | **Never unbounded** |
| Best use | Math research | Historical trigger research | Controlled correction research | Swing dual-path | Simple share law | Deterministic P/L engine | Lifecycle decisions | Strong-winner harvesting | **Master controller** |

**Architecture conclusion:** V8 is the control plane; V1–V7 are specialized engines. No V1–V7 formula is considered universally optimal. V8 must route, constrain, or block the engine based on validated state.


### 4.11 The Universal Reset Invariant
Across every edition of the TRM family, the governing state invariant is:
$$\boxed{\text{Upon Full Recovery Target Exit or Reversal Completion} \implies \text{UNCONDITIONAL RESET TO BASE } S}$$

- Prior tranche sizes, loss roll-overs, and escalation counters are **NEVER** carried forward into a new cycle.
- Each TRM cycle is an independent finite-state machine. Once the exit event triggers, state variables reset:
  $$\text{RoundNumber} \leftarrow 1$$
  $$\text{CumulativeInvested} \leftarrow 0$$
  $$\text{NextAllocation} \leftarrow S$$

```mermaid
stateDiagram-v2
    [*] --> Idle_State: Dedicated Bankroll Established (W)

    Idle_State --> Base_Active_T1: Signal Confirmed (Deploy Tranche 1 = S)
    
    state Base_Active_T1 {
        [*] --> Position_Live
        Position_Live --> Target_Hit: Price >= Target (T1 / T2)
        Position_Live --> Adverse_Drop: Price <= Dip Trigger (-10%)
        Position_Live --> Invalidation_Breach: Price <= Structural Invalidation
    }

    state Dip_Recovery_Loop {
        [*] --> Tranche_Calculated: Compute B_next = 2 × ActualLoss
        Tranche_Calculated --> Capital_Deployed: Buy Additional Shares
        Capital_Deployed --> Price_Monitoring: Cost Basis Shifted Down
        Price_Monitoring --> Recovery_Exit: Price >= Break-Even + Margin
        Price_Monitoring --> Further_Drop: Price drops next 10% (n < 5)
        Further_Drop --> Tranche_Calculated: Step n = n + 1
    }

    Base_Active_T1 --> Target_Hit: Standard Win
    Base_Active_T1 --> Dip_Recovery_Loop: Adverse Pullback (V2 Mode)
    Base_Active_T1 --> Invalidation_Breach: Stop Loss Triggered (V3 Mode)

    state Profit_Harvest_V3 {
        [*] --> Retained_Value: R = Profit × r
        Retained_Value --> Convert_Shares: Q_retain = R / P_current
        Convert_Shares --> Liquidate_Balance: Cash Credited to Bankroll
    }

    Target_Hit --> Profit_Harvest_V3: V3 Dual-Path Mode
    Profit_Harvest_V3 --> Runner_Vault: Transfer Free Runner Shares
    
    Recovery_Exit --> Universal_Reset: 100% Cash Liquidated at Profit
    Runner_Vault --> Universal_Reset: Retained Equity Locked in Vault
    
    Invalidation_Breach --> Loss_Recovery_V3: Realized Loss L = I × l
    Loss_Recovery_V3 --> Universal_Reset: Next Tranche A_next = A_base + (L × f)
    
    state Universal_Reset {
        [*] --> Purge_Cycle_State: Clear Tranche Counter
        Purge_Cycle_State --> Reset_Variables: RoundNumber = 1, Allocation = S
        Reset_Variables --> [*]
    }

    Universal_Reset --> Idle_State: Ready for Next Independent Cycle
```

> [!IMPORTANT]
> **Anti-Compounding Principle**:
> Never carry the allocation size of a completed recovery cycle into the next trade. If Round 4 deployed ₹45,000 to recover a dip, the very next trade must deploy **₹5,000 (Base $S$)**, NOT ₹45,000. Compounding tranche sizes post-recovery is the single most common cause of algorithmic capital destruction.

---

### 4.12 Cross-Model Comparison Matrix
| Architectural Dimension | TRM Original | TRM V1 (Drop Model) | TRM V2 (Actual Loss) | TRM V3 (Dual-Path) |
| :--- | :--- | :--- | :--- | :--- |
| **Primary Domain** | Theoretical Math | Equity Dip Averaging | Equity Dip Averaging | Reversal / Swing Trading |
| **Trigger Mechanism** | Prior Losing Step | Predefined Price Drop ($L$) | Unrealized Portfolio Loss | Target Gain ($p$) or Stop ($l$) |
| **Capital Progression** | Geometric ($O(3^n)$) | Geometric ($O(3^n)$) | Sub-Geometric Adaptive | Linear Surcharge ($A_{\text{base}} + \Delta$) |
| **Mark-to-Market Accounting**| Assumes 100% loss | Ignores Residual Equity | Real-Time Mark-to-Market | Real-Time Mark-to-Market |
| **Profit Handling** | 100% Exit & Reset | 100% Exit & Reset | 100% Exit & Reset | Partial Retention ($Q_{\text{retain}}$) + Reset |
| **Capital Required (10 Rounds)**| \$295,245 ($S=5$) | ₹5,904,900 ($S=100$) | **₹18,712** ($S=100$) | Controlled by $A_{\text{base}}$ |
| **Capital Efficiency** | Zero (Exponential) | Extremely Fragile | **99.68% Savings vs V1** | **Maximum Capital Efficiency** |
| **Production Suitability**| Educational / Math Only | Not Recommended | **Approved (Max 3-5 Rounds)** | **Approved (Standard Production)**|

---

# 5. Technical Market Regime & Confluence Engine (Trend-Regression-Momentum)

While Subsystem B dictates capital sizing, Subsystem A evaluates the price chart to determine trade timing and market regime qualification.

```mermaid
flowchart LR
    subgraph Regime ["1. Multi-Timeframe Trend Regime"]
        direction TB
        SMA200["Price > Rising 200 SMA (Weekly Macro)"]
        SuperTrend["Weekly SuperTrend (10, 3) Bullish"]
        EMAAlign["20 EMA > 50 SMA (Daily Alignment)"]
        SMA200 --- SuperTrend --- EMAAlign
    end

    subgraph Channels ["2. 100-Period Linear Regression Channel"]
        direction TB
        UpperBand["Upper Band (+2.0σ): Overbought (Strictly NO ENTRY)"]
        MeanBand["Regression Mean (0.0σ): Equilibrium (T1 Harvest)"]
        LowerBand["Lower Band (-2.0σ): Extreme Value (BUY DEMAND ZONE)"]
        UpperBand --- MeanBand --- LowerBand
    end

    subgraph Momentum ["3. Momentum & Volume Confluences"]
        direction TB
        RSIFilter["Daily RSI (14) > 50 (or Bouncing from < 35)"]
        MACDFilter["MACD Histogram Ticking Upward"]
        RVOLFilter["RVOL ≥ 1.2x at Bullish Rejection Candle"]
        RSIFilter --- MACDFilter --- RVOLFilter
    end

    Regime ==> Channels ==> Momentum

    classDef regimeBox fill:#0F172A,stroke:#38BDF8,stroke-width:2px,color:#F8FAFC;
    classDef channelBox fill:#1E1B4B,stroke:#818CF8,stroke-width:2px,color:#FFFFFF;
    classDef momentumBox fill:#064E3B,stroke:#34D399,stroke-width:2px,color:#ECFDF5;

    class SMA200,SuperTrend,EMAAlign regimeBox;
    class UpperBand,MeanBand,LowerBand channelBox;
    class RSIFilter,MACDFilter,RVOLFilter momentumBox;
```

### 5.1 Multi-Timeframe Trend Alignment (Macro Bias)
1. **Weekly Horizon (Macro Direction)**:
   - Asset must trade above its rising 200-day Simple Moving Average (SMA).
   - Weekly SuperTrend ($10, 3$) must be **Bullish (Green)**.
2. **Daily Horizon (Intermediate Trend)**:
   - 20-day Exponential Moving Average (EMA) must be aligned above the 50-day SMA.
   - Price structure must exhibit Higher Highs (HH) and Higher Lows (HL).

### 5.2 Linear Regression Channel Mapping (Structural Value Bounds)

A 100-period Linear Regression Channel ($N = 100$) is fitted to daily closing prices using Ordinary Least Squares (OLS) to define the equilibrium path and $\pm 2.0\sigma$ boundary envelopes.

##### A. Governing Linear Regression Mathematical Engine Reference

| Metric / Parameter | Symbol | Exact Governing Equation | Provenance | Operational Purpose |
| :--- | :---: | :--- | :---: | :--- |
| **Channel Lookback Length** | $N$ | $$N = 100 \text{ daily bars}$$ | `[ASSUMED]` | Standardization window |
| **OLS Slope Coefficient** | $\beta$ | $$\mathbf{\beta = \frac{N \sum_{t=1}^N (t \cdot P_t) - \left(\sum_{t=1}^N t\right)\left(\sum_{t=1}^N P_t\right)}{N \sum_{t=1}^N t^2 - \left(\sum_{t=1}^N t\right)^2}}$$ | `[CALCULATED]` | Directional drift of intermediate trend ($\beta > 0 \implies \text{Uptrend}$) |
| **OLS Intercept Constant** | $\alpha$ | $$\mathbf{\alpha = \frac{\sum_{t=1}^N P_t - \beta \sum_{t=1}^N t}{N}}$$ | `[CALCULATED]` | Baseline price origin at start of regression |
| **Regression Equilibrium Mean**| $\hat{P}_t$ | $$\mathbf{\hat{P}_t = \alpha + \beta t}$$ | `[CALCULATED]` | Fair-value centerline; Target 1 (T1) harvest line |
| **Standard Error of Estimate** | $S_e$ | $$\mathbf{S_e = \sqrt{\frac{\sum_{t=1}^N (P_t - \hat{P}_t)^2}{N - 2}}}$$ | `[CALCULATED]` | Volatility dispersion of residuals around mean |
| **Upper Channel Band (+2.0$\sigma$)**| $UB_t$ | $$\mathbf{UB_t = \hat{P}_t + 2.0 \times S_e}$$ | `[CALCULATED]` | Overbought resistance ceiling; Target 2 (T2) harvest line |
| **Lower Channel Band (-2.0$\sigma$)**| $LB_t$ | $$\mathbf{LB_t = \hat{P}_t - 2.0 \times S_e}$$ | `[CALCULATED]` | Extreme value demand floor; Preferred Tranche 1 / Dip entry |

---

##### B. Channel Boundary Operational Decision Matrix

| Channel Boundary Region | Price Condition | Market Regime | Tranche Authorization | Action Mandate |
| :--- | :--- | :--- | :---: | :--- |
| **Upper Boundary (+2.0$\sigma$)** | $P_{\text{current}} \ge UB_t$ | Overbought Exhaustion | ❌ **PROHIBITED** | Strictly **NO NEW ENTRIES**. Harvest T2 profits; trail stops. |
| **Equilibrium Mean (0.0$\sigma$)** | $\|P_{\text{current}} - \hat{P}_t\| \le 0.2 S_e$ | Fair Value Equilibrium | ⚠️ **CONDITIONAL** | Harvest T1 partial profit; shift stop to breakeven. |
| **Lower Boundary (-2.0$\sigma$)** | $P_{\text{current}} \le LB_t$ | Extreme Value / Demand | ✅ **AUTHORIZED** | High-probability entry zone for Tranche 1 ($B_1$) or V2 Dip tranches ($B_n$). |
| **Structural Breakdown** | $P_{\text{current}} < LB_t - 1.0 S_e$ | Channel Invalidation | 🛑 **STOP TRIGGER** | Technical thesis voided; execute structural invalidation stop. |

---

### 5.3 Momentum, Volatility & Mathematical Expectancy Engine

##### A. Technical Indicator Formulas Reference

| Indicator | Parameter | Governing Formula | Unit | Provenance | Operational Role |
| :--- | :---: | :--- | :---: | :---: | :--- |
| **True Range** | $\text{TR}_t$ | $$\mathbf{\text{TR}_t = \max(H_t - L_t, \|H_t - C_{t-1}\|, \|L_t - C_{t-1}\|)}$$ | $\$$ | `[CALCULATED]` | Raw inter-bar volatility |
| **Average True Range** | $\text{ATR}_{14}$ | $$\mathbf{\text{ATR}_t = \frac{13 \times \text{ATR}_{t-1} + \text{TR}_t}{14}}$$ | $\$$ | `[CALCULATED]` | Normalized daily volatility baseline |
| **Structural Stop Buffer** | $\text{Buffer}$ | $$\mathbf{\text{Buffer} = 0.5 \times \text{ATR}(14)}$$ | $\$$ | `[CALCULATED]` | Distance buffer below support swing low |
| **Relative Strength Index** | $\text{RSI}_{14}$ | $$\mathbf{\text{RSI} = 100 - \frac{100}{1 + \frac{\text{EMA}_{14}(\text{Up})}{\text{EMA}_{14}(\text{Down})}}}$$ | Points | `[CALCULATED]` | Momentum filter: Require $> 50$ (or bounce from $< 35$) |
| **Relative Volume** | $\text{RVOL}$ | $$\mathbf{\text{RVOL}_t = \frac{\text{Volume}_t}{\frac{1}{20}\sum_{i=1}^{20} \text{Volume}_{t-i}}}$$ | Ratio | `[CALCULATED]` | Volume confirmation: Require $\text{RVOL} \ge 1.2\times$ on bounce |

---

##### B. Mathematical Expectancy ($R:R$) Engine

Every proposed trade setup must satisfy the **Positive Mathematical Expectancy Invariant**:

$$\mathbf{\mathbb{E}[R] = (W_r \times R_w) - (L_r \times R_l) \ge +0.50 R \quad [CALCULATED]}$$

| Expectancy Parameter | Symbol | Definition / Benchmark Value | Provenance | Operational Requirement |
| :--- | :---: | :--- | :---: | :--- |
| **Historical Win Rate** | $W_r$ | Empirical setup success frequency ($0 \le W_r \le 1.0$) | `[ASSUMED]` | Baseline assumption: $W_r = 0.55$ ($55\%$) |
| **Payoff Ratio (Reward)** | $R_w$ | Target Reward in multiples of $R$ ($R_w = \frac{P_{\text{target}} - P_{\text{entry}}}{P_{\text{entry}} - P_{\text{stop}}}$) | `[CALCULATED]` | Minimum constraint: $R_w \ge 1.50 R$ |
| **Loss Rate** | $L_r$ | Probability of stop execution ($L_r = 1.0 - W_r$) | `[CALCULATED]` | Baseline: $L_r = 0.45$ ($45\%$) |
| **Risk Multiple (Loss)** | $R_l$ | Defined capital unit risked ($R_l \equiv 1.0 R$) | `[CALCULATED]` | Standard unit of account |
| **Net Setup Expectancy** | $\mathbb{E}[R]$ | Expected return per unit of risk deployed | `[CALCULATED]` | Threshold: $\mathbb{E}[R] \ge +0.50 R$ required for authorization |

---

# 6. TRM V8 Deterministic Execution Pipeline

The V8 pipeline is deterministic **given validated inputs**. Market outcomes themselves are not deterministic.

### Phase 0: Request & Scope Normalization
- Parse symbol/exchange, position, account, intended holding horizon, and user-specified parameters.
- Confirm cash-equity-only scope.
- Reject requests requiring fabricated live data or unsupported instruments.

### Phase 1: Data Integrity & Provenance
- Validate timestamp freshness.
- Validate price, OHLCV, volume, ATR, VWAP, moving averages, regression ordering, and corporate-action adjustments.
- Tag data as `[OBSERVED]`, `[CALCULATED]`, `[USER-PROVIDED]`, or `[ASSUMED]`.
- If key data is missing: `WAIT` or `DEGRADED_MODE`; never invent values.

### Phase 2: Eligibility & Fundamental Quality Gate
- Validate liquidity, solvency, governance and event-risk requirements.
- Ban averaging/recovery when structural or fundamental impairment is present.
- Apply configured market-cap, debt, pledging, liquidity and bankruptcy-immunity criteria.

### Phase 3: Portfolio State Calculation
Calculate:
$$I_0=P_{entry}Q_0$$
$$MV=P_{current}Q$$
$$PnL_{position}=MV-I_0$$
$$W=MV+HarvestedCash$$
$$PnL_{cycle}=W-I_0$$
$$W^{peak}=\max(W^{peak}_{previous},W)$$
$$DD=(W^{peak}-W)/W^{peak}$$

### Phase 4: Market Regime & Structural State
Classify:
`BULLISH_STRONG`, `BULLISH`, `BULLISH_EXTENDED`, `NEUTRAL`, `BEARISH_CORRECTION`, `BEARISH`, `STRUCTURAL_BREAKDOWN`, `UNKNOWN`.

Use trend, regression, momentum, volatility, volume, market/sector alignment, and validated external analysis.

### Phase 5: Master Model Router
Select the least aggressive suitable engine according to Section 4.9.4.

### Phase 6: Mathematical Engine
Run exactly one primary profit/recovery calculation path per action cycle. Store all intermediate state variables with units.

### Phase 7: Wealth / Risk / Capital Gate
Require:
- wealth floor compliance,
- risk budget compliance,
- single-stock cap compliance,
- available cash,
- maximum recovery rounds,
- liquidity feasibility,
- structural stop existence,
- economic minimum action value.

### Phase 8: Execution Construction
For cash equities:
$$Shares=\left\lfloor Allocation/LimitPrice\right\rfloor$$
$$CommittedCash=Shares\times LimitPrice$$
$$ResidualCash=Allocation-CommittedCash$$

For V7 harvest:
round executable sell quantity and reconcile residual position value against protected principal.

### Phase 9: Order & Anti-Churn Gate
- Prefer limit orders where liquidity/spread requires it.
- Apply cooldown/hysteresis.
- Reject duplicate orders for the same state event.
- Prevent simultaneous contradictory instructions.

### Phase 10: Post-Trade Reconciliation
Update:
- holdings,
- cash,
- harvested cash,
- realized P/L,
- protected principal,
- cycle wealth,
- peak wealth,
- drawdown,
- recovery round,
- cycle state.

### Phase 11: Monitoring & State Transition
Possible transitions:
```text
IDLE → ENTRY_CANDIDATE → POSITION_ACTIVE
POSITION_ACTIVE → V7_HARVEST
POSITION_ACTIVE → HOLD
POSITION_ACTIVE → REDUCE
POSITION_ACTIVE → RECOVERY_REVIEW
POSITION_ACTIVE → EXIT
EXIT → REENTRY_WAIT → REENTRY_CONFIRMATION → POSITION_ACTIVE
ANY_STATE → EMERGENCY_EXIT
ANY_STATE → WAIT / DATA_DEGRADED
```

### Phase 12: Audit Trail
Every action must store:
- input snapshot,
- timestamp,
- model selected,
- calculations,
- gates passed/failed,
- order intent,
- executed quantity/price if available,
- resulting ledger state,
- reason for state transition.

# 6A. Legacy 6-Phase Pipeline Reference (Preserved for Backward Compatibility)

Every trade evaluated under this skill must progress through the structured 6-phase pipeline without omitting any verification stage:

```mermaid
flowchart TD
    subgraph Phase1_Ingestion ["Phase 1: Ingestion & Fundamental Quality Gate"]
        direction TB
        P1["Data Intake & Ticker Normalization<br/>(LTP, OHLCV, Bankroll W, ADV, Debt/Equity)"]
        GateCheck{"Fundamental Quality Gate?<br/>(Large-Cap, Debt/Equity < 0.8, ADV > $500k)"}
        AbortGate["REJECT TRADE:<br/>Fails Fundamental Gate or Illiquid"]
        P1 --> GateCheck
        GateCheck -- "FAIL (High Debt / Penny)" --> AbortGate
    end

    subgraph Phase2_Selection ["Phase 2: Strategy & Model Dispatch"]
        direction TB
        ModelRoute{"Trading Intent & Regime?"}
        RouteV2["Select TRM V2 Engine<br/>(Core Blue-Chip Dip Averaging)"]
        RouteV3["Select TRM V3 Engine<br/>(Active Swing & Runner Retention)"]
        ModelRoute -- "Multi-Week Correction" --> RouteV2
        ModelRoute -- "Structural Reversal / Swing" --> RouteV3
    end

    subgraph Phase3_Calculation ["Phase 3: Mathematical Calculus & Dimensional Gate"]
        direction TB
        P3["Compute State Vector & Tranches<br/>(V2: B_next = 2 × ActualLoss | V3: Profit/Loss Path)"]
        UnitCheck{"Dimensional Integrity Valid?<br/>(Currency → % → Currency → Quantity)"}
        FixUnits["Recalibrate Unit Conversions<br/>(Strictly forbid dividing R by P_entry)"]
        P3 --> UnitCheck
        UnitCheck -- "FAIL" --> FixUnits --> P3
    end

    subgraph Phase4_Execution ["Phase 4: Sizing & Limit Order Generation"]
        direction TB
        P4["Calculate Integer Shares & Limit Orders<br/>(Shares = ⌊Allocation / LimitPrice⌋)"]
        OrderDispatch["Emit Limit Order & Bracket Stops<br/>(Day / GTC Order on Cash Equity)"]
        P4 --> OrderDispatch
    end

    subgraph Phase5_Lifecycle ["Phase 5: Trade Lifecycle Monitoring & Trailing"]
        direction TB
        P5["Active Market Position Tracking"]
        Outcome{"Real-Time Price Action Event?"}
        P5Reset["Target Reached (T1 / T2):<br/>Liquidate Cash Profit / Harvest Runner Q_retain"]
        P5Dip["Dip Trigger Hit (-10%, -20%):<br/>Prepare Tranche n + 1 (Max n ≤ 5)"]
        P5 --> Outcome
        Outcome -- "Profit Target Hit" --> P5Reset
        Outcome -- "Next Dip Level Hit" --> P5Dip
    end

    subgraph Phase6_Containment ["Phase 6: Invalidation & Circuit Breakers"]
        direction TB
        P6["EMERGENCY CUT-LOSS:<br/>Structural Invalidation or Max n=5 Breached"]
        Outcome -- "Stop Hit / Invalidation" --> P6
    end

    subgraph UniversalResetPhase ["Universal Cycle Reset"]
        direction TB
        ResetNode["UNCONDITIONAL RESET TO BASE S<br/>(Purge Tranche Counter, RoundNumber = 1)"]
    end

    GateCheck -- "PASS" --> ModelRoute
    RouteV2 --> P3
    RouteV3 --> P3
    UnitCheck -- "PASS" --> P4
    OrderDispatch --> P5
    P5Dip -.->|Loop Back to Recalculate State| P3
    P5Reset ==> ResetNode
    P6 ==> ResetNode
    ResetNode -.->|Ready for Next Trade| P1

    classDef primary fill:#1E293B,stroke:#38BDF8,stroke-width:2px,color:#F8FAFC;
    classDef decision fill:#312E81,stroke:#818CF8,stroke-width:2px,color:#FFFFFF;
    classDef success fill:#064E3B,stroke:#34D399,stroke-width:2px,color:#ECFDF5;
    classDef abort fill:#7F1D1D,stroke:#F87171,stroke-width:2px,color:#FEF2F2;
    classDef loop fill:#78350F,stroke:#FBBF24,stroke-width:2px,color:#FFFBEB;

    class P1,RouteV2,RouteV3,P3,FixUnits,P4,OrderDispatch,P5 primary;
    class GateCheck,ModelRoute,UnitCheck,Outcome decision;
    class P5Reset,ResetNode success;
    class AbortGate,P6 abort;
    class P5Dip loop;
```

### Phase 1: Intake, Ticker Validation & Fundamental Quality Gate
- Ingest ticker, exchange, current price, user bankroll ($W$), and trade history.
- Run the 5-point Fundamental Quality Gate (Market Cap, Debt/Equity, Pledging, ADV, Bankruptcy Immunity).
- Verify 100% unleveraged cash equity status. If non-equity or leveraged, abort immediately.

### Phase 2: TRM Model Selection
- **Select TRM V2 (Dip Averaging)**: When accumulating a core large-cap blue chip or index ETF during a multi-week correction. Tranches are planned at discrete price drops (-10%, -20%, -30%).
- **Select TRM V3 (Dual-Path)**: When executing swing trading setups where profit targets and structural stops are clearly defined, enabling runner equity accumulation on wins and calibrated recovery on losses.

### Phase 3: Mathematical State Calculation & Dimensional Unit Checks

The quantitative calculation engine processes the market quote through the designated TRM model equation:

| Execution Pipeline Mode | Mathematical Model | Primary Calculation Equation | Output Metric | Unit | Provenance |
| :--- | :---: | :--- | :---: | :---: | :---: |
| **Blue-Chip Dip Averaging** | TRM V2 | $$\mathbf{B_{\text{next}} = 2.0 \times \max(0, T_{\text{invested}} - Q_{\text{total}} P_{\text{current}})}$$ | $B_{\text{next}}$ | $\$$ | `[CALCULATED]` |
| **Profit Runner Retention** | TRM V3 | $$\mathbf{Q_{\text{retain}} = \frac{I \times p \times r}{P_{\text{current}}} \quad \text{with} \quad R = I \cdot p \cdot r}$$ | $Q_{\text{retain}}$ | shares | `[CALCULATED]` |
| **Stop-Loss Recovery Surcharge** | TRM V3 | $$\mathbf{A_{\text{next}} = A_{\text{base}} + (I \times l \times f)}$$ | $A_{\text{next}}$ | $\$$ | `[CALCULATED]` |

- **Dimensional Integrity Gate**: Confirm that share quantity conversions divide strictly by current execution price $P_{\text{current}}$ (never entry price $P_0$). Confirm that cash allocations do not mix currency with percentages.

### Phase 4: Execution Order Generation

Monetary allocations are translated into exact integer share quantities for delivery equity orders:

| Order Execution Parameter | Governing Mathematical Equation | Dimensional Unit | Provenance |
| :--- | :--- | :---: | :---: |
| **Integer Delivery Shares** | $$\mathbf{\text{Shares} = \left\lfloor \frac{\text{Allocation}}{P_{\text{limit}}} \right\rfloor}$$ | whole shares | `[CALCULATED]` |
| **Committed Cash Value** | $$\mathbf{\text{Cash Committed} = \text{Shares} \times P_{\text{limit}} \le \text{Allocation}}$$ | $\$$ | `[CALCULATED]` |
| **Residual Cash Buffer** | $$\mathbf{\text{Residual Buffer} = \text{Allocation} - \text{Cash Committed}}$$ | $\$$ | `[CALCULATED]` |

### Phase 5: Trade Lifecycle & Active Management
- Track active position against staged targets:
  - Target 1 (Conservative / Breakeven Shift): Regression Channel Mean.
  - Target 2 (Base Case): Regression Channel Upper Band (+2σ).
  - Target 3 (Runner): Trailing stop via 20 EMA or SuperTrend.
- Execute profit retention and runner transfer upon Target 2 hit.
- On completed recovery/exit, reset escalation state to base $S$; V7/V8 active harvesting remains in-cycle until a valid exit/reset event.

### Phase 6: Invalidation Enforcement & Hard Circuit Breakers
- If a daily bar closes below the **Structural Invalidation Stop**, execute full market liquidation.
- If averaging round reaches **$n = 5$** without price recovery, trigger emergency freeze and hard exit.
- If total cumulative allocation in a single stock reaches **15% of total portfolio equity**, freeze further tranches.

---

# 7. V8 Decision Scorecard & Calibration

The legacy 100-point rubric in this file remains available as a research scorecard. V8 must not treat an arbitrary score as a prediction of returns.

### 7.1 Evidence Dimensions
| Dimension | Suggested maximum | Evidence |
|---|---:|---|
| Data integrity | 15 | Freshness, completeness, consistency |
| Fundamental/solvency quality | 15 | Debt, cash flow, governance, event risk |
| Macro/market/sector regime | 10 | Index + sector trend alignment |
| Structural price action | 15 | Support/resistance, swing structure |
| Regression/valuation location | 10 | Mean/upper/lower channel context |
| Momentum/volume | 10 | RSI/MACD/ATR/volume/RVOL |
| Risk/reward & stop validity | 15 | Stop distance, expected payoff, bankroll |
| Liquidity/execution economics | 10 | Spread, depth, slippage, transaction cost |
| **Total** | **100** | **Research aid only** |

### 7.2 Gating vs Scoring
Hard gates are binary and override scores:
- invalid data,
- structural breakdown,
- insufficient cash,
- maximum position cap breached,
- prohibited instrument,
- unrecoverable event/fundamental impairment.

A high score can never override a hard gate.

### 7.3 Calibration Policy
Initial weights/thresholds are configuration values, not proven optimal constants. They MUST be calibrated through historical walk-forward validation and locked before out-of-sample evaluation to reduce overfitting.


### 7.4 Legacy Quantitative Setup Scoring System (100-Point Rubric)

Every trade evaluated under this skill is scored across 8 objective dimensions:

| Dimension | Max Points | Evaluation Criteria |
| :--- | :---: | :--- |
| **1. Fundamental Health & Bankruptcy Immunity** | 15 | Debt/Equity $< 0.5$ (15), D/E $0.5–0.8$ (10), Clean Promoter Holding (5 bonus), High Debt (0). |
| **2. Macro Trend Alignment** | 15 | Weekly SuperTrend Bullish + Price $> 200$ SMA (15), Price $> 50$ SMA (10), Below 200 SMA (0). |
| **3. Regression Channel Confluence** | 15 | Price at Lower Channel ($-2\sigma$) demand zone (15), At Mean (8), At Upper Channel (0). |
| **4. Price Action & Dip Confirmation** | 15 | Bullish reversal candlestick (Hammer, Bullish Engulfing) + high volume rejection at level (15). |
| **5. Indicator & Momentum Alignment** | 10 | Daily RSI rebounding from $< 35$ or $> 50$ + MACD histogram ticking upward (10), Bearish (0). |
| **6. Risk/Reward & Bankroll Safety** | 15 | $R:R \ge 1:2.5$ (15), $R:R \ge 1:1.5$ (10), Bankroll covers $n_{\max} \ge 5$ rounds (5 bonus). |
| **7. Market & Sector Tailwinds** | 10 | Benchmark index (Nifty / S&P 500) and Sector Index in confirmed uptrend (10), Choppy (5). |
| **8. Catalyst Cleanliness** | 5 | Clean calendar; no corporate earnings, SEBI/SEC actions, or binary events within 48h (5). |
| **TOTAL SCORE** | **100** | **Sum of all 8 dimensions** |

### Score Classifications & Action Thresholds
- **85 – 100 Points (`STRONG`)**: High-conviction setup. Authorize full Tranche 1 ($B_1$) execution.
- **65 – 84 Points (`GOOD`)**: Acceptable setup. Authorize execution provided $R:R \ge 1:1.5$.
- **50 – 64 Points (`WATCHLIST`)**: Marginal quality. Issue **WAIT** recommendation; await pullback or confirmation.
- **< 50 Points (`AVOID / NO TRADE`)**: Setup disqualified. Prohibit trade execution.

---


# 8. Pre-Flight Execution Checklist

Before publishing or routing any TRM trade order, the agent must verify every item in this interactive checklist:

```markdown
## TRM Pre-Flight Verification Checklist
- [ ] **1. Instrument Eligibility**: Is the ticker a high-liquidity, large/mid-cap equity passing the Bankruptcy Immunity Gate (Debt/Equity < 0.8, clean governance)?
- [ ] **2. Cash-Equity Scope**: Is the trade 100% unleveraged spot/cash equity (strictly ZERO F&O, MTF margin, or shorting)?
- [ ] **3. TRM Model Match**: Is the correct model designated (TRM V2 for dip averaging; TRM V3 for dual-path swing trading)?
- [ ] **4. Dimensional Arithmetic Validation**: Have all conversions adhered to Currency -> % -> Currency -> Shares? (Confirming $Q_{retain} = R / P_{current}$, NOT $P_0$).
- [ ] **5. Bankroll Capacity Verification**: Is the maximum surviving streak calculated ($n_{\max}$), and is cumulative allocation strictly capped at $\le 15\%$ of total portfolio equity?
- [ ] **6. Round Limit Guardrail**: Is the maximum averaging round hard-capped at $\le 5$ rounds?
- [ ] **7. Objective Structural Stop**: Is the invalidation stop anchored to a structural support swing floor / regression lower band plus ATR buffer?
- [ ] **8. Binary Event Clearance**: Is the company free of earnings releases or major regulatory rulings over the next 48 hours?
- [ ] **9. Universal Reset Invariant**: Is the execution order configured to liquidate target tranches and reset capital to base $S$ immediately upon recovery?
```

---

# 9. Comprehensive Edge Cases & Anomaly Protocols

| # | Edge Case / Market Hazard | Specific Condition | Mandatory System Protocol |
|---|---------------------------|--------------------|---------------------------|
| 1 | **Upper / Lower Circuit Limit** | Stock hits daily circuit breaker; trading halted or frozen | Freeze all pending TRM orders. Emit `⚠️ CIRCUIT BREAKER ACTIVE`. Do NOT place market orders. Provide structural levels to monitor once liquidity unfreezes. |
| 2 | **Severe Illiquidity / Spread Blowout** | RVOL $< 0.5$ or Bid-Ask Spread $> 0.50\%$ | Emit `⚠️ LOW LIQUIDITY WARNING`. Strictly prohibit market orders; mandate limit orders only. Reduce tranche size by 50%. |
| 3 | **Gap Down Beyond Invalidation Stop** | Stock gaps down overnight directly past the structural stop | Do NOT hold and hope. Mandate immediate market exit on opening bell (or within initial 15-minute price discovery); record realized loss and execute TRM V3 Loss Recovery. |
| 4 | **Gap Down Slicing Tranche Levels** | Price gaps down directly from Tranche 1 past Tranche 2 into Tranche 3 | Do NOT execute duplicate orders simultaneously. Wait for 15-minute candle close to confirm price stabilization; consolidate Tranches 2 and 3 at the stabilized market price within risk caps. |
| 5 | **Corporate Actions (Stock Split / Bonus)** | 1:5 or 1:10 price adjustment overnight | Emit `⚠️ PRICE DISCONTINUITY DETECTED`. Recalculate historical average cost and trigger levels using split-adjusted series before evaluating any TRM trigger. |
| 6 | **Corporate Dividend Payout** | Large dividend ex-date price drop | Dividend received in bank account is credited against cumulative invested capital ($T_{\text{invested}} \leftarrow T_{\text{invested}} - \text{DividendCash}$). Recalculate break-even average cost. |
| 7 | **Fundamental Impairment / Fraud** | Promoter fraud, default, pledging $> 50\%$, or accounting inquiry | **IMMEDIATE DISQUALIFICATION**. Permanently ban further TRM averaging; execute unconditional full position liquidation regardless of paper loss. |
| 8 | **Capital Exhaustion ($n > n_{\max}$)** | Maximum planned averaging rounds ($N=5$) reached without rebound | Trigger **HARD CYCLE TERMINATION**. Liquidate position to preserve remaining portfolio equity. Strictly prohibit borrowing or diverting capital from emergency reserves. |
| 9 | **Severe Price Anchoring** | User's historical entry price is $> 30\%$ above current price | Issue `Anti-Anchoring Alert`. State that historical purchase price is irrelevant to market structure. Base all analysis 100% on current regression channels and technical structure. |
| 10 | **Parabolic Climax Run** | Price extends $> 3.0\times$ ATR above 20 EMA or Upper Channel | Enforce `Anti-FOMO Protocol`. Strictly prohibit buying Tranche 1 at extended highs. Wait for mean reversion to 20 EMA or regression median. |
| 11 | **Binary Event Imminent** | Corporate earnings announcement scheduled within 48 hours | Enforce `⚠️ BINARY EVENT RISK`. Prohibit initiating new TRM cycles or deploying averaging tranches until post-earnings volatility crush clears. |
| 12 | **Indian F&O Expiry Week** | Analysis conducted within 48 hours of monthly derivative expiry | Emit warning that institutional rollover flows may create false breakouts/breakdowns in underlying cash equities. Require wider ATR stop buffers ($0.75\times$ ATR). |

---

# 10. Standardized Output Contracts (Copy-Paste ASCII Templates)

### 10.1 Mode A: TRM V2 Dip-Averaging Execution Report
```text
================================================================================
TRM V2 DIP-AVERAGING EXECUTION REPORT: [SYMBOL] ([EXCHANGE])
================================================================================

1. DATA PROVENANCE & CONTEXT
--------------------------------------------------------------------------------
• Last Traded Price (LTP)        : [₹/$ Price] [OBSERVED]
• Data Feed State                : [REAL-TIME / DELAYED / EOD] [OBSERVED]
• Session Volume & RVOL          : [Volume] (RVOL: [X.Xx]) [OBSERVED]
• Daily ATR (14-Period)          : [₹/$ Value] [CALCULATED]
• Account Allocated Bankroll (W) : [₹/$ Amount] [USER-PROVIDED]
• Max Permitted Rounds (N_max)   : [3 to 5 Rounds] [CALCULATED]

2. FUNDAMENTAL QUALITY & ELIGIBILITY GATE
--------------------------------------------------------------------------------
• Market Capitalization Tier     : [Large-Cap / High Mid-Cap] [OBSERVED]
• Debt-to-Equity Ratio           : [X.XX] (< 0.80 Requirement Passed) [OBSERVED]
• Promoter Pledging Status       : [X.X%] (Zero Pledging Passed) [OBSERVED]
• 20-Day Average Daily Turnover  : [₹/$ Amount] (> ₹10 Cr / $5M Passed) [OBSERVED]
• Eligibility Status             : [PASSED / REJECTED]

3. TECHNICAL REGIME & REGRESSION CONFLUENCE
--------------------------------------------------------------------------------
• Macro Trend (Weekly / Daily)   : [Uptrend / Pullback in Bull Trend] [OBSERVED]
• 200 SMA / 50 SMA Alignment     : [Price > 200 SMA; 20 EMA > 50 SMA] [OBSERVED]
• Linear Regression Channel Pos  : [Lower Band (-2σ) Demand Zone] [OBSERVED]
• Daily RSI (14) & Momentum      : [RSI Value] ([Oversold Bounce / Bullish]) [OBSERVED]

4. TRM V2 ADAPTIVE CAPITAL ALLOCATION PLAN
--------------------------------------------------------------------------------
Base Tranche (S): [₹/$ Amount] | Allocation Multiplier: 2.0x Actual Loss

| Tranche | Trigger Event | Trigger Price | Tranche Size (Bn) | Cumul. Capital (Tn) | Cumul. Shares | Avg Cost (AC) | Mkt Value | Actual Loss | Target (+5%) |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **B1** | Initial Entry | [₹/$ Price]   | [₹/$ B1]          | [₹/$ T1]            | [Q1]          | [₹/$ AC1]     | [₹/$ MV1] | ₹0.00       | [₹/$ Tgt1]   |
| **B2** | -10.0% Dip    | [₹/$ Price]   | [₹/$ B2]          | [₹/$ T2]            | [Q2]          | [₹/$ AC2]     | [₹/$ MV2] | [₹/$ Loss2] | [₹/$ Tgt2]   |
| **B3** | -20.0% Dip    | [₹/$ Price]   | [₹/$ B3]          | [₹/$ T3]            | [Q3]          | [₹/$ AC3]     | [₹/$ MV3] | [₹/$ Loss3] | [₹/$ Tgt3]   |
| **B4** | -30.0% Dip    | [₹/$ Price]   | [₹/$ B4]          | [₹/$ T4]            | [Q4]          | [₹/$ AC4]     | [₹/$ MV4] | [₹/$ Loss4] | [₹/$ Tgt4]   |

5. RISK MANAGEMENT & THESIS INVALIDATION
--------------------------------------------------------------------------------
• Structural Invalidation Stop   : [₹/$ Price] [CALCULATED]
• Stop-Loss Placement Logic      : [Below Major Swing Floor - 0.5x ATR buffer]
• Max Portfolio Risk at Stop     : [₹/$ Amount] ([X.X%] of Bankroll) [CALCULATED]
• Max Averaging Round Cap        : [Hard Cap at Round 4] [CALCULATED]

6. SETUP QUALITY SCORECARD (100-POINT RUBRIC)
--------------------------------------------------------------------------------
• Fundamental & Solvency Health  : [X]/15
• Macro Trend Alignment          : [X]/15
• Regression Channel Confluence  : [X]/15
• Price Action & Volume Reversal : [X]/15
• Indicator Momentum Alignment   : [X]/10
• Risk/Reward & Bankroll Safety  : [X]/15
• Market & Sector Tailwinds      : [X]/10
• Catalyst Cleanliness           : [X]/5
--------------------------------------------------------------------------------
• TOTAL SCORE                    : [XX]/100
• CLASSIFICATION                 : [STRONG / GOOD / WATCHLIST / AVOID]

================================================================================
MANDATORY EXECUTIVE ACTION DIRECTIVE
================================================================================
ACTION DIRECTIVE       : [EXECUTE TRANCHE B1 / DEPLOY TRANCHE B_N / WAIT / NO TRADE]
LIMIT ORDER PRICE      : [₹/$ Exact Price]
SHARE QUANTITY         : [Exact Integer Shares]
CAPITAL DEPLOYED       : [₹/$ Allocation]
RECOVERY TARGET EXIT   : [₹/$ Target Price]
STRUCTURAL STOP-LOSS   : [₹/$ Cut-Loss Price]
CYCLE RESET INVARIANT  : On recovery target fill, liquidate 100% and reset to S.
================================================================================
```

---

### 10.2 Mode B: TRM V3 Dual-Path Execution Report
```text
================================================================================
TRM V3 DUAL-PATH EXECUTION REPORT: [SYMBOL] ([EXCHANGE])
================================================================================

1. ACTIVE TRADE SPECIFICATION & UNIT AUDIT
--------------------------------------------------------------------------------
• Asset Symbol & Exchange        : [SYMBOL] ([EXCHANGE]) [OBSERVED]
• Entry Price (P_entry)          : [₹/$ Price] [USER-PROVIDED]
• Current Market Price (P_curr)  : [₹/$ Price] [OBSERVED]
• Original Position Shares (Q_0) : [Shares] [USER-PROVIDED]
• Initial Invested Capital (I)   : [₹/$ Capital = P_entry * Q_0] [CALCULATED]
• Configured Profit Target (p)   : [+XX.X%] [USER-PROVIDED]
• Configured Stop-Loss (l)       : [-XX.X%] [USER-PROVIDED]
• Profit Retention Factor (r)    : [XX.X%] [USER-PROVIDED]
• Loss Recovery Factor (f)       : [XX.X%] [USER-PROVIDED]
• Base Reversal Allocation (A_b) : [₹/$ Base Amount] [USER-PROVIDED]

2. PATH A: PROFIT RETENTION SCENARIO (TARGET REACHED)
--------------------------------------------------------------------------------
• Target Exit Price (P_target)   : [₹/$ Price = P_entry * (1 + p)] [CALCULATED]
• Gross Realized Profit Amount   : [₹/$ Amount = I * p] [CALCULATED]
• Retained Profit Value (R)      : [₹/$ Value = Gross Profit * r] [CALCULATED]
• Retained Shares (Q_retain)     : [Shares = R / P_target] (Current Price Rule!) [CALCULATED]
• Shares to Liquidate (Q_sell)   : [Shares = Q_0 - Q_retain] [CALCULATED]
• Cash Returned to Account       : [₹/$ Cash = Q_sell * P_target] [CALCULATED]
• Permanent Runner Equity Value  : [₹/$ Equity = Q_retain * P_target] [CALCULATED]
• Action Mandate                 : Sell Q_sell shares; vault Q_retain shares; RESET active trade to S.

3. PATH B: LOSS RECOVERY REVERSAL SCENARIO (STOP HIT)
--------------------------------------------------------------------------------
• Stop-Loss Exit Price (P_stop)  : [₹/$ Price = P_entry * (1 - l)] [CALCULATED]
• Gross Realized Loss Amount     : [₹/$ Amount = I * l] [CALCULATED]
• Recovery Surcharge Amount      : [₹/$ Surcharge = Loss * f] [CALCULATED]
• Next Reversal Trade Allocation : [₹/$ Next Capital = A_base + Surcharge] [CALCULATED]
• Next Trade Required Shares     : [Shares = Next Capital / Reversal Entry Price] [CALCULATED]
• Action Mandate                 : Close active position; route Next Capital into new setup.

================================================================================
MANDATORY EXECUTIVE ACTION DIRECTIVE
================================================================================
CURRENT STATE          : [ACTIVE / PROFIT TARGET HIT / STOP-LOSS HIT]
ACTIVE ORDER           : [ROUTE REVERSAL ORDER / HARVEST PROFIT & VAULT RUNNER]
CASH CREDITED          : [₹/$ Amount]
RUNNER SHARES VAULTED  : [Shares to Long-Term Compounding Vault]
NEXT CYCLE ALLOCATION  : [₹/$ Reset to Base S or Reversal Allocation]
================================================================================
```

---

### 10.3 Mode C: TRM State Machine & Multi-Round Portfolio Audit
```text
================================================================================
TRM STATE MACHINE & MULTI-ROUND PORTFOLIO AUDIT
================================================================================

PORTFOLIO RISK & BANKROLL CONTEXT
--------------------------------------------------------------------------------
• Total Portfolio Net Worth      : [₹/$ Amount] [USER-PROVIDED]
• Dedicated TRM Bankroll (W)     : [₹/$ Amount] ([XX%] of Portfolio) [USER-PROVIDED]
• Single-Asset Capital Cap       : [₹/$ Max 15% Cap] [CALCULATED]
• Active TRM Cycles Running      : [Count of Active Tickers]

ACTIVE CYCLE STATE VECTORS
--------------------------------------------------------------------------------
| Ticker | Mode | Cycle ID | Active Tranche | Invested (T) | Mkt Value (MV) | Unrealized P/L | Next Allocation | Max Cap Status |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| [TICK1] | V2 | #CYC-01 | B2 (-10%) | [₹/$ Tn] | [₹/$ MV] | [₹/$ PnL] | [₹/$ Next] | [X.X% / 15% Max] |
| [TICK2] | V3 | #CYC-02 | Active Long | [₹/$ Tn] | [₹/$ MV] | [₹/$ PnL] | [₹/$ Reset] | [X.X% / 15% Max] |

SYSTEM AUDIT & COMPLIANCE VERIFICATION
--------------------------------------------------------------------------------
• Leverage Status                : ZERO LEVERAGE CONFIRMED (100% Cash Equity)
• Dimensional Separation Check   : PASSED (Zero unit mixture detected)
• Bankroll Capacity Buffer       : PASSED ([X] Tranches of reserve remaining)
• Emergency Circuit Breakers     : ACTIVE (All stops programmed at exchange)
================================================================================
```

---

### 10.4 TRM V7/V8 Master Execution Report
```text
================================================================================
TRM V8 MASTER ADAPTIVE EXECUTION REPORT: [SYMBOL] ([EXCHANGE])
================================================================================

1. INPUT / PROVENANCE
--------------------------------------------------------------------------------
• Timestamp / Feed State       : [TIME] [REAL-TIME / DELAYED / EOD] [OBSERVED]
• Current Price                : [PRICE] [OBSERVED]
• Quantity Held                : [QTY] [USER-PROVIDED / OBSERVED]
• Average Entry               : [PRICE] [USER-PROVIDED / CALCULATED]
• Initial Investment           : [AMOUNT] [CALCULATED]
• Available Cash               : [AMOUNT] [USER-PROVIDED]

2. ACTUAL P/L + WEALTH LEDGER
--------------------------------------------------------------------------------
• Current Market Value         : [AMOUNT] [CALCULATED]
• Position P/L                 : [AMOUNT] [CALCULATED]
• Position P/L %               : [%] [CALCULATED]
• Harvested Cash               : [AMOUNT] [CALCULATED]
• Cycle Wealth                 : [AMOUNT] [CALCULATED]
• Cycle P/L                    : [AMOUNT] [CALCULATED]
• Peak Cycle Wealth            : [AMOUNT] [CALCULATED]
• Current Drawdown             : [%] [CALCULATED]
• Wealth Floor                 : [AMOUNT] [CALCULATED]

3. MARKET / STRUCTURAL STATE
--------------------------------------------------------------------------------
• Regime                       : [BULLISH_STRONG / ...]
• Trend State                  : [STATE]
• Support / Resistance         : [LEVELS] [OBSERVED]
• Regression Channel           : [MEAN / UPPER / LOWER] [CALCULATED]
• RSI / MACD / ATR / RVOL      : [VALUES] [OBSERVED/CALCULATED]
• Sector / Index Alignment     : [STATE]
• Structural Invalidation      : [PRICE]

4. V8 MODEL ROUTING
--------------------------------------------------------------------------------
• Selected Engine              : [V7 / V5 / V6 / V4 / V3 / V2 / WAIT]
• Routing Reason               : [FACTUAL EVIDENCE]
• External Analysis Validity   : [PASS / DEGRADED / FAIL]
• Hard Gate Status             : [PASS / BLOCKED]

5. V7 HARVEST (WHEN APPLICABLE)
--------------------------------------------------------------------------------
• Protected Principal          : [AMOUNT]
• Current Position Value       : [AMOUNT]
• Excess Above Principal       : [AMOUNT]
• Shares To Sell               : [QTY]
• Shares To Keep               : [QTY]
• Expected Residual Value      : [AMOUNT]
• Estimated Costs              : [AMOUNT]
• Economic Action Gate         : [PASS / BLOCKED]

6. LOSS / RECOVERY (WHEN APPLICABLE)
--------------------------------------------------------------------------------
• Actual Cycle Loss            : [AMOUNT]
• Recovery Engine              : [V2 / V3 / V4 / V5 / V6 / NONE]
• Risk Budget                  : [AMOUNT]
• Max Position Cap             : [AMOUNT]
• Candidate Allocation         : [AMOUNT]
• Final Approved Allocation    : [AMOUNT]
• Recovery Round               : [N / MAX]

7. EXECUTION DECISION
--------------------------------------------------------------------------------
ACTION : [NO_TRADE / WAIT / HOLD / HARVEST / REDUCE / ENTER / REENTER / RECOVER / EXIT / EMERGENCY_EXIT]
LIMIT PRICE                  : [PRICE]
QUANTITY                     : [INTEGER SHARES]
CAPITAL / CASH FLOW          : [AMOUNT]
STOP / INVALIDATION          : [PRICE]
COOLDOWN / HYSTERESIS        : [STATE]

8. AUDIT REASON
--------------------------------------------------------------------------------
[ONE-PARAGRAPH FACTUAL RATIONALE]

================================================================================
V8 SAFETY INVARIANTS: DATA • STRUCTURE • WEALTH FLOOR • COST • CAPITAL CAP
================================================================================
```


# 11. Production Reference Implementations

### 11.1 Python Production Engine (Decimal-Precise)
```python
"""
TRM Quantitative Engine - Production Reference Implementation
Precision: Decimal arithmetic conforming to repository dimensional integrity standards.
"""
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, NamedTuple

class TRMOriginalResult(NamedTuple):
    sequence: List[Decimal]
    total_invested: Decimal
    next_investment: Decimal
    bankroll_required: Decimal

class TRMV2Step(NamedTuple):
    round_num: int
    price: Decimal
    investment: Decimal
    shares_bought: Decimal
    total_invested: Decimal
    total_shares: Decimal
    average_cost: Decimal
    market_value: Decimal
    actual_loss: Decimal

class TRMV3ProfitResult(NamedTuple):
    entry_price: Decimal
    current_price: Decimal
    original_quantity: Decimal
    initial_investment: Decimal
    profit_amount: Decimal
    retained_profit_value: Decimal
    retained_quantity: Decimal
    quantity_to_sell: Decimal
    cash_realized: Decimal

class TRMV3LossResult(NamedTuple):
    entry_price: Decimal
    stop_price: Decimal
    original_quantity: Decimal
    initial_investment: Decimal
    loss_amount: Decimal
    recovery_addition: Decimal
    base_reversal_investment: Decimal
    next_reversal_investment: Decimal


class TRMCalculator:
    PRECISION = Decimal('0.00000001')
    CURRENCY_PREC = Decimal('0.01')

    @classmethod
    def calculate_trm_original(cls, base_investment: Decimal, max_losses: int) -> TRMOriginalResult:
        bet = base_investment
        total = Decimal('0')
        sequence = []
        for i in range(1, max_losses + 1):
            sequence.append(bet)
            total += bet
            bet = total * Decimal('2')
        next_investment = bet
        bankroll_required = total + next_investment
        return TRMOriginalResult(sequence, total, next_investment, bankroll_required)

    @classmethod
    def calculate_trm_v2_step(cls, 
                              current_price: Decimal, 
                              total_shares: Decimal, 
                              total_invested: Decimal, 
                              loss_multiplier: Decimal = Decimal('2.0')) -> Dict[str, Decimal]:
        market_value = (total_shares * current_price).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        actual_loss = max(Decimal('0'), total_invested - market_value)
        next_investment = (actual_loss * loss_multiplier).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        
        additional_shares = Decimal('0')
        if next_investment > Decimal('0') and current_price > Decimal('0'):
            additional_shares = (next_investment / current_price).quantize(cls.PRECISION, rounding=ROUND_HALF_UP)
            
        new_total_invested = total_invested + next_investment
        new_total_shares = total_shares + additional_shares
        new_avg_cost = (new_total_invested / new_total_shares).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP) if new_total_shares > 0 else Decimal('0')

        return {
            "market_value": market_value,
            "actual_loss": actual_loss,
            "next_investment": next_investment,
            "additional_shares": additional_shares,
            "new_total_invested": new_total_invested,
            "new_total_shares": new_total_shares,
            "new_average_cost": new_avg_cost
        }

    @classmethod
    def calculate_trm_v3_profit(cls, 
                                 entry_price: Decimal, 
                                 original_quantity: Decimal, 
                                 profit_percent: Decimal, 
                                 retention_factor: Decimal) -> TRMV3ProfitResult:
        initial_investment = (entry_price * original_quantity).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        current_price = (entry_price * (Decimal('1') + profit_percent)).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        profit_amount = (initial_investment * profit_percent).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        retained_profit_value = (profit_amount * retention_factor).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        
        # STRICT DIMENSIONAL RULE: Convert at current_price, NOT entry_price!
        retained_quantity = (retained_profit_value / current_price).quantize(cls.PRECISION, rounding=ROUND_HALF_UP)
        quantity_to_sell = original_quantity - retained_quantity
        cash_realized = (quantity_to_sell * current_price).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)

        return TRMV3ProfitResult(
            entry_price=entry_price,
            current_price=current_price,
            original_quantity=original_quantity,
            initial_investment=initial_investment,
            profit_amount=profit_amount,
            retained_profit_value=retained_profit_value,
            retained_quantity=retained_quantity,
            quantity_to_sell=quantity_to_sell,
            cash_realized=cash_realized
        )

    @classmethod
    def calculate_trm_v3_loss(cls, 
                               entry_price: Decimal, 
                               original_quantity: Decimal, 
                               loss_percent: Decimal, 
                               recovery_factor: Decimal, 
                               base_reversal_investment: Decimal) -> TRMV3LossResult:
        initial_investment = (entry_price * original_quantity).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        stop_price = (entry_price * (Decimal('1') - loss_percent)).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        loss_amount = (initial_investment * loss_percent).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        recovery_addition = (loss_amount * recovery_factor).quantize(cls.CURRENCY_PREC, rounding=ROUND_HALF_UP)
        next_reversal_investment = base_reversal_investment + recovery_addition

        return TRMV3LossResult(
            entry_price=entry_price,
            stop_price=stop_price,
            original_quantity=original_quantity,
            initial_investment=initial_investment,
            loss_amount=loss_amount,
            recovery_addition=recovery_addition,
            base_reversal_investment=base_reversal_investment,
            next_reversal_investment=next_reversal_investment
        )
```

---

### 11.2 Java 17+ Enterprise Implementation
```java
package com.trm.algorithm.engine;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

/**
 * Enterprise Production Engine for the Teja Recovery Model (TRM).
 * Enforces strict BigDecimal rounding, immutable record state, and cash-equity bounds.
 */
public final class TRMEngine {

    private static final int SHARE_SCALE = 8;
    private static final int CURRENCY_SCALE = 2;
    private static final RoundingMode ROUNDING = RoundingMode.HALF_UP;

    public record TRMOriginalResult(
        List<BigDecimal> sequence,
        BigDecimal totalInvested,
        BigDecimal nextInvestment,
        BigDecimal bankrollRequired
    ) {}

    public record TRMV2State(
        BigDecimal marketValue,
        BigDecimal actualLoss,
        BigDecimal nextInvestment,
        BigDecimal additionalShares,
        BigDecimal newTotalInvested,
        BigDecimal newTotalShares,
        BigDecimal newAverageCost
    ) {}

    public record TRMV3ProfitResult(
        BigDecimal entryPrice,
        BigDecimal currentPrice,
        BigDecimal originalQuantity,
        BigDecimal initialInvestment,
        BigDecimal profitAmount,
        BigDecimal retainedProfitValue,
        BigDecimal retainedQuantity,
        BigDecimal quantityToSell,
        BigDecimal cashRealized
    ) {}

    public record TRMV3LossResult(
        BigDecimal entryPrice,
        BigDecimal stopPrice,
        BigDecimal originalQuantity,
        BigDecimal initialInvestment,
        BigDecimal lossAmount,
        BigDecimal recoveryAddition,
        BigDecimal baseReversalInvestment,
        BigDecimal nextReversalInvestment
    ) {}

    public static TRMOriginalResult computeOriginal(BigDecimal baseInvestment, int maxLosses) {
        BigDecimal bet = baseInvestment;
        BigDecimal total = BigDecimal.ZERO;
        List<BigDecimal> sequence = new ArrayList<>();

        for (int i = 1; i <= maxLosses; i++) {
            sequence.add(bet);
            total = total.add(bet);
            bet = total.multiply(BigDecimal.valueOf(2));
        }
        BigDecimal nextInvestment = bet;
        BigDecimal bankrollRequired = total.add(nextInvestment);
        return new TRMOriginalResult(sequence, total, nextInvestment, bankrollRequired);
    }

    public static TRMV2State computeV2Step(
            BigDecimal currentPrice,
            BigDecimal totalShares,
            BigDecimal totalInvested,
            BigDecimal lossMultiplier) {

        BigDecimal marketValue = totalShares.multiply(currentPrice).setScale(CURRENCY_SCALE, ROUNDING);
        BigDecimal actualLoss = totalInvested.subtract(marketValue).max(BigDecimal.ZERO);
        BigDecimal nextInvestment = actualLoss.multiply(lossMultiplier).setScale(CURRENCY_SCALE, ROUNDING);

        BigDecimal additionalShares = BigDecimal.ZERO;
        if (nextInvestment.compareTo(BigDecimal.ZERO) > 0 && currentPrice.compareTo(BigDecimal.ZERO) > 0) {
            additionalShares = nextInvestment.divide(currentPrice, SHARE_SCALE, ROUNDING);
        }

        BigDecimal newTotalInvested = totalInvested.add(nextInvestment);
        BigDecimal newTotalShares = totalShares.add(additionalShares);
        BigDecimal newAverageCost = newTotalShares.compareTo(BigDecimal.ZERO) > 0
                ? newTotalInvested.divide(newTotalShares, CURRENCY_SCALE, ROUNDING)
                : BigDecimal.ZERO;

        return new TRMV2State(marketValue, actualLoss, nextInvestment, additionalShares, 
                              newTotalInvested, newTotalShares, newAverageCost);
    }

    public static TRMV3ProfitResult computeV3Profit(
            BigDecimal entryPrice,
            BigDecimal originalQuantity,
            BigDecimal profitPercent,
            BigDecimal retentionFactor) {

        BigDecimal initialInvestment = entryPrice.multiply(originalQuantity).setScale(CURRENCY_SCALE, ROUNDING);
        BigDecimal currentPrice = entryPrice.multiply(BigDecimal.ONE.add(profitPercent)).setScale(CURRENCY_SCALE, ROUNDING);
        BigDecimal profitAmount = initialInvestment.multiply(profitPercent).setScale(CURRENCY_SCALE, ROUNDING);
        BigDecimal retainedProfitValue = profitAmount.multiply(retentionFactor).setScale(CURRENCY_SCALE, ROUNDING);

        // MANDATORY: Divide by currentPrice, NOT entryPrice!
        BigDecimal retainedQuantity = retainedProfitValue.divide(currentPrice, SHARE_SCALE, ROUNDING);
        BigDecimal quantityToSell = originalQuantity.subtract(retainedQuantity);
        BigDecimal cashRealized = quantityToSell.multiply(currentPrice).setScale(CURRENCY_SCALE, ROUNDING);

        return new TRMV3ProfitResult(entryPrice, currentPrice, originalQuantity, initialInvestment,
                                     profitAmount, retainedProfitValue, retainedQuantity, quantityToSell, cashRealized);
    }

    public static TRMV3LossResult computeV3Loss(
            BigDecimal entryPrice,
            BigDecimal originalQuantity,
            BigDecimal lossPercent,
            BigDecimal recoveryFactor,
            BigDecimal baseReversalInvestment) {

        BigDecimal initialInvestment = entryPrice.multiply(originalQuantity).setScale(CURRENCY_SCALE, ROUNDING);
        BigDecimal stopPrice = entryPrice.multiply(BigDecimal.ONE.subtract(lossPercent)).setScale(CURRENCY_SCALE, ROUNDING);
        BigDecimal lossAmount = initialInvestment.multiply(lossPercent).setScale(CURRENCY_SCALE, ROUNDING);
        BigDecimal recoveryAddition = lossAmount.multiply(recoveryFactor).setScale(CURRENCY_SCALE, ROUNDING);
        BigDecimal nextReversalInvestment = baseReversalInvestment.add(recoveryAddition);

        return new TRMV3LossResult(entryPrice, stopPrice, originalQuantity, initialInvestment,
                                   lossAmount, recoveryAddition, baseReversalInvestment, nextReversalInvestment);
    }
}
```

---

### 11.3 V4–V8 Implementation Contracts

#### 11.3.1 Canonical domain model
Implement immutable calculation inputs and explicit state transitions. Recommended Java 17 types:
```text
record MarketSnapshot(...)
record PositionState(...)
record WealthLedger(...)
record RiskLimits(...)
record AnalysisSnapshot(...)
record V8Decision(...)
record ExecutionIntent(...)
```
Use `BigDecimal` for money, prices, percentages and share quantities that require high precision. Never use `double` for ledger reconciliation.

#### 11.3.2 Engine interface
```java
interface TrmEngine {
    TrmDecision evaluate(TrmContext context);
}
```
Recommended implementations:
```text
OriginalRecoveryEngine
V1PercentageTriggerEngine
V2ActualLossEngine
V3DualPathEngine
V4CurrentPriceShareEngine
V5AdaptivePnLEngine
V6ExitReentryEngine
V7PrincipalHarvestEngine
V8AdaptiveMetaController
```

#### 11.3.3 V8 controller contract
```java
V8Decision evaluate(V8Context context) {
    validate(context);
    if (context.hasHardDataFailure()) return WAIT;
    if (context.isStructurallyInvalid()) return EXIT;

    calculateActualPnL(context);
    calculateCycleWealth(context);
    calculateWealthFloor(context);
    classifyRegime(context);

    TrmEngine engine = routeToLeastAggressiveValidEngine(context);
    TrmDecision decision = engine.evaluate(context);

    decision = applyRiskCaps(decision, context);
    decision = applyLiquidityAndCostGate(decision, context);
    decision = applyAntiChurnGate(decision, context);
    decision = reconcileWholeShareQuantity(decision, context);

    return decision;
}
```

#### 11.3.4 Testing requirements
Mandatory unit/property tests:
- Original recurrence and closed forms.
- V1 price-trigger sequence.
- V2 actual-loss calculation.
- V3 actual observed P/L correction.
- V4 current-price share conversion.
- V5 actual P/L authority.
- V6 validation, clustering and re-entry confirmation.
- V7 principal preservation invariant.
- V8 wealth-floor invariant.
- No currency/share/percentage unit mixing.
- Whole-share rounding and residual reconciliation.
- No order when hard gates fail.
- No duplicate orders under anti-churn conditions.
- Cycle reset isolation.

#### 11.3.5 V7 invariant property test
For a successful harvest before costs/whole-share rounding:
$$Q_{keep}\times P_{current}=ProtectedPrincipal$$
Within execution tolerance, after order rounding and charges the residual difference must be recorded rather than silently discarded.

#### 11.3.6 V8 routing property test
Given identical validated inputs, routing MUST be deterministic. If input quality changes from valid to invalid, the decision must not become more aggressive; it must remain or move toward `WAIT`/`EXIT`.


# 12. End-to-End Concrete Numerical Case Studies

### 12.1 Case Study 1: Large-Cap Dip Averaging via TRM V2 (Reliance Industries / NSE)

##### A. Portfolio Account Context & Initial Parameters
- **Account Context**: Dedicated Cash Equity Account. Total Equity = ₹20,00,000.
- **Dedicated TRM Bankroll ($W$)**: ₹3,00,000 (Strict 15% Single-Stock Cap).
- **Target Asset**: Reliance Industries Ltd. (`RELIANCE.NS`).
- **Base Allocation ($S$)**: ₹28,000.00 (Round 1 Tranche).
- **Loss Multiplier ($m$)**: $2.0$.

##### B. Step-by-Step Execution Progression Ledger

| Execution Event | Stock Price ($P$) | Action Taken | Governing Math Equation | Tranche Deployed ($B_n$) | Total Invested ($T_n$) | Total Shares ($Q_n$) | Breakeven Cost ($AC_n$) | Market Value ($MV$) | Unrealized PnL |
| :--- | :---: | :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **Tranche 1 (Entry)** | ₹2,800.00 | Initial Entry at 200 SMA | $B_1 = S$ | ₹28,000.00 | ₹28,000.00 | 10.0000 | ₹2,800.00 | ₹28,000.00 | ₹0.00 |
| **Market Dip (-10%)** | ₹2,520.00 | Adverse Pullback Evaluated | $L_{\text{actual}} = \max(0, 28000 - 25200)$ | — | ₹28,000.00 | 10.0000 | ₹2,800.00 | ₹25,200.00 | -₹2,800.00 |
| **Tranche 2 (Averaging)**| ₹2,520.00 | Adaptive Tranche Deployed | $B_2 = 2.0 \times ₹2,800$ | ₹5,600.00 | ₹33,600.00 | 12.2222 | **₹2,749.09** | ₹30,800.00 | -₹2,800.00 |
| **Target Rebound (+9.5%)**| ₹2,760.00 | Crosses $AC$ (₹2,749.09) | Price $> AC \implies \text{Liquidate}$ | — | ₹33,600.00 | 12.2222 | ₹2,749.09 | ₹33,733.27 | **+₹133.27** |
| **Universal Reset** | ₹2,760.00 | 100% Cash Liquidated | Full Capital Returned to Bankroll | +₹33,733.27 | **Reset $T \leftarrow 0$** | **0.0000** | **Reset to S** | **Cash: ₹33,733.27**| **Net Gain: +₹133.27** |

*Key Takeaway*: By investing an incremental ₹5,600.00, the breakeven cost collapsed from ₹2,800.00 to ₹2,749.09. A minor rebound of $+9.5\%$ from the low completely recovered the position with a net profit.

---

### 12.2 Case Study 2: Swing Trade Execution with TRM V3 Dual-Path (Apple Inc. / NASDAQ)

##### A. Setup Parameters
- **Instrument**: Apple Inc. (`AAPL`).
- **Initial Trade**: Purchase $Q_0 = 10 \text{ shares}$ at $P_{\text{entry}} = \$200.00 \implies I = \$2,000.00$.
- **Model Parameters**: Profit Target $p = 10\%$, Stop Loss $l = 5\%$, Retention Factor $r = 20\%$, Recovery Factor $f = 20\%$, Base Allocation $A_{\text{base}} = \$1,000.00$.

##### B. Scenario A: Profit Path Execution (Runner Retention)

| Step / Ledger Item | Governing Mathematical Formula | Computed Numerical Output | Provenance | Financial Meaning |
| :--- | :--- | :---: | :---: | :--- |
| **Exit Target Trigger** | $P_{\text{current}} = \$200.00 \times (1 + 0.10)$ | **\$220.00** | `[CALCULATED]` | $+10\%$ take-profit limit hit |
| **Gross Profit Realized** | $P_{\text{amount}} = \$2,000.00 \times 10\%$ | **\$200.00** | `[CALCULATED]` | Total dollar gain generated |
| **Retained Profit Value** | $R = \$200.00 \times 20\%$ | **\$40.00** | `[CALCULATED]` | Profit allocated to runner equity |
| **Retained Free Shares** | $Q_{\text{retain}} = \frac{\$40.00}{\$220.00}$ | **0.181818 shares** | `[CALCULATED]` | **The Current Price Law**: Converted at $\$220$ |
| **Shares Sold to Cash** | $Q_{\text{sell}} = 10.000000 - 0.181818$ | **9.818182 shares** | `[CALCULATED]` | Liquidated on open market |
| **Cash Principal Returned** | Initial Capital $I$ | **\$2,000.00** | `[CALCULATED]` | 100% capital preservation |
| **Net Cash Profit Realized**| $\$200.00 \times (1 - 0.20)$ | **\$160.00** | `[CALCULATED]` | Liquid cash credited to account |
| **Total Liquid Cash Returned**| $\$2,000.00 + \$160.00$ | **\$2,160.00** | `[CALCULATED]` | $9.818182 \times \$220.00 = \$2,160.00$ |
| **Perpetual Runner Equity**| $0.181818 \times \$220.00$ | **\$40.00** | `[CALCULATED]` | Zero-cost basis shares vaulted |
| **Cycle Balance Reconciliation**| $\$2,160.00 \text{ (Cash)} + \$40.00 \text{ (Vault)}$ | **\$2,200.00** | `[CALCULATED]` | **100% Reconciled ($10 \times \$220.00$)** |
| **Universal Reset** | State reset to Base $S$ | **\$1,000.00** | `[CALCULATED]` | Next trade allocation resets to base |

##### C. Scenario B: Loss Path Execution (Linear Loss Surcharge)

| Step / Ledger Item | Governing Mathematical Formula | Computed Numerical Output | Provenance | Financial Meaning |
| :--- | :--- | :---: | :---: | :--- |
| **Structural Stop Trigger** | $P_{\text{current}} = \$200.00 \times (1 - 0.05)$ | **\$190.00** | `[CALCULATED]` | Hard stop hit; position terminated |
| **Liquidated Cash Preserved**| $10 \times \$190.00$ | **\$1,900.00** | `[CALCULATED]` | $95\%$ principal preserved in cash |
| **Realized Loss Amount** | $L_{\text{amount}} = \$2,000.00 - \$1,900.00$ | **\$100.00** | `[CALCULATED]` | Booked loss on execution |
| **Recovery Surcharge** | $R_{\text{loss}} = \$100.00 \times 20\%$ | **\$20.00** | `[CALCULATED]` | Calibrated linear surcharge |
| **Base Reversal Tranche** | Standard Base $A_{\text{base}}$ | **\$1,000.00** | `[USER-PROVIDED]` | Baseline setup allocation |
| **Next Trade Allocation** | $A_{\text{next}} = \$1,000.00 + \$20.00$ | **\$1,020.00** | `[CALCULATED]` | Deployed on next qualified setup |
| **Recovery Target Validation**| $10\% \text{ gain on } \$1,020.00$ | **+\$102.00** | `[CALCULATED]` | Easily recaptures $\$20$ surcharge |
| **Bankroll Cap Guardrail** | Cap check: $\$1,020 \le 0.15 \times W$ | **PASS** | `[CALCULATED]` | Well within portfolio risk limits |

---

# 13. Validation, Backtesting & Deployment Readiness

### 13.1 Required Validation Order
1. **Unit validation:** verify every equation independently.
2. **Ledger reconciliation:** ensure cash + position + realized/harvested values reconcile.
3. **Scenario tests:** trend, pullback, gap, circuit, split/bonus, dividend, illiquidity, stale data, structural break.
4. **Historical backtest:** same instruments, timestamps, costs and position limits for every TRM version.
5. **Walk-forward test:** re-fit/choose parameters only on the training window; evaluate on the next unseen window.
6. **Out-of-sample holdout:** do not tune on the final test segment.
7. **Stress test:** prolonged decline, repeated gaps, no-fill orders, spread blowout, missing data, multi-stock correlation.
8. **Paper-trading/shadow mode:** verify live feed handling before live routing.

### 13.2 Cross-Version Benchmark Metrics
Compare Original/V1/V2/V3/V4/V5/V6/V7/V8 using the same definition of:
- cumulative return,
- maximum drawdown,
- downside deviation,
- realized win/loss distribution,
- profit factor,
- recovery time,
- capital utilization,
- turnover,
- transaction-cost drag,
- tail loss,
- concentration,
- number of forced exits,
- number of blocked/WAIT decisions,
- slippage and fill rate.

Do not select parameters because they maximize only one metric. The comparison must preserve the user's chosen risk constraints.

### 13.3 Avoiding Look-Ahead Bias
- Only use data available at the timestamp of the simulated decision.
- Do not use future high/low, revised fundamentals unavailable at the time, or future corporate-action information.
- External AI/agent signals in historical tests must be represented by timestamped historical outputs, not today's retrospective analysis.

### 13.4 Deployment Readiness Gate
V8 is not live-ready merely because it compiles. Minimum readiness requires:
```text
All unit tests PASS
        AND
Ledger reconciliation PASS
        AND
Walk-forward test completed
        AND
Out-of-sample holdout completed
        AND
Cost/slippage model enabled
        AND
Risk limits enforced in code
        AND
Paper/shadow trading completed
```

If any item fails, output `NOT_PRODUCTION_READY`.

### 13.5 Configuration Governance
The following MUST be externally configurable and versioned:
- trigger thresholds,
- protected principal policy,
- wealth-floor drawdown,
- maximum recovery rounds,
- risk budget,
- single-stock cap,
- transaction-cost assumptions,
- slippage assumptions,
- cooldown/hysteresis,
- liquidity thresholds,
- indicator windows.

Never silently change a parameter after a backtest without incrementing the strategy/configuration version.


# 14. Mandatory Regulatory Compliance & Risk Disclaimer

> **⚠️ STATUTORY REGULATORY COMPLIANCE & RISK DISCLAIMER**:
> 
> This document, quantitative algorithm specification, mathematical framework, and decision-support architecture are strictly for **educational, academic research, and decision-support purposes only** and do not constitute registered investment, financial, tax, or trading advice under the Securities and Exchange Board of India (SEBI) Regulations, U.S. Securities and Exchange Commission (SEC) guidelines, or any applicable statutory jurisdiction.
> 
> Equities and financial securities carry inherent market risk, including the possible loss of principal capital. Past performance, backtested simulations, mathematical derivations, and probabilistic win rates do not guarantee future market returns. 
> 
> Averaging down and recursive capital-allocation strategies (including the Teja Recovery Model in all its editions) carry severe tail-risk if applied to declining, leveraged, illiquid, or fundamentally impaired assets. Unmanaged position accumulation without adherence to structural stop-losses, maximum round limits ($N \le 5$), and single-stock capital caps ($\le 15\%$) can lead to catastrophic capital exhaustion. 
> 
> The user retains sole, absolute, and exclusive responsibility for all trading actions, order routing, capital allocation, and risk management decisions.
