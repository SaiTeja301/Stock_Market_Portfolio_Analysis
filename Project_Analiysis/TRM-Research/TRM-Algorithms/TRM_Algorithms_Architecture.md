# Teja Recovery Model (TRM) Algorithmic Engine Documentation
## Multi-Model Quantitative Architecture: TRM Classic, V1, V2, and V3

### 1. Architectural Overview
This directory houses the formal specification and reference implementations for the **Teja Recovery Model (TRM)** algorithmic suite, integrated into the Spring Boot microservice [`nse_market_gateway_service`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service).

The system operates across two interlocking engines:
- **Subsystem A (Trend-Regression-Momentum)**: Evaluates macro trend health, calculates 100-period OLS Linear Regression Channel ($\pm 2.0\sigma$) value bounds, and filters for momentum exhaustion (**WHEN to trade**).
- **Subsystem B (Teja Recovery Model)**: Governs capital tranches, share allocations, break-even shift, and risk-adjusted recovery targets (**HOW MUCH capital to deploy**).

---

### 2. Summary of All 4 Algorithmic Editions

| Model | Formal Name | Governing Recurrence | Primary Purpose | Capital Scaling | Production Suitability |
|---|---|---|---|---|---|
| **TRM Classic** | Recursive Cumulative Allocation | $B_1 = S, B_{n+1} = 2 T_n$ | Pure recursive recovery | $\mathcal{O}(3^n)$ Exponential | Theoretical / Foundation |
| **TRM V1** | Percentage-Triggered Model | $P_n = P_0(1-L)^n, B_{n+1} = 2 T_n$ | Downside equity averaging | $\mathcal{O}(3^n)$ Exponential | Fragile (Over-allocates) |
| **TRM V2** | Actual-Portfolio-Loss Adaptive | $B_{next} = m \cdot \max(0, T - MV)$ | True-loss dynamic averaging | Sub-geometric adaptive | **Approved** (99.68% savings vs V1) |
| **TRM V3** | Asymmetric Dual-Path Model | Path A: $Q_{retain} = R / P_{curr}$<br/>Path B: $A_{next} = A_{base} + (L \cdot f)$ | Zero-cost runner retention + calibrated loss reversal | Controlled linear surcharge | **Approved** (Standard Production) |

---

### 3. REST API Endpoints in `nse-market-gateway-service`
Base URI: `http://localhost:8081/api/v1/trm`

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/v1/trm/classic` | Evaluates TRM Classic recursive tripling sequence for input base capital. |
| `GET` | `/api/v1/trm/classic/{symbol}` | Evaluates TRM Classic for a specific NSE equity symbol using live quote. |
| `POST` | `/api/v1/trm/v1` | Evaluates TRM V1 percentage-triggered drop model. |
| `GET` | `/api/v1/trm/v1/{symbol}` | Evaluates TRM V1 for a specific NSE equity symbol. |
| `POST` | `/api/v1/trm/v2` | Evaluates TRM V2 actual-loss adaptive model (calculates capital savings vs V1). |
| `GET` | `/api/v1/trm/v2/{symbol}` | Evaluates TRM V2 for a specific NSE equity symbol using live quote. |
| `POST` | `/api/v1/trm/v3` | Evaluates TRM V3 dual-path model (Path A profit runner & Path B loss recovery). |
| `GET` | `/api/v1/trm/v3/{symbol}` | Evaluates TRM V3 for a specific NSE equity symbol. |
| `POST` | `/api/v1/trm/evaluate-all` | Evaluates all 4 models simultaneously for comprehensive portfolio analysis. |
| `GET` | `/api/v1/trm/evaluate/{symbol}` | Live 360° evaluation fetching real-time NSE quote, 100-period OLS channel, and all 4 TRM models. |

---

### 4. Mathematical Rules & Invariants

#### 4.1 The Current Price Law (TRM V3)
Retained profit value ($R$) MUST be converted to runner shares by dividing by the **current market price** ($P_{\text{current}}$), NEVER the entry price:
$$Q_{\text{retain}} = \frac{R}{P_{\text{current}}} = \frac{I \cdot p \cdot r}{P_{\text{entry}}(1 + p)}$$
This guarantees that **100% of initial principal $I$ is extracted in cash**, along with $(1-r)$ net cash profit, leaving the retained runner shares at an **effective cost basis of exactly $0.00$**.

#### 4.2 The Universal Reset Invariant
Upon reaching recovery target or profit harvest, active capital allocation unconditionally resets to Base investment $S$. Tranche escalation sizes are NEVER compounded into the next independent trade cycle.

#### 4.3 Cash Equity Risk Guardrails
- 100% Unleverage spot equity only (no F&O, margin, or shorting).
- Maximum 5 averaging rounds ($N \le 5$).
- Strict 15% single-stock bankroll cap ($T \le 0.15 \times W$).
