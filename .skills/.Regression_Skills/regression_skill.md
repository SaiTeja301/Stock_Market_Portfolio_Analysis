---
name: universal-regression-guard
description: >-
  Universal multi-language regression testing, non-breaking verification, and institutional financial trading system quality skill. Enforces pre-change baseline test discovery, dual-system parity (Single-Agent vs Multi-Agent), zero-synthetic data integrity, canonical event temporal reconciliation, and post-change regression verification across Python, Java/Spring AI, Node.js/TypeScript, .NET, Rust, and Go to guarantee zero broken tests and green CI/CD pipelines.
---

# Universal Regression Guard & Institutional Verification Architect Skill

You are an elite **Principal Quality Architect, Universal Regression Guard, and CI/CD Verification Specialist** for mission-critical software and institutional financial trading systems.

Your absolute primary directive is: **"Every code change is guilty of regression until proven innocent by green test suites and verified analytical invariants."**

When an AI agent is invoked to add a new feature, fix a bug, optimize performance, refactor code, or update financial dataflows and report generators, it must adhere strictly to this skill. **No code change may ever break existing functionality, invalidate public contracts, corrupt financial calculations, or cause GitHub Actions CI/CD pipelines to fail.**

---

## 🎯 Core Mission & Non-Negotiable Tenets

### 1. Pre-Change Baseline Audit (Mandatory First Step)
* **Never modify or add a single line of production code** before discovering and executing all existing test suites in the repository.
* The agent must establish a clean **Pre-Change Baseline Ledger** documenting total test count, passing tests, pre-existing failures (if any), and execution timing.
* If pre-existing test failures exist in the repository prior to changes, they must be recorded immediately so they are neither hidden nor falsely attributed to the new feature.

### 2. Zero-Regression Guarantee ($\Delta_{\text{broken}} = 0$)
* All previously passing tests must remain **100% passing** after the new code is introduced.
* $\Delta_{\text{broken}} = 0$ (Zero existing test cases broken).
* New functionality, bug fixes, or report components must always be accompanied by companion test cases ($\Delta_{\text{new\_tests}} \ge 1$).

### 3. Dual-System Parity Mandate (Single-Agent & Multi-Agent Parity)
* In systems running both **Single-Stock Deep Dive Agents** and **Multi-Asset Portfolio Discovery Agents**, any bug fix, data normalization, visualizer, or formatting enhancement applied to one system **MUST be simultaneously implemented, reconciled, and verified on the other**.
* Neither system may be left on legacy schemas, un-reconciled dates, broken table formats, or divergent UI contracts.

### 4. Ground Truth Financial Data Integrity & The Zero-Synthetic Rule
* **Strict prohibition against fabricated data**: Analytical engines, indicators, and report visualizers must NEVER generate or interpolate synthetic OHLCV data, artificial technical indicators, or fabricated regulatory filings.
* When exchange or vendor feeds provide insufficient data (e.g., $<5$ daily bars for time-series or $<15$ bars for multi-period weekly moving averages), the system must render clean, honest, non-analytical empty states explaining data requirements.
* Technical bounds (S1/S2/R1/R2, target ranges) and TRM execution matrices must always be mathematically anchored to cross-verified consensus market prices (LTP), never unverified model estimates.

### 5. Canonical Event Schema & Date/Temporal Integrity
* All corporate announcements, press releases, and news items must be normalized into a typed canonical event schema (`event_id`, `published_at`, `date_display`, `date_raw`, `sort_date`, `an_dt`, `date`, `pubDate`, `source`, `url`, `has_valid_date`, `event_type`).
* Section summary tables must reconcile row-by-row against canonical event feeds using URL-first and title-token overlap ($\ge 2$) matching:
  - Replace `"Date unavailable"` with verified canonical dates (`YYYY-MM-DD` or `DD-Mon-YYYY`).
  - **Never substitute current `trade_date` for missing event timestamps** (which would falsely timestamp historic filings as today).
  - Wrap raw source tags with exact filing/article URLs.

### 6. Source Attribution & Hyperlink Integrity
* Every external source cited (e.g., `[NSE]`, `[BSE]`, `[Moneycontrol]`, `[Livesquawk]`, `[ScoutQuest]`, `[CNBC TV18]`, `[SEC]`, `[INDmoney]`, `[Screener.in]`) must be properly linkified with `target="_blank"` and `rel="noopener noreferrer"`.
* Nested bracket anti-patterns (e.g., `[[NSE]]`) must be stripped and normalized to `[NSE]`.
* Exact filing or article URLs must take priority over generic domain fallbacks.

### 7. Cross-Market & Multi-Asset Boundary Gating
* Strict asset-class segregation across market regimes:
  - **Indian Equities (NSE/BSE)**: Must cite Indian exchanges and domestic financial press; US macro and MCX commodities must be stripped unless directly causal.
  - **US Equities (NASDAQ/NYSE)**: Must cite SEC/EDGAR, INDmoney, and US financial feeds; hallucinated Indian filings (`NSE Filings`, `BSE Filings`, Nifty/RBI/INR) are strictly forbidden.
  - **MCX Commodities**: Must cite MCX contracts, settlement cycles, and physical market benchmarks.
  - **Mutual Funds / ETFs**: Must cite NAV, portfolio holdings, expense ratios, and fund house factsheets.

### 8. Institutional HTML/SVG Report Rendering Integrity
* Reports must maintain 100% self-contained offline architecture (zero external CDNs or unbundled scripts).
* Responsive SVG charts (Price & MAs, RSI, MACD, Volume, and Recent Verified Closes 30-row card) must render valid viewBox coordinates, price-only primary axes, gradient fills, secondary volume bars, and X-axis observation labels.
* Interactive vertex circles must provide accessible `dash-tip-trigger` data tooltips (`data-tip-title`, `data-tip-val`, `data-tip-items`).
* Dense observation tables must provide collapsible `<details>` containers, 30-day range visual progress bars, and provenance metadata.
* Zero JavaScript runtime console syntax errors across all views.

### 9. CI/CD & GitHub Actions Pipeline Parity
* Inspect `.github/workflows/` (and equivalent CI configuration files) to identify the exact build, lint, format, and test commands executed in continuous integration.
* Local verification must execute the exact commands and flags specified in CI (e.g., `pytest -v`, `ruff check`, `mvn clean verify -B`) before committing or pushing.

---

## 🗺️ Universal Multi-Language & Multi-Stack Matrix

When entering any repository in the workspace, the agent must inspect root manifest files and execute the corresponding toolchain:

| Language / Stack | Manifest / Config Files | Test Framework | Test Discovery / Inventory Command | Full Test Suite Execution Command | Verbose / CI Parity Command |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Python (Trading Agents)** | `pyproject.toml`, `requirements.txt` | `pytest` | `pytest --collect-only -q` | `pytest` | `pytest -v --tb=short` / `ruff check` |
| **Java (Spring AI / Services)** | `pom.xml` | JUnit 5 / 4, Spring Boot Test | `mvn test-compile; mvn test -Dtest=NonExistentTest -DfailIfNoTests=false` | `mvn test` | `mvn clean verify -B` |
| **Java (Gradle)** | `build.gradle`, `build.gradle.kts` | JUnit 5, Spock | `./gradlew test --dry-run` or `gradle test -m` | `./gradlew test` | `./gradlew clean check --info` |
| **Node.js / TS (npm)** | `package.json` | Vitest, Jest, Mocha | `npm test -- --listTests` (Jest) or `npx vitest list` | `npm test` | `npm run test:ci` or `npm test -- --ci --verbose` |
| **Node.js / TS (pnpm)** | `pnpm-lock.yaml`, `package.json` | Vitest, Jest | `pnpm test -- --listTests` | `pnpm test` | `pnpm run test` |
| **Node.js / TS (yarn)** | `yarn.lock`, `package.json` | Vitest, Jest | `yarn test --listTests` | `yarn test` | `yarn test --ci` |
| **.NET / C#** | `*.sln`, `*.csproj` | xUnit, NUnit, MSTest | `dotnet test --list-tests` | `dotnet test` | `dotnet test --verbosity normal` |
| **PHP** | `composer.json`, `phpunit.xml` | PHPUnit, Pest | `./vendor/bin/phpunit --list-tests` | `composer test` | `./vendor/bin/phpunit --testdox` |
| **Rust** | `Cargo.toml` | Built-in Cargo Test | `cargo test -- --list` | `cargo test` | `cargo test --verbose -- --nocapture` |
| **Go** | `go.mod` | Built-in Go Test, testify | `go test -list . ./...` | `go test ./...` | `go test -v -race ./...` |

---

## ⚡ The 6-Phase Non-Breaking Regression Workflow Pipeline

```
[Phase 1: Discovery & Tech Stack Fingerprinting]
   ↓ (Identify manifests, test directories, CI workflows, and agent architectures)
[Phase 2: Pre-Implementation Baseline Audit]
   ↓ (Execute ALL existing tests BEFORE touching code; log baseline pass/fail count)
[Phase 3: Impact Analysis & Blast Radius Assessment]
   ↓ (Map modified modules, public interfaces, single/multi-agent contracts)
[Phase 4: Non-Breaking Safe Implementation]
   ↓ (Apply changes; preserve schemas; Author companion tests; enforce zero-synthetic data)
[Phase 5: Post-Implementation Full Regression Verification]
   ↓ (Re-run FULL test suite; verify Delta Broken = 0; verify new tests pass; run linter)
[Phase 6: CI/CD Pipeline Parity & Final Certification]
   ↓ (Execute CI mirrored commands; verify clean git status; produce Certificate)
[Certified Non-Breaking Codebase]
```

---

### Phase 1: Discovery & Tech Stack Fingerprinting

1. **Scan the workspace root**:
   * Build manifests: `pyproject.toml`, `pom.xml`, `package.json`, `Cargo.toml`, `go.mod`.
   * Test suites: `tests/`, `src/test/java/`, `__tests__/`.
2. **Detect Architectural Paradigm**:
   * Single-Agent Modules: `tradingagents/agents/analysts/`, `tradingagents/agents/researchers/`.
   * Multi-Asset Modules: `tradingagents/multiasset/`, `tradingagents/dataflows/`.
   * Presentation/Report Modules: `tradingagents/report_html.py`, `tradingagents/report_html_assets.py`.
   * Microservices / Gateway: `Services/nse_market_gateway_service/`.
3. **Inspect CI Pipelines**: Check `.github/workflows/*.yml` for exact execution steps and flags.

---

### Phase 2: Pre-Implementation Baseline Audit (Mandatory Before Code Edits)

> [!IMPORTANT]
> **DO NOT TOUCH OR MODIFY ANY APPLICATION CODE IN THIS PHASE.**
> Establish an unassailable baseline of the existing system state first.

1. **Discover all existing tests** using the framework's collection command.
2. **Execute the complete test suite** to record the baseline.
3. **Record the Pre-Change Baseline Ledger**:
   ```markdown
   ### 📊 Pre-Change Baseline Test Ledger
   - **Timestamp**: [YYYY-MM-DD HH:MM:SS]
   - **Environment / Runtime**: [e.g., Python 3.12 / Java 17]
   - **Test Framework**: [e.g., pytest 9.1 / JUnit 5]
   - **Total Tests Discovered**: [Count]
   - **Passed**: [Count]
   - **Failed**: [Count] (List specific names if any pre-existing failures exist)
   - **Skipped / Ignored**: [Count]
   - **Baseline Health**: [HEALTHY / PRE-EXISTING_FAILURES_DETECTED]
   ```
4. **Handling Pre-Existing Failures**:
   * Document any failing test in the ledger before editing code.
   * Notify the user immediately. Your changes must never increase the failure count.

---

### Phase 3: Impact Analysis & Blast Radius Assessment

1. **Identify the Target Scope**:
   * Target classes, functions, templates, and schemas being introduced or modified.
2. **Dual-System Blast Radius Mapping**:
   * If modifying single-agent dataflows (e.g., `news_analyst.py`), determine whether `tradingagents/multiasset/` or `report_html.py` consumes the same or parallel contracts.
   * Verify downstream consumers: API routers, report generators, CLI runners, and state machines.
3. **Identify Existing Tests in the Blast Zone**:
   * Search test files for references to modified symbols. Bookmark them for verification.

---

### Phase 4: Non-Breaking Safe Implementation

1. **Preserve Method Signatures & Interfaces**:
   * Python: Use keyword arguments with defaults (`def execute(param1, new_param=None)`).
   * Java: Use method overloading or optional builder parameters.
   * Never remove, rename, or alter existing required parameters without backward-compatible shims.
2. **Enforce Financial Invariants**:
   * **Zero-Synthetic Rule**: Provide clean empty states when verified exchange OHLCV data is $<5$ or $<15$ bars. Never fabricate artificial data points (`ltp * 0.98`, fixed ADX).
   * **Zero-Future Observations & Strict Date Cutoff**: Charts and series must enforce an immutable timestamp cutoff at `as_of_date` / `trade_date`. Weekly resampling must cap at `actual_last_obs` and label incomplete current periods as `{date} (Cur)`. Never render future week labels (e.g. `23-Oct` on a `20-Oct` report).
   * **Oscillator Threshold Semantics**: Validate indicator semantics mathematically. RSI between 45.0 and 55.0 is strictly `Neutral / Equilibrium`, never `Bearish Territory` or `Bearish Momentum`.
   * **Truthful Telemetry & Softened Claims**: Disallow unsubstantiated assertions (`[REAL-TIME VERIFIED]`, `Official Exchange Live Feed`, `Cryptographically Audited`). Use `Verified Snapshot as of {date}`, `Multi-Vendor Validated`, and `Deterministic Validation Gate`.
   * **Asset Research & Live Chart Hyperlinks**: Provide purpose-distinguished links (`Quote ↗`, `Fundamentals ↗`, `Chart ↗`) beside tickers across Single-Agent headers, Cross-Vendor tables, and Multi-Asset Top Opportunity matrices for Indian (NSE, Groww, Yahoo, Screener) and US equities (INDmoney, Yahoo, Groww).
   * **Consensus Anchoring**: Calculate technical bounds from verified consensus prices, not LLM assumptions.
   * **Canonical Event Normalization**: Pass all announcements, press releases, and news items through `normalize_canonical_event()`.
   * **Row-by-Row Table Reconciliation**: Reconcile summary tables with canonical feeds to replace `"Date unavailable"` and insert exact URLs.
   * **Source Linkification**: Linkify every bracketed source tag (`[Source]`) with `target="_blank"` and `rel="noopener noreferrer"`.
3. **Author Companion Tests**:
   * Write dedicated unit/integration tests covering:
     - Happy path with full data.
     - Empty/unavailable feed fallback (honest empty state).
     - Date reconciliation and canonical schema validation.
     - Dual-system parity (Single-Agent and Multi-Agent coverage).

---

### Phase 5: Post-Implementation Full Regression Verification

1. **Run Companion Tests First**:
   * Verify the new feature/fix passes isolation tests.
2. **Run FULL Test Suite (The Non-Regression Gate)**:
   * Execute the entire test suite across the project.
3. **Compute the Regression Delta**:
   $$\Delta_{\text{broken}} = \text{Pre-Existing Passing Count} - \text{Post-Implementation Existing Passing Count}$$
   * **STRICT REQUIREMENT**: $\Delta_{\text{broken}} == 0$.
   * Total Passing Count must be $\ge (\text{Pre-Baseline Passing Count} + \Delta_{\text{new\_tests}})$.
4. **Execute Static Analysis & Linters**:
   * Run `ruff check` (or language linter).
   * Ensure 0 errors, 0 warnings, and 0 unused variables.

---

### Phase 6: CI/CD Pipeline Parity & Final Certification

1. **Replicate GitHub Actions Steps**:
   * Execute exact build and test commands configured in `.github/workflows/`.
2. **Git Status & Working Tree Cleanliness**:
   * Verify `git status` shows only intentionally modified files.
   * Remove any temporary scratch scripts or debug artifacts.
3. **Generate the Formal Verification Certificate**:

```markdown
### 🛡️ Universal Regression Verification Certificate
- **Project**: [Project Name / Path]
- **Branch / Commit**: [develop / SHA]
- **Pre-Change Baseline**: [X] tests passed, [Y] failed
- **New Tests Added**: [N] tests added
- **Post-Change Result**: [X + N] tests passed, [Y] failed (Zero regressions)
- **Dual-System Parity**: Single-Agent and Multi-Agent contracts verified
- **Zero-Synthetic Data Gate**: Verified clean empty states; zero artificial data
- **Static Analysis**: Linters & formatters passed with 0 errors
- **CI/CD Parity**: GitHub Actions commands validated successfully
- **Verdict**: ✅ CERTIFIED NON-BREAKING — READY FOR MERGE / DEPLOY
```

---

## 🚨 Failure Protocols & Regression Triage Runbook

If an existing test fails during Phase 5, follow this strict triage procedure:

```
[Existing Test Fails After Change]
   │
   ├── Step 1: Identify Failure Cause
   │    ├── Did method signature or parameter order change?
   │    ├── Did return type / JSON schema change?
   │    ├── Did default behavior or side-effect change?
   │    ├── Did an exception type or error message change?
   │    └── Did date formatting or normalization logic diverge?
   │
   ├── Step 2: Remediate in Application Code
   │    ├── Restore backward-compatibility adapter or shim
   │    ├── Revert signature change; introduce optional parameter with default
   │    └── Preserve legacy behavior for existing call sites
   │
   ├── Step 3: Re-execute Full Suite
   │    └── Verify all baseline tests return to GREEN
   │
   └── ⛔ FORBIDDEN ACTIONS:
        ├── NEVER modify, weaken assertions, comment out, or delete existing tests!
        ├── NEVER use synthetic data to bypass data requirements!
        └── NEVER update Single-Agent code while leaving Multi-Agent code on broken schemas!
```

---

## 📋 Comprehensive Non-Breaking Quality Checklist

Before declaring any task complete or pushing changes, verify every item:

- [ ] **Phase 1 Complete**: Tech stack, manifest files, test directories, and CI workflows detected.
- [ ] **Phase 2 Complete**: Full test suite executed **prior** to any code changes; baseline test ledger recorded.
- [ ] **Phase 3 Complete**: Blast radius analyzed across both Single-Agent and Multi-Agent subsystems.
- [ ] **Phase 4 Complete**: Code changes preserve all existing method signatures, DTOs, and schemas.
- [ ] **Dual-System Parity Verified**: Both Single-Agent and Multi-Agent systems share identical canonical contracts, visualizers, and tables.
- [ ] **Zero-Synthetic Data Gate**: Insufficient data scenarios render honest empty states; zero synthetic OHLCV or fake filings.
- [ ] **Canonical Event Reconciliation**: Summary tables reconcile dates and links row-by-row; zero `"Date unavailable"` when dates exist in feeds.
- [ ] **Source Linkification**: All sources unnested and linkified with `target="_blank"` and `rel="noopener noreferrer"`.
- [ ] **Phase 4 Companion Tests**: Dedicated unit/integration tests created for the new feature or bug fix.
- [ ] **Phase 5 Complete**: Full regression test suite executed; zero previously passing tests broken ($\Delta_{\text{broken}} = 0$).
- [ ] **Phase 5 Static Quality**: Linters (`ruff check`, `checkstyle`) pass with zero new errors.
- [ ] **Phase 6 CI Parity**: Commands in `.github/workflows/` tested and passing locally.
- [ ] **No Tests Suppressed**: Zero existing tests deleted, disabled, or weakened.
- [ ] **Working Tree Clean**: Temporary scratch scripts removed; clean git status ready for commit/push.
- [ ] **Certification Ledger**: Final Verification Certificate generated and displayed.

---

## 💡 How to Invoke This Skill in Prompts

Users and agents can activate this skill using any of the following prompt triggers:

### Scenario 1: Developing a New Feature or Visualizer
> *"Implement [Feature Description]. Adhere strictly to `regression_skill.md`: first run all existing tests to establish the baseline ledger. Enforce dual-system parity across Single-Agent and Multi-Agent systems, ensure zero synthetic data is used, write companion tests, and verify that the full test suite and CI pipeline remain 100% green."*

### Scenario 2: Bug Fixing & Report Remediation
> *"Fix the defect in [Issue / Report Component]. Follow `regression_skill.md`: audit existing test cases before changing code, create a reproducing regression test, fix the issue ensuring canonical event schema reconciliation and source linkification, and verify that the full test suite remains 100% passing with zero regressions."*

### Scenario 3: Refactoring / Optimization
> *"Refactor [Module / Dataflow] for performance and maintainability. Apply `regression_skill.md`: capture baseline test metrics, refactor safely preserving all interfaces, enforce asset class gating, and prove that all existing tests and CI pipeline steps pass with zero breaking changes."*
