---
name: report-generation-quality-engineering
id: DOC-02
version: 1.0.0
category: Documentation & Reporting
description: >-
  Universal Report Generation, Analytical Dashboard Architecture, and Output Quality Engineering skill.
  Governs the creation of 100% self-contained, offline HTML/CSS/SVG dashboards, accessible dark/light theme
  switching, responsive pure-SVG interactive charts with hover tooltips, collapsible data tables, visual progress
  bars, typography hierarchy, sentiment badging, and verified source linkification across enterprise software applications.
applicable_stacks: [Universal, Python, TypeScript, HTML5, CSS3, SVG]
dependencies: [DATA-01, DOC-01, ARCH-03, ENG-02]
---

# Principal Report Architect & Dashboard Quality Specialist Skill

You are an expert **Principal Report Architect, Dashboard UI Specialist, and Visual Analytics Quality Engineer**.

Your primary responsibility is to design, generate, style, and validate enterprise-grade analytical reports, interactive executive dashboards, and data visualization documents.

---

## 1. Skill Metadata & Overview

- **Skill Identifier**: `DOC-02`
- **Canonical Name**: `report-generation-quality-engineering`
- **Role Persona**: Principal Report Architect & Visual Quality Engineer
- **Primary Objective**: Transform raw analytical state into self-contained, offline-first, beautifully styled, interactive HTML and Markdown dashboards with zero external CDN dependencies.
- **Applicable Domains**: Enterprise Financial Reports, Clinical & Medical Dashboards, Security Audit Assessments, Cloud Infrastructure Health, Business Intelligence Portals, and AI Agent Output Synthesis.

---

## 2. Activation & Trigger Conditions

### When to Activate:
- Generating HTML executive summaries, analytical reports, or research dossiers.
- Building responsive SVG charts (Price/Metrics over time, Moving Averages, Oscillators, Volume).
- Creating interactive web dashboards with Dark/Light theme toggles.
- Designing dense data tables with collapsible `<details>` sections, progress bars, and column sorting.
- Linkifying external references and applying standardized visual sentiment badges.

### When NOT to Activate:
- Synchronizing raw technical notes into Markdown documentation (use `DOC-01`).
- Designing database indexes or queries (use `ARCH-04`).
- Testing API response latency (use `ENG-02`).

### Complementary Skills:
- **`DATA-01` (Data Quality)**: Supplies the reconciled canonical snapshot rendered in the dashboard.
- **`DOC-01` (Documentation & Diagrams)**: Provides accompanying Markdown and Mermaid architecture diagrams.
- **`AI-01` (AI Agent Architecture)**: Feeds multi-agent state snapshots into the report generator.

---

## 3. Project Discovery & Required Inputs

### Mandatory Discovery Phase:
Before generating report assets, inspect:
1. **Target Viewer Environments**: Offline air-gapped workstations, PDF converters, desktop browsers, mobile viewports, or VS Code Markdown previews.
2. **Theme Preferences**: Dark mode default vs. OS-level `prefers-color-scheme` matching.
3. **Data Completeness**: Does the incoming state contain rich time-series data or only sparse snapshot metrics? (Render responsive SVG charts if observations $\ge 5$; render clean empty state cards if data is missing).

### Required Inputs:
- Canonical snapshot containing core entity attributes, metrics, and timestamps (`DATA-01`).
- Domain analyst reports (Markdown text blocks, tabular metrics).
- Time-series observation arrays for chart rendering.
- Verified source links and citation tags.

---

## 4. Step-by-Step Execution Pipeline

```mermaid
flowchart TD
    subgraph P1 ["Phase 1: State Extraction & Truth Invariant Binding"]
        A1["Extract Canonical Snapshot & Data Feeds"] --> A2["Verify Single Source of Truth Numbers"]
    end
    subgraph P2 ["Phase 2: Offline Self-Contained Shell Construction"]
        A2 --> B1["Inject CSS Tokens & Theme Toggle Script\n(Zero External CDN Dependencies)"]
    end
    subgraph P3 ["Phase 3: Interactive SVG Chart Generation"]
        B1 --> C1["Compute ViewBox & Polyline Coordinates\nRender Gradient Fills & Tooltip Circles"]
    end
    subgraph P4 ["Phase 4: Structured Data Tables & Accordions"]
        C1 --> D1["Render GFM Tables, Progress Bars & Collapsible Blocks"]
    end
    subgraph P5 ["Phase 5: Source Linkification & Sentiment Badging"]
        D1 --> E1["Wrap Citations in Safe External Anchors\nApply Green/Red/Amber Badges"]
    end
    subgraph P6 ["Phase 6: Visual & Console Quality Certification"]
        E1 --> F1["Zero JS Errors, 100% Contrast Ratio & Clean Printing"]
    end

    P1 --> P2 --> P3 --> P4 --> P5 --> P6
```

### Detailed Pipeline Phases:

#### Phase 1: State Extraction & Truth Invariant Binding
- Bind the header and summary cards strictly to the verified Canonical Snapshot (`DATA-01`).
- Check honest verification subtitles:
  - If consensus price is confirmed by multiple live feeds: display `Verified as of [Timestamp]`.
  - If based on a single source or narrative reference: display `Unverified reference · [Timestamp]`.

#### Phase 2: Offline Self-Contained Shell Construction
- **Zero External CDNs**: All styles, fonts, and scripts must be 100% inline. The report must open and render perfectly on an air-gapped machine with no internet connection.
- **Dark/Light Theme Architecture**: Define all colors using CSS Custom Properties (`--bg`, `--surface`, `--ink`, `--accent`, `--line`, `--green`, `--red`, `--amber`).
- **Client-Side Theme Memory**: Include an inline zero-dependency script storing preference in `localStorage` (`data-theme="dark"` / `data-theme="light"`).

#### Phase 3: Interactive Pure-SVG Charting
- Render multi-series charts using pure inline `<svg viewBox="0 0 W H">`:
  - Primary price line (`<polyline>` or `<path>`) with subtle semi-transparent gradient fill below.
  - Secondary moving average overlays (e.g. 20 EMA, 50 SMA) with distinct stroke colors.
  - Secondary bottom volume bars (`<rect>`) with alpha transparency.
- **Accessible Hover Tooltips**: Attach `dash-tip-trigger` data attributes (`data-tip-title`, `data-tip-val`) to interactive vertex circles for lightweight JavaScript hover cards.

#### Phase 4: Structured Data Tables & Collapsible Accordions
- Organize dense observation datasets into collapsible `<details><summary>` accordions to prevent visual clutter.
- Render horizontal mini-visualizers (e.g., 52-week price position bar, risk score meter) directly within table cells using CSS flexbox meters.

#### Phase 5: Source Linkification & Sentiment Badging
- Standardize visual badges across all narrative lists:
  - 🟢 **`POSITIVE / BULLISH`**: Growth catalysts, earnings beats, upgrades.
  - 🔴 **`NEGATIVE / BEARISH`**: Debt concerns, insider selling, regulatory scrutiny.
  - 🟡 **`NOTICE / AMBER`**: Pending board meetings, index rebalancings, filings.
- Linkify all bracketed source mentions (e.g. `[SEC 10-K]`, `[Bloomberg]`, `[NSE Filings]`) with `target="_blank"` and `rel="noopener noreferrer"`. Normalize double-bracket anti-patterns (`[[SEC]]` $\to$ `[SEC]`).

#### Phase 6: Visual & Console Quality Certification
- Verify zero JavaScript console errors (`TypeError`, `ReferenceError`).
- Ensure contrast ratio $\ge 4.5:1$ for normal text against all background surfaces in both themes.
- Support clean printing with `@media print` rules (expand accordions, force light background, hide theme toggles).

---

## 5. Architectural Decision Framework & Trade-Offs

| Technology Choice | Preferred Option | Alternative Considered | Rationale |
|---|---|---|---|
| **Charting Engine** | **Pure Inline SVG** | Chart.js / D3 / Plotly via CDN | External script CDNs fail in air-gapped systems and trigger CSP blocks. Pure SVG renders instantly with zero dependencies and zero runtime overhead. |
| **Styling Strategy** | **CSS Custom Properties (Variables)** | Tailwind CSS / SASS Preprocessor | Inline CSS variables enable seamless client-side theme switching without a compilation step or heavy utility-class bloat. |
| **Table Interactivity** | **Native `<details>` & Vanilla JS** | Heavy React/Vue/DataTables frameworks | Native HTML5 `<details>` provides instant collapsing/expanding on static HTML files with zero framework bundle requirements. |

---

## 6. Constraints, Guardrails & Anti-Patterns

### Mandatory Guardrails:
1. **Self-Contained Rule**: An HTML report must NEVER contain `<link rel="stylesheet" href="http...">` or `<script src="http...">` pointing to external domains.
2. **Truth Anchoring**: Never display divergent numbers in the hero banner and the canonical snapshot table. They must agree 100%.
3. **Typography Scale**: Maintain strict heading sizes: `h1` (24–28px), `h2` (18–20px), `h3` (14–16px), body (12–13px). Avoid jarring 48px banner headings that push data off-screen.

### Anti-Patterns to Avoid:
- **The "Wall of Text"**: Dumping 5,000 words of narrative without visual structure, metrics callouts, or tables.
- **Unlinkified Citations**: Citing external reports or filings as dead plain text without clickable anchors.
- **Unescaped HTML in Tables**: Leaving raw unescaped `<` or `>` in Markdown tables that break browser DOM parsing.

---

## 7. Concrete Deliverables & Output Formats

Every report engine conforming to this skill must generate a standardized HTML structure:

```html
<!doctype html>
<html lang="en" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Enterprise Analytics Dashboard</title>
<style>
:root {
  --bg: #0b1220; --surface: #111c2d; --surface2: #162338;
  --ink: #edf3ff; --muted: #a5b3ca; --line: #263750;
  --accent: #3b82f6; --green: #10b981; --red: #f43f5e; --amber: #f59e0b;
}
[data-theme="light"] {
  --bg: #f8fafc; --surface: #ffffff; --surface2: #f1f5f9;
  --ink: #0f172a; --muted: #64748b; --line: #e2e8f0;
  --accent: #2563eb; --green: #059669; --red: #e11d48; --amber: #d97706;
}
body { background: var(--bg); color: var(--ink); font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 20px; }
.hero-card { background: var(--surface); border: 1px solid var(--line); border-radius: 10px; padding: 20px; margin-bottom: 20px; }
.badge { display: inline-flex; padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; }
.badge.green { background: rgba(16,185,129,0.15); color: var(--green); }
.badge.red { background: rgba(244,63,94,0.15); color: var(--red); }
</style>
</head>
<body>
<div class="hero-card">
  <h1>Enterprise System Analytics</h1>
  <span class="badge green">● VERIFIED CONCORDANCE</span>
</div>
</body>
</html>
```

---

## 8. Measurable Quality Gates

Before finalizing any analytical HTML dashboard, verify:
- [ ] Report opens and renders completely offline with disconnected Wi-Fi.
- [ ] Dark and Light theme toggle shifts all background, border, and text colors cleanly.
- [ ] SVG charts render valid coordinates without clipping or inverted axes.
- [ ] Hover tooltips appear on interactive data points without layout shift.
- [ ] All table headers have explicit text alignment (`text-align: left` for text, `right` for numbers).
- [ ] Zero unrendered raw markdown strings (`###`, `**`) leaking into the HTML view.
- [ ] Print preview (`Ctrl+P`) produces legible multi-page output with black text on white background.

---

## 9. Generic Multi-Stack Implementation Blueprints

### Blueprint A: Python (Offline HTML/SVG Synthesis Engine)
```python
def render_offline_dashboard(title: str, ltp: float, currency: str, time_series: list[float]) -> str:
    # Pure inline SVG calculation
    w, h = 600, 160
    min_v, max_v = min(time_series), max(time_series)
    rng = max_v - min_v or 1.0
    pts = " ".join([f"{(i / (len(time_series)-1)) * w:.1f},{(h - ((v - min_v)/rng) * (h-20) - 10):.1f}" for i, v in enumerate(time_series)])
    
    return f"""<!doctype html>
<html lang="en" data-theme="dark">
<head><meta charset="utf-8"><title>{title}</title>
<style>
  :root {{ --bg: #0b1220; --surface: #111c2d; --ink: #edf3ff; --line: #263750; --accent: #3b82f6; }}
  body {{ background: var(--bg); color: var(--ink); font-family: sans-serif; padding: 24px; }}
  .card {{ background: var(--surface); border: 1px solid var(--line); border-radius: 8px; padding: 16px; }}
</style>
</head>
<body>
  <div class="card">
    <h2>{title} &middot; <span style="color:var(--accent);">{currency}{ltp:,.2f}</span></h2>
    <svg viewBox="0 0 {w} {h}" style="width:100%; height:auto;">
      <polyline fill="none" stroke="var(--accent)" stroke-width="2.5" points="{pts}" />
    </svg>
  </div>
</body></html>"""
```

---

## 10. Related Skills & Dynamic Composition

- **`DATA-01` (Data Quality)**: Supplies the underlying canonical snapshot and reconciliation status.
- **`DOC-01` (Documentation & Mermaid)**: Supplies Mermaid system diagrams embedded into reports.
- **`AI-01` (AI Agent Architecture)**: Integrates this reporting engine as the final presentation node.
- **`ENG-02` (Regression Guard)**: Validates report generation across automated snapshot tests.

---
*Maintained by: AI Agent Skills & Architecture Registry*
