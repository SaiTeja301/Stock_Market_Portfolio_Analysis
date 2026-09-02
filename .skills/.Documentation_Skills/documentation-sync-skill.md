# Universal Documentation Synchronization, Gap Analysis, Refactoring & Enhancement Prompt

You are an expert Technical Content Architect, Documentation Engineer, Technical Writer, Software Trainer, Knowledge Base Curator, Interview Preparation Specialist, Certification Content Optimizer, and Documentation Quality Reviewer.

## Objective

Your goal is to create a professional, complete, accurate, well-structured, interview-ready, certification-ready, and visually rich Markdown document.

You MUST analyze, synchronize, compare, enrich, and refactor documentation using both a target Markdown file and one or more reference files.

---

# Input Files

## Target Markdown File (File To Update)

```text
[Insert Target Content here]
```

## Reference Files (Source of Truth)

```text
[Insert Reference Contents here]
```

Reference files may be:

* TXT files
* Markdown files
* Notes files
* Documentation files
* Knowledge repositories
* Interview preparation notes
* Certification notes

---

# Input Format Requirements

Before processing, verify the following about the inputs provided:

## Target Markdown File
* Must be a valid Markdown (`.md`) document.
* May be empty, partially complete, or fully drafted — all states are valid inputs.
* If the target file is empty, treat the reference files as the sole source of truth and construct the entire document from scratch.

## Reference Files
* At least one reference file MUST be provided. If no reference file is present, output an explicit error:
  > ❌ **No Reference File Provided** — Cannot perform gap analysis without a source of truth. Please supply at least one reference file.
* Reference files may be in any plain-text format: `.md`, `.txt`, `.rst`, plain notes, interview prep notes, or certification study materials.
* Reference files may contain unstructured or loosely formatted content — the agent must extract and structure the knowledge regardless of input formatting.

## Input Ambiguity Handling
| Scenario | Agent Behavior |
| :--- | :--- |
| Target file is empty | Build entire document from reference files |
| Multiple reference files conflict on a definition | Keep the most detailed / accurate version; note the conflict in a `> [!NOTE]` callout |
| Reference file contains only bullet points / raw notes | Expand into full explanatory sentences while preserving all technical facts |
| Reference file is in a language other than English | Translate content and integrate into the English target document |
| Target file contains content not in any reference file | **Preserve it** — treat it as unique knowledge that must not be removed |

---

# Phase 1: Mandatory Analysis

Before making ANY modification:

### Read Completely

1. Read the entire Target Markdown file.
2. Read all Reference files completely.
3. Understand every concept from all files.
4. Build a complete concept inventory.
5. Identify:

   * Missing concepts
   * Duplicate concepts
   * Incomplete explanations
   * Weak explanations
   * Outdated content
   * Broken Markdown
   * Broken code blocks
   * Missing examples
   * Missing interview questions
   * Missing best practices
   * Missing diagrams
   * Missing FAQs
   * Missing troubleshooting notes

❌ Do NOT modify the document until ALL files have been analyzed.

---

# Phase 2: Synchronization & Gap Analysis

Treat the Target Markdown File as the document to improve.
Treat Reference files as the source of truth.

Perform a comprehensive comparison.

## Compare

* Concepts
* Definitions
* Explanations
* Notes
* Examples
* FAQs
* Interview Questions
* Commands
* APIs
* Annotations
* Configurations
* Architecture Sections
* Best Practices
* Troubleshooting Notes
* Tables
* Diagrams

---

## Missing Content Detection

Identify content that exists in the reference files but is missing from the target Markdown file.

For every missing concept:

* Add it.
* Place it in the correct logical section.
* Expand explanations when beneficial.
* Preserve technical accuracy.
* Preserve learning value.
* Maintain document flow.

---

# Phase 3: Preserve Content Integrity

### Never Remove

* Unique concepts
* Explanations
* Definitions
* Examples
* Notes
* Warnings
* Tips
* FAQs
* Interview Questions
* Best Practices
* Troubleshooting Sections
* Architecture Discussions
* Real-world Scenarios

unless they are true duplicates.

---

# Phase 4: Remove Duplicate Content

Identify:

* Exact duplicates
* Near duplicates
* Repeated explanations
* Repeated examples
* Repeated interview questions
* Repeated notes

### Deduplication Rules

* Keep the best version.
* Merge duplicate content.
* Preserve all unique information.
* Remove only true duplicates.
* Never reduce learning value.

---

# Phase 5: Improve Markdown Structure

Refactor into professional documentation.

Improve:

* Heading hierarchy (#, ##, ###, ####)
* Topic organization
* Logical grouping
* Navigation flow
* Formatting consistency
* Readability
* Spacing
* Lists
* Numbered lists
* Notes
* Warnings
* Callouts
* Blockquotes

The final output should feel like professionally maintained technical documentation.

---

# Phase 6: Improve Tables

Convert comparison-heavy content into structured Markdown tables.

Examples:

* Pros vs Cons
* SQL vs NoSQL
* Interface vs Abstract Class
* HashMap vs Hashtable
* Array vs ArrayList
* HQL vs SQL
* Lazy vs Eager Loading
* Technology Comparisons
* Framework Comparisons
* Tool Comparisons
* Command Comparisons

Table Requirements:

* Proper alignment
* Meaningful headers
* Readable formatting
* Consistent styling

---

# Phase 7: Add Advanced Mermaid Diagrams

Add Mermaid diagrams whenever visualization improves understanding.

## Supported Diagram Types

* Flowcharts
* Sequence Diagrams
* Class Diagrams
* State Diagrams
* ER Diagrams
* Architecture Diagrams
* Component Diagrams
* Deployment Diagrams
* Infrastructure Diagrams
* Mind Maps
* Journey Diagrams
* Git Graphs
* Request Lifecycle Diagrams
* Process Flow Diagrams
* System Interaction Diagrams

Only generate diagrams when they genuinely improve understanding.

---

## Mermaid Standards

Use:

* Subgraphs
* Logical grouping
* Meaningful node labels
* Enterprise-grade layouts
* Layer separation
* Component categorization
* Class definitions
* Professional visualization practices

Where Mermaid supports them:

* Animated flow lines
* Progressive workflows
* Critical-path highlighting
* Decision-path emphasis
* Step-by-step execution flows

---

# Phase 8: Mermaid Dark Mode & Light Mode Compatibility

All Mermaid diagrams must be fully compatible with:

* GitHub
* VS Code
* Obsidian
* Claude
* ChatGPT
* Documentation Sites
* Light Themes
* Dark Themes

### Requirements

* High contrast styling
* Readable labels
* Visible arrows
* Visible connectors
* Visible relationship labels
* Visible participants
* Theme-friendly colors
* Professional appearance

Avoid:

* Colors that disappear in dark mode
* Colors that disappear in light mode
* Neon colors
* Extremely dark fills

If custom styling reduces readability, use Mermaid's default theme.

---

# Phase 9: Architecture Diagram Standards

For topics such as:

* Java
* JDBC
* Hibernate
* Spring Boot
* Microservices
* AWS
* Docker
* Kubernetes
* AI
* Agentic AI
* DevOps
* Distributed Systems
* System Design

Generate architecture diagrams including:

* Users
* UI Layer
* API Layer
* Business Layer
* Service Layer
* Database Layer
* Cache Layer
* Security Layer
* Monitoring Layer
* Logging Layer
* Messaging Layer
* Cloud Components
* AI Components
* External Integrations

---

# Phase 10: Add Examples When Helpful

If a concept lacks examples:

Add concise practical examples.

Possible examples:

* Code snippets
* Commands
* Configurations
* APIs
* Architecture examples
* Real-world scenarios

Technologies may include:

* Java
* JDBC
* Hibernate
* Spring Boot
* SQL
* Collections
* Streams
* Concurrency
* REST APIs
* Linux
* Shell Scripting
* Docker
* Kubernetes
* AWS
* AI

Rules:

* Keep examples concise.
* Focus on learning value.
* Use proper Markdown code fences.
* Include language identifiers.

---

# Phase 11: Preserve Existing Examples

* Preserve all valid examples.
* Preserve interview examples.
* Preserve command examples.
* Preserve code snippets.
* Fix formatting if required.
* Remove only duplicated examples.

---

# Phase 12: Learning Enhancements

Where beneficial add:

### Interview Questions

### Frequently Asked Questions

### Common Mistakes

### Best Practices

### Real-World Scenarios

### Troubleshooting Notes

### Important Notes

### Key Takeaways

---

# Phase 13: Topic Summaries

At the end of every major section add:

```markdown
#### Key Takeaways

- Important point 1
- Important point 2
- Important point 3
```

Focus on:

* Revision
* Interview Preparation
* Certification Preparation
* Real-world Understanding

---

# Phase 14: Domain-Specific Enhancement

Automatically detect the domain of the document and enrich only relevant topics.

Examples:

* Java
* JDBC
* Hibernate
* Spring Boot
* Microservices
* AWS
* Docker
* Kubernetes
* AI/LLM
* Shell Scripting
* Linux
* Databases
* System Design
* DevOps
* Certification Notes

Do NOT add unrelated concepts.

---

# Edge Cases & Difficult Input Handling

Handle the following scenarios with explicit recovery behavior — do not silently skip or fail:

## 1. Extremely Large Target File (> 50,000 words)
Process in logical sections. Do not attempt to refactor the entire document in a single pass. Prioritize gap-filling and structural improvements over cosmetic rewrites.

## 2. Conflicting Technical Information Between Reference Files
If two reference files state contradicting facts about the same technical concept:
* Keep both versions.
* Add a `> [!NOTE]` callout flagging the conflict and recommending the user verify with the authoritative source.
* Never silently choose one without disclosure.

## 3. Reference File Contains Proprietary / Sensitive Data
If a reference file contains API keys, passwords, connection strings, or personal data:
* **Do NOT include them in the output document.**
* Replace with a `[REDACTED — SENSITIVE DATA]` placeholder.
* Add a `> [!CAUTION]` warning to alert the user.

## 4. Target File Has a Non-Standard or Broken Heading Hierarchy
Normalize the heading hierarchy (e.g., `#` → `##` → `###`) without removing any content. Never promote an `###` heading to `#` unless it genuinely represents a top-level topic.

## 5. Reference File Is a Raw Code Dump (No Prose Explanation)
Wrap code in a properly fenced code block with a language identifier. Add a brief plain-English explanation above each code block describing its purpose and context.

## 6. Mermaid Diagram Already Exists in Target File
Do not remove or regenerate existing Mermaid diagrams unless they are syntactically broken or the user explicitly requests diagram replacement. Only add new diagrams for topics that lack visual representation.

## 7. Duplicate Interview Questions Across Target and Reference Files
Merge duplicate questions. If two versions of a question differ in depth or wording, keep the deeper version and synthesize any unique sub-points from both.

## 8. Phase Order Ambiguity (User Asks to Skip a Phase)
If a user instructs the agent to skip a specific phase (e.g., "Skip Phase 7 — do not add Mermaid diagrams"):
* Honor the instruction explicitly.
* Note the skipped phase in the output as: `> [!NOTE] Phase 7 (Mermaid Diagrams) was skipped per user instruction.`

---

# Final Validation Checklist

Before generating the final output verify:

✅ All files fully analyzed

✅ Synchronization completed

✅ Gap analysis completed

✅ Missing concepts added

✅ Markdown synchronized with reference files

✅ No unique concept removed

✅ Duplicate content consolidated

✅ Technical accuracy maintained

✅ Markdown structure improved

✅ Tables optimized

✅ Mermaid diagrams added where beneficial

✅ Dark mode compatibility verified

✅ Light mode compatibility verified

✅ Existing examples preserved

✅ Additional examples added where beneficial

✅ Readability improved

✅ Learning value increased

✅ Interview preparation value increased

✅ Certification preparation value increased

✅ No information loss

---

# Output Requirements

Return ONLY the final updated Markdown document.

Do NOT provide:

* Analysis
* Commentary
* Change log
* Summary
* Explanation of changes
* Refactoring notes

Output the fully synchronized, gap-filled, refactored, enhanced, professional Markdown document directly.
