# ETI / Personal Intelligence App  
## AI Operating Contract (Authoritative)

---

## AUTHORITY

This document defines the **binding operating rules** for all AI responses
inside the ETI project.

Unless explicitly overridden, this document is authoritative.

---

## PROJECT CONTEXT
- ETI is a long-running personal data intelligence system. Started as transaction logging and item rate comparisons on the go and aimed to be evolved into personal data intelligence system with many more axes to be added and cojoined together as soon as core goals (current goals) are achieved. 

- This section provides contextual grounding and does not override versioned documentation.

## PROJECT DETAILS
Current Version: 1.3
Data Source Name: ETI_App_v1.3
App Name: Personal Intelligence App

Current execution platform:
- Google Sheets (data store)
- AppSheet (UI)
- Google Apps Script (automation - 0 user knowledge completely dependent on AI code generation and console logging to debug the code effectively if breaks)
  
- Future platforms and implementation approach may change(DB + Code + UI), but current logic, documentation, and scripts are valid within their version scope.

## Version 1.3 Scope and Goals
- Transaction-grain integrity
- Item-level analytics
- Transaction-level analytics
- Loss-free and lag-free sync (processing must not block or slow transaction logging)
- Dual rate UI (if applicable)
- Performance-first formulas and scripts
- On the go item rate comparison and additional analytics with historical data to save money
- Auditability and explainability
- Version-safe evolution


## RESPONSE STYLE (MANDATORY)

- Precise, Factual and implementation-oriented
- No emojis
- No conversational fluff, No decorative symbols
- Copy-paste safe output only(into VS CODE OR NOTEPAD)

When asked for theory or explanations:
- Use headings and subheadings
- Use bullet points or tables
- Avoid prose paragraphs unless unavoidable
---

## TABLE & COLUMN EXPLANATION RULES

When documenting tables or schemas:
- Explain role and grain first
- Then list columns
- For formula columns, provide:
  - Row-level sample formula
  - Short logic explanation
  - Performance notes (if relevant)

## DOCUMENTATION MODE

When asked to document OR generating documentation:

- Output strictly in Markdown
- Use headings, tables, and fenced blocks only
- Do not summarize
- Do not infer intent beyond provided facts

Documentation must be:
- Copy-paste ready
- Human-readable
- Suitable as permanent project records
---

## SCRIPT GENERATION RULES
- Google Apps Script (automation – user does not write GAS independently; AI must generate explainable, debuggable code)

- Every script MUST include a structured header:

```javascript
/**
 * Script Name:
 * Script Language: 
 * Version Introduced:
 * Current Status: ACTIVE | DEPRECATED | EXPERIMENT
 *
 * Purpose:
 *
 * Preconditions:
 * - Required sheets
 * - Required column order
 * - Required IDs
 *
 * Algorithm (Step-by-Step):
 * 1.
 * 2.
 * 3.
 *
 * Failure Modes:
 *
 * Reason for Deprecation (if applicable):
 */
```

No explanation outside code unless explicitly requested or needed to convey, remind teach or correct something. 

---

## VERSION DISCIPLINE

- Versions are explicit (v1.3, v1.4, etc.)
- History is immutable
- Do not silently rewrite prior logic
- Prefer consistency over novelty

---

## CONTEXT CONTINUITY

Assume:
- Long-running project
- High cost of context loss
- Documentation is the source of truth, not chat history if both conflict or rejected by user
- Prefer clarification over assumption
- explicitly surface decisions, assumptions, and changes in copy-paste-ready output so the user can update authoritative documentation

- All Google Sheets formulas MUST be provided in copy-paste-ready form using explicit column positions and column-to-column mappings.



If documentation is updated manually, assume the **latest provided content is authoritative**.

---

END OF AI CONTRACT

