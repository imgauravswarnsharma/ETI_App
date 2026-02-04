# ETI / Personal Intelligence App  
## AI Operating Contract (Authoritative)

---

## AUTHORITY

This document defines the **binding operating rules** for all AI responses
inside the ETI project.

Unless explicitly overridden, this document is authoritative.

---

## PROJECT CONTEXT
- ETI is a long-running personal data intelligence system. Started as transaction logging and item rate comparisons on the go and aimed to be evolved into a personal data intelligence system with multiple analytical axes layered over time once core goals are achieved.

- This section provides contextual grounding and does not override versioned documentation.

---

## PROJECT DETAILS
Current Version: 1.3  
Data Source Name: ETI_App_v1.3  
App Name: Personal Intelligence App  

Current execution platform:
- Google Sheets (data store)
- AppSheet (UI)
- Google Apps Script (automation – user does not write GAS independently; AI-generated only)

Future platforms may change, but current logic and documentation are version-bound and authoritative.

---

## Version 1.3 Scope and Goals
- Transaction-grain integrity  
- Item-level analytics  
- Transaction-level analytics  
- Loss-free and lag-free sync  
- Dual rate UI (if applicable)  
- Performance-first formulas and scripts  
- On-the-go item rate comparison  
- Auditability and explainability  
- Version-safe evolution  

---

## RESPONSE STYLE (MANDATORY)

- Precise, factual, implementation-oriented
- No emojis
- No conversational fluff
- Copy-paste safe output only (VS Code / Notepad)

When explaining theory:
- Use headings and subheadings
- Use bullet points or tables
- Avoid prose paragraphs unless unavoidable

---

## TABLE & COLUMN EXPLANATION RULES

When documenting tables or schemas:
1. Explain **role and grain first**
2. Then list columns
3. For formula columns, provide:
   - Row-level sample formula
   - Short logic explanation
   - Performance notes (if relevant)

---

## DERIVED COLUMN LOGIC (MANDATORY PRE-CONTEXT)

- `Derived_Column_Logic` is a **first-class authoritative reference**
- AI MUST consult Derived_Column_Logic before:
  - Explaining any formula behavior
  - Suggesting formula changes or optimizations
  - Reasoning about column dependencies
  - Interpreting downstream effects
- No formula reasoning is allowed without respecting:
  - Semantic_Class
  - Resolved_References
  - Upstream dependency chains

Derived logic supersedes ad-hoc formula interpretation.

---

## APPSHEET CONTEXT & UI AUTHORITY

- AppSheet exported documentation and column mappings are **authoritative UI context**
- AI MUST consult them before:
  - Suggesting new views
  - Modifying UX behavior
  - Recommending slices or actions
  - Commenting on column visibility or editability

Rules:
- Do not duplicate existing views
- Do not contradict existing UI intent
- Do not infer missing UX if documentation exists

---

## DOCUMENTATION MODE

When asked to document or generate documentation:
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

All scripts MUST include the following header:

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
