# AGENTS.md

## Purpose

This document defines operational constraints and autonomy rules for AI coding agents (e.g., Codex) in Python-based research environments.

Goals:
- Maximum safety (data integrity, reproducibility, security)
- Minimal interruption (high autonomy)
- Scientific rigor and traceability

---

## Core Principles

### 1. Default to Forward Progress
The agent MUST proceed without asking questions if a reasonable interpretation exists.

Heuristic:
If confidence ≥ 80% → proceed.

---

### 2. Research-Specific Autonomy Scope

Allowed without confirmation:
- Writing Python scripts, notebooks, and modules
- Data preprocessing and transformation
- Statistical analysis and model training
- Visualization (matplotlib, seaborn if explicitly required)
- Writing unit tests and experiment scripts
- Refactoring for clarity or performance
- Using standard libraries (numpy, scipy, pandas, sklearn, torch, etc.)

Allowed with implicit caution:
- Creating new experiment pipelines
- Modifying data schemas (must preserve originals)
- Adding dependencies from trusted ecosystems (PyPI)

Requires explicit confirmation:
- Deleting datasets or overwriting raw data
- Changing evaluation protocols affecting comparability
- Introducing external data sources or APIs
- Modifying environment configuration (Docker, CUDA, system libs)

---

### 3. Data Safety Rules

The agent MUST:
- Never overwrite raw data
- Always create derived datasets (e.g., /processed, /tmp)
- Preserve reproducibility (fixed seeds when applicable)
- Log transformations clearly

---

### 4. Reproducibility Requirements

All outputs MUST:
- Be deterministic where possible
- Include seeds (numpy, torch, random)
- Use explicit versioning if relevant
- Avoid hidden state

---

### 5. Silent Assumptions

The agent MAY infer:
- Default libraries (pandas, numpy, matplotlib)
- Standard file structures (/data, /src, /notebooks)

The agent MUST document assumptions AFTER execution.

---

### 6. Output Structure

Each response MUST include:
1. Result (code / changes)
2. Assumptions
3. Risk level (Low / Medium / High)
4. Optional next steps

---

### 7. Error Handling

On failure:
- Retry once with correction
- Fall back to simpler method

Do NOT loop indefinitely.

---

### 8. Dependency Policy

Allowed:
- numpy, scipy, pandas, matplotlib
- scikit-learn
- PyTorch / TensorFlow
- statsmodels

Avoid:
- obscure or unmaintained packages

---

### 9. Security & Integrity

The agent MUST NEVER:
- Expose credentials
- Download arbitrary data without justification
- Execute unsafe code

---

### 10. Code Quality

All code MUST be:
- Readable
- Modular
- Documented
- Testable (where meaningful)

---

### 11. Autonomy Override

If unsure:
→ choose safe, reversible action and proceed

---

### 12. Interruption Policy

Interrupt ONLY if:
- Data loss risk
- Scientific invalidity
- Logical impossibility

---

## Summary Directive

Execute autonomously, preserve data, ensure reproducibility, and document assumptions without interrupting workflow.
