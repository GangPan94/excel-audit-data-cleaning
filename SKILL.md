---
name: excel-audit-data-cleaning
description: Use this skill whenever the user wants to clean, validate, prepare, normalize, reconcile, or document Excel data for audit, finance, accounting, ledger, trial balance, transaction, or AS 1105 evidence workflows in ChatGPT with an Excel plugin. Use it for dirty spreadsheets, GL exports, subsidiary ledgers, trial balances, debit/credit cleanup, completeness and accuracy checks, audit trails, source-data preservation, or reconciliation work. It follows a native-Excel-only workflow: inspect first, write a `Cleaning_Plan` worksheet, ask for approval, then clean, document, reconcile, and apply final quality gates.
license: MIT. See LICENSE and the disclaimer in README.md (https://github.com/GangPan94/excel-audit-data-cleaning#disclaimer).
---

# Excel Audit Data Cleaning

Use this skill to prepare audit, finance, or accounting data in Excel without overwriting source data. The workflow is non-destructive: inspect, document the plan, get approval, clean only approved items, reconcile, and preserve an audit trail.

This skill can support AS 1105-style evidence preparation, but it does not conclude compliance. The user or engagement team must confirm that the final data ties to the proper source records and fits their audit methodology.

## Core Rules

- Use only native Excel features available through the Excel plugin: worksheets, formulas, filters, sorting, formatting, tables, and PivotTables.
- For large ledgers, avoid broad full-sheet scans or writes. Prefer targeted reads, row counts, sample rows, PivotTables, chunked summaries, and small verified writes.
- Treat the user's initial request as permission to inspect and create `Cleaning_Plan` only. It is not approval to transform data.
- For every cleaning request, create or update a `Cleaning_Plan` worksheet before asking for approval.
- Stop after asking for approval. Do not clean, verify, reconcile, create working tabs, write formulas into data sheets, delete rows, convert values, or modify source data until the user approves the plan.
- After approval, apply only approved steps. Ask again before any change that could alter business meaning, such as interpreting ambiguous dates, changing account numbers, deleting rows, mapping currencies, or netting debit and credit columns.
- Preserve the original data and keep row-level traceability with `Source_Row_ID`.
- Log every approved transformation in `Modifications_Log`.
- Do not treat mixed original currencies as a reconciliation problem by itself. First determine whether the amount field has already been translated into a common reporting currency.
- Do not report a reconciliation difference as a true variance until agent error and missed workbook context have been ruled out.
- Before final delivery, apply final quality gates. If a required gate is blocked, fix it or ask for the missing evidence instead of calling the work complete.
- Before final chat delivery, create or update a `Final_Summary` worksheet that shows what was done, what changed, what reconciled, and what remains open.
- Keep detail tables, reconciliation summaries, quality checks, and final summaries in separate tabs or clearly separated ranges to avoid overwriting headers or results.
- Avoid saying the data is "compliant" or "audit-ready." Prefer "prepared for reconciliation review," "supports the audit trail," or "blocked pending user input."

## Planning and Approval

First inspect in read-only mode: sheet names, relevant headers, row and column counts, sample rows, visible quality issues, control sources, and context tabs such as `Report Contents`, mapping tables, legends, assumptions, instructions, or reconciliation notes.

Then create or update `Cleaning_Plan` with the full proposed plan. Keep it concise and structured:

| Section | Content |
|---|---|
| Scope | Target sheet, data type, control source, relevant context tabs |
| Issues Found | Data quality observations with affected column or range, example cell, observed value, and why it matters |
| Proposed Cleaning | Approved-action candidates with the exact native Excel method or formula where practical |
| Verification Plan | Row count, total checks, PivotTables, variance tolerance, bridge checks, and reconciliation approach |
| Currency Treatment | Original currency field, reporting currency evidence, amount field used for reconciliation, and whether mixed original currencies can be aggregated |
| Reconciliation Key Plan | Control-source grain, candidate account keys, and account-name formulas to test against the trial balance |
| Output Tabs | Planned tabs such as `Working_Data`, `Modifications_Log`, `Reconciliation`, bridge tabs, and `Final_Summary` |
| Judgment Points | Ambiguous dates, sign conventions, duplicates, exclusions, mappings, or missing context |
| Approval Status | `Pending Approval`, `Approved`, or `Revised` |

In chat, confirm that `Cleaning_Plan` was written, summarize only the highest-risk judgment points, and ask:

> Please review `Cleaning_Plan`. Do you approve me to proceed with these steps, or would you like changes first?

End the response there and wait.

## Approved Workbook Setup

After approval, establish:

- `Raw_Data`: untouched source data, or a preserved copy if the original tab cannot be renamed safely.
- `Working_Data`: copy of the source where cleaning occurs.
- `Modifications_Log`: audit trail of approved transformations.
- `Reconciliation`: required when reconciliation is in scope.
- Bridge tabs, such as `FY_Variance_Bridge`: use when multiple periods or control tabs must tie together.
- `Source_Row_ID`: leftmost column in `Working_Data` mapping each row back to the source row.

Use this `Modifications_Log` structure:

| Step | Timestamp | Sheet | Column or Range | Issue Found | Approved Action | Excel Method or Formula | Rows Affected | Validation Performed | Notes |
|---|---|---|---|---|---|---|---|---|---|

When exact row counts are unknown, log the filter or test used to identify affected rows.

## Cleaning Workflow

### 1. Scan

Start with core checks: headers, blank rows, report artifacts, repeated headers, numbers stored as text, negative values, date consistency, required fields, duplicates, debit/credit structure, empty or redundant columns, source-system identifiers, and currency fields.

Use deeper checks only when symptoms appear: nonbreaking spaces, encoding issues, merged cells, hidden rows, active filters, mixed units or currencies, character confusion, subtotal rows, or formula-driven totals. When original currencies differ, look for workbook evidence that the reconciliation amount is already translated into the reporting currency.

For each issue included in `Cleaning_Plan`, provide at least one example cell reference and observed value. Prefer formula-level fixes such as `=TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," ")))`, `=NUMBERVALUE(A2)`, `=YEAR(DateCell)`, `=IF(DebitCreditCell="S",AmountCell,-AmountCell)`, or account-key concatenation formulas.

### 2. Clean

Use helper columns when a transformation needs review before replacing values. Common native Excel methods include `TRIM/CLEAN/SUBSTITUTE` for text cleanup, `NUMBERVALUE` for numeric conversion, date standardization after interpretation is confirmed, `COUNTIF` duplicate flags, and debit/credit netting only after sign convention is approved.

Remove report artifacts with logged filters rather than undocumented manual deletion.

### 3. Validate

Run lightweight validation first:

- compare raw, working, and cleaned row counts
- confirm no unintended row loss
- recalculate key totals before and after cleaning
- check required fields and documented exceptions
- summarize by relevant account, entity, period, currency, or transaction type
- reconcile to the user-provided trial balance, GL, subledger, bank statement, or other control source

Before reconciling, infer the control source's aggregation grain before choosing a key. Test whether the control is summarized by base GL account, currency + account, GL account + cost center, company code + account, account description, or mixed grain where only some accounts use cost center detail. Use cost center suffixes only when the control account exists at that grain. Use helper formulas such as `=TRIM(A2)`, `=SUBSTITUTE(A2,"-","")`, `=TEXT(AccountCell,"000000")`, `=CompanyCodeCell&"-"&GLAccountCell`, `=GLAccountCell&"-"&CostCenterCell`, or `=GLAccountCell&" - "&AccountNameCell`.

Ask for or document the variance tolerance, including rounding threshold and whether account-level differences are acceptable when total-level controls tie. For multi-period reconciliations, add a bridge check such as all-period variance = sum of period-level variances before reporting a residual difference.

Before aggregating across original currencies, confirm the selected amount field is denominated in the reporting currency. If translation has already occurred, original currency may be informational and account-level aggregation can proceed. If translation is not evidenced, do not aggregate across currencies without approved FX translation logic or rate support.

If totals tie and no quality gates are blocked, do not run deeper variance procedures.

## Variance Investigation

Use this section only when reconciliation does not tie or the control source is ambiguous. Treat differences as preliminary until self-checks are complete.

Before reporting an unresolved difference:

- reread `Report Contents` and any instruction, mapping, legend, assumption, control, or reconciliation note
- confirm source tab, control tab, period, entity, account range, currency, business unit, population, and whether the control is balance, activity, subtotal, opening balance, ending balance, or filtered subset
- if account names or keys do not match, infer the control grain and test reasonable GL column combinations, mixed-grain keys, and normalized account-key formulas before reporting a mismatch
- compare each difference to the approved variance tolerance and, for multi-period controls, bridge all-period variance to period-level variances
- confirm whether the amount used for reconciliation is a reporting-currency amount, whether original currency is informational, or whether FX translation/rate support is missing
- verify debit/credit sign conventions, parenthetical negatives, normal balances, and whether debit and credit columns should be netted
- check hidden rows, filters, subtotals, repeated headers, blanks, merged cells, exclusions, and formulas affecting totals
- compare `Raw_Data`, `Working_Data`, and cleaned totals to confirm the cleaning did not introduce the difference
- trace the difference with `Source_Row_ID`, filters, and PivotTables
- review `Modifications_Log` for omitted exceptions or transformation errors
- inspect nearby tabs for mappings, adjustments, exclusions, opening balances, reclasses, or reconciliation support

If the agent caused the difference or missed context, correct it, update `Modifications_Log`, and rerun validation. If the difference remains, report it as an unresolved reconciliation difference with amount, direction, affected fields, context reviewed, tests performed, and remaining user input needed.

## Final Quality Gates

Before final delivery, run these pass/block gates. Do not assign points, ratings, or percentages.

| Gate | Required Result |
|---|---|
| Preservation and approval | Source data is preserved, `Source_Row_ID` exists, and all transformations were approved. |
| Documentation | `Modifications_Log` records transformations, formulas or methods, affected rows, validation, and exceptions. Required output tabs exist without overwritten headers or overlapping summary blocks. |
| Data quality | Identified issues are resolved or clearly documented as exceptions. |
| Reconciliation | Totals reconcile, or unresolved differences completed the variance investigation and are labeled as unresolved. Currency treatment, variance tolerance, bridge checks, and account-key logic are confirmed before account-level aggregation. |
| Communication | `Final_Summary` exists and assumptions, open items, blocked inputs, and limitations are clear; no audit compliance claim is made. |

If any gate is blocked, do not deliver final results. Fix the issue or ask for the missing approval, source view, control total, mapping, or explanation. Use "blocked pending user input" rather than "complete."

## Final Output

When validation is complete:

- create or update `Final_Summary` with sections for scope, completed steps, key formulas or methods, reconciliation result, quality gate status, unresolved exceptions, assumptions, and next user action
- summarize changes made
- point to `Modifications_Log`
- report final gate status
- identify unresolved exceptions, assumptions, or blocked inputs
- keep `Source_Row_ID` unless the user requests a clean export copy
- remind the user to preserve `Raw_Data`, `Working_Data`, `Modifications_Log`, and the conversation record if needed for the audit trail
