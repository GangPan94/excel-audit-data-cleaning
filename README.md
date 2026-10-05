# Excel Audit Data Cleaning

A reusable skill for cleaning and reconciling audit, finance, and accounting data in Excel through the ChatGPT Excel plugin.

The workflow preserves source data, inspects the workbook, writes a proposed `Cleaning_Plan` worksheet, and waits for approval before changing data. After approval, it documents transformations, reconciles results to user-provided controls, investigates differences, and writes a `Final_Summary` worksheet.

## Background

This skill accompanies the article [Using an Excel Agent to Clean, Validate, and Reconcile Data](https://editions.journalofaccountancy.com/article/Using+an+Excel+agent+to+clean%2C+validate%2C+and+reconcile+data/5207267/869709/article.html) by Gang Pan and Li Wang, *Journal of Accountancy*, 242(4), 20–27, October 2026.

The [`v1.0.0` release](../../releases/tag/v1.0.0) is the version described in the article. Later versions may differ; see `CHANGELOG.md`.

## Use

Load `SKILL.md` into the reusable skill or instruction mechanism available in your ChatGPT environment. Open the target workbook with the Excel plugin, then give a request such as:

> Clean the general ledger data on the GL Data tab and reconcile it to the trial balance. Read the workbook guide first.

The skill should create a `Cleaning_Plan` worksheet and ask for approval before it transforms workbook data. The exact behavior depends on the Excel plugin's available workbook actions and the connected model.

## Workflow

1. Inspect the workbook in read-only mode.
2. Write a `Cleaning_Plan` worksheet and wait for approval.
3. Preserve source data in `Raw_Data` and clean a copy in `Working_Data`, keeping a `Source_Row_ID` for each row.
4. Log every approved transformation in `Modifications_Log`.
5. Reconcile to the user-provided control source and investigate any differences.
6. Apply final quality gates and write a `Final_Summary` worksheet.

## Testing

Tested in July 2026 using ChatGPT for Excel in Microsoft Excel for Mac on macOS Tahoe, with the Heavy thinking-effort setting, on general ledger cleaning and reconciliation scenarios. OpenAI stated that ChatGPT for Excel was powered by GPT-5.5 as of May 5, 2026; GPT-5.5 is therefore the documented model for this test period. See [OpenAI's ChatGPT for Excel announcement](https://openai.com/index/chatgpt-for-excel/).

## Disclaimer

This skill helps prepare data for reconciliation review. It is not audit advice and does not establish compliance with AS 1105 or any other standard. Users and engagement teams remain responsible for verifying results against source records and applying professional judgment. Do not share confidential client data when reporting issues.

Your results may differ from those described in the article. Common reasons include:

- **Model changes.** OpenAI may update or replace the model behind ChatGPT for Excel, which can change how the skill is followed.
- **Thinking-effort setting.** Lower effort settings than the Heavy setting used in testing may skip steps or perform less thorough checks.
- **Excel platform and plugin version.** Excel for Windows, Mac, and the web, and different plugin versions, may support different workbook actions.
- **Workbook size and structure.** Very large ledgers, merged cells, hidden rows, external links, or unusual layouts may exceed what the agent can read or write reliably in one pass.
- **Missing context.** Results depend on the workbook containing the needed control totals, mappings, and notes, or on the user supplying them.
- **Prompt wording.** Vague requests, or requests that skip the approval step, can lead to different behavior.
- **How the skill is loaded.** Pasting the skill as custom instructions, uploading it as a file, or loading it as a skill may affect how consistently it is applied.
- **Run-to-run variation.** AI models are not deterministic; the same prompt on the same workbook can produce different plans or outputs.

Always review the `Cleaning_Plan` before approving it and check `Modifications_Log` and `Final_Summary` against the source data.

## Citation

If you use this skill in research or practice, please cite the article:

```bibtex
@article{pan2026using,
  author  = {Pan, Gang and Wang, Li},
  title   = {Using an Excel Agent to Clean, Validate, and Reconcile Data},
  journal = {Journal of Accountancy},
  volume  = {242},
  number  = {4},
  pages   = {20--27},
  month   = {10},
  year    = {2026},
  url     = {https://editions.journalofaccountancy.com/article/Using+an+Excel+agent+to+clean%2C+validate%2C+and+reconcile+data/5207267/869709/article.html}
}
```

## Contents

- `SKILL.md` - complete skill instructions
- `CHANGELOG.md` - version history
- `CITATION.cff` - citation metadata
- `LICENSE` - MIT License

This package contains no scripts or external dependencies.
