---
name: uat-script-generator
description: Generate a customer-ready UAT workbook from an integration workflow export (Tray-style JSON), first validating the build against the signed requirements spec. Use for HRIS user-sync or LMS-completion integrations, UAT test scripts, pre-UAT build checks, or "run the UAT demo".
---

# UAT Script Generator

Turns an integration workflow export into the UAT package a customer's HRIS/IT team actually tests from, and catches build-vs-requirements mismatches *before* the customer finds them.

**Input:** a workflow export (Tray-style JSON) and, ideally, the signed requirements spec for that integration.
**Output:** an .xlsx workbook with eight tabs:

| Tab | Purpose |
|---|---|
| Overview | Dashboard: live readiness banner, KPI tiles, test-progress bar, workflow details |
| Pre-UAT Validation | Every requirement in the signed spec checked against the build: ✓ PASS / ✗ FAIL / ⚠ WARN / ℹ INFO, problems sorted first, with stable check IDs |
| Field Mappings | Every field written to the target, transforms, sample values, a Verified column for the customer |
| Test Cases | Generated from what the workflow does: core flows plus lifecycle edge cases (rehire, termination, leave, email change, excluded classes, duplicates, time zones, rate limits) |
| Test Data | Fictional personas each test case references |
| Issue Log | Pre-formatted with an example row and dropdowns |
| Sign-off | Criteria are formulas (no failing checks, all Core cases passed, no open High issues, mappings verified) plus approver rows from the spec |
| AI Instructions | Rules for an agent maintaining the workbook between runs |

Supported integration types: `hris_user_sync` (HRIS → learning platform) and `lms_completions` (learning platform → LMS). The type comes from the workflow's tags, or is inferred from its connectors.

## Workflow

1. **Get the inputs.**
   - The user supplies a workflow export, plus a requirements spec if one exists. Input formats are described in `references/input-formats.md`.
   - If they ask for a demo, or supply no files, use the fictional example in `examples/halvorsen-freight/`. It covers both integration types, and its builds contain planted defects the validation step should catch. If they name one integration ("the user sync"), run only that one.
   - With no spec, the workbook still builds; tell the user the validation tab is skipped and why that matters.

2. **Run the generator** from your working directory, never from inside the skill folder. Write outputs to `outputs/` in the working directory. The generator needs openpyxl; install it if missing.
   ```bash
   python <skill-dir>/scripts/generate_uat.py --workflow <workflow.json> --spec <requirements.json> --out outputs/<name>.xlsx
   ```
   Demo output names: `UAT_Halvorsen_UserSync.xlsx` and `UAT_Halvorsen_Completions.xlsx`. The script prints a JSON summary with a `verdict`, validation counts, each non-PASS check, and test-case counts. Errors (bad JSON, unsupported type) exit non-zero with a readable message. Relay that message; don't patch the script.

3. **Brief the user before handing over the files.** Lead with the verdict.
   - **Verdict:** one combined verdict first ("Not ready: 5 blocking issues across 2 integrations"). Then a short section per integration.
   - **Each ✗ FAIL, in plain language:** what was agreed, what was built, and what goes wrong for the customer if it ships. For example: "Hire date reads from Original_Hire_Date, so rehired workers keep their 2017 date and never get new-hire onboarding."
   - **⚠ WARN and ℹ INFO items:** one line each.
   - **Test cases:** how many were generated, and how many are flagged as depending on a fix.
   - Keep it short. The workbook holds the detail, and its Overview tab shows the same verdict as a live banner.

4. **Deliver the workbook(s)** with whatever file-delivery mechanism the environment provides.

5. **Offer the fix-and-rerun loop.** This is the strongest part of a demo. If the user agrees:
   - Fix **FAILs only**. Copy the workflow to `<name>.fixed.json` and never edit the original.
   - Make each change mirror the spec exactly. Don't add transforms or values the spec doesn't state.
   - Leave WARN and INFO items alone. They need a customer decision (for example, whether an extra field should sync). List them as still open.
   - Re-run to `outputs/<name>_after-fixes.xlsx`. Check IDs are stable between runs, so "V-07" means the same check before and after.
   - Show the before/after: FAIL count dropping to 0, the Overview banner changing from ✗ NOT READY to ▶ READY FOR CUSTOMER UAT, and the Sign-off criterion "No failing pre-UAT checks" changing to ✓ Met. The other criteria stay "Not met" until the customer tests; say that's expected.

## Rules

- The signed spec is the source of truth. Never invent requirements, field names, or customer decisions. If the spec is ambiguous, list the question in your briefing instead of guessing.
- A "fix" means changing the workflow to match the spec. Never change the spec to match the workflow.
- Test data is fictional and belongs in a TEST tenant. Never put real employee records in the workbook.
- The workbook is customer-facing: short sentences, no internal jargon, no blame.
- Sign-off criteria are formulas. Don't hard-code "Met".
- If the workflow's integration type isn't supported, say so and stop. Don't force it into a template that doesn't fit.
