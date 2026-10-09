# Demo talk track (5 minutes)

**Setup:** Claude has the `uat-script-generator` skill enabled. Have `sample-output/` open in a second window as a fallback. For a no-install version, screen-share the interactive walkthrough (https://caseymullan.github.io/claude-skills/uat-script-generator/).

---

### 1. Frame the problem (30 sec)
> "On enterprise HRIS integrations, UAT is where trust gets won or lost. Two things kept going wrong. The build didn't quite match what the customer signed, and the test scripts were copied from the last customer, so they missed the edge cases that actually break: rehires, leave, contractors, time zones. I built a tool for that at my last role. This is a from-scratch rebuild with fictional data."

### 2. Run it (60 sec)
Type: **"Run the UAT demo for the Halvorsen Freight user sync."**

While it runs:
> "The input is two files: the workflow export from the iPaaS, and the requirements spec the customer signed in discovery."

### 3. The verdict (60 sec)
Claude leads with: *Not ready for customer UAT: 3 blocking issues.* Pick one and tell its story:
> "Hire date is mapped from Original_Hire_Date. Every field looks populated, so nobody notices, until a rehire comes back with a 2017 hire date and never gets onboarding. I've seen that one in production."

### 4. Open the workbook (90 sec)
- **Overview:** the red ✗ NOT READY banner and KPI tiles. "Every number here is a formula, so it updates as the customer's testers fill in results."
- **Pre-UAT Validation:** failures sorted to the top, each with a recommended action.
- **Test Cases:** point at the *"⚠ Will fail until V-07 is fixed"* notes in the Pre-UAT dependency column. "The customer never runs a test we already know is broken. That protects the relationship."
- **Sign-off:** "These criteria are formulas. Nobody can type 'Met'."

### 5. Fix and re-run (60 sec)
Type: **"Apply the fixes and re-run."**
Claude edits a copy of the workflow and regenerates. The FAIL count drops to 0, and the Overview banner turns from red ✗ NOT READY to ▶ READY FOR CUSTOMER UAT. Check IDs don't change, so V-07 is still V-07, now ✓ PASS.

### 6. Land it (30 sec)
> "The script is deterministic: same input, same workbook. Claude does the parts that need judgement: explaining what each failure means for the customer, and making the fix. Supporting another iPaaS means writing an adapter; the checks don't change."

---

**Likely questions**
- *Why not have Claude write the test cases directly?* Repeatability and auditability. A customer-facing artifact shouldn't change between two runs on the same input.
- *How would this work with real Tray or Workato?* Write an export adapter that maps to the input format in `references/input-formats.md`.
- *What would you add next?* A data-preview tab (a dry run of the first load against the test tenant) and more integration types (SSO, Teams provisioning).
