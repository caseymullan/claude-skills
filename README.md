# Claude Skills

Skills I've built for Claude, drawn from six years of enterprise integration and solutions engineering work: Workday, SuccessFactors, Oracle HCM, Microsoft Teams, SSO, and iPaaS platforms like Workato and Tray.ai.

Each folder is a self-contained skill: a `SKILL.md` that tells Claude when and how to use it, plus the scripts, examples and reference material it needs. All demo data is fictional.

## Skills

### [UAT Script Generator](uat-script-generator/)

Checks an integration build against the requirements spec the customer signed, then generates the UAT workbook their HRIS team tests from. Test cases are derived from what the workflow actually does: rehires, terminations, excluded worker classes, time zones, rate limits.

<a href="uat-script-generator/"><img src="docs/uat-script-generator/img/hero-not-ready.png" width="600" alt="UAT workbook Overview tab: NOT READY FOR CUSTOMER UAT, 3 blocking checks"></a>

**[Interactive walkthrough](https://caseymullan.github.io/claude-skills/uat-script-generator/)** · [Sample workbooks](uat-script-generator/sample-output/) · [5-minute demo script](uat-script-generator/DEMO.md)

## Using a skill

Upload a skill's folder to Claude as a skill, or put it in your Claude Code skills directory. Then ask in plain language, for example *"Run the UAT demo"*. Each skill's README covers its inputs and options.

## About

Built by **Casey Mullan**, a solutions engineer who works with customer IT and HRIS teams from requirements through UAT and go-live, and who writes the fix when it lives in the product.

Licensed under [MIT](LICENSE).
