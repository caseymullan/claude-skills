# Input formats

Both formats are deliberately small. The workflow format is a simplified stand-in for an iPaaS export (modeled loosely on Tray.ai). To support a real platform's export, write an adapter that produces this shape. The validation and workbook logic doesn't change.

## Workflow export (`tray-workflow-mock/v1`)

```json
{
  "workflow": {
    "id": "wf_…",
    "title": "Customer | Source -> Target | Purpose",
    "customer": "Customer name",
    "tags": ["hris_user_sync", "env:uat"],
    "trigger": { "type": "scheduled", "schedule": { "interval": "15m | 2h | daily | weekly" } },
    "steps": [
      { "name": "workday-1", "connector": "workday", "operation": "get_custom_report", "properties": {} },
      { "name": "boolean-1", "connector": "boolean-condition", "operation": "filter",
        "properties": { "conditions": [ { "field": "Employee_Class", "operator": "not_in", "value": ["CONTRACTOR"] } ] } },
      { "name": "data-mapper-1", "connector": "data-mapper", "operation": "map",
        "properties": { "mappings": [
          { "source": "Employee_ID", "target": "external_id" },
          { "source": "Worker_Status", "target": "status", "transform": "value_map", "value_map": { "Active": "active" } },
          { "source": "completed_at", "target": "COMPLETION_DATE", "transform": "datetime:epoch_ms", "timezone": "UTC" }
        ] } },
      { "name": "lp-1", "connector": "learning-platform-api", "operation": "upsert_user",
        "properties": { "match_on": "external_id", "on_missing_from_feed": "deactivate" } },
      { "name": "error-handler-1", "connector": "error-handler", "operation": "on_error",
        "properties": { "retry_on_status": [429, 500, 503] } }
    ]
  }
}
```

How the generator reads it:
- **Source and target:** the first and last steps whose connector isn't a utility. Loop, boolean-condition, data-mapper, error-handler, script and branch are utilities.
- **Integration type:** taken from `tags`. If the tags don't name one, it's inferred from the connectors.
- **Environment:** taken from the `env:` tag.

## Requirements spec (`integration-requirements/v1`)

This is the configuration agreed and signed with the customer during discovery.

| Key | Used for |
|---|---|
| `integration_type`, `source_system`, `target_system` | Scope check, labels |
| `signed_off[]` (`name`, `role`, `date`) | Overview + Sign-off approvers |
| `schedule` | Cadence check (faster than spec is WARN, slower is FAIL) |
| `timezone` | Checked against every `datetime:*` transform |
| `unique_key` (`source`, `target`) | Identity check against `match_on` |
| `required_mappings[]` (`source`, `target`, optional `note`) | One check per field; the `note` explains why the field matters |
| `filters[]` | Population-filter check (missing values are FAIL, extra values are WARN) |
| `lifecycle` (status value to outcome, plus `missing_from_feed`) | Lifecycle checks against the status `value_map` and `on_missing_from_feed` |
| `error_handling.retry_on_status[]` | Retry coverage check |
| `expected_volume.peak_completions_per_run` | Sizes the peak-volume test case |
