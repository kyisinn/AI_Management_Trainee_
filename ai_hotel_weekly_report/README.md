# Automated Weekly Hotel Performance Reports

This submission implements the assignment in two forms:

1. `hotel_weekly_report_n8n.json` - an n8n workflow for the production design.
2. `hotel_weekly_report.py` - a local reference implementation that can be tested immediately with the included CSV or the supplied Excel file.

The automation reads one row per hotel, calculates peer medians, classifies each hotel, creates a personalized message, and prepares or sends an email. The mock file has no date or goal columns, so each Monday run treats the current sheet as the previous week's snapshot. The benchmark is recalculated from all valid rows each run.

## Quick test

Python 3.10+ is sufficient for the included CSV:

```bash
python3 hotel_weekly_report.py \
  --input sample_input.csv \
  --output-dir run_output \
  --run-date 2026-10-05 \
  --test-email your.email@example.com
```

The default is dry-run mode. It creates:

- `results.csv` with classifications, benchmarks, and final recipients.
- `messages.json` with structured message data.
- one HTML email preview per hotel in `email_previews/`.

To read the supplied `.xlsx` directly, install `openpyxl` if it is not already available:

```bash
python3 -m pip install openpyxl
```

## Optional AI generation

Set an API key and request AI-generated wording. The program uses the OpenAI Responses API with a strict JSON schema and `store: false`. If the request fails, it records the error and uses the deterministic fallback message so one API failure does not stop the weekly batch.

```bash
export OPENAI_API_KEY="..."
export OPENAI_MODEL="gpt-5-mini"
python3 hotel_weekly_report.py \
  --input sample_input.csv \
  --output-dir run_output \
  --test-email your.email@example.com \
  --use-ai
```

## Optional email sending

Sending is off by default. Configure an SMTP account, keep a test address set, then add `--send`:

```bash
export SMTP_HOST="smtp.gmail.com"
export SMTP_PORT="587"
export SMTP_USERNAME="sender@gmail.com"
export SMTP_PASSWORD="app-password"
export SMTP_FROM="sender@gmail.com"
python3 hotel_weekly_report.py \
  --input sample_input.csv \
  --output-dir run_output \
  --test-email your.email@example.com \
  --use-ai \
  --send
```

The script refuses `--send` unless `--test-email` is present. This keeps all messages redirected to one controlled inbox during testing.

## n8n setup

1. Import `hotel_weekly_report_n8n.json` into n8n.
2. In **Read hotel rows**, select Google Sheets credentials, your spreadsheet, and `Sheet1`.
3. In **Generate message**, select an OpenAI credential. The request body uses the Responses API and structured output.
4. In **Send test email**, select Gmail credentials.
5. Open **Calculate benchmarks** and replace `YOUR_TEST_EMAIL@example.com` with your own email. Leave `TEST_MODE = true` for the first run.
6. Run manually and inspect every message in the test inbox.
7. When approved, change `TEST_MODE` to `false`, verify that all owner emails are valid, set the workflow timezone, and publish the workflow.

The Schedule Trigger runs at 08:00 every Monday. n8n uses the workflow timezone when set, otherwise the instance timezone.

## Classification rule

- **Doing well**: revenue is at or above the peer median and orders are at or above the peer median.
- **Needs improvement**: either metric is below its peer median.
- A zero-revenue and zero-order row is flagged as no activity and remains **Needs improvement**.

Using both metrics prevents a high average booking value or many low-value bookings from hiding weakness in the other measure. Peer medians are transparent and robust to extreme values. When real targets become available, the recommended next version is to add `revenue_goal` and `order_goal` columns and score against each hotel's own plan rather than only against peers.

## Input schema

The workflow accepts these headers from the supplied mock data:

| Column | Meaning |
|---|---|
| `brandname` | Hotel display name |
| `revenue` | Previous week's revenue |
| `Order` | Previous week's booking/order count |
| `Boss Gmail` | Owner email |

Rows with a missing hotel name, invalid metrics, or missing email are rejected by the local validator and should be fixed before a live send.

