# Cost Management Agent

Azure FinOps report agent that:

- Queries Azure Cost Management, Advisor, and Resource Graph
- Uses Claude to analyze findings and build a structured report JSON
- Generates a PowerPoint service review deck from that JSON

## What this project does

1. Collects Azure cost, trend, recommendation, and inventory data
2. Passes tool outputs to Claude for executive-style analysis
3. Writes report data to `report_data.json`
4. Builds a deck using `generate_deck.js`

## Project structure

- `agent.py` - main agent loop and Azure data collection tools
- `generate_deck.js` - PowerPoint generation script
- `report_data.json` - generated report payload used by deck script

## Prerequisites

- Python 3.10+
- Node.js 18+
- Azure CLI installed and authenticated
- Access to an Anthropic API key

## Setup

### 1. Create and activate Python environment

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 2. Install Python dependencies

```bash
pip install anthropic azure-identity azure-mgmt-costmanagement azure-mgmt-advisor azure-mgmt-resourcegraph azure-mgmt-consumption python-dateutil requests
```

### 3. Install Node dependencies

```bash
npm install
```

## Authentication and environment variables

Authenticate Azure locally:

```bash
az login
```

Set required environment variables:

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
export AZURE_SUBSCRIPTION_ID="<your-subscription-id>"
```

Optional:

```bash
export CLAUDE_MODEL="claude-sonnet-4-20250514"
export TEAMS_WEBHOOK_URL="https://..."
```

## Run

```bash
python agent.py
```

Expected output:

- `report_data.json` created
- `FinOps_Report_<period>.pptx` generated

## Common issues

- `429 Too many requests` from Cost Management API:
  - Script includes retry logic, but you may still need to retry later.
- `Cannot find module 'pptxgenjs'`:
  - Run `npm install`.
- Missing Azure access:
  - Ensure your account has at least reader access to required APIs.

## Security notes

- Never commit secrets (`ANTHROPIC_API_KEY`, webhook URLs, credentials).
- Review logs before sharing; they can include subscription identifiers.
- Data is sent to Anthropic for analysis when running the agent.

