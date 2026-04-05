"""
Azure FinOps Report Agent
Queries Azure Cost Management, Advisor, and Resource Graph,
then uses Claude to analyse and generate a PowerPoint deck.

Authentication: Uses DefaultAzureCredential (Managed Identity, CLI, SP via env vars)
  pip install azure-identity azure-mgmt-costmanagement azure-mgmt-advisor azure-mgmt-resourcegraph anthropic
"""

import json
import subprocess
import os
import time
from datetime import datetime, timedelta, timezone
from dateutil.relativedelta import relativedelta
import anthropic

from azure.identity import DefaultAzureCredential
from azure.mgmt.costmanagement import CostManagementClient
from azure.mgmt.costmanagement.models import (
    QueryDefinition,
    QueryTimePeriod,
    QueryDataset,
    QueryAggregation,
    QueryGrouping,
)
from azure.mgmt.advisor import AdvisorManagementClient
from azure.mgmt.resourcegraph import ResourceGraphClient
from azure.mgmt.resourcegraph.models import (
    QueryRequest,
    QueryRequestOptions,
)

# ---------------------------------------------------------------------------
# CLAUDE CLIENT
# ---------------------------------------------------------------------------
# Set this environment variable:
#   ANTHROPIC_API_KEY = "sk-ant-..."

credential = DefaultAzureCredential()

client = anthropic.Anthropic()

MODEL_NAME = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-20250514")


# ---------------------------------------------------------------------------
# HELPER: resolve time range to start/end dates
# ---------------------------------------------------------------------------

def _resolve_time_range(time_range: str) -> tuple[datetime, datetime]:
    """
    Translates a plain-English time range (e.g. "last_month") into exact
    calendar dates that the Azure billing APIs understand.

    For example, if today is 5 April 2026 and you ask for "last_month",
    this returns 1 March 2026 → 31 March 2026.
    """
    now = datetime.now(timezone.utc)
    if time_range == "last_month":
        first_of_this_month = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        end = first_of_this_month - timedelta(seconds=1)
        start = (first_of_this_month - relativedelta(months=1))
    elif time_range == "last_quarter":
        first_of_this_month = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        end = first_of_this_month - timedelta(seconds=1)
        start = first_of_this_month - relativedelta(months=3)
    elif time_range == "last_30_days":
        end = now
        start = now - timedelta(days=30)
    else:
        end = now
        start = now - timedelta(days=30)
    return start, end


def _previous_period(start: datetime, end: datetime) -> tuple[datetime, datetime]:
    """
    Works out the equivalent earlier time window so we can compare
    "this period vs last period". If we're looking at March, this
    gives us February — so we can show month-over-month changes.
    """
    delta = end - start
    return start - delta, start - timedelta(seconds=1)


# ---------------------------------------------------------------------------
# AZURE FUNCTIONS
# ---------------------------------------------------------------------------

def get_cost_by_scope(scope: str, time_range: str, group_by: str) -> dict:
    """
    Fetches a cost breakdown from Azure Cost Management.

    What it does:
    1. Connects to Azure's billing system for the given subscription
    2. Pulls the total spend for the requested time period (e.g. last month)
    3. Also pulls the previous period's spend for comparison
    4. Groups costs by service (e.g. Virtual Machines, Storage) or by team tag
    5. Calculates whether each group's spend is trending up, down, or stable
    6. Optionally checks if there's a budget set, so we can show utilisation

    Returns a dictionary with total cost, previous period cost, budget,
    and a breakdown showing each service/team's spend and trend.
    """
    cost_client = CostManagementClient(credential)
    start, end = _resolve_time_range(time_range)
    prev_start, prev_end = _previous_period(start, end)

    # Map the user-friendly group name to Azure's internal field name.
    # "service" groups by Azure service type (e.g. VMs, Storage, SQL).
    # "team" and "environment" group by resource tags that your team
    # should have applied to each Azure resource for cost tracking.
    group_dimension = {
        "service": "ServiceName",
        "team": "TagKey:team",
        "environment": "TagKey:environment",
    }.get(group_by, "ServiceName")

    if group_dimension.startswith("TagKey:"):
        tag_key = group_dimension.split(":")[1]
        grouping = [QueryGrouping(type="TagKey", name=tag_key)]
    else:
        grouping = [QueryGrouping(type="Dimension", name=group_dimension)]

    # This inner function sends the actual billing query to Azure.
    # It includes automatic retry logic — Azure sometimes rate-limits
    # requests (returns a 429 error), so we wait and try again.
    def _query_costs_with_retry(s: datetime, e: datetime, max_retries: int = 3) -> list[dict]:
        query = QueryDefinition(
            type="ActualCost",
            timeframe="Custom",
            time_period=QueryTimePeriod(from_property=s, to=e),
            dataset=QueryDataset(
                granularity="None",
                aggregation={
                    "totalCost": QueryAggregation(name="Cost", function="Sum"),
                    "totalCostUSD": QueryAggregation(name="CostUSD", function="Sum"),
                },
                grouping=grouping,
            ),
        )
        for attempt in range(max_retries):
            try:
                result = cost_client.query.usage(scope=scope, parameters=query)
                rows = []
                columns = [col.name for col in result.columns]
                for row in result.rows:
                    rows.append(dict(zip(columns, row)))
                return rows
            except Exception as ex:
                if "429" in str(ex) and attempt < max_retries - 1:
                    wait = (attempt + 1) * 10
                    print(f"  ⏳ Rate limited, waiting {wait}s before retry...")
                    time.sleep(wait)
                else:
                    raise
        return []

    # Fetch this period's costs and last period's costs.
    # The small delay between calls avoids Azure's rate limit.
    current_rows = _query_costs_with_retry(start, end)
    time.sleep(2)
    previous_rows = _query_costs_with_retry(prev_start, prev_end)

    # Build a lookup table of last period's costs so we can quickly
    # compare each service/team's current spend vs previous spend.
    prev_lookup = {}
    for row in previous_rows:
        group_key = None
        for k, v in row.items():
            if k not in ("Cost", "CostUSD", "Currency"):
                group_key = v
                break
        if group_key:
            prev_lookup[group_key] = row.get("Cost", 0.0)

    total_cost = sum(r.get("Cost", 0.0) for r in current_rows)
    previous_total = sum(prev_lookup.values())

    # Build the breakdown: for each service or team, record the current
    # cost, the previous period cost, and whether spend is going up or down.
    # A change of more than 5% is flagged as "up" or "down".
    breakdown = []
    for row in current_rows:
        group_key = None
        for k, v in row.items():
            if k not in ("Cost", "CostUSD", "Currency"):
                group_key = v
                break
        cost = row.get("Cost", 0.0)
        prev_cost = prev_lookup.get(group_key, 0.0)
        if prev_cost > 0:
            pct_change = ((cost - prev_cost) / prev_cost) * 100
            trend = "up" if pct_change > 5 else ("down" if pct_change < -5 else "stable")
        else:
            trend = "new"
        breakdown.append({
            "name": group_key or "Unknown",
            "cost": round(cost, 2),
            "previous": round(prev_cost, 2),
            "trend": trend,
        })

    breakdown.sort(key=lambda x: x["cost"], reverse=True)

    # Try to fetch the subscription's budget (if one has been set in Azure).
    # This lets us report "you've used X% of your budget".
    # If no budget exists or the API isn't available, we just skip it.
    budget = None
    try:
        from azure.mgmt.consumption import ConsumptionManagementClient
        consumption_client = ConsumptionManagementClient(credential)
        budgets = list(consumption_client.budgets.list(scope=scope))
        if budgets:
            budget = budgets[0].amount
    except Exception:
        budget = None

    result = {
        "scope": scope,
        "time_range": time_range,
        "currency": current_rows[0].get("Currency", "GBP") if current_rows else "GBP",
        "total_cost": round(total_cost, 2),
        "previous_period_cost": round(previous_total, 2),
        "budget": budget,
        "breakdown": breakdown,
    }

    if group_by == "team":
        result["by_team"] = [
            {"team": item["name"], "cost": item["cost"], "previous": item["previous"]}
            for item in breakdown
        ]

    return result


def get_cost_trend(scope: str, months: int) -> dict:
    """
    Fetches the month-by-month cost trend over the last N months.

    What it does:
    1. Asks Azure for monthly totals going back the requested number of months
    2. Returns a list like [{"month": "Jan 2026", "cost": 1200}, ...]
    3. This data is used to create the bar chart in the PowerPoint deck
       showing whether overall spend is rising, falling, or flat.

    Includes automatic retry if Azure rate-limits the request.
    """
    cost_client = CostManagementClient(credential)
    now = datetime.now(timezone.utc)
    first_of_this_month = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    start = first_of_this_month - relativedelta(months=months)
    end = now

    query = QueryDefinition(
        type="ActualCost",
        timeframe="Custom",
        time_period=QueryTimePeriod(from_property=start, to=end),
        dataset=QueryDataset(
            granularity="Monthly",
            aggregation={
                "totalCost": QueryAggregation(name="Cost", function="Sum"),
            },
        ),
    )

    for attempt in range(3):
        try:
            result = cost_client.query.usage(scope=scope, parameters=query)
            break
        except Exception as ex:
            if "429" in str(ex) and attempt < 2:
                wait = (attempt + 1) * 10
                print(f"  ⏳ Rate limited on trend query, waiting {wait}s...")
                time.sleep(wait)
            else:
                raise

    columns = [col.name for col in result.columns]
    rows = [dict(zip(columns, row)) for row in result.rows]

    monthly = []
    for row in rows:
        date_val = row.get("BillingMonth") or row.get("UsageDate")
        cost_val = row.get("Cost", 0.0)
        if date_val:
            try:
                if isinstance(date_val, (int, float)):
                    dt = datetime.strptime(str(int(date_val)), "%Y%m%d")
                else:
                    dt = datetime.fromisoformat(str(date_val))
                month_label = dt.strftime("%b %Y")
            except (ValueError, TypeError):
                month_label = str(date_val)
        else:
            month_label = "Unknown"
        monthly.append({"month": month_label, "cost": round(cost_val, 2)})

    monthly.sort(key=lambda x: x["month"])

    currency = "GBP"
    if rows and "Currency" in rows[0]:
        currency = rows[0]["Currency"]

    return {"scope": scope, "currency": currency, "monthly": monthly}


def get_advisor_recommendations(scope: str) -> dict:
    """
    Pulls cost-saving and reliability recommendations from Azure Advisor.

    What it does:
    1. Connects to Azure Advisor — Microsoft's built-in recommendation engine
    2. Filters for two categories:
       - Cost: "You could save money by doing X" (e.g. resize an underused VM)
       - High Availability: "This resource has a reliability risk"
    3. For each recommendation, extracts:
       - Priority (High / Medium / Low)
       - What the problem is and what action to take
       - Estimated annual saving in currency
    4. Sorts recommendations by biggest saving first

    This is essentially Azure telling you where you're wasting money.
    """
    sub_id = scope.replace("/subscriptions/", "").split("/")[0]
    advisor_client = AdvisorManagementClient(credential, sub_id)

    recommendations_raw = list(advisor_client.recommendations.list(
        filter="Category eq 'Cost' or Category eq 'HighAvailability'"
    ))

    total_savings = 0.0
    recommendations = []

    for rec in recommendations_raw:
        ext = rec.extended_properties or {}
        annual_saving = 0.0
        try:
            saving_str = (
                ext.get("annualSavingsAmount")
                or ext.get("savingsAmount")
                or ext.get("estimatedAnnualSavings")
                or "0"
            )
            annual_saving = float(saving_str)
        except (ValueError, TypeError):
            annual_saving = 0.0

        total_savings += annual_saving
        impact_map = {"High": "High", "Medium": "Medium", "Low": "Low"}
        priority = impact_map.get(str(rec.impact), "Medium")

        recommendations.append({
            "category": str(rec.category) if rec.category else "Cost",
            "priority": priority,
            "title": rec.short_description.problem if rec.short_description else "Recommendation",
            "description": rec.short_description.solution if rec.short_description else "",
            "estimated_annual_saving": round(annual_saving, 2),
            "affected_resources": 1,
            "action": rec.short_description.solution if rec.short_description else "",
        })

    recommendations.sort(key=lambda x: x["estimated_annual_saving"], reverse=True)

    return {
        "total_estimated_annual_savings": round(total_savings, 2),
        "currency": "GBP",
        "recommendations": recommendations,
    }


def get_resource_inventory(scope: str) -> dict:
    """
    Counts and categorises all Azure resources in the subscription.

    What it does:
    1. Queries Azure Resource Graph — a fast search engine across all your
       Azure resources (VMs, databases, storage accounts, etc.)
    2. Counts resources by type (e.g. "42 Virtual Machines, 15 Storage Accounts")
    3. Counts resources by region (e.g. "80 in UK South, 20 in West Europe")
    4. Checks tagging compliance — are resources properly labelled with:
       - cost-centre: which budget pays for this?
       - environment: is this production, dev, or test?
       - owner: who is responsible for this resource?
    5. Reports what percentage of resources have all required tags

    Tagging compliance is important because untagged resources can't be
    tracked back to a team or project, making cost allocation impossible.
    """
    sub_id = scope.replace("/subscriptions/", "").split("/")[0]
    rg_client = ResourceGraphClient(credential)

    # Helper to run a query against Azure Resource Graph.
    # Resource Graph uses a SQL-like language called KQL (Kusto Query Language)
    # to search across all resources in the subscription.
    def _run_query(query_str: str) -> list[dict]:
        request = QueryRequest(
            subscriptions=[sub_id],
            query=query_str,
            options=QueryRequestOptions(result_format="objectArray"),
        )
        response = rg_client.resources(request)
        return response.data if response.data else []

    # Query 1: How many of each resource type do we have?
    # e.g. "42 Virtual Machines, 15 Storage Accounts, 8 SQL Databases"
    by_type_raw = _run_query(
        "Resources | summarize count_ = count() by type "
        "| order by count_ desc | project type, count_"
    )

    # Translate Azure's internal type names into human-readable labels.
    # e.g. "microsoft.compute/virtualmachines" → "Virtual Machines"
    type_name_map = {
        "microsoft.compute/virtualmachines": "Virtual Machines",
        "microsoft.compute/disks": "Managed Disks",
        "microsoft.storage/storageaccounts": "Storage Accounts",
        "microsoft.sql/servers/databases": "SQL Databases",
        "microsoft.containerservice/managedclusters": "AKS Clusters",
        "microsoft.web/serverfarms": "App Service Plans",
        "microsoft.keyvault/vaults": "Key Vaults",
        "microsoft.network/virtualnetworks": "Virtual Networks",
        "microsoft.network/networkinterfaces": "Network Interfaces",
        "microsoft.network/networksecuritygroups": "NSGs",
        "microsoft.network/publicipaddresses": "Public IPs",
        "microsoft.network/loadbalancers": "Load Balancers",
    }

    total_resources = 0
    by_type = []
    for item in by_type_raw:
        raw_type = item.get("type", "unknown").lower()
        count = item.get("count_", 0)
        total_resources += count
        friendly_name = type_name_map.get(raw_type, raw_type.split("/")[-1].title())
        by_type.append({"type": friendly_name, "count": count})

    # Query 2: Where are our resources located geographically?
    # e.g. "80 in UK South, 20 in West Europe, 10 in East US"
    by_region_raw = _run_query(
        "Resources | summarize count_ = count() by location "
        "| order by count_ desc | project location, count_"
    )

    region_name_map = {
        "uksouth": "UK South", "ukwest": "UK West",
        "westeurope": "West Europe", "northeurope": "North Europe",
        "eastus": "East US", "eastus2": "East US 2",
        "westus": "West US", "westus2": "West US 2", "centralus": "Central US",
    }

    by_region = []
    for item in by_region_raw:
        loc = item.get("location", "unknown")
        friendly = region_name_map.get(loc.lower().replace(" ", ""), loc)
        by_region.append({"region": friendly, "count": item.get("count_", 0)})

    # Query 3: Tagging compliance check.
    # We check each resource for three required tags. Resources without
    # these tags are "untagged" and can't be properly tracked for billing.
    required_tags = ["cost-centre", "environment", "owner"]
    missing_tags = []
    for tag in required_tags:
        tag_result = _run_query(
            f"Resources | where isnull(tags['{tag}']) or tags['{tag}'] == '' "
            f"| summarize count_ = count()"
        )
        count = tag_result[0].get("count_", 0) if tag_result else 0
        if count > 0:
            missing_tags.append(tag)

    # Query 4: Count resources missing ANY of the required tags.
    # This gives us the overall compliance percentage.
    any_missing_result = _run_query(
        "Resources "
        "| where isnull(tags['cost-centre']) or tags['cost-centre'] == '' "
        "or isnull(tags['environment']) or tags['environment'] == '' "
        "or isnull(tags['owner']) or tags['owner'] == '' "
        "| summarize count_ = count()"
    )
    untagged_resources = any_missing_result[0].get("count_", 0) if any_missing_result else 0

    tagging_pct = (
        round(((total_resources - untagged_resources) / total_resources) * 100)
        if total_resources > 0 else 0
    )

    return {
        "scope": scope,
        "total_resources": total_resources,
        "by_type": by_type[:10],
        "by_region": by_region,
        "tagging_compliance": {
            "percentage": tagging_pct,
            "untagged_resources": untagged_resources,
            "missing_tags": missing_tags,
        },
    }


def raise_report_ready(report_path: str, summary: str) -> dict:
    """
    Sends a notification that the report has been generated.

    If a Microsoft Teams webhook URL is configured, it posts a message
    to the Teams channel so stakeholders know the report is ready.
    If no webhook is set up, this step is silently skipped.
    """
    teams_webhook = os.environ.get("TEAMS_WEBHOOK_URL")
    if teams_webhook:
        import requests
        payload = {
            "@type": "MessageCard",
            "summary": "FinOps Report Ready",
            "sections": [{"activityTitle": "📊 Azure FinOps Report Ready", "text": summary,
                          "facts": [{"name": "Report", "value": report_path}]}],
        }
        try:
            requests.post(teams_webhook, json=payload, timeout=10)
        except Exception:
            pass
    return {"status": "success", "message": f"Report ready at {report_path}"}


# ---------------------------------------------------------------------------
# TOOL DEFINITIONS (Anthropic format)
# ---------------------------------------------------------------------------

tools = [
    {
        "name": "get_cost_by_scope",
        "description": "Get Azure cost breakdown for a subscription or resource group scope.",
        "input_schema": {
            "type": "object",
            "properties": {
                "scope": {"type": "string", "description": "Full ARM scope, e.g. /subscriptions/<sub-id>"},
                "time_range": {"type": "string", "enum": ["last_month", "last_quarter", "last_30_days"]},
                "group_by": {"type": "string", "enum": ["service", "team", "environment"]}
            },
            "required": ["scope", "time_range", "group_by"]
        },
    },
    {
        "name": "get_cost_trend",
        "description": "Get month-over-month cost trend for the last N months.",
        "input_schema": {
            "type": "object",
            "properties": {
                "scope": {"type": "string", "description": "Full ARM scope"},
                "months": {"type": "integer", "description": "Number of months (3-12)"}
            },
            "required": ["scope", "months"]
        },
    },
    {
        "name": "get_advisor_recommendations",
        "description": "Get Azure Advisor cost optimisation recommendations.",
        "input_schema": {
            "type": "object",
            "properties": {
                "scope": {"type": "string", "description": "Subscription ID or full ARM scope"}
            },
            "required": ["scope"]
        },
    },
    {
        "name": "get_resource_inventory",
        "description": "Get Azure resource inventory by type, region, and tagging compliance.",
        "input_schema": {
            "type": "object",
            "properties": {
                "scope": {"type": "string", "description": "Subscription ID or full ARM scope"}
            },
            "required": ["scope"]
        },
    },
    {
        "name": "raise_report_ready",
        "description": "Signal report is ready. Call as the final step with the complete structured JSON.",
        "input_schema": {
            "type": "object",
            "properties": {
                "report_json": {"type": "string", "description": "Complete structured report data as JSON string"},
                "summary": {"type": "string", "description": "Two sentence executive summary"}
            },
            "required": ["report_json", "summary"]
        },
    },
]


# ---------------------------------------------------------------------------
# TOOL EXECUTOR
# ---------------------------------------------------------------------------

def execute_tool(name: str, inputs: dict) -> str:
    try:
        if name == "get_cost_by_scope":
            result = get_cost_by_scope(**inputs)
        elif name == "get_cost_trend":
            result = get_cost_trend(**inputs)
        elif name == "get_advisor_recommendations":
            result = get_advisor_recommendations(**inputs)
        elif name == "get_resource_inventory":
            result = get_resource_inventory(**inputs)
        elif name == "raise_report_ready":
            result = {"status": "captured", "data": inputs}
        else:
            result = {"error": f"Unknown tool: {name}"}
    except Exception as e:
        result = {"error": str(e), "tool": name}
        print(f"⚠️  Tool '{name}' failed: {e}")
    return json.dumps(result, default=str)


# ---------------------------------------------------------------------------
# AGENT LOOP
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """You are an Azure FinOps analyst agent. Your job is to gather 
cost, usage, and optimisation data from Azure and structure it into a clear, 
executive-ready service review report.

Guidelines:
- Audience is senior stakeholders — focus on business impact, not technical detail
- Always retrieve: current cost breakdown, trend data, Advisor recommendations, and resource inventory
- Calculate month-over-month change as a percentage
- Flag anything that represents >10% increase as a concern
- Prioritise recommendations by annual saving value, highest first
- Note budget utilisation (spend vs budget)
- When you have all data, call raise_report_ready with a fully structured JSON report
- Use full ARM scope format: /subscriptions/<sub-id>

The report JSON must have this structure:
{
  "period": "March 2026",
  "generated": "ISO date",
  "executive_summary": {
    "total_spend": float,
    "budget": float,
    "budget_utilisation_pct": float,
    "mom_change_pct": float,
    "mom_change_gbp": float,
    "trend_direction": "increasing|stable|decreasing",
    "top_concern": "string",
    "total_advisor_savings": float
  },
  "cost_breakdown": [{"service": str, "cost": float, "previous": float, "trend": str}],
  "team_breakdown": [{"team": str, "cost": float, "previous": float, "change_pct": float}],
  "monthly_trend": [{"month": str, "cost": float}],
  "recommendations": [{"priority": str, "title": str, "saving": float, "action": str}],
  "inventory_highlights": {"total": int, "top_types": list, "tagging_pct": int},
  "next_steps": ["string"]
}"""


def run_agent(subscription_id: str) -> dict:
    """
    The main agent loop — this is the "brain" of the system.

    How it works:
    1. We send Claude an initial instruction: "generate a FinOps report for this subscription"
    2. Claude decides which Azure tools to call (cost breakdown, trend, advisor, inventory)
    3. We execute those tool calls against real Azure APIs and return the results to Claude
    4. Claude reads the results, may call more tools, and eventually produces a structured report
    5. The loop repeats until Claude says "I'm done" (no more tool calls)

    This is called an "agentic loop" — Claude autonomously decides what data to gather
    and in what order, rather than us hardcoding the sequence of API calls.
    """
    print(f"\n🤖 FinOps Agent starting for subscription: {subscription_id}\n")

    # Start the conversation with an instruction telling Claude what to do.
    # This is the only human message — from here, Claude drives the process.
    messages = [
        {"role": "user", "content": f"""Generate a service review FinOps report for Azure subscription: {subscription_id}
        
        Use the full ARM scope /subscriptions/{subscription_id} when calling tools.
        Retrieve all cost, trend, recommendation, and inventory data, then call raise_report_ready 
        with the fully structured report JSON."""},
    ]

    # This will hold the final structured report once Claude produces it
    report_data = None

    # --- The Agent Loop ---
    # Each iteration: send messages to Claude → Claude responds (possibly with tool calls)
    # → we execute the tools → feed results back to Claude → repeat
    while True:

        # Send the full conversation history to Claude and get a response.
        # Claude sees: the system prompt (its role), the tools available,
        # and all previous messages including tool results.
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            tools=tools,
            messages=messages,
        )

        # Claude's response contains "blocks" — either text (thinking out loud)
        # or tool_use (requesting to call one of our Azure functions).
        # We print both so you can see what the agent is doing in real time.
        for block in response.content:
            if block.type == "text" and block.text:
                print(f"💭 {block.text[:200]}...")
            elif block.type == "tool_use":
                print(f"🔧 Calling: {block.name}({json.dumps(block.input)[:80]}...)")

        # Extract just the tool calls from Claude's response
        tool_uses = [b for b in response.content if b.type == "tool_use"]

        # If Claude didn't request any tools, it's finished — exit the loop.
        # "end_turn" means Claude has said everything it wants to say.
        if response.stop_reason == "end_turn" or not tool_uses:
            break

        # Add Claude's response to the conversation history.
        # This is important — Claude needs to "remember" what it already said and asked for.
        messages.append({"role": "assistant", "content": response.content})

        # Now execute each tool that Claude requested.
        # Claude might ask for multiple tools at once (e.g. cost + advisor + inventory
        # in parallel) — we run them all and collect the results.
        tool_results = []
        for tool_use in tool_uses:
            # Call the actual Azure function (e.g. get_cost_by_scope)
            result = execute_tool(tool_use.name, tool_use.input)
            result_data = json.loads(result)

            # Special case: when Claude calls raise_report_ready, it's sending us
            # the final structured report. We capture it here so we can save it
            # to report_data.json after the loop ends.
            if tool_use.name == "raise_report_ready" and "data" in result_data:
                report_json_str = result_data["data"].get("report_json", "{}")
                try:
                    report_data = json.loads(report_json_str)
                    print(f"\n✅ Report data captured successfully\n")
                except json.JSONDecodeError as e:
                    print(f"⚠️  JSON parse error: {e}")

            # Package the tool result in the format Claude expects.
            # The tool_use_id links this result back to the specific tool call
            # so Claude knows which request this is answering.
            tool_results.append({
                "type": "tool_result",
                "tool_use_id": tool_use.id,
                "content": result,
            })

        # Feed all tool results back to Claude as a "user" message.
        # Claude will read these results and decide what to do next —
        # either call more tools or produce the final report.
        messages.append({"role": "user", "content": tool_results})

    return report_data


# ---------------------------------------------------------------------------
# ENTRY POINT
# This is what runs when you execute: python agent.py
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Read the Azure subscription ID from an environment variable.
    # This is the subscription whose costs we'll analyse.
    SUBSCRIPTION_ID = os.environ.get("AZURE_SUBSCRIPTION_ID", "sub-xxxxxxxx-demo")

    # Step 1: Run the agent — it queries Azure and produces a structured report
    report = run_agent(SUBSCRIPTION_ID)

    if report:
        # Step 2: Save the report as a JSON file.
        # This is the bridge between Python (data gathering) and Node.js (deck generation).
        with open("report_data.json", "w") as f:
            json.dump(report, f, indent=2)
        print("📄 report_data.json written")

        # Step 3: Run the Node.js script that reads report_data.json
        # and generates a formatted PowerPoint deck (.pptx file).
        result = subprocess.run(
            ["node", "generate_deck.js"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print("✅ Deck generated successfully")
        else:
            print(f"❌ Deck error: {result.stderr}")
    else:
        print("❌ Agent did not return report data")
