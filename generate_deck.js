/**
 * Azure FinOps Service Review Deck Generator
 * Reads report_data.json produced by agent.py and renders a polished .pptx
 *
 * Palette: Midnight Executive
 *   Navy:     1E2761
 *   Ice Blue: CADCFC
 *   White:    FFFFFF
 *   Accent:   0D9488 (teal)
 *   Warning:  E85D04 (amber)
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ---------------------------------------------------------------------------
// Load report data
// ---------------------------------------------------------------------------
const raw = fs.readFileSync("report_data.json", "utf8");
const report = JSON.parse(raw);

// ---------------------------------------------------------------------------
// Detect currency from report
// ---------------------------------------------------------------------------
const currencySymbol = (() => {
  const c = (report.currency || "CHF").toUpperCase();
  if (c === "GBP") return "£";
  if (c === "USD") return "$";
  if (c === "EUR") return "€";
  return c + " ";  // e.g. "CHF 1,234"
})();

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function fmt(val) {
  if (val == null || isNaN(val)) return `${currencySymbol}0`;
  return `${currencySymbol}${Number(val).toLocaleString("en", { maximumFractionDigits: 0 })}`;
}

function pct(val) {
  if (val == null || isNaN(val)) return "0%";
  return `${Number(val).toFixed(1)}%`;
}

function safeStr(val, fallback = "N/A") {
  return val != null ? String(val) : fallback;
}

// Brand colours
const DARK = "1B2A4A";
const ACCENT = "0078D4";
const GREEN = "107C10";
const RED = "D13438";
const GREY = "F3F3F3";
const WHITE = "FFFFFF";

// ---------------------------------------------------------------------------
// Create presentation
// ---------------------------------------------------------------------------
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Azure FinOps Agent";
pres.subject = "Monthly Service Review";

// ---------------------------------------------------------------------------
// Slide 1: Title
// ---------------------------------------------------------------------------
const slide1 = pres.addSlide();
slide1.background = { fill: DARK };
slide1.addText("Azure FinOps\nService Review", {
  x: 0.8, y: 1.0, w: 8, h: 2.5,
  fontSize: 36, fontFace: "Segoe UI", color: WHITE, bold: true,
  lineSpacingMultiple: 1.2,
});
slide1.addText(safeStr(report.period, "Monthly Report"), {
  x: 0.8, y: 3.5, w: 5, h: 0.6,
  fontSize: 20, fontFace: "Segoe UI", color: "7EB8FF",
});
slide1.addText(`Generated: ${safeStr(report.generated, new Date().toISOString().split("T")[0])}`, {
  x: 0.8, y: 4.2, w: 5, h: 0.4,
  fontSize: 12, fontFace: "Segoe UI", color: "AAAAAA",
});

// ---------------------------------------------------------------------------
// Slide 2: Executive Summary
// ---------------------------------------------------------------------------
const slide2 = pres.addSlide();
slide2.addText("Executive Summary", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 24, fontFace: "Segoe UI", color: DARK, bold: true,
});

const es = report.executive_summary || {};

const kpiData = [
  { label: "Total Spend", value: fmt(es.total_spend), color: DARK },
  { label: "Budget", value: es.budget ? fmt(es.budget) : "Not Set", color: ACCENT },
  { label: "Budget Used", value: es.budget ? pct(es.budget_utilisation_pct) : "N/A", color: (es.budget_utilisation_pct || 0) > 90 ? RED : GREEN },
  { label: "MoM Change", value: pct(es.mom_change_pct), color: (es.mom_change_pct || 0) > 10 ? RED : GREEN },
  { label: "Advisor Savings", value: fmt(es.total_advisor_savings), color: GREEN },
];

kpiData.forEach((kpi, i) => {
  const xPos = 0.5 + i * 1.85;
  slide2.addShape(pres.ShapeType.roundRect, {
    x: xPos, y: 1.2, w: 1.7, h: 1.4,
    fill: { color: GREY }, rectRadius: 0.1,
  });
  slide2.addText(kpi.label, {
    x: xPos, y: 1.3, w: 1.7, h: 0.4,
    fontSize: 10, fontFace: "Segoe UI", color: "666666", align: "center",
  });
  slide2.addText(kpi.value, {
    x: xPos, y: 1.7, w: 1.7, h: 0.6,
    fontSize: 18, fontFace: "Segoe UI", color: kpi.color, bold: true, align: "center",
  });
});

// Trend & concern
const trendDir = safeStr(es.trend_direction, "stable");
const trendIcon = trendDir === "increasing" ? "📈" : trendDir === "decreasing" ? "📉" : "➡️";
slide2.addText(`${trendIcon}  Trend: ${trendDir.charAt(0).toUpperCase() + trendDir.slice(1)}`, {
  x: 0.5, y: 2.9, w: 9, h: 0.4,
  fontSize: 14, fontFace: "Segoe UI", color: DARK,
});

if (es.top_concern) {
  slide2.addText(`⚠️  ${es.top_concern}`, {
    x: 0.5, y: 3.4, w: 9, h: 0.5,
    fontSize: 12, fontFace: "Segoe UI", color: RED,
  });
}

// ---------------------------------------------------------------------------
// Slide 3: Cost Breakdown by Service
// ---------------------------------------------------------------------------
const slide3 = pres.addSlide();
slide3.addText("Cost Breakdown by Service", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 24, fontFace: "Segoe UI", color: DARK, bold: true,
});

const costRows = (report.cost_breakdown || []).slice(0, 10).map((item) => {
  const trendColor = item.trend === "up" ? RED : item.trend === "down" ? GREEN : "666666";
  const trendArrow = item.trend === "up" ? "▲" : item.trend === "down" ? "▼" : "●";
  return [
    { text: safeStr(item.service, "Unknown"), options: { fontSize: 10, fontFace: "Segoe UI", color: DARK } },
    { text: fmt(item.cost), options: { fontSize: 10, fontFace: "Segoe UI", color: DARK, align: "right" } },
    { text: fmt(item.previous), options: { fontSize: 10, fontFace: "Segoe UI", color: "888888", align: "right" } },
    { text: `${trendArrow} ${safeStr(item.trend)}`, options: { fontSize: 10, fontFace: "Segoe UI", color: trendColor, align: "center" } },
  ];
});

if (costRows.length > 0) {
  slide3.addTable(
    [
      [
        { text: "Service", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI" } },
        { text: "Current", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "right" } },
        { text: "Previous", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "right" } },
        { text: "Trend", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "center" } },
      ],
      ...costRows,
    ],
    { x: 0.5, y: 1.1, w: 9, colW: [4, 1.8, 1.8, 1.4], rowH: 0.35, border: { pt: 0.5, color: "CCCCCC" } }
  );
}

// ---------------------------------------------------------------------------
// Slide 4: Team Breakdown
// ---------------------------------------------------------------------------
const teamData = report.team_breakdown || [];
if (teamData.length > 0) {
  const slide4 = pres.addSlide();
  slide4.addText("Cost by Team", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 24, fontFace: "Segoe UI", color: DARK, bold: true,
  });

  const teamRows = teamData.slice(0, 10).map((item) => {
    const changePct = item.change_pct != null ? item.change_pct : 0;
    const changeColor = changePct > 10 ? RED : changePct < -5 ? GREEN : "666666";
    return [
      { text: safeStr(item.team, "Unknown"), options: { fontSize: 10, fontFace: "Segoe UI", color: DARK } },
      { text: fmt(item.cost), options: { fontSize: 10, fontFace: "Segoe UI", color: DARK, align: "right" } },
      { text: fmt(item.previous), options: { fontSize: 10, fontFace: "Segoe UI", color: "888888", align: "right" } },
      { text: pct(changePct), options: { fontSize: 10, fontFace: "Segoe UI", color: changeColor, align: "center" } },
    ];
  });

  slide4.addTable(
    [
      [
        { text: "Team", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI" } },
        { text: "Current", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "right" } },
        { text: "Previous", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "right" } },
        { text: "Change", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "center" } },
      ],
      ...teamRows,
    ],
    { x: 0.5, y: 1.1, w: 9, colW: [3.5, 2, 2, 1.5], rowH: 0.35, border: { pt: 0.5, color: "CCCCCC" } }
  );
}

// ---------------------------------------------------------------------------
// Slide 5: Monthly Trend
// ---------------------------------------------------------------------------
const trendData = report.monthly_trend || [];
if (trendData.length > 0) {
  const slide5 = pres.addSlide();
  slide5.addText("Monthly Cost Trend", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 24, fontFace: "Segoe UI", color: DARK, bold: true,
  });

  const chartData = [
    {
      name: "Cost",
      labels: trendData.map((d) => safeStr(d.month)),
      values: trendData.map((d) => d.cost || 0),
    },
  ];

  slide5.addChart(pres.ChartType.bar, chartData, {
    x: 0.5, y: 1.1, w: 9, h: 3.8,
    showValue: true,
    valueBarColors: true,
    chartColors: [ACCENT],
    catAxisLabelFontSize: 9,
    valAxisLabelFontSize: 9,
    dataLabelFontSize: 8,
  });
}

// ---------------------------------------------------------------------------
// Slide 6: Recommendations
// ---------------------------------------------------------------------------
const recs = (report.recommendations || []).slice(0, 8);
if (recs.length > 0) {
  const slide6 = pres.addSlide();
  slide6.addText("Optimisation Recommendations", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 24, fontFace: "Segoe UI", color: DARK, bold: true,
  });

  const recRows = recs.map((r) => {
    const prioColor = r.priority === "High" ? RED : r.priority === "Medium" ? "FF8C00" : GREEN;
    return [
      { text: safeStr(r.priority, "Medium"), options: { fontSize: 9, fontFace: "Segoe UI", color: prioColor, bold: true, align: "center" } },
      { text: safeStr(r.title, "Recommendation"), options: { fontSize: 9, fontFace: "Segoe UI", color: DARK } },
      { text: fmt(r.saving), options: { fontSize: 9, fontFace: "Segoe UI", color: GREEN, align: "right", bold: true } },
      { text: safeStr(r.action, ""), options: { fontSize: 8, fontFace: "Segoe UI", color: "666666" } },
    ];
  });

  slide6.addTable(
    [
      [
        { text: "Priority", options: { bold: true, fontSize: 9, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "center" } },
        { text: "Recommendation", options: { bold: true, fontSize: 9, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI" } },
        { text: "Annual Saving", options: { bold: true, fontSize: 9, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "right" } },
        { text: "Action", options: { bold: true, fontSize: 9, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI" } },
      ],
      ...recRows,
    ],
    { x: 0.3, y: 1.1, w: 9.4, colW: [1, 3, 1.5, 3.9], rowH: 0.35, border: { pt: 0.5, color: "CCCCCC" } }
  );
}

// ---------------------------------------------------------------------------
// Slide 7: Resource Inventory & Tagging
// ---------------------------------------------------------------------------
const inv = report.inventory_highlights || {};
const slide7 = pres.addSlide();
slide7.addText("Resource Inventory & Tagging", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 24, fontFace: "Segoe UI", color: DARK, bold: true,
});

slide7.addText(`Total Resources: ${inv.total || 0}`, {
  x: 0.5, y: 1.1, w: 4, h: 0.4,
  fontSize: 16, fontFace: "Segoe UI", color: DARK, bold: true,
});

const taggingPct = inv.tagging_pct || 0;
const taggingColor = taggingPct >= 90 ? GREEN : taggingPct >= 70 ? "FF8C00" : RED;
slide7.addText(`Tagging Compliance: ${taggingPct}%`, {
  x: 5, y: 1.1, w: 4.5, h: 0.4,
  fontSize: 16, fontFace: "Segoe UI", color: taggingColor, bold: true,
});

const topTypes = inv.top_types || [];
if (topTypes.length > 0) {
  const typeRows = topTypes.slice(0, 8).map((t) => {
    const name = typeof t === "string" ? t : safeStr(t.type || t.name, "Unknown");
    const count = typeof t === "object" ? (t.count || "") : "";
    return [
      { text: name, options: { fontSize: 10, fontFace: "Segoe UI", color: DARK } },
      { text: String(count), options: { fontSize: 10, fontFace: "Segoe UI", color: DARK, align: "right" } },
    ];
  });

  slide7.addTable(
    [
      [
        { text: "Resource Type", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI" } },
        { text: "Count", options: { bold: true, fontSize: 10, fill: { color: DARK }, color: WHITE, fontFace: "Segoe UI", align: "right" } },
      ],
      ...typeRows,
    ],
    { x: 0.5, y: 1.7, w: 5, colW: [3.5, 1.5], rowH: 0.32, border: { pt: 0.5, color: "CCCCCC" } }
  );
}

// ---------------------------------------------------------------------------
// Slide 8: Next Steps
// ---------------------------------------------------------------------------
const nextSteps = report.next_steps || [];
if (nextSteps.length > 0) {
  const slide8 = pres.addSlide();
  slide8.addText("Next Steps", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 24, fontFace: "Segoe UI", color: DARK, bold: true,
  });

  const stepsText = nextSteps
    .map((s, i) => `${i + 1}.  ${safeStr(s)}`)
    .join("\n\n");

  slide8.addText(stepsText, {
    x: 0.5, y: 1.1, w: 9, h: 4,
    fontSize: 13, fontFace: "Segoe UI", color: DARK, lineSpacingMultiple: 1.3,
    valign: "top",
  });
}

// ---------------------------------------------------------------------------
// Save
// ---------------------------------------------------------------------------
const filename = `FinOps_Report_${(report.period || "Report").replace(/\s+/g, "_")}.pptx`;
pres.writeFile({ fileName: filename }).then(() => {
  console.log(`✅ Deck saved: ${filename}`);
}).catch((err) => {
  console.error(`❌ Failed to save deck: ${err.message}`);
  process.exit(1);
});
