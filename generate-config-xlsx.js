/* Generates a starter agentic-dashboard-config.xlsx in this folder:
   one sheet per Division+Role. Run with:  node generate-config-xlsx.js
   (The dashboard's in-app "⬇ Template" button generates the same workbook
   from the live configs — this just seeds an editable file in the folder.) */
const XLSX = require("./xlsx.full.min.js");
const DC = require("./dashboard-config.js");

// Faithful MSL sheet (the live dashboard keeps the rich MSL behaviour; this is its editable mirror).
const MSL_STAGES = [
  { key: "d1", col: "Ident.", tile: "HCPs Identified", sub: "Discovered, Scored & Ranked", name: "KOL Identification & Scoring", tasks: ["Query KOL databases", "Score candidates", "Rank by composite score", "Compile ranked KOL list"] },
  { key: "d2", col: "360° Profile", tile: "Profiles Built", sub: "Aggregated Across 8 Dimensions", name: "HCP 360° Profile Build", tasks: ["Aggregate 8 profile dimensions", "Structure profile", "Score relationship maturity", "Generate artifacts"] },
  { key: "d3", col: "Outreach", tile: "Outreach Scheduled", sub: "Emails Sent · Meetings Booked", name: "Outreach Scheduling", tasks: ["Identify optimal slots", "Draft outreach email", "Send via CRM", "Confirm meeting"] },
  { key: "d4", col: "Pre-Call", tile: "Pre-Call Briefs", sub: "MLR-Cleared · Ready for Visit", name: "Pre-Call Briefing", tasks: ["Profile refresh", "MLR-cleared key messages", "Competitor table", "Compile pre-call brief"] },
  { key: "d5", col: "In-Call", tile: "In-Call Support", sub: "Live MLR · Notes · Signals", name: "In-Call Support & Capture", tasks: [] },
  { key: "d6", col: "Post-Call", tile: "Post-Call Reports", sub: "Minutes · CRM · MI/PV Routing", name: "Post-Call Documentation", tasks: [] },
  { key: "d7", col: "Insights", tile: "Insights Synthesised", sub: "Clustered · Heat Map", name: "Insight Synthesis & Taxonomy", tasks: [] },
  { key: "d8", col: "Strategy", tile: "Strategy Feedback", sub: "Prioritised · Routed", name: "Medical Strategy Feedback", tasks: [] },
  { key: "d9", col: "Advisory", tile: "Advisory Boards", sub: "Concept · Panel · Materials", name: "Advisory Board Convening", tasks: [] },
  { key: "d10", col: "IIS/RWE", tile: "IIS / RWE", sub: "Protocol · Budget · Package", name: "IIS / RWE Initiation", tasks: [] },
  { key: "d11", col: "Rel. Health", tile: "Relationship Health", sub: "Cadence · AQS · Alerts", name: "Relationship Health Monitoring", tasks: [] },
  { key: "d12", col: "Feedback", tile: "Feedback Loop", sub: "Closed-loop · Next Cycle", name: "Cyclical Feedback Loop", tasks: [] },
];
const MSL_IMPACT = [
  { head: "KOL Identification", big: "↓ ~50%", dir: "green", d: "faster KOL discovery", s: "multi-source automated" },
  { head: "Profile Build", big: "↓ ~60%", dir: "green", d: "less profiling time", s: "8 dimensions auto-aggregated" },
  { head: "Outreach", big: "↑ 2–4×", dir: "blue", d: "more HCP touchpoints", s: "drafted & scheduled" },
  { head: "Pre-Call Brief", big: "↑ +30%", dir: "blue", d: "brief quality score", s: "MLR pre-check automated" },
  { head: "In-Call & Capture", big: "↑ ~3×", dir: "blue", d: "insights captured / call", s: "ambient NLP · MIR & AE flagged" },
  { head: "Reporting", big: "↓ ~70%", dir: "green", d: "reduction in reporting", s: "8 hr/mo → 2 hr HITL" },
  { head: "Insight → Strategy", big: "↓ ~40%", dir: "green", d: "insight-to-action time", s: "auto-synthesised & routed" },
];
const MSL_SUBJECTS = [
  ["Dr. Anita Rao", "Hematology", "PNH", 2, "approval"],
  ["Dr. Liu Wen", "Hematology", "PNH", 2, "awaiting"],
  ["Dr. Marco Bianchi", "Hematology", "Aplastic Anemia", 2, "pending"],
  ["Dr. Sarah Chen", "Cardiology", "HFrEF", 3, "approval"],
  ["Dr. Elena Fischer", "Hematology", "PNH", 7, "progress"],
  ["Dr. Priya Nair", "Hematology", "PNH", 2, "awaiting"],
  ["Dr. Tomas Vidal", "Hematology", "Aplastic Anemia", 2, "pending"],
  ["Dr. Hana Kim", "Cardiology", "HFrEF", 3, "pending"],
  ["Dr. Omar Haddad", "Cardiology", "HFrEF", 2, "pending"],
  ["Dr. Brodsky RA", "", "", 1, "pending"],
].map((s) => ({ name: s[0], area: s[1], category: s[2], journey: DC.buildJourney(MSL_STAGES, DC.mk(MSL_STAGES.length, s[3], s[4]), "MSL") }));
const MSL_CFG = {
  mode: "generic", code: "MSL", div: "med",
  title: "MSL Activity Dashboard",
  subtitle: "Medical Science Liaisons (MSL) · MSL–HCP 360° Engagement Flow · live status",
  subjectNoun: "HCP", subjectPlural: "HCPs",
  colSubject: "HCP", colArea: "Therapeutic Area", colCategory: "Disease",
  assistant: "Hi! I'm synced to the HCP you select on the left. Ask about their status, score, outreach email or brief — or type approve to action a pending request.",
  footerLabel: "👤 Agentic MSL Employee",
  stages: MSL_STAGES, impact: MSL_IMPACT, subjects: MSL_SUBJECTS,
};
MSL_CFG.agents = DC.buildAgents(MSL_CFG);

const wb = XLSX.utils.book_new();
Object.keys(DC.ROLE_INDEX).forEach((code) => {
  const cfg = code === "MSL" ? MSL_CFG : DC.buildRoleConfig(code);
  if (!cfg) return;
  const ws = XLSX.utils.aoa_to_sheet(DC.configToAOA(cfg));
  ws["!cols"] = [{ wch: 22 }, { wch: 18 }, { wch: 22 }, { wch: 28 }, { wch: 30 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, ws, (cfg.div + "_" + code).slice(0, 31));
});
const buf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
require("fs").writeFileSync("agentic-dashboard-config.xlsx", buf);
console.log("Wrote agentic-dashboard-config.xlsx with " + wb.SheetNames.length + " sheets: " + wb.SheetNames.join(", "));
