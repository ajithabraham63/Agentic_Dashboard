/* =====================================================================
   dashboard-config.js  —  shared by the dashboard (browser) and the
   Excel template generator (Node).  Defines the per-Division / per-Role
   configuration model that drives the generic Agentic Employee dashboard,
   plus the Excel <-> config mapping (one sheet per Division+Role).

   UMD wrapper so it works as a <script> (window.DashboardConfig) and via
   require() in Node (module.exports).
   ===================================================================== */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.DashboardConfig = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var STLBL = { done: "Done", progress: "In Progress", approval: "Pending Approval", awaiting: "Awaiting Response", pending: "Pending" };

  /* ---- helpers -------------------------------------------------------- */
  // Build a "cells" array (one state per stage) from a progress index + active state.
  function mk(n, p, act) {
    var a = [];
    for (var i = 0; i < n; i++) a.push(i < p ? "done" : i === p ? act : "pending");
    return a;
  }
  function defaultDetail(stage, state, prevName) {
    switch (state) {
      case "done": return stage.name + " complete";
      case "approval": return stage.name + " — pending approval";
      case "awaiting": return stage.name + " — awaiting response";
      case "progress": return stage.name + " in progress";
      default: return "Awaits " + (prevName ? prevName : "upstream steps");
    }
  }
  function slugTool(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "").split("_").filter(Boolean).slice(0, 4).join("_"); }
  // Derive an Agent Workspace spec (orchestrator + one agent per stage, tools from tasks).
  function buildAgents(cfg) {
    var list = (cfg.stages || []).map(function (s, i) {
      var tools = (s.tasks && s.tasks.length ? s.tasks.slice(0, 4) : [s.name]).map(slugTool).filter(Boolean);
      if (!tools.length) tools = ["run_" + slugTool(s.name)];
      return { id: s.key + "_agent", label: s.name, cluster: s.col, group: i % 6, tools: tools };
    });
    return { root: { l1: (cfg.code || "agent").toLowerCase().slice(0, 9), l2: "workspace" }, list: list };
  }
  // Build a subject journey ([{key,name,state,detail,facts}]) from stages + cells.
  function buildJourney(stages, cells, ownerLabel) {
    return stages.map(function (s, i) {
      var state = cells[i] || "pending";
      return {
        key: s.key, name: s.name, state: state,
        detail: defaultDetail(s, state, i > 0 ? stages[i - 1].name : ""),
        facts: [["Stage", s.col], ["Status", STLBL[state]], ["Owner", ownerLabel || "Agent"]],
      };
    });
  }

  /* ---- Division themes ----------------------------------------------- */
  // Each theme defines the operating flow (stages), the column labels, the
  // subject vocabulary, sample subjects and the impact metrics for a division.
  var THEMES = {
    sff: {
      flow: "Field Engagement Flow",
      subjectNoun: "Account", subjectPlural: "Accounts",
      colSubject: "Account", colArea: "Region", colCategory: "Specialty",
      stages: [
        { key: "s1", col: "Target", tile: "Targets Identified", sub: "Segmented & prioritised", name: "Account Targeting & Segmentation", tasks: ["Pull territory universe", "Score & segment accounts", "Prioritise call list"] },
        { key: "s2", col: "Plan", tile: "Call Plans Built", sub: "Route & frequency set", name: "Call Planning & Routing", tasks: ["Set call frequency", "Optimise route", "Sync to calendar"] },
        { key: "s3", col: "Schedule", tile: "Visits Scheduled", sub: "Calendar booked", name: "Visit Scheduling", tasks: ["Request slots", "Confirm appointment", "Block calendar"] },
        { key: "s4", col: "Detail", tile: "Details Delivered", sub: "Approved messaging", name: "In-Field Detailing", tasks: ["Present approved messaging", "Capture HCP response", "Handle objections"] },
        { key: "s5", col: "Sample", tile: "Samples Logged", sub: "Compliance tracked", name: "Sampling & Compliance", tasks: ["Log sample drop", "Capture signature", "Sync compliance record"] },
        { key: "s6", col: "Capture", tile: "Calls Captured", sub: "CRM synced", name: "Call Capture & CRM Sync", tasks: ["Draft call notes", "Tag topics & sentiment", "Sync to CRM"] },
        { key: "s7", col: "Follow-up", tile: "Follow-ups Sent", sub: "Next-best-action", name: "Follow-up & Next-Best-Action", tasks: ["Generate NBA", "Send follow-up", "Schedule next touch"] },
      ],
      impact: [
        { head: "Targeting", big: "↓ ~45%", dir: "green", d: "faster territory planning", s: "data-driven segmentation" },
        { head: "Call Planning", big: "↑ ~2×", dir: "blue", d: "more effective calls", s: "route & frequency optimised" },
        { head: "Detailing", big: "↑ +25%", dir: "blue", d: "message adherence", s: "approved content only" },
        { head: "Sampling", big: "↓ ~80%", dir: "green", d: "compliance effort", s: "auto-logged & tracked" },
        { head: "Call Capture", big: "↓ ~70%", dir: "green", d: "admin time per call", s: "ambient notes → CRM" },
        { head: "Follow-up", big: "↑ ~3×", dir: "blue", d: "timely follow-ups", s: "next-best-action engine" },
        { head: "Coaching", big: "↑ +30%", dir: "blue", d: "rep performance", s: "insight-led coaching" },
      ],
      subjects: [
        { name: "Riverside Cardiology Group", area: "Northeast", category: "Cardiology", p: 6, act: "progress" },
        { name: "Dr. Alan Pierce", area: "Northeast", category: "Cardiology", p: 3, act: "approval" },
        { name: "Lakeside Medical Center", area: "Midwest", category: "Hospital", p: 2, act: "awaiting" },
        { name: "Dr. Maria Gomez", area: "West", category: "Endocrinology", p: 4, act: "progress" },
        { name: "Summit Specialty Clinic", area: "West", category: "Rheumatology", p: 7, act: "done" },
        { name: "Dr. Kevin Brooks", area: "South", category: "Cardiology", p: 1, act: "progress" },
        { name: "Harbor Health Network", area: "Northeast", category: "Hospital", p: 0, act: "pending" },
      ],
    },
    mkt: {
      flow: "Omnichannel Campaign Flow",
      subjectNoun: "Campaign", subjectPlural: "Campaigns",
      colSubject: "Campaign", colArea: "Brand", colCategory: "Channel",
      stages: [
        { key: "m1", col: "Brief", tile: "Briefs Created", sub: "Objectives & audience", name: "Campaign Brief", tasks: ["Define objective & KPI", "Define target audience", "Align budget"] },
        { key: "m2", col: "Content", tile: "Assets Drafted", sub: "Creative produced", name: "Content Development", tasks: ["Draft copy", "Produce creative", "Assemble assets"] },
        { key: "m3", col: "MLR", tile: "MLR Cleared", sub: "Med-legal-reg approved", name: "MLR Review", tasks: ["Submit to MLR", "Resolve comments", "Capture approval"] },
        { key: "m4", col: "Channel", tile: "Channels Set", sub: "Omnichannel plan", name: "Channel Planning", tasks: ["Select channels", "Set cadence", "Configure targeting"] },
        { key: "m5", col: "Launch", tile: "Campaigns Live", sub: "Deployed", name: "Campaign Launch", tasks: ["Schedule deployment", "QA & go-live", "Confirm delivery"] },
        { key: "m6", col: "Measure", tile: "Results Tracked", sub: "KPIs & attribution", name: "Performance Measurement", tasks: ["Track KPIs", "Run attribution", "Report results"] },
        { key: "m7", col: "Optimise", tile: "Optimisations", sub: "A/B & reallocation", name: "Optimisation", tasks: ["Analyse performance", "Run A/B tests", "Reallocate spend"] },
      ],
      impact: [
        { head: "Brief", big: "↓ ~50%", dir: "green", d: "faster brief turnaround", s: "templated & AI-assisted" },
        { head: "Content", big: "↑ ~3×", dir: "blue", d: "more assets produced", s: "generative creative" },
        { head: "MLR", big: "↓ ~60%", dir: "green", d: "review cycle time", s: "pre-checked submissions" },
        { head: "Channel", big: "↑ +35%", dir: "blue", d: "channel efficiency", s: "omnichannel orchestration" },
        { head: "Launch", big: "↓ ~40%", dir: "green", d: "time to live", s: "automated deployment" },
        { head: "Measurement", big: "↑ ~2×", dir: "blue", d: "insight cadence", s: "real-time dashboards" },
        { head: "Optimisation", big: "↑ +28%", dir: "blue", d: "campaign ROI", s: "continuous A/B testing" },
      ],
      subjects: [
        { name: "PNH HCP Awareness Q3", area: "Eculizumab", category: "Email + Web", p: 5, act: "progress" },
        { name: "HFrEF Patient Education", area: "Cardio Brand", category: "Social + Display", p: 2, act: "approval" },
        { name: "Congress Activation — EHA", area: "Eculizumab", category: "Events + Email", p: 4, act: "awaiting" },
        { name: "Rep-Triggered eDetailing", area: "Cardio Brand", category: "Field + Email", p: 6, act: "progress" },
        { name: "Disease-State Hub Refresh", area: "Portfolio", category: "Web", p: 7, act: "done" },
        { name: "Formulary Win Announce", area: "Eculizumab", category: "Email", p: 1, act: "progress" },
        { name: "New Indication Teaser", area: "Cardio Brand", category: "Display", p: 0, act: "pending" },
      ],
    },
    map: {
      flow: "Market Access Flow",
      subjectNoun: "Payer", subjectPlural: "Payers",
      colSubject: "Account", colArea: "Region", colCategory: "Account Type",
      stages: [
        { key: "a1", col: "Dossier", tile: "Dossiers Built", sub: "Evidence compiled", name: "Value Dossier Build", tasks: ["Compile clinical evidence", "Add HEOR data", "Assemble dossier"] },
        { key: "a2", col: "Value", tile: "Value Stories", sub: "Narrative framed", name: "Value Story Development", tasks: ["Frame value narrative", "Build budget-impact model", "Tailor to payer"] },
        { key: "a3", col: "Pricing", tile: "Pricing Models", sub: "Scenario modelled", name: "Pricing & Contracting", tasks: ["Model price scenarios", "Define contract terms", "Approve strategy"] },
        { key: "a4", col: "Negotiate", tile: "Negotiations", sub: "Terms agreed", name: "Payer Negotiation", tasks: ["Present value story", "Negotiate terms", "Agree contract"] },
        { key: "a5", col: "Formulary", tile: "Formulary Wins", sub: "Listed", name: "Formulary Access", tasks: ["Submit for review", "Address P&T questions", "Confirm listing"] },
        { key: "a6", col: "Monitor", tile: "Contracts Tracked", sub: "Rebates & compliance", name: "Contract Monitoring", tasks: ["Track utilisation", "Reconcile rebates", "Monitor compliance"] },
      ],
      impact: [
        { head: "Dossier", big: "↓ ~55%", dir: "green", d: "dossier build time", s: "evidence auto-compiled" },
        { head: "Value Story", big: "↑ +30%", dir: "blue", d: "payer relevance", s: "tailored narratives" },
        { head: "Pricing", big: "↓ ~40%", dir: "green", d: "modelling time", s: "scenario automation" },
        { head: "Negotiation", big: "↑ ~2×", dir: "blue", d: "deals supported", s: "data-ready negotiators" },
        { head: "Formulary", big: "↑ +22%", dir: "blue", d: "access wins", s: "stronger submissions" },
        { head: "Monitoring", big: "↓ ~65%", dir: "green", d: "reconciliation effort", s: "automated tracking" },
        { head: "Compliance", big: "↑ +35%", dir: "blue", d: "contract adherence", s: "continuous monitoring" },
      ],
      subjects: [
        { name: "National Health Plan A", area: "National", category: "Commercial Payer", p: 4, act: "progress" },
        { name: "Regional Plan — Midwest", area: "Midwest", category: "Commercial Payer", p: 2, act: "approval" },
        { name: "State Medicaid — CA", area: "West", category: "Government", p: 3, act: "awaiting" },
        { name: "IDN — Coastal Health", area: "Northeast", category: "IDN", p: 5, act: "progress" },
        { name: "PBM — Major", area: "National", category: "PBM", p: 6, act: "done" },
        { name: "Regional Plan — South", area: "South", category: "Commercial Payer", p: 1, act: "progress" },
        { name: "State Medicaid — TX", area: "South", category: "Government", p: 0, act: "pending" },
      ],
    },
    med: {
      flow: "Medical Engagement Flow",
      subjectNoun: "Engagement", subjectPlural: "Engagements",
      colSubject: "Item", colArea: "Therapeutic Area", colCategory: "Type",
      stages: [
        { key: "e1", col: "Intake", tile: "Items Received", sub: "Logged & triaged", name: "Intake & Triage", tasks: ["Log request", "Classify & triage", "Assign owner"] },
        { key: "e2", col: "Research", tile: "Researched", sub: "Evidence gathered", name: "Evidence Research", tasks: ["Search literature", "Gather data", "Summarise findings"] },
        { key: "e3", col: "Draft", tile: "Drafts Prepared", sub: "Content drafted", name: "Drafting", tasks: ["Draft response / content", "Add references", "Internal QC"] },
        { key: "e4", col: "Review", tile: "MLR Reviewed", sub: "Compliance cleared", name: "Scientific / MLR Review", tasks: ["Submit for review", "Resolve comments", "Capture approval"] },
        { key: "e5", col: "Deliver", tile: "Delivered", sub: "Sent / published", name: "Delivery", tasks: ["Send / publish", "Confirm receipt", "Notify stakeholders"] },
        { key: "e6", col: "Record", tile: "Recorded", sub: "Repository synced", name: "Documentation", tasks: ["Archive record", "Tag metadata", "Sync repository"] },
      ],
      impact: [
        { head: "Intake", big: "↓ ~60%", dir: "green", d: "triage time", s: "auto-classified" },
        { head: "Research", big: "↑ ~3×", dir: "blue", d: "evidence coverage", s: "AI literature search" },
        { head: "Drafting", big: "↓ ~50%", dir: "green", d: "drafting time", s: "generated drafts" },
        { head: "Review", big: "↓ ~55%", dir: "green", d: "review cycle", s: "pre-checked submissions" },
        { head: "Delivery", big: "↑ ~2×", dir: "blue", d: "response throughput", s: "orchestrated delivery" },
        { head: "Documentation", big: "↓ ~70%", dir: "green", d: "documentation effort", s: "auto-archived" },
        { head: "Quality", big: "↑ +30%", dir: "blue", d: "scientific quality", s: "consistent & cited" },
      ],
      subjects: [
        { name: "MI Request — PNH dosing", area: "Hematology", category: "Medical Info", p: 4, act: "progress" },
        { name: "Publication — RWE PNH", area: "Hematology", category: "Publication", p: 2, act: "approval" },
        { name: "MI Request — HFrEF safety", area: "Cardiology", category: "Medical Info", p: 3, act: "awaiting" },
        { name: "Congress Abstract — EHA", area: "Hematology", category: "Publication", p: 5, act: "progress" },
        { name: "Med Comms — Slide Deck", area: "Cardiology", category: "Med Comms", p: 6, act: "done" },
        { name: "MI Request — interactions", area: "Hematology", category: "Medical Info", p: 1, act: "progress" },
        { name: "Manuscript — durability", area: "Hematology", category: "Publication", p: 0, act: "pending" },
      ],
    },
    coa: {
      flow: "Analytics & Operations Flow",
      subjectNoun: "Initiative", subjectPlural: "Initiatives",
      colSubject: "Initiative", colArea: "Function", colCategory: "Type",
      stages: [
        { key: "c1", col: "Intake", tile: "Requests Logged", sub: "Scoped & prioritised", name: "Request Intake", tasks: ["Log request", "Scope & estimate", "Prioritise"] },
        { key: "c2", col: "Data", tile: "Data Prepared", sub: "Sourced & cleaned", name: "Data Preparation", tasks: ["Source data", "Clean & validate", "Build dataset"] },
        { key: "c3", col: "Model", tile: "Models Built", sub: "Analysis run", name: "Modelling & Analysis", tasks: ["Build model", "Run analysis", "Validate results"] },
        { key: "c4", col: "Insight", tile: "Insights Ready", sub: "Synthesised", name: "Insight Synthesis", tasks: ["Synthesise findings", "Frame recommendations", "Review with SME"] },
        { key: "c5", col: "Dashboard", tile: "Dashboards Live", sub: "Published", name: "Dashboard Delivery", tasks: ["Build visualisation", "QA & publish", "Grant access"] },
        { key: "c6", col: "Action", tile: "Actions Routed", sub: "Recommendations", name: "Action & Enablement", tasks: ["Route recommendations", "Enable owners", "Track adoption"] },
      ],
      impact: [
        { head: "Intake", big: "↓ ~50%", dir: "green", d: "request turnaround", s: "auto-scoped" },
        { head: "Data Prep", big: "↓ ~65%", dir: "green", d: "data wrangling time", s: "automated pipelines" },
        { head: "Modelling", big: "↑ ~2×", dir: "blue", d: "analyses delivered", s: "reusable models" },
        { head: "Insight", big: "↓ ~40%", dir: "green", d: "insight-to-action time", s: "auto-synthesised" },
        { head: "Dashboards", big: "↑ ~3×", dir: "blue", d: "self-serve coverage", s: "live dashboards" },
        { head: "Action", big: "↑ +30%", dir: "blue", d: "recommendation adoption", s: "routed & tracked" },
        { head: "Excellence", big: "↑ +25%", dir: "blue", d: "commercial performance", s: "capability enablement" },
      ],
      subjects: [
        { name: "Q3 Sales Performance Review", area: "Sales Ops", category: "Report", p: 5, act: "progress" },
        { name: "Territory Alignment Model", area: "Commercial Excellence", category: "Model", p: 2, act: "approval" },
        { name: "Brand Performance Dashboard", area: "Analytics/BI", category: "Dashboard", p: 3, act: "awaiting" },
        { name: "Incentive Compensation Plan", area: "Sales Ops", category: "Model", p: 4, act: "progress" },
        { name: "Field Force Sizing Study", area: "Commercial Excellence", category: "Analysis", p: 6, act: "done" },
        { name: "Next-Best-Action Engine", area: "Analytics/BI", category: "Model", p: 1, act: "progress" },
        { name: "Launch Readiness Tracker", area: "Commercial Excellence", category: "Dashboard", p: 0, act: "pending" },
      ],
    },
  };

  /* ---- Roles -> Division map (codes must match the role-selection screen) */
  var ROLE_INDEX = {
    SR: { div: "sff", name: "Sales Representatives" },
    KAM: { div: "sff", name: "Key Account Managers" },
    RDM: { div: "sff", name: "Regional/District Managers" },
    NSD: { div: "sff", name: "National Sales Director" },
    MM: { div: "mkt", name: "Brand/Marketing Managers" },
    PM: { div: "mkt", name: "Product Managers" },
    MKO: { div: "mkt", name: "Marketing Operations" },
    MLR: { div: "mkt", name: "MLR" },
    DM: { div: "mkt", name: "Digital Marketing" },
    LEX: { div: "mkt", name: "Launch Excellence" },
    PR: { div: "map", name: "Payer Relations" },
    PC: { div: "map", name: "Pricing & Contracting" },
    HEOR: { div: "map", name: "HEOR" },
    TD: { div: "map", name: "Trade & Distribution" },
    MSL: { div: "med", name: "Medical Science Liaisons" },
    MI: { div: "med", name: "Medical Information" },
    PMC: { div: "med", name: "Publications & Med Comms" },
    CBI: { div: "coa", name: "Commercial Analytics/BI" },
    SO: { div: "coa", name: "Sales Operations" },
    CE: { div: "coa", name: "Commercial Excellence" },
  };

  /* ---- Build a full config object for a role code --------------------- */
  function buildRoleConfig(code) {
    var idx = ROLE_INDEX[code];
    if (!idx) return null;
    var t = THEMES[idx.div];
    var stages = t.stages.map(function (s) { return { key: s.key, col: s.col, tile: s.tile, sub: s.sub, name: s.name, tasks: (s.tasks || []).slice() }; });
    var owner = idx.name;
    var subjects = t.subjects.map(function (sp) {
      var cells = sp.cells ? sp.cells.slice() : mk(stages.length, sp.p, sp.act);
      return { name: sp.name, area: sp.area, category: sp.category, journey: buildJourney(stages, cells, owner) };
    });
    var cfg = {
      mode: "generic", code: code, div: idx.div,
      title: idx.name + " Activity Dashboard",
      subtitle: idx.name + " (" + code + ") · " + t.flow + " · live status",
      subjectNoun: t.subjectNoun, subjectPlural: t.subjectPlural,
      colSubject: t.colSubject, colArea: t.colArea, colCategory: t.colCategory,
      assistant: "Hi! I'm synced to the " + t.subjectNoun.toLowerCase() + " you select on the left. Ask about its status or stage — or type “approve” to action a pending request.",
      footerLabel: "Agentic " + idx.name + " Employee",
      stages: stages, impact: t.impact.slice(), subjects: subjects,
    };
    cfg.agents = buildAgents(cfg);
    return cfg;
  }

  /* ===================================================================
     Excel mapping  —  one worksheet (array-of-arrays) per Division+Role.
     Sections are tagged in column A with #META / #STAGES / #IMPACT /
     #SUBJECTS.  For #STAGES/#IMPACT/#SUBJECTS the tag row also carries the
     column headers; for #META each row is a key / value pair.
     =================================================================== */
  function configToAOA(cfg) {
    var rows = [];
    rows.push(["#META", "value"]);
    rows.push(["title", cfg.title]);
    rows.push(["subtitle", cfg.subtitle]);
    rows.push(["subjectNoun", cfg.subjectNoun]);
    rows.push(["subjectPlural", cfg.subjectPlural]);
    rows.push(["colSubject", cfg.colSubject]);
    rows.push(["colArea", cfg.colArea]);
    rows.push(["colCategory", cfg.colCategory]);
    rows.push(["assistant", cfg.assistant]);
    rows.push(["footerLabel", cfg.footerLabel]);
    rows.push(["agentRootL1", cfg.agents && cfg.agents.root ? cfg.agents.root.l1 : ""]);
    rows.push(["agentRootL2", cfg.agents && cfg.agents.root ? cfg.agents.root.l2 : ""]);
    rows.push([]);

    rows.push(["#STAGES", "col", "tile", "sub", "name", "tasks"]);
    cfg.stages.forEach(function (s) { rows.push([s.key, s.col, s.tile, s.sub, s.name, (s.tasks || []).join("; ")]); });
    rows.push([]);

    rows.push(["#IMPACT", "value", "dir", "desc", "sub"]);
    cfg.impact.forEach(function (m) { rows.push([m.head, m.big, m.dir, m.d, m.s]); });
    rows.push([]);

    // Agent Workspace (tab 2): orchestrator agents + their tools
    rows.push(["#AGENTS", "label", "cluster", "tools"]);
    ((cfg.agents && cfg.agents.list) || []).forEach(function (a) { rows.push([a.id, a.label, a.cluster, (a.tools || []).join("; ")]); });
    rows.push([]);

    // Subjects: name, area, category, then one column per stage holding "state | detail"
    var head = ["#SUBJECTS", "area", "category"];
    cfg.stages.forEach(function (s) { head.push(s.key); });
    rows.push(head);
    cfg.subjects.forEach(function (su) {
      var r = [su.name, su.area || su.therapeutic_area || "", su.category || su.disease || ""];
      var j = su.journey || [];
      cfg.stages.forEach(function (s, i) {
        var st = j[i] || {};
        var cell = (st.state || "pending");
        if (st.detail && st.detail !== defaultDetail(s, st.state, i > 0 ? cfg.stages[i - 1].name : "")) cell += " | " + st.detail;
        r.push(cell);
      });
      rows.push(r);
    });
    return rows;
  }

  function aoaToConfig(aoa, sheetName) {
    if (!aoa || !aoa.length) return null;
    var meta = {}, stages = [], impact = [], subjRows = [], agentRows = [], section = "", head = [];
    for (var r = 0; r < aoa.length; r++) {
      var row = aoa[r] || [];
      var a = (row[0] == null ? "" : String(row[0])).trim();
      if (a.charAt(0) === "#") { section = a.toUpperCase(); head = row; continue; }
      if (a === "" && row.length === 0) continue;
      if (section === "#META") { if (a) meta[a] = row[1] != null ? String(row[1]) : ""; }
      else if (section === "#STAGES") { if (a) stages.push({ key: a, col: s2(row[1]), tile: s2(row[2]), sub: s2(row[3]), name: s2(row[4]) || a, tasks: s2(row[5]) ? String(row[5]).split(";").map(function (x) { return x.trim(); }).filter(Boolean) : [] }); }
      else if (section === "#IMPACT") { if (a) impact.push({ head: a, big: s2(row[1]), dir: s2(row[2]) || "green", d: s2(row[3]), s: s2(row[4]) }); }
      else if (section === "#SUBJECTS") { if (a) subjRows.push(row); }
      else if (section === "#AGENTS") { if (a) agentRows.push(row); }
    }
    if (!stages.length) return null;
    var code = (sheetName && sheetName.indexOf("_") >= 0) ? sheetName.split("_").pop() : sheetName;
    var idx = ROLE_INDEX[code] || {};
    var owner = meta.footerLabel || idx.name || code;
    var subjects = subjRows.map(function (row) {
      var cells = [];
      var details = [];
      for (var i = 0; i < stages.length; i++) {
        var raw = row[3 + i] != null ? String(row[3 + i]) : "pending";
        var parts = raw.split("|");
        cells.push((parts[0] || "pending").trim());
        details.push(parts[1] ? parts[1].trim() : null);
      }
      var journey = buildJourney(stages, cells, owner);
      journey.forEach(function (st, i) { if (details[i]) st.detail = details[i]; });
      return { name: s2(row[0]), area: s2(row[1]), category: s2(row[2]), journey: journey };
    });
    var cfg = {
      mode: "generic", code: code, div: idx.div || (sheetName && sheetName.split("_")[0]),
      title: meta.title || (code + " Dashboard"),
      subtitle: meta.subtitle || "",
      subjectNoun: meta.subjectNoun || "Item", subjectPlural: meta.subjectPlural || "Items",
      colSubject: meta.colSubject || "Item", colArea: meta.colArea || "Area", colCategory: meta.colCategory || "Type",
      assistant: meta.assistant || "Hi! Select an item on the left.",
      footerLabel: meta.footerLabel || ("Agentic " + (idx.name || code) + " Employee"),
      stages: stages, impact: impact, subjects: subjects,
    };
    if (agentRows.length) {
      cfg.agents = {
        root: { l1: meta.agentRootL1 || (code || "agent").toLowerCase().slice(0, 9), l2: meta.agentRootL2 || "workspace" },
        list: agentRows.map(function (row) { return { id: s2(row[0]), label: s2(row[1]) || s2(row[0]), cluster: s2(row[2]), group: 0, tools: s2(row[3]) ? String(row[3]).split(";").map(function (x) { return x.trim(); }).filter(Boolean) : [] }; }),
      };
      cfg.agents.list.forEach(function (a, i) { a.group = i % 6; });
    } else cfg.agents = buildAgents(cfg);
    return cfg;
  }
  function s2(v) { return v == null ? "" : String(v); }

  return {
    STLBL: STLBL, THEMES: THEMES, ROLE_INDEX: ROLE_INDEX,
    buildRoleConfig: buildRoleConfig, buildJourney: buildJourney, mk: mk, buildAgents: buildAgents,
    configToAOA: configToAOA, aoaToConfig: aoaToConfig,
  };
});
