#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GRC Shield layer for the Risk Register Consolidator
===================================================
Adds ShieldGRC-style capabilities ON TOP of consolidate_risks.py, without
touching it. Works on the same staging_output.xlsx and sources/ folder.

Built by Tosin Aina - Northbridge.  License: MIT

Commands:
    python grc_shield.py ai-pack     # write a curated AI-risk register into sources/
    python grc_shield.py enrich      # tag staged risks (category, ATT&CK, NIST, treatment)
    python grc_shield.py dashboard   # build a ShieldGRC-style dashboard.html
    python grc_shield.py ai-enrich   # optional: polish tags with Claude (needs ANTHROPIC_API_KEY)

Typical flow:
    python consolidate_risks.py run          # -> staging_output.xlsx
    python grc_shield.py enrich              # -> staging_enriched.xlsx
    python grc_shield.py dashboard           # -> dashboard.html
"""

import os
import sys
import html
import json

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required:  pip install openpyxl")

SOURCES_DIR = "sources"
STAGING_FILE = "staging_output.xlsx"
ENRICHED_FILE = "staging_enriched.xlsx"
DASHBOARD_FILE = "dashboard.html"
AI_PACK_FILE = os.path.join(SOURCES_DIR, "AI_Risk_Pack.xlsx")
TEAL = "0B8F70"

# ---------------------------------------------------------------------------
# Keyword -> (Category, ATT&CK/ATLAS id, NIST CSF ref, suggested treatment)
# Offline, transparent rules. First match wins; order matters (specific first).
# ---------------------------------------------------------------------------
RULES = [
    (("prompt injection", "llm", "language model", "chatbot", "genai", "gen ai"),
     ("AI / LLM", "ATLAS AML.T0051", "NIST AI RMF: Measure", "Guardrails + input/output validation")),
    (("model poison", "training data", "data poison"),
     ("AI / ML", "ATLAS AML.T0020", "NIST AI RMF: Manage", "Data provenance + integrity checks")),
    (("ai ", "artificial intelligence", "model drift", "hallucinat", "agentic", "autonomous agent"),
     ("AI Risk", "MITRE ATLAS", "NIST AI RMF", "Human-in-the-loop + monitoring")),
    (("phish", "spear", "social engineer", "business email"),
     ("Social Engineering", "ATT&CK T1566", "NIST PR.AT / DE.CM", "Awareness training + email filtering")),
    (("ransomware", "encrypt", "extortion"),
     ("Malware / Ransomware", "ATT&CK T1486", "NIST PR.IP / RC.RP", "Backups + EDR + segmentation")),
    (("malware", "virus", "trojan", "worm"),
     ("Malware", "ATT&CK T1204", "NIST DE.CM", "EDR + patching + least privilege")),
    (("credential", "password", "brute force", "mfa", "authentication"),
     ("Access Control", "ATT&CK T1110 / T1078", "NIST PR.AC", "MFA + strong password policy")),
    (("privilege", "admin rights", "escalat"),
     ("Privileged Access", "ATT&CK T1068 / T1078", "NIST PR.AC-4", "Least privilege + PAM")),
    (("exfil", "data leak", "data loss", "leakage"),
     ("Data Protection", "ATT&CK T1041 / T1530", "NIST PR.DS", "DLP + encryption + access controls")),
    (("cloud", "s3", "bucket", "misconfig", "iam"),
     ("Cloud Security", "ATT&CK T1530", "NIST PR.AC / PR.DS", "CSPM + least privilege + posture checks")),
    (("unpatched", "patch", "vulnerab", "cve", "outdated"),
     ("Vulnerability Mgmt", "ATT&CK T1190", "NIST ID.RA / PR.IP", "Patch SLA + vulnerability scanning")),
    (("ddos", "denial of service", "availability", "outage"),
     ("Availability", "ATT&CK T1498", "NIST PR.PT / RC", "Redundancy + DDoS protection")),
    (("backup", "recovery", "disaster", "continuity", "bcp", "dr "),
     ("Resilience", "-", "NIST RC.RP", "Tested backups + DR plan")),
    (("third party", "vendor", "supplier", "subprocessor", "outsourc"),
     ("Third-Party Risk", "-", "NIST ID.SC", "Vendor due diligence + contractual controls")),
    (("insider", "disgruntled", "malicious employee"),
     ("Insider Threat", "ATT&CK T1078", "NIST PR.AC / DE.CM", "Segregation of duties + monitoring")),
    (("physical", "premises", "building", "theft of"),
     ("Physical Security", "-", "NIST PR.AC-2", "Physical access controls")),
    (("logging", "monitor", "siem", "detection", "audit trail"),
     ("Detection & Monitoring", "-", "NIST DE.CM", "Centralised logging + alerting")),
    (("encryption", "at rest", "in transit", "tls"),
     ("Data Protection", "-", "NIST PR.DS-1/2", "Encryption standards + key management")),
    (("gdpr", "privacy", "personal data", "pii", "data protection act"),
     ("Privacy / Compliance", "-", "ISO 27701 / GDPR", "DPIA + data minimisation")),
    (("compliance", "regulat", "iso 27001", "soc 2", "pci", "hipaa"),
     ("Compliance", "-", "ISO 27001 / SOC 2", "Control mapping + evidence")),
]
DEFAULT_TAG = ("General", "-", "NIST ID.RA", "Assess + assign owner")


def classify(text):
    t = (text or "").lower()
    for keys, tag in RULES:
        if any(k in t for k in keys):
            return tag
    return DEFAULT_TAG


def severity(score):
    try:
        s = float(score)
    except (TypeError, ValueError):
        return "Unscored"
    if s >= 15:
        return "Critical"
    if s >= 10:
        return "High"
    if s >= 5:
        return "Medium"
    return "Low"


# ---------------------------------------------------------------------------
# ai-pack : curated AI / agentic risk register, Drata-shaped, into sources/
# ---------------------------------------------------------------------------
AI_RISKS = [
    ("AI-01", "Prompt injection into production LLM", "A crafted user input overrides the system prompt of a customer-facing LLM, leaking data or triggering unintended actions.", 4, 5),
    ("AI-02", "Sensitive data leakage via LLM output", "The model reveals PII, secrets, or internal data in its responses (OWASP LLM06).", 3, 5),
    ("AI-03", "Training / data poisoning", "Adversary manipulates training or fine-tuning data to bias or backdoor the model.", 2, 5),
    ("AI-04", "Model theft / extraction", "Repeated queries reconstruct or steal a proprietary model (OWASP LLM10 / ATLAS).", 2, 4),
    ("AI-05", "Reliance on hallucinated output", "Staff act on confident but wrong model output without verification.", 4, 4),
    ("AI-06", "Insecure LLM tool / plugin use", "An agent's tool call performs an unsafe action with excessive permissions (OWASP LLM07).", 3, 5),
    ("AI-07", "Excessive agency of autonomous agent", "An agent takes real-world actions without adequate human oversight (OWASP LLM08 / MAESTRO).", 3, 5),
    ("AI-08", "AI supply-chain compromise", "A compromised model, library, or dataset is pulled from an untrusted source (OWASP LLM05).", 3, 4),
    ("AI-09", "Insufficient human oversight (high-risk AI)", "A high-risk AI system operates without the human review the EU AI Act (Art. 14) requires.", 3, 5),
    ("AI-10", "Bias / discrimination in model outputs", "Model produces unfair outcomes across protected groups (NIST AI RMF fairness; EU AI Act).", 3, 4),
    ("AI-11", "No model monitoring / drift detection", "Model accuracy silently degrades in production with no monitoring (NIST AI RMF: Manage).", 4, 3),
    ("AI-12", "Unbounded consumption / LLM DoS", "Uncontrolled prompts drive cost or availability failures (OWASP LLM10).", 3, 3),
    ("AI-13", "Insufficient logging of AI decisions", "No record of AI inputs/outputs for audit or incident review (EU AI Act Art. 12).", 3, 4),
    ("AI-14", "Third-party AI processor without DPA", "PII sent to an external AI service with no data-processing agreement or assessment.", 3, 5),
]


def cmd_ai_pack():
    os.makedirs(SOURCES_DIR, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "AI Risks"
    heads = ["ID", "Risk Title", "Description", "Likelihood", "Impact", "Owner"]
    widths = [8, 42, 70, 12, 10, 18]
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=TEAL)
        ws.column_dimensions[get_column_letter(i)].width = w
    for rid, title, desc, like, imp in AI_RISKS:
        ws.append([rid, title, desc, like, imp, ""])
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(AI_PACK_FILE)
    print(f"Wrote {len(AI_RISKS)} AI/agentic risks -> {AI_PACK_FILE}")
    print("Next: fold it into your consolidation with")
    print("    python consolidate_risks.py automap   (then confirm mapping.xlsx)")
    print("    python consolidate_risks.py run")
    print("Columns are already scale 5x5: title=B, description=C, likelihood=D, impact=E, owner=F, ref=A")


# ---------------------------------------------------------------------------
# enrich : add Category / ATT&CK / NIST / Treatment to the staged register
# ---------------------------------------------------------------------------
def _read_staging(path):
    wb = load_workbook(path, data_only=True)
    ws = wb["STAGING"] if "STAGING" in wb.sheetnames else wb.active
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[idx.get("Title", 1)] is None:
            continue
        rows.append(r)
    return headers, idx, rows


def cmd_enrich():
    if not os.path.exists(STAGING_FILE):
        sys.exit(f"No {STAGING_FILE}. Run:  python consolidate_risks.py run")
    headers, idx, rows = _read_staging(STAGING_FILE)

    wb = Workbook()
    ws = wb.active
    ws.title = "STAGING (enriched)"
    new_cols = ["Category (auto)", "ATT&CK / ATLAS", "NIST / ISO", "Suggested Treatment"]
    out_headers = list(headers) + new_cols
    for i, h in enumerate(out_headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=TEAL)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"

    ti, di = idx.get("Title", 1), idx.get("Description", 2)
    cat_i = idx.get("Category")
    counts = {}
    for r, row in enumerate(rows, 2):
        text = f"{row[ti] or ''} {row[di] or ''}"
        cat, attack, nist, treat = classify(text)
        counts[cat] = counts.get(cat, 0) + 1
        vals = list(row)
        if cat_i is not None and not (vals[cat_i] and str(vals[cat_i]).strip()):
            vals[cat_i] = cat  # backfill the empty Category column too
        vals = vals + [cat, attack, nist, treat]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    # widths: copy sensible defaults
    for i in range(1, len(out_headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 50

    wb.save(ENRICHED_FILE)
    top = sorted(counts.items(), key=lambda x: -x[1])[:6]
    print(f"Enriched {len(rows)} risks -> {ENRICHED_FILE}")
    print("Top categories: " + ", ".join(f"{k} ({v})" for k, v in top))
    print("Every tag is a STARTING POINT - review before importing to Drata.")


# ---------------------------------------------------------------------------
# dashboard : ShieldGRC-style HTML from the staged (or enriched) register
# ---------------------------------------------------------------------------
def cmd_dashboard():
    src = ENRICHED_FILE if os.path.exists(ENRICHED_FILE) else STAGING_FILE
    if not os.path.exists(src):
        sys.exit(f"No staging file. Run:  python consolidate_risks.py run")
    headers, idx, rows = _read_staging(src)
    ti = idx.get("Title", 1)
    si = idx.get("Inherent Score")
    ri = idx.get("Source Register")
    sti = idx.get("Status")
    ci = idx.get("Category (auto)", idx.get("Category"))

    risks = []
    for row in rows:
        score = row[si] if si is not None else ""
        risks.append({
            "title": str(row[ti] or ""),
            "score": score,
            "sev": severity(score),
            "reg": str(row[ri] or "-") if ri is not None else "-",
            "status": str(row[sti] or "") if sti is not None else "",
            "cat": str(row[ci] or "") if ci is not None else "",
        })

    total = len(risks)
    sev_order = ["Critical", "High", "Medium", "Low", "Unscored"]
    sev_counts = {s: sum(1 for x in risks if x["sev"] == s) for s in sev_order}
    review = sum(1 for x in risks if "review" in x["status"].lower())
    by_reg = {}
    for x in risks:
        by_reg[x["reg"]] = by_reg.get(x["reg"], 0) + 1
    by_cat = {}
    for x in risks:
        if x["cat"]:
            by_cat[x["cat"]] = by_cat.get(x["cat"], 0) + 1
    ai_count = sum(v for k, v in by_cat.items() if "ai" in k.lower())
    top_risks = sorted(
        [x for x in risks if isinstance(x["score"], (int, float))],
        key=lambda x: -x["score"])[:10]

    html_out = _render_dashboard(total, sev_counts, sev_order, review, by_reg, by_cat, ai_count, top_risks, src)
    with open(DASHBOARD_FILE, "w", encoding="utf-8") as f:
        f.write(html_out)
    print(f"Dashboard -> {DASHBOARD_FILE}  ({total} risks from {src})")
    print("Open it in a browser. Regenerate any time after enrich/run.")


SEV_COLOR = {"Critical": "#c0392b", "High": "#e07b39", "Medium": "#c9a227", "Low": "#2f8f5b", "Unscored": "#8a94a6"}


def _bar_rows(data, color="#0d8b86", maxw=None):
    if not data:
        return '<div class="muted">No data.</div>'
    mx = maxw or max(data.values()) or 1
    out = []
    for k, v in sorted(data.items(), key=lambda x: -x[1]):
        pct = int(v / mx * 100)
        out.append(
            f'<div class="brow"><div class="blabel">{html.escape(str(k))}</div>'
            f'<div class="btrack"><span style="width:{pct}%;background:{color}"></span></div>'
            f'<div class="bval">{v}</div></div>')
    return "".join(out)


def _render_dashboard(total, sev_counts, sev_order, review, by_reg, by_cat, ai_count, top_risks, src):
    # severity donut (conic-gradient)
    stops, acc = [], 0
    for s in sev_order:
        n = sev_counts.get(s, 0)
        if not n:
            continue
        frac = n / total * 100 if total else 0
        stops.append(f"{SEV_COLOR[s]} {acc:.1f}% {acc+frac:.1f}%")
        acc += frac
    donut = "conic-gradient(" + ", ".join(stops) + ")" if stops else "#e2e8f0"
    sev_legend = "".join(
        f'<div class="legrow"><span class="dot" style="background:{SEV_COLOR[s]}"></span>{s}<b>{sev_counts.get(s,0)}</b></div>'
        for s in sev_order if sev_counts.get(s, 0))
    critical = sev_counts.get("Critical", 0) + sev_counts.get("High", 0)

    rows_html = "".join(
        f'<tr><td>{html.escape(x["title"][:70])}</td>'
        f'<td><span class="pill" style="background:{SEV_COLOR[x["sev"]]}22;color:{SEV_COLOR[x["sev"]]}">{x["sev"]}</span></td>'
        f'<td class="mono">{x["score"]}</td><td class="muted">{html.escape(x["reg"])}</td>'
        f'<td class="muted">{html.escape(x["cat"])}</td></tr>'
        for x in top_risks) or '<tr><td colspan="5" class="muted">No scored risks yet.</td></tr>'

    cat_chips = "".join(
        f'<span class="chip">{html.escape(k)} <b>{v}</b></span>'
        for k, v in sorted(by_cat.items(), key=lambda x: -x[1])) or '<span class="muted">Run enrich to populate categories.</span>'

    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ShieldGRC · Risk Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{{--bg:#eef2f6;--panel:#fff;--ink:#16202e;--soft:#54637a;--faint:#8291a5;--line:#e2e8f0;--accent:#0d8b86;--ai:#7c3aed}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font-family:Inter,system-ui,sans-serif}}
.app{{display:grid;grid-template-columns:230px 1fr;min-height:100vh}}
.side{{background:#0f1b2d;color:#cdd8e8;padding:22px 16px}}
.side .brand{{font-family:Sora;font-weight:700;font-size:17px;color:#fff;display:flex;gap:9px;align-items:center;margin-bottom:24px}}
.side .brand .mk{{width:26px;height:26px;border-radius:7px;background:linear-gradient(135deg,var(--accent),var(--ai));display:grid;place-items:center;font-size:14px}}
.side .grp{{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#6c7c94;margin:16px 0 6px}}
.side a{{display:block;color:#cdd8e8;text-decoration:none;padding:8px 10px;border-radius:8px;font-size:14px;font-weight:500}}
.side a.on,.side a:hover{{background:rgba(255,255,255,.08);color:#fff}}
.main{{padding:26px 30px}}
.h{{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px}}
.h h1{{font-family:Sora;font-size:22px;margin:0}}.h .src{{font-size:12px;color:var(--faint)}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}}
.kpi{{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:16px 18px;box-shadow:0 10px 26px -20px rgba(20,36,58,.4)}}
.kpi .n{{font-family:Sora;font-weight:700;font-size:28px}}.kpi .l{{font-size:12.5px;color:var(--faint);margin-top:3px}}
.kpi.red .n{{color:#c0392b}}.kpi.ai .n{{color:var(--ai)}}.kpi.amber .n{{color:#c9a227}}
.grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}}
.card{{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px 20px;box-shadow:0 10px 26px -20px rgba(20,36,58,.4)}}
.card h3{{font-family:Sora;font-size:14px;margin:0 0 14px}}
.donutwrap{{display:flex;gap:20px;align-items:center}}
.donut{{width:130px;height:130px;border-radius:50%;background:{donut};flex:none;
  -webkit-mask:radial-gradient(circle 40px at center,transparent 98%,#000);mask:radial-gradient(circle 40px at center,transparent 98%,#000)}}
.legrow{{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--soft);margin-bottom:6px}}
.legrow b{{margin-left:auto;color:var(--ink)}}.dot{{width:10px;height:10px;border-radius:3px}}
.brow{{display:grid;grid-template-columns:130px 1fr 34px;gap:10px;align-items:center;margin-bottom:9px;font-size:13px}}
.blabel{{color:var(--soft);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.btrack{{height:9px;background:#eef2f6;border-radius:6px;overflow:hidden}}.btrack span{{display:block;height:100%}}
.bval{{text-align:right;font-weight:600;font-size:12px}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{text-align:left;color:var(--faint);font-size:11px;text-transform:uppercase;letter-spacing:.05em;padding:8px 10px;border-bottom:1px solid var(--line)}}
td{{padding:9px 10px;border-bottom:1px solid var(--line)}}
.pill{{padding:2px 9px;border-radius:999px;font-size:11px;font-weight:600}}
.mono{{font-family:ui-monospace,monospace}}.muted{{color:var(--faint)}}
.chip{{display:inline-block;background:#f1f5f9;border:1px solid var(--line);border-radius:999px;padding:4px 11px;font-size:12px;color:var(--soft);margin:0 6px 6px 0}}
.chip b{{color:var(--accent)}}
.note{{font-size:11.5px;color:var(--faint);margin-top:14px}}
@media(max-width:820px){{.app{{grid-template-columns:1fr}}.side{{display:none}}.kpis{{grid-template-columns:1fr 1fr}}.grid{{grid-template-columns:1fr}}}}
</style></head><body><div class="app">
<aside class="side">
  <div class="brand"><span class="mk">◆</span> ShieldGRC</div>
  <div class="grp">Overview</div><a class="on">Dashboard</a>
  <div class="grp">Risk</div><a>Risk Register</a><a>Risk Assessment</a>
  <div class="grp">Compliance</div><a>Control Testing</a><a>Frameworks</a><a>AI Management</a>
  <div class="grp">Governance</div><a>Reports</a><a>Audit Trail</a>
</aside>
<main class="main">
  <div class="h"><h1>Risk dashboard</h1><div class="src">Source: {html.escape(os.path.basename(src))}</div></div>
  <div class="kpis">
    <div class="kpi"><div class="n">{total}</div><div class="l">Total risks staged</div></div>
    <div class="kpi red"><div class="n">{critical}</div><div class="l">Critical + High</div></div>
    <div class="kpi amber"><div class="n">{review}</div><div class="l">Need owner review</div></div>
    <div class="kpi ai"><div class="n">{ai_count}</div><div class="l">AI / agentic risks</div></div>
  </div>
  <div class="grid">
    <div class="card"><h3>Risk by severity</h3>
      <div class="donutwrap"><div class="donut"></div><div>{sev_legend or '<span class="muted">No scores yet.</span>'}</div></div></div>
    <div class="card"><h3>Risk by source register</h3>{_bar_rows(by_reg)}</div>
  </div>
  <div class="card" style="margin-bottom:16px"><h3>Top risks by inherent score</h3>
    <table><thead><tr><th>Risk</th><th>Severity</th><th>Score</th><th>Register</th><th>Category</th></tr></thead>
    <tbody>{rows_html}</tbody></table></div>
  <div class="card"><h3>Framework coverage (from enrich)</h3>{cat_chips}
    <div class="note">Auto-tags are starting points mapped to MITRE ATT&CK/ATLAS and NIST/ISO. Review with owners before importing to Drata.</div></div>
</main></div></body></html>"""


# ---------------------------------------------------------------------------
# ai-enrich : optional Claude polish (uses ANTHROPIC_API_KEY via stdlib urllib)
# ---------------------------------------------------------------------------
def cmd_ai_enrich():
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        sys.exit("Set ANTHROPIC_API_KEY to use ai-enrich (falls back to the offline 'enrich' otherwise).")
    if not os.path.exists(STAGING_FILE):
        sys.exit(f"No {STAGING_FILE}. Run:  python consolidate_risks.py run")
    import urllib.request
    headers, idx, rows = _read_staging(STAGING_FILE)
    ti, di = idx.get("Title", 1), idx.get("Description", 2)
    items = [{"i": n, "t": str(r[ti] or ""), "d": str(r[di] or "")} for n, r in enumerate(rows)]
    prompt = ("You are a GRC analyst. For each risk, return a JSON array of "
              '{"i":<index>,"category":"...","attack":"MITRE ATT&CK or ATLAS id or -","nist":"NIST CSF or ISO ref","treatment":"one short control"}. '
              "Base it ONLY on the text. Return ONLY the JSON array.\n\n" + json.dumps(items))
    body = json.dumps({
        "model": "claude-haiku-4-5-20251001", "max_tokens": 4000,
        "messages": [{"role": "user", "content": prompt}],
    }).encode()
    req = urllib.request.Request("https://api.anthropic.com/v1/messages", data=body, headers={
        "content-type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read())
        text = data["content"][0]["text"]
        tags = json.loads(text[text.find("["):text.rfind("]") + 1])
    except Exception as e:
        sys.exit(f"Claude call failed ({e}). Use the offline 'enrich' command instead.")
    by_i = {t["i"]: t for t in tags if isinstance(t, dict)}

    wb = Workbook(); ws = wb.active; ws.title = "STAGING (AI-enriched)"
    out_headers = list(headers) + ["Category (AI)", "ATT&CK / ATLAS (AI)", "NIST / ISO (AI)", "Treatment (AI)"]
    for i, h in enumerate(out_headers, 1):
        c = ws.cell(row=1, column=i, value=h); c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=TEAL)
    for r, row in enumerate(rows, 2):
        t = by_i.get(r - 2, {})
        vals = list(row) + [t.get("category", ""), t.get("attack", ""), t.get("nist", ""), t.get("treatment", "")]
        for c, v in enumerate(vals, 1):
            ws.cell(row=r, column=c, value=v)
    out = "staging_ai_enriched.xlsx"; wb.save(out)
    print(f"Claude-enriched {len(rows)} risks -> {out}")


if __name__ == "__main__":
    cmd = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    if cmd == "ai-pack":
        cmd_ai_pack()
    elif cmd == "enrich":
        cmd_enrich()
    elif cmd == "dashboard":
        cmd_dashboard()
    elif cmd == "ai-enrich":
        cmd_ai_enrich()
    else:
        print(__doc__)
