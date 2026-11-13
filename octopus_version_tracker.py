#!/usr/bin/env python3
"""
Harmony Version Tracker - Octopus Deploy sync
=============================================

Logs in to Octopus Deploy, reads the release that is currently deployed for every
project (component) in every environment, and writes / refreshes an Excel
tracker with one hyperlinked version per cell.

SETUP (once, in PowerShell)
    pip install requests openpyxl

AUTH (credentials are NEVER stored in this file or on disk)
    Option A - API key (recommended, required for unattended / scheduled runs)
        Octopus > your profile > My API Keys > New API key
        $env:OCTOPUS_API_KEY = "API-XXXXXXXX"
    Option B - username + password
        $env:OCTOPUS_USERNAME = "rajesh.m"
        $env:OCTOPUS_PASSWORD = "<your password>"       # or just run and you'll be prompted
    (Password login does not work if your Octopus uses SSO only - use an API key.)

RUN
    python octopus_version_tracker.py                 # one run: build / refresh the Excel
    python octopus_version_tracker.py --watch 15      # keep running, re-check every 15 min
    python octopus_version_tracker.py --list          # just list projects/environments found

FIRST RUN
    Creates tracker_config.json next to this script. It lists every Octopus project with
    a display name and an "include" flag. Edit it (rename rows, set include=false for
    projects you don't want), then run again. Environment names that could not be matched
    to a tracker column are printed - fix them in "env_overrides" in the same file.

AUTO REFRESH
    Run on a schedule so the Excel refreshes when new builds are deployed, e.g. every
    15 minutes (see the README text printed by --help-schedule), or use --watch.
"""
import argparse
import difflib
import getpass
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    import requests
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except ImportError:
    sys.exit("Missing packages. Run:  pip install requests openpyxl")

HERE = Path(__file__).resolve().parent
CONFIG_FILE = HERE / "tracker_config.json"
STATE_FILE = HERE / "octopus_state.json"

DEFAULT_CONFIG = {
    "base_url": "https://deploy.jitterbit.us",
    "space_id": "Spaces-1",
    "output_file": str(HERE / "Harmony_Version_Tracker_Octopus.xlsx"),
    "verify_ssl": True,
    # Octopus environment name -> tracker column (only needed when auto-matching is wrong)
    "env_overrides": {},
    # filled on first run: Octopus project name -> {"display": ..., "include": true}
    "components": {},
}

COLUMNS = ["QA-Red", "QA-Blue", "QA-Green", "Staging", "Prod NA", "Prod EMEA", "Prod APAC"]
COL_COLORS = {"QA-Red": "FF0000", "QA-Blue": "6D7BF0", "QA-Green": "6AC46A"}

# Names already used in the tracker, used to give nice display names on the first run.
KNOWN_MODULES = [
    "API Manager", "API Gateway", "Citizen Integrator", "Cloud Agent Group Windows",
    "Cloud Agent Group Sandbox Windows", "Cloud Agent Group Sandbox Linux",
    "Cloud Agent Group Production Linux", "Cloud Studio", "Harmony (CPP)",
    "Harmony (MC & Backend)", "Identity App", "WMC Management Console UI",
    "Client Config", "Routing Engine", "EDI App", "Log API", "Log Collector",
]


# --------------------------------------------------------------------------- helpers
def load_config():
    if CONFIG_FILE.exists():
        cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        return cfg
    return json.loads(json.dumps(DEFAULT_CONFIG))


def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def version_key(v):
    """'12.13.0.74' -> (12, 13, 0, 74). Non-numeric parts count as 0, prerelease tags are ignored."""
    parts = re.split(r"[.\-+]", str(v))
    out = []
    for p in parts[:4]:
        m = re.match(r"\d+", p)
        out.append(int(m.group()) if m else 0)
    while len(out) < 4:
        out.append(0)
    return tuple(out)


def map_environment(name, overrides):
    """Map an Octopus environment name to one of COLUMNS (or None)."""
    if name in overrides:
        return overrides[name]
    # Exact names only, so variants such as "QA in Prod-EMEA", "Prod-DR-US" or
    # "Prod-East - PCAG 2" are NOT mixed into the main columns.
    n = re.sub(r"[^a-z0-9]", "", name.lower())
    # Kubernetes cloud-agent environments ("Prod-NA-KCAG", "QA-Red-KCAG", ...) belong to the same column
    if n.endswith("kcag"):
        n = n[:-4]
    exact = {
        "prodna": "Prod NA",
        "qared": "QA-Red", "qablue": "QA-Blue", "qagreen": "QA-Green",
        "staging": "Staging",
        "prodeast": "Prod NA", "prodna": "Prod NA",
        "prodemea": "Prod EMEA",
        "prodapac": "Prod APAC",
    }
    return exact.get(n)


def get_credentials():
    api_key = os.environ.get("OCTOPUS_API_KEY")
    if api_key:
        return {"api_key": api_key}
    user = os.environ.get("OCTOPUS_USERNAME") or input("Octopus username [rajesh.m]: ").strip() or "rajesh.m"
    pwd = os.environ.get("OCTOPUS_PASSWORD") or getpass.getpass(f"Octopus password for {user}: ")
    return {"username": user, "password": pwd}


def make_session(cfg, creds):
    s = requests.Session()
    s.verify = cfg["verify_ssl"]
    base = cfg["base_url"].rstrip("/")
    if "api_key" in creds:
        s.headers["X-Octopus-ApiKey"] = creds["api_key"]
    else:
        r = s.post(f"{base}/api/users/login",
                   json={"Username": creds["username"], "Password": creds["password"]}, timeout=60)
        if r.status_code != 200:
            sys.exit(f"Login failed (HTTP {r.status_code}): {r.text[:300]}\n"
                     "If your Octopus uses SSO, create an API key and set $env:OCTOPUS_API_KEY.")
        token = s.cookies.get("Octopus-Csrf-Token")
        if token:
            s.headers["X-Octopus-Csrf-Token"] = token
    return s


def get_json(s, url):
    r = s.get(url, timeout=120)
    if r.status_code == 401:
        sys.exit("Octopus rejected the credentials (HTTP 401).")
    r.raise_for_status()
    return r.json()


def fetch_deployments(s, cfg):
    """Return (projects, environments, items) from Octopus."""
    base = cfg["base_url"].rstrip("/")
    sp = cfg["space_id"]
    try:
        d = get_json(s, f"{base}/api/{sp}/dashboard")
        return d["Projects"], d["Environments"], d["Items"]
    except Exception as e:  # newer Octopus versions: use the dynamic dashboard
        print(f"  (dashboard endpoint not usable: {e}; trying dynamic dashboard)")
    projects = get_json(s, f"{base}/api/{sp}/projects/all")
    envs = get_json(s, f"{base}/api/{sp}/environments/all")
    q = "&".join([f"projects={p['Id']}" for p in projects] + [f"environments={e['Id']}" for e in envs])
    d = get_json(s, f"{base}/api/{sp}/dashboard/dynamic?{q}")
    return d.get("Projects") or projects, d.get("Environments") or envs, d["Items"]


def build_matrix(cfg, projects, envs, items):
    """matrix[project_name][column] = {version, url, state, date}"""
    env_by_id = {e["Id"]: e for e in envs}
    base = cfg["base_url"].rstrip("/")
    sp = cfg["space_id"]
    best = {}
    unmatched = set()
    for it in items:
        env = env_by_id.get(it["EnvironmentId"])
        if not env:
            continue
        col = map_environment(env["Name"], cfg["env_overrides"])
        if not col:
            unmatched.add(env["Name"])
            continue
        # keep the newest successful deployment; a currently running/failed one is a fallback
        rank = (it.get("State") == "Success", it.get("IsCurrent", False), it.get("Created", ""))
        key = (it["ProjectId"], col)
        if key not in best or rank > best[key][0]:
            best[key] = (rank, it)
    matrix = {}
    for p in projects:
        row = {}
        for col in COLUMNS:
            hit = best.get((p["Id"], col))
            if not hit:
                continue
            it = hit[1]
            ver = it["ReleaseVersion"]
            row[col] = {
                "version": ver,
                # confirmed working format on deploy.jitterbit.us
                "url": f"{base}/app#/{sp}/projects/{p['Slug']}/deployments/releases/{ver}",
                "state": it.get("State", ""),
                "date": it.get("Created", ""),
            }
        matrix[p["Name"]] = {"slug": p["Slug"], "cells": row}
    return matrix, sorted(unmatched)


def init_components(cfg, projects):
    comps = cfg["components"]
    changed = False
    for p in projects:
        if p["Name"] in comps:
            continue
        guess = difflib.get_close_matches(p["Name"], KNOWN_MODULES, n=1, cutoff=0.6)
        # Only projects that look like a known tracker component start as included;
        # with 100+ projects in Octopus you opt the others in by setting include=true.
        comps[p["Name"]] = {"display": guess[0] if guess else p["Name"], "include": bool(guess)}
        changed = True
    return changed


# --------------------------------------------------------------------------- excel
def combine_rows(cfg, matrix):
    """One row per component: Octopus versions first, then v3 fallback values for any gap.
    v3 endpoint URLs (live version pages) become the hyperlink target where one exists."""
    fallback = cfg.get("fallback", {})
    endpoints = cfg.get("endpoints", {})
    rows = {}
    for pname, info in matrix.items():
        comp = cfg["components"].get(pname, {"display": pname, "include": True})
        if not (comp.get("include", True) or cfg.get("include_all")):
            continue
        cells = rows.setdefault(comp["display"], {})
        for col, d in info["cells"].items():
            cells[col] = {"version": d["version"], "url": d["url"], "source": "octopus"}
    for display, fb in fallback.items():
        cells = rows.setdefault(display, {})
        for col in COLUMNS:
            if col not in cells and fb.get(col):
                cells[col] = {"version": fb[col], "url": "", "source": "fallback"}
    for display, cells in rows.items():
        for col, d in cells.items():
            ep = endpoints.get(display, {}).get(col)
            if ep:
                d["url"] = ep
    return rows



def write_html(cfg, ordered, latest_of):
    """Write a self-contained index.html (same data as the Excel) for publishing on Jenkins."""
    import html as H
    colors = {"QA-Red": "#ff4d4d", "QA-Blue": "#8c96f5", "QA-Green": "#7fd37f"}
    head = "".join(f'<th style="background:{colors.get(c, "#d9d9d9")}">{c}</th>' for c in ["Latest Version"] + COLUMNS)
    body = []
    for display, cells in ordered:
        lat = latest_of(cells)
        tds = []
        for col in ["Latest"] + COLUMNS:
            d = lat if col == "Latest" else cells.get(col)
            if not d:
                tds.append('<td class="na">-</td>')
                continue
            cls = "" if col == "Latest" else ("ok" if version_key(d["version"]) == version_key(lat["version"]) else "old")
            if d["source"] == "fallback":
                cls += " fb"
            v = H.escape(d["version"])
            tds.append(f'<td class="{cls}">' + (f'<a href="{H.escape(d["url"])}" target="_blank">{v}</a>' if d["url"] else v) + "</td>")
        body.append(f"<tr><th class='m'>{H.escape(display)}</th>{''.join(tds)}</tr>")
    page = f"""<!doctype html><html><head><meta charset="utf-8"><title>Harmony Version Tracker</title>
<style>body{{font-family:Arial,sans-serif;margin:24px}}table{{border-collapse:collapse}}th,td{{border:1px solid #bbb;padding:6px 12px;text-align:center}}
th.m{{text-align:left}}td.ok{{background:#a8e6cf}}td.old{{background:#fff2cc}}td.fb a,td.fb{{color:#777;font-style:italic}}td.na{{color:#888}}small{{color:#666}}</style></head>
<body><h2>Harmony Version Tracker</h2><small>Last refreshed from Octopus: {datetime.now():%Y-%m-%d %H:%M}.
Green = matches latest in row, yellow = behind, grey italic = saved value (not found in Octopus).</small><br><br>
<table><tr><th>Modules</th>{head}</tr>{''.join(body)}</table></body></html>"""
    Path(cfg.get("html_file", str(HERE / "index.html"))).write_text(page, encoding="utf-8")

def write_excel(cfg, matrix):
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    F = "Arial"
    wb = Workbook()
    ws = wb.active
    ws.title = "Harmony Version Tracker"
    ws["A1"] = "Harmony Version Tracker"
    ws["A1"].font = Font(name=F, bold=True, size=12)
    ws["C1"] = f"Last refreshed from Octopus: {datetime.now():%Y-%m-%d %H:%M}"
    ws["C1"].font = Font(name=F, italic=True, size=9)
    headers = ["Modules", "Latest Version"] + COLUMNS
    for j, h in enumerate(headers, 1):
        c = ws.cell(2, j, h)
        c.font = Font(name=F, bold=True)
        c.fill = PatternFill("solid", fgColor=COL_COLORS.get(h, "D9D9D9"))
        c.border = bd
        c.alignment = Alignment(horizontal="left" if j == 1 else "center", vertical="center")

    green = PatternFill("solid", fgColor="A8E6CF")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    link_font = Font(name=F, color="0563C1", underline="single")
    fb_font = Font(name=F, color="7F7F7F", italic=True)

    rows = combine_rows(cfg, matrix)
    known = {m: i for i, m in enumerate(KNOWN_MODULES)}
    ordered = sorted(rows.items(), key=lambda r: (known.get(r[0], 999), r[0].lower()))

    r = 3
    for display, cells in ordered:
        c = ws.cell(r, 1, display)
        c.font = Font(name=F, bold=True)
        c.border = bd
        latest = None
        if cells:
            latest_col = max(cells, key=lambda k: version_key(cells[k]["version"]))
            latest = cells[latest_col]
        c = ws.cell(r, 2, latest["version"] if latest else "-")
        c.border = bd
        c.alignment = Alignment(horizontal="center")
        c.font = Font(name=F)
        if latest and latest["url"]:
            c.hyperlink = latest["url"]
            c.font = link_font
        for j, col in enumerate(COLUMNS, 3):
            cell = ws.cell(r, j)
            cell.border = bd
            cell.alignment = Alignment(horizontal="center")
            d = cells.get(col)
            if d:
                cell.value = d["version"]
                if d["url"]:
                    cell.hyperlink = d["url"]
                cell.font = link_font if d["url"] else Font(name=F)
                if d["source"] == "fallback":
                    cell.font = Font(name=F, color="7F7F7F", italic=True,
                                     underline="single" if d["url"] else None)
                cell.fill = green if version_key(d["version"]) == version_key(latest["version"]) else yellow
            else:
                cell.value = "-"
                cell.font = Font(name=F, color="808080")
        r += 1
    ws.column_dimensions["A"].width = 36
    for col in "BCDEFGHI":
        ws.column_dimensions[col].width = 16
    ws.freeze_panes = "B3"
    notes = [
        "Latest Version = highest version in that row. Green = matches latest, yellow = behind.",
        "Versions come from Octopus. A version links to its live version page where one is known, otherwise to its Octopus release page.",
        "Grey italic = not deployed/found in Octopus; showing the last saved value from the previous tracker (v3). '-' = no value anywhere.",
    ]
    for k, t in enumerate(notes):
        ws.cell(r + 1 + k, 1, t).font = Font(name=F, italic=True, size=9)

    def _latest(cells):
        return cells[max(cells, key=lambda k: version_key(cells[k]["version"]))] if cells else None
    ordered = [(d, c) for d, c in ordered if c]
    write_html(cfg, ordered, _latest)
    out = Path(cfg["output_file"])
    try:
        wb.save(out)
    except PermissionError:
        alt = out.with_name(out.stem + f"_{datetime.now():%H%M%S}" + out.suffix)
        wb.save(alt)
        print(f"  Excel file is open - saved to {alt} instead. Close Excel to update the main file.")
        return alt
    return out


# --------------------------------------------------------------------------- main
def run_once(cfg, session, list_only=False):
    print(f"[{datetime.now():%H:%M:%S}] Reading deployments from {cfg['base_url']} ...")
    projects, envs, items = fetch_deployments(session, cfg)
    print(f"  {len(projects)} projects, {len(envs)} environments, {len(items)} deployment records")
    if list_only:
        print("\nEnvironments -> tracker column")
        for e in envs:
            print(f"  {e['Name']:30s} -> {map_environment(e['Name'], cfg['env_overrides'])}")
        names = sorted(p["Name"] for p in projects)
        listing = HERE / "octopus_projects.txt"
        listing.write_text("\n".join(names), encoding="utf-8")
        print(f"\n{len(names)} projects written to {listing}  (send me this file)")
        return
    if init_components(cfg, projects):
        save_config(cfg)
        print(f"  Config written to {CONFIG_FILE} (edit display names / include flags if needed)")
    matrix, unmatched = build_matrix(cfg, projects, envs, items)
    if unmatched:
        print("  Environments not matched to a tracker column (add to env_overrides): " + ", ".join(unmatched))
    snapshot = {n: {c: d["version"] for c, d in m["cells"].items()} for n, m in matrix.items()}
    cfg_hash = hashlib.md5(json.dumps([cfg["components"], cfg.get("fallback"), cfg.get("endpoints"), cfg.get("include_all", False)],
                                      sort_keys=True).encode()).hexdigest()
    state = json.loads(STATE_FILE.read_text()) if STATE_FILE.exists() else {}
    old = state.get("versions")
    if old == snapshot and state.get("cfg") == cfg_hash and Path(cfg["output_file"]).exists():
        print("  No new builds since the last run - Excel left as is.")
        return
    out = write_excel(cfg, matrix)
    STATE_FILE.write_text(json.dumps({"versions": snapshot, "cfg": cfg_hash}, indent=1))
    if old:
        for n in snapshot:
            for c, v in snapshot[n].items():
                if old.get(n, {}).get(c) != v:
                    print(f"  UPDATED  {n} / {c}: {old.get(n, {}).get(c, '-')} -> {v}")
    print(f"  Wrote {out}")


def main():
    ap = argparse.ArgumentParser(description="Sync Octopus deployed versions into the Harmony Version Tracker.")
    ap.add_argument("--watch", type=int, metavar="MINUTES", help="keep running and re-check every N minutes")
    ap.add_argument("--all", action="store_true", help="show EVERY Octopus project, ignoring the include flags")
    ap.add_argument("--list", action="store_true", help="list projects and environment mapping, then exit")
    ap.add_argument("--help-schedule", action="store_true", help="show how to schedule automatic refresh")
    a = ap.parse_args()
    if a.help_schedule:
        print(f"""
Schedule an automatic refresh every 15 minutes (run once in PowerShell):

  [Environment]::SetEnvironmentVariable("OCTOPUS_API_KEY","API-XXXXXXXX","User")
  schtasks /Create /SC MINUTE /MO 15 /TN "HarmonyVersionTracker" /TR "python \\"{Path(__file__).resolve()}\\"" /F

Remove it with:  schtasks /Delete /TN "HarmonyVersionTracker" /F
""")
        return
    cfg = load_config()
    if a.all:
        cfg["include_all"] = True
    creds = get_credentials()
    session = make_session(cfg, creds)
    if a.list:
        run_once(cfg, session, list_only=True)
        return
    while True:
        try:
            run_once(cfg, session)
        except Exception as e:
            print(f"  Refresh failed: {e}")
            if not a.watch:
                raise
        if not a.watch:
            break
        time.sleep(a.watch * 60)


if __name__ == "__main__":
    main()
