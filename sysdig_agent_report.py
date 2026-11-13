#!/usr/bin/env python3
"""Report which nodes have a Sysdig Agent installed, and its status.

Pulls /api/cloud/v2/dataSources/agents (paginated), optionally filters by agent
status, and writes a CSV (node, cluster, status, version, last seen, ...).

Auth: Sysdig API token via SYSDIG_API_TOKEN env var (Settings > User Profile >
Sysdig API Token). Never pass it on the command line or commit it.

Examples:
  export SYSDIG_URL=https://<your-region>.app.sysdig.com
  export SYSDIG_API_TOKEN=...
  ./sysdig_agent_report.py                         # only "Up to Date" agents
  ./sysdig_agent_report.py --status all            # every agent, any status
  ./sysdig_agent_report.py -o report.xlsx          # Excel (needs openpyxl)
  ./sysdig_agent_report.py --input sample.json     # offline test with saved JSON
"""
import argparse
import csv
import json
import os
import ssl
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone

ENDPOINT = "/api/cloud/v2/dataSources/agents"

COLUMNS = [
    "hostname", "cluster", "agent_status", "agent_version", "last_seen",
    "connected", "deployment_type", "provider", "account_id", "agent_id",
]


def make_ssl_context(ca_bundle=None, insecure=False):
    """Default trust store; certifi if present; or a custom CA bundle (corporate proxy)."""
    if insecure:
        print("WARNING: TLS verification disabled", file=sys.stderr)
        return ssl._create_unverified_context()
    ca_bundle = ca_bundle or os.environ.get("SSL_CERT_FILE") or os.environ.get("REQUESTS_CA_BUNDLE")
    if ca_bundle:
        return ssl.create_default_context(cafile=ca_bundle)
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        return ssl.create_default_context()


SSL_CTX = None


def fetch_page(base, token, limit, offset):
    url = f"{base.rstrip('/')}{ENDPOINT}?limit={limit}&offset={offset}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as resp:
            return json.load(resp)
    except urllib.error.HTTPError as e:
        sys.exit(f"HTTP {e.code} from Sysdig: {e.read().decode(errors='replace')[:300]}")
    except urllib.error.URLError as e:
        hint = ""
        if "CERTIFICATE_VERIFY_FAILED" in str(e.reason):
            hint = ("\nTLS trust problem. Try: pip install certifi, or --ca-bundle <corp-ca.pem> "
                    "(needed behind a TLS-inspecting proxy such as Zscaler).")
        sys.exit(f"Connection error: {e.reason}{hint}")


def fetch_all(base, token, limit):
    details, offset = [], 0
    while True:
        page = fetch_page(base, token, limit, offset)
        batch = page.get("details") or []
        details.extend(batch)
        if len(batch) < limit:
            return details
        offset += limit


def to_row(d):
    labels = d.get("labels") or {}
    return {
        "hostname": labels.get("hostname", ""),
        "cluster": d.get("clusterName", ""),
        "agent_status": d.get("agentStatus", ""),
        "agent_version": d.get("agentVersion", ""),
        "last_seen": d.get("agentLastSeen", ""),
        "connected": d.get("connected", False),
        "deployment_type": d.get("deploymentType", ""),
        "provider": labels.get("provider", ""),
        "account_id": labels.get("accountId", ""),
        "agent_id": labels.get("agentId", ""),
    }


def write_csv(path, rows):
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        w.writeheader()
        w.writerows(rows)


def write_xlsx(path, rows, counts):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        sys.exit("XLSX output needs openpyxl: pip install openpyxl")
    wb = Workbook()
    ws = wb.active
    ws.title = "Agents"
    ws.append([c.replace("_", " ").title() for c in COLUMNS])
    for r in rows:
        ws.append([r[c] for c in COLUMNS])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(COLUMNS, 1):
        width = max(len(col), *(len(str(r[col])) for r in rows)) if rows else len(col)
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 50)
    summ = wb.create_sheet("Summary")
    summ.append(["Agent status (all fetched)", "Count"])
    for k, v in sorted(counts.items()):
        summ.append([k, v])
    summ.append(["Total", sum(counts.values())])
    summ["A1"].font = summ["B1"].font = Font(bold=True)
    summ.column_dimensions["A"].width = 30
    wb.save(path)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--base-url", default=os.environ.get("SYSDIG_URL"),
                   help="Sysdig region URL, e.g. https://us2.app.sysdig.com (or set SYSDIG_URL)")
    p.add_argument("--status", default="Up to Date",
                   help='Agent status to keep (case-insensitive), or "all". Default: "Up to Date"')
    p.add_argument("--limit", type=int, default=50, help="Page size (default 50)")
    p.add_argument("--format", choices=["csv", "xlsx"], default="csv",
                   help="Output format (default csv; inferred from --output extension if given)")
    p.add_argument("--output", "-o",
                   help="Output file (default sysdig_agents_<timestamp>.<format>)")
    p.add_argument("--ca-bundle", help="PEM file of trusted CAs (also read from SSL_CERT_FILE)")
    p.add_argument("--insecure", action="store_true", help="Disable TLS verification (last resort)")
    p.add_argument("--input", help="Read a saved JSON response instead of calling the API (testing)")
    args = p.parse_args()
    global SSL_CTX
    SSL_CTX = make_ssl_context(args.ca_bundle, args.insecure)

    if args.input:
        with open(args.input) as f:
            data = json.load(f)
        details = data.get("details") or []
    else:
        if not args.base_url:
            sys.exit("Set --base-url or SYSDIG_URL to your Sysdig region URL.")
        token = os.environ.get("SYSDIG_API_TOKEN")
        if not token:
            sys.exit("Set SYSDIG_API_TOKEN in the environment.")
        details = fetch_all(args.base_url, token, args.limit)

    fmt = args.format
    if args.output and args.output.lower().endswith(".xlsx"):
        fmt = "xlsx"
    elif args.output and args.output.lower().endswith(".csv"):
        fmt = "csv"
    out = args.output or f"sysdig_agents_{datetime.now(timezone.utc):%Y%m%d_%H%M%S}.{fmt}"

    rows = [to_row(d) for d in details]
    if args.status.lower() != "all":
        rows = [r for r in rows if r["agent_status"].lower() == args.status.lower()]
    rows.sort(key=lambda r: (r["cluster"], r["hostname"]))

    # Summary of everything fetched, so filtered-out agents are visible.
    counts = {}
    for d in details:
        s = d.get("agentStatus", "Unknown")
        counts[s] = counts.get(s, 0) + 1
    if fmt == "xlsx":
        write_xlsx(out, rows, counts)
    else:
        write_csv(out, rows)
    print(f"Fetched {len(details)} agents: " + ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))
    print(f"Wrote {len(rows)} rows (status filter: {args.status}) -> {out}")


if __name__ == "__main__":
    main()
