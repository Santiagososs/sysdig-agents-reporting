# sysdig-agent-report

Report which nodes have a Sysdig Agent installed, and in what state. It pulls
the agent inventory from the Sysdig API, optionally filters by agent status
(default: `Up to Date`), and writes a CSV or Excel report.

Columns: hostname, cluster, agent status, version, last seen, connected,
deployment type, provider, account ID, agent ID. The XLSX also has a Summary
sheet with counts by status for all agents fetched.

## Setup

Python 3.8+. The CSV output needs no dependencies; XLSX needs `openpyxl`.

```bash
pip install -r requirements.txt   # only for XLSX
export SYSDIG_URL=https://<your-region>.app.sysdig.com
export SYSDIG_API_TOKEN=<your Sysdig API token>
```

Use the URL of your Sysdig region. Never commit or pass the token on the command line.

## Usage

```bash
./sysdig_agent_report.py -o agents.csv              # Up to Date agents only
./sysdig_agent_report.py -o agents.xlsx             # Excel output
./sysdig_agent_report.py --status all -o all.xlsx   # every agent, any status
./sysdig_agent_report.py --status "Never Connected" # a specific status
./sysdig_agent_report.py --input sample.json        # offline test, no API call
```

## TLS errors

`CERTIFICATE_VERIFY_FAILED` usually means Python has no CA bundle (python.org
builds on macOS: run `Install Certificates.command`, or `pip install certifi`),
or you are behind a TLS-inspecting proxy: use `--ca-bundle corp-ca.pem` or
`SSL_CERT_FILE`. `--insecure` disables verification; use only for testing.

## Endpoint

`GET /api/cloud/v2/dataSources/agents?limit=&offset=` (paginated, with a bearer
API token). This is the endpoint behind the Agents page in the Sysdig UI and
may change without notice.

## License

MIT
