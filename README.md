# Okta Inventory

A tool that connects to an **Okta tenant**, pulls a full inventory of its configuration via the Okta Admin REST API, and presents it two ways: an interactive **Streamlit web app** and a downloadable **16-sheet styled Excel report**.

All API calls are **read-only** — the tool never modifies the tenant.

---

## What it does

The project has two parts:

1. **`app.py`** — a **Streamlit web app** (the deployed UI). Sidebar to enter the Okta URL + API token (or enable "Demo mode"), runs the collection live, shows the data in 13 tabbed tables plus summary metrics, and offers a **Download Excel** button.
2. **`files/okta_inventory.py`** — the **engine + CLI**. Contains the API client, all data collectors, the demo-data generator, and the Excel report builder. Can also be run standalone from the command line.

```
┌─────────────┐     imports      ┌──────────────────────┐     HTTPS      ┌──────────┐
│   app.py    │ ───────────────► │ files/okta_inventory │ ─────────────► │  Okta    │
│ (Streamlit) │                  │  OktaClient / collect │   SSWS token   │  Tenant  │
│   UI + tabs │ ◄─ data dict ──  │  build_excel / demo   │ ◄───JSON─────  │   API    │
└─────────────┘                  └──────────────────────┘                └──────────┘
```

---

## How it works — flow

1. The user enters an **Org URL** + **API token** in the sidebar (or ticks **Demo mode**) and clicks **▶ Run Inventory**.
2. `app.py` creates an `OktaClient(org_url, token)` and calls **`collect_all(client)`**, which hits ~25 Okta endpoints and returns one `data` dictionary. (Demo mode calls `generate_demo_data()` instead — no network.)
3. The `data` dict is cached in `st.session_state`, rendered into metrics + 13 tabs, and passed to **`build_excel()`**, which produces a 16-sheet `.xlsx` in memory for download.

---

## The API client (`OktaClient`)

- **Auth:** every request sends the header `Authorization: SSWS <token>` (Okta API token / "SSWS" scheme).
- **`_request()`** — GETs with a 30s timeout, retries up to 3× on 5xx errors and timeouts (exponential backoff), and handles connection errors gracefully.
- **`get_all()`** — handles **pagination** by following the `Link: rel="next"` header, and handles:
  - `401` → auth failed (bad token)
  - `403` → permission denied (skips the endpoint)
  - `429` → **rate limited** — reads the `X-Rate-Limit-Reset` header and sleeps until reset
- **`get_one()`** — single-object GET (e.g. org settings).
- **Parallelism** — `fetch_app_assignments_parallel()` and `fetch_policy_rules_parallel()` use a `ThreadPoolExecutor` (5 workers) to fetch per-app and per-policy sub-resources concurrently (the slow part of collection).

> **Note:** the client currently sets `verify=False` (TLS verification disabled, warnings suppressed). This is convenient for a quick internal tool but should be tightened for production use.

---

## APIs called (`collect_all`)

| Category | Endpoint | Purpose |
|---|---|---|
| Applications | `GET /api/v1/apps?limit=200` | All apps |
| Users (by status) | `GET /api/v1/users?...filter=status eq "ACTIVE"` (+ STAGED, DEPROVISIONED, SUSPENDED, LOCKED_OUT, PASSWORD_EXPIRED) | Users split by lifecycle status |
| Groups | `GET /api/v1/groups?limit=200` | All groups |
| Identity Providers | `GET /api/v1/idps` | External IdPs / directories |
| Directory agents | `GET /api/v1/agentPools` | AD/LDAP agent pools |
| Policies | `GET /api/v1/policies?type=ACCESS_POLICY` / `MFA_ENROLL` / `PASSWORD` / `OKTA_SIGN_ON` | 4 policy types |
| Policy rules | `GET /api/v1/policies/{id}/rules` (parallel) | Rules per policy |
| Features | `GET /api/v1/features` | Enabled/disabled features |
| Trusted origins | `GET /api/v1/trustedOrigins` | CORS/redirect origins |
| API tokens | `GET /api/v1/api-tokens` | Active admin tokens |
| Org settings | `GET /api/v1/org` | Company name, subdomain |
| Network zones | `GET /api/v1/zones` | IP zones |
| Authorization servers | `GET /api/v1/authorizationServers` | Custom OAuth servers |
| Authenticators | `GET /api/v1/authenticators` | MFA authenticators |
| Group rules | `GET /api/v1/groups/rules?limit=200` | Dynamic group assignment rules |
| Event hooks | `GET /api/v1/eventHooks` | Outbound event webhooks |
| Inline hooks | `GET /api/v1/inlineHooks` | Inline transform hooks |
| Log streams | `GET /api/v1/logStreams` | SIEM log streaming |
| ThreatInsight | `GET /api/v1/threats/configuration` | Threat protection config |
| App assignments | `GET /api/v1/apps/{id}/users` + `/apps/{id}/groups` (parallel) | Per-app user counts & assigned groups |

---

## The Excel builder (`build_excel`)

Uses **openpyxl** to generate a styled workbook with a "Charcoal & Teal" theme (colored headers, status color-coding, auto-filters, frozen panes, metric cards). 16 sheets:

**Summary, Applications, Users, Groups, Identity Providers, Features, Policies, Policy Rules, OAG Applications, Network Zones, Auth Servers, Authenticators, Group Rules, Hooks, Log Streams,** and a manual-entry template.

The OAG sheet heuristically flags Okta Access Gateway apps by label keywords.

---

## Demo mode

`generate_demo_data()` builds a realistic fake tenant (30 apps, 150 users, groups, policies, hooks, etc.) so the UI and Excel layout can be previewed **without any credentials**. Enable it with the sidebar checkbox in the web app, or `--demo` on the CLI.

---

## Running the app

### Web app (Streamlit)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL, enter your Okta Org URL + API token (or tick **Demo mode**), and click **▶ Run Inventory**.

### CLI

```bash
# Live tenant
python files/okta_inventory.py --url https://your-org.okta.com --token <SSWS_TOKEN>

# Sample data, no token
python files/okta_inventory.py --demo
```

Outputs an `.xlsx` report (default `Okta_Inventory_Ross.xlsx`).

---

## Deployment (Streamlit Community Cloud)

When deploying this repo, set:

- **Main file path:** `app.py`  *(not `okta_inventory.py` — that file is the backend engine and has no UI, which renders a blank screen)*
- **Branch:** `main`
- **Python version:** pin via the app's **Advanced settings** at deploy time (the `runtime.txt` file is ignored by Streamlit Cloud — it's a Heroku convention).

---

## Requirements

```
requests >= 2.31.0
openpyxl >= 3.1.0
pandas   >= 2.0.0
streamlit >= 1.28.0
urllib3  >= 2.0.0
```

---

## Security notes

- The API token is used **only** for the live API calls and is never stored by the app.
- The HTTP client currently disables TLS certificate verification (`verify=False`); consider enabling it for production.
- `.streamlit/secrets.toml` and `*.xlsx` outputs are git-ignored so tokens and generated reports are never committed.
