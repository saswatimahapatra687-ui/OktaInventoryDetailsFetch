#!/usr/bin/env python3
"""Streamlit web interface for the Okta Inventory script."""

import io
import sys
import os
import contextlib

import pandas as pd
import streamlit as st

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "files"))
try:
    from okta_inventory import OktaClient, collect_all, build_excel, generate_demo_data
except Exception as _import_err:
    import streamlit as st
    st.error(f"Failed to import okta_inventory: {_import_err}")
    st.stop()

st.set_page_config(
    page_title="Okta Inventory",
    page_icon="🔐",
    layout="wide",
)

# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("🔐 Okta Inventory")
    st.caption("Okta Tenant Report")
    st.divider()

    org_url   = st.text_input("Okta Org URL", placeholder="https://your-org.okta.com")
    api_token = st.text_input("API Token (SSWS)", type="password", placeholder="00...")
    demo_mode = st.checkbox("Demo mode (no token needed)")
    st.divider()
    run_btn   = st.button("▶  Run Inventory", type="primary", use_container_width=True)

    st.markdown(
        "<small>The token is used only for API calls and is never stored.</small>",
        unsafe_allow_html=True,
    )

# ── Run logic ─────────────────────────────────────────────────────────────────

if run_btn:
    if not demo_mode and (not org_url.strip() or not api_token.strip()):
        st.sidebar.error("Enter Org URL and API Token, or enable Demo mode.")
    else:
        log_buf = io.StringIO()
        with st.status("Collecting Okta inventory…", expanded=True) as status:
            try:
                with contextlib.redirect_stdout(log_buf):
                    if demo_mode:
                        data = generate_demo_data()
                    else:
                        client = OktaClient(org_url.strip(), api_token.strip())
                        data = collect_all(client)
                st.code(log_buf.getvalue() or "Collection complete.", language=None)
                status.update(label="✅ Collection complete", state="complete", expanded=False)
            except Exception as e:
                status.update(label="❌ Collection failed", state="error")
                st.error(str(e))
                st.stop()

        excel_buf = io.BytesIO()
        with contextlib.redirect_stdout(io.StringIO()):
            build_excel(data, excel_buf)
        excel_buf.seek(0)

        st.session_state["data"]        = data
        st.session_state["excel_bytes"] = excel_buf.getvalue()

# ── Guard: nothing collected yet ──────────────────────────────────────────────

if "data" not in st.session_state:
    st.markdown("""
## Welcome
Configure your Okta org in the sidebar and click **▶ Run Inventory**.

The report is collected live via the Okta API and displayed here.
You can then download the full 16-sheet Excel workbook.

> **Demo mode** generates sample data so you can preview the layout without an API token.
    """)
    st.stop()

data = st.session_state["data"]
org  = data.get("org", {})

# ── Header + download ─────────────────────────────────────────────────────────

col_title, col_dl = st.columns([5, 1])
with col_title:
    subdomain = org.get("subdomain", "")
    company   = org.get("companyName", "")
    domain    = f"{subdomain}.okta.com" if subdomain else ""
    # Avoid repeating the tenant name when companyName just mirrors the subdomain
    if company and subdomain and company.strip().lower() != subdomain.strip().lower():
        label = f"{company}  ·  {domain}"
    else:
        label = domain or company or "Okta Inventory"
    st.subheader(label)
with col_dl:
    st.download_button(
        label="📥 Download Excel",
        data=st.session_state["excel_bytes"],
        file_name="Okta_Inventory.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

st.divider()

# ── Summary metrics ───────────────────────────────────────────────────────────

apps      = data.get("apps", [])
all_users = (
    data.get("users", []) + data.get("users_staged", []) +
    data.get("users_suspended", []) + data.get("users_locked", []) +
    data.get("users_password_expired", []) + data.get("users_deprovisioned", [])
)

c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
c1.metric("Apps",           len(apps))
c2.metric("Active Users",   len(data.get("users", [])))
c3.metric("Total Users",    len(all_users))
c4.metric("Groups",         len(data.get("groups", [])))
c5.metric("IdPs",           len(data.get("idps", [])))
c6.metric("Auth Servers",   len(data.get("auth_servers", [])))
c7.metric("Authenticators", len(data.get("authenticators", [])))

st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────

tabs = st.tabs([
    "📱 Apps", "👥 Users", "🗂️ Groups", "🔗 Identity Providers",
    "📲 Authenticators", "🌐 Network Zones", "🔑 Auth Servers",
    "📏 Group Rules", "🪝 Hooks", "📡 Log Streams",
    "🛡️ Policies", "📜 Policy Rules", "⚙️ Features",
])

assignments = data.get("app_assignments", {})

# ── Apps ──────────────────────────────────────────────────────────────────────
with tabs[0]:
    rows = []
    for app in apps:
        asgn = assignments.get(app.get("id", ""), {})
        rows.append({
            "Name":         app.get("label", ""),
            "Status":       app.get("status", ""),
            "Sign-On":      app.get("signOnMode", ""),
            "App Type":     app.get("name", ""),
            "Users":        asgn.get("user_count", ""),
            "Groups":       ", ".join(asgn.get("group_names") or []),
            "Provisioning": ", ".join(app.get("features") or []),
            "Created":      (app.get("created") or "")[:10],
            "ID":           app.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Users ─────────────────────────────────────────────────────────────────────
with tabs[1]:
    rows = []
    for key, label in [
        ("users", "ACTIVE"), ("users_staged", "STAGED"),
        ("users_suspended", "SUSPENDED"), ("users_locked", "LOCKED_OUT"),
        ("users_password_expired", "PASSWORD_EXPIRED"),
        ("users_deprovisioned", "DEPROVISIONED"),
    ]:
        for u in data.get(key, []):
            p = u.get("profile", {})
            rows.append({
                "First Name":  p.get("firstName", ""),
                "Last Name":   p.get("lastName", ""),
                "Email":       p.get("email", p.get("login", "")),
                "Status":      label,
                "Created":     (u.get("created") or "")[:10],
                "Last Login":  (u.get("lastLogin") or "")[:10] if u.get("lastLogin") else "Never",
                "ID":          u.get("id", ""),
            })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Groups ────────────────────────────────────────────────────────────────────
with tabs[2]:
    rows = []
    for g in data.get("groups", []):
        p = g.get("profile", {})
        rows.append({
            "Name":        p.get("name", ""),
            "Type":        g.get("type", ""),
            "Description": p.get("description", ""),
            "Created":     (g.get("created") or "")[:10],
            "ID":          g.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Identity Providers ────────────────────────────────────────────────────────
with tabs[3]:
    rows = []
    for idp in data.get("idps", []):
        rows.append({
            "Name":         idp.get("name", ""),
            "Type":         idp.get("type", ""),
            "Status":       idp.get("status", ""),
            "Protocol":     (idp.get("protocol") or {}).get("type", ""),
            "Created":      (idp.get("created") or "")[:10],
            "Last Updated": (idp.get("lastUpdated") or "")[:10],
            "ID":           idp.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Authenticators ────────────────────────────────────────────────────────────
with tabs[4]:
    rows = []
    for a in data.get("authenticators", []):
        rows.append({
            "Name":         a.get("name", ""),
            "Key":          a.get("key", ""),
            "Type":         a.get("type", ""),
            "Status":       a.get("status", ""),
            "Created":      (a.get("created") or "")[:10],
            "Last Updated": (a.get("lastUpdated") or "")[:10],
            "ID":           a.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Network Zones ─────────────────────────────────────────────────────────────
with tabs[5]:
    rows = []
    for z in data.get("zones", []):
        rows.append({
            "Name":     z.get("name", ""),
            "Type":     z.get("type", ""),
            "Status":   z.get("status", ""),
            "Gateways": ", ".join(
                f"{g.get('type','')}: {g.get('value','')}" for g in (z.get("gateways") or [])
            ),
            "Proxies":  ", ".join(
                f"{p.get('type','')}: {p.get('value','')}" for p in (z.get("proxies") or [])
            ),
            "Created":  (z.get("created") or "")[:10],
            "ID":       z.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Auth Servers ──────────────────────────────────────────────────────────────
with tabs[6]:
    rows = []
    for s in data.get("auth_servers", []):
        rows.append({
            "Name":        s.get("name", ""),
            "Description": s.get("description", ""),
            "Audiences":   ", ".join(s.get("audiences") or []),
            "Issuer":      s.get("issuer", ""),
            "Status":      s.get("status", ""),
            "Created":     (s.get("created") or "")[:10],
            "ID":          s.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Group Rules ───────────────────────────────────────────────────────────────
with tabs[7]:
    rows = []
    for r in data.get("group_rules", []):
        rows.append({
            "Name":          r.get("name", ""),
            "Status":        r.get("status", ""),
            "Expression":    (r.get("conditions") or {}).get("expression", {}).get("value", ""),
            "Target Groups": ", ".join(
                (r.get("actions") or {}).get("assignUserToGroups", {}).get("groupIds") or []
            ),
            "ID":            r.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Hooks ─────────────────────────────────────────────────────────────────────
with tabs[8]:
    rows = []
    for h in data.get("event_hooks", []):
        rows.append({
            "Name":     h.get("name", ""),
            "Type":     "Event Hook",
            "Status":   h.get("status", ""),
            "Endpoint": (h.get("channel") or {}).get("config", {}).get("uri", ""),
            "Events":   ", ".join((h.get("events") or {}).get("items") or []),
            "Created":  (h.get("created") or "")[:10],
            "ID":       h.get("id", ""),
        })
    for h in data.get("inline_hooks", []):
        rows.append({
            "Name":     h.get("name", ""),
            "Type":     "Inline Hook",
            "Status":   h.get("status", ""),
            "Endpoint": (h.get("channel") or {}).get("config", {}).get("uri", ""),
            "Events":   h.get("type", ""),
            "Created":  (h.get("created") or "")[:10],
            "ID":       h.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Log Streams ───────────────────────────────────────────────────────────────
with tabs[9]:
    rows = []
    for ls in data.get("log_streams", []):
        rows.append({
            "Name":         ls.get("name", ""),
            "Type":         ls.get("type", ""),
            "Status":       ls.get("status", ""),
            "Created":      (ls.get("created") or "")[:10],
            "Last Updated": (ls.get("lastUpdated") or "")[:10],
            "ID":           ls.get("id", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Policies ──────────────────────────────────────────────────────────────────
with tabs[10]:
    rows = []
    for key, label in [
        ("auth_policies", "Authentication"), ("mfa_policies", "MFA Enrollment"),
        ("password_policies", "Password"),   ("signon_policies", "Sign-On"),
    ]:
        for p in data.get(key, []):
            rows.append({
                "Name":         p.get("name", ""),
                "Policy Type":  label,
                "Status":       p.get("status", ""),
                "Priority":     p.get("priority", ""),
                "Created":      (p.get("created") or "")[:10],
                "Last Updated": (p.get("lastUpdated") or "")[:10],
                "ID":           p.get("id", ""),
            })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Policy Rules ──────────────────────────────────────────────────────────────
with tabs[11]:
    pol_lookup = {}
    for key, label in [
        ("auth_policies", "Authentication"), ("mfa_policies", "MFA Enrollment"),
        ("password_policies", "Password"),   ("signon_policies", "Sign-On"),
    ]:
        for p in data.get(key, []):
            pol_lookup[p["id"]] = (p.get("name", ""), label)

    rows = []
    for pol_id, rules in (data.get("policy_rules") or {}).items():
        pol_name, pol_type = pol_lookup.get(pol_id, ("", ""))
        for rule in (rules or []):
            signon = (rule.get("actions") or {}).get("signon") or {}
            rows.append({
                "Policy Name":  pol_name,
                "Policy Type":  pol_type,
                "Rule Name":    rule.get("name", ""),
                "Priority":     rule.get("priority", ""),
                "Status":       rule.get("status", ""),
                "Access":       signon.get("access", ""),
                "Require MFA":  signon.get("requireFactor", ""),
                "ID":           rule.get("id", ""),
            })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Features ──────────────────────────────────────────────────────────────────
with tabs[12]:
    rows = []
    for f in data.get("features", []):
        stage = f.get("stage", {})
        rows.append({
            "Name":        f.get("name", ""),
            "Status":      f.get("status", ""),
            "Stage":       (stage.get("value", "") if isinstance(stage, dict) else stage or ""),
            "Type":        f.get("type", ""),
            "Description": f.get("description", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
