# powerplatform-copilot-credit-limits

Apply **Copilot Credit limits to Copilot Studio agents** across a single Power Platform environment,
a list of environments, or every environment in an **environment group** — with exclusions, a dry-run
mode, and a CSV report of exactly what changed.

Companion to [`powerplatform-tenant-pool-draw`](https://github.com/LaureVDH/powerplatform-tenant-pool-draw).

> **Verified.** Tested end to end against a live tenant, and confirmed in the Power Platform admin
> center: a limit written by this script appears under **Licensing → Copilot Studio → Manage Agents**
> with a *Within Limit* status. That is the enforcement surface, not just the underlying table.

---

## ⚠️ The API version trap

If you are writing your own automation against this API, this will save you a day.

The published licensing spec documents `api-version=2024-10-01`. The admin center calls the same
route with **`api-version=1`**. They do not behave the same:

| Version | Result |
| --- | --- |
| `2024-10-01` | `200 OK`, the row persists in `resourceThresholds` — but the limit **never appears in Manage Agents and does not enforce** |
| `1` | `200 OK`, and the limit appears in Manage Agents and enforces |

Worse, the endpoint **validates nothing**. It will accept a random GUID that belongs to no agent, and
store it. `DELETE` is not supported, so a bad row cannot be removed through the API.

**A `200` response is not evidence that anything was configured. Verify in the admin center.**

This script uses `api-version=1` and falls back to the documented version.

---

## What it does

| Scope | Behaviour |
| --- | --- |
| **Agents** | Discovered across the target environments, then written with a monthly credit limit, a notification threshold, and an optional hard stop. |
| **Flows** | Agent flows and cloud flows consume credits, so they are **inventoried and reported**. They are **not written to by default** — see [Flows](#flows). |

### Agent discovery

Three sources, merged:

| Source | Gives you | Caveat |
| --- | --- | --- |
| **Inventory** (`resourcequery`) | Every agent and flow, tenant-wide, in one call. **Admin-scoped.** | — |
| **Licensing** | Month-to-date consumption. | Only agents that have consumed credits. Can return 403. |
| **Dataverse** (`bots`) | Display names. | **Requires environment membership** — see below. |

**Why Inventory matters.** Being a Power Platform administrator does **not** grant access to the
Dataverse inside someone else's environment — the Copilot Studio portal fails there too. Personal
developer environments therefore return **HTTP 403** to a `bots` query. The inventory API is
control-plane and admin-scoped, so it sees into them anyway. Without it, PDEs are invisible — and
they are exactly the environments that tend to consume credits unintentionally.

---

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7
- `Az.Accounts` — `Install-Module Az.Accounts -Scope CurrentUser`
- Power Platform Administrator or Global Administrator

---

## Quick start

Run these in order. The first two change nothing.

```powershell
# 1. Inventory the agents and show their current limits. Read-only.
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -ReportOnly

# 2. Dry run - shows exactly which agents would be changed.
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -AgentCreditLimit 1000 -WhatIf

# 3. Apply to one environment.
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -AgentCreditLimit 1000

# 4. Dry run across a whole environment group.
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" -AgentCreditLimit 500 -WhatIf

# 5. Apply across the group, with exclusions and a hard stop.
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" `
    -AgentCreditLimit 500 -StopAgentIfOverLimit `
    -ExcludeAgentIdCsv .\samples\agent-exclusions.csv -Force
```

Then confirm in **Licensing → Copilot Studio → Manage Agents**.

---

## Parameters

| Parameter | Description |
| --- | --- |
| `-EnvironmentGroup` | Environment group ID or exact display name. All environments in the group are targeted. |
| `-Environment` | One or more environment IDs or exact display names. |
| `-AgentCreditLimit` | Monthly Copilot Credit limit applied to each agent. |
| `-AgentNotificationThreshold` | Notify at this **percentage** of the limit. Default 80. The admin center restricts this to **50-100**. |
| `-NotifyIfOverCapacity` | Notify when an agent exceeds its limit. On by default. |
| `-StopAgentIfOverLimit` | Hard stop: turn the agent off at the limit. |
| `-ExcludeAgentId` / `-ExcludeAgentIdCsv` | Agents to skip — inline IDs, or a CSV. |
| `-IncludeFlows` | Also write limits to flows. See [Flows](#flows). |
| `-AgentSource` | `Inventory`, `Licensing`, `Dataverse`, or `All` (default). |
| `-LookbackDays` | Consumption snapshot window. Default 90. |
| `-ReportPath` | Folder for the CSV report. Default `.\reports`. |
| `-Discover` | Read-only raw API dump for the first target environment. |
| `-ReportOnly` | Inventory current limits without writing. |
| `-TenantId` | Tenant to sign in to. |
| `-Force` | Skip confirmations; rewrite limits even when unchanged. |
| `-Diagnostics` | Extra detail about discovery and API fallbacks. |
| `-WhatIf` / `-Confirm` | Standard PowerShell safety switches. |

### Exclusions

A header named `AgentId`, `ResourceId`, `BotId` or `Id` is used when present; otherwise every
non-empty line is treated as an ID. See `samples\agent-exclusions.csv`. A single exclusion can be
passed inline: `-ExcludeAgentId 1111...-1111`.

> **Safety:** if an exclusion file is supplied but yields **zero** exclusions — because it is empty or
> its column was not recognised — the script **stops**. Passing the parameter means you intend to
> protect something, and continuing would apply the limit to *every* agent. Exclusions are also
> validated against the discovered resources **before** anything is written, so a typo or stale ID is
> reported rather than silently protecting nothing.

---

## Flows

Agent flows and cloud flows can consume Copilot Credits — AI Builder actions and agent flow actions
are common cost drivers — so they are **always inventoried and reported**, with a `ResourceType`
column in the CSV.

They are **not written to by default.** `-IncludeFlows` targets them deliberately.

The API accepts limits on flows and behaves identically to agents — including returning
`NoAvailableCapacitySource` for flows in environments without a capacity source, and succeeding once
one is present. That is good evidence they are handled by the same subsystem. **Enforcement has not
yet been confirmed in the admin center**, so verify there before relying on a flow limit.

---

## Report

Every run writes `reports\CopilotCreditLimits-yyyyMMdd-HHmmss.csv` and prints a summary.

`Timestamp, Scope, EnvironmentName, EnvironmentId, TargetId, TargetName, Source, ResourceType,
Action, PreviousLimit, NewLimit, NotificationThresholdPct, StopIfOverCapacity, Consumed, Message`

| `Action` | Meaning |
| --- | --- |
| `Set` | Limit applied. |
| `Skipped-Excluded` | Listed in the exclusions. |
| `Skipped-NoChange` | Already has this limit. |
| `Skipped-Flow` | A flow, and `-IncludeFlows` was not used. |
| `WhatIf` / `ReportOnly` | No change written. |
| `NoAgentsFound` | Environment inspected, genuinely empty. |
| `NotInspected` | Environment **could not be read** — its agents are **unknown**. |
| `Failed` | The write failed; see `Message`. |

> `NoAgentsFound` and `NotInspected` mean different things. Treat `NotInspected` rows as gaps in
> coverage, not clean results. Every run ends with an **INCOMPLETE COVERAGE** section if any exist.

---

## What this does not do

### Environments with no capacity source

A write can fail with:

```
HTTP 400 - NoAvailableCapacitySource - No environment allocation or tenant pool found.
```

An environment can consume Copilot Credits through exactly three routes: an **allocation**,
**tenant-pool draw**, or a linked **pay-as-you-go billing plan**. If it has none of them, there is
nothing for a per-agent limit to apply against, and the API says so.

**This failure is usually good news.** An environment with no capacity source cannot consume credits
at all — which is what a limit was trying to achieve. Treat it as confirmation, not a problem.

> Do **not** add a billing plan or allocation just to make this script succeed. That would enable
> spending in order to cap it. Give an environment a capacity source only when you intend it to
> consume.

Verified by controlled test: the same command against the same environments failed with
`NoAvailableCapacitySource` before a billing plan was linked, and succeeded for every agent and flow
afterwards. Nothing else changed.

### Per-user limits

**There is no way to cap Copilot Studio credits per user.** Tested against a live tenant: no control
in the Power Platform admin center, nothing in Entra or Azure, and every plausible API route returns
"not found".

- The admin center now has a per-user **consumption view** (Licensing → Copilot Studio → Users). That
  is reporting, not a limit.
- Microsoft 365 **Cost Management** spending policies *do* offer a hard per-user monthly cap scoped to
  an Entra group — but that experience is currently scoped to **Cowork and Work IQ API**, so it does
  not govern Copilot Studio agent credits.
- The older pay-as-you-go billing **budget only sends alerts. It does not stop consumption.**

The enforceable controls today are **environment allocation**, **tenant-pool draw**, and **per-agent
limits**.

---

## Scheduling

The script is idempotent (`Skipped-NoChange`) and non-interactive with `-Force`, so it can run on a
schedule to re-assert limits as new agents appear. Unattended runs need a non-interactive auth path
(service principal or managed identity) rather than the interactive `Connect-AzAccount` used here.

---

## Notes and limits

- Entitlement ID for Copilot Credits is `MCSMessages`.
- `resourceThresholds` is **tenant-wide**: records must be matched on `environmentId` **and**
  `resourceId`. Matching on `resourceId` alone lets a threshold from one environment be mistaken for
  another's.
- Date parameters are **camelCase** (`fromDate` / `toDate`). Kebab-case is the `pac` CLI flag
  spelling and is rejected with HTTP 400.
- A published environment-group rule can lock capacity settings (`TenantPoolLockedByPolicy`).
- Provided **as-is**, as personal tooling rather than a Microsoft product. Test it in a
  non-production environment first.

See [`docs/API-GUIDE.md`](docs/API-GUIDE.md) for a map of the licensing, governance, copilotstudio
and usermanagement namespaces.

## References

- [Manage Copilot Credits and capacity for Copilot Studio](https://learn.microsoft.com/power-platform/admin/manage-copilot-studio-copilot-credits-capacity)
- [Tutorial: Manage Copilot Credits allocations programmatically](https://learn.microsoft.com/power-platform/admin/programmability-tutorial-manage-copilot-credit-allocations)
- [Power Platform licensing OpenAPI spec](https://github.com/MicrosoftDocs/power-platform/blob/main/power-platform/developer/reference/licensing/licensing.json)

## Prior art

[`jameswh3/MW-Toolbox`](https://github.com/jameswh3/MW-Toolbox) includes
`Set-CopilotAgentConsumptionLimit.ps1`, which sets the threshold for a **single** agent. This repo
covers the bulk case: environment-group enumeration, multi-environment runs, admin-scoped discovery,
exclusions, dry runs and reporting.

## License

MIT
