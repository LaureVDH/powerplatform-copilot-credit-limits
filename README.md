# powerplatform-copilot-credit-limits

Automate **Copilot Credit limits** across Microsoft Copilot Studio / Power Platform — for every **agent**
in an environment or in every environment of an **environment group** — and get a CSV report of exactly
what was changed.

Companion to [`powerplatform-tenant-pool-draw`](https://github.com/LaureVDH/powerplatform-tenant-pool-draw):
same design (interactive `Az.Accounts` sign-in, Power Platform API, `-WhatIf`, tolerant property resolution).

> **Status:** tested end to end against a live tenant (19 environments, 18 agents).
> Write → read-back verified: limits applied through the API are returned by a subsequent read.

---

## What it does

| Scope | What happens | API |
| --- | --- | --- |
| **Agents** | Lists every agent in the target environment(s), then writes a per-agent monthly Copilot Credit limit, notification threshold and optional hard stop. | `PUT licensing/environments/{envId}/entitlements/MCSMessages/resources/{resourceId}/threshold` |
| **Users** | Lists every maker/user of the target environment(s), annotates Entra group membership and credit consumption, and reports. **Read-only — see below.** | `GET licensing/entitlements/MCSMessages/users`, Dataverse `systemusers` |

### Why users are report-only

There are **two separate Copilot Credit limit systems**, and only one of them is programmable.

**1. Copilot Studio / Power Platform credits (`MCSMessages`)** — what this script writes.
Limits here apply to **agents only**. Verified against the published Power Platform OpenAPI specs
(`licensing`, `copilotstudio`, `usermanagement`, `governance`, `environmentmanagement`): every `/users`
endpoint is `GET`, and the only threshold `PUT` in the whole licensing namespace is the per-resource one.
Microsoft's own guidance says the same — *"limits apply to agents rather than users."*

**2. Microsoft 365 Cost Management spending policies** — where **per-user monthly caps do exist**.
These are **hard limits**: the service stops for that user when the cap is reached.

| Aspect | Detail |
| --- | --- |
| Where | M365 admin center → **Copilot → Cost Management → Spending policies** |
| Scoping | **Entra group** or tenant. Scoping a policy to a single user is *not yet supported*. |
| Pattern | **Entra group → spending policy → per-user monthly limit** |
| Minimum | 2,000 credits/user/month (7,000 recommended) |
| Enforcement | Hard stop, reconciled periodically — a user can briefly exceed the cap; that overage isn't billed |
| API | **None public today** — admin-center UI only |

> The older PAYG billing-policy "budget" (Copilot Chat / SharePoint agents) only sends **alerts**.
> It does **not** stop consumption. Don't confuse the two.

So this script's user pass produces the **group-scoped inventory** you need to build those spending
policies, and `-ProbeUserThresholdApi` tests whether a user-threshold endpoint has appeared in your
tenant — at which point adding the write is a small change.

---

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7
- `Az.Accounts` module — `Install-Module Az.Accounts -Scope CurrentUser`
- Power Platform Administrator or Global Administrator
- Optional fallback for environment listing: `Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser`
- Optional, for `-UserGroup`: Entra permission to read groups (Graph `Group.Read.All` / `Directory.Read.All`)

---

## Quick start (recommended order)

```powershell
# 1. Read-only discovery - verifies auth and dumps the raw API shapes for one environment
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -Discover

# 2. Inventory current limits, no writes
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -ReportOnly

# 3. Dry run - shows every agent that would be changed
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -AgentCreditLimit 1000 -WhatIf

# 4. Apply to one environment
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -AgentCreditLimit 1000

# 5. Dry run across a whole environment group
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" -AgentCreditLimit 500 -WhatIf

# 6. Apply across the group, with exclusions and a hard stop
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" `
    -AgentCreditLimit 500 -StopAgentIfOverLimit `
    -ExcludeAgentIdCsv .\samples\agent-exclusions.csv -Force
```

Users pass:

```powershell
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" `
    -Scope Users -UserCreditLimit 100 -UserGroup "Copilot Makers" `
    -ExcludeUserIdCsv .\samples\user-exclusions.csv
```

Both in one run:

```powershell
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" `
    -Scope Both -AgentCreditLimit 500 -UserCreditLimit 100 -UserGroup "Copilot Makers"
```

---

## Parameters

| Parameter | Description |
| --- | --- |
| `-EnvironmentGroup` | Environment group ID or exact display name. All environments in the group are targeted. |
| `-Environment` | One or more environment IDs or exact display names. |
| `-Scope` | `Agents` (default), `Users`, or `Both`. |
| `-AgentCreditLimit` | Monthly Copilot Credit limit applied to each agent. |
| `-AgentNotificationThreshold` | Notify at this **percentage** (1-100) of the limit. Default 80. The API defines this field as a percentage, not a credit count. |
| `-NotifyIfOverCapacity` | Notify when the agent exceeds its limit. On by default (`-NotifyIfOverCapacity:$false` to disable). |
| `-StopAgentIfOverLimit` | Hard stop: turn the agent off at the limit. |
| `-UserCreditLimit` | Per-user limit recorded in the report (see "report-only" above). |
| `-UserGroup` | Entra group object ID or display name; scopes and labels the user report. |
| `-ExcludeAgentId` / `-ExcludeAgentIdCsv` | Agents to skip — inline IDs, or a CSV. |
| `-ExcludeUserId` / `-ExcludeUserIdCsv` | Users to skip — inline IDs/UPNs, or a CSV. |
| `-AgentSource` | `Licensing`, `Dataverse`, or `Both` (default). |
| `-LookbackDays` | Consumption snapshot window. Default 90. |
| `-ReportPath` | Folder for the CSV report. Default `.\reports`. |
| `-Discover` | Read-only raw API dump for the first target environment. |
| `-ReportOnly` | Inventory current limits without writing. |
| `-ProbeUserThresholdApi` | Test whether a user-threshold endpoint exists in your tenant. |
| `-TenantId` | Tenant to sign in to. |
| `-Force` | Skip confirmation and rewrite limits even when unchanged. |
| `-Diagnostics` | Extra detail about discovery and API fallbacks. |
| `-WhatIf` / `-Confirm` | Standard PowerShell safety switches. |

### Exclusion files

A header named `AgentId`, `ResourceId`, `BotId` or `Id` (users: `UserId`, `ObjectId`, `Upn`, `Email`, `Id`)
is used when present; otherwise every non-empty line of the file is treated as an ID. See `samples\`.
A single exclusion can be passed inline instead: `-ExcludeAgentId 1111...-1111`.

> **Safety:** if an exclusion file is supplied but yields **zero** exclusions — because it's empty or its
> ID column wasn't recognised — the script **stops with an error** instead of running. Passing the
> parameter means you intend to protect something, and silently continuing would apply the limit to
> *every* agent. Exclusion files are parsed before any network call, so mistakes fail immediately.

---

## Verified behaviour

Tested against a live tenant on 19 September 2026:

| Check | Result |
| --- | --- |
| Environment resolution (BAP admin API) | 19 environments listed |
| Agent discovery (licensing + Dataverse) | 18 agents, all named |
| Environment-scoped threshold lookup | thresholds belonging to other environments correctly ignored |
| `-WhatIf` | no writes issued |
| Write | 18/18 succeeded |
| **Read-back proof** | subsequent read returned the written limit for every agent |
| Exclusions | 17 targeted, 1 correctly skipped |
| Empty exclusion file | refused, as designed |

Two API details worth knowing, both confirmed against a live tenant:

- `notificationThreshold` is a **percentage (1-100)**, not a credit count.
- `licensing/entitlements/{id}/resourceThresholds` is **tenant-wide** — records must be matched on
  `environmentId` *and* `resourceId`. Matching on `resourceId` alone lets a threshold from one
  environment be mistaken for another environment's.

### Known API quirks

- `licensing/entitlements/MCSMessages/environments/{envId}/resources` can return **403** for a
  delegated admin token even when the same data renders in the admin center. Agents are still
  discovered from Dataverse and limits still write; only the `Consumed` column is blank.
- Date parameters are **camelCase** on the REST API (`fromDate` / `toDate`). The kebab-case spelling
  is the `pac` CLI flag name and is rejected with HTTP 400.

---

## How agents are discovered

| Source | Gives you | Caveat |
| --- | --- | --- |
| **Licensing** (`licensing/entitlements/MCSMessages/environments/{envId}/resources`) | Resource IDs that are guaranteed valid for the threshold API, plus month-to-date consumption. | Only agents that have consumed credits appear. |
| **Dataverse** (`bots` table) | Every agent in the environment, with display names. | Needs Dataverse access to each environment. |

`-AgentSource Both` (default) merges the two, so you get names *and* coverage of never-used agents.

---

## Report

Every run writes `reports\CopilotCreditLimits-yyyyMMdd-HHmmss.csv` and prints a summary. Columns:

`Timestamp, Scope, EnvironmentName, EnvironmentId, TargetId, TargetName, Source, Action,
PreviousLimit, NewLimit, NotificationThresholdPct, StopIfOverCapacity, Consumed, Message`

`Action` values: `Set`, `Skipped-Excluded`, `Skipped-NoChange`, `WhatIf`, `ReportOnly`,
`NoAgentsFound`, `Failed`.

---

## Scheduling

Because the script is idempotent (`Skipped-NoChange`) and non-interactive with `-Force`, it can run on a
schedule to re-assert limits as new agents appear:

```powershell
schtasks /create /tn "Copilot credit limits" /sc DAILY /st 06:00 /tr `
  "powershell -NoProfile -File \"C:\path\Set-CopilotCreditLimits.ps1\" -EnvironmentGroup \"Personal Productivity\" -AgentCreditLimit 500 -Force"
```

Unattended runs need a non-interactive auth path (service principal / managed identity) instead of the
interactive `Connect-AzAccount` used here.

---

## Notes and limits

- Entitlement ID for Copilot Credits is `MCSMessages`; API version `2024-10-01`.
- These licensing operations are **preview**. Run `-Discover` first if a tenant returns unexpected shapes —
  the script tries several documented path/parameter spellings and reports which one worked under `-Verbose`.
- Environment-group rules can lock capacity settings (`TenantPoolLockedByPolicy`). Agent-level thresholds are
  independent of the environment allocation, but the environment must still have capacity available.

## References

- [Manage Copilot Credits and capacity for Copilot Studio](https://learn.microsoft.com/power-platform/admin/manage-copilot-studio-copilot-credits-capacity)
- [Tutorial: Manage Copilot Credits allocations programmatically](https://learn.microsoft.com/power-platform/admin/programmability-tutorial-manage-copilot-credit-allocations)
- [`pac licensing` reference](https://learn.microsoft.com/power-platform/developer/cli/reference/licensing)
- [Managing AI experiences enabled by usage-based billing](https://learn.microsoft.com/microsoft-365/copilot/usage-based-billing-manage-copilot-credits)
- [Power Platform licensing OpenAPI spec](https://github.com/MicrosoftDocs/power-platform/blob/main/power-platform/developer/reference/licensing/licensing.json)

## Prior art

[`jameswh3/MW-Toolbox`](https://github.com/jameswh3/MW-Toolbox) includes
`Set-CopilotAgentConsumptionLimit.ps1`, which sets the threshold for a **single** agent. This repo
covers the bulk case: environment-group enumeration, multi-environment runs, agent discovery with
name enrichment, exclusions, dry runs and reporting.

## License

MIT
