# AgentPjM Automation — Step-by-Step Implementation Guide

> **Last updated**: April 2, 2026
> **Auth model**: Managed Identity only (zero secrets)

## Architecture Overview

```
┌──────────────────────┐
│   SharePoint Online   │  3 Excel workbooks updated manually by CPMs
│   (Document Library)  │
└──────────┬───────────┘
           │  Graph API (Managed Identity + Sites.Read.All)
           ▼
┌──────────────────────┐
│  Azure Function       │  Timer: every Monday 6 AM UTC
│  (Flex Consumption)   │  Python 3.11, zero secrets
│                       │
│  1. Download .xlsm    │
│  2. Generate reports  │  ~84 .txt files (reuses existing script)
│  3. Sync to Foundry   │
└──────────┬───────────┘
           │  azure-ai-projects SDK (Managed Identity)
           ▼
┌──────────────────────┐
│  Foundry Vector Store │  Agent sees fresh data
│  (agentpjm-foundry-01)│
└──────────────────────┘
```

**Auth**: A single system-assigned Managed Identity handles both SharePoint
(Graph API) and Foundry access. No App Registration secrets needed.

---

## Your Environment

| Resource | Value |
|---|---|
| Azure Tenant | `72f988bf-86f1-41af-91ab-2d7cd011db47` / `microsoft.onmicrosoft.com` |
| Subscription | `d5144933-5c29-4b8c-98a5-8f625fce9d58` |
| Resource Group | `rg-gcidinnvlab-agentPjM` |
| SharePoint Site | `microsoftapc.sharepoint.com/teams/DigitalEmployee_PjM` |
| Foundry Resource | `agentpjm-foundry-01` (swedencentral) |
| Foundry Project | `agent-pjm-project` |
| Function App (to create) | `func-agentpjm-ingest` |
| Storage Account (to create) | `stagentpjmfunc` |

---

## PHASE 1 — Get the Vector Store ID

### Step 1: Record the Vector Store ID

1. Go to https://ai.azure.com → your project (`agent-pjm-project`)
2. **Agents** → select your agent → **Tools** → **File Search**
3. Click on the linked vector store
4. Copy the **Vector Store ID** (looks like `vs_xxxxxxxxxxxxxxxxxxxx`)
5. Save it — you'll need it in Step 6

---

## PHASE 2 — Create Azure Resources

Open **Azure Cloud Shell** (https://shell.azure.com, PowerShell) or a local terminal with Azure CLI.

### Step 2: Create Storage Account

```powershell
$RG       = "rg-gcidinnvlab-agentPjM"
$LOCATION = "swedencentral"
$STORAGE  = "stagentpjmfunc"

az storage account create `
  --name $STORAGE `
  --resource-group $RG `
  --location $LOCATION `
  --sku Standard_LRS `
  --kind StorageV2
```

### Step 3: Create the Function App

```powershell
$FUNCAPP = "func-agentpjm-ingest"

# Option A: Flex Consumption (preferred — future-proof, per-function scaling)
az functionapp create `
  --name $FUNCAPP `
  --resource-group $RG `
  --storage-account $STORAGE `
  --flexconsumption-location $LOCATION `
  --runtime python `
  --runtime-version 3.11 `
  --functions-version 4
```

> If Flex Consumption is not available in swedencentral, use regular Consumption:
> ```powershell
> # Option B: Regular Consumption (fallback)
> az functionapp create `
>   --name $FUNCAPP `
>   --resource-group $RG `
>   --storage-account $STORAGE `
>   --consumption-plan-location $LOCATION `
>   --runtime python `
>   --runtime-version 3.11 `
>   --functions-version 4 `
>   --os-type Linux
> ```

### Step 4: Enable System-Assigned Managed Identity

```powershell
az functionapp identity assign --name $FUNCAPP --resource-group $RG
```

**Copy the `principalId`** from the output — you need it for Steps 5a and 5b.

Example output:
```json
{
  "principalId": "abcd1234-5678-....",
  "tenantId": "72f988bf-86f1-41af-91ab-2d7cd011db47",
  "type": "SystemAssigned"
}
```

---

## PHASE 3 — Grant Permissions to Managed Identity

The Managed Identity needs 3 permissions on 2 services.

### Step 5a: Grant Graph API `Sites.Read.All` (for SharePoint access)

This requires **Global Admin** or **Privileged Role Admin** privileges.
If you don't have admin rights, send this script to your tenant admin.

```powershell
# Install the Microsoft Graph PowerShell module (one-time)
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Connect with admin credentials
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All"

# Set your Managed Identity's principal ID from Step 4
$ManagedIdentityId = "<principalId-from-step-4>"

# Get Microsoft Graph's service principal (same in every tenant)
$GraphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

# Find the Sites.Read.All application role
$SitesReadAll = $GraphSP.AppRoles | Where-Object { $_.Value -eq "Sites.Read.All" }

# Grant it to the Managed Identity
New-MgServicePrincipalAppRoleAssignment `
  -ServicePrincipalId $ManagedIdentityId `
  -PrincipalId $ManagedIdentityId `
  -ResourceId $GraphSP.Id `
  -AppRoleId $SitesReadAll.Id
```

Expected output: a JSON object with `appRoleId` and `principalId` confirming the assignment.

> **What this does**: Gives your Function App read-only access to SharePoint sites
> via Microsoft Graph. Equivalent to `Sites.Read.All` application permission on an
> App Registration, but without any secrets.

### Step 5b: Grant Azure RBAC Roles (for Foundry access)

```powershell
$PRINCIPAL_ID = "<principalId-from-step-4>"

# ── Role 1: Azure AI Developer on the Foundry resource ──
# Allows managing vector stores, uploading files, invoking agents
az role assignment create `
  --assignee $PRINCIPAL_ID `
  --role "Azure AI Developer" `
  --scope "/subscriptions/d5144933-5c29-4b8c-98a5-8f625fce9d58/resourceGroups/rg-gcidinnvlab-agentPjM/providers/Microsoft.CognitiveServices/accounts/agentpjm-foundry-01"

# ── Role 2: Storage Blob Data Contributor on Foundry's storage account ──
# Allows writing file blobs that the vector store indexes
# Find the storage account: Foundry Portal → Project Settings → Connected resources
az role assignment create `
  --assignee $PRINCIPAL_ID `
  --role "Storage Blob Data Contributor" `
  --scope "<foundry-connected-storage-account-resource-id>"
```

> **How to find the Foundry storage resource ID**:
> 1. Foundry Portal → your project → **Settings** → **Connected resources**
> 2. Find the Storage Account listed there
> 3. Go to Azure Portal → that Storage Account → **Properties** → **Resource ID**
> 4. It looks like: `/subscriptions/.../resourceGroups/.../providers/Microsoft.Storage/storageAccounts/...`

### Step 5c: Verify All Permissions (optional but recommended)

```powershell
# Check RBAC roles
az role assignment list --assignee $PRINCIPAL_ID --output table

# Check Graph API app roles
Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityId | Format-Table AppRoleId, ResourceDisplayName
```

---

## PHASE 4 — Configure App Settings

### Step 6: Set App Settings (6 settings, zero secrets)

```powershell
az functionapp config appsettings set `
  --name func-agentpjm-ingest `
  --resource-group rg-gcidinnvlab-agentPjM `
  --settings `
    "TIMER_SCHEDULE=0 0 6 * * 1" `
    "SP_SITE_HOST=microsoftapc.sharepoint.com" `
    "SP_SITE_PATH=/teams/DigitalEmployee_PjM" `
    "SP_FOLDER_PATH=/Shared Documents/Delivery Quality Enablement Agent/TQP DataSet" `
    "FOUNDRY_ENDPOINT=https://agentpjm-foundry-01.services.ai.azure.com/api/projects/agent-pjm-project" `
    "VECTOR_STORE_ID=<from-step-1>"
```

| Setting | Purpose | Example |
|---|---|---|
| `TIMER_SCHEDULE` | When the function runs (NCRONTAB) | `0 0 6 * * 1` = Monday 6AM UTC |
| `SP_SITE_HOST` | SharePoint hostname | `microsoftapc.sharepoint.com` |
| `SP_SITE_PATH` | Team site path | `/teams/DigitalEmployee_PjM` |
| `SP_FOLDER_PATH` | Folder containing Excel files | `/Shared Documents/Delivery Quality...` |
| `FOUNDRY_ENDPOINT` | Foundry project endpoint | `https://agentpjm-foundry-01...` |
| `VECTOR_STORE_ID` | Which vector store to update | `vs_xxxxxxxxxxxxxxxxxxxx` |

> **Note**: `SP_FOLDER_PATH` should point to whichever folder the CPMs upload
> their latest Excel files to. If the folder name changes each month (e.g.,
> "TQP DataSet (12th March 2026)"), update this setting accordingly — or ask
> CPMs to always use a fixed folder name.

---

## PHASE 5 — Deploy the Function Code

### Step 7: Prepare the Deployment Package

```powershell
cd "c:\Users\srahlawat\OneDrive - Microsoft\agentpjm (2)\agentpjm"

# Copy the existing report script into the function folder
Copy-Item .\generate_sprint_reports.py .\azure_function\generate_sprint_reports.py
```

Verify the folder structure:
```
azure_function/
├── function_app.py            # Timer + HTTP triggers (entry point)
├── sharepoint_client.py       # Downloads Excel from SharePoint via Graph API
├── foundry_client.py          # Uploads .txt to Foundry vector store
├── report_generator.py        # Calls generate_sprint_reports.py
├── generate_sprint_reports.py # Your existing script (copied in)
├── requirements.txt           # azure-functions, azure-identity, azure-ai-projects, requests, openpyxl
├── host.json                  # Functions host config
├── local.settings.json        # Local dev only (NOT deployed)
└── .gitignore
```

### Step 8: Install Azure Functions Core Tools

```powershell
# If you don't have it already
npm install -g azure-functions-core-tools@4 --unsafe-perm true

# Verify
func --version
```

### Step 9: Deploy

```powershell
cd azure_function
func azure functionapp publish func-agentpjm-ingest --python
```

This packages your code, installs dependencies from `requirements.txt` in the
cloud, and deploys. Takes 2-3 minutes.

---

## PHASE 6 — Test & Validate

### Step 10: Trigger Manually via HTTP

```powershell
# Get the function URL
$URL = az functionapp function show `
  --name func-agentpjm-ingest `
  --resource-group rg-gcidinnvlab-agentPjM `
  --function-name SyncSprintDataHttp `
  --query "invokeUrlTemplate" -o tsv

# Get the function key
$KEY = az functionapp function keys list `
  --name func-agentpjm-ingest `
  --resource-group rg-gcidinnvlab-agentPjM `
  --function-name SyncSprintDataHttp `
  --query "default" -o tsv

# Trigger the pipeline
Invoke-RestMethod -Uri "$URL`?code=$KEY" -Method POST
```

Expected response: `Pipeline completed successfully.`

### Step 11: Monitor Logs

**Option A**: Live stream
```powershell
func azure functionapp logstream func-agentpjm-ingest
```

**Option B**: Azure Portal → Function App → **Monitor** → **Invocations**

You should see log entries like:
```
Step 1/3: Downloading Excel files from SharePoint...
  Downloaded: FY26 Americas Sprint Checkpoint Tracker v2.0.xlsm
  Downloaded: FY26 EMEA Sprint Checkpoint Tracker v2.0.xlsm
  Downloaded: FY26 Asia Sprint Checkpoint Tracker v2.0.xlsm
Step 2/3: Generating text reports...
  Generated 84 files.
Step 3/3: Syncing to Foundry vector store...
  Vector store sync complete.
Pipeline completed successfully.
```

### Step 12: Verify Agent Answers

Go to Teams or Foundry playground and ask your agent:

1. **"How many projects are there?"** — should match your current project count
2. **"Tell me about Medline"** — a new project; confirms new data was ingested
3. **"Show me the EMEA heatmap"** — confirms regional indexes are working
4. **"What is Lumen's health status?"** — another new project

If the agent answers correctly, the pipeline is working end-to-end.

---

## PHASE 7 — Ongoing Operations

### Timer Schedule

The function runs automatically every Monday at 6:00 AM UTC (11:30 AM IST).

To change the schedule, update the `TIMER_SCHEDULE` app setting:

| Schedule | NCRONTAB Value |
|----------|---------------|
| Every Monday 6 AM UTC | `0 0 6 * * 1` (default) |
| Every weekday 7 AM UTC | `0 0 7 * * 1-5` |
| Every 6 hours | `0 0 */6 * * *` |
| 1st of every month | `0 0 8 1 * *` |

```powershell
az functionapp config appsettings set `
  --name func-agentpjm-ingest `
  --resource-group rg-gcidinnvlab-agentPjM `
  --settings "TIMER_SCHEDULE=0 0 7 * * 1-5"
```

### When CPMs Update the Folder Path

If the SharePoint folder name changes (e.g., new dataset folder each quarter):

```powershell
az functionapp config appsettings set `
  --name func-agentpjm-ingest `
  --resource-group rg-gcidinnvlab-agentPjM `
  --settings "SP_FOLDER_PATH=/Shared Documents/Delivery Quality Enablement Agent/TQP DataSet (New Folder Name)"
```

### When You Update generate_sprint_reports.py

1. Copy the updated script: `Copy-Item .\generate_sprint_reports.py .\azure_function\generate_sprint_reports.py`
2. Re-deploy: `cd azure_function; func azure functionapp publish func-agentpjm-ingest --python`

---

## Completion Checklist

| # | Task | Status |
|---|------|--------|
| 1 | Vector Store ID recorded | ⬜ |
| 2 | Storage account created (`stagentpjmfunc`) | ⬜ |
| 3 | Function App created (Flex Consumption) | ⬜ |
| 4 | Managed Identity enabled | ⬜ |
| 5a | `Sites.Read.All` granted to MI (admin PowerShell) | ⬜ |
| 5b | `Azure AI Developer` role assigned to MI | ⬜ |
| 5c | `Storage Blob Data Contributor` role assigned to MI | ⬜ |
| 6 | App Settings configured (6 settings) | ⬜ |
| 7 | `generate_sprint_reports.py` copied to `azure_function/` | ⬜ |
| 8 | Azure Functions Core Tools installed | ⬜ |
| 9 | Code deployed via `func azure functionapp publish` | ⬜ |
| 10 | HTTP trigger test passed | ⬜ |
| 11 | Logs show all 3 steps completing | ⬜ |
| 12 | Agent returns correct answers from new data | ⬜ |

---

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `DefaultAzureCredential failed` | MI not enabled or token endpoint unreachable | Verify Step 4 completed; check `az functionapp identity show` |
| `403 Forbidden` on Graph API | `Sites.Read.All` not granted to MI | Run Step 5a PowerShell with admin |
| `401 Unauthorized` on Foundry | MI missing `Azure AI Developer` role | Run Step 5b |
| `404 Not Found` on SharePoint | Wrong `SP_SITE_HOST`, `SP_SITE_PATH`, or `SP_FOLDER_PATH` | Verify app settings match the actual SharePoint URL |
| `Vector store not found` | Wrong `VECTOR_STORE_ID` | Copy from Foundry portal |
| Cold start timeout (>30s) | Too many deps or large package | Minimize requirements.txt; use Flex Consumption always-ready instances |
| `File upload failed` | Wrong encoding or Foundry storage role missing | Ensure UTF-8 files; check Step 5c |
| `No module named generate_sprint_reports` | Forgot to copy the script | Run Step 7 and redeploy |

---

## Security Notes

1. **Zero secrets in app settings** — Managed Identity handles all auth
2. **Read-only SharePoint access** — `Sites.Read.All` cannot modify any files
3. **Function key required** for HTTP trigger — prevents unauthorized invocations
4. **`.gitignore` excludes `local.settings.json`** — prevents accidental secret commits
5. **Managed Identity is lifecycle-bound** — deleting the Function App automatically revokes all access
