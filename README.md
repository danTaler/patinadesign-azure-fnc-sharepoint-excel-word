# patinadesign-azure-fnc-sharepoint-excel-word

# Proposal generator (Azure Functions, Python v2)

On-demand HTTP function that reads values from a SharePoint-hosted Excel workbook via
Microsoft Graph and fills a SharePoint-hosted Word MASTER template, writing the result
as a **new** `.docx` next to the MASTER. The MASTER is never modified.

Sources (site `patinaau.sharepoint.com/sites/AdminConfidential`):

- Excel: `Patina_Design_Studio_Scope of Work Model 20260707.xlsx` (cell `B5` = client name)
- Word MASTER: `Patina_Design_Studio_Full_Service_Design_Proposal_Template_MASTER_20260618.docx`

## How templating works

The Word MASTER contains literal placeholders such as `[[client_name]]`. On each run the
function downloads the MASTER, renders it with `docxtpl` (Jinja delimiters `[[` / `]]`), and
uploads the result as `<ClientName>_Proposal.docx`. Placeholders are consumed only in the
rendered copy, so the MASTER keeps them and can be reused indefinitely.

To add a field: add an entry to `MAPPING` in `function_app.py`
(e.g. `"project_address": {"sheet": SHEET_NAME, "cell": "B7"}`) and put `[[project_address]]`
in the Word MASTER. Keys must not contain spaces.

## Develop in GitHub Codespaces

1. Open the repo on GitHub → **Code → Codespaces → Create codespace**.
2. The dev container (`.devcontainer/devcontainer.json`) preinstalls Python 3.11, Azure CLI and
   Azure Functions Core Tools and runs `pip install -r requirements.txt`.

## Run locally

```bash
pip install -r requirements.txt
# fill in local.settings.json (SHAREPOINT_SITE_ID, SHEET_NAME, EXCEL_ITEM_ID, WORD_ITEM_ID)
az login          # DefaultAzureCredential uses your CLI login locally
func start
curl -X POST http://localhost:7071/api/generate
```

Locally, `DefaultAzureCredential` falls back to your `az login` identity, which needs
read/write access to the SharePoint site. In Azure it uses the Managed Identity.

## Deploy

```bash
az login
func azure functionapp publish <app-name>
```

## One-time Azure wiring

```bash
RG=<resource-group>; LOC=<region>; APP=<app-name>; STORAGE=<storageaccount>

az group create -n $RG -l $LOC
az storage account create -n $STORAGE -g $RG -l $LOC --sku Standard_LRS

# Flex Consumption (recommended) ...
az functionapp create -g $RG -n $APP --storage-account $STORAGE --flexconsumption-location $LOC \
  --runtime python --runtime-version 3.11
# ... or classic Linux Consumption
# az functionapp create -g $RG -n $APP --storage-account $STORAGE --consumption-plan-location $LOC \
#   --os-type Linux --runtime python --runtime-version 3.11 --functions-version 4

# System-assigned Managed Identity
az functionapp identity assign -g $RG -n $APP
PRINCIPAL_ID=$(az functionapp identity show -g $RG -n $APP --query principalId -o tsv)

# App settings
az functionapp config appsettings set -g $RG -n $APP --settings \
  SHAREPOINT_SITE_ID=<site-id> SHEET_NAME=<sheet> EXCEL_ITEM_ID=<excel-item-id> WORD_ITEM_ID=<word-item-id>
```

Grant the Managed Identity the Graph application role `Sites.ReadWrite.All` (requires a
Global/Privileged Role admin; app-role assignments to a managed identity are implicitly consented):

```powershell
Connect-MgGraph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All"
$graph = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$role  = $graph.AppRoles | Where-Object { $_.Value -eq "Sites.ReadWrite.All" -and $_.AllowedMemberTypes -contains "Application" }
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId <PRINCIPAL_ID> `
  -PrincipalId <PRINCIPAL_ID> -ResourceId $graph.Id -AppRoleId $role.Id
```

(For least privilege, use `Sites.Selected` instead and grant the identity `write` on the
single site via `POST /sites/{site-id}/permissions`.)

## Finding the IDs

Use [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) or `az rest`.

Site ID:

```
GET https://graph.microsoft.com/v1.0/sites/patinaau.sharepoint.com:/sites/AdminConfidential
```

→ `id` (the full `hostname,siteCollectionId,webId` string) is `SHAREPOINT_SITE_ID`.

Item IDs (files in the default document library):

```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root/children
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/<Folder>:/children
```

→ `id` of the `.xlsx` is `EXCEL_ITEM_ID`, `id` of the MASTER `.docx` is `WORD_ITEM_ID`.
Alternatively search: `GET .../sites/{site-id}/drive/root/search(q='MASTER_20260618')`.

## Invoke

```bash
curl -X POST "https://<app-name>.azurewebsites.net/api/generate?code=<function-key>" \
  -H "Content-Type: application/json" \
  -d '{}'
```

Body values `SHAREPOINT_SITE_ID`, `EXCEL_ITEM_ID`, `WORD_ITEM_ID` override the app settings
if supplied. Get the key with `az functionapp function keys list -g $RG -n $APP --function-name generate`.

Response:

```json
{"status": "ok", "url": "https://patinaau.sharepoint.com/.../Acme_Pty_Ltd_Proposal.docx", "fields": {"client_name": "Acme Pty Ltd"}}
```

Errors return HTTP 500 with `{"status": "error", "error": "..."}`.
