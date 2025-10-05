# Power BI Extract Metadata

A small PowerShell utility to extract dataset metadata (tables, columns, measures, partitions, data sources and Power Query M where available) from all Power BI workspaces accessible to a Service Principal.

This project uses the Analysis Services Tabular client (XMLA) to connect to datasets and the Power BI REST API to enumerate workspaces. It's suitable for auditing and documenting datasets in workspaces that expose the XMLA endpoint.

## Contents

- `PBI script with Service Principal Multiple workspace With Parameter.ps1` — main script (reads credentials from `config.txt` and writes a CSV with metadata to the project folder).
- `config.txt` — template for providing Service Principal credentials (a masked example is included in the repo; do NOT commit real secrets).
- `requirements.txt` — runtime and tooling notes required to run the script (DLLs, PowerShell version, optional tools).
- `run_output.txt` — example run log (not required).

## Quick start

1. Clone this repository:

	git clone <repo-url>

2. Open the project folder in PowerShell (Windows) or PowerShell Core and create/update `config.txt` in the project root. A template will be created automatically the first time you run the script.

3. Edit `config.txt` and fill in your Service Principal details:

	tenantId=<your-tenant-id>
	clientId=<your-app-client-id>
	clientSecret=<your-client-secret>

4. Run the script from the project folder (PowerShell):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\"PBI script with Service Principal Multiple workspace With Parameter.ps1"
```

5. Output: a timestamped CSV will be created in the project folder, e.g. `PowerBI_Metadata_Extract_20251005_123451.csv`.

## Files added to the repository

- `config.txt` (masked example) — contains placeholders and should be updated locally with real credentials if you run the script. Do not commit real secrets.
- `requirements.txt` — lists runtime DLLs and optional tools required to run the script.

## Requirements

- Windows / PowerShell or PowerShell Core
- The following assemblies must be available on the machine where the script runs:
  - `Microsoft.AnalysisServices.AdomdClient.dll` (e.g. from ADOMD.NET)
  - `Microsoft.AnalysisServices.Tabular.dll` (e.g. from Tabular Editor or Analysis Services SDK)
  These are referenced in the script at the paths below; update the script if your installation paths differ:

```
Add-Type -Path "C:\Program Files\Microsoft.NET\ADOMD.NET\160\Microsoft.AnalysisServices.AdomdClient.dll"
Add-Type -Path "C:\Program Files (x86)\Tabular Editor\Microsoft.AnalysisServices.Tabular.dll"
```

## Azure / Power BI prerequisites

1. Register a Service Principal (App) in Azure AD. Create a client secret and note the tenantId, clientId and clientSecret.
2. Tenant admin: In the Power BI Admin portal, enable service principal access where required and ensure the app has the necessary API access.
3. Add the Service Principal as a Workspace Member (Contributor/Viewer) for any workspaces the script should read, or grant tenant-level access according to your governance.
4. Ensure the workspaces/datasets you want to extract support XMLA endpoint (Premium / Premium Per User or capacity with XMLA enabled). If the capacity is paused or XMLA is disabled you'll get `CapacityNotActive` errors and the script cannot connect via XMLA.

## What the script extracts

- Columns: name, data type, hidden flag, description (if present)
- Measures: name, expression (DAX), data type, hidden flag, description
- Partitions: partition name, M code (Power Query) when available, resolved datasource and database when parseable
- DataSources: stand-alone data sources in the model (connection strings, database/schema)

Note: If XMLA access is unavailable the script cannot retrieve model-level details (measures, partitions, M), but it can still enumerate workspaces via REST. Consider enabling the XMLA endpoint for full results.

## Security

- Never commit `config.txt` with real secrets to source control. This repository includes a `.gitignore` which excludes `config.txt` and CSV outputs.
- If a secret is accidentally exposed, rotate it immediately in Azure AD.
- For production use, consider using Azure Key Vault or environment-based secret injection instead of plain `config.txt`.

## Contribution

Contributions welcome. A suggested workflow:

1. Fork the repository.
2. Create a feature branch.
3. Add tests or documentation where appropriate.
4. Open a pull request with a clear description of changes.

Guidelines:

- Keep secrets out of commits.
- Keep changes small and focused.

## Troubleshooting

- `CapacityNotActive` — check Power BI capacity status and XMLA endpoint enablement.
- `The specified Power BI workspace ... is not found` — usually indicates the XMLA URI used didn't match what the service expects; the script defaults to using workspace name URIs. If you run into name/ID issues, open an issue with logs.
- Missing assemblies — verify the DLL paths in the script or install ADOMD.NET and Tabular client libraries.

## Future improvements

- Add a REST-based fallback to extract partial metadata when XMLA is not available (already planned).
- Support output formats other than CSV (JSON, Excel).
- Use MSAL (modern auth) instead of the v1 token endpoint for more robust authentication.

## License

Specify a license (e.g., MIT) if you want this repository public and allow contributions. Add a `LICENSE` file to the root.

---

If you'd like, I can also add the REST fallback and a CONTRIBUTING.md with a PR template. Let me know which features to add next.