# DocWeave Examples

This project is a runnable how-to for the main DocWeave APIs.

Run it from the repository root:

```powershell
dotnet run --project Examples\DocWeave.Examples\DocWeave.Examples.csproj -- Examples\Output
```

The command creates one folder per scenario. Template scenarios include:

- `template.json`, the DocWeave template.
- `api-request.json`, the equivalent Web API payload.
- sample source files or generated export files.
- `import-result.json` for import scenarios.

The examples are grouped in `Program.cs`:

- 5 XLSX template scenarios.
- 5 CSV template scenarios.
- direct XLSX and CSV fluent/read scenarios for comparison.

