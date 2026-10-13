using System.Text.Json;
using DocWeave.WebApi;

namespace DocWeave.Tests;

public sealed class WebApiContractTests
{
    [Fact]
    public void Export_request_contract_serializes_inline_template_and_datasources()
    {
        var payload = JsonSerializer.Deserialize<JsonElement>("""
        [
          { "Region": "London", "Amount": 125.5 },
          { "Region": "United States", "Amount": 300.0 }
        ]
        """);
        var request = new DocWeaveExportRequest(
            Format: "xlsx",
            TemplateId: null,
            TemplateJson: "{ \"schemaVersion\": \"1.0\" }",
            DataSources: new Dictionary<string, JsonElement> { ["sales"] = payload },
            Parameters: null,
            Delivery: new DocWeaveExportDeliveryRequest("download", "sales.xlsx"),
            Audit: new DocWeaveApiAuditMetadata(ClientId: "internal-service", RequestedBy: "tests", CorrelationId: "corr-1"));

        var json = JsonSerializer.Serialize(request);
        var roundTrip = JsonSerializer.Deserialize<DocWeaveExportRequest>(json);

        Assert.NotNull(roundTrip);
        Assert.Equal("xlsx", roundTrip!.Format);
        Assert.Equal("download", roundTrip.Delivery!.Kind);
        Assert.Equal("internal-service", roundTrip.Audit!.ClientId);
        Assert.True(roundTrip.DataSources.ContainsKey("sales"));
    }

    [Fact]
    public void Operation_status_contract_supports_async_downloads_and_warnings()
    {
        var response = new DocWeaveOperationStatusResponse(
            OperationId: "op-123",
            Status: "completed",
            PercentComplete: 100,
            Message: "Export completed",
            DownloadUrl: "/exports/op-123",
            Warnings: ["One optional column was not present."]);

        var json = JsonSerializer.Serialize(response);
        var roundTrip = JsonSerializer.Deserialize<DocWeaveOperationStatusResponse>(json);

        Assert.Equal("completed", roundTrip!.Status);
        Assert.Equal(100, roundTrip.PercentComplete);
        Assert.Equal("/exports/op-123", roundTrip.DownloadUrl);
        Assert.Single(roundTrip.Warnings!);
    }
}
