using System.Text.Json;

namespace DocWeave.WebApi;

/// <summary>
/// Request to export data to a file. This is the wire shape a Web API host accepts: it is plain data, so it serializes to and from JSON.
/// Supply either a stored template by id or an inline template, not both.
/// </summary>
/// <param name="Format">The output format, "xlsx" or "csv".</param>
/// <param name="TemplateId">The id of a stored, approved template. Null when an inline template is sent.</param>
/// <param name="TemplateJson">An inline template document (see template-api-contract-v1.md). Null when a stored template id is used.</param>
/// <param name="DataSources">The data the template refers to, keyed by data source name. Each value is the JSON array or object for that source.</param>
/// <param name="Parameters">Values for {{parameters.name}} placeholders in the template.</param>
/// <param name="Delivery">How the finished file should be returned. Null means the caller decides.</param>
/// <param name="Audit">Who is asking and why, recorded with the operation.</param>
public sealed record DocWeaveExportRequest(
    string Format,
    string? TemplateId,
    string? TemplateJson,
    IReadOnlyDictionary<string, JsonElement> DataSources,
    IReadOnlyDictionary<string, JsonElement>? Parameters = null,
    DocWeaveExportDeliveryRequest? Delivery = null,
    DocWeaveApiAuditMetadata? Audit = null);

/// <summary>
/// Request to read a workbook or CSV file into rows.
/// </summary>
/// <param name="Format">The source format, "xlsx" or "csv".</param>
/// <param name="Source">Where the file comes from.</param>
/// <param name="TemplateId">The id of a stored import template that maps columns. Null when an inline template is sent or none is used.</param>
/// <param name="TemplateJson">An inline import template. Null when a stored template id is used or none is needed.</param>
/// <param name="Audit">Who is asking and why, recorded with the operation.</param>
public sealed record DocWeaveImportRequest(
    string Format,
    DocWeaveImportSource Source,
    string? TemplateId = null,
    string? TemplateJson = null,
    DocWeaveApiAuditMetadata? Audit = null);

/// <summary>
/// Where an import reads its file from. Only the fields that match <paramref name="Kind"/> are used.
/// </summary>
/// <param name="Kind">The kind of source, for example a server path, a URI, or uploaded content.</param>
/// <param name="FileName">The original file name, used for labels and error messages.</param>
/// <param name="Path">A file path, for a path source.</param>
/// <param name="Uri">A location to fetch the file from, for a URI source. The host decides whether fetching is allowed.</param>
/// <param name="Content">The file bytes, for an upload.</param>
/// <param name="Password">The password to open an encrypted workbook. It is never logged.</param>
public sealed record DocWeaveImportSource(
    string Kind,
    string? FileName = null,
    string? Path = null,
    string? Uri = null,
    byte[]? Content = null,
    string? Password = null);

/// <summary>
/// How a finished export is handed back to the caller.
/// </summary>
/// <param name="Kind">The delivery kind, for example "download" to return the file in the response.</param>
/// <param name="FileName">The name to give the file.</param>
/// <param name="Target">Where to send the file for kinds that need a destination, such as a storage location.</param>
public sealed record DocWeaveExportDeliveryRequest(
    string Kind,
    string? FileName = null,
    string? Target = null);

/// <summary>
/// Who made a request and why. It is recorded with the operation so an export or import can be traced.
/// </summary>
/// <param name="ClientId">The calling system or service.</param>
/// <param name="RequestedBy">The user or account the request is made for.</param>
/// <param name="CorrelationId">An id that ties this operation to the caller's own logs.</param>
/// <param name="Tags">Extra key/value labels to record with the operation.</param>
public sealed record DocWeaveApiAuditMetadata(
    string? ClientId = null,
    string? RequestedBy = null,
    string? CorrelationId = null,
    IReadOnlyDictionary<string, string>? Tags = null);

/// <summary>
/// Response when an export or import runs in the background and has been accepted but not finished.
/// </summary>
/// <param name="OperationId">The id to poll for progress.</param>
/// <param name="StatusUrl">Where to poll for the operation's status.</param>
/// <param name="DownloadUrl">Where the result will be available once complete, when it is already known.</param>
public sealed record DocWeaveOperationAcceptedResponse(
    string OperationId,
    string StatusUrl,
    string? DownloadUrl = null);

/// <summary>
/// The current state of a background export or import.
/// </summary>
/// <param name="OperationId">The id of the operation.</param>
/// <param name="Status">The state, for example "running" or "completed".</param>
/// <param name="PercentComplete">Progress from 0 to 100, or null when it cannot be worked out.</param>
/// <param name="Message">A short human-readable description of the current step or outcome.</param>
/// <param name="DownloadUrl">Where to download the result once the operation has completed.</param>
/// <param name="Warnings">Non-fatal problems found along the way, such as an optional column that was missing.</param>
public sealed record DocWeaveOperationStatusResponse(
    string OperationId,
    string Status,
    int? PercentComplete = null,
    string? Message = null,
    string? DownloadUrl = null,
    IReadOnlyList<string>? Warnings = null);
