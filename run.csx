#r "Newtonsoft.Json"

using System;
using System.Net;
using Newtonsoft.Json;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;

// Main entry point for our Azure Function. Listens for webhooks from OneDrive and responds to the webhook with a 204 No Content.
public static async Task<object> Run(HttpRequestMessage req, CloudTable syncStateTable, TraceWriter log)
{
    log.Info($"Webhook was triggered!");

    Dictionary<string,string> qs = req.GetQueryNameValuePairs()
                            .ToDictionary(kv => kv.Key, kv=> kv.Value, StringComparer.OrdinalIgnoreCase);
    // Handle validation scenario for creating a new webhook subscription
    if (qs.ContainsKey("validationToken"))
    {
        var token = qs["validationToken"];
        log.Info($"Responding to validationToken: {token}");
        return PlainTextResponse(token);
    }                            

    // if not the validation scenario, read the body of the request and parse the notification
    string jsonContent = await req.Content.ReadAsStringAsync();
    log.Verbose($"Raw request content: {jsonContent}");
    
    dynamic data = JsonConvert.DeserializeObject(jsonContent);
    if (data.value != null)
    {
        foreach(var subscription in data.value)
        {
            var clientState = subscription.clientState;
            var resource = subscription.resource;
            string subscriptionId = (string)subscription.subscriptionId;
            log.Info($"Notification for subscription: '{subscriptionId}' Resource: '{resource}', clientState: '{clientState}'");
            await ProcessSubscriptionNotificationAsync(subscriptionId, syncStateTable, log);
        }
        return req.CreateResponse(HttpStatusCode.NoContent);
    }

    log.Info($"Request was incorrect. Returning bad request.");
    return req.CreateResponse(HttpStatusCode.BadRequest);
}

// Do the work to retrieve deltas from this subscription and then find any changed Excel files
private static async Task ProcessSubscriptionNotificationAsync(string subscriptionId, CloudTable table, TraceWriter log)
{
    // Retrieve our stored state from an Azure Table
    TableOperation operation = TableOperation.Retrieve<StoredState>("AAA", subscriptionId);
    TableResult result = table.Execute(operation);
    
    StoredState state = (StoredState)result.Result;
    if (state == null)
    {
        log.Info($"Missing data for subscription '{subscriptionId}'.");
        return;
    }

    log.Verbose($"Found subscription '{subscriptionId}' with lastDeltaUrl: '{state.LastDeltaToken}'.");

    state.LastNotificationDateTime = DateTime.UtcNow;

    try {
        await RenewAccessTokenAsync(state, log);
    } catch (Exception ex) {
        log.Info($"Unable to refresh access token: {ex.Message}");
    }


    // Make requests to Microsoft Graph to get changes
    List<string> changedExcelFileIds = await FindChangedExcelFilesInOneDrive(state, log);

    // Do work on the changed files
    foreach(var file in changedExcelFileIds)
    {
        log.Verbose($"Processing changes in file: {file}");
        try {
        await ScanExcelFileForPlaceholdersAsync(state, file, log);
        } catch (Exception ex)
        {
            log.Info($"Exception processing file: {ex.Message}");
        }
    }
    
    // Update our saved state for this subscription
    operation = TableOperation.Replace(state);
    table.Execute(operation);
}

// Use the Excel REST API to look for queries that we can replace with real data
private static async Task ScanExcelFileForPlaceholdersAsync(StoredState state, string fileId, TraceWriter log)
{
    string baseUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{fileId}/workbook";

    // Get the used range of the first sheet
    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {state.AccessToken}");
        client.DefaultRequestHeaders.Accept.Clear();

        var usedRangeUrl = baseUrl + "/worksheets/Sheet1/UsedRange?$select=address,cellCount,columnCount,values";
        log.Verbose($"UsedRangeUrl: {usedRangeUrl}");

        var response = await client.GetAsync(usedRangeUrl);
        response.EnsureSuccessStatusCode();

        dynamic data = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());

        var usedRangeId = data.address;
        log.Info($"Used range in file: {data.address}");

        // Find any placeholders in the used range with !roland, and record where they were and the value
        bool sendPatch = false;
        var range = data.values;
        for(int rowIndex = 0; rowIndex < range.Count; rowIndex++)
        {
            var rowValues = range[rowIndex];
            for(int columnIndex = 0; columnIndex < rowValues.Count; columnIndex++)
            {
                var value = (string)rowValues[columnIndex];
                if (value.StartsWith("!roland "))
                {
                    log.Verbose($"Found cell [{rowIndex},{columnIndex}] with value: {value} ");
                    rowValues[columnIndex] = await ReplacePlaceholderValue(value);
                    sendPatch = true;
                }
                else
                {
                    // Replace the value with null so we don't overwrite anything on the PATCH
                    rowValues[columnIndex] = null;
                }
            }
        }

        if (!sendPatch)
            return; // No placeholders, so nothing more to do.
        
        // Send the range back to the server as a patch request
        var patchUrl = baseUrl + $"/worksheets/Sheet1/range(address='{data.address}')";
        var request = new HttpRequestMessage(new HttpMethod("PATCH"), new Uri(patchUrl));
        var patchBody = JsonConvert.SerializeObject(new { values = range });
        log.Verbose($"Sending PATCH {patchUrl} with content: {patchBody}");
        request.Content = new StringContent(patchBody, System.Text.Encoding.UTF8, "application/json");
        var patchResponse = await client.SendAsync(request);
        patchResponse.EnsureSuccessStatusCode();
    }

}

// Make a request to retrieve a response based on the input value
private static async Task<string> ReplacePlaceholderValue(string inputValue)
{
    // This is merely an example. A real solution would do something much richer
    if (inputValue.StartsWith("!roland ") && inputValue.EndsWith(" stock quote"))
    {
        var tickerSymbol = inputValue.Substring(8, inputValue.Length -  20).Trim();
        var requestUrl = $"http://dev.markitondemand.com/MODApis/Api/v2/Quote/json?symbol={tickerSymbol}";
        using (var client = new HttpClient())
        {
            var response = await client.GetAsync(requestUrl);
            dynamic data = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
            return data.LastPrice;
        }
    }
    
    Random rndNum = new Random(int.Parse(Guid.NewGuid().ToString().Substring(0, 8), System.Globalization.NumberStyles.HexNumber));
    return rndNum.Next(20, 100).ToString(); 
}

// Request the delta stream from OneDrive to find files that have changed between notifications for this account
private static async Task<List<string>> FindChangedExcelFilesInOneDrive(StoredState state, TraceWriter log)
{
    List<string> changedFileIds = new List<string>();
    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {state.AccessToken}");
        client.DefaultRequestHeaders.Accept.Clear();

        string deltaUrl = "https://graph.microsoft.com/v1.0/me/drive/root/delta?token=latest";
        if (!String.IsNullOrEmpty(state.LastDeltaToken))
            deltaUrl = state.LastDeltaToken;

        while(true)
        {
            log.Verbose($"Making request for '{state.SubscriptionId}' to '{deltaUrl}' ");
            var response = await client.GetAsync(deltaUrl);
            response.EnsureSuccessStatusCode();
            var deltaResponse = JsonConvert.DeserializeObject<DeltaResponse>(await response.Content.ReadAsStringAsync());

            log.Verbose($"Found {deltaResponse.Value.Count()} files changed in this page.");
            log.Verbose("Changed files" + JsonConvert.SerializeObject(deltaResponse.Value));
            try {
            var changedExcelFiles = (from f in deltaResponse.Value
                                     where f.File != null && 
                                           f.Name != null && f.Name.EndsWith(".xlsx") &&
                                           f.Deleted == null
                                    select f.Id);
            log.Verbose($"Found {changedExcelFiles.Count()} changed Excel files in this page.");
            changedFileIds.AddRange(changedExcelFiles);
            } catch (Exception ex)
            {
                log.Info($"Exception enumerating changed files: {ex.ToString()}");
                throw;
            }

            if (!string.IsNullOrEmpty(deltaResponse.NextPageUrl))
            {
                // Get the next page of data from the service
                log.Verbose($"More pages of change data are available: nextPageUrl: {deltaResponse.NextPageUrl}");
                deltaUrl = deltaResponse.NextPageUrl;
            }
            else if (!string.IsNullOrEmpty(deltaResponse.NextDeltaUrl))
            {
                log.Verbose($"All changes requested, nextDeltaUrl: {deltaResponse.NextDeltaUrl}");
                state.LastDeltaToken = deltaResponse.NextDeltaUrl;
                return changedFileIds;
            }
        }
    }
}

private static HttpResponseMessage PlainTextResponse(string text)
{
    HttpResponseMessage response = new HttpResponseMessage()
    {
        StatusCode = HttpStatusCode.OK,
        Content = new StringContent(
                text,
                System.Text.Encoding.UTF8,
                "text/plain"
            )
    };
    return response;
}

// Retrieve a new access token from AAD
private static async Task RenewAccessTokenAsync(StoredState state, TraceWriter log) 
{
    log.Verbose("Refreshing access_token");
    var tokenServiceUri = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    var redirectUri = "https://onedriveapi.azurewebsites.net/o2c.html";
    var applicationId = "6f9f1420-26e0-498f-a3b9-25643198be8b";
    var applicationPassword = "tTh8hNEuKRTCt3k67T5ZqjS";

    var postBody = $"client_id={applicationId}&redirect_uri={redirectUri}&client_secret={applicationPassword}&grant_type=refresh_token&refresh_token={state.RefreshToken}";

    var request = new HttpRequestMessage(new HttpMethod("POST"), tokenServiceUri);
    request.Content = new StringContent(postBody, System.Text.Encoding.UTF8, "application/x-www-form-urlencoded");

    using (var client = new HttpClient())
    {
        var response = await client.SendAsync(request);
        response.EnsureSuccessStatusCode();

        dynamic data = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());

        state.AccessToken = data.access_token;
        state.RefreshToken = data.refresh_token;

        log.Verbose("New token acquired!");
    }
}

public class StoredState : TableEntity
{
    public StoredState() 
    {
        PartitionKey = "AAA";
    }
    public string AccessToken { get; set; }
    public string RefreshToken { get; set; }
    public string SubscriptionId { get { return RowKey;} set { RowKey = value;} }
    public string LastDeltaToken { get; set; }
    public DateTime LastNotificationDateTime { get; set; }
}

public class DeltaResponse 
{
    [JsonProperty("@odata.nextLink")]
    public string NextPageUrl {get;set;}

    [JsonProperty("@odata.deltaLink")]
    public string NextDeltaUrl {get;set;}

    [JsonProperty("value")]
    public DriveItem[] Value {get;set;}
    
}

public class DriveItem 
{
    [JsonProperty("id")]
    public string Id {get;set;}

    [JsonProperty("name")]
    public string Name {get;set;}

    [JsonProperty("size")]
    public int Size {get;set;}

    [JsonProperty("file")]
    public Dictionary<string, object> File {get;set;}

    [JsonProperty("deleted")]
    public Dictionary<string, object> Deleted {get;set;}
}