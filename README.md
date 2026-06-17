# UPS

C# — Bulk User Profile Properties via Microsoft Graph API
1. NuGet Packages
xml<PackageReference Include="Microsoft.Graph" Version="5.*" />
<PackageReference Include="Azure.Identity" Version="1.*" />
<PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="8.*" />

2. Model
csharppublic class UserProfileModel
{
    public string Id { get; set; }
    public string DisplayName { get; set; }
    public string Mail { get; set; }
    public string UserPrincipalName { get; set; }
    public string JobTitle { get; set; }
    public string Department { get; set; }
    public string OfficeLocation { get; set; }
    public string City { get; set; }
    public string Country { get; set; }
    public string MobilePhone { get; set; }
    public string EmployeeId { get; set; }
    public string ManagerDisplayName { get; set; }
    public string ManagerMail { get; set; }
}

3. Graph Service Client Setup (Certificate Auth — Pattern)
csharp

using Azure.Identity;
using Microsoft.Graph;

public static class GraphClientFactory
{
    public static GraphServiceClient Create(string tenantId, string clientId, string certThumbprint)
    {
        var credential = new ClientCertificateCredential(
            tenantId,
            clientId,
            certThumbprint,
            new ClientCertificateCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            }
        );

        return new GraphServiceClient(credential, new[]
        {
            "https://graph.microsoft.com/.default"
        });
    }

    // Alternate: PFX from Key Vault (your existing pattern)
    public static GraphServiceClient CreateFromCertificate(
        string tenantId, string clientId, X509Certificate2 certificate)
    {
        var credential = new ClientCertificateCredential(tenantId, clientId, certificate);
        return new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
    }
}

4. Bulk Fetch — All Users with Pagination
csharp

using Microsoft.Graph;
using Microsoft.Graph.Models;

public class UserProfileService
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<UserProfileService> _logger;

    public UserProfileService(GraphServiceClient graphClient, ILogger<UserProfileService> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    public async Task<List<UserProfileModel>> GetAllUserProfilesAsync(
        CancellationToken cancellationToken = default)
    {
        var profiles = new List<UserProfileModel>();

        try
        {
            // PageIterator handles @odata.nextLink automatically
            var usersPage = await _graphClient.Users
                .GetAsync(config =>
                {
                    config.QueryParameters.Select = new[]
                    {
                        "id", "displayName", "mail", "userPrincipalName",
                        "jobTitle", "department", "officeLocation",
                        "city", "country", "mobilePhone", "employeeId"
                    };
                    config.QueryParameters.Top = 999;
                    config.QueryParameters.Filter = "accountEnabled eq true";
                    config.QueryParameters.Orderby = new[] { "displayName" };
                }, cancellationToken);

            var pageIterator = PageIterator<User, UserCollectionResponse>
                .CreatePageIterator(
                    _graphClient,
                    usersPage,
                    user =>
                    {
                        profiles.Add(MapToModel(user));
                        return true; // return false to stop iteration early
                    });

            await pageIterator.IterateAsync(cancellationToken);

            _logger.LogInformation("Fetched {Count} user profiles", profiles.Count);
        }
        catch (ODataError ex)
        {
            _logger.LogError(ex, "Graph API error: {Code} - {Message}",
                ex.Error?.Code, ex.Error?.Message);
            throw;
        }

        return profiles;
    }
}

5. Fetch with Manager (Expand)
csharp

public async Task<List<UserProfileModel>> GetAllUsersWithManagerAsync(
    CancellationToken cancellationToken = default)
{
    var profiles = new List<UserProfileModel>();

    var usersPage = await _graphClient.Users
        .GetAsync(config =>
        {
            config.QueryParameters.Select = new[]
            {
                "id", "displayName", "mail", "jobTitle", "department", "officeLocation"
            };
            config.QueryParameters.Expand = new[] { "manager($select=displayName,mail)" };
            config.QueryParameters.Top = 500; // lower limit when using $expand
            config.QueryParameters.Filter = "accountEnabled eq true";
        }, cancellationToken);

    var pageIterator = PageIterator<User, UserCollectionResponse>
        .CreatePageIterator(
            _graphClient,
            usersPage,
            user =>
            {
                var model = MapToModel(user);

                if (user.Manager is User manager)
                {
                    model.ManagerDisplayName = manager.DisplayName;
                    model.ManagerMail = manager.Mail;
                }

                profiles.Add(model);
                return true;
            });

    await pageIterator.IterateAsync(cancellationToken);
    return profiles;
}

7. Batch API — Specific User IDs (Most Efficient)
csharp

using Microsoft.Graph.Models;
using System.Text.Json;

public async Task<List<UserProfileModel>> GetUserProfilesBatchAsync(
    IEnumerable<string> userIds,
    CancellationToken cancellationToken = default)
{
    var results = new List<UserProfileModel>();
    var idList = userIds.ToList();
    const int batchSize = 20;

    for (int i = 0; i < idList.Count; i += batchSize)
    {
        var chunk = idList.Skip(i).Take(batchSize).ToList();

        var batchRequestContent = new BatchRequestContentCollection(_graphClient);
        var requestIdMap = new Dictionary<string, string>(); // requestId -> userId

        foreach (var userId in chunk)
        {
            var requestInfo = _graphClient.Users[userId]
                .ToGetRequestInformation(config =>
                {
                    config.QueryParameters.Select = new[]
                    {
                        "id", "displayName", "mail", "userPrincipalName",
                        "jobTitle", "department", "officeLocation", "mobilePhone", "employeeId"
                    };
                });

            var requestId = await batchRequestContent.AddBatchRequestStepAsync(requestInfo);
            requestIdMap[requestId] = userId;
        }

        var batchResponse = await _graphClient.Batch
            .PostAsync(batchRequestContent, cancellationToken);

        foreach (var (requestId, userId) in requestIdMap)
        {
            try
            {
                var user = await batchResponse
                    .GetResponseByIdAsync<User>(requestId);

                if (user != null)
                    results.Add(MapToModel(user));
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to get profile for user {UserId}", userId);
            }
        }

        // Throttle between batches to avoid 429s
        if (i + batchSize < idList.Count)
            await Task.Delay(200, cancellationToken);
    }

    return results;
}

7. SharePoint UPA Custom Properties (via SharePoint REST)
For custom SharePoint profile fields not in Entra ID:
csharp

using System.Net.Http.Headers;

public class SharePointUpaService
{
    private readonly HttpClient _httpClient;
    private readonly string _tenantUrl;

    public SharePointUpaService(HttpClient httpClient, string tenantUrl)
    {
        _httpClient = httpClient;
        _tenantUrl = tenantUrl;
    }

    public async Task<Dictionary<string, string>> GetUpaPropertiesAsync(
        string accessToken, string accountName)
    {
        // accountName = "i:0#.f|membership|user@domain.com"
        var encodedAccount = Uri.EscapeDataString($"'{accountName}'");
        var url = $"{_tenantUrl}/_api/SP.UserProfiles.PeopleManager/" +
                  $"GetPropertiesFor(accountName=@v)?@v={encodedAccount}";

        _httpClient.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", accessToken);
        _httpClient.DefaultRequestHeaders.Accept
            .Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        var doc = JsonDocument.Parse(json);

        var properties = new Dictionary<string, string>();

        if (doc.RootElement.TryGetProperty("UserProfileProperties", out var props))
        {
            foreach (var prop in props.EnumerateArray())
            {
                var key = prop.GetProperty("Key").GetString();
                var value = prop.GetProperty("Value").GetString();
                if (key != null)
                    properties[key] = value ?? string.Empty;
            }
        }

        return properties;
    }
}

8. Mapper
csharp

private static UserProfileModel MapToModel(User user) => new()
{
    Id                 = user.Id,
    DisplayName        = user.DisplayName,
    Mail               = user.Mail,
    UserPrincipalName  = user.UserPrincipalName,
    JobTitle           = user.JobTitle,
    Department         = user.Department,
    OfficeLocation     = user.OfficeLocation,
    City               = user.City,
    Country            = user.Country,
    MobilePhone        = user.MobilePhone,
    EmployeeId         = user.EmployeeId
};

10. Azure Function Entry Point (.NET Isolated)
csharp

[Function("BulkUserProfileExport")]
public async Task Run([TimerTrigger("0 0 2 * * *")] TimerInfo timer)
{
    _logger.LogInformation("BulkUserProfileExport triggered at {Time}", DateTime.UtcNow);

    var profiles = await _userProfileService.GetAllUsersWithManagerAsync();

    _logger.LogInformation("Retrieved {Count} profiles", profiles.Count);

    // → Export to SharePoint List, Azure SQL, or Blob JSON
    await _exportService.WriteToSharePointListAsync(profiles);
}

Key Notes for  Context
ConcernRecommendationThrottlingUse PageIterator + 200ms delay between batches; handle 429 with Retry-After headerPermissionsUser.Read.All (App) via Entra App Registration — same cert you use for SharePointCustom UPA fieldsUse SharePoint REST /_api/SP.UserProfiles.PeopleManager/ — Graph doesn't expose theseLarge tenants (5k+ users)Use $filter by department/region to split load, or use $batch with known user listsOutputWrite to Azure SQL or SharePoint List via Graph POST /sites/{id}/lists/{id}/items
Want me to add the Azure SQL write layer or the SharePoint List export on top of this?
Sonnet 4.6 LowClaude is AI and can make mistakes. Please double-check responses.
