using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace LicenceValidator.Core
{
    public sealed class OAuthTokenService
    {
        private readonly HttpClient _http;
        private readonly ToolConfig _config;
        private readonly Dictionary<string, (string Token, DateTime Expires)> _cache = new Dictionary<string, (string, DateTime)>(StringComparer.OrdinalIgnoreCase);

        public OAuthTokenService(HttpClient http, ToolConfig config) { _http = http; _config = config; }

        public async Task<string> GetTokenAsync(string resourceRoot, CancellationToken ct)
        {
            var key = resourceRoot.TrimEnd('/');
            if (_cache.TryGetValue(key, out var cached) && cached.Expires > DateTime.UtcNow.AddMinutes(5)) return cached.Token;
            var url = "https://login.microsoftonline.com/" + _config.TenantId + "/oauth2/v2.0/token";
            var form = new Dictionary<string, string>
            {
                ["client_id"] = _config.ClientId, ["client_secret"] = _config.ClientSecret,
                ["grant_type"] = "client_credentials", ["scope"] = key + "/.default"
            };
            using (var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = new FormUrlEncodedContent(form) })
            using (var resp = await _http.SendAsync(req, ct).ConfigureAwait(false))
            {
                var payload = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!resp.IsSuccessStatusCode) throw new InvalidOperationException("Token request failed: " + (int)resp.StatusCode + " " + payload);
                using (var doc = JsonDocument.Parse(payload))
                {
                    var token = JsonHelper.GetRequiredString(doc.RootElement, "access_token");
                    var expiresIn = JsonHelper.GetInt32(doc.RootElement, "expires_in") ?? 3600;
                    _cache[key] = (token, DateTime.UtcNow.AddSeconds(expiresIn));
                    return token;
                }
            }
        }
    }

    public sealed class GraphService
    {
        private readonly HttpClient _http;
        private readonly OAuthTokenService _tokenService;
        private readonly ToolConfig _config;
        private readonly IAuditLogger _logger;
        private readonly string _baseUrl;
        private readonly string _resourceRoot;

        public GraphService(HttpClient http, OAuthTokenService tokenService, ToolConfig config, IAuditLogger logger)
        {
            _http = http; _tokenService = tokenService; _config = config; _logger = logger;
            _baseUrl = config.GraphBaseUrl.TrimEnd('/');
            _resourceRoot = new Uri(_baseUrl, UriKind.Absolute).GetLeftPart(UriPartial.Authority);
        }

        public bool EnabledByMode => _config.GraphEnabled;

        public async Task<List<TenantSkuRecord>> GetSubscribedSkusAsync(CancellationToken ct)
        {
            if (!EnabledByMode) return new List<TenantSkuRecord>();
            try
            {
                var rows = await GetODataCollectionAsync("/subscribedSkus?$select=skuId,skuPartNumber,consumedUnits,capabilityStatus,prepaidUnits", ct).ConfigureAwait(false);
                return rows.Select(row => new TenantSkuRecord
                {
                    SkuId = JsonHelper.GetRequiredGuid(row, "skuId"),
                    SkuPartNumber = JsonHelper.GetString(row, "skuPartNumber") ?? string.Empty,
                    ConsumedUnits = JsonHelper.GetInt32(row, "consumedUnits"),
                    CapabilityStatus = JsonHelper.GetString(row, "capabilityStatus"),
                    EnabledUnits = JsonHelper.GetInt32FromChild(row, "prepaidUnits", "enabled"),
                    SuspendedUnits = JsonHelper.GetInt32FromChild(row, "prepaidUnits", "suspended"),
                    WarningUnits = JsonHelper.GetInt32FromChild(row, "prepaidUnits", "warning")
                }).ToList();
            }
            catch (Exception ex)
            {
                if (_config.ShouldFailOnGraphError) throw;
                _logger.Warn("Graph subscribedSkus failed: " + TextHelper.Shorten(ex.Message));
                return new List<TenantSkuRecord>();
            }
        }

        public async Task<UserGraphLicenseSnapshot> GetUserLicenseSnapshotAsync(Guid aadObjectId, IReadOnlyDictionary<Guid, string> skuMap, CancellationToken ct)
            => await GetSnapshotByIdAsync(aadObjectId.ToString("D"), skuMap, ct).ConfigureAwait(false);

        public async Task<UserGraphLicenseSnapshot> GetUserLicenseSnapshotByUpnAsync(string upn, IReadOnlyDictionary<Guid, string> skuMap, CancellationToken ct)
            => await GetSnapshotByIdAsync(Uri.EscapeDataString(upn), skuMap, ct).ConfigureAwait(false);

        private async Task<UserGraphLicenseSnapshot> GetSnapshotByIdAsync(string userId, IReadOnlyDictionary<Guid, string> skuMap, CancellationToken ct)
        {
            if (!EnabledByMode) return UserGraphLicenseSnapshot.CreateUnknown("DisabledByMode", "Ist-Zustand konnte nicht geholt werden");
            try
            {
                var root = await GetRootAsync("/users/" + userId + "?$select=id,userPrincipalName,displayName,accountEnabled,assignedLicenses,licenseAssignmentStates", ct).ConfigureAwait(false);
                var snapshot = new UserGraphLicenseSnapshot
                {
                    EnabledByMode = true, Attempted = true, Found = true, GraphStatus = "Known",
                    ActualLicenseState = "Known", ActualLicenseMessage = "Ist-Zustand aus Graph geholt",
                    UserPrincipalName = JsonHelper.GetString(root, "userPrincipalName"),
                    DisplayName = JsonHelper.GetString(root, "displayName"),
                    AccountEnabled = JsonHelper.GetBoolean(root, "accountEnabled")
                };
                var assignmentModes = new Dictionary<Guid, string>();
                foreach (var state in JsonHelper.GetArray(root, "licenseAssignmentStates"))
                {
                    var skuId = JsonHelper.GetGuid(state, "skuId");
                    if (!skuId.HasValue) continue;
                    var assignedByGroup = JsonHelper.GetString(state, "assignedByGroup");
                    assignmentModes[skuId.Value] = string.IsNullOrWhiteSpace(assignedByGroup) ? "Direct" : "Group";
                }
                foreach (var license in JsonHelper.GetArray(root, "assignedLicenses"))
                {
                    var skuId = JsonHelper.GetGuid(license, "skuId");
                    if (!skuId.HasValue) continue;
                    var partNumber = skuMap.ContainsKey(skuId.Value) ? skuMap[skuId.Value] : "skuid:" + skuId.Value.ToString("D");
                    snapshot.ActualSkuPartNumbers.Add(partNumber);
                    if (assignmentModes.TryGetValue(skuId.Value, out var mode)) snapshot.AssignmentModeBySkuPartNumber[partNumber] = mode;
                }
                snapshot.ActualSkuPartNumbers = snapshot.ActualSkuPartNumbers.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
                if (snapshot.ActualSkuPartNumbers.Count == 0) snapshot.ActualLicenseMessage = "Ist-Zustand aus Graph geholt; keine zugewiesenen Lizenzen";
                return snapshot;
            }
            catch (InvalidOperationException ex) when (ex.Message.IndexOf("404", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new UserGraphLicenseSnapshot { EnabledByMode = true, Attempted = true, Found = false, GraphStatus = "UserNotFound", ActualLicenseState = "Unknown", ActualLicenseMessage = "Ist-Zustand konnte nicht geholt werden", Errors = new List<string> { "Graph user not found." } };
            }
            catch (Exception ex)
            {
                if (_config.ShouldFailOnGraphError) throw;
                var s = UserGraphLicenseSnapshot.CreateUnknown("Failed", "Ist-Zustand konnte nicht geholt werden", true);
                s.Attempted = true; s.Errors.Add("Graph lookup failed: " + TextHelper.Shorten(ex.Message));
                return s;
            }
        }

        private async Task<List<JsonElement>> GetODataCollectionAsync(string path, CancellationToken ct)
        {
            var results = new List<JsonElement>();
            string next = path;
            while (!string.IsNullOrWhiteSpace(next))
            {
                var root = await GetRootAsync(next, ct).ConfigureAwait(false);
                results.AddRange(JsonHelper.GetArray(root, "value").Select(x => x.Clone()));
                next = JsonHelper.GetString(root, "@odata.nextLink");
            }
            return results;
        }

        private async Task<JsonElement> GetRootAsync(string path, CancellationToken ct)
        {
            var token = await _tokenService.GetTokenAsync(_resourceRoot, ct).ConfigureAwait(false);
            int attempt = 0;
            while (true)
            {
                attempt++;
                var uri = path.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? new Uri(path) : new Uri(_baseUrl + path);
                using (var req = new HttpRequestMessage(HttpMethod.Get, uri))
                {
                    req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    using (var resp = await _http.SendAsync(req, ct).ConfigureAwait(false))
                    {
                        var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (resp.IsSuccessStatusCode)
                        {
                            using (var doc = JsonDocument.Parse(body)) return doc.RootElement.Clone();
                        }
                        if (attempt < _config.MaxRetryCount && ((int)resp.StatusCode == 429 || (int)resp.StatusCode >= 500))
                        {
                            var delay = TimeSpan.FromSeconds(Math.Min(30, Math.Pow(2, attempt)));
                            _logger.Warn("Graph GET failed " + (int)resp.StatusCode + ". Retrying in " + delay.TotalSeconds.ToString("N1") + "s.");
                            await Task.Delay(delay, ct).ConfigureAwait(false);
                            continue;
                        }
                        throw new InvalidOperationException("Graph request failed: " + (int)resp.StatusCode + " " + resp.ReasonPhrase + "\n" + body);
                    }
                }
            }
        }
    }
}