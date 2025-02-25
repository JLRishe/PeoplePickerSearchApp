using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;

namespace PeoplePickerSearchApp
{
    class Program
    {
        // Microsoft Entra registrations:
        // Specifies the Microsoft Entra tenant ID
        const string AadTenantId = "ENTER_HERE";
        // Specifies the Application (client) ID of the console application registration in Microsoft Entra ID
        const string ClientId = "ENTER_HERE";
        const string ClientCertificatePath = "ENTER_HERE";
        const string ClientCertificatePassword = "ENTER_HERE";
        // Specifies the redirect URL for the client that was configured for console application registration in Microsoft Entra ID
        const string ClientRedirectUrl = "http://localhost";

        const string SharePointTenantUrl = "ENTER_HERE";

        private static readonly (bool, bool) DelegatedNoSubstrate = (false, false);
        private static readonly (bool, bool) DelegatedUseSubstrate = (false, true);
        private static readonly (bool, bool) AppOnlyNoSubstrate = (true, false);
        private static readonly (bool, bool) AppOnlyUseSubstrate = (true, true);

        private static readonly Dictionary<(bool, bool), List<double>> Timings = new()
        {
            { DelegatedNoSubstrate, new List<double>()  },
            { DelegatedUseSubstrate, new List<double>() },
            { AppOnlyNoSubstrate, new List<double>() },
            { AppOnlyUseSubstrate, new List<double>() },
        };

        private static async Task<AuthenticationResult> GetToken(bool getAppOnlyToken)
        {
            var authority = "https://login.microsoftonline.com/" + AadTenantId;
            var scopes = new[] { $"{SharePointTenantUrl}/.default" };

            if (getAppOnlyToken)
            {
                var cert = new X509Certificate2(ClientCertificatePath, ClientCertificatePassword);
                var confidentialClient = ConfidentialClientApplicationBuilder
                        .Create(ClientId)
                        .WithCertificate(cert)
                        .WithAuthority(authority, false)
                        .Build();

                return await confidentialClient.AcquireTokenForClient(scopes).ExecuteAsync();

            }

            var publicClient = PublicClientApplicationBuilder
                    .Create(ClientId)
                    .WithAuthority(authority, false)
                    .WithRedirectUri(ClientRedirectUrl)
                    .Build();

            return await publicClient.AcquireTokenInteractive(scopes).ExecuteAsync();
        }

        private static async Task<PeopleSearchResult[]?> SearchPeopleRest(string token, string searchTerm, bool useSubstrate)
        {
            var parameters = new PeoplePickerSearchUserPayload
            {
                queryParams = new PeoplePickerSearchUserQueryParams
                {
                    PrincipalType = 1,
                    PrincipalSource = 15,
                    QueryString = searchTerm,
                    AllowMultipleEntities = false,
                    MaximumEntitySuggestions = 200,
                    UseSubstrateSearch = useSubstrate,
                },
            };

            var payload = JsonSerializer.Serialize(parameters);

            var client = new HttpClient();

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var response = await client.PostAsync(
                $"{SharePointTenantUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser",
                new StringContent(payload, Encoding.UTF8, "application/json")
            );

            var responseText = await response.Content.ReadAsStringAsync();

            if (responseText == null)
            {
                throw new ApplicationException("null response");
            }

            var odataResponse = JsonSerializer.Deserialize<OdataResponse>(responseText);

            if (odataResponse == null)
            {
                throw new ApplicationException("null response");
            }

            return JsonSerializer.Deserialize<PeopleSearchResult[]>(odataResponse.value);
        }

        private static async Task QueryAndListPeople(string token, (bool, bool) options, string searchTerm)
        {
            var isAppOnly = options.Item1;
            var useSubstrate = options.Item2;

            var startTime = DateTime.Now;
            var result = await SearchPeopleRest(token, searchTerm, useSubstrate);
            var endTime = DateTime.Now;

            var duration = endTime - startTime;
            Timings[options].Add(duration.TotalSeconds);

            Console.WriteLine($"App-only: {isAppOnly} | UseSubstrate: {useSubstrate} | Results: {(result is null ? "(null)" : result.Length.ToString())} | Time: {duration.TotalSeconds:0.00} seconds");

            if (result != null)
            {
                foreach (var item in result)
                {
                    Console.WriteLine($"- {item.DisplayText} ({item.Key})");
                }
            }
        }

        static async Task Main()
        {
            var delegatedToken = (await GetToken(false)).AccessToken;
            var appOnlyToken = (await GetToken(true)).AccessToken;

            for (; ; )
            {
                Console.Write("Enter search term (x to exit): ");
                var searchTerm = Console.ReadLine() ?? "";

                if (searchTerm == "x")
                {
                    break;
                }

                await QueryAndListPeople(delegatedToken, DelegatedNoSubstrate, searchTerm);
                await QueryAndListPeople(delegatedToken, DelegatedUseSubstrate, searchTerm);
                await QueryAndListPeople(appOnlyToken, AppOnlyNoSubstrate, searchTerm);
                await QueryAndListPeople(appOnlyToken, AppOnlyUseSubstrate, searchTerm);

                Console.WriteLine();
            }

            foreach (var (options, durations) in Timings)
            {
                var total = durations.Sum();
                var average = total / durations.Count;
                var (appOnly, useSubstrate) = options;

                Console.WriteLine($"App-only: {appOnly} | UseSubstrate: {useSubstrate} | Average: {average:0.00} seconds");
            }
        }
    }
}