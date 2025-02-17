using Microsoft.Identity.Client;
using Microsoft.SharePoint.ApplicationPages.ClientPickerQuery;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System.Security.Cryptography.X509Certificates;

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

        private static async Task<ClientContext> GetSharePointClient(bool useAppOnly)
        {
            var token = await GetToken(useAppOnly);

            return new AccessTokenClientContext(SharePointTenantUrl, token.AccessToken);
        }

        private static async Task<string> SearchPeople(ClientContext context, string searchTerm, bool useSubstrate)
        {
            var query = new ClientPeoplePickerQueryParameters
            {
                PrincipalType = PrincipalType.User,
                PrincipalSource = PrincipalSource.All,
                QueryString = searchTerm,

                AllowMultipleEntities = false,
                MaximumEntitySuggestions = 200,

                UseSubstrateSearch = useSubstrate,
            };

            var result = ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser(context, query);

            await context.ExecuteQueryAsync();

            return result.Value;

        }

        static async Task Main()
        {
            var appOnlyClient = await GetSharePointClient(true);

            var appOnlyResult = await SearchPeople(appOnlyClient, "Ailmel", false);

            Console.WriteLine(appOnlyResult);
        }
    }
}