using Microsoft.SharePoint.Client;

namespace PeoplePickerSearchApp
{
    internal class AccessTokenClientContext : ClientContext
    {
        private string AccessToken { get; }

        internal AccessTokenClientContext(string siteUrl, string accessToken) : base(siteUrl)
        {
            AccessToken = accessToken;
            ExecutingWebRequest += OnExecutingWebRequest;
        }

        private void OnExecutingWebRequest(object? sender, WebRequestEventArgs e)
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + AccessToken;
        }
    }
}
