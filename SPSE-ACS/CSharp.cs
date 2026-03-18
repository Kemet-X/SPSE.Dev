using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SPCSOM4
{
    class TestApp
    {
        static async Task Main(string[] args)
        {
            string siteUrl = "siteURL";
            string clientId = "app ID";
            string clientSecret = "client secret";
             
            string realm = GetRealmFromTargetUrl(new Uri(siteUrl));
            Console.WriteLine("releam: " + realm);
            string accessToken = await GetAppOnlyAccessToken(siteUrl, realm, clientId, clientSecret);

            using (var context = new ClientContext(siteUrl))
            {
                context.ExecutingWebRequest += (s, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                };

                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                Console.WriteLine("Connected. Site title: " + web.Title);


                context.Load(web.Lists,
                 lists => lists.Include(list => list.Title,
                                        list => list.Id));
                context.ExecuteQuery();
                    foreach (List list in web.Lists)
                    {
                        Console.WriteLine(list.Title);
                    }
               
            }
        }

        static string   GetRealmFromTargetUrl(Uri targetApplicationUri)
        {
            using (var client = new System.Net.WebClient())
            {
                client.Headers.Add("Authorization", "Bearer ");
                try { client.DownloadString(targetApplicationUri); }
                catch (System.Net.WebException e)
                {
                    string bearer = e.Response.Headers["WWW-Authenticate"];
                    int realmIndex = bearer.LastIndexOf("=");
                    return bearer.Substring(realmIndex + 1).Replace("\"", "");
                } 
            }  
            return null;
        }

        static async Task<string> GetAppOnlyAccessToken(string siteUrl, string realm, string clientId, string clientSecret)
        {
            string formattedClientId = string.Format("{0}@{1}", clientId, realm);
            string formattedPrincipal = string.Format("{0}@{1}",
                new Uri(siteUrl).Authority, realm);

            var body = new Dictionary<string, string>()
        {
            { "grant_type", "client_credentials" },
            { "client_id", formattedClientId },
            { "client_secret", clientSecret },
            { "resource", formattedPrincipal }
        };

            using (var http = new HttpClient())
            {
                string acsUrl = $"https://accounts.accesscontrol.windows.net/{realm}/tokens/OAuth/2";

                var response = await http.PostAsync(acsUrl, new FormUrlEncodedContent(body));
                var json = await response.Content.ReadAsStringAsync();

                //return json.access_token;
                return json.ToString();
            }
        }
    }
}
