using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SharePointSearch
{
    class Program
    {
        // Azure AD Application details
        private static readonly string tenantId = "";
        private static readonly string clientId = "";
        private static readonly string clientSecret = "";
        private static readonly string graphApiEndpoint = "https://graph.microsoft.com/beta/search/query";

        static async Task Main(string[] args)
        {
            try
            {
                string accessToken = await GetAccessTokenAsync();
                await SearchSharePointAsync(accessToken, "Test"); // Replace "test" with your query
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Method to acquire an OAuth 2.0 token
        private static async Task<string> GetAccessTokenAsync()
        {
            using (var client = new HttpClient())
            {
                var requestBody = new StringContent(
                    $"client_id={clientId}&client_secret={clientSecret}&scope=https://graph.microsoft.com/.default&grant_type=client_credentials",
                    Encoding.UTF8, "application/x-www-form-urlencoded"
                );

                var response = await client.PostAsync($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token", requestBody);
                response.EnsureSuccessStatusCode();

                var responseContent = await response.Content.ReadAsStringAsync();
                var responseJson = JsonConvert.DeserializeObject<dynamic>(responseContent);

                return responseJson.access_token;
            }
        }

        // Method to search SharePoint using Graph API
        private static async Task SearchSharePointAsync(string accessToken, string query)
        {
            using (var client = new HttpClient())
            {
                // Set Authorization header with the Bearer token
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Define the request body
                var requestBody = new
                {
                    requests = new[]
                    {
                        new
                        {
                            entityTypes = new[] { "driveItem", "listItem", "list", "drive", "site" },
                            query = new { queryString = query },
                            region = "NAM",
                            size = 50,
                            from = 0,
                            fields = new[] { "WebUrl", "lastModifiedBy", "name" }
                        }
                    }
                };

                var jsonContent = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                // Send the request to the Graph API
                var response = await client.PostAsync(graphApiEndpoint, content);
                response.EnsureSuccessStatusCode();

                // Parse and display the response
                var responseContent = await response.Content.ReadAsStringAsync();
                dynamic responseJson = JsonConvert.DeserializeObject(responseContent);

                // Output search results
                Console.WriteLine("Search Results:");
                foreach (var hit in responseJson.value[0].hitsContainers[0].hits)
                {
                    Console.WriteLine($"Name: {hit.resource.name}");
                    Console.WriteLine($"Web URL: {hit.resource.webUrl}");
                    //Console.WriteLine($"Last Modified: {hit.resource.lastModifiedDateTime}");
                    Console.WriteLine($"Last Modified By: {hit.resource.lastModifiedBy.user.displayName}");
                    Console.WriteLine(new string('-', 50));
                }

                // Check if more results are available
                bool moreResultsAvailable = responseJson.value[0].hitsContainers[0].moreResultsAvailable;
                Console.WriteLine($"More Results Available: {moreResultsAvailable}");
            }
        }
    }
}
