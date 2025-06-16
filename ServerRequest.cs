using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace moduleIntegrationFio
{
    internal class ServerRequest
    {
        public async Task<string> GetRequestAsync(string url)
        {
            using (HttpClient client = new HttpClient())
            {
                HttpResponseMessage response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();

                    return content;
                }
                else
                {
                    return $"Error: {response.StatusCode}";
                }
            }
        }
    }
}

