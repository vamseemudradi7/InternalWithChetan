using Newtonsoft.Json;
using System;
using System.Net.Http;

namespace IEXTradingExample
{
    public class CompanyInfoResponse
    {
        public string symbol { get; set; }
        public string companyName { get; set; }
        public string exchange { get; set; }
        public string industry { get; set; }
        public string website { get; set; }
        public string description { get; set; }
        public string CEO { get; set; }
        public string issueType { get; set; }
        public string sector { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            //Getting Forbidden Error for this one. Not of much use.
            var symbol = "msft";
            var IEXTrading_API_PATH = "https://api.iextrading.com/1.0/stock/{0}/company";

            IEXTrading_API_PATH = string.Format(IEXTrading_API_PATH, symbol);

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                //For IP-API
                client.BaseAddress = new Uri(IEXTrading_API_PATH);
                HttpResponseMessage response = client.GetAsync(IEXTrading_API_PATH).GetAwaiter().GetResult();
                if (response.IsSuccessStatusCode)
                {
                    var companysInfo = JsonConvert.DeserializeObject<CompanyInfoResponse>(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                    if (companysInfo != null)
                    {
                        Console.WriteLine("Company Name: " + companysInfo.companyName);
                        Console.WriteLine("Industry: " + companysInfo.industry);
                        Console.WriteLine("Sector: " + companysInfo.sector);
                        Console.WriteLine("Website: " + companysInfo.website);
                    }
                }
            }
        }
    }
}
