using System;
using System.Net.Http;
//Async programming
using System.Threading.Tasks;
using System.Xml.Linq;
//Into project needs to be installed Newtonsoft Json by PM console
using Newtonsoft.Json.Linq;

class Program
{
    //Creating static HttpClient 
    private static readonly HttpClient client = new HttpClient();

    static async Task Main(string[] args)
    {
        try
        {
            //Creating HttpResponseMessage that waits for the client request
            HttpResponseMessage response = await client.GetAsync("https://api.exchangerate-api.com/v4/latest/USD");

            //Creating string represents the response content
            string responseBody = await response.Content.ReadAsStringAsync();

            //
            //Creating collection provided with NewtonsoftJson
            var data = JObject.Parse(responseBody);
            //Showing the actual currencies ;)
            Console.WriteLine("1 USD is equal to " + data["rates"]["EUR"] + " Euros");
            Console.WriteLine("1 USD is equal to " + data["rates"]["PLN"] + " PLN");
        }
        catch (HttpRequestException e)
        {
          
            Console.WriteLine($"Message:{e.Message}");
        }




        Console.ReadKey();
    }
}