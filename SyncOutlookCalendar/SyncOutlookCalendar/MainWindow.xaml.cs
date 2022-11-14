using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using System;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Threading;
using System.Windows;
using static System.Formats.Asn1.AsnWriter;
using System.Net.Http;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Graph.Models.Security;
using System.Collections.Generic;
using System.Xml.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text;

namespace SyncOutlookCalendar
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Authenticate();
        }

        static async Task Authenticate()
        {
            var accounts = await App.ClientApplication.GetAccountsAsync();
            AuthenticationResult tokenRequest;
            if (accounts.Count() > 0)
            {
                tokenRequest = await App.ClientApplication.AcquireTokenSilent(App.Scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }
            else
            {
                tokenRequest = await App.ClientApplication.AcquireTokenInteractive(App.Scopes).ExecuteAsync();

            }

            HttpClient client = new HttpClient();
            var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, "https://graph.microsoft.com/v1.0");

            //Add the token in Authorization header
            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest?.AccessToken);
          
            //var response = await client.SendAsync(request);

            //HttpResponseMessage getEvents = await client.GetAsync("https://graph.microsoft.com/beta/me/calendar/events");
            //// if (getEvents.IsSuccessStatusCode)
            //{
            //    string json = await getEvents.Content.ReadAsStringAsync();
            //    var result1 = JsonConvert.DeserializeObject(json);

            //    // foreach (JProperty child in result.Properties().Where(p => !p.Name.StartsWith("@")))
            //} 

            Uri addEventUri = new Uri("https://graph.microsoft.com/beta/me/calendar/events");

            Event newEvent = new Event()
            {
                Subject = "Sample Data 1",
                Start = new DateTimeTimeZone() { DateTime = DateTime.Now.ToString(), TimeZone = "Eastern Standard Time" },
                End = new DateTimeTimeZone() { DateTime =  DateTime.Now.AddHours(1).ToString(), TimeZone = "Eastern Standard Time" },
            };
            HttpContent addEventContent = new StringContent(JsonConvert.SerializeObject(newEvent), Encoding.UTF8, "application/json");
            HttpResponseMessage addEventResponse = await client.PostAsync(addEventUri, addEventContent);

            if (addEventResponse.IsSuccessStatusCode)
            {
                Console.WriteLine("Event has been added successfully!");
            }
        }
    }
}
