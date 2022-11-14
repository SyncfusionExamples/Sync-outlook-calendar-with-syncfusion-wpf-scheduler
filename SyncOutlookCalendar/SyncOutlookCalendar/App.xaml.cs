using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Identity.Client;


namespace SyncOutlookCalendar
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        //public static UIParent UiParent;
        public static IPublicClientApplication ClientApplication;
        //You need to replace your Application ID
        public static string ClientID = "8d0bd3de-1b47-4844-9bd0-61744355cc3b";
        public static string[] Scopes = { "User.Read", "Calendars.Read", "Calendars.ReadWrite" };
        private const string Tenant = "77f1fe12-b049-4919-8c50-9fb41e5bb63b";
        private const string Authority = "https://login.microsoftonline.com/" + Tenant;

        public App()
        {
            ClientApplication = PublicClientApplicationBuilder.Create(ClientID)
            .WithAuthority(AzureCloudInstance.AzurePublic, Tenant)
            .WithDefaultRedirectUri()
            .Build();
        }


    }
}
