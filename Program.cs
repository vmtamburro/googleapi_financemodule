using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Configuration;

namespace GmailQuickstart
{
    class Program
    {
        private const string AVAIL_BAL_ID = "Available Balance:";
        private const string LEDG_BAL_ID = "Ledger Balance:";
        private const string TIME_ID = "ET on";
        private const string END_TIME_ID = ", your account";

        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Finance Module";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            UsersResource.MessagesResource.ListRequest inboxRequest = service.Users.Messages.List("me");
            IEnumerable<string> messageIds = inboxRequest.Execute().Messages.ToList().Select(x => x.Id);
            List<Message> messages = new List<Message>();
            foreach (var id in messageIds)
            {
                var message = service.Users.Messages.Get("me", id).Execute();
                messages.Add(message);

            }

            var pncAlerts = messages.Where(x => x.LabelIds.Contains(ConfigurationManager.AppSettings["LabelID"]));
            var BalanceDictionary = new Dictionary<DateTime, decimal>();
            foreach (var bankAlert in pncAlerts)
            {
                var availableBalanceIndex = bankAlert.Snippet.IndexOf(AVAIL_BAL_ID);
                var ledgerBalanceIndex = bankAlert.Snippet.IndexOf(LEDG_BAL_ID);

                var dateIndex = bankAlert.Snippet.IndexOf(TIME_ID);
                var endDateIndex = bankAlert.Snippet.IndexOf(END_TIME_ID);
                decimal balance = 0;
                DateTime balanceDate = DateTime.MinValue;

                balance = ParseBalance(bankAlert, availableBalanceIndex, ledgerBalanceIndex, balance);
                balanceDate = ParseBalanceDate(bankAlert, dateIndex, endDateIndex, balanceDate);
                BalanceDictionary.Add(balanceDate, balance);
            }

            foreach(var item in BalanceDictionary)
            {
                //Save to Sqlite DB
                Console.WriteLine($"Date: {item.Key}, Balance: {item.Value}");
                Console.ReadKey();
            }
        }

        private static DateTime ParseBalanceDate(Message bankAlert, int dateIndex, int endDateIndex, DateTime balanceDate)
        {
            if (dateIndex != -1 && endDateIndex != -1)
            {
                var startDate = dateIndex + TIME_ID.Length;
                var lengthDate = endDateIndex - startDate;
                var dateString = bankAlert.Snippet.Substring(startDate, lengthDate).Trim();
                DateTime.TryParse(dateString, out balanceDate);
            }

            return balanceDate;
        }

        private static decimal ParseBalance(Message bankAlert, int availableBalanceIndex, int ledgerBalanceIndex, decimal balance)
        {
            if (availableBalanceIndex != -1 && ledgerBalanceIndex != -1)
            {
                var startBalance = availableBalanceIndex + AVAIL_BAL_ID.Length;
                var lengthBalance = ledgerBalanceIndex - startBalance;
                var balanceString = bankAlert.Snippet.Substring(startBalance, lengthBalance).Trim().Replace("$", "");
                decimal.TryParse(balanceString, out balance);
            }

            return balance;
        }
    }
}