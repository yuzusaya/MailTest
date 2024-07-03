namespace MailTest;

using Microsoft.Office.Interop.Outlook;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
public class OutlookService
{
    public static void EnumerateAccounts(Outlook.Application application)
    {
        Outlook.Accounts accounts = application.Session.Accounts;
        foreach (Outlook.Account account in accounts)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Account: " + account.DisplayName);
                if (string.IsNullOrEmpty(account.SmtpAddress)
                    || string.IsNullOrEmpty(account.UserName))
                {
                    Outlook.AddressEntry oAE =
                        account.CurrentUser.AddressEntry
                        as Outlook.AddressEntry;
                    if (oAE.Type == "EX")
                    {
                        Outlook.ExchangeUser oEU =
                            oAE.GetExchangeUser()
                            as Outlook.ExchangeUser;
                        sb.AppendLine("UserName: " +
                            oEU.Name);
                        sb.AppendLine("SMTP: " +
                            oEU.PrimarySmtpAddress);
                        sb.AppendLine("Exchange Server: " +
                            account.ExchangeMailboxServerName);
                        sb.AppendLine("Exchange Server Version: " +
                            account.ExchangeMailboxServerVersion);
                    }
                    else
                    {
                        sb.AppendLine("UserName: " +
                            oAE.Name);
                        sb.AppendLine("SMTP: " +
                            oAE.Address);
                    }
                }
                else
                {
                    sb.AppendLine("UserName: " +
                        account.UserName);
                    sb.AppendLine("SMTP: " +
                        account.SmtpAddress);
                    if (account.AccountType ==
                        Outlook.OlAccountType.olExchange)
                    {
                        sb.AppendLine("Exchange Server: " +
                            account.ExchangeMailboxServerName);
                        sb.AppendLine("Exchange Server Version: " +
                            account.ExchangeMailboxServerVersion);
                    }
                }
                if (account.DeliveryStore != null)
                {
                    sb.AppendLine("Delivery Store: " +
                        account.DeliveryStore.DisplayName);
                }
                sb.AppendLine("---------------------------------");
                Console.Write(sb.ToString());
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
    public static ExchangeUser? GetCurrentUserInfo(Outlook.Application application)
    {
        Outlook.AddressEntry addrEntry = application.Session.CurrentUser.AddressEntry;
        if (addrEntry.Type == "EX")
        {
            Outlook.ExchangeUser currentUser =
                application.Session.CurrentUser.
                AddressEntry.GetExchangeUser();
            if (currentUser != null)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Name: "
                    + currentUser.Name);
                sb.AppendLine("STMP address: "
                    + currentUser.PrimarySmtpAddress);
                sb.AppendLine("Title: "
                    + currentUser.JobTitle);
                sb.AppendLine("Department: "
                    + currentUser.Department);
                sb.AppendLine("Location: "
                    + currentUser.OfficeLocation);
                sb.AppendLine("Business phone: "
                    + currentUser.BusinessTelephoneNumber);
                sb.AppendLine("Mobile phone: "
                    + currentUser.MobileTelephoneNumber);
                Console.WriteLine(sb.ToString());
                return currentUser;
            }
        }
        return null;
    }
    public static void SendEmailFromAccount(Outlook.Application application, string subject, string body, string to, string smtpAddress)
    {

        // Create a new MailItem and set the To, Subject, and Body properties.
        Outlook.MailItem newMail = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);
        newMail.To = to;
        newMail.Subject = subject;
        newMail.Body = body;

        // Retrieve the account that has the specific SMTP address.
        Outlook.Account account = GetAccountForEmailAddress(application, smtpAddress);
        // Use this account to send the email.
        newMail.SendUsingAccount = account;
        newMail.Send();
    }


    public static Outlook.Account GetAccountForEmailAddress(Outlook.Application application, string smtpAddress)
    {

        // Loop over the Accounts collection of the current Outlook session.
        Outlook.Accounts accounts = application.Session.Accounts;
        foreach (Outlook.Account account in accounts)
        {
            // When the email address matches, return the account.
            if (account.SmtpAddress == smtpAddress)
            {
                return account;
            }
        }
        throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!", smtpAddress));
    }

}