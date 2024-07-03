using MailTest;
using Outlook = Microsoft.Office.Interop.Outlook;

try
{
    Outlook.Application objOutlook = new ();
    OutlookService.EnumerateAccounts(objOutlook);
    var currentUser = OutlookService.GetCurrentUserInfo(objOutlook);
    var senderEmail = currentUser.PrimarySmtpAddress;
    var recipientEmail = senderEmail;
    var outlookAccount = OutlookService.GetAccountForEmailAddress(objOutlook, senderEmail);
    //OutlookService.SendEmailFromAccount(objOutlook,"test","test",recipientEmail, senderEmail);
}
catch (Exception e)
{
    Console.WriteLine(e.Message);
}