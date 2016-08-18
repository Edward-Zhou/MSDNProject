using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using Microsoft.Office365.OutlookServices;

namespace O365Mvc.Helpers
{
    internal class MailOperations
    {

        //get mail folder collection
        internal async Task<List<string>> GetEmailFolders()
        {

            // Make sure we have a reference to the Outlook Services client
            var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");


            List<string> EmailFolders = new List<string>();

            var mailFolders = await outlookServicesClient.Me.Folders.ExecuteAsync();

            foreach (var mailFolder in mailFolders.CurrentPage)
            {
                EmailFolders.Add(mailFolder.DisplayName); 
            }

            return EmailFolders;
        }
        //create mail folder
        internal async Task<List<string>> CreateEmailFolder(List<string> FolderName)
        {
            var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");
            List<string> EmailFolders = new List<string>();
            foreach (var item in FolderName)
            {
                Folder newFolder = new Folder 
                {
                    DisplayName=item.ToString()
                };
                await outlookServicesClient.Me.RootFolder.ChildFolders.AddFolderAsync(newFolder);
            }
            EmailFolders=await GetEmailFolders();
            return EmailFolders;
        }
        //delete mail folder
        internal async Task<List<string>> DeleteEmailFolder(string FolderName)
        {
            var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");
            List<string> EmailFolders = new List<string>();
            var Folders = await outlookServicesClient.Me.Folders.ExecuteAsync();
            foreach (var mailfolder in Folders.CurrentPage)
            {
                if (mailfolder.DisplayName == FolderName)
                {                   
                   await mailfolder.DeleteAsync();
                }
            }
            EmailFolders = await GetEmailFolders();
            return EmailFolders;
        }
        //create mail 
        internal async Task<Action> CreateEmail()
        {
            var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");
            ItemBody body = new ItemBody
            {
                Content = "It was <b>My first Email</b>!",
                ContentType = BodyType.HTML
            };
            List<Recipient> toRecipients = new List<Recipient>();            
            Recipient toRecipient = new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = "v-tazho@OfficeDevGroup.onmicrosoft.com"
                }
            };
            toRecipients.Add(toRecipient);
            Message newMessage = new Message { 
                Subject="First Email Test",
                Body=body,
                ToRecipients = toRecipients
            };
            await outlookServicesClient.Me.SendMailAsync(newMessage,true);
            return null;
        }
        //get attachment
        internal async Task<Action> GetAttachment()
        {
            var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");
            var messages = await outlookServicesClient.Me.Folders["Inbox"].Messages
                .Where(m => m.HasAttachments == true)
                .Expand(m => m.Attachments)
                .Take(10)
                .ExecuteAsync();
            foreach (var message in messages.CurrentPage)
            {
                var attachments = message.Attachments.CurrentPage;

                foreach (FileAttachment attachment in attachments)
                {
                    byte[] test = attachment.ContentBytes;
                }
            }
            return null;
        }

    }
}