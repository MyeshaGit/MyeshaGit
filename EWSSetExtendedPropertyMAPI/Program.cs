using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
/// This sample uses a service account to search through a list of mailboxes for a particular folder under the Mailbox root. When the folder is found, it will search through all messages for a specific MAPI property.
/// If the message does not contain the MAPI property, it will stamp it on the message. If the message already contains the MAPI property, it will just move on to the next message.
/// To confirm the property has been stamped, use MFCMapi. https://mfcmapi.codeplex.com/
/// The EWS Managed API 2.2 must be installed first and the EWS dll must be referenced in this project. --> https://www.microsoft.com/en-us/download/details.aspx?id=42951
/// The service account / impersonation accoutn must have been assigned impersonation priviledges ---> https://msdn.microsoft.com/en-us/library/office/dn722377(v=exchg.150).aspx
///
/// </summary>

namespace EWSSetExtendedPropertyMAPI
{
    class Program
    {
        static void Main(string[] args)

        {
            string valueOfProperty = "zzzTestLabel";
            string nameOfProperty = "ComplianceTag";
            string nameOfFoldertoSearch = "CustomFolderName";
            string impersonationAccount = "ServiceAccount1@domain.com";
            string impersonationAccountPassword = "Password123";
            
            Guid myPropertySetId = new Guid("{403FC56B-CD30-47C5-86F8-EDE9E35A022B}"); //This is the GUID stamped on messages in MFCMapi. 
            ExtendedPropertyDefinition extendedPropertyDefinition = new ExtendedPropertyDefinition(myPropertySetId, "ComplianceTag", MapiPropertyType.String); //Constructs the MAPI property
            ItemView myItemView = new ItemView(int.MaxValue); //ItemView determines how many pages of results to return
            List<string> listOfPropsOnMessage = new List<string>(); //Store the properties on each message to a List

            FindItemsResults<Item> messagesInDesiredFolder = null;
            Console.WriteLine("Connecting to Exchange...."); //Begin trying to connect to Exchange; 

            //Create the Exchange Service. https://msdn.microsoft.com/en-us/library/office/dn567668(v=exchg.150).aspx
            ExchangeService myService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            myService.Credentials = new WebCredentials(impersonationAccount, impersonationAccountPassword);

            //To use AutoDiscover Look-up process, use this line instead: myService.AutodiscoverUrl("bigsmall@myrice.onmicrosoft.com", RedirectionUrlValidationCallback);
            myService.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            FolderView myFolderView = new FolderView(int.MaxValue); //FolderView with page size  - https://msdn.microsoft.com/en-us/library/office/dd633627(v=exchg.80).aspx
            myFolderView.Traversal = FolderTraversal.Deep;


            //If searching multiple mailboxes, it would probably be best to store the list of mailboxes in a List, then emulate through the list. 
            string userMailboxGoingToAccess = "BigSmall@myrice.onmicrosoft.com";

            //Bind impersonation account to the UserMailboxGoingToAccess
            myService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userMailboxGoingToAccess);


            //Start searching for folder at the root
            FolderId rootFolderID = new FolderId(WellKnownFolderName.Root, userMailboxGoingToAccess);
            Folder rootFolder = Folder.Bind(myService, WellKnownFolderName.Root);
            Console.WriteLine("The " + rootFolder.DisplayName + " has " + rootFolder.ChildFolderCount + " child folders.");

            FindFoldersResults myFolderResults = myService.FindFolders(rootFolder.Id, myFolderView);
            Folder desiredFolder = null;

            //Locate the folder to be searched by it's DisplayName.
            foreach (Folder currentFolder in myFolderResults)
            {
                if (currentFolder.DisplayName == nameOfFoldertoSearch)
                {
                    desiredFolder = currentFolder;
                    break;
                }

            }

            //Store the messages in messagesInDesiredFolder    

            try
            {
                messagesInDesiredFolder = desiredFolder.FindItems(myItemView);
            }
            catch (Exception ex)
            {
                throw new Exception("A folder named: " + nameOfFoldertoSearch + " does not exit. Proceed to next mailbox.");

            }

            //Cycle through each message in the desired folder
            foreach (EmailMessage msg in messagesInDesiredFolder)
            {
                //Search and add each MAPI property to a list
                foreach (ExtendedProperty extendedprop in msg.ExtendedProperties)
                {
                    listOfPropsOnMessage.Add(extendedprop.PropertyDefinition.Name);
                }

                //If the specified property (nameofProperty) is not already stamped on the message, stamp it now. 
                if (!listOfPropsOnMessage.Contains(nameOfProperty))
                {
                    msg.SetExtendedProperty(extendedPropertyDefinition, valueOfProperty);
                    msg.Update(ConflictResolutionMode.AlwaysOverwrite);
                    listOfPropsOnMessage.Clear(); //clear list
                }
            }

            Console.Read();

        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

    }
}


