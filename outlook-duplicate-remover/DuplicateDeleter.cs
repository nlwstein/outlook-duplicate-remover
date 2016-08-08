using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace outlook_duplicate_remover
{
    class DuplicateDeleter
    {

        private MAPIFolder _deletedFolder;
        private MAPIFolder DeletedFolder
        {
            get
            {
                if (_deletedFolder == null)
                {
                    var folderExists = false;
                    foreach (MAPIFolder folder in CurrentFolder.Folders)
                    {
                        if (folder.Name == "Duplicates")
                        {
                            folderExists = true;
                            _deletedFolder = folder;
                            break;
                        }
                    }
                    if (!folderExists)
                    {
                        _deletedFolder = CurrentFolder.Folders.Add("Duplicates");
                    }
                }
                return _deletedFolder;
            }
        }
        private MAPIFolder _currentFolder;
        private MAPIFolder CurrentFolder
        {
            get
            {
                if (_currentFolder == null)
                {
                    _currentFolder = Application.ActiveExplorer().CurrentFolder;
                }
                return _currentFolder;
            }
        }
        private Microsoft.Office.Interop.Outlook.Application _application;
        private Microsoft.Office.Interop.Outlook.Application Application
        {
            get
            {
                if (_application == null)
                {
                    _application = Globals.ThisAddIn.Application;
                }
                return _application;
            }
        }
        private static DuplicateDeleter _instance;
        public static DuplicateDeleter Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new DuplicateDeleter();
                }
                return _instance;
            }
        }

        public void Reset()
        {
            _deletedFolder = null;
            _currentFolder = null;
        }

        public void DeleteDuplicates()
        {
            if (CurrentFolder.DefaultItemType != OlItemType.olMailItem)
            {
                MessageBox.Show("Selected folder is not for mail.");
            }
            var totalNumberOfItems = CurrentFolder.Items.Count;
            var totalNumberOfItemsString = totalNumberOfItems.ToString();
            var unicodeEncoding = new UnicodeEncoding();
            var sha = new SHA256Managed();
            var mailHashes = new List<byte[]>();
            var itemsToMove = new List<MailItem>();
            var progress = new Progress(totalNumberOfItems);
            progress.Show();
            var numberOfItems = 0;
            foreach (var item in CurrentFolder.Items)
            {
                progress.UpdateProgressBar();
                numberOfItems += 1;
                progress.ProgressMessage = string.Format("Processing email {0} of {1}...", numberOfItems.ToString(), totalNumberOfItemsString);
                // Cast to being a mailitem, so we can grab mail-specific info: 
                var mailItem = (MailItem)item;
                // Create a hash of the identifying fields (from, sent timestamp, message body, subject)
                string mailBody;
                switch (mailItem.BodyFormat)
                {
                    case OlBodyFormat.olFormatHTML:
                        mailBody = mailItem.HTMLBody;
                        break;
                    case OlBodyFormat.olFormatRichText:
                        mailBody = mailItem.RTFBody;
                        break;
                    case OlBodyFormat.olFormatPlain:
                    case OlBodyFormat.olFormatUnspecified:
                    default:
                        mailBody = mailItem.Body;
                        break;
                }
                // Create a unique identifier that consists of the fields mentioned in prev. comment
                var uniqueIdentifier = string.Format("{0}_{1}_{2}_{3}", mailItem.SenderEmailAddress, mailItem.SentOn.ToString(), mailBody, mailItem.Subject);
                // Hash it
                byte[] hash;
                var messageBytes = unicodeEncoding.GetBytes(uniqueIdentifier);
                hash = sha.ComputeHash(messageBytes);
                if (mailHashes.Any(x => x.SequenceEqual(hash)))
                {
                    itemsToMove.Add(mailItem);
                }
                else
                {
                    mailHashes.Add(hash);
                }
            }
            var numOfEmails = 0;
            if (itemsToMove.Count > 0)
            {
                progress.ResetProgressBar(itemsToMove.Count);
            }
            CurrentFolder.InAppFolderSyncObject = false;
            foreach (var itemToMove in itemsToMove)
            {
                numOfEmails += 1;
                progress.UpdateProgressBar();
                progress.ProgressMessage = string.Format("Email # {0} moving...", numOfEmails);
                itemToMove.Move(DeletedFolder);
            }
            progress.Close();
            CurrentFolder.InAppFolderSyncObject = true;
        }
    }
}
