using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Exception = System.Exception;
using System.Net;


namespace PromotionsToPrimary
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }




        public void MoveToPrimary()
        {
            //Create list to store emails
            List<MailItem> emails = new List<MailItem>();


            //email address to search;
            List<string> accountNames = new List<string> {"John.Miller052792@gmail.com", "Mary.Poppins19877@gmail.com",
                "Walter.White8777@gmail.com", "James.McGill247@gmail.com", "William.Hayes4555@gmail.com" };

            //get list of email accounts in Outlook
            Stores stores = Application.Session.Stores;


            //cycle through email accounts until the target account is found
            foreach (string accountName in accountNames)
            {
                foreach (Store store in stores)
                {
                    if (store.DisplayName.ToLower() == accountName.ToLower())
                    {
                        //set target account as default
                        MAPIFolder root = store.GetRootFolder();

                        //set Source Folder
                        MAPIFolder source = root.Folders["Promotions_tab"];

                        //set Destination Folder 
                        MAPIFolder destination = root.Folders["Primary_tab"];

                        //Find all emails in source Folder
                        Items items = source.Items;
                        MailItem mail = items.Find("[Size] > 0");

                        if (mail == null)
                        {
                            MessageBox.Show(String.Format("Sorry, no emails to move for {0}", accountName), "Promotions to Primary");
                            continue;
                        }

                        //scan folder for emails and add to the list
                        while (mail != null)
                        {
                            emails.Add(mail);
                            mail = items.FindNext();
                        }

                        try
                        {
                            //Move emails to the Destination Folder
                            foreach (MailItem email in emails)
                            {
                                email.Move(destination);
                            }

                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("We were unable to move some messages.  Please try again shortly.", accountName);
                            continue;
                        }

                    }
                }
            }

            //get count of emails moved and generate confirmation message
            MessageBox.Show(String.Format("We have moved {0} emails to Primary Tab folder", emails.Count), "Move Successful");

        }

        public void ClickThoseLinks()
        {
            //email addresses to search;
            List<string> accountNames = new List<string> { "John.Miller052792@gmail.com", "Mary.Poppins19877@gmail.com",
            "Walter.White8777@gmail.com", "James.McGill247@gmail.com", "William.Hayes4555@gmail.com",
            "James.McGill24@hotmail.com (1)", "John.Miller7277@outlook.com", "Walter.White877@hotmail.com",
            "William.Hayes45@hotmail.com"};

                        
            //get list of email accounts in Outlook
            Stores stores = Application.Session.Stores;

            //cycle through email accounts until the target account is found
            foreach (string accountName in accountNames)
            {
                foreach (Store store in stores)
                {
                    if (store.DisplayName.ToLower() == accountName.ToLower())
                    {
                        
                        List<string> links = new List<string>();
                        //reset links count to zero when going to next account in list
                        links.Clear();

                        List<MailItem> emails = new List<MailItem>();
                        //reset emails count when going to next account in list
                        emails.Clear();

                        //set target account as default
                        MAPIFolder root = store.GetRootFolder();

                        //set Source Folder
                        MAPIFolder source = root.Folders["Inbox"];

                        //Find all emails in source Folder
                        Items items = source.Items;

                        //find all emails received within the last 24 hours
                        DateTime now = DateTime.Now;

                        MailItem mail = items.Find("[ReceivedTime] >= '" + now.AddDays(-1).ToShortDateString() + "' AND [ReceivedTime] <=" +
                            "'" + now.ToString() + "'");

                        if (mail == null)
                        {
                            MessageBox.Show("Sorry, no Emails Received Within last 24 Hours", accountName);
                            continue;
                        }

                        //scan folder for emails and add to the list
                        while (mail != null)
                        {
                            emails.Add(mail);
                            mail = items.FindNext();
                        }

                        //extract a link in each email
                        foreach (MailItem message in emails)
                        {
                            //get each link in the email
                            string url = GetLink(message.HTMLBody, "href=\"", "\"");

                            //only add urls beginning with http
                            if (url.StartsWith("http") == false)
                            {
                                continue;
                            }

                            //add urls to list
                            links.Add(url);
                        }

                        //create web request to be sent once link clicked
                        foreach (string link in links)
                        {
                            try
                            {
                                HttpWebRequest req =
                                        (HttpWebRequest)WebRequest.Create(link);
                                                                                             
                                //send web request and wait for response
                                HttpWebResponse response =
                                    (HttpWebResponse)req.GetResponse();
                                response.Close();
                            }
                            //if website doesn't respond (e.g. Not Found, Forbidden), go to next link
                            catch (WebException ex)
                            {
                                HttpWebResponse webResponse = (HttpWebResponse)ex.Response;
                                if(webResponse != null)
                                {
                                    if (webResponse.StatusCode != HttpStatusCode.OK)
                                    {
                                        MessageBox.Show(String.Format("Cannot connect to {0}", link), "Sorry About Your Luck");

                                    }
                                }
                            }
                        }

                       
                        MessageBox.Show(String.Format("Clicked {0} links for {1}", links.Count, accountName), "Links Clicked");
                    }
                
                }
            }
            MessageBox.Show("Links Clicked!", "Complete");
        }

        //Method to get url from a tag
        public string GetLink(string str, string firstString, string lastString)
        {
            int pos1 = str.IndexOf(firstString) + firstString.Length;
            int pos2 = str.Substring(pos1).IndexOf(lastString);
            return str.Substring(pos1, pos2);

        }
    }
}
