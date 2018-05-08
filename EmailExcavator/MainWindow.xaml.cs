using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using EmailExcavator.ObjectClasses;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace EmailExcavator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Set the API Endpoint to Graph 'me' endpoint
        string _graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages?$select=body,subject,from,receivedDateTime&$top=15&$orderby=receivedDateTime%20DESC&$count=true";//Set the scope for API call to user.read
        public static JToken globalToken = new JObject();
        int globalIndex = 0;
        int globalEmailCount = 0;
        int globalEmailsRemaining = 0;
        AuthenticationResult globalAuthResult = null; 
        // ADD THE FOLLOWING FOR ONLY UNREAD INBOX MESSAGES --- "&$filter=isRead eq false"  - Remove Quotes

        string[] _scopes = new string[] { "Mail.Read" };//
        public MainWindow()
        {
            this.InitializeComponent();
        }

        // Lines 45 through 59: Get Between Function to attempt to gather the data between seperate emails
        public static string GetBetween(string strSource, string strStart, string strEnd)
        {
            int Start, End;
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
        }

        // Lines 61 through 93: Get Between function that gets the keywords in the body of the email
        public static string BodyGetBetween(string origString, string startString, string endString)
        {
            int Start, End;
            string Holder;
            if (origString != null)
            {

                if (origString.Contains(startString) && origString.Contains(endString))
                {
                    Start = origString.IndexOf(startString, 0) + startString.Length;
                    End = origString.IndexOf(endString, Start);
                    Holder = origString.Substring(Start, End - Start);
                    if (Holder.Contains(@"&nbsp;"))
                    {
                        Holder = Holder.Replace(@"&nbsp;", "");
                        return Holder;
                    }
                    else
                    {
                        return Holder;
                    }
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
         

        // Lines 96 through 130: Removes the HTML tags
        internal static string RemoveUnwantedTags(string data)
        {
            if (string.IsNullOrEmpty(data)) return string.Empty;

            var document = new HtmlDocument();
            document.LoadHtml(data);

            var acceptableTags = new String[] { "strong", "em", "u" };

            var nodes = new Queue<HtmlNode>(document.DocumentNode.SelectNodes("./*|./text()"));
            while (nodes.Count > 0)
            {
                var node = nodes.Dequeue();
                var parentNode = node.ParentNode;

                if (!acceptableTags.Contains(node.Name) && node.Name != "#text")
                {
                    var childNodes = node.SelectNodes("./*|./text()");

                    if (childNodes != null)
                    {
                        foreach (var child in childNodes)
                        {
                            nodes.Enqueue(child);
                            parentNode.InsertBefore(child, node);
                        }
                    }

                    parentNode.RemoveChild(node);

                }
            }
            return document.DocumentNode.InnerHtml;
        }


        /// Lines 135 through 235: Sign in, data gathering
        /// Call AcquireTokenAsync - to acquire a token requiring user to sign-in
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationResult authResult = null;

            try
            {
                authResult = await App.PublicClientApp.AcquireTokenSilentAsync(_scopes, App.PublicClientApp.Users.FirstOrDefault());
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilentAsync. This indicates you need to call AcquireTokenAsync to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await App.PublicClientApp.AcquireTokenAsync(_scopes);
                }
                catch (MsalException msalex)
                {
                    ResultText.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return;
            }

            if (authResult != null)
            {

                var resultBeforeStringify = await GetHttpContentWithToken(_graphAPIEndpoint, authResult.AccessToken);
                dynamic resultAfterStringify = Newtonsoft.Json.JsonConvert.DeserializeObject(resultBeforeStringify);
               
                resultBeforeStringify = resultBeforeStringify.Replace(@"\r\n", "");
                string cleanedString = RemoveUnwantedTags(resultBeforeStringify);

                JToken token = JObject.Parse(cleanedString);
                MainWindow.globalToken = JObject.Parse(cleanedString);
                
                string count5 = (string)token.SelectToken("value.@odata.count");
                string body = (string)token.SelectToken("value[0].body.content");
                string subject = (string)token.SelectToken("value[0].subject");
                string emailName = (string)token.SelectToken("value[0].from.emailAddress.name");
                string emailAddress = (string)token.SelectToken("value[0].from.emailAddress.address");

                string count6 = GetBetween(cleanedString, "odata.count\":", ",\"@odata.nextLink"); 
                int emailCount = Int32.Parse(count6);
                globalEmailCount = emailCount;
                if (globalEmailCount > 15)
                {
                    globalEmailCount = 15;
                }

               
                string address = "";

                string name = BodyGetBetween(body, "Caller :", "Phone:");
                string address2 = BodyGetBetween(body, "Address:", "Within 1/4 mile:");
                string address3 = BodyGetBetween(body, "Address :", "Within 1/4 mile:");
                if (address2 == "")
                {
                    address = address3.Replace("Street :", " ");
                }
                else
                {
                    address = address2.Replace("Street :", " ");
                }


               
                string workType = BodyGetBetween(body, "Work type :", "Done for :");
                string contact = BodyGetBetween(body, "Contact:", "- CELL");
                string location = BodyGetBetween(body, "Location:", "Grids :");
                string startDate = BodyGetBetween(body, "Start date:", "Time:");

                DisplayInfo(name, address, workType, contact, location, startDate);

                

                foreach (JObject x in token.SelectToken("value"))
                {
                    JToken type = x.SelectToken("subject");
                    string typeStr = type.ToString().ToLower();
                }

              
                string displayEmailCount = (globalEmailCount - 1).ToString();
                globalEmailsRemaining = globalEmailCount - 1;
                ResultText.Text = displayEmailCount;
               





                this.SignOutButton.Visibility = Visibility.Visible;
            }

            browser.Navigate("http://www.google.com");
        }

        /// Lines 238 through 274: Next button logic
        public async void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (globalIndex <= globalEmailCount - 2)
            {
                globalIndex = globalIndex + 1;
               
                globalEmailsRemaining = globalEmailsRemaining - 1;
                ResultText.Text = globalEmailsRemaining.ToString();
                string body = (string)MainWindow.globalToken.SelectToken("value[" + globalIndex + "].body.content");
                string subject = (string)MainWindow.globalToken.SelectToken("value[" + globalIndex + "].subject");
                string emailName = (string)MainWindow.globalToken.SelectToken("value[" + globalIndex + "].from.emailAddress.name");
                string emailAddress = (string)MainWindow.globalToken.SelectToken("value[" + globalIndex + "].from.emailAddress.address");

               
                string address = "";

                string name = BodyGetBetween(body, "Caller :", "Phone:");
                string address2 = BodyGetBetween(body, "Address:", "Within 1/4 mile:");
                string address3 = BodyGetBetween(body, "Address :", "Within 1/4 mile:");
                if (address2 == "")
                {
                    address = address3.Replace("Street :", " ");
                }
                else
                {
                    address = address2.Replace("Street :", " ");
                }

             
                string workType = BodyGetBetween(body, "Work type :", "Done for :");
                string contact = BodyGetBetween(body, "Contact:", "- CELL");
                string location = BodyGetBetween(body, "Location:", "Grids :");
                string startDate = BodyGetBetween(body, "Start date:", "Time:");

                DisplayInfo(name, address, workType, contact, location, startDate);
            }
        }
        /// Lines 276 through 314: Previous button logic
        public async void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            if (globalIndex != 0)
            {
                globalIndex = globalIndex - 1;
              
                globalEmailsRemaining = globalEmailsRemaining + 1;
                ResultText.Text = globalEmailsRemaining.ToString();

                string body = (string)MainWindow.globalToken.SelectToken("value[" + globalIndex + "].body.content");
                string subject = (string)MainWindow.globalToken.SelectToken("value[" + globalIndex + "].subject");
                string emailName = (string)MainWindow.globalToken.SelectToken("value[" + globalIndex + "].from.emailAddress.name");
                string emailAddress = (string)MainWindow.globalToken.SelectToken("value[" + globalIndex + "].from.emailAddress.address");

            
                string address = "";

                string name = BodyGetBetween(body, "Caller :", "Phone:");
                string address2 = BodyGetBetween(body, "Address:", "Within 1/4 mile:");
                string address3 = BodyGetBetween(body, "Address :", "Within 1/4 mile:");
                if (address2 == "")
                {
                    address = address3.Replace("Street :", " ");
                }
                else
                {
                    address = address2.Replace("Street :", " ");
                }


            
                string workType = BodyGetBetween(body, "Work type :", "Done for :");
                string contact = BodyGetBetween(body, "Contact:", "- CELL");
                string location = BodyGetBetween(body, "Location:", "Grids :");
                string startDate = BodyGetBetween(body, "Start date:", "Time:");

                DisplayInfo(name, address, workType, contact, location, startDate);
            }
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        /// <summary>
        /// Lines 344 through 372: Sign out the current user
        /// </summary>
        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            if (App.PublicClientApp.Users.Any())
            {
                try
                {
                    App.PublicClientApp.Remove(App.PublicClientApp.Users.FirstOrDefault());
                    this.ResultText.Text = "User has signed-out";
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                    this.nameText.Text = "";
                    this.addressText.Text = "";
                    this.workTypeText.Text = "";
                    this.contactText.Text = "";
                    this.locationText.Text = "";
                    this.startText.Text = "";
                    addressText.Background = Brushes.White;
                    workTypeText.Background = Brushes.White;
                    contactText.Background = Brushes.White;
                    locationText.Background = Brushes.White;
                    startText.Background = Brushes.White;
                    nameText.Background = Brushes.White;
                }
                catch (MsalException ex)
                {
                    ResultText.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }

        /// <summary>
        /// Lines 379 through 493: Display basic information contained in the body, change text box color on click, code for about and instructions drop downs
        /// </summary>
        /// 

        private void DisplayInfo(string name, string address, string wType, string contPerson, string location, string strDate)
        {
            nameText.Text = $"{name}";
            addressText.Text = $"{address}";
            workTypeText.Text = $"{wType}";
            contactText.Text = $"{contPerson}";
            locationText.Text = $"{location}";
            startText.Text = $"{strDate}";
        }

        private void exitBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(addressText.Text);
            addressText.Background = Brushes.Red;
            workTypeText.Background = Brushes.White;
            contactText.Background = Brushes.White;
            locationText.Background = Brushes.White;
            startText.Background = Brushes.White;
            nameText.Background = Brushes.White;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(workTypeText.Text);
            addressText.Background = Brushes.White;
            workTypeText.Background = Brushes.Red;
            contactText.Background = Brushes.White;
            locationText.Background = Brushes.White;
            startText.Background = Brushes.White;
            nameText.Background = Brushes.White;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(contactText.Text);
            addressText.Background = Brushes.White;
            workTypeText.Background = Brushes.White;
            contactText.Background = Brushes.Red;
            locationText.Background = Brushes.White;
            startText.Background = Brushes.White;
            nameText.Background = Brushes.White;
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(locationText.Text);
            addressText.Background = Brushes.White;
            workTypeText.Background = Brushes.White;
            contactText.Background = Brushes.White;
            locationText.Background = Brushes.Red;
            startText.Background = Brushes.White;
            nameText.Background = Brushes.White;
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(startText.Text);
            addressText.Background = Brushes.White;
            workTypeText.Background = Brushes.White;
            contactText.Background = Brushes.White;
            locationText.Background = Brushes.White;
            startText.Background = Brushes.Red;
            nameText.Background = Brushes.White;
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(nameText.Text);
            addressText.Background = Brushes.White;
            workTypeText.Background = Brushes.White;
            contactText.Background = Brushes.White;
            locationText.Background = Brushes.White;
            startText.Background = Brushes.White;
            nameText.Background = Brushes.Red;
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {

            MessageBox.Show("Email Excavator is a WPF application coded in C# by three students at Purdue University Northwest" +
                " for the Michigan City Sanitary Department. Email Excavator uses Microsoft Outlook REST API to upload 811 Locate emails"
                + " into the program. Then, by using get between functions, the application extracts key information from the emails and places" +
                " the info in text boxes to then be entered into the departments Web Based Order system. The application uses a Web Browser Control" +
                " to allow the department to easily put the info into the Web Based Order system.", "About Email Excavator");


        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("1. Click Upload Emails Button" +
                "\n" +
                "2. Login to Outlook Email Account" +
                "\n" +
                "3. Allow application access to emails when prompted" +
                "\n" +
                "4. Drag/Drop or Copy/Paste information into Web Based Order System" +
                "\n" +
                "5. Use next email button to move to next email" +
                "\n" +
                "6. Repeat steps 4 and 5 until all emails are entered" +
                "\n" +
                "7. Click sign out button to log out" +
                "\n" +
                "8. Use Exit tab in navigation bar to exit the application", 
                "Email Excavator Instructions");
        }
    }
}
