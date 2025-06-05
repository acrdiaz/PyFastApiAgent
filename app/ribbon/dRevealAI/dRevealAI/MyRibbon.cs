using dReveal.Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace dRevealAI
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private readonly AIServiceProvider _aiServiceProvider;

        public MyRibbon()
        {
            _aiServiceProvider = new AIServiceProvider();
        }

        public async void MyButton_Click(Office.IRibbonControl control)
        {
            //MessageBox.Show("Hi", "My Add-in", MessageBoxButtons.OK, MessageBoxIcon.Information);
            try
            {
                // Disable the button during operation
                ribbon.InvalidateControl(control.Id);

                Outlook.MailItem mailItem = GetSelectedMailItem();
                if (mailItem != null)
                {
                    await ProcessEmailWithAI(mailItem);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Re-enable the button
                ribbon.InvalidateControl(control.Id);
            }
        }

        private Outlook.MailItem GetSelectedMailItem()
        {
            try
            {
                Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector?.CurrentItem is Outlook.MailItem mailItem)
                {
                    return mailItem;
                }
            }
            catch { /* Ignore errors */ }

            MessageBox.Show("Please open an email first", "No Email Selected",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }

        private async Task ProcessEmailWithAI(Outlook.MailItem mailItem)
        {
            try
            {
                var aiService = _aiServiceProvider.GetDefaultService();
                string emailContent = mailItem.Body;

                // Show processing indicator
                using (var progressForm = new Form { Text = "Processing...", Width = 300, Height = 100 })
                {
                    progressForm.Show();
                    progressForm.Refresh();

                    // Process with AI (async)
                    string analysis = await aiService.AnalyzeContentAsync(emailContent);

                    // Insert analysis at top of email
                    mailItem.Body = $"AI ANALYSIS:\n{analysis}\n\n{emailContent}";
                    mailItem.Save();

                    progressForm.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"AI Processing Error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region working
        //private static async Task Add()
        //{
        //    using (var client = new HttpClient())
        //    {
        //        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {DR_APIKEY}");

        //        var num1 = 5;
        //        var num2 = 3;

        //        //var requestData = new
        //        //{
        //        //    model = DR_MODEL,
        //        //    messages = new[]
        //        //    {
        //        //        new { role = "system", content = "You are a bot that performs addition." },
        //        //        new { role = "user", content = $"What is {num1} + {num2}?" }
        //        //    }
        //        //};
        //        var requestData = new
        //        {
        //            prompt = $"What is {num1} + {num2}?"
        //        };

        //        var requestBody = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");

        //        var response = await client.PostAsync(DR_API_PROMPT, requestBody);
        //        if (!response.IsSuccessStatusCode)
        //        {
        //            throw new Exception($"API request failed with status: {response.StatusCode}");
        //        }
        //        var responseContent = await response.Content.ReadAsStringAsync();

        //        //dynamic responseJson = Newtonsoft.Json.JsonConvert.DeserializeObject(responseContent);
        //        //var generatedText = responseJson.choices[0].message.content;

        //        string generatedText = string.Empty;
        //        if (responseContent.Contains("success"))
        //        {
        //            bool loop = true;
        //            while (loop)
        //            {
        //                response = await client.GetAsync(DR_API_RESPONSE);
        //                if (!response.IsSuccessStatusCode)
        //                {
        //                    throw new Exception($"API request failed with status: {response.StatusCode}");
        //                }
        //                responseContent = await response.Content.ReadAsStringAsync();

        //                if (responseContent.Contains("No response available."))
        //                {
        //                    continue;
        //                }
        //                {
        //                    loop = false; // Exit loop if no success in response
        //                }

        //                var result = JsonConvert.DeserializeObject<ApiResponse>(responseContent);
        //                string message = result.message;
        //                MessageBox.Show("Generated text: " + message);
        //            }
        //        }
        //    }
        //}
        #endregion working

        private static async Task Add()
        {
            try
            {
                var num1 = 5;
                var num2 = 3;

                var requestData = new
                {
                    prompt = $"What is {num1} + {num2}?"
                };

                string message = await ApiHelper.GetApiResponseWithPollingAsync(
                    DR_API_PROMPT,
                    DR_API_RESPONSE,
                    DR_API_CLEAR,
                    requestData);

                MessageBox.Show("Generated text: " + message);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
        }

        #region delete
        //public async void GetInboxEmails_Click(Office.IRibbonControl control)
        //{
        //    try
        //    {
        //        //string folderName = "Drafts"; // Specify the folder name you want to access
        //        Outlook.OlDefaultFolders folderType = Outlook.OlDefaultFolders.olFolderDrafts;

        //        Outlook.Application outlookApp = new Outlook.Application();
        //        Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

        //        Outlook.MAPIFolder emailFolder = outlookNs.GetDefaultFolder(folderType);

        //        Outlook.Items emailItems = emailFolder.Items;

        //        // Sort by received time (newest first)
        //        emailItems.Sort("[ReceivedTime]", true);

        //        // Build the list of subjects
        //        //StringBuilder sb = new StringBuilder();
        //        //sb.AppendLine($"{folderName} Emails:");
        //        //sb.AppendLine("====================");

        //        foreach (object item in emailItems)
        //        {
        //            if (item is Outlook.MailItem mail)
        //            {
        //                //sb.AppendLine($"- {mail.Subject}");
        //                await HandleEmail(mail);
        //                Marshal.ReleaseComObject(mail); // Important for Outlook COM objects
        //            }
        //        }

        //        // Show the results
        //        //MessageBox.Show(sb.ToString(), "Draft Emails",
        //        //    MessageBoxButtons.OK, MessageBoxIcon.Information);

        //        // Clean up COM objects
        //        Marshal.ReleaseComObject(emailItems);
        //        Marshal.ReleaseComObject(emailFolder);
        //        Marshal.ReleaseComObject(outlookNs);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Error accessing specified email folder: {ex.Message}", "Error",
        //            MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        #endregion

        public async Task GetGeneralSummary_ClickAsync(Office.IRibbonControl control)
        {
            Outlook.Application outlookApp = null;
            Outlook.NameSpace outlookNs = null;
            Outlook.MAPIFolder emailFolder = null;
            Outlook.Items emailItems = null;

            try
            {
                Outlook.OlDefaultFolders folderType = Outlook.OlDefaultFolders.olFolderDrafts;
                outlookApp = new Outlook.Application();
                outlookNs = outlookApp.GetNamespace("MAPI");
                emailFolder = outlookNs.GetDefaultFolder(folderType);
                emailItems = emailFolder.Items;

                // Sort by received time (newest first)
                emailItems.Sort("[ReceivedTime]", true);

                StringBuilder sb = new StringBuilder();
                //sb.AppendLine("====================");

                // Process items
                foreach (object item in emailItems)
                {
                    if (item is Outlook.MailItem mail)
                    {
                        try
                        {
                            sb.AppendLine($"Email To: {mail.To}");
                            sb.AppendLine($"Subject: {mail.Subject}");
                            //sb.AppendLine($"...: {mail.RetentionExpirationDate}");
                            //sb.AppendLine($"...: {mail.CreationTime}");
                            //sb.AppendLine($"...: {mail.ReceivedTime}");
                            sb.AppendLine($".\n");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(mail);
                        }
                    }
                }

                await HandleGeneralEmail(sb.ToString());
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error accessing specified email folder: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up COM objects in reverse order
                if (emailItems != null) Marshal.ReleaseComObject(emailItems);
                if (emailFolder != null) Marshal.ReleaseComObject(emailFolder);
                if (outlookNs != null) Marshal.ReleaseComObject(outlookNs);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }
        }

        private static async Task HandleGeneralEmail(string EmailList)
        {
            try
            {
                var requestData = new
                {
                    prompt = $"Give a general summary of this email list, how many are from each person: {EmailList}?"
                };

                string message = await ApiHelper.GetApiResponseWithPollingAsync(
                    DR_API_PROMPT,
                    DR_API_RESPONSE,
                    DR_API_CLEAR,
                    requestData);

                MessageBox.Show(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
        }



        public async Task GetInboxEmails_ClickAsync(Office.IRibbonControl control)
        {
            Outlook.Application outlookApp = null;
            Outlook.NameSpace outlookNs = null;
            Outlook.MAPIFolder emailFolder = null;
            Outlook.Items emailItems = null;

            try
            {
                Outlook.OlDefaultFolders folderType = Outlook.OlDefaultFolders.olFolderDrafts;
                outlookApp = new Outlook.Application();
                outlookNs = outlookApp.GetNamespace("MAPI");
                emailFolder = outlookNs.GetDefaultFolder(folderType);
                emailItems = emailFolder.Items;

                // Sort by received time (newest first)
                emailItems.Sort("[ReceivedTime]", true);

                // Process items
                foreach (object item in emailItems)
                {
                    if (item is Outlook.MailItem mail)
                    {
                        try
                        {
                            await HandleEmail(mail);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(mail);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error accessing specified email folder: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up COM objects in reverse order
                if (emailItems != null) Marshal.ReleaseComObject(emailItems);
                if (emailFolder != null) Marshal.ReleaseComObject(emailFolder);
                if (outlookNs != null) Marshal.ReleaseComObject(outlookNs);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }
        }

        private static async Task HandleEmail(Outlook.MailItem mail)
        {
            try
            {
                string emailText = $"Subject: {mail.Subject}\n.EmailMessage: {mail.Body}";

                var requestData = new
                {
                    prompt = $"Summarize one sentence this email: {emailText}?"
                };

                string message = await ApiHelper.GetApiResponseWithPollingAsync(
                    DR_API_PROMPT,
                    DR_API_RESPONSE,
                    DR_API_CLEAR,
                    requestData);

                HandleResponse(message, mail.Subject);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
        }

        private static void HandleResponse(string message, string subject)
        {
            //MessageBox.Show("Generated text:\n" + message);
            Outlook.Application outlookApp = null;
            Outlook.MailItem newMail = null;

            try
            {
                // Get the Outlook application
                outlookApp = new Outlook.Application();

                // Create a new mail item
                newMail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                // Set the message body
                newMail.Body = message;

                // You can set other properties as needed
                newMail.Subject = $"{subject} -- AI";
                // newMail.To = "recipient@example.com";
                // newMail.CC = "cc@example.com";

                // Display the email (opens in Outlook's UI)
                newMail.Display(false); // false = non-modal window

                // Alternatively, to send immediately:
                // newMail.Send();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error creating email: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up COM objects
                if (newMail != null) Marshal.ReleaseComObject(newMail);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }
        }

        #region delete
        //sb.AppendLine($"- {mail.Subject}");
        //sb.AppendLine($"- {mail.Body}");
        //MessageBox.Show(response, "Draft Emails",
        //                MessageBoxButtons.OK, MessageBoxIcon.Information);


        //private async Task GetInboxEmails()
        //{

        //}

        //public Bitmap GetButtonImage(Office.IRibbonControl control)
        //{
        //    try
        //    {
        //        MessageBox.Show("Current control ID: " + control.Id);

        //        switch (control.Id)
        //        {
        //            case "MyButton":
        //                return Properties.Resources.MyButtonIcon; // If using project resources
        //                                                          // OR for embedded resources:
        //                                                          // return LoadImageFromResource("MyOfficeAddIn.Images.happy.png");

        //            case "ListInboxButton":
        //                return Properties.Resources.DraftsIcon;
        //            // OR:
        //            // return LoadImageFromResource("MyOfficeAddIn.Images.drafts.png");

        //            default:
        //                return null;
        //        }
        //    }
        //    catch
        //    {
        //        return null;
        //    }
        //}
        #endregion

        private Bitmap LoadImageFromResource(string resourcePath)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourcePath))
            {
                return stream != null ? new Bitmap(stream) : null;
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("dRevealAI.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    public class ApiResponse
    {
        public string message { get; set; }
    }

    public static class ApiHelper
    {
        private static readonly HttpClient client = new HttpClient();

        static ApiHelper()
        {
            //client.DefaultRequestHeaders.Add("Authorization", $"Bearer {DR_APIKEY}");
        }

        public static async Task<string> PostApiRequestAsync(string endpoint, object requestData)
        {
            var requestBody = new StringContent(
                JsonConvert.SerializeObject(requestData),
                Encoding.UTF8,
                "application/json");

            var response = await client.PostAsync(endpoint, requestBody);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        public static async Task<string> GetApiRequestAsync(string endpoint)
        {
            var response = await client.GetAsync(endpoint);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        public static async Task<string> GetApiResponseWithPollingAsync(string promptEndpoint, string responseEndpoint, string clearEndpoint, object requestData)
        {
            var responseContent = await PostApiRequestAsync(clearEndpoint, null);
            await Task.Delay(700);

            // Initial prompt request
            responseContent = await PostApiRequestAsync(promptEndpoint, requestData);

            if (responseContent.Contains("success"))
            {
                while (true)
                {
                    // Poll for response
                    responseContent = await GetApiRequestAsync(responseEndpoint);

                    if (!responseContent.Contains("No response available."))
                    {
                        var result = JsonConvert.DeserializeObject<ApiResponse>(responseContent);
                        return result.message;
                    }

                    await Task.Delay(1000); // Add delay between polling attempts
                }
            }

            throw new Exception("Initial API request didn't return success status");
        }
    }
}
